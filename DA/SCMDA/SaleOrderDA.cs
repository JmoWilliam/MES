using Dapper;
using Helpers;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Transactions;
using System.Web;

namespace SCMDA
{
    public class SaleOrderDA
    {
        public string MainConnectionStrings = "";
        public string ErpConnectionStrings = "";
        public string HrmConnectionStrings = "";
        public string BpmDbConnectionStrings = "";
        public string OfficialConnectionStrings = "";

        public string BpmServerPath = "";
        public string BpmAccount = "";
        public string BpmPassword = "";

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
        public BpmHelper bpmHelper = new BpmHelper();
        public MamoHelper mamoHelper = new MamoHelper();


        public SaleOrderDA()
        {
            MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];
            HrmConnectionStrings = ConfigurationManager.AppSettings["HrmDb"];
            BpmDbConnectionStrings = ConfigurationManager.AppSettings["BpmDb"];
            OfficialConnectionStrings = ConfigurationManager.AppSettings["OfficialDb"];

            BpmServerPath = ConfigurationManager.AppSettings["BpmServerPath"];
            BpmAccount = ConfigurationManager.AppSettings["BpmAccount"];
            BpmPassword = ConfigurationManager.AppSettings["BpmPassword"];

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
        #region //GetSaleOrder -- 取得訂單單頭資料 -- Zoey 2022.07.07
        public string GetSaleOrder(int SoId, string SoErpPrefix, string SoErpNo, string SoErpFullNo, int CustomerId
            , int SalesmenId, string MtlItemNo, string StartDate, string EndDate, string ConfirmStatus, string ClosureStatus
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.SoId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.CompanyId, a.DepartmentId, a.SoErpPrefix, a.SoErpNo, a.Version, FORMAT(a.DocDate, 'yyyy-MM-dd') DocDate
                        , FORMAT(a.SoDate, 'yyyy-MM-dd') SoDate, a.SoRemark, a.CustomerId, a.SalesmenId, a.CustomerAddressFirst
                        , a.CustomerAddressSecond, a.CustomerPurchaseOrder, a.DepositPartial, a.DepositRate, a.Currency
                        , ROUND(a.ExchangeRate, 3) ExchangeRate, ROUND(a.BusinessTaxRate, 2) BusinessTaxRate
                        , a.TaxNo, a.Taxation, a.DetailMultiTax, a.TotalQty, a.Amount, a.TaxAmount, a.ShipMethod
                        , a.TradeTerm, a.PaymentTerm, a.PriceTerm, a.ConfirmStatus, a.ConfirmUserId, a.TransferStatus, a.TransferDate
                        , a.SoErpPrefix + '-' + a.SoErpNo SoErpFullNo
                        , ISNULL(b.CustomerName,'') CustomerName, b.CustomerNo + ' ' + b.CustomerShortName CustomerWithNo
                        , ISNULL(c.UserNo, '') SalesmenNo, ISNULL(c.UserName,'') SalesmenName, c.Gender
                        , ISNULL(d.SomId, '') SomId
                        , e.StatusName ConfirmName
                        , f.StatusName TransferName
                        , g.TotalAmount
                        , ISNULL(h.UserNo, '') ConfirmUserNo, ISNULL(h.UserName,'') ConfirmUserName, ISNULL(h.Gender, '') ConfirmUserGender
                        , (
                            SELECT x.SoSequence, ISNULL(y.MtlItemNo, '') MtlItemNo, ISNULL(x.SoMtlItemName, '') MtlItemName
                            , ISNULL(x.CustomerMtlItemNo, '') CustomerMtlItemNo, ISNULL(y.MtlItemSpec, '') MtlItemSpec
                            , x.SoQty, x.UnitPrice, x.Amount, ISNULL(FORMAT(x.PromiseDate, 'yyyy-MM-dd'), '') PromiseDate
                            FROM SCM.SoDetail x
                            LEFT JOIN PDM.MtlItem y ON y.MtlItemId = x.MtlItemId
                            WHERE x.SoId = a.SoId
                            FOR JSON PATH, ROOT('data')
                        ) SoDetail";
                    sqlQuery.mainTables =
                        @"FROM SCM.SaleOrder a
                        LEFT JOIN SCM.Customer b ON b.CustomerId = a.CustomerId
                        LEFT JOIN BAS.[User] c ON c.UserId = a.SalesmenId
                        LEFT JOIN SCM.SoModification d ON d.SoId = a.SoId
                        LEFT JOIN BAS.[Status] e ON e.StatusSchema = 'ConfirmStatus' AND e.StatusNo = a.ConfirmStatus
                        LEFT JOIN BAS.[Status] f ON f.StatusSchema = 'Boolean' AND f.StatusNo = a.TransferStatus
                        OUTER APPLY(
                            SELECT CASE a.Taxation
                                WHEN 1 THEN ROUND(ISNULL(SUM(x.Amount), 0) / (1 + a.BusinessTaxRate) + (ISNULL(SUM(x.Amount), 0) / (1 + a.BusinessTaxRate) * a.BusinessTaxRate), 2)
                                ELSE ISNULL(SUM(x.Amount), 0) + (ISNULL(SUM(x.Amount), 0) * a.BusinessTaxRate)
                                END TotalAmount
                            FROM SCM.SoDetail x
                            WHERE x.SoId = a.SoId
                        ) g
                        LEFT JOIN BAS.[User] h ON a.ConfirmUserId = h.UserId";
                    string queryTable =
                        @"FROM (
                            SELECT a.SoId, a.CompanyId, a.SoErpPrefix, a.SoErpNo, FORMAT(a.SoDate, 'yyyy-MM-dd') SoDate, a.CustomerId, a.SalesmenId, a.ConfirmStatus
                            ,(
                                SELECT x.SoSequence, y.MtlItemNo, x.SoMtlItemName MtlItemName, x.CustomerMtlItemNo
                                , x.SoQty, ISNULL(FORMAT(x.PromiseDate, 'yyyy-MM-dd'), '') PromiseDate, x.ClosureStatus
                                FROM SCM.SoDetail x
                                LEFT JOIN PDM.MtlItem y ON y.MtlItemId = x.MtlItemId
                                WHERE x.SoId = a.SoId
                                FOR JSON PATH, ROOT('data')
                            ) SoDetail
                            FROM SCM.SaleOrder a
                        ) a 
                        OUTER APPLY (
                            SELECT TOP 1 x.MtlItemNo, x.MtlItemName, x.CustomerMtlItemNo, x.ClosureStatus
                            FROM OPENJSON(a.SoDetail, '$.data')
                            WITH (
                                MtlItemNo NVARCHAR(40) N'$.MtlItemNo', 
                                MtlItemName NVARCHAR(120) N'$.MtlItemName', 
                                CustomerMtlItemNo NVARCHAR(40) N'$.CustomerMtlItemNo',
                                ClosureStatus NVARCHAR(2) N'$.ClosureStatus'
                            ) x
                        )b";

                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SoId", @" AND a.SoId = @SoId", SoId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SoErpPrefix", @" AND a.SoErpPrefix = @SoErpPrefix", SoErpPrefix);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SoErpNo", @" AND a.SoErpNo LIKE '%' + @SoErpNo + '%'", SoErpNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SoErpFullNo", @" AND (a.SoErpPrefix + '-' + a.SoErpNo) LIKE '%' + @SoErpFullNo + '%'", SoErpFullNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerId", @" AND a.CustomerId = @CustomerId", CustomerId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SalesmenId", @" AND a.SalesmenId = @SalesmenId", SalesmenId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemNo", @" AND (b.MtlItemNo LIKE '%' + @MtlItemNo + '%' OR b.MtlItemName LIKE '%' + @MtlItemNo + '%' OR b.CustomerMtlItemNo LIKE '%' + @MtlItemNo + '%')", MtlItemNo);
                    if (ConfirmStatus.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ConfirmStatus", @" AND a.ConfirmStatus IN @ConfirmStatus", ConfirmStatus.Split(','));
                    if (ClosureStatus.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ClosureStatus", @" AND ISNULL(b.ClosureStatus, 'N') IN @ClosureStatus", ClosureStatus.Split(','));
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.SoDate >= @StartDate ", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.SoDate <= @EndDate ", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.SoId DESC";
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

        #region //GetSaleOrder02 -- 取得訂單資料 -- Zoey 2022.07.07
        public string GetSaleOrder02(int SoId, string SoErpPrefix, string SoErpNo, int CustomerId
            , int SalesmenId, string SearchKey, string ConfirmStatus, string ClosureStatus
            , string StartDate, string EndDate
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                List<SaleOrder> sos = new List<SaleOrder>();

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //判斷公司資料是否正確
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 CompanyId, CompanyName, ErpDb
                            FROM BAS.Company
                            WHERE CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0) throw new SystemException("【公司】資料錯誤!");

                    foreach (var item in result)
                    {
                        ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                    }
                    #endregion

                    sqlQuery.mainKey = "a.SoId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.CompanyId, a.DepartmentId, a.SoErpPrefix, a.SoErpNo, FORMAT(a.DocDate, 'yyyy-MM-dd') DocDate
                        , a.SoRemark, a.CustomerId, a.SalesmenId, a.CustomerPurchaseOrder, a.DepositPartial, a.Currency
                        , a.ConfirmStatus, a.ConfirmUserId, a.TransferStatus, FORMAT(a.TransferDate, 'yyyy-MM-dd') TransferDate
                        , a.SoErpPrefix + '-' + a.SoErpNo SoErpFullNo, a.Amount
                        , b.CustomerNo, b.CustomerShortName
                        , ISNULL(c.UserNo, '') SalesmenNo, ISNULL(c.UserName, '') SalesmenName, ISNULL(c.Gender, '') SalesmenGender
                        , (
                            SELECT aa.SoDetailId, aa.SoSequence, aa.MtlItemId, aa.SoMtlItemName, aa.SoMtlItemSpec
                            , aa.SoQty, aa.ProductType, ISNULL(aa.FreebieQty, '0') FreebieQty, ISNULL(aa.SpareQty, '0') SpareQty
                            , aa.UnitPrice, aa.Amount, FORMAT(aa.PromiseDate, 'yyyy-MM-dd') PromiseDate
                            , aa.SoDetailRemark, aa.ClosureStatus
                            , ISNULL(ab.MtlItemNo, '') MtlItemNo, ac.TypeName ProductTypeName
                            , dbo.MassProductionReviewDocVerify(ab.MtlItemNo) DocVerify
                            FROM SCM.SoDetail aa
                            LEFT JOIN PDM.MtlItem ab ON aa.MtlItemId = ab.MtlItemId
                            INNER JOIN BAS.[Type] ac ON aa.ProductType = ac.TypeNo AND ac.TypeSchema = 'SoDetail.ProductType'
                            WHERE aa.SoId = a.SoId
                            ORDER BY aa.SoSequence
                            FOR JSON PATH, ROOT('data')
                        ) SoDetail
                        , ISNULL(d.UserNo, '') ConfirmUserNo, ISNULL(d.UserName, '') ConfirmUserName, ISNULL(d.Gender, '') ConfirmUserGender
                        , e.BpmTransferStatus
                        , ISNULL(f.StatusName, '尚未拋轉BPM') BpmTransferStatusName";
                    sqlQuery.mainTables =
                        @"FROM SCM.SaleOrder a
                        INNER JOIN SCM.Customer b ON a.CustomerId = b.CustomerId
                        INNER JOIN BAS.[User] c ON a.SalesmenId = c.UserId
                        LEFT JOIN BAS.[User] d ON a.ConfirmUserId = d.UserId
                        LEFT JOIN SCM.SoBpmInfo e ON a.SoId = e.SoId
                        LEFT JOIN BAS.[Status] f ON e.BpmTransferStatus = f.StatusNo AND f.StatusSchema = 'SoBpmInfo.BpmTransferStatus'";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SoId", @" AND a.SoId = @SoId", SoId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SoErpPrefix", @" AND a.SoErpPrefix = @SoErpPrefix", SoErpPrefix);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SoErpNo", @" AND a.SoErpNo LIKE '%' + @SoErpNo + '%'", SoErpNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerId", @" AND a.CustomerId = @CustomerId", CustomerId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SalesmenId", @" AND a.SalesmenId = @SalesmenId", SalesmenId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey", @" AND EXISTS (
                                                                                                            SELECT TOP 1 1
                                                                                                            FROM SCM.SoDetail aa
                                                                                                            LEFT JOIN PDM.MtlItem ab ON aa.MtlItemId = ab.MtlItemId
                                                                                                            WHERE aa.SoId = a.SoId
                                                                                                            AND (ab.MtlItemNo LIKE '%' + @SearchKey + '%' OR aa.SoMtlItemName LIKE '%' + @SearchKey + '%')
                                                                                                       )", SearchKey);
                    if (ConfirmStatus.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ConfirmStatus", @" AND a.ConfirmStatus IN @ConfirmStatus", ConfirmStatus.Split(','));
                    if (ClosureStatus.Length > 0)
                    {
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ClosureStatus", @" AND EXISTS (
                                                                                                                    SELECT TOP 1 1
                                                                                                                    FROM SCM.SoDetail aa
                                                                                                                    WHERE aa.SoId = a.SoId
                                                                                                                    AND aa.ClosureStatus IN @ClosureStatus
                                                                                                               )", ClosureStatus.Split(','));
                    }                    
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.DocDate >= @StartDate ", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.DocDate <= @EndDate ", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.DocDate DESC, a.SoErpPrefix, a.SoErpNo DESC";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    sos = BaseHelper.SqlQuery<SaleOrder>(sqlConnection, dynamicParameters, sqlQuery);
                }

                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                {
                    List<ErpDocStatus> erpDocStatuses = new List<ErpDocStatus>();

                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT LTRIM(RTRIM(TC001)) ErpPrefix, LTRIM(RTRIM(TC002)) ErpNo, TC027 DocStatus
                            FROM COPTC
                            WHERE (TC001 + '-' + TC002) IN @ErpFullNo";
                    dynamicParameters.Add("ErpFullNo", sos.Select(x => x.SoErpPrefix + '-' + x.SoErpNo).ToArray());
                    erpDocStatuses = sqlConnection.Query<ErpDocStatus>(sql, dynamicParameters).ToList();
                    sos = sos.GroupJoin(erpDocStatuses, x => x.SoErpPrefix + '-' + x.SoErpNo, y => y.ErpPrefix + '-' + y.ErpNo, (x, y) => { x.DocStatus = y.FirstOrDefault()?.DocStatus ?? ""; return x; }).ToList();
                    sos = sos.OrderByDescending(x => x.DocDate).ToList();
                    sos.ForEach(x =>
                    {
                        x.DocDateStr = x.DocDate?.ToString("yyyy-MM-dd");
                    });
                }

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    data = sos
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

        #region //GetSoStatus -- 取得訂單狀態 -- Ben Ma 2023.06.15
        public string GetSoStatus(int SoId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.ConfirmStatus, a.TransferStatus
                            FROM SCM.SaleOrder a
                            WHERE a.SoId = @SoId
                            AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("SoId", SoId);
                    dynamicParameters.Add("CompanyId", CurrentCompany);

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

        #region //GetSoDetail -- 取得訂單單身資料 -- Zoey 2022.07.12
        public string GetSoDetail(int SoDetailId, int SoId, int MtlItemId, string SoErpFullNo, string CustomerMtlItemNo, string TransferStatus, string SearchKey, string MtlItemIdIsNull, int CompanyId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.SoDetailId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.SoId, a.SoSequence, a.MtlItemId, ISNULL(a.SoMtlItemName, d.MtlItemName) MtlItemName
                          , ISNULL(a.SoMtlItemSpec, d.MtlItemSpec) MtlItemSpec, a.InventoryId, a.UomId, a.SoQty
                          , a.SiQty, a.ProductType, a.FreebieQty, a.FreebieSiQty, a.SpareQty, a.SpareSiQty
                          , a.UnitPrice, a.Amount, a.Project
                          , ISNULL(FORMAT(a.PromiseDate, 'yyyy-MM-dd'), '') PromiseDate
                          , ISNULL(FORMAT(a.PcPromiseDate, 'yyyy-MM-dd'), '') PcPromiseDate
                          , a.SoDetailRemark, a.SoPriceQty, a.SoPriceUomId, a.TaxNo
                          , a.BusinessTaxRate, a.DiscountRate, a.DiscountAmount, a.ConfirmStatus, a.ClosureStatus,a.QuotationErp
                          , b.SoErpPrefix, b.SoErpNo, b.TransferStatus, b.SoErpPrefix + '-' + b.SoErpNo SoErpFullNo
                          , c.CustomerNo, c.CustomerShortName, c.CustomerName
                          , ISNULL(d.MtlItemNo, '') MtlItemNo
                          , e.UomNo, f.UomNo SoPriceUomNo
                          , h.SoDetailTempId
                          , ISNULL(g.CustomerMtlItemNo, a.CustomerMtlItemNo) CustomerMtlItemNo";
                    sqlQuery.mainTables =
                          @"FROM SCM.SoDetail a
                            INNER JOIN SCM.SaleOrder b ON a.SoId = b.SoId
                            INNER JOIN SCM.Customer c ON b.CustomerId = c.CustomerId
                            LEFT JOIN PDM.MtlItem d ON a.MtlItemId = d.MtlItemId
                            LEFT JOIN PDM.UnitOfMeasure e ON a.UomId = e.UomId
                            LEFT JOIN PDM.UnitOfMeasure f ON a.SoPriceUomId = f.UomId
                            LEFT JOIN SCM.SoDetailTemp h ON h.SoDetailId = a.SoDetailId
                            OUTER APPLY (
                                SELECT TOP 1 ga.CustomerMtlItemNo
                                FROM PDM.CustomerMtlItem ga
                                WHERE ga.CustomerId = c.CustomerId 
                                AND ga.MtlItemId = d.MtlItemId
                                ORDER BY ga.LastModifiedDate DESC
                            ) g";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    if (MtlItemIdIsNull == "Y") queryCondition += @" AND a.MtlItemId IS NULL";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SoDetailId", @" AND a.SoDetailId = @SoDetailId", SoDetailId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SoId", @" AND a.SoId = @SoId", SoId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemId", @" AND a.MtlItemId = @MtlItemId", MtlItemId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SoErpFullNo", @" AND (b.SoErpPrefix + '-' + b.SoErpNo) LIKE '%' + @SoErpFullNo + '%'", SoErpFullNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerMtlItemNo", @" AND a.CustomerMtlItemNo LIKE '%' + @CustomerMtlItemNo + '%'", CustomerMtlItemNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TransferStatus", @" AND b.TransferStatus = @TransferStatus", TransferStatus);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CompanyId", @" AND b.CompanyId = @CompanyId", CompanyId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey",
                          @" AND (d.MtlItemNo LIKE '%' + @SearchKey + '%' 
                          OR d.MtlItemName LIKE '%' + @SearchKey + '%' 
                          OR a.SoMtlItemName LIKE '%' + @SearchKey + '%' 
                          OR a.CustomerMtlItemNo LIKE '%' + @SearchKey + '%' 
                          OR g.CustomerMtlItemNo LIKE '%' + @SearchKey + '%')", SearchKey);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.SoId DESC, a.SoSequence";
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

        #region //GetSoDetail02 -- 取得訂單單身資料 -- Zoey 2022.07.12
        public string GetSoDetail02(int SoDetailId, int SoId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.SoDetailId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.SoId, a.SoSequence, ISNULL(a.MtlItemId, -1) MtlItemId, a.SoMtlItemName, a.SoMtlItemSpec
                        , a.InventoryId, ISNULL(a.UomId, -1) UomId, a.SoQty, a.SiQty, a.ProductType, e.TypeName ProductTypeName, a.FreebieQty, a.FreebieSiQty
                        , a.SpareQty, a.SpareSiQty, a.UnitPrice, a.Amount, a.BusinessTaxRate
                        , FORMAT(a.PromiseDate, 'yyyy-MM-dd') PromiseDate, FORMAT(a.PcPromiseDate, 'yyyy-MM-dd') PcPromiseDate
                        , a.Project, a.CustomerMtlItemNo, a.SoDetailRemark, a.SoPriceQty, a.SoPriceUomId, a.ClosureStatus,a.QuotationErp
                        , b.SoErpPrefix, b.SoErpNo, b.Currency, b.ConfirmStatus, b.TransferStatus, b.CustomerId
                        , ISNULL(c.UomNo, '') UomNo
                        , ISNULL(d.MtlItemNo, '') MtlItemNo
                        , f.SoDetailTempId";
                    sqlQuery.mainTables =
                        @"FROM SCM.SoDetail a
                        INNER JOIN SCM.SaleOrder b ON a.SoId = b.SoId
                        LEFT JOIN PDM.UnitOfMeasure c ON a.UomId = c.UomId
                        LEFT JOIN PDM.MtlItem d ON a.MtlItemId = d.MtlItemId
                        LEFT JOIN BAS.[Type] e ON a.ProductType = e.TypeNo AND e.TypeSchema = 'SoDetail.ProductType'
                        LEFT JOIN SCM.SoDetailTemp f ON f.SoDetailId = a.SoDetailId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND b.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SoDetailId", @" AND a.SoDetailId = @SoDetailId", SoDetailId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SoId", @" AND a.SoId = @SoId", SoId);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.SoId DESC, a.SoSequence";
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

        #region //GetSaleOrderSimple -- 取得訂單單頭資料(簡單) -- Shintokuro 2022-11-28
        public string GetSaleOrderSimple(int MtlItemId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    if (MtlItemId <= 0) throw new SystemException("缺少產品品號!!!");

                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sql = @"SELECT b.SoDetailId, a.SoErpPrefix + '-' + a.SoErpNo + '-' + b.SoSequence  SoErpFullNo
                            FROM SCM.SaleOrder a
                            INNER JOIN SCM.SoDetail b on a.SoId = b.SoId
                            INNER JOIN PDM.MtlItem c on b.MtlItemId = c.MtlItemId
                            WHERE a.CompanyId = @CompanyId
                            AND b.MtlItemId = @MtlItemId
                            AND a.ConfirmStatus = 'Y'
                            AND b.ClosureStatus = 'N'";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
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

        #region //GetTotal -- 取得訂單單身加總資料 -- Zoey 2022.07.13
        public string GetTotal(int SoId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                { 
                    sql = @"SELECT ISNULL(SUM(a.SoQty), 0) + ISNULL(SUM(a.FreebieQty), 0) + ISNULL(SUM(a.SpareQty), 0) TotalQty
                            , CASE b.Taxation
                                WHEN 1 THEN c.Amount
                                ELSE d.Amount
                                END Amount
                            , CASE b.Taxation
                                WHEN 1 THEN ROUND(c.Amount * b.BusinessTaxRate, 2)
                                ELSE ROUND(d.Amount * b.BusinessTaxRate, 2)
                                END TaxAmount
                            , CASE b.Taxation
                                WHEN 1 THEN ROUND(c.Amount + (c.Amount * b.BusinessTaxRate), 2)
                                ELSE d.Amount + (d.Amount * b.BusinessTaxRate)
                                END TotalAmount
                            FROM SCM.SoDetail a
                            INNER JOIN SCM.SaleOrder b ON b.SoId = a.SoId
                            OUTER APPLY(
                                SELECT ROUND(ISNULL(SUM(x.Amount), 0) / (1 + b.BusinessTaxRate), 2) Amount
                                FROM SCM.SoDetail x
                                WHERE x.SoId = a.SoId
                                GROUP BY x.SoId
                            ) c
                            OUTER APPLY(
                                SELECT ISNULL(SUM(x.Amount), 0) Amount
                                FROM SCM.SoDetail x
                                WHERE x.SoId = a.SoId
                                GROUP BY x.SoId
                            ) d
                            WHERE a.SoId = @SoId
                            GROUP BY b.Taxation, c.Amount, d.Amount, b.BusinessTaxRate";

                    dynamicParameters.Add("SoId", SoId);

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

        #region //GetSoSequence -- 取得訂單單身流水號資料 -- Zoey 2022.07.13
        public string GetSoSequence(int SoId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.SoDetailId";
                    sqlQuery.auxKey = 
                        @", ISNULL(MAX(a.SoSequence), 0) SoSequence";
                    sqlQuery.columns = "";
                    sqlQuery.mainTables =
                        @"FROM SCM.SoDetail a";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SoId", @" AND a.SoId = @SoId", SoId);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.groupBy = "GROUP BY a.SoDetailId";

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

        #region //GetSoModification -- 取得訂單變更單頭資料 -- Zoey 2022.08.16
        public string GetSoModification(int SomId, int SoId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.SomId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.SoId, a.SomDetail, a.Version, a.DepartmentId, a.DocDate, a.SoRemark, a.SalesmenId, a.CustomerAddressFirst
                          , a.CustomerAddressSecond, a.CustomerPurchaseOrder, a.DepositPartial, a.DepositRate, a.Currency, a.ExchangeRate, a.TaxNo
                          , a.Taxation, a.BusinessTaxRate, a.DetailMultiTax, a.ShipMethod, a.TradeTerm, a.PaymentTerm, a.PriceTerm, a.OriVersion
                          , a.OriDepartmentId, a.OriSoRemark, a.OriSalesmenId, a.OriCustomerAddressFirst, a.OriCustomerAddressSecond
                          , a.OriCustomerPurchaseOrder, a.OriDepositPartial, a.OriDepositRate, a.OriCurrency, a.OriExchangeRate, a.OriTaxNo
                          , a.OriTaxation, a.OriBusinessTaxRate, a.OriDetailMultiTax, a.OriShipMethod, a.OriTradeTerm, a.OriPaymentTerm, a.OriPriceTerm
                          , a.ClosureStatus, a.ModiReason, a.ConfirmStatus, a.ConfirmUserId, a.TransferStatus, a.TransferDate
                          , b.SoErpNo, b.SoErpPrefix, b.CustomerId, b.SoErpPrefix + '-' + b.SoErpNo SoErpFullNo
                          ,(
                              SELECT x.SomDetailId, x.SoSequence, x.MtlItemId, y.MtlItemNo, x.SoMtlItemName MtlItemName, x.SoMtlItemSpec MtlItemSpec, x.CustomerMtlItemNo
                              , x.InventoryId, x.UomId, x.SoQty, x.FreebieQty, x.UnitPrice, x.Amount, x.SoPriceQty, x.SoPriceUomId
                              , x.TaxNo, x.BusinessTaxRate, FORMAT(x.PromiseDate, 'yyyy-MM-dd') PromiseDate, x.Project, x.SoDetailRemark
                              FROM SCM.SomDetail x
                              LEFT JOIN PDM.MtlItem y ON y.MtlItemId = x.MtlItemId
                              WHERE x.SomId = a.SomId
                              FOR JSON PATH, ROOT('data')    
                          ) SomDetail";
                    sqlQuery.mainTables =
                        @"FROM SCM.SoModification a
                          LEFT JOIN SCM.SaleOrder b ON b.SoId = a.SoId";
                    string queryTable =
                        @"FROM (
                            SELECT a.SomId, a.SoId
                            ,(
                                SELECT x.SomDetailId, x.SoSequence, x.MtlItemId, y.MtlItemNo, x.SoMtlItemName MtlItemName, x.SoMtlItemSpec MtlItemSpec, x.CustomerMtlItemNo
                                , x.InventoryId, x.UomId, x.SoQty, x.FreebieQty, x.UnitPrice, x.Amount, x.SoPriceQty, x.SoPriceUomId
                                , x.TaxNo, x.BusinessTaxRate, FORMAT(x.PromiseDate, 'yyyy-MM-dd') PromiseDate, x.Project, x.SoDetailRemark
                                FROM SCM.SomDetail x
                                LEFT JOIN PDM.MtlItem y ON y.MtlItemId = x.MtlItemId
                                WHERE x.SomId = a.SomId
                                FOR JSON PATH, ROOT('data')
                            ) SomDetail
                            FROM SCM.SoModification a
                        ) a ";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SomId", @" AND a.SomId = @SomId", SomId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SoId", @" AND a.SoId = @SoId", SoId);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.SomId";
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

        #region //GetSoErpFullNo -- 取得訂單單別+單號 -- Zoey 2022.09.13
        public string GetSoErpFullNo(string DeliveryStatus)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT x.SoId, x.SoErpPrefix + '-' + x.SoErpNo SoErpFullNo
                            FROM SCM.SaleOrder x
                            WHERE x.SoId IN (
                                SELECT DISTINCT a.SoId
                                FROM SCM.SoDetail a
                                INNER JOIN SCM.SaleOrder b ON a.SoId = b.SoId
                                INNER JOIN PDM.MtlItem c ON a.MtlItemId = c.MtlItemId
                                WHERE 1=1
                                AND a.ConfirmStatus = @ConfirmStatus
                                AND a.ClosureStatus IN @ClosureStatus
                                AND b.ConfirmStatus = @ConfirmStatus
                                AND c.MtlItemNo LIKE @MtlItemNo + '%'
                            )
                            AND x.CompanyId = @CompanyId";
                    dynamicParameters.Add("ConfirmStatus", "Y");
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    switch (DeliveryStatus)
                    {
                        case "tracked":
                            dynamicParameters.Add("ClosureStatus", "N".Split(','));
                            dynamicParameters.Add("MtlItemNo", "");
                            break;
                        case "triTrade":
                            dynamicParameters.Add("ClosureStatus", "Y,y,N".Split(','));
                            dynamicParameters.Add("MtlItemNo", "5");
                            break;
                        default:
                            dynamicParameters.Add("ClosureStatus", "Y,y,N".Split(','));
                            dynamicParameters.Add("MtlItemNo", "");
                            break;
                    }
                    sql += @" ORDER BY x.DocDate DESC, x.SoErpPrefix + '-' + x.SoErpNo";

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

        #region //GetSaleOrderPdf -- 取得訂單PDF資料 -- Zoey 2022.09.13
        public string GetSaleOrderPdf(int SoId)
        {
            try
            {
                List<SaleOrderPdf> saleOrderPdfs = new List<SaleOrderPdf>();

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //判斷公司資料是否正確
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 CompanyId, CompanyName, ErpDb
                                FROM BAS.Company
                                WHERE CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0) throw new SystemException("【公司】資料錯誤!");

                    foreach (var item in result)
                    {
                        ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                    }
                    #endregion

                    sql = @"SELECT a.SoId, a.SoErpPrefix SoErpPrefixNo, a.SoErpNo, FORMAT(a.DocDate, 'yyyy/MM/dd') DocDate, d.DepartmentNo, d.DepartmentName
                            , e.UserNo SalesmenNo, e.UserName SalesmenName, c.CustomerNo, c.CustomerShortName CustomerName, c.TelNoFirst TelNo
                            , c.FaxNo, c.GuiNumber, a.PaymentTerm PaymentTermNo, a.CustomerPurchaseOrder, a.CustomerAddressFirst CustomerAddress
                            , a.Currency, a.ExchangeRate, Cast(Cast(a.BusinessTaxRate * 100 as decimal) as varchar) + '%' BusinessTaxRate, a.TaxNo, a.SoRemark
                            , a.TotalQty, a.Amount SoAmount, a.TaxAmount, a.Amount + a.TaxAmount TotalAmount, a.SoErpPrefix + '-' + a.SoErpNo SoFullNo
                            , (
                                   SELECT w.SoSequence, x.MtlItemNo, w.SoMtlItemName MtlItemName, x.MtlItemSpec, w.SoQty
                                   , y.UomNo, CAST(w.UnitPrice as nvarchar) UnitPrice, CAST(w.Amount as nvarchar) Amount
                                   , FORMAT(w.PromiseDate, 'yyyy/MM/dd') PromiseDate, z.InventoryNo, w.SoDetailRemark
                                   FROM SCM.SoDetail w
                                   LEFT JOIN PDM.MtlItem x ON x.MtlItemId = w.MtlItemId
                                   LEFT JOIN PDM.UnitOfMeasure y ON y.UomId = w.UomId
                                   LEFT JOIN SCM.Inventory z ON z.InventoryId = w.InventoryId
                                   WHERE w.SoId = a.SoId
                                   FOR JSON PATH, ROOT('data')
                               ) SoDetail
                            , a.Taxation, ISNULL(f.TypeName, '') TaxationName
                            FROM SCM.SaleOrder a
                            INNER JOIN SCM.Customer c ON c.CustomerId = a.CustomerId
                            INNER JOIN BAS.Department d ON d.DepartmentId = a.DepartmentId
                            INNER JOIN BAS.[User] e ON e.UserId = a.SalesmenId
                            LEFT JOIN BAS.[Type] f ON a.Taxation = f.TypeNo AND f.TypeSchema = 'Customer.Taxation'
                            WHERE a.SoId = @SoId";
                    dynamicParameters.Add("@SoId", SoId);

                    saleOrderPdfs = sqlConnection.Query<SaleOrderPdf>(sql, dynamicParameters).ToList();
                }

                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                {
                    List<Erp> erps = new List<Erp>();

                    #region //撈取單別資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT LTRIM(RTRIM(MQ001)) SoErpPrefixNo, LTRIM(RTRIM(MQ002)) SoErpPrefixName 
                            FROM CMSMQ
                            LEFT JOIN CMSMU ON MQ001 = MU001 
                            WHERE 1=1
                            AND (MQ003 ='22' OR (MQ003 = '27' AND MQ038 = '1'))
                            AND ((MU003 = 'DS' AND MQ029 ='Y') OR MQ029='N')
                            AND LTRIM(RTRIM(MQ001)) = @SoErpPrefix";
                    dynamicParameters.Add("@SoErpPrefix", saleOrderPdfs.Select(x => x.SoErpPrefixNo).FirstOrDefault());

                    erps = sqlConnection.Query<Erp>(sql, dynamicParameters).ToList();

                    saleOrderPdfs = saleOrderPdfs.GroupJoin(erps, x => x.SoErpPrefixNo, y => y.SoErpPrefixNo, (x, y) => { x.SoErpPrefixName = y.FirstOrDefault()?.SoErpPrefixName; return x; }).ToList();
                    #endregion

                    #region //撈取出貨廠別
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT LTRIM(RTRIM(a.TC001)) SoErpPrefixNo,LTRIM(RTRIM(a.TC002)) SoErpNo
                            , LTRIM(RTRIM(a.TC007)) ShippingSiteNo, LTRIM(RTRIM(b.MB002)) ShippingSiteName
                            FROM COPTC a
                            INNER JOIN CMSMB b ON b.MB001 = a.TC007
                            WHERE 1=1
                            AND a.TC001 = @SoErpPrefixNo
                            AND a.TC002 = @SoErpNo";
                    dynamicParameters.Add("@SoErpPrefixNo", saleOrderPdfs.Select(x => x.SoErpPrefixNo).FirstOrDefault());
                    dynamicParameters.Add("@SoErpNo", saleOrderPdfs.Select(x => x.SoErpNo).FirstOrDefault());

                    erps = sqlConnection.Query<Erp>(sql, dynamicParameters).ToList();

                    saleOrderPdfs = saleOrderPdfs.GroupJoin(erps, x => (x.SoErpPrefixNo, x.SoErpNo), y => (y.SoErpPrefixNo, y.SoErpNo), (x, y) => { x.ShippingSiteNo = y.FirstOrDefault()?.ShippingSiteNo; x.ShippingSiteName = y.FirstOrDefault()?.ShippingSiteName; return x; }).ToList();
                    #endregion

                    #region //撈取付款條件
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT LTRIM(RTRIM(NA002)) PaymentTermNo, LTRIM(RTRIM(NA003)) PaymentTermName 
                            FROM CMSNA
                            WHERE 1=1
                            AND LTRIM(RTRIM(NA002)) = @PaymentTerm";
                    dynamicParameters.Add("@PaymentTerm", saleOrderPdfs.Select(x => x.PaymentTermNo).FirstOrDefault());

                    erps = sqlConnection.Query<Erp>(sql, dynamicParameters).ToList();

                    saleOrderPdfs = saleOrderPdfs.GroupJoin(erps, x => x.PaymentTermNo, y => y.PaymentTermNo, (x, y) => { x.PaymentTermName = y.FirstOrDefault()?.PaymentTermName; return x; }).ToList();
                    #endregion

                    #region //撈取課稅別
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT LTRIM(RTRIM(NN001)) TaxNo, LTRIM(RTRIM(NN002)) TaxName
                            FROM CMSNN 
                            WHERE 1=1
                            AND LTRIM(RTRIM(NN001)) = @TaxNo";
                    dynamicParameters.Add("@TaxNo", saleOrderPdfs.Select(x => x.TaxNo).FirstOrDefault());

                    erps = sqlConnection.Query<Erp>(sql, dynamicParameters).ToList();

                    saleOrderPdfs = saleOrderPdfs.GroupJoin(erps, x => x.TaxNo, y => y.TaxNo, (x, y) => { x.TaxName = y.FirstOrDefault()?.TaxName; return x; }).ToList();
                    #endregion
                }

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    data = saleOrderPdfs
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

        #region //GetSaleOrderDoc -- 取得訂單單據Doc -- Ellie 2023.06.09
        public string GetSaleOrderDoc(int SoId)
        {
            try
            {
                string companyNo = "", erpPrefix = "", erpNo = "", userName = "";                

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //公司別資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.ErpNo, a.ErpDb
                            FROM BAS.Company a
                            WHERE CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    var resultCompany = sqlConnection.Query(sql, dynamicParameters);
                    if (resultCompany.Count() <= 0) throw new SystemException("公司別資料錯誤!");

                    foreach (var item in resultCompany)
                    {
                        ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        companyNo = item.ErpNo;
                    }
                    #endregion

                    #region //訂單單據資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.SoErpPrefix, a.SoErpNo, ISNULL(b.UserName, '未確認') UserName
                            FROM SCM.SaleOrder a
                            LEFT JOIN BAS.[User] b ON a.ConfirmUserId = b.UserId
                            WHERE SoId = @SoId";
                    dynamicParameters.Add("SoId", SoId);

                    var resultSo = sqlConnection.Query(sql, dynamicParameters);
                    if (resultSo.Count() <= 0) throw new SystemException("訂單單據資料錯誤!");

                    foreach (var item in resultSo)
                    {
                        erpPrefix = item.SoErpPrefix;
                        erpNo = item.SoErpNo;
                        userName = item.UserName;
                    }
                    #endregion
                }

                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                {
                    #region //COPTC資料 
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT b.MQ002
                            , a.TC001 + ' ' + b.MQ002 TC001Order
                            , a.TC001
                            , a.TC002
                            , FORMAT(CAST(LTRIM(RTRIM(a.TC039)) as date), 'yyyy/MM/dd') TC039Date
                            , LTRIM(RTRIM(a.TC007)) + ' ' + c.MB002 TC007Fac, a.TC005 + ' ' + d.ME002 TC005Dep
                            , a.TC006 + ' ' + e.MV002 TC006Clerk, a.TC004 + ' ' + f.MA002 TC004Clt
                            , f.MA006, f.MA008, f.MA010, a.TC014, a.TC012, a.TC010
                            , a.TC008, CONVERT(real, a.TC009) TC009Rate  
                            , FORMAT(a.TC041, 'P0') TC041
                            , CASE a.TC016
                                WHEN 1 THEN '應稅內含'
                                WHEN 2 THEN '應稅外加'
                                WHEN 3 THEN '零稅率'
                                WHEN 4 THEN '免稅'
                                WHEN 9 THEN '不計稅'
                            END TC016
                            , a.TC015
                            , FORMAT(a.TC031, '#,#.00') TC031Qua
                            , FORMAT(a.TC029, '#,#.00') TC029Amount
                            , FORMAT(a.TC030, '#,#') TC030Tax
                            , FORMAT(a.TC029 + a.TC030, '#,#.00') TotalAmount
                            , a.TC039
                            , a.TC070
                            FROM COPTC a
                            INNER JOIN CMSMQ b ON b.MQ001 = a.TC001
                            INNER JOIN CMSMB c ON c.MB001 = a.TC007
                            INNER JOIN CMSME d ON d.ME001 = a.TC005
                            INNER JOIN (
                                SELECT a.MJ001, a.MJ003, b.MK002, c.MV002
                                FROM CMSMJ a
                                    INNER JOIN CMSMK b ON a.MJ001 = b.MK001
                                    INNER JOIN CMSMV c ON b.MK002 = c.MV001
                                WHERE MJ002 = '3'        
                            ) e ON e.MK002=a.TC006
                            INNER JOIN COPMA f ON f.MA001 = a.TC004
                            WHERE a.TC001 = @ErpPrefix
                            AND a.TC002 = @ErpNo";
                    dynamicParameters.Add("ErpPrefix", erpPrefix);
                    dynamicParameters.Add("ErpNo", erpNo);
                    
                    var resultCoptc = sqlConnection.Query(sql, dynamicParameters);
                    if (resultCoptc.Count() <= 0) throw new SystemException("ERP客戶訂單資料錯誤!");
                    #endregion

                    #region //COPTD資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.TD003 Seq, a.TD004 MtlItemNo, a.TD005 MtlItemName, a.TD006 MtlItemSpec, a.TD014 CutMtlItemName
                            , CONVERT(real, a.TD008) Qty, CONVERT(real, a.TD050) PQty, CONVERT(real, a.TD024) GQty
                            , a.TD010 Uom
                            --, FORMAT(a.TD011, '#,#.00') UnPri
                            , a.TD011 UnPri
                            , FORMAT(a.TD012, '#,#.00') Amount
                            , FORMAT(CAST(LTRIM(RTRIM(a.TD013)) as date), 'yyyy/MM/dd') PrDate
                            , a.TD007 InvNo, a.TD020 Remark
                            FROM COPTD a
                            WHERE a.TD001 = @ErpPrefix
                            AND a.TD002 = @ErpNo
                            ORDER BY a.TD003";
                    dynamicParameters.Add("ErpPrefix", erpPrefix);
                    dynamicParameters.Add("ErpNo", erpNo);

                    var resultCoptd = sqlConnection.Query(sql, dynamicParameters);
                    if (resultCoptd.Count() <= 0) throw new SystemException("ERP客戶訂單單單身資料錯誤!");
                    #endregion

                    #region //COPUC資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT UC003 Seq, CAST(UC004 * 100 AS float) Ratio, UC005 Amount
                            , CASE WHEN LEN(LTRIM(RTRIM(UC006))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(UC006)) AS date), 'yyyy/MM/dd') ELSE '' END PaymentDate
                            , UC007 ClosingStatus
                            FROM COPUC
                            WHERE UC001 = @ErpPrefix
                            AND UC002 = @ErpNo
                            ORDER BY UC003";
                    dynamicParameters.Add("ErpPrefix", erpPrefix);
                    dynamicParameters.Add("ErpNo", erpNo);

                    var resultCopuc = sqlConnection.Query(sql, dynamicParameters);
                    //if (resultCopuc.Count() <= 0) throw new SystemException("ERP訂金資料錯誤!");
                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = resultCoptc,
                        dataDetail = resultCoptd,
                        dataDeposit = resultCopuc,
                        user = userName
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

        #region //GetSoDeposit -- 取得訂單訂金資料 -- Ben Ma 2023.06.27
        public string GetSoDeposit(int SoId)
        {
            try
            {
                string soErpPrefix = "", soErpNo = "", transferStatus = "";

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //判斷公司資料是否正確
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 CompanyId, CompanyName, ErpDb
                            FROM BAS.Company
                            WHERE CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0) throw new SystemException("【公司】資料錯誤!");

                    foreach (var item in result)
                    {
                        ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                    }
                    #endregion

                    #region //判斷訂單資料是否正確
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 SoErpPrefix, SoErpNo, ConfirmStatus, TransferStatus
                            FROM SCM.SaleOrder
                            WHERE SoId = @SoId
                            AND CompanyId = @CompanyId";
                    dynamicParameters.Add("SoId", SoId);
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    var resultSo = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0) throw new SystemException("訂單資料錯誤!");

                    foreach (var item in resultSo)
                    {
                        soErpPrefix = item.SoErpPrefix;
                        soErpNo = item.SoErpNo;
                        transferStatus = item.TransferStatus;
                    }

                    #region //判斷訂單狀態
                    if (transferStatus != "Y") throw new SystemException("訂單尚未拋轉，無法取得訂金資料!");
                    #endregion
                    #endregion
                }

                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT UC003 'Sequence', CAST(UC004 * 100 AS float) Ratio, UC005 Amount
                            , CASE WHEN LEN(LTRIM(RTRIM(UC006))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(UC006)) AS date), 'yyyy-MM-dd') ELSE '' END PaymentDate
                            , UC007 ClosingStatus
                            FROM COPUC
                            WHERE UC001 = @ErpPrefix
                            AND UC002 = @ErpNo
                            ORDER BY UC003";
                    dynamicParameters.Add("ErpPrefix", soErpPrefix);
                    dynamicParameters.Add("ErpNo", soErpNo);

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

        #region //GetSoDetailTxAmount -- 試算訂單單身金額與稅額 -- Ann 2023-12-26
        public string GetSoDetailTxAmount(int SoId, double UnitPrice, double SoPriceQty)
        {
            try
            {
                string currency = "";
                string taxNo = "";
                int unitRound = 0;
                int totalRound = 0;
                double amount = 0;
                double exciseTax = 0;
                string taxation = "";
                double taxAmount = 0;
                string customerNo = "";

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
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

                    #region //判斷訂單資料是否正確
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 a.Currency, a.TaxNo
                            , b.CustomerNo
                            FROM SCM.SaleOrder a 
                            INNER JOIN SCM.Customer b ON a.CustomerId = b.CustomerId
                            WHERE a.SoId = @SoId
                            AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("SoId", SoId);
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0) throw new SystemException("訂單資料錯誤!");

                    foreach (var item in result)
                    {
                        currency = item.Currency;
                        taxNo = item.TaxNo;
                        customerNo = item.CustomerNo;
                    }
                    #endregion
                }

                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                {
                    #region //交易幣別設定
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT MF003, MF004, MF005, MF006
                            FROM CMSMF
                            WHERE MF001 = @Currency";
                    dynamicParameters.Add("Currency", currency);

                    var resultCurrencySetting = sqlConnection.Query(sql, dynamicParameters);
                    if (resultCurrencySetting.Count() <= 0) throw new SystemException("ERP交易幣別設定不存在!");

                    foreach (var item in resultCurrencySetting)
                    {
                        unitRound = Convert.ToInt32(item.MF003); //單價取位
                        totalRound = Convert.ToInt32(item.MF004); //金額取位
                    }
                    #endregion

                    #region //稅別碼設定
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT RTRIM(LTRIM(NN001)) NN001, RTRIM(LTRIM(NN002)) NN002, NN004, RTRIM(LTRIM(NN006)) NN006
                            FROM CMSNN
                            WHERE NN001 = @TaxNo";
                    dynamicParameters.Add("TaxNo", taxNo);

                    var resultTax = sqlConnection.Query(sql, dynamicParameters);
                    if (resultTax.Count() <= 0) throw new SystemException("ERP稅別碼設定不存在!");

                    foreach (var item in resultTax)
                    {
                        exciseTax = Convert.ToDouble(item.NN004); //營業稅率
                        taxation = item.NN006; //課稅別
                    }
                    #endregion

                    #region //計算訂單單身金額
                    amount += Math.Round(Math.Round(Convert.ToDouble(UnitPrice), unitRound, MidpointRounding.AwayFromZero) * Convert.ToDouble(SoPriceQty), totalRound, MidpointRounding.AwayFromZero);
                    #endregion

                    #region //計算稅額
                    #region //計算數量與金額
                    switch (taxation)
                    {
                        case "1":
                            taxAmount = amount - Math.Round(amount / (1 + exciseTax), totalRound, MidpointRounding.AwayFromZero);
                            amount = Math.Round(amount / (1 + exciseTax), totalRound);
                            break;
                        case "2":
                            taxAmount = Math.Round(amount * exciseTax, totalRound, MidpointRounding.AwayFromZero);
                            break;
                        case "3":
                            break;
                        case "4":
                            break;
                        case "9":
                            break;
                    }
                    #endregion
                    #endregion

                    JObject data = JObject.FromObject(new
                    {
                        TaxAmount = taxAmount,
                        Amount = amount,
                        CustomerNo = customerNo,
                        Currency = currency
                    });

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data
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

        #region//GetPendingSaleOrder --未結訂單
        public string GetPendingSaleOrder(string Company, string OrderDate, string CustomerNo, string CustomerShortName
            , string CustomerOrderNo, string Currency, string ExchangeRate, string SalesmenName
            , string MtlItemNo, string MtlItemName, string PromiseDate
            , string OrderBy = "", int PageIndex = -1, int PageSize = -1)
        {
            try
            {
                int CompanyId = -1;
                string ErpConnectionStrings = "", ErpNo = "";
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {

                    #region //判斷公司資料是否正確
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 CompanyId, CompanyName, ErpDb,ErpNo
                                FROM BAS.Company
                                WHERE CompanyNo = @CompanyNo";
                    dynamicParameters.Add("CompanyNo", Company);

                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() != 1) throw new SystemException("【公司】資料錯誤!");

                    foreach (var item in result)
                    {
                        CompanyId = Convert.ToInt32(item.CompanyId);
                        ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        ErpNo = item.ErpNo;
                    }
                    #endregion
                }

                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                {                    
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.TC003,a.TC004,c.MA002,a.TC012,a.TC008,a.TC009,d.MF002,a.TC029
                            ,b.TD004,b.TD005,b.TD008,b.TD009,b.TD011,b.TD013,b.TD016
                            FROM COPTC a
                            INNER JOIN COPTD b ON a.TC001=b.TD001 AND a.TC002=b.TD002
                            INNER JOIN COPMA c ON a.TC004=c.MA001
                            INNER JOIN ADMMF d ON a.TC006=d.MF001
                            WHERE 1=1 AND b.TD016='N' AND b.TD016!='y' AND a.TC027!='V'  ";                   
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

        #region //GetSoDetailTemp --取得採購單匯入資料 -- Chia Yuan 2024.04.18
        public string GetSoDetailTemp(int SoDetailTempId, int SoId, int SoDetailId, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.SoDetailTempId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.SoId, ISNULL(a.SoDetailId, '') AS SoDetailId, a.BaseDocType
                        , a.BaseErpPrefix AS TD001, a.BaseErpNo AS TD002, a.BaseErpSNo AS TD003
                        , a.BaseErpPrefix + '-' + a.BaseErpNo AS PoErpPrefixNo
                        , a.UomNo AS TD009, a.SoQty AS TD008, a.UnitPrice AS TD010, a.Amount AS TD011
                        , a.BaseMtlItemNo AS TD004, a.BaseMtlItemName AS TD005, a.BaseMtlItemSpec AS TD006, a.SortNumber
                        , a.[Status] AS SoDetailTempStatus, s.StatusName AS SoDetailTempStatusName
                        , ISNULL(CONVERT(VARCHAR(10), b.CustomerMtlItemId), '') AS CustomerMtlItemId
                        , ISNULL(b.CustomerMtlItemNo, '') AS CustomerMtlItemNo
                        , ISNULL(b.CustomerMtlItemName, '') AS CustomerMtlItemName
                        , ISNULL(CONVERT(VARCHAR(10), b.MtlItemId), '') AS MtlItemId
                        , ISNULL(b.MtlItemNo, '') AS MtlItemNo
                        , ISNULL(b.MtlItemName, '') AS MtlItemName
                        , ISNULL(b.SaleUomNo, '') AS SaleUomNo
                        , ISNULL(b.InventoryId, -1) AS InventoryId
                        , ISNULL(b.InventoryNo, '') AS InventoryNo
                        , ISNULL(b.InventoryName, '') AS InventoryName
                        , c.CompanyId, c.CompanyNo, c.CompanyName
                        , ISNULL(k.ProductType, '') AS ProductType, FORMAT(k.PromiseDate, 'yyyy-MM-dd') AS PromiseDate, ISNULL(k.Project, '') AS Project
                        , k.SoDetailRemark, k.FreebieOrSpareQty, k.ProductTypeName
                        , k.MainKey, k.cboText";
                    sqlQuery.mainTables =
                        @"FROM SCM.SoDetailTemp a
                        INNER JOIN BAS.[Status] s ON s.StatusNo = a.[Status] AND s.StatusSchema = 'Boolean'
                        INNER JOIN BAS.Company c ON c.CompanyId = a.CompanyId
                        OUTER APPLY (
                            SELECT ka.BaseErpPrefix + ka.BaseErpNo + ka.BaseErpSNo AS MainKey
                            , ka.BaseErpPrefix + '-' + ka.BaseErpNo + '-' + ka.BaseErpSNo + CASE WHEN kb.SoSequence IS NULL THEN '' ELSE ' 訂單單身流水號(' + kb.SoSequence + ')' END AS cboText
                            , kb.ProductType, kb.PromiseDate, kb.Project, kb.SoDetailRemark
                            , kb.FreebieQty, kb.FreebieSiQty, kb.SpareQty, kb.SpareSiQty
                            , CASE kb.ProductType WHEN '1' THEN kb.FreebieQty ELSE kb.SpareQty END AS FreebieOrSpareQty
                            , kt.TypeName AS ProductTypeName
                            --, '採購單品號：' + a.BaseMtlItemNo + '(' + CASE ISNULL(b.CustomerNo, '') WHEN '' THEN '查無客戶部番' ELSE b.CustomerNo END + ')'
                            FROM SCM.SoDetailTemp ka
                            LEFT JOIN SCM.SoDetail kb ON kb.SoDetailId = ka.SoDetailId
                            LEFT JOIN BAS.[Type] kt ON kt.TypeNo = kb.ProductType AND kt.TypeSchema = 'SoDetail.ProductType'
                            WHERE ka.SoDetailTempId = a.SoDetailTempId
                            --ka.BaseErpPrefix = a.BaseErpPrefix AND ka.BaseErpNo = a.BaseErpNo AND ka.BaseErpSNo = a.BaseErpSNo AND a.SoId = @SoId
                        ) k
                        LEFT JOIN (
	                        SELECT a.CustomerMtlItemId, a.CustomerMtlItemNo, a.CustomerMtlItemName, a.[Status] AS CustMtlItemStatus
	                        , a.TransferStatus AS CustMtlItemTransferStatus, s1.StatusName AS CustMtlItemTransferStatusName
	                        , b.MtlItemId, b.MtlItemNo, b.MtlItemName, b.MtlItemDesc, b.[Status] AS MtlItemStatus
                            , c.CustomerId, c.CustomerName, c.CustomerNo
	                        , b.TransferStatus AS MtlItemTransferStatus, s2.StatusName AS MtlItemTransferStatusName
                            , ISNULL(u.UomNo, '') AS SaleUomNo, ISNULL(u.UomName, '') AS SaleUomName
                            , i.InventoryId, i.InventoryNo, i.InventoryName
	                        , ROW_NUMBER() OVER (PARTITION BY a.CustomerMtlItemNo ORDER BY a.CustomerMtlItemId DESC) AS SortCustMtlItem
	                        FROM PDM.CustomerMtlItem a 
	                        INNER JOIN PDM.MtlItem b ON b.MtlItemId = a.MtlItemId AND b.[Status] = a.[Status]
                            INNER JOIN SCM.Customer c ON c.CustomerId = a.CustomerId
                            INNER JOIN SCM.Inventory i ON i.InventoryId = b.InventoryId
                            LEFT JOIN PDM.UnitOfMeasure u ON u.UomId = b.SaleUomId
	                        INNER JOIN BAS.[Status] s1 ON s1.StatusNo = a.TransferStatus AND s1.StatusSchema = 'Boolean'
	                        INNER JOIN BAS.[Status] s2 ON s2.StatusNo = b.TransferStatus AND s2.StatusSchema = 'Boolean'
                        ) b ON b.CustomerMtlItemId = a.CustomerMtlItemId";
                    string queryTable = "";

                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @" AND a.SoId = @SoId";
                    dynamicParameters.Add("SoId", SoId);

                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SoDetailTempId", @" AND a.SoDetailTempId = @SoDetailTempId ", SoDetailTempId);

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

        #region//GetQuotationErp --取得ERP報價單 -- Shintokuro 2024.11.28
        public string GetQuotationErp(int SoId,int MtlItemId)
        {
            try
            {
                int CompanyId = -1;
                string ErpConnectionStrings = "", ErpNo = "";
                string MtlItemNo = "", Currency = "", PaymentTerm = "", TaxNo = "", DocDate = "", CustomerNo = "";
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //判斷公司資料是否正確
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 CompanyId, CompanyName, ErpDb,ErpNo
                                FROM BAS.Company
                                WHERE CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() != 1) throw new SystemException("【公司】資料錯誤!");

                    foreach (var item in result)
                    {
                        CompanyId = Convert.ToInt32(item.CompanyId);
                        ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        ErpNo = item.ErpNo;
                    }
                    #endregion

                    #region //判斷訂單資料是否正確
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 a.Currency,a.PaymentTerm,a.TaxNo,CONVERT(VARCHAR(8), a.DocDate, 112) AS DocDate,b.CustomerNo
                            FROM SCM.SaleOrder a
                            INNER JOIN SCM.Customer b on a.CustomerId = b.CustomerId
                            WHERE SoId = @SoId";
                    dynamicParameters.Add("SoId", SoId);

                    result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() > 0)
                    {
                        foreach (var item in result)
                        {
                            Currency = item.Currency;
                            PaymentTerm = item.PaymentTerm;
                            TaxNo = item.TaxNo;
                            DocDate = item.DocDate;
                            CustomerNo = item.CustomerNo;
                        }
                    } 
                    #endregion

                    #region //判斷品號資料是否正確
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 MtlItemNo
                            FROM PDM.MtlItem
                            WHERE MtlItemId = @MtlItemId";
                    dynamicParameters.Add("MtlItemId", MtlItemId);

                    result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() > 0)
                    {
                        foreach (var item in result)
                        {
                            MtlItemNo = item.MtlItemNo;
                        }
                    }
                    #endregion
                }

                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    //目前只單純做報價單綁定訂單,先不卡控太多條件
                    sql = @"SELECT a.TA001 +'-' + a.TA002 + '-' +b.TB003 QuotationFullNoErp
                            FROM COPTA a
                            INNER JOIN COPTB b on a.TA001 = b.TB001 AND a.TA002 = b.TB002
                            WHERE 1=1
                            AND a.TA016 = 'Y' 
                            AND a.TA019 = 'Y' 
                            --AND a.TA007 = @Currency
                            --AND a.TA026 = @PaymentTerm
                            --AND a.TA049 = @TaxNo
                            AND (b.TB016 = '' OR b.TB016 <= @DocDate)
                            AND (b.TB017 = '' OR b.TB017 > @DocDate)
                            AND a.TA004 = @CustomerNo
                            AND b.TB004 = @MtlItemNo
                            ";
                    dynamicParameters.Add("Currency", Currency);
                    dynamicParameters.Add("PaymentTerm", PaymentTerm);
                    dynamicParameters.Add("TaxNo", TaxNo);
                    dynamicParameters.Add("DocDate", DocDate);
                    dynamicParameters.Add("CustomerNo", CustomerNo);
                    dynamicParameters.Add("MtlItemNo", MtlItemNo);

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

        #region //GetPMDOrderData -- 取得 PMD 訂單資料 GPAI -- 2025-05-06
        public string GetPMDOrderData(int TypeId, int SendCustomer ,string Company)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    int companyId = -1;
                    if (Company == "")
                    {
                        #region //確認公司別DB
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb, a.CompanyId
                            FROM BAS.Company a
                            WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var companyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (companyResult.Count() <= 0) throw new SystemException("公司別錯誤!");

                        foreach (var item in companyResult)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            companyId = item.CompanyId;

                        }
                        #endregion
                    }
                    else
                    {
                        #region //確認公司別DB
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb, a.CompanyId
                            FROM BAS.Company a
                            WHERE CompanyNo = @Company";
                        dynamicParameters.Add("Company", Company);

                        var companyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (companyResult.Count() <= 0) throw new SystemException("公司別錯誤!");

                        foreach (var item in companyResult)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            companyId = item.CompanyId;
                        }
                        #endregion
                    }


                    var CustomerId = -1;
                    if (SendCustomer > 0) {
                        dynamicParameters = new DynamicParameters();
                        sql = @"select * 
                            from SCM.Customer
                            where CustomerId = @CustomerId";
                        dynamicParameters.Add("CustomerId", SendCustomer);
                        var result1 = sqlConnection.Query(sql, dynamicParameters);
                        foreach (var iten in result1) {
                            CustomerId = iten.CustomerId;
                        }
                    }
                    
                    

                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.PMDOrderTempId, a.MtlItemId, a.SoQty, a.SpareQty, a.OrderDate, a.DrawingAttachmentDate, a.CurrentProcess, a.CustomerDueDate, a.ConfirmedDueDate
                            , a.DelayRemark, a.MaterialNumber, a.Remark, a.TrackingNumber, a.ShipFrom, a.SemiDelivery, a.TypeId, a.PMDItem01, a.PMDItem02, a.PMDItem03, a.PMDItem04, a.PlannedShipmentDate
                            , a.ActualShipmentDate, a.PlannedShipmentQty, a.ActualShipmentQty, a.MoldFrameCtrlNo, a.ReturnQty, a.ReturnIssue, a.PrevReturnQty, a.PrevShortageQty, a.TotalReplenishmentQty
                            , a.CAV, a.ReturnDate, a.SoDetailId
                            , a.CreateDate, a.LastModifiedDate, a.CreateBy, a.LastModifiedBy, b.MtlItemNo, a.MtlItemName, a.CustomerMtlItemNo, c.CustomerNo, c.CustomerName, c.CustomerId
                            FROM SCM.PMDOrderTemp a 
                            LEFT JOIN PDM.MtlItem b on a.MtlItemId = b.MtlItemId 
							INNER JOIN SCM.Customer c ON a.CustomerId = c.CustomerId
                            WHERE 1=1 and c.CompanyId =  @CompanyId and a.SendType = 0
                            ";
                    dynamicParameters.Add("CompanyId", companyId);

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TypeId", " AND a.TypeId = @TypeId", TypeId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CustomerId", " AND a.CustomerId = @CustomerId", CustomerId);

                    sql += " ORDER BY PMDOrderTempId DESC";

                    var result = sqlConnection.Query(sql, dynamicParameters);

                    #region 暫留
                    //#region //新增資料
                    //#region //初始化Data
                    //List<Data> datas = new List<Data>();
                    //Data data = new Data();
                    //#region //設定單身量測標準
                    //int row = 2;
                    //foreach (var item2 in result)
                    //{
                    //    #region //item2 Data
                    //    var PMDOrderId =  (item2.PMDOrderId ?? 0) == 0 ? "" : item2.PMDOrderId?.ToString();

                    //    var MtlItemId = (item2.MtlItemId ?? 0) == 0 ? "" : item2.MtlItemId?.ToString();
                    //    var SoQty = (item2.SoQty ?? 0) == 0 ? "" : item2.SoQty?.ToString();
                    //    var SpareQty = (item2.SpareQty ?? 0) == 0 ? "" : item2.SpareQty?.ToString();
                    //    var OrderDate = item2.OrderDate ?? "";
                    //    var DrawingAttachmentDate = item2.DrawingAttachmentDate ?? "";
                    //    var CurrentProcess = item2.CurrentProcess ?? "";
                    //    var CustomerDueDate = item2.CustomerDueDate ?? "";
                    //    var ConfirmedDueDate = item2.ConfirmedDueDate ?? "";
                    //    var DelayRemark = item2.DelayRemark ?? "";
                    //    var MaterialNumber = item2.MaterialNumber ?? "";
                    //    var Remark = item2.Remark ?? "";
                    //    var TrackingNumber = item2.TrackingNumber ?? "";
                    //    var ShipFrom = item2.ShipFrom ?? "";
                    //    var SemiDelivery = item2.SemiDelivery ?? "";
                    //    //var TypeId = item2.TypeId == 0 ? "" : item2.TypeId.ToString();
                    //    var PMDItem01 = item2.PMDItem01 ?? "";
                    //    var PMDItem02 = item2.PMDItem02 ?? "";
                    //    var PMDItem03 = item2.PMDItem03 ?? "";
                    //    var PMDItem04 = item2.PMDItem04 ?? "";
                    //    var PlannedShipmentDate = item2.PlannedShipmentDate ?? "";
                    //    var ActualShipmentDate = item2.ActualShipmentDate ?? "";
                    //    var PlannedShipmentQty = (item2.PlannedShipmentQty ?? 0) == 0 ? "" : item2.PlannedShipmentQty?.ToString();
                    //    var ActualShipmentQty = (item2.ActualShipmentQty ?? 0) == 0 ? "" : item2.ActualShipmentQty?.ToString();
                    //    var MoldFrameCtrlNo = item2.MoldFrameCtrlNo ?? "";
                    //    var ReturnQty = (item2.ReturnQty ?? 0) == 0 ? "" : item2.ReturnQty?.ToString();
                    //    var ReturnIssue = item2.ReturnIssue ?? "";
                    //    var PrevReturnQty = (item2.PrevReturnQty ?? 0) == 0 ? "" : item2.PrevReturnQty?.ToString();
                    //    var PrevShortageQty = (item2.PrevShortageQty ?? 0) == 0 ? "" : item2.PrevShortageQty?.ToString();
                    //    var TotalReplenishmentQty = (item2.TotalReplenishmentQty ?? 0) == 0 ? "" : item2.TotalReplenishmentQty?.ToString();
                    //    var MtlItemNo = "";
                    //    var MtlItemName =  "";
                    //    var CustomerMtlItemNo = item2.CustomerMtlItemNo ?? "";
                    //    var CAV = item2.CAV ?? "";
                    //    var ReturnDate = item2.ReturnDate ?? "";

                    //    #region //找品號最新過站資料
                    //    dynamicParameters = new DynamicParameters();
                    //    sql = @"select MtlItemNo, MtlItemName 
                    //            from PDM.MtlItem
                    //            where MtlItemId = @MtlItemId
                    //            ";
                    //    dynamicParameters.Add("MtlItemId", MtlItemId);

                    //    var MtlIresult = sqlConnection.Query(sql, dynamicParameters);
                    //    // CurrentProcess = Processresult.Any() ? Processresult.First() ?? "" : "";
                    //    foreach (var item3 in MtlIresult)
                    //    {
                    //         MtlItemNo = item3.MtlItemNo ?? "";
                    //         MtlItemName = item3.MtlItemName ?? "";
                    //    }
                    //        #endregion


                    //        #endregion

                    //        #region //找品號最新過站資料
                    //    dynamicParameters = new DynamicParameters();
                    //    sql = @"select top 1 b.ProcessAlias 
                    //            from MES.BarcodeProcess a
                    //            inner join MES.MoProcess b on a.MoProcessId = b.MoProcessId
                    //            inner join MES.ManufactureOrder c on a.MoId = c.MoId
                    //            inner join MES.WipOrder d on c.WoId = d.WoId
                    //            where d.MtlItemId = @MtlItemId
                    //            ORDER BY BarcodeProcessId DESC";
                    //    dynamicParameters.Add("MtlItemId", MtlItemId);

                    //    var Processresult = sqlConnection.Query(sql, dynamicParameters);
                    //   // CurrentProcess = Processresult.Any() ? Processresult.First() ?? "" : "";

                    //    #endregion

                    //    switch (item2.TypeId)
                    //    {
                    //        case 1:
                    //            #region //初始化Data欄位

                    //            data = new Data()
                    //            {
                    //                cell = "A1",
                    //                css = "imported_class1",
                    //                format = "text",
                    //                value = "序號",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "B1",
                    //                css = "imported_class2",
                    //                format = "text",
                    //                value = "品號",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "C1",
                    //                css = "imported_class2",
                    //                format = "text",
                    //                value = "模號/品號",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "D1",
                    //                css = "imported_class2",
                    //                format = "text",
                    //                value = "品名",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "E1",
                    //                css = "imported_class2",
                    //                format = "text",
                    //                value = "訂單數",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "F1",
                    //                css = "imported_class3",
                    //                format = "text",
                    //                value = "備品數",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "G1",
                    //                css = "imported_class4",
                    //                format = "text",
                    //                value = "接單日",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "H1",
                    //                css = "imported_class5",
                    //                format = "text",
                    //                value = "下圖日",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "I1",
                    //                css = "imported_class6",
                    //                format = "text",
                    //                value = "當前進度",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "J1",
                    //                css = "imported_class7",
                    //                format = "text",
                    //                value = "客戶交期",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "K1",
                    //                css = "imported_class8",
                    //                format = "text",
                    //                value = "回復交期",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "L1",
                    //                css = "imported_class8",
                    //                format = "text",
                    //                value = "延誤交期時間",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "M1",
                    //                css = "imported_class9",
                    //                format = "text",
                    //                value = "物料編碼",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "N1",
                    //                css = "imported_class10",
                    //                format = "text",
                    //                value = "備註",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "O1",
                    //                css = "imported_class10",
                    //                format = "text",
                    //                value = "物流單號",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "P1",
                    //                css = "imported_class10",
                    //                format = "text",
                    //                value = "發貨地",
                    //            };
                    //            datas.Add(data); data = new Data()
                    //            {
                    //                cell = "Q1",
                    //                css = "imported_class10",
                    //                format = "text",
                    //                value = "半成品發貨(ZY-JC)",
                    //            };
                    //            datas.Add(data); data = new Data()
                    //            {
                    //                cell = "R1",
                    //                css = "imported_class10",
                    //                format = "text",
                    //                value = "保留欄位1",
                    //            };
                    //            datas.Add(data); data = new Data()
                    //            {
                    //                cell = "S1",
                    //                css = "imported_class10",
                    //                format = "text",
                    //                value = "保留欄位2",
                    //            };
                    //            datas.Add(data); data = new Data()
                    //            {
                    //                cell = "T1",
                    //                css = "imported_class10",
                    //                format = "text",
                    //                value = "保留欄位3",
                    //            };
                    //            datas.Add(data); data = new Data()
                    //            {
                    //                cell = "U1",
                    //                css = "imported_class10",
                    //                format = "text",
                    //                value = "保留欄位4",
                    //            };



                    //            #endregion

                    //            #region //資料
                    //            #region //設定

                    //            data = new Data()
                    //            {
                    //                cell = "A" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = PMDOrderId,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "B" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = MtlItemNo,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "C" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = CustomerMtlItemNo,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "D" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = MtlItemName,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "E" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = SoQty,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "F" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = SpareQty,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "G" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = OrderDate,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "H" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = DrawingAttachmentDate,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "I" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = CurrentProcess,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "J" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = CustomerDueDate,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "K" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = ConfirmedDueDate,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "L" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = DelayRemark,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "M" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = MaterialNumber,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "N" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = Remark,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "O" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = TrackingNumber,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "P" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = ShipFrom,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "Q" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = SemiDelivery,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "R" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = PMDItem01,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "S" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = PMDItem02,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "T" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = PMDItem03,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "U" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = PMDItem04,
                    //            };
                    //            datas.Add(data);

                    //            row++;
                    //            #endregion

                    //            #endregion
                    //            break;
                    //        case 2:
                    //            #region //初始化Data欄位

                    //            data = new Data()
                    //            {
                    //                cell = "A1",
                    //                css = "imported_class1",
                    //                format = "text",
                    //                value = "序號",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "B1",
                    //                css = "imported_class2",
                    //                format = "text",
                    //                value = "品號",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "C1",
                    //                css = "imported_class2",
                    //                format = "text",
                    //                value = "模號/品名",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "D1",
                    //                css = "imported_class2",
                    //                format = "text",
                    //                value = "接單日",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "E1",
                    //                css = "imported_class2",
                    //                format = "text",
                    //                value = "訂單數",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "F1",
                    //                css = "imported_class4",
                    //                format = "text",
                    //                value = "預計出貨日期",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "G1",
                    //                css = "imported_class5",
                    //                format = "text",
                    //                value = "預計出貨數量",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "H1",
                    //                css = "imported_class6",
                    //                format = "text",
                    //                value = "實際出貨日期",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "I1",
                    //                css = "imported_class7",
                    //                format = "text",
                    //                value = "實際出貨數量",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "J1",
                    //                css = "imported_class8",
                    //                format = "text",
                    //                value = "備註",
                    //            };
                    //            datas.Add(data);
                                
                    //            data = new Data()
                    //            {
                    //                cell = "K1",
                    //                css = "imported_class10",
                    //                format = "text",
                    //                value = "物流單號",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "L1",
                    //                css = "imported_class10",
                    //                format = "text",
                    //                value = "發貨地",
                    //            };
                    //            datas.Add(data);

                              
                    //            data = new Data()
                    //            {
                    //                cell = "M1",
                    //                css = "imported_class10",
                    //                format = "text",
                    //                value = "保留欄位1",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "N1",
                    //                css = "imported_class10",
                    //                format = "text",
                    //                value = "保留欄位2",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "O1",
                    //                css = "imported_class10",
                    //                format = "text",
                    //                value = "保留欄位3",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "P1",
                    //                css = "imported_class10",
                    //                format = "text",
                    //                value = "保留欄位4",
                    //            };
                    //            datas.Add(data);



                    //            #endregion

                    //            #region //資料
                    //            #region //設定

                    //            data = new Data()
                    //            {
                    //                cell = "A" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = PMDOrderId,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "B" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = MtlItemNo,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "C" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = CustomerMtlItemNo,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "D" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = OrderDate,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "E" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = SoQty,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "F" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = PlannedShipmentDate,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "G" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = PlannedShipmentQty,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "H" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = ActualShipmentDate,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "I" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = ActualShipmentQty,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "J" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = Remark,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "K" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = TrackingNumber,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "L" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = ShipFrom,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "M" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = PMDItem01,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "N" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = PMDItem02,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "O" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = PMDItem03,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "P" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = PMDItem04,
                    //            };
                    //            datas.Add(data);

                    //            row++;
                    //            #endregion

                    //            #endregion
                    //            break;
                    //        case 3:
                    //            #region //初始化Data欄位

                    //            data = new Data()
                    //            {
                    //                cell = "A1",
                    //                css = "imported_class1",
                    //                format = "text",
                    //                value = "序號",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "B1",
                    //                css = "imported_class2",
                    //                format = "text",
                    //                value = "品號",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "C1",
                    //                css = "imported_class2",
                    //                format = "text",
                    //                value = "模號/品名",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "D1",
                    //                css = "imported_class2",
                    //                format = "text",
                    //                value = "品名",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "E1",
                    //                css = "imported_class2",
                    //                format = "text",
                    //                value = "QTY",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "F1",
                    //                css = "imported_class4",
                    //                format = "text",
                    //                value = "CAV",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "G1",
                    //                css = "imported_class5",
                    //                format = "text",
                    //                value = "接單日期",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "H1",
                    //                css = "imported_class5",
                    //                format = "text",
                    //                value = "預計出貨數量",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "I1",
                    //                css = "imported_class6",
                    //                format = "text",
                    //                value = "實際出貨日期",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "J1",
                    //                css = "imported_class7",
                    //                format = "text",
                    //                value = "實際出貨數量",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "K1",
                    //                css = "imported_class10",
                    //                format = "text",
                    //                value = "模架管理號",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "L1",
                    //                css = "imported_class8",
                    //                format = "text",
                    //                value = "備註",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "M1",
                    //                css = "imported_class10",
                    //                format = "text",
                    //                value = "物流單號",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "N1",
                    //                css = "imported_class10",
                    //                format = "text",
                    //                value = "發貨地",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "O1",
                    //                css = "imported_class10",
                    //                format = "text",
                    //                value = "半成品發貨(ZY-JC)",
                    //            };
                    //            datas.Add(data);


                    //            data = new Data()
                    //            {
                    //                cell = "P1",
                    //                css = "imported_class10",
                    //                format = "text",
                    //                value = "保留欄位1",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "Q1",
                    //                css = "imported_class10",
                    //                format = "text",
                    //                value = "保留欄位2",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "R1",
                    //                css = "imported_class10",
                    //                format = "text",
                    //                value = "保留欄位3",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "S1",
                    //                css = "imported_class10",
                    //                format = "text",
                    //                value = "保留欄位4",
                    //            };
                    //            datas.Add(data);



                    //            #endregion

                    //            #region //資料
                    //            #region //設定

                    //            data = new Data()
                    //            {
                    //                cell = "A" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = PMDOrderId,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "B" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = MtlItemNo,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "C" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = CustomerMtlItemNo,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "D" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = MtlItemName,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "E" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = SoQty,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "F" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = CAV,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "G" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = OrderDate,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "H" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = PlannedShipmentDate,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "I" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = ActualShipmentDate,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "J" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = ActualShipmentQty,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "K" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = MoldFrameCtrlNo,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "L" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = Remark,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "M" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = TrackingNumber,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "N" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = ShipFrom,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "O" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = SemiDelivery,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "P" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = PMDItem01,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "Q" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = PMDItem02,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "R" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = PMDItem03,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "S" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = PMDItem04,
                    //            };
                    //            datas.Add(data);

                    //            row++;
                    //            #endregion

                    //            #endregion
                    //            break;
                    //        case 4:
                    //            #region //初始化Data欄位

                    //            data = new Data()
                    //            {
                    //                cell = "A1",
                    //                css = "imported_class1",
                    //                format = "text",
                    //                value = "序號",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "B1",
                    //                css = "imported_class2",
                    //                format = "text",
                    //                value = "退貨日期",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "C1",
                    //                css = "imported_class2",
                    //                format = "text",
                    //                value = "模號/品名",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "D1",
                    //                css = "imported_class2",
                    //                format = "text",
                    //                value = "品名",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "E1",
                    //                css = "imported_class2",
                    //                format = "text",
                    //                value = "退貨數量",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "F1",
                    //                css = "imported_class4",
                    //                format = "text",
                    //                value = "退貨問題點",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "G1",
                    //                css = "imported_class5",
                    //                format = "text",
                    //                value = "前期退貨數量",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "H1",
                    //                css = "imported_class5",
                    //                format = "text",
                    //                value = "前期欠貨數量",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "I1",
                    //                css = "imported_class6",
                    //                format = "text",
                    //                value = "共需補貨數量",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "J1",
                    //                css = "imported_class7",
                    //                format = "text",
                    //                value = "預計出貨日",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "K1",
                    //                css = "imported_class10",
                    //                format = "text",
                    //                value = "實際出貨日",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "L1",
                    //                css = "imported_class8",
                    //                format = "text",
                    //                value = "備註",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "M1",
                    //                css = "imported_class10",
                    //                format = "text",
                    //                value = "物流單號",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "N1",
                    //                css = "imported_class10",
                    //                format = "text",
                    //                value = "發貨地",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "O1",
                    //                css = "imported_class10",
                    //                format = "text",
                    //                value = "保留欄位1",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "P1",
                    //                css = "imported_class10",
                    //                format = "text",
                    //                value = "保留欄位2",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "Q1",
                    //                css = "imported_class10",
                    //                format = "text",
                    //                value = "保留欄位3",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "R1",
                    //                css = "imported_class10",
                    //                format = "text",
                    //                value = "保留欄位4",
                    //            };
                    //            datas.Add(data);



                    //            #endregion

                    //            #region //資料
                    //            #region //設定

                    //            data = new Data()
                    //            {
                    //                cell = "A" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = PMDOrderId,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "B" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = ReturnDate,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "C" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = CustomerMtlItemNo,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "D" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = MtlItemName,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "E" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = ReturnQty,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "F" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = ReturnIssue,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "G" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = PrevReturnQty,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "H" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = PrevShortageQty,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "I" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = TotalReplenishmentQty,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "J" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = PlannedShipmentDate,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "K" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = ActualShipmentDate,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "L" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = Remark,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "M" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = TrackingNumber,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "N" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = ShipFrom,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "O" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = PMDItem01,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "P" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = PMDItem02,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "Q" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = PMDItem03,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "R" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = PMDItem04,
                    //            };
                    //            datas.Add(data);

                    //            row++;
                    //            #endregion

                    //            #endregion
                    //            break;

                    //        case 5:
                    //            #region //初始化Data欄位

                    //            data = new Data()
                    //            {
                    //                cell = "A1",
                    //                css = "imported_class1",
                    //                format = "text",
                    //                value = "序號",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "B1",
                    //                css = "imported_class2",
                    //                format = "text",
                    //                value = "退貨日期",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "C1",
                    //                css = "imported_class2",
                    //                format = "text",
                    //                value = "退貨數量(套)",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "D1",
                    //                css = "imported_class2",
                    //                format = "text",
                    //                value = "模號/品名",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "E1",
                    //                css = "imported_class2",
                    //                format = "text",
                    //                value = "退貨問題點",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "F1",
                    //                css = "imported_class4",
                    //                format = "text",
                    //                value = "預計出貨日期",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "G1",
                    //                css = "imported_class5",
                    //                format = "text",
                    //                value = "預計出貨數量(套)",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "H1",
                    //                css = "imported_class10",
                    //                format = "text",
                    //                value = "實際出貨日",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "I1",
                    //                css = "imported_class10",
                    //                format = "text",
                    //                value = "實際出貨數量(套)",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "J1",
                    //                css = "imported_class8",
                    //                format = "text",
                    //                value = "備註",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "K1",
                    //                css = "imported_class10",
                    //                format = "text",
                    //                value = "物流單號",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "L1",
                    //                css = "imported_class10",
                    //                format = "text",
                    //                value = "發貨地",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "M1",
                    //                css = "imported_class10",
                    //                format = "text",
                    //                value = "保留欄位1",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "N1",
                    //                css = "imported_class10",
                    //                format = "text",
                    //                value = "保留欄位2",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "O1",
                    //                css = "imported_class10",
                    //                format = "text",
                    //                value = "保留欄位3",
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "P1",
                    //                css = "imported_class10",
                    //                format = "text",
                    //                value = "保留欄位4",
                    //            };
                    //            datas.Add(data);



                    //            #endregion

                    //            #region //資料
                    //            #region //設定

                    //            data = new Data()
                    //            {
                    //                cell = "A" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = PMDOrderId,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "B" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = ReturnDate,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "C" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = ReturnQty,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "D" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = MtlItemName,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "E" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = ReturnIssue,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "F" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = PlannedShipmentDate,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "G" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = PlannedShipmentQty,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "H" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = ActualShipmentDate,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "I" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = ActualShipmentQty,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "J" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = Remark,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "K" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = TrackingNumber,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "L" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = ShipFrom,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "M" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = PMDItem01,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "N" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = PMDItem02,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "O" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = PMDItem03,
                    //            };
                    //            datas.Add(data);

                    //            data = new Data()
                    //            {
                    //                cell = "P" + row,
                    //                css = "",
                    //                format = "text",
                    //                value = PMDItem04,
                    //            };
                    //            datas.Add(data);
                    //            row++;
                    //            #endregion

                    //            #endregion
                    //            break;
                    //    }





                    //}
                    //#endregion
                    //#region //整合Spreadsheet格式
                    //List<Sheets> sheetss = new List<Sheets>();

                    //Sheets sheets = new Sheets()
                    //{
                    //    name = "sheet1",
                    //    data = datas
                    //};
                    //sheetss.Add(sheets);

                    //#region //更新SpreadsheetData
                    //SpreadsheetModel spreadsheetModel = new SpreadsheetModel()
                    //{
                    //    sheets = sheetss
                    //};
                    //string SpreadsheetData = JsonConvert.SerializeObject(spreadsheetModel);

                    //#endregion
                    //#endregion
                    //#endregion

                    //#endregion
                    #endregion

                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = result
                    });
                }
            }
            catch (Exception e)
            {
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //GetPMDOrderId 取得 PMD 訂單ID GPAI -- 2025-05-06
        public string GetPMDOrderId(int TypeId)
        {

            var pMDOrderId = -1;
            try
            {

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //判斷公司資料是否正確
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 CompanyId, CompanyName, ErpDb
                            FROM BAS.Company
                            WHERE CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0) throw new SystemException("【公司】資料錯誤!");

                    foreach (var item in result)
                    {
                        ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                    }
                    #endregion

                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT ISNULL((
    SELECT TOP 1 a.PMDOrderTempId
    FROM SCM.PMDOrderTemp a 
    INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId 
    ORDER BY a.PMDOrderTempId DESC
), 0) AS PMDOrderId
                            ";

                    //dynamicParameters.Add("TypeId", TypeId);
                    //sql += " ORDER BY PMDOrderId DESC";

                    var PMDresult = sqlConnection.Query(sql, dynamicParameters);
                    pMDOrderId = sqlConnection.QueryFirstOrDefault<int>(sql, dynamicParameters);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = pMDOrderId
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

        #region //GetCurrentProcess 取得當前品號最新加工站 GPAI -- 2025-05-06
        public string GetCurrentProcess(int MtlItemId)
        {

            var pMDOrderId = "";
            try
            {

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //判斷公司資料是否正確
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 CompanyId, CompanyName, ErpDb
                            FROM BAS.Company
                            WHERE CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0) throw new SystemException("【公司】資料錯誤!");

                    foreach (var item in result)
                    {
                        ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                    }
                    #endregion

                    dynamicParameters = new DynamicParameters();
                    sql = @"select top 1 b.ProcessAlias
from MES.BarcodeProcess a
inner join MES.MoProcess b on a.MoProcessId = b.MoProcessId
inner join MES.ManufactureOrder c on a.MoId = c.MoId
inner join MES.WipOrder d on c.WoId = d.WoId
where d.MtlItemId = @MtlItemId
ORDER BY BarcodeProcessId DESC
                            ";

                    dynamicParameters.Add("MtlItemId", MtlItemId);
                    //sql += " ORDER BY PMDOrderId DESC";

                    var PMDresult = sqlConnection.Query(sql, dynamicParameters);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = PMDresult
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

        #region //GetCurrentEdit 取得當前頁面編輯者 GPAI -- 2025-05-06
        public string GetCurrentEdit(string PageNames)
        {

            var pMDOrderId = "";
            try
            {

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //判斷公司資料是否正確
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 CompanyId, CompanyName, ErpDb
                            FROM BAS.Company
                            WHERE CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0) throw new SystemException("【公司】資料錯誤!");

                    foreach (var item in result)
                    {
                        ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                    }
                    #endregion

                    dynamicParameters = new DynamicParameters();
                    sql = @"select top 1 a.* , (b.UserNo + '-' + b.UserName) UserName
                            from SCM.PageLock a
							inner join BAS.[User] b on a.LockedBy = b.UserId
                            where a.PageName = @PageName
                            order by PageLockId desc
                            ";
                    dynamicParameters.Add("PageName", PageNames);
                    //sql += " ORDER BY PMDOrderId DESC";

                    var PMDresult = sqlConnection.Query(sql, dynamicParameters);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = PMDresult
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

        #region //GetCustomer -- 取得客戶資料  GPAI -- 2025-05-06
        public string GetCustomer(int CustomerId, string CustomerNo, string CustomerName
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
                        @", a.CustomerNo, a.CustomerName, a.CustomerEnglishName, a.CustomerShortName
                          , a.RelatedPerson, FORMAT(a.PermitDate, 'yyyy-MM-dd') PermitDate, a.Version
                          , a.ResponsiblePerson, a.Contact, a.TelNoFirst, a.TelNoSecond, a.FaxNo, a.Email
                          , a.GuiNumber, a.Capital, a.AnnualTurnover, a.Headcount, a.HomeOffice, a.Currency
                          , a.DepartmentId, a.CustomerKind, a.SalesmenId, a.PaymentSalesmenId
                          , FORMAT(a.InauguateDate, 'yyyy-MM-dd') InauguateDate, FORMAT(a.CloseDate, 'yyyy-MM-dd') CloseDate
                          , a.ZipCodeRegister, a.RegisterAddressFirst, a.RegisterAddressSecond
                          , a.ZipCodeInvoice, a.InvoiceAddressFirst, a.InvoiceAddressSecond
                          , a.ZipCodeDelivery, a.DeliveryAddressFirst, a.DeliveryAddressSecond
                          , a.ZipCodeDocument, a.DocumentAddressFirst, a.DocumentAddressSecond
                          , a.BillReceipient, a.ZipCodeBill, a.BillAddressFirst, a.BillAddressSecond
                          , a.InvocieAttachedStatus, a.DepositRate, a.TaxAmountCalculateType, a.SaleRating, a.CreditRating
                          , a.TradeTerm, a.PaymentTerm, a.PricingType, a.ClearanceType, a.DocumentDeliver
                          , a.ReceiptReceive, a.PaymentType, a.TaxNo, a.InvoiceCount, a.Taxation
                          , a.Country, a.Region, a.Route, a.UploadType, a.PaymentBankFirst, a.BankAccountFirst
                          , a.PaymentBankSecond, a.BankAccountSecond, a.PaymentBankThird, a.BankAccountThird
                          , a.Account, a.AccountInvoice, a.AccountDay, a.ShipMethod, a.ShipType, a.ForwarderId
                          , a.CustomerRemark, a.CreditLimit, a.CreditLimitControl, a.CreditLimitControlCurrency
                          , a.SoCreditAuditType, a.SiCreditAuditType, a.DoCreditAuditType, a.InTransitCreditAuditType
                          , a.TransferStatus, a.TransferDate, a.Status
                          , a.CustomerNo + ' ' + a.CustomerShortName CustomerWithNo";
                    sqlQuery.mainTables =
                        @"FROM SCM.Customer a";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerId", @" AND a.CustomerId = @CustomerId", CustomerId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerNo", @" AND a.CustomerNo LIKE '%' + @CustomerNo + '%'", CustomerNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerName", @" AND a.CustomerName LIKE '%' + @CustomerName + '%'", CustomerName);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.CustomerNo";
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

        #region //GetPMDPage 取得頁面 GPAI -- 2025-05-06
        public string GetPMDPage()
        {

            var pMDOrderId = "";
            try
            {

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //判斷公司資料是否正確
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 CompanyId, CompanyName, ErpDb
                            FROM BAS.Company
                            WHERE CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0) throw new SystemException("【公司】資料錯誤!");

                    foreach (var item in result)
                    {
                        ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                    }
                    #endregion

                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT 
                            PMDPageId,
                            PageName,
                            PageNo,
                            CreateDate,
                            LastModifiedDate,
                            CreateBy,
                            LastModifiedBy
                            FROM SCM.PMDPage;
                            ";
                    //dynamicParameters.Add("PageName", PageNames);
                    //sql += " ORDER BY PMDOrderId DESC";

                    var PMDresult = sqlConnection.Query(sql, dynamicParameters);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = PMDresult
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

        #region //PMDAlertMamo 進度表推送MAMO異常通知 GPAI -- 2025-05-06
        public string PMDAlertMamo(string Company, string PageName, string ItemName, string CustomerMtlItem, string MtlItemName)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //確認公司資訊
                    int companyId = -1;
                    string CompanyNo = "";
                    if (Company == "")
                    {
                        #region //確認公司別DB
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb, a.CompanyId, a.CompanyNo
                            FROM BAS.Company a
                            WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var companyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (companyResult.Count() <= 0) throw new SystemException("公司別錯誤!");

                        foreach (var item in companyResult)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            companyId = item.CompanyId;
                            CompanyNo = item.CompanyNo;
                        }
                        #endregion
                    }
                    else
                    {
                        #region //確認公司別DB
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb, a.CompanyId, a.CompanyNo
                            FROM BAS.Company a
                            WHERE CompanyNo = @Company";
                        dynamicParameters.Add("Company", Company);

                        var companyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (companyResult.Count() <= 0) throw new SystemException("公司別錯誤!");

                        foreach (var item in companyResult)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            companyId = item.CompanyId;
                            CompanyNo = item.CompanyNo;

                        }
                        #endregion
                    }
                    #endregion

                    #region //MAMO推播通知
                    string Content = "";


                    Content = "### 【進度寄送失敗通知】\n" +
                                       "** 單據資料異常，請確認該筆資料各欄位是否填寫正確!! **" + "\n" +
                                       "- 表單名: " + PageName + "\n" +
                                       "- 異常欄位: " + ItemName + "\n" +
                                       "- 模号/品号: "  + "\n" +
                                       "- 品名: " + "\n" ;

                    #region //確認推播群組
                    //string SendId = "";
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.ChannelId
                                FROM MAMO.Channels a
                                WHERE a.ChannelId = @SendId";
                    dynamicParameters.Add("SendId", 247);

                    var MamoChannelResult = sqlConnection.Query(sql, dynamicParameters);

                    if (MamoChannelResult.Count() <= 0) throw new SystemException("請確認是否已設定推播群組!!");

                    #endregion

                    #region //取得標記USER資料
                    List<string> Tags = new List<string>();
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.UserId
                                , b.UserNo, b.UserName
                                FROM MAMO.ChannelMembers a 
                                INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                                WHERE a.ChannelId = @SendId";
                    dynamicParameters.Add("SendId", 247);

                    var UserResult = sqlConnection.Query(sql, dynamicParameters);

                    foreach (var item in UserResult)
                    {
                        Tags.Add(item.UserNo);
                    }
                    #endregion

                    List<int> Files = new List<int>();

                    string MamoResult = mamoHelper.SendMessage(CompanyNo, 945, "Channel", "247", Content, Tags, Files);

                    JObject MamoResultJson = JObject.Parse(MamoResult);
                    if (MamoResultJson["status"].ToString() != "success")
                    {
                        throw new SystemException(MamoResultJson["msg"].ToString());
                    }
                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = MamoResult
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

        #region //GetPMDOrderCustomerList 取得PMD客戶列表 GPAI -- 2025-05-06
        public string GetPMDOrderCustomerList(int TypeId)
        {

            var pMDOrderId = "";
            try
            {

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //判斷公司資料是否正確
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 CompanyId, CompanyName, ErpDb
                            FROM BAS.Company
                            WHERE CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0) throw new SystemException("【公司】資料錯誤!");

                    foreach (var item in result)
                    {
                        ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                    }
                    #endregion

                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DISTINCT c.CustomerNo, c.CustomerName, c.CustomerId, c.CustomerNo + ' ' + c.CustomerShortName CustomerWithNo
FROM SCM.PMDOrderTemp a 
INNER JOIN SCM.Customer c ON a.CustomerId = c.CustomerId
WHERE c.CompanyId = @CompanyId AND a.SendType = 0 AND a.TypeId = @TypeId
ORDER BY c.CustomerId
                            ";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    dynamicParameters.Add("TypeId", TypeId);

                    //sql += " ORDER BY PMDOrderId DESC";

                    var PMDresult = sqlConnection.Query(sql, dynamicParameters);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = PMDresult
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

        #region //GetPMDOrderDataSend -- 取得 PMD 訂單資料 GPAI -- 2025-05-06
        public string GetPMDOrderDataSend(int TypeId, int SendCustomer, string Company)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    int companyId = -1;
                    if (Company == "")
                    {
                        #region //確認公司別DB
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb, a.CompanyId
                            FROM BAS.Company a
                            WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var companyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (companyResult.Count() <= 0) throw new SystemException("公司別錯誤!");

                        foreach (var item in companyResult)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            companyId = item.CompanyId;

                        }
                        #endregion
                    }
                    else
                    {
                        #region //確認公司別DB
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb, a.CompanyId
                            FROM BAS.Company a
                            WHERE CompanyNo = @Company";
                        dynamicParameters.Add("Company", Company);

                        var companyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (companyResult.Count() <= 0) throw new SystemException("公司別錯誤!");

                        foreach (var item in companyResult)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            companyId = item.CompanyId;
                        }
                        #endregion
                    }


                    var CustomerId = -1;
                    if (SendCustomer > 0)
                    {
                        dynamicParameters = new DynamicParameters();
                        sql = @"select * 
                            from SCM.Customer
                            where CustomerId = @CustomerId";
                        dynamicParameters.Add("CustomerId", SendCustomer);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        foreach (var iten in result)
                        {
                            CustomerId = iten.CustomerId;
                        }
                    }



                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.PMDOrderTempId, a.MtlItemId, a.SoQty, a.SpareQty, a.OrderDate, a.DrawingAttachmentDate, a.CurrentProcess, a.CustomerDueDate, a.ConfirmedDueDate
                            , a.DelayRemark, a.MaterialNumber, a.Remark, a.TrackingNumber, a.ShipFrom, a.SemiDelivery, a.TypeId, a.PMDItem01, a.PMDItem02, a.PMDItem03, a.PMDItem04, a.PlannedShipmentDate
                            , a.ActualShipmentDate, a.PlannedShipmentQty, a.ActualShipmentQty, a.MoldFrameCtrlNo, a.ReturnQty, a.ReturnIssue, a.PrevReturnQty, a.PrevShortageQty, a.TotalReplenishmentQty
                            , a.CAV, a.ReturnDate, a.SoDetailId
                            , a.CreateDate, a.LastModifiedDate, a.CreateBy, a.LastModifiedBy, b.MtlItemNo, a.MtlItemName, a.CustomerMtlItemNo, c.CustomerNo, c.CustomerName, c.CustomerId
                            FROM SCM.PMDOrderTemp a 
                            LEFT JOIN PDM.MtlItem b on a.MtlItemId = b.MtlItemId 
							INNER JOIN SCM.Customer c ON a.CustomerId = c.CustomerId
                            WHERE 1=1 and c.CompanyId =  @CompanyId and a.SendType = 0
                            ";
                    dynamicParameters.Add("CompanyId", companyId);

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TypeId", " AND a.TypeId = @TypeId", TypeId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CustomerId", " AND a.CustomerId = @CustomerId", CustomerId);

                    sql += " ORDER BY a.PMDOrderTempId DESC, a.SoDetailId ";

                    var result1 = sqlConnection.Query(sql, dynamicParameters);

                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.PMDOrderId, a.MtlItemId, a.SoQty, a.SpareQty, a.OrderDate, a.DrawingAttachmentDate, a.CurrentProcess, a.CustomerDueDate, a.ConfirmedDueDate
                            , a.DelayRemark, a.MaterialNumber, a.Remark, a.TrackingNumber, a.ShipFrom, a.SemiDelivery, a.TypeId, a.PMDItem01, a.PMDItem02, a.PMDItem03, a.PMDItem04, a.PlannedShipmentDate
                            , a.ActualShipmentDate, a.PlannedShipmentQty, a.ActualShipmentQty, a.MoldFrameCtrlNo, a.ReturnQty, a.ReturnIssue, a.PrevReturnQty, a.PrevShortageQty, a.TotalReplenishmentQty
                            , a.CAV, a.ReturnDate, a.SoDetailId, a.PMDOrderTempId
                            , a.CreateDate, a.LastModifiedDate, a.CreateBy, a.LastModifiedBy, b.MtlItemNo, a.MtlItemName, a.CustomerMtlItemNo, c.CustomerNo, c.CustomerName, c.CustomerId
                            FROM SCM.PMDOrder a 
                            LEFT JOIN PDM.MtlItem b on a.MtlItemId = b.MtlItemId 
							INNER JOIN SCM.Customer c ON a.CustomerId = c.CustomerId
                            WHERE 1=1 and c.CompanyId =  @CompanyId and a.SendType = 0
                            ";
                    dynamicParameters.Add("CompanyId", companyId);

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TypeId", " AND a.TypeId = @TypeId", TypeId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CustomerId", " AND a.CustomerId = @CustomerId", CustomerId);

                    sql += " ORDER BY a.PMDOrderId , a.SoDetailId ";

                    var result2 = sqlConnection.Query(sql, dynamicParameters);


                    #region //比較兩邊寄送數據是否相同
                    // 將查詢結果轉換為匿名物件，排除主鍵欄位
                    var list1 = result1.Select(r => new {
                        MtlItemId = r.MtlItemId,
                        SoQty = r.SoQty,
                        SpareQty = r.SpareQty,
                        OrderDate = r.OrderDate,
                        DrawingAttachmentDate = r.DrawingAttachmentDate,
                        CurrentProcess = r.CurrentProcess,
                        CustomerDueDate = r.CustomerDueDate,
                        ConfirmedDueDate = r.ConfirmedDueDate,
                        DelayRemark = r.DelayRemark,
                        MaterialNumber = r.MaterialNumber,
                        Remark = r.Remark,
                        TrackingNumber = r.TrackingNumber,
                        ShipFrom = r.ShipFrom,
                        SemiDelivery = r.SemiDelivery,
                        TypeId = r.TypeId,
                        PMDItem01 = r.PMDItem01,
                        PMDItem02 = r.PMDItem02,
                        PMDItem03 = r.PMDItem03,
                        PMDItem04 = r.PMDItem04,
                        PlannedShipmentDate = r.PlannedShipmentDate,
                        ActualShipmentDate = r.ActualShipmentDate,
                        PlannedShipmentQty = r.PlannedShipmentQty,
                        ActualShipmentQty = r.ActualShipmentQty,
                        MoldFrameCtrlNo = r.MoldFrameCtrlNo,
                        ReturnQty = r.ReturnQty,
                        ReturnIssue = r.ReturnIssue,
                        PrevReturnQty = r.PrevReturnQty,
                        PrevShortageQty = r.PrevShortageQty,
                        TotalReplenishmentQty = r.TotalReplenishmentQty,
                        CAV = r.CAV,
                        ReturnDate = r.ReturnDate,
                        SoDetailId = r.SoDetailId,
                        MtlItemNo = r.MtlItemNo,
                        MtlItemName = r.MtlItemName,
                        CustomerMtlItemNo = r.CustomerMtlItemNo,
                        CustomerNo = r.CustomerNo,
                        CustomerName = r.CustomerName,
                        CustomerId = r.CustomerId
                    }).OrderBy(x => x.MtlItemId).ToList();

                    var list2 = result2.Select(r => new {
                        MtlItemId = r.MtlItemId,
                        SoQty = r.SoQty,
                        SpareQty = r.SpareQty,
                        OrderDate = r.OrderDate,
                        DrawingAttachmentDate = r.DrawingAttachmentDate,
                        CurrentProcess = r.CurrentProcess,
                        CustomerDueDate = r.CustomerDueDate,
                        ConfirmedDueDate = r.ConfirmedDueDate,
                        DelayRemark = r.DelayRemark,
                        MaterialNumber = r.MaterialNumber,
                        Remark = r.Remark,
                        TrackingNumber = r.TrackingNumber,
                        ShipFrom = r.ShipFrom,
                        SemiDelivery = r.SemiDelivery,
                        TypeId = r.TypeId,
                        PMDItem01 = r.PMDItem01,
                        PMDItem02 = r.PMDItem02,
                        PMDItem03 = r.PMDItem03,
                        PMDItem04 = r.PMDItem04,
                        PlannedShipmentDate = r.PlannedShipmentDate,
                        ActualShipmentDate = r.ActualShipmentDate,
                        PlannedShipmentQty = r.PlannedShipmentQty,
                        ActualShipmentQty = r.ActualShipmentQty,
                        MoldFrameCtrlNo = r.MoldFrameCtrlNo,
                        ReturnQty = r.ReturnQty,
                        ReturnIssue = r.ReturnIssue,
                        PrevReturnQty = r.PrevReturnQty,
                        PrevShortageQty = r.PrevShortageQty,
                        TotalReplenishmentQty = r.TotalReplenishmentQty,
                        CAV = r.CAV,
                        ReturnDate = r.ReturnDate,
                        SoDetailId = r.SoDetailId,
                        MtlItemNo = r.MtlItemNo,
                        MtlItemName = r.MtlItemName,
                        CustomerMtlItemNo = r.CustomerMtlItemNo,
                        CustomerNo = r.CustomerNo,
                        CustomerName = r.CustomerName,
                        CustomerId = r.CustomerId
                    }).OrderBy(x => x.MtlItemId).ToList();

                    // 比較兩個列表
                    bool isEqual = list1.SequenceEqual(list2);

                    if (isEqual)
                    {
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            data = result2
                        });
                    }
                    else
                    {
                        throw new SystemException("Excel資料與後台資料不符，無法寄送!");
                    }
                    #endregion
                    
                }
            }
            catch (Exception e)
            {
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion


        #endregion

        #region //Add
        #region //AddSaleOrder -- 訂單單頭資料新增 -- Zoey 2022.07.07
        public string AddSaleOrder(int SoId, int DepartmentId, string SoErpPrefix, string SoErpNo, string SoDate
            , string DocDate, string SoRemark, int CustomerId, int SalesmenId, string CustomerAddressFirst, string CustomerAddressSecond
            , string CustomerPurchaseOrder, string DepositPartial, double DepositRate, string Currency, double ExchangeRate
            , string TaxNo, string Taxation, double BusinessTaxRate, string DetailMultiTax, double TotalQty, double Amount
            , double TaxAmount, string ShipMethod, string TradeTerm, string PaymentTerm, string PriceTerm, string ConfirmStatus
            , int ConfirmUserId, string TransferStatus)
        {
            try
            {
                if (SoErpPrefix.Length <= 0) throw new SystemException("【訂單單別】不能為空!");
                if (DocDate.Length <= 0) throw new SystemException("【單據日期】不能為空!");
                if (CustomerId <= 0) throw new SystemException("【客戶名稱】不能為空!");
                if (SoRemark.Length > 255) throw new SystemException("【單頭備註】長度錯誤!");
                if (DepartmentId <= 0) throw new SystemException("【部門名稱】不能為空!");
                if (SalesmenId <= 0) throw new SystemException("【業務人員】不能為空!");
                if (CustomerAddressFirst.Length <= 0) throw new SystemException("【客戶聯絡地址(一)】不能為空!");
                if (CustomerAddressFirst.Length > 200) throw new SystemException("【客戶聯絡地址(一)】長度錯誤!");
                if (CustomerAddressSecond.Length > 200) throw new SystemException("【客戶聯絡地址(二)】長度錯誤!");
                if (Currency.Length <= 0) throw new SystemException("【交易幣別】不能為空!");
                if (ExchangeRate <= 0) throw new SystemException("【匯率】不能為空!");
                if (TaxNo.Length <= 0) throw new SystemException("【稅別碼】不能為空!");
                if (Taxation.Length <= 0) throw new SystemException("【課稅別】不能為空!");
                if (BusinessTaxRate < 0) throw new SystemException("【營業稅率】不能為空!");
                if (DetailMultiTax.Length <= 0) throw new SystemException("【單身多稅率】不能為空!");
                if (TradeTerm.Length <= 0) throw new SystemException("【交易條件】不能為空!");
                if (PaymentTerm.Length <= 0) throw new SystemException("【付款條件】不能為空!");
                if (DepositPartial.Length <= 0) throw new SystemException("【訂金分批】不能為空!");
                if (DepositRate < 0) throw new SystemException("【訂金比率】不能為空!");
                if (CustomerPurchaseOrder.Length > 20) throw new SystemException("【客戶採購單】長度錯誤!");
                if (ShipMethod.Length <= 0) throw new SystemException("【運輸方式】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        SoErpNo = BaseHelper.RandomCode(11);

                        #region //判斷單別+單號是否重複
                        sql = @"SELECT TOP 1 1
                                FROM SCM.SaleOrder
                                WHERE SoErpPrefix = @SoErpPrefix
                                AND SoErpNo = @SoErpNo
                                AND SoId != @SoId";
                        dynamicParameters.Add("SoErpPrefix", SoErpPrefix);
                        dynamicParameters.Add("SoErpNo", SoErpNo);
                        dynamicParameters.Add("SoId", SoId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【單別單號】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.SaleOrder (CompanyId, DepartmentId, SoErpPrefix, SoErpNo, Version, SoDate, DocDate, SoRemark
                                , CustomerId, SalesmenId, CustomerAddressFirst, CustomerAddressSecond, CustomerPurchaseOrder, DepositPartial
                                , DepositRate, Currency, ExchangeRate, TaxNo, Taxation, BusinessTaxRate, DetailMultiTax, TotalQty, Amount
                                , TaxAmount, ShipMethod, TradeTerm, PaymentTerm, PriceTerm, ConfirmStatus, ConfirmUserId, TransferStatus
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.SoId, INSERTED.SoErpNo
                                VALUES (@CompanyId, @DepartmentId, @SoErpPrefix, @SoErpNo, @Version, @SoDate, @DocDate, @SoRemark
                                , @CustomerId, @SalesmenId, @CustomerAddressFirst, @CustomerAddressSecond, @CustomerPurchaseOrder, @DepositPartial
                                , @DepositRate, @Currency, @ExchangeRate, @TaxNo, @Taxation, @BusinessTaxRate, @DetailMultiTax, @TotalQty, @Amount
                                , @TaxAmount, @ShipMethod, @TradeTerm, @PaymentTerm, @PriceTerm, @ConfirmStatus, @ConfirmUserId, @TransferStatus
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                DepartmentId,
                                SoErpPrefix,
                                SoErpNo,
                                Version = "0000",
                                SoDate = DocDate,
                                DocDate,
                                SoRemark,
                                CustomerId,
                                SalesmenId,
                                CustomerAddressFirst,
                                CustomerAddressSecond,
                                CustomerPurchaseOrder,
                                DepositPartial,
                                DepositRate,
                                Currency,
                                ExchangeRate,
                                TaxNo,
                                Taxation,
                                BusinessTaxRate,
                                DetailMultiTax,
                                TotalQty = 0,
                                Amount = 0,
                                TaxAmount = 0,
                                ShipMethod,
                                TradeTerm,
                                PaymentTerm,
                                PriceTerm,
                                ConfirmStatus = 'N',
                                ConfirmUserId = -1,
                                TransferStatus = 'N',
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

        #region //AddSaleOrder02 -- 訂單資料新增 -- Zoey 2022.07.07
        public string AddSaleOrder02(string SoErpPrefix, string DocDate, int CustomerId
            , string CustomerPurchaseOrder, int DepartmentId, int SalesmenId, string SoRemark, string Currency)
        {
            try
            {
                if (SoErpPrefix.Length <= 0) throw new SystemException("【訂單單別】不能為空!");
                if (!DateTime.TryParse(DocDate, out DateTime tempDate)) throw new SystemException("【單據日期】格式錯誤!");
                if (CustomerId <= 0) throw new SystemException("【客戶】不能為空!");
                if (CustomerPurchaseOrder.Length > 20) throw new SystemException("【客戶採購單】長度錯誤!");
                if (DepartmentId <= 0) throw new SystemException("【部門】不能為空!");
                if (SalesmenId <= 0) throw new SystemException("【業務】不能為空!");
                if (SoRemark.Length > 255) throw new SystemException("【單頭備註】長度錯誤!");
                if (Currency.Length <= 0) throw new SystemException("【幣別】不能為空!");


                using (TransactionScope transactionScope = new TransactionScope())
                {
                    string companyNo = "", exchangeRateSource = "", taxation = "";
                    double exchangeRate = 0, exciseTax = 0;
                    List<Customer> resultCustomer = new List<Customer>();

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //公司別資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var resultCompany = sqlConnection.Query(sql, dynamicParameters);
                        if (resultCompany.Count() <= 0) throw new SystemException("【公司別】資料錯誤!");

                        foreach (var item in resultCompany)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            companyNo = item.ErpNo;
                        }
                        #endregion

                        #region //判斷客戶資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 *
                                FROM SCM.Customer
                                WHERE CustomerId = @CustomerId
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("CustomerId", CustomerId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        resultCustomer = sqlConnection.Query<Customer>(sql, dynamicParameters).ToList();
                        if (resultCustomer.Count() <= 0) throw new SystemException("【客戶】資料錯誤!");
                        #endregion

                        #region //判斷部門資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Department
                                WHERE DepartmentId = @DepartmentId
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("DepartmentId", DepartmentId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var resultDepartment = sqlConnection.Query(sql, dynamicParameters);
                        if (resultDepartment.Count() <= 0) throw new SystemException("【部門】資料錯誤!");
                        #endregion

                        #region //判斷業務資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
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
                                AND a.DetailCode = @DetailCode
                                AND c.Authority > 0";
                        dynamicParameters.Add("UserId", SalesmenId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("Status", "A");
                        dynamicParameters.Add("FunctionCode", "SaleOrderManagement");
                        dynamicParameters.Add("DetailCode", "add");

                        var resultSalesmen = sqlConnection.Query(sql, dynamicParameters);
                        if (resultSalesmen.Count() <= 0) throw new SystemException("【業務】資料錯誤!");
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        #region //單據設定
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MQ004, a.MQ005, a.MQ006, a.MQ044
                                FROM CMSMQ a
                                WHERE a.COMPANY = @CompanyNo
                                AND a.MQ001 = @ErpPrefix";
                        dynamicParameters.Add("CompanyNo", companyNo);
                        dynamicParameters.Add("ErpPrefix", SoErpPrefix);

                        var resultDocSetting = sqlConnection.Query(sql, dynamicParameters);
                        if (resultDocSetting.Count() <= 0) throw new SystemException("ERP單據設定資料不存在!");

                        foreach (var item in resultDocSetting)
                        {
                            exchangeRateSource = item.MQ044; //匯率來源
                        }
                        #endregion

                        #region //目前匯率
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 MG003, MG004, MG005, MG006
                                FROM CMSMG
                                WHERE MG001 = @Currency
                                AND MG002 <= @DateNow
                                ORDER BY MG002 DESC";
                        dynamicParameters.Add("Currency", Currency);
                        dynamicParameters.Add("DateNow", DateTime.Now.ToString("yyyyMMdd"));

                        var resultExchangeRate = sqlConnection.Query(sql, dynamicParameters);
                        if (resultExchangeRate.Count() <= 0) throw new SystemException("ERP交易幣別匯率不存在!");

                        foreach (var item in resultExchangeRate)
                        {
                            switch (exchangeRateSource)
                            {
                                case "I": //銀行買進匯率
                                    exchangeRate = Convert.ToDouble(item.MG003);
                                    break;
                                case "O": //銀行賣出匯率
                                    exchangeRate = Convert.ToDouble(item.MG004);
                                    break;
                                case "E": //報關買進匯率
                                    exchangeRate = Convert.ToDouble(item.MG005);
                                    break;
                                case "W": //報關賣出匯率
                                    exchangeRate = Convert.ToDouble(item.MG006);
                                    break;
                            }
                        }
                        #endregion

                        #region //稅別碼設定
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT RTRIM(LTRIM(NN001)) NN001, RTRIM(LTRIM(NN002)) NN002, NN004, RTRIM(LTRIM(NN006)) NN006
                                FROM CMSNN
                                WHERE NN001 = @TaxNo";
                        dynamicParameters.Add("TaxNo", resultCustomer.Select(x => x.TaxNo).FirstOrDefault());

                        var resultTax = sqlConnection.Query(sql, dynamicParameters);
                        if (resultTax.Count() <= 0) throw new SystemException("ERP稅別碼設定不存在!");

                        foreach (var item in resultTax)
                        {
                            exciseTax = Convert.ToDouble(item.NN004); //營業稅率
                            taxation = item.NN006; //課稅別
                        }
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //隨機取得單號資料
                        string SoErpNo = "";

                        bool checkSoErpNo = true;
                        while (checkSoErpNo)
                        {
                            SoErpNo = BaseHelper.RandomCode(11);

                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM SCM.SaleOrder
                                    WHERE SoErpPrefix = @SoErpPrefix
                                    AND SoErpNo = @SoErpNo";
                            dynamicParameters.Add("SoErpPrefix", SoErpPrefix);
                            dynamicParameters.Add("SoErpNo", SoErpNo);

                            var resultSoErpNo = sqlConnection.Query(sql, dynamicParameters);
                            checkSoErpNo = resultSoErpNo.Count() > 0;
                        }
                        #endregion

                        #region //單頭新增
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.SaleOrder (CompanyId, DepartmentId, SoErpPrefix, SoErpNo, Version
                                , SoDate, DocDate, SoRemark, CustomerId, SalesmenId, CustomerAddressFirst
                                , CustomerAddressSecond, CustomerPurchaseOrder, DepositPartial, DepositRate
                                , Currency, ExchangeRate, TaxNo, Taxation, BusinessTaxRate, DetailMultiTax
                                , TotalQty, Amount, TaxAmount, ShipMethod, TradeTerm, PaymentTerm, PriceTerm
                                , ConfirmStatus, TransferStatus
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.SoId, INSERTED.SoErpNo
                                VALUES (@CompanyId, @DepartmentId, @SoErpPrefix, @SoErpNo, @Version
                                , @SoDate, @DocDate, @SoRemark, @CustomerId, @SalesmenId, @CustomerAddressFirst
                                , @CustomerAddressSecond, @CustomerPurchaseOrder, @DepositPartial, @DepositRate
                                , @Currency, @ExchangeRate, @TaxNo, @Taxation, @BusinessTaxRate, @DetailMultiTax
                                , @TotalQty, @Amount, @TaxAmount, @ShipMethod, @TradeTerm, @PaymentTerm, @PriceTerm
                                , @ConfirmStatus, @TransferStatus
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                DepartmentId = DepartmentId <= 0 ? (int?)null : DepartmentId,
                                SoErpPrefix,
                                SoErpNo,
                                Version = "0000", //版次
                                SoDate = DocDate,
                                DocDate,
                                SoRemark,
                                CustomerId,
                                SalesmenId,
                                CustomerAddressFirst = resultCustomer.Select(x => x.RegisterAddressFirst).FirstOrDefault(),
                                CustomerAddressSecond = resultCustomer.Select(x => x.RegisterAddressSecond).FirstOrDefault(),
                                CustomerPurchaseOrder,
                                DepositPartial = "N", //訂金分批
                                DepositRate = 0, //訂金比率
                                Currency = Currency,
                                ExchangeRate = exchangeRate,
                                TaxNo = resultCustomer.Select(x => x.TaxNo).FirstOrDefault(),
                                Taxation = taxation,
                                BusinessTaxRate = exciseTax,
                                DetailMultiTax = "N", //單身多稅率
                                TotalQty = 0, //訂單總數量
                                Amount = 0, //訂單總金額
                                TaxAmount = 0, //訂單總稅額
                                ShipMethod = resultCustomer.Select(x => x.ShipMethod).FirstOrDefault(),
                                TradeTerm = resultCustomer.Select(x => x.TradeTerm).FirstOrDefault().Length <= 0 ? "1" : resultCustomer.Select(x => x.TradeTerm).FirstOrDefault(), //交易條件
                                PaymentTerm = resultCustomer.Select(x => x.PaymentTerm).FirstOrDefault(),
                                PriceTerm = "", //價格條件
                                ConfirmStatus = "N", //確認碼
                                TransferStatus = "N", //拋轉ERP狀態
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = insertResult.Count();
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

        #region //AddSoDetail -- 訂單單身資料新增 -- Zoey 2022.07.14
        public string AddSoDetail(int SoId, int SoDetailTempId, string SoSequence, int MtlItemId, string SoMtlItemName
            , int InventoryId, double SoQty, double UnitPrice, double Amount
            , string ProductType, float FreebieOrSpareQty, string PromiseDate
            , string Project, string SoDetailRemark, string QuotationErp)
        {
            try
            {
                if (SoSequence.Length <= 0) throw new SystemException("【流水號】不能為空!");
                if (SoSequence.Length > 4) throw new SystemException("【流水號】長度錯誤!");
                if (InventoryId <= 0) throw new SystemException("【庫別】不能為空!");
                //if (SoQty <= 0) throw new SystemException("【訂單數量】不能為空!");
                if (UnitPrice < 0) throw new SystemException("【單價】不能為空!");
                if (Amount < 0) throw new SystemException("【金額】不能為空!");
                if (ProductType.Length <= 0) throw new SystemException("【類型】不能為空!");
                if (FreebieOrSpareQty < 0) throw new SystemException("【贈/備品量】不能為空!");
                if (PromiseDate.Length <= 0) throw new SystemException("【預計出貨日】不能為空!");
                if (!DateTime.TryParse(PromiseDate, out DateTime tempDateTime)) throw new SystemException("【預計出貨日】格式錯誤!");
                if (Project.Length > 20) throw new SystemException("【專案代號】長度錯誤!");
                if (SoDetailRemark.Length > 255) throw new SystemException("【備註】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    string companyNo = "", currency = "", taxNo = "", soMtlItemSpec = "", taxation = "", customerNo = "";
                    int uomId = -1, unitRound = 0, totalRound = 0, rowsAffected = 0;
                    double exciseTax = 0;

                    decimal creditBalance = 0, totalAmount = 0;

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //公司別資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var resultCompany = sqlConnection.Query(sql, dynamicParameters);
                        if (resultCompany.Count() <= 0) throw new SystemException("【公司別】資料錯誤!");

                        foreach (var item in resultCompany)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            companyNo = item.ErpNo;
                        }
                        #endregion

                        #region //判斷訂單資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.Currency, a.TaxNo, a.ConfirmStatus
                                , b.CustomerNo
                                FROM SCM.SaleOrder a 
                                INNER JOIN SCM.Customer b ON a.CustomerId = b.CustomerId
                                WHERE a.SoId = @SoId
                                AND a.CompanyId = @CompanyId";
                        dynamicParameters.Add("SoId", SoId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("訂單資料錯誤!");

                        
                        string confirmStatus = "";
                        foreach (var item in result)
                        {
                            currency = item.Currency;
                            taxNo = item.TaxNo;
                            confirmStatus = item.ConfirmStatus;
                            customerNo = item.CustomerNo;
                        }

                        #region //判斷確認過無法再確認
                        if (confirmStatus != "N") throw new SystemException("訂單不是未確認狀態，無法新增!");
                        #endregion
                        #endregion

                        #region //判斷序號是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.SoDetail
                                WHERE SoId = @SoId
                                AND SoSequence = @SoSequence";
                        dynamicParameters.Add("SoId", SoId);
                        dynamicParameters.Add("SoSequence", SoSequence);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("流水號重複，請重新取號!");
                        #endregion

                        #region //判斷品號資料是否正確
                        if (MtlItemId > 0)
                        {
                            if (CurrentCompany==4) {
                                #region//晶彩-返修品品項不允訂單出售
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 MtlItemName, MtlItemSpec, InventoryId, ISNULL(SaleUomId, InventoryUomId) SaleUomId
                                    FROM PDM.MtlItem
                                    WHERE MtlItemId = @MtlItemId AND TypeThree!='313' AND TypeThree!=''
                                    AND CompanyId = @CompanyId";
                                dynamicParameters.Add("MtlItemId", MtlItemId);
                                dynamicParameters.Add("CompanyId", CurrentCompany);

                                var resultMtlItem = sqlConnection.Query(sql, dynamicParameters);
                                if (resultMtlItem.Count() <= 0) throw new SystemException("請確認品號資料是否正確! <br/> ps:晶彩-返修品品項不允訂單出售");

                                foreach (var item in resultMtlItem)
                                {
                                    SoMtlItemName = item.MtlItemName;
                                    soMtlItemSpec = item.MtlItemSpec;
                                    uomId = Convert.ToInt32(item.SaleUomId);
                                }
                                #endregion
                            }
                            else {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 MtlItemName, MtlItemSpec, InventoryId, ISNULL(SaleUomId, InventoryUomId) SaleUomId
                                    FROM PDM.MtlItem
                                    WHERE MtlItemId = @MtlItemId
                                    AND CompanyId = @CompanyId";
                                dynamicParameters.Add("MtlItemId", MtlItemId);
                                dynamicParameters.Add("CompanyId", CurrentCompany);

                                var resultMtlItem = sqlConnection.Query(sql, dynamicParameters);
                                if (resultMtlItem.Count() <= 0) throw new SystemException("品號資料錯誤!");

                                foreach (var item in resultMtlItem)
                                {
                                    SoMtlItemName = item.MtlItemName;
                                    soMtlItemSpec = item.MtlItemSpec;
                                    uomId = Convert.ToInt32(item.SaleUomId);
                                }
                            }

                        }
                        else
                        {
                            if (SoMtlItemName.Length <= 0) throw new SystemException("【品名】不能為空!");
                            if (SoMtlItemName.Length > 120) throw new SystemException("【品名】長度錯誤!");
                        }
                        #endregion

                        #region //判斷庫別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.Inventory
                                WHERE InventoryId = @InventoryId
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("InventoryId", InventoryId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var resultInventory = sqlConnection.Query(sql, dynamicParameters);
                        if (resultInventory.Count() <= 0) throw new SystemException("庫別資料錯誤!");
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        #region //交易幣別設定
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT MF003, MF004, MF005, MF006
                                FROM CMSMF
                                WHERE MF001 = @Currency";
                        dynamicParameters.Add("Currency", currency);

                        var resultCurrencySetting = sqlConnection.Query(sql, dynamicParameters);
                        if (resultCurrencySetting.Count() <= 0) throw new SystemException("ERP交易幣別設定不存在!");

                        foreach (var item in resultCurrencySetting)
                        {
                            unitRound = Convert.ToInt32(item.MF003); //單價取位
                            totalRound = Convert.ToInt32(item.MF004); //金額取位
                        }
                        #endregion

                        #region //稅別碼設定
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT RTRIM(LTRIM(NN001)) NN001, RTRIM(LTRIM(NN002)) NN002, NN004, RTRIM(LTRIM(NN006)) NN006
                                FROM CMSNN
                                WHERE NN001 = @TaxNo";
                        dynamicParameters.Add("TaxNo", taxNo);

                        var resultTax = sqlConnection.Query(sql, dynamicParameters);
                        if (resultTax.Count() <= 0) throw new SystemException("ERP稅別碼設定不存在!");

                        foreach (var item in resultTax)
                        {
                            exciseTax = Convert.ToDouble(item.NN004); //營業稅率
                            taxation = item.NN006; //課稅別
                        }
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //單身新增
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.SoDetail (SoId, SoSequence, MtlItemId, SoMtlItemName
                                , SoMtlItemSpec, InventoryId, UomId, SoQty, SiQty, ProductType, FreebieQty
                                , FreebieSiQty, SpareQty, SpareSiQty, UnitPrice, Amount, PromiseDate, PcPromiseDate
                                , Project, CustomerMtlItemNo, SoDetailRemark, SoPriceQty, SoPriceUomId
                                , TaxNo, BusinessTaxRate, DiscountRate, DiscountAmount
                                , ConfirmStatus, ClosureStatus, QuotationErp
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.SoDetailId
                                VALUES (@SoId, @SoSequence, @MtlItemId, @SoMtlItemName
                                , @SoMtlItemSpec, @InventoryId, @UomId, @SoQty, @SiQty, @ProductType, @FreebieQty
                                , @FreebieSiQty, @SpareQty, @SpareSiQty, @UnitPrice, @Amount, @PromiseDate, @PcPromiseDate
                                , @Project, @CustomerMtlItemNo, @SoDetailRemark, @SoPriceQty, @SoPriceUomId
                                , @TaxNo, @BusinessTaxRate, @DiscountRate, @DiscountAmount
                                , @ConfirmStatus, @ClosureStatus, @QuotationErp
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SoId,
                                SoSequence,
                                MtlItemId = MtlItemId <= 0 ? (int?)null : MtlItemId,
                                SoMtlItemName,
                                SoMtlItemSpec = soMtlItemSpec,
                                InventoryId,
                                UomId = uomId <= 0 ? (int?)null : uomId,
                                SoQty,
                                SiQty = 0, //已交數量
                                ProductType,
                                FreebieQty = ProductType == "1" ? FreebieOrSpareQty : 0,
                                FreebieSiQty = 0, //贈品已交數量
                                SpareQty = ProductType == "2" ? FreebieOrSpareQty : 0,
                                SpareSiQty = 0, //備品已交數量
                                UnitPrice = Math.Round(UnitPrice, unitRound, MidpointRounding.AwayFromZero),
                                Amount = Math.Round(Math.Round(UnitPrice, unitRound) * SoQty, totalRound, MidpointRounding.AwayFromZero),
                                PromiseDate,
                                PcPromiseDate = PromiseDate, //排定交貨日
                                Project,
                                CustomerMtlItemNo = "", //客戶品號
                                SoDetailRemark,
                                SoPriceQty = SoQty, //計價數量
                                SoPriceUomId = uomId <= 0 ? (int?)null : uomId, //計價單位
                                TaxNo = taxNo, //稅別碼
                                BusinessTaxRate = exciseTax, //營業稅率
                                DiscountRate = 1, //折扣率
                                DiscountAmount = 0, //折扣金額
                                ConfirmStatus = "N", //確認碼
                                ClosureStatus = "N", //結案碼
                                QuotationErp,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);
                        rowsAffected += insertResult.Count();
                        int soDetailId = -1;
                        foreach (var item in insertResult)
                        {
                            soDetailId = item.SoDetailId;
                        }
                        #endregion

                        if (SoDetailTempId > 0)
                        {
                            #region //判斷採購單匯入資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM SCM.SoDetailTemp
                                    WHERE SoDetailTempId = @SoDetailTempId";
                            dynamicParameters.Add("SoDetailTempId", SoDetailTempId);
                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【採購單】匯入錯誤!");
                            #endregion

                            #region //訂單身暫存綁定
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.SoDetailTemp SET 
                                    SoDetailId = @SoDetailId,
                                    Status = @Status,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE SoDetailTempId = @SoDetailTempId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    SoDetailTempId,
                                    Status = "Y", //綁定訂單身
                                    SoDetailId = soDetailId,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }

                        #region //單頭資料更新
                        double totalQty = 0, amount = 0, taxAmount = 0;

                        #region //所有單身資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT SoQty, SoPriceQty, ProductType, FreebieQty, SpareQty, UnitPrice, Amount
                                FROM SCM.SoDetail
                                WHERE SoId = @SoId
                                ORDER BY SoSequence";
                        dynamicParameters.Add("SoId", SoId);

                        var resultSoDetail = sqlConnection.Query(sql, dynamicParameters);
                        if (resultSoDetail.Count() <= 0) throw new SystemException("無單身資料!");

                        foreach (var item in resultSoDetail)
                        {
                            totalQty += Convert.ToDouble(item.SoQty) + Convert.ToDouble(item.FreebieQty) + Convert.ToDouble(item.SpareQty);
                            amount += Math.Round(Math.Round(Convert.ToDouble(item.UnitPrice), unitRound, MidpointRounding.AwayFromZero) * Convert.ToDouble(item.SoPriceQty), totalRound, MidpointRounding.AwayFromZero);
                        }

                        #region //計算數量與金額
                        switch (taxation)
                        {
                            case "1":
                                taxAmount = amount - Math.Round(amount / (1 + exciseTax), totalRound, MidpointRounding.AwayFromZero);
                                amount = Math.Round(amount / (1 + exciseTax), totalRound, MidpointRounding.AwayFromZero);
                                break;
                            case "2":
                                taxAmount = Math.Round(amount * exciseTax, totalRound, MidpointRounding.AwayFromZero);
                                break;
                            case "3":
                                break;
                            case "4":
                                break;
                            case "9":
                                break;
                        }
                        #endregion
                        #endregion

                        totalAmount = Convert.ToDecimal(amount + taxAmount);

                        #region //單頭更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.SaleOrder SET
                                TotalQty = @TotalQty,
                                Amount = @Amount,
                                TaxAmount = @TaxAmount,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SoId = @SoId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TotalQty = totalQty,
                                Amount = amount,
                                TaxAmount = taxAmount,
                                LastModifiedDate,
                                LastModifiedBy,
                                SoId
                            });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //AddSoModification -- 訂單變更單頭資料新增 -- Zoey 2022.08.09
        public string AddSoModification(int SomId, int SoId, string Version, int DepartmentId, string DocDate, string SoRemark
            , int SalesmenId, string CustomerAddressFirst, string CustomerAddressSecond, string CustomerPurchaseOrder
            , string DepositPartial, double DepositRate, string Currency, double ExchangeRate, string TaxNo, string Taxation
            , double BusinessTaxRate, string DetailMultiTax, string ShipMethod, string TradeTerm, string PaymentTerm, string PriceTerm            
            , string ClosureStatus, string ModiReason, string ConfirmStatus, int ConfirmUserId, string TransferStatus, string TransferDate)
        {
            try
            {
                if (DocDate.Length <= 0) throw new SystemException("【單據日期】不能為空!");
                if (CustomerPurchaseOrder.Length > 20) throw new SystemException("【客戶採購單】長度錯誤!");
                if (DepartmentId <= 0) throw new SystemException("【部門名稱】不能為空!");
                if (SalesmenId <= 0) throw new SystemException("【業務人員】不能為空!");
                if (SoRemark.Length > 255) throw new SystemException("【備註】長度錯誤!");
                if (CustomerAddressFirst.Length <= 0) throw new SystemException("【客戶聯絡地址(一)】不能為空!");
                if (CustomerAddressFirst.Length > 200) throw new SystemException("【客戶聯絡地址(一)】長度錯誤!");
                if (CustomerAddressSecond.Length > 200) throw new SystemException("【客戶聯絡地址(二)】長度錯誤!");
                if (Currency.Length <= 0) throw new SystemException("【交易幣別】不能為空!");
                if (ExchangeRate <= 0) throw new SystemException("【匯率】不能為空!");
                if (TradeTerm.Length <= 0) throw new SystemException("【交易條件】不能為空!");
                if (TaxNo.Length <= 0) throw new SystemException("【稅別碼】不能為空!");
                if (Taxation.Length <= 0) throw new SystemException("【課稅別】不能為空!");
                if (BusinessTaxRate < 0) throw new SystemException("【營業稅率】不能為空!");                
                if (PaymentTerm.Length <= 0) throw new SystemException("【付款條件】不能為空!");
                if (PriceTerm.Length > 40) throw new SystemException("【價格條件】長度錯誤!");
                if (DepositPartial.Length <= 0) throw new SystemException("【訂金分批】不能為空!");
                if (DetailMultiTax.Length <= 0) throw new SystemException("【單身多稅率】不能為空!");
                if (DepositRate < 0) throw new SystemException("【訂金比率】不能為空!");
                if (ShipMethod.Length <= 0) throw new SystemException("【運輸方式】不能為空!");
                if (ModiReason.Length > 255) throw new SystemException("【變更原因】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        sql = @"DECLARE @OriVersion NVARCHAR(4), @OriDepartmentId INT, @OriSoRemark NVARCHAR(255)
                                , @OriSalesmenId INT, @OriCustomerAddressFirst NVARCHAR(200)
                                , @OriCustomerAddressSecond NVARCHAR(200), @OriCustomerPurchaseOrder NVARCHAR(20)
                                , @OriDepositPartial NVARCHAR(1), @OriDepositRate FLOAT, @OriCurrency NVARCHAR(4)
                                , @OriExchangeRate FLOAT, @OriTaxNo  NVARCHAR(3), @OriTaxation  NVARCHAR(1)
                                , @OriBusinessTaxRate FLOAT, @OriDetailMultiTax NVARCHAR(1), @OriShipMethod NVARCHAR(1)
                                , @OriTradeTerm NVARCHAR(1), @OriPaymentTerm NVARCHAR(6), @OriPriceTerm NVARCHAR(40)

                                SELECT @OriVersion = Version
                                , @OriDepartmentId = DepartmentId
                                , @OriSoRemark = SoRemark
                                , @OriSalesmenId = SalesmenId
                                , @OriCustomerAddressFirst = CustomerAddressFirst
                                , @OriCustomerAddressSecond = CustomerAddressSecond
                                , @OriCustomerPurchaseOrder = CustomerPurchaseOrder
                                , @OriDepositPartial = DepositPartial
                                , @OriDepositRate = DepositRate
                                , @OriCurrency = Currency
                                , @OriExchangeRate = ExchangeRate
                                , @OriTaxNo = TaxNo
                                , @OriTaxation = Taxation
                                , @OriBusinessTaxRate = BusinessTaxRate
                                , @OriDetailMultiTax = DetailMultiTax
                                , @OriShipMethod = ShipMethod
                                , @OriTradeTerm = TradeTerm
                                , @OriPaymentTerm = PaymentTerm
                                , @OriPriceTerm = PriceTerm
                                FROM SCM.SaleOrder a
                                WHERE SoId = @SoId

                                INSERT INTO SCM.SoModification (SoId, Version, DepartmentId, DocDate, SoRemark, SalesmenId, CustomerAddressFirst 
                                , CustomerAddressSecond, CustomerPurchaseOrder, DepositPartial, DepositRate, Currency, ExchangeRate, TaxNo
                                , Taxation, BusinessTaxRate, DetailMultiTax, ShipMethod, TradeTerm, PaymentTerm, PriceTerm, OriVersion
                                , OriDepartmentId, OriSoRemark, OriSalesmenId, OriCustomerAddressFirst, OriCustomerAddressSecond
                                , OriCustomerPurchaseOrder, OriDepositPartial, OriDepositRate, OriCurrency, OriExchangeRate, OriTaxNo 
                                , OriTaxation, OriBusinessTaxRate, OriDetailMultiTax, OriShipMethod, OriTradeTerm, OriPaymentTerm, OriPriceTerm
                                , ClosureStatus, ModiReason, ConfirmStatus, ConfirmUserId, TransferStatus, TransferDate
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.SomId
                                VALUES (@SoId, @Version, @DepartmentId, @DocDate, @SoRemark, @SalesmenId, @CustomerAddressFirst
                                , @CustomerAddressSecond, @CustomerPurchaseOrder, @DepositPartial, @DepositRate, @Currency, @ExchangeRate, @TaxNo
                                , @Taxation, @BusinessTaxRate, @DetailMultiTax, @ShipMethod, @TradeTerm, @PaymentTerm, @PriceTerm, @OriVersion
                                , @OriDepartmentId, @OriSoRemark, @OriSalesmenId, @OriCustomerAddressFirst, @OriCustomerAddressSecond
                                , @OriCustomerPurchaseOrder, @OriDepositPartial, @OriDepositRate, @OriCurrency, @OriExchangeRate, @OriTaxNo
                                , @OriTaxation, @OriBusinessTaxRate, @OriDetailMultiTax, @OriShipMethod, @OriTradeTerm, @OriPaymentTerm, @OriPriceTerm
                                , @ClosureStatus, @ModiReason, @ConfirmStatus, @ConfirmUserId, @TransferStatus, @TransferDate
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";

                                //UPDATE SCM.SaleOrder SET
                                //Version = @Version,
                                //DepartmentId = @DepartmentId,
                                //SoRemark = @SoRemark,
                                //SalesmenId = @SalesmenId,
                                //CustomerAddressFirst = @CustomerAddressFirst,
                                //CustomerAddressSecond = @CustomerAddressSecond,
                                //CustomerPurchaseOrder = @CustomerPurchaseOrder,
                                //DepositPartial = @DepositPartial,
                                //DepositRate = @DepositRate,
                                //Currency = @Currency,
                                //ExchangeRate = @ExchangeRate,
                                //TaxNo = @TaxNo,
                                //Taxation = @Taxation,
                                //BusinessTaxRate = @BusinessTaxRate,
                                //DetailMultiTax = @DetailMultiTax,
                                //ShipMethod = @ShipMethod,
                                //TradeTerm = @TradeTerm,
                                //PaymentTerm = @PaymentTerm,
                                //PriceTerm = @PriceTerm
                                //WHERE SoId = @SoId
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SoId,
                                Version = "0001",
                                DepartmentId,
                                DocDate,
                                SoRemark,
                                SalesmenId,
                                CustomerAddressFirst,
                                CustomerAddressSecond,
                                CustomerPurchaseOrder,
                                DepositPartial,
                                DepositRate,
                                Currency,
                                ExchangeRate,
                                TaxNo,
                                Taxation,
                                BusinessTaxRate,
                                DetailMultiTax,
                                ShipMethod,
                                TradeTerm,
                                PaymentTerm,
                                PriceTerm,
                                ClosureStatus,
                                ModiReason,
                                ConfirmStatus = 'N',
                                ConfirmUserId,
                                TransferStatus = 'N',
                                TransferDate,
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

        #region //AddSoBpmLog -- 新增訂單LOG紀錄 -- Ann 2023-12-06
        public string AddSoBpmLog(int SoId, string BpmNo, string Status, string RootId, string ConfirmUser)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認訂單資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT b.BpmTransferDate
                                FROM SCM.SaleOrder a
                                INNER JOIN SCM.SoBpmInfo b ON a.SoId = b.SoId
                                WHERE a.SoId = @SoId";
                        dynamicParameters.Add("SoId", SoId);

                        var SaleOrderResult = sqlConnection.Query(sql, dynamicParameters);

                        if (SaleOrderResult.Count() <= 0) throw new SystemException("訂單資料錯誤!!");

                        DateTime BpmTransferDate = new DateTime();
                        foreach (var item in SaleOrderResult)
                        {
                            BpmTransferDate = item.BpmTransferDate;
                        }
                        #endregion

                        #region //INSERT SCM.SoBpmLog
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.SoBpmLog (SoId, RootId, BpmNo, TransferBpmDate, BpmStatus, ConfirmUser
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.SoBpmLogId
                                VALUES (@SoId, @RootId, @BpmNo, @TransferBpmDate, @BpmStatus, @ConfirmUser
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                          new
                          {
                              SoId,
                              RootId,
                              BpmNo,
                              TransferBpmDate = BpmTransferDate,
                              BpmStatus = Status,
                              ConfirmUser,
                              CreateDate,
                              LastModifiedDate,
                              CreateBy,
                              LastModifiedBy
                          });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = insertResult.Count();
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
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

        #region //AddPMDOrderLog -- 新增PMD訂單LOG紀錄 -- GPAI 2025/05/06
        public string AddPMDOrderLog(int UserId, int EditItem, string EditTable)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        #region //INSERT SCM.SoBpmLog
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.PMDOrderLog (UserId, EditTime, EditItem, EditTable
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.PMDOrderLogId
                                VALUES (@UserId, @EditTime, @EditItem, @EditTable
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";

                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                UserId,                    // 編輯者ID
                                EditTime = CreateDate,                  // 編輯時間
                                EditItem,                  // 編輯欄位名稱
                                EditTable,                 // 編輯表格名稱 (例如: "PMDOrder")
                                CreateDate,                // 新增日期
                                LastModifiedDate,          // 修改日期
                                CreateBy,                  // 新增人員
                                LastModifiedBy             // 修改人員
                            });
                    
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);
                        int rowsAffected = insertResult.Count();
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

        #region //Update
        #region //UpdateSaleOrder -- 訂單單頭資料更新 -- Zoey 2022.07.12
        public string UpdateSaleOrder(int SoId, int DepartmentId, string SoErpPrefix, string SoErpNo, string SoDate
            , string DocDate, string SoRemark, int CustomerId, int SalesmenId, string CustomerAddressFirst, string CustomerAddressSecond
            , string CustomerPurchaseOrder, string DepositPartial, double DepositRate, string Currency, double ExchangeRate
            , string TaxNo, string Taxation, double BusinessTaxRate, string DetailMultiTax, double TotalQty, double Amount
            , double TaxAmount, string ShipMethod, string TradeTerm, string PaymentTerm, string PriceTerm, string ConfirmStatus
            , int ConfirmUserId, string TransferStatus)
        {
            try
            {
                if (SoErpPrefix.Length <= 0) throw new SystemException("【訂單單別】不能為空!");
                if (SoErpNo.Length <= 0) throw new SystemException("【訂單單號】不能為空!");
                if (DocDate.Length <= 0) throw new SystemException("【單據日期】不能為空!");
                if (CustomerId <= 0) throw new SystemException("【客戶名稱】不能為空!");
                if (SoRemark.Length > 255) throw new SystemException("【單頭備註】長度錯誤!");
                if (DepartmentId <= 0) throw new SystemException("【部門名稱】不能為空!");
                if (SalesmenId <= 0) throw new SystemException("【業務人員】不能為空!");
                if (CustomerAddressFirst.Length <= 0) throw new SystemException("【客戶聯絡地址(一)】不能為空!");
                if (CustomerAddressFirst.Length > 200) throw new SystemException("【客戶聯絡地址(一)】長度錯誤!");
                if (CustomerAddressSecond.Length > 200) throw new SystemException("【客戶聯絡地址(二)】長度錯誤!");
                if (Currency.Length <= 0) throw new SystemException("【交易幣別】不能為空!");
                if (ExchangeRate <= 0) throw new SystemException("【匯率】不能為空!");
                if (TaxNo.Length <= 0) throw new SystemException("【稅別碼】不能為空!");
                if (Taxation.Length <= 0) throw new SystemException("【課稅別】不能為空!");
                if (BusinessTaxRate < 0) throw new SystemException("【營業稅率】不能為空!");
                if (DetailMultiTax.Length <= 0) throw new SystemException("【單身多稅率】不能為空!");
                if (TradeTerm.Length <= 0) throw new SystemException("【交易條件】不能為空!");
                if (PaymentTerm.Length <= 0) throw new SystemException("【付款條件】不能為空!");
                if (DepositPartial.Length <= 0) throw new SystemException("【訂金分批】不能為空!");
                if (DepositRate < 0) throw new SystemException("【訂金比率】不能為空!");
                if (CustomerPurchaseOrder.Length > 20) throw new SystemException("【客戶採購單】長度錯誤!");
                if (ShipMethod.Length <= 0) throw new SystemException("【運輸方式】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;

                        #region //判斷訂單單頭資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.SaleOrder
                                WHERE SoId = @SoId";
                        dynamicParameters.Add("SoId", SoId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("訂單單頭資料錯誤!");
                        #endregion

                        #region //判斷單別+單號是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.SaleOrder
                                WHERE SoErpPrefix = @SoErpPrefix
                                AND SoErpNo = @SoErpNo
                                AND SoId != @SoId";
                        dynamicParameters.Add("SoErpPrefix", SoErpPrefix);
                        dynamicParameters.Add("SoErpNo", SoErpNo);
                        dynamicParameters.Add("SoId", SoId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【單別單號】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.SaleOrder SET
                                DepartmentId = @DepartmentId,
                                SoErpNo = @SoErpNo,
                                SoDate = @SoDate,
                                DocDate = @DocDate,
                                SoRemark = @SoRemark,
                                CustomerId = @CustomerId,
                                SalesmenId = @SalesmenId,
                                CustomerAddressFirst = @CustomerAddressFirst,
                                CustomerAddressSecond = @CustomerAddressSecond,
                                CustomerPurchaseOrder = @CustomerPurchaseOrder,
                                DepositPartial = @DepositPartial,
                                DepositRate = @DepositRate,
                                Currency = @Currency,
                                ExchangeRate = @ExchangeRate,
                                TaxNo = @TaxNo,
                                Taxation = @Taxation,
                                BusinessTaxRate = @BusinessTaxRate,
                                DetailMultiTax = @DetailMultiTax,
                                ShipMethod = @ShipMethod,
                                TradeTerm = @TradeTerm,
                                PaymentTerm = @PaymentTerm,
                                PriceTerm = @PriceTerm,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SoId = @SoId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                DepartmentId,
                                SoErpNo,
                                SoDate = DocDate,
                                DocDate,
                                SoRemark,
                                CustomerId,
                                SalesmenId,
                                CustomerAddressFirst,
                                CustomerAddressSecond,
                                CustomerPurchaseOrder,
                                DepositPartial,
                                DepositRate,
                                Currency,
                                ExchangeRate,
                                TaxNo,
                                Taxation,
                                BusinessTaxRate,
                                DetailMultiTax,
                                ShipMethod,
                                TradeTerm,
                                PaymentTerm,
                                PriceTerm,
                                TransferStatus,
                                LastModifiedDate,
                                LastModifiedBy,
                                SoId
                            });
                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.SoDetail SET
                                TaxNo = @TaxNo,
                                BusinessTaxRate = @BusinessTaxRate,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SoId = @SoId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TaxNo,
                                BusinessTaxRate,
                                LastModifiedDate,
                                LastModifiedBy,
                                SoId
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

        #region //UpdateSaleOrder02 -- 訂單資料更新 -- Zoey 2022.07.12
        public string UpdateSaleOrder02(int SoId, string SoErpPrefix, string DocDate, int CustomerId
            , string CustomerPurchaseOrder, int DepartmentId, int SalesmenId, string SoRemark, string Currency)
        {
            try
            {
                if (SoErpPrefix.Length <= 0) throw new SystemException("【訂單單別】不能為空!");
                if (!DateTime.TryParse(DocDate, out DateTime tempDate)) throw new SystemException("【單據日期】格式錯誤!");
                if (CustomerId <= 0) throw new SystemException("【客戶】不能為空!");
                if (CustomerPurchaseOrder.Length > 20) throw new SystemException("【客戶採購單】長度錯誤!");
                if (DepartmentId <= 0) throw new SystemException("【部門】不能為空!");
                if (SalesmenId <= 0) throw new SystemException("【業務】不能為空!");
                if (SoRemark.Length > 255) throw new SystemException("【單頭備註】長度錯誤!");
                if (Currency.Length <= 0) throw new SystemException("【幣別】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    string companyNo = "", exchangeRateSource = "", taxation = "";
                    double exchangeRate = 0, exciseTax = 0;
                    int unitRound = 0, totalRound = 0;
                    List<Customer> resultCustomer = new List<Customer>();

                    string originalSoErpPrefix = "", originalSoErpNo = "";
                    int originalCustomerId = -1;

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷訂單資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.SoErpPrefix, a.SoErpNo, a.CustomerId, a.ConfirmStatus
                                , ISNULL(b.BpmTransferStatus, 'N') BpmTransferStatus, a.TransferStatus
                                FROM SCM.SaleOrder a 
                                LEFT JOIN SCM.SoBpmInfo b ON a.SoId = b.SoId
                                WHERE a.SoId = @SoId
                                AND a.CompanyId = @CompanyId";
                        dynamicParameters.Add("SoId", SoId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("訂單資料錯誤!");

                        //string TransferStatus = result.First().TransferStatus;
                        //if (TransferStatus == "Y") throw new SystemException("訂單已拋轉ERP，無法修改!");

                        string confirmStatus = "";
                        foreach (var item in result)
                        {
                            if (CurrentCompany == 4 && item.BpmTransferStatus != "N" && item.BpmTransferStatus != "E")
                            {
                                throw new SystemException("此訂單目前拋轉BPM簽核中，無法更改!!");
                            }
                            originalSoErpPrefix = item.SoErpPrefix;
                            originalSoErpNo = item.SoErpNo;
                            originalCustomerId = Convert.ToInt32(item.CustomerId);
                            confirmStatus = item.ConfirmStatus;
                        }

                        #region //判斷訂單狀態
                        if (confirmStatus != "N") throw new SystemException("訂單不是未確認狀態，無法修改!");
                        #endregion
                        #endregion

                        #region //公司別資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var resultCompany = sqlConnection.Query(sql, dynamicParameters);
                        if (resultCompany.Count() <= 0) throw new SystemException("【公司別】資料錯誤!");

                        foreach (var item in resultCompany)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            companyNo = item.ErpNo;
                        }
                        #endregion

                        #region //判斷客戶資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 *
                                FROM SCM.Customer
                                WHERE CustomerId = @CustomerId
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("CustomerId", CustomerId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        resultCustomer = sqlConnection.Query<Customer>(sql, dynamicParameters).ToList();
                        if (resultCustomer.Count() <= 0) throw new SystemException("【客戶】資料錯誤!");
                        #endregion

                        #region //判斷部門資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Department
                                WHERE DepartmentId = @DepartmentId
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("DepartmentId", DepartmentId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var resultDepartment = sqlConnection.Query(sql, dynamicParameters);
                        if (resultDepartment.Count() <= 0) throw new SystemException("【部門】資料錯誤!");
                        #endregion

                        #region //判斷業務資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
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
                                AND a.DetailCode = @DetailCode
                                AND c.Authority > 0";
                        dynamicParameters.Add("UserId", SalesmenId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("Status", "A");
                        dynamicParameters.Add("FunctionCode", "SaleOrderManagement");
                        dynamicParameters.Add("DetailCode", "add");

                        var resultSalesmen = sqlConnection.Query(sql, dynamicParameters);
                        if (resultSalesmen.Count() <= 0) throw new SystemException("【業務】資料錯誤!");
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        #region //單據設定
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MQ004, a.MQ005, a.MQ006, a.MQ044
                                FROM CMSMQ a
                                WHERE a.COMPANY = @CompanyNo
                                AND a.MQ001 = @ErpPrefix";
                        dynamicParameters.Add("CompanyNo", companyNo);
                        dynamicParameters.Add("ErpPrefix", SoErpPrefix);

                        var resultDocSetting = sqlConnection.Query(sql, dynamicParameters);
                        if (resultDocSetting.Count() <= 0) throw new SystemException("ERP單據設定資料不存在!");

                        foreach (var item in resultDocSetting)
                        {
                            exchangeRateSource = item.MQ044; //匯率來源
                        }
                        #endregion

                        #region //目前匯率
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 MG003, MG004, MG005, MG006
                                FROM CMSMG
                                WHERE MG001 = @Currency
                                AND MG002 <= @DateNow
                                ORDER BY MG002 DESC";
                        dynamicParameters.Add("Currency", Currency);
                        dynamicParameters.Add("DateNow", DateTime.Now.ToString("yyyyMMdd"));

                        var resultExchangeRate = sqlConnection.Query(sql, dynamicParameters);
                        if (resultExchangeRate.Count() <= 0) throw new SystemException("ERP交易幣別匯率不存在!");

                        foreach (var item in resultExchangeRate)
                        {
                            switch (exchangeRateSource)
                            {
                                case "I": //銀行買進匯率
                                    exchangeRate = Convert.ToDouble(item.MG003);
                                    break;
                                case "O": //銀行賣出匯率
                                    exchangeRate = Convert.ToDouble(item.MG004);
                                    break;
                                case "E": //報關買進匯率
                                    exchangeRate = Convert.ToDouble(item.MG005);
                                    break;
                                case "W": //報關賣出匯率
                                    exchangeRate = Convert.ToDouble(item.MG006);
                                    break;
                            }
                        }
                        #endregion

                        #region //稅別碼設定
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT RTRIM(LTRIM(NN001)) NN001, RTRIM(LTRIM(NN002)) NN002, NN004, RTRIM(LTRIM(NN006)) NN006
                                FROM CMSNN
                                WHERE NN001 = @TaxNo";
                        dynamicParameters.Add("TaxNo", resultCustomer.Select(x => x.TaxNo).FirstOrDefault());

                        var resultTax = sqlConnection.Query(sql, dynamicParameters);
                        if (resultTax.Count() <= 0) throw new SystemException("ERP稅別碼設定不存在!");

                        foreach (var item in resultTax)
                        {
                            exciseTax = Convert.ToDouble(item.NN004); //營業稅率
                            taxation = item.NN006; //課稅別
                        }
                        #endregion

                        #region //交易幣別設定
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT MF003, MF004, MF005, MF006
                                FROM CMSMF
                                WHERE MF001 = @Currency";
                        dynamicParameters.Add("Currency", Currency);

                        var resultCurrencySetting = sqlConnection.Query(sql, dynamicParameters);
                        if (resultCurrencySetting.Count() <= 0) throw new SystemException("ERP交易幣別設定不存在!");

                        foreach (var item in resultCurrencySetting)
                        {
                            unitRound = Convert.ToInt32(item.MF003); //單價取位
                            totalRound = Convert.ToInt32(item.MF004); //金額取位
                        }
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //隨機取得單號資料
                        string SoErpNo = originalSoErpNo;

                        if (SoErpPrefix != originalSoErpPrefix)
                        {
                            //bool checkSoErpNo = true;
                            //while (checkSoErpNo)
                            //{
                            //    dynamicParameters = new DynamicParameters();
                            //    sql = @"SELECT TOP 1 1
                            //            FROM SCM.SaleOrder
                            //            WHERE SoErpPrefix = @SoErpPrefix
                            //            AND SoErpNo = @SoErpNo";
                            //    dynamicParameters.Add("SoErpPrefix", SoErpPrefix);
                            //    dynamicParameters.Add("SoErpNo", SoErpNo);

                            //    var resultSoErpNo = sqlConnection.Query(sql, dynamicParameters);
                            //    checkSoErpNo = resultSoErpNo.Count() > 0;

                            //    if (checkSoErpNo) SoErpNo = BaseHelper.RandomCode(11);
                            //}

                            throw new SystemException("【單別】不能修改!");
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.SaleOrder SET
                                DepartmentId = @DepartmentId,
                                SoErpPrefix = @SoErpPrefix,
                                SoErpNo = @SoErpNo,
                                SoDate = @SoDate,
                                DocDate = @DocDate,
                                SoRemark = @SoRemark,
                                CustomerId = @CustomerId,
                                SalesmenId = @SalesmenId,
                                CustomerAddressFirst = @CustomerAddressFirst,
                                CustomerAddressSecond = @CustomerAddressSecond,
                                CustomerPurchaseOrder = @CustomerPurchaseOrder,
                                Currency = @Currency,
                                ExchangeRate = @ExchangeRate,
                                TaxNo = @TaxNo,
                                Taxation = @Taxation,
                                BusinessTaxRate = @BusinessTaxRate,
                                ShipMethod = @ShipMethod,
                                TradeTerm = @TradeTerm,
                                PaymentTerm = @PaymentTerm,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SoId = @SoId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                DepartmentId = DepartmentId <= 0 ? (int?)null : DepartmentId,
                                SoErpPrefix,
                                SoErpNo,
                                SoDate = DocDate,
                                DocDate,
                                SoRemark,
                                CustomerId,
                                SalesmenId,
                                CustomerAddressFirst = resultCustomer.Select(x => x.RegisterAddressFirst).FirstOrDefault(),
                                CustomerAddressSecond = resultCustomer.Select(x => x.RegisterAddressSecond).FirstOrDefault(),
                                CustomerPurchaseOrder,
                                Currency = Currency,
                                ExchangeRate = exchangeRate,
                                TaxNo = resultCustomer.Select(x => x.TaxNo).FirstOrDefault(),
                                Taxation = taxation,
                                BusinessTaxRate = exciseTax,
                                ShipMethod = resultCustomer.Select(x => x.ShipMethod).FirstOrDefault(),
                                TradeTerm = resultCustomer.Select(x => x.TradeTerm).FirstOrDefault().Length <= 0 ? "1" : resultCustomer.Select(x => x.TradeTerm).FirstOrDefault(), //交易條件
                                PaymentTerm = resultCustomer.Select(x => x.PaymentTerm).FirstOrDefault(),
                                LastModifiedDate,
                                LastModifiedBy,
                                SoId
                            });

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);


                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.SoDetail SET
                                TaxNo = @TaxNo,
                                BusinessTaxRate = @BusinessTaxRate,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SoId = @SoId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TaxNo = resultCustomer.Select(x => x.TaxNo).FirstOrDefault(),
                                BusinessTaxRate = exciseTax,
                                LastModifiedDate,
                                LastModifiedBy,
                                SoId
                            });

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                        #region //單頭資料更新
                        double totalQty = 0, amount = 0, taxAmount = 0;

                        #region //所有單身資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT SoQty, SoPriceQty, FreebieQty, SpareQty, UnitPrice, Amount
                                FROM SCM.SoDetail
                                WHERE SoId = @SoId
                                ORDER BY SoSequence";
                        dynamicParameters.Add("SoId", SoId);

                        var resultSoDetail = sqlConnection.Query(sql, dynamicParameters);
                        //if (resultSoDetail.Count() <= 0) throw new SystemException("無單身資料!");

                        foreach (var item in resultSoDetail)
                        {
                            totalQty += Convert.ToDouble(item.SoQty) + Convert.ToDouble(item.FreebieQty) + Convert.ToDouble(item.SpareQty);
                            amount += Math.Round(Math.Round(Convert.ToDouble(item.UnitPrice), unitRound, MidpointRounding.AwayFromZero) * Convert.ToDouble(item.SoPriceQty), totalRound, MidpointRounding.AwayFromZero);
                        }

                        #region //計算數量與金額
                        switch (taxation)
                        {
                            case "1":
                                taxAmount = amount - Math.Round(amount / (1 + exciseTax), totalRound, MidpointRounding.AwayFromZero);
                                amount = Math.Round(amount / (1 + exciseTax), totalRound, MidpointRounding.AwayFromZero);
                                break;
                            case "2":
                                taxAmount = Math.Round(amount * exciseTax, totalRound, MidpointRounding.AwayFromZero);
                                break;
                            case "3":
                                break;
                            case "4":
                                break;
                            case "9":
                                break;
                        }
                        #endregion
                        #endregion

                        #region //單頭更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.SaleOrder SET
                                TotalQty = @TotalQty,
                                Amount = @Amount,
                                TaxAmount = @TaxAmount,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SoId = @SoId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TotalQty = totalQty,
                                Amount = amount,
                                TaxAmount = taxAmount,
                                LastModifiedDate,
                                LastModifiedBy,
                                SoId
                            });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

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

        #region //UpdateSoDetail -- 訂單單身資料更新 -- Zoey 2022.07.15
        public string UpdateSoDetail(int SoDetailId, int SoDetailTempId, string SoSequence, int MtlItemId, string SoMtlItemName
            , int InventoryId, double SoQty, double UnitPrice, double Amount
            , string ProductType, float FreebieOrSpareQty, string PromiseDate
            , string Project, string SoDetailRemark, string QuotationErp)
        {
            try
            {
                if (SoSequence.Length <= 0) throw new SystemException("【流水號】不能為空!");
                if (SoSequence.Length > 4) throw new SystemException("【流水號】長度錯誤!");
                if (InventoryId <= 0) throw new SystemException("【庫別】不能為空!");
                //if (SoQty <= 0) throw new SystemException("【訂單數量】不能為空!");
                if (UnitPrice < 0) throw new SystemException("【單價】不能為空!");
                if (Amount < 0) throw new SystemException("【金額】不能為空!");
                if (ProductType.Length <= 0) throw new SystemException("【類型】不能為空!");
                if (FreebieOrSpareQty < 0) throw new SystemException("【贈/備品數】不能為空!");
                if (PromiseDate.Length <= 0) throw new SystemException("【預計出貨日】不能為空!");
                if (!DateTime.TryParse(PromiseDate, out DateTime tempDateTime)) throw new SystemException("【預計出貨日】格式錯誤!");
                if (Project.Length > 20) throw new SystemException("【專案代號】長度錯誤!");
                if (SoDetailRemark.Length > 255) throw new SystemException("【備註】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    string companyNo = "", currency = "", taxNo = "", soMtlItemSpec = "", taxation = "";
                    int soId = -1, uomId = -1, unitRound = 0, totalRound = 0;
                    double exciseTax = 0;

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //公司別資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var resultCompany = sqlConnection.Query(sql, dynamicParameters);
                        if (resultCompany.Count() <= 0) throw new SystemException("【公司別】資料錯誤!");

                        foreach (var item in resultCompany)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            companyNo = item.ErpNo;
                        }
                        #endregion

                        #region //判斷訂單單身資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 b.SoId, b.Currency, b.TaxNo, b.ConfirmStatus
                                , ISNULL(c.BpmTransferStatus, 'N') BpmTransferStatus, b.TransferStatus
                                FROM SCM.SoDetail a
                                INNER JOIN SCM.SaleOrder b ON a.SoId = b.SoId
                                LEFT JOIN SCM.SoBpmInfo c ON a.SoId = c.SoId
                                WHERE a.SoDetailId = @SoDetailId
                                AND b.CompanyId = @CompanyId";
                        dynamicParameters.Add("SoDetailId", SoDetailId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("訂單資料錯誤!");

                        //string TransferStatus = result.First().TransferStatus;
                        //if (TransferStatus == "Y") throw new SystemException("訂單已拋轉ERP，無法修改!");

                        string confirmStatus = "";
                        foreach (var item in result)
                        {
                            if (CurrentCompany == 4 && item.BpmTransferStatus != "N" && item.BpmTransferStatus != "E")
                            {
                                throw new SystemException("此訂單目前拋轉BPM簽核中，無法更改!!");
                            }

                            soId = Convert.ToInt32(item.SoId);
                            currency = item.Currency;
                            taxNo = item.TaxNo;
                            confirmStatus = item.ConfirmStatus;
                        }

                        #region //判斷訂單狀態
                        if (confirmStatus != "N") throw new SystemException("訂單不是未確認狀態，無法更新!");
                        #endregion
                        #endregion

                        #region //判斷序號是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.SoDetail
                                WHERE SoId = @SoId
                                AND SoSequence = @SoSequence
                                AND SoDetailId != @SoDetailId";
                        dynamicParameters.Add("SoId", soId);
                        dynamicParameters.Add("SoSequence", SoSequence);
                        dynamicParameters.Add("SoDetailId", SoDetailId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("流水號重複，請重新取號!");
                        #endregion

                        #region //判斷品號資料是否正確
                        if (MtlItemId > 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 MtlItemName, MtlItemSpec, InventoryId, ISNULL(SaleUomId, InventoryUomId) SaleUomId
                                    FROM PDM.MtlItem
                                    WHERE MtlItemId = @MtlItemId
                                    AND CompanyId = @CompanyId";
                            dynamicParameters.Add("MtlItemId", MtlItemId);
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            var resultMtlItem = sqlConnection.Query(sql, dynamicParameters);
                            if (resultMtlItem.Count() <= 0) throw new SystemException("品號資料錯誤!");

                            foreach (var item in resultMtlItem)
                            {
                                SoMtlItemName = item.MtlItemName;
                                soMtlItemSpec = item.MtlItemSpec;
                                uomId = Convert.ToInt32(item.SaleUomId);
                            }
                        }
                        else
                        {
                            if (SoMtlItemName.Length <= 0) throw new SystemException("【品名】不能為空!");
                            if (SoMtlItemName.Length > 120) throw new SystemException("【品名】長度錯誤!");
                        }
                        #endregion

                        #region //判斷庫別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.Inventory
                                WHERE InventoryId = @InventoryId
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("InventoryId", InventoryId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var resultInventory = sqlConnection.Query(sql, dynamicParameters);
                        if (resultInventory.Count() <= 0) throw new SystemException("庫別資料錯誤!");
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        #region //交易幣別設定
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT MF003, MF004, MF005, MF006
                                FROM CMSMF
                                WHERE MF001 = @Currency";
                        dynamicParameters.Add("Currency", currency);

                        var resultCurrencySetting = sqlConnection.Query(sql, dynamicParameters);
                        if (resultCurrencySetting.Count() <= 0) throw new SystemException("ERP交易幣別設定不存在!");

                        foreach (var item in resultCurrencySetting)
                        {
                            unitRound = Convert.ToInt32(item.MF003); //單價取位
                            totalRound = Convert.ToInt32(item.MF004); //金額取位
                        }
                        #endregion

                        #region //稅別碼設定
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT RTRIM(LTRIM(NN001)) NN001, RTRIM(LTRIM(NN002)) NN002, NN004, RTRIM(LTRIM(NN006)) NN006
                                FROM CMSNN
                                WHERE NN001 = @TaxNo";
                        dynamicParameters.Add("TaxNo", taxNo);

                        var resultTax = sqlConnection.Query(sql, dynamicParameters);
                        if (resultTax.Count() <= 0) throw new SystemException("ERP稅別碼設定不存在!");

                        foreach (var item in resultTax)
                        {
                            exciseTax = Convert.ToDouble(item.NN004); //營業稅率
                            taxation = item.NN006; //課稅別
                        }
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;

                        #region //單身更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.SoDetail SET
                                SoSequence = @SoSequence,
                                MtlItemId = @MtlItemId,
                                SoMtlItemName = @SoMtlItemName,
                                SoMtlItemSpec = @SoMtlItemSpec,
                                InventoryId = @InventoryId,
                                UomId = @UomId,
                                SoQty = @SoQty,
                                SiQty = @SiQty,
                                ProductType = @ProductType,
                                FreebieQty = @FreebieQty,
                                FreebieSiQty = @FreebieSiQty,
                                SpareQty = @SpareQty,
                                SpareSiQty = @SpareSiQty,
                                UnitPrice = @UnitPrice,
                                Amount = @Amount,
                                PromiseDate = @PromiseDate,
                                PcPromiseDate = @PcPromiseDate,
                                Project = @Project,
                                CustomerMtlItemNo = @CustomerMtlItemNo,
                                SoDetailRemark = @SoDetailRemark,
                                SoPriceQty = @SoPriceQty,
                                SoPriceUomId = @SoPriceUomId,
                                TaxNo = @TaxNo,
                                BusinessTaxRate = @BusinessTaxRate,
                                DiscountRate = @DiscountRate,
                                DiscountAmount = @DiscountAmount,
                                ConfirmStatus = @ConfirmStatus,
                                ClosureStatus = @ClosureStatus,
                                QuotationErp = @QuotationErp,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SoDetailId = @SoDetailId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SoSequence,
                                MtlItemId = MtlItemId <= 0 ? (int?)null : MtlItemId,
                                SoMtlItemName,
                                SoMtlItemSpec = soMtlItemSpec,
                                InventoryId,
                                UomId = uomId <= 0 ? (int?)null : uomId,
                                SoQty,
                                SiQty = 0, //已交數量
                                ProductType,
                                FreebieQty = ProductType == "1" ? FreebieOrSpareQty : 0, //贈品數量
                                FreebieSiQty = 0, //贈品已交數量
                                SpareQty = ProductType == "2" ? FreebieOrSpareQty : 0,
                                SpareSiQty = 0, //備品已交數量
                                UnitPrice = Math.Round(UnitPrice, unitRound),
                                Amount = Math.Round(Math.Round(UnitPrice, unitRound) * SoQty, totalRound),
                                PromiseDate,
                                PcPromiseDate = PromiseDate, //排定交貨日
                                Project,
                                CustomerMtlItemNo = "", //客戶品號
                                SoDetailRemark,
                                SoPriceQty = SoQty, //計價數量
                                SoPriceUomId = uomId <= 0 ? (int?)null : uomId, //計價單位
                                TaxNo = taxNo, //稅別碼
                                BusinessTaxRate = exciseTax, //營業稅率
                                DiscountRate = 1, //折扣率
                                DiscountAmount = 0, //折扣金額
                                ConfirmStatus = "N", //確認碼
                                ClosureStatus = "N", //結案碼
                                QuotationErp,
                                LastModifiedDate,
                                LastModifiedBy,
                                SoDetailId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += insertResult.Count();
                        #endregion

                        if (SoDetailTempId > 0)
                        {
                            #region //判斷採購單匯入資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM SCM.SoDetailTemp
                                    WHERE SoDetailTempId = @SoDetailTempId";
                            dynamicParameters.Add("SoDetailTempId", SoDetailTempId);
                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【採購單】匯入錯誤!");
                            #endregion

                            #region //更新綁定的暫存檔
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.SoDetailTemp SET 
                                    SoDetailId = @SoDetailId,
                                    Status = @Status,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE SoDetailTempId = @SoDetailTempId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    SoDetailId,
                                    SoDetailTempId,
                                    Status = "Y", //綁定訂單身
                                    LastModifiedDate,
                                    LastModifiedBy
                                });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }
                        else
                        {
                            #region //解除訂單身暫存綁定
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.SoDetailTemp SET 
                                    SoDetailId = NULL,
                                    Status = @Status,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE SoDetailId = @SoDetailId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    SoDetailId,
                                    Status = "N", //解除訂單身綁定
                                    LastModifiedDate,
                                    LastModifiedBy
                                });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }

                        #region //單頭資料更新
                        double totalQty = 0, amount = 0, taxAmount = 0;

                        #region //所有單身資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT SoQty, SoPriceQty, FreebieQty, SpareQty, UnitPrice, Amount
                                FROM SCM.SoDetail
                                WHERE SoId = @SoId
                                ORDER BY SoSequence";
                        dynamicParameters.Add("SoId", soId);

                        var resultSoDetail = sqlConnection.Query(sql, dynamicParameters);
                        if (resultSoDetail.Count() <= 0) throw new SystemException("無單身資料!");

                        foreach (var item in resultSoDetail)
                        {
                            totalQty += Convert.ToDouble(item.SoQty) + Convert.ToDouble(item.FreebieQty) + Convert.ToDouble(item.SpareQty);
                            amount += Math.Round(Math.Round(Convert.ToDouble(item.UnitPrice), unitRound) * Convert.ToDouble(item.SoPriceQty), totalRound);
                        }

                        #region //計算數量與金額
                        switch (taxation)
                        {
                            case "1":
                                taxAmount = amount - Math.Round(amount / (1 + exciseTax), totalRound);
                                amount = Math.Round(amount / (1 + exciseTax), totalRound);
                                break;
                            case "2":
                                taxAmount = Math.Round(amount * exciseTax, totalRound);
                                break;
                            case "3":
                                break;
                            case "4":
                                break;
                            case "9":
                                break;
                        }
                        #endregion
                        #endregion

                        #region //單頭更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.SaleOrder SET
                                TotalQty = @TotalQty,
                                Amount = @Amount,
                                TaxAmount = @TaxAmount,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SoId = @SoId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TotalQty = totalQty,
                                Amount = amount,
                                TaxAmount = taxAmount,
                                LastModifiedDate,
                                LastModifiedBy,
                                SoId = soId
                            });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion
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

        #region //UpdateSaleOrderImport -- 訂單拋轉ERP -- Ben Ma 2023.06.14
        public string UpdateSaleOrderImport(int SoId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0, totalDetail = 0;
                    string confirmStatus = "", transferStatus = "", companyNo = "", soErpPrefix = "", soErpNo = "", docDate = "", departmentNo = "", userNo = "", userName = "";
                    List<COPTC> cOPTCs = new List<COPTC>();
                    List<COPTD> cOPTDs = new List<COPTD>();

                    string dateNow = DateTime.Now.ToString("yyyyMMdd");
                    string timeNow = DateTime.Now.ToString("HH:mm:ss");

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷訂單資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 SoErpPrefix, SoErpNo, ConfirmStatus, TransferStatus
                                FROM SCM.SaleOrder
                                WHERE SoId = @SoId
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("SoId", SoId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("訂單資料錯誤!");

                        foreach (var item in result)
                        {
                            soErpPrefix = item.SoErpPrefix;
                            soErpNo = item.SoErpNo;
                            confirmStatus = item.ConfirmStatus;
                            transferStatus = item.TransferStatus;
                        }

                        #region //判斷訂單狀態
                        if (confirmStatus != "N") throw new SystemException("訂單不是未確認狀態，無法拋轉!");
                        #endregion
                        #endregion

                        #region //公司別資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var resultCompany = sqlConnection.Query(sql, dynamicParameters);
                        if (resultCompany.Count() <= 0) throw new SystemException("【公司別】資料錯誤!");

                        foreach (var item in resultCompany)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            companyNo = item.ErpNo;
                        }
                        #endregion

                        #region //使用者資料
                        // 目前僅晶彩使用BPM
                        string DetailCode = "";
                        if (CurrentCompany == 4)
                        {
                            DetailCode = "bpm-transfer";
                        }
                        else
                        {
                            DetailCode = "import";
                        }

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 d.UserNo, d.UserName, d.DepartmentNo
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
                                OUTER APPLY (
                                    SELECT da.UserNo, da.UserName, db.DepartmentNo
                                    FROM BAS.[User] da
                                    INNER JOIN BAS.Department db ON da.DepartmentId = db.DepartmentId
                                    WHERE da.UserId = @UserId
                                ) d
                                WHERE a.[Status] = @Status
                                AND b.[Status] = @Status
                                AND b.FunctionCode = @FunctionCode
                                AND a.DetailCode = @DetailCode
                                AND c.Authority > 0";
                        dynamicParameters.Add("UserId", CurrentUser);
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("Status", "A");
                        dynamicParameters.Add("FunctionCode", "SaleOrderManagement");
                        dynamicParameters.Add("DetailCode", DetailCode);

                        var resultUser = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUser.Count() <= 0) throw new SystemException("使用者資料錯誤!");

                        foreach (var item in resultUser)
                        {
                            userNo = item.UserNo;
                            userName = item.UserName;
                            departmentNo = item.DepartmentNo;
                        }
                        #endregion

                        #region //單身筆數
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT COUNT(1) TotalDetail
                                FROM SCM.SoDetail
                                WHERE SoId = @SoId";
                        dynamicParameters.Add("SoId", SoId);

                        var resultDetailExist = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in resultDetailExist)
                        {
                            totalDetail = Convert.ToInt32(item.TotalDetail);
                            if (totalDetail <= 0) throw new SystemException("請先建立單身!");
                        }
                        #endregion

                        #region //判斷單身是否皆有品號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.SoDetail
                                WHERE SoId = @SoId
                                AND MtlItemId IS NULL";
                        dynamicParameters.Add("SoId", SoId);

                        var resultDetailMtlItemExist = sqlConnection.Query(sql, dynamicParameters);
                        if (resultDetailMtlItemExist.Count() > 0) throw new SystemException("部分單身無綁定品號!");
                        #endregion

                        #region //訂單單頭
                        sql = @"SELECT a.SoId, c.DepartmentNo TC005, a.SoErpPrefix TC001, a.SoErpNo TC002
                                , a.Version TC069, FORMAT(a.SoDate, 'yyyyMMdd') TC003, FORMAT(a.DocDate, 'yyyyMMdd') TC039
                                , a.SoRemark TC015, b.CustomerNo TC004, b.CustomerNo TC032, d.UserNo TC006
                                , a.CustomerAddressFirst TC010, a.CustomerAddressSecond TC011
                                , a.CustomerPurchaseOrder TC012, a.DepositPartial TC070, a.DepositRate TC045
                                , a.Currency TC008, a.ExchangeRate TC009, a.TaxNo TC078, a.Taxation TC016
                                , a.BusinessTaxRate TC041, a.DetailMultiTax TC091, a.TotalQty TC031
                                , a.Amount TC029, a.TaxAmount TC030, a.ShipMethod TC019
                                , a.TradeTerm TC068, a.PaymentTerm TC042, a.PriceTerm TC013
                                , a.TransferStatus, a.ConfirmStatus TC027
                                , b.CustomerName TC053, b.CustomerName TC065, b.CustomerEnglishName TC071
                                , b.Country TC081, b.Region TC080, b.Route TC083, b.Contact TC018
                                , b.TelNoFirst TC066, b.FaxNo TC067, b.InvoiceAddressFirst TC063
                                , b.InvoiceAddressSecond TC064, b.PaymentBankFirst TC036
                                FROM SCM.SaleOrder a
                                INNER JOIN SCM.Customer b ON b.CustomerId = a.CustomerId
                                LEFT JOIN BAS.Department c ON c.DepartmentId = a.DepartmentId
                                LEFT JOIN BAS.[User] d ON d.UserId = a.SalesmenId
                                LEFT JOIN BAS.[User] e ON e.UserId = a.ConfirmUserId
                                WHERE a.SoId = @SoId";
                        dynamicParameters.Add("SoId", SoId);

                        cOPTCs = sqlConnection.Query<COPTC>(sql, dynamicParameters).ToList();
                        if (cOPTCs.Count() <= 0) throw new SystemException("訂單單頭資料錯誤!");
                        #endregion

                        #region //訂單單身
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.SoDetailId, a.SoId, b.SoErpPrefix TD001, b.SoErpNo TD002, a.SoSequence TD003
                                , c.MtlItemNo TD004, a.SoMtlItemName TD005, a.SoMtlItemSpec TD006,d.InventoryNo TD007
                                , e.UomNo TD010, a.SoQty TD008, a.SiQty TD009
                                , a.ProductType TD049, a.FreebieQty TD024, a.FreebieSiQty TD025, a.SpareQty TD050, a.SpareSiQty TD051
                                , a.UnitPrice TD011, a.Amount TD012, FORMAT(a.PromiseDate, 'yyyyMMdd') TD013
                                , FORMAT(a.PcPromiseDate, 'yyyyMMdd') TD048, a.Project TD027, a.CustomerMtlItemNo TD014
                                , a.SoDetailRemark TD020, a.SoPriceQty TD076
                                , f.UomNo TD077, a.TaxNo TD079, a.BusinessTaxRate TD070, a.DiscountRate TD026
                                , a.DiscountAmount TD080, a.ConfirmStatus TD021, a.ClosureStatus TD016,a.QuotationErp
                                ,CASE 
                                    WHEN ISNULL(a.QuotationErp, '') != '' THEN '1'
                                    ELSE ''
                                END AS TD045
                                ,CASE 
                                    WHEN ISNULL(a.QuotationErp, '') != '' AND CHARINDEX('-', ISNULL(a.QuotationErp, '')) > 0 
                                    THEN LEFT(a.QuotationErp, CHARINDEX('-', a.QuotationErp) - 1)
                                    ELSE ''
                                END AS TD017
                                ,CASE 
                                    WHEN ISNULL(a.QuotationErp, '') != '' AND CHARINDEX('-', ISNULL(a.QuotationErp, '')) > 0 
                                    THEN SUBSTRING(
                                        a.QuotationErp, 
                                        CHARINDEX('-', a.QuotationErp) + 1, 
                                        CHARINDEX('-', a.QuotationErp, CHARINDEX('-', a.QuotationErp) + 1) - CHARINDEX('-', a.QuotationErp) - 1
                                    )
                                    ELSE ''
                                END AS TD018
                                ,CASE 
                                    WHEN ISNULL(a.QuotationErp, '') != '' AND CHARINDEX('-', ISNULL(a.QuotationErp, '')) > 0 
                                    THEN RIGHT(a.QuotationErp, LEN(a.QuotationErp) - CHARINDEX('-', a.QuotationErp, CHARINDEX('-', a.QuotationErp) + 1))
                                    ELSE ''
                                END AS TD019
                                FROM SCM.SoDetail a
                                INNER JOIN SCM.SaleOrder b ON b.SoId = a.SoId
                                INNER JOIN PDM.MtlItem c ON c.MtlItemId = a.MtlItemId
                                INNER JOIN SCM.Inventory d ON d.InventoryId = a.InventoryId
                                INNER JOIN PDM.UnitOfMeasure e ON e.UomId = a.UomId
                                INNER JOIN PDM.UnitOfMeasure f ON f.UomId = a.SoPriceUomId
                                WHERE a.SoId = @SoId
                                AND a.MtlItemId IS NOT NULL
                                ORDER BY a.SoId, a.SoSequence";
                        dynamicParameters.Add("SoId", SoId);

                        cOPTDs = sqlConnection.Query<COPTD>(sql, dynamicParameters).ToList();
                        if (cOPTDs.Count() <= 0) throw new SystemException("訂單單身資料錯誤!");
                        if (cOPTDs.Count() != totalDetail) throw new SystemException("訂單單身數量錯誤!");
                        #endregion

                        #region //判斷量產訂單是否上傳驗證文件(紘立先導入)
                        if (companyNo == "ETG")
                        {
                            if (soErpPrefix == "2201") //量產訂單卡控
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM (
                                            SELECT dbo.MassProductionReviewDocVerify(b.MtlItemNo) DocVerify
                                            FROM SCM.SoDetail a
                                            LEFT JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
                                            WHERE a.SoId = @SoId
                                        ) doc
                                        WHERE doc.DocVerify != 'F'";
                                dynamicParameters.Add("SoId", SoId);

                                var resultDocVerify = sqlConnection.Query(sql, dynamicParameters);
                                if (resultDocVerify.Count() > 0) throw new SystemException("部分品號量產審查文件不齊全!");
                            }
                        }
                        #endregion

                        #region //基本資料設定
                        #region //COPTC
                        cOPTCs
                            .ToList()
                            .ForEach(x =>
                            {
                                x.COMPANY = companyNo;
                                x.CREATOR = userNo;
                                x.USR_GROUP = "";
                                x.CREATE_DATE = dateNow;
                                x.MODIFIER = "";
                                x.MODI_DATE = "";
                                x.FLAG = "1";
                                x.CREATE_TIME = timeNow;
                                x.CREATE_AP = userNo + "PC";
                                x.CREATE_PRID = "BM";
                                x.MODI_TIME = "";
                                x.MODI_AP = "";
                                x.MODI_PRID = "";
                            });
                        #endregion

                        #region //COPTD
                        cOPTDs
                            .ToList()
                            .ForEach(x =>
                            {
                                x.COMPANY = companyNo;
                                x.CREATOR = userNo;
                                x.USR_GROUP = "";
                                x.CREATE_DATE = dateNow;
                                x.MODIFIER = "";
                                x.MODI_DATE = "";
                                x.FLAG = "1";
                                x.CREATE_TIME = timeNow;
                                x.CREATE_AP = userNo + "PC";
                                x.CREATE_PRID = "BM";
                                x.MODI_TIME = "";
                                x.MODI_AP = "";
                                x.MODI_PRID = "";
                            });
                        #endregion
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        int currentNum = 0, yearLength = 0, lineLength = 0;
                        string encode = "", paymentTerm = "", factory = "";
                        DateTime referenceTime = default(DateTime);

                        #region //判斷ERP使用者是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 MF005
                                FROM ADMMF
                                WHERE COMPANY = @CompanyNo
                                AND MF001 = @User";
                        dynamicParameters.Add("CompanyNo", companyNo);
                        dynamicParameters.Add("User", userNo);

                        var resultUserExist = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUserExist.Count() <= 0) throw new SystemException("ERP使用者不存在!");

                        string adminUser = "N";
                        foreach (var item in resultUserExist)
                        {
                            adminUser = item.MF005; //超級使用者
                        }
                        #endregion

                        #region //判斷ERP使用者是否有權限
                        //晶彩不用卡控要ERP訂單確認權限
                        //if (companyNo!= "DGJMO") {
                        //    if (adminUser != "Y")
                        //    {
                        //        dynamicParameters = new DynamicParameters();
                        //        sql = @"SELECT TOP 1 MG006
                        //                FROM ADMMG
                        //                WHERE COMPANY = @CompanyNo
                        //                AND MG001 = @User
                        //                AND MG002 = @Function
                        //                AND MG004 = 'Y'
                        //                AND MG006 LIKE @Auth";
                        //        dynamicParameters.Add("CompanyNo", companyNo);
                        //        dynamicParameters.Add("User", userNo);
                        //        dynamicParameters.Add("Function", "COPI06");
                        //        dynamicParameters.Add("Auth", "_Y__________"); //修改/查詢/新增

                        //        var resultUserAuthExist = sqlConnection.Query(sql, dynamicParameters);
                        //        if (resultUserAuthExist.Count() <= 0) throw new SystemException("ERP使用者無權限!");
                        //    }
                        //}
                        #endregion

                        #region //判斷品號是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT COUNT(1) TotalMtlItems
                                FROM INVMB
                                WHERE RTRIM(LTRIM(MB001)) IN @MtlItems";
                        dynamicParameters.Add("MtlItems", cOPTDs.Select(x => x.TD004).Distinct().ToArray());

                        var resultMtlItem = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in resultMtlItem)
                        {
                            if (Convert.ToInt32(item.TotalMtlItems) != cOPTDs.Select(x => x.TD004).Distinct().Count()) throw new SystemException("部分ERP品號不存在!");
                        }
                        #endregion

                        #region //訂單拋轉
                        docDate = cOPTCs.Select(x => x.TC039).FirstOrDefault();
                        referenceTime = DateTime.ParseExact(docDate, "yyyyMMdd", CultureInfo.InvariantCulture);

                        #region //單據設定
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MQ004, a.MQ005, a.MQ006, a.MQ044
                                FROM CMSMQ a
                                WHERE a.COMPANY = @CompanyNo
                                AND a.MQ001 = @ErpPrefix";
                        dynamicParameters.Add("CompanyNo", companyNo);
                        dynamicParameters.Add("ErpPrefix", soErpPrefix);

                        var resultDocSetting = sqlConnection.Query(sql, dynamicParameters);
                        if (resultDocSetting.Count() <= 0) throw new SystemException("ERP單據設定資料不存在!");

                        foreach (var item in resultDocSetting)
                        {
                            encode = item.MQ004; //編碼方式
                            yearLength = Convert.ToInt32(item.MQ005); //年碼數
                            lineLength = Convert.ToInt32(item.MQ006); //流水號碼數
                        }
                        #endregion

                        #region //單號取號
                        if (transferStatus == "N")
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT CONVERT(int, RIGHT(ISNULL(MAX(RTRIM(LTRIM(TC002))), '" + new string('0', lineLength) + @"'), " + lineLength + @")) + 1 CurrentNum
                                    FROM COPTC
                                    WHERE TC001 = @ErpPrefix";
                            dynamicParameters.Add("ErpPrefix", soErpPrefix);

                            #region //編碼方式
                            string dateFormat = "";
                            switch (encode)
                            {
                                case "1": //日編
                                    dateFormat = new string('y', yearLength) + "MMdd";
                                    sql += @" AND RTRIM(LTRIM(TC002)) LIKE @ReferenceTime + '" + new string('_', lineLength) + @"'";
                                    dynamicParameters.Add("ReferenceTime", referenceTime.ToString(dateFormat));
                                    soErpNo = referenceTime.ToString(dateFormat);
                                    break;
                                case "2": //月編
                                    dateFormat = new string('y', yearLength) + "MM";
                                    sql += @" AND RTRIM(LTRIM(TC002)) LIKE @ReferenceTime + '" + new string('_', lineLength) + @"'";
                                    dynamicParameters.Add("ReferenceTime", referenceTime.ToString(dateFormat));
                                    soErpNo = referenceTime.ToString(dateFormat);
                                    break;
                                case "3": //流水號
                                    break;
                                case "4": //手動編號
                                    break;
                                default:
                                    throw new SystemException("編碼方式錯誤!");
                            }
                            #endregion

                            currentNum = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).CurrentNum;
                            soErpNo += string.Format("{0:" + new string('0', lineLength) + "}", currentNum);
                        }
                        #endregion

                        #region //付款條件檔
                        sql = @"SELECT RTRIM(LTRIM(NA002)) NA002, RTRIM(LTRIM(NA003)) NA003
                                FROM CMSNA
                                WHERE 1=1
                                AND NA001 = '2'
                                AND NA002 = @PaymentTerm";
                        dynamicParameters.Add("PaymentTerm", cOPTCs.Select(x => x.TC042).FirstOrDefault());

                        var resultPaymentTerm = sqlConnection.Query(sql, dynamicParameters);
                        if (resultPaymentTerm.Count() <= 0) throw new SystemException("ERP付款條件不存在!");

                        foreach (var item in resultPaymentTerm)
                        {
                            paymentTerm = item.NA003; //付款條件名稱
                        }
                        #endregion

                        #region //廠別資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 MB001
                                FROM CMSMB
                                WHERE COMPANY = @COMPANY";
                        dynamicParameters.Add("COMPANY", companyNo);

                        var resultFactory = sqlConnection.Query(sql, dynamicParameters);
                        if (resultFactory.Count() <= 0) throw new SystemException("ERP廠別資料不存在!");

                        foreach (var item in resultFactory)
                        {
                            factory = item.MB001; //廠別
                        }
                        #endregion

                        switch (transferStatus)
                        {
                            case "N": //未拋轉新增
                                #region //COPTC
                                #region //判斷單號是否重複
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM COPTC
                                        WHERE TC001 = @ErpPrefix
                                        AND TC002 = @ErpNo";
                                dynamicParameters.Add("ErpPrefix", soErpPrefix);
                                dynamicParameters.Add("ErpNo", soErpNo);

                                var resultRepeatExist = sqlConnection.Query(sql, dynamicParameters);
                                if (resultRepeatExist.Count() > 0) throw new SystemException("【訂單單號】重複，請重新取號!");
                                #endregion

                                cOPTCs
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.TC001 = soErpPrefix; //訂單單別
                                        x.TC002 = soErpNo; //訂單單號
                                        x.TC007 = factory; //出貨廠別
                                        x.TC014 = paymentTerm; //付款條件
                                        x.TC017 = ""; //L/CNO.
                                        x.TC020 = ""; //起始港口
                                        x.TC021 = ""; //目的港口
                                        x.TC022 = ""; //代理商
                                        x.TC023 = ""; //報關行
                                        x.TC024 = ""; //驗貨公司
                                        x.TC025 = ""; //運輸公司
                                        x.TC026 = 0; //佣金比率
                                        x.TC027 = "N"; //確認碼
                                        x.TC028 = 0; //列印次數
                                        x.TC033 = ""; //NOTIFY
                                        x.TC034 = ""; //嘜頭代號
                                        x.TC035 = ""; //目的地
                                        x.TC037 = ""; //INVOICE備註
                                        x.TC038 = ""; //PACKING-LIST備註
                                        x.TC040 = ""; //確認者
                                        x.TC043 = 0; //總毛重(Kg)
                                        x.TC044 = 0; //總材積(CUFT)
                                        x.TC046 = 0; //總包裝數量
                                        x.TC047 = ""; //押匯銀行
                                        x.TC048 = "N"; //簽核狀態碼
                                        x.TC049 = ""; //流程代號(多角貿易)
                                        x.TC050 = "N"; //拋轉狀態
                                        x.TC051 = ""; //下游廠商
                                        x.TC052 = 0; //傳送次數
                                        x.TC054 = ""; //正嘜
                                        x.TC055 = ""; //側嘜
                                        x.TC056 = "1"; //材積單位
                                        x.TC057 = "N"; //EBC確認碼
                                        x.TC058 = ""; //EBC訂單號碼
                                        x.TC059 = ""; //EBC訂單版次
                                        x.TC060 = "N"; //匯至EBC
                                        x.TC061 = ""; //正嘜文管代號
                                        x.TC062 = ""; //側嘜文管代號
                                        x.TC072 = 0; //預留欄位
                                        x.TC073 = 0; //收入遞延天數
                                        x.TC074 = ""; //預留欄位
                                        x.TC075 = ""; //預留欄位
                                        x.TC076 = ""; //預留欄位
                                        x.TC077 = "N"; //不控管信用額度
                                        x.TC079 = ""; //通路別
                                        x.TC082 = ""; //型態別
                                        x.TC084 = ""; //其他別
                                        x.TC085 = ""; //出口港
                                        x.TC086 = ""; //經過港口
                                        x.TC087 = ""; //目的港口
                                        x.TC088 = ""; //最上游客戶
                                        x.TC089 = ""; //最上游交易幣別
                                        x.TC090 = ""; //最上游稅別碼
                                        x.TC092 = ""; //來源
                                        x.TC093 = ""; //送貨國家
                                        x.UDF01 = "";
                                        x.UDF02 = "";
                                        x.UDF03 = "";
                                        x.UDF04 = "";
                                        x.UDF05 = "";
                                        x.UDF06 = 0;
                                        x.UDF07 = 0;
                                        x.UDF08 = 0;
                                        x.UDF09 = 0;
                                        x.UDF10 = 0;
                                    });

                                sql = @"INSERT INTO COPTC (COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE
                                        , FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID
                                        , TC001, TC002, TC003, TC004, TC005, TC006, TC007, TC008, TC009, TC010
                                        , TC011, TC012, TC013, TC014, TC015, TC016, TC017, TC018, TC019, TC020
                                        , TC021, TC022, TC023, TC024, TC025, TC026, TC027, TC028, TC029, TC030
                                        , TC031, TC032, TC033, TC034, TC035, TC036, TC037, TC038, TC039, TC040
                                        , TC041, TC042, TC043, TC044, TC045, TC046, TC047, TC048, TC049, TC050
                                        , TC051, TC052, TC053, TC054, TC055, TC056, TC057, TC058, TC059, TC060
                                        , TC061, TC062, TC063, TC064, TC065, TC066, TC067, TC068, TC069, TC070
                                        , TC071, TC072, TC073, TC074, TC075, TC076, TC077, TC078, TC079, TC080
                                        , TC081, TC082, TC083, TC084, TC085, TC086, TC087, TC088, TC089, TC090
                                        , TC091, TC092, TC093
                                        , UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10)
                                        VALUES (@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE
                                        , @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID, @MODI_TIME, @MODI_AP, @MODI_PRID
                                        , @TC001, @TC002, @TC003, @TC004, @TC005, @TC006, @TC007, @TC008, @TC009, @TC010
                                        , @TC011, @TC012, @TC013, @TC014, @TC015, @TC016, @TC017, @TC018, @TC019, @TC020
                                        , @TC021, @TC022, @TC023, @TC024, @TC025, @TC026, @TC027, @TC028, @TC029, @TC030
                                        , @TC031, @TC032, @TC033, @TC034, @TC035, @TC036, @TC037, @TC038, @TC039, @TC040
                                        , @TC041, @TC042, @TC043, @TC044, @TC045, @TC046, @TC047, @TC048, @TC049, @TC050
                                        , @TC051, @TC052, @TC053, @TC054, @TC055, @TC056, @TC057, @TC058, @TC059, @TC060
                                        , @TC061, @TC062, @TC063, @TC064, @TC065, @TC066, @TC067, @TC068, @TC069, @TC070
                                        , @TC071, @TC072, @TC073, @TC074, @TC075, @TC076, @TC077, @TC078, @TC079, @TC080
                                        , @TC081, @TC082, @TC083, @TC084, @TC085, @TC086, @TC087, @TC088, @TC089, @TC090
                                        , @TC091, @TC092, @TC093
                                        , @UDF01, @UDF02, @UDF03, @UDF04, @UDF05, @UDF06, @UDF07, @UDF08, @UDF09, @UDF10)";
                                rowsAffected += sqlConnection.Execute(sql, cOPTCs);
                                #endregion

                                #region //COPTD
                                cOPTDs
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.TD001 = soErpPrefix; //訂單單別
                                        x.TD002 = soErpNo; //訂單單號
                                        x.TD015 = ""; //預測代號
                                        //x.TD017 = ""; //前置單據-單別
                                        //x.TD018 = ""; //前置單據-單號
                                        //x.TD019 = ""; //前置單據-序號
                                        x.TD021 = "N"; //確認碼
                                        x.TD022 = 0; //庫存數量
                                        x.TD023 = ""; //小單位
                                        x.TD028 = ""; //預測序號
                                        x.TD029 = ""; //包裝方式
                                        x.TD030 = 0; //毛重
                                        x.TD031 = 0; //材積
                                        x.TD032 = 0; //訂單包裝數量
                                        x.TD033 = 0; //已交包裝數量
                                        x.TD034 = 0; //贈品包裝量
                                        x.TD035 = 0; //贈品已交包裝量
                                        x.TD036 = ""; //包裝單位
                                        x.TD037 = ""; //原始客戶
                                        x.TD038 = ""; //請採購廠商
                                        x.TD039 = ""; //圖號
                                        x.TD040 = ""; //預留欄位
                                        x.TD041 = ""; //預留欄位
                                        x.TD042 = 0; //預留欄位
                                        x.TD043 = ""; //EBC訂單號碼
                                        x.TD044 = ""; //EBC訂單版次
                                        //x.TD045 = "9"; //來源
                                        x.TD046 = ""; //圖號版次
                                        x.TD047 = ""; //原預交日
                                        x.TD052 = 0; //備品包裝量
                                        x.TD053 = 0; //備品已交包裝量
                                        x.TD054 = 0; //預留欄位
                                        x.TD055 = 0; //預留欄位
                                        x.TD056 = ""; //預留欄位
                                        x.TD057 = ""; //預留欄位
                                        x.TD058 = ""; //預留欄位
                                        x.TD059 = 0; //贈品率
                                        x.TD060 = ""; //預留欄位
                                        x.TD061 = 0; //RFQ
                                        x.TD062 = ""; //NewCode
                                        x.TD063 = ""; //測試備註一
                                        x.TD064 = ""; //測試備註二
                                        x.TD065 = ""; //最終客戶代號
                                        x.TD066 = ""; //計畫批號
                                        x.TD067 = ""; //優先順序
                                        x.TD068 = ""; //預留欄位
                                        x.TD069 = ""; //鎖定交期
                                        x.TD071 = ""; //CRM單別
                                        x.TD072 = ""; //CRM單號
                                        x.TD073 = ""; //CRM序號
                                        x.TD074 = ""; //CRM合約代號
                                        x.TD075 = ""; //業務品號
                                        x.TD078 = 0; //已交計價數量
                                        x.TD500 = ""; //排程日期
                                        x.TD501 = 0; //可排量
                                        x.TD502 = ""; //產品系列
                                        x.TD503 = ""; //客戶需求日
                                        x.TD504 = ""; //以包裝單位計價
                                        x.UDF01 = "";
                                        x.UDF02 = "";
                                        x.UDF03 = "";
                                        x.UDF04 = "";
                                        x.UDF05 = "";
                                        x.UDF06 = 0;
                                        x.UDF07 = 0;
                                        x.UDF08 = 0;
                                        x.UDF09 = 0;
                                        x.UDF10 = 0;
                                    });

                                sql = @"INSERT INTO COPTD(COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE
                                        , FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID
                                        , TD001, TD002, TD003, TD004, TD005, TD006, TD007, TD008, TD009, TD010
                                        , TD011, TD012, TD013, TD014, TD015, TD016, TD017, TD018, TD019, TD020
                                        , TD021, TD022, TD023, TD024, TD025, TD026, TD027, TD028, TD029, TD030
                                        , TD031, TD032, TD033, TD034, TD035, TD036, TD037, TD038, TD039, TD040
                                        , TD041, TD042, TD043, TD044, TD045, TD046, TD047, TD048, TD049, TD050
                                        , TD051, TD052, TD053, TD054, TD055, TD056, TD057, TD058, TD059, TD060
                                        , TD061, TD062, TD063, TD064, TD065, TD066, TD067, TD068, TD069, TD070
                                        , TD071, TD072, TD073, TD074, TD075, TD076, TD077, TD078, TD079, TD080
                                        , TD500, TD501, TD502, TD503, TD504
                                        , UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10)
                                        VALUES(@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE
                                        , @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID, @MODI_TIME, @MODI_AP, @MODI_PRID
                                        , @TD001, @TD002, @TD003, @TD004, @TD005, @TD006, @TD007, @TD008, @TD009, @TD010
                                        , @TD011, @TD012, @TD013, @TD014, @TD015, @TD016, @TD017, @TD018, @TD019, @TD020
                                        , @TD021, @TD022, @TD023, @TD024, @TD025, @TD026, @TD027, @TD028, @TD029, @TD030
                                        , @TD031, @TD032, @TD033, @TD034, @TD035, @TD036, @TD037, @TD038, @TD039, @TD040
                                        , @TD041, @TD042, @TD043, @TD044, @TD045, @TD046, @TD047, @TD048, @TD049, @TD050
                                        , @TD051, @TD052, @TD053, @TD054, @TD055, @TD056, @TD057, @TD058, @TD059, @TD060
                                        , @TD061, @TD062, @TD063, @TD064, @TD065, @TD066, @TD067, @TD068, @TD069, @TD070
                                        , @TD071, @TD072, @TD073, @TD074, @TD075, @TD076, @TD077, @TD078, @TD079, @TD080
                                        , @TD500, @TD501, @TD502, @TD503, @TD504
                                        , @UDF01, @UDF02, @UDF03, @UDF04, @UDF05, @UDF06, @UDF07, @UDF08, @UDF09, @UDF10)";
                                rowsAffected += sqlConnection.Execute(sql, cOPTDs);
                                #endregion
                                break;
                            case "Y": //已拋轉修改
                                #region //COPTC
                                #region //判斷原單據是否存在
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 TC027
                                        FROM COPTC
                                        WHERE TC001 = @ErpPrefix
                                        AND TC002 = @ErpNo";
                                dynamicParameters.Add("ErpPrefix", soErpPrefix);
                                dynamicParameters.Add("ErpNo", soErpNo);

                                var resultExist = sqlConnection.Query(sql, dynamicParameters);
                                if (resultExist.Count() <= 0) throw new SystemException("原單據不存在!");

                                foreach (var item in resultExist)
                                {
                                    if (item.TC027 != "N") throw new SystemException("原單據確認碼為無法更動狀態!");
                                }
                                #endregion

                                cOPTCs
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.MODIFIER = userNo;
                                        x.MODI_DATE = dateNow;
                                        x.MODI_TIME = timeNow;
                                        x.MODI_AP = userNo + "PC";
                                        x.MODI_PRID = "BM";
                                        x.TC014 = paymentTerm; //付款條件
                                        x.TC027 = "N"; //確認碼
                                    });

                                sql = @"UPDATE COPTC SET
                                        MODIFIER = @MODIFIER,
                                        MODI_DATE = @MODI_DATE,
                                        MODI_TIME = @MODI_TIME,
                                        MODI_AP = @MODI_AP,
                                        MODI_PRID = @MODI_PRID,
                                        TC003 = @TC003, 
                                        TC004 = @TC004, 
                                        TC005 = @TC005, 
                                        TC006 = @TC006, 
                                        TC008 = @TC008, 
                                        TC009 = @TC009, 
                                        TC010 = @TC010, 
                                        TC011 = @TC011, 
                                        TC012 = @TC012, 
                                        TC013 = @TC013, 
                                        TC014 = @TC014, 
                                        TC015 = @TC015, 
                                        TC016 = @TC016, 
                                        TC018 = @TC018, 
                                        TC019 = @TC019,
                                        TC027 = @TC027,
                                        TC029 = @TC029,
                                        TC030 = @TC030,
                                        TC031 = @TC031,
                                        TC032 = @TC032,
                                        TC036 = @TC036,
                                        TC039 = @TC039,
                                        TC041 = @TC041,
                                        TC042 = @TC042,
                                        TC045 = @TC045,
                                        TC053 = @TC053,
                                        TC063 = @TC063,
                                        TC064 = @TC064,
                                        TC065 = @TC065,
                                        TC066 = @TC066,
                                        TC067 = @TC067,
                                        TC068 = @TC068,
                                        TC069 = @TC069,
                                        TC070 = @TC070,
                                        TC071 = @TC071,
                                        TC078 = @TC078,
                                        TC080 = @TC080,
                                        TC081 = @TC081,
                                        TC083 = @TC083,
                                        TC091 = @TC091
                                        WHERE TC001 = @TC001
                                        AND TC002 = @TC002";
                                rowsAffected += sqlConnection.Execute(sql, cOPTCs);
                                #endregion

                                #region //COPTD
                                #region //刪除原單據單身
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE COPTD
                                        WHERE TD001 = @ErpPrefix
                                        AND TD002 = @ErpNo";
                                dynamicParameters.Add("ErpPrefix", soErpPrefix);
                                dynamicParameters.Add("ErpNo", soErpNo);

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                cOPTDs
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.TD015 = ""; //預測代號
                                        //x.TD017 = ""; //前置單據-單別
                                        //x.TD018 = ""; //前置單據-單號
                                        //x.TD019 = ""; //前置單據-序號
                                        x.TD021 = "N"; //確認碼
                                        x.TD022 = 0; //庫存數量
                                        x.TD023 = ""; //小單位
                                        x.TD028 = ""; //預測序號
                                        x.TD029 = ""; //包裝方式
                                        x.TD030 = 0; //毛重
                                        x.TD031 = 0; //材積
                                        x.TD032 = 0; //訂單包裝數量
                                        x.TD033 = 0; //已交包裝數量
                                        x.TD034 = 0; //贈品包裝量
                                        x.TD035 = 0; //贈品已交包裝量
                                        x.TD036 = ""; //包裝單位
                                        x.TD037 = ""; //原始客戶
                                        x.TD038 = ""; //請採購廠商
                                        x.TD039 = ""; //圖號
                                        x.TD040 = ""; //預留欄位
                                        x.TD041 = ""; //預留欄位
                                        x.TD042 = 0; //預留欄位
                                        x.TD043 = ""; //EBC訂單號碼
                                        x.TD044 = ""; //EBC訂單版次
                                        //x.TD045 = "9"; //來源
                                        x.TD046 = ""; //圖號版次
                                        x.TD047 = ""; //原預交日
                                        x.TD052 = 0; //備品包裝量
                                        x.TD053 = 0; //備品已交包裝量
                                        x.TD054 = 0; //預留欄位
                                        x.TD055 = 0; //預留欄位
                                        x.TD056 = ""; //預留欄位
                                        x.TD057 = ""; //預留欄位
                                        x.TD058 = ""; //預留欄位
                                        x.TD059 = 0; //贈品率
                                        x.TD060 = ""; //預留欄位
                                        x.TD061 = 0; //RFQ
                                        x.TD062 = ""; //NewCode
                                        x.TD063 = ""; //測試備註一
                                        x.TD064 = ""; //測試備註二
                                        x.TD065 = ""; //最終客戶代號
                                        x.TD066 = ""; //計畫批號
                                        x.TD067 = ""; //優先順序
                                        x.TD068 = ""; //預留欄位
                                        x.TD069 = ""; //鎖定交期
                                        x.TD071 = ""; //CRM單別
                                        x.TD072 = ""; //CRM單號
                                        x.TD073 = ""; //CRM序號
                                        x.TD074 = ""; //CRM合約代號
                                        x.TD075 = ""; //業務品號
                                        x.TD078 = 0; //已交計價數量
                                        x.TD500 = ""; //排程日期
                                        x.TD501 = 0; //可排量
                                        x.TD502 = ""; //產品系列
                                        x.TD503 = ""; //客戶需求日
                                        x.TD504 = ""; //以包裝單位計價
                                        x.UDF01 = "";
                                        x.UDF02 = "";
                                        x.UDF03 = "";
                                        x.UDF04 = "";
                                        x.UDF05 = "";
                                        x.UDF06 = 0;
                                        x.UDF07 = 0;
                                        x.UDF08 = 0;
                                        x.UDF09 = 0;
                                        x.UDF10 = 0;
                                    });

                                sql = @"INSERT INTO COPTD(COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE
                                        , FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID
                                        , TD001, TD002, TD003, TD004, TD005, TD006, TD007, TD008, TD009, TD010
                                        , TD011, TD012, TD013, TD014, TD015, TD016, TD017, TD018, TD019, TD020
                                        , TD021, TD022, TD023, TD024, TD025, TD026, TD027, TD028, TD029, TD030
                                        , TD031, TD032, TD033, TD034, TD035, TD036, TD037, TD038, TD039, TD040
                                        , TD041, TD042, TD043, TD044, TD045, TD046, TD047, TD048, TD049, TD050
                                        , TD051, TD052, TD053, TD054, TD055, TD056, TD057, TD058, TD059, TD060
                                        , TD061, TD062, TD063, TD064, TD065, TD066, TD067, TD068, TD069, TD070
                                        , TD071, TD072, TD073, TD074, TD075, TD076, TD077, TD078, TD079, TD080
                                        , TD500, TD501, TD502, TD503, TD504
                                        , UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10)
                                        VALUES(@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE
                                        , @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID, @MODI_TIME, @MODI_AP, @MODI_PRID
                                        , @TD001, @TD002, @TD003, @TD004, @TD005, @TD006, @TD007, @TD008, @TD009, @TD010
                                        , @TD011, @TD012, @TD013, @TD014, @TD015, @TD016, @TD017, @TD018, @TD019, @TD020
                                        , @TD021, @TD022, @TD023, @TD024, @TD025, @TD026, @TD027, @TD028, @TD029, @TD030
                                        , @TD031, @TD032, @TD033, @TD034, @TD035, @TD036, @TD037, @TD038, @TD039, @TD040
                                        , @TD041, @TD042, @TD043, @TD044, @TD045, @TD046, @TD047, @TD048, @TD049, @TD050
                                        , @TD051, @TD052, @TD053, @TD054, @TD055, @TD056, @TD057, @TD058, @TD059, @TD060
                                        , @TD061, @TD062, @TD063, @TD064, @TD065, @TD066, @TD067, @TD068, @TD069, @TD070
                                        , @TD071, @TD072, @TD073, @TD074, @TD075, @TD076, @TD077, @TD078, @TD079, @TD080
                                        , @TD500, @TD501, @TD502, @TD503, @TD504
                                        , @UDF01, @UDF02, @UDF03, @UDF04, @UDF05, @UDF06, @UDF07, @UDF08, @UDF09, @UDF10)";
                                rowsAffected += sqlConnection.Execute(sql, cOPTDs);
                                #endregion
                                break;
                            default:
                                throw new SystemException("拋轉狀態錯誤!");
                        }
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region//單頭資料更新
                        switch (transferStatus)
                        {
                            case "N": //未拋轉新增
                                sql = @"UPDATE SCM.SaleOrder SET
                                        SoErpNo = @SoErpNo,
                                        TransferStatus = @TransferStatus,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE SoId = @SoId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        SoErpNo = soErpNo,
                                        TransferStatus = "Y",
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        SoId
                                    });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                break;
                            case "Y": //已拋轉修改
                                sql = @"UPDATE SCM.SaleOrder SET
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE SoId = @SoId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        SoId
                                    });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                break;
                            default:
                                throw new SystemException("拋轉狀態錯誤!");
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

        #region //UpdateSaleOrderConfirm -- 訂單ERP確認 -- Ben Ma 2023.06.15
        public string UpdateSaleOrderConfirm(int SoId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0, totalDetail = 0;
                    string confirmStatus = "", transferStatus = "", companyNo = "", soErpPrefix = "", soErpNo = "", departmentNo = "", userNo = "", userName = "";

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷訂單資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 SoErpPrefix, SoErpNo, ConfirmStatus, TransferStatus
                                FROM SCM.SaleOrder
                                WHERE SoId = @SoId
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("SoId", SoId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("訂單資料錯誤!");
                        
                        foreach (var item in result)
                        {
                            soErpPrefix = item.SoErpPrefix;
                            soErpNo = item.SoErpNo;
                            confirmStatus = item.ConfirmStatus;
                            transferStatus = item.TransferStatus;
                        }

                        #region //判斷訂單狀態
                        if (confirmStatus != "N") throw new SystemException("訂單不是未確認狀態，無法確認!");
                        if (transferStatus != "Y") throw new SystemException("訂單尚未拋轉，無法確認!");
                        #endregion
                        #endregion

                        #region //公司別資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var resultCompany = sqlConnection.Query(sql, dynamicParameters);
                        if (resultCompany.Count() <= 0) throw new SystemException("【公司別】資料錯誤!");

                        foreach (var item in resultCompany)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            companyNo = item.ErpNo;
                        }
                        #endregion

                        #region //使用者資料
                        // 目前僅晶彩使用BPM
                        string DetailCode = "";
                        if (CurrentCompany == 4)
                        {
                            DetailCode = "bpm-transfer";
                        }
                        else
                        {
                            DetailCode = "confirm";
                        }

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 d.UserNo, d.UserName, d.DepartmentNo
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
                                OUTER APPLY (
                                    SELECT da.UserNo, da.UserName, db.DepartmentNo
                                    FROM BAS.[User] da
                                    INNER JOIN BAS.Department db ON da.DepartmentId = db.DepartmentId
                                    WHERE da.UserId = @UserId
                                ) d
                                WHERE a.[Status] = @Status
                                AND b.[Status] = @Status
                                AND b.FunctionCode = @FunctionCode
                                AND a.DetailCode = @DetailCode
                                AND c.Authority > 0";
                        dynamicParameters.Add("UserId", CurrentUser);
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("Status", "A");
                        dynamicParameters.Add("FunctionCode", "SaleOrderManagement");
                        dynamicParameters.Add("DetailCode", DetailCode);

                        var resultUser = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUser.Count() <= 0) throw new SystemException("使用者資料錯誤!");

                        foreach (var item in resultUser)
                        {
                            userNo = item.UserNo;
                            userName = item.UserName;
                            departmentNo = item.DepartmentNo;
                        }
                        #endregion

                        #region //單身筆數
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT COUNT(1) TotalDetail
                                FROM SCM.SoDetail
                                WHERE SoId = @SoId";
                        dynamicParameters.Add("SoId", SoId);

                        var resultDetailExist = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in resultDetailExist)
                        {
                            totalDetail = Convert.ToInt32(item.TotalDetail);
                            if (totalDetail <= 0) throw new SystemException("單身筆數錯誤!");
                        }
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        #region //判斷ERP使用者是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 MF005
                                FROM ADMMF
                                WHERE COMPANY = @CompanyNo
                                AND MF001 = @User";
                        dynamicParameters.Add("CompanyNo", companyNo);
                        dynamicParameters.Add("User", userNo);

                        var resultUserExist = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUserExist.Count() <= 0) throw new SystemException("ERP使用者不存在!");

                        string adminUser = "N";
                        foreach (var item in resultUserExist)
                        {
                            adminUser = item.MF005; //超級使用者
                        }
                        #endregion

                        #region //判斷ERP使用者是否有權限
                        //晶彩不用卡控要ERP訂單確認權限
                        //if (companyNo != "DGJMO")
                        //{
                        //    if (adminUser != "Y")
                        //    {
                        //        dynamicParameters = new DynamicParameters();
                        //        sql = @"SELECT TOP 1 MG006
                        //            FROM ADMMG
                        //            WHERE COMPANY = @CompanyNo
                        //            AND MG001 = @User
                        //            AND MG002 = @Function
                        //            AND MG004 = 'Y'
                        //            AND MG006 LIKE @Auth";
                        //        dynamicParameters.Add("CompanyNo", companyNo);
                        //        dynamicParameters.Add("User", userNo);
                        //        dynamicParameters.Add("Function", "COPI06");
                        //        dynamicParameters.Add("Auth", "_Y__________"); //確認

                        //        var resultUserAuthExist = sqlConnection.Query(sql, dynamicParameters);
                        //        if (resultUserAuthExist.Count() <= 0) throw new SystemException("ERP使用者無權限!");
                        //    }
                        //}                        
                        #endregion

                        #region //判斷ERP單據是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 TC027
                                FROM COPTC
                                WHERE TC001 = @ErpPrefix
                                AND TC002 = @ErpNo";
                        dynamicParameters.Add("ErpPrefix", soErpPrefix);
                        dynamicParameters.Add("ErpNo", soErpNo);

                        var resultDocExist = sqlConnection.Query(sql, dynamicParameters);
                        if (resultDocExist.Count() <= 0) throw new SystemException("ERP單據不存在!");

                        foreach (var item in resultDocExist)
                        {
                            if (item.TC027 != "N") throw new SystemException("原單據確認碼為無法更動狀態!");
                        }
                        #endregion

                        int tempRowsAffected = 0;
                        #region //ERP確認
                        #region //COPTC
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE COPTC SET
                                TC027 = @ConfirmStatus,
                                TC040 = @ConfirmUser
                                WHERE TC001 = @ErpPrefix
                                AND TC002 = @ErpNo";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmStatus = "Y",
                                ConfirmUser = userNo,
                                ErpPrefix = soErpPrefix,
                                ErpNo = soErpNo
                            });

                        tempRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        if (tempRowsAffected != 1) throw new SystemException(string.Format("ERP單據 {0}-{1} 確認失敗!", soErpPrefix, soErpNo));
                        rowsAffected += tempRowsAffected;
                        #endregion

                        #region //COPTD
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE COPTD SET
                                TD021 = @ConfirmStatus
                                WHERE TD001 = @ErpPrefix
                                AND TD002 = @ErpNo";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmStatus = "Y",
                                ErpPrefix = soErpPrefix,
                                ErpNo = soErpNo
                            });

                        tempRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        if (tempRowsAffected != totalDetail) throw new SystemException(string.Format("ERP單據 {0}-{1} 單身確認失敗!", soErpPrefix, soErpNo));
                        rowsAffected += tempRowsAffected;
                        #endregion
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int tempRowsAffected = 0;
                        #region //BM確認
                        #region //SCM.SaleOrder
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.SaleOrder SET
                                ConfirmStatus = @ConfirmStatus,
                                ConfirmUserId = @ConfirmUserId,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SoId = @SoId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmStatus = "Y",
                                ConfirmUserId = CurrentUser,
                                LastModifiedDate,
                                LastModifiedBy,
                                SoId
                            });
                        tempRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        if (tempRowsAffected != 1) throw new SystemException(string.Format("BM單據 {0}-{1} 確認失敗!", soErpPrefix, soErpNo));
                        rowsAffected += tempRowsAffected;
                        #endregion

                        #region //SCM.SoDetail
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.SoDetail SET
                                ConfirmStatus = @ConfirmStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SoId = @SoId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmStatus = "Y",
                                LastModifiedDate,
                                LastModifiedBy,
                                SoId
                            });
                        tempRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        if (tempRowsAffected != totalDetail) throw new SystemException(string.Format("BM單據 {0}-{1} 單身確認失敗!", soErpPrefix, soErpNo));
                        rowsAffected += tempRowsAffected;
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

        #region //UpdateSaleOrderReConfirm -- 訂單ERP反確認 -- Ben Ma 2023.06.15
        public string UpdateSaleOrderReConfirm(int SoId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0, totalDetail = 0;
                    string confirmStatus = "", transferStatus = "", companyNo = "", soErpPrefix = "", soErpNo = "", departmentNo = "", userNo = "", userName = "";

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷訂單資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 SoErpPrefix, SoErpNo, ConfirmStatus, TransferStatus
                                FROM SCM.SaleOrder
                                WHERE SoId = @SoId
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("SoId", SoId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("訂單資料錯誤!");

                        #region //判斷訂單狀態
                        foreach (var item in result)
                        {
                            soErpPrefix = item.SoErpPrefix;
                            soErpNo = item.SoErpNo;
                            confirmStatus = item.ConfirmStatus;
                            transferStatus = item.TransferStatus;
                        }

                        if (confirmStatus != "Y") throw new SystemException("訂單尚未確認，無法反確認!");
                        if (transferStatus != "Y") throw new SystemException("訂單尚未拋轉，無法反確認!");
                        #endregion
                        #endregion

                        #region //公司別資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var resultCompany = sqlConnection.Query(sql, dynamicParameters);
                        if (resultCompany.Count() <= 0) throw new SystemException("【公司別】資料錯誤!");

                        foreach (var item in resultCompany)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            companyNo = item.ErpNo;
                        }
                        #endregion

                        #region //使用者資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 d.UserNo, d.UserName, d.DepartmentNo
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
                                OUTER APPLY (
                                    SELECT da.UserNo, da.UserName, db.DepartmentNo
                                    FROM BAS.[User] da
                                    INNER JOIN BAS.Department db ON da.DepartmentId = db.DepartmentId
                                    WHERE da.UserId = @UserId
                                ) d
                                WHERE a.[Status] = @Status
                                AND b.[Status] = @Status
                                AND b.FunctionCode = @FunctionCode
                                AND a.DetailCode = @DetailCode
                                AND c.Authority > 0";
                        dynamicParameters.Add("UserId", CurrentUser);
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("Status", "A");
                        dynamicParameters.Add("FunctionCode", "SaleOrderManagement");
                        dynamicParameters.Add("DetailCode", "reconfirm");

                        var resultUser = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUser.Count() <= 0) throw new SystemException("使用者資料錯誤!");

                        foreach (var item in resultUser)
                        {
                            userNo = item.UserNo;
                            userName = item.UserName;
                            departmentNo = item.DepartmentNo;
                        }
                        #endregion

                        #region //單身筆數
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT COUNT(1) TotalDetail
                                FROM SCM.SoDetail
                                WHERE SoId = @SoId";
                        dynamicParameters.Add("SoId", SoId);

                        var resultDetailExist = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in resultDetailExist)
                        {
                            totalDetail = Convert.ToInt32(item.TotalDetail);
                            if (totalDetail <= 0) throw new SystemException("單身筆數錯誤!");
                        }
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        #region //判斷ERP使用者是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 MF005
                                FROM ADMMF
                                WHERE COMPANY = @CompanyNo
                                AND MF001 = @User";
                        dynamicParameters.Add("CompanyNo", companyNo);
                        dynamicParameters.Add("User", userNo);

                        var resultUserExist = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUserExist.Count() <= 0) throw new SystemException("ERP使用者不存在!");

                        string adminUser = "N";
                        foreach (var item in resultUserExist)
                        {
                            adminUser = item.MF005; //超級使用者
                        }
                        #endregion

                        #region //判斷ERP使用者是否有權限
                        //晶彩不用卡控要ERP訂單確認權限
                        if (companyNo != "DGJMO") {
                            if (adminUser != "Y")
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 MG006
                                    FROM ADMMG
                                    WHERE COMPANY = @CompanyNo
                                    AND MG001 = @User
                                    AND MG002 = @Function
                                    AND MG004 = 'Y'
                                    AND MG006 LIKE @Auth";
                                dynamicParameters.Add("CompanyNo", companyNo);
                                dynamicParameters.Add("User", userNo);
                                dynamicParameters.Add("Function", "COPI06");
                                dynamicParameters.Add("Auth", "____Y_______"); //反確認
                                var resultUserAuthExist = sqlConnection.Query(sql, dynamicParameters);
                                if (resultUserAuthExist.Count() <= 0) throw new SystemException("ERP使用者無權限!");
                            }
                        }
                        #endregion

                        #region //判斷ERP單據是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 TC027
                                FROM COPTC
                                WHERE TC001 = @ErpPrefix
                                AND TC002 = @ErpNo";
                        dynamicParameters.Add("ErpPrefix", soErpPrefix);
                        dynamicParameters.Add("ErpNo", soErpNo);

                        var resultDocExist = sqlConnection.Query(sql, dynamicParameters);
                        if (resultDocExist.Count() <= 0) throw new SystemException("ERP單據不存在!");

                        foreach (var item in resultDocExist)
                        {
                            if (item.TC027 != "Y") throw new SystemException("原單據確認碼為無法更動狀態!");
                        }
                        #endregion

                        #region //判斷ERP變更單單據是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM COPTE
                                WHERE TE001 = @ErpPrefix
                                AND TE002 = @ErpNo";
                        dynamicParameters.Add("ErpPrefix", soErpPrefix);
                        dynamicParameters.Add("ErpNo", soErpNo);

                        var resultChangeNoteExist = sqlConnection.Query(sql, dynamicParameters);
                        if (resultChangeNoteExist.Count() > 0) throw new SystemException("ERP變更單單據已存在，無法反確認!");
                        #endregion

                        #region //判斷ERP暫出單單據是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM INVTG
                                WHERE TG014 = @ErpPrefix
                                AND TG015 = @ErpNo";
                        dynamicParameters.Add("ErpPrefix", soErpPrefix);
                        dynamicParameters.Add("ErpNo", soErpNo);

                        var resultTempShippingNoteExist = sqlConnection.Query(sql, dynamicParameters);
                        if (resultTempShippingNoteExist.Count() > 0) throw new SystemException("ERP暫出單單據已存在，無法反確認!");
                        #endregion

                        #region //判斷ERP銷貨單單據是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM COPTG
                                WHERE TG048 = @ErpPrefix
                                AND TG049 = @ErpNo";
                        dynamicParameters.Add("ErpPrefix", soErpPrefix);
                        dynamicParameters.Add("ErpNo", soErpNo);

                        var resultShippinigOrderExist = sqlConnection.Query(sql, dynamicParameters);
                        if (resultShippinigOrderExist.Count() > 0) throw new SystemException("ERP銷貨單單據已存在，無法反確認!");
                        #endregion

                        #region //判斷ERP結帳單是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM ACRTB a
                                INNER JOIN ACRTA b ON a.TB001 = b.TA001 AND a.TB002 = b.TA002
                                WHERE a.TB005 = @ErpPrefix
                                AND a.TB006 = @ErpNo";
                        dynamicParameters.Add("ErpPrefix", soErpPrefix);
                        dynamicParameters.Add("ErpNo", soErpNo);

                        var resultARSheetExist = sqlConnection.Query(sql, dynamicParameters);
                        if (resultARSheetExist.Count() > 0) throw new SystemException("ERP結帳單單據已存在，無法反確認!");
                        #endregion

                        int tempRowsAffected = 0;
                        #region //ERP反確認
                        #region //COPTC
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE COPTC SET
                                TC027 = @ConfirmStatus
                                WHERE TC001 = @ErpPrefix
                                AND TC002 = @ErpNo";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmStatus = "N",
                                ErpPrefix = soErpPrefix,
                                ErpNo = soErpNo
                            });

                        tempRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        if (tempRowsAffected != 1) throw new SystemException(string.Format("ERP單據 {0}-{1} 反確認失敗!", soErpPrefix, soErpNo));
                        rowsAffected += tempRowsAffected;
                        #endregion

                        #region //COPTD
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE COPTD SET
                                TD021 = @ConfirmStatus
                                WHERE TD001 = @ErpPrefix
                                AND TD002 = @ErpNo";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmStatus = "N",
                                ErpPrefix = soErpPrefix,
                                ErpNo = soErpNo
                            });

                        tempRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        if (tempRowsAffected != totalDetail) throw new SystemException(string.Format("ERP單據 {0}-{1} 單身反確認失敗!", soErpPrefix, soErpNo));
                        rowsAffected += tempRowsAffected;
                        #endregion
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int tempRowsAffected = 0;
                        #region //BM反確認
                        #region //SCM.SaleOrder
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.SaleOrder SET
                                ConfirmStatus = @ConfirmStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SoId = @SoId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmStatus = "N",
                                LastModifiedDate,
                                LastModifiedBy,
                                SoId
                            });
                        tempRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        if (tempRowsAffected != 1) throw new SystemException(string.Format("BM單據 {0}-{1} 反確認失敗!", soErpPrefix, soErpNo));
                        rowsAffected += tempRowsAffected;
                        #endregion

                        #region //SCM.SoDetail
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.SoDetail SET
                                ConfirmStatus = @ConfirmStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SoId = @SoId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmStatus = "N",
                                LastModifiedDate,
                                LastModifiedBy,
                                SoId
                            });
                        tempRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        if (tempRowsAffected != totalDetail) throw new SystemException(string.Format("BM單據 {0}-{1} 單身反確認失敗!", soErpPrefix, soErpNo));
                        rowsAffected += tempRowsAffected;
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

        #region //UpdateSaleOrderVoid -- 訂單ERP作廢 -- Ben Ma 2023.06.14
        public string UpdateSaleOrderVoid(int SoId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0, totalDetail = 0;
                    string confirmStatus = "", transferStatus = "", companyNo = "", soErpPrefix = "", soErpNo = "", departmentNo = "", userNo = "", userName = "";

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷訂單資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 SoErpPrefix, SoErpNo, ConfirmStatus, TransferStatus
                                FROM SCM.SaleOrder
                                WHERE SoId = @SoId
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("SoId", SoId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("訂單資料錯誤!");

                        foreach (var item in result)
                        {
                            soErpPrefix = item.SoErpPrefix;
                            soErpNo = item.SoErpNo;
                            confirmStatus = item.ConfirmStatus;
                            transferStatus = item.TransferStatus;
                        }

                        #region //判斷訂單狀態
                        if (confirmStatus != "N") throw new SystemException("訂單不是未確認狀態，無法作廢!");
                        if (transferStatus != "Y") throw new SystemException("訂單尚未拋轉，無法作廢!");
                        #endregion
                        #endregion

                        #region //公司別資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var resultCompany = sqlConnection.Query(sql, dynamicParameters);
                        if (resultCompany.Count() <= 0) throw new SystemException("【公司別】資料錯誤!");

                        foreach (var item in resultCompany)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            companyNo = item.ErpNo;
                        }
                        #endregion

                        #region //使用者資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 d.UserNo, d.UserName, d.DepartmentNo
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
                                OUTER APPLY (
                                    SELECT da.UserNo, da.UserName, db.DepartmentNo
                                    FROM BAS.[User] da
                                    INNER JOIN BAS.Department db ON da.DepartmentId = db.DepartmentId
                                    WHERE da.UserId = @UserId
                                ) d
                                WHERE a.[Status] = @Status
                                AND b.[Status] = @Status
                                AND b.FunctionCode = @FunctionCode
                                AND a.DetailCode = @DetailCode
                                AND c.Authority > 0";
                        dynamicParameters.Add("UserId", CurrentUser);
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("Status", "A");
                        dynamicParameters.Add("FunctionCode", "SaleOrderManagement");
                        dynamicParameters.Add("DetailCode", "void");

                        var resultUser = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUser.Count() <= 0) throw new SystemException("使用者資料錯誤!");

                        foreach (var item in resultUser)
                        {
                            userNo = item.UserNo;
                            userName = item.UserName;
                            departmentNo = item.DepartmentNo;
                        }
                        #endregion

                        #region //單身筆數
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT COUNT(1) TotalDetail
                                FROM SCM.SoDetail
                                WHERE SoId = @SoId";
                        dynamicParameters.Add("SoId", SoId);

                        var resultDetailExist = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in resultDetailExist)
                        {
                            totalDetail = Convert.ToInt32(item.TotalDetail);
                            if (totalDetail <= 0) throw new SystemException("單身筆數錯誤!");
                        }
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        #region //判斷ERP使用者是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 MF005
                                FROM ADMMF
                                WHERE COMPANY = @CompanyNo
                                AND MF001 = @User";
                        dynamicParameters.Add("CompanyNo", companyNo);
                        dynamicParameters.Add("User", userNo);

                        var resultUserExist = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUserExist.Count() <= 0) throw new SystemException("ERP使用者不存在!");

                        string adminUser = "N";
                        foreach (var item in resultUserExist)
                        {
                            adminUser = item.MF005; //超級使用者
                        }
                        #endregion

                        #region //判斷ERP使用者是否有權限
                        //晶彩不用卡控要ERP訂單確認權限
                        if (companyNo != "DGJMO")
                        {
                            if (adminUser != "Y")
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 MG006
                                    FROM ADMMG
                                    WHERE COMPANY = @CompanyNo
                                    AND MG001 = @User
                                    AND MG002 = @Function
                                    AND MG004 = 'Y'
                                    AND MG006 LIKE @Auth";
                                dynamicParameters.Add("CompanyNo", companyNo);
                                dynamicParameters.Add("User", userNo);
                                dynamicParameters.Add("Function", "COPI06");
                                dynamicParameters.Add("Auth", "_________Y__"); //作廢
                                var resultUserAuthExist = sqlConnection.Query(sql, dynamicParameters);
                                if (resultUserAuthExist.Count() <= 0) throw new SystemException("ERP使用者無權限!");
                            }
                        }
                        #endregion

                        #region //判斷ERP單據是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 TC027
                                FROM COPTC
                                WHERE TC001 = @ErpPrefix
                                AND TC002 = @ErpNo";
                        dynamicParameters.Add("ErpPrefix", soErpPrefix);
                        dynamicParameters.Add("ErpNo", soErpNo);

                        var resultExist = sqlConnection.Query(sql, dynamicParameters);
                        if (resultExist.Count() <= 0) throw new SystemException("ERP單據不存在!");

                        if (resultExist.Count() > 0)
                        {
                            foreach (var item in resultExist)
                            {
                                if (item.TC027 != "N") throw new SystemException("原單據確認碼為無法更動狀態!");
                            }
                        }
                        #endregion

                        int tempRowsAffected = 0;
                        #region //ERP作廢
                        #region //COPTC
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE COPTC SET
                                TC027 = @ConfirmStatus,
                                TC040 = @ConfirmUser
                                WHERE TC001 = @ErpPrefix
                                AND TC002 = @ErpNo";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmStatus = "V",
                                ConfirmUser = userNo,
                                ErpPrefix = soErpPrefix,
                                ErpNo = soErpNo
                            });

                        tempRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        if (tempRowsAffected != 1) throw new SystemException(string.Format("ERP單據 {0}-{1} 作廢失敗!", soErpPrefix, soErpNo));
                        rowsAffected += tempRowsAffected;
                        #endregion

                        #region //COPTD
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE COPTD SET
                                TD021 = @ConfirmStatus
                                WHERE TD001 = @ErpPrefix
                                AND TD002 = @ErpNo";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmStatus = "V",
                                ErpPrefix = soErpPrefix,
                                ErpNo = soErpNo
                            });

                        tempRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        if (tempRowsAffected != totalDetail) throw new SystemException(string.Format("ERP單據 {0}-{1} 單身作廢失敗!", soErpPrefix, soErpNo));
                        rowsAffected += tempRowsAffected;
                        #endregion
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int tempRowsAffected = 0;
                        #region //BM作廢
                        #region //SCM.SaleOrder
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.SaleOrder SET
                                ConfirmStatus = @ConfirmStatus,
                                ConfirmUserId = @ConfirmUserId,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SoId = @SoId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmStatus = "V",
                                ConfirmUserId = CurrentUser,
                                LastModifiedDate,
                                LastModifiedBy,
                                SoId
                            });
                        tempRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        if (tempRowsAffected != 1) throw new SystemException(string.Format("BM單據 {0}-{1} 作廢失敗!", soErpPrefix, soErpNo));
                        rowsAffected += tempRowsAffected;
                        #endregion

                        #region //SCM.SoDetail
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.SoDetail SET
                                ConfirmStatus = @ConfirmStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SoId = @SoId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmStatus = "V",
                                LastModifiedDate,
                                LastModifiedBy,
                                SoId
                            });
                        tempRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        if (tempRowsAffected != totalDetail) throw new SystemException(string.Format("BM單據 {0}-{1} 單身作廢失敗!", soErpPrefix, soErpNo));
                        rowsAffected += tempRowsAffected;
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

        #region //UpdateSaleOrderModify -- 訂單資料變更 -- Ben Ma 2023.06.26
        public string UpdateSaleOrderModify(string SoDetails, string Deposits)
        {
            try
            {
                JObject tempJObject = new JObject();
                if (!SoDetails.TryParseJson(out tempJObject)) throw new SystemException("訂單資料格式錯誤");
                if (!Deposits.TryParseJson(out tempJObject)) throw new SystemException("訂金分批資料格式錯誤");

                JObject soJson = JObject.Parse(SoDetails);
                JObject depositJson = JObject.Parse(Deposits);

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    string confirmStatus = "", transferStatus = "", companyNo = "", soErpPrefix = "", soErpNo = "", departmentNo = "", userNo = "", userName = ""
                        , detailConfirmStatus = "", detailClosureStatus = "";
                    double totalRate = 0;
                    List<SoDetail> soDetails = new List<SoDetail>();
                    List<SoDeposit> soDeposits = new List<SoDeposit>();
                    List<SoDeposit> oriSoDeposits = new List<SoDeposit>();

                    string dateNow = DateTime.Now.ToString("yyyyMMdd");
                    string timeNow = DateTime.Now.ToString("HH:mm:ss");

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷訂單資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 SoErpPrefix, SoErpNo, ConfirmStatus, TransferStatus
                                FROM SCM.SaleOrder
                                WHERE SoId = @SoId
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("SoId", Convert.ToInt32(soJson["soId"]));
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("訂單資料錯誤!");

                        foreach (var item in result)
                        {
                            soErpPrefix = item.SoErpPrefix;
                            soErpNo = item.SoErpNo;
                            confirmStatus = item.ConfirmStatus;
                            transferStatus = item.TransferStatus;
                        }

                        #region //判斷訂單狀態
                        if (confirmStatus != "Y") throw new SystemException("訂單不是確認狀態，無法變更!");
                        if (transferStatus != "Y") throw new SystemException("訂單尚未拋轉，無法變更!");
                        #endregion
                        #endregion

                        #region //判斷訂單單身資料是否正確
                        if (soJson["data"].Count() > 0)
                        {
                            for (int i = 0; i < soJson["data"].Count(); i++)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 SoSequence, ConfirmStatus, ClosureStatus
                                        FROM SCM.SoDetail
                                        WHERE SoDetailId = @SoDetailId
                                        AND SoId = @SoId";
                                dynamicParameters.Add("SoDetailId", Convert.ToInt32(soJson["data"][i]["soDetailId"]));
                                dynamicParameters.Add("SoId", Convert.ToInt32(soJson["soId"]));

                                var resultDetail = sqlConnection.Query(sql, dynamicParameters);
                                if (resultDetail.Count() <= 0) throw new SystemException("訂單單身資料錯誤!");

                                foreach (var item in resultDetail)
                                {
                                    soDetails.Add(new SoDetail
                                    {
                                        SoSequence = item.SoSequence,
                                        SoDetailRemark = soJson["data"][i]["soDetailRemark"].ToString()
                                    });

                                    detailConfirmStatus = item.ConfirmStatus;
                                    detailClosureStatus = item.ClosureStatus;
                                }

                                #region //判斷訂單單身狀態
                                if (detailConfirmStatus != "Y") throw new SystemException("訂單單身不是確認狀態，無法變更!");
                                if (detailClosureStatus != "N") throw new SystemException("訂單單身已結案，無法變更!");
                                #endregion
                            }
                        }
                        #endregion

                        #region //公司別資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var resultCompany = sqlConnection.Query(sql, dynamicParameters);
                        if (resultCompany.Count() <= 0) throw new SystemException("【公司別】資料錯誤!");

                        foreach (var item in resultCompany)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            companyNo = item.ErpNo;
                        }
                        #endregion

                        #region //使用者資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 d.UserNo, d.UserName, d.DepartmentNo
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
                                OUTER APPLY (
                                    SELECT da.UserNo, da.UserName, db.DepartmentNo
                                    FROM BAS.[User] da
                                    INNER JOIN BAS.Department db ON da.DepartmentId = db.DepartmentId
                                    WHERE da.UserId = @UserId
                                ) d
                                WHERE a.[Status] = @Status
                                AND b.[Status] = @Status
                                AND b.FunctionCode = @FunctionCode
                                AND a.DetailCode = @DetailCode
                                AND c.Authority > 0";
                        dynamicParameters.Add("UserId", CurrentUser);
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("Status", "A");
                        dynamicParameters.Add("FunctionCode", "SaleOrderManagement");
                        dynamicParameters.Add("DetailCode", "modify");

                        var resultUser = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUser.Count() <= 0) throw new SystemException("使用者資料錯誤!");

                        foreach (var item in resultUser)
                        {
                            userNo = item.UserNo;
                            userName = item.UserName;
                            departmentNo = item.DepartmentNo;
                        }
                        #endregion

                        #region //訂金資料
                        if (depositJson["data"].Count() > 0)
                        {
                            for (int i = 0; i < depositJson["data"].Count(); i++)
                            {
                                soDeposits.Add(new SoDeposit
                                {
                                    Sequence = depositJson["data"][i]["Sequence"].ToString().Length > 0 ? depositJson["data"][i]["Sequence"].ToString() : i.ToString(),
                                    Ratio = Convert.ToDouble(depositJson["data"][i]["Ratio"]),
                                    Amount = Convert.ToDouble(depositJson["data"][i]["Amount"]),
                                    PaymentDate = depositJson["data"][i]["PaymentDate"].ToString().Length > 0 ? Convert.ToDateTime(depositJson["data"][i]["PaymentDate"]) : (DateTime?)null,
                                    ClosingStatus = depositJson["data"][i]["ClosingStatus"].ToString()
                                });
                            }
                        }
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        #region //判斷ERP使用者是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 MF005
                                FROM ADMMF
                                WHERE COMPANY = @CompanyNo
                                AND MF001 = @User";
                        dynamicParameters.Add("CompanyNo", companyNo);
                        dynamicParameters.Add("User", userNo);

                        var resultUserExist = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUserExist.Count() <= 0) throw new SystemException("ERP使用者不存在!");

                        string adminUser = "N";
                        foreach (var item in resultUserExist)
                        {
                            adminUser = item.MF005; //超級使用者
                        }
                        #endregion

                        #region //判斷ERP使用者是否有權限
                        //晶彩不用卡控要ERP訂單確認權限
                        if (companyNo != "DGJMO")
                        {
                            if (adminUser != "Y")
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 MG006
                                    FROM ADMMG
                                    WHERE COMPANY = @CompanyNo
                                    AND MG001 = @User
                                    AND MG002 = @Function
                                    AND MG004 = 'Y'
                                    AND MG006 LIKE @Auth";
                                dynamicParameters.Add("CompanyNo", companyNo);
                                dynamicParameters.Add("User", userNo);
                                dynamicParameters.Add("Function", "COPI06");
                                dynamicParameters.Add("Auth", "Y___________"); //修改
                                var resultUserAuthExist = sqlConnection.Query(sql, dynamicParameters);
                                if (resultUserAuthExist.Count() <= 0) throw new SystemException("ERP使用者無權限!");
                            }
                        }
                        #endregion

                        #region //判斷ERP單據是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 TC027
                                FROM COPTC
                                WHERE TC001 = @ErpPrefix
                                AND TC002 = @ErpNo";
                        dynamicParameters.Add("ErpPrefix", soErpPrefix);
                        dynamicParameters.Add("ErpNo", soErpNo);

                        var resultExist = sqlConnection.Query(sql, dynamicParameters);
                        if (resultExist.Count() <= 0) throw new SystemException("ERP單據不存在!");

                        if (resultExist.Count() > 0)
                        {
                            foreach (var item in resultExist)
                            {
                                if (item.TC027 != "Y") throw new SystemException("原單據未確認!");
                            }
                        }
                        #endregion

                        #region //判斷ERP單據單身是否存在
                        if (soDetails.Count > 0)
                        {
                            foreach (var soDetail in soDetails)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 TD016, TD021
                                        FROM COPTD
                                        WHERE TD001 = @ErpPrefix
                                        AND TD002 = @ErpNo
                                        AND TD003 = @Sequence";
                                dynamicParameters.Add("ErpPrefix", soErpPrefix);
                                dynamicParameters.Add("ErpNo", soErpNo);
                                dynamicParameters.Add("Sequence", soDetail.SoSequence);

                                var resultDetailExist = sqlConnection.Query(sql, dynamicParameters);
                                if (resultDetailExist.Count() <= 0) throw new SystemException("ERP單據單身不存在!");

                                if (resultDetailExist.Count() > 0)
                                {
                                    foreach (var item in resultDetailExist)
                                    {
                                        if (item.TD016 != "N") throw new SystemException("原單據單身已結案!");
                                        if (item.TD021 != "Y") throw new SystemException("原單據單身未確認!");
                                    }
                                }
                            }
                        }
                        #endregion

                        #region //判斷ERP訂金資料狀態
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT UC003 'Sequence', CAST(UC004 * 100 AS float) Ratio, UC005 Amount
                                , CASE WHEN LEN(LTRIM(RTRIM(UC006))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(UC006)) AS date), 'yyyy-MM-dd') ELSE '' END PaymentDate
                                , UC007 ClosingStatus
                                FROM COPUC
                                WHERE UC001 = @ErpPrefix
                                AND UC002 = @ErpNo
                                ORDER BY UC003";
                        dynamicParameters.Add("ErpPrefix", soErpPrefix);
                        dynamicParameters.Add("ErpNo", soErpNo);

                        var resultDeposit = sqlConnection.Query(sql, dynamicParameters);
                        if (resultDeposit.Count() > 0)
                        {
                            foreach (var item in resultDeposit)
                            {
                                oriSoDeposits.Add(new SoDeposit
                                {
                                    Sequence = item.Sequence.ToString(),
                                    Ratio = Convert.ToDouble(item.Ratio),
                                    Amount = Convert.ToDouble(item.Amount),
                                    PaymentDate = item.PaymentDate.ToString().Length > 0 ? Convert.ToDateTime(item.PaymentDate) : (DateTime?)null,
                                    ClosingStatus = item.ClosingStatus
                                });
                            }
                        }

                        #region //篩選異動訂金資料
                        var dictionaryErpDeposit = oriSoDeposits.ToDictionary(x => x.Sequence, x => x.Sequence);
                        var dictionaryBmDeposit = soDeposits.ToDictionary(x => x.Sequence, x => x.Sequence);
                        var changeDeposit = dictionaryErpDeposit.Where(x => !dictionaryBmDeposit.ContainsKey(x.Key)).ToList();
                        var changeDepositList = changeDeposit.Select(x => x.Value).ToList();
                        #endregion

                        if (changeDepositList.Count > 0)
                        {
                            foreach (var deposit in changeDepositList)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM ACRTB a
                                        INNER JOIN ACRTA b ON a.TB001 = b.TA001 AND a.TB002 = b.TA002
                                        WHERE a.TB005 = @ErpPrefix
                                        AND a.TB006 = @ErpNo
                                        AND a.TB023 = @Sequence";
                                dynamicParameters.Add("ErpPrefix", soErpPrefix);
                                dynamicParameters.Add("ErpNo", soErpNo);
                                dynamicParameters.Add("Sequence", deposit);

                                var resultARSheetExist = sqlConnection.Query(sql, dynamicParameters);
                                if (resultARSheetExist.Count() > 0) throw new SystemException("訂金序號【" + deposit + "】已開立結帳單，無法刪除!");
                            }
                        }

                        switch (soJson["depositPartial"].ToString())
                        {
                            case "N":
                                if (oriSoDeposits.Count() > 0) throw new SystemException("已有訂金資料，無法更改定金分批!");
                                break;
                            case "Y":
                                break;
                            default:
                                throw new SystemException("訂金分批資料錯誤");
                        }
                        #endregion

                        #region //ERP資料修改
                        #region //COPUC
                        #region //刪除未結帳訂金資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE COPUC
                                WHERE UC001 = @ErpPrefix
                                AND UC002 = @ErpNo
                                AND UC007 = @ClosingStatus";
                        dynamicParameters.Add("ErpPrefix", soErpPrefix);
                        dynamicParameters.Add("ErpNo", soErpNo);
                        dynamicParameters.Add("ClosingStatus", "N");

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //新增訂金資料
                        #region //目前序號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ISNULL(CAST(MAX(UC003) AS INT), 0) + 1 MaxSequence
                                FROM COPUC
                                WHERE UC001 = @ErpPrefix
                                AND UC002 = @ErpNo";
                        dynamicParameters.Add("ErpPrefix", soErpPrefix);
                        dynamicParameters.Add("ErpNo", soErpNo);

                        var resultSequence = sqlConnection.Query(sql, dynamicParameters);

                        int maxSequence = 1;
                        if (resultSequence.Count() > 0)
                        {
                            foreach (var item in resultSequence)
                            {
                                maxSequence = Convert.ToInt32(item.MaxSequence);
                            }
                        }
                        #endregion

                        var insertData = soDeposits.Where(x => x.ClosingStatus != "Y").ToList();
                        foreach (var item in insertData)
                        {
                            totalRate += item.Ratio / 100;

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO COPUC (COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE
                                    , FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID
                                    , UC001, UC002, UC003, UC004, UC005, UC006, UC007, UC008, UC009, UC010
                                    , UC011, UC012
                                    , UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10)
                                    VALUES (@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE
                                    , @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID, @MODI_TIME, @MODI_AP, @MODI_PRID
                                    , @UC001, @UC002, @UC003, @UC004, @UC005, @UC006, @UC007, @UC008, @UC009, @UC010
                                    , @UC011, @UC012
                                    , @UDF01, @UDF02, @UDF03, @UDF04, @UDF05, @UDF06, @UDF07, @UDF08, @UDF09, @UDF10)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    COMPANY = companyNo,
                                    CREATOR = userNo,
                                    USR_GROUP = "",
                                    CREATE_DATE = dateNow,
                                    MODIFIER = "",
                                    MODI_DATE = "",
                                    FLAG = "1",
                                    CREATE_TIME = timeNow,
                                    CREATE_AP = userNo + "PC",
                                    CREATE_PRID = "BM",
                                    MODI_TIME = "",
                                    MODI_AP = "",
                                    MODI_PRID = "",
                                    UC001 = soErpPrefix,
                                    UC002 = soErpNo,
                                    UC003 = string.Format("{0:0000}", maxSequence),
                                    UC004 = item.Ratio / 100,
                                    UC005 = item.Amount,
                                    UC006 = item.PaymentDate != null ? Convert.ToDateTime(item.PaymentDate).ToString("yyyyMMdd") : "",
                                    UC007 = "N",
                                    UC008 = 0,
                                    UC009 = 0,
                                    UC010 = "",
                                    UC011 = "",
                                    UC012 = "",
                                    UDF01 = "",
                                    UDF02 = "",
                                    UDF03 = "",
                                    UDF04 = "",
                                    UDF05 = "",
                                    UDF06 = 0,
                                    UDF07 = 0,
                                    UDF08 = 0,
                                    UDF09 = 0,
                                    UDF10 = 0
                                });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();
                            maxSequence++;
                        }
                        #endregion
                        #endregion

                        #region //COPTC
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE COPTC SET
                                TC012 = @CustomerPurchaseOrder,
                                TC015 = @SoRemark,
                                TC045 = @DepositRate,
                                TC070 = @DepositPartial
                                WHERE TC001 = @ErpPrefix
                                AND TC002 = @ErpNo";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CustomerPurchaseOrder = soJson["customerPurchaseOrder"].ToString(),
                                SoRemark = soJson["soRemark"].ToString(),
                                DepositRate = totalRate,
                                DepositPartial = soJson["depositPartial"].ToString(),
                                ErpPrefix = soErpPrefix,
                                ErpNo = soErpNo
                            });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //COPTD
                        if (soDetails.Count > 0)
                        {
                            foreach (var soDetail in soDetails)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE COPTD SET
                                        TD020 = @SoDetailRemark
                                        WHERE TD001 = @ErpPrefix
                                        AND TD002 = @ErpNo
                                        AND TD003 = @Sequence";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        soDetail.SoDetailRemark,
                                        ErpPrefix = soErpPrefix,
                                        ErpNo = soErpNo,
                                        Sequence = soDetail.SoSequence
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            }
                        }
                        #endregion
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //BM資料修改
                        #region //SCM.SaleOrder
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.SaleOrder SET
                                SoRemark = @SoRemark,
                                CustomerPurchaseOrder = @CustomerPurchaseOrder,
                                DepositPartial = @DepositPartial,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SoId = @SoId";
                        sql = @"UPDATE SCM.SaleOrder SET
                                SoRemark = @SoRemark,
                                CustomerPurchaseOrder = @CustomerPurchaseOrder,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SoId = @SoId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SoRemark = soJson["soRemark"].ToString(),
                                CustomerPurchaseOrder = soJson["customerPurchaseOrder"].ToString(),
                                DepositRate = totalRate,
                                DepositPartial = soJson["depositPartial"].ToString(),
                                LastModifiedDate,
                                LastModifiedBy,
                                SoId = Convert.ToInt32(soJson["soId"])
                            });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //SCM.SoDetail
                        if (soJson["data"].Count() > 0)
                        {
                            for (int i = 0; i < soJson["data"].Count(); i++)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE SCM.SoDetail SET
                                        SoDetailRemark = @SoDetailRemark,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE SoDetailId = @SoDetailId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        SoDetailRemark = soJson["data"][i]["soDetailRemark"].ToString(),
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        SoDetailId = Convert.ToInt32(soJson["data"][i]["soDetailId"])
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            }
                        }
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

        #region //UpdateSaleOrderSynchronize -- 訂單資料同步 -- Zoey 2022.07.18
        public string UpdateSaleOrderSynchronize(string CompanyNo, string UpdateDate)
        {
            try
            {
                List<SaleOrder> saleOrders = new List<SaleOrder>();
                List<SoDetail> soDetails = new List<SoDetail>();

                using (TransactionScope transactionScope = new TransactionScope(TransactionScopeOption.Required, new TimeSpan(0, 180, 0)))
                {
                    int CompanyId = -1;
                    string ErpConnectionStrings = "";

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷公司資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 CompanyId, CompanyName, ErpDb
                                FROM BAS.Company
                                WHERE CompanyNo = @CompanyNo";
                        dynamicParameters.Add("CompanyNo", CompanyNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【公司】資料錯誤!");

                        foreach (var item in result)
                        {
                            CompanyId = Convert.ToInt32(item.CompanyId);
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        }
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        #region //撈取ERP訂單單頭資料
                        sql = @"SELECT LTRIM(RTRIM(TC005)) DepartmentNo, LTRIM(RTRIM(TC001)) SoErpPrefix, LTRIM(RTRIM(TC002)) SoErpNo, LTRIM(RTRIM(TC069)) Version
                                , CASE WHEN LEN(LTRIM(RTRIM(TC003))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TC003)) as date), 'yyyy-MM-dd') ELSE NULL END SoDate
                                , CASE WHEN LEN(LTRIM(RTRIM(TC039))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TC039)) as date), 'yyyy-MM-dd') ELSE NULL END DocDate
                                , LTRIM(RTRIM(TC015)) SoRemark, LTRIM(RTRIM(TC004)) CustomerNo, LTRIM(RTRIM(TC006)) SalesmenNo
                                , LTRIM(RTRIM(TC010)) CustomerAddressFirst, LTRIM(RTRIM(TC011)) CustomerAddressSecond, LTRIM(RTRIM(TC012)) CustomerPurchaseOrder
                                , LTRIM(RTRIM(TC070)) DepositPartial, LTRIM(RTRIM(TC045)) DepositRate, LTRIM(RTRIM(TC008)) Currency, LTRIM(RTRIM(TC009)) ExchangeRate
                                , LTRIM(RTRIM(TC078)) TaxNo, LTRIM(RTRIM(TC016)) Taxation, LTRIM(RTRIM(TC041)) BusinessTaxRate, LTRIM(RTRIM(TC091)) DetailMultiTax
                                , LTRIM(RTRIM(TC031)) TotalQty, LTRIM(RTRIM(TC029)) Amount, LTRIM(RTRIM(TC030)) TaxAmount, LTRIM(RTRIM(TC019)) ShipMethod
                                , LTRIM(RTRIM(TC068)) TradeTerm, LTRIM(RTRIM(TC042)) PaymentTerm, LTRIM(RTRIM(TC013)) PriceTerm, LTRIM(RTRIM(TC027)) ConfirmStatus
                                , LTRIM(RTRIM(TC040)) ConfirmUserNo
                                , CASE WHEN LEN(LTRIM(RTRIM(CREATE_DATE))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_DATE)) as date), 'yyyy-MM-dd') ELSE NULL END TransferDate
                                , CASE WHEN LEN(LTRIM(RTRIM(CREATE_TIME))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_TIME)) as datetime), 'HH:mm:ss') ELSE NULL END TransferTime
                                FROM COPTC
                                WHERE 1=1";
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "UpdateDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME > @UpdateDate OR MODI_DATE + ' ' + MODI_TIME > @UpdateDate)", UpdateDate);
                        sql += @" ORDER BY TransferDate, TransferTime";

                        saleOrders = sqlConnection.Query<SaleOrder>(sql, dynamicParameters).ToList();
                        #endregion

                        #region //撈取ERP訂單單身資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT LTRIM(RTRIM(TD001)) SoErpPrefix, LTRIM(RTRIM(TD002)) SoErpNo, LTRIM(RTRIM(TD003)) SoSequence, LTRIM(RTRIM(TD004)) MtlItemNo, LTRIM(RTRIM(TD005)) SoMtlItemName
                                , LTRIM(RTRIM(TD006)) SoMtlItemSpec, LTRIM(RTRIM(TD007)) InventoryNo, LTRIM(RTRIM(TD010)) UomNo
                                , LTRIM(RTRIM(CAST(TD008 AS INT))) SoQty, LTRIM(RTRIM(CAST(TD009 AS INT))) SiQty, LTRIM(RTRIM(CAST(TD024 AS INT))) FreebieQty
                                , LTRIM(RTRIM(CAST(TD025 AS INT))) FreebieSiQty, LTRIM(RTRIM(TD011)) UnitPrice, LTRIM(RTRIM(TD012)) Amount
                                , LTRIM(RTRIM(TD049)) ProductType, LTRIM(RTRIM(CAST(TD050 AS INT))) SpareQty, LTRIM(RTRIM(CAST(TD051 AS INT))) SpareSiQty
                                , CASE WHEN LEN(LTRIM(RTRIM(TD013))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TD013)) as date), 'yyyy-MM-dd') ELSE NULL END PromiseDate
                                , CASE WHEN LEN(LTRIM(RTRIM(TD048))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TD048)) as date), 'yyyy-MM-dd') ELSE NULL END PcPromiseDate
                                , LTRIM(RTRIM(TD027)) Project, LTRIM(RTRIM(TD014)) CustomerMtlItemNo
                                , LTRIM(RTRIM(TD020)) SoDetailRemark, LTRIM(RTRIM(CAST(TD076 AS INT))) SoPriceQty
                                , LTRIM(RTRIM(TD077)) SoPriceUomrNo, LTRIM(RTRIM(TD079)) TaxNo, LTRIM(RTRIM(TD070)) BusinessTaxRate
                                , LTRIM(RTRIM(TD026)) DiscountRate, LTRIM(RTRIM(TD080)) DiscountAmount, LTRIM(RTRIM(TD021)) ConfirmStatus
                                , LTRIM(RTRIM(TD016)) ClosureStatus , LTRIM(RTRIM(TD017)) + '-' + LTRIM(RTRIM(TD018))+ '-' + LTRIM(RTRIM(TD019)) QuotationErp
                                , CASE WHEN LEN(LTRIM(RTRIM(CREATE_DATE))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_DATE)) as date), 'yyyy-MM-dd') ELSE NULL END TransferDate
                                , CASE WHEN LEN(LTRIM(RTRIM(CREATE_TIME))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_TIME)) as datetime), 'HH:mm:ss') ELSE NULL END TransferTime
                                FROM COPTD
                                WHERE 1=1 AND TD010!='pcs' AND TD077!='pcs'";
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "UpdateDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME > @UpdateDate OR MODI_DATE + ' ' + MODI_TIME > @UpdateDate)", UpdateDate);
                        sql += @" ORDER BY TransferDate, TransferTime";

                        soDetails = sqlConnection.Query<SoDetail>(sql, dynamicParameters).ToList();
                        #endregion
                    }

                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //撈取部門ID
                        sql = @"SELECT DepartmentId, DepartmentNo
                                FROM BAS.Department
                                WHERE [Status]='A' AND CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<Department> resultDepartments = sqlConnection.Query<Department>(sql, dynamicParameters).ToList();

                        saleOrders = saleOrders.Join(resultDepartments, x => x.DepartmentNo, y => y.DepartmentNo, (x, y) => { x.DepartmentId = y.DepartmentId; return x; }).ToList();
                        #endregion

                        #region //撈取客戶ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT CustomerId, CustomerNo 
                                FROM SCM.Customer
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<Customer> resultCustomers = sqlConnection.Query<Customer>(sql, dynamicParameters).ToList();

                        saleOrders = saleOrders.Join(resultCustomers, x => x.CustomerNo, y => y.CustomerNo, (x, y) => { x.CustomerId = y.CustomerId; return x; }).ToList();
                        #endregion

                        #region //撈取業務人員ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserId, a.UserNo 
                                FROM BAS.[User] a";

                        List<User> resultSalesmens = sqlConnection.Query<User>(sql, dynamicParameters).ToList();

                        saleOrders = saleOrders.Join(resultSalesmens, x => x.SalesmenNo, y => y.UserNo, (x, y) => { x.SalesmenId = y.UserId; return x; }).ToList();
                        #endregion

                        #region //撈取確認者ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserId, a.UserNo 
                                FROM BAS.[User] a";

                        List<User> resultConfirmUsers = sqlConnection.Query<User>(sql, dynamicParameters).ToList();

                        saleOrders = saleOrders.Join(resultConfirmUsers, x => x.ConfirmUserNo, y => y.UserNo, (x, y) => { x.ConfirmUserId = y.UserId; return x; }).ToList();
                        #endregion

                        #region //撈取單位ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT UomId, UomNo
                                FROM PDM.UnitOfMeasure
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<Uom> resultSoPriceUomrNos = sqlConnection.Query<Uom>(sql, dynamicParameters).ToList();

                        soDetails = soDetails.Join(resultSoPriceUomrNos, x => x.SoPriceUomrNo.ToUpper(), y => y.UomNo, (x, y) => { x.SoPriceUomId = y.UomId; return x; }).ToList();
                        soDetails = soDetails.Join(resultSoPriceUomrNos, x => x.UomNo.ToUpper(), y => y.UomNo, (x, y) => { x.UomId = y.UomId; return x; }).ToList();
                        #endregion

                        #region //撈取品號ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT MtlItemId, MtlItemNo
                                FROM PDM.MtlItem
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<MtlItem> resultMtlitems = sqlConnection.Query<MtlItem>(sql, dynamicParameters).ToList();

                        soDetails = soDetails.Join(resultMtlitems, x => x.MtlItemNo, y => y.MtlItemNo, (x, y) => { x.MtlItemId = y.MtlItemId; return x; }).Select(x => x).ToList();
                        #endregion

                        #region //撈取庫別ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT InventoryId, InventoryNo
                                FROM SCM.Inventory
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<Inventory> resultInventories = sqlConnection.Query<Inventory>(sql, dynamicParameters).ToList();

                        soDetails = soDetails.Join(resultInventories, x => x.InventoryNo, y => y.InventoryNo, (x, y) => { x.InventoryId = y.InventoryId; return x; }).ToList();
                        #endregion

                        #region //判斷SCM.SaleOrder是否有資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT SoId, SoErpPrefix, SoErpNo
                                FROM SCM.SaleOrder
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<SaleOrder> resultSaleOrder = sqlConnection.Query<SaleOrder>(sql, dynamicParameters).ToList();

                        saleOrders = saleOrders.GroupJoin(resultSaleOrder, x => new { x.SoErpPrefix, x.SoErpNo }, y => new { y.SoErpPrefix, y.SoErpNo }, (x, y) => { x.SoId = y.FirstOrDefault()?.SoId; return x; }).Select(x => x).ToList();
                        #endregion

                        #region //訂單單頭(新增/修改)
                        List<SaleOrder> addsaleOrders = saleOrders.Where(x => x.SoId == null).ToList();
                        List<SaleOrder> updatesaleOrders = saleOrders.Where(x => x.SoId != null).ToList();

                        #region //新增
                        dynamicParameters = new DynamicParameters();
                        if (addsaleOrders.Count > 0)
                        {
                            addsaleOrders
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.CompanyId = CompanyId;
                                    x.TransferStatus = "Y";
                                    x.CreateDate = CreateDate;
                                    x.LastModifiedDate = LastModifiedDate;
                                    x.CreateBy = CreateBy;
                                    x.LastModifiedBy = LastModifiedBy;
                                });

                            sql = @"INSERT INTO SCM.SaleOrder (CompanyId, DepartmentId, SoErpPrefix, SoErpNo, Version, SoDate, DocDate
                                    , SoRemark, CustomerId, SalesmenId, CustomerAddressFirst, CustomerAddressSecond, CustomerPurchaseOrder
                                    , DepositPartial, DepositRate, Currency, ExchangeRate, TaxNo, Taxation, BusinessTaxRate
                                    , DetailMultiTax, TotalQty, Amount, TaxAmount, ShipMethod, TradeTerm, PaymentTerm, PriceTerm
                                    , ConfirmStatus, ConfirmUserId, TransferStatus
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@CompanyId, @DepartmentId, @SoErpPrefix, @SoErpNo, @Version, @SoDate, @DocDate, @SoRemark
                                    , @CustomerId, @SalesmenId, @CustomerAddressFirst, @CustomerAddressSecond, @CustomerPurchaseOrder
                                    , @DepositPartial, @DepositRate, @Currency, @ExchangeRate, @TaxNo, @Taxation, @BusinessTaxRate
                                    , @DetailMultiTax, @TotalQty, @Amount, @TaxAmount, @ShipMethod, @TradeTerm, @PaymentTerm, @PriceTerm
                                    , @ConfirmStatus, @ConfirmUserId, @TransferStatus
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            rowsAffected += sqlConnection.Execute(sql, addsaleOrders);
                        }
                        #endregion

                        #region //修改
                        dynamicParameters = new DynamicParameters();
                        if (updatesaleOrders.Count > 0)
                        {
                            updatesaleOrders
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.LastModifiedDate = LastModifiedDate;
                                    x.LastModifiedBy = LastModifiedBy;
                                });

                            sql = @"UPDATE SCM.SaleOrder SET
                                    DepartmentId = @DepartmentId,
                                    SoErpPrefix = @SoErpPrefix,
                                    SoErpNo = @SoErpNo,
                                    Version = @Version,
                                    SoDate = @SoDate,
                                    DocDate = @DocDate,
                                    SoRemark = @SoRemark,
                                    CustomerId = @CustomerId,
                                    SalesmenId = @SalesmenId,
                                    CustomerAddressFirst = @CustomerAddressFirst,
                                    CustomerAddressSecond = @CustomerAddressSecond,
                                    CustomerPurchaseOrder = @CustomerPurchaseOrder,
                                    DepositPartial = @DepositPartial,
                                    DepositRate = @DepositRate,
                                    Currency = @Currency,
                                    ExchangeRate = @ExchangeRate,
                                    TaxNo = @TaxNo,
                                    Taxation = @Taxation,
                                    BusinessTaxRate = @BusinessTaxRate,
                                    DetailMultiTax = @DetailMultiTax,
                                    TotalQty = @TotalQty,
                                    Amount = @Amount,
                                    TaxAmount = @TaxAmount,
                                    ShipMethod = @ShipMethod,
                                    TradeTerm = @TradeTerm,
                                    PaymentTerm = @PaymentTerm,
                                    PriceTerm = @PriceTerm,
                                    ConfirmStatus = @ConfirmStatus,
                                    ConfirmUserId = @ConfirmUserId,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE SoId = @SoId";
                            rowsAffected += sqlConnection.Execute(sql, updatesaleOrders);
                        }
                        #endregion
                        #endregion

                        #region //訂單單身(新增/修改)
                        #region //撈取訂單單頭ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT SoId, SoErpPrefix, SoErpNo
                                FROM  SCM.SaleOrder
                                WHERE CompanyId = @CompanyId
                                ORDER BY SoId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        resultSaleOrder = sqlConnection.Query<SaleOrder>(sql, dynamicParameters).ToList();

                        soDetails = soDetails.Join(resultSaleOrder, x => new { x.SoErpPrefix, x.SoErpNo }, y => new { y.SoErpPrefix, y.SoErpNo }, (x, y) => { x.SoId = y.SoId; return x; }).Select(x => x).ToList();
                        #endregion

                        #region //判斷SCM.SoDetail是否有資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.SoDetailId, a.SoId, a.SoSequence
                                FROM SCM.SoDetail a
                                INNER JOIN SCM.SaleOrder b ON a.SoId = b.SoId
                                WHERE b.CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<SoDetail> resultSoDetails = sqlConnection.Query<SoDetail>(sql, dynamicParameters).ToList();

                        soDetails = soDetails.GroupJoin(resultSoDetails, x => new { x.SoId, x.SoSequence }, y => new { y.SoId, y.SoSequence }, (x, y) => { x.SoDetailId = y.FirstOrDefault()?.SoDetailId; return x; }).ToList();
                        #endregion

                        List<SoDetail> addSoDetails = soDetails.Where(x => x.SoDetailId == null).ToList();
                        List<SoDetail> updateSoDetails = soDetails.Where(x => x.SoDetailId != null).ToList();

                        #region //新增
                        dynamicParameters = new DynamicParameters();
                        if (addSoDetails.Count > 0)
                        {
                            addSoDetails
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.CreateDate = CreateDate;
                                    x.LastModifiedDate = LastModifiedDate;
                                    x.CreateBy = CreateBy;
                                    x.LastModifiedBy = LastModifiedBy;
                                });

                            sql = @"INSERT INTO SCM.SoDetail (SoId, SoSequence, MtlItemId, SoMtlItemName, SoMtlItemSpec
                                    , InventoryId, UomId, SoQty, SiQty, FreebieQty, FreebieSiQty, ProductType, SpareQty, SpareSiQty
                                    , UnitPrice, Amount, PromiseDate, PcPromiseDate, Project, CustomerMtlItemNo
                                    , SoDetailRemark, SoPriceQty, SoPriceUomId, TaxNo, BusinessTaxRate
                                    , DiscountRate, DiscountAmount, ConfirmStatus, ClosureStatus, QuotationErp
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@SoId, @SoSequence, @MtlItemId, @SoMtlItemName, @SoMtlItemSpec
                                    , @InventoryId, @UomId, @SoQty, @SiQty, @FreebieQty, @FreebieSiQty, @ProductType, @SpareQty, @SpareSiQty
                                    , @UnitPrice, @Amount, @PromiseDate, @PcPromiseDate, @Project, @CustomerMtlItemNo
                                    , @SoDetailRemark, @SoPriceQty, @SoPriceUomId, @TaxNo, @BusinessTaxRate
                                    , @DiscountRate, @DiscountAmount, @ConfirmStatus, @ClosureStatus, @QuotationErp
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            rowsAffected += sqlConnection.Execute(sql, addSoDetails);
                        }
                        #endregion

                        #region //修改
                        dynamicParameters = new DynamicParameters();
                        if (updateSoDetails.Count > 0)
                        {
                            updateSoDetails
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.LastModifiedDate = LastModifiedDate;
                                    x.LastModifiedBy = LastModifiedBy;
                                });

                            sql = @"UPDATE SCM.SoDetail SET
                                    MtlItemId = @MtlItemId, 
                                    SoMtlItemName = @SoMtlItemName, 
                                    SoMtlItemSpec = @SoMtlItemSpec, 
                                    InventoryId = @InventoryId, 
                                    UomId = @UomId, 
                                    SoQty = @SoQty, 
                                    FreebieQty = @FreebieQty, 
                                    ProductType = @ProductType, 
                                    SpareQty = @SpareQty, 
                                    UnitPrice = @UnitPrice, 
                                    Amount = @Amount, 
                                    PromiseDate = @PromiseDate, 
                                    PcPromiseDate = @PcPromiseDate, 
                                    Project = @Project, 
                                    CustomerMtlItemNo = @CustomerMtlItemNo, 
                                    SoDetailRemark = @SoDetailRemark, 
                                    SoPriceQty = @SoPriceQty, 
                                    SoPriceUomId = @SoPriceUomId, 
                                    TaxNo = @TaxNo, 
                                    BusinessTaxRate = @BusinessTaxRate, 
                                    DiscountRate = @DiscountRate, 
                                    DiscountAmount = @DiscountAmount, 
                                    ConfirmStatus = @ConfirmStatus, 
                                    ClosureStatus = @ClosureStatus, 
                                    QuotationErp = @QuotationErp, 
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE SoDetailId = @SoDetailId";
                            rowsAffected += sqlConnection.Execute(sql, updateSoDetails);
                        }
                        #endregion
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

        #region //UpdateSaleOrderManualSynchronize -- 訂單資料手動同步 -- Ben Ma 2023.06.07
        public string UpdateSaleOrderManualSynchronize(string SoErpFullNo, string SyncStartDate, string SyncEndDate
            , string NormalSync, string TranSync)
        {
            try
            {
                List<SaleOrder> saleOrders = new List<SaleOrder>();
                List<SoDetail> soDetails = new List<SoDetail>();

                int mainAffected = 0, detailAffected = 0, mainDelAffected = 0, detailDelAffected = 0, tempAffected = 0;
                string companyNo = "";

                using (TransactionScope transactionScope = new TransactionScope(TransactionScopeOption.Required, new TimeSpan(0, 180, 0)))
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //公司別資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var resultCompany = sqlConnection.Query(sql, dynamicParameters);
                        if (resultCompany.Count() <= 0) throw new SystemException("公司別資料錯誤!");

                        foreach (var item in resultCompany)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            companyNo = item.ErpNo;
                        }
                        #endregion
                    }

                    #region //正常同步
                    if (NormalSync == "Y")
                    {
                        using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //判斷ERP訂單資料是否存在
                            if (SoErpFullNo.Length > 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM COPTC
                                        WHERE (LTRIM(RTRIM(TC001)) + '-' + LTRIM(RTRIM(TC002))) LIKE '%' + @SoErpFullNo + '%'";
                                dynamicParameters.Add("SoErpFullNo", SoErpFullNo);

                                var resultTsn = sqlConnection.Query(sql, dynamicParameters);
                                if (resultTsn.Count() <= 0) throw new SystemException("ERP訂單資料錯誤");
                            }
                            #endregion

                            #region //撈取ERP訂單單頭資料
                            sql = @"SELECT LTRIM(RTRIM(TC005)) DepartmentNo, LTRIM(RTRIM(TC001)) SoErpPrefix, LTRIM(RTRIM(TC002)) SoErpNo, LTRIM(RTRIM(TC069)) Version
                                    , CASE WHEN LEN(LTRIM(RTRIM(TC003))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TC003)) as date), 'yyyy-MM-dd') ELSE NULL END SoDate
                                    , CASE WHEN LEN(LTRIM(RTRIM(TC039))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TC039)) as date), 'yyyy-MM-dd') ELSE NULL END DocDate
                                    , LTRIM(RTRIM(TC015)) SoRemark, LTRIM(RTRIM(TC004)) CustomerNo, LTRIM(RTRIM(TC006)) SalesmenNo
                                    , LTRIM(RTRIM(TC010)) CustomerAddressFirst, LTRIM(RTRIM(TC011)) CustomerAddressSecond, LTRIM(RTRIM(TC012)) CustomerPurchaseOrder
                                    , LTRIM(RTRIM(TC070)) DepositPartial, LTRIM(RTRIM(TC045)) DepositRate, LTRIM(RTRIM(TC008)) Currency, LTRIM(RTRIM(TC009)) ExchangeRate
                                    , LTRIM(RTRIM(TC078)) TaxNo, LTRIM(RTRIM(TC016)) Taxation, LTRIM(RTRIM(TC041)) BusinessTaxRate, LTRIM(RTRIM(TC091)) DetailMultiTax
                                    , LTRIM(RTRIM(TC031)) TotalQty, LTRIM(RTRIM(TC029)) Amount, LTRIM(RTRIM(TC030)) TaxAmount, LTRIM(RTRIM(TC019)) ShipMethod
                                    , LTRIM(RTRIM(TC068)) TradeTerm, LTRIM(RTRIM(TC042)) PaymentTerm, LTRIM(RTRIM(TC013)) PriceTerm, LTRIM(RTRIM(TC027)) ConfirmStatus
                                    , LTRIM(RTRIM(TC040)) ConfirmUserNo
                                    , CASE WHEN LEN(LTRIM(RTRIM(CREATE_DATE))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_DATE)) as date), 'yyyy-MM-dd') ELSE NULL END TransferDate
                                    , CASE WHEN LEN(LTRIM(RTRIM(CREATE_TIME))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_TIME)) as datetime), 'HH:mm:ss') ELSE NULL END TransferTime
                                    FROM COPTC
                                    WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SoErpFullNo", @" AND (LTRIM(RTRIM(TC001)) + '-' + LTRIM(RTRIM(TC002))) LIKE '%' + @SoErpFullNo + '%'", SoErpFullNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SyncStartDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME >= @SyncStartDate OR MODI_DATE + ' ' + MODI_TIME >= @SyncStartDate)", SyncStartDate.Length > 0 ? Convert.ToDateTime(SyncStartDate).ToString("yyyyMMdd 00:00:00") : "");
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SyncEndDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME <= @SyncEndDate OR MODI_DATE + ' ' + MODI_TIME <= @SyncEndDate)", SyncEndDate.Length > 0 ? Convert.ToDateTime(SyncEndDate).ToString("yyyyMMdd 23:59:59") : ""); sql += @" ORDER BY TransferDate, TransferTime";

                            saleOrders = sqlConnection.Query<SaleOrder>(sql, dynamicParameters).ToList();
                            #endregion

                            #region //撈取ERP訂單單身資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(TD001)) SoErpPrefix, LTRIM(RTRIM(TD002)) SoErpNo, LTRIM(RTRIM(TD003)) SoSequence, LTRIM(RTRIM(TD004)) MtlItemNo, LTRIM(RTRIM(TD005)) SoMtlItemName
                                    , LTRIM(RTRIM(TD006)) SoMtlItemSpec, LTRIM(RTRIM(TD007)) InventoryNo, LTRIM(RTRIM(TD010)) UomNo
                                    , LTRIM(RTRIM(CAST(TD008 AS INT))) SoQty, LTRIM(RTRIM(CAST(TD009 AS INT))) SiQty, LTRIM(RTRIM(CAST(TD024 AS INT))) FreebieQty
                                    , LTRIM(RTRIM(CAST(TD025 AS INT))) FreebieSiQty, LTRIM(RTRIM(TD011)) UnitPrice, LTRIM(RTRIM(TD012)) Amount
                                    , LTRIM(RTRIM(TD049)) ProductType, LTRIM(RTRIM(CAST(TD050 AS INT))) SpareQty, LTRIM(RTRIM(CAST(TD051 AS INT))) SpareSiQty
                                    , CASE WHEN LEN(LTRIM(RTRIM(TD013))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TD013)) as date), 'yyyy-MM-dd') ELSE NULL END PromiseDate
                                    , CASE WHEN LEN(LTRIM(RTRIM(TD048))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TD048)) as date), 'yyyy-MM-dd') ELSE NULL END PcPromiseDate
                                    , LTRIM(RTRIM(TD027)) Project, LTRIM(RTRIM(TD014)) CustomerMtlItemNo
                                    , LTRIM(RTRIM(TD020)) SoDetailRemark, LTRIM(RTRIM(CAST(TD076 AS INT))) SoPriceQty
                                    , LTRIM(RTRIM(TD077)) SoPriceUomrNo, LTRIM(RTRIM(TD079)) TaxNo, LTRIM(RTRIM(TD070)) BusinessTaxRate
                                    , LTRIM(RTRIM(TD026)) DiscountRate, LTRIM(RTRIM(TD080)) DiscountAmount, LTRIM(RTRIM(TD021)) ConfirmStatus
                                    , LTRIM(RTRIM(TD016)) ClosureStatus, LTRIM(RTRIM(TD017)) + '-' + LTRIM(RTRIM(TD018))+ '-' + LTRIM(RTRIM(TD019)) QuotationErp
                                    , CASE WHEN LEN(LTRIM(RTRIM(CREATE_DATE))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_DATE)) as date), 'yyyy-MM-dd') ELSE NULL END TransferDate
                                    , CASE WHEN LEN(LTRIM(RTRIM(CREATE_TIME))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_TIME)) as datetime), 'HH:mm:ss') ELSE NULL END TransferTime
                                    FROM COPTD
                                    WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SoErpFullNo", @" AND (LTRIM(RTRIM(TD001)) + '-' + LTRIM(RTRIM(TD002))) LIKE '%' + @SoErpFullNo + '%'", SoErpFullNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SyncStartDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME >= @SyncStartDate OR MODI_DATE + ' ' + MODI_TIME >= @SyncStartDate)", SyncStartDate.Length > 0 ? Convert.ToDateTime(SyncStartDate).ToString("yyyyMMdd 00:00:00") : "");
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SyncEndDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME <= @SyncEndDate OR MODI_DATE + ' ' + MODI_TIME <= @SyncEndDate)", SyncEndDate.Length > 0 ? Convert.ToDateTime(SyncEndDate).ToString("yyyyMMdd 23:59:59") : "");
                            sql += @" ORDER BY TransferDate, TransferTime";

                            soDetails = sqlConnection.Query<SoDetail>(sql, dynamicParameters).ToList();
                            #endregion
                        }

                        using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                        {
                            #region //撈取部門ID
                            sql = @"SELECT DepartmentId, DepartmentNo
                                    FROM BAS.Department
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<Department> resultDepartments = sqlConnection.Query<Department>(sql, dynamicParameters).ToList();

                            saleOrders = saleOrders.Join(resultDepartments, x => x.DepartmentNo, y => y.DepartmentNo, (x, y) => { x.DepartmentId = y.DepartmentId; return x; }).ToList();
                            #endregion

                            #region //撈取客戶ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT CustomerId, CustomerNo 
                                    FROM SCM.Customer
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<Customer> resultCustomers = sqlConnection.Query<Customer>(sql, dynamicParameters).ToList();

                            saleOrders = saleOrders.Join(resultCustomers, x => x.CustomerNo, y => y.CustomerNo, (x, y) => { x.CustomerId = y.CustomerId; return x; }).ToList();
                            #endregion

                            #region //撈取業務人員ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.UserId, a.UserNo 
                                    FROM BAS.[User] a";

                            List<User> resultSalesmens = sqlConnection.Query<User>(sql, dynamicParameters).ToList();

                            saleOrders = saleOrders.Join(resultSalesmens, x => x.SalesmenNo, y => y.UserNo, (x, y) => { x.SalesmenId = y.UserId; return x; }).ToList();
                            #endregion

                            #region //撈取確認者ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.UserId, a.UserNo 
                                    FROM BAS.[User] a";

                            List<User> resultConfirmUsers = sqlConnection.Query<User>(sql, dynamicParameters).ToList();

                            saleOrders = saleOrders.Join(resultConfirmUsers, x => x.ConfirmUserNo, y => y.UserNo, (x, y) => { x.ConfirmUserId = y.UserId; return x; }).ToList();
                            #endregion

                            #region //撈取單位ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT UomId, UomNo
                                    FROM PDM.UnitOfMeasure
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<Uom> resultSoPriceUomrNos = sqlConnection.Query<Uom>(sql, dynamicParameters).ToList();

                            soDetails = soDetails.Join(resultSoPriceUomrNos, x => x.SoPriceUomrNo.ToUpper(), y => y.UomNo, (x, y) => { x.SoPriceUomId = y.UomId; return x; }).ToList();
                            soDetails = soDetails.Join(resultSoPriceUomrNos, x => x.UomNo.ToUpper(), y => y.UomNo, (x, y) => { x.UomId = y.UomId; return x; }).ToList();
                            #endregion

                            #region //撈取品號ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT MtlItemId, MtlItemNo
                                    FROM PDM.MtlItem
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<MtlItem> resultMtlitems = sqlConnection.Query<MtlItem>(sql, dynamicParameters).ToList();

                            soDetails = soDetails.Join(resultMtlitems, x => x.MtlItemNo, y => y.MtlItemNo, (x, y) => { x.MtlItemId = y.MtlItemId; return x; }).Select(x => x).ToList();
                            #endregion

                            #region //撈取庫別ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT InventoryId, InventoryNo
                                    FROM SCM.Inventory
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<Inventory> resultInventories = sqlConnection.Query<Inventory>(sql, dynamicParameters).ToList();

                            soDetails = soDetails.Join(resultInventories, x => x.InventoryNo, y => y.InventoryNo, (x, y) => { x.InventoryId = y.InventoryId; return x; }).ToList();
                            #endregion

                            #region //判斷SCM.SaleOrder是否有資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT SoId, SoErpPrefix, SoErpNo
                                    FROM SCM.SaleOrder
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<SaleOrder> resultSaleOrder = sqlConnection.Query<SaleOrder>(sql, dynamicParameters).ToList();

                            saleOrders = saleOrders.GroupJoin(resultSaleOrder, x => new { x.SoErpPrefix, x.SoErpNo }, y => new { y.SoErpPrefix, y.SoErpNo }, (x, y) => { x.SoId = y.FirstOrDefault()?.SoId; return x; }).Select(x => x).ToList();
                            #endregion

                            #region //訂單單頭(新增/修改)
                            List<SaleOrder> addsaleOrders = saleOrders.Where(x => x.SoId == null).ToList();
                            List<SaleOrder> updatesaleOrders = saleOrders.Where(x => x.SoId != null).ToList();

                            #region //新增
                            dynamicParameters = new DynamicParameters();
                            if (addsaleOrders.Count > 0)
                            {
                                addsaleOrders
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.CompanyId = CurrentCompany;
                                        x.TransferStatus = "Y";
                                        x.CreateDate = CreateDate;
                                        x.LastModifiedDate = LastModifiedDate;
                                        x.CreateBy = CreateBy;
                                        x.LastModifiedBy = LastModifiedBy;
                                    });

                                sql = @"INSERT INTO SCM.SaleOrder (CompanyId, DepartmentId, SoErpPrefix, SoErpNo, Version, SoDate, DocDate
                                        , SoRemark, CustomerId, SalesmenId, CustomerAddressFirst, CustomerAddressSecond, CustomerPurchaseOrder
                                        , DepositPartial, DepositRate, Currency, ExchangeRate, TaxNo, Taxation, BusinessTaxRate
                                        , DetailMultiTax, TotalQty, Amount, TaxAmount, ShipMethod, TradeTerm, PaymentTerm, PriceTerm
                                        , ConfirmStatus, ConfirmUserId, TransferStatus
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        VALUES (@CompanyId, @DepartmentId, @SoErpPrefix, @SoErpNo, @Version, @SoDate, @DocDate, @SoRemark
                                        , @CustomerId, @SalesmenId, @CustomerAddressFirst, @CustomerAddressSecond, @CustomerPurchaseOrder
                                        , @DepositPartial, @DepositRate, @Currency, @ExchangeRate, @TaxNo, @Taxation, @BusinessTaxRate
                                        , @DetailMultiTax, @TotalQty, @Amount, @TaxAmount, @ShipMethod, @TradeTerm, @PaymentTerm, @PriceTerm
                                        , @ConfirmStatus, @ConfirmUserId, @TransferStatus
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                int addMain = sqlConnection.Execute(sql, addsaleOrders);
                                mainAffected += addMain;
                            }
                            #endregion

                            #region //修改
                            dynamicParameters = new DynamicParameters();
                            if (updatesaleOrders.Count > 0)
                            {
                                updatesaleOrders
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.LastModifiedDate = LastModifiedDate;
                                        x.LastModifiedBy = LastModifiedBy;
                                    });

                                sql = @"UPDATE SCM.SaleOrder SET
                                        DepartmentId = @DepartmentId,
                                        SoErpPrefix = @SoErpPrefix,
                                        SoErpNo = @SoErpNo,
                                        Version = @Version,
                                        SoDate = @SoDate,
                                        DocDate = @DocDate,
                                        SoRemark = @SoRemark,
                                        CustomerId = @CustomerId,
                                        SalesmenId = @SalesmenId,
                                        CustomerAddressFirst = @CustomerAddressFirst,
                                        CustomerAddressSecond = @CustomerAddressSecond,
                                        CustomerPurchaseOrder = @CustomerPurchaseOrder,
                                        DepositPartial = @DepositPartial,
                                        DepositRate = @DepositRate,
                                        Currency = @Currency,
                                        ExchangeRate = @ExchangeRate,
                                        TaxNo = @TaxNo,
                                        Taxation = @Taxation,
                                        BusinessTaxRate = @BusinessTaxRate,
                                        DetailMultiTax = @DetailMultiTax,
                                        TotalQty = @TotalQty,
                                        Amount = @Amount,
                                        TaxAmount = @TaxAmount,
                                        ShipMethod = @ShipMethod,
                                        TradeTerm = @TradeTerm,
                                        PaymentTerm = @PaymentTerm,
                                        PriceTerm = @PriceTerm,
                                        ConfirmStatus = @ConfirmStatus,
                                        ConfirmUserId = @ConfirmUserId,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE SoId = @SoId";
                                int updateMain = sqlConnection.Execute(sql, updatesaleOrders);
                                mainAffected += updateMain;
                            }
                            #endregion
                            #endregion

                            #region //訂單單身(新增/修改)
                            #region //撈取訂單單頭ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT SoId, SoErpPrefix, SoErpNo
                                    FROM SCM.SaleOrder
                                    WHERE CompanyId = @CompanyId
                                    ORDER BY SoId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            resultSaleOrder = sqlConnection.Query<SaleOrder>(sql, dynamicParameters).ToList();

                            soDetails = soDetails.Join(resultSaleOrder, x => new { x.SoErpPrefix, x.SoErpNo }, y => new { y.SoErpPrefix, y.SoErpNo }, (x, y) => { x.SoId = y.SoId; return x; }).Select(x => x).ToList();
                            #endregion

                            #region //判斷SCM.SoDetail是否有資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.SoDetailId, a.SoId, a.SoSequence
                                    FROM SCM.SoDetail a
                                    INNER JOIN SCM.SaleOrder b ON a.SoId = b.SoId
                                    WHERE b.CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<SoDetail> resultSoDetails = sqlConnection.Query<SoDetail>(sql, dynamicParameters).ToList();

                            soDetails = soDetails.GroupJoin(resultSoDetails, x => new { x.SoId, x.SoSequence }, y => new { y.SoId, y.SoSequence }, (x, y) => { x.SoDetailId = y.FirstOrDefault()?.SoDetailId; return x; }).ToList();
                            #endregion

                            List<SoDetail> addSoDetails = soDetails.Where(x => x.SoDetailId == null).ToList();
                            List<SoDetail> updateSoDetails = soDetails.Where(x => x.SoDetailId != null).ToList();

                            #region //新增
                            dynamicParameters = new DynamicParameters();
                            if (addSoDetails.Count > 0)
                            {
                                addSoDetails
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.CreateDate = CreateDate;
                                        x.LastModifiedDate = LastModifiedDate;
                                        x.CreateBy = CreateBy;
                                        x.LastModifiedBy = LastModifiedBy;
                                    });

                                sql = @"INSERT INTO SCM.SoDetail (SoId, SoSequence, MtlItemId, SoMtlItemName, SoMtlItemSpec
                                        , InventoryId, UomId, SoQty, SiQty, FreebieQty, FreebieSiQty, ProductType, SpareQty, SpareSiQty
                                        , UnitPrice, Amount, PromiseDate, PcPromiseDate, Project, CustomerMtlItemNo
                                        , SoDetailRemark, SoPriceQty, SoPriceUomId, TaxNo, BusinessTaxRate
                                        , DiscountRate, DiscountAmount, ConfirmStatus, ClosureStatus, QuotationErp
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        VALUES (@SoId, @SoSequence, @MtlItemId, @SoMtlItemName, @SoMtlItemSpec
                                        , @InventoryId, @UomId, @SoQty, @SiQty, @FreebieQty, @FreebieSiQty, @ProductType, @SpareQty, @SpareSiQty
                                        , @UnitPrice, @Amount, @PromiseDate, @PcPromiseDate, @Project, @CustomerMtlItemNo
                                        , @SoDetailRemark, @SoPriceQty, @SoPriceUomId, @TaxNo, @BusinessTaxRate
                                        , @DiscountRate, @DiscountAmount, @ConfirmStatus, @ClosureStatus, @QuotationErp
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                int addDetail = sqlConnection.Execute(sql, addSoDetails);
                                detailAffected += addDetail;
                            }
                            #endregion

                            #region //修改
                            dynamicParameters = new DynamicParameters();
                            if (updateSoDetails.Count > 0)
                            {
                                updateSoDetails
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.LastModifiedDate = LastModifiedDate;
                                        x.LastModifiedBy = LastModifiedBy;
                                    });

                                sql = @"UPDATE SCM.SoDetail SET
                                        MtlItemId = @MtlItemId, 
                                        SoMtlItemName = @SoMtlItemName, 
                                        SoMtlItemSpec = @SoMtlItemSpec, 
                                        InventoryId = @InventoryId, 
                                        UomId = @UomId, 
                                        SoQty = @SoQty, 
                                        SiQty = @SiQty, 
                                        FreebieQty = @FreebieQty, 
                                        ProductType = @ProductType, 
                                        SpareQty = @SpareQty, 
                                        UnitPrice = @UnitPrice, 
                                        Amount = @Amount, 
                                        PromiseDate = @PromiseDate, 
                                        PcPromiseDate = @PcPromiseDate, 
                                        Project = @Project, 
                                        CustomerMtlItemNo = @CustomerMtlItemNo, 
                                        SoDetailRemark = @SoDetailRemark, 
                                        SoPriceQty = @SoPriceQty, 
                                        SoPriceUomId = @SoPriceUomId, 
                                        TaxNo = @TaxNo, 
                                        BusinessTaxRate = @BusinessTaxRate, 
                                        DiscountRate = @DiscountRate, 
                                        DiscountAmount = @DiscountAmount, 
                                        ConfirmStatus = @ConfirmStatus, 
                                        ClosureStatus = @ClosureStatus, 
                                        QuotationErp = @QuotationErp, 
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE SoDetailId = @SoDetailId";
                                int updateDetail = sqlConnection.Execute(sql, updateSoDetails);
                                detailAffected += updateDetail;
                            }
                            #endregion
                            #endregion
                        }
                    }
                    #endregion

                    #region //異動同步
                    if (TranSync == "Y")
                    {
                        using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //ERP訂單單頭資料
                            sql = @"SELECT TC001 SoErpPrefix, TC002 SoErpNo
                                    FROM COPTC
                                    WHERE 1=1
                                    ORDER BY TC001, TC002";
                            var resultErpSo = erpConnection.Query<SaleOrder>(sql, dynamicParameters);
                            #endregion

                            using (SqlConnection bmConnection = new SqlConnection(MainConnectionStrings))
                            {
                                #region //BM訂單單頭資料
                                sql = @"SELECT SoErpPrefix, SoErpNo
                                        FROM SCM.SaleOrder
                                        WHERE CompanyId = @CompanyId
                                        ORDER BY SoErpPrefix, SoErpNo";
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                var resultBmIt = bmConnection.Query<SaleOrder>(sql, dynamicParameters);
                                #endregion

                                #region //篩選ERP不存在但BM還存在的訂單單頭
                                var dictionaryErpSo = resultErpSo.ToDictionary(x => x.SoErpPrefix + "-" + x.SoErpNo, x => x.SoErpPrefix + "-" + x.SoErpNo);
                                var dictionaryBmSo = resultBmIt.ToDictionary(x => x.SoErpPrefix + "-" + x.SoErpNo, x => x.SoErpPrefix + "-" + x.SoErpNo);
                                var changeSo = dictionaryBmSo.Where(x => !dictionaryErpSo.ContainsKey(x.Key)).ToList();
                                var changeSoList = changeSo.Select(x => x.Value).ToList();
                                #endregion

                                #region //刪除異動訂單單頭
                                if (changeSoList.Count > 0)
                                {
                                    #region //取得單身資料
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT SoDetailId
                                            FROM SCM.SoDetail
                                            WHERE SoId IN (
                                                SELECT SoId
                                                FROM SCM.SaleOrder
                                                WHERE SoErpPrefix + '-' + SoErpNo IN @SoErpFullNo
                                            )";
                                    dynamicParameters.Add("SoErpFullNo", changeSoList.Select(x => x).ToArray());
                                    var resultSoDetail = bmConnection.Query(sql, dynamicParameters);
                                    #endregion

                                    #region //更新綁定的暫存檔 (停用)
                                    //dynamicParameters = new DynamicParameters();
                                    //sql = @"UPDATE SCM.SoDetailTemp SET 
                                    //        SoDetailId = NULL,
                                    //        Status = @Status,
                                    //        LastModifiedDate = @LastModifiedDate,
                                    //        LastModifiedBy = @LastModifiedBy
                                    //        WHERE SoDetailId IN @SoDetailIds";
                                    //dynamicParameters.AddDynamicParams(
                                    //    new
                                    //    {
                                    //        SoDetailIds = resultSoDetail.Select(s => s.SoDetailId).ToArray(),
                                    //        Status = "N", //解除訂單身綁定
                                    //        LastModifiedDate,
                                    //        LastModifiedBy
                                    //    });
                                    //mainAffected += bmConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //刪除綁定的暫存檔
                                    sql = @"DELETE a 
                                            FROM SCM.SoDetailTemp a
                                            WHERE EXISTS (
	                                            SELECT TOP 1 1 
	                                            FROM SCM.SaleOrder aa 
	                                            WHERE aa.SoId = a.SoId 
	                                            AND aa.SoErpPrefix + '-' + aa.SoErpNo IN @SoErpFullNo)";
                                    tempAffected += bmConnection.Execute(sql, new { SoErpFullNo = changeSoList.Select(x => x).ToArray() });
                                    #endregion

                                    #region //單身刪除
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE SCM.SoDetail
                                            WHERE SoId IN (
                                                SELECT SoId
                                                FROM SCM.SaleOrder
                                                WHERE SoErpPrefix + '-' + SoErpNo IN @SoErpFullNo
                                            )";
                                    dynamicParameters.Add("SoErpFullNo", changeSoList.Select(x => x).ToArray());
                                    mainDelAffected += bmConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //單頭刪除
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE SCM.SaleOrder
                                            WHERE SoErpPrefix + '-' + SoErpNo IN @SoErpFullNo";
                                    dynamicParameters.Add("SoErpFullNo", changeSoList.Select(x => x).ToArray());
                                    detailDelAffected += bmConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                                #endregion
                            }

                            #region //ERP訂單單身資料
                            sql = @"SELECT TD001 SoErpPrefix, TD002 SoErpNo, TD003 SoSequence
                                    FROM COPTD
                                    WHERE 1=1
                                    ORDER BY TD001, TD002, TD003";
                            var resultErpSoDetail = erpConnection.Query<SoDetail>(sql, dynamicParameters);
                            #endregion

                            using (SqlConnection bmConnection = new SqlConnection(MainConnectionStrings))
                            {
                                #region //BM訂單單身資料
                                sql = @"SELECT b.SoErpPrefix, b.SoErpNo, a.SoSequence
                                        FROM SCM.SoDetail a
                                        INNER JOIN SCM.SaleOrder b ON a.SoId = b.SoId
                                        WHERE b.CompanyId = @CompanyId
                                        ORDER BY b.SoErpPrefix, b.SoErpNo, a.SoSequence";
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                var resultBmSoDetail = bmConnection.Query<SoDetail>(sql, dynamicParameters);
                                #endregion

                                #region //篩選ERP不存在但BM還存在的訂單單身
                                var dictionaryErpSoDetail = resultErpSoDetail.ToDictionary(x => x.SoErpPrefix + "-" + x.SoErpNo + "-" + x.SoSequence, x => x.SoErpPrefix + "-" + x.SoErpNo + "-" + x.SoSequence);
                                var dictionaryBmSoDetail = resultBmSoDetail.ToDictionary(x => x.SoErpPrefix + "-" + x.SoErpNo + "-" + x.SoSequence, x => x.SoErpPrefix + "-" + x.SoErpNo + "-" + x.SoSequence);
                                var changeSoDetail = dictionaryBmSoDetail.Where(x => !dictionaryErpSoDetail.ContainsKey(x.Key)).ToList();
                                var changeSoDetailList = changeSoDetail.Select(x => x.Value).ToList();
                                #endregion

                                #region //刪除訂單單身
                                if (changeSoDetailList.Count > 0)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE a
                                            FROM SCM.SoDetail a
                                            INNER JOIN SCM.SaleOrder b ON a.SoId = b.SoId
                                            WHERE b.SoErpPrefix + '-' + b.SoErpNo + '-' + RIGHT('0000' + a.SoSequence, 4) IN @SoErpFullNo";
                                    dynamicParameters.Add("SoErpFullNo", changeSoDetailList.Select(x => x).ToArray());
                                    detailDelAffected += bmConnection.Execute(sql, dynamicParameters, null, 300);
                                }
                                #endregion
                            }
                        }
                    }
                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "(" + mainAffected + detailAffected + " rows affected)",
                        data = string.Format(@"已更新資料<br/>【{0}】筆單頭<br/>【{1}】筆單身<br/>刪除<br/>【{2}】筆單頭<br/>【{3}】筆單身<br/>【{4}】筆暫存", 
                        mainAffected, detailAffected, mainDelAffected, detailDelAffected, tempAffected)
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

        #region //UpdateSoAmount -- 訂單單頭金額更新 -- Zoey 2022.07.19
        private void UpdateSoAmount(int SoId, SqlConnection sqlConnection, DynamicParameters dynamicParameters)
        {
            try
            {
                #region //判斷訂單單頭資料是否正確
                dynamicParameters = new DynamicParameters();
                sql = @"SELECT TOP 1 1
                        FROM SCM.SaleOrder
                        WHERE SoId = @SoId";
                dynamicParameters.Add("SoId", SoId);

                var result = sqlConnection.Query(sql, dynamicParameters);
                if (result.Count() <= 0) throw new SystemException("訂單單頭資料錯誤!");
                #endregion

                dynamicParameters = new DynamicParameters();
                sql = @"DECLARE @TotalQty FLOAT, @Amount FLOAT, @TaxAmount FLOAT

                        SELECT @TotalQty = ISNULL(SUM(a.SoQty), 0)
                        , @Amount = CASE b.Taxation
                            WHEN 1 THEN c.Amount
                            ELSE d.Amount
                            END 
                        , @TaxAmount = CASE b.Taxation
                            WHEN 1 THEN ROUND(c.Amount * b.BusinessTaxRate, 2)
                            ELSE ROUND(d.Amount * b.BusinessTaxRate, 2)
                            END
                        FROM SCM.SoDetail a
                        INNER JOIN SCM.SaleOrder b ON b.SoId = a.SoId
                        OUTER APPLY(
                            SELECT ROUND(ISNULL(SUM(x.Amount), 0) / (1 + b.BusinessTaxRate), 2) Amount
                            FROM SCM.SoDetail x
                            WHERE x.SoId = a.SoId
                            GROUP BY x.SoId
                        ) c
                        OUTER APPLY(
                            SELECT ISNULL(SUM(x.Amount), 0) Amount
                            FROM SCM.SoDetail x
                            WHERE x.SoId = a.SoId
                            GROUP BY x.SoId
                        ) d
                        WHERE a.SoId = @SoId
                        GROUP BY b.Taxation, c.Amount, d.Amount, b.BusinessTaxRate

                        UPDATE SCM.SaleOrder SET 
                        TotalQty = @TotalQty,
                        Amount = @Amount,
                        TaxAmount = @TaxAmount,
                        LastModifiedDate = @LastModifiedDate,
                        LastModifiedBy = @LastModifiedBy
                        WHERE SoId = @SoId";
                dynamicParameters.AddDynamicParams(
                    new
                    {
                        LastModifiedDate,
                        LastModifiedBy,
                        SoId
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
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }
        }

        public static implicit operator SaleOrderDA(DeliveryDA v)
        {
            throw new NotImplementedException();
        }
        #endregion

        #region //UpdateSaleOrderToERP -- 訂單資料拋轉 -- Zoey 2022.07.21
        public string UpdateSaleOrderToERP(int SoId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    List<COPTC> saleOrders = new List<COPTC>();
                    List<COPTD> soDetails = new List<COPTD>();

                    string dateNow = DateTime.Now.ToString("yyyyMMdd");
                    string timeNow = DateTime.Now.ToString("HH:mm:ss");

                    string SoErpPrefix = "", SoErpNo = "", SoSequence = "", SoDate = "", ErpNo = "";

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

                        if (companyResult.Count() <= 0) throw new SystemException("公司別錯誤!");

                        foreach (var item in companyResult)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            ErpNo = item.ErpNo;
                        }
                        #endregion

                        #region //撈取BM訂單單頭資料
                        sql = @"SELECT a.SoId, c.DepartmentNo TC005, a.SoErpPrefix TC001, a.SoErpNo TC002
                                , a.Version TC069, FORMAT(a.SoDate, 'yyyyMMdd') TC003, FORMAT(a.DocDate, 'yyyyMMdd') TC039
                                , a.SoRemark TC015, b.CustomerNo TC004, b.CustomerNo TC032, d.UserNo TC006
                                , a.CustomerAddressFirst TC010, a.CustomerAddressSecond TC011
                                , a.CustomerPurchaseOrder TC012, a.DepositPartial TC070, a.DepositRate TC045
                                , a.Currency TC008, a.ExchangeRate TC009, a.TaxNo TC078, a.Taxation TC016
                                , a.BusinessTaxRate TC041, a.DetailMultiTax TC091, a.TotalQty TC031
                                , a.Amount TC029, a.TaxAmount TC030, a.ShipMethod TC019
                                , a.TradeTerm TC068, a.PaymentTerm TC042, a.PriceTerm TC013
                                , a.TransferStatus, a.ConfirmStatus TC027
                                , b.CustomerName TC053, b.CustomerName TC065, b.CustomerEnglishName TC071
                                , b.Country TC081, b.Region TC080, b.Route TC083, b.Contact TC018
                                , b.TelNoFirst TC066, b.FaxNo TC067, b.InvoiceAddressFirst TC063
                                , b.InvoiceAddressSecond TC064, b.PaymentBankFirst TC036
                                FROM SCM.SaleOrder a
                                INNER JOIN SCM.Customer b ON b.CustomerId = a.CustomerId
                                LEFT JOIN BAS.Department c ON c.DepartmentId = a.DepartmentId
                                LEFT JOIN BAS.[User] d ON d.UserId = a.SalesmenId
                                LEFT JOIN BAS.[User] e ON e.UserId = a.ConfirmUserId
                                WHERE a.SoId = @SoId";
                        dynamicParameters.Add("SoId", SoId);

                        saleOrders = sqlConnection.Query<COPTC>(sql, dynamicParameters).ToList();

                        if (saleOrders.Count() <= 0) throw new SystemException("訂單資料錯誤!");
                        #endregion

                        #region //撈取BM訂單單身資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.SoDetailId, a.SoId, b.SoErpPrefix TD001, b.SoErpNo TD002, a.SoSequence TD003
                                , c.MtlItemNo TD004, a.SoMtlItemName TD005, a.SoMtlItemSpec TD006,d.InventoryNo TD007
                                , e.UomNo TD010, a.SoQty TD008, a.SiQty TD009
                                , a.ProductType TD049, a.FreebieQty TD024, a.FreebieSiQty TD025, a.SpareQty TD050, a.SpareSiQty TD051
                                , a.UnitPrice TD011, a.Amount TD012, FORMAT(a.PromiseDate, 'yyyyMMdd') TD013
                                , FORMAT(a.PcPromiseDate, 'yyyyMMdd') TD048, a.Project TD027, a.CustomerMtlItemNo TD014
                                , a.SoDetailRemark TD020, a.SoPriceQty TD076
                                , f.UomNo TD077, a.TaxNo TD079, a.BusinessTaxRate TD070, a.DiscountRate TD026
                                , a.DiscountAmount TD080, a.ConfirmStatus TD021, a.ClosureStatus TD016
                                FROM SCM.SoDetail a
                                INNER JOIN SCM.SaleOrder b ON b.SoId = a.SoId
                                INNER JOIN PDM.MtlItem c ON c.MtlItemId = a.MtlItemId
                                INNER JOIN SCM.Inventory d ON d.InventoryId = a.InventoryId
                                INNER JOIN PDM.UnitOfMeasure e ON e.UomId = a.UomId
                                INNER JOIN PDM.UnitOfMeasure f ON f.UomId = a.SoPriceUomId
                                WHERE a.SoId = @SoId
                                AND a.MtlItemId IS NOT NULL
                                ORDER BY a.SoId, a.SoSequence";
                        dynamicParameters.Add("SoId", SoId);

                        soDetails = sqlConnection.Query<COPTD>(sql, dynamicParameters).ToList();

                        if (soDetails.Count() <= 0) throw new SystemException("訂單單身資料錯誤!");
                        #endregion

                        #region //撈取單別/單號/流水號/訂單日期
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.SoErpPrefix, a.SoErpNo, b.SoSequence, a.SoDate 
                                FROM SCM.SaleOrder a
                                LEFT JOIN SCM.SoDetail b ON b.SoId = a.SoId
                                WHERE a.SoId = @SoId";
                        dynamicParameters.Add("SoId", SoId);
                        var result = sqlConnection.Query(sql, dynamicParameters);

                        if (result.Count() <= 0) throw new SystemException("訂單資料錯誤!");

                        foreach (var item in result)
                        {
                            SoErpPrefix = item.SoErpPrefix;
                            SoErpNo = item.SoErpNo;
                            SoSequence = item.SoSequence;
                            SoDate = Convert.ToDateTime(item.SoDate).ToString("yyyyMMdd");
                        }
                        #endregion

                        #region //撈取使用者No
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserNo
                                FROM BAS.[User] a
                                WHERE a.UserId = @userId";
                        dynamicParameters.Add("userId", CurrentUser);
                        var result2 = sqlConnection.Query(sql, dynamicParameters);

                        if (result2.Count() <= 0) throw new SystemException("使用者資料錯誤!");

                        foreach (var item in result2)
                        {
                            saleOrders
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.CREATOR = item.UserNo;
                                    x.CREATE_AP = item.UserNo + "PC";
                                    x.TC040 = item.UserNo;
                                });

                            soDetails
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.CREATOR = item.UserNo;
                                    x.CREATE_AP = item.UserNo + "PC";
                                });
                        }
                        #endregion
                    }

                    if (saleOrders.Count() > 0)
                    {
                        using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //撈取付款條件
                            sql = @"SELECT LTRIM(RTRIM(NA002)) PaymentTerm, LTRIM(RTRIM(NA003)) PaymentTermName
                                    FROM CMSNA
                                    WHERE 1=1
                                    AND NA001 = '2'";
                            List<SaleOrder> resultPaymentTerms = sqlConnection.Query<SaleOrder>(sql, dynamicParameters).ToList();

                            if (resultPaymentTerms.Count() <= 0) throw new SystemException("ERP付款條件不存在!");

                            saleOrders = saleOrders.GroupJoin(resultPaymentTerms, x => x.TC042, y => y.PaymentTerm, (x, y) => { x.TC014 = y.FirstOrDefault()?.PaymentTermName; return x; }).ToList();
                            #endregion

                            #region //廠別資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 MB001
                                    FROM CMSMB
                                    WHERE COMPANY = @COMPANY";
                            dynamicParameters.Add("COMPANY", ErpNo);

                            var resultFactory = sqlConnection.Query(sql, dynamicParameters);
                            if (resultFactory.Count() <= 0) throw new SystemException("【廠別資料】錯誤!");

                            string factory = "";
                            foreach (var item in resultFactory)
                            {
                                factory = item.MB001;
                            }
                            #endregion

                            if (saleOrders.Where(x => x.TransferStatus == "N" && x.TC027 == "N").Any())
                            {
                                #region //單號自動取號
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT CONVERT(int, RIGHT(ISNULL(MAX(RTRIM(LTRIM(TC002))), '000'), 3)) + 1 CurrentNum
                                        FROM COPTC
                                        WHERE TC001 = @SoErpPrefix
                                        AND TC002 LIKE '' + @SoDate + '___'";
                                dynamicParameters.Add("SoErpPrefix", SoErpPrefix);
                                dynamicParameters.Add("SoDate", SoDate);

                                int currentNum = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).CurrentNum;
                                SoErpNo = SoDate + string.Format("{0:000}", currentNum);

                                saleOrders
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.TC002 = SoErpNo;
                                        x.CFIELD01 = x.TC001 + "-" + SoErpNo;
                                    });

                                soDetails
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.TD002 = SoErpNo;
                                    });
                                #endregion

                                #region //判斷單號是否重複
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM COPTC
                                        WHERE TC001 = @TC001
                                        AND TC002 = @TC002";
                                dynamicParameters.Add("TC001", SoErpPrefix);
                                dynamicParameters.Add("TC002", SoErpNo);

                                var result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() > 0) throw new SystemException("【訂單單號】重複，請重新取號!");
                                #endregion

                                #region //單頭新增
                                if (saleOrders.Count > 0)
                                {
                                    saleOrders
                                        .ToList()
                                        .ForEach(x =>
                                        {
                                            x.COMPANY = ErpNo;
                                            x.USR_GROUP = "";
                                            x.CREATE_DATE = dateNow;
                                            x.MODIFIER = "";
                                            x.MODI_DATE = "";
                                            x.FLAG = "1";
                                            x.CREATE_TIME = timeNow;
                                            x.CREATE_PRID = "BM";
                                            x.MODI_TIME = "";
                                            x.MODI_AP = "";
                                            x.MODI_PRID = "";
                                            x.TC007 = factory; //出貨廠別
                                            x.TC017 = ""; //L/CNO.
                                            x.TC020 = ""; //起始港口
                                            x.TC021 = ""; //目的港口
                                            x.TC022 = ""; //代理商
                                            x.TC023 = ""; //報關行
                                            x.TC024 = ""; //驗貨公司
                                            x.TC025 = ""; //運輸公司
                                            x.TC026 = 0; //佣金比率
                                            x.TC027 = "Y"; //確認碼
                                            x.TC028 = 0; //列印次數
                                            x.TC033 = ""; //NOTIFY
                                            x.TC034 = ""; //嘜頭代號
                                            x.TC035 = ""; //目的地
                                            x.TC037 = ""; //INVOICE備註
                                            x.TC038 = ""; //PACKING-LIST備註
                                            x.TC043 = 0; //總毛重(Kg)
                                            x.TC044 = 0; //總材積(CUFT)
                                            x.TC046 = 0; //總包裝數量
                                            x.TC047 = ""; //押匯銀行
                                            x.TC048 = "N"; //簽核狀態碼
                                            x.TC049 = ""; //流程代號(多角貿易)
                                            x.TC050 = "N"; //拋轉狀態
                                            x.TC051 = ""; //下游廠商
                                            x.TC052 = 0; //傳送次數
                                            x.TC054 = ""; //正嘜
                                            x.TC055 = ""; //側嘜
                                            x.TC056 = "1"; //材積單位
                                            x.TC057 = "N"; //EBC確認碼
                                            x.TC058 = ""; //EBC訂單號碼
                                            x.TC059 = ""; //EBC訂單版次
                                            x.TC060 = "N"; //匯至EBC
                                            x.TC061 = ""; //正嘜文管代號
                                            x.TC062 = ""; //側嘜文管代號
                                            x.TC072 = 0; //預留欄位
                                            x.TC073 = 0; //收入遞延天數
                                            x.TC074 = ""; //預留欄位
                                            x.TC075 = ""; //預留欄位
                                            x.TC076 = ""; //預留欄位
                                            x.TC077 = "N"; //不控管信用額度
                                            x.TC079 = ""; //通路別
                                            x.TC082 = ""; //型態別
                                            x.TC084 = ""; //其他別
                                            x.TC085 = ""; //出口港
                                            x.TC086 = ""; //經過港口
                                            x.TC087 = ""; //目的港口
                                            x.TC088 = ""; //最上游客戶
                                            x.TC089 = ""; //最上游交易幣別
                                            x.TC090 = ""; //最上游稅別碼
                                            x.TC092 = ""; //來源
                                            x.TC093 = ""; //送貨國家
                                            x.UDF01 = "";
                                            x.UDF02 = "";
                                            x.UDF03 = "";
                                            x.UDF04 = "";
                                            x.UDF05 = "";
                                            x.UDF06 = 0;
                                            x.UDF07 = 0;
                                            x.UDF08 = 0;
                                            x.UDF09 = 0;
                                            x.UDF10 = 0;
                                        });

                                    sql = @"INSERT INTO COPTC (COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE
                                          , FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID
                                          , TC001, TC002, TC003, TC004, TC005, TC006, TC007, TC008, TC009, TC010
                                          , TC011, TC012, TC013, TC014, TC015, TC016, TC017, TC018, TC019, TC020
                                          , TC021, TC022, TC023, TC024, TC025, TC026, TC027, TC028, TC029, TC030
                                          , TC031, TC032, TC033, TC034, TC035, TC036, TC037, TC038, TC039, TC040
                                          , TC041, TC042, TC043, TC044, TC045, TC046, TC047, TC048, TC049, TC050
                                          , TC051, TC052, TC053, TC054, TC055, TC056, TC057, TC058, TC059, TC060
                                          , TC061, TC062, TC063, TC064, TC065, TC066, TC067, TC068, TC069, TC070
                                          , TC071, TC072, TC073, TC074, TC075, TC076, TC077, TC078, TC079, TC080
                                          , TC081, TC082, TC083, TC084, TC085, TC086, TC087, TC088, TC089, TC090
                                          , TC091, TC092, TC093
                                          , UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10)
                                          VALUES (@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE
                                          , @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID, @MODI_TIME, @MODI_AP, @MODI_PRID
                                          , @TC001, @TC002, @TC003, @TC004, @TC005, @TC006, @TC007, @TC008, @TC009, @TC010
                                          , @TC011, @TC012, @TC013, @TC014, @TC015, @TC016, @TC017, @TC018, @TC019, @TC020
                                          , @TC021, @TC022, @TC023, @TC024, @TC025, @TC026, @TC027, @TC028, @TC029, @TC030
                                          , @TC031, @TC032, @TC033, @TC034, @TC035, @TC036, @TC037, @TC038, @TC039, @TC040
                                          , @TC041, @TC042, @TC043, @TC044, @TC045, @TC046, @TC047, @TC048, @TC049, @TC050
                                          , @TC051, @TC052, @TC053, @TC054, @TC055, @TC056, @TC057, @TC058, @TC059, @TC060
                                          , @TC061, @TC062, @TC063, @TC064, @TC065, @TC066, @TC067, @TC068, @TC069, @TC070
                                          , @TC071, @TC072, @TC073, @TC074, @TC075, @TC076, @TC077, @TC078, @TC079, @TC080
                                          , @TC081, @TC082, @TC083, @TC084, @TC085, @TC086, @TC087, @TC088, @TC089, @TC090
                                          , @TC091, @TC092, @TC093
                                          , @UDF01, @UDF02, @UDF03, @UDF04, @UDF05, @UDF06, @UDF07, @UDF08, @UDF09, @UDF10)";
                                    rowsAffected += sqlConnection.Execute(sql, saleOrders);
                                }
                                #endregion

                                #region //單身新增
                                if (soDetails.Count > 0)
                                {
                                    soDetails
                                        .ToList()
                                        .ForEach(x =>
                                        {
                                            x.COMPANY = ErpNo;
                                            x.USR_GROUP = "";
                                            x.CREATE_DATE = dateNow;
                                            x.MODIFIER = "";
                                            x.MODI_DATE = "";
                                            x.FLAG = "1";
                                            x.CREATE_TIME = timeNow;
                                            x.CREATE_PRID = "BM";
                                            x.MODI_TIME = "";
                                            x.MODI_AP = "";
                                            x.MODI_PRID = "";
                                            x.TD015 = ""; //預測代號
                                            x.TD017 = ""; //前置單據-單別
                                            x.TD018 = ""; //前置單據-單號
                                            x.TD019 = ""; //前置單據-序號
                                            x.TD021 = "Y"; //確認碼
                                            x.TD022 = 0; //庫存數量
                                            x.TD023 = ""; //小單位
                                            x.TD028 = ""; //預測序號
                                            x.TD029 = ""; //包裝方式
                                            x.TD030 = 0; //毛重
                                            x.TD031 = 0; //材積
                                            x.TD032 = 0; //訂單包裝數量
                                            x.TD033 = 0; //已交包裝數量
                                            x.TD034 = 0; //贈品包裝量
                                            x.TD035 = 0; //贈品已交包裝量
                                            x.TD036 = ""; //包裝單位
                                            x.TD037 = ""; //原始客戶
                                            x.TD038 = ""; //請採購廠商
                                            x.TD040 = ""; //預留欄位
                                            x.TD041 = ""; //預留欄位
                                            x.TD042 = 0; //預留欄位
                                            x.TD043 = ""; //EBC訂單號碼
                                            x.TD044 = ""; //EBC訂單版次
                                            x.TD045 = "9"; //來源
                                            x.TD046 = ""; //圖號版次
                                            x.TD047 = ""; //原預交日
                                            x.TD052 = 0; //備品包裝量
                                            x.TD053 = 0; //備品已交包裝量
                                            x.TD054 = 0; //預留欄位
                                            x.TD055 = 0; //預留欄位
                                            x.TD056 = ""; //預留欄位
                                            x.TD057 = ""; //預留欄位
                                            x.TD058 = ""; //預留欄位
                                            x.TD059 = 0; //贈品率
                                            x.TD060 = ""; //預留欄位
                                            x.TD061 = 0; //RFQ
                                            x.TD062 = ""; //NewCode
                                            x.TD063 = ""; //測試備註一
                                            x.TD064 = ""; //測試備註二
                                            x.TD065 = ""; //最終客戶代號
                                            x.TD066 = ""; //計畫批號
                                            x.TD067 = ""; //優先順序
                                            x.TD068 = ""; //預留欄位
                                            x.TD069 = ""; //鎖定交期
                                            x.TD071 = ""; //CRM單別
                                            x.TD072 = ""; //CRM單號
                                            x.TD073 = ""; //CRM序號
                                            x.TD074 = ""; //CRM合約代號
                                            x.TD075 = ""; //業務品號
                                            x.TD078 = 0; //已交計價數量
                                            x.TD500 = ""; //排程日期
                                            x.TD501 = 0; //可排量
                                            x.TD502 = ""; //產品系列
                                            x.TD503 = ""; //客戶需求日
                                            x.TD504 = ""; //以包裝單位計價
                                            x.UDF01 = "";
                                            x.UDF02 = "";
                                            x.UDF03 = "";
                                            x.UDF04 = "";
                                            x.UDF05 = "";
                                            x.UDF06 = 0;
                                            x.UDF07 = 0;
                                            x.UDF08 = 0;
                                            x.UDF09 = 0;
                                            x.UDF10 = 0;
                                        });

                                    sql = @"INSERT INTO COPTD(COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE
                                            , FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID
                                            , TD001, TD002, TD003, TD004, TD005, TD006, TD007, TD008, TD009, TD010
                                            , TD011, TD012, TD013, TD014, TD015, TD016, TD017, TD018, TD019, TD020
                                            , TD021, TD022, TD023, TD024, TD025, TD026, TD027, TD028, TD029, TD030
                                            , TD031, TD032, TD033, TD034, TD035, TD036, TD037, TD038, TD039, TD040
                                            , TD041, TD042, TD043, TD044, TD045, TD046, TD047, TD048, TD049, TD050
                                            , TD051, TD052, TD053, TD054, TD055, TD056, TD057, TD058, TD059, TD060
                                            , TD061, TD062, TD063, TD064, TD065, TD066, TD067, TD068, TD069, TD070
                                            , TD071, TD072, TD073, TD074, TD075, TD076, TD077, TD078, TD079, TD080
                                            , TD500, TD501, TD502, TD503, TD504
                                            , UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10)
                                            VALUES(@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE
                                            , @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID, @MODI_TIME, @MODI_AP, @MODI_PRID
                                            , @TD001, @TD002, @TD003, @TD004, @TD005, @TD006, @TD007, @TD008, @TD009, @TD010
                                            , @TD011, @TD012, @TD013, @TD014, @TD015, @TD016, @TD017, @TD018, @TD019, @TD020
                                            , @TD021, @TD022, @TD023, @TD024, @TD025, @TD026, @TD027, @TD028, @TD029, @TD030
                                            , @TD031, @TD032, @TD033, @TD034, @TD035, @TD036, @TD037, @TD038, @TD039, @TD040
                                            , @TD041, @TD042, @TD043, @TD044, @TD045, @TD046, @TD047, @TD048, @TD049, @TD050
                                            , @TD051, @TD052, @TD053, @TD054, @TD055, @TD056, @TD057, @TD058, @TD059, @TD060
                                            , @TD061, @TD062, @TD063, @TD064, @TD065, @TD066, @TD067, @TD068, @TD069, @TD070
                                            , @TD071, @TD072, @TD073, @TD074, @TD075, @TD076, @TD077, @TD078, @TD079, @TD080
                                            , @TD500, @TD501, @TD502, @TD503, @TD504
                                            , @UDF01, @UDF02, @UDF03, @UDF04, @UDF05, @UDF06, @UDF07, @UDF08, @UDF09, @UDF10)";
                                    rowsAffected += sqlConnection.Execute(sql, soDetails);
                                }

                                #endregion
                            }

                            if (saleOrders.Where(x => x.TransferStatus == "Y" && x.TC027 == "N").Any())
                            {
                                #region //單頭修改
                                if (saleOrders.Count > 0)
                                {
                                    saleOrders
                                        .ToList()
                                        .ForEach(x =>
                                        {
                                            x.MODI_DATE = dateNow;
                                            x.MODI_TIME = timeNow;
                                            x.TC027 = "Y"; //確認碼
                                        });

                                    sql = @"UPDATE COPTC SET
                                            MODI_DATE = @MODI_DATE,
                                            MODI_TIME = @MODI_TIME,
                                            TC003 = @TC003, 
                                            TC004 = @TC004, 
                                            TC005 = @TC005, 
                                            TC006 = @TC006, 
                                            TC008 = @TC008, 
                                            TC009 = @TC009, 
                                            TC010 = @TC010, 
                                            TC011 = @TC011, 
                                            TC012 = @TC012, 
                                            TC013 = @TC013, 
                                            TC014 = @TC014, 
                                            TC015 = @TC015, 
                                            TC016 = @TC016, 
                                            TC018 = @TC018, 
                                            TC019 = @TC019,
                                            TC027 = @TC027,
                                            TC029 = @TC029,
                                            TC030 = @TC030,
                                            TC031 = @TC031,
                                            TC032 = @TC032,
                                            TC036 = @TC036,
                                            TC039 = @TC039,
                                            TC040 = @TC040,
                                            TC041 = @TC041,
                                            TC042 = @TC042,
                                            TC045 = @TC045,
                                            TC053 = @TC053,
                                            TC063 = @TC063,
                                            TC064 = @TC064,
                                            TC065 = @TC065,
                                            TC066 = @TC066,
                                            TC067 = @TC067,
                                            TC068 = @TC068,
                                            TC069 = @TC069,
                                            TC070 = @TC070,
                                            TC071 = @TC071,
                                            TC078 = @TC078,
                                            TC080 = @TC080,
                                            TC081 = @TC081,
                                            TC083 = @TC083,
                                            TC091 = @TC091
                                            WHERE TC001 = @TC001
                                            AND TC002 = @TC002";
                                    rowsAffected += sqlConnection.Execute(sql, saleOrders);
                                }
                                #endregion

                                #region //單身修改
                                if (soDetails.Count > 0)
                                {
                                    #region //刪除所有單身
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE COPTD
                                            WHERE TD001 = @TD001
                                            AND TD002 = @TD002 ";
                                    dynamicParameters.Add("@TD001", SoErpPrefix);
                                    dynamicParameters.Add("@TD002", SoErpNo);

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //單身新增
                                    if (soDetails.Count > 0)
                                    {
                                        soDetails
                                            .ToList()
                                            .ForEach(x =>
                                            {
                                                x.COMPANY = ErpNo;
                                                x.USR_GROUP = "";
                                                x.CREATE_DATE = dateNow;
                                                x.MODIFIER = "";
                                                x.MODI_DATE = "";
                                                x.FLAG = "1";
                                                x.CREATE_TIME = timeNow;
                                                x.CREATE_PRID = "BM";
                                                x.MODI_TIME = "";
                                                x.MODI_AP = "";
                                                x.MODI_PRID = "";
                                                x.TD015 = ""; //預測代號
                                                x.TD017 = ""; //前置單據-單別
                                                x.TD018 = ""; //前置單據-單號
                                                x.TD019 = ""; //前置單據-序號
                                                x.TD021 = "Y"; //確認碼
                                                x.TD022 = 0; //庫存數量
                                                x.TD023 = ""; //小單位
                                                x.TD028 = ""; //預測序號
                                                x.TD029 = ""; //包裝方式
                                                x.TD030 = 0; //毛重
                                                x.TD031 = 0; //材積
                                                x.TD032 = 0; //訂單包裝數量
                                                x.TD033 = 0; //已交包裝數量
                                                x.TD034 = 0; //贈品包裝量
                                                x.TD035 = 0; //贈品已交包裝量
                                                x.TD036 = ""; //包裝單位
                                                x.TD037 = ""; //原始客戶
                                                x.TD038 = ""; //請採購廠商
                                                x.TD040 = ""; //預留欄位
                                                x.TD041 = ""; //預留欄位
                                                x.TD042 = 0; //預留欄位
                                                x.TD043 = ""; //EBC訂單號碼
                                                x.TD044 = ""; //EBC訂單版次
                                                x.TD045 = "9"; //來源
                                                x.TD046 = ""; //圖號版次
                                                x.TD047 = ""; //原預交日
                                                x.TD052 = 0; //備品包裝量
                                                x.TD053 = 0; //備品已交包裝量
                                                x.TD054 = 0; //預留欄位
                                                x.TD055 = 0; //預留欄位
                                                x.TD056 = ""; //預留欄位
                                                x.TD057 = ""; //預留欄位
                                                x.TD058 = ""; //預留欄位
                                                x.TD059 = 0; //贈品率
                                                x.TD060 = ""; //預留欄位
                                                x.TD061 = 0; //RFQ
                                                x.TD062 = ""; //NewCode
                                                x.TD063 = ""; //測試備註一
                                                x.TD064 = ""; //測試備註二
                                                x.TD065 = ""; //最終客戶代號
                                                x.TD066 = ""; //計畫批號
                                                x.TD067 = ""; //優先順序
                                                x.TD068 = ""; //預留欄位
                                                x.TD069 = ""; //鎖定交期
                                                x.TD071 = ""; //CRM單別
                                                x.TD072 = ""; //CRM單號
                                                x.TD073 = ""; //CRM序號
                                                x.TD074 = ""; //CRM合約代號
                                                x.TD075 = ""; //業務品號
                                                x.TD078 = 0; //已交計價數量
                                                x.TD500 = ""; //排程日期
                                                x.TD501 = 0; //可排量
                                                x.TD502 = ""; //產品系列
                                                x.TD503 = ""; //客戶需求日
                                                x.TD504 = ""; //以包裝單位計價
                                                x.UDF01 = "";
                                                x.UDF02 = "";
                                                x.UDF03 = "";
                                                x.UDF04 = "";
                                                x.UDF05 = "";
                                                x.UDF06 = 0;
                                                x.UDF07 = 0;
                                                x.UDF08 = 0;
                                                x.UDF09 = 0;
                                                x.UDF10 = 0;
                                            });

                                        sql = @"INSERT INTO COPTD(COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE
                                                , FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID
                                                , TD001, TD002, TD003, TD004, TD005, TD006, TD007, TD008, TD009, TD010
                                                , TD011, TD012, TD013, TD014, TD015, TD016, TD017, TD018, TD019, TD020
                                                , TD021, TD022, TD023, TD024, TD025, TD026, TD027, TD028, TD029, TD030
                                                , TD031, TD032, TD033, TD034, TD035, TD036, TD037, TD038, TD039, TD040
                                                , TD041, TD042, TD043, TD044, TD045, TD046, TD047, TD048, TD049, TD050
                                                , TD051, TD052, TD053, TD054, TD055, TD056, TD057, TD058, TD059, TD060
                                                , TD061, TD062, TD063, TD064, TD065, TD066, TD067, TD068, TD069, TD070
                                                , TD071, TD072, TD073, TD074, TD075, TD076, TD077, TD078, TD079, TD080
                                                , TD500, TD501, TD502, TD503, TD504
                                                , UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10)
                                                VALUES(@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE
                                                , @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID, @MODI_TIME, @MODI_AP, @MODI_PRID
                                                , @TD001, @TD002, @TD003, @TD004, @TD005, @TD006, @TD007, @TD008, @TD009, @TD010
                                                , @TD011, @TD012, @TD013, @TD014, @TD015, @TD016, @TD017, @TD018, @TD019, @TD020
                                                , @TD021, @TD022, @TD023, @TD024, @TD025, @TD026, @TD027, @TD028, @TD029, @TD030
                                                , @TD031, @TD032, @TD033, @TD034, @TD035, @TD036, @TD037, @TD038, @TD039, @TD040
                                                , @TD041, @TD042, @TD043, @TD044, @TD045, @TD046, @TD047, @TD048, @TD049, @TD050
                                                , @TD051, @TD052, @TD053, @TD054, @TD055, @TD056, @TD057, @TD058, @TD059, @TD060
                                                , @TD061, @TD062, @TD063, @TD064, @TD065, @TD066, @TD067, @TD068, @TD069, @TD070
                                                , @TD071, @TD072, @TD073, @TD074, @TD075, @TD076, @TD077, @TD078, @TD079, @TD080
                                                , @TD500, @TD501, @TD502, @TD503, @TD504
                                                , @UDF01, @UDF02, @UDF03, @UDF04, @UDF05, @UDF06, @UDF07, @UDF08, @UDF09, @UDF10)";
                                        rowsAffected += sqlConnection.Execute(sql, soDetails);
                                    }
                                    #endregion
                                }
                                #endregion
                            }
                        }

                        using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                        {
                            #region//修改BM單頭單號/拋轉狀態/確認碼
                            sql = @"UPDATE SCM.SaleOrder SET
                                    SoErpNo = @SoErpNo,
                                    ConfirmStatus = @ConfirmStatus,
                                    ConfirmUserId = @CurrentUser,
                                    TransferStatus = @TransferStatus,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE SoId = @SoId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    SoErpNo = saleOrders.Select(x => x.TC002).FirstOrDefault(),
                                    ConfirmStatus = "Y",
                                    CurrentUser,
                                    TransferStatus = "Y",
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    SoId
                                });
                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region//修改BM單身確認碼
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.SoDetail SET
                                    ConfirmStatus = @ConfirmStatus,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE SoId = @SoId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    ConfirmStatus = "Y",
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    SoId
                                });
                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }
                    }
                    transactionScope.Complete();
                }

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    msg = "拋轉成功"
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

        #region //UpdateSoModification -- 訂單變更單頭資料更新 -- Zoey 2022.08.12
        public string UpdateSoModification(int SomId, int SoId, string Version, int DepartmentId, string DocDate, string SoRemark
            , int SalesmenId, string CustomerAddressFirst, string CustomerAddressSecond, string CustomerPurchaseOrder
            , string DepositPartial, double DepositRate, string Currency, double ExchangeRate, string TaxNo, string Taxation
            , double BusinessTaxRate, string DetailMultiTax, string ShipMethod, string TradeTerm, string PaymentTerm, string PriceTerm
            , string ClosureStatus, string ModiReason, string ConfirmStatus, int ConfirmUserId, string TransferStatus, string TransferDate)
        {
            try
            {
                if (DocDate.Length <= 0) throw new SystemException("【單據日期】不能為空!");
                if (CustomerPurchaseOrder.Length > 20) throw new SystemException("【客戶採購單】長度錯誤!");
                if (DepartmentId <= 0) throw new SystemException("【部門名稱】不能為空!");
                if (SalesmenId <= 0) throw new SystemException("【業務人員】不能為空!");
                if (SoRemark.Length > 255) throw new SystemException("【備註】長度錯誤!");
                if (CustomerAddressFirst.Length <= 0) throw new SystemException("【客戶聯絡地址(一)】不能為空!");
                if (CustomerAddressFirst.Length > 200) throw new SystemException("【客戶聯絡地址(一)】長度錯誤!");
                if (CustomerAddressSecond.Length > 200) throw new SystemException("【客戶聯絡地址(二)】長度錯誤!");
                if (Currency.Length <= 0) throw new SystemException("【交易幣別】不能為空!");
                if (ExchangeRate <= 0) throw new SystemException("【匯率】不能為空!");
                if (TradeTerm.Length <= 0) throw new SystemException("【交易條件】不能為空!");
                if (TaxNo.Length <= 0) throw new SystemException("【稅別碼】不能為空!");
                if (Taxation.Length <= 0) throw new SystemException("【課稅別】不能為空!");
                if (BusinessTaxRate < 0) throw new SystemException("【營業稅率】不能為空!");
                if (PaymentTerm.Length <= 0) throw new SystemException("【付款條件】不能為空!");
                if (PriceTerm.Length > 40) throw new SystemException("【價格條件】長度錯誤!");
                if (DepositPartial.Length <= 0) throw new SystemException("【訂金分批】不能為空!");
                if (DetailMultiTax.Length <= 0) throw new SystemException("【單身多稅率】不能為空!");
                if (DepositRate < 0) throw new SystemException("【訂金比率】不能為空!");
                if (ShipMethod.Length <= 0) throw new SystemException("【運輸方式】不能為空!");
                if (ModiReason.Length > 255) throw new SystemException("【變更原因】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷訂單變更單頭資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.SoModification
                                WHERE SomId = @SomId";
                        dynamicParameters.Add("SomId", SomId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("訂單變更單頭資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.SoModification SET
                                DepartmentId = @DepartmentId,
                                DocDate = @DocDate,
                                SoRemark = @SoRemark,
                                SalesmenId = @SalesmenId,
                                CustomerAddressFirst = @CustomerAddressFirst,
                                CustomerAddressSecond = @CustomerAddressSecond,
                                CustomerPurchaseOrder = @CustomerPurchaseOrder,
                                DepositPartial = @DepositPartial,
                                DepositRate = @DepositRate,
                                Currency = @Currency,
                                ExchangeRate = @ExchangeRate,
                                TaxNo = @TaxNo,
                                Taxation = @Taxation,
                                BusinessTaxRate = @BusinessTaxRate,
                                DetailMultiTax = @DetailMultiTax,
                                ShipMethod = @ShipMethod,
                                TradeTerm = @TradeTerm,
                                PaymentTerm = @PaymentTerm,
                                PriceTerm = @PriceTerm,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SoId = @SoId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                DepartmentId,
                                DocDate,
                                SoRemark,
                                SalesmenId,
                                CustomerAddressFirst,
                                CustomerAddressSecond,
                                CustomerPurchaseOrder,
                                DepositPartial,
                                DepositRate,
                                Currency,
                                ExchangeRate,
                                TaxNo,
                                Taxation,
                                BusinessTaxRate,
                                DetailMultiTax,
                                ShipMethod,
                                TradeTerm,
                                PaymentTerm,
                                PriceTerm,
                                TransferStatus,
                                LastModifiedDate,
                                LastModifiedBy,
                                SoId
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

        #region //UpdateSaleOrderStatus -- 訂單資料狀態更新 -- Zoey 2022.08.15
        public string UpdateSaleOrderStatus(int SoId, string StatusName)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    string dateNow = DateTime.Now.ToString("yyyyMMdd");
                    string timeNow = DateTime.Now.ToString("HH:mm:ss");
                    string userNo = "";

                    string Responsemsg = "";

                    string SoErpPrefix = "", SoErpNo = "";

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷公司資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 CompanyId, CompanyName, ErpDb
                                FROM BAS.Company
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【公司】資料錯誤!");

                        foreach (var item in result)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        }
                        #endregion

                        #region //撈取單別+單號
                        sql = @"SELECT a.SoErpPrefix, a.SoErpNo 
                                FROM SCM.SaleOrder a
                                WHERE a.SoId = @SoId";
                        dynamicParameters.Add("SoId", SoId);
                        var result1 = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in result1)
                        {
                            SoErpPrefix = item.SoErpPrefix;
                            SoErpNo = item.SoErpNo;
                        }
                        #endregion

                        #region //撈取使用者No
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserNo
                               FROM BAS.[User] a
                               WHERE a.UserId = @userId";
                        dynamicParameters.Add("userId", CurrentUser);

                        userNo = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).UserNo;
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        switch (StatusName)
                        {
                            case "N": //訂單反確認
                                #region //判斷SCM.SoModification是否有資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM SCM.SoModification
                                        WHERE SoId = @SoId";
                                dynamicParameters.Add("SoId", SoId);
                                var result2 = sqlConnection.Query(sql, dynamicParameters);
                                if (result2.Count() > 0) throw new SystemException("該訂單已變更，無法進行反確認!");
                                #endregion

                                #region //判斷SCM.DoDetail是否有資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.SoDetailId FROM SCM.DoDetail a
                                        INNER JOIN SCM.SoDetail b ON b.SoDetailId = a.SoDetailId
                                        INNER JOIN SCM.SaleOrder c ON c.SoId = b.SoId
                                        WHERE c.SoId = @SoId";
                                dynamicParameters.Add("SoId", SoId);
                                var result3 = sqlConnection.Query(sql, dynamicParameters);
                                if (result3.Count() > 0) throw new SystemException("該訂單已開立出貨單，無法進行反確認!");
                                #endregion

                                sql = @"UPDATE SCM.SaleOrder SET
                                        ConfirmStatus = 'N',
                                        ConfirmUserId = @CurrentUser,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE SoId = @SoId ";

                                sql += @"UPDATE SCM.SoDetail SET
                                        ConfirmStatus = 'N',
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE SoId = @SoId";
                                Responsemsg = "反確認成功";
                                break;
                            case "V": //訂單作廢
                                #region //判斷確認碼是否為Y
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM SCM.SaleOrder
                                        WHERE SoId = @SoId
                                        AND ConfirmStatus = 'Y'";
                                dynamicParameters.Add("SoId", SoId);

                                var result4 = sqlConnection.Query(sql, dynamicParameters);
                                if (result4.Count() > 0) throw new SystemException("訂單已確認，無法進行作廢!");
                                #endregion

                                sql = @"UPDATE SCM.SaleOrder SET
                                        ConfirmStatus = 'V',
                                        ConfirmUserId = @CurrentUser,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE SoId = @SoId ";

                                sql += @"UPDATE SCM.SoDetail SET
                                        ConfirmStatus = 'V',
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE SoId = @SoId";
                                Responsemsg = "作廢成功";
                                break;
                        }

                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CurrentUser,
                                LastModifiedDate,
                                LastModifiedBy,
                                SoId
                            });

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        switch (StatusName)
                        {
                            case "N": //訂單反確認
                                #region //判斷是否有訂單變更單(COPTE)
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM COPTE
                                        WHERE TE001 = @SoErpPrefix
                                        AND TE002 = @SoErpNo";
                                dynamicParameters.Add("SoErpPrefix", SoErpPrefix);
                                dynamicParameters.Add("SoErpNo", SoErpNo);

                                var result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() > 0) throw new SystemException("訂單已變更，無法進行反確認!");
                                #endregion

                                #region //判斷是否有暫出單(INVTF)
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM INVTG
                                        WHERE TG014 = @SoErpPrefix
                                        AND TG015 = @SoErpNo";
                                dynamicParameters.Add("SoErpPrefix", SoErpPrefix);
                                dynamicParameters.Add("SoErpNo", SoErpNo);

                                var result2 = sqlConnection.Query(sql, dynamicParameters);
                                if (result2.Count() > 0) throw new SystemException("已開立暫出單，無法進行反確認!");
                                #endregion

                                sql = @"UPDATE COPTC SET
                                        TC027 = 'N',
                                        TC040 = @userNo,
                                        MODIFIER = @userNo,
                                        MODI_DATE = @dateNow,
                                        MODI_TIME = @timeNow
                                        WHERE TC001 = @SoErpPrefix
                                        AND TC002 = @SoErpNo ";

                                sql += @"UPDATE COPTD SET
                                         TD021 = 'N',
                                         MODIFIER = @userNo,
                                         MODI_DATE = @dateNow,
                                         MODI_TIME = @timeNow
                                         WHERE TD001 = @SoErpPrefix
                                         AND TD002 = @SoErpNo";
                                break;
                            case "V": //訂單作廢
                                sql = @"UPDATE COPTC SET
                                        TC027 = 'V',
                                        TC040 = @userNo,
                                        MODIFIER = @userNo,
                                        MODI_DATE = @dateNow,
                                        MODI_TIME = @timeNow
                                        WHERE TC001 = @SoErpPrefix
                                        AND TC002 = @SoErpNo ";

                                sql += @"UPDATE COPTD SET
                                         TD021 = 'V',
                                         MODIFIER = @userNo,
                                         MODI_DATE = @dateNow,
                                         MODI_TIME = @timeNow
                                         WHERE TD001 = @SoErpPrefix
                                         AND TD002 = @SoErpNo";
                                break;
                        }

                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                userNo,
                                dateNow,
                                timeNow,
                                SoErpPrefix,
                                SoErpNo
                            });

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                    }

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = Responsemsg
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

        #region //UpdateSoTransferBpm -- 拋轉訂單單據至BPM(若ERP未拋，連同ERP一起拋) -- Ann 2023-12-05
        public string UpdateSoTransferBpm(int SoId)
        {
            try
            {
                string token = "";
                int rowsAffected = 0;
                string CompanyNo = "";
                int totalDetail = 0;
                int CounterSign = 0; //單身是否有加簽筆數 0:沒有 1:有
                List<SaleOrder> saleOrders = new List<SaleOrder>();
                List<SoDetail> soDetails = new List<SoDetail>();

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //使用者資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 d.UserNo, d.UserName, d.DepartmentNo, d.DepartmentName
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
                                OUTER APPLY (
                                    SELECT da.UserNo, da.UserName, db.DepartmentNo, db.DepartmentName
                                    FROM BAS.[User] da
                                    INNER JOIN BAS.Department db ON da.DepartmentId = db.DepartmentId
                                    WHERE da.UserId = @UserId
                                ) d
                                WHERE a.[Status] = @Status
                                AND b.[Status] = @Status
                                AND b.FunctionCode = @FunctionCode
                                AND a.DetailCode = @DetailCode
                                AND c.Authority > 0";
                        dynamicParameters.Add("UserId", CurrentUser);
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("Status", "A");
                        dynamicParameters.Add("FunctionCode", "SaleOrderManagement");
                        dynamicParameters.Add("DetailCode", "bpm-transfer");

                        var resultUser = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUser.Count() <= 0) throw new SystemException("使用者資料錯誤!");

                        string UserNo = "";
                        string DepartmentNo = "";
                        string DepartmentName = "";
                        foreach (var item in resultUser)
                        {
                            UserNo = item.UserNo;
                            DepartmentNo = item.DepartmentNo;
                            DepartmentName = item.DepartmentName;
                        }
                        #endregion

                        #region //確認公司別DB
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CompanyNo, a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var companyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (companyResult.Count() <= 0) throw new SystemException("公司別錯誤!");

                        foreach (var item in companyResult)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            CompanyNo = item.CompanyNo;
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //查詢ERP是否有此帳號
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(a.MF004)) USR_GROUP
                                    FROM ADMMF a
                                    WHERE MF001 = @MF001";
                            dynamicParameters.Add("MF001", UserNo);

                            var ADMMFResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (ADMMFResult.Count() <= 0) throw new SystemException("查無ERP帳號權限，無法進行拋轉及核單!");
                            #endregion

                            #region //判斷訂單資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.TransferStatus, a.ConfirmStatus
                                    , b.SoBpmInfoId, ISNULL(b.BpmTransferStatus, 'N') BpmTransferStatus
                                    FROM SCM.SaleOrder a 
                                    LEFT JOIN SCM.SoBpmInfo b ON a.SoId = b.SoId
                                    WHERE a.SoId = @SoId
                                    AND a.CompanyId = @CompanyId";
                            dynamicParameters.Add("SoId", SoId);
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("訂單資料錯誤!");

                            int? SoBpmInfoId = -1;
                            foreach (var item in result)
                            {
                                if (item.TransferStatus != "Y")
                                {
                                    string response = UpdateSaleOrderImport(SoId);
                                    JObject responseJson = JObject.Parse(response);
                                    if (responseJson["status"].ToString() != "success")
                                    {
                                        throw new SystemException(responseJson["msg"].ToString());
                                    }
                                }

                                if (item.ConfirmStatus != "N") throw new SystemException("訂單不是未確認狀態，無法拋轉BPM!!");
                                if (item.BpmTransferStatus != "N" && item.BpmTransferStatus != "E") throw new SystemException("訂單狀態無法重複拋轉BPM!!");

                                SoBpmInfoId = item.SoBpmInfoId;
                            }
                            #endregion

                            #region //取得BPM TOKEN
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.IpAddress, a.CompanyId, a.Token, FORMAT(a.VerifyDate, 'yyyy-MM-dd HH:mm:ss') VerifyDate
                                    FROM BPM.SystemToken a
                                    WHERE IpAddress = @IpAddress
                                    AND CompanyId = @CompanyId";
                            dynamicParameters.Add("IpAddress", BpmServerPath);
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            var systemTokenResult = sqlConnection.Query(sql, dynamicParameters);
                            if (systemTokenResult.Count() <= 0) throw new SystemException("查無此憑證!");

                            foreach (var item in systemTokenResult)
                            {
                                DateTime verifyDate = Convert.ToDateTime(item.VerifyDate);
                                DateTime nowDate = DateTime.Now;
                                var CheckMin = (nowDate - verifyDate).TotalMinutes;
                                if (CheckMin >= 30)
                                {
                                    #region //取得新BPM TOKEN
                                    string tokenResponse = BpmHelper.GetBpmToken(BpmServerPath, BpmAccount, BpmPassword);
                                    var tokenJson = JObject.Parse(tokenResponse);
                                    foreach (var item2 in tokenJson)
                                    {
                                        if (item2.Key == "status")
                                        {
                                            if (item2.Value.ToString() != "success") throw new SystemException("取得token失敗!");
                                        }
                                        else if (item2.Key == "data")
                                        {
                                            token = item2.Value.ToString();
                                        }
                                    }
                                    #endregion

                                    #region //將新的TOKEN更新回MES
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE BPM.SystemToken SET
                                            Token = @Token,
                                            VerifyDate = @VerifyDate,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE IpAddress = @IpAddress
                                            AND CompanyId = @CompanyId";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            Token = token,
                                            VerifyDate = nowDate,
                                            LastModifiedDate,
                                            LastModifiedBy,
                                            IpAddress = BpmServerPath,
                                            CompanyId = CurrentCompany
                                        });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                                else
                                {
                                    token = item.Token;
                                }
                            }
                            #endregion

                            #region //取得BpmUser資料
                            string BpmUserId = "";
                            string BpmRoleId = "";
                            string BpmUserNo = "";
                            string UserName = "";
                            string BpmDepNo = "";
                            string BpmDepName = "";
                            using (SqlConnection sqlConnection3 = new SqlConnection(BpmDbConnectionStrings))
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"WITH BasicUserInfo(MemID, LoginID, UserName, MainRoleID, RolID, RolName, ParentRol) AS(
                                        SELECT a.MemID, a.LoginID, a.UserName, a.MainRoleID
                                        , b.RolID, b.Name, b.DepID AS ParentRol
                                        FROM Mem_GenInf a
                                        LEFT JOIN Rol_GenInf b ON b.RolID = a.MainRoleID
                                        WHERE a.LoginID = @LoginID
                                        UNION ALL
                                        SELECT a.MemID, a.LoginID, a.UserName, a.MainRoleID
                                        , b.RolID, b.Name, b.DepID AS ParentRol
                                        FROM BasicUserInfo a, Rol_GenInf b
                                        WHERE a.ParentRol = b.RolID
                                        )
                                        SELECT a.MemID, a.LoginID, a.UserName, a.MainRoleID
                                        , b.Name AS RolName
                                        , c.DepID, c.ID AS DepNo, c.Name AS DepName
                                        FROM BasicUserInfo a
                                        LEFT JOIN Rol_GenInf AS parentRol_GenInf ON a.RolID = parentRol_GenInf.RolID
                                        LEFT JOIN Dep_GenInf c ON parentRol_GenInf.DepID = c.DepID
                                        , Rol_GenInf b
                                        WHERE a.MainRoleID = b.RolID
                                        AND c.DepID IS NOT NULL
                                        UNION
                                        SELECT a.MemID, a.LoginID, a.UserName, a.MainRoleID
                                        , b.Name AS RolName
                                        , c.ComID AS DepID, c.ID AS DepNo, c.Name AS DepName
                                        FROM Mem_GenInf a
                                        LEFT JOIN Rol_GenInf b ON a.MainRoleID = b.RolID
                                        LEFT JOIN Company c ON b.DepID = c.ComID
                                        WHERE c.ComID IS NOT NULL
                                        AND a.LoginID = @LoginID
                                        ORDER BY a.LoginID";
                                dynamicParameters.Add("LoginID", UserNo);

                                var MemGenInfResult = sqlConnection3.Query(sql, dynamicParameters);

                                if (MemGenInfResult.Count() <= 0) throw new SystemException("取得BPM使用者資訊時發生錯誤!!");

                                foreach (var item in MemGenInfResult)
                                {
                                    BpmUserId = item.MemID;
                                    BpmRoleId = item.MainRoleID;
                                    BpmUserNo = item.LoginID;
                                    UserName = item.UserName;
                                    BpmDepNo = item.DepNo;
                                    BpmDepName = item.DepName;
                                }
                            }
                            #endregion

                            #region //單身筆數
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT COUNT(1) TotalDetail
                                FROM SCM.SoDetail
                                WHERE SoId = @SoId";
                            dynamicParameters.Add("SoId", SoId);

                            var resultDetailExist = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item in resultDetailExist)
                            {
                                totalDetail = Convert.ToInt32(item.TotalDetail);
                                if (totalDetail <= 0) throw new SystemException("請先建立單身!");
                            }
                            #endregion

                            #region //訂單單頭
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.SoId, a.SoErpPrefix + '-' + a.SoErpNo SoErpFullNo, FORMAT(a.DocDate, 'yyyy-MM-dd') DocDate
                                    , a.SoErpPrefix, a.SoErpNo, a.TradeTerm, a.PaymentTerm
                                    , a.CustomerPurchaseOrder, a.SoRemark, a.Amount, a.TaxAmount, a.Currency, a.ExchangeRate, a.BusinessTaxRate
                                    , b.UserNo SalesmenNo, b.UserName SalesmenName
                                    , b.UserNo, b.UserName
                                    , d.DepartmentNo SalesmenDepNo, d.DepartmentName SalesmenDepName
                                    , e.DepartmentNo, e.DepartmentName
                                    , f.CustomerShortName CustomerFullNo, f.TelNoFirst, f.FaxNo, f.DeliveryAddressFirst, f.GuiNumber
                                    , f.CustomerNo, f.CustomerShortName
                                    , g.TypeName Taxation
                                    FROM SCM.SaleOrder a 
                                    INNER JOIN BAS.[User] b ON a.SalesmenId = b.UserId
                                    INNER JOIN BAS.[User] c ON a.CreateBy = c.UserId
                                    INNER JOIN BAS.Department d ON b.DepartmentId = d.DepartmentId
                                    INNER JOIN BAS.Department e ON c.DepartmentId = e.DepartmentId
                                    INNER JOIN SCM.Customer f ON a.CustomerId = f.CustomerId
                                    INNER JOIN BAS.[Type] g ON a.Taxation = g.TypeNo AND g.TypeSchema = 'Customer.Taxation'
                                    WHERE SoId = @SoId";
                            dynamicParameters.Add("SoId", SoId);

                            saleOrders = sqlConnection.Query<SaleOrder>(sql, dynamicParameters).ToList();
                            if (saleOrders.Count() <= 0) throw new SystemException("訂單單頭資料錯誤!");
                            #endregion

                            #region //訂單單身
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.SoDetailId, a.SoSequence, a.SoMtlItemName, a.SoMtlItemSpec, a.SoQty, a.ProductType
                                    , a.FreebieQty, a.SpareQty, a.SoPriceQty, a.UnitPrice, a.Amount, FORMAT(a.PromiseDate, 'yyyy-MM-dd') PromiseDate
                                    , a.BusinessTaxRate, a.SoDetailRemark, a.Project, a.DiscountAmount
                                    , b.MtlItemNo, b.TypeThree, b.TypeFour
                                    , c.InventoryNo + ' ' + c.InventoryName InventoryFullNo
                                    , d.UomNo
                                    , f.Currency
                                    , g.UomNo SoPriceUomNo
                                    FROM SCM.SoDetail a 
                                    INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
                                    INNER JOIN SCM.Inventory c ON a.InventoryId = c.InventoryId
                                    INNER JOIN PDM.UnitOfMeasure d ON a.UomId = d.UomId
                                    INNER JOIN SCM.SaleOrder e ON a.SoId = e.SoId
                                    INNER JOIN SCM.Customer f ON e.CustomerId = f.CustomerId
                                    INNER JOIN PDM.UnitOfMeasure g ON a.SoPriceUomId = g.UomId
                                    WHERE a.SoId = @SoId
                                    ORDER BY a.SoId, a.SoSequence";
                            dynamicParameters.Add("SoId", SoId);

                            soDetails = sqlConnection.Query<SoDetail>(sql, dynamicParameters).ToList();
                            if (soDetails.Count() <= 0) throw new SystemException("訂單單身資料錯誤!");
                            if (soDetails.Count() != totalDetail) throw new SystemException("訂單單身數量錯誤!");
                            #endregion

                            #region //取得ERP訂單資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(a.TC007)) + ' ' + LTRIM(RTRIM(c.MB002)) Site
                                    , LTRIM(RTRIM(a.TC014)) TC014
                                    , LTRIM(RTRIM(b.MQ002)) MQ002
                                    FROM COPTC a    
                                    INNER JOIN CMSMQ b ON a.TC001 = b.MQ001
                                    INNER JOIN CMSMB c ON a.TC007 = c.MB001
                                    WHERE a.TC001 = @TC001
                                    AND a.TC002 = @TC002";
                            dynamicParameters.Add("TC001", saleOrders.FirstOrDefault().SoErpPrefix);
                            dynamicParameters.Add("TC002", saleOrders.FirstOrDefault().SoErpNo);

                            var COPTCResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (COPTCResult.Count() <= 0) throw new SystemException("查無ERP訂單資料!!");

                            string Site = COPTCResult.FirstOrDefault().Site;
                            saleOrders[0].PaymentTermName = COPTCResult.FirstOrDefault().TC014;
                            string SoErpPrefixName = COPTCResult.FirstOrDefault().MQ002;
                            #endregion

                            #region //組單身資料
                            JArray soDetailModels = new JArray();
                            List<SoDetailModel> soDetailDatas = new List<SoDetailModel>();
                            foreach (var item in soDetails)
                            {
                                #region //檢核卡控條件: 訂單數量>0 單價<0
                                if (item.SoPriceQty > 0 && item.UnitPrice <= 0)
                                {
                                    throw new SystemException("若計價數量大於0，單價及金額不可小於等於0!!");
                                }
                                #endregion

                                if (CompanyNo == "JMO")
                                {
                                    SoDetailModel soDetailModel = new SoDetailModel()
                                    {
                                        SoSequence = item.SoSequence,
                                        MtlItemNo = item.MtlItemNo,
                                        SoMtlItemName = item.SoMtlItemName,
                                        SoMtlItemSpec = item.SoMtlItemSpec,
                                        InventoryFullNo = item.InventoryFullNo,
                                        UomNo = item.UomNo,
                                        SoQty = item.SoQty,
                                        SpareQty = item.SpareQty,
                                        FreebieQty = item.FreebieQty,
                                        UnitPrice = item.UnitPrice,
                                        Amount = item.Amount,
                                        SoPriceQty = item.SoPriceQty,
                                        PromiseDate = Convert.ToDateTime(item.PromiseDate).ToString("yyyy/MM/dd"),
                                        SoDetailRemark = item.SoDetailRemark,
                                    };
                                    soDetailDatas.Add(soDetailModel);
                                }
                                else
                                {
                                    List<string> CounterSignDep = new List<string>();

                                    #region //處理加簽部分
                                    //以下為晶彩邏輯，中揚邏輯待確認
                                    if (item.Amount <= 0 && CurrentCompany == 4)
                                    {
                                        bool CheckFlag = false;
                                        CounterSign = 1;

                                        string TypeThree = item.TypeThree;
                                        string TypeFour = item.TypeFour;
                                        if (TypeThree == "301")
                                        {
                                            CounterSignDep.Add("J550");
                                            CounterSignDep.Add("J540");
                                            CheckFlag = true;
                                        }

                                        if (TypeThree == "302" || TypeThree == "308" || TypeThree == "309" || TypeThree == "310" || TypeThree == "315" || TypeThree == "316" || TypeThree == "307")
                                        {
                                            CounterSignDep.Add("J550");
                                            CheckFlag = true;
                                        }

                                        if (TypeThree == "303" || TypeThree == "312")
                                        {
                                            CounterSignDep.Add("J540");
                                            CheckFlag = true;
                                        }

                                        if (TypeThree == "304")
                                        {
                                            CounterSignDep.Add("J231");
                                            CheckFlag = true;
                                        }

                                        if (TypeThree == "318" || TypeThree == "319" || TypeThree == "320" || TypeThree == "321" || TypeThree == "322" || TypeThree == "323")
                                        {
                                            if (TypeFour == "403")
                                            {
                                                CounterSignDep.Add("J231");
                                                CheckFlag = true;
                                            }
                                        }

                                        if (TypeThree == "306" || TypeThree == "324")
                                        {
                                            CounterSignDep.Add("J630");
                                            CheckFlag = true;
                                        }

                                        if (TypeThree == "318" || TypeThree == "319" || TypeThree == "320" || TypeThree == "321" || TypeThree == "322" || TypeThree == "323")
                                        {
                                            if (TypeFour == "402")
                                            {
                                                CounterSignDep.Add("J630");
                                                CheckFlag = true;
                                            }
                                        }

                                        if (TypeThree == "326")
                                        {
                                            CounterSignDep.Add("J730");
                                            CounterSignDep.Add("J731");
                                            CheckFlag = true;
                                        }

                                        if (TypeThree == "305")
                                        {
                                            CounterSignDep.Add("J660");
                                            CounterSignDep.Add("J740");
                                            CheckFlag = true;
                                        }

                                        if (TypeThree == "318" || TypeThree == "319" || TypeThree == "320" || TypeThree == "321" || TypeThree == "322" || TypeThree == "323")
                                        {
                                            if (TypeFour == "401")
                                            {
                                                CounterSignDep.Add("J660");
                                                CounterSignDep.Add("J740");
                                                CheckFlag = true;
                                            }
                                        }

                                        if (TypeThree == "311" || TypeThree == "314" || TypeThree == "317" || TypeThree == "325")
                                        {
                                            CheckFlag = true;
                                        }

                                        if (TypeThree.Contains("37") || TypeThree.Contains("38") || TypeThree.Contains("39"))
                                        {
                                            CheckFlag = true;
                                        }

                                        if (CheckFlag == false)
                                        {
                                            throw new SystemException("此品號【" + item.MtlItemNo + "】類別尚未符合目前加簽規則，無法拋轉!!");
                                        }

                                        CounterSignDep = CounterSignDep.Distinct().ToList();
                                    }
                                    #endregion

                                    soDetailModels.Add(JObject.FromObject(new
                                    {
                                        item.SoSequence,
                                        item.MtlItemNo,
                                        item.SoMtlItemName,
                                        item.SoMtlItemSpec,
                                        item.InventoryFullNo,
                                        item.UomNo,
                                        item.SoQty,
                                        item.SpareQty,
                                        item.FreebieQty,
                                        item.UnitPrice,
                                        item.Amount,
                                        item.SoPriceQty,
                                        PromiseDate = Convert.ToDateTime(item.PromiseDate).ToString("yyyy/MM/dd"),
                                        item.SoDetailRemark,
                                        CounterSignDep
                                    }));
                                }
                            }
                            #endregion

                            #region //依公司別取得ProId
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.TypeName ProId
                                    FROM BAS.[Type] a
                                    WHERE a.TypeSchema = 'BPM.SoProId'
                                    AND a.TypeNo = @CompanyNo";
                            dynamicParameters.Add("CompanyNo", CompanyNo);

                            var ProIdResult = sqlConnection.Query(sql, dynamicParameters);

                            if (ProIdResult.Count() <= 0) throw new SystemException("此公司別【" + CompanyNo + "】尚未建立拋轉起單碼，請聯繫資訊人員!!");

                            string proId = "";
                            foreach (var item in ProIdResult)
                            {
                                proId = item.ProId;
                            }
                            #endregion

                            #region //計算訂單合計金額
                            int TotalQty = -1;
                            double Amount = -1;
                            double TaxAmount = -1;
                            double TotalAmount = -1;
                            string getTotalResponse = GetTotal(SoId);
                            JObject getTotalResponseJson = JObject.Parse(getTotalResponse);
                            if (getTotalResponseJson["status"].ToString() != "success")
                            {
                                throw new SystemException(getTotalResponseJson["msg"].ToString());
                            }
                            else
                            {
                                foreach (var item in getTotalResponseJson["data"])
                                {
                                    TotalQty = Convert.ToInt32(item["TotalQty"]);
                                    Amount = Convert.ToDouble(item["Amount"]);
                                    TaxAmount = Convert.ToDouble(item["TaxAmount"]);
                                    TotalAmount = Convert.ToDouble(item["TotalAmount"]);
                                }
                            }
                            #endregion

                            #region //組訂單BPM資料
                            string memId = BpmUserId;
                            string rolId = BpmRoleId;
                            string startMethod = "NoOpFirst";

                            JObject artInsAppData = new JObject();
                            foreach (var item in saleOrders)
                            {
                                if (CompanyNo == "JMO")
                                {
                                    SaleOrderModel data = new SaleOrderModel()
                                    {
                                        Title = "MES訂單:" + item.SoErpFullNo,
                                        Filler = UserName,
                                        FillerDepartment = DepartmentNo + "-" + DepartmentName,
                                        FillerNo = UserNo,
                                        FillCalendar = DateTime.Now.ToString("yyyy-MM-dd"),
                                        SoErpFullNo = item.SoErpFullNo,
                                        SoErpFullCa = item.SoErpPrefix + " " + SoErpPrefixName,
                                        SoErpNo = item.SoErpNo,
                                        createDate = Convert.ToDateTime(item.DocDate).ToString("yyyy/MM/dd"),
                                        DocDate = Convert.ToDateTime(item.DocDate).ToString("yyyy/MM/dd"),
                                        SalesmenNo = item.SalesmenNo,
                                        SalesmenName = item.SalesmenName,
                                        SalesName = item.SalesmenName,
                                        SalesNo = item.SalesmenNo,
                                        SalesmenDepNo = item.SalesmenDepNo,
                                        SalesmenDepName = item.SalesmenDepName,
                                        SalesDep = item.SalesmenDepNo + "-" + item.SalesmenDepName,
                                        CustomerFullNo = item.CustomerFullNo,
                                        CustomerPurchaseOrder = item.CustomerPurchaseOrder,
                                        Site = Site,
                                        SoRemark = item.SoRemark,
                                        Amount = Amount,
                                        TaxAmount = TaxAmount,
                                        DbTable = "SCM.SaleOrder",
                                        SoId = SoId,
                                        CompanyNo = CompanyNo,
                                        SoDetailModelList = soDetailDatas,
                                        CounterSign = CounterSign,
                                        Currency = item.Currency,
                                        ExchangeRate = item.ExchangeRate.ToString(),
                                        BusinessTaxRate = (item.BusinessTaxRate * 100).ToString() + "%",
                                        Taxation = item.Taxation,
                                        TotalDetail = TotalQty.ToString(),
                                        TotalAmount = TotalAmount,
                                        Customer = item.CustomerNo + "-" + item.CustomerShortName,
                                        CustomerOrder = item.CustomerPurchaseOrder,
                                        PhoneNo = item.TelNoFirst,
                                        Fax = item.FaxNo,
                                        DeliveryAdd = item.DeliveryAddressFirst,
                                        TaxRate = (item.BusinessTaxRate * 100).ToString() + "%",
                                        TaxID = item.GuiNumber,
                                        Payment = item.PaymentTermName,
                                        SoTotalAmount = TotalAmount + TaxAmount
                                    };

                                    string soDataJsonString = JsonConvert.SerializeObject(data);

                                    artInsAppData = JObject.FromObject(new
                                    {
                                        SoDataJson = soDataJsonString
                                    });
                                }
                                else
                                {
                                    artInsAppData = JObject.FromObject(new
                                    {
                                        Title = "MES訂單:" + item.SoErpFullNo,
                                        Filler = UserName,
                                        FillerDepartment = DepartmentNo + "-" + DepartmentName,
                                        FillerNo = UserNo,
                                        FillCalendar = DateTime.Now.ToString("yyyy-MM-dd"),
                                        item.SoErpFullNo,
                                        SoErpFullCa = item.SoErpPrefix,
                                        item.SoErpNo,
                                        createDate = Convert.ToDateTime(item.DocDate).ToString("yyyy/MM/dd"),
                                        DocDate = Convert.ToDateTime(item.DocDate).ToString("yyyy/MM/dd"),
                                        item.SalesmenNo,
                                        item.SalesmenName,
                                        SalesName = item.SalesmenName,
                                        SalesNo = item.SalesmenNo,
                                        item.SalesmenDepNo,
                                        item.SalesmenDepName,
                                        SalesDep = item.SalesmenDepNo + "-" + item.SalesmenDepName,
                                        item.CustomerFullNo,
                                        item.CustomerPurchaseOrder,
                                        Site = Site,
                                        item.SoRemark,
                                        Amount,
                                        TaxAmount,
                                        DbTable = "SCM.SaleOrder",
                                        SoId,
                                        CompanyNo,
                                        SoDetailModelList = JsonConvert.SerializeObject(soDetailModels),
                                        CounterSign,
                                        item.Currency,
                                        ExchangeRate = item.ExchangeRate.ToString(),
                                        BusinessTaxRate = (item.BusinessTaxRate * 100).ToString() + "%",
                                        item.Taxation,
                                        TotalDetail = TotalQty.ToString(),
                                        TotalAmount,
                                        Customer = item.CustomerNo + "-" + item.CustomerShortName,
                                        CustomerOrder = item.CustomerPurchaseOrder,
                                        PhoneNo = item.TelNoFirst,
                                        Fax = item.FaxNo,
                                        DeliveryAdd = item.DeliveryAddressFirst,
                                        TaxRate = (item.BusinessTaxRate * 100).ToString() + "%",
                                        TaxID = item.GuiNumber,
                                        Payment = item.PaymentTermName,
                                        SoTotalAmount = TotalAmount + TaxAmount
                                    });
                                }
                            }
                            #endregion

                            string sData = BpmHelper.PostFormToBpm(token, proId, memId, rolId, startMethod, artInsAppData, BpmServerPath);
                            //string sData = "false";
                            if (sData == "true")
                            {
                                if (SoBpmInfoId == null)
                                {
                                    #region //INSERT SCM.SoBpmInfo
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO SCM.SoBpmInfo (SoId, BpmNo, BpmTransferStatus, BpmTransferUserId, BpmTransferDate
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            VALUES (@SoId, @BpmNo, @BpmTransferStatus, @BpmTransferUserId, @BpmTransferDate
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                      new
                                      {
                                          SoId,
                                          BpmNo = "",
                                          BpmTransferStatus = "P",
                                          BpmTransferUserId = CreateBy,
                                          BpmTransferDate = CreateDate,
                                          CreateDate,
                                          LastModifiedDate,
                                          CreateBy,
                                          LastModifiedBy
                                      });
                                    var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult.Count();
                                    #endregion
                                }
                                else
                                {
                                    #region //UPDATE SCM.SoBpmInfo
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE SCM.SoBpmInfo SET
                                            BpmTransferStatus = 'P',
                                            BpmTransferUserId = @BpmTransferUserId,
                                            BpmTransferDate = @BpmTransferDate,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE SoId = @SoId";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            BpmTransferUserId = CreateBy,
                                            BpmTransferDate = CreateDate,
                                            LastModifiedDate,
                                            LastModifiedBy,
                                            SoId
                                        });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                            }
                            else
                            {
                                throw new SystemException("訂單拋轉BPM失敗!");
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

        #region //UpdateSaleOrderConfirmByBpm -- 訂單ERP確認(BPM流程) -- Ann 2023-12-06
        public string UpdateSaleOrderConfirmByBpm(int SoId, string CompanyNo, string ConfirmUser)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0, totalDetail = 0;
                    string confirmStatus = "", transferStatus = "", companyNo = "", soErpPrefix = "", soErpNo = "", departmentNo = "", userNo = "", userName = "";

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //公司別資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CompanyId, a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        dynamicParameters.Add("CompanyNo", CompanyNo);

                        var resultCompany = sqlConnection.Query(sql, dynamicParameters);
                        if (resultCompany.Count() <= 0) throw new SystemException("【公司別】資料錯誤!");

                        int CompanyId = -1;
                        foreach (var item in resultCompany)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            companyNo = item.ErpNo;
                            CompanyId = item.CompanyId;
                        }
                        #endregion

                        #region //判斷訂單資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 SoErpPrefix, SoErpNo, ConfirmStatus, TransferStatus
                                FROM SCM.SaleOrder
                                WHERE SoId = @SoId
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("SoId", SoId);
                        dynamicParameters.Add("CompanyId", CompanyId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("訂單資料錯誤!");

                        foreach (var item in result)
                        {
                            soErpPrefix = item.SoErpPrefix;
                            soErpNo = item.SoErpNo;
                            confirmStatus = item.ConfirmStatus;
                            transferStatus = item.TransferStatus;
                        }

                        #region //判斷訂單狀態
                        if (confirmStatus != "N") throw new SystemException("訂單不是未確認狀態，無法確認!");
                        if (transferStatus != "Y") throw new SystemException("訂單尚未拋轉，無法確認!");
                        #endregion
                        #endregion

                        #region //確認核單者資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT UserId, UserNo
                                FROM BAS.[User]
                                WHERE UserNo = @ConfirmUser";
                        dynamicParameters.Add("ConfirmUser", ConfirmUser);

                        var UserResult = sqlConnection.Query(sql, dynamicParameters);

                        if (UserResult.Count() <= 0) throw new SystemException("核單者資料錯誤!");

                        int UserId = -1;
                        foreach (var item in UserResult)
                        {
                            UserId = item.UserId;
                            userNo = item.UserNo;
                        }
                        #endregion

                        #region //單身筆數
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT COUNT(1) TotalDetail
                                FROM SCM.SoDetail
                                WHERE SoId = @SoId";
                        dynamicParameters.Add("SoId", SoId);

                        var resultDetailExist = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in resultDetailExist)
                        {
                            totalDetail = Convert.ToInt32(item.TotalDetail);
                            if (totalDetail <= 0) throw new SystemException("單身筆數錯誤!");
                        }
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        #region //判斷ERP使用者是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 MF005
                                FROM ADMMF
                                WHERE COMPANY = @CompanyNo
                                AND MF001 = @User";
                        dynamicParameters.Add("CompanyNo", companyNo);
                        dynamicParameters.Add("User", userNo);

                        var resultUserExist = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUserExist.Count() <= 0) throw new SystemException("ERP使用者不存在!");

                        string adminUser = "N";
                        foreach (var item in resultUserExist)
                        {
                            adminUser = item.MF005; //超級使用者
                        }
                        #endregion

                        #region //判斷ERP使用者是否有權限
                        if (companyNo != "DGJMO")
                        {
                            if (adminUser != "Y")
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 MG006
                                    FROM ADMMG
                                    WHERE COMPANY = @CompanyNo
                                    AND MG001 = @User
                                    AND MG002 = @Function
                                    AND MG004 = 'Y'
                                    AND MG006 LIKE @Auth";
                                dynamicParameters.Add("CompanyNo", companyNo);
                                dynamicParameters.Add("User", userNo);
                                dynamicParameters.Add("Function", "COPI06");
                                dynamicParameters.Add("Auth", "___Y________"); //確認

                                var resultUserAuthExist = sqlConnection.Query(sql, dynamicParameters);
                                if (resultUserAuthExist.Count() <= 0) throw new SystemException("ERP使用者無權限!");
                            }
                        }
                        #endregion

                        #region //判斷ERP單據是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 TC027
                                FROM COPTC
                                WHERE TC001 = @ErpPrefix
                                AND TC002 = @ErpNo";
                        dynamicParameters.Add("ErpPrefix", soErpPrefix);
                        dynamicParameters.Add("ErpNo", soErpNo);

                        var resultDocExist = sqlConnection.Query(sql, dynamicParameters);
                        if (resultDocExist.Count() <= 0) throw new SystemException("ERP單據不存在!");

                        foreach (var item in resultDocExist)
                        {
                            if (item.TC027 != "N") throw new SystemException("原單據確認碼為無法更動狀態!");
                        }
                        #endregion

                        int tempRowsAffected = 0;
                        #region //ERP確認
                        #region //COPTC
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE COPTC SET
                                TC027 = @ConfirmStatus,
                                TC040 = @ConfirmUser
                                WHERE TC001 = @ErpPrefix
                                AND TC002 = @ErpNo";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmStatus = "Y",
                                ConfirmUser = userNo,
                                ErpPrefix = soErpPrefix,
                                ErpNo = soErpNo
                            });

                        tempRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        if (tempRowsAffected != 1) throw new SystemException(string.Format("ERP單據 {0}-{1} 確認失敗!", soErpPrefix, soErpNo));
                        rowsAffected += tempRowsAffected;
                        #endregion

                        #region //COPTD
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE COPTD SET
                                TD021 = @ConfirmStatus
                                WHERE TD001 = @ErpPrefix
                                AND TD002 = @ErpNo";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmStatus = "Y",
                                ErpPrefix = soErpPrefix,
                                ErpNo = soErpNo
                            });

                        tempRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        if (tempRowsAffected != totalDetail) throw new SystemException(string.Format("ERP單據 {0}-{1} 單身確認失敗!", soErpPrefix, soErpNo));
                        rowsAffected += tempRowsAffected;
                        #endregion
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int tempRowsAffected = 0;
                        #region //BM確認
                        #region //SCM.SaleOrder
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.SaleOrder SET
                                ConfirmStatus = @ConfirmStatus,
                                ConfirmUserId = @ConfirmUserId,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SoId = @SoId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmStatus = "Y",
                                ConfirmUserId = CurrentUser,
                                LastModifiedDate,
                                LastModifiedBy,
                                SoId
                            });
                        tempRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        if (tempRowsAffected != 1) throw new SystemException(string.Format("BM單據 {0}-{1} 確認失敗!", soErpPrefix, soErpNo));
                        rowsAffected += tempRowsAffected;
                        #endregion

                        #region //SCM.SoDetail
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.SoDetail SET
                                ConfirmStatus = @ConfirmStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SoId = @SoId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmStatus = "Y",
                                LastModifiedDate,
                                LastModifiedBy,
                                SoId
                            });
                        tempRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        if (tempRowsAffected != totalDetail) throw new SystemException(string.Format("BM單據 {0}-{1} 單身確認失敗!", soErpPrefix, soErpNo));
                        rowsAffected += tempRowsAffected;
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

        #region //UpdateSoBpmInfo -- 更新訂單BPM簽核狀態 -- Ann 2023-12-06
        public string UpdateSoBpmInfo(int SoId, string BpmNo, string Status, string RootId, string ConfirmUser, string ErpFlag)
        {
            try
            {
                int rowsAffected = 0;
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷訂單資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT b.BpmTransferDate, b.BpmTransferStatus
                                , a.SoErpPrefix + '-' + a.SoErpNo SoErpFullNo
                                , a.SalesmenId
                                FROM SCM.SaleOrder a
                                INNER JOIN SCM.SoBpmInfo b ON a.SoId = b.SoId
                                WHERE a.SoId = @SoId";
                        dynamicParameters.Add("SoId", SoId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("訂單資料錯誤!");

                        DateTime BpmTransferDate = new DateTime();
                        string BpmTransferStatus = "";
                        string SoErpFullNo = "";
                        int SalesmenId = -1;
                        foreach (var item in result)
                        {
                            BpmTransferDate = item.BpmTransferDate;
                            BpmTransferStatus = item.BpmTransferStatus;
                            SoErpFullNo = item.SoErpFullNo;
                            SalesmenId = item.SalesmenId;
                        }
                        #endregion

                        #region //確認核單者資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT UserId
                                FROM BAS.[User]
                                WHERE UserNo = @ConfirmUser";
                        dynamicParameters.Add("ConfirmUser", ConfirmUser);

                        var UserResult = sqlConnection.Query(sql, dynamicParameters);

                        if (UserResult.Count() <= 0) throw new SystemException("核單者資料錯誤!");

                        int UserId = -1;
                        foreach (var item in UserResult)
                        {
                            UserId = item.UserId;
                        }
                        #endregion

                        #region //UPDATE SCM.SoBpmInfo
                        if (RootId.Length > 0) //BPM結束流程(E、Y)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.SoBpmInfo SET
                                    BpmTransferStatus = @BpmTransferStatus,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @UserId
                                    WHERE SoId = @SoId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    BpmTransferStatus = ErpFlag != "F" ? Status : "F",
                                    LastModifiedDate,
                                    UserId,
                                    SoId
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        }
                        else //BPM流程開始(P)、拋轉失敗(R)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.SoBpmInfo SET
                                    BpmTransferStatus = @BpmTransferStatus,
                                    BpmTransferDate = @BpmTransferDate,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @UserId
                                    WHERE SoId = @SoId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    BpmTransferStatus = Status,
                                    BpmTransferDate = DateTime.Now.ToString("yyyy-MM-dd"),
                                    LastModifiedDate,
                                    UserId,
                                    SoId
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        }
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            data = SoErpFullNo
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

        #region //UpdateSoDetailTempBatch --取得採購單匯入資料 -- Chia Yuan 2024.04.18
        public string UpdateSoDetailTempBatch(int SoId, int CompanyId, string SpreadsheetData)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        string currency = "", taxNo = "", taxation = "", customerNo = "";
                        int unitRound = 0, totalRound = 0;
                        decimal exciseTax = 0m;

                        var jsonData = JObject.Parse(SpreadsheetData);
                        if (jsonData["spreadsheetInfo"] != null)
                        {
                            #region //判斷欄位值是否有誤
                            List<string> msgs = new List<string>();
                            var vCol = new List<string> { "PoErpPrefixNo", "TD003", "TD008", "TD010", "TD011" };
                            vCol.ForEach(v =>
                            {
                                var vData = jsonData["spreadsheetInfo"].Select(s => new { MainKey = s["MainKey"]?.ToString(), t = decimal.TryParse(s[v].ToString(), out decimal d), value = d })
                                    .Where(w => w.value == 0m).ToList();

                                if (vData.Any()) msgs.Add(string.Join("</br>", vData.Select(s => string.Format("{0} {1}值有誤", s.MainKey, v))));
                            });
                            //if (msgs.Any()) throw new SystemException(string.Join("</br>", msgs));
                            #endregion

                            var companyIds = jsonData["spreadsheetInfo"].Select(s => { int.TryParse(s["CompanyId"]?.ToString(), out int i); return i; }).Distinct().ToList();
                            var customerMtlItemIds = jsonData["spreadsheetInfo"].Select(s => { int.TryParse(s["CustomerMtlItemId"]?.ToString(), out int i); return i; }).Distinct().ToList();

                            #region //判斷客戶部番是否正確
                            if (customerMtlItemIds.Count > 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT 1
                                        FROM PDM.CustomerMtlItem
                                        WHERE CustomerMtlItemId IN @CustomerMtlItemIds";
                                dynamicParameters.Add("CustomerMtlItemIds", customerMtlItemIds);
                                var resultCustMtlItem = sqlConnection.Query(sql, dynamicParameters);
                                //if (resultCustMtlItem.Count() != customerMtlItemIds.Count) throw new SystemException("【部番】資料錯誤!");
                            }
                            #endregion

                            #region //取得訂單單頭資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.Currency, a.TaxNo, a.ConfirmStatus
                                    , b.CustomerNo
                                    FROM SCM.SaleOrder a 
                                    INNER JOIN SCM.Customer b ON a.CustomerId = b.CustomerId
                                    WHERE a.SoId = @SoId
                                    AND a.CompanyId = @CompanyId";
                            dynamicParameters.Add("SoId", SoId);
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【訂單】資料錯誤!");
                            string confirmStatus = "";
                            foreach (var item in result)
                            {
                                currency = item.Currency;
                                taxNo = item.TaxNo;
                                confirmStatus = item.ConfirmStatus;
                                customerNo = item.CustomerNo;
                            }
                            if (confirmStatus != "N") throw new SystemException("【訂單】已確認，無法新增!");
                            #endregion

                            #region //公司別資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.ErpNo, a.ErpDb
                                    FROM BAS.Company a
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            var resultCompany = sqlConnection.Query(sql, dynamicParameters);
                            if (resultCompany.Count() <= 0) throw new SystemException("【公司別】資料錯誤!");
                            string companyNo = string.Empty;
                            foreach (var item in resultCompany)
                            {
                                ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                                companyNo = item.ErpNo;
                            }
                            #endregion

                            List<dynamic> resultCMSNB = new List<dynamic>();
                            using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                            {
                                #region //交易幣別設定
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT MF003, MF004, MF005, MF006
                                        FROM CMSMF
                                        WHERE MF001 = @Currency";
                                dynamicParameters.Add("Currency", currency);

                                var resultCMSMF = erpConnection.Query(sql, dynamicParameters);
                                if (resultCMSMF.Count() <= 0) throw new SystemException("ERP交易幣別設定不存在!");

                                foreach (var item in resultCMSMF)
                                {
                                    unitRound = Convert.ToInt32(item.MF003); //單價取位
                                    totalRound = Convert.ToInt32(item.MF004); //金額取位
                                }
                                #endregion

                                #region //稅別碼設定
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT RTRIM(LTRIM(NN001)) NN001, RTRIM(LTRIM(NN002)) NN002, NN004, RTRIM(LTRIM(NN006)) NN006
                                        FROM CMSNN
                                        WHERE NN001 = @TaxNo";
                                dynamicParameters.Add("TaxNo", taxNo);

                                var resultCMSNN = erpConnection.Query(sql, dynamicParameters);
                                if (resultCMSNN.Count() <= 0) throw new SystemException("ERP稅別碼設定不存在!");

                                foreach (var item in resultCMSNN)
                                {
                                    exciseTax = Convert.ToDecimal(item.NN004); //營業稅率
                                    taxation = item.NN006; //課稅別
                                }
                                #endregion

                                var projects = jsonData["spreadsheetInfo"].Where(w => !string.IsNullOrWhiteSpace(w["Project"]?.ToString())).Select(s => s["Project"]?.ToString()).Distinct().ToList();

                                #region //取得專案資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT LTRIM(RTRIM(NB001)) Project, LTRIM(RTRIM(NB002)) ProjectName 
                                        FROM CMSNB
                                        WHERE LTRIM(RTRIM(NB001)) IN @Projects";
                                dynamicParameters.Add("Projects", projects);
                                resultCMSNB = erpConnection.Query(sql, dynamicParameters).ToList();
                                #endregion
                            }

                            #region //取得客戶公司別DB
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.CompanyId, a.ErpNo, a.ErpDb
                                    FROM BAS.Company a
                                    WHERE CompanyId = @CompanyIds";
                            dynamicParameters.Add("CompanyIds", companyIds);
                            var resultCorp = sqlConnection.Query(sql, dynamicParameters);
                            if (resultCorp.Count() <= 0) throw new SystemException("【公司別】資料錯誤!");
                            #endregion

                            List<dynamic> resultPURTD = new List<dynamic>();
                            foreach (var corp in resultCorp)
                            {
                                var tempData = jsonData["spreadsheetInfo"].Where(w => w["CompanyId"] == corp.CompanyId.ToString());
                                var newMainKeys = tempData.Select(s => s["MainKey"]?.ToString()).ToList();

                                ErpConnectionStrings = ConfigurationManager.AppSettings[corp.ErpDb];
                                using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                                {
                                    #region //取得ERP客戶採購單
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT @CompanyId, c.MainKey, a.COMPANY
                                            , LTRIM(RTRIM(a.TD001)) AS TD001, LTRIM(RTRIM(a.TD002)) AS TD002, LTRIM(RTRIM(a.TD003)) AS TD003
                                            , LTRIM(RTRIM(a.TD001)) + '-' + LTRIM(RTRIM(a.TD002)) AS PoErpPrefixNo
                                            , LTRIM(RTRIM(a.TD004)) + ' ' + LTRIM(RTRIM(a.TD005)) AS MtlItemNameWithNo
                                            , LTRIM(RTRIM(b.ML002)) AS ML002, LTRIM(RTRIM(b.ML003)) AS ML003
                                            , LTRIM(RTRIM(a.TD004)) AS TD004, LTRIM(RTRIM(a.TD005)) AS TD005, LTRIM(RTRIM(a.TD006)) AS TD006
                                            , a.TD007, a.TD008, LTRIM(RTRIM(a.TD009)) AS TD009, a.TD010, a.TD011
                                            , a.TD012, a.TD045, a.TD046, a.TD015
                                            , a.TD024, a.TD057, a.TD058, a.TD059, a.TD060
                                            , a.TD061, a.TD062, a.TD063
                                            FROM PURTD a
                                            INNER JOIN CMSML b ON b.COMPANY = a.COMPANY
                                            OUTER APPLY(SELECT LTRIM(RTRIM(TD001)) + LTRIM(RTRIM(TD002)) + LTRIM(RTRIM(TD003)) AS MainKey FROM PURTD WHERE TD001 = a.TD001 AND TD002 = a.TD002 AND TD003 = a.TD003) c
                                            WHERE c.MainKey IN @MainKeys";
                                    dynamicParameters.Add("CompanyId", corp.CompanyId);
                                    dynamicParameters.Add("MainKeys", newMainKeys);
                                    resultPURTD.AddRange(erpConnection.Query(sql, dynamicParameters).ToList());
                                    #endregion
                                }

                                #region //取得採購單暫存資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.SoDetailTempId
                                        , a.SoId, a.SoDetailId, a.BaseDocType, a.BaseErpPrefix, a.BaseErpNo, a.BaseErpSNo
                                        , a.CustomerMtlItemId, a.UomNo, a.SoQty, a.UnitPrice, a.Amount
                                        , a.BaseMtlItemNo, a.BaseMtlItemName, a.BaseMtlItemSpec, a.SortNumber
                                        , a.[Status] AS SoDetailTempStatus, s.StatusName AS SoDetailTempStatusName
                                        , b.*
                                        , c.CompanyId, c.CompanyNo, c.CompanyName, k.MainKey
                                        FROM SCM.SoDetailTemp a
                                        INNER JOIN BAS.[Status] s ON s.StatusNo = a.[Status] AND s.StatusSchema = 'Boolean'
                                        INNER JOIN BAS.Company c ON c.CompanyId = a.CompanyId
                                        OUTER APPLY(SELECT BaseErpPrefix + BaseErpNo + BaseErpSNo AS MainKey FROM SCM.SoDetailTemp WHERE BaseErpPrefix = a.BaseErpPrefix AND BaseErpNo = a.BaseErpNo AND BaseErpSNo = a.BaseErpSNo) k
                                        LEFT JOIN (
	                                        SELECT a.CustomerMtlItemId, a.CustomerMtlItemNo, a.CustomerMtlItemName, a.[Status] AS CustMtlItemStatus
	                                        , a.TransferStatus AS CustMtlItemTransferStatus, s1.StatusName AS CustMtlItemTransferStatusName
	                                        , b.MtlItemId, b.MtlItemNo, b.MtlItemName, b.MtlItemSpec, b.MtlItemDesc, b.[Status] AS MtlItemStatus
	                                        , b.TransferStatus AS MtlItemTransferStatus, s2.StatusName AS MtlItemTransferStatusName
                                            , ISNULL(u.UomNo, '') AS SaleUomNo, ISNULL(u.UomName, '') AS SaleUomName
	                                        , ROW_NUMBER() OVER (PARTITION BY a.CustomerMtlItemNo ORDER BY a.CustomerMtlItemId DESC) AS SortCustMtlItem
	                                        FROM PDM.CustomerMtlItem a 
	                                        INNER JOIN PDM.MtlItem b ON b.MtlItemId = a.MtlItemId AND b.[Status] = a.[Status]
                                            LEFT JOIN PDM.UnitOfMeasure u ON u.UomId = b.SaleUomId
	                                        INNER JOIN BAS.[Status] s1 ON s1.StatusNo = a.TransferStatus AND s1.StatusSchema = 'Boolean'
	                                        INNER JOIN BAS.[Status] s2 ON s2.StatusNo = b.TransferStatus AND s2.StatusSchema = 'Boolean'
                                        ) b ON b.CustomerMtlItemId = a.CustomerMtlItemId
                                        WHERE a.SoId = @SoId 
                                        AND a.CompanyId = @CompanyId";
                                dynamicParameters.Add("SoId", SoId);
                                dynamicParameters.Add("CompanyId", corp.CompanyId);
                                var resultOld = sqlConnection.Query(sql, dynamicParameters);
                                #endregion

                                var oldMainKeys = resultOld.Select(s => { return (string)s.MainKey; }).ToList();

                                #region //刪除舊資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE SCM.SoDetailTemp
                                        WHERE SoDetailTempId IN @SoDetailTempIds";
                                dynamicParameters.Add("SoDetailTempIds", resultOld.Where(w => !newMainKeys.Contains(w.MainKey)).Select(s => s.SoDetailTempId).ToList());
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region //排序
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE SCM.SoDetailTemp SET
                                        SortNumber = SortNumber * -1,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE SoId = @SoId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        SoId
                                    });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region //更新現有資料
                                var updMainKeys = oldMainKeys.Where(w => newMainKeys.Contains(w)).ToList();
                                foreach (var item in jsonData["spreadsheetInfo"].Where(w => updMainKeys.Contains(w["MainKey"]?.ToString())))
                                {
                                    var erpData = resultPURTD.FirstOrDefault(f => f.MainKey == item["MainKey"]?.ToString());
                                    if (erpData == null) throw new SystemException("【ERP採購單】資料錯誤!");

                                    var data = resultOld.FirstOrDefault(f => f.MainKey == item["MainKey"]?.ToString());

                                    int.TryParse(item["Row"]?.ToString(), out int sortNumber);
                                    int? customerMtlItemId = int.TryParse(item["CustomerMtlItemId"]?.ToString(), out int i) ? i : (int?)null;
                                    int.TryParse(item["TD008"]?.ToString(), out int soQty);
                                    double.TryParse(item["TD010"]?.ToString(), out double unitPrice);
                                    double.TryParse(item["TD011"]?.ToString(), out double amount);

                                    if (unitPrice == 0) throw new SystemException("【單價】不能為空!");
                                    if (amount == 0) throw new SystemException("【金額】不能為空!");

                                    #region //更新
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE SCM.SoDetailTemp SET
                                            CompanyId = @CompanyId,
                                            BaseDocType = @BaseDocType,
                                            CustomerMtlItemId = @CustomerMtlItemId,
                                            UomNo = @UomNo,
                                            SoQty = @SoQty,
                                            UnitPrice = @UnitPrice,
                                            Amount = @Amount,
                                            BaseMtlItemSpec = @BaseMtlItemSpec,
                                            BaseMtlItemNo = @BaseMtlItemNo,
                                            BaseMtlItemName = @BaseMtlItemName,
                                            SortNumber = @SortNumber,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE SoDetailTempId = @SoDetailTempId";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            data.SoDetailTempId,
                                            corp.CompanyId,
                                            BaseDocType = "E", //客戶委託
                                            CustomerMtlItemId = customerMtlItemId,
                                            UomNo = erpData.TD009,
                                            SoQty = soQty,
                                            UnitPrice = unitPrice,
                                            Amount = amount,
                                            BaseMtlItemNo = erpData.TD004,
                                            BaseMtlItemName = erpData.TD005,
                                            BaseMtlItemSpec = erpData.TD006,
                                            SortNumber = sortNumber,
                                            LastModifiedDate,
                                            LastModifiedBy
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                                #endregion

                                #region //加入新資料
                                var addMainKeys = newMainKeys.Where(w => !oldMainKeys.Contains(w)).ToList();
                                foreach (var item in jsonData["spreadsheetInfo"].Where(w => addMainKeys.Contains(w["MainKey"]?.ToString())))
                                {
                                    var erpData = resultPURTD.FirstOrDefault(f => f.MainKey == item["MainKey"]?.ToString());
                                    if (erpData == null) throw new SystemException("【ERP採購單】資料錯誤!");

                                    var data = resultOld.FirstOrDefault(f => f.MainKey == item["MainKey"]?.ToString());

                                    int.TryParse(item["Row"]?.ToString(), out int sortNumber);
                                    int? customerMtlItemId = int.TryParse(item["CustomerMtlItemId"]?.ToString(), out int i) ? i : (int?)null;
                                    int.TryParse(item["TD008"]?.ToString(), out int soQty);
                                    double.TryParse(item["TD010"]?.ToString(), out double unitPrice);
                                    double.TryParse(item["TD011"]?.ToString(), out double amount);

                                    if (unitPrice == 0) throw new SystemException("【單價】不能為空!");
                                    if (amount == 0) throw new SystemException("【金額】不能為空!");

                                    #region //新增
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO SCM.SoDetailTemp (SoId, CompanyId
                                            , BaseDocType, BaseErpPrefix, BaseErpNo, BaseErpSNo, CustomerMtlItemId
                                            , UomNo, SoQty, UnitPrice, Amount, BaseMtlItemNo, BaseMtlItemName, BaseMtlItemSpec
                                            , SortNumber, Status
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.SoDetailTempId
                                            VALUES (@SoId, @CompanyId
                                            , @BaseDocType, @BaseErpPrefix, @BaseErpNo, @BaseErpSNo, @CustomerMtlItemId
                                            , @UomNo, @SoQty, @UnitPrice, @Amount, @BaseMtlItemNo, @BaseMtlItemName, @BaseMtlItemSpec
                                            , @SortNumber, @Status
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            SoId,
                                            corp.CompanyId,
                                            BaseDocType = "E", //客戶委託
                                            BaseErpPrefix = erpData.TD001,
                                            BaseErpNo = erpData.TD002,
                                            BaseErpSNo = erpData.TD003,
                                            CustomerMtlItemId = customerMtlItemId,
                                            UomNo = erpData.TD009,
                                            SoQty = soQty,
                                            UnitPrice = unitPrice,
                                            Amount = amount,
                                            BaseMtlItemNo = erpData.TD004,
                                            BaseMtlItemName = erpData.TD005,
                                            BaseMtlItemSpec = erpData.TD006,
                                            SortNumber = sortNumber,
                                            Status = "N", //未綁定
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy,

                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                                #endregion
                            }

                            #region //解除訂單身綁定 (停用)
                            //dynamicParameters = new DynamicParameters();
                            //sql = @"UPDATE SCM.SoDetailTemp SET
                            //        SoDetailId = NULL,
                            //        Status = 'N'
                            //        WHERE SoId = @SoId";
                            //dynamicParameters.AddDynamicParams(
                            //    new
                            //    {
                            //        SoId
                            //    });
                            #endregion

                            #region //重新取得訂單明細暫存資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.SoDetailTempId
                                    ,a.SoId,a.CompanyId,ISNULL(a.SoDetailId, 0) AS SoDetailId,a.BaseDocType,a.BaseErpPrefix,a.BaseErpNo,a.BaseErpSNo
                                    ,a.CustomerMtlItemId,a.UomNo,a.SoQty,a.UnitPrice,a.Amount
                                    ,a.BaseMtlItemNo,a.BaseMtlItemName,a.BaseMtlItemSpec,a.SortNumber,a.Status, k.MainKey
                                    ,b.*
                                    FROM SCM.SoDetailTemp a
                                    OUTER APPLY(SELECT BaseErpPrefix + BaseErpNo + BaseErpSNo AS MainKey FROM SCM.SoDetailTemp WHERE BaseErpPrefix = a.BaseErpPrefix AND BaseErpNo = a.BaseErpNo AND BaseErpSNo = a.BaseErpSNo AND SoId = a.SoId) k
                                    LEFT JOIN (
	                                    SELECT a.CustomerMtlItemId, a.CustomerMtlItemNo, a.CustomerMtlItemName, a.[Status] AS CustMtlItemStatus
	                                    , a.TransferStatus AS CustMtlItemTransferStatus, s1.StatusName AS CustMtlItemTransferStatusName
	                                    , b.MtlItemId, b.MtlItemNo, b.MtlItemName, b.MtlItemSpec, b.MtlItemDesc, b.[Status] AS MtlItemStatus
	                                    , b.TransferStatus AS MtlItemTransferStatus, s2.StatusName AS MtlItemTransferStatusName
                                        , ISNULL(b.SaleUomId, b.InventoryUomId) SaleUomId
                                        , ISNULL(u.UomNo, '') AS SaleUomNo, ISNULL(u.UomName, '') AS SaleUomName
                                        , i.InventoryId, i.InventoryNo, i.InventoryName
	                                    , ROW_NUMBER() OVER (PARTITION BY a.CustomerMtlItemNo ORDER BY a.CustomerMtlItemId DESC) AS SortCustMtlItem
	                                    FROM PDM.CustomerMtlItem a 
	                                    INNER JOIN PDM.MtlItem b ON b.MtlItemId = a.MtlItemId AND b.[Status] = a.[Status]
                                        INNER JOIN SCM.Customer c ON c.CustomerId = a.CustomerId
                                        INNER JOIN SCM.Inventory i ON i.InventoryId = b.InventoryId
                                        LEFT JOIN PDM.UnitOfMeasure u ON u.UomId = b.SaleUomId
	                                    INNER JOIN BAS.[Status] s1 ON s1.StatusNo = a.TransferStatus AND s1.StatusSchema = 'Boolean'
	                                    INNER JOIN BAS.[Status] s2 ON s2.StatusNo = b.TransferStatus AND s2.StatusSchema = 'Boolean'
                                    ) b ON b.CustomerMtlItemId = a.CustomerMtlItemId
                                    WHERE a.SoId = @SoId
                                    ORDER BY a.SortNumber";
                            dynamicParameters.Add("SoId", SoId);
                            var resultSoDetailTemp = sqlConnection.Query(sql, dynamicParameters);
                            if (resultSoDetailTemp.Count() <= 0) throw new SystemException("【採購單暫存資料】資料錯誤!");
                            #endregion

                            #region //刪除未綁定的訂單單身
                            dynamicParameters = new DynamicParameters(); //Delete SCM.SoDetail
                            sql = @"DELETE a
                                    FROM SCM.SoDetail a
                                    WHERE a.SoId = @SoId 
                                    AND NOT EXISTS(
	                                    SELECT DISTINCT ba.SoDetailId 
	                                    FROM SCM.SoDetailTemp ba
	                                    WHERE ba.SoId = a.SoId 
	                                    AND ba.SoDetailId = a.SoDetailId
                                    )";
                            dynamicParameters.Add("SoId", SoId);
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            foreach (var item in resultSoDetailTemp.Where(w => w.MtlItemId != null))
                            {
                                //if (item.MtlItemId == null) throw new SystemException("【品號/部番】上未勾稽!"); //判斷採購資料是否與品號有關連

                                var tempData = jsonData["spreadsheetInfo"].FirstOrDefault(w => w["MainKey"] == item.MainKey && w["CompanyId"] == item.CompanyId); //取得傳入的excel資料
                                if (tempData == null) throw new SystemException("【採購單】資料錯誤!");

                                int uomId = Convert.ToInt32(item.SaleUomId);
                                int soQty = Convert.ToInt32(item.SoQty);
                                decimal unitPrice = Convert.ToDecimal(item.UnitPrice);
                                decimal amount = Convert.ToDecimal(item.Amount);

                                int.TryParse(tempData["FreebieOrSpareQty"]?.ToString(), out int freebieOrSpareQty);
                                
                                string productType = tempData["ProductType"]?.ToString();
                                string project = tempData["Project"]?.ToString();
                                string soDetailRemark = tempData["SoDetailRemark"]?.ToString();

                                if (string.IsNullOrWhiteSpace(item.MtlItemName)) throw new SystemException("【品名】不能為空!");
                                if (string.IsNullOrWhiteSpace(item.MtlItemSpec)) throw new SystemException("【規格】不能為空!");
                                if (unitPrice < 0) throw new SystemException("【單價】不能為空!");
                                if (amount < 0) throw new SystemException("【金額】不能為空!");
                                if (productType.Length <= 0) throw new SystemException("【類型】不能為空!");
                                if (freebieOrSpareQty < 0) throw new SystemException("【贈/備品量】不能為空!");
                                if (!DateTime.TryParse(tempData["PromiseDate"]?.ToString(), out DateTime promiseDate)) throw new SystemException("【預計出貨日】格式錯誤!");
                                if (project.Length > 20) throw new SystemException("【專案代號】長度錯誤!");
                                if (soDetailRemark.Length > 255) throw new SystemException("【備註】長度錯誤!");

                                var CMSNB = resultCMSNB.FirstOrDefault(w => w.Project == project); //ERP專案資料

                                if (item.SoDetailId > 0) //Update SCM.SoDetail
                                {
                                    #region //更新已綁定的訂單單身
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE SCM.SoDetail SET
                                            MtlItemId = @MtlItemId,
                                            SoMtlItemName = @SoMtlItemName,
                                            SoMtlItemSpec = @SoMtlItemSpec,
                                            InventoryId = @InventoryId,
                                            UomId = @UomId,
                                            SoQty = @SoQty,
                                            ProductType = @ProductType,
                                            FreebieQty = @FreebieQty,
                                            SpareQty = @SpareQty,
                                            UnitPrice = @UnitPrice,
                                            Amount = @Amount,
                                            PromiseDate = @PromiseDate,
                                            PcPromiseDate = @PromiseDate,
                                            Project = @Project,
                                            CustomerMtlItemNo = @CustomerMtlItemNo,
                                            SoDetailRemark = @SoDetailRemark,
                                            SoPriceQty = @SoQty,
                                            SoPriceUomId = @UomId,
                                            TaxNo = @TaxNo,
                                            BusinessTaxRate = @BusinessTaxRate,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE SoDetailId = @SoDetailId";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            item.SoDetailId,
                                            item.MtlItemId,
                                            SoMtlItemName = item.MtlItemName,
                                            SoMtlItemSpec = item.MtlItemSpec,
                                            item.InventoryId,
                                            UomId = uomId <= 0 ? (int?)null : uomId,
                                            SoQty = soQty,
                                            ProductType = productType,
                                            FreebieQty = productType == "1" ? freebieOrSpareQty : 0,
                                            SpareQty = productType == "2" ? freebieOrSpareQty : 0,
                                            UnitPrice = Math.Round(unitPrice, unitRound, MidpointRounding.AwayFromZero),
                                            Amount = Math.Round(Math.Round(unitPrice, unitRound, MidpointRounding.AwayFromZero) * soQty, totalRound, MidpointRounding.AwayFromZero),
                                            PromiseDate = promiseDate.ToString("yyyy-MM-dd"),
                                            PcPromiseDate = promiseDate.ToString("yyyy-MM-dd"),
                                            Project = CMSNB == null ? "" : CMSNB.Project,
                                            item.CustomerMtlItemNo,
                                            SoDetailRemark = soDetailRemark.Trim(),
                                            TaxNo = taxNo,
                                            BusinessTaxRate = exciseTax,
                                            LastModifiedDate,
                                            LastModifiedBy
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                                else //Add SCM.SoDetail
                                {
                                    #region //取得訂單單身目前序號
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT ISNULL(CAST(MAX(SoSequence) AS INT), 0) + 1 MaxSequence
                                            FROM SCM.SoDetail
                                            WHERE SoId = @SoId";
                                    dynamicParameters.Add("SoId", SoId);
                                    var result2 = sqlConnection.Query(sql, dynamicParameters);
                                    int maxSequence = 1;
                                    foreach (var seq in result2)
                                    {
                                        maxSequence = Convert.ToInt32(seq.MaxSequence);
                                    }
                                    #endregion

                                    #region //新增訂單單身
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO SCM.SoDetail (SoId, SoSequence, MtlItemId, SoMtlItemName
                                            , SoMtlItemSpec, InventoryId, UomId, SoQty, SiQty, ProductType, FreebieQty
                                            , FreebieSiQty, SpareQty, SpareSiQty, UnitPrice, Amount, PromiseDate, PcPromiseDate
                                            , Project, CustomerMtlItemNo, SoDetailRemark, SoPriceQty, SoPriceUomId
                                            , TaxNo, BusinessTaxRate, DiscountRate, DiscountAmount
                                            , ConfirmStatus, ClosureStatus
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.SoDetailId
                                            VALUES (@SoId, @SoSequence, @MtlItemId, @SoMtlItemName
                                            , @SoMtlItemSpec, @InventoryId, @UomId, @SoQty, @SiQty, @ProductType, @FreebieQty
                                            , @FreebieSiQty, @SpareQty, @SpareSiQty, @UnitPrice, @Amount, @PromiseDate, @PcPromiseDate
                                            , @Project, @CustomerMtlItemNo, @SoDetailRemark, @SoPriceQty, @SoPriceUomId
                                            , @TaxNo, @BusinessTaxRate, @DiscountRate, @DiscountAmount
                                            , @ConfirmStatus, @ClosureStatus
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            SoId,
                                            SoSequence = string.Format("{0:0000}", maxSequence),
                                            item.MtlItemId,
                                            SoMtlItemName = item.MtlItemName,
                                            SoMtlItemSpec = item.MtlItemSpec,
                                            item.InventoryId,
                                            UomId = uomId <= 0 ? (int?)null : uomId,
                                            SoQty = soQty,
                                            SiQty = 0, //已交數量
                                            ProductType = productType,
                                            FreebieQty = productType == "1" ? freebieOrSpareQty : 0,
                                            FreebieSiQty = 0, //贈品已交數量
                                            SpareQty = productType == "2" ? freebieOrSpareQty : 0,
                                            SpareSiQty = 0, //備品已交數量
                                            UnitPrice = Math.Round(unitPrice, unitRound, MidpointRounding.AwayFromZero),
                                            Amount = Math.Round(Math.Round(unitPrice, unitRound, MidpointRounding.AwayFromZero) * soQty, totalRound, MidpointRounding.AwayFromZero),
                                            PromiseDate = promiseDate.ToString("yyyy-MM-dd"),
                                            PcPromiseDate = promiseDate.ToString("yyyy-MM-dd"), //排定交貨日
                                            Project = CMSNB == null ? "" : CMSNB.Project,
                                            CustomerMtlItemNo = "", //客戶品號
                                            SoDetailRemark = soDetailRemark.Trim(),
                                            SoPriceQty = soQty, //計價數量
                                            SoPriceUomId = uomId <= 0 ? (int?)null : uomId, //計價單位
                                            TaxNo = taxNo, //稅別碼
                                            BusinessTaxRate = exciseTax, //營業稅率
                                            DiscountRate = 1, //折扣率
                                            DiscountAmount = 0, //折扣金額
                                            ConfirmStatus = "N", //確認碼
                                            ClosureStatus = "N", //結案碼
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });
                                    var insertResult = sqlConnection.Query(sql, dynamicParameters);
                                    rowsAffected += insertResult.Count();
                                    int soDetailId = -1;
                                    foreach (var soDetail in insertResult)
                                    {
                                        soDetailId = soDetail.SoDetailId;
                                    }
                                    #endregion

                                    #region //訂單身暫存綁定
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE SCM.SoDetailTemp SET 
                                            SoDetailId = @SoDetailId,
                                            Status = @Status,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE SoDetailTempId = @SoDetailTempId";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            item.SoDetailTempId,
                                            SoDetailId = soDetailId,
                                            Status = "Y", //綁定訂單身
                                            LastModifiedDate,
                                            LastModifiedBy
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                            }

                            int totalQty = 0;
                            decimal SoTotalAmount = 0m, SoAmount = 0m, SoTaxAmount = 0m;

                            #region //計算訂單單身單價數量及總金額
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT SoQty, SoPriceQty, ProductType, FreebieQty, SpareQty, UnitPrice, Amount
                                    FROM SCM.SoDetail
                                    WHERE SoId = @SoId
                                    ORDER BY SoSequence";
                            dynamicParameters.Add("SoId", SoId);
                            var resultSoDetail = sqlConnection.Query(sql, dynamicParameters);
                            //if (resultSoDetail.Count() <= 0) throw new SystemException("無單身資料!");

                            foreach (var item in resultSoDetail)
                            {
                                totalQty += Convert.ToInt32(item.SoQty) + Convert.ToInt32(item.FreebieQty) + Convert.ToInt32(item.SpareQty);
                                SoAmount += Math.Round(Math.Round(Convert.ToDecimal(item.UnitPrice), unitRound, MidpointRounding.AwayFromZero) * Convert.ToInt32(item.SoPriceQty), totalRound, MidpointRounding.AwayFromZero);
                            }

                            switch (taxation)
                            {
                                case "1":
                                    SoTaxAmount = SoAmount - Math.Round(SoAmount / (1 + exciseTax), totalRound, MidpointRounding.AwayFromZero);
                                    SoAmount = Math.Round(SoAmount / (1 + exciseTax), totalRound, MidpointRounding.AwayFromZero);
                                    break;
                                case "2":
                                    SoTaxAmount = Math.Round(SoAmount * exciseTax, totalRound, MidpointRounding.AwayFromZero);
                                    break;
                                case "3":
                                    break;
                                case "4":
                                    break;
                                case "9":
                                    break;
                            }

                            SoTotalAmount = Convert.ToDecimal(SoAmount + SoTaxAmount);
                            #endregion

                            #region //訂單單頭資料更新
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.SaleOrder SET
                                    TotalQty = @TotalQty,
                                    Amount = @Amount,
                                    TaxAmount = @TaxAmount,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE SoId = @SoId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    TotalQty = totalQty,
                                    Amount = SoTotalAmount,
                                    TaxAmount = SoTaxAmount,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    SoId
                                });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }
                        else
                        {
                            #region //訂單明細暫存資料刪除
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE a
                                    FROM SCM.SoDetailTemp a
                                    WHERE a.SoId = @SoId";
                            dynamicParameters.Add("SoId", SoId);
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

        #region //UpdateSoBpmLogErrorMessage -- 更新訂單拋轉BPM LOG回傳資訊 -- Ann 2024-05-20
        public string UpdateSoBpmLogErrorMessage(int SoBpmLogId, string ErrorMessage)
        {
            try
            {
                int rowsAffected = 0;
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷訂單LOG紀錄資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.SoBpmLog a 
                                WHERE a.SoBpmLogId = @SoBpmLogId";
                        dynamicParameters.Add("SoBpmLogId", SoBpmLogId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("訂單LOG資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.SoBpmLog SET
                                ErrorMessage = @ErrorMessage,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SoBpmLogId = @SoBpmLogId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ErrorMessage,
                                LastModifiedDate,
                                LastModifiedBy,
                                SoBpmLogId
                            });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

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

        #region //UploadPmdOrderData -- 上傳PMD資料 -- GPAI 2025-05-06
        public string UploadPmdOrderData(string PmdJsonData, int TypeId, int customerId)
        {
            JObject jsonResponse;
            int rowsAffected = 0;
            string ErpConnectionStrings = "", ErpNo = "";
            int logPmdId = -1;
            try
            {


                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {

                    #region //判斷公司資料是否正確
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 CompanyId, CompanyName, ErpDb,ErpNo
                                FROM BAS.Company
                                WHERE CompanyId = @CurrentCompany";
                    dynamicParameters.Add("CurrentCompany", CurrentCompany);

                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() != 1) throw new SystemException("【公司】資料錯誤!");


                    #endregion


                    #region //解析SpreadsheetData
                    var spreadsheetJson = JObject.Parse(PmdJsonData);
                    //取得jSON id
                    var pmdOrderIds = spreadsheetJson["spreadsheetInfo"]
                    .Select(x => (int)x["PMDOrderId"])
                    .ToList();

                    //取得該Type id
                    var parameters = new DynamicParameters();
                    parameters.Add("@TypeId", TypeId);
                    parameters.Add("@CustomerId", customerId);

                    var backendIds = sqlConnection.Query<int>(
                        @"SELECT PMDOrderTempId 
                          FROM SCM.PMDOrder
                          WHERE TypeId = @TypeId AND CustomerId = @CustomerId",
                        parameters
                    ).ToList();

                    //比對與刪除
                    var idsToDelete = backendIds.Except(pmdOrderIds).ToList();

                    if (idsToDelete.Any())
                    {
                        var deleteParams = new DynamicParameters();
                        deleteParams.Add("@Ids", idsToDelete);

                        var deleteSql = @"DELETE FROM SCM.PMDOrder WHERE PMDOrderTempId IN @Ids";
                        sqlConnection.Execute(deleteSql, deleteParams);
                    }

                    foreach (var pmdid in pmdOrderIds)
                    {
                        #region //驗證是否有該筆PMD TEMP資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.PMDOrderTempId, a.MtlItemId, a.SoQty, a.SpareQty, a.OrderDate, a.DrawingAttachmentDate, a.CurrentProcess, a.CustomerDueDate, a.ConfirmedDueDate
        , a.DelayRemark, a.MaterialNumber, a.Remark, a.TrackingNumber, a.ShipFrom, a.SemiDelivery, a.TypeId, a.PMDItem01, a.PMDItem02, a.PMDItem03, a.PMDItem04, a.PlannedShipmentDate
        , a.ActualShipmentDate, a.PlannedShipmentQty, a.ActualShipmentQty, a.MoldFrameCtrlNo, a.ReturnQty, a.ReturnIssue, a.PrevReturnQty, a.PrevShortageQty, a.TotalReplenishmentQty
        , a.CAV, a.ReturnDate, a.SoDetailId
        , a.CreateDate, a.LastModifiedDate, a.CreateBy, a.LastModifiedBy, b.MtlItemNo, a.MtlItemName, a.CustomerMtlItemNo, c.CustomerNo, c.CustomerName, c.CustomerId
        FROM SCM.PMDOrderTemp a 
        LEFT JOIN PDM.MtlItem b on a.MtlItemId = b.MtlItemId 
        INNER JOIN SCM.Customer c ON a.CustomerId = c.CustomerId
        WHERE 1=1 and c.CompanyId = @CompanyId and a.SendType = 0 AND PMDOrderTempId = @PMDOrderId";

                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("PMDOrderId", pmdid);

                        var pmdTempData = sqlConnection.Query(sql, dynamicParameters).FirstOrDefault();
                        int? soDetailId = pmdTempData.SoDetailId ?? -1;

                        #endregion

                        // 檢查是否有找到對應的 PMD TEMP 資料
                        if (pmdTempData != null)
                        {
                            // 找到對應的 PMD Order ID (透過 SoDetailId)


                                #region //查詢對應的 PMD Order ID
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT PMDOrderId FROM SCM.PMDOrder WHERE SoDetailId = @SoDetailId AND TypeId = @TypeId";
                                dynamicParameters.Add("SoDetailId", soDetailId.Value);
                                dynamicParameters.Add("TypeId", TypeId);

                                var pmdOrderId = sqlConnection.QuerySingleOrDefault<int?>(sql, dynamicParameters);
                                #endregion

                                if (pmdOrderId.HasValue)
                                {
                                    #region //UPDATE SCM.PMDOrder
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"
                                            UPDATE SCM.PMDOrder SET
                                                MtlItemId = @MtlItemId,
                                                SoQty = @SoQty,
                                                SpareQty = @SpareQty,
                                                OrderDate = @OrderDate,
                                                DrawingAttachmentDate = @DrawingAttachmentDate,
                                                CurrentProcess = @CurrentProcess,
                                                CustomerDueDate = @CustomerDueDate,
                                                ConfirmedDueDate = @ConfirmedDueDate,
                                                DelayRemark = @DelayRemark,
                                                MaterialNumber = @MaterialNumber,
                                                Remark = @Remark,
                                                TrackingNumber = @TrackingNumber,
                                                ShipFrom = @ShipFrom,
                                                SemiDelivery = @SemiDelivery,
                                                TypeId = @TypeId,
                                                PMDItem01 = @PMDItem01,
                                                PMDItem02 = @PMDItem02,
                                                PMDItem03 = @PMDItem03,
                                                PMDItem04 = @PMDItem04,
                                                MtlItemName = @MtlItemName,
                                                PlannedShipmentDate = @PlannedShipmentDate,
                                                ActualShipmentDate = @ActualShipmentDate,
                                                PlannedShipmentQty = @PlannedShipmentQty,
                                                ActualShipmentQty = @ActualShipmentQty,
                                                MoldFrameCtrlNo = @MoldFrameCtrlNo,
                                                ReturnQty = @ReturnQty,
                                                ReturnIssue = @ReturnIssue,
                                                PrevReturnQty = @PrevReturnQty,
                                                PrevShortageQty = @PrevShortageQty,
                                                TotalReplenishmentQty = @TotalReplenishmentQty,
                                                CAV = @CAV,
                                                ReturnDate = @ReturnDate,
                                                CustomerMtlItemNo = @CustomerMtlItemNo,
                                                CustomerId = @CustomerId,
                                                LastModifiedDate = @LastModifiedDate,
                                                LastModifiedBy = @LastModifiedBy
                                            WHERE PMDOrderId = @PMDOrderId";

                                    dynamicParameters.AddDynamicParams(new
                                    {
                                        // 使用 PMD TEMP 資料填入參數
                                        MtlItemId = pmdTempData.MtlItemId,
                                        SoQty = pmdTempData.SoQty ?? 0,
                                        SpareQty = pmdTempData.SpareQty ?? 0,
                                        OrderDate = pmdTempData.OrderDate,
                                        DrawingAttachmentDate = pmdTempData.DrawingAttachmentDate,
                                        CurrentProcess = pmdTempData.CurrentProcess,
                                        CustomerDueDate = pmdTempData.CustomerDueDate,
                                        ConfirmedDueDate = pmdTempData.ConfirmedDueDate,
                                        DelayRemark = pmdTempData.DelayRemark,
                                        MaterialNumber = pmdTempData.MaterialNumber,
                                        Remark = pmdTempData.Remark,
                                        TrackingNumber = pmdTempData.TrackingNumber,
                                        ShipFrom = pmdTempData.ShipFrom,
                                        SemiDelivery = pmdTempData.SemiDelivery,
                                        TypeId = pmdTempData.TypeId,
                                        PMDItem01 = pmdTempData.PMDItem01,
                                        PMDItem02 = pmdTempData.PMDItem02,
                                        PMDItem03 = pmdTempData.PMDItem03,
                                        PMDItem04 = pmdTempData.PMDItem04,
                                        MtlItemName = pmdTempData.MtlItemName,
                                        PlannedShipmentDate = pmdTempData.PlannedShipmentDate,
                                        ActualShipmentDate = pmdTempData.ActualShipmentDate,
                                        PlannedShipmentQty = pmdTempData.PlannedShipmentQty,
                                        ActualShipmentQty = pmdTempData.ActualShipmentQty,
                                        MoldFrameCtrlNo = pmdTempData.MoldFrameCtrlNo,
                                        ReturnQty = pmdTempData.ReturnQty,
                                        ReturnIssue = pmdTempData.ReturnIssue,
                                        PrevReturnQty = pmdTempData.PrevReturnQty,
                                        PrevShortageQty = pmdTempData.PrevShortageQty,
                                        TotalReplenishmentQty = pmdTempData.TotalReplenishmentQty,
                                        CAV = pmdTempData.CAV,
                                        ReturnDate = pmdTempData.ReturnDate,
                                        CustomerMtlItemNo = pmdTempData.CustomerMtlItemNo,
                                        CustomerId = pmdTempData.CustomerId,
                                        LastModifiedDate = LastModifiedDate, // 或使用你的 LastModifiedDate 變數
                                        LastModifiedBy = LastModifiedBy,  // 或使用你的 LastModifiedBy 變數
                                        PMDOrderId = pmdOrderId.Value
                                    });

                                    rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                    logPmdId = pmdOrderId.Value;

                                    // 可以在這裡加入成功更新的 log 或其他處理
                                    #endregion
                                }
                                else
                                {
                                #region //找不到對應的 PMD TEMP 資料，執行 INSERT
                                dynamicParameters = new DynamicParameters();
                                sql = @"
INSERT INTO SCM.PMDOrder (
    MtlItemId, SoQty, SpareQty, OrderDate, DrawingAttachmentDate,
    CurrentProcess, CustomerDueDate, ConfirmedDueDate, DelayRemark,
    MaterialNumber, Remark, TrackingNumber, ShipFrom, SemiDelivery, TypeId,
    PMDItem01, PMDItem02, PMDItem03, PMDItem04, PlannedShipmentDate,
    ActualShipmentDate, PlannedShipmentQty, ActualShipmentQty, SoDetailId, PMDOrderTempId,
    MoldFrameCtrlNo, ReturnQty, ReturnIssue, PrevReturnQty, PrevShortageQty,
    TotalReplenishmentQty, CAV, ReturnDate, CustomerMtlItemNo, MtlItemName,
    CustomerId, SendType, CreateBy, LastModifiedDate, LastModifiedBy
)
OUTPUT INSERTED.PMDOrderId
VALUES (
    @MtlItemId, @SoQty, @SpareQty, @OrderDate, @DrawingAttachmentDate,
    @CurrentProcess, @CustomerDueDate, @ConfirmedDueDate, @DelayRemark,
    @MaterialNumber, @Remark, @TrackingNumber, @ShipFrom, @SemiDelivery, @TypeId,
    @PMDItem01, @PMDItem02, @PMDItem03, @PMDItem04, @PlannedShipmentDate,
    @ActualShipmentDate, @PlannedShipmentQty, @ActualShipmentQty, @SoDetailId, @PMDOrderTempId,
    @MoldFrameCtrlNo, @ReturnQty, @ReturnIssue, @PrevReturnQty, @PrevShortageQty,
    @TotalReplenishmentQty, @CAV, @ReturnDate, @CustomerMtlItemNo, @MtlItemName,
    @CustomerId, @SendType, @CreateBy, @LastModifiedDate, @LastModifiedBy
)";

                                dynamicParameters.AddDynamicParams(new
                                {
                                    MtlItemId = pmdTempData.MtlItemId,
                                    SoQty = pmdTempData.SoQty ?? 0,
                                    SpareQty = pmdTempData.SpareQty ?? 0,
                                    OrderDate = pmdTempData.OrderDate,
                                    DrawingAttachmentDate = pmdTempData.DrawingAttachmentDate,
                                    CurrentProcess = pmdTempData.CurrentProcess,
                                    CustomerDueDate = pmdTempData.CustomerDueDate,
                                    ConfirmedDueDate = pmdTempData.ConfirmedDueDate,
                                    DelayRemark = pmdTempData.DelayRemark,
                                    MaterialNumber = pmdTempData.MaterialNumber,
                                    Remark = pmdTempData.Remark,
                                    TrackingNumber = pmdTempData.TrackingNumber,
                                    ShipFrom = pmdTempData.ShipFrom,
                                    SemiDelivery = pmdTempData.SemiDelivery,
                                    TypeId = TypeId,
                                    PMDItem01 = pmdTempData.PMDItem01,
                                    PMDItem02 = pmdTempData.PMDItem02,
                                    PMDItem03 = pmdTempData.PMDItem03,
                                    PMDItem04 = pmdTempData.PMDItem04,
                                    PlannedShipmentDate = pmdTempData.PlannedShipmentDate,
                                    ActualShipmentDate = pmdTempData.ActualShipmentDate,
                                    PlannedShipmentQty = pmdTempData.PlannedShipmentQty,
                                    ActualShipmentQty = pmdTempData.ActualShipmentQty,
                                    SoDetailId = pmdTempData.SoDetailId,
                                    PMDOrderTempId = pmdTempData.PMDOrderTempId,
                                    MoldFrameCtrlNo = pmdTempData.MoldFrameCtrlNo,
                                    ReturnQty = pmdTempData.ReturnQty,
                                    ReturnIssue = pmdTempData.ReturnIssue,
                                    PrevReturnQty = pmdTempData.PrevReturnQty,
                                    PrevShortageQty = pmdTempData.PrevShortageQty,
                                    TotalReplenishmentQty = pmdTempData.TotalReplenishmentQty,
                                    CAV = pmdTempData.CAV,
                                    ReturnDate = pmdTempData.ReturnDate,
                                    CustomerMtlItemNo = pmdTempData.CustomerMtlItemNo,
                                    MtlItemName = pmdTempData.MtlItemName,
                                    CustomerId = pmdTempData.CustomerId,
                                    SendType = 0,
                                    CreateBy = CreateBy,
                                    LastModifiedDate = LastModifiedDate,
                                    LastModifiedBy = LastModifiedBy
                                });

                                var insertResult = sqlConnection.Query<int>(sql, dynamicParameters);
                                int newPMDOrderId = insertResult.FirstOrDefault(); // 取得新插入的 PMDOrderId
                                logPmdId = newPMDOrderId;
                                rowsAffected = insertResult.Count();

                                // Console.WriteLine($"找不到對應的 PMD TEMP 資料，已新增 PMDOrder ID: {newPMDOrderId}");
                                #endregion

                            }

                        }
                        else
                        {
                            #region //找不到對應的 PMD TEMP 資料，執行 INSERT
                            dynamicParameters = new DynamicParameters();
                            sql = @"
INSERT INTO SCM.PMDOrder (
    MtlItemId, SoQty, SpareQty, OrderDate, DrawingAttachmentDate,
    CurrentProcess, CustomerDueDate, ConfirmedDueDate, DelayRemark,
    MaterialNumber, Remark, TrackingNumber, ShipFrom, SemiDelivery, TypeId,
    PMDItem01, PMDItem02, PMDItem03, PMDItem04, PlannedShipmentDate,
    ActualShipmentDate, PlannedShipmentQty, ActualShipmentQty, SoDetailId,
    MoldFrameCtrlNo, ReturnQty, ReturnIssue, PrevReturnQty, PrevShortageQty,
    TotalReplenishmentQty, CAV, ReturnDate, CustomerMtlItemNo, MtlItemName,
    CustomerId, SendType, CreateBy, LastModifiedDate, LastModifiedBy
)
OUTPUT INSERTED.PMDOrderId
VALUES (
    @MtlItemId, @SoQty, @SpareQty, @OrderDate, @DrawingAttachmentDate,
    @CurrentProcess, @CustomerDueDate, @ConfirmedDueDate, @DelayRemark,
    @MaterialNumber, @Remark, @TrackingNumber, @ShipFrom, @SemiDelivery, @TypeId,
    @PMDItem01, @PMDItem02, @PMDItem03, @PMDItem04, @PlannedShipmentDate,
    @ActualShipmentDate, @PlannedShipmentQty, @ActualShipmentQty, @SoDetailId,
    @MoldFrameCtrlNo, @ReturnQty, @ReturnIssue, @PrevReturnQty, @PrevShortageQty,
    @TotalReplenishmentQty, @CAV, @ReturnDate, @CustomerMtlItemNo, @MtlItemName,
    @CustomerId, @SendType, @CreateBy, @LastModifiedDate, @LastModifiedBy
)";

                            dynamicParameters.AddDynamicParams(new
                            {
                                MtlItemId = pmdTempData.MtlItemId,
                                SoQty = pmdTempData.SoQty ?? 0,
                                SpareQty = pmdTempData.SpareQty ?? 0,
                                OrderDate = pmdTempData.OrderDate,
                                DrawingAttachmentDate = pmdTempData.DrawingAttachmentDate,
                                CurrentProcess = pmdTempData.CurrentProcess,
                                CustomerDueDate = pmdTempData.CustomerDueDate,
                                ConfirmedDueDate = pmdTempData.ConfirmedDueDate,
                                DelayRemark = pmdTempData.DelayRemark,
                                MaterialNumber = pmdTempData.MaterialNumber,
                                Remark = pmdTempData.Remark,
                                TrackingNumber = pmdTempData.TrackingNumber,
                                ShipFrom = pmdTempData.ShipFrom,
                                SemiDelivery = pmdTempData.SemiDelivery,
                                TypeId = TypeId,
                                PMDItem01 = pmdTempData.PMDItem01,
                                PMDItem02 = pmdTempData.PMDItem02,
                                PMDItem03 = pmdTempData.PMDItem03,
                                PMDItem04 = pmdTempData.PMDItem04,
                                PlannedShipmentDate = pmdTempData.PlannedShipmentDate,
                                ActualShipmentDate = pmdTempData.ActualShipmentDate,
                                PlannedShipmentQty = pmdTempData.PlannedShipmentQty,
                                ActualShipmentQty = pmdTempData.ActualShipmentQty,
                                SoDetailId = pmdTempData.SoDetailId,
                                MoldFrameCtrlNo = pmdTempData.MoldFrameCtrlNo,
                                ReturnQty = pmdTempData.ReturnQty,
                                ReturnIssue = pmdTempData.ReturnIssue,
                                PrevReturnQty = pmdTempData.PrevReturnQty,
                                PrevShortageQty = pmdTempData.PrevShortageQty,
                                TotalReplenishmentQty = pmdTempData.TotalReplenishmentQty,
                                CAV = pmdTempData.CAV,
                                ReturnDate = pmdTempData.ReturnDate,
                                CustomerMtlItemNo = pmdTempData.CustomerMtlItemNo,
                                MtlItemName = pmdTempData.MtlItemName,
                                CustomerId = pmdTempData.CustomerId,
                                SendType = 0,
                                CreateBy = CreateBy,
                                LastModifiedDate = LastModifiedDate,
                                LastModifiedBy = LastModifiedBy
                            });

                            var insertResult = sqlConnection.Query<int>(sql, dynamicParameters);
                            int newPMDOrderId = insertResult.FirstOrDefault(); // 取得新插入的 PMDOrderId
                            logPmdId = newPMDOrderId;
                            rowsAffected = insertResult.Count();

                           // Console.WriteLine($"找不到對應的 PMD TEMP 資料，已新增 PMDOrder ID: {newPMDOrderId}");
                            #endregion
                        }

                        #region //INSERT SCM.PMDOrderLog
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.PMDOrderLog (UserId, EditTime, EditItem, EditTable
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.PMDOrderLogId
                                VALUES (@UserId, @EditTime, @EditItem, @EditTable
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";

                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                UserId = CreateBy,
                                EditTime = CreateDate,
                                EditItem = logPmdId,
                                EditTable = TypeId,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });

                        var loginsertResult = sqlConnection.Query(sql, dynamicParameters);
                        rowsAffected = loginsertResult.Count();
                        #endregion
                    }

                    #endregion


                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = $"共 {rowsAffected} 筆PMD資料上傳成功"
                    });
                }
            }
            catch (Exception ex)
            {
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = ex.Message
                });
                logger.Error(ex.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //UpdatePageLock -- 鎖定頁更新 -- GPAI 20250506
        public string UpdatePageLock(string PageNames, int isLock)
        {
            try
            {
                
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;

                        #region //判斷資料是否正確
                        sql = @"select top 1 * 
                                from SCM.PageLock
                                where PageName = @PageName
                                order by PageLockId desc";
                        dynamicParameters.Add("PageName", PageNames);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0)
                        {
                            #region //INSERT SCM.SoBpmLog
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO SCM.PageLock (
                                    PageName, IsLocked, LockedBy, LockTime,
                                    CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@PageName, @IsLocked, @LockedBy, @LockTime
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                              new
                              {
                                  PageName= PageNames,
                                  IsLocked = isLock,
                                  LockedBy = CreateBy,
                                  LockTime = CreateDate,
                                  CreateDate,
                                  LastModifiedDate,
                                  CreateBy,
                                  LastModifiedBy
                              });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected = insertResult.Count();
                            #endregion
                        }
                        else
                        {

                            foreach (var item in result)
                            {
                                //if (item.LockedBy == LastModifiedBy && item.IsLocked == 1)
                                //{
                                //    dynamicParameters = new DynamicParameters();
                                //    sql = @"UPDATE SCM.PageLock
                                //    SET IsLocked = @IsLocked,
                                //        LockedBy = @LockedBy,
                                //        LockTime = @LockTime,
                                //        LastModifiedDate = @LastModifiedDate,
                                //        LastModifiedBy = @LastModifiedBy
                                //    WHERE PageName = @PageName;";
                                //    dynamicParameters.AddDynamicParams(
                                //        new
                                //        {
                                //            IsLocked = 0,
                                //            LockedBy = LastModifiedBy,
                                //            LockTime = LastModifiedDate,
                                //            LastModifiedDate,
                                //            LastModifiedBy,
                                //            PageName = PageNames
                                //        });

                                //    rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                //}
                                //else {
                                //    dynamicParameters = new DynamicParameters();
                                //    sql = @"UPDATE SCM.PageLock
                                //    SET IsLocked = @IsLocked,
                                //        LockedBy = @LockedBy,
                                //        LockTime = @LockTime,
                                //        LastModifiedDate = @LastModifiedDate,
                                //        LastModifiedBy = @LastModifiedBy
                                //    WHERE PageName = @PageName;";
                                //    dynamicParameters.AddDynamicParams(
                                //        new
                                //        {
                                //            IsLocked = isLock,
                                //            LockedBy = LastModifiedBy,
                                //            LockTime = LastModifiedDate,
                                //            LastModifiedDate,
                                //            LastModifiedBy,
                                //            PageName = PageNames
                                //        });
                                //    rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                                //}

                            }


                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.PageLock
                                    SET IsLocked = @IsLocked,
                                        LockedBy = @LockedBy,
                                        LockTime = @LockTime,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                    WHERE PageName = @PageName;";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    IsLocked = isLock,
                                    LockedBy = LastModifiedBy,
                                    LockTime = LastModifiedDate,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    PageName = PageNames
                                });
                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

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

        #region //UpdatePageUnLock -- 鎖定頁解鎖 -- GPAI 20250506
        public string UpdatePageUnLock(string PageNames)
        {
            try
            {

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;

                        #region //判斷資料是否正確
                        sql = @"select top 1 * 
                                from SCM.PageLock
                                where PageName = @PageName
                                order by PageLockId desc";
                        dynamicParameters.Add("PageName", PageNames);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0)
                        {
                        }
                        else
                        {
                            foreach (var item in result) {
                                if (item.LockedBy == LastModifiedBy) {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE SCM.PageLock
                                    SET IsLocked = @IsLocked,
                                        LockedBy = @LockedBy,
                                        LockTime = @LockTime,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                    WHERE PageName = @PageName;";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            IsLocked = 0,
                                            LockedBy = LastModifiedBy,
                                            LockTime = LastModifiedDate,
                                            LastModifiedDate,
                                            LastModifiedBy,
                                            PageName = PageNames
                                        });

                                    rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                }
                                
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

        #region //UpdatePageSended -- PMD修改為已送 -- GPAI 20250506
        public string UpdatePageSended(string SendedIds, string pmdOrdertempIds)
        {
            try
            {

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;
                        var doId = SendedIds.Split(',');
                        var tempdoId = pmdOrdertempIds.Split(',');


                        #region //UPDATE
                        foreach (var item in doId)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.PMDOrder
                                    SET SendType = 1,
                                       
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                    WHERE PMDOrderId = @PMDOrderId;";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                   
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    PMDOrderId = item
                                });

                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                        }

                        foreach (var item in tempdoId)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.PMDOrderTemp
                                    SET SendType = 1,
                                       
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                    WHERE PMDOrderTempId = @PMDOrderId;";
                            dynamicParameters.AddDynamicParams(
                                new
                                {

                                    LastModifiedDate,
                                    LastModifiedBy,
                                    PMDOrderId = item
                                });

                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

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

        #region //UpdateOrderToPMDOrder 訂單更新至PMD進度表
        public string UpdateOrderToPMDOrder(int TypeId)
        {
            JObject jsonResponse;
            int rowsAffected = 0;
            string ErpConnectionStrings = "", ErpNo = "";

            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {

                    #region //判斷公司資料是否正確
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 CompanyId, CompanyName, ErpDb,ErpNo
                        FROM BAS.Company
                        WHERE CompanyId = @CurrentCompany";
                    dynamicParameters.Add("CurrentCompany", CurrentCompany);

                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() != 1) throw new SystemException("【公司】資料錯誤!");
                    #endregion

                    #region //取得所有PMD訂單的SoDetailId
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DISTINCT a.SoDetailId
                    FROM SCM.PMDOrderTemp a 
                    INNER JOIN SCM.Customer c ON a.CustomerId = c.CustomerId
                    WHERE c.CompanyId = @CompanyId and a.SendType = 0 and a.TypeId = @TypeId
                    AND a.SoDetailId IS NOT NULL";

                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    dynamicParameters.Add("TypeId", TypeId);

                    var existingPmdSoDetailIds = sqlConnection.Query<int>(sql, dynamicParameters).ToList();

                    // 除錯輸出
                    Console.WriteLine($"現有PMDOrderTemp中的SoDetailId數量: {existingPmdSoDetailIds.Count}");
                    Console.WriteLine($"現有PMDOrderTemp中的SoDetailId: [{string.Join(",", existingPmdSoDetailIds)}]");
                    #endregion

                    #region //取得所有未結案訂單
                    dynamicParameters = new DynamicParameters();
                    sql = @"select a.SoDetailId, a.SoQty, a.SiQty, a.SpareQty, FORMAT(a.PromiseDate, 'yyyy-MM-dd') PromiseDate
                    , b.SoErpPrefix,b.SoErpNo, a.SoSequence, FORMAT(b.SoDate, 'yyyy-MM-dd') SoDate
                    , c.CustomerNo, c.CustomerName, c.CustomerId
                    , d.MtlItemNo, d.MtlItemName, d.MtlItemId
                    from SCM.SoDetail a
                    inner join SCM.SaleOrder b on a.SoId = b.SoId
                    inner join SCM.Customer c on b.CustomerId = c.CustomerId
                    inner join PDM.MtlItem d on a.MtlItemId = d.MtlItemId
                    where a.ClosureStatus = 'N'  and a.SoQty > a.SiQty
                    and c.CompanyId = @CompanyId
                    order by c.CustomerId, a.SoDetailId desc";

                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    var orderresult = sqlConnection.Query(sql, dynamicParameters).ToList();

                    #endregion

                    #region //建立品號與最新ProcessAlias的對應字典
                    // 取得所有需要查詢的品號
                    var mtlItemIds = orderresult
                        .Select(x => (int)x.MtlItemId)
                        .Distinct()
                        .ToList();

                    var processAliasDict = new Dictionary<int, string>();

                    if (mtlItemIds.Any())
                    {
                        // 使用 IN 條件一次查詢所有品號的最新ProcessAlias
                        dynamicParameters = new DynamicParameters();
                        sql = @"
                            WITH LatestProcess AS (
                                SELECT 
                                    d.MtlItemId,
                                    b.ProcessAlias,
                                    ROW_NUMBER() OVER (PARTITION BY d.MtlItemId ORDER BY a.BarcodeProcessId DESC) as rn
                                FROM MES.BarcodeProcess a
                                INNER JOIN MES.MoProcess b ON a.MoProcessId = b.MoProcessId
                                INNER JOIN MES.ManufactureOrder c ON a.MoId = c.MoId
                                INNER JOIN MES.WipOrder d ON c.WoId = d.WoId
                                WHERE d.MtlItemId IN @MtlItemIds
                            )
                            SELECT MtlItemId, ProcessAlias
                            FROM LatestProcess
                            WHERE rn = 1";

                        dynamicParameters.Add("MtlItemIds", mtlItemIds);

                        var processResults = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var result2 in processResults)
                        {
                            int mtlItemId = (int)result2.MtlItemId;
                            string processAlias = result2.ProcessAlias;

                            processAliasDict[mtlItemId] = processAlias;
                            Console.WriteLine($"品號 {mtlItemId} 的最新ProcessAlias: {processAlias}");
                        }

                        Console.WriteLine($"共查詢到 {processAliasDict.Count} 個品號的ProcessAlias");
                    }
                    #endregion

                    #region //使用LINQ Except找出新的SoDetailId
                    // 取得所有訂單的SoDetailId
                    var allOrderSoDetailIds = orderresult
                        .Select(x => (int)x.SoDetailId)
                        .Distinct()
                        .ToList();


                    // 找出不存在於PMDOrderTemp的SoDetailId
                    var newSoDetailIds = allOrderSoDetailIds
                        .Except(existingPmdSoDetailIds)
                        .ToHashSet();

                    var newOrders = new List<dynamic>();

                    // 只處理新的SoDetailId
                    foreach (var order in orderresult)
                    {
                        int soDetailId = (int)order.SoDetailId;

                        if (newSoDetailIds.Contains(soDetailId))
                        {
                            Console.WriteLine($"準備插入 SoDetailId: {soDetailId}");

                            // 取得對應的CurrentProcess
                            int mtlItemId = (int)order.MtlItemId;
                            string currentProcess = processAliasDict.ContainsKey(mtlItemId)
                                ? processAliasDict[mtlItemId]
                                : null;

                            // 加入INSERT列表
                            newOrders.Add(new
                            {
                                SoDetailId = soDetailId,
                                MtlItemId = order.MtlItemId,
                                SoQty = order.SoQty ?? 0,
                                SpareQty = order.SpareQty ?? 0,
                                OrderDate = order.SoDate,
                                DrawingAttachmentDate = (string)null,
                                CurrentProcess = currentProcess, // 使用查詢到的ProcessAlias
                                CustomerDueDate = order.PromiseDate,
                                ConfirmedDueDate = (string)null,
                                DelayRemark = (string)null,
                                MaterialNumber = (string)null,
                                Remark = (string)null,
                                TrackingNumber = (string)null,
                                ShipFrom = (string)null,
                                SemiDelivery = (string)null,
                                TypeId = TypeId,
                                PMDItem01 = (string)null,
                                PMDItem02 = (string)null,
                                PMDItem03 = (string)null,
                                PMDItem04 = (string)null,
                                PlannedShipmentDate = (string)null,
                                ActualShipmentDate = (string)null,
                                PlannedShipmentQty = (int?)null,
                                ActualShipmentQty = (int?)null,
                                MoldFrameCtrlNo = (string)null,
                                ReturnQty = (int?)null,
                                ReturnIssue = (string)null,
                                PrevReturnQty = (int?)null,
                                PrevShortageQty = (int?)null,
                                TotalReplenishmentQty = (int?)null,
                                CAV = (int?)null,
                                ReturnDate = (string)null,
                                CustomerMtlItemNo = (string)null,
                                MtlItemName = order.MtlItemName,
                                CustomerId = order.CustomerId,
                                SendType = 0,
                                CreateDate = LastModifiedDate,
                                CreateBy = LastModifiedBy,
                                LastModifiedDate = LastModifiedDate,
                                LastModifiedBy = LastModifiedBy
                            });

                            // 從HashSet中移除，避免同一個SoDetailId被多次處理（如果有重複資料）
                            newSoDetailIds.Remove(soDetailId);
                        }
                    }

                    #endregion

                    #region //INSERT 新資料到 PMDOrderTemp
                    int insertedRows = 0;
                    if (newOrders.Any())
                    {
                        var insertSql = @"
INSERT INTO SCM.PMDOrderTemp (
    MtlItemId, SoQty, SpareQty, OrderDate, DrawingAttachmentDate,
    CurrentProcess, CustomerDueDate, ConfirmedDueDate, DelayRemark,
    MaterialNumber, Remark, TrackingNumber, ShipFrom, SemiDelivery, TypeId,
    PMDItem01, PMDItem02, PMDItem03, PMDItem04, PlannedShipmentDate,
    ActualShipmentDate, PlannedShipmentQty, ActualShipmentQty, SoDetailId,
    MoldFrameCtrlNo, ReturnQty, ReturnIssue, PrevReturnQty, PrevShortageQty,
    TotalReplenishmentQty, CAV, ReturnDate, CustomerMtlItemNo, MtlItemName,
    CustomerId, SendType, CreateBy, LastModifiedDate, LastModifiedBy
)
VALUES (
    @MtlItemId, @SoQty, @SpareQty, @OrderDate, @DrawingAttachmentDate,
    @CurrentProcess, @CustomerDueDate, @ConfirmedDueDate, @DelayRemark,
    @MaterialNumber, @Remark, @TrackingNumber, @ShipFrom, @SemiDelivery, @TypeId,
    @PMDItem01, @PMDItem02, @PMDItem03, @PMDItem04, @PlannedShipmentDate,
    @ActualShipmentDate, @PlannedShipmentQty, @ActualShipmentQty, @SoDetailId,
    @MoldFrameCtrlNo, @ReturnQty, @ReturnIssue, @PrevReturnQty, @PrevShortageQty,
    @TotalReplenishmentQty, @CAV, @ReturnDate, @CustomerMtlItemNo, @MtlItemName,
    @CustomerId, @SendType, @CreateBy, @LastModifiedDate, @LastModifiedBy
)";

                        insertedRows = sqlConnection.Execute(insertSql, newOrders);
                        Console.WriteLine($"新增了 {insertedRows} 筆PMD資料");
                    }
                    #endregion

                    #region //UPDATE 現有資料的 CurrentProcess
                    int updatedRows = 0;
                    if (processAliasDict.Any() && existingPmdSoDetailIds.Any())
                    {
                        // 準備更新現有資料的 CurrentProcess
                        var updateList = new List<dynamic>();

                        // 只處理存在於訂單資料中的SoDetailId
                        foreach (var order in orderresult)
                        {
                            int soDetailId = (int)order.SoDetailId;
                            int mtlItemId = (int)order.MtlItemId;

                            // 只有存在於PMDOrderTemp且有對應ProcessAlias的資料才更新
                            if (existingPmdSoDetailIds.Contains(soDetailId) &&
                                processAliasDict.ContainsKey(mtlItemId))
                            {
                                updateList.Add(new
                                {
                                    SoDetailId = soDetailId,
                                    CurrentProcess = processAliasDict[mtlItemId],
                                    LastModifiedDate = LastModifiedDate,
                                    LastModifiedBy = LastModifiedBy
                                });

                                Console.WriteLine($"準備更新 SoDetailId: {soDetailId}, ProcessAlias: {processAliasDict[mtlItemId]}");
                            }
                        }

                        if (updateList.Any())
                        {
                            var updateSql = @"
                                UPDATE SCM.PMDOrderTemp 
                                SET CurrentProcess = @CurrentProcess,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                WHERE SoDetailId = @SoDetailId 
                                AND SendType = 0 
                                AND TypeId = @TypeId";

                            // 為每個更新項目加上TypeId
                            var updateListWithTypeId = updateList.Select(item => new
                            {
                                SoDetailId = item.SoDetailId,
                                CurrentProcess = item.CurrentProcess,
                                LastModifiedDate = item.LastModifiedDate,
                                LastModifiedBy = item.LastModifiedBy,
                                TypeId = TypeId
                            }).ToList();

                            updatedRows = sqlConnection.Execute(updateSql, updateListWithTypeId);
                            Console.WriteLine($"更新了 {updatedRows} 筆PMD資料的CurrentProcess");
                        }
                    }
                    #endregion

                    rowsAffected = insertedRows + updatedRows;

                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = $"共處理 {rowsAffected} 筆PMD資料 (新增: {insertedRows}, 更新: {updatedRows})"
                    });
                }
            }
            catch (Exception ex)
            {
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = ex.Message
                });
                logger.Error(ex.Message);
                Console.WriteLine($"錯誤: {ex.Message}");
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //UploadPmdOrderDataTemp -- 上傳暫存PMD資料 -- GPAI 2025-05-06
        public string UploadPmdOrderDataTemp(string PmdJsonData, int TypeId, int customerId)
        {
            JObject jsonResponse;
            int rowsAffected = 0;
            string ErpConnectionStrings = "", ErpNo = "";

            try
            {


                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {

                    #region //判斷公司資料是否正確
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 CompanyId, CompanyName, ErpDb,ErpNo
                                FROM BAS.Company
                                WHERE CompanyId = @CurrentCompany";
                    dynamicParameters.Add("CurrentCompany", CurrentCompany);

                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() != 1) throw new SystemException("【公司】資料錯誤!");


                    #endregion


                    #region //解析SpreadsheetData
                    var spreadsheetJson = JObject.Parse(PmdJsonData);
                    //取得jSON id
                    var pmdOrderIds = spreadsheetJson["spreadsheetInfo"]
                    .Select(x => (int)x["PMDOrderId"])
                    .ToList();

                    //取得該Type id
                    var parameters = new DynamicParameters();
                    parameters.Add("@TypeId", TypeId);
                    parameters.Add("@CustomerId", customerId);

                    var backendIds = sqlConnection.Query<int>(
                        @"SELECT PMDOrderTempId 
                          FROM SCM.PMDOrderTemp
                          WHERE TypeId = @TypeId AND CustomerId = @CustomerId",
                        parameters
                    ).ToList();

                    //比對與刪除
                    var idsToDelete = backendIds.Except(pmdOrderIds).ToList();

                    if (idsToDelete.Any())
                    {
                        var deleteParams = new DynamicParameters();
                        deleteParams.Add("@Ids", idsToDelete);

                        var deleteSql = @"DELETE FROM SCM.PMDOrderTemp WHERE PMDOrderTempId IN @Ids";
                        sqlConnection.Execute(deleteSql, deleteParams);
                    }


                    foreach (var item in spreadsheetJson["spreadsheetInfo"])
                    {
                        var MtlItemId = -1;
                        var CustomerId = -1;

                        var PMDOrderId = item["PMDOrderId"]?.ToString();
                        var MtlItemNo = item["MtlItemNo"]?.ToString();
                        var MtlItemName = item["MtlItemName"]?.ToString();
                        var SoQty = item["SoQty"]?.ToObject<int?>();
                        var SpareQty = item["SpareQty"]?.ToObject<int?>();
                        var OrderDate = item["OrderDate"]?.ToString();
                        var DrawingAttachmentDate = item["DrawingAttachmentDate"]?.ToString();
                        var CurrentProcess = item["CurrentProcess"]?.ToString();
                        var CustomerDueDate = item["CustomerDueDate"]?.ToString();
                        var ConfirmedDueDate = item["ConfirmedDueDate"]?.ToString();
                        var DelayRemark = item["DelayRemark"]?.ToString();
                        var MaterialNumber = item["MaterialNumber"]?.ToString();
                        var Remark = item["Remark"]?.ToString();
                        var TrackingNumber = item["TrackingNumber"]?.ToString();
                        var ShipFrom = item["ShipFrom"]?.ToString();
                        var SemiDelivery = item["SemiDelivery"]?.ToString();
                        var PMDItem01 = item["PMDItem01"]?.ToString();
                        var PMDItem02 = item["PMDItem02"]?.ToString();
                        var PMDItem03 = item["PMDItem03"]?.ToString();
                        var PMDItem04 = item["PMDItem04"]?.ToString();
                        var PlannedShipmentDate = item["PlannedShipmentDate"]?.ToString();
                        var ActualShipmentDate = item["ActualShipmentDate"]?.ToString();
                        var PlannedShipmentQty = item["PlannedShipmentQty"]?.ToObject<int?>();
                        var ActualShipmentQty = item["ActualShipmentQty"]?.ToObject<int?>();
                        var MoldFrameCtrlNo = item["MoldFrameCtrlNo"]?.ToString();
                        var ReturnQty = item["ReturnQty"]?.ToObject<int?>();
                        var ReturnIssue = item["ReturnIssue"]?.ToString();
                        var PrevReturnQty = item["PrevReturnQty"]?.ToObject<int?>();
                        var PrevShortageQty = item["PrevShortageQty"]?.ToObject<int?>();
                        var TotalReplenishmentQty = item["TotalReplenishmentQty"]?.ToObject<int?>();
                        var CAV = item["CAV"]?.ToString();
                        var ReturnDate = item["ReturnDate"]?.ToString();
                        var CustomerMtlItemNo = item["CustomerMtlItemNo"]?.ToString();
                        var CustomerNo = item["CustomerNo"]?.ToString();
                        var CustomerName = item["CustomerName"]?.ToString();

                        //日期格式 yyyy/mm/dd  mm/dd
                        #region //驗證是否有該筆PMD資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT *
                                FROM SCM.PMDOrderTemp
                                WHERE PMDOrderTempId = @PMDOrderId";
                        dynamicParameters.Add("PMDOrderId", PMDOrderId);

                        var pmdresult = sqlConnection.Query(sql, dynamicParameters);
                        #endregion

                        #region //驗證品號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT *
                                FROM PDM.MtlItem
                                WHERE MtlItemNo = @MtlItemNo";
                        dynamicParameters.Add("MtlItemNo", MtlItemNo);

                        var mtiresult = sqlConnection.Query(sql, dynamicParameters);
                        if (TypeId == 1 || TypeId == 2 || TypeId == 3)
                        {

                            if (mtiresult.Count() <= 0) throw new SystemException("【品號】資料錯誤! 品號: " + MtlItemNo);
                            foreach (var item2 in mtiresult)
                            {
                                MtlItemId = item2.MtlItemId;
                            }
                        }

                        #endregion

                        #region //驗證客戶
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT *
                                FROM SCM.Customer
                                WHERE CustomerNo = @CustomerNo and CompanyId = @CompanyId";
                        dynamicParameters.Add("CustomerNo", CustomerNo);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var customresult = sqlConnection.Query(sql, dynamicParameters);

                        if (customresult.Count() <= 0) throw new SystemException("【客戶】資料錯誤! 客戶代號: " + CustomerNo);
                        foreach (var item3 in customresult)
                        {
                            CustomerId = item3.CustomerId;
                        }
                        #endregion

                        #region//若有PMD資料 UPDATE
                        if (pmdresult.Count() > 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"
    UPDATE SCM.PMDOrderTemp SET
        MtlItemId = @MtlItemId,
        SoQty = @SoQty,
        SpareQty = @SpareQty,
        OrderDate = @OrderDate,
        DrawingAttachmentDate = @DrawingAttachmentDate,
        CurrentProcess = @CurrentProcess,
        CustomerDueDate = @CustomerDueDate,
        ConfirmedDueDate = @ConfirmedDueDate,
        DelayRemark = @DelayRemark,
        MaterialNumber = @MaterialNumber,
        Remark = @Remark,
        TrackingNumber = @TrackingNumber,
        ShipFrom = @ShipFrom,
        SemiDelivery = @SemiDelivery,
        TypeId = @TypeId,
        PMDItem01 = @PMDItem01,
        PMDItem02 = @PMDItem02,
        PMDItem03 = @PMDItem03,
        PMDItem04 = @PMDItem04,
MtlItemName = @MtlItemName,
        PlannedShipmentDate = @PlannedShipmentDate,
        ActualShipmentDate = @ActualShipmentDate,
        PlannedShipmentQty = @PlannedShipmentQty,
        ActualShipmentQty = @ActualShipmentQty,
        MoldFrameCtrlNo = @MoldFrameCtrlNo,
        ReturnQty = @ReturnQty,
        ReturnIssue = @ReturnIssue,
        PrevReturnQty = @PrevReturnQty,
        PrevShortageQty = @PrevShortageQty,
        TotalReplenishmentQty = @TotalReplenishmentQty,
        CAV = @CAV,
        ReturnDate = @ReturnDate,
        CustomerMtlItemNo = @CustomerMtlItemNo,
        CustomerId = @CustomerId,

        LastModifiedDate = @LastModifiedDate,
        LastModifiedBy = @LastModifiedBy
    WHERE PMDOrderTempId = @PMDOrderId";

                            dynamicParameters.AddDynamicParams(new
                            {
                                MtlItemId,
                                SoQty,
                                SpareQty,
                                OrderDate,
                                DrawingAttachmentDate,
                                CurrentProcess,
                                CustomerDueDate,
                                ConfirmedDueDate,
                                DelayRemark,
                                MaterialNumber,
                                Remark,
                                TrackingNumber,
                                ShipFrom,
                                SemiDelivery,
                                TypeId,
                                PMDItem01,
                                PMDItem02,
                                PMDItem03,
                                PMDItem04,
                                MtlItemName,
                                PlannedShipmentDate,
                                ActualShipmentDate,
                                PlannedShipmentQty,
                                ActualShipmentQty,
                                MoldFrameCtrlNo,
                                ReturnQty,
                                ReturnIssue,
                                PrevReturnQty,
                                PrevShortageQty,
                                TotalReplenishmentQty,
                                CAV,
                                ReturnDate,
                                CustomerMtlItemNo,
                                CustomerId,
                                LastModifiedDate,
                                LastModifiedBy,
                                PMDOrderId
                            });

                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                            #endregion

                            #region//若無PMD資料 INSERT
                        }
                        else
                        {
                            dynamicParameters = new DynamicParameters();

                            sql = @"
INSERT INTO SCM.PMDOrderTemp (
    MtlItemId, SoQty, SpareQty, OrderDate, DrawingAttachmentDate,
    CurrentProcess, CustomerDueDate, ConfirmedDueDate, DelayRemark,
    MaterialNumber, Remark, TrackingNumber, ShipFrom, SemiDelivery, TypeId,
    PMDItem01, PMDItem02, PMDItem03, PMDItem04, PlannedShipmentDate,
    ActualShipmentDate, PlannedShipmentQty, ActualShipmentQty,
    MoldFrameCtrlNo, ReturnQty, ReturnIssue, PrevReturnQty, PrevShortageQty,
    TotalReplenishmentQty, CAV, ReturnDate, CustomerMtlItemNo, MtlItemName,
    CustomerId, SendType, CreateBy, LastModifiedDate, LastModifiedBy
)
VALUES (
    @MtlItemId, @SoQty, @SpareQty, @OrderDate, @DrawingAttachmentDate,
    @CurrentProcess, @CustomerDueDate, @ConfirmedDueDate, @DelayRemark,
    @MaterialNumber, @Remark, @TrackingNumber, @ShipFrom, @SemiDelivery, @TypeId,
    @PMDItem01, @PMDItem02, @PMDItem03, @PMDItem04, @PlannedShipmentDate,
    @ActualShipmentDate, @PlannedShipmentQty, @ActualShipmentQty,
    @MoldFrameCtrlNo, @ReturnQty, @ReturnIssue, @PrevReturnQty, @PrevShortageQty,
    @TotalReplenishmentQty, @CAV, @ReturnDate, @CustomerMtlItemNo, @MtlItemName,
    @CustomerId, @SendType, @CreateBy, @LastModifiedDate, @LastModifiedBy
)";

                            dynamicParameters.AddDynamicParams(new
                            {
                                MtlItemId,
                                SoQty,
                                SpareQty,
                                OrderDate,
                                DrawingAttachmentDate,
                                CurrentProcess,
                                CustomerDueDate,
                                ConfirmedDueDate,
                                DelayRemark,
                                MaterialNumber,
                                Remark,
                                TrackingNumber,
                                ShipFrom,
                                SemiDelivery,
                                TypeId,
                                PMDItem01,
                                PMDItem02,
                                PMDItem03,
                                PMDItem04,
                                PlannedShipmentDate,
                                ActualShipmentDate,
                                PlannedShipmentQty,
                                ActualShipmentQty,
                                MoldFrameCtrlNo,
                                ReturnQty,
                                ReturnIssue,
                                PrevReturnQty,
                                PrevShortageQty,
                                TotalReplenishmentQty,
                                CAV,
                                ReturnDate,
                                CustomerMtlItemNo,
                                MtlItemName,
                                CustomerId,
                                SendType = 0,
                                CreateBy,
                                LastModifiedDate,
                                LastModifiedBy
                            });


                            var insertResult = sqlConnection.Query(sql, dynamicParameters);
                            rowsAffected = insertResult.Count();

                        }
                        #endregion

                        #endregion

                    }



                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = $"共 {rowsAffected} 筆PMD資料上傳成功"
                    });
                }
            }
            catch (Exception ex)
            {
                jsonResponse = JObject.FromObject(new
                {
                    status = "error",
                    msg = ex.Message
                });
                logger.Error(ex.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion



        #endregion

        #region //Delete
        #region //DeleteSaleOrder -- 訂單資料刪除 -- Zoey 2022.07.15
        public string DeleteSaleOrder(int SoId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷訂單資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 ConfirmStatus, TransferStatus
                                FROM SCM.SaleOrder
                                WHERE SoId = @SoId
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("SoId", SoId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("訂單資料錯誤!");

                        #region //判斷確認過無法再確認
                        string confirmStatus = "", transferStatus = "";
                        foreach (var item in result)
                        {
                            confirmStatus = item.ConfirmStatus;
                            transferStatus = item.TransferStatus;
                        }

                        if (confirmStatus == "Y") throw new SystemException("訂單已經確認，無法刪除!");
                        if (transferStatus == "Y") throw new SystemException("訂單已經拋轉，無法刪除!");
                        #endregion
                        #endregion

                        #region //取得單身資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT SoDetailId
                                FROM SCM.SoDetail
                                WHERE SoId = @SoId";
                        dynamicParameters.Add("SoId", SoId);
                        var resultSoDetail = sqlConnection.Query(sql, dynamicParameters);
                        #endregion

                        #region //更新綁定的暫存檔
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.SoDetailTemp
                                WHERE SoId = @SoId";
                        dynamicParameters.Add("SoId", SoId);
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除關聯table
                        #region //訂單單身刪除
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.SoDetail
                                WHERE SoId = @SoId";
                        dynamicParameters.Add("SoId", SoId);
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.SaleOrder
                                WHERE SoId = @SoId";
                        dynamicParameters.Add("SoId", SoId);
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

        #region //DeleteSoDetail -- 訂單單身資料刪除 -- Zoey 2022.07.15
        public string DeleteSoDetail(int SoDetailId, int SoId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷訂單單身資料是否正確
                        sql = @"SELECT TOP 1 SoDetailId
                                FROM SCM.SoDetail
                                WHERE SoDetailId = @SoDetailId";
                        dynamicParameters.Add("SoDetailId", SoDetailId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("訂單單身資料錯誤!");
                        #endregion

                        #region //更新綁定的暫存檔
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.SoDetailTemp SET 
                                SoDetailId = NULL,
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SoDetailId = @SoDetailId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SoDetailId,
                                Status = "N", //解除訂單身綁定
                                LastModifiedDate,
                                LastModifiedBy
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.SoDetail
                                WHERE SoDetailId = @SoDetailId";
                        dynamicParameters.Add("SoDetailId", SoDetailId);

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

        #region //DeleteSoDetail02 -- 訂單單身資料刪除 -- Zoey 2022.07.15
        public string DeleteSoDetail02(int SoDetailId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    string companyNo = "", currency = "", taxNo = "", taxation = "";
                    int soId = -1, unitRound = 0, totalRound = 0;
                    double exciseTax = 0;

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //公司別資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var resultCompany = sqlConnection.Query(sql, dynamicParameters);
                        if (resultCompany.Count() <= 0) throw new SystemException("【公司別】資料錯誤!");

                        foreach (var item in resultCompany)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            companyNo = item.ErpNo;
                        }
                        #endregion

                        #region //判斷訂單單身資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 b.SoId, b.Currency, b.TaxNo, b.ConfirmStatus, b.TransferStatus
                                FROM SCM.SoDetail a
                                INNER JOIN SCM.SaleOrder b ON a.SoId = b.SoId
                                WHERE a.SoDetailId = @SoDetailId";
                        dynamicParameters.Add("SoDetailId", SoDetailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("訂單單身資料錯誤!");

                        #region //判斷確認過無法再刪除
                        string confirmStatus = "", transferStatus = "";
                        foreach (var item in result)
                        {
                            soId = Convert.ToInt32(item.SoId);
                            currency = item.Currency;
                            taxNo = item.TaxNo;
                            confirmStatus = item.ConfirmStatus;
                            transferStatus = item.TransferStatus;
                        }

                        if (confirmStatus == "Y") throw new SystemException("訂單已經確認，無法刪除!");
                        if (transferStatus == "Y") throw new SystemException("訂單已經拋轉，無法刪除!");
                        #endregion
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        #region //交易幣別設定
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT MF003, MF004, MF005, MF006
                                FROM CMSMF
                                WHERE MF001 = @Currency";
                        dynamicParameters.Add("Currency", currency);

                        var resultCurrencySetting = sqlConnection.Query(sql, dynamicParameters);
                        if (resultCurrencySetting.Count() <= 0) throw new SystemException("ERP交易幣別設定不存在!");

                        foreach (var item in resultCurrencySetting)
                        {
                            unitRound = Convert.ToInt32(item.MF003); //單價取位
                            totalRound = Convert.ToInt32(item.MF004); //金額取位
                        }
                        #endregion

                        #region //稅別碼設定
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT RTRIM(LTRIM(NN001)) NN001, RTRIM(LTRIM(NN002)) NN002, NN004, RTRIM(LTRIM(NN006)) NN006
                                FROM CMSNN
                                WHERE NN001 = @TaxNo";
                        dynamicParameters.Add("TaxNo", taxNo);

                        var resultTax = sqlConnection.Query(sql, dynamicParameters);
                        if (resultTax.Count() <= 0) throw new SystemException("ERP稅別碼設定不存在!");

                        foreach (var item in resultTax)
                        {
                            exciseTax = Convert.ToDouble(item.NN004); //營業稅率
                            taxation = item.NN006; //課稅別
                        }
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;

                        #region //更新綁定的暫存檔
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.SoDetailTemp SET 
                                SoDetailId = NULL,
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SoDetailId = @SoDetailId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SoDetailId,
                                Status = "N", //解除訂單身綁定
                                LastModifiedDate,
                                LastModifiedBy
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.SoDetail
                                WHERE SoDetailId = @SoDetailId";
                        dynamicParameters.Add("SoDetailId", SoDetailId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //單頭資料更新
                        double totalQty = 0, amount = 0, taxAmount = 0;

                        #region //所有單身資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT SoQty, SoPriceQty, UnitPrice, Amount
                                FROM SCM.SoDetail
                                WHERE SoId = @SoId
                                ORDER BY SoSequence";
                        dynamicParameters.Add("SoId", soId);

                        var resultSoDetail = sqlConnection.Query(sql, dynamicParameters);
                        //if (resultSoDetail.Count() <= 0) throw new SystemException("無單身資料!");

                        foreach (var item in resultSoDetail)
                        {
                            totalQty += Convert.ToDouble(item.SoQty);
                            amount += Math.Round(Math.Round(Convert.ToDouble(item.UnitPrice), unitRound, MidpointRounding.AwayFromZero) * Convert.ToDouble(item.SoPriceQty), totalRound, MidpointRounding.AwayFromZero);
                        }

                        #region //計算數量與金額
                        switch (taxation)
                        {
                            case "1":
                                taxAmount = amount - Math.Round(amount / (1 + exciseTax), totalRound, MidpointRounding.AwayFromZero);
                                amount = Math.Round(amount / (1 + exciseTax), totalRound, MidpointRounding.AwayFromZero);
                                break;
                            case "2":
                                taxAmount = Math.Round(amount * exciseTax, totalRound, MidpointRounding.AwayFromZero);
                                break;
                            case "3":
                                break;
                            case "4":
                                break;
                            case "9":
                                break;
                        }
                        #endregion
                        #endregion

                        #region //單頭更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.SaleOrder SET
                                TotalQty = @TotalQty,
                                Amount = @Amount,
                                TaxAmount = @TaxAmount,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SoId = @SoId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TotalQty = totalQty,
                                Amount = amount,
                                TaxAmount = taxAmount,
                                LastModifiedDate,
                                LastModifiedBy,
                                SoId = soId
                            });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion
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

        #region //DeleteSoModification -- 訂單變更單頭資料刪除 -- Zoey 2022.08.17
        public string DeleteSoModification(int SomId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷訂單變更單頭資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.SoModification
                                WHERE SomId = @SomId";
                        dynamicParameters.Add("SomId", SomId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("訂單變更單頭資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除連動table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.SomDetail
                                WHERE SomId = @SomId";
                        dynamicParameters.Add("SomId", SomId);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.SoModification
                                WHERE SomId = @SomId";
                        dynamicParameters.Add("SomId", SomId);

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

        #region //DeleteSoDetailTemp -- 採購單匯入資料刪除 -- Chia Yuan 2024.04.24
        public string DeleteSoDetailTemp(int SoId, int CompanyId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷採購單匯入資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.SoDetailTemp
                                WHERE SoId = @SoId";
                        dynamicParameters.Add("SoId", SoId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【採購單】匯入錯誤!");
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.SoDetailTemp
                                WHERE SoId = @SoId";
                        dynamicParameters.Add("SoId", SoId);
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

        #region //API
        #region //GetCreditLimit -- 檢核客戶信用額度 -- Ann 2023-12-22
        public string GetCreditLimit(string CustomerNo, SqlConnection sqlConnection, decimal totalAmount, string docType)
        {
            try
            {
                decimal notesReceivableRate = 0; //未兌現應收票據比率
                decimal accountsReceivableRate = 0; //應收帳款比率
                decimal unbilledSalesAmountRate = 0; //未結帳銷貨金額比率
                decimal orderAmountRate = 0; //未出貨訂單金額比率
                decimal temporaryDisbursementAmountRate = 0; //未歸還暫出金額比率

                decimal creditBalance = 0; //信用餘額
                decimal excessCreditLimit = 0; //可超出額度
                decimal totalAccountsReceivable = 0; //應收合計

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success"
                });
                #endregion

                #region //確認此客戶的信用額度管制方式 COPMA
                dynamicParameters = new DynamicParameters();
                sql = @"SELECT LTRIM(RTRIM(MA032)) MA032, MA033, MA034
                        , LTRIM(RTRIM(MA088)) MA088, LTRIM(RTRIM(MA089)) MA089
                        , MA091, MA092, MA093, MA094, MA102
                        , LTRIM(RTRIM(MA120)) MA120, LTRIM(RTRIM(MA132)) MA132
                        FROM COPMA
                        WHERE MA001 = @CustomerNo";
                dynamicParameters.Add("CustomerNo", CustomerNo);

                var COPMAResult = sqlConnection.Query(sql, dynamicParameters);

                if (COPMAResult.Count() <= 0) throw new SystemException("客戶【" + CustomerNo + "】資料錯誤!!");

                string creditLimitControl = "";
                decimal creaditLimit = 0;
                decimal creditLimitTolerance = 0;
                string MA088 = "";
                string MA089 = "";
                string MA120 = "";
                string MA132 = "";
                decimal MA091 = 0;
                decimal MA092 = 0;
                decimal MA093 = 0;
                decimal MA094 = 0;
                decimal MA102 = 0;
                foreach (var item in COPMAResult)
                {
                    creditLimitControl = item.MA032;
                    creaditLimit = item.MA033;
                    creditLimitTolerance = item.MA034;
                    MA088 = item.MA088;
                    MA089 = item.MA089;
                    MA120 = item.MA120;
                    MA132 = item.MA132;
                    MA091 = item.MA091;
                    MA092 = item.MA091;
                    MA093 = item.MA093;
                    MA094 = item.MA094;
                    MA102 = item.MA102;
                }
                #endregion

                #region //針對信用額度管控方式進行不同流程
                if (creditLimitControl == "Y")
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT LTRIM(RTRIM(MH001)) MH001, LTRIM(RTRIM(MH002)) MH002
                            , MH004, MH005, MH006, MH007, MH008
                            , LTRIM(RTRIM(MH014)) MH014, LTRIM(RTRIM(MH015)) MH015
                            FROM COPMH
                            WHERE 1=1";

                    var COPMHResult = sqlConnection.Query(sql, dynamicParameters);

                    if (COPMHResult.Count() <= 0) throw new SystemException("查詢信用控制參數設定資料錯誤!!");

                    foreach (var item in COPMHResult)
                    {
                        #region //先確認檢核卡控參數
                        switch (docType)
                        {
                            case "SaleOrder":
                                if (item.MH001 != "3")
                                {
                                    return jsonResponse.ToString();
                                }
                                break;
                            case "ShippingOrder":
                                if (item.MH002 != "3")
                                {
                                    return jsonResponse.ToString();
                                }
                                break;
                            case "ShippingNotice":
                                if (item.MH014 != "3")
                                {
                                    return jsonResponse.ToString();
                                }
                                break;
                            case "TempShippingNote":
                                if (item.MH015 != "3")
                                {
                                    return jsonResponse.ToString();
                                }
                                break;
                            default:
                                throw new SystemException("單據名稱【" + docType + "】尚未設定邏輯!!");
                        }
                        #endregion

                        notesReceivableRate = item.MH004;
                        accountsReceivableRate = item.MH005;
                        unbilledSalesAmountRate = item.MH006;
                        orderAmountRate = item.MH007;
                        temporaryDisbursementAmountRate = item.MH008;
                    }
                }
                else if (creditLimitControl == "y")
                {
                    #region //先確認檢核卡控參數
                    switch (docType)
                    {
                        case "SaleOrder":
                            if (MA088 != "3")
                            {
                                return jsonResponse.ToString();
                            }
                            break;
                        case "ShippingOrder":
                            if (MA089 != "3")
                            {
                                return jsonResponse.ToString();
                            }
                            break;
                        case "ShippingNotice":
                            if (MA120 != "3")
                            {
                                return jsonResponse.ToString();
                            }
                            break;
                        case "TempShippingNote":
                            if (MA132 != "3")
                            {
                                return jsonResponse.ToString();
                            }
                            break;
                        default:
                            throw new SystemException("單據名稱【" + docType + "】尚未設定邏輯!!");
                    }
                    #endregion

                    notesReceivableRate = MA091;
                    accountsReceivableRate = MA092;
                    unbilledSalesAmountRate = MA093;
                    orderAmountRate = MA094;
                    temporaryDisbursementAmountRate = MA102;
                }
                else
                {
                    return jsonResponse.ToString();
                }
                #endregion

                #region //計算可超出額度
                //可超出額度 = 信用額度(COPMA.MA033) * (1 + 可超出率%(COPMA.MA034))
                excessCreditLimit = creaditLimit * (1 + creditLimitTolerance);
                #endregion

                #region //計算應收合計
                //【應收合計】=【應收票據】+【應收帳款】+【未結帳銷貨】+【訂貨金額】+【暫出金額】

                #region //計算應收票據 
                // * 未兌現應收票據比率 (COPMA. MA032 = Y)COPMH.MH004 (COPMA. MA032 = y)COPMA. MA091
                decimal notesReceivable = 0;
                dynamicParameters = new DynamicParameters();
                sql = @"SELECT TC012 AS status,TC013 AS customer,SUM(ISNULL(TC003,0)) AS Amount
                        FROM NOTTC 
                        WHERE TC012 IN ('1','2','3','5','9') 
                        AND TC013 = @CustomerNo
                        GROUP BY TC012,TC013";
                dynamicParameters.Add("CustomerNo", CustomerNo);

                var NOTTCResult = sqlConnection.Query(sql, dynamicParameters);

                foreach (var item in NOTTCResult)
                {
                    notesReceivable = item.Amount;
                }
                notesReceivable = notesReceivable * notesReceivableRate;
                #endregion

                #region //計算應收帳款
                //【應收帳款】= (結帳單-溢收待抵單) * 應收帳款比率 (COPMA. MA032 = Y)COPMH.MH005 (COPMA. MA032 = y)COPMA. MA092

                decimal accountsReceivable = 0;

                #region //結帳單統計(原幣)
                decimal invoiceAmount = 0;
                dynamicParameters = new DynamicParameters();
                sql = @"SELECT TA004 AS customer,SUM(ISNULL(a.ARtotal,0)) AS SumARtotal
                        FROM ACRTA a1
                        INNER JOIN (
                        SELECT TA001,TA002,TA029 AS AR,TA030 AS tax,TA031 AS ARreceived,(ISNULL(a.TA029,0) + ISNULL(a.TA030,0) - ISNULL(a.TA031,0)) AS ARtotal,TA004 AS customer
                        FROM ACRTA a
                        WHERE TA025 = 'Y' AND TA027 = 'N' AND TA019 = 'N' AND a.TA001 LIKE '61%')a ON a1.TA001 = a.TA001 AND a1.TA002 = a.TA002
                        WHERE a1.TA004 = @CustomerNo
                        GROUP BY TA004";
                dynamicParameters.Add("CustomerNo", CustomerNo);

                var ACRTAResult = sqlConnection.Query(sql, dynamicParameters);

                foreach (var item in ACRTAResult)
                {
                    invoiceAmount = item.SumARtotal;
                }
                #endregion

                #region //溢收待抵單統計(原幣)
                decimal overpaymentToBeOffset = 0;
                dynamicParameters = new DynamicParameters();
                sql = @"SELECT TA004 AS customer,SUM(ISNULL(a.ARtotal,0)) AS SumARtotal
                        FROM ACRTA a1
                        INNER JOIN (
                        SELECT TA001,TA002,TA029 AS AR,TA030 AS tax,TA031 AS ARreceived,(ISNULL(a.TA029,0) + ISNULL(a.TA030,0) - ISNULL(a.TA031,0)) AS ARtotal,TA004 AS customer
                        FROM ACRTA a
                        WHERE TA025 = 'Y' AND TA027 = 'N' AND TA019 = 'N' AND a.TA001 LIKE '62%')a ON a1.TA001 = a.TA001 AND a1.TA002 = a.TA002
                        WHERE a1.TA004 = @CustomerNo
                        GROUP BY TA004";
                dynamicParameters.Add("CustomerNo", CustomerNo);

                var ACRTAResult2 = sqlConnection.Query(sql, dynamicParameters);

                foreach (var item in ACRTAResult2)
                {
                    overpaymentToBeOffset = item.SumARtotal;
                }
                #endregion

                accountsReceivable = (invoiceAmount - overpaymentToBeOffset) * accountsReceivableRate;
                #endregion

                #region //未結帳銷貨
                //【未結帳銷貨】= (銷貨單 - 銷退單) * 未結帳銷貨金額比率(COPMA. MA032 = Y)COPMH.MH006 (COPMA. MA032 = y)COPMA. MA093

                decimal unbilledSalesAmount = 0;

                #region //銷貨單統計(原幣)
                decimal salesAmount = 0;
                dynamicParameters = new DynamicParameters();
                sql = @"SELECT TG004 AS customer,SUM(ISNULL(salesTotal,0)) AS SumsalesTotal
                        FROM COPTG a1
                        INNER JOIN(SELECT TG001,TG002,TG013 AS 原幣銷貨金額sales,TG025 AS 原幣銷貨稅額salesTax, (ISNULL(TG013,0) + ISNULL(TG025,0)) AS salesTotal 
                        FROM COPTG a
                        LEFT JOIN COPTH b ON a.TG001 = b.TH001 AND a.TG002 = b.TH002 
                        WHERE TG023 = 'Y' AND TH026 = 'N'
                        GROUP BY TG001,TG002,TG013,TG025
                        )c ON c.TG001 = a1.TG001 AND c.TG002 = a1.TG002
                        WHERE TG004 = @CustomerNo
                        GROUP BY TG004";
                dynamicParameters.Add("CustomerNo", CustomerNo);

                var COPTGResult = sqlConnection.Query(sql, dynamicParameters);

                foreach (var item in COPTGResult)
                {
                    salesAmount = item.SumsalesTotal;
                }
                #endregion

                #region //銷退單統計(原幣)
                decimal returnAmount = 0;
                dynamicParameters = new DynamicParameters();
                sql = @"SELECT TI004 AS customer,SUM(salesReturnTotal) AS SumsalesReturnTotal
                        FROM COPTI a1 
                        INNER JOIN (SELECT TI001,TI002,TI004 AS customer,TI010 AS 原幣銷退金額salesReturn,TI011 AS 原幣銷退稅額salesReturnTax,(ISNULL(TI010,0) + ISNULL(TI011,0)) AS salesReturnTotal
                        FROM COPTI a 
                        LEFT JOIN COPTJ b ON a.TI001 = b.TJ001 AND a.TI002 = b.TJ002
                        WHERE TI019 = 'Y' AND TJ024 = 'N'
                        GROUP BY TI004,TI010,TI011,TI001,TI002) c ON c.TI001 = a1.TI001 AND c.TI002 = a1.TI002
                        WHERE TI004 = @CustomerNo
                        GROUP BY TI004";
                dynamicParameters.Add("CustomerNo", CustomerNo);

                var COPTIResult = sqlConnection.Query(sql, dynamicParameters);

                foreach (var item in COPTIResult)
                {
                    returnAmount = item.SumsalesReturnTotal;
                }
                #endregion

                unbilledSalesAmount = (salesAmount - returnAmount) * unbilledSalesAmountRate;

                #endregion

                #region //訂貨金額統計(原幣)
                // * 未出貨訂單金額比率(COPMA. MA032 = Y)COPMH.MH007 (COPMA. MA032 = y)COPMA. MA094
                decimal orderAmount = 0;
                dynamicParameters = new DynamicParameters();
                sql = @"SELECT c.TC004 AS customer,SUM(c.unsoldOrder) AS SumunsoldOrder
                        FROM COPTC a1
                        INNER JOIN (
                            SELECT a.TC001,a.TC002,b.TD003,a.TC004,b.TD008 訂單數量,b.TD009 已交數量,b.TD011 單價,(ISNULL(b.TD008,0) - ISNULL(b.TD009,0))AS unsoldQTY, 
                            ((ISNULL(b.TD008,0) - ISNULL(b.TD009,0)) * ISNULL(b.TD011,0))AS unsoldOrder
                        FROM COPTC a
                        LEFT JOIN COPTD b ON a.TC001 = b.TD001 AND a.TC002 = b.TD002
                        WHERE a.TC027 = 'Y' AND b.TD016 = 'N'   AND TD012 <> 0 
                        ) c ON c.TC001 = a1.TC001 AND c.TC002 = a1.TC002
                        WHERE  a1.TC004 = @CustomerNo
                        GROUP BY c.TC004";
                dynamicParameters.Add("CustomerNo", CustomerNo);

                var COPTCResult = sqlConnection.Query(sql, dynamicParameters);

                foreach (var item in COPTCResult)
                {
                    orderAmount = item.SumunsoldOrder;
                }

                orderAmount = orderAmount * orderAmountRate;
                #endregion

                #region //暫出金額統計(原幣)
                // * 未歸還暫出金額比率(COPMA. MA032 = Y)COPMH.MH008 (COPMA. MA032 = y)COPMA. MA102
                decimal temporaryDisbursementAmount = 0;
                dynamicParameters = new DynamicParameters();
                sql = @"SELECT TF005 AS customer,SUM(c.unreturned) AS SumUnreturned
                        FROM INVTF a1
                        INNER JOIN (
                            SELECT TG001,TG002,TG003,TG009 數量,TG012 單價,TG020 轉進銷量,TG021 歸還量,((ISNULL(TG009,0)-ISNULL(TG021,0))) AS unreturnedQTY,
                        ((ISNULL(TG009,0)-ISNULL(TG021,0)) * ISNULL(TG012,0)) AS unreturned
                        FROM INVTF a
                        LEFT JOIN INVTG b ON a.TF001 = b.TG001 AND a.TF002 = b.TG002
                        WHERE  TF020 = 'Y' AND TG024 = 'N'
                        )c ON a1.TF001 = c.TG001 AND a1.TF002 = c.TG002
                        WHERE a1.TF005 = @CustomerNo
                        GROUP BY TF005";
                dynamicParameters.Add("CustomerNo", CustomerNo);

                var INVTFResult = sqlConnection.Query(sql, dynamicParameters);

                foreach (var item in INVTFResult)
                {
                    //temporaryDisbursementAmount = item.SumUnreturned; //先預設是0，待君葦修改邏輯後修正
                }

                //temporaryDisbursementAmount = temporaryDisbursementAmount * temporaryDisbursementAmountRate;
                #endregion

                totalAccountsReceivable = notesReceivable + accountsReceivable + unbilledSalesAmount + orderAmount + temporaryDisbursementAmount;
                #endregion

                #region //計算最終信用餘額
                //信用餘額 = 可超出額度 - 應收合計
                creditBalance = excessCreditLimit - totalAccountsReceivable;
                #endregion

                if (totalAmount > creditBalance)
                {
                    throw new SystemException("訂單合計金額【" + totalAmount.ToString() + "】已超過剩餘信用額度【" + creditBalance.ToString() + "】");
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion
        
        #region //CheckCreditLimit -- 檢核客戶信用額度 -- Ann 2023-12-22
        public string CheckCreditLimit(string CustomerNo, string Currency, decimal TotalAmount, string DocType, decimal Amount, string CompanyNo)
        {
            try
            {
                decimal notesReceivableRate = 0; //未兌現應收票據比率
                decimal accountsReceivableRate = 0; //應收帳款比率
                decimal unbilledSalesAmountRate = 0; //未結帳銷貨金額比率
                decimal orderAmountRate = 0; //未出貨訂單金額比率
                decimal temporaryDisbursementAmountRate = 0; //未歸還暫出金額比率

                decimal creditBalance = 0; //信用餘額
                decimal excessCreditLimit = 0; //可超出額度
                decimal totalAccountsReceivable = 0; //應收合計

                decimal allTotalAmount = TotalAmount + Amount; //總金額(含稅額)+單筆單身金額(含稅額)

                jsonResponse = JObject.FromObject(new
                {
                    status = "ok",
                    msg = "滿足信用餘額。"
                });

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    if (CompanyNo.Length > 0)
                    {
                        #region //確認公司別DB
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        dynamicParameters.Add("CompanyNo", CompanyNo);

                        var companyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (companyResult.Count() <= 0) throw new SystemException("公司別錯誤!");

                        foreach (var item in companyResult)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        }
                        #endregion
                    }
                    else
                    {
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
                }

                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                {
                    #region //確認此客戶的信用額度管制方式 COPMA
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT LTRIM(RTRIM(MA032)) MA032, MA033, MA034
                            , LTRIM(RTRIM(MA088)) MA088, LTRIM(RTRIM(MA089)) MA089
                            , MA091, MA092, MA093, MA094, MA102
                            , LTRIM(RTRIM(MA118)) MA118
                            , LTRIM(RTRIM(MA120)) MA120, LTRIM(RTRIM(MA132)) MA132
                            FROM COPMA
                            WHERE MA001 = @CustomerNo";
                    dynamicParameters.Add("CustomerNo", CustomerNo);

                    var COPMAResult = sqlConnection.Query(sql, dynamicParameters);

                    if (COPMAResult.Count() <= 0) throw new SystemException("客戶【" + CustomerNo + "】資料錯誤!!");

                    string creditLimitControl = "";
                    decimal creaditLimit = 0;
                    decimal creditLimitTolerance = 0;
                    string MA088 = "";
                    string MA089 = "";
                    string MA118 = "";
                    string MA120 = "";
                    string MA132 = "";
                    decimal MA091 = 0;
                    decimal MA092 = 0;
                    decimal MA093 = 0;
                    decimal MA094 = 0;
                    decimal MA102 = 0;
                    foreach (var item in COPMAResult)
                    {
                        creditLimitControl = item.MA032;
                        creaditLimit = item.MA033;
                        creditLimitTolerance = item.MA034;
                        MA088 = item.MA088;
                        MA089 = item.MA089;
                        MA118 = item.MA118;
                        MA120 = item.MA120;
                        MA132 = item.MA132;
                        MA091 = item.MA091;
                        MA092 = item.MA092;
                        MA093 = item.MA093;
                        MA094 = item.MA094;
                        MA102 = item.MA102;
                    }
                    #endregion

                    #region //針對信用額度管控方式進行不同流程
                    if (creditLimitControl == "Y")
                    {
                        //Y情況找COPMH設定參數
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT LTRIM(RTRIM(MH001)) MH001, LTRIM(RTRIM(MH002)) MH002
                                , MH004, MH005, MH006, MH007, MH008
                                , LTRIM(RTRIM(MH014)) MH014, LTRIM(RTRIM(MH015)) MH015
                                FROM COPMH
                                WHERE 1=1";

                        var COPMHResult = sqlConnection.Query(sql, dynamicParameters);

                        if (COPMHResult.Count() <= 0) throw new SystemException("查詢信用控制參數設定資料錯誤!!");

                        foreach (var item in COPMHResult)
                        {
                            notesReceivableRate = item.MH004;
                            accountsReceivableRate = item.MH005;
                            unbilledSalesAmountRate = item.MH006;
                            orderAmountRate = item.MH007;
                            temporaryDisbursementAmountRate = item.MH008;

                            #region //先確認檢核卡控參數
                            switch (DocType)
                            {
                                #region //訂單SaleOrder
                                case "SaleOrder":
                                    if (item.MH001 == "1")
                                    {
                                        return jsonResponse.ToString();
                                    }
                                    else if (item.MH001 == "2")
                                    {
                                        decimal resultCredit = GetCreditBalance();
                                        if (allTotalAmount > resultCredit)
                                        {
                                            jsonResponse = JObject.FromObject(new
                                            {
                                                status = "warning",
                                                msg = "訂單總金額【" + allTotalAmount + "】已超過信用餘額【" + resultCredit + "】，請確認是否繼續執行!!"
                                            });
                                        }

                                        return jsonResponse.ToString();
                                    }
                                    if (item.MH001 == "3")
                                    {
                                        decimal resultCredit = GetCreditBalance();
                                        if (allTotalAmount > resultCredit)
                                        {
                                            throw new SystemException("訂單總金額【" + allTotalAmount + "】已超過信用餘額【" + resultCredit + "】，無法新增!!");
                                        }

                                        return jsonResponse.ToString();
                                    }
                                    break;
                                #endregion
                                #region //銷貨單ShippingOrder
                                case "ShippingOrder":
                                    if (item.MH002 == "1")
                                    {
                                        return jsonResponse.ToString();
                                    }
                                    else if (item.MH002 == "2")
                                    {
                                        decimal resultCredit = GetCreditBalance();
                                        if (allTotalAmount > resultCredit)
                                        {
                                            jsonResponse = JObject.FromObject(new
                                            {
                                                status = "warning",
                                                msg = "銷貨單總金額【" + allTotalAmount + "】已超過信用餘額【" + resultCredit + "】，請確認是否繼續執行!!"
                                            });
                                        }

                                        return jsonResponse.ToString();
                                    }
                                    if (item.MH002 == "3")
                                    {
                                        decimal resultCredit = GetCreditBalance();
                                        if (allTotalAmount > resultCredit)
                                        {
                                            throw new SystemException("銷貨單總金額【" + allTotalAmount + "】已超過信用餘額【" + resultCredit + "】，無法新增!!");
                                        }

                                        return jsonResponse.ToString();
                                    }
                                    break;
                                #endregion
                                #region //出貨通知單ShippingNotice
                                case "ShippingNotice":
                                    if (item.MH014 == "1")
                                    {
                                        return jsonResponse.ToString();
                                    }
                                    else if (item.MH014 == "2")
                                    {
                                        decimal resultCredit = GetCreditBalance();
                                        if (allTotalAmount > resultCredit)
                                        {
                                            jsonResponse = JObject.FromObject(new
                                            {
                                                status = "warning",
                                                msg = "出貨通知單總金額【" + allTotalAmount + "】已超過信用餘額【" + resultCredit + "】，請確認是否繼續執行!!"
                                            });
                                        }

                                        return jsonResponse.ToString();
                                    }
                                    if (item.MH014 == "3")
                                    {
                                        decimal resultCredit = GetCreditBalance();
                                        if (allTotalAmount > resultCredit)
                                        {
                                            throw new SystemException("出貨通知單總金額【" + allTotalAmount + "】已超過信用餘額【" + resultCredit + "】，無法新增!!");
                                        }

                                        return jsonResponse.ToString();
                                    }
                                    break;
                                #endregion
                                #region //暫出單TempShippingNote
                                case "TempShippingNote":
                                    if (item.MH015 == "1")
                                    {
                                        return jsonResponse.ToString();
                                    }
                                    else if (item.MH015 == "2")
                                    {
                                        decimal resultCredit = GetCreditBalance();
                                        if (allTotalAmount > resultCredit)
                                        {
                                            jsonResponse = JObject.FromObject(new
                                            {
                                                status = "warning",
                                                msg = "暫出單總金額【" + allTotalAmount + "】已超過信用餘額【" + resultCredit + "】，請確認是否繼續執行!!"
                                            });
                                        }

                                        return jsonResponse.ToString();
                                    }
                                    if (item.MH015 == "3")
                                    {
                                        decimal resultCredit = GetCreditBalance();
                                        if (allTotalAmount > resultCredit)
                                        {
                                            throw new SystemException("暫出單總金額【" + allTotalAmount + "】已超過信用餘額【" + resultCredit + "】，無法新增!!");
                                        }

                                        return jsonResponse.ToString();
                                    }
                                    break;
                                #endregion
                                #region //未設定邏輯
                                default:
                                    throw new SystemException("單據名稱【" + DocType + "】尚未設定邏輯!!");
                                    #endregion
                            }
                            #endregion
                        }
                    }
                    else if (creditLimitControl == "y")
                    {
                        //y情況找COPMA設定參數
                        notesReceivableRate = MA091;
                        accountsReceivableRate = MA092;
                        unbilledSalesAmountRate = MA093;
                        orderAmountRate = MA094;
                        temporaryDisbursementAmountRate = MA102;

                        #region //先確認檢核卡控參數
                        switch (DocType)
                        {
                            #region //訂單SaleOrder
                            case "SaleOrder":
                                if (MA088 == "1")
                                {
                                    return jsonResponse.ToString();
                                }
                                else if (MA088 == "2")
                                {
                                    decimal resultCredit = GetCreditBalance();
                                    if (allTotalAmount > resultCredit)
                                    {
                                        jsonResponse = JObject.FromObject(new
                                        {
                                            status = "warning",
                                            msg = "訂單總金額【" + allTotalAmount + "】已超過信用餘額【" + resultCredit + "】，請確認是否繼續執行!!"
                                        });
                                    }

                                    return jsonResponse.ToString();
                                }
                                if (MA088 == "3")
                                {
                                    decimal resultCredit = GetCreditBalance();
                                    if (allTotalAmount > resultCredit)
                                    {
                                        throw new SystemException("訂單總金額【" + allTotalAmount + "】已超過信用餘額【" + resultCredit + "】，無法新增!!");
                                    }

                                    return jsonResponse.ToString();
                                }
                                break;
                            #endregion
                            #region //銷貨單ShippingOrder
                            case "ShippingOrder":
                                if (MA089 == "1")
                                {
                                    return jsonResponse.ToString();
                                }
                                else if (MA089 == "2")
                                {
                                    decimal resultCredit = GetCreditBalance();
                                    if (allTotalAmount > resultCredit)
                                    {
                                        jsonResponse = JObject.FromObject(new
                                        {
                                            status = "warning",
                                            msg = "銷貨單總金額【" + allTotalAmount + "】已超過信用餘額【" + resultCredit + "】，請確認是否繼續執行!!"
                                        });
                                    }

                                    return jsonResponse.ToString();
                                }
                                if (MA089 == "3")
                                {
                                    decimal resultCredit = GetCreditBalance();
                                    if (allTotalAmount > resultCredit)
                                    {
                                        throw new SystemException("銷貨單總金額【" + allTotalAmount + "】已超過信用餘額【" + resultCredit + "】，無法新增!!");
                                    }

                                    return jsonResponse.ToString();
                                }
                                break;
                            #endregion
                            #region //出貨通知單ShippingNotice
                            case "ShippingNotice":
                                if (MA120 == "1")
                                {
                                    return jsonResponse.ToString();
                                }
                                else if (MA120 == "2")
                                {
                                    decimal resultCredit = GetCreditBalance();
                                    if (allTotalAmount > resultCredit)
                                    {
                                        jsonResponse = JObject.FromObject(new
                                        {
                                            status = "warning",
                                            msg = "出貨通知單總金額【" + allTotalAmount + "】已超過信用餘額【" + resultCredit + "】，請確認是否繼續執行!!"
                                        });
                                    }

                                    return jsonResponse.ToString();
                                }
                                if (MA120 == "3")
                                {
                                    decimal resultCredit = GetCreditBalance();
                                    if (allTotalAmount > resultCredit)
                                    {
                                        throw new SystemException("出貨通知單總金額【" + allTotalAmount + "】已超過信用餘額【" + resultCredit + "】，無法新增!!");
                                    }

                                    return jsonResponse.ToString();
                                }
                                break;
                            #endregion
                            #region //暫出單TempShippingNote
                            case "TempShippingNote":
                                if (MA132 == "1")
                                {
                                    return jsonResponse.ToString();
                                }
                                else if (MA132 == "2")
                                {
                                    decimal resultCredit = GetCreditBalance();
                                    if (allTotalAmount > resultCredit)
                                    {
                                        jsonResponse = JObject.FromObject(new
                                        {
                                            status = "warning",
                                            msg = "暫出單總金額【" + allTotalAmount + "】已超過信用餘額【" + resultCredit + "】，請確認是否繼續執行!!"
                                        });
                                    }

                                    return jsonResponse.ToString();
                                }
                                if (MA132 == "3")
                                {
                                    decimal resultCredit = GetCreditBalance();
                                    if (allTotalAmount > resultCredit)
                                    {
                                        throw new SystemException("暫出單總金額【" + allTotalAmount + "】已超過信用餘額【" + resultCredit + "】，無法新增!!");
                                    }

                                    return jsonResponse.ToString();
                                }
                                break;
                            #endregion
                            #region //未設定邏輯
                            default:
                                throw new SystemException("單據名稱【" + DocType + "】尚未設定邏輯!!");
                            #endregion
                        }
                        #endregion
                    }
                    else
                    {
                        return jsonResponse.ToString();
                    }
                    #endregion

                    #region //計算信用額度
                    decimal GetCreditBalance()
                    {
                        #region //計算可超出額度
                        //可超出額度 = 信用額度(COPMA.MA033) * (1 + 可超出率%(COPMA.MA034))
                        excessCreditLimit = creaditLimit * (1 + creditLimitTolerance);
                        #endregion

                        bool versionFlag = true;

                        if (versionFlag == true)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.MA001,a.MA002
                                    ,ROUND((c1.TrSum * a.MA091 + c2.ChSum * a.MA092 + c3.RoSum * a.MA093  + c4.SoSum * a.MA094  + c5.TsSum * a.MA102 - c6.RbSum * a.MA093) , CAST(b.decimalLoAmoute AS INTEGER)) TattolSum
                                    --,c1.TrSum * a.MA091 TrSum
                                    --,c2.ChSum * a.MA092 ChSum
                                    --,c3.RoSum * a.MA093 RoSum
                                    --,c4.SoSum * a.MA094 SoSum
                                    --,c5.TsSum * a.MA102 TsSum
                                    --,a.MA091 TrRate,a.MA092 ChRate,a.MA093 RoRate,a.MA094 SoRate,a.MA102 TsRate
                                    FROM COPMA a
                                    OUTER APPLY(
                                        SELECT x1.MA003 CurrencyLocal,x2.MF003 decimalLoPrice,x2.MF004 decimalLoAmoute
                                        FROM CMSMA x1
                                        INNER JOIN CMSMF x2 on x1.MA003 = x2.MF001
                                    ) b
                                    OUTER APPLY(
                                        SELECT ISNULL(SUM(TC027 * TC003),0) TrSum
                                        FROM NOTTC 
                                        WHERE 1=1
                                        AND TC013 = a.MA001
                                        AND TC012 IN ('1','2','3','5','9','A')
                                    ) c1
                                    OUTER APPLY(
                                        SELECT 
                                        ISNULL(SUM(
		                                    (x1.TA041 + x1.TA042 - x1.TA058  ) * x2.MQ010
	                                    ),0) AS ChSum 
                                        FROM ACRTA x1
                                        INNER JOIN CMSMQ x2 on  x1.TA001 = x2.MQ001
                                        WHERE 1=1
                                        AND x1.TA004 = a.MA001
                                        AND x1.TA025 = 'Y'
                                        AND x1.TA027 = 'N'
                                    ) c2
                                    OUTER APPLY(
                                        SELECT ISNULL(SUM(x1.TH037 + x1.TH038 - x1.TH047- x1.TH048),0) RoSum
                                        FROM COPTH x1
                                        INNER JOIN COPTG x2 on x1.TH001 = x2.TG001 AND x1.TH002 = x2.TG002
                                        WHERE 1=1
                                        AND x1.TH026 = 'N'
                                        AND x2.TG023 = 'Y'
                                        AND x2.TG004 = a.MA001
                                    ) c3
                                    OUTER APPLY(
                                        SELECT  
                                        ISNULL(SUM(
                                            CASE 
                                                WHEN x2.TC016 = '2' THEN 
                                                    CASE 
                                                        WHEN x1.TD078 = 0 THEN 
                                                            (
                                                            ROUND(((x1.TD012)) , CAST(y.decimalLoAmoute AS INTEGER)) + 
                                                            ROUND(((x1.TD012) * x2.TC041), CAST(y.decimalLoAmoute AS INTEGER))
                                                            )
                                                        ELSE
                                                            (
                                                            ROUND(((x1.TD076 - x1.TD078) * x1.TD011 * x1.TD026) , CAST(y.decimalLoAmoute AS INTEGER)) + 
                                                            ROUND(((x1.TD076 - x1.TD078) * x1.TD011 * x1.TD026 * x2.TC041), CAST(y.decimalLoAmoute AS INTEGER))
                                                            )
                                                    END 
                                                WHEN x2.TC016 IN ('1','3','4','9') THEN 
                                                    CASE 
                                                        WHEN x1.TD078 = 0 THEN 
                                                            (
                                                            ROUND(((x1.TD012)) , CAST(y.decimalLoAmoute AS INTEGER)) 
                                                            )
                                                        ELSE
                                                            (
                                                            ROUND(((x1.TD076 - x1.TD078) * x1.TD011 * x1.TD026) , CAST(y.decimalLoAmoute AS INTEGER))
                                                            )
                                                    END 
                                            END),0) AS SoSum
                                        FROM COPTD x1
                                        INNER JOIN COPTC x2 on x1.TD001 = x2.TC001 AND x1.TD002 = x2.TC002
                                        OUTER APPLY(
                                            SELECT y1.MA003 CurrencyLocal,y2.MF003 decimalLoPrice,y2.MF004 decimalLoAmoute
                                            FROM CMSMA y1
                                            INNER JOIN CMSMF y2 on y1.MA003 = y2.MF001
                                        ) y
                                        WHERE 1=1
                                        AND x1.TD016 = 'N'
                                        AND x2.TC027 = 'Y'
                                        AND x2.TC004 = a.MA001
                                    ) c4
                                    OUTER APPLY(
                                        SELECT 
                                        ISNULL(SUM(
                                        CASE 
                                            WHEN x2.TF010 = '2' THEN
                                                    CASE
                                                        WHEN x1.TG020 = 0 THEN
                                                                (
                                                                    ROUND(((x1.TG013) * x2.TF012) , CAST(y.decimalLoAmoute AS INTEGER)) + 
                                                                    ROUND(((x1.TG013) * x2.TF012 * x2.TF026), CAST(y.decimalLoAmoute AS INTEGER))
                                                                )
                                                        ELSE
                                                                (
                                                                    ROUND(((x1.TG052 - x1.TG054 - x1.TG055) * x1.TG012 * x2.TF012) , CAST(y.decimalLoAmoute AS INTEGER)) + 
                                                                    ROUND(((x1.TG052 - x1.TG054 - x1.TG055) * x1.TG012 * x2.TF012 * x2.TF026), CAST(y.decimalLoAmoute AS INTEGER))
                                                                )
                                                    END 
                                            ELSE
                                                    CASE
                                                        WHEN x1.TG020 = 0 THEN
                                                                (
                                                                    ROUND(((x1.TG013) * x2.TF012) , CAST(y.decimalLoAmoute AS INTEGER))
                                                                )
                                                        ELSE
                                                                (
                                                                    ROUND(((x1.TG052 - x1.TG054 - x1.TG055) * x1.TG012 * x2.TF012) , CAST(y.decimalLoAmoute AS INTEGER))
                                                                )
                                                    END 
                                        END),0) AS TsSum
                                        FROM INVTG x1
                                        INNER JOIN INVTF x2 on x1.TG001 = x2.TF001 AND x1.TG002 = x2.TF002
                                        OUTER APPLY(
                                            SELECT y1.MA003 CurrencyLocal,y2.MF003 decimalLoPrice,y2.MF004 decimalLoAmoute
                                            FROM CMSMA y1
                                            INNER JOIN CMSMF y2 on y1.MA003 = y2.MF001
                                        ) y
                                        WHERE 1=1
                                        AND x1.TG014 =''
                                        AND x1.TG015 =''
                                        AND x1.TG016 =''
                                        AND x1.TG024 = 'N'
                                        AND x2.TF020 = 'Y'
                                        AND x2.TF005 = a.MA001 
                                    ) c5
                                    OUTER APPLY(
	                                    SELECT ISNULL(SUM(x2.TJ033 + x2.TJ034),0) RbSum
                                        FROM COPTI x1
                                        INNER JOIN COPTJ x2 on x1.TI001 = x2.TJ001 AND x1.TI002 = x2.TJ002
                                        WHERE 1=1
                                        AND x2.TJ024 = 'N'
                                        AND x2.TJ021 IN ('Y','N') 
                                        AND x1.TI004 = a.MA001 
                                    )c6
                                    WHERE a.MA001 = @CustomerNo";
                            dynamicParameters.Add("CustomerNo", CustomerNo);

                            var COPMAResult2 = sqlConnection.Query(sql, dynamicParameters);

                            if (COPMAResult2.Count() <= 0) throw new SystemException("計算應收合計資料錯誤!!");

                            foreach (var item in COPMAResult2)
                            {
                                totalAccountsReceivable = Convert.ToDecimal(item.TattolSum);
                            }
                        }
                        else
                        {
                            #region //計算應收合計
                            //【應收合計】=【應收票據】+【應收帳款】+【未結帳銷貨】+【訂貨金額】+【暫出金額】

                            #region //計算應收票據 
                            // * 未兌現應收票據比率 (COPMA. MA032 = Y)COPMH.MH004 (COPMA. MA032 = y)COPMA. MA091
                            decimal notesReceivable = 0;
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TC012 AS status,TC013 AS customer,SUM(ISNULL(TC003,0)) AS Amount
                                FROM NOTTC 
                                WHERE TC012 IN ('1','2','3','5','9') 
                                AND TC013 = @CustomerNo
                                GROUP BY TC012,TC013";
                            dynamicParameters.Add("CustomerNo", CustomerNo);

                            var NOTTCResult = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item in NOTTCResult)
                            {
                                notesReceivable = Convert.ToDecimal(item.Amount);
                            }
                            notesReceivable = notesReceivable * notesReceivableRate;
                            #endregion

                            #region //計算應收帳款
                            //【應收帳款】= (結帳單-溢收待抵單) * 應收帳款比率 (COPMA. MA032 = Y)COPMH.MH005 (COPMA. MA032 = y)COPMA. MA092

                            decimal accountsReceivable = 0;

                            #region //結帳單統計(原幣)
                            decimal invoiceAmount = 0;
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TA004 AS customer,SUM(ISNULL(a.ARtotal,0)) AS SumARtotal
                                FROM ACRTA a1
                                INNER JOIN (
                                SELECT TA001,TA002,TA029 AS AR,TA030 AS tax,TA031 AS ARreceived,(ISNULL(a.TA029,0) + ISNULL(a.TA030,0) - ISNULL(a.TA031,0)) AS ARtotal,TA004 AS customer
                                FROM ACRTA a
                                WHERE TA025 = 'Y' AND TA027 = 'N' AND TA019 = 'N' AND a.TA001 LIKE '61%')a ON a1.TA001 = a.TA001 AND a1.TA002 = a.TA002
                                WHERE a1.TA004 = @CustomerNo
                                GROUP BY TA004";
                            dynamicParameters.Add("CustomerNo", CustomerNo);

                            var ACRTAResult = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item in ACRTAResult)
                            {
                                invoiceAmount = Convert.ToDecimal(item.SumARtotal);
                            }
                            #endregion

                            #region //溢收待抵單統計(原幣)
                            decimal overpaymentToBeOffset = 0;
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TA004 AS customer,SUM(ISNULL(a.ARtotal,0)) AS SumARtotal
                                FROM ACRTA a1
                                INNER JOIN (
                                SELECT TA001,TA002,TA029 AS AR,TA030 AS tax,TA031 AS ARreceived,(ISNULL(a.TA029,0) + ISNULL(a.TA030,0) - ISNULL(a.TA031,0)) AS ARtotal,TA004 AS customer
                                FROM ACRTA a
                                WHERE TA025 = 'Y' AND TA027 = 'N' AND TA019 = 'N' AND a.TA001 LIKE '62%')a ON a1.TA001 = a.TA001 AND a1.TA002 = a.TA002
                                WHERE a1.TA004 = @CustomerNo
                                GROUP BY TA004";
                            dynamicParameters.Add("CustomerNo", CustomerNo);

                            var ACRTAResult2 = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item in ACRTAResult2)
                            {
                                overpaymentToBeOffset = Convert.ToDecimal(item.SumARtotal);
                            }
                            #endregion

                            accountsReceivable = (invoiceAmount - overpaymentToBeOffset) * accountsReceivableRate;
                            #endregion

                            #region //未結帳銷貨
                            //【未結帳銷貨】= (銷貨單 - 銷退單) * 未結帳銷貨金額比率(COPMA. MA032 = Y)COPMH.MH006 (COPMA. MA032 = y)COPMA. MA093

                            decimal unbilledSalesAmount = 0;

                            #region //銷貨單統計(原幣)
                            decimal salesAmount = 0;
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TG004 AS customer,SUM(ISNULL(salesTotal,0)) AS SumsalesTotal
                                FROM COPTG a1
                                INNER JOIN(SELECT TG001,TG002,TG013 AS 原幣銷貨金額sales,TG025 AS 原幣銷貨稅額salesTax, (ISNULL(TG013,0) + ISNULL(TG025,0)) AS salesTotal 
                                FROM COPTG a
                                LEFT JOIN COPTH b ON a.TG001 = b.TH001 AND a.TG002 = b.TH002 
                                WHERE TG023 = 'Y' AND TH026 = 'N'
                                GROUP BY TG001,TG002,TG013,TG025
                                )c ON c.TG001 = a1.TG001 AND c.TG002 = a1.TG002
                                WHERE TG004 = @CustomerNo
                                GROUP BY TG004";
                            dynamicParameters.Add("CustomerNo", CustomerNo);

                            var COPTGResult = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item in COPTGResult)
                            {
                                salesAmount = Convert.ToDecimal(item.SumsalesTotal);
                            }
                            #endregion

                            #region //銷退單統計(原幣)
                            decimal returnAmount = 0;
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TI004 AS customer,SUM(salesReturnTotal) AS SumsalesReturnTotal
                                FROM COPTI a1 
                                INNER JOIN (SELECT TI001,TI002,TI004 AS customer,TI010 AS 原幣銷退金額salesReturn,TI011 AS 原幣銷退稅額salesReturnTax,(ISNULL(TI010,0) + ISNULL(TI011,0)) AS salesReturnTotal
                                FROM COPTI a 
                                LEFT JOIN COPTJ b ON a.TI001 = b.TJ001 AND a.TI002 = b.TJ002
                                WHERE TI019 = 'Y' AND TJ024 = 'N'
                                GROUP BY TI004,TI010,TI011,TI001,TI002) c ON c.TI001 = a1.TI001 AND c.TI002 = a1.TI002
                                WHERE TI004 = @CustomerNo
                                GROUP BY TI004";
                            dynamicParameters.Add("CustomerNo", CustomerNo);

                            var COPTIResult = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item in COPTIResult)
                            {
                                returnAmount = Convert.ToDecimal(item.SumsalesReturnTotal);
                            }
                            #endregion

                            unbilledSalesAmount = (salesAmount - returnAmount) * unbilledSalesAmountRate;

                            #endregion

                            #region //訂貨金額統計(原幣)
                            // * 未出貨訂單金額比率(COPMA. MA032 = Y)COPMH.MH007 (COPMA. MA032 = y)COPMA. MA094
                            decimal orderAmount = 0;
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT c.TC004 AS customer,SUM(c.unsoldOrder) AS SumunsoldOrder
                                FROM COPTC a1
                                INNER JOIN (
                                    SELECT a.TC001,a.TC002,b.TD003,a.TC004,b.TD008 訂單數量,b.TD009 已交數量,b.TD011 單價,(ISNULL(b.TD008,0) - ISNULL(b.TD009,0))AS unsoldQTY, 
                                    ((ISNULL(b.TD008,0) - ISNULL(b.TD009,0)) * ISNULL(b.TD011,0))AS unsoldOrder
                                FROM COPTC a
                                LEFT JOIN COPTD b ON a.TC001 = b.TD001 AND a.TC002 = b.TD002
                                WHERE a.TC027 = 'Y' AND b.TD016 = 'N'   AND TD012 <> 0 
                                ) c ON c.TC001 = a1.TC001 AND c.TC002 = a1.TC002
                                WHERE  a1.TC004 = @CustomerNo
                                GROUP BY c.TC004";
                            dynamicParameters.Add("CustomerNo", CustomerNo);

                            var COPTCResult = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item in COPTCResult)
                            {
                                orderAmount = Convert.ToDecimal(item.SumunsoldOrder);
                            }

                            orderAmount = orderAmount * orderAmountRate;
                            #endregion

                            #region //暫出金額統計(原幣)
                            // * 未歸還暫出金額比率(COPMA. MA032 = Y)COPMH.MH008 (COPMA. MA032 = y)COPMA. MA102
                            decimal temporaryDisbursementAmount = 0;
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TF005 AS customer, SUM(c.unreturned) AS SumUnreturned
                                FROM INVTF a1
                                    INNER JOIN (
                                    SELECT TG001, TG002, TG003, TG009 數量, TG012 單價, TG020 轉進銷量, TG021 歸還量, ((ISNULL(TG009,0)-ISNULL(TG020,0)-ISNULL(TG021,0))) AS unreturnedQTY,
                                        ((ISNULL(TG009,0)-ISNULL(TG020,0)-ISNULL(TG021,0)) * ISNULL(TG012,0)) AS unreturned,TG014,TG015,TG016
                                    FROM INVTF a
                                        LEFT JOIN INVTG b ON a.TF001 = b.TG001 AND a.TF002 = b.TG002
                                    WHERE  TF020 = 'Y' AND TG024 = 'N' AND TG014 = ''
                                )c ON a1.TF001 = c.TG001 AND a1.TF002 = c.TG002
                                WHERE a1.TF005 = @CustomerNo --客戶代號
                                GROUP BY TF005";
                            dynamicParameters.Add("CustomerNo", CustomerNo);

                            var INVTFResult = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item in INVTFResult)
                            {
                                temporaryDisbursementAmount = Convert.ToDecimal(item.SumUnreturned);
                            }

                            temporaryDisbursementAmount = temporaryDisbursementAmount * temporaryDisbursementAmountRate;
                            #endregion

                            totalAccountsReceivable = notesReceivable + accountsReceivable + unbilledSalesAmount + orderAmount + temporaryDisbursementAmount;
                            #endregion
                        }

                        #region //計算最終信用餘額
                        //信用餘額 = 可超出額度 - 應收合計
                        creditBalance = excessCreditLimit - totalAccountsReceivable;
                        #endregion

                        #region //小數點後取位
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MF004
                                FROM CMSMF a 
                                WHERE a.MF001 = @Currency";
                        dynamicParameters.Add("Currency", Currency);

                        var CMSMFResult = sqlConnection.Query(sql, dynamicParameters);

                        if (CMSMFResult.Count() <= 0) throw new SystemException("金額小數點取位資料錯誤!!");

                        foreach (var item in CMSMFResult)
                        {
                            creditBalance = Decimal.Round(creditBalance, Convert.ToInt32(item.MF004));
                        }
                        #endregion

                        return creditBalance;
                    }
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

        #region //GetSaleOrderEIP -- 取得客戶訂單資料 -- Chia Yuan -- 2024.4.3
        public string GetSaleOrderEIP(int SoId, string SoErpNo, string SoErpFullNo, string CustomerPurchaseOrder, string SearchKey, string StartDate, string EndDate
            , int MemberId, int[] CustomerIds
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                if (CustomerIds == null) throw new SystemException("客戶資料錯誤!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //取得訂單頭資料
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.SoId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.CompanyId, a.SoErpPrefix, a.SoErpNo, a.[Version], a.CustomerId, a.SalesmenId, a.DepositPartial, a.DepositRate
                        , a.Currency, a.ExchangeRate, a.TaxNo, a.Taxation, a.BusinessTaxRate, a.DetailMultiTax, a.TotalQty
                        , CONVERT(decimal(18, 6), a.Amount) AS Amount
                        , CONVERT(decimal(18, 6), a.TaxAmount) AS TaxAmount
                        , a.ShipMethod, a.TradeTerm, a.PaymentTerm, a.PriceTerm, a.ConfirmStatus, a.ConfirmUserId, a.TransferStatus, a.TransferDate
                        , a.CustomerAddressFirst, a.CustomerAddressSecond, a.CustomerPurchaseOrder
                        , ISNULL(a.SoRemark, '') AS SoRemark
                        , FORMAT(a.SoDate, 'yyyy-MM-dd') AS SoDate
                        , FORMAT(a.DocDate, 'yyyy-MM-dd') AS DocDate
                        , a.SoErpPrefix + '-' + a.SoErpNo AS SoErpFullNo
                        , ISNULL(c.CustomerName,'') CustomerName, c.CustomerNo + ' ' + c.CustomerShortName AS CustomerWithNo
                        , ISNULL(d.SomId, '') SomId
                        , e.StatusName ConfirmName
                        , f.StatusName TransferName
                        , CONVERT(decimal(18,6), g.TotalAmount) AS TotalAmount
                        , t.TypeName AS CurrencyName
                        , ISNULL(u.UserNo, '') AS SalesmenNo, ISNULL(u.UserName,'') AS SalesmenName, ISNULL(u.Gender, '') AS SalesmenGender
                        , ISNULL(h.UserNo, '') AS ConfirmUserNo, ISNULL(h.UserName,'') AS ConfirmUserName, ISNULL(h.Gender, '') AS ConfirmUserGender";
                    sqlQuery.mainTables =
                        @"FROM (
	                        SELECT ROW_NUMBER() OVER (PARTITION BY a.SoErpNo ORDER BY a.[Version] DESC) as SortVersion, a.SoId
	                        FROM SCM.SaleOrder a
	                        WHERE a.CustomerId IN @CustomerIds
                        ) x
                        INNER JOIN SCM.SaleOrder a ON a.SoId = x.SoId
                        INNER JOIN SCM.Customer c ON c.CustomerId = a.CustomerId
                        LEFT JOIN BAS.[User] u ON u.UserId = a.SalesmenId
                        LEFT JOIN BAS.[User] h ON h.UserId = a.ConfirmUserId
                        LEFT JOIN SCM.SoModification d ON d.SoId = a.SoId
                        INNER JOIN BAS.[Status] e ON e.StatusSchema = 'ConfirmStatus' AND e.StatusNo = a.ConfirmStatus
                        INNER JOIN BAS.[Status] f ON f.StatusSchema = 'Boolean' AND f.StatusNo = a.TransferStatus
                        INNER JOIN BAS.[Type] t ON  t.TypeSchema = 'ExchangeRate.Currency' AND t.TypeNo = a.Currency
                        OUTER APPLY(
                            SELECT CASE a.Taxation
                                WHEN 1 THEN ROUND(ISNULL(SUM(x.Amount), 0) / (1 + a.BusinessTaxRate) + (ISNULL(SUM(x.Amount), 0) / (1 + a.BusinessTaxRate) * a.BusinessTaxRate), 2)
                                ELSE ISNULL(SUM(x.Amount), 0) + (ISNULL(SUM(x.Amount), 0) * a.BusinessTaxRate)
                                END TotalAmount
                            FROM SCM.SoDetail x
                            WHERE x.SoId = a.SoId
                        ) g";
                        sqlQuery.auxTables = "";
                        string queryCondition = @" AND x.SortVersion = 1";
                        dynamicParameters.Add("CustomerIds", CustomerIds);

                    if (!string.IsNullOrWhiteSpace(SearchKey))
                    {
                        SearchKey = SearchKey.Trim();
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey", @" AND EXISTS (
	                                            SELECT TOP 1 1
	                                            FROM SCM.SoDetail ba
	                                            INNER JOIN PDM.MtlItem bb ON bb.MtlItemId = ba.MtlItemId
	                                            WHERE a.SoId = ba.SoId
	                                                AND (ISNULL(bb.MtlItemNo, '') LIKE '%' + @SearchKey + '%' 
                                                    OR ISNULL(bb.MtlItemName, '') LIKE '%' + @SearchKey + '%'
                                                    OR ISNULL(ba.SoMtlItemName, '') LIKE '%' + @SearchKey + '%' 
                                                    OR ISNULL(ba.CustomerMtlItemNo, '') LIKE '%' + @SearchKey + '%'))", SearchKey);
                    }
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SoId", @" AND a.SoId = @SoId", SoId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SoErpNo", @" AND a.SoErpNo = @SoErpNo", SoErpNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerPurchaseOrder", @" AND a.CustomerPurchaseOrder LIKE '%' + @CustomerPurchaseOrder + '%'", CustomerPurchaseOrder);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SoErpFullNo", @" AND a.SoErpPrefix + '-' + a.SoErpNo LIKE '%' + @SoErpFullNo + '%'", SoErpFullNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.DocDate >= @StartDate ", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.DocDate <= @EndDate ", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.DocDate DESC";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    var resultOrder = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery).ToList();
                    #endregion

                    List<dynamic> resultCMSNJ = new List<dynamic>();
                    List<dynamic> resultCMSNA = new List<dynamic>();
                    List<ErpDocStatus> resultCOPTC = new List<ErpDocStatus>();

                    #region //判斷公司資料是否正確
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DISTINCT CompanyId, CompanyName, ErpDb
                            FROM BAS.Company
                            WHERE CompanyId IN @CompanyIds";
                    dynamicParameters.Add("CompanyIds", resultOrder.Select(s => s.CompanyId).Distinct());
                    var resultCorp = sqlConnection.Query(sql, dynamicParameters);
                    if (resultCorp.Count() <= 0) throw new SystemException("【公司】資料錯誤!");
                    foreach (var corp in resultCorp)
                    {
                        ErpConnectionStrings = ConfigurationManager.AppSettings[corp.ErpDb];
                        using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //取得運輸方式資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT 
                                    LTRIM(RTRIM(NJ001)) AS NJ001,
                                    LTRIM(RTRIM(NJ002)) AS NJ002,
                                    CompanyId = @CompanyId
                                    FROM CMSNJ
                                    WHERE NJ001 IN @NJ001s";
                            dynamicParameters.Add("NJ001s", resultOrder.Select(s => s.ShipMethod).Distinct());
                            dynamicParameters.Add("CompanyId", corp.CompanyId);
                            resultCMSNJ.AddRange(erpConnection.Query(sql, dynamicParameters).ToList());
                            if (resultCMSNJ.Count() <= 0) throw new SystemException("【運輸方式】資料錯誤!");
                            #endregion

                            #region //取得付款條件資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT 
                                    LTRIM(RTRIM(NA002)) AS NA002,
                                    LTRIM(RTRIM(NA003)) AS NA003,
                                    CompanyId = @CompanyId
                                    FROM CMSNA
                                    WHERE NA002 IN @NA002s";
                            dynamicParameters.Add("NA002s", resultOrder.Select(s => s.PaymentTerm).Distinct());
                            dynamicParameters.Add("CompanyId", corp.CompanyId);
                            resultCMSNA.AddRange(erpConnection.Query(sql, dynamicParameters).ToList());
                            if (resultCMSNA.Count() <= 0) throw new SystemException("【付款條件】資料錯誤!");
                            #endregion

                            #region //取得ERP客戶訂單頭資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT 
                                    LTRIM(RTRIM(TC001)) AS ErpPrefix,
                                    LTRIM(RTRIM(TC002)) AS ErpNo, 
                                    LTRIM(RTRIM(TC027)) AS DocStatus
                                    FROM COPTC
                                    WHERE (TC001 + '-' + TC002) IN @ErpFullNo";
                            dynamicParameters.Add("ErpFullNo", resultOrder.Select(s => s.SoErpFullNo).Distinct());
                            resultCOPTC = erpConnection.Query<ErpDocStatus>(sql, dynamicParameters).ToList();

                            //resultOrder = resultOrder.GroupJoin(resultCOPTC, x => x.SoErpFullNo, y => y.ErpPrefix + '-' + y.ErpNo, (x, y) => { x.DocStatus = y.FirstOrDefault()?.DocStatus ?? ""; return x; }).ToList();
                            //resultOrder = resultOrder.OrderByDescending(x => x.DocDate).ToList();
                            #endregion
                        }
                    }
                    #endregion

                    List<dynamic> resultCust = new List<dynamic>();
                    using (SqlConnection officialConnection = new SqlConnection(OfficialConnectionStrings))
                    {
                        int memberId = -1;
                        #region //取得會員資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.MemberId
                                FROM EIP.[Member] a
                                WHERE a.MemberId = @MemberId";
                        dynamicParameters.Add("MemberId", MemberId);
                        var resutl = officialConnection.Query(sql, dynamicParameters).ToList();
                        if (resutl.Count() <= 0) throw new SystemException("【會員】資料錯誤!");
                        foreach (var item in resutl)
                        {
                            memberId = item.MemberId;
                        }
                        #endregion

                        #region //取得會員客戶資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT DISTINCT b.CustomerId
                                , ISNULL(a.CustomerName, '') AS CustomerName
                                , ISNULL(a.CustomerEnglishName, '') AS CustomerEnglishName
                                , ISNULL(a.CustomerName + ISNULL(' ' + a.CustomerEnglishName, ''), '') AS CNameAndEName
                                , ISNULL(c.MemberIcon, -1) AS MemberIcon
                                FROM EIP.CsCustomer a
                                INNER JOIN EIP.CsCustomerDetail b ON b.CsCustId = a.CsCustId
                                LEFT JOIN EIP.[Member] c ON c.CsCustId = a.CsCustId
                                WHERE b.CustomerId IN @CustomerIds
                                AND c.MemberId = @MemberId";
                        dynamicParameters.Add("CustomerIds", CustomerIds);
                        dynamicParameters.Add("MemberId", memberId);
                        resultCust = officialConnection.Query(sql, dynamicParameters).ToList();
                        if (resultCust.Count() <= 0) throw new SystemException("【會員客戶】資料錯誤!");
                        #endregion
                    }

                    #region //資料集合處理
                    var result = (from s1 in resultOrder
                                  join s2 in resultCust 
                                     on s1.CustomerId equals s2.CustomerId
                                  join s3 in resultCMSNJ
                                    on new { s1.ShipMethod, s1.CompanyId } equals new { ShipMethod = s3.NJ001, s3.CompanyId } into a1
                                  from s3 in a1.DefaultIfEmpty()
                                　join s4 in resultCMSNA
                                    on new { s1.PaymentTerm, s1.CompanyId } equals new { PaymentTerm = s4.NA002, s4.CompanyId } into a2
                                  from s4 in a2.DefaultIfEmpty()
                                  join s5 in resultCOPTC
                                    on new { ErpPrefix = (string)s1.SoErpPrefix, ErpNo = (string)s1.SoErpNo } equals new { s5.ErpPrefix, s5.ErpNo } into a3
                                  from s5 in a3.DefaultIfEmpty()
                                  select new
                                  {
                                      s1.SoId,
                                      s1.CompanyId,
                                      s1.SoErpPrefix,
                                      s1.SoErpNo,
                                      s1.Version,
                                      s1.CustomerId,
                                      s1.SalesmenId,
                                      s1.DepositPartial,
                                      s1.DepositRate,
                                      s1.Currency,
                                      s1.CurrencyName,
                                      s1.ExchangeRate,
                                      s1.TaxNo,
                                      s1.Taxation,
                                      s1.BusinessTaxRate,
                                      s1.DetailMultiTax,
                                      s1.TotalQty,
                                      s1.Amount,
                                      s1.TaxAmount,
                                      s1.ShipMethod,
                                      ShipMethodName = s3?.NJ002, //ERP
                                      s1.TradeTerm,
                                      s1.PaymentTerm,
                                      PaymentTermName = s4?.NA003, //ERP
                                      s5?.DocStatus, //ERP
                                      s1.PriceTerm,
                                      s1.ConfirmStatus,
                                      s1.ConfirmUserId,
                                      s1.TransferStatus,
                                      s1.TransferDate,
                                      s1.CustomerAddressFirst,
                                      s1.CustomerAddressSecond,
                                      s1.CustomerPurchaseOrder,
                                      s1.SoRemark,
                                      s1.SoDate,
                                      s1.DocDate,
                                      s1.SoErpFullNo,
                                      s1.SalesmenNo,
                                      s1.SalesmenName,
                                      s1.SalesmenGender,
                                      s1.ConfirmUserNo,
                                      s1.ConfirmUserName,
                                      s1.ConfirmUserGender,
                                      s1.SomId,
                                      s1.ConfirmName,
                                      s1.TransferName,
                                      s1.TotalAmount,
                                      s2.CustomerName,
                                      s2.CustomerEnglishName,
                                      s2.CNameAndEName,
                                      s2.MemberIcon,
                                      s1.TotalCount
                                  })
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

        #region //GetSoDetailEIP -- 取得客戶訂單身資料 -- Chia Yuan 2024.4.8
        public string GetSoDetailEIP(int SoDetailId, int SoId, string SoErpFullNo, string TransferStatus, string SearchKey
            , int MemberId, int[] CustomerIds
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.SoDetailId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.SoId, a.SoSequence, ISNULL(a.MtlItemId, -1) MtlItemId, a.SoMtlItemName, a.SoMtlItemSpec
                        , a.InventoryId, a.SoQty, a.SiQty, a.ProductType, a.FreebieQty, a.FreebieSiQty
                        , a.SpareQty, a.SpareSiQty, a.UnitPrice, a.Amount, a.BusinessTaxRate
                        , a.Project, a.SoDetailRemark, a.SoPriceQty, a.SoPriceUomId
                        , LTRIM(RTRIM(ISNULL(a.CustomerMtlItemNo, ''))) AS CustomerMtlItemNo
                        , ISNULL(a.UomId, -1) AS UomId
                        , FORMAT(a.PromiseDate, 'yyyy-MM-dd') AS PromiseDate
                        , FORMAT(a.PcPromiseDate, 'yyyy-MM-dd') AS PcPromiseDate
                        , b.SoErpPrefix, b.SoErpNo, b.Currency, b.ConfirmStatus, b.TransferStatus, b.CustomerId
                        , ISNULL(c.UomNo, '') AS UomNo
                        , LTRIM(RTRIM(ISNULL(d.MtlItemNo, ''))) AS MtlItemNo
                        , e.TypeName AS ProductTypeName
                        , ISNULL(f.InventoryName, '') AS InventoryName";
                    sqlQuery.mainTables =
                        @"FROM SCM.SoDetail a
                        INNER JOIN SCM.SaleOrder b ON a.SoId = b.SoId
                        LEFT JOIN SCM.Inventory f ON f.InventoryId = a.InventoryId
                        LEFT JOIN PDM.UnitOfMeasure c ON a.UomId = c.UomId
                        LEFT JOIN PDM.MtlItem d ON a.MtlItemId = d.MtlItemId
                        LEFT JOIN BAS.[Type] e ON a.ProductType = e.TypeNo AND e.TypeSchema = 'SoDetail.ProductType'";
                    sqlQuery.auxTables = "";
                    string queryCondition = @" AND b.CustomerId IN @CustomerIds";
                    dynamicParameters.Add("CustomerIds", CustomerIds);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SoDetailId", @" AND a.SoDetailId = @SoDetailId", SoDetailId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SoId", @" AND a.SoId = @SoId", SoId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey", @" AND (
                                                                                        d.MtlItemName LIKE '%' + @SearchKey + '%' 
                                                                                        OR a.SoMtlItemName LIKE '%' + @SearchKey + '%' 
                                                                                        OR a.CustomerMtlItemNo LIKE '%' + @SearchKey + '%')", SearchKey);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.SoId DESC, a.SoSequence";
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

        #region//Mail樣版-客戶寄送
        public string SendCustomerMail(int CustomerId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    var customerNo = "";
                    #region
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.CustomerId, a.CustomerNo, a.CustomerName
                            FROM SCM.Customer a
                            WHERE a.CustomerId = @CustomerId";
                    dynamicParameters.Add("CustomerId", CustomerId);
                    var customerresult = sqlConnection.Query(sql, dynamicParameters);
                    if (customerresult.Count() <= 0) throw new SystemException("無客戶資料! 無法寄送");
                    foreach (var item in customerresult) {
                        customerNo = item.CustomerNo;
                    }
                    #endregion



                    #region //Mail資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.CustomerMailId, a.CustomerId, a.CustomerNo, a.CustomerName, a.SalesName, a.SalesMail, a.SalesMailCc, a.SalesMailBcc, a.SalesTax, a.ServerId, a.MailFrom
                            , b.Host, b.Port, b.SendMode, b.Account, b.Password
                            FROM SCM.CustomerMail a
                            LEFT JOIN BAS.MailServer b ON b.ServerId = a.ServerId
                            WHERE a.CustomerNo = @CustomerNo";
                    dynamicParameters.Add("CustomerNo", customerNo);
                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0) throw new SystemException("無客戶寄送資料! 無法寄送");
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
    }
}
