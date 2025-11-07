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
using System.Text;
using System.Transactions;
using System.Web;

namespace MESDA
{
    public class OutsourcingProductionDA
    {
        public string MainConnectionStrings = "";
        public string ErpConnectionStrings = "";
        public string HrmConnectionStrings = "";
        public string BpmDbConnectionStrings = "";

        public string BpmServerPath = "";
        public string BpmAccount = "";
        public string BpmPassword = "";

        public int CurrentCompany = -1;
        public int CurrentUser = -1;
        public int CreateBy = -1;
        public int LastModifiedBy = -1;
        public string UserLock = "";
        public string UserNo = "";
        public DateTime CreateDate = default(DateTime);
        public DateTime LastModifiedDate = default(DateTime);

        public string sql = "";
        public JObject jsonResponse = new JObject();
        public Logger logger = LogManager.GetCurrentClassLogger();
        public SqlQuery sqlQuery = new SqlQuery();
        public DynamicParameters dynamicParameters = new DynamicParameters();
        public MamoHelper mamoHelper = new MamoHelper();

        public OutsourcingProductionDA()
        {
            MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];
            HrmConnectionStrings = ConfigurationManager.AppSettings["HrmDb"];
            BpmDbConnectionStrings = ConfigurationManager.AppSettings["BpmDb"];

            BpmServerPath = ConfigurationManager.AppSettings["BpmServerPath"];
            BpmAccount = ConfigurationManager.AppSettings["BpmAccount"];
            BpmPassword = ConfigurationManager.AppSettings["BpmPassword"];

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
                UserLock = Convert.ToString(HttpContext.Current.Session["UserLock"]);
                UserNo = Convert.ToString(HttpContext.Current.Session["UserNo"]);

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
        #region //GetOutsourcingProduction -- 取得託外生產單資料 -- Ann 2022-09-06
        public string GetOutsourcingProduction(int OspId, int DepartmentId, string OspNo, int SupplierId, string OspStatus, string Status, string OspStartDate, string OspEndDate, string WoErpFullNo, string MtlItemNo
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.OspId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",a.CompanyId, a.DepartmentId, a.OspNo, FORMAT(a.OspDate, 'yyyy-MM-dd') OspDate, a.SupplierId, a.OspStatus, a.OspDesc, a.[Status]
                        , ISNULL(b.SupplierNo, '') SupplierNo, ISNULL(b.SupplierName, '') SupplierName, ISNULL(b.SupplierShortName, '') SupplierShortName
                        ,b.TelNoFirst,b.FaxNo,b.PassStationControl
                        , c.DepartmentNo, c.DepartmentName
                        , d.Address";
                    sqlQuery.mainTables =
                        @"FROM MES.OutsourcingProduction a
                        LEFT JOIN SCM.Supplier b ON a.SupplierId = b.SupplierId
                        INNER JOIN BAS.Department c ON a.DepartmentId = c.DepartmentId
                        INNER JOIN BAS.Company d ON a.CompanyId=d.CompanyId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "OspId", @" AND a.OspId = @OspId", OspId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DepartmentId", @" AND a.DepartmentId = @DepartmentId", DepartmentId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "OspNo", @" AND a.OspNo LIKE '%' + @OspNo + '%'", OspNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SupplierId", @" AND a.SupplierId = @SupplierId", SupplierId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "OspStatus", @" AND a.OspStatus = @OspStatus", OspStatus);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status = @Status", Status);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "OspStartDate", @" AND a.CreateDate >= @OspStartDate", OspStartDate.Length > 0 ? Convert.ToDateTime(OspStartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "OspEndDate", @" AND a.CreateDate <= @OspEndDate", OspEndDate.Length > 0 ? Convert.ToDateTime(OspEndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemNo", @" AND EXISTS (SELECT TOP 1 1
                                                                                                                   FROM MES.OspDetail x
                                                                                                                   INNER JOIN MES.ManufactureOrder xa ON x.MoId = xa.MoId
                                                                                                                   INNER JOIN MES.WipOrder xb ON xa.WoId = xb.WoId
                                                                                                                   INNER JOIN PDM.MtlItem xc ON xb.MtlItemId = xc.MtlItemId
                                                                                                                   WHERE x.OspId = a.OspId AND xc.MtlItemNo LIKE '%' + @MtlItemNo + '%')", MtlItemNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "WoErpFullNo", @" AND EXISTS (SELECT TOP 1 1
                                                                                                                     FROM MES.OspDetail x
                                                                                                                     INNER JOIN MES.ManufactureOrder xa ON x.MoId = xa.MoId
                                                                                                                     INNER JOIN MES.WipOrder xb ON xa.WoId = xb.WoId
                                                                                                                     WHERE x.OspId = a.OspId AND (xb.WoErpPrefix + '-' + xb.WoErpNo) LIKE '%' + @WoErpFullNo + '%')", WoErpFullNo);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.OspId DESC";
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

        #region //GetOspDetail -- 取得託外生產單詳細資料 -- Ann 2022-09-07
        public string GetOspDetail(int OspDetailId, int OspId, int MoId, string Status, int SupplierId, string WoErpFullNo, string OspNo
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.OspDetailId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.OspId, a.MoId, a.MoProcessId, a.ProcessCheckStatus, a.ProcessCheckType, a.OspQty
                        , a.SuppliedQty, ISNULL(a.ProcessCode, '') ProcessCode, ISNULL(a.ProcessCodeName, '') ProcessCodeName, FORMAT(a.ExpectedDate, 'yyyy-MM-dd') ExpectedDate, ISNULL(a.Remark, '') Remark
                        ,b.WoSeq
                        , c.WoErpPrefix, c.WoErpNo, c.UomId
                        , d.MtlItemNo, d.MtlItemName, d.MtlItemSpec, d.InventoryId
                        , e.ProcessAlias, e.ProcessId
                        , f.OspNo, f.SupplierId
                        , ISNULL(g.SupplierNo, '') SupplierNo, ISNULL(g.SupplierName, '') SupplierName
                        , h.DepartmentNo, h.DepartmentName,e.RoutingItemProcessDesc";
                    sqlQuery.mainTables =
                        @"FROM MES.OspDetail a
                        INNER JOIN MES.ManufactureOrder b ON a.MoId = b.MoId
                        INNER JOIN MES.WipOrder c ON b.WoId = c.WoId
                        INNER JOIN PDM.MtlItem d ON c.MtlItemId = d.MtlItemId
                        INNER JOIN MES.MoProcess e ON a.MoProcessId = e.MoProcessId
                        INNER JOIN MES.OutsourcingProduction f ON a.OspId = f.OspId
                        LEFT JOIN SCM.Supplier g ON f.SupplierId = g.SupplierId
                        INNER JOIN BAS.Department h ON f.DepartmentId = h.DepartmentId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND 1=1";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "OspDetailId", @" AND a.OspDetailId = @OspDetailId", OspDetailId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "OspId", @" AND a.OspId = @OspId", OspId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MoId", @" AND a.MoId = @MoId", MoId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND f.Status = @Status", Status);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SupplierId", @" AND f.SupplierId = @SupplierId", SupplierId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "WoErpFullNo", @" AND (c.WoErpPrefix + '-' + c.WoErpNo) LIKE '%' + @WoErpFullNo + '%'", WoErpFullNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "OspNo", @" AND f.OspNo LIKE '%' + @OspNo + '%'", OspNo);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "b.MoId,e.SortNumber ASC";
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

        #region //GetOspDetailForPop -- 取得託外生產單詳細資料(For Pop) -- Ann 2024-07-19
        public string GetOspDetailForPop(int OspDetailId, int OspId, int MoId, string Status
            , int SupplierId, string WoErpFullNo, string OspNo, string ProcessAlias
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.OspDetailId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.OspId, a.MoId, a.MoProcessId, a.ProcessCheckStatus, a.ProcessCheckType, a.OspQty
                        , a.SuppliedQty, a.ProcessCode, a.ProcessCodeName, FORMAT(a.ExpectedDate, 'yyyy-MM-dd') ExpectedDate
                        , b.WoSeq
                        , c.WoErpPrefix, c.WoErpNo, c.UomId
                        , d.MtlItemNo, d.MtlItemName, d.MtlItemSpec, d.InventoryId
                        , e.ProcessAlias, e.ProcessId
                        , f.OspNo, f.SupplierId
                        , g.SupplierNo, g.SupplierName
                        , h.DepartmentNo, h.DepartmentName,e.RoutingItemProcessDesc";
                    sqlQuery.mainTables =
                        @"FROM MES.OspDetail a
                        INNER JOIN MES.ManufactureOrder b ON a.MoId = b.MoId
                        INNER JOIN MES.WipOrder c ON b.WoId = c.WoId
                        INNER JOIN PDM.MtlItem d ON c.MtlItemId = d.MtlItemId
                        INNER JOIN MES.MoProcess e ON a.MoProcessId = e.MoProcessId
                        INNER JOIN MES.OutsourcingProduction f ON a.OspId = f.OspId
                        INNER JOIN SCM.Supplier g ON f.SupplierId = g.SupplierId
                        INNER JOIN BAS.Department h ON f.DepartmentId = h.DepartmentId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND NOT EXISTS (
                                                SELECT 1
                                                FROM MES.OspReceiptDetail x 
                                                WHERE x.OspDetailId = a.OspDetailId
                                                GROUP BY x.OspDetailId
                                                HAVING SUM(x.ReceiptQty) >= a.OspQty
                                            )";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "OspDetailId", @" AND a.OspDetailId = @OspDetailId", OspDetailId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "OspId", @" AND a.OspId = @OspId", OspId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MoId", @" AND a.MoId = @MoId", MoId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND f.Status = @Status", Status);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SupplierId", @" AND f.SupplierId = @SupplierId", SupplierId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "WoErpFullNo", @" AND (c.WoErpPrefix + '-' + c.WoErpNo +  '(' + CONVERT(VARCHAR(10), b.WoSeq) + ')') LIKE '%' + @WoErpFullNo + '%'", WoErpFullNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "OspNo", @" AND f.OspNo LIKE '%' + @OspNo + '%'", OspNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProcessAlias", @" AND e.ProcessAlias LIKE N'%' + @ProcessAlias + '%'", ProcessAlias);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "b.MoId,f.OspNo,e.SortNumber ASC";
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

        #region //GetBarcode -- 取得製令條碼資料 -- Ann 2023-08-31
        public string GetBarcode(int MoId, int OspDetailId)
        {
            try
            {
                if (MoId <= 0) throw new SystemException("製令ID不能為空!!");
                if (OspDetailId <= 0) throw new SystemException("託外生產單詳細資料ID不能為空!!");
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.BarcodeId, a.BarcodeNo, a.BarcodeQty, a.CurrentProdStatus, a.BarcodeStatus
                            , b.ProcessAlias CurrentProcessAlias
                            , ISNULL(c.ProcessAlias, '已完工') NextProcessAlias
                            , d.OspBarcodeId 
                            , e.StatusName
                            , f.TypeName
                            FROM MES.Barcode a
                            INNER JOIN MES.MoProcess b ON a.CurrentMoProcessId = b.MoProcessId
                            LEFT JOIN MES.MoProcess c ON a.NextMoProcessId = c.MoProcessId
                            LEFT JOIN MES.OspBarcode d ON a.BarcodeNo = d.BarcodeNo AND d.OspDetailId = @OspDetailId
                            INNER JOIN BAS.[Status] e ON a.BarcodeStatus = e.StatusNo AND e.StatusSchema = 'Barcode.BarcodeStatus'
                            INNER JOIN BAS.[Type] f ON a.CurrentProdStatus = f.TypeNo AND f.TypeSchema = 'Barcode.CurrentProdStatus'
                            WHERE a.MoId = @MoId
                            AND a.BarcodeStatus = '1'";
                    dynamicParameters.Add("MoId", MoId);
                    dynamicParameters.Add("OspDetailId", OspDetailId);

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

        #region //GetOspBarcode -- 取得託外生產條碼 -- Ann 2022-09-07
        public string GetOspBarcode(int OspBarcodeId, int OspDetailId, string BarcodeNo, int MoProcessId, int MoId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                if (OspDetailId <= 0) throw new SystemException("託外生產單詳細資料ID不能為空!!");
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.OspBarcodeId
                            , b.BarcodeId, b.BarcodeNo, b.BarcodeQty, b.CurrentProdStatus, b.BarcodeStatus
                            , d.ProcessAlias CurrentProcessAlias
                            , ISNULL(e.ProcessAlias, '已完工') NextProcessAlias
                            , a.OspBarcodeId
                            , f.StatusName
                            , g.TypeName
                            FROM MES.OspBarcode a 
                            INNER JOIN MES.Barcode b ON a.BarcodeNo = b.BarcodeNo
                            INNER JOIN MES.OspDetail c ON a.OspDetailId = c.OspDetailId
                            INNER JOIN MES.MoProcess d ON b.CurrentMoProcessId = d.MoProcessId
                            LEFT JOIN MES.MoProcess e ON b.NextMoProcessId = e.MoProcessId
                            INNER JOIN BAS.[Status] f ON b.BarcodeStatus = f.StatusNo AND f.StatusSchema = 'Barcode.BarcodeStatus'
                            INNER JOIN BAS.[Type] g ON b.CurrentProdStatus = g.TypeNo AND g.TypeSchema = 'Barcode.CurrentProdStatus'
                            INNER JOIN MES.OutsourcingProduction h ON c.OspId = h.OspId
                            WHERE h.CompanyId = @CompanyId 
                            AND a.OspDetailId = @OspDetailId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    dynamicParameters.Add("OspDetailId", OspDetailId);

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "BarcodeNo", @" AND a.BarcodeNo LIKE '%' + @BarcodeNo + '%'", BarcodeNo);

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

        #region //GetOspReceipt -- 取得託外入庫資料 -- Ann 2022-09-12
        public string GetOspReceipt(int OsrId, string OsrErpFullNo, int SupplierId, string OsrStartDate, string OsrEndDate, string OspNo, string WoErpFullNo, string MtlItemNo
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.OsrId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.OsrErpPrefix, a.OsrErpNo, FORMAT(a.ReceiptDate, 'yyyy-MM-dd') ReceiptDate, a.FactoryCode, a.SupplierId, a.SupplierSo, a.CurrencyCode
                        , a.Exchange, a.RowCnt, a.Remark, a.UiNo, a.InvoiceType, FORMAT(a.InvoiceDate, 'yyyy-MM-dd') InvoiceDate, a.InvoiceNo, a.TaxType, a.DeductType
                        , a.ReserveFlag, a.OrigAmount, a.DeductAmount, a.OrigTax, a.ReceiptAmount, a.Quantity, a.ConfirmStatus, a.ConfirmUserId
                        , a.RenewFlag, a.PrintCnt, a.AutoMaterialBilling, a.OrigPreTaxAmount, FORMAT(a.ApplyYYMM, 'yyyy-MM-dd') ApplyYYMM, FORMAT(a.DocumentDate, 'yyyy-MM-dd') DocumentDate
                        , a.TaxRate, a.PretaxAmount, a.TaxAmount, a.PaymentTerm, a.PackageQuantity, a.SupplierPicking, a.ApproveStatus
                        , a.ReserveTaxCode, a.SendCount, a.NoticeFlag, a.TaxCode, a.TaxExchange, a.QcFlag, a.FlowStatus, a.TransferStatus
                        , b.StatusNo, b.StatusName
                        , c.SupplierNo, c.SupplierName, c.PassStationControl
                        , d.UserNo, d.UserName
                        , e.UserNo ConfirmUserNo, e.UserName ConfirmUserName
                        , (
                            SELECT f.ProcessStatus
                            FROM MES.OspReceiptDetail f
                            WHERE f.OsrId = a.OsrId
                            FOR JSON PATH, ROOT('data')
                        ) ProcessStatus
                        , (
                            SELECT TOP 1 1
                            FROM MES.OspReceiptBarcode x 
                            INNER JOIN MES.OspBarcode xa ON x.OspBarcodeId = xa.OspBarcodeId
                            INNER JOIN MES.Barcode xb ON xa.BarcodeNo = xb.BarcodeNo
                            INNER JOIN MES.BarcodeProcess xc ON xb.BarcodeId = xc.BarcodeId
                            INNER JOIN MES.OspReceiptDetail xd ON x.OsrDetailId = xd.OsrDetailId
                            INNER JOIN MES.OspDetail xe ON xa.OspDetailId = xe.OspDetailId
                            INNER JOIN BAS.[User] xf ON xc.StartUserId = xf.UserId
                            WHERE xd.OsrId = a.OsrId
                            AND xc.MoProcessId = xe.MoProcessId
                            AND xf.UserNo = 'Z100'
                        ) SupplierMode";
                    sqlQuery.mainTables =
                        @"FROM MES.OspReceipt a
                        INNER JOIN BAS.[Status] b ON a.ConfirmStatus = b.StatusNo AND b.StatusSchema = 'ConfirmStatus'
                        INNER JOIN SCM.Supplier c ON a.SupplierId = c.SupplierId
                        INNER JOIN BAS.[User] d ON a.CreateBy = d.UserId
                        LEFT JOIN BAS.[User] e ON a.ConfirmUserId = e.UserId";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "OsrId", @" AND a.OsrId = @OsrId", OsrId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "OsrErpFullNo", @" AND (a.OsrErpPrefix + '-' + a.OsrErpNo) LIKE '%' + @OsrErpFullNo + '%'", OsrErpFullNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SupplierId", @" AND c.SupplierId = @SupplierId", SupplierId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "OsrStartDate", @" AND a.CreateDate >= @OsrStartDate", OsrStartDate.Length > 0 ? Convert.ToDateTime(OsrStartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "OsrEndDate", @" AND a.CreateDate <= @OsrEndDate", OsrEndDate.Length > 0 ? Convert.ToDateTime(OsrEndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemNo", @" AND EXISTS (SELECT TOP 1 1
                                                                                                                   FROM MES.OspReceiptDetail x
                                                                                                                   INNER JOIN MES.OspDetail xa ON x.OspDetailId = xa.OspDetailId
                                                                                                                   INNER JOIN MES.ManufactureOrder xb ON xa.MoId = xb.MoId
                                                                                                                   INNER JOIN MES.WipOrder xc ON xb.WoId = xc.WoId
                                                                                                                   INNER JOIN PDM.MtlItem xd ON xc.MtlItemId = xd.MtlItemId
                                                                                                                   WHERE x.OsrId = a.OsrId AND xd.MtlItemNo LIKE '%' + @MtlItemNo + '%')", MtlItemNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "OspNo", @" AND EXISTS (SELECT TOP 1 1
                                                                                                                   FROM MES.OspReceiptDetail x
                                                                                                                   INNER JOIN MES.OspDetail xa ON x.OspDetailId = xa.OspDetailId
                                                                                                                   INNER JOIN MES.OutsourcingProduction xb ON xa.OspId = xb.OspId
                                                                                                                   WHERE x.OsrId = a.OsrId AND xb.OspNo LIKE '%' + @OspNo + '%')", OspNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "WoErpFullNo", @" AND EXISTS (SELECT TOP 1 1
                                                                                                               FROM MES.OspReceiptDetail x
                                                                                                               INNER JOIN MES.OspDetail xa ON x.OspDetailId = xa.OspDetailId
                                                                                                               INNER JOIN MES.ManufactureOrder xb ON xa.MoId = xb.MoId
                                                                                                               INNER JOIN MES.WipOrder xc ON xb.WoId = xc.WoId
                                                                                                               WHERE x.OsrId = a.OsrId AND (xc.WoErpPrefix + '-' + xc.WoErpNo +  '(' + CONVERT(VARCHAR(10), xb.WoSeq) + ')') LIKE '%' + @WoErpFullNo + '%')", WoErpFullNo);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.OsrId DESC";
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

        #region //GetOspReceiptDetail -- 取得託外入庫詳細資料 -- Ann 2022-09-16
        public string GetOspReceiptDetail(int OsrId, int OsrDetailId, string OspNo, string WoErpFullNo, string MtlItemNo, string MtlItemName, int InventoryId, string OsrIdList
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.OsrDetailId, a.OsrId, a.OspDetailId, a.OsrErpPrefix, a.OsrErpNo, a.OsrSeq, a.MtlItemId, a.ReceiptQty, a.UomId, a.InventoryId
                            , a.LotNumber, FORMAT(a.AvailableDate, 'yyyy-MM-dd') AvailableDate, a.ReCheckDate, a.MoId, a.MoProcessId, a.ProcessCode, a.ReceiptPackageQty
                            , a.AcceptancePackageQty, FORMAT(a.AcceptanceDate, 'yyyy-MM-dd') AcceptanceDate, a.AcceptQty, a.AvailableQty, a.ScriptQty, a.ReturnQty, a.AvailableUom, a.OrigUnitPrice
                            , a.OrigAmount, a.OrigDiscountAmt, a.ReceiptExpense, a.DiscountDescription, a.ProjectCode, a.PaymentHold, a.Overdue, a.QcStatus
                            , a.ReturnCode, a.ConfirmStatus, a.CloseStatus, a.ReNewStatus, a.Remark, a.ConfirmUserId, a.ConfirmUser, a.OrigPreTaxAmt
                            , a.OrigTaxAmt, a.PreTaxAmt, a.TaxAmt, a.ProcessStatus
                            , (a1.OsrErpPrefix + '-' + a1.OsrErpNo) OsrFullNo,a2.SupplierName OsrSupplierName,a3.UserName OsrUserName, a1.TransferStatus, a1.ConfirmStatus OsrConfirmStatus, FORMAT(a1.DocumentDate, 'yyyy-MM-dd') DocumentDate, FORMAT(a1.ReceiptDate, 'yyyy-MM-dd') ReceiptDate
                            , b.MoProcessId, b.ProcessCode, b.ProcessCodeName
                            , c.OspNo
                            , d.MoId
                            , e.WoErpPrefix, e.WoErpNo
                            , f.MtlItemNo, f.MtlItemName, f.MtlItemSpec
                            , g.UomNo, g.UomName
                            , h.InventoryNo, h.InventoryName
                            , i.ProcessAlias
                            , j.StatusName
                            , k.UserNo ConfirmUserNo, k.UserName ConfirmUserName
                            , l.DepartmentNo, l.DepartmentName
                            , m.SupplierNo, m.SupplierName
                            ,(
                              SELECT na.OspBarcodeId
                              , nb.BarcodeQty
                              , ISNULL(nd.BarcodeProcessId, -1) BarcodeProcessId
                              FROM MES.OspReceiptBarcode na
                              INNER JOIN MES.OspBarcode nb ON na.OspBarcodeId = nb.OspBarcodeId
                              LEFT JOIN MES.Barcode nc ON nb.BarcodeNo = nc.BarcodeNo
                              LEFT JOIN MES.BarcodeProcess nd ON nc.BarcodeId = nd.BarcodeId AND nd.MoId = d.MoId AND nd.MoProcessId = i.MoProcessId
                              WHERE na.OsrDetailId = a.OsrDetailId
                              FOR JSON PATH, ROOT('data')
                            ) BarcodeProcessInfo
                            FROM MES.OspReceiptDetail a
                            INNER JOIN MES.OspReceipt a1  ON a.OsrId = a1.OsrId
                            INNER JOIN SCM.Supplier a2 ON a1.SupplierId = a2.SupplierId
                            INNER JOIN BAS.[User] a3 ON a1.CreateBy = a3.UserId
                            INNER JOIN MES.OspDetail b ON a.OspDetailId = b.OspDetailId
                            INNER JOIN MES.OutsourcingProduction c ON b.OspId = c.OspId
                            INNER JOIN MES.ManufactureOrder d ON a.MoId = d.MoId
                            INNER JOIN MES.WipOrder e ON d.WoId = e.WoId
                            INNER JOIN PDM.MtlItem f ON e.MtlItemId = f.MtlItemId
                            INNER JOIN PDM.UnitOfMeasure g ON a.UomId = g.UomId
                            INNER JOIN SCM.Inventory h ON a.InventoryId = h.InventoryId
                            INNER JOIN MES.MoProcess i ON b.MoProcessId = i.MoProcessId
                            INNER JOIN BAS.[Status] j ON a.ConfirmStatus = j.StatusNo AND StatusSchema = 'ConfirmStatus'
                            LEFT JOIN BAS.[User] k ON a.ConfirmUserId = k.UserId
                            INNER JOIN BAS.Department l ON c.DepartmentId = l.DepartmentId
                            INNER JOIN SCM.Supplier m ON c.SupplierId = m.SupplierId
                            WHERE e.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "OsrDetailId", @" AND a.OsrDetailId = @OsrDetailId", OsrDetailId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "OsrId", @" AND a.OsrId = @OsrId", OsrId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "OspNo", @" AND c.OspNo LIKE '%' + @OspNo + '%'", OspNo);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemNo", @" AND f.MtlItemNo LIKE '%' + @MtlItemNo + '%'", MtlItemNo);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemName", @" AND f.MtlItemName LIKE '%' + @MtlItemName + '%'", MtlItemName);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "WoErpFullNo", @" AND (e.WoErpPrefix + '-' + e.WoErpNo) LIKE '%' + @WoErpFullNo + '%'", WoErpFullNo);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "InventoryId", @" AND h.InventoryId = @InventoryId", InventoryId);
                    if (OsrIdList.Length > 0) BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "OsrIdList", @" AND a.OsrId IN @OsrIdList", OsrIdList.Split(','));

                    sql += @" ORDER BY a.OsrDetailId, a.OsrSeq";

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

        #region //GetOspReceiptBarcode -- 取得託外入庫條碼 -- Ann 2022-09-19
        public string GetOspReceiptBarcode(int OsrDetailId, int OspBarcodeId, string BarcodeNo, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.OsrDetailId, a.OspBarcodeId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.OspBarcodeId, FORMAT(a.CreateDate, 'yyyy-MM-dd') CreateDate, a.Status
                        , b.BarcodeNo, b.BarcodeQty
                        , c.UserNo, c.UserName";
                    sqlQuery.mainTables =
                        @"FROM MES.OspReceiptBarcode a
                        INNER JOIN MES.OspBarcode b ON a.OspBarcodeId = b.OspBarcodeId
                        INNER JOIN BAS.[User] c ON a.CreateBy = c.UserId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND 1=1";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "OsrDetailId", @" AND a.OsrDetailId = @OsrDetailId", OsrDetailId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "OspBarcodeId", @" AND a.OspBarcodeId = @OspBarcodeId", OspBarcodeId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status = @Status", Status);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "BarcodeNo", @" AND b.BarcodeNo LIKE '%' + @BarcodeNo + '%'", BarcodeNo);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.CreateDate DESC";
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

        #region //GetSupplierMachine -- 取得供應商對應機台資料 -- Ann 2024-03-25
        public string GetSupplierMachine(int SupplierId, int ProcessId, int MachineId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.SmId , a.SupplierId, a.ProcessId, a.MachineId
                            , b.MachineNo, b.MachineName + '(' + b.MachineDesc + ')' MachineName, b.MachineDesc
                            FROM SCM.SupplierMachine a 
                            INNER JOIN MES.Machine b ON a.MachineId = b.MachineId
                            INNER JOIN SCM.Supplier c ON a.SupplierId = c.SupplierId
                            WHERE 1=1";

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SupplierId", @" AND a.SupplierId = @SupplierId", SupplierId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ProcessId", @" AND a.ProcessId = @ProcessId", ProcessId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ProcMachineIdessId", @" AND a.MachineId = @MachineId", MachineId);

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

        #region //GetBarcodeExcel -- 取得製令條碼資料(Excel用) -- Shintokuro 2024-07-19
        public string GetBarcodeExcel(int OspId, string MoFullNo, string MoProcess, int OspDetailId)
        {
            try
            {
                if (MoFullNo.Length <= 0) throw new SystemException("製令ID不能為空!!");
                if (OspId <= 0) throw new SystemException("託外生產單ID不能為空!!");
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    int MoId = -1;
                    int MoProcessId = Convert.ToInt32(MoProcess);
                    string Status = "";
                    #region //判斷MES製令資料是否存在
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 MoId
                            FROM MES.ManufactureOrder a1
                            INNER JOIN MES.WipOrder a2 ON a1.WoId = a2.WoId
                            WHERE a2.WoErpPrefix +'-'+a2.WoErpNo +'('+CONVERT(VARCHAR, a1.WoSeq)  +')' = @MoFullNo
                            ";
                    dynamicParameters.Add("MoFullNo", MoFullNo);

                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0) throw new SystemException("找不到所選取製令【" + MoFullNo + "】,請重新確認!");
                    foreach (var item in result)
                    {
                        MoId = item.MoId;
                    }
                    #endregion

                    #region //判斷託外生產單是否存在
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 Status
                            FROM MES.OutsourcingProduction a
                            WHERE a.OspId = @OspId
                            ";
                    dynamicParameters.Add("OspId", OspId);

                    result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0) throw new SystemException("找不到託外生產單,請重新確認!");
                    foreach(var item in result)
                    {
                        Status = item.Status;
                    }
                    #endregion

                    #region //判斷製程是否存在
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 1
                            FROM MES.MoProcess a
                            WHERE a.MoProcessId = @MoProcessId
                            AND a.MoId = @MoId
                            ";
                    dynamicParameters.Add("MoProcessId", MoProcessId);
                    dynamicParameters.Add("MoId", MoId);

                    result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0) throw new SystemException("該製令找不到所選取製程【" + MoProcessId + "】,請重新確認!");
                    #endregion

                    #region //判斷託外生產單單身是否存在
                    if (OspDetailId > 0)
                    {
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                            FROM MES.OspDetail a
                            WHERE a.OspDetailId = @OspDetailId
                            ";
                        dynamicParameters.Add("OspDetailId", OspDetailId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.BarcodeNo
                                    FROM MES.OspBarcode a
                                    WHERE a.OspDetailId = @OspDetailId
                                    ";
                            dynamicParameters.Add("OspDetailId", OspDetailId);

                            result = sqlConnection.Query(sql, dynamicParameters);
                        }
                    }
                    #endregion

                    if(Status != "Y")
                    {
                        dynamicParameters = new DynamicParameters();
                        sql = @" SELECT a.MoId,a.BarcodeId,a.BarcodeNo,a.BarcodeQty
                             , b.ProcessAlias CurrentProcessAlias
                             , ISNULL(b1.ProcessAlias, '已完工') NextProcessAlias
                             , c.StatusName
                             , c1.TypeName
                             ,CASE 
                                WHEN  x.BarcodeNo is not null THEN 'Y'
                                ELSE  'N'
                             END CurrentBind
                             ,CASE 
                                WHEN y.ProcessStatus is null AND y.OspDetailId is null THEN 'N' 
                                WHEN y.ProcessStatus is null AND y.OspDetailId is NOT null THEN 'Y' 
                                WHEN y.ProcessStatus = 'Y' THEN 'N' 
                                WHEN y.ProcessStatus = 'N' THEN 'Y'   
                             END OtherBind
                             ,CASE 
                                WHEN z.OspBind is not null THEN 'Y' 
                                ELSE  'N'
                             END OspBind
                             FROM MES.Barcode a
                             INNER JOIN MES.ManufactureOrder a1 ON a.MoId = a1.MoId
                             INNER JOIN MES.WipOrder a2 ON a1.WoId = a2.WoId
                             INNER JOIN MES.MoProcess b ON a.CurrentMoProcessId = b.MoProcessId
                             LEFT JOIN MES.MoProcess b1 ON a.NextMoProcessId = b1.MoProcessId
                             INNER JOIN BAS.[Status] c ON a.BarcodeStatus = c.StatusNo AND c.StatusSchema = 'Barcode.BarcodeStatus'
                             INNER JOIN BAS.[Type] c1 ON a.CurrentProdStatus = c1.TypeNo AND c1.TypeSchema = 'Barcode.CurrentProdStatus'
                             OUTER APPLY(
                                 SELECT x1.BarcodeNo
                                 FROM MES.OspBarcode x1
                                 INNER JOIN MES.OspDetail x2 on x1.OspDetailId =x2.OspDetailId
                                 INNER JOIN MES.OutsourcingProduction x3 on x2.OspId =x3.OspId
                                 WHERE x1.BarcodeNo = a.BarcodeNo 
                                 AND x3.OspId = @OspId
                                 AND x2.MoProcessId = @MoProcessId
                             ) x
                             OUTER APPLY(
                                SELECT TOP 1 x4.ProcessStatus, x2.OspDetailId
                                FROM MES.OspBarcode x1
                                INNER JOIN MES.OspDetail x2 ON x1.OspDetailId = x2.OspDetailId
                                LEFT JOIN MES.OspReceiptBarcode x3 ON x1.OspBarcodeId = x3.OspBarcodeId
                                LEFT JOIN MES.OspReceiptDetail x4 ON x3.OsrDetailId = x4.OsrDetailId
                                WHERE x1.BarcodeNo = a.BarcodeNo 
                                AND x2.MoProcessId = @MoProcessId
                                ORDER BY x4.ProcessStatus ASC
                              ) y
                              OUTER APPLY(
                                  SELECT TOP 1 1 OspBind
                                  FROM MES.OspBarcode x1
                                  INNER JOIN MES.OspDetail x2 on x1.OspDetailId =x2.OspDetailId
                                  INNER JOIN MES.OutsourcingProduction x3 on x2.OspId =x3.OspId
                                  WHERE x3.OspId = @OspId
                                  AND x2.MoProcessId = @MoProcessId
                              ) z
                             WHERE a.MoId = @MoId
                             AND a.BarcodeStatus = '1'
                            ";
                        dynamicParameters.Add("MoId", MoId);
                        dynamicParameters.Add("OspId", OspId);
                        dynamicParameters.Add("MoProcessId", MoProcessId);
                        dynamicParameters.Add("OspDetailId", OspDetailId);

                    }
                    else
                    {
                        dynamicParameters = new DynamicParameters();
                        sql = @" SELECT b.BarcodeNo, b.BarcodeQty
                                , b1.ProcessAlias CurrentProcessAlias
                                , ISNULL(b2.ProcessAlias, '已完工') NextProcessAlias
                                ,b.BarcodeStatus,b.CurrentProdStatus
                                , c.StatusName
                                , c1.TypeName
                                ,x.OtherBind
                                FROM MES.OspBarcode a
                                INNER JOIN MES.OspDetail a1 on a.OspDetailId = a1.OspDetailId
                                LEFT JOIN MES.Barcode b on a.BarcodeNo = b.BarcodeNo AND a1.MoId =b.MoId
                                LEFT JOIN MES.MoProcess b1 ON b.CurrentMoProcessId = b1.MoProcessId
                                LEFT JOIN MES.MoProcess b2 ON b.NextMoProcessId = b2.MoProcessId
                                INNER JOIN BAS.[Status] c ON b.BarcodeStatus = c.StatusNo AND c.StatusSchema = 'Barcode.BarcodeStatus'
                                INNER JOIN BAS.[Type] c1 ON b.CurrentProdStatus = c1.TypeNo AND c1.TypeSchema = 'Barcode.CurrentProdStatus'
                                OUTER APPLY(
                                    SELECT TOP 1 'Y' OtherBind
                                    FROM MES.OspReceiptBarcode x1
                                    WHERE x1.OspBarcodeId = a.OspBarcodeId
                                ) x
                                WHERE a.OspDetailId = @OspDetailId
                                AND a1.MoId = @MoId
                            ";
                        dynamicParameters.Add("MoId", MoId);
                        dynamicParameters.Add("OspDetailId", OspDetailId);
                    }

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

        #region //托外加工商過站系統使用
        #region //GetOspDetailMachine -- 取得託外生產製程機台 -- Shintokuro 2024-03-14
        public string GetOspDetailMachine(int MoId, int MoProcessId)
        {
            try
            {
                if (MoId <= 0) throw new SystemException("MES製令ID不能為空!!");
                if (MoProcessId <= 0) throw new SystemException("製程ID不能為空!!");
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.MoId,c.MachineId FROM MES.MoProcess a
                            INNER JOIN MES.ProcessParameter b on a.ProcessId = b.ProcessId
                            INNER JOIN MES.ProcessMachine c on b.ParameterId = c.ParameterId
                            INNER JOIN MES.Machine d on c.MachineId = d.MachineId
                            WHERE MoProcessId = @MoProcessId
                            AND MoId =@MoId";
                    dynamicParameters.Add("MoProcessId", MoProcessId);
                    dynamicParameters.Add("MoId", MoId);

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

        #region //GetOspUser -- 取得託外加工商專屬使用者 -- Shintokuro 2024-03-14
        public string GetOspUser(int UserId,string Situation)
        {
            try
            {
                if (UserId <= 0) throw new SystemException("使用者ID不能為空!!");
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.UserId,a.UserNo, a.PasswordExpire
                            FROM BAS.[User] a
                            WHERE UserId = @UserId
                            AND UserId = 6438";
                    dynamicParameters.Add("UserId", UserId);

                    var result = sqlConnection.Query(sql, dynamicParameters);

                    if(result.Count()<=0 && Situation != "PassWeb") throw new SystemException("使用者查詢不到請重新確認!!");
                    foreach(var item in result)
                    {
                        
                        DateTime PasswordExpire = item.PasswordExpire;
                        if(PasswordExpire < DateTime.Now)
                        {
                            #region //將金額回寫單頭
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE BAS.[User] SET
                                    PasswordExpire = @PasswordExpire,
                                    LastModifiedDate = @LastModifiedDate
                                    WHERE UserId = @UserId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    PasswordExpire = DateTime.Now.AddDays(90),
                                    LastModifiedDate,
                                    UserId
                                });
                            var insertResult2 = sqlConnection.Query(sql, dynamicParameters);
                            #endregion
                        }
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

        #region //GetOspDetailData -- 托外過站介面-取得托外單身相關資料 -- Shintokuro 2024-03-15
        public string GetOspDetailData(int OspDetailId)
        {
            try
            {
                if (OspDetailId <= 0) throw new SystemException("加工商登入碼不能為空!!");
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.OspId, a.ProcessCode, a.ProcessCodeName,a.MoProcessId, 
                            FORMAT(a.ExpectedDate, 'yyyy-MM-dd') ExpectedDate, 
                            a1.OspNo,b.Quantity,
                            (b1.WoErpPrefix + '-' + b1.WoErpNo +'('+ CAST(b.WoSeq AS VARCHAR(255))+')') WoFullNo,
                            b1.MtlItemId,
                            b2.MtlItemNo,b2.MtlItemName,
                            c.ProcessAlias
                            FROM MES.OspDetail a
                            INNER JOIN MES.OutsourcingProduction a1 on a.OspId = a1.OspId
                            INNER JOIN MES.ManufactureOrder b on a.MoId = b.MoId
                            INNER JOIN MES.WipOrder b1 on b.WoId = b1.WoId
                            INNER JOIN PDM.MtlItem b2 on b1.MtlItemId = b2.MtlItemId
                            INNER JOIN MES.MoProcess c on a.MoProcessId = c.MoProcessId
                            WHERE a.OspDetailId = @OspDetailId";
                    dynamicParameters.Add("OspDetailId", OspDetailId);

                    var result = sqlConnection.Query(sql, dynamicParameters);

                    if (result.Count() <= 0) throw new SystemException("找不到相關托外生產資料,請重新確認!!");

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

        #region //GetOspMoProcess -- 取得託外加工製程(過站下拉用)-- Shintokuro 2024-03-18
        public string GetOspMoProcess(int OspDetailId)
        {
            try
            {
                //if (OspDetailId <= 0) throw new SystemException("托外單深不能為空!!");
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    int OspId = -1;
                    int MoId = -1;
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.MoId , a.OspId
                            FROM MES.OspDetail a
                            WHERE a.OspDetailId = @OspDetailId";
                    dynamicParameters.Add("OspDetailId", OspDetailId);
                    var result = sqlConnection.Query(sql, dynamicParameters);
                    foreach (var item in result)
                    {
                        OspId = item.OspId;
                        MoId = item.MoId;
                    }

                    dynamicParameters = new DynamicParameters();

                    sql = @"SELECT b.ProcessAlias,a.MoProcessId,a.OspId,a.MoId 
                            FROM MES.OspDetail a
                            INNER JOIN MES.MoProcess b on a.MoProcessId = b.MoProcessId
                            WHERE a.OspId = @OspId
                            AND a.MoId = @MoId";
                    dynamicParameters.Add("OspId", OspId);
                    dynamicParameters.Add("MoId", MoId);

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

        #region //UpdateOspNextMoProcessForLotMode-- 更新條碼下一站製程(批量修改) -- Shintokuro 2024-03-18
        public string UpdateOspNextMoProcessForLotMode(int OspDetailId, string BarcodeNoListString, int MoProcessId, int NewMoProcessId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int UserId = CreateBy;
                        int rowsAffected = 0;

                        var BarcodeNoList = BarcodeNoListString.Split(',');

                        foreach (var barcodeNo in BarcodeNoList)
                        {
                            string BarcodeNo = barcodeNo;

                            #region //判斷條碼資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.BarcodeId, a.MoId, a.BarcodeStatus, a.NextMoProcessId OriNextMoProcessId
                                    , a.BarcodeQty, a.CurrentProdStatus OriCurrentProdStatus
                                    FROM MES.Barcode a
                                    WHERE BarcodeNo = @BarcodeNo";
                            dynamicParameters.Add("BarcodeNo", BarcodeNo);

                            var barcodeResult = sqlConnection.Query(sql, dynamicParameters);
                            if (barcodeResult.Count() <= 0) throw new SystemException("條碼資料錯誤!");

                            string BarcodeStatus = "";
                            int BarcodeId = -1;
                            int OriNextMoProcessId = -1;
                            int BarcodeQty = -1;
                            string OriCurrentProdStatus = "";
                            int MoId = -1;
                            foreach (var item in barcodeResult)
                            {
                                if (item.BarcodeStatus != "1" && item.BarcodeStatus != "0")
                                {
                                    string ErrorMsg = "條碼目前狀態無法更改製程";
                                    switch (item.BarcodeStatus)
                                    {
                                        case "8":
                                            ErrorMsg = "條碼目前狀態【製令完工入庫】無法變更下製程";
                                            break;
                                        case "10":
                                            ErrorMsg = "條碼目前狀態【已出貨】無法變更下製程";
                                            break;
                                        case "3":
                                            ErrorMsg = "條碼目前狀態【可上料，但未綁定】無法變更下製程";
                                            break;
                                        case "4":
                                            ErrorMsg = "條碼目前狀態【已綁定上料】無法變更下製程";
                                            break;
                                        case "7":
                                            ErrorMsg = "條碼目前狀態【強制結束流程】無法變更下製程，請向【生管、製程】確認";
                                            break;
                                        default:
                                            ErrorMsg = "條碼目前狀態無法更改製程";
                                            break;
                                    }
                                    throw new SystemException(ErrorMsg);
                                }
                                else
                                {
                                    BarcodeStatus = item.BarcodeStatus;
                                    BarcodeId = item.BarcodeId;
                                    OriNextMoProcessId = item.OriNextMoProcessId;
                                    BarcodeQty = item.BarcodeQty;
                                    OriCurrentProdStatus = item.OriCurrentProdStatus;
                                    MoId = item.MoId;
                                    if (BarcodeStatus == "0") BarcodeStatus = "1";
                                }
                            }
                            #endregion

                            #region //如果TrayBarcode = Y , Replace BarcodeNo
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT b.TrayBarcode
                                      FROM MES.MoProcess a
                                           INNER JOIN MES.MoSetting b ON a.MoId = b.MoId
                                     WHERE a.MoProcessId = @MoProcessId";
                            dynamicParameters.Add("MoProcessId", MoProcessId);
                            var MoSettingResult = sqlConnection.Query(sql, dynamicParameters);
                            if (MoSettingResult.Count() <= 0) throw new SystemException("找不到制令設定資訊!");
                            string TrayBarcode = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).TrayBarcode;

                            if (TrayBarcode.Equals("Y"))
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.TrayId,a.TrayNo,a.BarcodeNo,a.TrayCapacity
                                          FROM MES.Tray a
                                               INNER JOIN MES.Barcode b ON a.BarcodeNo = b.BarcodeNo
                                         WHERE a.BarcodeNo = @BarcodeNo
                                           AND a.[Status] = 'A'
                                           AND a.CompanyId = @CompanyId";

                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                dynamicParameters.Add("BarcodeNo", BarcodeNo);
                                var ToBarcodeResult = sqlConnection.Query(sql, dynamicParameters);
                                if (ToBarcodeResult.Count() <= 0) throw new SystemException("該Tray盤Barcode資料錯誤,找不到對應資料!");
                                BarcodeNo = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).BarcodeNo;
                            }

                            #endregion

                            #region //判斷此條碼目前是否為非加工狀態
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.BarcodeProcess a
                                    WHERE BarcodeId = @BarcodeId AND a.FinishDate IS NULL";
                            dynamicParameters.Add("BarcodeId", BarcodeId);

                            var barcodeProcessResult = sqlConnection.Query(sql, dynamicParameters);
                            if (barcodeProcessResult.Count() > 0) throw new SystemException("條碼目前為【加工中】狀態，無法更改!");
                            #endregion

                            #region //判斷原製程資訊是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.MoProcess a
                                    WHERE MoProcessId = @MoProcessId";
                            dynamicParameters.Add("MoProcessId", MoProcessId);

                            var moProcessResult = sqlConnection.Query(sql, dynamicParameters);
                            if (moProcessResult.Count() <= 0) throw new SystemException("原製程資訊資料錯誤，請確認製令是否正確");
                            #endregion

                            #region //判斷條碼是否在品異單中且未判定
                            sql = @"SELECT Top 1 a.AqBarcodeId ,a.JudgeStatus 
                                    FROM QMS.AqBarcode a
                                    LEFT JOIN MES.Barcode b ON a.BarcodeId = b.BarcodeId
                                    WHERE b.BarcodeNo = @BarcodeNo
                                    Order By a.LastModifiedDate DESC";
                            dynamicParameters.Add("BarcodeNo", BarcodeNo);
                            var result = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item in result)
                            {
                                if (item.JudgeStatus == null) throw new SystemException("條碼【" + BarcodeNo + "】目前為品異判定中，無法更改!");
                            }
                            #endregion

                            #region //判斷托外加工商指向製程是否可以指向
                            int OspId = -1;
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.MoId , a.OspId
                                    FROM MES.OspDetail a
                                    WHERE a.OspDetailId = @OspDetailId";
                            dynamicParameters.Add("OspDetailId", OspDetailId);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            foreach (var item in result)
                            {
                                OspId = item.OspId;
                                MoId = item.MoId;
                            }

                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT b.ProcessAlias,a.MoProcessId,a.OspId,a.MoId 
                                    FROM MES.OspDetail a
                                    INNER JOIN MES.MoProcess b on a.MoProcessId = b.MoProcessId
                                    WHERE a.OspId = @OspId
                                    AND a.MoId = @MoId
                                    AND b.MoProcessId = @MoProcessId";
                            dynamicParameters.Add("OspId", OspId);
                            dynamicParameters.Add("MoId", MoId);
                            dynamicParameters.Add("MoProcessId", NewMoProcessId);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("該製程不再可以指定的站別內，請重新確認");
                            #endregion

                            if (NewMoProcessId != -1)
                            {
                                #region //判斷新製程資訊是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM MES.MoProcess a
                                        WHERE MoProcessId = @NewMoProcessId";
                                dynamicParameters.Add("NewMoProcessId", NewMoProcessId);

                                var newMoProcessResult = sqlConnection.Query(sql, dynamicParameters);
                                if (newMoProcessResult.Count() <= 0) throw new SystemException("新製程資訊資料錯誤，請確認製令是否正確");
                                #endregion

                                BarcodeStatus = "1";
                            }
                            else
                            {
                                BarcodeStatus = "0";
                            }

                            #region //更新條碼製程
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE MES.Barcode SET
                                    NextMoProcessId = @NewMoProcessId,
                                    BarcodeStatus = @BarcodeStatus,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE BarcodeNo = @BarcodeNo";
                            dynamicParameters.AddDynamicParams(
                              new
                              {
                                  NewMoProcessId,
                                  BarcodeStatus,
                                  LastModifiedDate,
                                  LastModifiedBy = UserId,
                                  BarcodeNo
                              });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //INSERT MES.AdvancedLog
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.AdvancedLog (BarcodeId, MoProcessId, OriNextMoProcessId, NextMoProcessId, OriQty, Qty, OriCurrentProdStatus, CurrentProdStatus
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.LogId
                                    VALUES (@BarcodeId, @MoProcessId, @OriNextMoProcessId, @NextMoProcessId, @OriQty, @Qty, @OriCurrentProdStatus, @CurrentProdStatus
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    BarcodeId,
                                    MoProcessId,
                                    OriNextMoProcessId,
                                    NextMoProcessId = NewMoProcessId,
                                    OriQty = BarcodeQty,
                                    Qty = BarcodeQty,
                                    OriCurrentProdStatus,
                                    CurrentProdStatus = OriCurrentProdStatus,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy = UserId,
                                    LastModifiedBy = UserId
                                });

                            var insertResult2 = sqlConnection.Query(sql, dynamicParameters);

                            if (insertResult2.Count() <= 0) throw new SystemException("進階功能紀錄LOG失敗，請重新再試一次!!");

                            rowsAffected += insertResult2.Count();
                            #endregion

                            #region //重新計算條碼狀態報表資料
                            UpdateMoQtyInformation(MoId);
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

        #region //Update MoProcessQty
        public void UpdateMoQtyInformation(int MoId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;

                        #region //更新MoProcess PassQty,NgQty,ScrapQty,InputQty
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.MoProcess
                                   SET TotalPassQty = y.PassQty,
                                       TotalNgQty = y.NgQty,
	                                   TotalScrapQty = y.ScrapQty,
	                                   TotalInputQty = y.InputQty
                                  FROM MES.MoProcess x,
	                                   (SELECT t.MoProcessId,
			                                   t.ProcessAlias,
			                                   SUM(t.PassQty) PassQty,
			                                   SUM(t.NgQty) NgQty,
			                                   SUM(t.ScrapQty) ScrapQty,
			                                   SUM(t.PassQty) + SUM(t.NgQty) + SUM(t.ScrapQty) InputQty
		                                  FROM (
				                                SELECT a.MoProcessId,a.ProcessAlias,
					                                   CASE WHEN b.ProdStatus = 'P' THEN SUM(b.StationQty) ELSE 0 END PassQty,
					                                   CASE WHEN b.ProdStatus = 'F' THEN SUM(b.StationQty) ELSE 0 END NgQty,
					                                   CASE WHEN b.ProdStatus = 'S' THEN SUM(b.StationQty) ELSE 0 END ScrapQty
				                                  FROM MES.MoProcess a
					                                   INNER JOIN MES.BarcodeProcess b ON a.MoProcessId = b.MoProcessId
				                                 GROUP BY a.MoProcessId,a.ProcessAlias,b.ProdStatus) t
		                                 GROUP BY t.MoProcessId,t.ProcessAlias) y
                                 WHERE x.MoProcessId  = y.MoProcessId
                                   AND x.MoId = @MoId";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            MoId
                        });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        #endregion

                        #region //找出所有MoProcess依順序更新WipQty
                        //目前無法有邏輯推算每站在製數量, 等確定邏輯再補
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.MoProcess
                                   SET WipQty = ISNULL(b.WipQty,0)
                                  FROM MES.MoProcess a
                                       LEFT JOIN (SELECT x.CurrentMoProcessId,SUM(x.BarcodeQty) WipQty
	                                                FROM MES.Barcode x
				                                   WHERE x.BarcodeStatus = 1
				                                     AND EXISTS(SELECT 1
					                                              FROM MES.BarcodeProcess z
								                                 WHERE x.BarcodeId = z.BarcodeId
								                                   AND z.MoId = x.MoId)
				                                   GROUP BY x.CurrentMoProcessId) b ON a.MoProcessId = b.CurrentMoProcessId
                                 WHERE a.MoId = @MoId";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            MoId
                        });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //更新Mo Scrap Quantity
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.ManufactureOrder
                                   SET ScrapQty = y.ScrapQty
                                  FROM MES.ManufactureOrder x,
                                       (SELECT a.MoId,ISNULL(SUM(b.TotalScrapQty),0) ScrapQty
                                          FROM MES.ManufactureOrder a
                                               INNER JOIN MES.MoProcess b ON a.MoId = b.MoId
	                                     GROUP BY a.MoId) y
                                 WHERE x.MoId = y.MoId
                                   AND x.MoId = @MoId";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            MoId
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
        }
        #endregion 
        #endregion
        #endregion

        #region //Add
        #region //AddOutsourcingProduction -- 新增託外生產單資料 -- Ann 2022-09-06
        public string AddOutsourcingProduction(int DepartmentId, string OspNo, string OspDate, int SupplierId, string OspStatus
            , string OspDesc, string Status)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (OspDate.Length <= 0) throw new SystemException("【托外生產日期】不能為空!");
                        if (OspStatus.Length <= 0) throw new SystemException("【供料狀態】不能為空!");
                        if (OspDesc.Length > 100) throw new SystemException("【描述】狀態錯誤!");
                        if (Status.Length <= 0) throw new SystemException("【狀態】狀態錯誤!");

                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷部門資料是否有誤
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Department
                                WHERE CompanyId = @CompanyId
                                AND DepartmentId = @DepartmentId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("DepartmentId", DepartmentId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("部門資料有誤!");
                        #endregion

                        #region //判斷供應商資料是否有誤
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 SupplierNo
                                FROM SCM.Supplier
                                WHERE CompanyId = @CompanyId
                                AND SupplierId = @SupplierId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("SupplierId", SupplierId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("供應商資料有誤!");

                        string SupplierNo = "";
                        foreach (var item in result2)
                        {
                            SupplierNo = item.SupplierNo;
                        }
                        #endregion

                        #region //取得單日供應商序號
                        dynamicParameters = new DynamicParameters();
                        string nowDate = DateTime.Now.ToString("yyyyMMdd");
                        sql = @"SELECT ISNULL(MAX(OspNo), '1-0') MaxOspNo
                                FROM MES.OutsourcingProduction a
                                WHERE a.CompanyId = @CompanyId
                                AND a.OspNo LIKE '" + nowDate + SupplierNo + "%'";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in result3)
                        {
                            string maxOspNo = item.MaxOspNo;
                            int count = Convert.ToInt32(maxOspNo.Split('-')[1]) + 1;
                            OspNo = nowDate + SupplierNo + "-" + count.ToString("D4");
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.OutsourcingProduction (CompanyId, DepartmentId, OspNo, OspDate, SupplierId, OspStatus, OspDesc, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.OspId, INSERTED.OspNo, INSERTED.CreateBy, INSERTED.CompanyId
                                VALUES (@CompanyId, @DepartmentId, @OspNo, @OspDate, @SupplierId, @OspStatus, @OspDesc, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                DepartmentId,
                                OspNo,
                                OspDate,
                                SupplierId,
                                OspStatus,
                                OspDesc,
                                Status,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });

                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = insertResult.Count();

                        foreach (var item in insertResult)
                        {
                            if (BaseHelper.CheckUserAuthority(item.CreateBy, item.CompanyId, "A", "OutsourcingProduction", "add", sqlConnection).Equals("N")) throw new SystemException("人員權限檢核有異常，請嘗試重新登入!");
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

        #region //AddOspDetail -- 新增託外生產單詳細資料 -- Ann 2022-09-07
        public string AddOspDetail(int OspId, int MoId, int MoProcessId, string ProcessCheckStatus, string ProcessCheckType, int OspQty, int SuppliedQty
            , string ProcessCode, string ExpectedDate, string Remark)
        {
            try
            {
                int rowsAffected = 0;

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (ProcessCheckStatus.Length <= 0) throw new SystemException("【是否支援工程檢】不能為空!");
                        if (ProcessCheckStatus == "Y" && ProcessCheckType.Length <= 0) throw new SystemException("【工程檢頻率】不能為空!");
                        if (OspQty < 0) throw new SystemException("【託外生產數量】不能為空!");
                        if (SuppliedQty < 0) throw new SystemException("【供貨生產數量】不能為空!");
                        if (ExpectedDate.Length <= 0) throw new SystemException("【預計回廠日期】不能為空!");

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

                        #region //判斷託外生產單資料是否有誤
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status, SupplierId
                                FROM MES.OutsourcingProduction
                                WHERE OspId = @OspId";
                        dynamicParameters.Add("OspId", OspId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("託外生產單資料有誤!");

                        int SupplierId = -1;
                        foreach (var item in result)
                        {
                            if (item.Status != "N" && item.Status != "E") throw new SystemException("託外生產單已確認，無法更改!");
                            SupplierId = item.SupplierId;
                        }
                        #endregion

                        #region //判斷MES製令資料是否有誤
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ModeId
                                FROM MES.ManufactureOrder
                                WHERE MoId = @MoId";
                        dynamicParameters.Add("MoId", MoId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("MES製令資料有誤!");

                        int ModeId = -1;
                        foreach (var item in result2)
                        {
                            ModeId = item.ModeId;
                        }
                        #endregion

                        #region //確認MES製令製程資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ProcessId
                                FROM MES.MoProcess 
                                WHERE MoProcessId = @MoProcessId";
                        dynamicParameters.Add("MoProcessId", MoProcessId);

                        var MoProcessResult = sqlConnection.Query(sql, dynamicParameters);

                        if (MoProcessResult.Count() <= 0) throw new SystemException("【製令製程】資料錯誤!!");

                        int ProcessId = -1;
                        foreach (var item in MoProcessResult)
                        {
                            ProcessId = item.ProcessId;
                        }
                        #endregion

                        string ProcessCodeName = "";
                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //判斷ERP製程代號資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 MW002
                                    FROM CMSMW
                                    WHERE MW001 = @MW001";
                            dynamicParameters.Add("MW001", ProcessCode);

                            var result3 = sqlConnection2.Query(sql, dynamicParameters);
                            if (result3.Count() <= 0) throw new SystemException("ERP製程代號資料有誤!");

                            foreach (var item in result3)
                            {
                                ProcessCodeName = item.MW002;
                            }
                            #endregion
                        }

                        #region //判斷製令製程是否已存在且已過站
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.OspReceiptDetail a 
                                INNER JOIN MES.OspDetail b ON a.OspDetailId = b.OspDetailId
                                WHERE b.MoProcessId = @MoProcessId
                                AND a.ProcessStatus = 'N'";
                        dynamicParameters.Add("MoProcessId", MoProcessId);

                        var result4 = sqlConnection.Query(sql, dynamicParameters);
                        
                        if (result4.Count() > 0) throw new SystemException("此製令製程已存在託外單據，且未過站!!");
                        #endregion

                        #region //若需工程檢，自動開立量測單據
                        if (ProcessCheckStatus == "Y")
                        {
                            #region //取得預設SpreadsheetData
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.DefaultSpreadsheetData
                                    FROM MES.QcRecord a";

                            var DefaultFileIdResult = sqlConnection.Query(sql, dynamicParameters);

                            string DefaultSpreadsheetData = "";
                            foreach (var item2 in DefaultFileIdResult)
                            {
                                DefaultSpreadsheetData = item2.DefaultSpreadsheetData;
                            }
                            #endregion

                            #region //取得此生產模式下IPQC資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.QcTypeId
                                    FROM QMS.QcType a 
                                    WHERE a.ModeId = @ModeId
                                    AND a.QcTypeNo = 'IPQC'";
                            dynamicParameters.Add("ModeId", ModeId);

                            var QcTypeResult = sqlConnection.Query(sql, dynamicParameters);

                            if (QcTypeResult.Count() <= 0) throw new SystemException("此生產模式尚未設定量測類型，請通知系統開發室!!");

                            int QcTypeId = -1;
                            foreach (var item2 in QcTypeResult)
                            {
                                QcTypeId = item2.QcTypeId;
                            }
                            #endregion

                            #region //INSERT MES.QcRecord
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.QcRecord (QcTypeId, InputType, MoId, MoProcessId, Remark, DefaultFileId, CurrentFileId, DefaultSpreadsheetData, CheckQcMeasureData
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.QcRecordId
                                    VALUES (@QcTypeId, @InputType, @MoId, @MoProcessId, @Remark, @DefaultFileId, @CurrentFileId, @DefaultSpreadsheetData, @CheckQcMeasureData
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    QcTypeId,
                                    InputType = "BarcodeNo",
                                    MoId,
                                    MoProcessId,
                                    Remark = "由託外生產單自動建立",
                                    DefaultFileId = -1,
                                    CurrentFileId = -1,
                                    DefaultSpreadsheetData,
                                    CheckQcMeasureData = "N",
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }
                        #endregion

                        #region //確認供應商是否為自主過站模式，若是則檢核是否已設定過站機台
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.PassStationControl
                                FROM SCM.Supplier a 
                                WHERE a.SupplierId = @SupplierId";
                        dynamicParameters.Add("SupplierId", SupplierId);

                        var SupplierResult = sqlConnection.Query(sql, dynamicParameters);

                        if (SupplierResult.Count() <= 0) throw new SystemException("供應商資料錯誤!!");

                        foreach (var item in SupplierResult)
                        {
                            if (item.PassStationControl == "Y")
                            {
                                #region //檢核是否已設定供應商過站機台
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM SCM.SupplierMachine a 
                                        WHERE a.SupplierId = @SupplierId
                                        AND a.ProcessId = @ProcessId";
                                dynamicParameters.Add("SupplierId", SupplierId);
                                dynamicParameters.Add("ProcessId", ProcessId);

                                var SupplierMachineResult = sqlConnection.Query(sql, dynamicParameters);

                                if (SupplierMachineResult.Count() <= 0) throw new SystemException("此供應商尚未維護此製程機台!!");
                                #endregion
                            }
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.OspDetail (OspId, MoId, MoProcessId, ProcessCheckStatus, ProcessCheckType, OspQty, SuppliedQty, ProcessCode, ProcessCodeName, ExpectedDate, Remark
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.OspDetailId
                                VALUES (@OspId, @MoId, @MoProcessId, @ProcessCheckStatus, @ProcessCheckType, @OspQty, @SuppliedQty, @ProcessCode, @ProcessCodeName, @ExpectedDate, @Remark
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                OspId,
                                MoId,
                                MoProcessId,
                                ProcessCheckStatus,
                                ProcessCheckType,
                                OspQty,
                                SuppliedQty,
                                ProcessCode,
                                ProcessCodeName,
                                ExpectedDate,
                                Remark,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });

                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected = insertResult.Count();

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

        #region //AddOspBarcode -- 新增託外生產條碼 -- Ann 2022-09-07
        public string AddOspBarcode(int OspDetailId, string BarcodeNo, int MoProcessId)
        {
            try
            {
                int BarcodeQty = 1;
                if (BarcodeNo.Length <= 0) throw new SystemException("條碼不能為空!");
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷託外生產單詳細資料是否有誤
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.MoId, a.OspQty, a.SuppliedQty
                                , b.Status
                                FROM MES.OspDetail a
                                INNER JOIN MES.OutsourcingProduction b ON a.OspId = b.OspId
                                WHERE a.OspDetailId = @OspDetailId";
                        dynamicParameters.Add("OspDetailId", OspDetailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("託外生產單詳細資料有誤!");

                        int MoId = -1;
                        int OspQty = -1;
                        int SuppliedQty = -1;
                        foreach (var item in result)
                        {
                            if (item.Status != "N" && item.Status != "E") throw new SystemException("託外生產單已確認，無法更改!");
                            MoId = item.MoId;
                            OspQty = item.OspQty;
                            SuppliedQty = item.SuppliedQty;
                        }
                        #endregion

                        #region //判斷MES製令製程資料是否有誤
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.MoProcess
                                WHERE MoProcessId = @MoProcessId";
                        dynamicParameters.Add("MoProcessId", MoProcessId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("MES製令製程資料有誤!");
                        #endregion

                        var barcodeList = BarcodeNo.Split(',');
                        foreach (var barcodeNo in barcodeList)
                        {
                            #region //判斷 託外生產單詳細資料 + 條碼 是否重複
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.OspBarcode
                                    WHERE OspDetailId = @OspDetailId
                                    AND BarcodeNo = @BarcodeNo";
                            dynamicParameters.Add("OspDetailId", OspDetailId);
                            dynamicParameters.Add("BarcodeNo", barcodeNo);

                            var result3 = sqlConnection.Query(sql, dynamicParameters);
                            if (result3.Count() > 0) throw new SystemException("【託外生產單詳細資料 + 條碼】重複，請重新輸入!");
                            #endregion

                            #region //判斷 託外條碼有沒有在該製令上
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT 1
                                  FROM MES.MoProcess a
                                 WHERE a.MoProcessId = @MoProcessId
                                   AND ((a.SortNumber = 1) or
                                        EXISTS (SELECT 1
		                                          FROM MES.Barcode b
		                                         WHERE a.MoId = b.MoId
				                                   AND b.BarcodeNo = @BarcodeNo))";
                            dynamicParameters.Add("MoProcessId", MoProcessId);
                            dynamicParameters.Add("BarcodeNo", barcodeNo);

                            var result4 = sqlConnection.Query(sql, dynamicParameters);
                            if (result4.Count() <= 0) throw new SystemException("條碼不在此製令, 不可託外!");
                            #endregion

                            #region //判斷是否已有相同製程、條碼存在於託外生產單中，且尚未被託外入庫單註冊並過站
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 c.OspNo
                                    FROM MES.OspBarcode a 
                                    INNER JOIN MES.OspDetail b ON a.OspDetailId = b.OspDetailId
                                    INNER JOIN MES.OutsourcingProduction c ON b.OspId = c.OspId
                                    WHERE a.BarcodeNo = @BarcodeNo 
                                    AND b.MoProcessId = @MoProcessId 
                                    AND NOT EXISTS (
                                        SELECT TOP 1 1 FROM MES.OspReceiptBarcode x
                                        INNER JOIN MES.OspReceiptDetail xa ON x.OsrDetailId = xa.OsrDetailId 
                                        WHERE x.OspBarcodeId = a.OspBarcodeId
                                        AND xa.ProcessStatus = 'Y'
                                    )";
                            dynamicParameters.Add("BarcodeNo", barcodeNo);
                            dynamicParameters.Add("MoProcessId", MoProcessId);

                            var OspBarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item in OspBarcodeResult)
                            {
                                throw new SystemException("已存在相同製程、條碼的託外生產單【" + item.OspNo + "】，且尚未完成託外入庫程序!!");
                            }
                            #endregion

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.OspBarcode (OspDetailId, BarcodeNo, BarcodeQty
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.OspBarcodeId
                                    VALUES (@OspDetailId, @BarcodeNo, @BarcodeQty
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    OspDetailId,
                                    barcodeNo,
                                    BarcodeQty = 1,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            int rowsAffected = insertResult.Count();

                            #region //同步更新OspDetail託外生產量、供貨量
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE MES.OspDetail SET
                                    OspQty = OspQty + @OspQty,
                                    SuppliedQty = SuppliedQty + @SuppliedQty,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE OspDetailId = @OspDetailId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    OspQty = BarcodeQty,
                                    SuppliedQty = BarcodeQty,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    OspDetailId
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //AddOspReceipt -- 新增託外入庫資料 -- Ann 2022-09-15
        public string AddOspReceipt(int SupplierId, string SupplierSo, string OsrErpPrefix, string OsrErpNo, string DocumentDate, string ReceiptDate, string ReserveTaxCode, string TaxCode
            , string TaxType, string InvoiceType, double TaxRate, string UiNo, string InvoiceDate, string InvoiceNo, string ApplyYYMM, string CurrencyCode, double Exchange, int RowCnt
            , string Remark, string DeductType, string PaymentTerm)
        {
            try
            {
                string FactoryCode = "";
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (OsrErpPrefix.Length <= 0) throw new SystemException("【ERP託外入庫單別】不能為空!");
                        if (OsrErpPrefix != "5902") throw new SystemException("【ERP託外入庫單別】錯誤!");
                        if (DocumentDate.Length <= 0) throw new SystemException("【單據日期】不能為空!");
                        if (ReceiptDate.Length <= 0) throw new SystemException("【託外入庫日期】不能為空!");
                        if (ReserveTaxCode.Length <= 0) throw new SystemException("【保稅碼】不能為空!");
                        //if (TaxRate <= 0) throw new SystemException("【營業稅率】不能為空!");
                        if (ApplyYYMM.Length <= 0) throw new SystemException("【申報日期】不能為空!");
                        //if (ApplyYYMM.Length > 6) throw new SystemException("【申報日期】長度錯誤!");
                        if (Exchange <= 0) throw new SystemException("【匯率】不能為空!");
                        if (RowCnt < 0) throw new SystemException("【件數】不能為空!");
                        if (DeductType.Length <= 0) throw new SystemException("【扣抵區分】不能為空!");

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

                        #region //判斷供應商資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.Supplier
                                WHERE SupplierId = @SupplierId";
                        dynamicParameters.Add("SupplierId", SupplierId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("供應商資料有誤!");
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //判斷稅別碼資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM CMSNN
                                    WHERE NN001 = @NN001";
                            dynamicParameters.Add("NN001", TaxCode);

                            var result2 = sqlConnection2.Query(sql, dynamicParameters);
                            if (result2.Count() <= 0) throw new SystemException("稅別碼資料有誤!");
                            #endregion

                            #region //判斷發票聯數資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM CMSNM
                                    WHERE NM001 = @NM001";
                            dynamicParameters.Add("NM001", InvoiceType);

                            var result3 = sqlConnection2.Query(sql, dynamicParameters);
                            if (result3.Count() <= 0) throw new SystemException("發票聯數資料有誤!");
                            #endregion

                            #region //判斷幣別資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM CMSMF
                                    WHERE MF001 = @MF001";
                            dynamicParameters.Add("MF001", CurrencyCode);

                            var result4 = sqlConnection2.Query(sql, dynamicParameters);
                            if (result4.Count() <= 0) throw new SystemException("幣別資料有誤!");
                            #endregion

                            #region //判斷付款條件資料是否正確
                            if (PaymentTerm != "")
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                    FROM CMSNA
                                    WHERE NA002 = @NA002";
                                dynamicParameters.Add("NA002", PaymentTerm);

                                var result5 = sqlConnection2.Query(sql, dynamicParameters);
                                if (result5.Count() <= 0) throw new SystemException("付款條件資料有誤!");
                            }
                            #endregion

                            #region //查詢廠別
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MB001)) MB001 FROM CMSMB";
                            var CMSMBResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (CMSMBResult.Count() > 1) throw new SystemException("廠別數有多個，請與資訊人員確認!!");

                            foreach (var item in CMSMBResult)
                            {
                                FactoryCode = item.MB001;
                            }
                            #endregion
                        }

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.OspReceipt (CompanyId, OsrErpPrefix, OsrErpNo, ReceiptDate, FactoryCode, SupplierId, SupplierSo, CurrencyCode
                                , Exchange, RowCnt, Remark, UiNo, InvoiceType, InvoiceDate, InvoiceNo, TaxType, DeductType, ReserveFlag, OrigAmount
                                , DeductAmount, OrigTax, ReceiptAmount, Quantity, ConfirmStatus, RenewFlag, PrintCnt, AutoMaterialBilling, OrigPreTaxAmount
                                , ApplyYYMM, DocumentDate, TaxRate, PretaxAmount, TaxAmount, PaymentTerm, PackageQuantity, SupplierPicking
                                , ApproveStatus, ReserveTaxCode, SendCount, NoticeFlag, TaxCode, TaxExchange, QcFlag, FlowStatus
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.OsrId, INSERTED.CreateBy, INSERTED.CompanyId
                                VALUES (@CompanyId, @OsrErpPrefix, @OsrErpNo, @ReceiptDate, @FactoryCode, @SupplierId, @SupplierSo, @CurrencyCode
                                , @Exchange, @RowCnt, @Remark, @UiNo, @InvoiceType, @InvoiceDate, @InvoiceNo, @TaxType, @DeductType, @ReserveFlag, @OrigAmount
                                , @DeductAmount, @OrigTax, @ReceiptAmount, @Quantity, @ConfirmStatus, @RenewFlag, @PrintCnt, @AutoMaterialBilling, @OrigPreTaxAmount
                                , @ApplyYYMM, @DocumentDate, @TaxRate, @PretaxAmount, @TaxAmount, @PaymentTerm, @PackageQuantity, @SupplierPicking
                                , @ApproveStatus, @ReserveTaxCode, @SendCount, @NoticeFlag, @TaxCode, @TaxExchange, @QcFlag, @FlowStatus
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                OsrErpPrefix,
                                OsrErpNo = BaseHelper.RandomCode(11),
                                ReceiptDate,
                                FactoryCode,
                                SupplierId,
                                SupplierSo,
                                CurrencyCode,
                                Exchange,
                                RowCnt,
                                Remark,
                                UiNo,
                                InvoiceType,
                                InvoiceDate,
                                InvoiceNo,
                                TaxType,
                                DeductType,
                                ReserveFlag = "N",
                                OrigAmount = 0.0000,
                                DeductAmount = 0.0000,
                                OrigTax = 0.0000,
                                ReceiptAmount = 0.0000,
                                Quantity = 0,
                                ConfirmStatus = "N",
                                RenewFlag = "N",
                                PrintCnt = 0,
                                AutoMaterialBilling = "N",
                                OrigPreTaxAmount = 0,
                                ApplyYYMM,
                                DocumentDate,
                                TaxRate,
                                PretaxAmount = 0,
                                TaxAmount = 0,
                                PaymentTerm,
                                PackageQuantity = 0,
                                SupplierPicking = "N",
                                ApproveStatus = "0",
                                ReserveTaxCode,
                                SendCount = 0,
                                NoticeFlag = "N",
                                TaxCode,
                                TaxExchange = 1,
                                QcFlag = "N",
                                FlowStatus = "1",
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });

                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = insertResult.Count();

                        foreach (var item in insertResult)
                        {
                            if (BaseHelper.CheckUserAuthority(item.CreateBy, item.CompanyId, "A", "OspReceipt", "add", sqlConnection).Equals("N")) throw new SystemException("人員權限檢核有異常，請嘗試重新登入!");
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

        #region //AddOspReceiptDetail -- 新增託外入庫詳細資料 -- Ann 2022-09-19
        public string AddOspReceiptDetail(int OsrId, int OspDetailId, string OsrSeq, string LotNumber, int InventoryId, string AcceptanceDate
            , int ReceiptQty, string AvailableUom, int AcceptQty, int ScriptQty, int ReturnQty, int AvailableQty, double OrigUnitPrice, double OrigAmount
            , string DiscountDescription, string Overdue, string Remark, double OrigPreTaxAmt, double OrigTaxAmt, double OrigDiscountAmt, double PreTaxAmt, double TaxAmt)
        {
            try
            {
                if (LotNumber.Length > 20) throw new SystemException("【批號】長度錯誤!");
                if (AcceptanceDate.Length <= 0) throw new SystemException("【驗收日期】不能為空!");
                if (ReceiptQty < 0) throw new SystemException("【進貨數量】不能為空!");
                if (AcceptQty < 0) throw new SystemException("【驗收數量】不能為空!");
                if (ScriptQty < 0) throw new SystemException("【報廢數量】不能為空!");
                if (ReturnQty < 0) throw new SystemException("【驗退數量】不能為空!");
                if (AvailableQty < 0) throw new SystemException("【計價數量】不能為空!");
                if (OrigUnitPrice < 0) throw new SystemException("【原幣加工單價】不能為空!");
                if (OrigAmount < 0) throw new SystemException("【原幣加工金額】不能為空!");
                if (DiscountDescription.Length > 40) throw new SystemException("【扣款說明】長度錯誤!");
                if (Overdue.Length <= 0) throw new SystemException("【逾期碼】不能為空!");
                if (Remark.Length > 255) throw new SystemException("【備註】長度錯誤!");
                if (OrigDiscountAmt < 0) throw new SystemException("【原幣扣款金額】不能為空!");
                if (OrigPreTaxAmt < 0) throw new SystemException("【原幣未稅金額】不能為空!");
                if (OrigTaxAmt < 0) throw new SystemException("【原幣稅金額】不能為空!");
                if (PreTaxAmt < 0) throw new SystemException("【本幣未稅金額】不能為空!");
                if (TaxAmt < 0) throw new SystemException("【本幣稅金額】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷託外入庫資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT OsrErpPrefix, OsrErpNo, ReceiptDate, ReserveTaxCode, ConfirmStatus
                                FROM MES.OspReceipt
                                WHERE OsrId = @OsrId";
                        dynamicParameters.Add("OsrId", OsrId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("託外入庫資料有誤!");

                        string OsrErpPrefix = "";
                        string OsrErpNo = "";
                        string ReceiptDate = new DateTime().ToString("yyyy-MM-dd");
                        string ReserveTaxCode = "";
                        foreach (var item in result)
                        {
                            if (item.ConfirmStatus != "N") throw new SystemException("託外入庫單已確認，無法更改!");
                            OsrErpPrefix = item.OsrErpPrefix;
                            OsrErpNo = item.OsrErpNo;
                            ReceiptDate = item.ReceiptDate.ToString("yyyy-MM-dd");
                            ReserveTaxCode = item.ReserveTaxCode;
                        }
                        #endregion

                        #region //確認此托外入庫單是否有其他單身狀態已過站
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.OspReceiptDetail a 
                                WHERE a.OsrId = @OsrId
                                AND a.ProcessStatus = 'Y'";
                        dynamicParameters.Add("OsrId", OsrId);

                        var ProcessStatusResult = sqlConnection.Query(sql, dynamicParameters);

                        if (ProcessStatusResult.Count() > 0)
                        {
                            #region //確認過站模式是否為供應商過站
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.OspReceiptBarcode x 
                                    INNER JOIN MES.OspBarcode xa ON x.OspBarcodeId = xa.OspBarcodeId
                                    INNER JOIN MES.Barcode xb ON xa.BarcodeNo = xb.BarcodeNo
                                    INNER JOIN MES.BarcodeProcess xc ON xb.BarcodeId = xc.BarcodeId
                                    INNER JOIN MES.OspDetail xd ON xa.OspDetailId = xd.OspDetailId
                                    INNER JOIN BAS.[User] xe ON xc.FinishUserId = xe.UserId
                                    INNER JOIN MES.OspReceiptDetail xf ON x.OsrDetailId = xf.OsrDetailId
                                    WHERE xf.OsrId = @OsrId
                                    AND xc.MoProcessId = xd.MoProcessId
                                    AND xe.UserNo = 'Z100'";
                            dynamicParameters.Add("OsrId", OsrId);

                            var OspReceiptBarcodeResult3 = sqlConnection.Query(sql, dynamicParameters);

                            if (OspReceiptBarcodeResult3.Count() <= 0) throw new SystemException("此託外入庫單已有其他單身資料已過站，無法繼續新增!!");
                            #endregion
                        }
                        #endregion

                        #region //檢查此託外生產單是否已經被其他單據綁定
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.OspReceiptDetail a 
                                WHERE a.OspDetailId = @OspDetailId";
                        dynamicParameters.Add("OspDetailId", OspDetailId);

                        var OspReceiptDetailResult = sqlConnection.Query(sql, dynamicParameters);

                        //if (OspReceiptDetailResult.Count() > 0) throw new SystemException("此託外生產單已被其他託外入庫單綁定!!");
                        #endregion

                        #region //判斷託外生產詳細資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MoId, a.MoProcessId, a.ProcessCode, a.OspQty
                                , c.MtlItemId, c.UomId
                                FROM MES.OspDetail a
                                INNER JOIN MES.ManufactureOrder b ON a.MoId = b.MoId
                                INNER JOIN MES.WipOrder c ON b.WoId = c.WoId
                                WHERE a.OspDetailId = @OspDetailId";
                        dynamicParameters.Add("OspDetailId", OspDetailId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("託外生產詳細資料有誤!");

                        int MtlItemId = -1;
                        int UomId = -1;
                        int MoId = -1;
                        int MoProcessId = -1;
                        string ProcessCode = "";
                        int OspQty = -1;
                        foreach (var item in result2)
                        {
                            MtlItemId = item.MtlItemId;
                            UomId = item.UomId;
                            MoId = item.MoId;
                            MoProcessId = item.MoProcessId;
                            ProcessCode = item.ProcessCode;
                            OspQty = item.OspQty;
                        }
                        #endregion

                        #region //判斷庫別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.Inventory
                                WHERE InventoryId = @InventoryId";
                        dynamicParameters.Add("InventoryId", InventoryId);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() <= 0) throw new SystemException("庫別資料有誤!");
                        #endregion

                        #region //判斷計價單位資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.UnitOfMeasure
                                WHERE UomId = @UomId";
                        dynamicParameters.Add("UomId", AvailableUom);

                        var result4 = sqlConnection.Query(sql, dynamicParameters);
                        if (result4.Count() <= 0) throw new SystemException("計價單位資料有誤!");
                        #endregion

                        #region //取得目前入庫序號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT COUNT(1) Count
                                FROM MES.OspReceiptDetail a
                                WHERE a.OsrId = @OsrId";
                        dynamicParameters.Add("OsrId", OsrId);

                        var result5 = sqlConnection.Query(sql, dynamicParameters);
                        if (result5.Count() <= 0) throw new SystemException("取得託外入庫序號時發生錯誤!!");

                        foreach (var item in result5)
                        {
                            OsrSeq = (item.Count + 1).ToString().PadLeft(4, '0');
                        }
                        #endregion

                        #region //檢查【託外入庫單 + 托外生產詳細資料】是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.OspReceiptDetail a
                                WHERE a.OsrId = @OsrId
                                AND a.OspDetailId = @OspDetailId";
                        dynamicParameters.Add("OsrId", OsrId);
                        dynamicParameters.Add("OspDetailId", OspDetailId);

                        var result6 = sqlConnection.Query(sql, dynamicParameters);
                        if (result6.Count() > 0) throw new SystemException("【託外入庫單 + 托外生產詳細資料】重複!");
                        #endregion

                        #region //檢查【託外入庫單 + 序號】是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.OspReceiptDetail a
                                WHERE a.OsrId = @OsrId
                                AND a.OsrSeq = @OsrSeq";
                        dynamicParameters.Add("OsrId", OsrId);
                        dynamicParameters.Add("OsrSeq", OsrSeq);

                        var result7 = sqlConnection.Query(sql, dynamicParameters);
                        if (result7.Count() > 0) throw new SystemException("【託外入庫單 + 序號】重複!");
                        #endregion

                        #region //INSERT MES.OspReceiptDetail
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.OspReceiptDetail (OsrId, OspDetailId, OsrErpPrefix, OsrErpNo, OsrSeq, MtlItemId, ReceiptQty, UomId, InventoryId
                                , LotNumber, AvailableDate, ReCheckDate, MoId, MoProcessId, ProcessCode, ReceiptPackageQty, AcceptancePackageQty, AcceptanceDate
                                , AcceptQty, AvailableQty, ScriptQty, ReturnQty, AvailableUom, OrigUnitPrice, OrigAmount, OrigDiscountAmt, ReceiptExpense
                                , DiscountDescription, ProjectCode, PaymentHold, Overdue, QcStatus, ReturnCode, ConfirmStatus, CloseStatus, ReNewStatus
                                , Remark, CostEntry, ExpenseEntry, ConfirmUserId, ConfirmUser, OrigPreTaxAmt, OrigTaxAmt, PreTaxAmt, TaxAmt, UrgentMtl, ReserveTaxCode
                                , ApproveStatus, TransferStatus, ProcessStatus
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.OsrDetailId
                                VALUES (@OsrId, @OspDetailId, @OsrErpPrefix, @OsrErpNo, @OsrSeq, @MtlItemId, @ReceiptQty, @UomId, @InventoryId
                                , @LotNumber, @AvailableDate, @ReCheckDate, @MoId, @MoProcessId, @ProcessCode, @ReceiptPackageQty, @AcceptancePackageQty, @AcceptanceDate
                                , @AcceptQty, @AvailableQty, @ScriptQty, @ReturnQty, @AvailableUom, @OrigUnitPrice, @OrigAmount, @OrigDiscountAmt, @ReceiptExpense
                                , @DiscountDescription, @ProjectCode, @PaymentHold, @Overdue, @QcStatus, @ReturnCode, @ConfirmStatus, @CloseStatus, @ReNewStatus
                                , @Remark, @CostEntry, @ExpenseEntry, @ConfirmUserId, @ConfirmUser, @OrigPreTaxAmt, @OrigTaxAmt, @PreTaxAmt, @TaxAmt, @UrgentMtl, @ReserveTaxCode
                                , @ApproveStatus, @TransferStatus, @ProcessStatus
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                OsrId,
                                OspDetailId,
                                OsrErpPrefix,
                                OsrErpNo,
                                OsrSeq,
                                MtlItemId,
                                ReceiptQty,
                                UomId,
                                InventoryId,
                                LotNumber,
                                AvailableDate = ReceiptDate,
                                ReCheckDate = ReceiptDate,
                                MoId,
                                MoProcessId,
                                ProcessCode,
                                ReceiptPackageQty = 0,
                                AcceptancePackageQty = 0,
                                AcceptanceDate,
                                AcceptQty,
                                AvailableQty,
                                ScriptQty,
                                ReturnQty,
                                AvailableUom,
                                OrigUnitPrice,
                                OrigAmount,
                                OrigDiscountAmt,
                                ReceiptExpense = 0,
                                DiscountDescription,
                                ProjectCode = "",
                                PaymentHold = "N",
                                Overdue,
                                QcStatus = "2",
                                ReturnCode = "N",
                                ConfirmStatus = "N",
                                CloseStatus = "N",
                                ReNewStatus = "N",
                                Remark,
                                CostEntry = "N",
                                ExpenseEntry = "N",
                                ConfirmUserId = (int?)null,
                                ConfirmUser = "",
                                OrigPreTaxAmt,
                                OrigTaxAmt,
                                PreTaxAmt,
                                TaxAmt,
                                UrgentMtl = "N",
                                ReserveTaxCode,
                                ApproveStatus = "N",
                                TransferStatus = "N",
                                ProcessStatus = "N",
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });

                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = insertResult.Count();
                        #endregion

                        #region //更新託外入庫單頭資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.OspReceipt SET
                                Quantity = Quantity + @AcceptQty,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE OsrId = @OsrId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                OrigAmount,
                                OrigTaxAmt,
                                AcceptQty,
                                PreTaxAmt,
                                TaxAmt,
                                LastModifiedDate,
                                LastModifiedBy,
                                OsrId
                            });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //AddOspReceiptBarcode-- 新增託外入庫條碼資料 -- Ann 2022-09-19
        public string AddOspReceiptBarcode(string OsrReceiptBarcodeData)
        {
            try
            {
                if (OsrReceiptBarcodeData.Length <= 0) throw new SystemException("【託外入庫條碼】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();
                        var osrReceiptBarcodeJson = JObject.Parse(OsrReceiptBarcodeData);
                        int rowsAffected = 0;
                        int OsrId = -1;
                        foreach (var item in osrReceiptBarcodeJson["osrReceiptBarcodeInfo"])
                        {
                            string BarcodeNo = "";

                            #region //判斷託外入庫詳細資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.OsrId, a.ReceiptQty, a.AcceptQty, a.AvailableQty
                                    , a.ConfirmStatus, a.ProcessStatus, a.MoProcessId, a.OspDetailId
                                    , c.PassStationControl
                                    , d.MoId
                                    , e.CompanyId
                                    , f.SupplierNo
                                    FROM MES.OspReceiptDetail a 
                                    INNER JOIN MES.OspReceipt b ON a.OsrId = b.OsrId
                                    INNER JOIN SCM.Supplier c ON b.SupplierId = c.SupplierId
                                    INNER JOIN MES.OspDetail d ON a.OspDetailId = d.OspDetailId
                                    INNER JOIN MES.OutsourcingProduction e ON d.OspId = e.OspId
                                    INNER JOIN SCM.Supplier f ON e.SupplierId = f.SupplierId
                                    WHERE a.OsrDetailId = @OsrDetailId";
                            dynamicParameters.Add("OsrDetailId", Convert.ToInt32(item["OsrDetailId"]));

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("託外入庫詳細資料錯誤!");

                            int ReceiptQty = -1;
                            int AcceptQty = -1;
                            int AvailableQty = -1;
                            int MoProcessId = -1;
                            int OspDetailId = -1;
                            string PassStationControl = "";
                            int CompanyId = -1;
                            string SupplierNo = "";
                            int MoId = -1;
                            foreach (var item2 in result)
                            {
                                if (item2.ProcessStatus != "N" && item2.PassStationControl != "Y")
                                {
                                    throw new SystemException("託外入庫碼已過站，無法更改!");
                                }

                                if (item2.ConfirmStatus != "N") throw new SystemException("託外入庫單已確認，無法更改!");

                                ReceiptQty = item2.ReceiptQty;
                                AcceptQty = item2.AcceptQty;
                                AvailableQty = item2.AvailableQty;
                                MoProcessId = item2.MoProcessId;
                                OsrId = item2.OsrId;
                                OspDetailId = item2.OspDetailId;
                                PassStationControl = item2.PassStationControl;
                                CompanyId = item2.CompanyId;
                                SupplierNo = item2.SupplierNo;
                                MoId = item2.MoId;
                            }
                            #endregion

                            #region //確認條碼資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.BarcodeStatus, a.CurrentProdStatus, a.BarcodeNo
                                    FROM MES.Barcode a 
                                    WHERE a.BarcodeNo = @BarcodeNo";
                            dynamicParameters.Add("BarcodeNo", item["BarcodeNo"].ToString());

                            var BarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                            if (BarcodeResult.Count() <= 0) throw new SystemException("條碼【" + item["BarcodeNo"].ToString() + "】資料錯誤!!");

                            string BarcodeStatus = "";
                            foreach (var item2 in BarcodeResult)
                            {
                                if (item2.BarcodeStatus != "0" && item2.BarcodeStatus != "1")
                                {
                                    if (!(CurrentCompany == 4 && SupplierNo == "200B0002" && item2.BarcodeStatus == "10"))
                                    {
                                        throw new SystemException("條碼【" + item["BarcodeNo"].ToString() + "】狀態錯誤!!");
                                    }
                                }
                                BarcodeNo = item2.BarcodeNo;
                                BarcodeStatus = item2.BarcodeStatus;
                            }
                            #endregion

                            #region //判斷託外生產條碼資料是否正確
                            //先判斷此製程是否為首站
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT SortNumber
                                    FROM MES.MoProcess
                                    WHERE MoProcessId = @MoProcessId";
                            dynamicParameters.Add("MoProcessId", MoProcessId);

                            var sortNumberResult = sqlConnection.Query(sql, dynamicParameters);

                            if (sortNumberResult.Count() <= 0) throw new SystemException("製令製程資料錯誤!");

                            foreach (var item2 in sortNumberResult)
                            {
                                if (item2.SortNumber != 1)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 BarcodeNo
                                            FROM MES.OspBarcode
                                            WHERE OspBarcodeId = @OspBarcodeId";
                                    dynamicParameters.Add("OspBarcodeId", Convert.ToInt32(item["OspBarcodeId"]));

                                    var result2 = sqlConnection.Query(sql, dynamicParameters);
                                    if (result2.Count() <= 0) throw new SystemException("託外生產條碼資料錯誤!");

                                    foreach (var item3 in result2)
                                    {
                                        #region //確認條碼資料是否正確
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT TOP 1 1
                                                FROM MES.Barcode
                                                WHERE BarcodeNo = @BarcodeNo";
                                        dynamicParameters.Add("BarcodeNo", item3.BarcodeNo);

                                        var result4 = sqlConnection.Query(sql, dynamicParameters);
                                        if (result4.Count() <= 0) throw new SystemException("條碼【" + item3.BarcodeNo + "】資料錯誤!");

                                        BarcodeNo = item3.BarcodeNo;
                                        #endregion
                                    }
                                }
                            }
                            #endregion

                            #region //檢查是否已經有這筆託外生產條碼
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.OspReceiptBarcode
                                    WHERE OspBarcodeId = @OspBarcodeId
                                    AND OsrDetailId = @OsrDetailId";
                            dynamicParameters.Add("OspBarcodeId", Convert.ToInt32(item["OspBarcodeId"]));
                            dynamicParameters.Add("OsrDetailId", Convert.ToInt32(item["OsrDetailId"]));

                            var result3 = sqlConnection.Query(sql, dynamicParameters);
                            if (result3.Count() > 0) continue;
                            #endregion

                            #region //檢查該條碼是否需要做工程檢且尚未完成，或是工程檢結果非良品
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.ProcessCheckStatus, a.ProcessCheckType
                                    FROM MES.MoProcess a
                                    WHERE MoProcessId = @MoProcessId";
                            dynamicParameters.Add("MoProcessId", MoProcessId);

                            var processCheckStatusResult = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item2 in processCheckStatusResult)
                            {
                                if (item2.ProcessCheckStatus == "Y")
                                {
                                    if (item2.ProcessCheckType == "1") //全檢
                                    {
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT TOP 1 a.QcStatus
                                                FROM MES.QcBarcode a
                                                INNER JOIN MES.Barcode b ON a.BarcodeId = b.BarcodeId
                                                INNER JOIN MES.QcRecord c ON a.QcRecordId = c.QcRecordId
                                                WHERE b.BarcodeNo = @BarcodeNo
                                                AND c.MoProcessId = b.CurrentMoProcessId
                                                ORDER BY a.CreateDate DESC";
                                        dynamicParameters.Add("BarcodeNo", BarcodeNo);

                                        var qcStatusResult = sqlConnection.Query(sql, dynamicParameters);
                                        if (qcStatusResult.Count() <= 0) throw new SystemException("該條碼查無工程檢紀錄!");

                                        foreach (var item3 in qcStatusResult)
                                        {
                                            if (item3.QcStatus != "P") throw new SystemException("該條碼工程檢結果非良品且尚未修正!");
                                        }
                                    }
                                    else if (item2.ProcessCheckType == "2") //抽檢
                                    {
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT TOP 1 1
                                                FROM MES.QcBarcode a
                                                INNER JOIN MES.QcRecord b ON a.QcRecordId = b.QcRecordId
                                                INNER JOIN MES.Barcode c ON a.BarcodeId = c.BarcodeId
                                                WHERE b.MoProcessId = c.CurrentMoProcessId AND a.QcStatus = 'P'";
                                        dynamicParameters.Add("MoProcessId", MoProcessId);

                                        var qcBarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                                        if (qcBarcodeResult.Count() <= 0) throw new SystemException("該製令製程查無任何一筆工程檢紀錄!");
                                    }
                                }
                            }
                            #endregion

                            #region //檢查是否有其他張相同託外生產單單身及條碼
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 d.BarcodeNo
                                    FROM MES.OspReceiptBarcode a 
                                    INNER JOIN MES.OspReceiptDetail b ON a.OsrDetailId = b.OsrDetailId
                                    INNER JOIN MES.OspBarcode c ON a.OspBarcodeId = c.OspBarcodeId
                                    INNER JOIN MES.Barcode d ON c.BarcodeNo = d.BarcodeNo
                                    WHERE a.OspBarcodeId = @OspBarcodeId
                                    AND b.OspDetailId = @OspDetailId";
                            dynamicParameters.Add("OspBarcodeId", Convert.ToInt32(item["OspBarcodeId"]));
                            dynamicParameters.Add("OspDetailId", OspDetailId);

                            var OspReceiptBarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                            if (OspReceiptBarcodeResult.Count() > 0)
                            {
                                foreach (var item2 in OspReceiptBarcodeResult)
                                {
                                    throw new SystemException("此條碼【" + item2.BarcodeNo + "】及此託外生產單據已有建立紀錄，無法重複建立!!");
                                }
                            }
                            #endregion

                            #region //若為供應商過站模式，確認條碼已過站
                            if (PassStationControl == "Y")
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.FinishDate
                                        FROM MES.BarcodeProcess a 
                                        INNER JOIN MES.Barcode b ON a.BarcodeId = b.BarcodeId
                                        WHERE b.BarcodeNo = @BarcodeNo
                                        AND a.MoProcessId = @MoProcessId";
                                dynamicParameters.Add("BarcodeNo", BarcodeNo);
                                dynamicParameters.Add("MoProcessId", MoProcessId);

                                var BarcodeProcessResult = sqlConnection.Query(sql, dynamicParameters);

                                //if (BarcodeProcessResult.Count() <= 0) throw new SystemException("條碼【" + item["BarcodeNo"].ToString() + "】尚未過站，<br>請通知供應商完成過站程序!!");

                                foreach (var item2 in BarcodeProcessResult)
                                {
                                    if (item2.FinishDate == null) throw new SystemException("條碼【" + item["BarcodeNo"].ToString() + "】尚在開工中，<br>請通知供應商完成過站程序!!");
                                }
                            }
                            #endregion

                            #region //INSERT MES.OspReceiptBarcode
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.OspReceiptBarcode (OsrDetailId, OspBarcodeId, Status
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@OsrDetailId, @OspBarcodeId, @Status
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    OsrDetailId = Convert.ToInt32(item["OsrDetailId"]),
                                    OspBarcodeId = Convert.ToInt32(item["OspBarcodeId"]),
                                    Status = "I",
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();
                            #endregion

                            #region //確認目前此單身條碼是否已都過站，且為供應商過站模式，若已過站則更改單據狀態
                            string ProcessStatus = "N";
                            if (PassStationControl == "Y")
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT b.BarcodeNo
                                        , (
                                            SELECT x.FinishDate
                                            FROM MES.BarcodeProcess x 
                                            INNER JOIN MES.Barcode xa ON x.BarcodeId = xa.BarcodeId
                                            INNER JOIN BAS.[User] xb ON x.CreateBy = xb.UserId
                                            WHERE xa.BarcodeNo = b.BarcodeNo
                                            AND x.MoProcessId = c.MoProcessId
                                            FOR JSON PATH, ROOT('data')
                                        ) BarcodeProcess
                                        FROM MES.OspReceiptBarcode a 
                                        INNER JOIN MES.OspBarcode b ON a.OspBarcodeId = b.OspBarcodeId
                                        INNER JOIN MES.OspDetail c ON b.OspDetailId = c.OspDetailId
                                        WHERE c.OspDetailId = @OspDetailId";
                                dynamicParameters.Add("OspDetailId", OspDetailId);

                                var OspReceiptBarcodeResult2 = sqlConnection.Query(sql, dynamicParameters);

                                foreach (var item2 in OspReceiptBarcodeResult2)
                                {
                                    if (item2.BarcodeProcess == null)
                                    {
                                        ProcessStatus = "N";
                                        break;
                                    }
                                    else
                                    {
                                        JObject barcodeProcess = JObject.Parse(item2.BarcodeProcess);
                                        bool finishDataFlag = false;
                                        foreach (var item3 in barcodeProcess["data"])
                                        {
                                            if (item3["FinishDate"] == null)
                                            {
                                                finishDataFlag = true;
                                                break;
                                            }
                                        }

                                        if (finishDataFlag == true)
                                        {
                                            ProcessStatus = "N";
                                            break;
                                        }
                                        else
                                        {
                                            #region //確認此條碼是否下一站已指定回此托外站別
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"SELECT TOP 1 1
                                                    FROM MES.Barcode a 
                                                    WHERE a.BarcodeNo = @BarcodeNo
                                                    AND a.NextMoProcessId = @MoProcessId";
                                            dynamicParameters.Add("BarcodeNo", item2.BarcodeNo);
                                            dynamicParameters.Add("MoProcessId", MoProcessId);

                                            var BarcodeCheckResult = sqlConnection.Query(sql, dynamicParameters);

                                            if (BarcodeCheckResult.Count() > 0)
                                            {
                                                ProcessStatus = "N";
                                            }
                                            else
                                            {
                                                ProcessStatus = "Y";
                                            }
                                            #endregion

                                            break;
                                        }
                                    }
                                }
                            }
                            #endregion

                            #region //同步新增OsrDetail入庫數量
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE MES.OspReceiptDetail SET
                                    ReceiptQty = @ReceiptQty,
                                    AcceptQty = @AcceptQty,
                                    AvailableQty = @AvailableQty,
                                    ProcessStatus = @ProcessStatus,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE OsrDetailId = @OsrDetailId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    ReceiptQty = ReceiptQty + 1,
                                    AcceptQty = AcceptQty + 1,
                                    AvailableQty = AvailableQty + 1,
                                    ProcessStatus,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    OsrDetailId = Convert.ToInt32(item["OsrDetailId"])
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //若為委外中揚製作，且條碼為出貨狀態，則改回晶彩相關資料
                            if (CompanyId == 4 && SupplierNo == "200B0002" && BarcodeStatus == "10")
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE MES.Barcode SET
                                        CurrentMoProcessId = @MoProcessId,
                                        NextMoProcessId = @MoProcessId,
                                        MoId = @MoId,
                                        BarcodeStatus = @BarcodeStatus,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE BarcodeNo = @BarcodeNo";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        MoProcessId,
                                        MoId,
                                        BarcodeStatus = "1",
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        BarcodeNo
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            }
                            #endregion
                        }

                        #region //計算目前此託外入庫單總數量
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT SUM(a.AcceptQty) SumAcceptQty
                                FROM MES.OspReceiptDetail a
                                INNER JOIN MES.OspReceipt b ON a.OsrId = b.OsrId
                                WHERE a.OsrId = @OsrId";
                        dynamicParameters.Add("OsrId", OsrId);
                        var SumAcceptQtyResult = sqlConnection.Query(sql, dynamicParameters);

                        int SumAcceptQty = -1;
                        foreach (var item in SumAcceptQtyResult)
                        {
                            SumAcceptQty = item.SumAcceptQty;
                        }
                        #endregion

                        #region //UPDATE MES.OspReceipt Quantity
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.OspReceipt SET
                                Quantity = @SumAcceptQty,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE OsrId = @OsrId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SumAcceptQty,
                                LastModifiedDate,
                                LastModifiedBy,
                                OsrId
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

        #region //BatchAddOspReceiptDetail -- 批量新增託外入庫資料 -- Ann 2024-07-18
        public string BatchAddOspReceiptDetail(int OsrId, string OspDetailIdList)
        {
            try
            {
                if (OspDetailIdList.Length <= 0) throw new SystemException("【託外生產】資料錯誤!");
                int rowsAffected = 0;

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();

                        #region //判斷託外入庫資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.OsrErpPrefix, a.OsrErpNo, a.ReceiptDate, a.ReserveTaxCode, a.ConfirmStatus, a.SupplierId
                                , b.PassStationControl
                                FROM MES.OspReceipt a 
                                INNER JOIN SCM.Supplier b ON a.SupplierId = b.SupplierId
                                WHERE OsrId = @OsrId";
                        dynamicParameters.Add("OsrId", OsrId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("託外入庫資料有誤!");

                        string OsrErpPrefix = "";
                        string OsrErpNo = "";
                        string ReceiptDate = new DateTime().ToString("yyyy-MM-dd");
                        string ReserveTaxCode = "";
                        string PassStationControl = "";
                        foreach (var item in result)
                        {
                            if (item.ConfirmStatus != "N") throw new SystemException("託外入庫單已確認，無法更改!");
                            OsrErpPrefix = item.OsrErpPrefix;
                            OsrErpNo = item.OsrErpNo;
                            ReceiptDate = item.ReceiptDate.ToString("yyyy-MM-dd");
                            ReserveTaxCode = item.ReserveTaxCode;
                            PassStationControl = item.PassStationControl;
                        }
                        #endregion

                        #region //確認此托外入庫單是否有其他單身狀態已過站
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.OspReceiptDetail a 
                                WHERE a.OsrId = @OsrId
                                AND a.ProcessStatus = 'Y'";
                        dynamicParameters.Add("OsrId", OsrId);

                        var ProcessStatusResult = sqlConnection.Query(sql, dynamicParameters);

                        if (ProcessStatusResult.Count() > 0)
                        {
                            #region //確認過站模式是否為供應商過站
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.OspReceiptBarcode x 
                                    INNER JOIN MES.OspBarcode xa ON x.OspBarcodeId = xa.OspBarcodeId
                                    INNER JOIN MES.Barcode xb ON xa.BarcodeNo = xb.BarcodeNo
                                    INNER JOIN MES.BarcodeProcess xc ON xb.BarcodeId = xc.BarcodeId
                                    INNER JOIN MES.OspDetail xd ON xa.OspDetailId = xd.OspDetailId
                                    INNER JOIN BAS.[User] xe ON xc.FinishUserId = xe.UserId
                                    INNER JOIN MES.OspReceiptDetail xf ON x.OsrDetailId = xf.OsrDetailId
                                    WHERE xf.OsrId = @OsrId
                                    AND xc.MoProcessId = xd.MoProcessId
                                    AND xe.UserNo = 'Z100'";
                            dynamicParameters.Add("OsrId", OsrId);

                            var OspReceiptBarcodeResult3 = sqlConnection.Query(sql, dynamicParameters);

                            if (OspReceiptBarcodeResult3.Count() <= 0) throw new SystemException("此託外入庫單已有其他單身資料已過站，無法繼續新增!!");
                            #endregion
                        }
                        #endregion

                        string[] ospDetailIds = OspDetailIdList.Split(',');
                        foreach (var ospDetailId in ospDetailIds)
                        {
                            #region //判斷託外生產詳細資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.MoId, a.MoProcessId, a.ProcessCode, a.OspQty
                                    , b.WoSeq
                                    , c.MtlItemId, c.UomId, c.WoErpPrefix, c.WoErpNo
                                    , d.InventoryId
                                    , e.OspNo, e.CompanyId
                                    , f.ProcessAlias
                                    , g.SupplierNo
                                    FROM MES.OspDetail a
                                    INNER JOIN MES.ManufactureOrder b ON a.MoId = b.MoId
                                    INNER JOIN MES.WipOrder c ON b.WoId = c.WoId
                                    INNER JOIN PDM.MtlItem d ON c.MtlItemId = d.MtlItemId
                                    INNER JOIN MES.OutsourcingProduction e ON a.OspId = e.OspId
                                    INNER JOIN MES.MoProcess f ON a.MoProcessId = f.MoProcessId
                                    INNER JOIN SCM.Supplier g ON e.SupplierId = g.SupplierId
                                    WHERE a.OspDetailId = @OspDetailId";
                            dynamicParameters.Add("OspDetailId", ospDetailId);

                            var OspDetailResult = sqlConnection.Query(sql, dynamicParameters);
                            if (OspDetailResult.Count() <= 0) throw new SystemException("託外生產詳細資料有誤!");

                            int MtlItemId = -1;
                            int UomId = -1;
                            int MoId = -1;
                            int MoProcessId = -1;
                            string ProcessCode = "";
                            int InventoryId = -1;
                            int OspQty = -1;
                            string OspNo = "";
                            string ProcessAlias = "";
                            int CompanyId = -1;
                            string SupplierNo = "";
                            foreach (var item in OspDetailResult)
                            {
                                MtlItemId = item.MtlItemId;
                                UomId = item.UomId;
                                MoId = item.MoId;
                                MoProcessId = item.MoProcessId;
                                ProcessCode = item.ProcessCode;
                                InventoryId = item.InventoryId;
                                OspQty = item.OspQty;
                                OspNo = item.OspNo;
                                ProcessAlias = item.ProcessAlias;
                                CompanyId = item.CompanyId;
                                SupplierNo = item.SupplierNo;
                            }
                            #endregion

                            #region //取得目前入庫序號
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT COUNT(1) Count
                                    FROM MES.OspReceiptDetail a
                                    WHERE a.OsrId = @OsrId";
                            dynamicParameters.Add("OsrId", OsrId);

                            var result5 = sqlConnection.Query(sql, dynamicParameters);
                            if (result5.Count() <= 0) throw new SystemException("取得託外入庫序號時發生錯誤!!");

                            string OsrSeq = "";
                            foreach (var item in result5)
                            {
                                OsrSeq = (item.Count + 1).ToString().PadLeft(4, '0');
                            }
                            #endregion

                            #region //檢查【託外入庫單 + 托外生產詳細資料】是否重複
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.OspReceiptDetail a
                                    WHERE a.OsrId = @OsrId
                                    AND a.OspDetailId = @OspDetailId";
                            dynamicParameters.Add("OsrId", OsrId);
                            dynamicParameters.Add("OspDetailId", ospDetailId);

                            var result6 = sqlConnection.Query(sql, dynamicParameters);
                            if (result6.Count() > 0) throw new SystemException("【託外入庫單 + 托外生產詳細資料】重複!<br>託外生產單號:" + OspNo + "<br>製程:" + ProcessAlias + "");
                            #endregion

                            #region //檢查【託外入庫單 + 序號】是否重複
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.OspReceiptDetail a
                                    WHERE a.OsrId = @OsrId
                                    AND a.OsrSeq = @OsrSeq";
                            dynamicParameters.Add("OsrId", OsrId);
                            dynamicParameters.Add("OsrSeq", OsrSeq);

                            var result7 = sqlConnection.Query(sql, dynamicParameters);
                            if (result7.Count() > 0) throw new SystemException("【託外入庫單 + 序號】重複!<br>託外生產單號:" + OspNo + "<br>製程:" + ProcessAlias + "");
                            #endregion

                            #region //INSERT MES.OspReceiptDetail
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.OspReceiptDetail (OsrId, OspDetailId, OsrErpPrefix, OsrErpNo, OsrSeq, MtlItemId, ReceiptQty, UomId, InventoryId
                                    , LotNumber, AvailableDate, ReCheckDate, MoId, MoProcessId, ProcessCode, ReceiptPackageQty, AcceptancePackageQty, AcceptanceDate
                                    , AcceptQty, AvailableQty, ScriptQty, ReturnQty, AvailableUom, OrigUnitPrice, OrigAmount, OrigDiscountAmt, ReceiptExpense
                                    , DiscountDescription, ProjectCode, PaymentHold, Overdue, QcStatus, ReturnCode, ConfirmStatus, CloseStatus, ReNewStatus
                                    , Remark, CostEntry, ExpenseEntry, ConfirmUserId, ConfirmUser, OrigPreTaxAmt, OrigTaxAmt, PreTaxAmt, TaxAmt, UrgentMtl, ReserveTaxCode
                                    , ApproveStatus, TransferStatus, ProcessStatus
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.OsrDetailId
                                    VALUES (@OsrId, @OspDetailId, @OsrErpPrefix, @OsrErpNo, @OsrSeq, @MtlItemId, @ReceiptQty, @UomId, @InventoryId
                                    , @LotNumber, @AvailableDate, @ReCheckDate, @MoId, @MoProcessId, @ProcessCode, @ReceiptPackageQty, @AcceptancePackageQty, @AcceptanceDate
                                    , @AcceptQty, @AvailableQty, @ScriptQty, @ReturnQty, @AvailableUom, @OrigUnitPrice, @OrigAmount, @OrigDiscountAmt, @ReceiptExpense
                                    , @DiscountDescription, @ProjectCode, @PaymentHold, @Overdue, @QcStatus, @ReturnCode, @ConfirmStatus, @CloseStatus, @ReNewStatus
                                    , @Remark, @CostEntry, @ExpenseEntry, @ConfirmUserId, @ConfirmUser, @OrigPreTaxAmt, @OrigTaxAmt, @PreTaxAmt, @TaxAmt, @UrgentMtl, @ReserveTaxCode
                                    , @ApproveStatus, @TransferStatus, @ProcessStatus
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    OsrId,
                                    OspDetailId = ospDetailId,
                                    OsrErpPrefix,
                                    OsrErpNo,
                                    OsrSeq,
                                    MtlItemId,
                                    ReceiptQty = 0,
                                    UomId,
                                    InventoryId,
                                    LotNumber = "",
                                    AvailableDate = ReceiptDate,
                                    ReCheckDate = ReceiptDate,
                                    MoId,
                                    MoProcessId,
                                    ProcessCode,
                                    ReceiptPackageQty = 0,
                                    AcceptancePackageQty = 0,
                                    AcceptanceDate = DateTime.Now,
                                    AcceptQty = 0,
                                    AvailableQty = 0,
                                    ScriptQty = 0,
                                    ReturnQty = 0,
                                    AvailableUom = UomId,
                                    OrigUnitPrice = 0,
                                    OrigAmount = 0,
                                    OrigDiscountAmt = 0,
                                    ReceiptExpense = 0,
                                    DiscountDescription = "",
                                    ProjectCode = "",
                                    PaymentHold = "N",
                                    Overdue = "N",
                                    QcStatus = "2",
                                    ReturnCode = "N",
                                    ConfirmStatus = "N",
                                    CloseStatus = "N",
                                    ReNewStatus = "N",
                                    Remark = "",
                                    CostEntry = "N",
                                    ExpenseEntry = "N",
                                    ConfirmUserId = (int?)null,
                                    ConfirmUser = "",
                                    OrigPreTaxAmt = 0,
                                    OrigTaxAmt = 0,
                                    PreTaxAmt = 0,
                                    TaxAmt = 0,
                                    UrgentMtl = "N",
                                    ReserveTaxCode,
                                    ApproveStatus = "N",
                                    TransferStatus = "N",
                                    ProcessStatus = "N",
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();

                            int OsrDetailId = -1;
                            foreach (var item in insertResult)
                            {
                                OsrDetailId = item.OsrDetailId;
                            }
                            #endregion

                            #region //註冊託外入庫條碼流程
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.OspBarcodeId, a.BarcodeNo, a.BarcodeQty
                                    FROM MES.OspBarcode a 
                                    WHERE a.OspDetailId = @OspDetailId";
                            dynamicParameters.Add("OspDetailId", ospDetailId);

                            var OspBarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                            if (OspBarcodeResult.Count() <= 0)
                            {
                                #region //確認是否有批量開立流程
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1 
                                        FROM MES.OspList a 
                                        WHERE a.MoProcessId = @MoProcessId
                                        AND a.[Status] = 'Y'";
                                dynamicParameters.Add("MoProcessId", MoProcessId);

                                var OspListResult = sqlConnection.Query(sql, dynamicParameters);

                                if (OspListResult.Count() <= 0)
                                {
                                    throw new SystemException("取得託外條碼資料錯誤!!<br>託外生產單號:" + OspNo + "<br>製程:" + ProcessAlias + "");
                                }
                                else
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE MES.OspReceiptDetail SET
                                            ReceiptQty = @OspQty,
                                            AcceptQty = @OspQty,
                                            AvailableQty = @OspQty,
                                            ProcessStatus = 'Y',
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE OsrDetailId = @OsrDetailId";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            OspQty,
                                            LastModifiedDate,
                                            LastModifiedBy,
                                            OsrDetailId
                                        });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                }
                                #endregion
                            }

                            int OspBarcodeId = -1;
                            string BarcodeNo = "";
                            string BarcodeStatus = "";
                            foreach (var item in OspBarcodeResult)
                            {
                                BarcodeNo = item.BarcodeNo;
                                OspBarcodeId = item.OspBarcodeId;

                                #region //確認條碼資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.BarcodeStatus, a.CurrentProdStatus, a.BarcodeNo
                                        FROM MES.Barcode a 
                                        WHERE a.BarcodeNo = @BarcodeNo";
                                dynamicParameters.Add("BarcodeNo", BarcodeNo);

                                var BarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                                if (BarcodeResult.Count() <= 0) throw new SystemException("條碼【" + BarcodeNo + "】資料錯誤!!<br>託外生產單號:" + OspNo + "<br>製程:" + ProcessAlias + "");

                                foreach (var item2 in BarcodeResult)
                                {
                                    if (item2.BarcodeStatus != "0" && item2.BarcodeStatus != "1")
                                    {
                                        if (!(CurrentCompany == 4 && SupplierNo == "200B0002" && item2.BarcodeStatus == "10"))
                                        {
                                            throw new SystemException("條碼【" + item2.BarcodeNo + "】狀態錯誤!!");
                                        }
                                    }
                                    BarcodeNo = item2.BarcodeNo;
                                    BarcodeStatus = item2.BarcodeStatus;
                                }
                                #endregion

                                #region //判斷託外生產條碼資料是否正確
                                //先判斷此製程是否為首站
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT SortNumber
                                        FROM MES.MoProcess
                                        WHERE MoProcessId = @MoProcessId";
                                dynamicParameters.Add("MoProcessId", MoProcessId);

                                var sortNumberResult = sqlConnection.Query(sql, dynamicParameters);

                                if (sortNumberResult.Count() <= 0) throw new SystemException("製令製程資料錯誤!<br>託外生產單號:" + OspNo + "<br>製程:" + ProcessAlias + "");

                                foreach (var item2 in sortNumberResult)
                                {
                                    if (item2.SortNumber != 1)
                                    {
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT TOP 1 BarcodeNo
                                                FROM MES.OspBarcode
                                                WHERE OspBarcodeId = @OspBarcodeId";
                                        dynamicParameters.Add("OspBarcodeId", OspBarcodeId);

                                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                                        if (result2.Count() <= 0) throw new SystemException("託外生產條碼資料錯誤!<br>託外生產單號:" + OspNo + "<br>製程:" + ProcessAlias + "");

                                        foreach (var item3 in result2)
                                        {
                                            #region //確認條碼資料是否正確
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"SELECT TOP 1 1
                                                FROM MES.Barcode
                                                WHERE BarcodeNo = @BarcodeNo";
                                            dynamicParameters.Add("BarcodeNo", item3.BarcodeNo);

                                            var result4 = sqlConnection.Query(sql, dynamicParameters);
                                            if (result4.Count() <= 0) throw new SystemException("條碼【" + item3.BarcodeNo + "】資料錯誤!<br>託外生產單號:" + OspNo + "<br>製程:" + ProcessAlias + "");

                                            BarcodeNo = item3.BarcodeNo;
                                            #endregion
                                        }
                                    }
                                }
                                #endregion

                                #region //檢查是否已經有這筆託外生產條碼
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM MES.OspReceiptBarcode
                                        WHERE OspBarcodeId = @OspBarcodeId
                                        AND OsrDetailId = @OsrDetailId";
                                dynamicParameters.Add("OspBarcodeId", OspBarcodeId);
                                dynamicParameters.Add("OsrDetailId", OsrDetailId);

                                var result3 = sqlConnection.Query(sql, dynamicParameters);
                                if (result3.Count() > 0) continue;
                                #endregion

                                #region //檢查該條碼是否需要做工程檢且尚未完成，或是工程檢結果非良品
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.ProcessCheckStatus, a.ProcessCheckType
                                        FROM MES.MoProcess a
                                        WHERE MoProcessId = @MoProcessId";
                                dynamicParameters.Add("MoProcessId", MoProcessId);

                                var processCheckStatusResult = sqlConnection.Query(sql, dynamicParameters);

                                foreach (var item2 in processCheckStatusResult)
                                {
                                    if (item2.ProcessCheckStatus == "Y")
                                    {
                                        if (item2.ProcessCheckType == "1") //全檢
                                        {
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"SELECT TOP 1 a.QcStatus
                                                FROM MES.QcBarcode a
                                                INNER JOIN MES.Barcode b ON a.BarcodeId = b.BarcodeId
                                                INNER JOIN MES.QcRecord c ON a.QcRecordId = c.QcRecordId
                                                WHERE b.BarcodeNo = @BarcodeNo
                                                AND c.MoProcessId = b.CurrentMoProcessId
                                                ORDER BY a.CreateDate DESC";
                                            dynamicParameters.Add("BarcodeNo", BarcodeNo);

                                            var qcStatusResult = sqlConnection.Query(sql, dynamicParameters);
                                            if (qcStatusResult.Count() <= 0) throw new SystemException("該條碼【" + BarcodeNo + "】查無工程檢紀錄!<br>託外生產單號:" + OspNo + "<br>製程:" + ProcessAlias + "");

                                            foreach (var item3 in qcStatusResult)
                                            {
                                                if (item3.QcStatus != "P") throw new SystemException("該條碼【" + BarcodeNo + "】工程檢結果非良品且尚未修正!<br>託外生產單號:" + OspNo + "<br>製程:" + ProcessAlias + "");
                                            }
                                        }
                                        else if (item2.ProcessCheckType == "2") //抽檢
                                        {
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"SELECT TOP 1 1
                                                FROM MES.QcBarcode a
                                                INNER JOIN MES.QcRecord b ON a.QcRecordId = b.QcRecordId
                                                INNER JOIN MES.Barcode c ON a.BarcodeId = c.BarcodeId
                                                WHERE b.MoProcessId = c.CurrentMoProcessId AND a.QcStatus = 'P'";
                                            dynamicParameters.Add("MoProcessId", MoProcessId);

                                            var qcBarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                                            if (qcBarcodeResult.Count() <= 0) throw new SystemException("該製令製程查無任何一筆工程檢紀錄!<br>託外生產單號:" + OspNo + "<br>製程:" + ProcessAlias + "");
                                        }
                                    }
                                }
                                #endregion

                                #region //檢查是否有其他張相同託外生產單單身及條碼
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 d.BarcodeNo
                                        FROM MES.OspReceiptBarcode a 
                                        INNER JOIN MES.OspReceiptDetail b ON a.OsrDetailId = b.OsrDetailId
                                        INNER JOIN MES.OspBarcode c ON a.OspBarcodeId = c.OspBarcodeId
                                        INNER JOIN MES.Barcode d ON c.BarcodeNo = d.BarcodeNo
                                        WHERE a.OspBarcodeId = @OspBarcodeId
                                        AND b.OspDetailId = @OspDetailId";
                                dynamicParameters.Add("OspBarcodeId", OspBarcodeId);
                                dynamicParameters.Add("OspDetailId", ospDetailId);

                                var OspReceiptBarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                                if (OspReceiptBarcodeResult.Count() > 0)
                                {
                                    foreach (var item2 in OspReceiptBarcodeResult)
                                    {
                                        throw new SystemException("此條碼【" + item2.BarcodeNo + "】及此託外生產單據【" + OspNo + "】【" + ProcessAlias + "】已有建立紀錄，無法重複建立!!");
                                    }
                                }
                                #endregion

                                #region //若為供應商過站模式，確認條碼已過站
                                if (PassStationControl == "Y")
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.FinishDate
                                            FROM MES.BarcodeProcess a 
                                            INNER JOIN MES.Barcode b ON a.BarcodeId = b.BarcodeId
                                            WHERE b.BarcodeNo = @BarcodeNo
                                            AND a.MoProcessId = @MoProcessId";
                                    dynamicParameters.Add("BarcodeNo", BarcodeNo);
                                    dynamicParameters.Add("MoProcessId", MoProcessId);

                                    var BarcodeProcessResult = sqlConnection.Query(sql, dynamicParameters);

                                    foreach (var item2 in BarcodeProcessResult)
                                    {
                                        if (item2.FinishDate == null) throw new SystemException("條碼【" + BarcodeNo + "】尚在開工中，<br>請通知供應商完成過站程序!!");
                                    }
                                }
                                #endregion

                                #region //INSERT MES.OspReceiptBarcode
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO MES.OspReceiptBarcode (OsrDetailId, OspBarcodeId, Status
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        VALUES (@OsrDetailId, @OspBarcodeId, @Status
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        OsrDetailId,
                                        OspBarcodeId,
                                        Status = "I",
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });

                                insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();
                                #endregion

                                #region //確認目前此單身條碼是否已都過站，若已過站則更改單據狀態
                                string ProcessStatus = "N";
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT b.BarcodeNo
                                        , (
                                            SELECT x.FinishDate
                                            FROM MES.BarcodeProcess x 
                                            INNER JOIN MES.Barcode xa ON x.BarcodeId = xa.BarcodeId
                                            INNER JOIN BAS.[User] xb ON x.CreateBy = xb.UserId
                                            WHERE xa.BarcodeNo = b.BarcodeNo
                                            AND x.MoProcessId = c.MoProcessId
                                            FOR JSON PATH, ROOT('data')
                                        ) BarcodeProcess
                                        FROM MES.OspReceiptBarcode a 
                                        INNER JOIN MES.OspBarcode b ON a.OspBarcodeId = b.OspBarcodeId
                                        INNER JOIN MES.OspDetail c ON b.OspDetailId = c.OspDetailId
                                        WHERE c.OspDetailId = @OspDetailId";
                                dynamicParameters.Add("OspDetailId", ospDetailId);

                                var OspReceiptBarcodeResult2 = sqlConnection.Query(sql, dynamicParameters);

                                foreach (var item2 in OspReceiptBarcodeResult2)
                                {
                                    if (item2.BarcodeProcess == null)
                                    {
                                        ProcessStatus = "N";
                                        break;
                                    }
                                    else
                                    {
                                        JObject barcodeProcess = JObject.Parse(item2.BarcodeProcess);
                                        bool finishDataFlag = false;
                                        foreach (var item3 in barcodeProcess["data"])
                                        {
                                            if (item3["FinishDate"] == null)
                                            {
                                                finishDataFlag = true;
                                                break;
                                            }
                                        }

                                        if (finishDataFlag == true)
                                        {
                                            ProcessStatus = "N";
                                            break;
                                        }
                                        else
                                        {
                                            #region //確認此條碼是否下一站已指定回此托外站別
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"SELECT TOP 1 1
                                                    FROM MES.Barcode a 
                                                    WHERE a.BarcodeNo = @BarcodeNo
                                                    AND a.NextMoProcessId = @MoProcessId";
                                            dynamicParameters.Add("BarcodeNo", item2.BarcodeNo);
                                            dynamicParameters.Add("MoProcessId", MoProcessId);

                                            var BarcodeCheckResult = sqlConnection.Query(sql, dynamicParameters);

                                            if (BarcodeCheckResult.Count() > 0)
                                            {
                                                ProcessStatus = "N";
                                            }
                                            else
                                            {
                                                ProcessStatus = "Y";
                                            }
                                            #endregion

                                            break;
                                        }
                                    }
                                }
                                #endregion

                                #region //同步新增OsrDetail入庫數量
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE MES.OspReceiptDetail SET
                                        ReceiptQty = ReceiptQty + 1,
                                        AcceptQty = AcceptQty + 1,
                                        AvailableQty = AvailableQty + 1,
                                        ProcessStatus = @ProcessStatus,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE OsrDetailId = @OsrDetailId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        ProcessStatus,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        OsrDetailId
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region //若為委外中揚製作，且條碼為出貨狀態，則改回晶彩相關資料
                                if (CompanyId == 4 && SupplierNo == "200B0002" && BarcodeStatus == "10")
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE MES.Barcode SET
                                            CurrentMoProcessId = @MoProcessId,
                                            NextMoProcessId = @MoProcessId,
                                            MoId = @MoId,
                                            BarcodeStatus = @BarcodeStatus,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE BarcodeNo = @BarcodeNo";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            MoProcessId,
                                            MoId,
                                            BarcodeStatus = "1",
                                            LastModifiedDate,
                                            LastModifiedBy,
                                            BarcodeNo
                                        });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                }
                                #endregion
                            }
                            #endregion
                        }

                        #region //先取得此託外入庫單共有幾張製令
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT DISTINCT a.MoId
                                FROM MES.OspReceiptDetail a
                                WHERE a.OsrId = @OsrId";
                        dynamicParameters.Add("OsrId", OsrId);
                        var OspReceiptDetailMoIdResult = sqlConnection.Query(sql, dynamicParameters);
                        #endregion

                        foreach (var item in OspReceiptDetailMoIdResult)
                        {
                            #region 確認託外製程是否為連續順序
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT b.SortNumber
                                    , d.OspNo
                                    FROM MES.OspReceiptDetail a
                                    INNER JOIN MES.MoProcess b ON a.MoProcessId = b.MoProcessId
                                    INNER JOIN MES.OspDetail c ON a.OspDetailId = c.OspDetailId
                                    INNER JOIN MES.OutsourcingProduction d ON c.OspId = d.OspId
                                    WHERE a.OsrId = @OsrId
                                    AND a.MoId = @MoId";
                            dynamicParameters.Add("OsrId", OsrId);
                            dynamicParameters.Add("MoId", item.MoId);
                            var OspReceiptSortNumberResult = sqlConnection.Query(sql, dynamicParameters);

                            List<int> SortNumberList = new List<int>();
                            int CurrentMoProcessSortNumber = 0;
                            string OspNo = "";
                            foreach (var item2 in OspReceiptSortNumberResult)
                            {
                                SortNumberList.Add(item2.SortNumber);
                                OspNo = item2.OspNo;
                            }
                            SortNumberList.Sort();
                            int i = 1;
                            foreach (var item3 in SortNumberList)
                            {
                                if (i == 1)
                                {
                                    CurrentMoProcessSortNumber = item3;
                                }
                                else
                                {
                                    if (item3 - CurrentMoProcessSortNumber != 1)
                                    {
                                        #region //確認中間站別是否為非必要過站
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT TOP 1 1 
                                                    FROM MES.MoProcess
                                                    WHERE MoId = @MoId
                                                    AND SortNumber > @CurrentSortNumber
                                                    AND SortNumber < @SortNumber
                                                    AND NecessityStatus = 'Y'";
                                        dynamicParameters.Add("MoId", item.MoId);
                                        dynamicParameters.Add("CurrentSortNumber", CurrentMoProcessSortNumber);
                                        dynamicParameters.Add("SortNumber", item3);

                                        var SortNumberResult = sqlConnection.Query(sql, dynamicParameters);

                                        if (SortNumberResult.Count() > 0)
                                        {
                                            throw new SystemException("託外生產單號【" + OspNo + "】過站製程中順序沒有連續!!");
                                        }
                                        #endregion
                                    }
                                    else
                                    {
                                        CurrentMoProcessSortNumber = item3;
                                    }
                                }
                                i++;
                            }
                            #endregion
                        }

                        #region //更新託外入庫單頭資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a 
                                SET Quantity = (
                                    SELECT SUM(x.AcceptQty) 
                                    FROM MES.OspReceiptDetail x
                                    WHERE x.OsrId = a.OsrId
                                ),
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                FROM MES.OspReceipt a
                                WHERE a.OsrId = @OsrId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                LastModifiedDate,
                                LastModifiedBy,
                                OsrId
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

        #region //AddOspDetailExcel -- 新增託外生產單詳細資料 -- Ann 2022-09-07
        public string AddOspDetailExcel(int OspId, string ExcelData)
        {
            try
            {
                string ErrMeg = "";
                int rowsAffected = 0;
                int row = 1;
                int SupplierId = -1;
                if (OspId <= 0) throw new SystemException("託外生產單不能為空!");
                if (ExcelData.Length <= 0) throw new SystemException("Excel欲新增資料不能為空!");

                List<int> MoIdList = new List<int>();
                List<Dictionary<string, string>> ExcelJsonList = JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(ExcelData);


                using (TransactionScope transactionScope = new TransactionScope())
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

                        if (companyResult.Count() <= 0) throw new SystemException("公司別錯誤!");

                        foreach (var item in companyResult)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        }
                        #endregion

                        #region //判斷託外生產單資料是否有誤
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status,SupplierId
                                FROM MES.OutsourcingProduction
                                WHERE OspId = @OspId";
                        dynamicParameters.Add("OspId", OspId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("託外生產單資料有誤!");

                        foreach (var item in result)
                        {
                            if (item.Status != "N" && item.Status != "E") throw new SystemException("託外生產單已確認，無法更改!");
                            SupplierId = item.SupplierId;
                        }

                        foreach (var item in ExcelJsonList)
                        {
                            row++;
                            string Action = item["Action"] != null ? item["Action"].ToString() : throw new SystemException("【資料維護不完整】操作欄位資料不可以為空,請重新確認~~");
                            string MoFullNo = item["MoFullNo"] != null ? item["MoFullNo"].ToString() : throw new SystemException("【資料維護不完整】製令欄位資料不可以為空,請重新確認~~");
                            string MoProcess = item["MoProcess"] != null ? item["MoProcess"].ToString() : throw new SystemException("【資料維護不完整】加工製程欄位資料不可以為空,請重新確認~~");
                            string ProcessCode = item["ProcessCode"] != null ? item["ProcessCode"].ToString() : throw new SystemException("【資料維護不完整】加工代碼欄位資料不可以為空,請重新確認~~");
                            string ExpectedDate = item["ExpectedDate"] != null ? item["ExpectedDate"].ToString() : throw new SystemException("【資料維護不完整】預計回廠日期欄位資料不可以為空,請重新確認~~");
                            string ProcessCheckStatus = item["ProcessCheckStatus"] != null ? item["ProcessCheckStatus"].ToString() : throw new SystemException("【資料維護不完整】是否支援工程檢欄位資料不可以為空,請重新確認~~");
                            string ProcessCheckType = item["ProcessCheckType"] != null ? item["ProcessCheckType"].ToString() : throw new SystemException("【資料維護不完整】工程檢頻率欄位資料不可以為空,請重新確認~~");
                            string BarcodeList = item["BarcodeList"] != null ? item["BarcodeList"].ToString() : "";
                            int OspQty = item["OspQty"] != null ? Convert.ToInt32(item["OspQty"]) : 0;
                            int SuppliedQty = item["SuppliedQty"] != null ? Convert.ToInt32(item["SuppliedQty"]) : 0;
                            int OspDetailId = item["OspDetailId"] != null ? Convert.ToInt32(item["OspDetailId"]) : 0;
                            if (OspQty < 0) throw new SystemException("【託外生產數量】不能為負!");
                            if (SuppliedQty < 0) throw new SystemException("【供貨生產數量】不能為負!");
                            int MoProcessId = Convert.ToInt32(MoProcess.Split(':')[0]);

                            #region //判斷MES製令資料是否有誤
                            int MoId = -1;
                            int ModeId = -1;
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.ModeId,a.MoId
                                    FROM MES.ManufactureOrder a
                                    INNER JOIN MES.WipOrder a1 ON a.WoId = a1.WoId
                                    WHERE a1.WoErpPrefix +'-'+a1.WoErpNo +'('+CONVERT(VARCHAR, a.WoSeq)  +')' = @MoFullNo";
                            dynamicParameters.Add("MoFullNo", MoFullNo);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0)
                            {
                                ErrMeg += "找不到MES製令" + MoFullNo + ",請重新確認!!<br>";
                                continue;
                            }
                            else
                            {
                                foreach (var item1 in result)
                                {
                                    MoId = item1.MoId;
                                    ModeId = item1.ModeId;
                                    if (!MoIdList.Contains(MoId))
                                    {
                                        MoIdList.Add(MoId);
                                    }
                                }
                            }
                            #endregion

                            #region //判斷製令製程是否存在
                            int nowSortNumber = -1;
                            int ProcessId = -1;
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 SortNumber,ProcessId
                                    FROM MES.MoProcess a 
                                    WHERE a.MoProcessId = @MoProcessId
                                    AND a.MoId = @MoId";
                            dynamicParameters.Add("MoProcessId", MoProcessId);
                            dynamicParameters.Add("MoId", MoId);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0)
                            {
                                ErrMeg += "資料異常，此製程" + MoProcess + "不存在製令" + MoFullNo + "，請重新確認!!<br>";
                                continue;
                            }
                            else
                            {
                                foreach (var item1 in result)
                                {
                                    nowSortNumber = item1.SortNumber;
                                    ProcessId = item1.ProcessId;
                                }
                            }
                            #endregion

                            #region //確認供應商是否為自主過站模式，若是則檢核是否已設定過站機台
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.PassStationControl
                                FROM SCM.Supplier a 
                                WHERE a.SupplierId = @SupplierId";
                            dynamicParameters.Add("SupplierId", SupplierId);

                            var SupplierResult = sqlConnection.Query(sql, dynamicParameters);

                            if (SupplierResult.Count() <= 0) throw new SystemException("供應商資料錯誤!!");

                            foreach (var item1 in SupplierResult)
                            {
                                if (item1.PassStationControl == "Y")
                                {
                                    #region //檢核是否已設定供應商過站機台
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                        FROM SCM.SupplierMachine a 
                                        WHERE a.SupplierId = @SupplierId
                                        AND a.ProcessId = @ProcessId";
                                    dynamicParameters.Add("SupplierId", SupplierId);
                                    dynamicParameters.Add("ProcessId", ProcessId);

                                    var SupplierMachineResult = sqlConnection.Query(sql, dynamicParameters);

                                    if (SupplierMachineResult.Count() <= 0) throw new SystemException("此供應商尚未維護此製程機台!!");
                                    #endregion
                                }
                            }
                            #endregion




                            #region //判斷製令製程是否已存在且已過站
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.OspReceiptDetail a 
                                    INNER JOIN MES.OspDetail b ON a.OspDetailId = b.OspDetailId
                                    WHERE b.MoProcessId = @MoProcessId
                                    AND a.ProcessStatus = 'N'";
                            dynamicParameters.Add("MoProcessId", MoProcessId);

                            result = sqlConnection.Query(sql, dynamicParameters);

                            if (result.Count() > 0)
                            {
                                ErrMeg += "此製令" + MoFullNo + "製程" + MoProcess + "已存在託外單據，且未過站!!<br>";
                                continue;
                            }
                            #endregion

                            string ProcessCodeName = "";
                            using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                            {
                                #region //判斷ERP製程代號資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 MW002
                                        FROM CMSMW
                                        WHERE MW001 = @MW001";
                                dynamicParameters.Add("MW001", ProcessCode);

                                result = sqlConnection2.Query(sql, dynamicParameters);
                                if (result.Count() <= 0)
                                {
                                    ErrMeg += "找不到ERP製程代號" + ProcessCode + "資料有誤!!<br>";
                                    continue;
                                }
                                else
                                {
                                    foreach (var item1 in result)
                                    {
                                        ProcessCodeName = item1.MW002;
                                    }
                                }
                                #endregion
                            }

                            #region //若需工程檢，自動開立量測單據
                            if (ProcessCheckStatus == "Y")
                            {
                                #region //取得預設SpreadsheetData
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 a.DefaultSpreadsheetData
                                        FROM MES.QcRecord a";

                                var DefaultFileIdResult = sqlConnection.Query(sql, dynamicParameters);

                                string DefaultSpreadsheetData = "";
                                foreach (var item2 in DefaultFileIdResult)
                                {
                                    DefaultSpreadsheetData = item2.DefaultSpreadsheetData;
                                }
                                #endregion

                                #region //取得此生產模式下IPQC資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.QcTypeId
                                        FROM QMS.QcType a 
                                        WHERE a.ModeId = @ModeId
                                        AND a.QcTypeNo = 'IPQC'";
                                dynamicParameters.Add("ModeId", ModeId);

                                var QcTypeResult = sqlConnection.Query(sql, dynamicParameters);

                                if (QcTypeResult.Count() <= 0)
                                {
                                    ErrMeg += "該製令" + MoFullNo + "採用生產模式尚未設定量測類型，請通知系統開發室!!<br>";
                                    continue;

                                }

                                int QcTypeId = -1;
                                foreach (var item2 in QcTypeResult)
                                {
                                    QcTypeId = item2.QcTypeId;
                                }
                                #endregion

                                #region //INSERT MES.QcRecord
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO MES.QcRecord (QcTypeId, InputType, MoId, MoProcessId, Remark, DefaultFileId, CurrentFileId, DefaultSpreadsheetData, CheckQcMeasureData
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.QcRecordId
                                    VALUES (@QcTypeId, @InputType, @MoId, @MoProcessId, @Remark, @DefaultFileId, @CurrentFileId, @DefaultSpreadsheetData, @CheckQcMeasureData
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        QcTypeId,
                                        InputType = "BarcodeNo",
                                        MoId,
                                        MoProcessId,
                                        Remark = "由託外生產單自動建立",
                                        DefaultFileId = -1,
                                        CurrentFileId = -1,
                                        DefaultSpreadsheetData,
                                        CheckQcMeasureData = "N",
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }
                            #endregion

                            
                            if (OspDetailId <= 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO MES.OspDetail (OspId, MoId, MoProcessId, ProcessCheckStatus, ProcessCheckType, OspQty, SuppliedQty, ProcessCode, ProcessCodeName, ExpectedDate
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.OspDetailId
                                        VALUES (@OspId, @MoId, @MoProcessId, @ProcessCheckStatus, @ProcessCheckType, @OspQty, @SuppliedQty, @ProcessCode, @ProcessCodeName, @ExpectedDate
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        OspId,
                                        MoId,
                                        MoProcessId,
                                        ProcessCheckStatus = "N",
                                        ProcessCheckType = "",
                                        OspQty,
                                        SuppliedQty,
                                        ProcessCode,
                                        ProcessCodeName,
                                        ExpectedDate,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });

                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();

                                OspDetailId = -1;
                                foreach (var item1 in insertResult)
                                {
                                    OspDetailId = item1.OspDetailId;
                                }
                            }
                            else
                            {
                                if(Action == "U")
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE MES.OspDetail SET
                                            OspId = @OspId,
                                            MoId = @MoId,
                                            MoProcessId = @MoProcessId,
                                            ProcessCheckStatus = @ProcessCheckStatus,
                                            ProcessCheckType = @ProcessCheckType,
                                            OspQty = @OspQty,
                                            SuppliedQty = @SuppliedQty,
                                            ProcessCode = @ProcessCode,
                                            ProcessCodeName = @ProcessCodeName,
                                            ExpectedDate = @ExpectedDate,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE OspDetailId = @OspDetailId";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            OspId,
                                            MoId,
                                            MoProcessId,
                                            ProcessCheckStatus = "N",
                                            ProcessCheckType = "",
                                            OspQty,
                                            SuppliedQty,
                                            ProcessCode,
                                            ProcessCodeName,
                                            ExpectedDate,
                                            LastModifiedDate,
                                            LastModifiedBy,
                                            OspDetailId
                                        });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                                    if (BarcodeList != "")
                                    {
                                        #region //刪除主要table
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"DELETE MES.OspBarcode
                                        WHERE OspDetailId = @OspDetailId";
                                        dynamicParameters.Add("OspDetailId", OspDetailId);

                                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                        #endregion

                                    }
                                }
                                else if (Action == "D")
                                {
                                    #region //刪除主要table
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE MES.OspBarcode
                                        WHERE OspDetailId = @OspDetailId";
                                    dynamicParameters.Add("OspDetailId", OspDetailId);

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //刪除主要table
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE MES.OspDetail
                                        WHERE OspDetailId = @OspDetailId";
                                    dynamicParameters.Add("OspDetailId", OspDetailId);

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                                else
                                {
                                    throw new SystemException("操作欄位異常,請重新確認~~");
                                }

                            }

                            if (BarcodeList != "" && Action != "D")
                            {
                                string mesErr = "";

                                foreach (var item1 in BarcodeList.Split(','))
                                {
                                    int BarcodeId = Convert.ToInt32(item1);
                                    string BarcodeNo = "";

                                    #region //撈取BarcodeNo
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 BarcodeNo
                                            FROM MES.Barcode
                                            WHERE BarcodeId = @BarcodeId";
                                    dynamicParameters.Add("BarcodeId", BarcodeId);

                                    result = sqlConnection.Query(sql, dynamicParameters);
                                    if (result.Count() <= 0)
                                    {
                                        ErrMeg += "該條碼ID" + BarcodeId + "不存在，請重新確認!!<br>";
                                        continue;
                                    }
                                    else
                                    {
                                        foreach (var item2 in result)
                                        {
                                            BarcodeNo = item2.BarcodeNo;
                                        }
                                    }
                                    #endregion

                                    #region //判斷 託外生產單詳細資料 + 條碼 是否重複
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM MES.OspBarcode
                                            WHERE OspDetailId = @OspDetailId
                                            AND BarcodeNo = @BarcodeNo";
                                    dynamicParameters.Add("OspDetailId", OspDetailId);
                                    dynamicParameters.Add("BarcodeNo", BarcodeNo);

                                    var result3 = sqlConnection.Query(sql, dynamicParameters);
                                    if (result3.Count() > 0) throw new SystemException("【託外生產單詳細資料 + 條碼】重複，請重新輸入!");
                                    #endregion

                                    #region //判斷 託外條碼有沒有在該製令上
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT 1
                                             FROM MES.MoProcess a
                                             WHERE a.MoProcessId = @MoProcessId
                                             AND ((a.SortNumber = 1) or
                                                    EXISTS (SELECT 1
		                                                      FROM MES.Barcode b
		                                                     WHERE a.MoId = b.MoId
				                                               AND b.BarcodeNo = @BarcodeNo))";
                                    dynamicParameters.Add("MoProcessId", MoProcessId);
                                    dynamicParameters.Add("BarcodeNo", BarcodeNo);

                                    var result4 = sqlConnection.Query(sql, dynamicParameters);
                                    if (result4.Count() <= 0)
                                    {
                                        ErrMeg += "該條碼ID" + BarcodeId + "不存在製令，請重新確認!!<br>";
                                        continue;
                                    }
                                    #endregion

                                    #region //判斷 託外生產單詳細資料 + 條碼 是否重複
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.MoId,a.BarcodeId,a.BarcodeNo,a.BarcodeQty
                                                 , b.ProcessAlias CurrentProcessAlias
                                                 , ISNULL(b1.ProcessAlias, '已完工') NextProcessAlias
                                                 , c.StatusName
                                                 , c1.TypeName
                                                 ,CASE 
                                                    WHEN  x.BarcodeNo is not null THEN 'Y'
                                                    ELSE  'N'
                                                 END CurrentBind
                                                 ,CASE 
                                                    WHEN y.ProcessStatus is null AND y.OspDetailId is null THEN 'N' 
                                                    WHEN y.ProcessStatus is null AND y.OspDetailId is NOT null THEN 'Y' 
                                                    WHEN y.ProcessStatus = 'Y' THEN 'N' 
                                                    WHEN y.ProcessStatus = 'N' THEN 'Y'  
                                                 END OtherBind
                                                 ,CASE 
                                                    WHEN z.OspBind is not null THEN 'Y' 
                                                    ELSE  'N'
                                                 END OspBind
                                                 FROM MES.Barcode a
                                                 INNER JOIN MES.ManufactureOrder a1 ON a.MoId = a1.MoId
                                                 INNER JOIN MES.WipOrder a2 ON a1.WoId = a2.WoId
                                                 INNER JOIN MES.MoProcess b ON a.CurrentMoProcessId = b.MoProcessId
                                                 LEFT JOIN MES.MoProcess b1 ON a.NextMoProcessId = b1.MoProcessId
                                                 INNER JOIN BAS.[Status] c ON a.BarcodeStatus = c.StatusNo AND c.StatusSchema = 'Barcode.BarcodeStatus'
                                                 INNER JOIN BAS.[Type] c1 ON a.CurrentProdStatus = c1.TypeNo AND c1.TypeSchema = 'Barcode.CurrentProdStatus'
                                                 OUTER APPLY(
                                                     SELECT x1.BarcodeNo
                                                     FROM MES.OspBarcode x1
                                                     INNER JOIN MES.OspDetail x2 on x1.OspDetailId =x2.OspDetailId
                                                     INNER JOIN MES.OutsourcingProduction x3 on x2.OspId =x3.OspId
                                                     WHERE x1.BarcodeNo = a.BarcodeNo 
                                                     AND x3.OspId = @OspId
                                                     AND x2.MoProcessId = @MoProcessId
                                                 ) x
                                                 OUTER APPLY(
                                                     SELECT TOP 1 x4.ProcessStatus, x2.OspDetailId
                                                     FROM MES.OspBarcode x1
                                                     INNER JOIN MES.OspDetail x2 ON x1.OspDetailId = x2.OspDetailId
                                                     LEFT JOIN MES.OspReceiptBarcode x3 ON x1.OspBarcodeId = x3.OspBarcodeId
                                                     LEFT JOIN MES.OspReceiptDetail x4 ON x3.OsrDetailId = x4.OsrDetailId
                                                     WHERE x1.BarcodeNo = a.BarcodeNo 
                                                     AND x2.MoProcessId = @MoProcessId
                                                     ORDER BY x4.ProcessStatus ASC
                                                 ) y
                                                 OUTER APPLY(
                                                     SELECT TOP 1 1 OspBind
                                                     FROM MES.OspBarcode x1
                                                     INNER JOIN MES.OspDetail x2 on x1.OspDetailId =x2.OspDetailId
                                                     INNER JOIN MES.OutsourcingProduction x3 on x2.OspId =x3.OspId
                                                     WHERE x3.OspId = @OspId
                                                     AND x2.MoProcessId = @MoProcessId
                                                 ) z
                                                 WHERE a.BarcodeId = @BarcodeId
                                                 AND a.BarcodeStatus = '1'";
                                    dynamicParameters.Add("BarcodeId", BarcodeId);
                                    dynamicParameters.Add("OspId", OspId);
                                    dynamicParameters.Add("MoProcessId", MoProcessId);

                                    result = sqlConnection.Query(sql, dynamicParameters);
                                    foreach (var item2 in result)
                                    {
                                        if (item2.OtherBind == "Y")
                                        {
                                            mesErr += "第" + row + "列的條碼【" + item2.BarcodeNo + "】目前已被綁定或是尚未過站,不能綁定!!<br>";
                                            continue;
                                        }
                                        else
                                        {
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"INSERT INTO MES.OspBarcode (OspDetailId, BarcodeNo, BarcodeQty
                                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                                        OUTPUT INSERTED.OspBarcodeId
                                                        VALUES (@OspDetailId, @BarcodeNo, @BarcodeQty
                                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                            dynamicParameters.AddDynamicParams(
                                                new
                                                {
                                                    OspDetailId,
                                                    BarcodeNo = item2.BarcodeNo,
                                                    BarcodeQty = item2.BarcodeQty,
                                                    CreateDate,
                                                    LastModifiedDate,
                                                    CreateBy,
                                                    LastModifiedBy
                                                });

                                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                            rowsAffected += insertResult.Count();
                                        }
                                    }
                                    #endregion
                                }
                                if (mesErr != "") throw new SystemException(mesErr);

                            }
                        }

                        #endregion

                       


                        #region //判斷製令製程是否有斷站
                        foreach (var MoId in MoIdList)
                        {
                            string MoFullNo = "";
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a1.WoErpPrefix +'-'+a1.WoErpNo +'('+CONVERT(VARCHAR, a.WoSeq)  +')' MoFullNo
                                    FROM MES.ManufactureOrder a
                                    INNER JOIN MES.WipOrder a1 on a.WoId = a1.WoId
                                    WHERE a.MoId = @MoId";
                            dynamicParameters.Add("MoId", MoId);
                            result = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item in result)
                            {
                                MoFullNo = item.MoFullNo;
                            }

                            dynamicParameters = new DynamicParameters();
                            sql = @"
                                    WITH A AS (
                                        SELECT 
                                            b.SortNumber,
                                            b.NecessityStatus,
                                            ROW_NUMBER() OVER (ORDER BY b.SortNumber ASC) AS rn_asc,
                                            ROW_NUMBER() OVER (ORDER BY b.SortNumber DESC) AS rn_desc
                                        FROM MES.OspDetail a
                                        INNER JOIN MES.MoProcess b ON a.MoProcessId = b.MoProcessId
                                        OUTER APPLY (
                                            SELECT TOP 1 SortNumber LastSortNumber
                                            FROM MES.MoProcess x1
                                            WHERE x1.MoId = b.MoId
                                            AND x1.SortNumber < b.SortNumber
                                            AND x1.NecessityStatus = 'Y'
                                            ORDER BY x1.SortNumber DESC
                                        ) x
                                        WHERE a.OspId = @OspId
                                        AND a.MoId = @MoId
                                    ),
                                    FirstAndLast AS (
                                        SELECT
                                            MIN(CASE WHEN rn_asc = 1 THEN SortNumber ELSE NULL END) AS FirstSortNumber,
                                            MAX(CASE WHEN rn_desc = 1 THEN SortNumber ELSE NULL END) AS LastSortNumber
                                        FROM A
                                    )
                                    SELECT 
                                        CASE 
                                            WHEN EXISTS (
                                                SELECT 1
                                                FROM MES.MoProcess b
                                                INNER JOIN FirstAndLast fal ON 1 = 1
                                                WHERE 
                                                    b.MoId = @MoId
                                                    AND b.SortNumber > fal.FirstSortNumber
                                                    AND b.SortNumber < fal.LastSortNumber
                                                    AND b.SortNumber NOT IN (
                                                        SELECT SortNumber
                                                        FROM A
                                                    )
                                                    AND b.NecessityStatus = 'Y'
                                            ) THEN 'N'
                                            ELSE 'Y'
                                        END AS Result;

                                    ";
                            dynamicParameters.Add("OspId", OspId);
                            dynamicParameters.Add("MoId", MoId);
                            result = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item in result)
                            {
                                if (item.Result == "N")
                                {
                                    ErrMeg += "目前開立的託外單,其中製令【" + MoFullNo + "】具有非連續製程資料，不能開立請重新確認!!<br>";
                                }

                                    //throw new SystemException("目前開立的託外單,其中製令:" + MoFullNo + "具有非連續製程資料，不能開立請重新確認!!");
                            }
                        }

                        #endregion

                        if (ErrMeg != "")
                        {
                            throw new SystemException(ErrMeg);
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


        #endregion

        #region //Update
        #region //UpdateOutsourcingProduction -- 更新託外生產單資料 -- Ann 2022-09-06
        public string UpdateOutsourcingProduction(int OspId, int DepartmentId, string OspNo, string OspDate, int SupplierId, string OspStatus
            , string OspDesc, string Status)
        {
            try
            {
                if (OspDate.Length <= 0) throw new SystemException("【托外生產日期】不能為空!");
                if (OspStatus.Length <= 0) throw new SystemException("【供料狀態】不能為空!");
                if (OspDesc.Length > 100) throw new SystemException("【描述】狀態錯誤!");
                if (Status.Length <= 0) throw new SystemException("【狀態】狀態錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
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

                        if (companyResult.Count() <= 0) throw new SystemException("公司別錯誤!");

                        foreach (var item in companyResult)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        }
                        #endregion

                        #region //判斷託外生產資料是否有誤
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM MES.OutsourcingProduction
                                WHERE OspId = @OspId";
                        dynamicParameters.Add("OspId", OspId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("託外生產資料有誤!");

                        foreach (var item in result)
                        {
                            if (item.Status != "N" && item.Status != "E") throw new SystemException("託外生產單已確認，無法更改!");
                        }
                        #endregion

                        #region //託外生產單詳細資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT *
                                FROM MES.OspDetail
                                WHERE OspId = @OspId";
                        dynamicParameters.Add("OspId", OspId);

                        var ospDetailresult = sqlConnection.Query(sql, dynamicParameters);
                        //if (ospDetailresult.Count() <= 0) throw new SystemException("託外生產單詳細資料有誤!");
                        #endregion

                        #region //判斷部門資料是否有誤
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Department
                                WHERE CompanyId = @CompanyId
                                AND DepartmentId = @DepartmentId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("DepartmentId", DepartmentId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("部門資料有誤!");
                        #endregion

                        #region //判斷供應商資料是否有誤
                        string SupplierNo = "";
                        if (SupplierId > 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 SupplierNo
                                FROM SCM.Supplier
                                WHERE CompanyId = @CompanyId
                                AND SupplierId = @SupplierId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("SupplierId", SupplierId);

                            var result3 = sqlConnection.Query(sql, dynamicParameters);
                            if (result3.Count() <= 0) throw new SystemException("供應商資料有誤!");

                            foreach (var item in result3)
                            {
                                SupplierNo = item.SupplierNo;
                            }
                        }
                        #endregion

                        #region //判斷單號是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.OutsourcingProduction
                                WHERE OspNo = @OspNo
                                AND OspId != @OspId";
                        dynamicParameters.Add("OspNo", OspNo);
                        dynamicParameters.Add("OspId", OspId);

                        var result4 = sqlConnection.Query(sql, dynamicParameters);
                        if (result4.Count() > 0) throw new SystemException("託外生產單單號重複!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.OutsourcingProduction SET
                                DepartmentId = @DepartmentId,
                                OspNo = @OspNo,
                                OspStatus = @OspStatus,
                                OspDate = @OspDate,
                                SupplierId = @SupplierId,
                                OspDesc = @OspDesc,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE OspId = @OspId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                DepartmentId,
                                OspNo,
                                OspStatus,
                                OspDate,
                                SupplierId = SupplierId < 0 ? (int?)null : SupplierId,
                                OspDesc,
                                LastModifiedDate,
                                LastModifiedBy,
                                OspId
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

        #region //UpdateOspDetail -- 更新託外生產單詳細資料 -- Ann 2022-09-07
        public string UpdateOspDetail(int OspDetailId, int OspId, int MoId, int MoProcessId, string ProcessCheckStatus, string ProcessCheckType, int OspQty, int SuppliedQty
            , string ProcessCode, string ExpectedDate, string Remark)
        {
            try
            {
                if (ProcessCheckStatus.Length <= 0) throw new SystemException("【是否支援工程檢】不能為空!");
                if (ProcessCheckStatus == "Y" && ProcessCheckType.Length <= 0) throw new SystemException("【工程檢頻率】不能為空!");
                if (OspQty < 0) throw new SystemException("【託外生產數量】不能為空!");
                if (SuppliedQty < 0) throw new SystemException("【供貨生產數量】不能為空!");
                if (ExpectedDate.Length <= 0) throw new SystemException("【預計回廠日期】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
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

                        if (companyResult.Count() <= 0) throw new SystemException("公司別錯誤!");

                        foreach (var item in companyResult)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        }
                        #endregion

                        #region //判斷託外生產單詳細資料是否有誤
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 MoProcessId
                                FROM MES.OspDetail
                                WHERE OspDetailId = @OspDetailId";
                        dynamicParameters.Add("OspDetailId", OspDetailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("託外生產單詳細資料有誤!");

                        int preMoProcessId = -1;
                        foreach (var item in result)
                        {
                            preMoProcessId = item.MoProcessId;
                        }
                        #endregion

                        #region //判斷託外生產單資料是否有誤
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM MES.OutsourcingProduction
                                WHERE OspId = @OspId";
                        dynamicParameters.Add("OspId", OspId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("託外生產單資料有誤!");

                        foreach (var item in result2)
                        {
                            if (item.Status != "N" && item.Status != "E") throw new SystemException("託外生產單已確認，無法更改!");
                        }
                        #endregion

                        #region //判斷MES製令資料是否有誤
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.ManufactureOrder
                                WHERE MoId = @MoId";
                        dynamicParameters.Add("MoId", MoId);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() <= 0) throw new SystemException("MES製令資料有誤!");
                        #endregion

                        string ProcessCodeName = "";
                        if (ProcessCode != "")
                        {
                            using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                            {
                                #region //判斷ERP製程代號資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 MW002
                                    FROM CMSMW
                                    WHERE MW001 = @MW001";
                                dynamicParameters.Add("MW001", ProcessCode);

                                var ProcessCodeResult = sqlConnection2.Query(sql, dynamicParameters);
                                if (ProcessCodeResult.Count() <= 0) throw new SystemException("ERP製程代號資料有誤!");

                                foreach (var item in ProcessCodeResult)
                                {
                                    ProcessCodeName = item.MW002;
                                }
                                #endregion
                            }
                        }

                        #region //判斷此託外生產單是否已註冊條碼
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.OspBarcode
                                WHERE OspDetailId = @OspDetailId";
                        dynamicParameters.Add("OspDetailId", OspDetailId);

                        var result6 = sqlConnection.Query(sql, dynamicParameters);
                        if (preMoProcessId != MoProcessId && result6.Count() > 0) throw new SystemException("此託外生產單已註冊條碼，無法更改加工製程!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.OspDetail SET
                                OspId = @OspId,
                                MoId = @MoId,
                                MoProcessId = @MoProcessId,
                                ProcessCheckStatus = @ProcessCheckStatus,
                                ProcessCheckType = @ProcessCheckType,
                                OspQty = @OspQty,
                                SuppliedQty = @SuppliedQty,
                                ProcessCode = @ProcessCode,
                                ProcessCodeName = @ProcessCodeName,
                                ExpectedDate = @ExpectedDate,
                                Remark = @Remark,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE OspDetailId = @OspDetailId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                OspId,
                                MoId,
                                MoProcessId,
                                ProcessCheckStatus,
                                ProcessCheckType = ProcessCheckType != "" ? ProcessCheckType : "N",
                                OspQty,
                                SuppliedQty,
                                ProcessCode,
                                ProcessCodeName,
                                ExpectedDate,
                                Remark,
                                LastModifiedDate,
                                LastModifiedBy,
                                OspDetailId
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

        #region //UpdateOspConfirm -- 確認託外生產單資料 -- Ann 2022-09-12
        public string UpdateOspConfirm(int OspId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷託外生產單資料是否有誤
                        DynamicParameters dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM MES.OutsourcingProduction
                                WHERE OspId = @OspId";
                        dynamicParameters.Add("OspId", OspId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("託外生產單詳細資料有誤!");

                        foreach (var item in result)
                        {
                            if (item.Status != "N") throw new SystemException("託外生產單已確認!");
                        }
                        #endregion

                        #region //判斷託外生產單詳細資料是否有誤
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.OspQty, a.SuppliedQty
                                , b.OspBarcodeId
                                FROM MES.OspDetail a
                                LEFT JOIN MES.OspBarcode b ON a.OspDetailId = b.OspDetailId
                                WHERE a.OspId = @OspId";
                        dynamicParameters.Add("OspId", OspId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("尚未建立託外生產單詳細資料!");

                        foreach (var item in result2)
                        {
                            if (item.OspQty == 0) throw new SystemException("託外生產數量不能為0!");
                            if (item.OspBarcodeId == null) throw new SystemException("尚未新增託外條碼!");
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.OutsourcingProduction SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE OspId = @OspId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = "Y",
                                LastModifiedDate,
                                LastModifiedBy,
                                OspId
                            });

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                        #region //更新MoProcess工程檢參數
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.MoProcess
                                   SET ProcessCheckStatus = b.ProcessCheckStatus,
                                       ProcessCheckType = b.ProcessCheckType,
                                       LastModifiedDate = @LastModifiedDate,
                                       LastModifiedBy = @LastModifiedBy
                                  FROM MES.MoProcess a
                                       INNER JOIN MES.OspDetail b ON a.MoProcessId = b.MoProcessId
	                                   INNER JOIN MES.OutsourcingProduction c ON b.OspId = c.OspId
                                 WHERE c.OspId = @OspId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                LastModifiedDate,
                                LastModifiedBy,
                                OspId
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

        #region //UpdateOspReConfirm -- 反確認託外生產單資料 -- Ann 2022-09-12
        public string UpdateOspReConfirm(int OspId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷託外生產單資料是否有誤
                        DynamicParameters dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.Status
                                , b.OspDetailId, b.MoProcessId
                                FROM MES.OutsourcingProduction a
                                INNER JOIN MES.OspDetail b ON a.OspId = b.OspId
                                WHERE a.OspId = @OspId";
                        dynamicParameters.Add("OspId", OspId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("託外生產單詳細資料有誤!");

                        int OspDetailId = -1;
                        int MoProcessId = -1;
                        foreach (var item in result)
                        {
                            if (item.Status != "Y") throw new SystemException("託外生產單尚未確認!");
                            OspDetailId = item.OspDetailId;
                            MoProcessId = item.MoProcessId;
                        }
                        #endregion

                        #region //判斷條碼狀態是否有誤
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT b.NextMoProcessId,
                                       (SELECT COUNT(1) BPCount
                                          FROM MES.BarcodeProcess b
		                                       INNER JOIN MES.OspDetail c ON b.MoProcessId = c.MoProcessId
                                         WHERE c.OspDetailId = a.OspDetailId) BPCount
                                            , (
                                                SELECT TOP 1 1 FROM MES.OspReceiptDetail x 
                                                    WHERE x.OspDetailId = a.OspDetailId
                                                    AND x.ProcessStatus = 'Y'
                                            ) OsrDetailFlag
                                 FROM MES.OspBarcode a
                                      INNER JOIN MES.Barcode b ON a.BarcodeNo = b.BarcodeNo
                                WHERE a.OspDetailId = @OspDetailId";
                        dynamicParameters.Add("OspDetailId", OspDetailId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("託外生產條碼有誤!");

                        foreach (var item in result2)
                        {
                            if (Convert.ToInt32(item.BPCount) > 0 && item.OsrDetailFlag != null) throw new SystemException("託外生產條碼已過站，無法反確認!");
                        }
                        #endregion

                        #region //檢查此託外生產單詳細資料是否已綁定託外入庫單
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.OutsourcingProduction a
                                INNER JOIN MES.OspDetail b ON a.OspId = b.OspId
                                INNER JOIN MES.OspReceiptDetail c ON b.OspDetailId = c.OspDetailId
                                WHERE a.OspId = @OspId";
                        dynamicParameters.Add("OspId", OspId);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() > 0) throw new SystemException("此託外生產單已被綁定，無法反確認!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.OutsourcingProduction SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE OspId = @OspId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = "N",
                                LastModifiedDate,
                                LastModifiedBy,
                                OspId
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

        #region //UpdateOspReceipt -- 更新託外入庫單資料 -- Ann 2022-09-15
        public string UpdateOspReceipt(int OsrId, int SupplierId, string SupplierSo, string OsrErpPrefix, string OsrErpNo, string DocumentDate, string ReceiptDate, string ReserveTaxCode, string TaxCode
            , string TaxType, string InvoiceType, double TaxRate, string UiNo, string InvoiceDate, string InvoiceNo, string ApplyYYMM, string CurrencyCode, double Exchange, int RowCnt
            , string Remark, string DeductType, string PaymentTerm)
        {
            try
            {
                if (OsrErpPrefix.Length <= 0) throw new SystemException("【ERP託外入庫單別】不能為空!");
                if (OsrErpPrefix != "5902") throw new SystemException("【ERP託外入庫單別】錯誤!");
                if (DocumentDate.Length <= 0) throw new SystemException("【單據日期】不能為空!");
                if (ReceiptDate.Length <= 0) throw new SystemException("【託外入庫日期】不能為空!");
                if (ReserveTaxCode.Length <= 0) throw new SystemException("【保稅碼】不能為空!");
                //if (TaxRate < 0) throw new SystemException("【營業稅率】不能為空!");
                if (ApplyYYMM.Length <= 0) throw new SystemException("【申報日期】不能為空!");
                //if (ApplyYYMM.Length > 6) throw new SystemException("【申報日期】長度錯誤!");
                if (Exchange <= 0) throw new SystemException("【匯率】不能為空!");
                if (RowCnt < 0) throw new SystemException("【件數】不能為空!");
                if (DeductType.Length <= 0) throw new SystemException("【扣抵區分】不能為空!");

                string LocalCurrencyCode = "";
                int UnitDecimal = 0;
                int PriceDecimal = 0;
                int OriUnitDecimal = 0;
                int OriPriceDecimal = 0;

                using (TransactionScope transactionScope = new TransactionScope())
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

                        if (companyResult.Count() <= 0) throw new SystemException("公司別錯誤!");

                        foreach (var item in companyResult)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        }
                        #endregion

                        #region //判斷託外入庫單資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.FactoryCode, a.ReserveFlag, a.OrigAmount, a.DeductAmount, a.OrigTax, a.ReceiptAmount
                                , a.Quantity, a.ConfirmStatus, a.RenewFlag, a.PrintCnt, a.AutoMaterialBilling, a.OrigPreTaxAmount
                                , a.PretaxAmount, a.TaxAmount, a.PackageQuantity, a.SupplierPicking, a.ApproveStatus, a.SendCount
                                , a.NoticeFlag, a.QcFlag, a.FlowStatus, a.TaxType, a.ConfirmStatus
                                FROM MES.OspReceipt a
                                WHERE OsrId = @OsrId";
                        dynamicParameters.Add("OsrId", OsrId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("託外入庫單資料有誤!");

                        string FactoryCode = "";
                        string ReserveFlag = "";
                        double OrigAmount = -1;
                        double DeductAmount = -1;
                        double OrigTax = -1;
                        double ReceiptAmount = -1;
                        int Quantity = -1;
                        string ConfirmStatus = "";
                        string RenewFlag = "";
                        int PrintCnt = -1;
                        string AutoMaterialBilling = "";
                        double OrigPreTaxAmount = -1;
                        double PretaxAmount = -1;
                        double TaxAmount = -1;
                        int PackageQuantity = -1;
                        string SupplierPicking = "";
                        string ApproveStatus = "";
                        int SendCount = -1;
                        string NoticeFlag = "";
                        string QcFlag = "";
                        string FlowStatus = "";
                        string currentTaxType = "";
                        foreach (var item in result)
                        {
                            if (item.ConfirmStatus != "N") throw new SystemException("託外入庫單已確認，無法更改!");
                            FactoryCode = item.FactoryCode;
                            ReserveFlag = item.ReserveFlag;
                            OrigAmount = item.OrigAmount;
                            DeductAmount = item.DeductAmount;
                            OrigTax = item.OrigTax;
                            ReceiptAmount = item.ReceiptAmount;
                            Quantity = item.Quantity;
                            ConfirmStatus = item.ConfirmStatus;
                            RenewFlag = item.RenewFlag;
                            PrintCnt = item.PrintCnt;
                            AutoMaterialBilling = item.AutoMaterialBilling;
                            OrigPreTaxAmount = item.OrigPreTaxAmount;
                            PretaxAmount = item.PretaxAmount;
                            TaxAmount = item.TaxAmount;
                            PackageQuantity = item.PackageQuantity;
                            SupplierPicking = item.SupplierPicking;
                            ApproveStatus = item.ApproveStatus;
                            SendCount = item.SendCount;
                            NoticeFlag = item.NoticeFlag;
                            QcFlag = item.QcFlag;
                            FlowStatus = item.FlowStatus;
                            currentTaxType = item.TaxType;
                        }
                        #endregion

                        #region //判斷供應商資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.Supplier
                                WHERE SupplierId = @SupplierId";
                        dynamicParameters.Add("SupplierId", SupplierId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("供應商資料有誤!");
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //判斷稅別碼資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM CMSNN
                                    WHERE NN001 = @NN001";
                            dynamicParameters.Add("NN001", TaxCode);

                            var result3 = sqlConnection2.Query(sql, dynamicParameters);
                            if (result3.Count() <= 0) throw new SystemException("稅別碼資料有誤!");
                            #endregion

                            #region //判斷發票聯數資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM CMSNM
                                    WHERE NM001 = @NM001";
                            dynamicParameters.Add("NM001", InvoiceType);

                            var result4 = sqlConnection2.Query(sql, dynamicParameters);
                            if (result4.Count() <= 0) throw new SystemException("發票聯數資料有誤!");
                            #endregion

                            #region //判斷幣別資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM CMSMF
                                    WHERE MF001 = @MF001";
                            dynamicParameters.Add("MF001", CurrencyCode);

                            var result5 = sqlConnection2.Query(sql, dynamicParameters);
                            if (result5.Count() <= 0) throw new SystemException("幣別資料有誤!");
                            #endregion

                            #region //判斷付款條件資料是否正確
                            if (PaymentTerm != "")
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                    FROM CMSNA
                                    WHERE NA002 = @NA002";
                                dynamicParameters.Add("NA002", PaymentTerm);

                                var result6 = sqlConnection2.Query(sql, dynamicParameters);
                                if (result6.Count() <= 0) throw new SystemException("付款條件資料有誤!");
                            }
                            #endregion

                            #region //取得本幣幣別
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 LTRIM(RTRIM(MA003)) CurrencyCode
                                    FROM CMSMA
                                    WHERE 1=1";

                            var CMSMAResult = sqlConnection2.Query(sql, dynamicParameters);

                            foreach (var item in CMSMAResult)
                            {
                                LocalCurrencyCode = item.CurrencyCode;
                            }
                            #endregion

                            #region //本幣小數點取位資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.MF003, a.MF004
                                    FROM CMSMF a 
                                    WHERE a.MF001 = @Currency";
                            dynamicParameters.Add("Currency", CurrencyCode);

                            var CMSMFResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (CMSMFResult.Count() <= 0) throw new SystemException("金額小數點取位資料錯誤!!");

                            foreach (var item in CMSMFResult)
                            {
                                UnitDecimal = Convert.ToInt32(item.MF003);
                                PriceDecimal = Convert.ToInt32(item.MF004);
                            }
                            #endregion

                            #region //原幣小數點取位資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.MF003, a.MF004
                                    FROM CMSMF a 
                                    WHERE a.MF001 = @Currency";
                            dynamicParameters.Add("Currency", CurrencyCode);

                            var CMSMFResult2 = sqlConnection2.Query(sql, dynamicParameters);

                            if (CMSMFResult2.Count() <= 0) throw new SystemException("金額小數點取位資料錯誤!!");

                            foreach (var item in CMSMFResult2)
                            {
                                OriUnitDecimal = Convert.ToInt32(item.MF003);
                                OriPriceDecimal = Convert.ToInt32(item.MF004);
                            }
                            #endregion
                        }

                        #region /Update MES.OspReceipt
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.OspReceipt SET
                                OsrErpPrefix = @OsrErpPrefix,
                                OsrErpNo = @OsrErpNo,
                                ReceiptDate = @ReceiptDate,
                                FactoryCode = @FactoryCode,
                                SupplierId = @SupplierId,
                                SupplierSo = @SupplierSo,
                                CurrencyCode = @CurrencyCode,
                                Exchange = @Exchange,
                                RowCnt = @RowCnt,
                                Remark = @Remark,
                                UiNo = @UiNo,
                                InvoiceType = @InvoiceType,
                                InvoiceDate = @InvoiceDate,
                                InvoiceNo = @InvoiceNo,
                                TaxType = @TaxType,
                                DeductType = @DeductType,
                                ReserveFlag = @ReserveFlag,
                                OrigAmount = @OrigAmount,
                                DeductAmount = @DeductAmount,
                                OrigTax = @OrigTax,
                                ReceiptAmount = @ReceiptAmount,
                                Quantity = @Quantity,
                                ConfirmStatus = @ConfirmStatus,
                                RenewFlag = @RenewFlag,
                                PrintCnt = @PrintCnt,
                                AutoMaterialBilling = @AutoMaterialBilling,
                                OrigPreTaxAmount = @OrigPreTaxAmount,
                                ApplyYYMM = @ApplyYYMM,
                                DocumentDate = @DocumentDate,
                                TaxRate = @TaxRate,
                                PretaxAmount = @PretaxAmount,
                                TaxAmount = @TaxAmount,
                                PaymentTerm = @PaymentTerm,
                                PackageQuantity = @PackageQuantity,
                                SupplierPicking = @SupplierPicking,
                                ApproveStatus = @ApproveStatus,
                                ReserveTaxCode = @ReserveTaxCode,
                                SendCount = @SendCount,
                                NoticeFlag = @NoticeFlag,
                                TaxCode = @TaxCode,
                                TaxExchange = @TaxExchange,
                                QcFlag = @QcFlag,
                                FlowStatus = @FlowStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE OsrId = @OsrId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                OsrErpPrefix,
                                OsrErpNo,
                                ReceiptDate,
                                FactoryCode,
                                SupplierId,
                                SupplierSo,
                                CurrencyCode,
                                Exchange,
                                RowCnt,
                                Remark,
                                UiNo,
                                InvoiceType,
                                InvoiceDate,
                                InvoiceNo,
                                TaxType,
                                DeductType,
                                ReserveFlag,
                                OrigAmount,
                                DeductAmount,
                                OrigTax,
                                ReceiptAmount,
                                Quantity,
                                ConfirmStatus,
                                RenewFlag,
                                PrintCnt,
                                AutoMaterialBilling,
                                OrigPreTaxAmount,
                                ApplyYYMM,
                                DocumentDate,
                                TaxRate,
                                PretaxAmount,
                                TaxAmount,
                                PaymentTerm,
                                PackageQuantity,
                                SupplierPicking,
                                ApproveStatus,
                                ReserveTaxCode,
                                SendCount,
                                NoticeFlag,
                                TaxCode,
                                TaxExchange = 1,
                                QcFlag,
                                FlowStatus,
                                LastModifiedDate,
                                LastModifiedBy,
                                OsrId
                            });

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //同步更新單身價格
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT OsrDetailId, OrigPreTaxAmt, OrigTaxAmt, PreTaxAmt, TaxAmt
                                , AvailableQty, OrigUnitPrice, OrigDiscountAmt
                                FROM MES.OspReceiptDetail
                                WHERE OsrId = @OsrId";
                        dynamicParameters.Add("OsrId", OsrId);

                        var result7 = sqlConnection.Query(sql, dynamicParameters);
                        if (result7.Count() > 0)
                        {
                            double origPreTaxAmt = -1;
                            double origTaxAmt = -1;
                            double preTaxAmt = -1;
                            double taxAmt = -1;
                            foreach (var item in result7)
                            {
                                dynamicParameters = new DynamicParameters();
                                if (currentTaxType == TaxType)
                                {
                                    origPreTaxAmt = Math.Round(item.OrigPreTaxAmt, OriPriceDecimal);
                                    origTaxAmt = Math.Round(item.OrigTaxAmt, OriPriceDecimal);
                                    preTaxAmt = Math.Round(item.OrigPreTaxAmt * Exchange, PriceDecimal);
                                    taxAmt = Math.Round(item.OrigTaxAmt * Exchange, PriceDecimal);
                                }
                                else if (TaxType == "1")
                                {
                                    double t1OrigTaxAmt = Math.Round((item.AvailableQty * item.OrigUnitPrice - item.OrigDiscountAmt) * TaxRate);
                                    origPreTaxAmt = Math.Round((item.AvailableQty * item.OrigUnitPrice - item.OrigDiscountAmt) - t1OrigTaxAmt, OriPriceDecimal);
                                    origTaxAmt = Math.Round(t1OrigTaxAmt, OriPriceDecimal);
                                    preTaxAmt = Math.Round(origPreTaxAmt * Exchange, PriceDecimal);
                                    taxAmt = Math.Round(origTaxAmt * Exchange, PriceDecimal);
                                }
                                else if (TaxType == "2")
                                {
                                    origPreTaxAmt = Math.Round(item.AvailableQty * item.OrigUnitPrice - item.OrigDiscountAmt, OriPriceDecimal);
                                    origTaxAmt = Math.Round((item.AvailableQty * item.OrigUnitPrice - item.OrigDiscountAmt) * TaxRate, OriPriceDecimal);
                                    preTaxAmt = Math.Round(origPreTaxAmt * Exchange, PriceDecimal);
                                    taxAmt = Math.Round(origTaxAmt * Exchange, PriceDecimal);
                                }
                                else if (TaxType == "3" || TaxType == "4" || TaxType == "9")
                                {
                                    origPreTaxAmt = Math.Round(item.AvailableQty * item.OrigUnitPrice - item.OrigDiscountAmt, OriPriceDecimal);
                                    origTaxAmt = 0;
                                    preTaxAmt = Math.Round(origPreTaxAmt * Exchange, PriceDecimal);
                                    taxAmt = 0;
                                }

                                sql = @"UPDATE MES.OspReceiptDetail SET
                                        OrigPreTaxAmt = @OrigPreTaxAmt,
                                        OrigTaxAmt = @OrigTaxAmt,
                                        PreTaxAmt = @PreTaxAmt,
                                        TaxAmt = @TaxAmt,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE OsrDetailId = @OsrDetailId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        OrigPreTaxAmt = origPreTaxAmt,
                                        OrigTaxAmt = origTaxAmt,
                                        PreTaxAmt = preTaxAmt,
                                        TaxAmt = taxAmt,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        item.OsrDetailId
                                    });

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

        #region //UpdateOspReceiptDetail -- 更新託外入庫單詳細資料 -- Ann 2022-09-19
        public string UpdateOspReceiptDetail(int OsrDetailId, int OsrId, int OspDetailId, string OsrSeq, string LotNumber, int InventoryId, string AcceptanceDate
            , int ReceiptQty, string AvailableUom, int AcceptQty, int ScriptQty, int ReturnQty, int AvailableQty, double OrigUnitPrice, double OrigAmount
            , string DiscountDescription, string Overdue, string Remark, double OrigPreTaxAmt, double OrigTaxAmt, double OrigDiscountAmt, double PreTaxAmt, double TaxAmt)
        {
            try
            {
                if (LotNumber.Length > 20) throw new SystemException("【批號】長度錯誤!");
                if (AcceptanceDate.Length <= 0) throw new SystemException("【驗收日期】不能為空!");
                if (ReceiptQty < 0) throw new SystemException("【進貨數量】不能為空!");
                if (AcceptQty < 0) throw new SystemException("【驗收數量】不能為空!");
                if (ScriptQty < 0) throw new SystemException("【報廢數量】不能為空!");
                if (ReturnQty < 0) throw new SystemException("【驗退數量】不能為空!");
                if (AvailableQty < 0) throw new SystemException("【計價數量】不能為空!");
                if (OrigUnitPrice < 0) throw new SystemException("【原幣加工單價】不能為空!");
                if (OrigAmount < 0) throw new SystemException("【原幣加工金額】不能為空!");
                if (DiscountDescription.Length > 40) throw new SystemException("【扣款說明】長度錯誤!");
                if (Overdue.Length <= 0) throw new SystemException("【逾期碼】不能為空!");
                if (Remark.Length > 255) throw new SystemException("【備註】長度錯誤!");
                if (OrigDiscountAmt < 0) throw new SystemException("【原幣扣款金額】不能為空!");
                if (OrigPreTaxAmt < 0) throw new SystemException("【原幣未稅金額】不能為空!");
                if (OrigTaxAmt < 0) throw new SystemException("【原幣稅金額】不能為空!");
                if (PreTaxAmt < 0) throw new SystemException("【本幣未稅金額】不能為空!");
                if (TaxAmt < 0) throw new SystemException("【本幣稅金額】不能為空!");

                if (ReceiptQty < AvailableQty) throw new SystemException("計價數量不能超過進貨數量!!");

                //進貨數量 = 驗收數量 + 報廢數量 + 驗退數量
                if (ReceiptQty != (AcceptQty + ScriptQty + ReturnQty)) throw new SystemException("數量錯誤，進貨數量 = 驗收數量 + 報廢數量 + 驗退數量!");

                using (TransactionScope transactionScope = new TransactionScope())
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
                        if (resultCompany.Count() <= 0) throw new SystemException("【公司別】資料錯誤!");

                        foreach (var item in resultCompany)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //判斷託外入庫詳細資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT AvailableDate, ReCheckDate, ReceiptPackageQty, AcceptancePackageQty, ReceiptExpense, OsrSeq, AcceptQty OriAcceptQty
                                    , ProjectCode, PaymentHold, QcStatus, ReturnCode, ConfirmStatus, CloseStatus, ReNewStatus, CostEntry, ExpenseEntry
                                    , ISNULL(ConfirmUserId, -1) ConfirmUserId, ConfirmUser, UrgentMtl
                                    , ApproveStatus, TransferStatus, ProcessStatus, OrigAmount OriOrigAmount, OrigTaxAmt OriOrigTaxAmt, PreTaxAmt OriPreTaxAmt
                                    , TaxAmt OriTaxAmt, OrigPreTaxAmt OriOrigPreTaxAmt
                                    FROM MES.OspReceiptDetail
                                    WHERE OsrDetailId = @OsrDetailId";
                            dynamicParameters.Add("OsrDetailId", OsrDetailId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("託外入庫詳細資料有誤!");

                            string AvailableDate = new DateTime().ToString("yyyy-MM-dd");
                            string ReCheckDate = new DateTime().ToString("yyyy-MM-dd");
                            int ReceiptPackageQty = -1;
                            int AcceptancePackageQty = -1;
                            double ReceiptExpense = -1;
                            string ProjectCode = "";
                            string PaymentHold = "";
                            string QcStatus = "";
                            string ReturnCode = "";
                            string ConfirmStatus = "";
                            string CloseStatus = "";
                            string ReNewStatus = "";
                            string CostEntry = "";
                            string ExpenseEntry = "";
                            int? ConfirmUserId = -1;
                            string ConfirmUser = "";
                            string UrgentMtl = "";
                            string ApproveStatus = "";
                            string TransferStatus = "";
                            string ProcessStatus = "";
                            int OriAcceptQty = -1;
                            double OriOrigAmount = -1;
                            double OriOrigTaxAmt = -1;
                            double OriPreTaxAmt = -1;
                            double OriTaxAmt = -1;
                            double OriOrigPreTaxAmt = -1;
                            foreach (var item in result)
                            {
                                AvailableDate = item.AvailableDate.ToString("yyyy-MM-dd");
                                ReCheckDate = item.ReCheckDate.ToString("yyyy-MM-dd");
                                ReceiptPackageQty = item.ReceiptPackageQty;
                                AcceptancePackageQty = item.AcceptancePackageQty;
                                ReceiptExpense = item.ReceiptExpense;
                                ProjectCode = item.ProjectCode;
                                PaymentHold = item.PaymentHold;
                                QcStatus = item.QcStatus;
                                ReturnCode = item.ReturnCode;
                                ConfirmStatus = item.ConfirmStatus;
                                CloseStatus = item.CloseStatus;
                                ReNewStatus = item.ReNewStatus;
                                CostEntry = item.CostEntry;
                                ExpenseEntry = item.ExpenseEntry;
                                ConfirmUserId = item.ConfirmUserId != -1 ? item.ConfirmUserId : (int?)null;
                                ConfirmUser = item.ConfirmUser;
                                UrgentMtl = item.UrgentMtl;
                                ApproveStatus = item.ApproveStatus;
                                TransferStatus = item.TransferStatus;
                                ProcessStatus = item.ProcessStatus;
                                OsrSeq = item.OsrSeq;
                                OriAcceptQty = item.OriAcceptQty;
                                OriOrigAmount = item.OriOrigAmount;
                                OriOrigTaxAmt = item.OriOrigTaxAmt;
                                OriPreTaxAmt = item.OriPreTaxAmt;
                                OriTaxAmt = item.OriTaxAmt;
                                OriOrigPreTaxAmt = item.OriOrigPreTaxAmt;
                            }
                            #endregion

                            #region //判斷託外入庫資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT OsrErpPrefix, OsrErpNo, ReceiptDate, ReserveTaxCode, ConfirmStatus
                                    , OrigTax OriOrigTax, Quantity OriQuantity, PretaxAmount OriPretaxAmount
                                    , Quantity OriQuantity, PretaxAmount OriPretaxAmount, TaxAmount OriTaxAmount
                                    , CurrencyCode, Exchange
                                    FROM MES.OspReceipt
                                    WHERE OsrId = @OsrId";
                            dynamicParameters.Add("OsrId", OsrId);

                            var result2 = sqlConnection.Query(sql, dynamicParameters);
                            if (result2.Count() <= 0) throw new SystemException("託外入庫資料有誤!");

                            string OsrErpPrefix = "";
                            string OsrErpNo = "";
                            string ReceiptDate = new DateTime().ToString("yyyy-MM-dd");
                            string ReserveTaxCode = "";
                            string OriCurrencyCode = "";
                            double Exchange = 0;
                            foreach (var item in result2)
                            {
                                if (item.ConfirmStatus != "N") throw new SystemException("託外入庫單已確認，無法更改!");
                                OsrErpPrefix = item.OsrErpPrefix;
                                OsrErpNo = item.OsrErpNo;
                                ReceiptDate = item.ReceiptDate.ToString("yyyy-MM-dd");
                                ReserveTaxCode = item.ReserveTaxCode;
                                OriCurrencyCode = item.CurrencyCode;
                                Exchange = item.Exchange;
                            }
                            #endregion

                            #region //判斷託外生產詳細資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.MoId, a.MoProcessId, a.ProcessCode, a.OspQty
                                    , c.MtlItemId, c.UomId
                                    FROM MES.OspDetail a
                                    INNER JOIN MES.ManufactureOrder b ON a.MoId = b.MoId
                                    INNER JOIN MES.WipOrder c ON b.WoId = c.WoId
                                    WHERE a.OspDetailId = @OspDetailId";
                            dynamicParameters.Add("OspDetailId", OspDetailId);

                            var result3 = sqlConnection.Query(sql, dynamicParameters);
                            if (result3.Count() <= 0) throw new SystemException("託外生產詳細資料有誤!");

                            int MtlItemId = -1;
                            int UomId = -1;
                            int MoId = -1;
                            int MoProcessId = -1;
                            string ProcessCode = "";
                            foreach (var item in result3)
                            {
                                if (ReceiptQty > item.OspQty)
                                {
                                    throw new SystemException("進貨數量不可超過原託外數量!!");
                                }

                                MtlItemId = item.MtlItemId;
                                UomId = item.UomId;
                                MoId = item.MoId;
                                MoProcessId = item.MoProcessId;
                                ProcessCode = item.ProcessCode;
                            }
                            #endregion

                            #region //判斷庫別資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM SCM.Inventory
                                    WHERE InventoryId = @InventoryId";
                            dynamicParameters.Add("InventoryId", InventoryId);

                            var result4 = sqlConnection.Query(sql, dynamicParameters);
                            if (result4.Count() <= 0) throw new SystemException("庫別資料有誤!");
                            #endregion

                            #region //判斷計價單位資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM PDM.UnitOfMeasure
                                    WHERE UomId = @UomId";
                            dynamicParameters.Add("UomId", AvailableUom);

                            var result5 = sqlConnection.Query(sql, dynamicParameters);
                            if (result5.Count() <= 0) throw new SystemException("計價單位資料有誤!");
                            #endregion

                            #region //檢查【託外入庫單 + 托外生產詳細資料】是否重複
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.OspReceiptDetail a
                                    WHERE a.OsrId = @OsrId
                                    AND a.OspDetailId = @OspDetailId
                                    AND a.OsrDetailId != @OsrDetailId";
                            dynamicParameters.Add("OsrId", OsrId);
                            dynamicParameters.Add("OspDetailId", OspDetailId);
                            dynamicParameters.Add("OsrDetailId", OsrDetailId);

                            var result7 = sqlConnection.Query(sql, dynamicParameters);
                            if (result7.Count() > 0) throw new SystemException("【託外入庫單 + 托外生產詳細資料】重複!");
                            #endregion

                            #region //檢查【託外入庫單 + 序號】是否重複
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.OspReceiptDetail a
                                    WHERE a.OsrId = @OsrId
                                    AND a.OsrSeq = @OsrSeq
                                    AND a.OsrDetailId != @OsrDetailId";
                            dynamicParameters.Add("OsrId", OsrId);
                            dynamicParameters.Add("OsrSeq", OsrSeq);
                            dynamicParameters.Add("OsrDetailId", OsrDetailId);

                            var result8 = sqlConnection.Query(sql, dynamicParameters);
                            if (result8.Count() > 0) throw new SystemException("【託外入庫單 + 序號】重複!");
                            #endregion

                            #region //UPDATE MES.OspReceiptDetail
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE MES.OspReceiptDetail SET
                                    OsrId = @OsrId,
                                    OspDetailId = @OspDetailId,
                                    OsrErpPrefix = @OsrErpPrefix,
                                    OsrErpNo = @OsrErpNo,
                                    OsrSeq = @OsrSeq,
                                    MtlItemId = @MtlItemId,
                                    ReceiptQty = @ReceiptQty,
                                    UomId = @UomId,
                                    InventoryId = @InventoryId,
                                    LotNumber = @LotNumber,
                                    AvailableDate = @AvailableDate,
                                    ReCheckDate = @ReCheckDate,
                                    MoId = @MoId,
                                    MoProcessId = @MoProcessId,
                                    ProcessCode = @ProcessCode,
                                    ReceiptPackageQty = @ReceiptPackageQty,
                                    AcceptancePackageQty = @AcceptancePackageQty,
                                    AcceptanceDate = @AcceptanceDate,
                                    AcceptQty = @AcceptQty,
                                    AvailableQty = @AvailableQty,
                                    ScriptQty = @ScriptQty,
                                    ReturnQty = @ReturnQty,
                                    AvailableUom = @AvailableUom,
                                    OrigUnitPrice = @OrigUnitPrice,
                                    OrigAmount = @OrigAmount,
                                    OrigDiscountAmt = @OrigDiscountAmt,
                                    ReceiptExpense = @ReceiptExpense,
                                    DiscountDescription = @DiscountDescription,
                                    ProjectCode = @ProjectCode,
                                    PaymentHold = @PaymentHold,
                                    Overdue = @Overdue,
                                    QcStatus = @QcStatus,
                                    ReturnCode = @ReturnCode,
                                    ConfirmStatus = @ConfirmStatus,
                                    CloseStatus = @CloseStatus,
                                    ReNewStatus = @ReNewStatus,
                                    Remark = @Remark,
                                    CostEntry = @CostEntry,
                                    ExpenseEntry = @ExpenseEntry,
                                    ConfirmUserId = @ConfirmUserId,
                                    ConfirmUser = @ConfirmUser,
                                    OrigPreTaxAmt = @OrigPreTaxAmt,
                                    OrigTaxAmt = @OrigTaxAmt,
                                    PreTaxAmt = @PreTaxAmt,
                                    TaxAmt = @TaxAmt,
                                    UrgentMtl = @UrgentMtl,
                                    ReserveTaxCode = @ReserveTaxCode,
                                    ApproveStatus = @ApproveStatus,
                                    TransferStatus = @TransferStatus,
                                    ProcessStatus = @ProcessStatus,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE OsrDetailId = @OsrDetailId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    OsrId,
                                    OspDetailId,
                                    OsrErpPrefix,
                                    OsrErpNo,
                                    OsrSeq,
                                    MtlItemId,
                                    ReceiptQty,
                                    UomId,
                                    InventoryId,
                                    LotNumber,
                                    AvailableDate,
                                    ReCheckDate,
                                    MoId,
                                    MoProcessId,
                                    ProcessCode,
                                    ReceiptPackageQty,
                                    AcceptancePackageQty,
                                    AcceptanceDate,
                                    AcceptQty,
                                    AvailableQty,
                                    ScriptQty,
                                    ReturnQty,
                                    AvailableUom,
                                    OrigUnitPrice,
                                    OrigAmount,
                                    OrigDiscountAmt,
                                    ReceiptExpense,
                                    DiscountDescription,
                                    ProjectCode,
                                    PaymentHold,
                                    Overdue,
                                    QcStatus,
                                    ReturnCode,
                                    ConfirmStatus,
                                    CloseStatus,
                                    ReNewStatus,
                                    Remark,
                                    CostEntry,
                                    ExpenseEntry,
                                    ConfirmUserId,
                                    ConfirmUser,
                                    OrigPreTaxAmt,
                                    OrigTaxAmt,
                                    PreTaxAmt,
                                    TaxAmt,
                                    UrgentMtl,
                                    ReserveTaxCode,
                                    ApproveStatus,
                                    TransferStatus,
                                    ProcessStatus,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    OsrDetailId
                                });

                            int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //取得本幣幣別
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 LTRIM(RTRIM(MA003)) CurrencyCode
                                    FROM CMSMA
                                    WHERE 1=1";

                            var CMSMAResult = sqlConnection2.Query(sql, dynamicParameters);

                            string CurrencyCode = "";
                            foreach (var item in CMSMAResult)
                            {
                                CurrencyCode = item.CurrencyCode;
                            }
                            #endregion

                            #region //本幣小數點取位資料
                            int UnitDecimal = 0;
                            int PriceDecimal = 0;
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.MF003, a.MF004
                                    FROM CMSMF a 
                                    WHERE a.MF001 = @Currency";
                            dynamicParameters.Add("Currency", CurrencyCode);

                            var CMSMFResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (CMSMFResult.Count() <= 0) throw new SystemException("金額小數點取位資料錯誤!!");

                            foreach (var item in CMSMFResult)
                            {
                                UnitDecimal = Convert.ToInt32(item.MF003);
                                PriceDecimal = Convert.ToInt32(item.MF004);
                            }
                            #endregion

                            #region //原幣小數點取位資料
                            int OriUnitDecimal = 0;
                            int OriPriceDecimal = 0;
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.MF003, a.MF004
                                    FROM CMSMF a 
                                    WHERE a.MF001 = @Currency";
                            dynamicParameters.Add("Currency", OriCurrencyCode);

                            var CMSMFResult2 = sqlConnection2.Query(sql, dynamicParameters);

                            if (CMSMFResult2.Count() <= 0) throw new SystemException("金額小數點取位資料錯誤!!");

                            foreach (var item in CMSMFResult2)
                            {
                                OriUnitDecimal = Convert.ToInt32(item.MF003);
                                OriPriceDecimal = Convert.ToInt32(item.MF004);
                            }
                            #endregion

                            #region //計算目前此託外入庫單相關資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT SUM(a.OrigAmount) SumOrigAmount
                                    , SUM(a.OrigTaxAmt) SumOrigTaxAmt
                                    , SUM(a.AcceptQty) SumAcceptQty
                                    , SUM(a.PreTaxAmt) SumPreTaxAmt
                                    , SUM(a.TaxAmt) SumTaxAmt
                                    , SUM(a.OrigPreTaxAmt) SumOrigPreTaxAmt
                                    FROM MES.OspReceiptDetail a
                                    INNER JOIN MES.OspReceipt b ON a.OsrId = b.OsrId
                                    WHERE a.OsrId = @OsrId";
                            dynamicParameters.Add("OsrId", OsrId);
                            var SumAcceptQtyResult = sqlConnection.Query(sql, dynamicParameters);

                            double SumOrigAmount = 1;
                            double SumOrigTaxAmt = 1;
                            int SumAcceptQty = 1;
                            double SumPreTaxAmt = 1;
                            double SumTaxAmt = 1;
                            double SumOrigPreTaxAmt = 1;
                            foreach (var item in SumAcceptQtyResult)
                            {
                                SumOrigAmount = Math.Round(item.SumOrigAmount, OriPriceDecimal);
                                SumOrigTaxAmt = Math.Round(item.SumOrigTaxAmt, OriPriceDecimal);
                                SumAcceptQty = item.SumAcceptQty;
                                SumPreTaxAmt = Math.Round(item.SumPreTaxAmt, PriceDecimal);
                                SumTaxAmt = Math.Round(item.SumTaxAmt, PriceDecimal);
                                SumOrigPreTaxAmt = Math.Round(item.SumOrigPreTaxAmt, OriPriceDecimal);
                            }
                            #endregion

                            #region //更新託外入庫單頭資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE MES.OspReceipt SET
                                    OrigAmount = @SumOrigAmount,
                                    OrigTax = @SumOrigTaxAmt,
                                    Quantity = @SumAcceptQty,
                                    PretaxAmount = @SumPreTaxAmt,
                                    TaxAmount = @SumTaxAmt,
                                    OrigPreTaxAmount = @SumOrigPreTaxAmt,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE OsrId = @OsrId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    SumOrigAmount,
                                    SumOrigTaxAmt,
                                    SumAcceptQty,
                                    SumPreTaxAmt,
                                    SumTaxAmt,
                                    SumOrigPreTaxAmt,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    OsrId
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

        #region //UpdateTransferOspReceipt -- 拋轉託外入庫單資料(不核單) -- Ann 2022-09-21
        public string UpdateTransferOspReceipt(int OsrId = -1)
        {
            try
            {
                List<OspReceipt> ospReceipts = new List<OspReceipt>();
                List<OspReceiptDetail> ospReceiptDetails = new List<OspReceiptDetail>();

                List<MOCTH> mocths = new List<MOCTH>();
                List<MOCTI> moctis = new List<MOCTI>();

                string OsrErpNo = "";
                int rowsAffected = 0;
                string ErpNo = "";
                string UserNo = "";
                string ErpDbName = "";

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (BaseHelper.CheckUserAuthority(CreateBy, CurrentCompany, "A", "OspReceipt", "update", sqlConnection).Equals("N")) throw new SystemException("人員權限檢核有異常，請嘗試重新登入!");

                        #region //確認公司別DB
                        DynamicParameters dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var companyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (companyResult.Count() <= 0) throw new SystemException("公司別錯誤:" + UserNo + ":" + UserLock + ":" + CurrentUser + ":" + CurrentCompany);

                        foreach (var item in companyResult)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            ErpNo = item.ErpNo;
                            ErpDbName = ErpConnectionStrings.Split('=')[2].Split(';')[0];
                        }
                        #endregion

                        #region //查詢UserNo
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserNo
                                    FROM BAS.[User] a
                                    WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", CreateBy);

                        var userResult = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in userResult)
                        {
                            UserNo = item.UserNo;
                        }
                        #endregion

                        #region //取得MES.OspReceipt資料(MOCTH)
                        dynamicParameters = new DynamicParameters();
                        sqlQuery.mainKey = "a.OsrId";
                        sqlQuery.auxKey = "";
                        sqlQuery.columns =
                            @", a.OsrErpPrefix, a.OsrErpNo, FORMAT(a.ReceiptDate, 'yyyyMMdd') ReceiptDate, a.FactoryCode, a.SupplierId, a.SupplierSo, a.CurrencyCode
                            , a.Exchange, a.RowCnt, a.Remark, a.UiNo, a.InvoiceType, FORMAT(a.InvoiceDate, 'yyyyMMdd') InvoiceDate, a.InvoiceNo, a.TaxType, a.DeductType
                            , a.ReserveFlag, a.OrigAmount, a.DeductAmount, a.OrigTax, a.ReceiptAmount, a.Quantity, a.ConfirmStatus, a.ConfirmUserId
                            , a.RenewFlag, a.PrintCnt, a.AutoMaterialBilling, a.OrigPreTaxAmount, FORMAT(a.ApplyYYMM, 'yyyyMM') ApplyYYMM, FORMAT(a.DocumentDate, 'yyyyMMdd') DocumentDate
                            , a.TaxRate, a.PretaxAmount, a.TaxAmount, a.PaymentTerm, a.PackageQuantity, a.SupplierPicking, a.ApproveStatus, a.CompanyId, FORMAT(a.DocumentDate, 'yyyy-MM-dd') MesDocumentDate
                            , a.ReserveTaxCode, a.SendCount, a.NoticeFlag, a.TaxCode, a.TaxExchange, a.QcFlag, a.FlowStatus, FORMAT(a.CreateDate, 'HH:mm:ss') CreateTime
                            , b.UserNo ConfirmUserNo
                            , c.SupplierNo
                            , d.UserNo CreateUserNo";
                        sqlQuery.mainTables =
                            @"FROM MES.OspReceipt a
                            LEFT JOIN BAS.[User] b ON a.ConfirmUserId = b.UserId
                            INNER JOIN SCM.Supplier c ON a.SupplierId = c.SupplierId
                            INNER JOIN BAS.[User] d ON a.CreateBy = d.UserId";
                        sqlQuery.auxTables = "";
                        string queryCondition = @"AND a.OsrId=@OsrId";
                        dynamicParameters.Add("OsrId", OsrId);
                        sqlQuery.conditions = queryCondition;
                        sqlQuery.orderBy = "";

                        ospReceipts = BaseHelper.SqlQuery<OspReceipt>(sqlConnection, dynamicParameters, sqlQuery);
                        if (ospReceipts.Count() <= 0) throw new SystemException("託外入庫資料有誤!");

                        if (ospReceipts[0].InvoiceDate == "19000101") ospReceipts[0].InvoiceDate = "";

                        foreach (var item in ospReceipts)
                        {
                            if (item.ConfirmStatus != "N") throw new SystemException("單據非未確認狀態，無法修改!!");
                            if (item.CompanyId != CurrentCompany) throw new SystemException("單據公司別與使用者公司別不同，請嘗試重新登入!!");
                        }
                        #endregion

                        #region //取得MES.OspReceiptDetail(MOCTI)
                        dynamicParameters = new DynamicParameters();
                        sqlQuery.mainKey = "a.OsrDetailId";
                        sqlQuery.auxKey = "";
                        sqlQuery.columns =
                            @", a.OsrId, a.OspDetailId, a.OsrErpPrefix, a.OsrErpNo, a.OsrSeq, a.MtlItemId, a.ReceiptQty, a.UomId, a.InventoryId
                            , a.LotNumber, FORMAT(a.AvailableDate, 'yyyyMMdd') AvailableDate, FORMAT(a.ReCheckDate, 'yyyyMMdd') ReCheckDate, a.MoId, a.MoProcessId, a.ProcessCode, a.ReceiptPackageQty
                            , a.AcceptancePackageQty, FORMAT(a.AcceptanceDate, 'yyyyMMdd') AcceptanceDate, a.AcceptQty, a.AvailableQty, a.ScriptQty, a.ReturnQty, a.AvailableUom, a.OrigUnitPrice
                            , a.OrigAmount, a.OrigDiscountAmt, a.ReceiptExpense, a.DiscountDescription, a.ProjectCode, a.PaymentHold, a.Overdue, a.QcStatus
                            , a.ReturnCode, a.ConfirmStatus, a.CloseStatus, a.ReNewStatus, a.Remark, a.CostEntry, a.ExpenseEntry, a.ConfirmUserId, a.ConfirmUser, a.OrigPreTaxAmt
                            , a.OrigTaxAmt, a.PreTaxAmt, a.TaxAmt, a.UrgentMtl, a.ReserveTaxCode, a.ApproveStatus, a.TransferStatus, a.ProcessStatus, FORMAT(a.CreateDate, 'HH:mm:ss') CreateTime
                            , b.UserNo CreateUserNo
                            , c.MtlItemNo, c.MtlItemName, c.MtlItemSpec
                            , d.UomNo
                            , e.InventoryNo
                            , g.WoErpPrefix, g.WoErpNo
                            , h.Status MoStatus";
                        sqlQuery.mainTables =
                            @"FROM MES.OspReceiptDetail a
                            INNER JOIN BAS.[User] b ON a.CreateBy = b.UserId
                            INNER JOIN PDM.MtlItem c ON a.MtlItemId = c.MtlItemId
                            INNER JOIN PDM.UnitOfMeasure d ON a.UomId = d.UomId
                            INNER JOIN SCM.Inventory e ON a.InventoryId = e.InventoryId
                            INNER JOIN MES.ManufactureOrder f ON a.MoId = f.MoId
                            INNER JOIN MES.WipOrder g ON f.WoId = g.WoId
                            INNER JOIN MES.ManufactureOrder h ON a.MoId = h.MoId";
                        sqlQuery.auxTables = "";
                        queryCondition = @"AND OsrId = @OsrId";
                        dynamicParameters.Add("OsrId", OsrId);
                        sqlQuery.conditions = queryCondition;
                        sqlQuery.orderBy = "";

                        ospReceiptDetails = BaseHelper.SqlQuery<OspReceiptDetail>(sqlConnection, dynamicParameters, sqlQuery);
                        if (ospReceiptDetails.Count() <= 0) throw new SystemException("託外入庫詳細資料有誤!");
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //審核ERP權限
                            string USR_GROUP = BaseHelper.CheckErpAuthority(UserNo, sqlConnection2, "MOCI06", "CREATE");
                            #endregion

                            #region //檢查ERP單頭是否已核單
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(TH023)) ConfirmStatus
                                    FROM MOCTH
                                    WHERE TH001 = @TH001
                                    AND TH002 = @TH002";
                            dynamicParameters.Add("TH001", ospReceipts[0].OsrErpPrefix);
                            dynamicParameters.Add("TH002", ospReceipts[0].OsrErpNo);

                            var THResult = sqlConnection2.Query(sql, dynamicParameters);

                            foreach (var item in THResult)
                            {
                                if (item.ConfirmStatus != "N") throw new SystemException("此單據已核准，無法再次拋準!!");
                            }
                            #endregion

                            #region //檢查ERP單身是否已核單
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(TI037)) ConfirmStatus
                                    FROM MOCTI
                                    WHERE TI001 = @TI001
                                    AND TI002 = @TI002";
                            dynamicParameters.Add("TI001", ospReceipts[0].OsrErpPrefix);
                            dynamicParameters.Add("TI002", ospReceipts[0].OsrErpNo);

                            var TIResult = sqlConnection2.Query(sql, dynamicParameters);

                            foreach (var item in TIResult)
                            {
                                if (item.ConfirmStatus != "N") throw new SystemException("此單據已核准，無法再次拋準!!");
                            }
                            #endregion

                            #region //查詢ERP資料庫中是否已經有此托外進貨單
                            dynamicParameters = new DynamicParameters();
                            sqlQuery.mainKey = "a.TH001, a.TH002";
                            sqlQuery.auxKey = "";
                            sqlQuery.columns =
                                @", LTRIM(RTRIM(a.COMPANY)) COMPANY, LTRIM(RTRIM(a.MODIFIER)) MODIFIER, LTRIM(RTRIM(a.USR_GROUP)) USR_GROUP, LTRIM(RTRIM(a.MODI_DATE)) MODI_DATE
                                , LTRIM(RTRIM(a.FLAG)) FLAG, LTRIM(RTRIM(a.MODI_TIME)) MODI_TIME, LTRIM(RTRIM(a.TH009)) TH009
                                , LTRIM(RTRIM(a.TH010)) TH010, LTRIM(RTRIM(a.TH015)) TH015, LTRIM(RTRIM(a.TH016)) TH016, LTRIM(RTRIM(a.TH017)) TH017, LTRIM(RTRIM(a.TH018)) TH018
                                , LTRIM(RTRIM(a.TH020)) TH020, LTRIM(RTRIM(a.TH021)) TH021, LTRIM(RTRIM(a.TH022)) TH022, LTRIM(RTRIM(a.TH023)) TH023";
                            sqlQuery.mainTables =
                                @"FROM MOCTH a";
                            sqlQuery.auxTables = "";
                            queryCondition = @"AND a.TH001 = @TH001 AND a.TH002 = @TH002";
                            dynamicParameters.Add("TH001", ospReceipts[0].OsrErpPrefix);
                            dynamicParameters.Add("TH002", ospReceipts[0].OsrErpNo);
                            sqlQuery.conditions = queryCondition;
                            sqlQuery.orderBy = "";

                            mocths = BaseHelper.SqlQuery<MOCTH>(sqlConnection2, dynamicParameters, sqlQuery);
                            #endregion

                            #region //比對ERP關帳日期
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MA013)) MA013
                                    FROM CMSMA";
                            var cmsmaResult = sqlConnection2.Query(sql, dynamicParameters);
                            if (cmsmaResult.Count() < 0) throw new SystemException("EPR關帳日期資料錯誤!!");

                            foreach (var item in cmsmaResult)
                            {
                                string eprDate = item.MA013;
                                string erpYear = eprDate.Substring(0, 4);
                                string erpMonth = eprDate.Substring(4, 2);
                                string erpDay = eprDate.Substring(6, 2);
                                string erpFullDate = erpYear + "-" + erpMonth + "-" + erpDay;
                                DateTime erpDateTime = Convert.ToDateTime(erpFullDate);
                                DateTime DocDateDateTime = Convert.ToDateTime(ospReceipts[0].MesDocumentDate);
                                int compare = DocDateDateTime.CompareTo(erpDateTime);
                                //if (compare <= 0) throw new SystemException("ERP已關帳(" + erpFullDate + ")，無法開立此日期(" + ospReceipts[0].MesDocumentDate + ")之單據!!");
                            }
                            #endregion

                            if (mocths.Count() > 0)
                            {
                                #region //UPDATE ERP TABLE
                                #region //MOCTH 託外入庫單頭
                                foreach (var item in mocths)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE MOCTH SET
                                            COMPANY = @COMPANY,
                                            CREATOR = @CREATOR,
                                            MODIFIER = @MODIFIER,
                                            USR_GROUP = @USR_GROUP,
                                            MODI_DATE = @MODI_DATE,
                                            FLAG += 1,
                                            MODI_TIME = @MODI_TIME,
                                            TH001 = @TH001,
                                            TH002 = @TH002,
                                            TH003 = @TH003,
                                            TH004 = @TH004,
                                            TH005 = @TH005,
                                            TH006 = @TH006,
                                            TH007 = @TH007,
                                            TH008 = @TH008,
                                            TH009 = @TH009,
                                            TH010 = @TH010,
                                            TH011 = @TH011,
                                            TH012 = @TH012,
                                            TH013 = @TH013,
                                            TH014 = @TH014,
                                            TH015 = @TH015,
                                            TH016 = @TH016,
                                            TH017 = @TH017,
                                            TH018 = @TH018,
                                            TH019 = @TH019,
                                            TH020 = @TH020,
                                            TH021 = @TH021,
                                            TH022 = @TH022,
                                            TH023 = @TH023,
                                            TH024 = @TH024,
                                            TH025 = @TH025,
                                            TH026 = @TH026,
                                            TH027 = @TH027,
                                            TH028 = @TH028,
                                            TH029 = @TH029,
                                            TH030 = @TH030,
                                            TH031 = @TH031,
                                            TH032 = @TH032,
                                            TH033 = @TH033,
                                            TH034 = @TH034,
                                            TH035 = @TH035,
                                            TH036 = @TH036,
                                            TH037 = @TH037,
                                            TH038 = @TH038,
                                            TH039 = @TH039,
                                            TH040 = @TH040,
                                            TH041 = @TH041,
                                            TH042 = @TH042,
                                            TH043 = @TH043,
                                            TH044 = @TH044,
                                            TH045 = @TH045,
                                            TH046 = @TH046,
                                            TH047 = @TH047,
                                            TH048 = @TH048,
                                            TH049 = @TH049,
                                            TH050 = @TH050,
                                            UDF01 = @UDF01
                                            WHERE TH001 = @TH001
                                            AND TH002 = @TH002";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            COMPANY = ErpNo,
                                            CREATOR = ospReceipts[0].CreateUserNo,
                                            item.MODIFIER,
                                            USR_GROUP,
                                            item.MODI_DATE,
                                            item.MODI_TIME,
                                            item.TH001,
                                            item.TH002,
                                            TH003 = ospReceipts[0].ReceiptDate,
                                            TH004 = ospReceipts[0].FactoryCode,
                                            TH005 = ospReceipts[0].SupplierNo,
                                            TH006 = ospReceipts[0].SupplierSo,
                                            TH007 = ospReceipts[0].CurrencyCode,
                                            TH008 = ospReceipts[0].Exchange,
                                            TH009 = ospReceipts[0].RowCnt,
                                            TH010 = ospReceipts[0].Remark,
                                            TH011 = ospReceipts[0].UiNo,
                                            TH012 = ospReceipts[0].InvoiceType,
                                            TH013 = ospReceipts[0].InvoiceDate,
                                            TH014 = ospReceipts[0].InvoiceNo,
                                            TH015 = ospReceipts[0].TaxType,
                                            TH016 = ospReceipts[0].DeductType,
                                            TH017 = ospReceipts[0].ReserveFlag,
                                            TH018 = ospReceipts[0].OrigAmount,
                                            TH019 = ospReceipts[0].DeductAmount,
                                            TH020 = ospReceipts[0].OrigTax,
                                            TH021 = ospReceipts[0].ReceiptAmount,
                                            TH022 = ospReceipts[0].Quantity,
                                            TH023 = ospReceipts[0].ConfirmStatus,
                                            TH024 = ospReceipts[0].RenewFlag,
                                            TH025 = ospReceipts[0].PrintCnt,
                                            TH026 = ospReceipts[0].AutoMaterialBilling,
                                            TH027 = ospReceipts[0].OrigPreTaxAmount,
                                            TH028 = ospReceipts[0].ApplyYYMM,
                                            TH029 = ospReceipts[0].DocumentDate,
                                            TH030 = ospReceipts[0].TaxRate,
                                            TH031 = ospReceipts[0].PretaxAmount,
                                            TH032 = ospReceipts[0].TaxAmount,
                                            TH033 = ospReceipts[0].PaymentTerm,
                                            TH034 = ospReceipts[0].PackageQuantity,
                                            TH035 = ospReceipts[0].SupplierPicking,
                                            TH036 = ospReceipts[0].ApproveStatus,
                                            TH037 = ospReceipts[0].ReserveTaxCode,
                                            TH038 = ospReceipts[0].SendCount,
                                            item.TH039,
                                            item.TH040,
                                            TH041 = ospReceipts[0].NoticeFlag,
                                            item.TH042,
                                            item.TH043,
                                            TH044 = ospReceipts[0].TaxCode,
                                            item.TH045,
                                            item.TH046,
                                            TH047 = ospReceipts[0].TaxExchange,
                                            item.TH048,
                                            item.TH049,
                                            item.TH050,
                                            UDF01 = "" //費用部門
                                        });

                                    rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);
                                }
                                #endregion

                                #region //MOCTI 託外入庫單身設定
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE MOCTI
                                        WHERE TI001 = @TI001
                                        AND TI002 = @TI002";
                                dynamicParameters.Add("TI001", ospReceipts[0].OsrErpPrefix);
                                dynamicParameters.Add("TI002", ospReceipts[0].OsrErpNo);

                                rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);

                                foreach (var item in ospReceiptDetails)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO MOCTI (COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE, FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID
                                            , TI001, TI002, TI003, TI004, TI005, TI006, TI007, TI008, TI009, TI010, TI011, TI012, TI013, TI014, TI015, TI016, TI017, TI018, TI019, TI020, TI021, TI022
                                            , TI023, TI024, TI025, TI026, TI027, TI028, TI032, TI033, TI034, TI035, TI036, TI037, TI038, TI039, TI040, TI041, TI042, TI043, TI044
                                            , TI045, TI046, TI047, TI048, TI049, TI050, TI051, TI052, TI053, TI054, TI055, TI056, TI057, TI058, TI059, TI060, TI061, TI062, TI063, TI064, TI065, TI066
                                            , TI067, TI068, TI069, TI070, TI071, TI503, TI504, TI505, TI506, TI551, TI552, TI553, TI554, TI555, TI556
                                            , UDF06, UDF07, UDF08, UDF09, UDF10)
                                            VALUES (@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE, @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID, @MODI_TIME, @MODI_AP, @MODI_PRID
                                            , @TI001, @TI002, @TI003, @TI004, @TI005, @TI006, @TI007, @TI008, @TI009, @TI010, @TI011, @TI012, @TI013, @TI014, @TI015, @TI016, @TI017, @TI018, @TI019, @TI020, @TI021, @TI022
                                            , @TI023, @TI024, @TI025, @TI026, @TI027, @TI028, @TI032, @TI033, @TI034, @TI035, @TI036, @TI037, @TI038, @TI039, @TI040, @TI041, @TI042, @TI043, @TI044
                                            , @TI045, @TI046, @TI047, @TI048, @TI049, @TI050, @TI051, @TI052, @TI053, @TI054, @TI055, @TI056, @TI057, @TI058, @TI059, @TI060, @TI061, @TI062, @TI063, @TI064, @TI065, @TI066
                                            , @TI067, @TI068, @TI069, @TI070, @TI071, @TI503, @TI504, @TI505, @TI506, @TI551, @TI552, @TI553, @TI554, @TI555, @TI556
                                            , @UDF06, @UDF07, @UDF08, @UDF09, @UDF10)";

                                    dynamicParameters.AddDynamicParams(
                                      new
                                      {
                                          COMPANY = ErpNo,
                                          CREATOR = ospReceipts[0].CreateUserNo,
                                          USR_GROUP,
                                          CREATE_DATE = DateTime.Now.ToString("yyyyMMdd"),
                                          MODIFIER = "",
                                          MODI_DATE = "",
                                          FLAG = "1",
                                          CREATE_TIME = item.CreateTime,
                                          CREATE_AP = "ERP-API",
                                          CREATE_PRID = "BM",
                                          MODI_TIME = "",
                                          MODI_AP = "",
                                          MODI_PRID = "",
                                          TI001 = item.OsrErpPrefix,
                                          TI002 = item.OsrErpNo,
                                          TI003 = item.OsrSeq,
                                          TI004 = item.MtlItemNo,
                                          TI005 = item.MtlItemName,
                                          TI006 = item.MtlItemSpec,
                                          TI007 = item.ReceiptQty,
                                          TI008 = item.UomNo,
                                          TI009 = "",
                                          TI010 = item.LotNumber,
                                          TI011 = item.AvailableDate,
                                          TI012 = item.ReCheckDate,
                                          TI013 = item.WoErpPrefix,
                                          TI014 = item.WoErpNo,
                                          TI015 = item.ProcessCode,
                                          TI016 = item.ReceiptPackageQty,
                                          TI017 = item.AcceptancePackageQty,
                                          TI018 = item.AcceptanceDate,
                                          TI019 = item.AcceptQty,
                                          TI020 = item.AvailableQty,
                                          TI021 = item.ScriptQty,
                                          TI022 = item.ReturnQty,
                                          TI023 = item.UomNo,
                                          TI024 = item.OrigUnitPrice,
                                          TI025 = item.OrigAmount,
                                          TI026 = item.OrigDiscountAmt,
                                          TI027 = item.ReceiptExpense,
                                          TI028 = item.DiscountDescription,
                                          TI032 = item.ProjectCode,
                                          TI033 = item.PaymentHold,
                                          TI034 = item.Overdue,
                                          TI035 = item.QcStatus,
                                          TI036 = item.ReturnCode,
                                          TI037 = item.ConfirmStatus,
                                          TI038 = item.CloseStatus,
                                          TI039 = item.ReNewStatus,
                                          TI040 = item.Remark,
                                          TI041 = item.CostEntry,
                                          TI042 = item.ExpenseEntry,
                                          TI043 = item.ConfirmUser,
                                          TI044 = item.OrigPreTaxAmt,
                                          TI045 = item.OrigTaxAmt,
                                          TI046 = item.PreTaxAmt,
                                          TI047 = item.TaxAmt,
                                          TI048 = item.UrgentMtl,
                                          TI049 = "",
                                          TI050 = 0,
                                          TI051 = 0,
                                          TI052 = item.ReserveTaxCode,
                                          TI053 = item.ApproveStatus,
                                          TI054 = "N",
                                          TI055 = "0",
                                          TI056 = "2",
                                          TI057 = "",
                                          TI058 = 0,
                                          TI059 = 0,
                                          TI060 = "",
                                          TI061 = "",
                                          TI062 = "",
                                          TI063 = "",
                                          TI064 = "",
                                          TI065 = 0,
                                          TI066 = 0,
                                          TI067 = 0,
                                          TI068 = 0,
                                          TI069 = "",
                                          TI070 = "",
                                          TI071 = 0,
                                          TI503 = 0,
                                          TI504 = 0,
                                          TI505 = 0,
                                          TI506 = 0,
                                          TI551 = 0,
                                          TI552 = 0,
                                          TI553 = 0,
                                          TI554 = 0,
                                          TI555 = 0,
                                          TI556 = 0,
                                          UDF06 = 0,
                                          UDF07 = 0,
                                          UDF08 = 0,
                                          UDF09 = 0,
                                          UDF10 = 0
                                      });
                                    var insertResult = sqlConnection2.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult.Count();
                                }
                                #endregion
                                #endregion
                            }
                            else
                            {
                                int currentNum = 0, yearLength = 0, lineLength = 0;
                                string encode = "", paymentTerm = "", factory = "";
                                DateTime referenceTime = default(DateTime);

                                referenceTime = DateTime.ParseExact(ospReceipts[0].DocumentDate, "yyyyMMdd", CultureInfo.InvariantCulture);

                                #region //單據設定
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.MQ004, a.MQ005, a.MQ006, a.MQ044
                                        FROM CMSMQ a
                                        WHERE a.COMPANY = @CompanyNo
                                        AND a.MQ001 = @ErpPrefix";
                                dynamicParameters.Add("CompanyNo", ErpNo);
                                dynamicParameters.Add("ErpPrefix", ospReceipts[0].OsrErpPrefix);

                                var resultDocSetting = sqlConnection2.Query(sql, dynamicParameters);
                                if (resultDocSetting.Count() <= 0) throw new SystemException("ERP單據設定資料不存在!");

                                foreach (var item in resultDocSetting)
                                {
                                    encode = item.MQ004; //編碼方式
                                    yearLength = Convert.ToInt32(item.MQ005); //年碼數
                                    lineLength = Convert.ToInt32(item.MQ006); //流水號碼數
                                }
                                #endregion

                                #region //單號取號
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT CONVERT(int, RIGHT(ISNULL(MAX(RTRIM(LTRIM(TH002))), '" + new string('0', lineLength) + @"'), " + lineLength + @")) + 1 CurrentNum
                                        FROM MOCTH
                                        WHERE TH001 = @ErpPrefix";
                                dynamicParameters.Add("ErpPrefix", ospReceipts[0].OsrErpPrefix);

                                #region //編碼方式
                                string dateFormat = "";
                                switch (encode)
                                {
                                    case "1": //日編
                                        dateFormat = new string('y', yearLength) + "MMdd";
                                        sql += @" AND RTRIM(LTRIM(TH002)) LIKE @ReferenceTime + '" + new string('_', lineLength) + @"'";
                                        dynamicParameters.Add("ReferenceTime", referenceTime.ToString(dateFormat));
                                        OsrErpNo = referenceTime.ToString(dateFormat);
                                        break;
                                    case "2": //月編
                                        dateFormat = new string('y', yearLength) + "MM";
                                        sql += @" AND RTRIM(LTRIM(TH002)) LIKE @ReferenceTime + '" + new string('_', lineLength) + @"'";
                                        dynamicParameters.Add("ReferenceTime", referenceTime.ToString(dateFormat));
                                        OsrErpNo = referenceTime.ToString(dateFormat);
                                        break;
                                    case "3": //流水號
                                        break;
                                    case "4": //手動編號
                                        break;
                                    default:
                                        throw new SystemException("編碼方式錯誤!");
                                }
                                #endregion

                                currentNum = sqlConnection2.QueryFirstOrDefault(sql, dynamicParameters).CurrentNum;
                                OsrErpNo += string.Format("{0:" + new string('0', lineLength) + "}", currentNum);
                                ospReceipts[0].OsrErpNo = OsrErpNo;
                                #endregion

                                #region //找單據日期託外入庫單數
                                //dynamicParameters = new DynamicParameters();
                                //string nowDate = DateTime.Now.ToString("yyyyMMdd");
                                //sql = @"SELECT CONVERT(int, RIGHT(ISNULL(MAX(RTRIM(LTRIM(TH002))), '000'), 3)) + 1 CurrentNum
                                //        FROM MOCTH
                                //        WHERE TH001 = @TH001
                                //        AND TH002 LIKE '%" + ospReceipts[0].DocumentDate + "%'";
                                //dynamicParameters.Add("TH001", ospReceipts[0].OsrErpPrefix);

                                //var result2 = sqlConnection2.Query(sql, dynamicParameters);

                                //int CurrentNum = 0;
                                //foreach (var item in result2)
                                //{
                                //    CurrentNum = item.CurrentNum;
                                //}

                                //string seq = CurrentNum.ToString().PadLeft(3, '0');
                                //OsrErpNo = ospReceipts[0].DocumentDate + seq;
                                //ospReceipts[0].OsrErpNo = OsrErpNo;
                                #endregion

                                #region //MOCTH 託外入庫單單頭
                                foreach (var item in ospReceipts)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO MOCTH (COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE, FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID
                                            , TH001, TH002, TH003, TH004, TH005, TH006, TH007, TH008, TH009, TH010, TH011, TH012, TH013, TH014, TH015, TH016, TH017, TH018, TH019, TH020, TH021, TH022
                                            , TH023, TH024, TH025, TH026, TH027, TH028, TH029, TH030, TH031, TH032, TH033, TH034, TH035, TH036, TH037, TH038, TH039, TH040, TH041, TH042, TH043
                                            , TH044, TH045, TH046, TH047, TH048, TH049, TH050, UDF01)
                                            VALUES (@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE, @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID, @MODI_TIME, @MODI_AP, @MODI_PRID
                                            , @TH001, @TH002, @TH003, @TH004, @TH005, @TH006, @TH007, @TH008, @TH009, @TH010, @TH011, @TH012, @TH013, @TH014, @TH015, @TH016, @TH017, @TH018, @TH019, @TH020, @TH021, @TH022
                                            , @TH023, @TH024, @TH025, @TH026, @TH027, @TH028, @TH029, @TH030, @TH031, @TH032, @TH033, @TH034, @TH035, @TH036, @TH037, @TH038, @TH039, @TH040, @TH041, @TH042, @TH043
                                            , @TH044, @TH045, @TH046, @TH047, @TH048, @TH049, @TH050, @UDF01)";

                                    dynamicParameters.AddDynamicParams(
                                      new
                                      {
                                          COMPANY = ErpNo,
                                          CREATOR = ospReceipts[0].CreateUserNo,
                                          USR_GROUP,
                                          CREATE_DATE = DateTime.Now.ToString("yyyyMMdd"),
                                          MODIFIER = "",
                                          MODI_DATE = "",
                                          FLAG = "1",
                                          CREATE_TIME = item.CreateTime,
                                          CREATE_AP = "ERP-API",
                                          CREATE_PRID = "BM",
                                          MODI_TIME = "",
                                          MODI_AP = "",
                                          MODI_PRID = "",
                                          TH001 = item.OsrErpPrefix,
                                          TH002 = item.OsrErpNo,
                                          TH003 = item.ReceiptDate,
                                          TH004 = item.FactoryCode,
                                          TH005 = item.SupplierNo,
                                          TH006 = item.SupplierSo,
                                          TH007 = item.CurrencyCode,
                                          TH008 = item.Exchange,
                                          TH009 = item.RowCnt,
                                          TH010 = item.Remark,
                                          TH011 = item.UiNo,
                                          TH012 = item.InvoiceType,
                                          TH013 = item.InvoiceDate,
                                          TH014 = item.InvoiceNo,
                                          TH015 = item.TaxType,
                                          TH016 = item.DeductType,
                                          TH017 = item.ReserveFlag,
                                          TH018 = item.OrigAmount,
                                          TH019 = item.DeductAmount,
                                          TH020 = item.OrigTax,
                                          TH021 = item.ReceiptAmount,
                                          TH022 = item.Quantity,
                                          TH023 = item.ConfirmStatus,
                                          TH024 = item.RenewFlag,
                                          TH025 = item.PrintCnt,
                                          TH026 = item.AutoMaterialBilling,
                                          TH027 = item.OrigPreTaxAmount,
                                          TH028 = item.ApplyYYMM,
                                          TH029 = item.DocumentDate,
                                          TH030 = item.TaxRate,
                                          TH031 = item.PretaxAmount,
                                          TH032 = item.TaxAmount,
                                          TH033 = item.PaymentTerm,
                                          TH034 = item.PackageQuantity,
                                          TH035 = item.SupplierPicking,
                                          TH036 = item.ApproveStatus,
                                          TH037 = item.ReserveTaxCode,
                                          TH038 = item.SendCount,
                                          TH039 = 0,
                                          TH040 = 0,
                                          TH041 = item.NoticeFlag,
                                          TH042 = "",
                                          TH043 = "",
                                          TH044 = item.TaxCode,
                                          TH045 = "",
                                          TH046 = "",
                                          TH047 = item.TaxExchange,
                                          TH048 = 0,
                                          TH049 = 0,
                                          TH050 = "",
                                          UDF01 = "" //費用部門
                                      });
                                    var insertResult = sqlConnection2.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult.Count();
                                }
                                #endregion

                                #region //MOCTI 託外入庫單身
                                foreach (var item in ospReceiptDetails)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO MOCTI (COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE, FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID
                                            , TI001, TI002, TI003, TI004, TI005, TI006, TI007, TI008, TI009, TI010, TI011, TI012, TI013, TI014, TI015, TI016, TI017, TI018, TI019, TI020, TI021, TI022
                                            , TI023, TI024, TI025, TI026, TI027, TI028, TI032, TI033, TI034, TI035, TI036, TI037, TI038, TI039, TI040, TI041, TI042, TI043, TI044
                                            , TI045, TI046, TI047, TI048, TI049, TI050, TI051, TI052, TI053, TI054, TI055, TI056, TI057, TI058, TI059, TI060, TI061, TI062, TI063, TI064, TI065, TI066
                                            , TI067, TI068, TI069, TI070, TI071, TI503, TI504, TI505, TI506, TI551, TI552, TI553, TI554, TI555, TI556
                                            , UDF06, UDF07, UDF08, UDF09, UDF10)
                                            VALUES (@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE, @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID, @MODI_TIME, @MODI_AP, @MODI_PRID
                                            , @TI001, @TI002, @TI003, @TI004, @TI005, @TI006, @TI007, @TI008, @TI009, @TI010, @TI011, @TI012, @TI013, @TI014, @TI015, @TI016, @TI017, @TI018, @TI019, @TI020, @TI021, @TI022
                                            , @TI023, @TI024, @TI025, @TI026, @TI027, @TI028, @TI032, @TI033, @TI034, @TI035, @TI036, @TI037, @TI038, @TI039, @TI040, @TI041, @TI042, @TI043, @TI044
                                            , @TI045, @TI046, @TI047, @TI048, @TI049, @TI050, @TI051, @TI052, @TI053, @TI054, @TI055, @TI056, @TI057, @TI058, @TI059, @TI060, @TI061, @TI062, @TI063, @TI064, @TI065, @TI066
                                            , @TI067, @TI068, @TI069, @TI070, @TI071, @TI503, @TI504, @TI505, @TI506, @TI551, @TI552, @TI553, @TI554, @TI555, @TI556
                                            , @UDF06, @UDF07, @UDF08, @UDF09, @UDF10)";

                                    dynamicParameters.AddDynamicParams(
                                      new
                                      {
                                          COMPANY = ErpNo,
                                          CREATOR = ospReceipts[0].CreateUserNo,
                                          USR_GROUP,
                                          CREATE_DATE = DateTime.Now.ToString("yyyyMMdd"),
                                          MODIFIER = "",
                                          MODI_DATE = "",
                                          FLAG = "1",
                                          CREATE_TIME = item.CreateTime,
                                          CREATE_AP = "ERP-API",
                                          CREATE_PRID = "BM",
                                          MODI_TIME = "",
                                          MODI_AP = "",
                                          MODI_PRID = "",
                                          TI001 = ospReceipts[0].OsrErpPrefix,
                                          TI002 = ospReceipts[0].OsrErpNo,
                                          TI003 = item.OsrSeq,
                                          TI004 = item.MtlItemNo,
                                          TI005 = item.MtlItemName,
                                          TI006 = item.MtlItemSpec,
                                          TI007 = item.ReceiptQty,
                                          TI008 = item.UomNo,
                                          TI009 = "",
                                          TI010 = item.LotNumber,
                                          TI011 = item.AvailableDate,
                                          TI012 = item.ReCheckDate,
                                          TI013 = item.WoErpPrefix,
                                          TI014 = item.WoErpNo,
                                          TI015 = item.ProcessCode,
                                          TI016 = item.ReceiptPackageQty,
                                          TI017 = item.AcceptancePackageQty,
                                          TI018 = item.AcceptanceDate,
                                          TI019 = item.AcceptQty,
                                          TI020 = item.AvailableQty,
                                          TI021 = item.ScriptQty,
                                          TI022 = item.ReturnQty,
                                          TI023 = item.UomNo,
                                          TI024 = item.OrigUnitPrice,
                                          TI025 = item.OrigAmount,
                                          TI026 = item.OrigDiscountAmt,
                                          TI027 = item.ReceiptExpense,
                                          TI028 = item.DiscountDescription,
                                          TI032 = item.ProjectCode,
                                          TI033 = item.PaymentHold,
                                          TI034 = item.Overdue,
                                          TI035 = item.QcStatus,
                                          TI036 = item.ReturnCode,
                                          TI037 = item.ConfirmStatus,
                                          TI038 = item.CloseStatus,
                                          TI039 = item.ReNewStatus,
                                          TI040 = item.Remark,
                                          TI041 = item.CostEntry,
                                          TI042 = item.ExpenseEntry,
                                          TI043 = item.ConfirmUser,
                                          TI044 = item.OrigPreTaxAmt,
                                          TI045 = item.OrigTaxAmt,
                                          TI046 = item.PreTaxAmt,
                                          TI047 = item.TaxAmt,
                                          TI048 = item.UrgentMtl,
                                          TI049 = "",
                                          TI050 = 0,
                                          TI051 = 0,
                                          TI052 = item.ReserveTaxCode,
                                          TI053 = item.ApproveStatus,
                                          TI054 = "N",
                                          TI055 = "0",
                                          TI056 = "2",
                                          TI057 = "",
                                          TI058 = 0,
                                          TI059 = 0,
                                          TI060 = "",
                                          TI061 = "",
                                          TI062 = "",
                                          TI063 = "",
                                          TI064 = "",
                                          TI065 = 0,
                                          TI066 = 0,
                                          TI067 = 0,
                                          TI068 = 0,
                                          TI069 = "",
                                          TI070 = "",
                                          TI071 = 0,
                                          TI503 = 0,
                                          TI504 = 0,
                                          TI505 = 0,
                                          TI506 = 0,
                                          TI551 = 0,
                                          TI552 = 0,
                                          TI553 = 0,
                                          TI554 = 0,
                                          TI555 = 0,
                                          TI556 = 0,
                                          UDF06 = 0,
                                          UDF07 = 0,
                                          UDF08 = 0,
                                          UDF09 = 0,
                                          UDF10 = 0
                                      });
                                    var insertResult = sqlConnection2.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult.Count();
                                }
                                #endregion
                            }

                            #region //將ERP單據資料寫回MES
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE MES.OspReceipt SET
                                    OsrErpNo = @OsrErpNo,
                                    TransferStatus = 'Y',
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE OsrId = @OsrId";
                            dynamicParameters.AddDynamicParams(
                              new
                              {
                                  ospReceipts[0].OsrErpNo,
                                  LastModifiedDate,
                                  LastModifiedBy,
                                  OsrId
                              });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE MES.OspReceiptDetail SET
                                    OsrErpNo = @OsrErpNo,
                                    TransferStatus = 'Y',
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE OsrId = @OsrId";
                            dynamicParameters.AddDynamicParams(
                              new
                              {
                                  ospReceipts[0].OsrErpNo,
                                  LastModifiedDate,
                                  LastModifiedBy,
                                  OsrId
                              });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                msg = "拋轉成功!"
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

        #region //UpdateConfirmOspReceipt -- 核准託外入庫單資料 -- Ann 2023-03-20
        public string UpdateConfirmOspReceipt(int OsrId = -1)
        {
            try
            {
                List<OspReceipt> ospReceipts = new List<OspReceipt>();
                List<OspReceiptDetail> ospReceiptDetails = new List<OspReceiptDetail>();

                List<MOCTH> mocths = new List<MOCTH>();
                List<MOCTI> moctis = new List<MOCTI>();

                int rowsAffected = 0;
                string ErpNo = "";
                string UserNo = "";
                string ErpDbName = "";

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (BaseHelper.CheckUserAuthority(CreateBy, CurrentCompany, "A", "OspReceipt", "confirm", sqlConnection).Equals("N")) throw new SystemException("人員權限檢核有異常，請嘗試重新登入!");

                        #region //確認公司別DB
                        DynamicParameters dynamicParameters = new DynamicParameters();
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
                            ErpDbName = ErpConnectionStrings.Split('=')[2].Split(';')[0];
                        }
                        #endregion

                        #region //查詢UserNo
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserNo
                                FROM BAS.[User] a
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", CreateBy);

                        var userResult = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in userResult)
                        {
                            UserNo = item.UserNo;
                        }
                        #endregion

                        #region //取得MES.OspReceipt資料(MOCTH)
                        dynamicParameters = new DynamicParameters();
                        sqlQuery.mainKey = "a.OsrId";
                        sqlQuery.auxKey = "";
                        sqlQuery.columns =
                            @", a.OsrErpPrefix, a.OsrErpNo, FORMAT(a.ReceiptDate, 'yyyyMMdd') ReceiptDate, a.FactoryCode, a.SupplierId, a.SupplierSo, a.CurrencyCode
                            , a.Exchange, a.RowCnt, a.Remark, a.UiNo, a.InvoiceType, FORMAT(a.InvoiceDate, 'yyyyMMdd') InvoiceDate, a.InvoiceNo, a.TaxType, a.DeductType
                            , a.ReserveFlag, a.OrigAmount, a.DeductAmount, a.OrigTax, a.ReceiptAmount, a.Quantity, a.ConfirmStatus, a.ConfirmUserId
                            , a.RenewFlag, a.PrintCnt, a.AutoMaterialBilling, a.OrigPreTaxAmount, FORMAT(a.ApplyYYMM, 'yyyyMM') ApplyYYMM, FORMAT(a.DocumentDate, 'yyyyMMdd') DocumentDate
                            , a.TaxRate, a.PretaxAmount, a.TaxAmount, a.PaymentTerm, a.PackageQuantity, a.SupplierPicking, a.ApproveStatus, a.CompanyId
                            , a.ReserveTaxCode, a.SendCount, a.NoticeFlag, a.TaxCode, a.TaxExchange, a.QcFlag, a.FlowStatus, FORMAT(a.CreateDate, 'HH:mm:ss') CreateTime
                            , b.UserNo ConfirmUserNo
                            , c.SupplierNo
                            , d.UserNo CreateUserNo";
                        sqlQuery.mainTables =
                            @"FROM MES.OspReceipt a
                            LEFT JOIN BAS.[User] b ON a.ConfirmUserId = b.UserId
                            INNER JOIN SCM.Supplier c ON a.SupplierId = c.SupplierId
                            INNER JOIN BAS.[User] d ON a.CreateBy = d.UserId";
                        sqlQuery.auxTables = "";
                        string queryCondition = @"AND a.OsrId=@OsrId";
                        dynamicParameters.Add("OsrId", OsrId);
                        sqlQuery.conditions = queryCondition;
                        sqlQuery.orderBy = "";

                        ospReceipts = BaseHelper.SqlQuery<OspReceipt>(sqlConnection, dynamicParameters, sqlQuery);
                        if (ospReceipts.Count() <= 0) throw new SystemException("託外入庫資料有誤!");

                        if (ospReceipts[0].InvoiceDate == "19000101") ospReceipts[0].InvoiceDate = "";

                        foreach (var item in ospReceipts)
                        {
                            if (item.ConfirmStatus != "N") throw new SystemException("單據非未確認狀態，無法重複核單!!");
                            if (item.CompanyId != CurrentCompany) throw new SystemException("單據公司別與使用者公司別不同，請嘗試重新登入!!");
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //審核ERP權限
                            string USR_GROUP = BaseHelper.CheckErpAuthority(UserNo, sqlConnection2, "MOCI06", "CONFIRM");
                            #endregion

                            #region //確認並拋轉ERP
                            string COMPANY = ErpDbName;
                            string TYPE = "MOCI06";
                            string DOC = ospReceipts[0].OsrErpPrefix;
                            string DOC_NO = ospReceipts[0].OsrErpNo;
                            string TRANS = "1";
                            string USER = UserNo;
                            string DATE = DateTime.Now.ToString("yyyyMMdd");

                            var result = BaseHelper.TransferErpAPI(COMPANY, TYPE, DOC, DOC_NO, TRANS, USER, DATE);
                            #endregion

                            #region //若成功則將相關資料寫回MES
                            var resultJson = JToken.Parse(result);
                            string CODE = resultJson["CODE"].ToString();
                            string ERROR_CODE = resultJson["ERROR_CODE"].ToString(); //錯誤代碼
                            if (CODE == "200")
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE MES.OspReceipt SET
                                        ConfirmStatus = @ConfirmStatus,
                                        ConfirmUserId = @ConfirmUserId,
                                        ApproveStatus = @ApproveStatus,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE OsrId = @OsrId";
                                dynamicParameters.AddDynamicParams(
                                  new
                                  {
                                      ConfirmStatus = "Y",
                                      ConfirmUserId = CreateBy,
                                      ApproveStatus = "3",
                                      LastModifiedDate,
                                      LastModifiedBy,
                                      OsrId
                                  });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE MES.OspReceiptDetail SET
                                        ConfirmStatus = @ConfirmStatus,
                                        ConfirmUserId = @ConfirmUserId,
                                        ConfirmUser = @ConfirmUser,
                                        ApproveStatus = @ApproveStatus,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE OsrId = @OsrId";
                                dynamicParameters.AddDynamicParams(
                                  new
                                  {
                                      ConfirmStatus = "Y",
                                      ConfirmUserId = CreateBy,
                                      ConfirmUser = ospReceipts[0].ConfirmUserNo,
                                      ApproveStatus = "3",
                                      LastModifiedDate,
                                      LastModifiedBy,
                                      OsrId
                                  });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                                #region //Response
                                jsonResponse = JObject.FromObject(new
                                {
                                    status = "success",
                                    msg = "拋轉成功!"
                                });
                                #endregion
                            }
                            else
                            {
                                throw new SystemException("確認失敗，請聯繫資訊人員進行問題排解!");
                            }
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

        #region //UpdateTACOspReceipt -- 拋轉並核准託外入庫單資料 -- Ann 2023-03-20
        public string UpdateTACOspReceipt(int OsrId = -1)
        {
            try
            {
                List<OspReceipt> ospReceipts = new List<OspReceipt>();
                List<OspReceiptDetail> ospReceiptDetails = new List<OspReceiptDetail>();

                List<MOCTH> mocths = new List<MOCTH>();
                List<MOCTI> moctis = new List<MOCTI>();

                string OsrErpNo = "";
                int rowsAffected = 0;
                string ErpNo = "";
                string UserNo = "";
                string ErpDbName = "";

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認公司別DB
                        DynamicParameters dynamicParameters = new DynamicParameters();
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
                            ErpDbName = ErpConnectionStrings.Split('=')[2].Split(';')[0];
                        }
                        #endregion

                        #region //查詢UserNo
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserNo
                                    FROM BAS.[User] a
                                    WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", CreateBy);

                        var userResult = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in userResult)
                        {
                            UserNo = item.UserNo;
                        }
                        #endregion

                        #region //取得MES.OspReceipt資料(MOCTH)
                        dynamicParameters = new DynamicParameters();
                        sqlQuery.mainKey = "a.OsrId";
                        sqlQuery.auxKey = "";
                        sqlQuery.columns =
                            @", a.OsrErpPrefix, a.OsrErpNo, FORMAT(a.ReceiptDate, 'yyyyMMdd') ReceiptDate, a.FactoryCode, a.SupplierId, a.SupplierSo, a.CurrencyCode
                            , a.Exchange, a.RowCnt, a.Remark, a.UiNo, a.InvoiceType, FORMAT(a.InvoiceDate, 'yyyyMMdd') InvoiceDate, a.InvoiceNo, a.TaxType, a.DeductType
                            , a.ReserveFlag, a.OrigAmount, a.DeductAmount, a.OrigTax, a.ReceiptAmount, a.Quantity, a.ConfirmStatus, a.ConfirmUserId
                            , a.RenewFlag, a.PrintCnt, a.AutoMaterialBilling, a.OrigPreTaxAmount, FORMAT(a.ApplyYYMM, 'yyyyMM') ApplyYYMM, FORMAT(a.DocumentDate, 'yyyyMMdd') DocumentDate
                            , a.TaxRate, a.PretaxAmount, a.TaxAmount, a.PaymentTerm, a.PackageQuantity, a.SupplierPicking, a.ApproveStatus
                            , a.ReserveTaxCode, a.SendCount, a.NoticeFlag, a.TaxCode, a.TaxExchange, a.QcFlag, a.FlowStatus, FORMAT(a.CreateDate, 'HH:mm:ss') CreateTime
                            , b.UserNo ConfirmUserNo
                            , c.SupplierNo
                            , d.UserNo CreateUserNo";
                        sqlQuery.mainTables =
                            @"FROM MES.OspReceipt a
                            LEFT JOIN BAS.[User] b ON a.ConfirmUserId = b.UserId
                            INNER JOIN SCM.Supplier c ON a.SupplierId = c.SupplierId
                            INNER JOIN BAS.[User] d ON a.CreateBy = d.UserId";
                        sqlQuery.auxTables = "";
                        string queryCondition = @"AND a.OsrId=@OsrId";
                        dynamicParameters.Add("OsrId", OsrId);
                        sqlQuery.conditions = queryCondition;
                        sqlQuery.orderBy = "";

                        ospReceipts = BaseHelper.SqlQuery<OspReceipt>(sqlConnection, dynamicParameters, sqlQuery);
                        if (ospReceipts.Count() <= 0) throw new SystemException("託外入庫資料有誤!");

                        if (ospReceipts[0].InvoiceDate == "19000101") ospReceipts[0].InvoiceDate = "";

                        foreach (var item in ospReceipts)
                        {
                            if (item.ConfirmStatus != "N") throw new SystemException("單據非未確認狀態，無法修改!!");
                        }
                        #endregion

                        #region //取得MES.OspReceiptDetail(MOCTI)
                        dynamicParameters = new DynamicParameters();
                        sqlQuery.mainKey = "a.OsrDetailId";
                        sqlQuery.auxKey = "";
                        sqlQuery.columns =
                            @", a.OsrId, a.OspDetailId, a.OsrErpPrefix, a.OsrErpNo, a.OsrSeq, a.MtlItemId, a.ReceiptQty, a.UomId, a.InventoryId
                            , a.LotNumber, FORMAT(a.AvailableDate, 'yyyyMMdd') AvailableDate, FORMAT(a.ReCheckDate, 'yyyyMMdd') ReCheckDate, a.MoId, a.MoProcessId, a.ProcessCode, a.ReceiptPackageQty
                            , a.AcceptancePackageQty, FORMAT(a.AcceptanceDate, 'yyyyMMdd') AcceptanceDate, a.AcceptQty, a.AvailableQty, a.ScriptQty, a.ReturnQty, a.AvailableUom, a.OrigUnitPrice
                            , a.OrigAmount, a.OrigDiscountAmt, a.ReceiptExpense, a.DiscountDescription, a.ProjectCode, a.PaymentHold, a.Overdue, a.QcStatus
                            , a.ReturnCode, a.ConfirmStatus, a.CloseStatus, a.ReNewStatus, a.Remark, a.CostEntry, a.ExpenseEntry, a.ConfirmUserId, a.ConfirmUser, a.OrigPreTaxAmt
                            , a.OrigTaxAmt, a.PreTaxAmt, a.TaxAmt, a.UrgentMtl, a.ReserveTaxCode, a.ApproveStatus, a.TransferStatus, a.ProcessStatus, FORMAT(a.CreateDate, 'HH:mm:ss') CreateTime
                            , b.UserNo CreateUserNo
                            , c.MtlItemNo, c.MtlItemName, c.MtlItemSpec
                            , d.UomNo
                            , e.InventoryNo
                            , g.WoErpPrefix, g.WoErpNo
                            , h.Status MoStatus";
                        sqlQuery.mainTables =
                            @"FROM MES.OspReceiptDetail a
                            INNER JOIN BAS.[User] b ON a.CreateBy = b.UserId
                            INNER JOIN PDM.MtlItem c ON a.MtlItemId = c.MtlItemId
                            INNER JOIN PDM.UnitOfMeasure d ON a.UomId = d.UomId
                            INNER JOIN SCM.Inventory e ON a.InventoryId = e.InventoryId
                            INNER JOIN MES.ManufactureOrder f ON a.MoId = f.MoId
                            INNER JOIN MES.WipOrder g ON f.WoId = g.WoId
                            INNER JOIN MES.ManufactureOrder h ON a.MoId = h.MoId";
                        sqlQuery.auxTables = "";
                        queryCondition = @"AND OsrId = @OsrId";
                        dynamicParameters.Add("OsrId", OsrId);
                        sqlQuery.conditions = queryCondition;
                        sqlQuery.orderBy = "";

                        ospReceiptDetails = BaseHelper.SqlQuery<OspReceiptDetail>(sqlConnection, dynamicParameters, sqlQuery);
                        if (ospReceiptDetails.Count() <= 0) throw new SystemException("託外入庫詳細資料有誤!");
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //審核ERP權限
                            string USR_GROUP = BaseHelper.CheckErpAuthority(UserNo, sqlConnection2, "MOCI06", "C-CONFIRM");
                            #endregion

                            #region //查詢ERP資料庫中是否已經有此托外進貨單
                            dynamicParameters = new DynamicParameters();
                            sqlQuery.mainKey = "a.TH001, a.TH002";
                            sqlQuery.auxKey = "";
                            sqlQuery.columns =
                                @", LTRIM(RTRIM(a.COMPANY)) COMPANY, LTRIM(RTRIM(a.MODIFIER)) MODIFIER, LTRIM(RTRIM(a.USR_GROUP)) USR_GROUP, LTRIM(RTRIM(a.MODI_DATE)) MODI_DATE
                                , LTRIM(RTRIM(a.FLAG)) FLAG, LTRIM(RTRIM(a.MODI_TIME)) MODI_TIME, LTRIM(RTRIM(a.TH009)) TH009
                                , LTRIM(RTRIM(a.TH010)) TH010, LTRIM(RTRIM(a.TH015)) TH015, LTRIM(RTRIM(a.TH016)) TH016, LTRIM(RTRIM(a.TH017)) TH017, LTRIM(RTRIM(a.TH018)) TH018
                                , LTRIM(RTRIM(a.TH020)) TH020, LTRIM(RTRIM(a.TH021)) TH021, LTRIM(RTRIM(a.TH022)) TH022, LTRIM(RTRIM(a.TH023)) TH023";
                            sqlQuery.mainTables =
                                @"FROM MOCTH a";
                            sqlQuery.auxTables = "";
                            queryCondition = @"AND a.TH001 = @TH001 AND a.TH002 = @TH002";
                            dynamicParameters.Add("TH001", ospReceipts[0].OsrErpPrefix);
                            dynamicParameters.Add("TH002", ospReceipts[0].OsrErpNo);
                            sqlQuery.conditions = queryCondition;
                            sqlQuery.orderBy = "";

                            mocths = BaseHelper.SqlQuery<MOCTH>(sqlConnection2, dynamicParameters, sqlQuery);
                            #endregion

                            if (mocths.Count() > 0)
                            {
                                #region //UPDATE ERP TABLE
                                #region //MOCTH 託外入庫單頭
                                foreach (var item in mocths)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE MOCTH SET
                                            COMPANY = @COMPANY,
                                            MODIFIER = @MODIFIER,
                                            USR_GROUP = @USR_GROUP,
                                            MODI_DATE = @MODI_DATE,
                                            FLAG = @FLAG,
                                            MODI_TIME = @MODI_TIME,
                                            TH001 = @TH001,
                                            TH002 = @TH002,
                                            TH003 = @TH003,
                                            TH004 = @TH004,
                                            TH005 = @TH005,
                                            TH006 = @TH006,
                                            TH007 = @TH007,
                                            TH008 = @TH008,
                                            TH009 = @TH009,
                                            TH010 = @TH010,
                                            TH011 = @TH011,
                                            TH012 = @TH012,
                                            TH013 = @TH013,
                                            TH014 = @TH014,
                                            TH015 = @TH015,
                                            TH016 = @TH016,
                                            TH017 = @TH017,
                                            TH018 = @TH018,
                                            TH019 = @TH019,
                                            TH020 = @TH020,
                                            TH021 = @TH021,
                                            TH022 = @TH022,
                                            TH023 = @TH023,
                                            TH024 = @TH024,
                                            TH025 = @TH025,
                                            TH026 = @TH026,
                                            TH027 = @TH027,
                                            TH028 = @TH028,
                                            TH029 = @TH029,
                                            TH030 = @TH030,
                                            TH031 = @TH031,
                                            TH032 = @TH032,
                                            TH033 = @TH033,
                                            TH034 = @TH034,
                                            TH035 = @TH035,
                                            TH036 = @TH036,
                                            TH037 = @TH037,
                                            TH038 = @TH038,
                                            TH039 = @TH039,
                                            TH040 = @TH040,
                                            TH041 = @TH041,
                                            TH042 = @TH042,
                                            TH043 = @TH043,
                                            TH044 = @TH044,
                                            TH045 = @TH045,
                                            TH046 = @TH046,
                                            TH047 = @TH047,
                                            TH048 = @TH048,
                                            TH049 = @TH049,
                                            TH050 = @TH050,
                                            UDF06 = @UDF06,
                                            UDF07 = @UDF07,
                                            UDF08 = @UDF08,
                                            UDF09 = @UDF09,
                                            UDF10 = @UDF10
                                            WHERE TH001 = @TH001
                                            AND TH002 = @TH002";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            COMPANY = ErpNo,
                                            item.MODIFIER,
                                            USR_GROUP,
                                            item.MODI_DATE,
                                            FLAG = "3",
                                            item.MODI_TIME,
                                            item.TH001,
                                            item.TH002,
                                            TH003 = ospReceipts[0].ReceiptDate,
                                            TH004 = ospReceipts[0].FactoryCode,
                                            TH005 = ospReceipts[0].SupplierNo,
                                            TH006 = ospReceipts[0].SupplierSo,
                                            TH007 = ospReceipts[0].CurrencyCode,
                                            TH008 = ospReceipts[0].Exchange,
                                            TH009 = ospReceipts[0].RowCnt,
                                            TH010 = ospReceipts[0].Remark,
                                            TH011 = ospReceipts[0].UiNo,
                                            TH012 = ospReceipts[0].InvoiceType,
                                            TH013 = ospReceipts[0].InvoiceDate,
                                            TH014 = ospReceipts[0].InvoiceNo,
                                            TH015 = ospReceipts[0].TaxType,
                                            TH016 = ospReceipts[0].DeductType,
                                            TH017 = ospReceipts[0].ReserveFlag,
                                            TH018 = ospReceipts[0].OrigAmount,
                                            TH019 = ospReceipts[0].DeductAmount,
                                            TH020 = ospReceipts[0].OrigTax,
                                            TH021 = ospReceipts[0].ReceiptAmount,
                                            TH022 = ospReceipts[0].Quantity,
                                            TH023 = ospReceipts[0].ConfirmStatus,
                                            TH024 = ospReceipts[0].RenewFlag,
                                            TH025 = ospReceipts[0].PrintCnt,
                                            TH026 = ospReceipts[0].AutoMaterialBilling,
                                            TH027 = ospReceipts[0].OrigPreTaxAmount,
                                            TH028 = ospReceipts[0].ApplyYYMM,
                                            TH029 = ospReceipts[0].DocumentDate,
                                            TH030 = ospReceipts[0].TaxRate,
                                            TH031 = ospReceipts[0].PretaxAmount,
                                            TH032 = ospReceipts[0].TaxAmount,
                                            TH033 = ospReceipts[0].PaymentTerm,
                                            TH034 = ospReceipts[0].PackageQuantity,
                                            TH035 = ospReceipts[0].SupplierPicking,
                                            TH036 = ospReceipts[0].ApproveStatus,
                                            TH037 = ospReceipts[0].ReserveTaxCode,
                                            TH038 = ospReceipts[0].SendCount,
                                            item.TH039,
                                            item.TH040,
                                            TH041 = ospReceipts[0].NoticeFlag,
                                            item.TH042,
                                            item.TH043,
                                            TH044 = ospReceipts[0].TaxCode,
                                            item.TH045,
                                            item.TH046,
                                            TH047 = ospReceipts[0].TaxExchange,
                                            item.TH048,
                                            item.TH049,
                                            item.TH050,
                                            item.UDF06,
                                            item.UDF07,
                                            item.UDF08,
                                            item.UDF09,
                                            item.UDF10
                                        });

                                    rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);
                                }
                                #endregion

                                #region //MOCTI 託外入庫單身設定
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE MOCTI
                                        WHERE TI001 = @TI001
                                        AND TI002 = @TI002";
                                dynamicParameters.Add("TI001", ospReceipts[0].OsrErpPrefix);
                                dynamicParameters.Add("TI002", ospReceipts[0].OsrErpNo);

                                rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);

                                foreach (var item in ospReceiptDetails)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO MOCTI (COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE, FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID
                                            , TI001, TI002, TI003, TI004, TI005, TI006, TI007, TI008, TI009, TI010, TI011, TI012, TI013, TI014, TI015, TI016, TI017, TI018, TI019, TI020, TI021, TI022
                                            , TI023, TI024, TI025, TI026, TI027, TI028, TI032, TI033, TI034, TI035, TI036, TI037, TI038, TI039, TI040, TI041, TI042, TI043, TI044
                                            , TI045, TI046, TI047, TI048, TI049, TI050, TI051, TI052, TI053, TI054, TI055, TI056, TI057, TI058, TI059, TI060, TI061, TI062, TI063, TI064, TI065, TI066
                                            , TI067, TI068, TI069, TI070, TI071, TI503, TI504, TI505, TI506, TI551, TI552, TI553, TI554, TI555, TI556
                                            , UDF06, UDF07, UDF08, UDF09, UDF10)
                                            VALUES (@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE, @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID, @MODI_TIME, @MODI_AP, @MODI_PRID
                                            , @TI001, @TI002, @TI003, @TI004, @TI005, @TI006, @TI007, @TI008, @TI009, @TI010, @TI011, @TI012, @TI013, @TI014, @TI015, @TI016, @TI017, @TI018, @TI019, @TI020, @TI021, @TI022
                                            , @TI023, @TI024, @TI025, @TI026, @TI027, @TI028, @TI032, @TI033, @TI034, @TI035, @TI036, @TI037, @TI038, @TI039, @TI040, @TI041, @TI042, @TI043, @TI044
                                            , @TI045, @TI046, @TI047, @TI048, @TI049, @TI050, @TI051, @TI052, @TI053, @TI054, @TI055, @TI056, @TI057, @TI058, @TI059, @TI060, @TI061, @TI062, @TI063, @TI064, @TI065, @TI066
                                            , @TI067, @TI068, @TI069, @TI070, @TI071, @TI503, @TI504, @TI505, @TI506, @TI551, @TI552, @TI553, @TI554, @TI555, @TI556
                                            , @UDF06, @UDF07, @UDF08, @UDF09, @UDF10)";

                                    dynamicParameters.AddDynamicParams(
                                      new
                                      {
                                          COMPANY = ErpNo,
                                          CREATOR = ospReceipts[0].CreateUserNo,
                                          USR_GROUP,
                                          CREATE_DATE = DateTime.Now.ToString("yyyyMMdd"),
                                          MODIFIER = "",
                                          MODI_DATE = "",
                                          FLAG = "1",
                                          CREATE_TIME = item.CreateTime,
                                          CREATE_AP = "ERP-API",
                                          CREATE_PRID = "BM",
                                          MODI_TIME = "",
                                          MODI_AP = "",
                                          MODI_PRID = "",
                                          TI001 = item.OsrErpPrefix,
                                          TI002 = item.OsrErpNo,
                                          TI003 = item.OsrSeq,
                                          TI004 = item.MtlItemNo,
                                          TI005 = item.MtlItemName,
                                          TI006 = item.MtlItemSpec,
                                          TI007 = item.ReceiptQty,
                                          TI008 = item.UomNo,
                                          TI009 = "",
                                          TI010 = item.LotNumber,
                                          TI011 = item.AvailableDate,
                                          TI012 = item.ReCheckDate,
                                          TI013 = item.WoErpPrefix,
                                          TI014 = item.WoErpNo,
                                          TI015 = item.ProcessCode,
                                          TI016 = item.ReceiptPackageQty,
                                          TI017 = item.AcceptancePackageQty,
                                          TI018 = item.AcceptanceDate,
                                          TI019 = item.AcceptQty,
                                          TI020 = item.AvailableQty,
                                          TI021 = item.ScriptQty,
                                          TI022 = item.ReturnQty,
                                          TI023 = item.UomNo,
                                          TI024 = item.OrigUnitPrice,
                                          TI025 = item.OrigAmount,
                                          TI026 = item.OrigDiscountAmt,
                                          TI027 = item.ReceiptExpense,
                                          TI028 = item.DiscountDescription,
                                          TI032 = item.ProjectCode,
                                          TI033 = item.PaymentHold,
                                          TI034 = item.Overdue,
                                          TI035 = item.QcStatus,
                                          TI036 = item.ReturnCode,
                                          TI037 = item.ConfirmStatus,
                                          TI038 = item.CloseStatus,
                                          TI039 = item.ReNewStatus,
                                          TI040 = item.Remark,
                                          TI041 = item.CostEntry,
                                          TI042 = item.ExpenseEntry,
                                          TI043 = item.ConfirmUser,
                                          TI044 = item.OrigPreTaxAmt,
                                          TI045 = item.OrigTaxAmt,
                                          TI046 = item.PreTaxAmt,
                                          TI047 = item.TaxAmt,
                                          TI048 = item.UrgentMtl,
                                          TI049 = "",
                                          TI050 = 0,
                                          TI051 = 0,
                                          TI052 = item.ReserveTaxCode,
                                          TI053 = item.ApproveStatus,
                                          TI054 = "N",
                                          TI055 = "0",
                                          TI056 = "2",
                                          TI057 = "",
                                          TI058 = 0,
                                          TI059 = 0,
                                          TI060 = "",
                                          TI061 = "",
                                          TI062 = "",
                                          TI063 = "",
                                          TI064 = "",
                                          TI065 = 0,
                                          TI066 = 0,
                                          TI067 = 0,
                                          TI068 = 0,
                                          TI069 = "",
                                          TI070 = "",
                                          TI071 = 0,
                                          TI503 = 0,
                                          TI504 = 0,
                                          TI505 = 0,
                                          TI506 = 0,
                                          TI551 = 0,
                                          TI552 = 0,
                                          TI553 = 0,
                                          TI554 = 0,
                                          TI555 = 0,
                                          TI556 = 0,
                                          UDF06 = 0,
                                          UDF07 = 0,
                                          UDF08 = 0,
                                          UDF09 = 0,
                                          UDF10 = 0
                                      });
                                    var insertResult = sqlConnection2.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult.Count();
                                }
                                #endregion
                                #endregion
                            }
                            else
                            {
                                #region //找單據日期託外入庫單數
                                dynamicParameters = new DynamicParameters();
                                string nowDate = DateTime.Now.ToString("yyyyMMdd");
                                sql = @"SELECT COUNT(1) Count
                                        FROM MOCTH
                                        WHERE TH001 = @TH001
                                        AND TH002 LIKE '%" + ospReceipts[0].DocumentDate + "%'";
                                dynamicParameters.Add("TH001", ospReceipts[0].OsrErpPrefix);

                                var result2 = sqlConnection2.Query(sql, dynamicParameters);
                                int count = 0;
                                foreach (var item in result2)
                                {
                                    count = item.Count + 1;
                                }

                                string seq = count.ToString().PadLeft(3, '0');
                                OsrErpNo = ospReceipts[0].DocumentDate + seq;
                                ospReceipts[0].OsrErpNo = OsrErpNo;
                                #endregion

                                #region //MOCTH 託外入庫單單頭
                                foreach (var item in ospReceipts)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO MOCTH (COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE, FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID
                                            , TH001, TH002, TH003, TH004, TH005, TH006, TH007, TH008, TH009, TH010, TH011, TH012, TH013, TH014, TH015, TH016, TH017, TH018, TH019, TH020, TH021, TH022
                                            , TH023, TH024, TH025, TH026, TH027, TH028, TH029, TH030, TH031, TH032, TH033, TH034, TH035, TH036, TH037, TH038, TH039, TH040, TH041, TH042, TH043
                                            , TH044, TH045, TH046, TH047, TH048, TH049, TH050
                                            , UDF06, UDF07, UDF08, UDF09, UDF10)
                                            VALUES (@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE, @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID, @MODI_TIME, @MODI_AP, @MODI_PRID
                                            , @TH001, @TH002, @TH003, @TH004, @TH005, @TH006, @TH007, @TH008, @TH009, @TH010, @TH011, @TH012, @TH013, @TH014, @TH015, @TH016, @TH017, @TH018, @TH019, @TH020, @TH021, @TH022
                                            , @TH023, @TH024, @TH025, @TH026, @TH027, @TH028, @TH029, @TH030, @TH031, @TH032, @TH033, @TH034, @TH035, @TH036, @TH037, @TH038, @TH039, @TH040, @TH041, @TH042, @TH043
                                            , @TH044, @TH045, @TH046, @TH047, @TH048, @TH049, @TH050
                                            , @UDF06, @UDF07, @UDF08, @UDF09, @UDF10)";

                                    dynamicParameters.AddDynamicParams(
                                      new
                                      {
                                          COMPANY = ErpNo,
                                          CREATOR = ospReceipts[0].CreateUserNo,
                                          USR_GROUP,
                                          CREATE_DATE = DateTime.Now.ToString("yyyyMMdd"),
                                          MODIFIER = "",
                                          MODI_DATE = "",
                                          FLAG = "1",
                                          CREATE_TIME = item.CreateTime,
                                          CREATE_AP = "ERP-API",
                                          CREATE_PRID = "BM",
                                          MODI_TIME = "",
                                          MODI_AP = "",
                                          MODI_PRID = "",
                                          TH001 = item.OsrErpPrefix,
                                          TH002 = item.OsrErpNo,
                                          TH003 = item.ReceiptDate,
                                          TH004 = item.FactoryCode,
                                          TH005 = item.SupplierNo,
                                          TH006 = item.SupplierSo,
                                          TH007 = item.CurrencyCode,
                                          TH008 = item.Exchange,
                                          TH009 = item.RowCnt,
                                          TH010 = item.Remark,
                                          TH011 = item.UiNo,
                                          TH012 = item.InvoiceType,
                                          TH013 = item.InvoiceDate,
                                          TH014 = item.InvoiceNo,
                                          TH015 = item.TaxType,
                                          TH016 = item.DeductType,
                                          TH017 = item.ReserveFlag,
                                          TH018 = item.OrigAmount,
                                          TH019 = item.DeductAmount,
                                          TH020 = item.OrigTax,
                                          TH021 = item.ReceiptAmount,
                                          TH022 = item.Quantity,
                                          TH023 = item.ConfirmStatus,
                                          TH024 = item.RenewFlag,
                                          TH025 = item.PrintCnt,
                                          TH026 = item.AutoMaterialBilling,
                                          TH027 = item.OrigPreTaxAmount,
                                          TH028 = item.ApplyYYMM,
                                          TH029 = item.DocumentDate,
                                          TH030 = item.TaxRate,
                                          TH031 = item.PretaxAmount,
                                          TH032 = item.TaxAmount,
                                          TH033 = item.PaymentTerm,
                                          TH034 = item.PackageQuantity,
                                          TH035 = item.SupplierPicking,
                                          TH036 = item.ApproveStatus,
                                          TH037 = item.ReserveTaxCode,
                                          TH038 = item.SendCount,
                                          TH039 = 0,
                                          TH040 = 0,
                                          TH041 = item.NoticeFlag,
                                          TH042 = "",
                                          TH043 = "",
                                          TH044 = item.TaxCode,
                                          TH045 = "",
                                          TH046 = "",
                                          TH047 = item.TaxExchange,
                                          TH048 = 0,
                                          TH049 = 0,
                                          TH050 = "",
                                          UDF06 = 0.000000,
                                          UDF07 = 0.000000,
                                          UDF08 = 0.000000,
                                          UDF09 = 0.000000,
                                          UDF10 = 0.000000,
                                      });
                                    var insertResult = sqlConnection2.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult.Count();
                                }
                                #endregion

                                #region //MOCTI 託外入庫單身
                                foreach (var item in ospReceiptDetails)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO MOCTI (COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE, FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID
                                            , TI001, TI002, TI003, TI004, TI005, TI006, TI007, TI008, TI009, TI010, TI011, TI012, TI013, TI014, TI015, TI016, TI017, TI018, TI019, TI020, TI021, TI022
                                            , TI023, TI024, TI025, TI026, TI027, TI028, TI032, TI033, TI034, TI035, TI036, TI037, TI038, TI039, TI040, TI041, TI042, TI043, TI044
                                            , TI045, TI046, TI047, TI048, TI049, TI050, TI051, TI052, TI053, TI054, TI055, TI056, TI057, TI058, TI059, TI060, TI061, TI062, TI063, TI064, TI065, TI066
                                            , TI067, TI068, TI069, TI070, TI071, TI503, TI504, TI505, TI506, TI551, TI552, TI553, TI554, TI555, TI556
                                            , UDF06, UDF07, UDF08, UDF09, UDF10)
                                            VALUES (@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE, @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID, @MODI_TIME, @MODI_AP, @MODI_PRID
                                            , @TI001, @TI002, @TI003, @TI004, @TI005, @TI006, @TI007, @TI008, @TI009, @TI010, @TI011, @TI012, @TI013, @TI014, @TI015, @TI016, @TI017, @TI018, @TI019, @TI020, @TI021, @TI022
                                            , @TI023, @TI024, @TI025, @TI026, @TI027, @TI028, @TI032, @TI033, @TI034, @TI035, @TI036, @TI037, @TI038, @TI039, @TI040, @TI041, @TI042, @TI043, @TI044
                                            , @TI045, @TI046, @TI047, @TI048, @TI049, @TI050, @TI051, @TI052, @TI053, @TI054, @TI055, @TI056, @TI057, @TI058, @TI059, @TI060, @TI061, @TI062, @TI063, @TI064, @TI065, @TI066
                                            , @TI067, @TI068, @TI069, @TI070, @TI071, @TI503, @TI504, @TI505, @TI506, @TI551, @TI552, @TI553, @TI554, @TI555, @TI556
                                            , @UDF06, @UDF07, @UDF08, @UDF09, @UDF10)";

                                    dynamicParameters.AddDynamicParams(
                                      new
                                      {
                                          COMPANY = ErpNo,
                                          CREATOR = ospReceipts[0].CreateUserNo,
                                          USR_GROUP,
                                          CREATE_DATE = DateTime.Now.ToString("yyyyMMdd"),
                                          MODIFIER = "",
                                          MODI_DATE = "",
                                          FLAG = "1",
                                          CREATE_TIME = item.CreateTime,
                                          CREATE_AP = "ERP-API",
                                          CREATE_PRID = "BM",
                                          MODI_TIME = "",
                                          MODI_AP = "",
                                          MODI_PRID = "",
                                          TI001 = ospReceipts[0].OsrErpPrefix,
                                          TI002 = ospReceipts[0].OsrErpNo,
                                          TI003 = item.OsrSeq,
                                          TI004 = item.MtlItemNo,
                                          TI005 = item.MtlItemName,
                                          TI006 = item.MtlItemSpec,
                                          TI007 = item.ReceiptQty,
                                          TI008 = item.UomNo,
                                          TI009 = "",
                                          TI010 = item.LotNumber,
                                          TI011 = item.AvailableDate,
                                          TI012 = item.ReCheckDate,
                                          TI013 = item.WoErpPrefix,
                                          TI014 = item.WoErpNo,
                                          TI015 = item.ProcessCode,
                                          TI016 = item.ReceiptPackageQty,
                                          TI017 = item.AcceptancePackageQty,
                                          TI018 = item.AcceptanceDate,
                                          TI019 = item.AcceptQty,
                                          TI020 = item.AvailableQty,
                                          TI021 = item.ScriptQty,
                                          TI022 = item.ReturnQty,
                                          TI023 = item.UomNo,
                                          TI024 = item.OrigUnitPrice,
                                          TI025 = item.OrigAmount,
                                          TI026 = item.OrigDiscountAmt,
                                          TI027 = item.ReceiptExpense,
                                          TI028 = item.DiscountDescription,
                                          TI032 = item.ProjectCode,
                                          TI033 = item.PaymentHold,
                                          TI034 = item.Overdue,
                                          TI035 = item.QcStatus,
                                          TI036 = item.ReturnCode,
                                          TI037 = item.ConfirmStatus,
                                          TI038 = item.CloseStatus,
                                          TI039 = item.ReNewStatus,
                                          TI040 = item.Remark,
                                          TI041 = item.CostEntry,
                                          TI042 = item.ExpenseEntry,
                                          TI043 = item.ConfirmUser,
                                          TI044 = item.OrigPreTaxAmt,
                                          TI045 = item.OrigTaxAmt,
                                          TI046 = item.PreTaxAmt,
                                          TI047 = item.TaxAmt,
                                          TI048 = item.UrgentMtl,
                                          TI049 = "",
                                          TI050 = 0,
                                          TI051 = 0,
                                          TI052 = item.ReserveTaxCode,
                                          TI053 = item.ApproveStatus,
                                          TI054 = "N",
                                          TI055 = "0",
                                          TI056 = "2",
                                          TI057 = "",
                                          TI058 = 0,
                                          TI059 = 0,
                                          TI060 = "",
                                          TI061 = "",
                                          TI062 = "",
                                          TI063 = "",
                                          TI064 = "",
                                          TI065 = 0,
                                          TI066 = 0,
                                          TI067 = 0,
                                          TI068 = 0,
                                          TI069 = "",
                                          TI070 = "",
                                          TI071 = 0,
                                          TI503 = 0,
                                          TI504 = 0,
                                          TI505 = 0,
                                          TI506 = 0,
                                          TI551 = 0,
                                          TI552 = 0,
                                          TI553 = 0,
                                          TI554 = 0,
                                          TI555 = 0,
                                          TI556 = 0,
                                          UDF06 = 0,
                                          UDF07 = 0,
                                          UDF08 = 0,
                                          UDF09 = 0,
                                          UDF10 = 0
                                      });
                                    var insertResult = sqlConnection2.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult.Count();
                                }
                                #endregion
                            }

                            #region //確認並拋轉ERP
                            string COMPANY = ErpDbName;
                            string TYPE = "MOCI06";
                            string DOC = ospReceipts[0].OsrErpPrefix;
                            string DOC_NO = ospReceipts[0].OsrErpNo;
                            string TRANS = "1";
                            string USER = UserNo;
                            //string DATE = DateTime.Now.ToString("yyyyMMdd");
                            string DATE = ospReceipts[0].ReceiptDate;

                            var result = BaseHelper.TransferErpAPI(COMPANY, TYPE, DOC, DOC_NO, TRANS, USER, DATE);
                            #endregion

                            #region //若成功則將相關資料寫回MES
                            var resultJson = JToken.Parse(result);
                            string CODE = resultJson["CODE"].ToString();
                            string ERROR_CODE = resultJson["ERROR_CODE"].ToString(); //錯誤代碼
                            if (CODE == "200")
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE MES.OspReceipt SET
                                        ConfirmStatus = @ConfirmStatus,
                                        ConfirmUserId = @ConfirmUserId,
                                        ApproveStatus = @ApproveStatus,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE OsrId = @OsrId";
                                dynamicParameters.AddDynamicParams(
                                  new
                                  {
                                      ConfirmStatus = "Y",
                                      ConfirmUserId = CreateBy,
                                      ApproveStatus = "3",
                                      LastModifiedDate,
                                      LastModifiedBy,
                                      OsrId
                                  });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE MES.OspReceiptDetail SET
                                        ConfirmStatus = @ConfirmStatus,
                                        ConfirmUserId = @ConfirmUserId,
                                        ConfirmUser = @ConfirmUser,
                                        ApproveStatus = @ApproveStatus,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE OsrId = @OsrId";
                                dynamicParameters.AddDynamicParams(
                                  new
                                  {
                                      ConfirmStatus = "Y",
                                      ConfirmUserId = CreateBy,
                                      ConfirmUser = ospReceipts[0].ConfirmUserNo,
                                      ApproveStatus = "3",
                                      LastModifiedDate,
                                      LastModifiedBy,
                                      OsrId
                                  });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                                #region //Response
                                jsonResponse = JObject.FromObject(new
                                {
                                    status = "success",
                                    msg = "拋轉成功!"
                                });
                                #endregion
                            }
                            else
                            {
                                throw new SystemException("確認失敗，請聯繫資訊人員進行問題排解!");
                            }
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

        #region //UpdateReConfirmOspReceipt -- 反確認託外入庫單資料 -- Ann 2022-09-21
        public string UpdateReConfirmOspReceipt(int OsrId = -1)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    string OsrErpPrefix = "";
                    string OsrErpNo = "";
                    string ConfirmUser = "";
                    string ErpDbName = "";
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
                            ErpDbName = ErpConnectionStrings.Split('=')[2].Split(';')[0];
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //判斷託外入庫單資料是否有誤
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.ConfirmStatus, a.OsrErpPrefix, a.OsrErpNo
                                , b.ConfirmUser
                                FROM MES.OspReceipt a
                                INNER JOIN MES.OspReceiptDetail b ON a.OsrId = b.OsrId
                                WHERE a.OsrId = @OsrId";
                            dynamicParameters.Add("OsrId", OsrId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("託外入庫單資料有誤!");

                            foreach (var item in result)
                            {
                                if (item.ConfirmStatus != "Y") throw new SystemException("領退料單非確認狀態!");
                                OsrErpPrefix = item.OsrErpPrefix;
                                OsrErpNo = item.OsrErpNo;
                                ConfirmUser = item.ConfirmUser;
                            }
                            #endregion

                            #region //查詢UserNo
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.UserNo
                                    FROM BAS.[User] a
                                    WHERE UserId = @UserId";
                            dynamicParameters.Add("UserId", CreateBy);

                            var userResult = sqlConnection.Query(sql, dynamicParameters);

                            string UserNo = "";
                            foreach (var item in userResult)
                            {
                                UserNo = item.UserNo;
                            }
                            #endregion

                            #region //審核ERP權限
                            string USR_GROUP = BaseHelper.CheckErpAuthority(UserNo, sqlConnection2, "MOCI06", "RE-CONFIRM");
                            #endregion

                            #region //進行反確認
                            string COMPANY = ErpDbName;
                            string TYPE = "MOCI06";
                            string DOC = OsrErpPrefix;
                            string DOC_NO = OsrErpNo;
                            string TRANS = "2";
                            string USER = UserNo;
                            string DATE = DateTime.Now.ToString("yyyyMMdd");

                            var result2 = BaseHelper.TransferErpAPI(COMPANY, TYPE, DOC, DOC_NO, TRANS, USER, DATE);
                            #endregion

                            #region //若成功則將相關資料寫回MES
                            var resultJson = JToken.Parse(result2);
                            string CODE = resultJson["CODE"].ToString();
                            string ERROR_CODE = resultJson["ERROR_CODE"].ToString(); //錯誤代碼
                            if (CODE == "200")
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE MES.OspReceipt SET
                                        ConfirmStatus = @ConfirmStatus,
                                        ConfirmUserId = @ConfirmUserId,
                                        ApproveStatus = @ApproveStatus,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE OsrId = @OsrId";
                                dynamicParameters.AddDynamicParams(
                                  new
                                  {
                                      ConfirmStatus = "N",
                                      ConfirmUserId = CreateBy,
                                      ApproveStatus = "0",
                                      LastModifiedDate,
                                      LastModifiedBy,
                                      OsrId
                                  });

                                int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE MES.OspReceiptDetail SET
                                        ConfirmStatus = @ConfirmStatus,
                                        ConfirmUserId = @ConfirmUserId,
                                        ConfirmUser = @ConfirmUser,
                                        ApproveStatus = @ApproveStatus,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE OsrId = @OsrId";
                                dynamicParameters.AddDynamicParams(
                                  new
                                  {
                                      ConfirmStatus = "N",
                                      ConfirmUserId = CreateBy,
                                      ConfirmUser,
                                      ApproveStatus = "0",
                                      LastModifiedDate,
                                      LastModifiedBy,
                                      OsrId
                                  });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                                #region //Response
                                jsonResponse = JObject.FromObject(new
                                {
                                    status = "success",
                                    msg = "(" + rowsAffected + " rows affected)"
                                });
                                #endregion
                            }
                            else
                            {
                                throw new SystemException("反確認失敗!錯誤代碼:" + ERROR_CODE + "");
                            }
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

        #region //UpdateVoidOspReceipt -- 作廢託外入庫單 -- Ann 2022-09-22
        public string UpdateVoidOspReceipt(int OsrId = -1)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    string OsrErpPrefix = "";
                    string OsrErpNo = "";
                    string ConfirmUserNo = "";
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

                        #region //判斷領退料單詳細資料是否有誤
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 ConfirmStatus, OsrErpPrefix,  OsrErpNo
                                , b.UserNo ConfirmUserNo
                                FROM MES.OspReceipt a
                                LEFT JOIN BAS.[User] b ON a.ConfirmUserId = b.UserId
                                WHERE a.OsrId = @OsrId";
                        dynamicParameters.Add("OsrId", OsrId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("託外入庫單資料有誤!");

                        foreach (var item in result)
                        {
                            if (item.ConfirmStatus != "N" || item.ConfirmUserNo == null) throw new SystemException("託外入庫單狀態無法作廢!");
                            OsrErpPrefix = item.OsrErpPrefix;
                            OsrErpNo = item.OsrErpNo;
                            ConfirmUserNo = item.ConfirmUserNo;
                        }
                        #endregion
                    }

                    #region //進行作廢
                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MOCTH SET
                                TH023 = @TH023
                                WHERE TH001 = @TH001
                                AND TH002 = @TH002";
                        dynamicParameters.AddDynamicParams(
                          new
                          {
                              TH023 = "V",
                              TH001 = OsrErpPrefix,
                              TH002 = OsrErpNo
                          });

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                    }
                    #endregion

                    #region //若成功則將相關資料寫回MES
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.OspReceipt SET
                                ConfirmStatus = @ConfirmStatus,
                                ConfirmUserId = @ConfirmUserId,
                                ApproveStatus = @ApproveStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE OsrId = @OsrId";
                        dynamicParameters.AddDynamicParams(
                          new
                          {
                              ConfirmStatus = "V",
                              ConfirmUserId = CreateBy,
                              ApproveStatus = "3",
                              LastModifiedDate,
                              LastModifiedBy,
                              OsrId
                          });

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.OspReceiptDetail SET
                                ConfirmStatus = @ConfirmStatus,
                                ConfirmUserId = @ConfirmUserId,
                                ConfirmUser = @ConfirmUser,
                                ApproveStatus = @ApproveStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE OsrId = @OsrId";
                        dynamicParameters.AddDynamicParams(
                          new
                          {
                              ConfirmStatus = "V",
                              ConfirmUserId = CreateBy,
                              ConfirmUser = ConfirmUserNo,
                              ApproveStatus = "3",
                              LastModifiedDate,
                              LastModifiedBy,
                              OsrId
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

        #region //UpdateBarcodeProcess -- 託外入庫條碼過站 -- Ann 2022-09-27
        public string UpdateBarcodeProcess(int OsrDetailId = -1)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //取得託外過站條碼資訊
                        DynamicParameters dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ConfirmStatus, a.ReceiptQty, a.MoProcessId, a.MoId, a.ProcessStatus
                                , ISNULL(c.BarcodeId, -1) BarcodeId
                                , c.BarcodeQty
                                , d.BarcodeStatus
                                , FORMAT(e.CreateDate, 'yyyy-MM-dd') StartDate, e.CreateBy StartUserId
                                FROM MES.OspReceiptDetail a
                                LEFT JOIN MES.OspReceiptBarcode b ON a.OsrDetailId = b.OsrDetailId
                                LEFT JOIN MES.OspBarcode c ON b.OspBarcodeId = c.OspBarcodeId
                                LEFT JOIN MES.Barcode d ON c.BarcodeId = d.BarcodeId
                                INNER JOIN MES.OspDetail e ON a.OspDetailId = e.OspDetailId
                                WHERE a.OsrDetailId = @OsrDetailId";
                        dynamicParameters.Add("OsrDetailId", OsrDetailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("託外入庫單資料有誤!");

                        string BarcodeStatus = "";
                        int rowsAffected = -1;
                        foreach (var item in result)
                        {
                            if (item.ProcessStatus != "N") throw new SystemException("已過站，無法重複過站!");
                            if (item.ConfirmStatus != "N") throw new SystemException("託外入庫單已確認，無法更改!");
                            if (item.ReceiptQty <= 0) throw new SystemException("尚未維護進貨數量!");
                            if (item.BarcodeId == -1) throw new SystemException("尚未維護託外入庫條碼!");

                            BarcodeStatus = item.BarcodeStatus;
                            #region //判斷條碼狀態
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                FROM MES.OspReceiptBarcode a
                                INNER JOIN MES.OspBarcode b ON a.OspBarcodeId = b.OspBarcodeId
                                INNER JOIN MES.OspReceiptDetail c ON a.OsrDetailId = c.OsrDetailId
                                INNER JOIN MES.ManufactureOrder d ON c.MoId = d.MoId
                                INNER JOIN MES.MoProcess e ON c.MoId = e.MoId
                                INNER JOIN MES.BarcodeProcess f ON f.BarcodeId = b.BarcodeId AND f.MoId = c.MoId AND f.MoProcessId = c.MoProcessId
                                WHERE f.BarcodeId = @BarcodeId";
                            dynamicParameters.Add("BarcodeId", item.BarcodeId);

                            var result2 = sqlConnection.Query(sql, dynamicParameters);
                            if (result2.Count() > 0) throw new SystemException("此託外入庫條碼已過站，無法重複過站!");
                            #endregion

                            #region //取得下一站製程
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.MoProcessId, a.SortNumber
                                    , ISNULL(b.MoProcessId, -1) NextMoProcessId, ISNULL(b.SortNumber, -1) NextSortNumber
                                    , ISNULL(c.ProcessCheckStatus, 'Y') ProcessCheckStatus
                                    FROM MES.MoProcess a
                                    OUTER APPLY(
                                      SELECT ba.MoProcessId, ba.SortNumber
                                      FROM MES.MoProcess ba
                                      WHERE ba.MoId = @MoId AND SortNumber = a.SortNumber + 1
                                    ) b
                                    LEFT JOIN MES.ProcessParameter c ON a.ProcessId = c.ProcessId
                                    WHERE a.MoId = @MoId AND a.MoProcessId = @MoProcessId
                                    ORDER BY a.SortNumber";
                            dynamicParameters.Add("MoId", item.MoId);
                            dynamicParameters.Add("MoProcessId", item.MoProcessId);

                            var result4 = sqlConnection.Query(sql, dynamicParameters);

                            int NextMoProcessId = -1;
                            string ProcessCheckStatus = "";
                            foreach (var item2 in result4)
                            {
                                NextMoProcessId = item2.NextMoProcessId;
                                ProcessCheckStatus = item2.ProcessCheckStatus;
                                if (NextMoProcessId == -1) BarcodeStatus = "0"; //若為最後一站，則狀態改為完工
                            }
                            #endregion

                            #region //過站核心主程式
                            #region //先UPDATE MES.Barcode
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE MES.Barcode SET
                                    CurrentMoProcessId = @CurrentMoProcessId,
                                    NextMoProcessId = @NextMoProcessId,
                                    BarcodeStatus = @BarcodeStatus,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE BarcodeId = @BarcodeId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    CurrentMoProcessId = item.MoProcessId,
                                    NextMoProcessId,
                                    BarcodeStatus,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    item.BarcodeId
                                });

                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //INSERT MES.BarcodeProcess
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.BarcodeProcess (BarcodeId, MoId, MachineId, MoProcessId, ProdStatus, IpqcStatus, StationQty, PassQty, NgQty
                                    , StartDate, FinishDate, ModeShiftId, JigId
                                    , StartUserId, FinishUserId, CreateDate, CreateBy)
                                    VALUES (@BarcodeId, @MoId, @MachineId, @MoProcessId, @ProdStatus, @IpqcStatus, @StationQty, @PassQty, @NgQty
                                    , @StartDate, @FinishDate, @ModeShiftId, @JigId
                                    , @StartUserId, @FinishUserId, @CreateDate, @CreateBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    item.BarcodeId,
                                    item.MoId,
                                    MachineId = -1,
                                    item.MoProcessId,
                                    ProdStatus = 'P',
                                    IpqcStatus = ProcessCheckStatus == "Y" ? '0' : '1',
                                    StationQty = item.BarcodeQty,
                                    PassQty = item.BarcodeQty,
                                    NgQty = 0,
                                    item.StartDate,
                                    FinishDate = CreateDate,
                                    ModeShiftId = -1,
                                    JigId = -1,
                                    item.StartUserId,
                                    FinishUserId = CreateBy,
                                    CreateDate,
                                    CreateBy
                                });

                            var insertResult = sqlConnection.Query(sql, dynamicParameters);
                            #endregion
                            #endregion
                        }
                        #endregion

                        #region //回寫MES.OspReceiptDetail狀態
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.OspReceiptDetail SET
                                ProcessStatus = @ProcessStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE OsrDetailId = @OsrDetailId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ProcessStatus = "Y",
                                LastModifiedDate,
                                LastModifiedBy,
                                OsrDetailId
                            });

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //UpdateOsrQuantity -- 更新託外進貨單頭數量 -- Ann 2023-04-01
        public string UpdateOsrQuantity(int OsrDetailId = -1)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認託外入庫單單頭資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.OsrId, a.AcceptQty
                                FROM MES.OspReceiptDetail a
                                WHERE OsrDetailId = @OsrDetailId";
                        dynamicParameters.Add("OsrDetailId", OsrDetailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("託外入庫單詳細資料有誤!");

                        int OsrId = -1;
                        int AcceptQty = -1;
                        foreach (var item in result)
                        {
                            OsrId = item.OsrId;
                            AcceptQty = item.AcceptQty;
                        }
                        #endregion

                        #region //計算目前此託外入庫單總數量
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT SUM(a.AcceptQty) SumAcceptQty
                                FROM MES.OspReceiptDetail a
                                INNER JOIN MES.OspReceipt b ON a.OsrId = b.OsrId
                                WHERE a.OsrId = @OsrId";
                        dynamicParameters.Add("OsrId", OsrId);
                        var SumAcceptQtyResult = sqlConnection.Query(sql, dynamicParameters);

                        int SumAcceptQty = -1;
                        foreach (var item in SumAcceptQtyResult)
                        {
                            SumAcceptQty = item.SumAcceptQty;
                        }
                        #endregion

                        #region //UPDATE MES.OspReceipt Quantity
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.OspReceipt SET
                                Quantity = @SumAcceptQty,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE OsrId = @OsrId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SumAcceptQty,
                                LastModifiedDate,
                                LastModifiedBy,
                                OsrId
                            });

                        var rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //UpdateOsrDetailScriptQty -- 更新託外入庫單身報廢數量 -- Ann 2024-07-19
        public string UpdateOsrDetailScriptQty(int OsrDetailId, double ScriptQty)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認託外入庫單單頭資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.OsrId, a.ReceiptQty, a.AvailableQty, a.ScriptQty OriScriptQty 
                                FROM MES.OspReceiptDetail a 
                                WHERE a.OsrDetailId = @OsrDetailId";
                        dynamicParameters.Add("OsrDetailId", OsrDetailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("託外入庫單詳細資料有誤!");

                        int OsrId = 0;
                        double ReceiptQty = 0;
                        double AvailableQty = 0;
                        double OriScriptQty = 0;
                        foreach (var item in result)
                        {
                            OsrId = item.OsrId;
                            ReceiptQty = item.ReceiptQty;
                            AvailableQty = item.AvailableQty;
                            OriScriptQty = item.OriScriptQty;
                        }
                        #endregion

                        #region //數量卡控檢核
                        if (ScriptQty > ReceiptQty)
                        {
                            throw new SystemException("報廢數量 > 進貨數量!!");
                        }
                        #endregion

                        #region //Update MES.OspReceiptDetail
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.OspReceiptDetail SET
                                AcceptQty = ReceiptQty - @ScriptQty,
                                AvailableQty = ReceiptQty - @ScriptQty,
                                ScriptQty = @ScriptQty,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE OsrDetailId = @OsrDetailId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ScriptQty,
                                LastModifiedDate,
                                LastModifiedBy,
                                OsrDetailId
                            });

                        var rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //更新託外入庫單頭資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a 
                                SET Quantity = (
                                    SELECT SUM(x.AcceptQty) 
                                    FROM MES.OspReceiptDetail x
                                    WHERE x.OsrId = a.OsrId
                                ),
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                FROM MES.OspReceipt a
                                WHERE a.OsrId = @OsrId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                LastModifiedDate,
                                LastModifiedBy,
                                OsrId
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

        #region //UpdateOspVoid -- 託外生產單作廢 -- Ann 2025-05-20
        public string UpdateOspVoid(int OspId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認託外生產單資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.Status
                                FROM MES.OutsourcingProduction a 
                                WHERE a.OspId = @OspId";
                        dynamicParameters.Add("OspId", OspId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("託外生產單資料有誤!");

                        foreach (var item in result)
                        {
                            if (item.Status != "E" && item.Status != "Y") throw new SystemException("託外生產單狀態無法作廢!");
                        }
                        #endregion

                        #region //確認是否已開立託外入庫單
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.OspReceiptDetail a
                                INNER JOIN MES.OspDetail b ON a.OspDetailId = b.OspDetailId
                                WHERE b.OspId = @OspId";
                        dynamicParameters.Add("OspId", OspId);

                        var ospReceiptDetailResult = sqlConnection.Query(sql, dynamicParameters);

                        if (ospReceiptDetailResult.Count() > 0) throw new SystemException("此托外生產單已開立託外入庫單，無法作廢!");
                        #endregion

                        #region //Update MES.OspReceiptDetail
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.OutsourcingProduction SET
                                Status = 'V',
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE OspId = @OspId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                LastModifiedDate,
                                LastModifiedBy,
                                OspId
                            });

                        var rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //將此托外生產下的所有單身條碼刪除
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a 
                                FROM MES.OspBarcode a
                                INNER JOIN MES.OspDetail b ON a.OspDetailId = b.OspDetailId
                                WHERE b.OspId = @OspId";
                        dynamicParameters.Add("OspId", OspId);

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

        #region //UpdateOspConfirmBatch -- 確認託外生產單資料(批量) -- Shintokuro 2024-07-26
        public string UpdateOspConfirmBatch(string OspIdList)
        {
            try
            {
                int rowsAffected = 0;
                List<int> OspIdListData = OspIdList.Split(',').Select(int.Parse).ToList();

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        foreach (var OspId in OspIdListData)
                        {
                            #region //判斷託外生產單資料是否存在
                            DynamicParameters dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 Status
                                FROM MES.OutsourcingProduction
                                WHERE OspId = @OspId";
                            dynamicParameters.Add("OspId", OspId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("託外生產單資料找不到,請重新確認!");
                            foreach (var item in result)
                            {
                                if (item.Status != "N") throw new SystemException("託外生產單已確認!");
                            }
                            #endregion

                            #region //判斷託外生產單詳細資料是否有誤
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.OspQty, a.SuppliedQty
                                , b.OspBarcodeId
                                FROM MES.OspDetail a
                                LEFT JOIN MES.OspBarcode b ON a.OspDetailId = b.OspDetailId
                                WHERE a.OspId = @OspId";
                            dynamicParameters.Add("OspId", OspId);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("尚未建立託外生產單詳細資料!");

                            foreach (var item in result)
                            {
                                if (item.OspQty == 0) throw new SystemException("託外生產數量不能為0!");
                                if (item.OspBarcodeId == null) throw new SystemException("尚未新增託外條碼!");
                            }
                            #endregion

                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE MES.OutsourcingProduction SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE OspId = @OspId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    Status = "Y",
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    OspId
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                            #region //更新MoProcess工程檢參數
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE MES.MoProcess
                                   SET ProcessCheckStatus = b.ProcessCheckStatus,
                                       ProcessCheckType = b.ProcessCheckType,
                                       LastModifiedDate = @LastModifiedDate,
                                       LastModifiedBy = @LastModifiedBy
                                  FROM MES.MoProcess a
                                       INNER JOIN MES.OspDetail b ON a.MoProcessId = b.MoProcessId
	                                   INNER JOIN MES.OutsourcingProduction c ON b.OspId = c.OspId
                                 WHERE c.OspId = @OspId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    OspId
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

        #region //UpdateOspReConfirmBatch -- 反確認託外生產單資料(批量) -- Shintokuro 2024-07-26
        public string UpdateOspReConfirmBatch(string OspIdList)
        {
            try
            {
                int rowsAffected = 0;
                List<int> OspIdListData = OspIdList.Split(',').Select(int.Parse).ToList();

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        foreach (var OspId in OspIdListData)
                        {
                            #region //判斷託外生產單資料是否有誤
                            DynamicParameters dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.Status
                                , b.OspDetailId, b.MoProcessId
                                FROM MES.OutsourcingProduction a
                                INNER JOIN MES.OspDetail b ON a.OspId = b.OspId
                                WHERE a.OspId = @OspId";
                            dynamicParameters.Add("OspId", OspId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("託外生產單詳細資料有誤!");

                            int OspDetailId = -1;
                            int MoProcessId = -1;
                            foreach (var item in result)
                            {
                                if (item.Status != "Y") throw new SystemException("託外生產單尚未確認!");
                                OspDetailId = item.OspDetailId;
                                MoProcessId = item.MoProcessId;


                                #region //判斷條碼狀態是否有誤
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT b.NextMoProcessId,
                                       (SELECT COUNT(1) BPCount
                                          FROM MES.BarcodeProcess b
		                                       INNER JOIN MES.OspDetail c ON b.MoProcessId = c.MoProcessId
                                         WHERE c.OspDetailId = a.OspDetailId) BPCount
                                 FROM MES.OspBarcode a
                                      INNER JOIN MES.Barcode b ON a.BarcodeNo = b.BarcodeNo
                                WHERE a.OspDetailId = @OspDetailId";
                                dynamicParameters.Add("OspDetailId", OspDetailId);

                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("託外生產條碼有誤!");

                                foreach (var item1 in result)
                                {
                                    if (Convert.ToInt32(item1.BPCount) > 0) throw new SystemException("託外生產條碼已過站，無法反確認!");
                                }
                                #endregion
                            }
                            #endregion



                            #region //檢查此託外生產單詳細資料是否已綁定託外入庫單
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                FROM MES.OutsourcingProduction a
                                INNER JOIN MES.OspDetail b ON a.OspId = b.OspId
                                INNER JOIN MES.OspReceiptDetail c ON b.OspDetailId = c.OspDetailId
                                WHERE a.OspId = @OspId";
                            dynamicParameters.Add("OspId", OspId);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() > 0) throw new SystemException("此託外生產單已被綁定，無法反確認!");
                            #endregion

                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE MES.OutsourcingProduction SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE OspId = @OspId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    Status = "N",
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    OspId
                                });

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
        #endregion

        #region //Delete
        #region //DeleteOutsourcingProduction -- 刪除託外生產單資料 -- Ann 2022-09-07
        public string DeleteOutsourcingProduction(int OspId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷託外生產單資料是否正確
                        sql = @"SELECT TOP 1 Status  
                                FROM MES.OutsourcingProduction
                                WHERE OspId = @OspId";
                        dynamicParameters.Add("OspId", OspId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("託外生產單資料錯誤!");

                        foreach (var item in result)
                        {
                            if (item.Status != "N") throw new SystemException("託外生產單已確認或拋轉BPM，無法刪除!");
                        }
                        #endregion

                        #region //先刪除關聯Table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM MES.OspBarcode a
                                INNER JOIN MES.OspDetail b ON a.OspDetailId = b.OspDetailId
                                WHERE b.OspId = @OspId";
                        dynamicParameters.Add("OspId", OspId);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE FROM MES.OspDetail
                                WHERE OspId = @OspId";
                        dynamicParameters.Add("OspId", OspId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.OutsourcingProduction
                                WHERE OspId = @OspId";
                        dynamicParameters.Add("OspId", OspId);

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

        #region //DeleteOspBarcode -- 刪除託外生產條碼 -- Ann 2022-09-08
        public string DeleteOspBarcode(string OspBarcodeId)
        {
            try
            {
                if (OspBarcodeId.Length <= 0) throw new SystemException("託外生產條碼資訊錯誤!");
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();

                        var ospBarcodeIdList = OspBarcodeId.Split(',');

                        foreach (var ospBarcodeId in ospBarcodeIdList)
                        {
                            #region //判斷託外生產條碼資料是否正確
                            sql = @"SELECT TOP 1 a.BarcodeQty
                                    , b.OspDetailId, b.OspQty, b.SuppliedQty
                                    , c.[Status]
                                    FROM MES.OspBarcode a
                                    INNER JOIN MES.OspDetail b ON a.OspDetailId = b.OspDetailId
                                    INNER JOIN MES.OutsourcingProduction c ON b.OspId = c.OspId
                                    WHERE a.OspBarcodeId = @OspBarcodeId";
                            dynamicParameters.Add("OspBarcodeId", ospBarcodeId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("託外生產單資料錯誤!");

                            int BarcodeQty = -1;
                            int OspDetailId = -1;
                            int OspQty = -1;
                            int SuppliedQty = -1;
                            foreach (var item in result)
                            {
                                if (item.Status != "N" && item.Status != "E") throw new SystemException("託外生產單已確認，無法刪除!");
                                BarcodeQty = item.BarcodeQty;
                                OspDetailId = item.OspDetailId;
                                OspQty = item.OspQty;
                                SuppliedQty = item.SuppliedQty;
                            }
                            #endregion

                            #region //刪除主要table
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE MES.OspBarcode
                                    WHERE OspBarcodeId = @OspBarcodeId";
                            dynamicParameters.Add("OspBarcodeId", ospBarcodeId);

                            int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //同步更新OspDetail託外生產數量、供貨數量
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE MES.OspDetail SET
                                    OspQty = OspQty - @OspQty,
                                    SuppliedQty = SuppliedQty - @SuppliedQty,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE OspDetailId = @OspDetailId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    OspQty = BarcodeQty,
                                    SuppliedQty = BarcodeQty,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    OspDetailId
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

        #region //DeleteOspDetail -- 刪除託外生產詳細資料 -- Ann 2022-09-12
        public string DeleteOspDetail(int OspDetailId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷託外生產詳細資料是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM MES.OspDetail a
                                INNER JOIN MES.OutsourcingProduction b ON a.OspId = b.OspId
                                WHERE OspDetailId = @OspDetailId";
                        dynamicParameters.Add("OspDetailId", OspDetailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("託外生產單詳細資料錯誤!");

                        foreach (var item in result)
                        {
                            if (item.Status != "N" && item.Status != "E") throw new SystemException("該託外單已確認，無法更改!");
                        }
                        #endregion

                        #region //先刪除關聯Table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE FROM MES.OspBarcode
                                WHERE OspDetailId = @OspDetailId";
                        dynamicParameters.Add("OspDetailId", OspDetailId);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.OspDetail
                                WHERE OspDetailId = @OspDetailId";
                        dynamicParameters.Add("OspDetailId", OspDetailId);

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

        #region //DeleteOspReceipt -- 刪除託外入庫資料 -- Ann 2022-09-16
        public string DeleteOspReceipt(int OsrId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
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
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //判斷託外入庫資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.ConfirmStatus, a.TransferStatus, a.OsrErpPrefix, a.OsrErpNo
                                    FROM MES.OspReceipt a
                                    WHERE OsrId = @OsrId";
                            dynamicParameters.Add("OsrId", OsrId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("託外入庫資料錯誤!");

                            string OsrErpPrefix = "";
                            string OsrErpNo = "";
                            foreach (var item in result)
                            {
                                OsrErpPrefix = item.OsrErpPrefix;
                                OsrErpNo = item.OsrErpNo;
                                //if (item.TransferStatus != "N") throw new SystemException("該託外入庫單已拋轉，無法刪除單據!!");
                            }
                            #endregion

                            #region //檢查此托外入庫單是否存在於ERP
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MOCTH
                                    WHERE TH001 = @TH001
                                    AND TH002 = @TH002";
                            dynamicParameters.Add("TH001", OsrErpPrefix);
                            dynamicParameters.Add("TH002", OsrErpNo);

                            var MOCTHResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (MOCTHResult.Count() > 0) throw new SystemException("此單據已拋轉ERP，無法刪除!!");
                            #endregion

                            #region //檢查是否有單身已過站
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.ProcessStatus
                                FROM MES.OspReceiptDetail a
                                WHERE OsrId = @OsrId";
                            dynamicParameters.Add("OsrId", OsrId);

                            var result2 = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item in result2)
                            {
                                if (item.ProcessStatus != "N") throw new SystemException("託外入庫條碼已過站，無法更改!");
                            }
                            #endregion

                            #region //先刪除關聯Table
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE a
                                FROM MES.OspReceiptBarcode a
                                INNER JOIN MES.OspReceiptDetail b ON a.OsrDetailId = b.OsrDetailId
                                WHERE b.OsrId = @OsrId";
                            dynamicParameters.Add("OsrId", OsrId);

                            int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE FROM MES.OspReceiptDetail
                                WHERE OsrId = @OsrId";
                            dynamicParameters.Add("OsrId", OsrId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //刪除主要table
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE MES.OspReceipt
                                WHERE OsrId = @OsrId";
                            dynamicParameters.Add("OsrId", OsrId);

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

        #region //DeleteOspReceiptDetail -- 刪除託外入庫詳細資料 -- Ann 2022-09-19
        public string DeleteOspReceiptDetail(int OsrDetailId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷託外入庫詳細資料是否正確
                        sql = @"SELECT a.ConfirmStatus, a.ProcessStatus, a.AcceptQty, a.OsrId
                                , OrigAmount, OrigTaxAmt, PreTaxAmt, TaxAmt, OrigPreTaxAmt
                                FROM MES.OspReceiptDetail a
                                WHERE OsrDetailId = @OsrDetailId";
                        dynamicParameters.Add("OsrDetailId", OsrDetailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("託外入庫詳細資料錯誤!");

                        int AcceptQty = -1;
                        int OsrId = -1;
                        double OrigAmount = -1;
                        double OrigTaxAmt = -1;
                        double PreTaxAmt = -1;
                        double TaxAmt = -1;
                        double OrigPreTaxAmt = -1;
                        foreach (var item in result)
                        {
                            if (item.ConfirmStatus != "N") throw new SystemException("託外入庫單已確認，無法更改!");
                            if (item.ProcessStatus != "N")
                            {
                                #region //確認過站模式是否為供應商過站
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM MES.OspReceiptBarcode x 
                                        INNER JOIN MES.OspBarcode xa ON x.OspBarcodeId = xa.OspBarcodeId
                                        INNER JOIN MES.Barcode xb ON xa.BarcodeNo = xb.BarcodeNo
                                        INNER JOIN MES.BarcodeProcess xc ON xb.BarcodeId = xc.BarcodeId
                                        INNER JOIN MES.OspDetail xd ON xa.OspDetailId = xd.OspDetailId
                                        INNER JOIN BAS.[User] xe ON xc.FinishUserId = xe.UserId
                                        WHERE x.OsrDetailId = @OsrDetailId
                                        AND xc.MoProcessId = xd.MoProcessId
                                        AND xe.UserNo = 'Z100'";
                                dynamicParameters.Add("OsrDetailId", OsrDetailId);

                                var OspReceiptBarcodeResult3 = sqlConnection.Query(sql, dynamicParameters);

                                if (OspReceiptBarcodeResult3.Count() <= 0) throw new SystemException("託外入庫碼已過站，無法更改!");
                                #endregion
                            }
                            AcceptQty = item.AcceptQty;
                            OsrId = item.OsrId;
                            OrigAmount = item.OrigAmount;
                            OrigTaxAmt = item.OrigTaxAmt;
                            PreTaxAmt = item.PreTaxAmt;
                            TaxAmt = item.TaxAmt;
                            OrigPreTaxAmt = item.OrigPreTaxAmt;
                        }
                        #endregion

                        #region //先刪除關聯Table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.OspReceiptBarcode
                                WHERE OsrDetailId = @OsrDetailId";
                        dynamicParameters.Add("OsrDetailId", OsrDetailId);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.OspReceiptDetail
                                WHERE OsrDetailId = @OsrDetailId";
                        dynamicParameters.Add("OsrDetailId", OsrDetailId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //更新託外入庫單頭資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.OspReceipt SET
                                OrigAmount = OrigAmount - @OrigAmount,
                                OrigTax = OrigTax - @OrigTaxAmt,
                                Quantity = Quantity - @AcceptQty,
                                PretaxAmount = PretaxAmount - @PreTaxAmt,
                                TaxAmount = TaxAmount - @TaxAmt,
                                OrigPreTaxAmount = OrigPreTaxAmount - @OrigPreTaxAmt,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE OsrId = @OsrId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                OrigAmount,
                                OrigTaxAmt,
                                AcceptQty,
                                PreTaxAmt,
                                TaxAmt,
                                OrigPreTaxAmt,
                                LastModifiedDate,
                                LastModifiedBy,
                                OsrId
                            });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //重新調整單身序號
                        dynamicParameters = new DynamicParameters();
                        sql = @"With UpdateSort As
                                (
                                    SELECT OsrSeq,
                                    ROW_NUMBER() OVER(ORDER BY OsrSeq) NewOsrSeq
                                    FROM MES.OspReceiptDetail
                                    WHERE OsrId = @OsrId
                                )
                                UPDATE MES.OspReceiptDetail
                                SET OsrSeq = Right('0000' + Cast(NewOsrSeq as varchar), 4)
                                FROM MES.OspReceiptDetail
                                INNER JOIN UpdateSort ON MES.OspReceiptDetail.OsrSeq = UpdateSort.OsrSeq
                                WHERE OsrId = @OsrId";
                        dynamicParameters.Add("OsrId", OsrId);

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

        #region //DeleteOspReceiptBarcode -- 刪除託外入庫條碼資料 -- Ann 2022-09-20
        public string DeleteOspReceiptBarcode(int OsrDetailId, int OspBarcodeId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();

                        #region //判斷託外入庫詳細資料是否正確
                        sql = @"SELECT a.OsrId, a.ConfirmStatus, a.ReceiptQty, a.AcceptQty, a.AvailableQty, a.ProcessStatus, a.ScriptQty, a.ReturnQty
                                , a.OrigAmount, a.OrigTaxAmt, a.PreTaxAmt, a.TaxAmt
                                , c.PassStationControl
                                , d.MoProcessId
                                FROM MES.OspReceiptDetail a
                                INNER JOIN MES.OspReceipt b ON a.OsrId = b.OsrId
                                INNER JOIN SCM.Supplier c ON b.SupplierId = c.SupplierId
                                INNER JOIN MES.OspDetail d ON a.OspDetailId = d.OspDetailId
                                WHERE a.OsrDetailId = @OsrDetailId";
                        dynamicParameters.Add("OsrDetailId", OsrDetailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("託外入庫詳細資料錯誤!");

                        int OsrId = -1;
                        int ReceiptQty = -1;
                        int AcceptQty = -1;
                        int AvailableQty = -1;
                        int ScriptQty = -1;
                        int ReturnQty = -1;
                        double OrigAmount = -1;
                        double OrigTaxAmt = -1;
                        double PreTaxAmt = -1;
                        double TaxAmt = -1;
                        int MoProcessId = -1;
                        foreach (var item in result)
                        {
                            if (item.ConfirmStatus != "N") throw new SystemException("託外入庫單已確認，無法更改!");
                            if (item.ProcessStatus != "N" && item.PassStationControl != "Y")
                            {
                                throw new SystemException("託外入庫碼已過站，無法更改!");
                            }
                            if (item.ReceiptQty <= 0 || item.AcceptQty <= 0 || item.AvailableQty <= 0) throw new SystemException("託外入庫數量異常!");

                            OsrId = item.OsrId;
                            ReceiptQty = item.ReceiptQty;
                            AcceptQty = item.AcceptQty;
                            AvailableQty = item.AvailableQty;
                            ScriptQty = item.ScriptQty;
                            ReturnQty = item.ReturnQty;
                            OrigAmount = item.OrigAmount;
                            OrigTaxAmt = item.OrigTaxAmt;
                            PreTaxAmt = item.PreTaxAmt;
                            TaxAmt = item.TaxAmt;
                            MoProcessId = item.MoProcessId;
                        }
                        #endregion

                        #region //判斷託外生產條碼資料是否正確
                        sql = @"SELECT a.BarcodeQty, a.OspDetailId
                                FROM MES.OspBarcode a
                                WHERE a.OspBarcodeId = @OspBarcodeId";
                        dynamicParameters.Add("OspBarcodeId", OspBarcodeId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("託外生產條碼資料錯誤!");

                        int BarcodeQty = -1;
                        int OspDetailId = -1;
                        foreach (var item in result2)
                        {
                            BarcodeQty = item.BarcodeQty;
                            OspDetailId = item.OspDetailId;
                        }
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.OspReceiptBarcode
                                WHERE OsrDetailId = @OsrDetailId
                                AND OspBarcodeId = @OspBarcodeId";
                        dynamicParameters.Add("OsrDetailId", OsrDetailId);
                        dynamicParameters.Add("OspBarcodeId", OspBarcodeId);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //確認目前此單身條碼是否已都過站，若已過站則更改單據狀態
                        string ProcessStatus = "N";
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT b.BarcodeNo
                                 , (
                                    SELECT x.FinishDate
                                    FROM MES.BarcodeProcess x 
                                    INNER JOIN MES.Barcode xa ON x.BarcodeId = xa.BarcodeId
                                    INNER JOIN BAS.[User] xb ON x.CreateBy = xb.UserId
                                    WHERE xa.BarcodeNo = b.BarcodeNo
                                    AND x.MoProcessId = c.MoProcessId
                                    FOR JSON PATH, ROOT('data')
                                ) BarcodeProcess
                                FROM MES.OspReceiptBarcode a 
                                INNER JOIN MES.OspBarcode b ON a.OspBarcodeId = b.OspBarcodeId
                                INNER JOIN MES.OspDetail c ON b.OspDetailId = c.OspDetailId
                                WHERE c.OspDetailId = @OspDetailId";
                        dynamicParameters.Add("OspDetailId", OspDetailId);

                        var OspReceiptBarcodeResult2 = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item2 in OspReceiptBarcodeResult2)
                        {
                            if (item2.BarcodeProcess == null)
                            {
                                ProcessStatus = "N";
                                break;
                            }
                            else
                            {
                                JObject barcodeProcess = JObject.Parse(item2.BarcodeProcess);
                                bool finishDataFlag = false;
                                foreach (var item3 in barcodeProcess["data"])
                                {
                                    if (item3["FinishDate"] == null)
                                    {
                                        finishDataFlag = true;
                                        break;
                                    }
                                }

                                if (finishDataFlag == true)
                                {
                                    ProcessStatus = "N";
                                    break;
                                }
                                else
                                {
                                    #region //確認此條碼是否下一站已指定回此托外站別
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM MES.Barcode a 
                                            WHERE a.BarcodeNo = @BarcodeNo
                                            AND a.NextMoProcessId = @MoProcessId";
                                    dynamicParameters.Add("BarcodeNo", item2.BarcodeNo);
                                    dynamicParameters.Add("MoProcessId", MoProcessId);

                                    var BarcodeCheckResult = sqlConnection.Query(sql, dynamicParameters);

                                    if (BarcodeCheckResult.Count() > 0)
                                    {
                                        ProcessStatus = "N";
                                    }
                                    else
                                    {
                                        ProcessStatus = "Y";
                                    }
                                    #endregion

                                    break;
                                }
                            }
                        }
                        #endregion

                        #region //同步更新OspReceiptDetail入庫數量
                        //從驗收數量開始嘗試扣，以此類推
                        if (AcceptQty - BarcodeQty >= 0) AcceptQty -= BarcodeQty;
                        else if (ScriptQty - BarcodeQty >= 0) ScriptQty -= BarcodeQty;
                        else if (ReturnQty - BarcodeQty >= 0) ReturnQty -= BarcodeQty;
                        else throw new SystemException("進貨數量異常!!");

                        //若原計價數量超過目前進貨數量，自動修正為進貨數量
                        if (AvailableQty > (ReceiptQty - BarcodeQty)) AvailableQty = ReceiptQty - BarcodeQty;

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.OspReceiptDetail SET
                                ReceiptQty = ReceiptQty - @BarcodeQty,
                                AcceptQty = @AcceptQty,
                                ScriptQty = @ScriptQty,
                                ReturnQty = @ReturnQty,
                                AvailableQty = @AvailableQty,
                                ProcessStatus = @ProcessStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE OsrDetailId = @OsrDetailId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                BarcodeQty,
                                AcceptQty,
                                ScriptQty,
                                ReturnQty,
                                AvailableQty,
                                ProcessStatus,
                                LastModifiedDate,
                                LastModifiedBy,
                                OsrDetailId
                            });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region // 調整單頭總驗收數量(Quantity)

                        #region //計算目前此託外入庫單總數量
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT SUM(a.AcceptQty) SumAcceptQty
                                FROM MES.OspReceiptDetail a
                                INNER JOIN MES.OspReceipt b ON a.OsrId = b.OsrId
                                WHERE a.OsrId = @OsrId";
                        dynamicParameters.Add("OsrId", OsrId);
                        var SumAcceptQtyResult = sqlConnection.Query(sql, dynamicParameters);

                        int SumAcceptQty = -1;
                        foreach (var item in SumAcceptQtyResult)
                        {
                            SumAcceptQty = item.SumAcceptQty;
                        }
                        #endregion

                        #region //UPDATE MES.OspReceipt Quantity
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.OspReceipt SET
                                Quantity = @SumAcceptQty,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE OsrId = @OsrId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SumAcceptQty,
                                LastModifiedDate,
                                LastModifiedBy,
                                OsrId
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

        #region //DeleteOspBatch -- 刪除託外生產單資料(批量) -- Shintokuro 2024-07-26
        public string DeleteOspBatch(string OspIdList)
        {
            try
            {
                int rowsAffected = 0;
                List<int> OspIdListData = OspIdList.Split(',').Select(int.Parse).ToList();

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        foreach (var OspId in OspIdListData)
                        {
                            #region //判斷託外生產單資料是否正確
                            sql = @"SELECT TOP 1 Status  
                                    FROM MES.OutsourcingProduction
                                    WHERE OspId = @OspId";
                            dynamicParameters.Add("OspId", OspId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("託外生產單資料錯誤!");

                            foreach (var item in result)
                            {
                                if (item.Status != "N" && item.Status != "E") throw new SystemException("託外生產單已確認，無法刪除!");
                            }
                            #endregion

                            #region //先刪除關聯Table
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE a
                                FROM MES.OspBarcode a
                                INNER JOIN MES.OspDetail b ON a.OspDetailId = b.OspDetailId
                                WHERE b.OspId = @OspId";
                            dynamicParameters.Add("OspId", OspId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE FROM MES.OspDetail
                                WHERE OspId = @OspId";
                            dynamicParameters.Add("OspId", OspId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //刪除主要table
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE MES.OutsourcingProduction
                                WHERE OspId = @OspId";
                            dynamicParameters.Add("OspId", OspId);

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

        #region //DeleteOspDetailExcel -- 刪除託外生產詳細資料(Excel) -- Shintokuro 2024-07-26
        public string DeleteOspDetailExcel(string ExcelData)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        int rowsAffected = 0;
                        List<int> ExcelDataList = ExcelData.Split(',').Select(int.Parse).ToList();

                        DynamicParameters dynamicParameters = new DynamicParameters();

                        foreach (var OspDetailId in ExcelDataList)
                        {
                            #region //判斷託外生產詳細資料是否正確
                            sql = @"SELECT TOP 1 Status
                                FROM MES.OspDetail a
                                INNER JOIN MES.OutsourcingProduction b ON a.OspId = b.OspId
                                WHERE OspDetailId = @OspDetailId";
                            dynamicParameters.Add("OspDetailId", OspDetailId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("託外生產單詳細資料錯誤!");

                            foreach (var item in result)
                            {
                                if (item.Status != "N" && item.Status != "E") throw new SystemException("該托外單已確認，無法更改!");
                            }
                            #endregion

                            #region //先刪除關聯Table
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE FROM MES.OspBarcode
                                    WHERE OspDetailId = @OspDetailId";
                            dynamicParameters.Add("OspDetailId", OspDetailId);

                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //刪除主要table
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE MES.OspDetail
                                    WHERE OspDetailId = @OspDetailId";
                            dynamicParameters.Add("OspDetailId", OspDetailId);

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

        #endregion

        #region //BPM
        #region //TransferOspToBpm -- 拋轉托外生產單至BPM -- Ann 2025-04-10
        public string TransferOspToBpm(string OspIds)
        {
            try
            {
                string token = "";
                int rowsAffected = 0;
                string CompanyNo = "";

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //取得USER資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.UserNo
                                FROM BAS.[User] a
                                WHERE a.UserId = @UserId";
                    dynamicParameters.Add("UserId", CreateBy);

                    var UserResult = sqlConnection.Query(sql, dynamicParameters);

                    string UserNo = "";
                    foreach (var item in UserResult)
                    {
                        UserNo = item.UserNo;
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
                        string[] OspIdArray = OspIds.Split(',');
                        foreach (var OspId in OspIdArray)
                        {
                            using (TransactionScope transactionScope = new TransactionScope())
                            {
                                #region //判斷委外單資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.OspId, a.CompanyId, a.DepartmentId, a.OspNo, FORMAT(a.OspDate, 'yyyy-MM-dd') OspDate
                                    , a.SupplierId, a.OspStatus, a.OspDesc, a.[Status]
                                    , b.DepartmentNo, b.DepartmentName
                                    , ISNULL(c.SupplierNo, '') SupplierNo, ISNULL(c.SupplierShortName, '') SupplierShortName
                                    FROM MES.OutsourcingProduction a 
                                    INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                                    LEFT JOIN SCM.Supplier c ON a.SupplierId = c.SupplierId
                                    WHERE a.OspId = @OspId";
                                dynamicParameters.Add("OspId", OspId);

                                var result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("委外單【" + OspId + "】資料錯誤!");

                                string OspNo = "";
                                string DepartmentNo = "";
                                string OspDesc = "";
                                string OspStatus = "";
                                string OspDate = "";
                                string DepartmentName = "";
                                string SupplierNo = "";
                                string SupplierName = "";
                                foreach (var item in result)
                                {
                                    if (item.Status != "N" && item.Status != "E") throw new SystemException("委外單狀態無法拋轉!");
                                    OspNo = item.OspNo;
                                    DepartmentNo = item.DepartmentNo;
                                    OspDesc = item.OspDesc;
                                    OspStatus = item.OspStatus == "Y" ? "供料" : "不供料";
                                    OspDate = item.OspDate;
                                    DepartmentName = item.DepartmentName;
                                    SupplierNo = item.SupplierNo;
                                    SupplierName = item.SupplierShortName;
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

                                #region //取得單身資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.OspDetailId, a.OspId, a.MoId, a.MoProcessId, a.ProcessCheckStatus, a.ProcessCheckType, a.OspQty, a.SuppliedQty
                                        , ISNULL(a.ProcessCode, '') ProcessCode, ISNULL(a.ProcessCodeName, '') ProcessCodeName
                                        , FORMAT(a.ExpectedDate, 'yyyy-MM-dd') ExpectedDate
                                        , b.ProcessAlias
                                        , c.WoSeq
                                        , d.WoErpPrefix, d.WoErpNo
                                        , e.MtlItemId, e.MtlItemNo, e.MtlItemName, e.MtlItemSpec
                                        FROM MES.OspDetail a 
                                        INNER JOIN MES.MoProcess b ON a.MoProcessId = b.MoProcessId
                                        INNER JOIN MES.ManufactureOrder c ON a.MoId = c.MoId
                                        INNER JOIN MES.WipOrder d ON c.WoId = d.WoId
                                        INNER JOIN PDM.MtlItem e ON d.MtlItemId = e.MtlItemId
                                        WHERE a.OspId = @OspId 
                                        ORDER BY OspDetailId";
                                dynamicParameters.Add("OspId", OspId);

                                var ospDetailResult = sqlConnection.Query(sql, dynamicParameters);
                                if (ospDetailResult.Count() <= 0) throw new SystemException("委外單【" + OspId + "】單身資料錯誤!");
                                #endregion

                                #region //組單身資料
                                JArray tabPR_Line_Data = new JArray();

                                int count = 1;
                                foreach (var item in ospDetailResult)
                                {
                                    string woErpFullNo = item.WoErpPrefix + item.WoErpNo + "(" + item.WoSeq.ToString() + ")";
                                    string processCheckStatus = item.ProcessCheckStatus == "Y" ? "是" : "否";
                                    string processCheckType = item.ProcessCheckType == "1" ? "全檢" : "抽檢";

                                    tabPR_Line_Data.Add(JObject.FromObject(new
                                    {
                                        PRNumber_SN = count.ToString("D4"),
                                        ItemNo = item.MtlItemNo.ToString(),
                                        ItemName = item.MtlItemName.ToString(),
                                        Spec = item.MtlItemSpec.ToString(),
                                        WoErpFullNo = woErpFullNo,
                                        ProcessAlias = item.ProcessAlias.ToString(),
                                        OspQty = item.OspQty.ToString(),
                                        SuppliedQty = item.SuppliedQty.ToString(),
                                        ProcessCode = item.ProcessCode.ToString(),
                                        ExpectedDate = item.ExpectedDate.ToString(),
                                        ProcessCheckStatus = processCheckStatus,
                                        ProcessCheckType = processCheckType
                                    }));

                                    count++;
                                }
                                #endregion

                                #region //依公司別取得ProId
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.TypeName ProId
                                        FROM BAS.[Type] a
                                        WHERE a.TypeSchema = 'BPM.OspProId'
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

                                #region //組請購BPM資料
                                string memId = BpmUserId;
                                string rolId = BpmRoleId;
                                string startMethod = "NoOpFirst";

                                JObject artInsAppData = JObject.FromObject(new
                                {
                                    dbTable = "MES.OutsourcingProduction",
                                    Title = "MES託外生產單 = " + OspNo,
                                    OspDesc,
                                    OspStatus,
                                    OspDate,
                                    OspDeptName = DepartmentName,
                                    SupplierName,
                                    mesID = OspId,
                                    company = CompanyNo,
                                    SupplierNo,
                                    PRNumber_MES = OspNo,
                                    OspDeptNo = DepartmentNo,
                                    tabPR_Line_Data = JsonConvert.SerializeObject(tabPR_Line_Data),
                                });
                                #endregion

                                string sData = BpmHelper.PostFormToBpm(token, proId, memId, rolId, startMethod, artInsAppData, BpmServerPath);

                                if (sData == "true")
                                {
                                    #region //更改單據狀態
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE MES.OutsourcingProduction SET
                                            Status = 'P',
                                            TransferBpmDate = @TransferBpmDate,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE OspId = @OspId";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            TransferBpmDate = DateTime.Now,
                                            LastModifiedDate,
                                            LastModifiedBy,
                                            OspId
                                        });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                                else
                                {
                                    throw new SystemException("委外單【" + OspId + "】拋轉BPM失敗!");
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
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //AddOspBpmLog -- 新增委外單LOG紀錄 -- Ann 2025-04-11
        public string AddOspBpmLog(int OspId, string BpmNo, string Status, string RootId, string ConfirmUser)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認委外單資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.TransferBpmDate
                                FROM MES.OutsourcingProduction a
                                WHERE OspId = @OspId";
                        dynamicParameters.Add("OspId", OspId);

                        var OutsourcingProductionResult = sqlConnection.Query(sql, dynamicParameters);

                        if (OutsourcingProductionResult.Count() <= 0) throw new SystemException("委外單資料錯誤!!");

                        DateTime TransferBpmDate = new DateTime();
                        foreach (var item in OutsourcingProductionResult)
                        {
                            TransferBpmDate = item.TransferBpmDate;
                        }
                        #endregion

                        #region //INSERT MES.OspBpmLog
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.OspBpmLog (OspId, RootId, BpmNo, TransferBpmDate, BpmStatus, ConfirmUser
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.LogId
                                VALUES (@OspId, @RootId, @BpmNo, @TransferBpmDate, @BpmStatus, @ConfirmUser
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                          new
                          {
                              OspId,
                              RootId,
                              BpmNo,
                              TransferBpmDate,
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

        #region //UpdateOspStatus -- 更新委外單狀態 -- Ann 2025-04-11
        public string UpdateOspStatus(int OspId, string BpmNo, string Status, string RootId, string ConfirmUser)
        {
            try
            {
                int rowsAffected = 0;
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷委外單資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.OutsourcingProduction a
                                WHERE a.OspId = @OspId";
                        dynamicParameters.Add("OspId", OspId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("請購單資料錯誤!");
                        #endregion

                        #region //確認簽核者資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT UserId FROM BAS.[User]
                                WHERE UserNo = @UserNo";
                        dynamicParameters.Add("UserNo", ConfirmUser);

                        var UserResult = sqlConnection.Query(sql, dynamicParameters);

                        if (UserResult.Count() <= 0) throw new SystemException("簽核者資料錯誤!!");

                        int UserId = -1;
                        foreach (var item in UserResult)
                        {
                            UserId = item.UserId;
                        }
                        #endregion

                        #region //UPDATE SCM.PurchaseRequisition
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.OutsourcingProduction SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @UserId
                                WHERE OspId = @OspId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status,
                                LastModifiedDate,
                                UserId,
                                OspId
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

        #region //API
        #region //OspAlertMamo //託外逾時未加工MAMO通知 --GPAI 240828
        public string OspAlertMamo(string Company/*, string ChannelId*/)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //確認公司資訊
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.CompanyNo, a.CompanyId
                                FROM BAS.Company a
                                WHERE a.CompanyNo = @Company";
                    dynamicParameters.Add("Company", Company);

                    var CompanyResult = sqlConnection.Query(sql, dynamicParameters);

                    if (CompanyResult.Count() <= 0) throw new SystemException("公司資料錯誤!!");

                    string CompanyNo = "";
                    int CompanyId = 0;
                    foreach (var item in CompanyResult)
                    {
                        CompanyNo = item.CompanyNo;
                        CompanyId = item.CompanyId;
                    }
                    #endregion

                    #region //取得資料

                    //找出異常託外生產單號
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DISTINCT a.OspNo 
                            ,a.OspId, FORMAT(a.CreateDate, 'yyyy-MM-dd') DocDate
                            , b.OspDetailId, b.OspQty, FORMAT(b.ExpectedDate, 'yyyy-MM-dd') ExpectedDate, b.MoProcessId
                            , (d.WoErpPrefix + '-' + d.WoErpNo) WoErpFullNo
                            , e.MtlItemNo, e.MtlItemName
                            , f.SupplierName
                            , b.CreateDate , (DATEADD(HOUR,8,b.CreateDate)) fff
                            , h.ProcessAlias
                            FROM MES.OutsourcingProduction  a
                            INNER JOIN MES.OspDetail b ON b.OspId = a.OspId
                            INNER JOIN MES.ManufactureOrder c ON b.MoId = c.MoId
                            INNER JOIN MES.WipOrder d ON c.WoId = d.WoId
                            INNER JOIN PDM.MtlItem e ON d.MtlItemId = e.MtlItemId
                            INNER JOIN SCM.Supplier f ON a.SupplierId = f.SupplierId
                            INNER JOIN MES.OspBarcode g1 ON g1.OspDetailId = b.OspDetailId
                            INNER JOIN MES.Barcode g2 ON g1.BarcodeNo = g2.BarcodeNo
                            LEFT JOIN MES.OspReceiptBarcode g3 ON g3.OspBarcodeId = g1.OspBarcodeId
                            LEFT JOIN MES.BarcodeProcess g ON g.BarcodeId = g2.BarcodeId AND b.MoProcessId = g.MoProcessId
                            INNER JOIN MES.MoProcess h ON b.MoProcessId = h.MoProcessId
                            WHERE f.PassStationControl = 'Y'
                            AND (DATEADD(HOUR,8,b.CreateDate)) <= getdate()
                            AND g.FinishDate IS NULL
							AND b.CreateDate >= DATEADD(dd, -7, getdate()) 
							AND b.CreateDate <= getdate()
                            AND g3.OspBarcodeId IS NULL
                            AND a.[Status] = 'Y'
                            AND a.CompanyId = @CompanyId
                            ORDER BY a.OspId, f.SupplierName, e.MtlItemNo DESC";

                        dynamicParameters.Add("CompanyId", CompanyId);
                    //sql += @" ORDER BY wo.WoId DESC";
                    #endregion

                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() > 0)
                    {
                        #region //MAMO推播通知
                        string Content = "";

                        #region //取得清單
                        string barcodeInfoDesc = "| 託外單號 | 製令 | 品名 | 製程 | 託外數 | 預計回廠日 | 開單日期 | 加工廠商 |\n|--- | --- | --- | --- | --- | --- | --- |\n";
                        int count = 0;

                        foreach (var item in result)
                        {
                            barcodeInfoDesc += "| " + item.OspNo + " | " + item.WoErpFullNo + " | " + item.MtlItemName + " | " + item.ProcessAlias + " | " + item.OspQty + " | " + item.ExpectedDate + " | " + item.DocDate + " | " + item.SupplierName + " |\n";
                            count++;
                        }
                        #endregion
                        Content = "### 【託外加工逾時通知】\n" +
                                           "- 逾時單據如下:\n" + barcodeInfoDesc + "\n";
                        

                        #region //確認推播群組
                        //string SendId = "";
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ChannelId
                                FROM MAMO.Channels a
                                WHERE a.ChannelId = @SendId";
                        dynamicParameters.Add("SendId", 246);
                        //dynamicParameters.Add("PushType", CheckQcMeasureData);

                        var MamoChannelResult = sqlConnection.Query(sql, dynamicParameters);

                        if (MamoChannelResult.Count() <= 0) throw new SystemException("請確認是否已設定推播群組!!");


                        //if (QcMamoChannelResult.Count() <= 0) throw new SystemException("收貨確認/駁回推播群組資料錯誤!!<br>請確認生產模式【" + ModeName + "】是否已設定推播群組!!");

                        //foreach (var item in MamoChannelResult)
                        //{
                        //    ChannelId = item.ChannelId;
                        //    SendId = ChannelId.ToString();
                        //}
                        #endregion

                        #region //取得標記USER資料(原送測人員部門)
                        List<string> Tags = new List<string>();
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserId
                                , b.UserNo, b.UserName
                                FROM MAMO.ChannelMembers a 
                                INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                                WHERE a.ChannelId = @SendId";
                        dynamicParameters.Add("SendId", 246);

                        var UserResult = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in UserResult)
                        {
                            Tags.Add(item.UserNo);
                        }
                        #endregion

                        List<int> Files = new List<int>();

                        string MamoResult = mamoHelper.SendMessage(CompanyNo, 945, "Channel", "246", Content, Tags, Files);

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
                    else
                    {
                        throw new SystemException("查無資料！");
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

        #region //OspInAlertMamo //託外回廠前MAMO通知 --GPAI 240828
        public string OspInAlertMamo(string Company/*, string ChannelId*/)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //確認公司資訊
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.CompanyNo, a.CompanyId
                                FROM BAS.Company a
                                WHERE a.CompanyNo = @Company";
                    dynamicParameters.Add("Company", Company);

                    var CompanyResult = sqlConnection.Query(sql, dynamicParameters);

                    if (CompanyResult.Count() <= 0) throw new SystemException("公司資料錯誤!!");

                    string CompanyNo = "";
                    int CompanyId = 0;
                    foreach (var item in CompanyResult)
                    {
                        CompanyNo = item.CompanyNo;
                        CompanyId = item.CompanyId;
                    }
                    #endregion

                    #region //取得資料

                    //找出異常託外生產單號
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DISTINCT a.OspNo 
                            ,a.OspId, FORMAT(a.CreateDate, 'yyyy-MM-dd') DocDate
                            , b.OspDetailId, b.OspQty, FORMAT(b.ExpectedDate, 'yyyy-MM-dd') ExpectedDate, b.MoProcessId
                            , (d.WoErpPrefix + '-' + d.WoErpNo) WoErpFullNo
                            , e.MtlItemNo, e.MtlItemName
                            , f.SupplierName
                            , b.CreateDate , (DATEADD(DAY,-1,b.ExpectedDate)) fff
                            , h.ProcessAlias
                            FROM MES.OutsourcingProduction  a
                            INNER JOIN MES.OspDetail b ON b.OspId = a.OspId
                            INNER JOIN MES.ManufactureOrder c ON b.MoId = c.MoId
                            INNER JOIN MES.WipOrder d ON c.WoId = d.WoId
                            INNER JOIN PDM.MtlItem e ON d.MtlItemId = e.MtlItemId
                            INNER JOIN SCM.Supplier f ON a.SupplierId = f.SupplierId
                            INNER JOIN MES.OspBarcode g1 ON g1.OspDetailId = b.OspDetailId
                            INNER JOIN MES.Barcode g2 ON g1.BarcodeNo = g2.BarcodeNo
                            LEFT JOIN MES.OspReceiptBarcode g3 ON g3.OspBarcodeId = g1.OspBarcodeId
                            LEFT JOIN MES.BarcodeProcess g ON g.BarcodeId = g2.BarcodeId AND b.MoProcessId = g.MoProcessId
                            INNER JOIN MES.MoProcess h ON b.MoProcessId = h.MoProcessId
                            WHERE f.PassStationControl = 'Y'
							AND (DATEADD(DAY,-1,b.ExpectedDate)) <= getdate()
                            --AND g.FinishDate IS  NULL
							AND b.ExpectedDate >=  getdate()
							--AND b.ExpectedDate <= getdate()
                            AND g3.OspBarcodeId IS NULL
                            AND a.[Status] = 'Y'
                            AND a.CompanyId = @CompanyId
                            ORDER BY FORMAT(b.ExpectedDate, 'yyyy-MM-dd'), a.OspId, f.SupplierName, e.MtlItemNo DESC";

                    dynamicParameters.Add("CompanyId", CompanyId);
                    //sql += @" ORDER BY wo.WoId DESC";
                    #endregion

                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() > 0)
                    {
                        #region //MAMO推播通知
                        string Content = "";

                        #region //取得清單
                        string barcodeInfoDesc = "| 託外單號 | 製令 | 品名 | 製程 | 託外數 | 預計回廠日 | 開單日期 | 加工廠商 |\n|--- | --- | --- | --- | --- | --- | --- |\n";
                        int count = 0;

                        foreach (var item in result)
                        {
                            barcodeInfoDesc += "| " + item.OspNo + " | " + item.WoErpFullNo + " | " + item.MtlItemName + " | " + item.ProcessAlias + " | " + item.OspQty + " | " + item.ExpectedDate + " | " + item.DocDate + " | " + item.SupplierName + " |\n";
                            count++;
                        }
                        #endregion
                        Content = "### 【託外回廠前通知】\n" +
                                           "- 單據如下:\n" + barcodeInfoDesc + "\n";


                        #region //確認推播群組
                        //string SendId = "";
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ChannelId
                                FROM MAMO.Channels a
                                WHERE a.ChannelId = @SendId";
                        dynamicParameters.Add("SendId", 246);

                        var MamoChannelResult = sqlConnection.Query(sql, dynamicParameters);

                        if (MamoChannelResult.Count() <= 0) throw new SystemException("請確認是否已設定推播群組!!");

                        //int ChannelId = -1;
                        //foreach (var item in MamoChannelResult)
                        //{
                        //    ChannelId = item.ChannelId;
                        //    SendId = ChannelId.ToString();
                        //}
                        #endregion

                        #region //取得標記USER資料(原送測人員部門)
                        List<string> Tags = new List<string>();
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserId
                                , b.UserNo, b.UserName
                                FROM MAMO.ChannelMembers a 
                                INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                                WHERE a.ChannelId = @SendId";
                        dynamicParameters.Add("SendId", 246);

                        var UserResult = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in UserResult)
                        {
                            Tags.Add(item.UserNo);
                        }
                        #endregion

                        List<int> Files = new List<int>();

                        string MamoResult = mamoHelper.SendMessage(CompanyNo, 945, "Channel", "246", Content, Tags, Files);

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
                    else
                    {
                        throw new SystemException("查無資料！");
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
    }
}
