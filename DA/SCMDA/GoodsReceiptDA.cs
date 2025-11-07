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
using System.IO;
using System.Linq;
using System.Text;
using System.Transactions;
using System.Web;
using System.Text.RegularExpressions;


namespace SCMDA
{
    public class GoodsReceiptDA
    {
        public string MainConnectionStrings = "";
        public string ErpConnectionStrings = "";
        public string HrmConnectionStrings = "";
        public string ErpSysDbConnectionStrings = "";
        public string BpmDbConnectionStrings = "";
        public string SrmDbConnectionStrings = "";

        public int CurrentCompany = -1;
        public int CurrentUser = -1;
        public int CreateBy = -1;
        public int LastModifiedBy = -1;
        public DateTime CreateDate = default(DateTime);
        public DateTime LastModifiedDate = default(DateTime);

        public string sql = "";
        public JObject jsonResponse = new JObject();
        public Logger logger = LogManager.GetCurrentClassLogger();
        public SqlQuery sqlQuery = new SqlQuery();
        public DynamicParameters dynamicParameters = new DynamicParameters();

        public GoodsReceiptDA()
        {
            MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];
            HrmConnectionStrings = ConfigurationManager.AppSettings["HrmDb"];
            ErpSysDbConnectionStrings = ConfigurationManager.AppSettings["ErpSysDb"];
            BpmDbConnectionStrings = ConfigurationManager.AppSettings["BpmDb"];
            SrmDbConnectionStrings = ConfigurationManager.AppSettings["SrmDb"];

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

        #region //Get 
        #region //GetGoodsReceipt -- 取得進貨單資料 -- Ann 2023-02-24
        public string GetGoodsReceipt(int GrId, string GrErpFullNo, int SupplierId, int ConfirmUserId, string SearchKey, string ConfirmStatus, string ClosureStatus
            , string StartDate, string EndDate
            , string PoErpFullNo, int PoUserId, string PrErpFullNo, int PrUserId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.GrId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.CompanyId, a.GrErpPrefix, a.GrErpNo, GrErpPrefix + '-' + GrErpNo GrErpFullNo, a.ReceiptDate, a.SupplierId, a.SupplierNo
                        , a.SupplierSo, a.CurrencyCode, a.Exchange, a.InvoiceType, a.TaxType, a.InvoiceNo, a.PrintCnt, a.ConfirmStatus, a.DocDate
                        , a.RenewFlag, a.Remark, a.TotalAmount, a.DeductAmount, a.OrigTax, a.ReceiptAmount, a.SupplierName, a.UiNo, a.DeductType
                        , a.ObaccoAndLiquorFlag, a.RowCnt, a.Quantity, FORMAT(a.InvoiceDate, 'yyyy-MM-dd') InvoiceDate, a.OrigPreTaxAmount, a.ApplyYYMM
                        , a.TaxRate, a.PreTaxAmount, a.TaxAmount, a.PaymentTerm, a.PoId, a.PrepaidErpPrefix, a.PrepaidErpNo, a.OffsetAmount, a.TaxOffset, a.PackageQuantity
                        , a.PremiumAmount, a.ApproveStatus, a.InvoiceFlag, a.ReserveTaxCode, a.SendCount, a.OrigPremiumAmount, a.EbcErpPreNo, a.EbcEdition, a.EbcFlag
                        , a.FromDocType, a.NoticeFlag, a.TaxCode, a.TradeTerm, a.DetailMultiTax, a.TaxExchange, a.ContactUser, a.TransferStatus, a.TransferDate
                        , a.CreateDate, a.LastModifiedDate, a.CreateBy, a.LastModifiedBy
                        , b.SupplierNo, b.SupplierName
                        , (
                            SELECT x.GrDetailId,x.GrSeq, x.MtlItemNo, x.GrMtlItemName, x.GrMtlItemSpec
                            , x.PoErpPrefix, x.PoErpNo, x.PoSeq,x.LotNumber
                            , x.ReceiptQty, x.UomNo, x.InventoryNo
                            , FORMAT(x.AcceptanceDate, 'yyyy-MM-dd') AcceptanceDate, x.AcceptQty, x.AvailableQty, x.ReturnQty, x.Remark
                            , xa.InventoryName
							, (xb.PoErpPrefix + '-' + xb.PoErpNo + '-' + xb.PoSeq) PoErpFullNo, (xd.UserNo + '-' + xd.UserName) PoConfidmUser
							, (xf.PrErpPrefix + '-' + xf.PrErpNo + '-' + xe.PrSequence) PrErpFullNo, (xg.UserNo + '-' + xg.UserName) PrConfidmUser
                    FROM SCM.GrDetail x 
                            INNER JOIN SCM.Inventory xa ON x.InventoryId=  xa.InventoryId
                            INNER JOIN SCM.PoDetail xb ON (x.PoErpPrefix + '-' + x.PoErpNo + '-' + x.PoSeq) =  (xb.PoErpPrefix + '-' + xb.PoErpNo + '-' + xb.PoSeq)
                            INNER JOIN SCM.PurchaseOrder xc ON xc.PoId = xb.PoId
                            left JOIN BAS.[User] xd ON xc.PoUserId = xd.UserId
                            left JOIN SCM.PurchaseRequisition xf ON (xb.PrErpPrefix + '-' + xb.PrErpNo) =  (xf.PrErpPrefix + '-' + xf.PrErpNo)
							left JOIN SCM.PrDetail xe ON xe.PrId = xf.PrId AND xb.PrSequence = xe.PrSequence
                            left JOIN BAS.[User] xg ON xf.UserId = xg.UserId
                            WHERE x.GrId = a.GrId
                            FOR JSON PATH, ROOT('data')
                        ) GrDetail
                        , (
                            SELECT TOP 1 xa.UserNo
                            FROM SCM.GrDetail x 
                            INNER JOIN BAS.[User] xa ON x.ConfirmUser = xa.UserId
                            WHERE x.GrId = a.GrId
                        ) ConfirmUserNo
                        , (
                            SELECT TOP 1 xa.UserName
                            FROM SCM.GrDetail x 
                            INNER JOIN BAS.[User] xa ON x.ConfirmUser = xa.UserId
                            WHERE x.GrId = a.GrId
                        ) ConfirmUserName";
                    sqlQuery.mainTables =
                        @"FROM SCM.GoodsReceipt a 
                        INNER JOIN SCM.Supplier b ON a.SupplierId = b.SupplierId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "GrId", @" AND a.GrId = @GrId", GrId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SupplierId", @" AND a.SupplierId = @SupplierId", SupplierId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ConfirmUserId", @" AND EXISTS (
                                                                                                                SELECT TOP 1 1
                                                                                                                FROM SCM.GrDetail x 
                                                                                                                INNER JOIN BAS.[User] xa ON x.ConfirmUser = xa.UserNo
                                                                                                                WHERE x.GrId = a.GrId
                                                                                                                AND xa.UserId = @ConfirmUserId
                                                                                                       )", ConfirmUserId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey", @" AND EXISTS (
                                                                                                            SELECT TOP 1 1
                                                                                                            FROM SCM.GrDetail x
                                                                                                            LEFT JOIN PDM.MtlItem xa ON x.MtlItemId = xa.MtlItemId
                                                                                                            WHERE x.GrId = a.GrId
                                                                                                            AND (xa.MtlItemNo LIKE '%' + @SearchKey + '%' OR x.GrMtlItemName LIKE '%' + @SearchKey + '%')
                                                                                                       )", SearchKey);
                    if (ConfirmStatus.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ConfirmStatus", @" AND a.ConfirmStatus IN @ConfirmStatus", ConfirmStatus.Split(','));
                    if (ClosureStatus.Length > 0)
                    {
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ClosureStatus", @" AND EXISTS (
                                                                                                                    SELECT TOP 1 1
                                                                                                                    FROM SCM.GrDetail x
                                                                                                                    WHERE x.GrId = a.GrId
                                                                                                                    AND x.CloseStatus IN @ClosureStatus
                                                                                                               )", ClosureStatus.Split(','));
                    }
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "GrErpFullNo", @" AND (a.GrErpPrefix + '-' + a.GrErpNo) LIKE '%' + @GrErpFullNo + '%'", GrErpFullNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.DocDate >= @StartDate", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.DocDate <= @EndDate", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");

                    // 新增的搜尋條件
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PoErpFullNo", @" AND EXISTS (
                                                                                                            SELECT TOP 1 1
                                                                                                            FROM SCM.GrDetail x
                                                                                                            INNER JOIN SCM.PoDetail xb ON (x.PoErpPrefix + '-' + x.PoErpNo + '-' + x.PoSeq) = (xb.PoErpPrefix + '-' + xb.PoErpNo + '-' + xb.PoSeq)
                                                                                                            WHERE x.GrId = a.GrId
                                                                                                            AND (xb.PoErpPrefix + '-' + xb.PoErpNo) LIKE '%' + @PoErpFullNo + '%'
                                                                                                       )", PoErpFullNo);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PoUserId", @" AND EXISTS (
                                                                                                            SELECT TOP 1 1
                                                                                                            FROM SCM.GrDetail x
                                                                                                            INNER JOIN SCM.PoDetail xb ON (x.PoErpPrefix + '-' + x.PoErpNo + '-' + x.PoSeq) = (xb.PoErpPrefix + '-' + xb.PoErpNo + '-' + xb.PoSeq)
                                                                                                            INNER JOIN SCM.PurchaseOrder xc ON xc.PoId = xb.PoId
                                                                                                            WHERE x.GrId = a.GrId
                                                                                                            AND xc.PoUserId = @PoUserId
                                                                                                       )", PoUserId);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PrErpFullNo", @" AND EXISTS (
                                                                                                            SELECT TOP 1 1
                                                                                                            FROM SCM.GrDetail x
                                                                                                            INNER JOIN SCM.PoDetail xb ON (x.PoErpPrefix + '-' + x.PoErpNo + '-' + x.PoSeq) = (xb.PoErpPrefix + '-' + xb.PoErpNo + '-' + xb.PoSeq)
                                                                                                            INNER JOIN SCM.PurchaseRequisition xf ON (xb.PrErpPrefix + '-' + xb.PrErpNo) = (xf.PrErpPrefix + '-' + xf.PrErpNo)
                                                                                                            WHERE x.GrId = a.GrId
                                                                                                            AND (xf.PrErpPrefix + '-' + xf.PrErpNo) LIKE '%' + @PrErpFullNo + '%'
                                                                                                       )", PrErpFullNo);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PrUserId", @" AND EXISTS (
                                                                                                            SELECT TOP 1 1
                                                                                                            FROM SCM.GrDetail x
                                                                                                            INNER JOIN SCM.PoDetail xb ON (x.PoErpPrefix + '-' + x.PoErpNo + '-' + x.PoSeq) = (xb.PoErpPrefix + '-' + xb.PoErpNo + '-' + xb.PoSeq)
                                                                                                            INNER JOIN SCM.PurchaseRequisition xf ON (xb.PrErpPrefix + '-' + xb.PrErpNo) = (xf.PrErpPrefix + '-' + xf.PrErpNo)
                                                                                                            WHERE x.GrId = a.GrId
                                                                                                            AND xf.UserId = @PrUserId
                                                                                                       )", PrUserId);


                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.GrId DESC";
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

        #region //GetGrDetail -- 取得進貨單單身資料 -- Ann 2023-02-29
        public string GetGrDetail(int GrDetailId, int GrId, string GrErpFullNo, string SearchKey, int SupplierId, string ConfirmStatus, string CloseStatus, string GrErpFullNoWithSeq, string TransferStatus
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.GrDetailId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.GrId, a.GrErpPrefix, a.GrErpNo, a.GrSeq, a.MtlItemId, a.MtlItemNo, a.GrMtlItemName, a.GrMtlItemSpec
                        , a.ReceiptQty, a.UomId, a.UomNo, a.InventoryId, a.InventoryNo, a.LotNumber, a.PoErpPrefix, a.PoErpNo, a.PoSeq
                        , FORMAT(a.AcceptanceDate, 'yyyy-MM-dd') AcceptanceDate
                        , a.AcceptQty, a.AvailableQty, a.ReturnQty, a.OrigUnitPrice, a.OrigAmount, a.OrigDiscountAmt, a.TrErpPrefix, a.TrErpNo, a.TrSeq
                        , a.ReceiptExpense, a.DiscountDescription, a.PaymentHold, a.Overdue, a.QcStatus, a.ReturnCode, a.ConfirmStatus, a.CloseStatus
                        , a.ReNewStatus, a.Remark, a.InventoryQty, a.SmallUnit, a.AvailableDate, a.ReCheckDate, a.ConfirmUser, a.ApvErpPrefix, a.ApvErpNo
                        , a.ApvSeq, a.ProjectCode, a.ExpenseEntry, a.PremiumAmountFlag, a.OrigPreTaxAmt, a.OrigTaxAmt, a.PreTaxAmt, a.TaxAmt, a.ReceiptPackageQty
                        , a.AcceptancePackageQty, a.ReturnPackageQty, a.PremiumAmount, a.PackageUnit, a.ReserveTaxCode, a.OrigPremiumAmount, a.AvailableUomId
                        , a.AvailableUomNo, a.OrigCustomer, a.ApproveStatus, a.EbcErpPreNo, a.EbcEdition, a.ProductSeqAmount, a.MtlItemType, a.Loaction
                        , a.TaxRate, a.MultipleFlag, a.GrErpStatus, a.TaxCode, a.DiscountRate, a.DiscountPrice, a.QcType, a.TransferStatus, a.TransferDate
                        , a.CreateDate, a.LastModifiedDate, a.CreateBy, a.LastModifiedBy
                        , c.InventoryName
                        , d.PoDetailId, d.PoErpPrefix + '-' + d.PoErpNo PoErpFullNo, d.PoSeq, d.Quantity PoQuantity, d.SiQty, d.PoMtlItemName, d.PoMtlItemSpec
                        , e.MtlItemNo PoMtlItemNo, e.MeasureType, e.LotManagement
                        , f.StatusName QcStatusName
                        , g.SupplierNo, g.SupplierName
                        , h.MtlItemName, h.MtlItemSpec
                        , i.QcGoodsReceiptId, i.QcRecordId
                        , j.LogId";
                    sqlQuery.mainTables =
                        @"FROM SCM.GrDetail a 
                        INNER JOIN SCM.GoodsReceipt b ON a.GrId = b.GrId
                        INNER JOIN SCM.Inventory c ON a.InventoryId = c.InventoryId
                        LEFT JOIN SCM.PoDetail d ON a.PoErpPrefix = d.PoErpPrefix AND a.PoErpNo = d.PoErpNo AND a.PoSeq = d.PoSeq
                        LEFT JOIN PDM.MtlItem e ON d.MtlItemId = e.MtlItemId
                        INNER JOIN BAS.[Status] f ON a.QcStatus = f.StatusNo AND f.StatusSchema = 'GrDetail.QcStatus'
                        INNER JOIN SCM.Supplier g ON b.SupplierId = g.SupplierId
                        INNER JOIN PDM.MtlItem h ON a.MtlItemId = h.MtlItemId
                        LEFT JOIN MES.QcGoodsReceipt i ON a.GrDetailId = i.GrDetailId
                        LEFT JOIN MES.QcGoodsReceiptLog j ON i.QcGoodsReceiptId = j.QcGoodsReceiptId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND b.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "GrDetailId", @" AND a.GrDetailId = @GrDetailId", GrDetailId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "GrId", @" AND a.GrId = @GrId", GrId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SupplierId", @" AND b.SupplierId = @SupplierId", SupplierId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "GrErpFullNo", @" AND b.GrErpPrefix + '-' + b.GrErpNo LIKE '%' + @GrErpFullNo + '%'", GrErpFullNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey", @" AND (e.MtlItemNo LIKE '%' + @SearchKey + '%' OR a.GrMtlItemName LIKE '%' + @SearchKey + '%' OR a.GrMtlItemSpec LIKE '%' + @SearchKey + '%')", SearchKey);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ConfirmStatus", @" AND b.ConfirmStatus = @ConfirmStatus", ConfirmStatus);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CloseStatus", @" AND a.CloseStatus = @CloseStatus", CloseStatus);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TransferStatus", @" AND b.TransferStatus = @TransferStatus", TransferStatus);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "GrErpFullNoWithSeq", @" AND b.GrErpPrefix + '-' + b.GrErpNo + '(' + a.GrSeq + ')' LIKE '%' + @GrErpFullNoWithSeq + '%'", GrErpFullNoWithSeq);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.GrDetailId";
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

        #region //GetTotal -- 取得進貨單統整資料 -- Ann 2024-03-06
        public string GetTotal(int GrId)
        {
            try
            {
                if (GrId <= 0) throw new SystemException("【進貨單ID】格式錯誤!");
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.Quantity, a.TotalAmount, a.DeductAmount, a.OrigPreTaxAmount, a.OrigTax, a.ReceiptAmount
                            , a.OrigPreTaxAmount + a.OrigTax OriTotal
                            , a.PreTaxAmount, a.TaxAmount
                            , a.PreTaxAmount + a.TaxAmount Total
                            FROM SCM.GoodsReceipt a 
                            WHERE a.GrId = @GrId";
                    dynamicParameters.Add("GrId", GrId);

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

        #region //GetDocumentVerification -- 取得單據性質驗證資料 -- Ann 2024-07-16
        public string GetDocumentVerification(int GrId)
        {
            try
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
                        #region //取得MES進貨單資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.GrErpPrefix
                                FROM SCM.GoodsReceipt a 
                                WHERE a.GrId = @GrId";
                        dynamicParameters.Add("GrId", GrId);

                        var GoodsReceiptResult = sqlConnection.Query(sql, dynamicParameters);

                        if (GoodsReceiptResult.Count() <= 0) throw new SystemException("進貨單據資料錯誤!!");

                        string GrErpPrefix = "";
                        foreach (var item in GoodsReceiptResult)
                        {
                            GrErpPrefix = item.GrErpPrefix;
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT LTRIM(RTRIM(MQ019)) MQ019
                                FROM CMSMQ
                                WHERE MQ001 = @MQ001";
                        dynamicParameters.Add("MQ001", GrErpPrefix);

                        var result = sqlConnection2.Query(sql, dynamicParameters);

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

        #region //GetLnBarcode 取得批號綁定條碼 GPAI 20240412
        public string GetLnBarcode(string LotNumberNo, string LnBarcodeNo, int MtlItemId, int LotNumberId, int GrDetailId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.GrBarcodeId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.BarcodeNo, a.GrDetailId, b.BarcodeQty, b.BarcodeStatus, c.MtlItemId, d.LotNumberNo, d.LotNumberId
                        ,  FORMAT(a.CreateDate, 'yyyy-MM-dd') CreateDate, e.UserNo, e.UserName";
                    sqlQuery.mainTables =
                        @"from SCM.GrBarcode a
                          INNER JOIN MES.Barcode b ON a.BarcodeNo = b.BarcodeNo
                          INNER JOIN SCM.GrDetail c ON a.GrDetailId = c.GrDetailId
                          INNER JOIN SCM.LotNumber d ON d.LotNumberNo = c.LotNumber AND d.MtlItemId = c.MtlItemId
                          INNER JOIN BAS.[User] e ON a.CreateBy = e.UserId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"";
                    //dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "LotNumberId", @" AND d.LotNumberId = @LotNumberId", LotNumberId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "LotNumberNo", @" AND d.LotNumberNo = @LotNumberNo", LotNumberNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "LnBarcodeNo", @" AND a.BarcodeNo = @LnBarcodeNo", LnBarcodeNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemId", @" AND c.MtlItemId = @MtlItemId", MtlItemId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "GrDetailId", @" AND c.GrDetailId = @GrDetailId", GrDetailId);


                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.GrBarcodeId";
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

        #region //GetFileInfo -- 取得檔案資料 -- GPAI 2024-05-22
        public string GetFileInfo(string FileIdList, string FileType)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.FileId, a.[FileName], a.FileExtension
                            , a.[FileName] + a.FileExtension FileFullName, a.FileSize, a.FileContent
                            FROM BAS.[File] a
                            WHERE a.FileId IN (" + FileIdList + ")";

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
        #region //AddGoodsReceipt -- 新增進貨單資料 -- Ann 2024-02-27
        public string AddGoodsReceipt(string GrErpPrefix, string DocDate, string ReceiptDate, int SupplierId, string Remark
            , string CurrencyCode, double Exchange, string PaymentTerm, int RowCnt
            , string TaxCode, string TaxType, string InvoiceType, double TaxRate, string UiNo, string InvoiceDate
            , string InvoiceNo, string ApplyYYMM, string DeductType, string ContactUser)
        {
            try
            {
                if (GrErpPrefix.Length <= 0) throw new SystemException("【進貨單別】不能為空!");
                if (!DateTime.TryParse(DocDate, out DateTime tempDate)) throw new SystemException("【單據日期】格式錯誤!");
                if (!DateTime.TryParse(ReceiptDate, out DateTime tempDate2)) throw new SystemException("【進貨日期】格式錯誤!");
                if (Exchange < 0) throw new SystemException("【匯率】格式錯誤!");
                if (PaymentTerm.Length <= 0) throw new SystemException("【交易方式】格式錯誤!");
                if (RowCnt < 0) throw new SystemException("【件數】格式錯誤!");
                if (TaxRate < 0) throw new SystemException("【營業稅率】格式錯誤!");
                if (ApplyYYMM.Length <= 0) throw new SystemException("【申報日期】格式錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    string companyNo = "";

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

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //比對ERP關帳日期(庫存關帳)
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MA012)) MA012
                                    FROM CMSMA";
                            var cmsmaResult = sqlConnection2.Query(sql, dynamicParameters);
                            if (cmsmaResult.Count() < 0) throw new SystemException("EPR關帳日期資料錯誤!!");

                            foreach (var item in cmsmaResult)
                            {
                                string eprDate = item.MA012;
                                string erpYear = eprDate.Substring(0, 4);
                                string erpMonth = eprDate.Substring(4, 2);
                                string erpFullDate = erpYear + "-" + erpMonth;
                                DateTime erpDateTime = Convert.ToDateTime(erpFullDate);
                                DateTime DocDateDateTime = Convert.ToDateTime(DocDate.Split('-')[0] + "-" + DocDate.Split('-')[1]);
                                int compare = DocDateDateTime.CompareTo(erpDateTime);
                                if (compare <= 0) throw new SystemException("ERP已關帳(" + erpFullDate + ")，無法開立此日期(" + DocDate + ")之單據!!");
                            }
                            #endregion

                            #region //判斷供應商資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 *
                                    FROM SCM.Supplier a 
                                    WHERE a.SupplierId = @SupplierId";
                            dynamicParameters.Add("SupplierId", SupplierId);

                            List<Supplier> resultSuppliers = sqlConnection.Query<Supplier>(sql, dynamicParameters).ToList();

                            if (resultSuppliers.Count() <= 0) throw new SystemException("【供應商】資料錯誤!");
                            #endregion

                            #region //確認幣別資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM CMSMF
                                    WHERE MF001 = @MF001";
                            dynamicParameters.Add("MF001", CurrencyCode);

                            var CMSMFResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (CMSMFResult.Count() <= 0) throw new SystemException("【幣別】資料錯誤!!");
                            #endregion

                            #region //確認付款條件資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(NA002)) PaymentTermNo, LTRIM(RTRIM(NA003)) PaymentTermName 
                                    FROM CMSNA
                                    WHERE NA002 = @NA002";
                            dynamicParameters.Add("NA002", PaymentTerm);

                            var CMSNAResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (CMSNAResult.Count() <= 0) throw new SystemException("【付款條件】資料錯誤!!");
                            #endregion

                            #region //確認稅別碼資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(NN001)) TaxNo, LTRIM(RTRIM(NN002)) TaxName, ROUND(LTRIM(RTRIM(NN004)),2) BusinessTaxRate, LTRIM(RTRIM(NN006)) Taxation
                                    FROM CMSNN 
                                    WHERE NN001 = @NN001";
                            dynamicParameters.Add("NN001", TaxCode);

                            var CMSNNResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (CMSNNResult.Count() <= 0) throw new SystemException("【課稅別】資料錯誤!!");
                            #endregion

                            #region //確認課稅別資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.[Type]
                                    WHERE TypeSchema = 'Customer.Taxation'
                                    AND TypeNo = @TaxType";
                            dynamicParameters.Add("TaxType", TaxType);

                            var TaxationResult = sqlConnection.Query(sql, dynamicParameters);

                            if (TaxationResult.Count() <= 0) throw new SystemException("【課稅別】資料錯誤!!");
                            #endregion

                            #region //確認發票聯數資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(NM001)) InvoiceCountNo, LTRIM(RTRIM(NM002)) InvoiceCountName 
                                    FROM CMSNM a
                                    INNER JOIN CMSNN b ON a.NM001 = b.NN005
                                    WHERE NM001 = @NM001";
                            dynamicParameters.Add("NM001", InvoiceType);

                            var CMSNMResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (CMSNMResult.Count() <= 0) throw new SystemException("【發票聯數】資料錯誤!!");
                            #endregion

                            #region //確認扣抵區分資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.[Type]
                                    WHERE TypeSchema = 'OspReceipt.DeductType'
                                    AND TypeNo = @DeductType";
                            dynamicParameters.Add("DeductType", DeductType);

                            var DeductTypeResult = sqlConnection.Query(sql, dynamicParameters);

                            if (DeductTypeResult.Count() <= 0) throw new SystemException("【扣抵區分】資料錯誤!!");
                            #endregion

                            #region //隨機取得單號資料
                            string GrErpNo = "";

                            bool checkSoErpNo = true;
                            while (checkSoErpNo)
                            {
                                GrErpNo = BaseHelper.RandomCode(11);

                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM SCM.GoodsReceipt
                                        WHERE GrErpPrefix = @GrErpPrefix
                                        AND GrErpNo = @GrErpNo";
                                dynamicParameters.Add("GrErpPrefix", GrErpPrefix);
                                dynamicParameters.Add("GrErpNo", GrErpNo);

                                var resultGrErpNo = sqlConnection.Query(sql, dynamicParameters);
                                checkSoErpNo = resultGrErpNo.Count() > 0;
                            }
                            #endregion

                            #region //處理申報日期
                            string[] applyYYMMArray = ApplyYYMM.Split('-');
                            string year = applyYYMMArray[0];
                            string month = applyYYMMArray[1];
                            ApplyYYMM = year + month;
                            #endregion

                            #region //INSERT SCM.GoodsReceipt
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO SCM.GoodsReceipt (CompanyId, GrErpPrefix, GrErpNo, ReceiptDate, SupplierId, SupplierNo
                                    , SupplierSo, CurrencyCode, Exchange, InvoiceType, TaxType, InvoiceNo, ConfirmStatus, DocDate
                                    , Remark, SupplierName, UiNo, DeductType, InvoiceDate, ApplyYYMM, TaxRate, PaymentTerm, FromDocType, TaxCode, TradeTerm, ContactUser
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.GrId, INSERTED.GrErpNo
                                    VALUES (@CompanyId, @GrErpPrefix, @GrErpNo, @ReceiptDate, @SupplierId, @SupplierNo
                                    , @SupplierSo, @CurrencyCode, @Exchange, @InvoiceType, @TaxType, @InvoiceNo, @ConfirmStatus, @DocDate
                                    , @Remark, @SupplierName, @UiNo, @DeductType, @InvoiceDate, @ApplyYYMM, @TaxRate, @PaymentTerm, @FromDocType, @TaxCode, @TradeTerm, @ContactUser
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    CompanyId = CurrentCompany,
                                    GrErpPrefix,
                                    GrErpNo,
                                    ReceiptDate,
                                    SupplierId,
                                    SupplierNo = resultSuppliers.Select(x => x.SupplierNo).FirstOrDefault(),
                                    SupplierSo = "",
                                    CurrencyCode,
                                    Exchange,
                                    InvoiceType,
                                    TaxType,
                                    InvoiceNo,
                                    ConfirmStatus = "N",
                                    DocDate,
                                    Remark,
                                    SupplierName = resultSuppliers.Select(x => x.SupplierName).FirstOrDefault(),
                                    UiNo,
                                    DeductType,
                                    InvoiceDate,
                                    ApplyYYMM,
                                    TaxRate,
                                    PaymentTerm,
                                    FromDocType = "1",
                                    TaxCode,
                                    TradeTerm = resultSuppliers.Select(x => x.TradeTerm).FirstOrDefault(),
                                    ContactUser,
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

        #region //ImportPurchaseOrder -- 複製前置單據作業(來源:採購單) -- Ann 2024-03-05
        public string ImportPurchaseOrder(int GrId, int PoId, string CopyExchange)
        {
            try
            {
                if (CopyExchange.Length <= 0) throw new SystemException("【是否複製前置單據匯率】不能為空!!");

                string companyNo = "";
                double rowsAffected = 0;

                List<SpreadsheetModel> spreadsheetModels = new List<SpreadsheetModel>();
                Dictionary<string, string> BarcodeInfo = new Dictionary<string, string>();

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
                            companyNo = item.ErpNo;
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //確認進貨單資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.GrId, a.CompanyId, a.GrErpPrefix, a.GrErpNo, GrErpPrefix + '-' + GrErpNo GrErpFullNo, a.ReceiptDate, a.SupplierId, a.SupplierNo
                                    , a.SupplierSo, a.CurrencyCode, a.Exchange, a.InvoiceType, a.TaxType, a.InvoiceNo, a.PrintCnt, a.ConfirmStatus, a.DocDate
                                    , a.RenewFlag, a.Remark, a.TotalAmount, a.DeductAmount, a.OrigTax, a.ReceiptAmount, a.SupplierName, a.UiNo, a.DeductType
                                    , a.ObaccoAndLiquorFlag, a.RowCnt, a.Quantity, FORMAT(a.InvoiceDate, 'yyyy-MM-dd') InvoiceDate, a.OrigPreTaxAmount, a.ApplyYYMM
                                    , a.TaxRate, a.PreTaxAmount, a.TaxAmount, a.PaymentTerm, a.PoId, a.PrepaidErpPrefix, a.PrepaidErpNo, a.OffsetAmount, a.TaxOffset, a.PackageQuantity
                                    , a.PremiumAmount, a.ApproveStatus, a.InvoiceFlag, a.ReserveTaxCode, a.SendCount, a.OrigPremiumAmount, a.EbcErpPreNo, a.EbcEdition, a.EbcFlag
                                    , a.FromDocType, a.NoticeFlag, a.TaxCode, a.TradeTerm, a.DetailMultiTax, a.TaxExchange, a.ContactUser, a.TransferStatus, a.TransferDate
                                    , a.CreateDate, a.LastModifiedDate, a.CreateBy, a.LastModifiedBy
                                    FROM SCM.GoodsReceipt a 
                                    WHERE a.GrId = @GrId";
                            dynamicParameters.Add("GrId", GrId);

                            List<GoodsReceipt> goodsReceipts = sqlConnection.Query<GoodsReceipt>(sql, dynamicParameters).ToList();

                            if (goodsReceipts.Count() <= 0) throw new SystemException("進貨單資料錯誤!!");
                            #endregion

                            #region //確認此進貨單身是否有其他筆資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM SCM.GrDetail a 
                                    WHERE a.GrId = @GrId";
                            dynamicParameters.Add("GrId", GrId);

                            var GrDetailResult = sqlConnection.Query(sql, dynamicParameters);

                            if (GrDetailResult.Count() > 0) throw new SystemException("此進貨單已有單身資料，無法使用複製前置單據功能!!");
                            #endregion

                            #region //判斷採購單資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT CurrencyCode, Exchange, PaymentTermNo
                                    , Taxation, TaxRate, PaymentTermNo, TaxNo, TradeTerm
                                    FROM SCM.PurchaseOrder
                                    WHERE PoId = @PoId";
                            dynamicParameters.Add("PoId", PoId);

                            var PurchaseOrderResult = sqlConnection.Query(sql, dynamicParameters);

                            if (PurchaseOrderResult.Count() <= 0) throw new SystemException("採購單資料錯誤!!");

                            double PreExchange = 0;
                            string PaymentTermNo = "";
                            string Taxation = "";
                            double TaxRate = 0;
                            string TaxNo = "";
                            string TradeTerm = "";
                            foreach (var item in PurchaseOrderResult)
                            {
                                goodsReceipts[0].CurrencyCode = item.CurrencyCode;
                                PreExchange = item.Exchange;
                                PaymentTermNo = item.PaymentTermNo;
                                Taxation = item.Taxation;
                                TaxRate = item.TaxRate;
                                TaxNo = item.TaxNo;
                                TradeTerm = item.TradeTerm;
                            }
                            #endregion

                            #region //確認來源單據單身資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.PoDetailId, a.PoId, a.PoErpPrefix, a.PoErpNo, a.PoSeq, a.MtlItemId, a.PoMtlItemName, a.PoMtlItemSpec, a.InventoryId, a.Quantity
                                    , a.UomId, a.PoUnitPrice, a.PoPrice, a.PromiseDate, a.ReferencePrefix, a.Remark, a.SiQty, a.ClosureStatus, a.ConfirmStatus 
                                    , a.InventoryQty, a.SmallUnit, a.ReferenceNo, a.Project, a.ReferenceSeq, a.UrgentMtl, a.PrErpPrefix, a.PrErpNo, a.PrSequence 
                                    , a.FromDocType, a.MtlItemType, a.PartialSeq, a.OriPromiseDate, a.DeliveryDate, a.PoPriceQty, a.PoPriceUomId, a.SiPriceQty 
                                    , a.DiscountRate, a.DiscountAmount, a.CreateDate, a.LastModifiedDate, a.CreateBy, a.LastModifiedBy
                                    , b.UomNo
                                    , c.InventoryNo
                                    , d.MtlItemNo
                                    , e.TaxNo
                                    FROM SCM.PoDetail a
                                    INNER JOIN PDM.UnitOfMeasure b ON a.UomId = b.UomId
                                    INNER JOIN SCM.Inventory c ON a.InventoryId = c.InventoryId
                                    INNER JOIN PDM.MtlItem d ON a.MtlItemId = d.MtlItemId
                                    INNER JOIN SCM.PurchaseOrder e ON a.PoId = e.PoId
                                    WHERE a.PoId = @PoId";
                            dynamicParameters.Add("PoId", PoId);

                            List<PoDetail> poDetails = sqlConnection.Query<PoDetail>(sql, dynamicParameters).ToList();

                            if (poDetails.Count() <= 0) throw new SystemException("此採購單無任何單身資料，無法複製!!");
                            #endregion

                            #region //確認ERP採購單狀態是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(TC014)) TC014
                                    FROM PURTC 
                                    WHERE TC001 = @PoErpPrefix
                                    AND TC002 = @PoErpNo";
                            dynamicParameters.Add("PoErpPrefix", poDetails[0].PoErpPrefix);
                            dynamicParameters.Add("PoErpNo", poDetails[0].PoErpNo);

                            var PURTCResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (PURTCResult.Count() <= 0) throw new SystemException("ERP採購單資料錯誤!!");

                            foreach (var item in PURTCResult)
                            {
                                if (item.TC014 != "Y") throw new SystemException("ERP採購單狀態錯誤!!");
                            }
                            #endregion

                            #region //原幣小數點取位
                            int OriUnitDecimal = 0;
                            int OriPriceDecimal = 0;
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.MF003, a.MF004
                                    FROM CMSMF a 
                                    WHERE a.MF001 = @Currency";
                            dynamicParameters.Add("Currency", goodsReceipts[0].CurrencyCode);

                            var CMSMFResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (CMSMFResult.Count() <= 0) throw new SystemException("金額小數點取位資料錯誤!!");

                            foreach (var item in CMSMFResult)
                            {
                                OriUnitDecimal = Convert.ToInt32(item.MF003);
                                OriPriceDecimal = Convert.ToInt32(item.MF004);
                            }
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

                            #region //本幣小數點取位
                            int UnitDecimal = 0;
                            int PriceDecimal = 0;
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.MF003, a.MF004
                                    FROM CMSMF a 
                                    WHERE a.MF001 = @Currency";
                            dynamicParameters.Add("Currency", CurrencyCode);

                            var CMSMFResult2 = sqlConnection2.Query(sql, dynamicParameters);

                            if (CMSMFResult2.Count() <= 0) throw new SystemException("金額小數點取位資料錯誤!!");

                            foreach (var item in CMSMFResult2)
                            {
                                UnitDecimal = Convert.ToInt32(item.MF003);
                                PriceDecimal = Convert.ToInt32(item.MF004);
                            }
                            #endregion

                            #region //處理匯率
                            if (CopyExchange == "Y")
                            {
                                goodsReceipts[0].Exchange = PreExchange;
                            }
                            #endregion

                            #region //新增單身相關資料
                            int count = 1;
                            foreach (var item in poDetails)
                            {
                                if (item.PoPriceQty - item.SiPriceQty == 0)
                                {
                                    continue;
                                }

                                #region //確認ERP採購單身結案碼
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT LTRIM(RTRIM(TD016)) TD016
                                        , LTRIM(RTRIM(TD018)) TD018
                                        FROM PURTD
                                        WHERE TD001 = @PoErpPrefix
                                        AND TD002 = @PoErpNo
                                        AND TD003 = @PoSeq";
                                dynamicParameters.Add("PoErpPrefix", item.PoErpPrefix);
                                dynamicParameters.Add("PoErpNo", item.PoErpNo);
                                dynamicParameters.Add("PoSeq", item.PoSeq);

                                var PURTDResult = sqlConnection2.Query(sql, dynamicParameters);

                                if (PURTDResult.Count() <= 0) throw new SystemException("ERP採購單單身資料錯誤!!");

                                foreach (var item2 in PURTDResult)
                                {
                                    if (item2.TD016 != "N") throw new SystemException("ERP採購單單身結案碼狀態錯誤!!");
                                    if (item2.TD018 != "Y") throw new SystemException("ERP採購單單身狀態錯誤!!");
                                }
                                #endregion

                                #region //處理進貨單身序號
                                string GrSeq = count.ToString("D4");
                                #endregion

                                #region //確認品號檢驗狀態
                                string QcStatus = "";
                                string QcType = "";
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT MB043
                                        FROM INVMB
                                        WHERE MB001 = @MB001";
                                dynamicParameters.Add("MB001", item.MtlItemNo);

                                var INVMBResult = sqlConnection2.Query(sql, dynamicParameters);

                                if (INVMBResult.Count() <= 0) throw new SystemException("此進貨品號【" + item.MtlItemNo + "】資料錯誤!!");

                                foreach (var item2 in INVMBResult)
                                {
                                    if (item2.MB043 == "0")
                                    {
                                        QcStatus = "0";
                                        QcType = "N";
                                    }
                                    else
                                    {
                                        QcStatus = "1";
                                        QcType = "1";
                                    }
                                }
                                #endregion

                                #region //查詢目前庫存數
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT ISNULL(MC007, 0) MC007
                                        FROM INVMC
                                        WHERE MC001 = @MC001
                                        AND MC002 = @MC002";
                                dynamicParameters.Add("MC001", item.MtlItemNo);
                                dynamicParameters.Add("MC002", item.InventoryNo);

                                var INVMCResult = sqlConnection2.Query(sql, dynamicParameters);

                                double InventoryQty = 0;
                                foreach (var item2 in INVMCResult)
                                {
                                    InventoryQty = Convert.ToDouble(item2.MC007);
                                }
                                #endregion

                                #region //處理【原幣進貨金額】、【原幣未稅金額】、【原幣稅金額】、【本幣未稅金額】、【本幣稅金額】
                                double? AvailableQty = QcStatus == "0" ? item.PoPriceQty - item.SiPriceQty : 0; //計價數量
                                double? OrigUnitPrice = Math.Round(Convert.ToDouble(item.PoUnitPrice), OriUnitDecimal); //原幣進貨單價
                                double? OrigDiscountAmt = 0; //原幣扣款金額
                                double? OrigAmount = Math.Round(Convert.ToDouble(AvailableQty * OrigUnitPrice), OriPriceDecimal); //原幣進貨金額

                                double OrigPreTaxAmt = 0; //原幣未稅金額
                                double OrigTaxAmt = 0; //原幣稅金額
                                double PreTaxAmt = 0; //本幣未稅金額
                                double TaxAmt = 0; //本幣稅金額

                                JObject calculateTaxAmtResult = JObject.Parse(CalculateTaxAmt(AvailableQty, OrigUnitPrice, OrigDiscountAmt, OrigAmount, goodsReceipts[0].TaxRate, goodsReceipts[0].TaxType, goodsReceipts[0].Exchange
                                                                                , OriUnitDecimal, OriPriceDecimal, UnitDecimal, PriceDecimal));

                                if (calculateTaxAmtResult["status"].ToString() == "success")
                                {
                                    OrigPreTaxAmt = Convert.ToDouble(calculateTaxAmtResult["OrigPreTaxAmt"]);
                                    OrigTaxAmt = Convert.ToDouble(calculateTaxAmtResult["OrigTaxAmt"]);
                                    PreTaxAmt = Convert.ToDouble(calculateTaxAmtResult["PreTaxAmt"]);
                                    TaxAmt = Convert.ToDouble(calculateTaxAmtResult["TaxAmt"]);
                                }
                                else
                                {
                                    throw new SystemException(calculateTaxAmtResult["msg"].ToString());
                                }
                                #endregion

                                #region //INSERT SCM.GrDetail
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO SCM.GrDetail (GrId, GrErpPrefix, GrErpNo, GrSeq, MtlItemId, MtlItemNo, GrMtlItemName, GrMtlItemSpec
                                        , ReceiptQty, UomId, UomNo, InventoryId, InventoryNo, LotNumber, PoErpPrefix, PoErpNo, PoSeq, AcceptanceDate
                                        , AcceptQty, AvailableQty, ReturnQty, OrigUnitPrice, OrigAmount, OrigDiscountAmt, TrErpPrefix, TrErpNo, TrSeq
                                        , ReceiptExpense, DiscountDescription, PaymentHold, Overdue, QcStatus, ReturnCode, ConfirmStatus, CloseStatus
                                        , ReNewStatus, Remark, InventoryQty, SmallUnit, AvailableDate, ReCheckDate, ConfirmUser, ApvErpPrefix, ApvErpNo
                                        , ApvSeq, ProjectCode, ExpenseEntry, PremiumAmountFlag, OrigPreTaxAmt, OrigTaxAmt, PreTaxAmt, TaxAmt, ReceiptPackageQty
                                        , AcceptancePackageQty, ReturnPackageQty, PremiumAmount, PackageUnit, ReserveTaxCode, OrigPremiumAmount, AvailableUomId
                                        , AvailableUomNo, OrigCustomer, ApproveStatus, EbcErpPreNo, EbcEdition, ProductSeqAmount, MtlItemType, Loaction
                                        , TaxRate, MultipleFlag, GrErpStatus, TaxCode, DiscountRate, DiscountPrice, QcType, TransferStatus, TransferDate
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.GrDetailId
                                        VALUES (@GrId, @GrErpPrefix, @GrErpNo, @GrSeq, @MtlItemId, @MtlItemNo, @GrMtlItemName, @GrMtlItemSpec
                                        , @ReceiptQty, @UomId, @UomNo, @InventoryId, @InventoryNo, @LotNumber, @PoErpPrefix, @PoErpNo, @PoSeq, @AcceptanceDate
                                        , @AcceptQty, @AvailableQty, @ReturnQty, @OrigUnitPrice, @OrigAmount, @OrigDiscountAmt, @TrErpPrefix, @TrErpNo, @TrSeq
                                        , @ReceiptExpense, @DiscountDescription, @PaymentHold, @Overdue, @QcStatus, @ReturnCode, @ConfirmStatus, @CloseStatus
                                        , @ReNewStatus, @Remark, @InventoryQty, @SmallUnit, @AvailableDate, @ReCheckDate, @ConfirmUser, @ApvErpPrefix, @ApvErpNo
                                        , @ApvSeq, @ProjectCode, @ExpenseEntry, @PremiumAmountFlag, @OrigPreTaxAmt, @OrigTaxAmt, @PreTaxAmt, @TaxAmt, @ReceiptPackageQty
                                        , @AcceptancePackageQty, @ReturnPackageQty, @PremiumAmount, @PackageUnit, @ReserveTaxCode, @OrigPremiumAmount, @AvailableUomId
                                        , @AvailableUomNo, @OrigCustomer, @ApproveStatus, @EbcErpPreNo, @EbcEdition, @ProductSeqAmount, @MtlItemType, @Loaction
                                        , @TaxRate, @MultipleFlag, @GrErpStatus, @TaxCode, @DiscountRate, @DiscountPrice, @QcType, @TransferStatus, @TransferDate
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        GrId,
                                        goodsReceipts[0].GrErpPrefix,
                                        goodsReceipts[0].GrErpNo,
                                        GrSeq,
                                        item.MtlItemId,
                                        item.MtlItemNo,
                                        GrMtlItemName = item.PoMtlItemName,
                                        GrMtlItemSpec = item.PoMtlItemSpec,
                                        ReceiptQty = item.Quantity - item.SiQty,
                                        item.UomId,
                                        item.UomNo,
                                        item.InventoryId,
                                        item.InventoryNo,
                                        LotNumber = "",
                                        item.PoErpPrefix,
                                        item.PoErpNo,
                                        item.PoSeq,
                                        AcceptanceDate = DateTime.Now.ToString("yyyyMMdd"),
                                        AcceptQty = QcStatus != "0" ? 0 :item.Quantity - item.SiQty,
                                        AvailableQty,
                                        ReturnQty = 0,
                                        OrigUnitPrice,
                                        OrigAmount,
                                        OrigDiscountAmt,
                                        TrErpPrefix = "",
                                        TrErpNo = "",
                                        TrSeq = "",
                                        ReceiptExpense = 0,
                                        DiscountDescription = "",
                                        PaymentHold = "N",
                                        Overdue = "N",
                                        QcStatus,
                                        ReturnCode = "N",
                                        ConfirmStatus = "N",
                                        CloseStatus = "N",
                                        ReNewStatus = "N",
                                        item.Remark,
                                        InventoryQty,
                                        SmallUnit = "",
                                        AvailableDate = "",
                                        ReCheckDate = "",
                                        ConfirmUser = (int?)null,
                                        ApvErpPrefix = "",
                                        ApvErpNo = "",
                                        ApvSeq = "",
                                        ProjectCode = item.Project,
                                        ExpenseEntry = "N",
                                        PremiumAmountFlag = "N",
                                        OrigPreTaxAmt,
                                        OrigTaxAmt,
                                        PreTaxAmt,
                                        TaxAmt,
                                        ReceiptPackageQty = 0,
                                        AcceptancePackageQty = 0,
                                        ReturnPackageQty = 0,
                                        PremiumAmount = 0,
                                        PackageUnit = "",
                                        ReserveTaxCode = "N",
                                        OrigPremiumAmount = "",
                                        AvailableUomId = item.UomId,
                                        AvailableUomNo = item.UomNo,
                                        OrigCustomer = "",
                                        ApproveStatus = "N",
                                        EbcErpPreNo = "",
                                        EbcEdition = "",
                                        ProductSeqAmount = 0,
                                        MtlItemType = "2",
                                        Loaction = "",
                                        goodsReceipts[0].TaxRate,
                                        MultipleFlag = "N",
                                        GrErpStatus = "0",
                                        goodsReceipts[0].TaxCode,
                                        item.DiscountRate,
                                        DiscountPrice = item.DiscountAmount,
                                        QcType,
                                        TransferStatus = "N",
                                        TransferDate = (DateTime?)null,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();

                                int GrDetailId = -1;
                                foreach (var item2 in insertResult)
                                {
                                    GrDetailId = item2.GrDetailId;
                                }
                                #endregion

                                #region //若品項為待驗品，自動新增進貨檢單據
                                if (QcStatus == "1")
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

                                    #region //取得進貨檢QcType
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 QcTypeId
                                            FROM QMS.QcType
                                            WHERE QcTypeNo = 'IQC'";

                                    var QcTypeResult = sqlConnection.Query(sql, dynamicParameters);

                                    if (QcTypeResult.Count() <= 0) throw new SystemException("量測類型資料錯誤!!");

                                    int QcTypeId = -1;
                                    foreach (var item2 in QcTypeResult)
                                    {
                                        QcTypeId = item2.QcTypeId;
                                    }
                                    #endregion

                                    #region //INSERT MES.QcRecord
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO MES.QcRecord (QcTypeId, InputType, Remark, DefaultFileId, DefaultSpreadsheetData, ResolveFileJson
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.QcRecordId
                                            VALUES (@QcTypeId, @InputType, @Remark, @DefaultFileId, @DefaultSpreadsheetData, @ResolveFileJson
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            QcTypeId,
                                            InputType = "LotNumber",
                                            Remark = "由進貨單自動建立",
                                            DefaultFileId = -1,
                                            DefaultSpreadsheetData,
                                            ResolveFileJson = "",
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    var insertQcRecordResult = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertQcRecordResult.Count();

                                    int QcRecordId = -1;
                                    foreach (var item2 in insertQcRecordResult)
                                    {
                                        QcRecordId = item2.QcRecordId;
                                    }
                                    #endregion

                                    #region //INSERT MES.QcGoodsReceipt
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO MES.QcGoodsReceipt (QcRecordId, GrDetailId
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.QcGoodsReceiptId
                                            VALUES (@QcRecordId, @GrDetailId
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            QcRecordId,
                                            GrDetailId,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //建立量測單據Spreadsheet DATA
                                    #region //取得品號允收標準
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.MtlQcItemId, a.MtlItemId, a.QcItemId, a.DesignValue, a.UpperTolerance, a.LowerTolerance
                                            , a.QcItemDesc, a.SortNumber, a.QmmDetailId
                                            , b.QcItemNo, b.QcItemName, b.QcItemDesc OriQcItemDesc
                                            , c.MachineNumber
                                            , d.MachineDesc
                                            FROM PDM.MtlQcItem a
                                            INNER JOIN QMS.QcItem b ON a.QcItemId = b.QcItemId
                                            LEFT JOIN QMS.QmmDetail c ON a.QmmDetailId = c.QmmDetailId
                                            LEFT JOIN MES.Machine d ON c.MachineId = d.MachineId
                                            WHERE a.MtlItemId = @MtlItemId
                                            AND a.[Status] = 'A' 
                                            ORDER BY a.SortNumber";
                                    dynamicParameters.Add("MtlItemId", item.MtlItemId);

                                    var MtlQcItemResult = sqlConnection.Query(sql, dynamicParameters);
                                    #endregion

                                    #region //初始化Data
                                    List<Data> datas = new List<Data>();
                                    Data data = new Data();
                                    data = new Data()
                                    {
                                        cell = "A1",
                                        css = "imported_class1",
                                        format = "common",
                                        value = "序號",
                                    };
                                    datas.Add(data);

                                    data = new Data()
                                    {
                                        cell = "B1",
                                        css = "imported_class2",
                                        format = "common",
                                        value = "檢測項目",
                                    };
                                    datas.Add(data);

                                    data = new Data()
                                    {
                                        cell = "C1",
                                        css = "imported_class3",
                                        format = "common",
                                        value = "檢測備註",
                                    };
                                    datas.Add(data);

                                    data = new Data()
                                    {
                                        cell = "D1",
                                        css = "imported_class4",
                                        format = "common",
                                        value = "量測設備",
                                    };
                                    datas.Add(data);

                                    data = new Data()
                                    {
                                        cell = "E1",
                                        css = "imported_class5",
                                        format = "common",
                                        value = "設計值",
                                    };
                                    datas.Add(data);

                                    data = new Data()
                                    {
                                        cell = "F1",
                                        css = "imported_class6",
                                        format = "common",
                                        value = "上限",
                                    };
                                    datas.Add(data);

                                    data = new Data()
                                    {
                                        cell = "G1",
                                        css = "imported_class7",
                                        format = "common",
                                        value = "下限",
                                    };
                                    datas.Add(data);

                                    data = new Data()
                                    {
                                        cell = "H1",
                                        css = "imported_class8",
                                        format = "common",
                                        value = "Z軸",
                                    };
                                    datas.Add(data);

                                    data = new Data()
                                    {
                                        cell = "I1",
                                        css = "imported_class9",
                                        format = "common",
                                        value = "量測人員",
                                    };
                                    datas.Add(data);

                                    data = new Data()
                                    {
                                        cell = "J1",
                                        css = "imported_class10",
                                        format = "common",
                                        value = "量測工時",
                                    };
                                    datas.Add(data);
                                    #endregion

                                    #region //設定單身量測標準
                                    int row = 2;
                                    foreach (var item2 in MtlQcItemResult)
                                    {
                                        #region //若有機台，整理序號格式
                                        string QcItemNo = item2.QcItemNo;
                                        if (item2.MachineNumber != null)
                                        {
                                            QcItemNo = item2.QcItemNo;
                                            string firstPart = QcItemNo.Substring(0, 3);
                                            string secondPart = QcItemNo.Substring(3, 4);
                                            QcItemNo = firstPart + item2.MachineNumber + secondPart;
                                        }
                                        #endregion

                                        string QcItemName = item2.QcItemName;

                                        #region //設定量測項目、備註、上下限
                                        data = new Data()
                                        {
                                            cell = "A" + row,
                                            css = "",
                                            format = "common",
                                            value = QcItemNo,
                                        };
                                        datas.Add(data);

                                        data = new Data()
                                        {
                                            cell = "B" + row,
                                            css = "",
                                            format = "common",
                                            value = QcItemName,
                                        };
                                        datas.Add(data);

                                        data = new Data()
                                        {
                                            cell = "C" + row,
                                            css = "",
                                            format = "common",
                                            value = item2.QcItemDesc,
                                        };
                                        datas.Add(data);

                                        data = new Data()
                                        {
                                            cell = "D" + row,
                                            css = "",
                                            format = "common",
                                            value = item2.MachineDesc != null ? item2.MachineDesc : "",
                                        };
                                        datas.Add(data);

                                        data = new Data()
                                        {
                                            cell = "E" + row,
                                            css = "",
                                            format = "common",
                                            value = item2.DesignValue != null ? item2.DesignValue.ToString() : "",
                                        };
                                        datas.Add(data);

                                        data = new Data()
                                        {
                                            cell = "F" + row,
                                            css = "",
                                            format = "common",
                                            value = item2.UpperTolerance != null ? item2.UpperTolerance.ToString() : "",
                                        };
                                        datas.Add(data);

                                        data = new Data()
                                        {
                                            cell = "G" + row,
                                            css = "",
                                            format = "common",
                                            value = item2.LowerTolerance != null ? item2.LowerTolerance.ToString() : "",
                                        };
                                        datas.Add(data);

                                        row++;
                                        #endregion
                                    }
                                    #endregion

                                    #region //整合Spreadsheet格式
                                    List<Sheets> sheetss = new List<Sheets>();

                                    Sheets sheets = new Sheets()
                                    {
                                        name = "sheet1",
                                        data = datas
                                    };
                                    sheetss.Add(sheets);

                                    #region //更新至QcRecord SpreadsheetData
                                    SpreadsheetModel spreadsheetModel = new SpreadsheetModel()
                                    {
                                        sheets = sheetss
                                    };

                                    string SpreadsheetData = JsonConvert.SerializeObject(spreadsheetModel);

                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE MES.QcRecord SET
                                            SpreadsheetData = @SpreadsheetData,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE QcRecordId = @QcRecordId";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            SpreadsheetData,
                                            LastModifiedDate,
                                            LastModifiedBy,
                                            QcRecordId
                                        });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                    #endregion
                                    #endregion
                                }
                                #endregion

                                count++;
                            }
                            #endregion

                            #region //重新計算進貨單資料
                            JObject updateGrDetailResult = JObject.Parse(UpdateGrTotal(GrId, goodsReceipts[0].Exchange, sqlConnection, sqlConnection2));
                            if (updateGrDetailResult["status"].ToString() != "success")
                            {
                                throw new SystemException(updateGrDetailResult["msg"].ToString());
                            }
                            #endregion

                            #region //依稅別碼更新發票聯數
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(NM001)) InvoiceCountNo, LTRIM(RTRIM(NM002)) InvoiceCountName 
                                    FROM CMSNM a
                                    INNER JOIN CMSNN b ON a.NM001 = b.NN005
                                    WHERE b.NN001 = @TaxNo";
                            dynamicParameters.Add("TaxNo", TaxNo);

                            var InvoiceTypeResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (InvoiceTypeResult.Count() <= 0) throw new SystemException("發票聯數資料錯誤!!");

                            string invoiceType = InvoiceTypeResult.FirstOrDefault().InvoiceCountNo;
                            #endregion

                            #region //更新進貨單頭交易條件
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.GoodsReceipt SET
                                    CurrencyCode = @CurrencyCode,
                                    PaymentTerm = @PaymentTerm,
                                    TaxType = @TaxType,
                                    TaxRate = @TaxRate,
                                    TaxCode = @TaxCode,
                                    TradeTerm = @TradeTerm,
                                    InvoiceType = @InvoiceType,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE GrId = @GrId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    goodsReceipts[0].CurrencyCode,
                                    PaymentTerm = PaymentTermNo,
                                    TaxType = Taxation,
                                    TaxRate,
                                    TaxCode = TaxNo,
                                    TradeTerm,
                                    InvoiceType = invoiceType,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    GrId
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

        #region //ImportPoDetail -- 從採購單身帶出進貨單身 -- Ann 2024-03-06
        public string ImportPoDetail(int GrId, string PoDetailIdList)
        {
            try
            {
                if (PoDetailIdList.Length <= 0) throw new SystemException("【採購單身】不能為空!!");

                string companyNo = "";
                int rowsAffected = 0;

                List<PoDetail> poDetails = new List<PoDetail>();

                List<SpreadsheetModel> spreadsheetModels = new List<SpreadsheetModel>();
                List<Sheets> sheetss = new List<Sheets>();
                List<Data> datas = new List<Data>();
                Dictionary<string, string> BarcodeInfo = new Dictionary<string, string>();
                #region //初始化Data
                Data data = new Data();
                data = new Data()
                {
                    cell = "A1",
                    css = "imported_class1",
                    format = "common",
                    value = "序號",
                };
                datas.Add(data);

                data = new Data()
                {
                    cell = "B1",
                    css = "imported_class2",
                    format = "common",
                    value = "檢測項目",
                };
                datas.Add(data);

                data = new Data()
                {
                    cell = "C1",
                    css = "imported_class3",
                    format = "common",
                    value = "檢測備註",
                };
                datas.Add(data);

                data = new Data()
                {
                    cell = "D1",
                    css = "imported_class4",
                    format = "common",
                    value = "量測設備",
                };
                datas.Add(data);

                data = new Data()
                {
                    cell = "E1",
                    css = "imported_class5",
                    format = "common",
                    value = "設計值",
                };
                datas.Add(data);

                data = new Data()
                {
                    cell = "F1",
                    css = "imported_class6",
                    format = "common",
                    value = "上限",
                };
                datas.Add(data);

                data = new Data()
                {
                    cell = "G1",
                    css = "imported_class7",
                    format = "common",
                    value = "下限",
                };
                datas.Add(data);

                data = new Data()
                {
                    cell = "H1",
                    css = "imported_class8",
                    format = "common",
                    value = "Z軸",
                };
                datas.Add(data);

                data = new Data()
                {
                    cell = "I1",
                    css = "imported_class9",
                    format = "common",
                    value = "量測人員",
                };
                datas.Add(data);

                data = new Data()
                {
                    cell = "J1",
                    css = "imported_class10",
                    format = "common",
                    value = "量測工時",
                };
                datas.Add(data);
                #endregion

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
                            companyNo = item.ErpNo;
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //確認進貨單資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.GrId, a.CompanyId, a.GrErpPrefix, a.GrErpNo, GrErpPrefix + '-' + GrErpNo GrErpFullNo, a.ReceiptDate, a.SupplierId, a.SupplierNo
                                    , a.SupplierSo, a.CurrencyCode, a.Exchange, a.InvoiceType, a.TaxType, a.InvoiceNo, a.PrintCnt, a.ConfirmStatus, a.DocDate
                                    , a.RenewFlag, a.Remark, a.TotalAmount, a.DeductAmount, a.OrigTax, a.ReceiptAmount, a.SupplierName, a.UiNo, a.DeductType
                                    , a.ObaccoAndLiquorFlag, a.RowCnt, a.Quantity, FORMAT(a.InvoiceDate, 'yyyy-MM-dd') InvoiceDate, a.OrigPreTaxAmount, a.ApplyYYMM
                                    , a.TaxRate, a.PreTaxAmount, a.TaxAmount, a.PaymentTerm, a.PoId, a.PrepaidErpPrefix, a.PrepaidErpNo, a.OffsetAmount, a.TaxOffset, a.PackageQuantity
                                    , a.PremiumAmount, a.ApproveStatus, a.InvoiceFlag, a.ReserveTaxCode, a.SendCount, a.OrigPremiumAmount, a.EbcErpPreNo, a.EbcEdition, a.EbcFlag
                                    , a.FromDocType, a.NoticeFlag, a.TaxCode, a.TradeTerm, a.DetailMultiTax, a.TaxExchange, a.ContactUser, a.TransferStatus, a.TransferDate
                                    , a.CreateDate, a.LastModifiedDate, a.CreateBy, a.LastModifiedBy
                                    FROM SCM.GoodsReceipt a 
                                    WHERE a.GrId = @GrId";
                            dynamicParameters.Add("GrId", GrId);

                            List<GoodsReceipt> goodsReceipts = sqlConnection.Query<GoodsReceipt>(sql, dynamicParameters).ToList();

                            if (goodsReceipts.Count() <= 0) throw new SystemException("進貨單資料錯誤!!");
                            #endregion

                            string[] PoDetailIds = PoDetailIdList.Split(',');
                            foreach (var PoDetailId in PoDetailIds)
                            {
                                sheetss = new List<Sheets>();

                                #region //確認採購單身資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.PoDetailId, a.PoId, a.PoErpPrefix, a.PoErpNo, a.PoSeq, a.MtlItemId, a.PoMtlItemName, a.PoMtlItemSpec, a.InventoryId, a.Quantity
                                        , a.UomId, a.PoUnitPrice, a.PoPrice, a.PromiseDate, a.ReferencePrefix, a.Remark, a.SiQty, a.ClosureStatus, a.ConfirmStatus 
                                        , a.InventoryQty, a.SmallUnit, a.ReferenceNo, a.Project, a.ReferenceSeq, a.UrgentMtl, a.PrErpPrefix, a.PrErpNo, a.PrSequence 
                                        , a.FromDocType, a.MtlItemType, a.PartialSeq, a.OriPromiseDate, a.DeliveryDate, a.PoPriceQty, a.PoPriceUomId, a.SiPriceQty 
                                        , a.DiscountRate, a.DiscountAmount, a.CreateDate, a.LastModifiedDate, a.CreateBy, a.LastModifiedBy
                                        , b.UomNo
                                        , c.InventoryNo
                                        , d.MtlItemNo
                                        , e.TaxNo, e.PaymentTermNo
                                        FROM SCM.PoDetail a
                                        INNER JOIN PDM.UnitOfMeasure b ON a.UomId = b.UomId
                                        INNER JOIN SCM.Inventory c ON a.InventoryId = c.InventoryId
                                        INNER JOIN PDM.MtlItem d ON a.MtlItemId = d.MtlItemId
                                        INNER JOIN SCM.PurchaseOrder e ON a.PoId = e.PoId
                                        WHERE a.PoDetailId = @PoDetailId";
                                dynamicParameters.Add("PoDetailId", PoDetailId);

                                poDetails = sqlConnection.Query<PoDetail>(sql, dynamicParameters).ToList();
                                #endregion

                                #region //原幣小數點取位
                                int OriUnitDecimal = 0;
                                int OriPriceDecimal = 0;
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.MF003, a.MF004
                                        FROM CMSMF a 
                                        WHERE a.MF001 = @Currency";
                                dynamicParameters.Add("Currency", goodsReceipts[0].CurrencyCode);

                                var CMSMFResult = sqlConnection2.Query(sql, dynamicParameters);

                                if (CMSMFResult.Count() <= 0) throw new SystemException("金額小數點取位資料錯誤!!");

                                foreach (var item in CMSMFResult)
                                {
                                    OriUnitDecimal = Convert.ToInt32(item.MF003);
                                    OriPriceDecimal = Convert.ToInt32(item.MF004);
                                }
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

                                #region //本幣小數點取位
                                int UnitDecimal = 0;
                                int PriceDecimal = 0;
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.MF003, a.MF004
                                        FROM CMSMF a 
                                        WHERE a.MF001 = @Currency";
                                dynamicParameters.Add("Currency", CurrencyCode);

                                var CMSMFResult2 = sqlConnection2.Query(sql, dynamicParameters);

                                if (CMSMFResult2.Count() <= 0) throw new SystemException("金額小數點取位資料錯誤!!");

                                foreach (var item in CMSMFResult2)
                                {
                                    UnitDecimal = Convert.ToInt32(item.MF003);
                                    PriceDecimal = Convert.ToInt32(item.MF004);
                                }
                                #endregion

                                #region //INSERT SCM.GrDetail
                                int count = 1;
                                foreach (var item in poDetails)
                                {
                                    #region //確認結案碼狀態
                                    if (item.ClosureStatus != "N") throw new SystemException("採購單單身【" + item.PoErpPrefix + "-" + item.PoErpNo + "(" + item.PoSeq + ")】已結案!!");
                                    #endregion

                                    #region //確認ERP採購單狀態是否正確
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT LTRIM(RTRIM(TC014)) TC014
                                            FROM PURTC 
                                            WHERE TC001 = @PoErpPrefix
                                            AND TC002 = @PoErpNo";
                                    dynamicParameters.Add("PoErpPrefix", item.PoErpPrefix);
                                    dynamicParameters.Add("PoErpNo", item.PoErpNo);

                                    var PURTCResult = sqlConnection2.Query(sql, dynamicParameters);

                                    if (PURTCResult.Count() <= 0) throw new SystemException("ERP採購單資料錯誤!!");

                                    foreach (var item2 in PURTCResult)
                                    {
                                        if (item2.TC014 != "Y") throw new SystemException("ERP採購單【" + item.PoErpPrefix + "-" + item.PoErpNo + "】狀態錯誤!!");
                                    }
                                    #endregion

                                    #region //確認ERP採購單身結案碼
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT LTRIM(RTRIM(TD016)) TD016
                                            , LTRIM(RTRIM(TD018)) TD018
                                            FROM PURTD
                                            WHERE TD001 = @PoErpPrefix
                                            AND TD002 = @PoErpNo
                                            AND TD003 = @PoSeq";
                                    dynamicParameters.Add("PoErpPrefix", item.PoErpPrefix);
                                    dynamicParameters.Add("PoErpNo", item.PoErpNo);
                                    dynamicParameters.Add("PoSeq", item.PoSeq);

                                    var PURTDResult = sqlConnection2.Query(sql, dynamicParameters);

                                    if (PURTDResult.Count() <= 0) throw new SystemException("ERP採購單單身【" + item.PoErpPrefix + "-" + item.PoErpNo + "(" + item.PoSeq + ")】資料錯誤!!");

                                    foreach (var item2 in PURTDResult)
                                    {
                                        if (item2.TD016 != "N") throw new SystemException("ERP採購單單身【" + item.PoErpPrefix + "-" + item.PoErpNo + "(" + item.PoSeq + ")】已結案!!");
                                        if (item2.TD018 != "Y") throw new SystemException("ERP採購單單身【" + item.PoErpPrefix + "-" + item.PoErpNo + "(" + item.PoSeq + ")】狀態錯誤!!");
                                    }
                                    #endregion

                                    #region //查詢單身序號
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT ISNULL(MAX(a.GrSeq), 0) MaxGrSeq
                                            FROM SCM.GrDetail a 
                                            WHERE a.GrId = @GrId";
                                    dynamicParameters.Add("GrId", GrId);

                                    var GrSeqResult = sqlConnection.Query(sql, dynamicParameters);

                                    int MaxGrSeq = 0;
                                    foreach (var item2 in GrSeqResult)
                                    {
                                        MaxGrSeq = Convert.ToInt32(item2.MaxGrSeq);
                                    }

                                    string GrSeq = (MaxGrSeq + 1).ToString("D4");
                                    #endregion

                                    #region //確認品號檢驗狀態
                                    string QcStatus = "";
                                    string QcType = "";
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT MB043
                                            FROM INVMB
                                            WHERE MB001 = @MB001";
                                    dynamicParameters.Add("MB001", item.MtlItemNo);

                                    var INVMBResult = sqlConnection2.Query(sql, dynamicParameters);

                                    if (INVMBResult.Count() <= 0) throw new SystemException("此進貨品號【" + item.MtlItemNo + "】資料錯誤!!");

                                    foreach (var item2 in INVMBResult)
                                    {
                                        if (item2.MB043 == "0")
                                        {
                                            QcStatus = "0";
                                            QcType = "N";
                                        }
                                        else
                                        {
                                            QcStatus = "1";
                                            QcType = "1";
                                        }
                                    }
                                    #endregion

                                    #region //查詢目前庫存數
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT ISNULL(MC007, 0) MC007
                                            FROM INVMC
                                            WHERE MC001 = @MC001
                                            AND MC002 = @MC002";
                                    dynamicParameters.Add("MC001", item.MtlItemNo);
                                    dynamicParameters.Add("MC002", item.InventoryNo);

                                    var INVMCResult = sqlConnection2.Query(sql, dynamicParameters);

                                    double InventoryQty = 0;
                                    foreach (var item2 in INVMCResult)
                                    {
                                        InventoryQty = Convert.ToDouble(item2.MC007);
                                    }
                                    #endregion

                                    #region //處理【原幣進貨金額】、【原幣未稅金額】、【原幣稅金額】、【本幣未稅金額】、【本幣稅金額】
                                    double? AvailableQty = QcStatus == "0" ? item.PoPriceQty - item.SiPriceQty : 0; //計價數量
                                    double? OrigUnitPrice = Math.Round(Convert.ToDouble(item.PoUnitPrice), OriUnitDecimal); //原幣進貨單價
                                    double? OrigDiscountAmt = 0; //原幣扣款金額
                                    double? OrigAmount = Math.Round(Convert.ToDouble(AvailableQty * OrigUnitPrice), OriPriceDecimal); //原幣進貨金額

                                    double OrigPreTaxAmt = 0; //原幣未稅金額
                                    double OrigTaxAmt = 0; //原幣稅金額
                                    double PreTaxAmt = 0; //本幣未稅金額
                                    double TaxAmt = 0; //本幣稅金額

                                    JObject calculateTaxAmtResult = JObject.Parse(CalculateTaxAmt(AvailableQty, OrigUnitPrice, OrigDiscountAmt, OrigAmount, goodsReceipts[0].TaxRate, goodsReceipts[0].TaxType, goodsReceipts[0].Exchange
                                                                                , OriUnitDecimal, OriPriceDecimal, UnitDecimal, PriceDecimal));

                                    if (calculateTaxAmtResult["status"].ToString() == "success")
                                    {
                                        OrigPreTaxAmt = Convert.ToDouble(calculateTaxAmtResult["OrigPreTaxAmt"]);
                                        OrigTaxAmt = Convert.ToDouble(calculateTaxAmtResult["OrigTaxAmt"]);
                                        PreTaxAmt = Convert.ToDouble(calculateTaxAmtResult["PreTaxAmt"]);
                                        TaxAmt = Convert.ToDouble(calculateTaxAmtResult["TaxAmt"]);
                                    }
                                    else
                                    {
                                        throw new SystemException(calculateTaxAmtResult["msg"].ToString());
                                    }
                                    #endregion

                                    #region //確認此進貨單身無相同採購單身
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT ISNULL(SUM(a.ReceiptQty), 0) TotalReceiptQty
                                            FROM SCM.GrDetail a 
                                            WHERE a.PoErpPrefix = @PoErpPrefix
                                            AND a.PoErpNo = @PoErpNo
                                            AND a.PoSeq = @PoSeq
                                            AND a.GrId = @GrId";
                                    dynamicParameters.Add("PoErpPrefix", item.PoErpPrefix);
                                    dynamicParameters.Add("PoErpNo", item.PoErpNo);
                                    dynamicParameters.Add("PoSeq", item.PoSeq);
                                    dynamicParameters.Add("GrId", GrId);

                                    var PoDetailResult = sqlConnection.Query(sql, dynamicParameters);

                                    double? TotalReceiptQty = 0;
                                    foreach (var item2 in PoDetailResult)
                                    {
                                        TotalReceiptQty = item2.TotalReceiptQty;
                                    }
                                    #endregion

                                    #region //INSERT SCM.GrDetail
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO SCM.GrDetail (GrId, GrErpPrefix, GrErpNo, GrSeq, MtlItemId, MtlItemNo, GrMtlItemName, GrMtlItemSpec
                                            , ReceiptQty, UomId, UomNo, InventoryId, InventoryNo, LotNumber, PoErpPrefix, PoErpNo, PoSeq, AcceptanceDate
                                            , AcceptQty, AvailableQty, ReturnQty, OrigUnitPrice, OrigAmount, OrigDiscountAmt, TrErpPrefix, TrErpNo, TrSeq
                                            , ReceiptExpense, DiscountDescription, PaymentHold, Overdue, QcStatus, ReturnCode, ConfirmStatus, CloseStatus
                                            , ReNewStatus, Remark, InventoryQty, SmallUnit, AvailableDate, ReCheckDate, ConfirmUser, ApvErpPrefix, ApvErpNo
                                            , ApvSeq, ProjectCode, ExpenseEntry, PremiumAmountFlag, OrigPreTaxAmt, OrigTaxAmt, PreTaxAmt, TaxAmt, ReceiptPackageQty
                                            , AcceptancePackageQty, ReturnPackageQty, PremiumAmount, PackageUnit, ReserveTaxCode, OrigPremiumAmount, AvailableUomId
                                            , AvailableUomNo, OrigCustomer, ApproveStatus, EbcErpPreNo, EbcEdition, ProductSeqAmount, MtlItemType, Loaction
                                            , TaxRate, MultipleFlag, GrErpStatus, TaxCode, DiscountRate, DiscountPrice, QcType, TransferStatus, TransferDate
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.GrDetailId
                                            VALUES (@GrId, @GrErpPrefix, @GrErpNo, @GrSeq, @MtlItemId, @MtlItemNo, @GrMtlItemName, @GrMtlItemSpec
                                            , @ReceiptQty, @UomId, @UomNo, @InventoryId, @InventoryNo, @LotNumber, @PoErpPrefix, @PoErpNo, @PoSeq, @AcceptanceDate
                                            , @AcceptQty, @AvailableQty, @ReturnQty, @OrigUnitPrice, @OrigAmount, @OrigDiscountAmt, @TrErpPrefix, @TrErpNo, @TrSeq
                                            , @ReceiptExpense, @DiscountDescription, @PaymentHold, @Overdue, @QcStatus, @ReturnCode, @ConfirmStatus, @CloseStatus
                                            , @ReNewStatus, @Remark, @InventoryQty, @SmallUnit, @AvailableDate, @ReCheckDate, @ConfirmUser, @ApvErpPrefix, @ApvErpNo
                                            , @ApvSeq, @ProjectCode, @ExpenseEntry, @PremiumAmountFlag, @OrigPreTaxAmt, @OrigTaxAmt, @PreTaxAmt, @TaxAmt, @ReceiptPackageQty
                                            , @AcceptancePackageQty, @ReturnPackageQty, @PremiumAmount, @PackageUnit, @ReserveTaxCode, @OrigPremiumAmount, @AvailableUomId
                                            , @AvailableUomNo, @OrigCustomer, @ApproveStatus, @EbcErpPreNo, @EbcEdition, @ProductSeqAmount, @MtlItemType, @Loaction
                                            , @TaxRate, @MultipleFlag, @GrErpStatus, @TaxCode, @DiscountRate, @DiscountPrice, @QcType, @TransferStatus, @TransferDate
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            GrId,
                                            goodsReceipts[0].GrErpPrefix,
                                            goodsReceipts[0].GrErpNo,
                                            GrSeq,
                                            item.MtlItemId,
                                            item.MtlItemNo,
                                            GrMtlItemName = item.PoMtlItemName,
                                            GrMtlItemSpec = item.PoMtlItemSpec,
                                            ReceiptQty = item.Quantity - item.SiQty - TotalReceiptQty > 0 ? item.Quantity - item.SiQty - TotalReceiptQty : 0,
                                            item.UomId,
                                            item.UomNo,
                                            item.InventoryId,
                                            item.InventoryNo,
                                            LotNumber = "",
                                            item.PoErpPrefix,
                                            item.PoErpNo,
                                            item.PoSeq,
                                            AcceptanceDate = DateTime.Now.ToString("yyyyMMdd"),
                                            AcceptQty = QcStatus != "0" ? 0 : item.Quantity - item.SiQty,
                                            AvailableQty,
                                            ReturnQty = 0,
                                            OrigUnitPrice,
                                            OrigAmount,
                                            OrigDiscountAmt,
                                            TrErpPrefix = "",
                                            TrErpNo = "",
                                            TrSeq = "",
                                            ReceiptExpense = 0,
                                            DiscountDescription = "",
                                            PaymentHold = "N",
                                            Overdue = "N",
                                            QcStatus,
                                            ReturnCode = "N",
                                            ConfirmStatus = "N",
                                            CloseStatus = "N",
                                            ReNewStatus = "N",
                                            item.Remark,
                                            InventoryQty,
                                            SmallUnit = "",
                                            AvailableDate = "",
                                            ReCheckDate = "",
                                            ConfirmUser = (int?)null,
                                            ApvErpPrefix = "",
                                            ApvErpNo = "",
                                            ApvSeq = "",
                                            ProjectCode = item.Project,
                                            ExpenseEntry = "N",
                                            PremiumAmountFlag = "N",
                                            OrigPreTaxAmt,
                                            OrigTaxAmt,
                                            PreTaxAmt,
                                            TaxAmt,
                                            ReceiptPackageQty = 0,
                                            AcceptancePackageQty = 0,
                                            ReturnPackageQty = 0,
                                            PremiumAmount = 0,
                                            PackageUnit = "",
                                            ReserveTaxCode = "N",
                                            OrigPremiumAmount = "",
                                            AvailableUomId = item.UomId,
                                            AvailableUomNo = item.UomNo,
                                            OrigCustomer = "",
                                            ApproveStatus = "N",
                                            EbcErpPreNo = "",
                                            EbcEdition = "",
                                            ProductSeqAmount = 0,
                                            MtlItemType = "2",
                                            Loaction = "",
                                            goodsReceipts[0].TaxRate,
                                            MultipleFlag = "N",
                                            GrErpStatus = "0",
                                            goodsReceipts[0].TaxCode,
                                            item.DiscountRate,
                                            DiscountPrice = item.DiscountAmount,
                                            QcType,
                                            TransferStatus = "N",
                                            TransferDate = (DateTime?)null,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });
                                    var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult.Count();

                                    int GrDetailId = -1;
                                    foreach (var item2 in insertResult)
                                    {
                                        GrDetailId = item2.GrDetailId;
                                    }
                                    #endregion

                                    #region //若品項為待驗品，自動新增進貨檢單據
                                    if (QcStatus == "1")
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

                                        #region //取得進貨檢QcType
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT TOP 1 QcTypeId
                                                FROM QMS.QcType
                                                WHERE QcTypeNo = 'IQC'";

                                        var QcTypeResult = sqlConnection.Query(sql, dynamicParameters);

                                        int QcTypeId = -1;
                                        foreach (var item2 in QcTypeResult)
                                        {
                                            QcTypeId = item2.QcTypeId;
                                        }
                                        #endregion

                                        #region //INSERT MES.QcRecord
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO MES.QcRecord (QcTypeId, InputType, Remark, DefaultFileId, DefaultSpreadsheetData, ResolveFileJson
                                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                                OUTPUT INSERTED.QcRecordId
                                                VALUES (@QcTypeId, @InputType, @Remark, @DefaultFileId, @DefaultSpreadsheetData, @ResolveFileJson
                                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                QcTypeId,
                                                InputType = "LotNumber",
                                                Remark = "由進貨單自動建立",
                                                DefaultFileId = -1,
                                                DefaultSpreadsheetData,
                                                ResolveFileJson = "",
                                                CreateDate,
                                                LastModifiedDate,
                                                CreateBy,
                                                LastModifiedBy
                                            });

                                        var insertQcRecordResult = sqlConnection.Query(sql, dynamicParameters);

                                        rowsAffected += insertQcRecordResult.Count();

                                        int QcRecordId = -1;
                                        foreach (var item2 in insertQcRecordResult)
                                        {
                                            QcRecordId = item2.QcRecordId;
                                        }
                                        #endregion

                                        #region //INSERT MES.QcGoodsReceipt
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO MES.QcGoodsReceipt (QcRecordId, GrDetailId
                                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                                OUTPUT INSERTED.QcGoodsReceiptId
                                                VALUES (@QcRecordId, @GrDetailId
                                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                QcRecordId,
                                                GrDetailId,
                                                CreateDate,
                                                LastModifiedDate,
                                                CreateBy,
                                                LastModifiedBy
                                            });

                                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                        #endregion

                                        #region //建立量測單據Spreadsheet DATA
                                        #region //取得品號允收標準
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT a.MtlQcItemId, a.MtlItemId, a.QcItemId, a.DesignValue, a.UpperTolerance, a.LowerTolerance
                                                , a.QcItemDesc, a.SortNumber, a.QmmDetailId
                                                , b.QcItemNo, b.QcItemName, b.QcItemDesc OriQcItemDesc
                                                , c.MachineNumber
                                                , d.MachineDesc
                                                FROM PDM.MtlQcItem a
                                                INNER JOIN QMS.QcItem b ON a.QcItemId = b.QcItemId
                                                LEFT JOIN QMS.QmmDetail c ON a.QmmDetailId = c.QmmDetailId
                                                LEFT JOIN MES.Machine d ON c.MachineId = d.MachineId
                                                WHERE a.MtlItemId = @MtlItemId
                                                AND a.[Status] = 'A' 
                                                ORDER BY a.SortNumber";
                                        dynamicParameters.Add("MtlItemId", item.MtlItemId);

                                        var MtlQcItemResult = sqlConnection.Query(sql, dynamicParameters);
                                        #endregion

                                        #region //設定單身量測標準
                                        int row = 2;
                                        foreach (var item2 in MtlQcItemResult)
                                        {
                                            #region //若有機台，整理序號格式
                                            string QcItemNo = item2.QcItemNo;
                                            if (item2.MachineNumber != null)
                                            {
                                                QcItemNo = item2.QcItemNo;
                                                string firstPart = QcItemNo.Substring(0, 3);
                                                string secondPart = QcItemNo.Substring(3, 4);
                                                QcItemNo = firstPart + item2.MachineNumber + secondPart;
                                            }
                                            #endregion

                                            string QcItemName = item2.QcItemName;

                                            #region //設定量測項目、備註、上下限
                                            data = new Data()
                                            {
                                                cell = "A" + row,
                                                css = "",
                                                format = "common",
                                                value = QcItemNo,
                                            };
                                            datas.Add(data);

                                            data = new Data()
                                            {
                                                cell = "B" + row,
                                                css = "",
                                                format = "common",
                                                value = QcItemName,
                                            };
                                            datas.Add(data);

                                            data = new Data()
                                            {
                                                cell = "C" + row,
                                                css = "",
                                                format = "common",
                                                value = item2.QcItemDesc,
                                            };
                                            datas.Add(data);

                                            data = new Data()
                                            {
                                                cell = "D" + row,
                                                css = "",
                                                format = "common",
                                                value = item2.MachineDesc != null ? item2.MachineDesc : "",
                                            };
                                            datas.Add(data);

                                            data = new Data()
                                            {
                                                cell = "E" + row,
                                                css = "",
                                                format = "common",
                                                value = item2.DesignValue != null ? item2.DesignValue.ToString() : "",
                                            };
                                            datas.Add(data);

                                            data = new Data()
                                            {
                                                cell = "F" + row,
                                                css = "",
                                                format = "common",
                                                value = item2.UpperTolerance != null ? item2.UpperTolerance.ToString() : "",
                                            };
                                            datas.Add(data);

                                            data = new Data()
                                            {
                                                cell = "G" + row,
                                                css = "",
                                                format = "common",
                                                value = item2.LowerTolerance != null ? item2.LowerTolerance.ToString() : "",
                                            };
                                            datas.Add(data);

                                            row++;
                                            #endregion
                                        }
                                        #endregion

                                        #region //整合Spreadsheet格式
                                        Sheets sheets = new Sheets()
                                        {
                                            name = "sheet1",
                                            data = datas
                                        };
                                        sheetss.Add(sheets);

                                        SpreadsheetModel spreadsheetModel = new SpreadsheetModel()
                                        {
                                            sheets = sheetss
                                        };

                                        string SpreadsheetData = JsonConvert.SerializeObject(spreadsheetModel);

                                        #region //更新至QcRecord SpreadsheetData
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"UPDATE MES.QcRecord SET
                                                SpreadsheetData = @SpreadsheetData,
                                                LastModifiedDate = @LastModifiedDate,
                                                LastModifiedBy = @LastModifiedBy
                                                WHERE QcRecordId = @QcRecordId";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                SpreadsheetData,
                                                LastModifiedDate,
                                                LastModifiedBy,
                                                QcRecordId
                                            });

                                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                        #endregion
                                        #endregion
                                        #endregion
                                    }
                                    #endregion

                                    count++;
                                }
                                #endregion
                            }

                            #region //重新計算進貨單資料
                            JObject updateGrDetailResult = JObject.Parse(UpdateGrTotal(GrId, goodsReceipts[0].Exchange, sqlConnection, sqlConnection2));
                            if (updateGrDetailResult["status"].ToString() != "success")
                            {
                                throw new SystemException(updateGrDetailResult["msg"].ToString());
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

        #region //AddGrDetail -- 新增進貨單單身 -- Ann 2024-03-08
        public string AddGrDetail(int GrId, int PoDetailId, int InventoryId, string AcceptanceDate, double ReceiptQty, double ReceiptExpense
            , double AcceptQty, double AvailableQty, double ReturnQty, int UomId, string QcStatus
            , double OrigUnitPrice, double OrigAmount, double OrigDiscountAmt, int? MtlItemId
            , string DiscountDescription, string Remark, string LotNumber)
        {
            try
            {
                string companyNo = "";
                int rowsAffected = 0;
                string MtlItemNo = "", MtlItemName = "", MtlItemSpec = "";
                double? Quantity = 0;
                double? SiQty = 0;
                string PoErpPrefix = "", PoErpNo = "", PoSeq = "";
                string PoRemark = "", Project = "";
                double? DiscountRate = 0, DiscountAmount = 0;

                #region //基本判斷
                if (ReceiptQty < 0 || ReceiptExpense < 0 || AcceptQty < 0 || AvailableQty < 0 || ReturnQty < 0) throw new SystemException("【進貨數量】【驗收數量】【驗退數量】【計價數量】不可小於0!!");
                if ((AcceptQty + ReturnQty) != ReceiptQty && (AcceptQty != 0 && ReturnQty != 0)) throw new SystemException("【驗收數量】+【驗退數量】需等於【進貨數量】!!");
                if (AvailableQty > ReceiptQty) throw new SystemException("【計價數量】不可大於【進貨數量】!!");
                if (QcStatus.Length <= 0) throw new SystemException("【檢驗狀態】格式錯誤!!");
                if (QcStatus == "1" && AcceptQty != 0) throw new SystemException("此品號需檢驗，不可提前輸入驗收數量!!");
                if (QcStatus == "1" && ReturnQty != 0) throw new SystemException("此品號需檢驗，不可提前輸入驗退數量!!");
                if (QcStatus == "1" && AvailableQty != 0) throw new SystemException("此品號需檢驗，不可提前輸入計價數量!!");
                #endregion

                List<PoDetail> poDetails = new List<PoDetail>();

                List<SpreadsheetModel> spreadsheetModels = new List<SpreadsheetModel>();
                List<Sheets> sheetss = new List<Sheets>();
                List<Data> datas = new List<Data>();
                Dictionary<string, string> BarcodeInfo = new Dictionary<string, string>();
                #region //初始化Data
                Data data = new Data();
                data = new Data()
                {
                    cell = "A1",
                    css = "imported_class1",
                    format = "common",
                    value = "序號",
                };
                datas.Add(data);

                data = new Data()
                {
                    cell = "B1",
                    css = "imported_class2",
                    format = "common",
                    value = "檢測項目",
                };
                datas.Add(data);

                data = new Data()
                {
                    cell = "C1",
                    css = "imported_class3",
                    format = "common",
                    value = "檢測備註",
                };
                datas.Add(data);

                data = new Data()
                {
                    cell = "D1",
                    css = "imported_class4",
                    format = "common",
                    value = "量測設備",
                };
                datas.Add(data);

                data = new Data()
                {
                    cell = "E1",
                    css = "imported_class5",
                    format = "common",
                    value = "設計值",
                };
                datas.Add(data);

                data = new Data()
                {
                    cell = "F1",
                    css = "imported_class6",
                    format = "common",
                    value = "上限",
                };
                datas.Add(data);

                data = new Data()
                {
                    cell = "G1",
                    css = "imported_class7",
                    format = "common",
                    value = "下限",
                };
                datas.Add(data);

                data = new Data()
                {
                    cell = "H1",
                    css = "imported_class8",
                    format = "common",
                    value = "Z軸",
                };
                datas.Add(data);

                data = new Data()
                {
                    cell = "I1",
                    css = "imported_class9",
                    format = "common",
                    value = "量測人員",
                };
                datas.Add(data);

                data = new Data()
                {
                    cell = "J1",
                    css = "imported_class10",
                    format = "common",
                    value = "量測工時",
                };
                datas.Add(data);
                #endregion

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
                            companyNo = item.ErpNo;
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //確認進貨單資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.GrId, a.CompanyId, a.GrErpPrefix, a.GrErpNo, GrErpPrefix + '-' + GrErpNo GrErpFullNo, a.ReceiptDate, a.SupplierId, a.SupplierNo
                                    , a.SupplierSo, a.CurrencyCode, a.Exchange, a.InvoiceType, a.TaxType, a.InvoiceNo, a.PrintCnt, a.ConfirmStatus, a.DocDate
                                    , a.RenewFlag, a.Remark, a.TotalAmount, a.DeductAmount, a.OrigTax, a.ReceiptAmount, a.SupplierName, a.UiNo, a.DeductType
                                    , a.ObaccoAndLiquorFlag, a.RowCnt, a.Quantity, FORMAT(a.InvoiceDate, 'yyyy-MM-dd') InvoiceDate, a.OrigPreTaxAmount, a.ApplyYYMM
                                    , a.TaxRate, a.PreTaxAmount, a.TaxAmount, a.PaymentTerm, a.PoId, a.PrepaidErpPrefix, a.PrepaidErpNo, a.OffsetAmount, a.TaxOffset, a.PackageQuantity
                                    , a.PremiumAmount, a.ApproveStatus, a.InvoiceFlag, a.ReserveTaxCode, a.SendCount, a.OrigPremiumAmount, a.EbcErpPreNo, a.EbcEdition, a.EbcFlag
                                    , a.FromDocType, a.NoticeFlag, a.TaxCode, a.TradeTerm, a.DetailMultiTax, a.TaxExchange, a.ContactUser, a.TransferStatus, a.TransferDate
                                    , a.CreateDate, a.LastModifiedDate, a.CreateBy, a.LastModifiedBy
                                    FROM SCM.GoodsReceipt a 
                                    WHERE a.GrId = @GrId";
                            dynamicParameters.Add("GrId", GrId);

                            List<GoodsReceipt> goodsReceipts = sqlConnection.Query<GoodsReceipt>(sql, dynamicParameters).ToList();

                            if (goodsReceipts.Count() <= 0) throw new SystemException("進貨單資料錯誤!!");
                            #endregion

                            #region //確認ERP單據性質資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MQ019)) MQ019 
                                    FROM CMSMQ
                                    WHERE MQ001 = @GrErpPrefix";
                            dynamicParameters.Add("GrErpPrefix", goodsReceipts[0].GrErpPrefix);

                            var CMSMQResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (CMSMQResult.Count() <= 0) throw new SystemException("單據性質資料錯誤!!");

                            string MQ019 = "";
                            foreach (var item in CMSMQResult)
                            {
                                MQ019 = item.MQ019;

                                if (MQ019 == "Y")
                                {
                                    if (PoDetailId <= 0) throw new SystemException("此單據需【核對採購單】，採購單不能為空!!");
                                    if (MtlItemId > 0) throw new SystemException("此單據需【核對採購單】，無法自行輸入品號!!");
                                }
                            }
                            #endregion

                            #region //確認品號資料是否正確
                            if (MtlItemId > 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.MtlItemNo, a.MtlItemName, a.MtlItemSpec
                                        FROM PDM.MtlItem a 
                                        WHERE a.MtlItemId = @MtlItemId";
                                dynamicParameters.Add("MtlItemId", MtlItemId);

                                var MtlItemResult = sqlConnection.Query(sql, dynamicParameters);

                                if (MtlItemResult.Count() <= 0) throw new SystemException("品號資料錯誤!!");

                                foreach (var item in MtlItemResult)
                                {
                                    MtlItemNo = item.MtlItemNo;
                                    MtlItemName = item.MtlItemName;
                                    MtlItemSpec = item.MtlItemSpec;
                                }
                            }
                            #endregion

                            #region //確認採購單身資料是否正確
                            if (PoDetailId > 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.PoDetailId, a.PoId, a.PoErpPrefix, a.PoErpNo, a.PoSeq, a.MtlItemId, a.PoMtlItemName, a.PoMtlItemSpec, a.InventoryId, a.Quantity
                                        , a.UomId, a.PoUnitPrice, a.PoPrice, a.PromiseDate, a.ReferencePrefix, a.Remark, a.SiQty, a.ClosureStatus, a.ConfirmStatus 
                                        , a.InventoryQty, a.SmallUnit, a.ReferenceNo, a.Project, a.ReferenceSeq, a.UrgentMtl, a.PrErpPrefix, a.PrErpNo, a.PrSequence 
                                        , a.FromDocType, a.MtlItemType, a.PartialSeq, a.OriPromiseDate, a.DeliveryDate, a.PoPriceQty, a.PoPriceUomId, a.SiPriceQty 
                                        , a.DiscountRate, a.DiscountAmount, a.CreateDate, a.LastModifiedDate, a.CreateBy, a.LastModifiedBy
                                        , b.UomNo
                                        , c.InventoryNo
                                        , d.MtlItemNo
                                        , e.TaxNo
                                        FROM SCM.PoDetail a
                                        INNER JOIN PDM.UnitOfMeasure b ON a.UomId = b.UomId
                                        INNER JOIN SCM.Inventory c ON a.InventoryId = c.InventoryId
                                        INNER JOIN PDM.MtlItem d ON a.MtlItemId = d.MtlItemId
                                        INNER JOIN SCM.PurchaseOrder e ON a.PoId = e.PoId
                                        WHERE a.PoDetailId = @PoDetailId";
                                dynamicParameters.Add("PoDetailId", PoDetailId);

                                poDetails = sqlConnection.Query<PoDetail>(sql, dynamicParameters).ToList();

                                #region //確認ERP採購單狀態是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT LTRIM(RTRIM(TC014)) TC014
                                        FROM PURTC 
                                        WHERE TC001 = @PoErpPrefix
                                        AND TC002 = @PoErpNo";
                                dynamicParameters.Add("PoErpPrefix", poDetails[0].PoErpPrefix);
                                dynamicParameters.Add("PoErpNo", poDetails[0].PoErpNo);

                                var PURTCResult = sqlConnection2.Query(sql, dynamicParameters);

                                if (PURTCResult.Count() <= 0) throw new SystemException("ERP採購單資料錯誤!!");

                                foreach (var item in PURTCResult)
                                {
                                    if (item.TC014 != "Y") throw new SystemException("ERP採購單狀態錯誤!!");
                                }
                                #endregion
                            }
                            #endregion

                            #region //原幣小數點取位
                            int OriUnitDecimal = 0;
                            int OriPriceDecimal = 0;
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.MF003, a.MF004
                                    FROM CMSMF a 
                                    WHERE a.MF001 = @Currency";
                            dynamicParameters.Add("Currency", goodsReceipts[0].CurrencyCode);

                            var CMSMFResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (CMSMFResult.Count() <= 0) throw new SystemException("金額小數點取位資料錯誤!!");

                            foreach (var item in CMSMFResult)
                            {
                                OriUnitDecimal = Convert.ToInt32(item.MF003);
                                OriPriceDecimal = Convert.ToInt32(item.MF004);
                            }
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

                            #region //本幣小數點取位
                            int UnitDecimal = 0;
                            int PriceDecimal = 0;
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.MF003, a.MF004
                                    FROM CMSMF a 
                                    WHERE a.MF001 = @Currency";
                            dynamicParameters.Add("Currency", CurrencyCode);

                            var CMSMFResult2 = sqlConnection2.Query(sql, dynamicParameters);

                            if (CMSMFResult2.Count() <= 0) throw new SystemException("金額小數點取位資料錯誤!!");

                            foreach (var item in CMSMFResult2)
                            {
                                UnitDecimal = Convert.ToInt32(item.MF003);
                                PriceDecimal = Convert.ToInt32(item.MF004);
                            }
                            #endregion

                            #region //確認庫別資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT InventoryNo
                                    FROM SCM.Inventory
                                    WHERE InventoryId = @InventoryId";
                            dynamicParameters.Add("InventoryId", InventoryId);

                            var InventoryResult = sqlConnection.Query(sql, dynamicParameters);

                            if (InventoryResult.Count() <= 0) throw new SystemException("庫別資料錯誤!!");

                            string InventoryNo = "";
                            foreach (var item in InventoryResult)
                            {
                                InventoryNo = item.InventoryNo;
                            }
                            #endregion

                            #region //確認進貨單位資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT UomNo
                                    FROM PDM.UnitOfMeasure
                                    WHERE UomId = @UomId";
                            dynamicParameters.Add("UomId", UomId);

                            var UnitOfMeasureResult = sqlConnection.Query(sql, dynamicParameters);

                            if (UnitOfMeasureResult.Count() <= 0) throw new SystemException("進貨單位資料錯誤!!");

                            string UomNo = "";
                            foreach (var item in UnitOfMeasureResult)
                            {
                                UomNo = item.UomNo;
                            }
                            #endregion

                            #region //儲存單身相關資料
                            foreach (var item in poDetails)
                            {
                                if (MtlItemId <= 0)
                                {
                                    MtlItemId = item.MtlItemId;
                                    MtlItemNo = item.MtlItemNo;
                                    MtlItemName = item.PoMtlItemName;
                                    MtlItemSpec = item.PoMtlItemSpec;
                                }
                                Quantity = item.Quantity;
                                SiQty = item.SiQty;
                                PoErpPrefix = item.PoErpPrefix;
                                PoErpNo = item.PoErpNo;
                                PoSeq = item.PoSeq;
                                PoRemark = item.Remark;
                                Project = item.Project;
                                DiscountRate = item.DiscountRate;
                                DiscountAmount = item.DiscountAmount;

                                #region //確認ERP採購單身結案碼
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT LTRIM(RTRIM(TD016)) TD016
                                        , LTRIM(RTRIM(TD018)) TD018
                                        FROM PURTD
                                        WHERE TD001 = @PoErpPrefix
                                        AND TD002 = @PoErpNo
                                        AND TD003 = @PoSeq";
                                dynamicParameters.Add("PoErpPrefix", item.PoErpPrefix);
                                dynamicParameters.Add("PoErpNo", item.PoErpNo);
                                dynamicParameters.Add("PoSeq", item.PoSeq);

                                var PURTDResult = sqlConnection2.Query(sql, dynamicParameters);

                                if (PURTDResult.Count() <= 0) throw new SystemException("ERP採購單單身資料錯誤!!");

                                foreach (var item2 in PURTDResult)
                                {
                                    if (item2.TD016 != "N") throw new SystemException("ERP採購單單身結案碼狀態錯誤!!");
                                    if (item2.TD018 != "Y") throw new SystemException("ERP採購單單身狀態錯誤!!");
                                }
                                #endregion

                                #region //確認此進貨單身無相同採購單身
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT ISNULL(SUM(a.ReceiptQty), 0) TotalReceiptQty
                                        FROM SCM.GrDetail a 
                                        WHERE a.PoErpPrefix = @PoErpPrefix
                                        AND a.PoErpNo = @PoErpNo
                                        AND a.PoSeq = @PoSeq
                                        AND a.GrId = @GrId";
                                dynamicParameters.Add("PoErpPrefix", item.PoErpPrefix);
                                dynamicParameters.Add("PoErpNo", item.PoErpNo);
                                dynamicParameters.Add("PoSeq", item.PoSeq);
                                dynamicParameters.Add("GrId", GrId);

                                var PoDetailResult = sqlConnection.Query(sql, dynamicParameters);

                                foreach (var item2 in PoDetailResult)
                                {
                                    #region //確認數量沒有超過剩餘可進貨數量
                                    double poQty = Convert.ToDouble(item.Quantity);
                                    double siQty = Convert.ToDouble(item.SiQty);
                                    double TotalReceiptQty = item2.TotalReceiptQty;
                                    double allowQty = poQty - siQty - TotalReceiptQty > 0 ? poQty - siQty - TotalReceiptQty : 0;

                                    if (ReceiptQty > allowQty)
                                    {
                                        throw new SystemException("此採購單已超過剩餘可進貨數量: " + allowQty + "!!");
                                    }
                                    #endregion
                                }
                                #endregion
                            }
                            #endregion

                            #region //確認ERP品號相關資料是否正確
                            string QcType = "";
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MB004)) MB004
                                    , LTRIM(RTRIM(MB030)) MB030
                                    , LTRIM(RTRIM(MB031)) MB031
                                    , LTRIM(RTRIM(MB022)) MB022
                                    , LTRIM(RTRIM(MB043)) MB043
                                    , LTRIM(RTRIM(MB044)) MB044
                                    FROM INVMB
                                    WHERE MB001 = @MB001";
                            dynamicParameters.Add("MB001", MtlItemNo);

                            var INVMBResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (INVMBResult.Count() <= 0) throw new SystemException("ERP品號基本資料錯誤!!");

                            foreach (var item2 in INVMBResult)
                            {
                                #region //判斷ERP品號生效日與失效日
                                if (item2.MB030 != "" && item2.MB030 != null)
                                {
                                    #region //判斷單據日期需大於或等於生效日
                                    string EffectiveDate = item2.MB030;
                                    string effYear = EffectiveDate.Substring(0, 4);
                                    string effMonth = EffectiveDate.Substring(4, 2);
                                    string effDay = EffectiveDate.Substring(6, 2);
                                    DateTime effFullDate = Convert.ToDateTime(effYear + "-" + effMonth + "-" + effDay);
                                    int effresult = DateTime.Compare((DateTime)goodsReceipts[0].DocDate, effFullDate);
                                    if (effresult < 0) throw new SystemException("此品號尚未生效，無法使用!!");
                                    #endregion
                                }

                                if (item2.MB031 != "" && item2.MB031 != null)
                                {
                                    #region //判斷日期需小於或等於失效日
                                    string ExpirationDate = item2.MB031;
                                    string effYear = ExpirationDate.Substring(0, 4);
                                    string effMonth = ExpirationDate.Substring(4, 2);
                                    string effDay = ExpirationDate.Substring(6, 2);
                                    DateTime effFullDate = Convert.ToDateTime(effYear + "-" + effMonth + "-" + effDay);
                                    int effresult = DateTime.Compare((DateTime)goodsReceipts[0].DocDate, effFullDate);
                                    if (effresult > 0) throw new SystemException("此品號已失效，無法使用!!");
                                    #endregion
                                }
                                #endregion

                                #region //確認此品號是否需要批號管理
                                if (item2.MB022 == "Y" || item2.MB022 == "T")
                                {
                                    if (LotNumber.Length <= 0) throw new SystemException("品號【" + MtlItemNo + "】需設定批號!!");
                                }
                                else
                                {
                                    if (LotNumber.Length > 0) throw new SystemException("品號【" + MtlItemNo + "】不需設定批號!!");
                                }
                                #endregion

                                #region //確認檢驗狀態
                                if (item2.MB043 == "0")
                                {
                                    if (QcStatus != "0") throw new SystemException("品號【" + MtlItemNo + "】設定為免檢，無法修改!!");
                                }
                                else
                                {
                                    if (QcStatus == "0") throw new SystemException("品號【" + MtlItemNo + "】非免檢類型，無法修改!!");

                                    #region //確認進貨檢狀態
                                    if (QcStatus == "2") //合格
                                    {
                                        dynamicParameters = new DynamicParameters();
                                        //單據狀態也需要為合格
                                    }
                                    else if (QcStatus == "3") //不良
                                    {
                                        //驗收數量需為0，驗退要等於進貨數量
                                    }
                                    else if (QcStatus == "4") //特採
                                    {
                                        //單據狀態也需要為特採
                                        //若驗收、驗退數量與品異單特採數量不一致，警告視窗
                                    }
                                    #endregion
                                }
                                #endregion

                                #region //檢核數量是否超收
                                if (item2.MB044 == "Y" && poDetails.Count() > 0) //卡控超收
                                {
                                    if (ReceiptQty > (Quantity - SiQty)) throw new SystemException("此次數量已超過剩餘可進貨數量(" + (Quantity - SiQty).ToString() + ")!!");
                                }
                                #endregion
                            }
                            #endregion

                            #region //查詢單身序號
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT ISNULL(MAX(a.GrSeq), 0) MaxGrSeq
                                    FROM SCM.GrDetail a 
                                    WHERE a.GrId = @GrId";
                            dynamicParameters.Add("GrId", GrId);

                            var GrSeqResult = sqlConnection.Query(sql, dynamicParameters);

                            int MaxGrSeq = 0;
                            foreach (var item2 in GrSeqResult)
                            {
                                MaxGrSeq = Convert.ToInt32(item2.MaxGrSeq);
                            }

                            string GrSeq = (MaxGrSeq + 1).ToString("D4");
                            #endregion

                            #region //查詢目前庫存數
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT ISNULL(MC007, 0) MC007
                                    FROM INVMC
                                    WHERE MC001 = @MC001
                                    AND MC002 = @MC002";
                            dynamicParameters.Add("MC001", MtlItemNo);
                            dynamicParameters.Add("MC002", InventoryNo);

                            var INVMCResult = sqlConnection2.Query(sql, dynamicParameters);

                            double InventoryQty = 0;
                            foreach (var item2 in INVMCResult)
                            {
                                InventoryQty = Convert.ToDouble(item2.MC007);
                            }
                            #endregion

                            #region //處理【原幣進貨金額】、【原幣未稅金額】、【原幣稅金額】、【本幣未稅金額】、【本幣稅金額】
                            double OrigPreTaxAmt = 0; //原幣未稅金額
                            double OrigTaxAmt = 0; //原幣稅金額
                            double PreTaxAmt = 0; //本幣未稅金額
                            double TaxAmt = 0; //本幣稅金額

                            JObject calculateTaxAmtResult = JObject.Parse(CalculateTaxAmt(AvailableQty, OrigUnitPrice, OrigDiscountAmt, OrigAmount, goodsReceipts[0].TaxRate, goodsReceipts[0].TaxType, goodsReceipts[0].Exchange
                                                                            , OriUnitDecimal, OriPriceDecimal, UnitDecimal, PriceDecimal));

                            if (calculateTaxAmtResult["status"].ToString() == "success")
                            {
                                OrigPreTaxAmt = Convert.ToDouble(calculateTaxAmtResult["OrigPreTaxAmt"]);
                                OrigTaxAmt = Convert.ToDouble(calculateTaxAmtResult["OrigTaxAmt"]);
                                PreTaxAmt = Convert.ToDouble(calculateTaxAmtResult["PreTaxAmt"]);
                                TaxAmt = Convert.ToDouble(calculateTaxAmtResult["TaxAmt"]);
                            }
                            else
                            {
                                throw new SystemException(calculateTaxAmtResult["msg"].ToString());
                            }
                            #endregion

                            #region //INSERT SCM.GrDetail
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO SCM.GrDetail (GrId, GrErpPrefix, GrErpNo, GrSeq, MtlItemId, MtlItemNo, GrMtlItemName, GrMtlItemSpec
                                    , ReceiptQty, UomId, UomNo, InventoryId, InventoryNo, LotNumber, PoErpPrefix, PoErpNo, PoSeq, AcceptanceDate
                                    , AcceptQty, AvailableQty, ReturnQty, OrigUnitPrice, OrigAmount, OrigDiscountAmt, TrErpPrefix, TrErpNo, TrSeq
                                    , ReceiptExpense, DiscountDescription, PaymentHold, Overdue, QcStatus, ReturnCode, ConfirmStatus, CloseStatus
                                    , ReNewStatus, Remark, InventoryQty, SmallUnit, AvailableDate, ReCheckDate, ConfirmUser, ApvErpPrefix, ApvErpNo
                                    , ApvSeq, ProjectCode, ExpenseEntry, PremiumAmountFlag, OrigPreTaxAmt, OrigTaxAmt, PreTaxAmt, TaxAmt, ReceiptPackageQty
                                    , AcceptancePackageQty, ReturnPackageQty, PremiumAmount, PackageUnit, ReserveTaxCode, OrigPremiumAmount, AvailableUomId
                                    , AvailableUomNo, OrigCustomer, ApproveStatus, EbcErpPreNo, EbcEdition, ProductSeqAmount, MtlItemType, Loaction
                                    , TaxRate, MultipleFlag, GrErpStatus, TaxCode, DiscountRate, DiscountPrice, QcType, TransferStatus, TransferDate
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.GrDetailId
                                    VALUES (@GrId, @GrErpPrefix, @GrErpNo, @GrSeq, @MtlItemId, @MtlItemNo, @GrMtlItemName, @GrMtlItemSpec
                                    , @ReceiptQty, @UomId, @UomNo, @InventoryId, @InventoryNo, @LotNumber, @PoErpPrefix, @PoErpNo, @PoSeq, @AcceptanceDate
                                    , @AcceptQty, @AvailableQty, @ReturnQty, @OrigUnitPrice, @OrigAmount, @OrigDiscountAmt, @TrErpPrefix, @TrErpNo, @TrSeq
                                    , @ReceiptExpense, @DiscountDescription, @PaymentHold, @Overdue, @QcStatus, @ReturnCode, @ConfirmStatus, @CloseStatus
                                    , @ReNewStatus, @Remark, @InventoryQty, @SmallUnit, @AvailableDate, @ReCheckDate, @ConfirmUser, @ApvErpPrefix, @ApvErpNo
                                    , @ApvSeq, @ProjectCode, @ExpenseEntry, @PremiumAmountFlag, @OrigPreTaxAmt, @OrigTaxAmt, @PreTaxAmt, @TaxAmt, @ReceiptPackageQty
                                    , @AcceptancePackageQty, @ReturnPackageQty, @PremiumAmount, @PackageUnit, @ReserveTaxCode, @OrigPremiumAmount, @AvailableUomId
                                    , @AvailableUomNo, @OrigCustomer, @ApproveStatus, @EbcErpPreNo, @EbcEdition, @ProductSeqAmount, @MtlItemType, @Loaction
                                    , @TaxRate, @MultipleFlag, @GrErpStatus, @TaxCode, @DiscountRate, @DiscountPrice, @QcType, @TransferStatus, @TransferDate
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    GrId,
                                    goodsReceipts[0].GrErpPrefix,
                                    goodsReceipts[0].GrErpNo,
                                    GrSeq,
                                    MtlItemId,
                                    MtlItemNo,
                                    GrMtlItemName = MtlItemName,
                                    GrMtlItemSpec = MtlItemSpec,
                                    ReceiptQty,
                                    UomId,
                                    UomNo,
                                    InventoryId,
                                    InventoryNo,
                                    LotNumber,
                                    PoErpPrefix,
                                    PoErpNo,
                                    PoSeq,
                                    AcceptanceDate,
                                    AcceptQty,
                                    AvailableQty,
                                    ReturnQty,
                                    OrigUnitPrice,
                                    OrigAmount,
                                    OrigDiscountAmt,
                                    TrErpPrefix = "",
                                    TrErpNo = "",
                                    TrSeq = "",
                                    ReceiptExpense,
                                    DiscountDescription,
                                    PaymentHold = "N",
                                    Overdue = "N",
                                    QcStatus,
                                    ReturnCode = "N",
                                    ConfirmStatus = "N",
                                    CloseStatus = "N",
                                    ReNewStatus = "N",
                                    Remark = PoRemark,
                                    InventoryQty,
                                    SmallUnit = "",
                                    AvailableDate = "",
                                    ReCheckDate = "",
                                    ConfirmUser = (int?)null,
                                    ApvErpPrefix = "",
                                    ApvErpNo = "",
                                    ApvSeq = "",
                                    ProjectCode = Project,
                                    ExpenseEntry = "N",
                                    PremiumAmountFlag = "N",
                                    OrigPreTaxAmt,
                                    OrigTaxAmt,
                                    PreTaxAmt,
                                    TaxAmt,
                                    ReceiptPackageQty = 0,
                                    AcceptancePackageQty = 0,
                                    ReturnPackageQty = 0,
                                    PremiumAmount = 0,
                                    PackageUnit = "",
                                    ReserveTaxCode = "N",
                                    OrigPremiumAmount = "",
                                    AvailableUomId = UomId,
                                    AvailableUomNo = UomNo,
                                    OrigCustomer = "",
                                    ApproveStatus = "N",
                                    EbcErpPreNo = "",
                                    EbcEdition = "",
                                    ProductSeqAmount = 0,
                                    MtlItemType = "2",
                                    Loaction = "",
                                    goodsReceipts[0].TaxRate,
                                    MultipleFlag = "N",
                                    GrErpStatus = "0",
                                    goodsReceipts[0].TaxCode,
                                    DiscountRate,
                                    DiscountPrice = DiscountAmount,
                                    QcType,
                                    TransferStatus = "N",
                                    TransferDate = (DateTime?)null,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();

                            int GrDetailId = -1;
                            foreach (var item2 in insertResult)
                            {
                                GrDetailId = item2.GrDetailId;
                            }
                            #endregion

                            #region //若品項為待驗品，自動新增進貨檢單據
                            if (QcStatus == "1")
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

                                #region //取得進貨檢QcType
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 QcTypeId
                                            FROM QMS.QcType
                                            WHERE QcTypeNo = 'IQC'";

                                var QcTypeResult = sqlConnection.Query(sql, dynamicParameters);

                                int QcTypeId = -1;
                                foreach (var item2 in QcTypeResult)
                                {
                                    QcTypeId = item2.QcTypeId;
                                }
                                #endregion

                                #region //INSERT MES.QcRecord
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO MES.QcRecord (QcTypeId, InputType, Remark, DefaultFileId, DefaultSpreadsheetData, ResolveFileJson
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.QcRecordId
                                        VALUES (@QcTypeId, @InputType, @Remark, @DefaultFileId, @DefaultSpreadsheetData, @ResolveFileJson
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        QcTypeId,
                                        InputType = "LotNumber",
                                        Remark = "由進貨單自動建立",
                                        DefaultFileId = -1,
                                        DefaultSpreadsheetData,
                                        ResolveFileJson = "",
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });

                                var insertQcRecordResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertQcRecordResult.Count();

                                int QcRecordId = -1;
                                foreach (var item2 in insertQcRecordResult)
                                {
                                    QcRecordId = item2.QcRecordId;
                                }
                                #endregion

                                #region //INSERT MES.QcGoodsReceipt
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO MES.QcGoodsReceipt (QcRecordId, GrDetailId
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.QcGoodsReceiptId
                                        VALUES (@QcRecordId, @GrDetailId
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        QcRecordId,
                                        GrDetailId,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region //建立量測單據Spreadsheet DATA
                                #region //取得品號允收標準
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.MtlQcItemId, a.MtlItemId, a.QcItemId, a.DesignValue, a.UpperTolerance, a.LowerTolerance
                                        , a.QcItemDesc, a.SortNumber, a.QmmDetailId
                                        , b.QcItemNo, b.QcItemName, b.QcItemDesc OriQcItemDesc
                                        , c.MachineNumber
                                        , d.MachineDesc
                                        FROM PDM.MtlQcItem a
                                        INNER JOIN QMS.QcItem b ON a.QcItemId = b.QcItemId
                                        LEFT JOIN QMS.QmmDetail c ON a.QmmDetailId = c.QmmDetailId
                                        LEFT JOIN MES.Machine d ON c.MachineId = d.MachineId
                                        WHERE a.MtlItemId = @MtlItemId
                                        AND a.[Status] = 'A' 
                                        ORDER BY a.SortNumber";
                                dynamicParameters.Add("MtlItemId", MtlItemId);

                                var MtlQcItemResult = sqlConnection.Query(sql, dynamicParameters);
                                #endregion

                                #region //設定單身量測標準
                                int row = 2;
                                foreach (var item2 in MtlQcItemResult)
                                {
                                    #region //若有機台，整理序號格式
                                    string QcItemNo = item2.QcItemNo;
                                    if (item2.MachineNumber != null)
                                    {
                                        QcItemNo = item2.QcItemNo;
                                        string firstPart = QcItemNo.Substring(0, 3);
                                        string secondPart = QcItemNo.Substring(3, 4);
                                        QcItemNo = firstPart + item2.MachineNumber + secondPart;
                                    }
                                    #endregion
                                    
                                    string QcItemName = item2.QcItemName;

                                    #region //設定量測項目、備註、上下限
                                    data = new Data()
                                    {
                                        cell = "A" + row,
                                        css = "",
                                        format = "common",
                                        value = QcItemNo,
                                    };
                                    datas.Add(data);

                                    data = new Data()
                                    {
                                        cell = "B" + row,
                                        css = "",
                                        format = "common",
                                        value = QcItemName,
                                    };
                                    datas.Add(data);

                                    data = new Data()
                                    {
                                        cell = "C" + row,
                                        css = "",
                                        format = "common",
                                        value = item2.QcItemDesc,
                                    };
                                    datas.Add(data);

                                    data = new Data()
                                    {
                                        cell = "D" + row,
                                        css = "",
                                        format = "common",
                                        value = item2.MachineDesc != null ? item2.MachineDesc : "",
                                    };
                                    datas.Add(data);

                                    data = new Data()
                                    {
                                        cell = "E" + row,
                                        css = "",
                                        format = "common",
                                        value = item2.DesignValue != null ? item2.DesignValue.ToString() : "",
                                    };
                                    datas.Add(data);

                                    data = new Data()
                                    {
                                        cell = "F" + row,
                                        css = "",
                                        format = "common",
                                        value = item2.UpperTolerance != null ? item2.UpperTolerance.ToString() : "",
                                    };
                                    datas.Add(data);

                                    data = new Data()
                                    {
                                        cell = "G" + row,
                                        css = "",
                                        format = "common",
                                        value = item2.LowerTolerance != null ? item2.LowerTolerance.ToString() : "",
                                    };
                                    datas.Add(data);

                                    row++;
                                    #endregion
                                }
                                #endregion

                                #region //處理批號部分
                                int startIndex = 11;
                                if (LotNumber.Length > 0)
                                {
                                    List<string> LotNumberList = new List<string>();
                                    LotNumberList.Add(LotNumber);

                                    #region //綁定SpreadsheetData
                                    for (var i = 0; i < LotNumberList.Count(); i++)
                                    {
                                        string cell = Convert.ToChar(64 + startIndex + i).ToString();

                                        data = new Data()
                                        {
                                            cell = cell + "1",
                                            css = "",
                                            format = "common",
                                            value = LotNumberList[i],
                                        };
                                        datas.Add(data);
                                    }
                                    #endregion
                                }
                                #endregion

                                #region //整合Spreadsheet格式
                                Sheets sheets = new Sheets()
                                {
                                    name = "sheet1",
                                    data = datas
                                };
                                sheetss.Add(sheets);

                                SpreadsheetModel spreadsheetModel = new SpreadsheetModel()
                                {
                                    sheets = sheetss
                                };

                                string SpreadsheetData = JsonConvert.SerializeObject(spreadsheetModel);

                                #region //更新至QcRecord SpreadsheetData
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE MES.QcRecord SET
                                            SpreadsheetData = @SpreadsheetData,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE QcRecordId = @QcRecordId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        SpreadsheetData,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        QcRecordId
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                                #endregion
                                #endregion
                            }
                            #endregion

                            #region //重新計算進貨單資料
                            JObject updateGrDetailResult = JObject.Parse(UpdateGrTotal(GrId, goodsReceipts[0].Exchange, sqlConnection, sqlConnection2));
                            if (updateGrDetailResult["status"].ToString() != "success")
                            {
                                throw new SystemException(updateGrDetailResult["msg"].ToString());
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

        #region //ImportDeliveryOrder -- 匯入供應商出貨單 -- Ann 2025-05-13
        public string ImportDeliveryOrder(int GrId, string DoNo, string CopyExchange)
        {
            try
            {
                if (GrId <= 0) throw new SystemException("【進貨單】不能為空!!");
                if (DoNo.Length <= 0) throw new SystemException("【供應商出貨單】不能為空!!");
                if (CopyExchange.Length <= 0) throw new SystemException("【匯率複製設定】不能為空!!");

                string companyNo = "";
                int rowsAffected = 0;

                List<SpreadsheetModel> spreadsheetModels = new List<SpreadsheetModel>();
                List<Sheets> sheetss = new List<Sheets>();
                List<Data> datas = new List<Data>();
                Dictionary<string, string> BarcodeInfo = new Dictionary<string, string>();
                #region //初始化Data
                Data data = new Data();
                data = new Data()
                {
                    cell = "A1",
                    css = "imported_class1",
                    format = "common",
                    value = "序號",
                };
                datas.Add(data);

                data = new Data()
                {
                    cell = "B1",
                    css = "imported_class2",
                    format = "common",
                    value = "檢測項目",
                };
                datas.Add(data);

                data = new Data()
                {
                    cell = "C1",
                    css = "imported_class3",
                    format = "common",
                    value = "檢測備註",
                };
                datas.Add(data);

                data = new Data()
                {
                    cell = "D1",
                    css = "imported_class4",
                    format = "common",
                    value = "量測設備",
                };
                datas.Add(data);

                data = new Data()
                {
                    cell = "E1",
                    css = "imported_class5",
                    format = "common",
                    value = "設計值",
                };
                datas.Add(data);

                data = new Data()
                {
                    cell = "F1",
                    css = "imported_class6",
                    format = "common",
                    value = "上限",
                };
                datas.Add(data);

                data = new Data()
                {
                    cell = "G1",
                    css = "imported_class7",
                    format = "common",
                    value = "下限",
                };
                datas.Add(data);

                data = new Data()
                {
                    cell = "H1",
                    css = "imported_class8",
                    format = "common",
                    value = "Z軸",
                };
                datas.Add(data);

                data = new Data()
                {
                    cell = "I1",
                    css = "imported_class9",
                    format = "common",
                    value = "量測人員",
                };
                datas.Add(data);

                data = new Data()
                {
                    cell = "J1",
                    css = "imported_class10",
                    format = "common",
                    value = "量測工時",
                };
                datas.Add(data);
                #endregion

                DeliveryOrder deliveryOrder = new DeliveryOrder();
                List<DoDetail> doDetails = new List<DoDetail>();

                GoodsReceipt goodsReceipt = new GoodsReceipt();

                string currentCompanyNo = "";

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //公司別資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CompanyNo, a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var resultCompany = sqlConnection.Query(sql, dynamicParameters);
                        if (resultCompany.Count() <= 0) throw new SystemException("【公司別】資料錯誤!");

                        foreach (var item in resultCompany)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            companyNo = item.ErpNo;
                            currentCompanyNo = item.CompanyNo;
                        }
                        #endregion

                        #region //確認此進貨單身是否有其他筆資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.GrDetail a 
                                WHERE a.GrId = @GrId";
                        dynamicParameters.Add("GrId", GrId);

                        var GrDetailResult = sqlConnection.Query(sql, dynamicParameters);

                        if (GrDetailResult.Count() > 0) throw new SystemException("此進貨單已有單身資料，無法使用複製前置單據功能!!");
                        #endregion

                        using (SqlConnection srmConnection = new SqlConnection(SrmDbConnectionStrings))
                        {
                            #region //取得供應商出貨單頭資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.DoId, a.CompanyId, a.SupplierGroupId, a.DoNo, a.PoErpPrefix, a.PoErpNo
                                    , a.DocDate, a.DoAddressFirst, a.DoAddressSecond, a.InvoiceType, a.InvoiceNo
                                    , a.InvoiceDate, a.PaymentType, a.PaymentTerm, a.ShippingFree, a.TaxAmount
                                    , a.TotalAmount, a.DoUser, a.Remark, a.ConfirmStatus, a.ConfirmUser, a.ConfirmDate
                                    , a.ClosureStatus, a.DocStatus
                                    , c.SupplierNo
                                    FROM SCM.DeliveryOrder a 
                                    INNER JOIN BAS.SupplierGroup b ON a.SupplierGroupId = b.SupplierGroupId
                                    INNER JOIN BAS.SupplierCompany c ON b.SupplierGroupId = c.SupplierGroupId
                                    WHERE a.DoNo = @DoNo
                                    AND c.CompanyNo = @CompanyNo";
                            dynamicParameters.Add("DoNo", DoNo);
                            dynamicParameters.Add("CompanyNo", currentCompanyNo);

                            deliveryOrder = srmConnection.QueryFirstOrDefault<DeliveryOrder>(sql, dynamicParameters);

                            if (deliveryOrder == null)
                            {
                                throw new SystemException("查無此出貨單資料!!");
                            }
                            #endregion

                            #region //取得供應商出貨單身資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.DoDetailId, a.DoId, a.DoSequence, a.DpId, a.BatchId, a.Quantity, a.FreebieQty, a.SpareQty, a.UnitPrice, a.Tax, a.Price, a.Unit, a.Remark 
                                    , a.PoErpPrefix, a.PoErpNo, a.PoSeq, a.DoItemNo, a.DoItemName, a.DoItemSpec, a.DoStatus, a.CreateDate, a.LastModifiedDate
                                    , a.CreateBy, a.LastModifiedBy
                                    FROM SCM.DoDetail a
                                    WHERE a.DoId = @DoId";
                            dynamicParameters.Add("DoId", deliveryOrder.DoId);

                            doDetails = srmConnection.Query<DoDetail>(sql, dynamicParameters).ToList();

                            if (doDetails.Count() <= 0)
                            {
                                throw new SystemException("查無出貨單身資料!!");
                            }
                            #endregion
                        }

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //確認進貨單資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.GrId, a.CompanyId, a.GrErpPrefix, a.GrErpNo, GrErpPrefix + '-' + GrErpNo GrErpFullNo, a.ReceiptDate, a.SupplierId
                                    , a.SupplierSo, a.CurrencyCode, a.Exchange, a.InvoiceType, a.TaxType, a.InvoiceNo, a.PrintCnt, a.ConfirmStatus, a.DocDate
                                    , a.RenewFlag, a.Remark, a.TotalAmount, a.DeductAmount, a.OrigTax, a.ReceiptAmount, a.SupplierName, a.UiNo, a.DeductType
                                    , a.ObaccoAndLiquorFlag, a.RowCnt, a.Quantity, FORMAT(a.InvoiceDate, 'yyyy-MM-dd') InvoiceDate, a.OrigPreTaxAmount, a.ApplyYYMM
                                    , a.TaxRate, a.PreTaxAmount, a.TaxAmount, a.PaymentTerm, a.PoId, a.PrepaidErpPrefix, a.PrepaidErpNo, a.OffsetAmount, a.TaxOffset, a.PackageQuantity
                                    , a.PremiumAmount, a.ApproveStatus, a.InvoiceFlag, a.ReserveTaxCode, a.SendCount, a.OrigPremiumAmount, a.EbcErpPreNo, a.EbcEdition, a.EbcFlag
                                    , a.FromDocType, a.NoticeFlag, a.TaxCode, a.TradeTerm, a.DetailMultiTax, a.TaxExchange, a.ContactUser, a.TransferStatus, a.TransferDate
                                    , a.CreateDate, a.LastModifiedDate, a.CreateBy, a.LastModifiedBy
                                    , b.SupplierNo
                                    FROM SCM.GoodsReceipt a 
                                    INNER JOIN SCM.Supplier b ON a.SupplierId = b.SupplierId
                                    WHERE a.GrId = @GrId";
                            dynamicParameters.Add("GrId", GrId);

                            goodsReceipt = sqlConnection.QueryFirstOrDefault<GoodsReceipt>(sql, dynamicParameters);

                            if (goodsReceipt == null) throw new SystemException("進貨單資料錯誤!!");

                            if (goodsReceipt.SupplierNo != deliveryOrder.SupplierNo)
                            {
                                throw new SystemException("進貨單與出貨單供應商不同!!");
                            }
                            #endregion

                            foreach (var doDetail in doDetails)
                            {
                                sheetss = new List<Sheets>();

                                #region //取得採購單身資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.PoDetailId, a.PoId, a.PoErpPrefix, a.PoErpNo, a.PoSeq, a.MtlItemId, a.PoMtlItemName, a.PoMtlItemSpec, a.InventoryId, a.Quantity
                                        , a.UomId, a.PoUnitPrice, a.PoPrice, a.PromiseDate, a.ReferencePrefix, a.Remark, a.SiQty, a.ClosureStatus, a.ConfirmStatus 
                                        , a.InventoryQty, a.SmallUnit, a.ReferenceNo, a.Project, a.ReferenceSeq, a.UrgentMtl, a.PrErpPrefix, a.PrErpNo, a.PrSequence 
                                        , a.FromDocType, a.MtlItemType, a.PartialSeq, a.OriPromiseDate, a.DeliveryDate, a.PoPriceQty, a.PoPriceUomId, a.SiPriceQty 
                                        , a.DiscountRate, a.DiscountAmount, a.CreateDate, a.LastModifiedDate, a.CreateBy, a.LastModifiedBy
                                        , b.UomNo
                                        , c.InventoryNo
                                        , d.MtlItemNo
                                        , e.TaxNo, e.CurrencyCode, e.Exchange, e.PaymentTermNo, e.Taxation, e.TaxRate, e.TaxNo
                                        FROM SCM.PoDetail a
                                        INNER JOIN PDM.UnitOfMeasure b ON a.UomId = b.UomId
                                        INNER JOIN SCM.Inventory c ON a.InventoryId = c.InventoryId
                                        INNER JOIN PDM.MtlItem d ON a.MtlItemId = d.MtlItemId
                                        INNER JOIN SCM.PurchaseOrder e ON a.PoId = e.PoId
                                        WHERE a.PoErpPrefix = @PoErpPrefix
                                        AND a.PoErpNo = @PoErpNo
                                        AND a.PoSeq = @PoSeq";
                                dynamicParameters.Add("PoErpPrefix", doDetail.PoErpPrefix);
                                dynamicParameters.Add("PoErpNo", doDetail.PoErpNo);
                                dynamicParameters.Add("PoSeq", doDetail.PoSeq);

                                PoDetail poDetail = sqlConnection.QueryFirstOrDefault<PoDetail>(sql, dynamicParameters);

                                if (poDetail == null) throw new SystemException($"查無此採購單資料【{doDetail.PoErpPrefix}-{doDetail.PoErpNo}({doDetail.PoSeq})】!!");

                                goodsReceipt.CurrencyCode = poDetail.CurrencyCode;
                                goodsReceipt.Exchange = CopyExchange == "Y" ? poDetail.Exchange : goodsReceipt.Exchange;
                                goodsReceipt.PaymentTerm = poDetail.PaymentTermNo;
                                goodsReceipt.TaxCode = poDetail.TaxNo;
                                goodsReceipt.TaxType = poDetail.Taxation;
                                goodsReceipt.TaxRate = poDetail.TaxRate;

                                if (goodsReceipt.TaxType == "1" || goodsReceipt.TaxType == "2")
                                {
                                    goodsReceipt.DeductType = "1";
                                }
                                else
                                {
                                    goodsReceipt.DeductType = "3";
                                }
                                #endregion

                                #region //重新取得發票設定資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT LTRIM(RTRIM(NM001)) InvoiceCountNo, LTRIM(RTRIM(NM002)) InvoiceCountName 
                                        FROM CMSNM a
                                        INNER JOIN CMSNN b ON a.NM001 = b.NN005
                                        WHERE b.NN001 = @TaxCode";
                                dynamicParameters.Add("TaxCode", goodsReceipt.TaxCode);

                                var CMSNMResult = sqlConnection2.Query(sql, dynamicParameters);

                                if (CMSNMResult.Count() <= 0) throw new SystemException($"查無稅別碼【{goodsReceipt.TaxCode}】的發票設定資料!!");

                                foreach (var item in CMSNMResult)
                                {
                                    goodsReceipt.InvoiceType = item.InvoiceCountNo;
                                }
                                #endregion

                                #region //原幣小數點取位
                                int OriUnitDecimal = 0;
                                int OriPriceDecimal = 0;
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.MF003, a.MF004
                                        FROM CMSMF a 
                                        WHERE a.MF001 = @Currency";
                                dynamicParameters.Add("Currency", poDetail.CurrencyCode);

                                var CMSMFResult = sqlConnection2.Query(sql, dynamicParameters);

                                if (CMSMFResult.Count() <= 0) throw new SystemException("金額小數點取位資料錯誤!!");

                                foreach (var item in CMSMFResult)
                                {
                                    OriUnitDecimal = Convert.ToInt32(item.MF003);
                                    OriPriceDecimal = Convert.ToInt32(item.MF004);
                                }
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

                                #region //本幣小數點取位
                                int UnitDecimal = 0;
                                int PriceDecimal = 0;
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.MF003, a.MF004
                                        FROM CMSMF a 
                                        WHERE a.MF001 = @Currency";
                                dynamicParameters.Add("Currency", CurrencyCode);

                                var CMSMFResult2 = sqlConnection2.Query(sql, dynamicParameters);

                                if (CMSMFResult2.Count() <= 0) throw new SystemException("金額小數點取位資料錯誤!!");

                                foreach (var item in CMSMFResult2)
                                {
                                    UnitDecimal = Convert.ToInt32(item.MF003);
                                    PriceDecimal = Convert.ToInt32(item.MF004);
                                }
                                #endregion

                                #region //INSERT SCM.GrDetail
                                #region //確認結案碼狀態
                                if (poDetail.ClosureStatus != "N") throw new SystemException("採購單單身【" + poDetail.PoErpPrefix + "-" + poDetail.PoErpNo + "(" + poDetail.PoSeq + ")】已結案!!");
                                #endregion

                                #region //確認ERP採購單狀態是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT LTRIM(RTRIM(TC014)) TC014
                                        FROM PURTC 
                                        WHERE TC001 = @PoErpPrefix
                                        AND TC002 = @PoErpNo";
                                dynamicParameters.Add("PoErpPrefix", poDetail.PoErpPrefix);
                                dynamicParameters.Add("PoErpNo", poDetail.PoErpNo);

                                var PURTCResult = sqlConnection2.Query(sql, dynamicParameters);

                                if (PURTCResult.Count() <= 0) throw new SystemException("ERP採購單資料錯誤!!");

                                foreach (var item2 in PURTCResult)
                                {
                                    if (item2.TC014 != "Y") throw new SystemException("ERP採購單【" + poDetail.PoErpPrefix + "-" + poDetail.PoErpNo + "】狀態錯誤!!");
                                }
                                #endregion

                                #region //確認ERP採購單身結案碼
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT LTRIM(RTRIM(TD016)) TD016
                                        , LTRIM(RTRIM(TD018)) TD018
                                        FROM PURTD
                                        WHERE TD001 = @PoErpPrefix
                                        AND TD002 = @PoErpNo
                                        AND TD003 = @PoSeq";
                                dynamicParameters.Add("PoErpPrefix", poDetail.PoErpPrefix);
                                dynamicParameters.Add("PoErpNo", poDetail.PoErpNo);
                                dynamicParameters.Add("PoSeq", poDetail.PoSeq);

                                var PURTDResult = sqlConnection2.Query(sql, dynamicParameters);

                                if (PURTDResult.Count() <= 0) throw new SystemException("ERP採購單單身【" + poDetail.PoErpPrefix + "-" + poDetail.PoErpNo + "(" + poDetail.PoSeq + ")】資料錯誤!!");

                                foreach (var item2 in PURTDResult)
                                {
                                    if (item2.TD016 != "N") throw new SystemException("ERP採購單單身【" + poDetail.PoErpPrefix + "-" + poDetail.PoErpNo + "(" + poDetail.PoSeq + ")】已結案!!");
                                    if (item2.TD018 != "Y") throw new SystemException("ERP採購單單身【" + poDetail.PoErpPrefix + "-" + poDetail.PoErpNo + "(" + poDetail.PoSeq + ")】狀態錯誤!!");
                                }
                                #endregion

                                #region //查詢單身序號
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT ISNULL(MAX(a.GrSeq), 0) MaxGrSeq
                                        FROM SCM.GrDetail a 
                                        WHERE a.GrId = @GrId";
                                dynamicParameters.Add("GrId", GrId);

                                var GrSeqResult = sqlConnection.Query(sql, dynamicParameters);

                                int MaxGrSeq = 0;
                                foreach (var item2 in GrSeqResult)
                                {
                                    MaxGrSeq = Convert.ToInt32(item2.MaxGrSeq);
                                }

                                string GrSeq = (MaxGrSeq + 1).ToString("D4");
                                #endregion

                                #region //確認品號檢驗狀態
                                string QcStatus = "";
                                string QcType = "";
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT MB043
                                        FROM INVMB
                                        WHERE MB001 = @MB001";
                                dynamicParameters.Add("MB001", poDetail.MtlItemNo);

                                var INVMBResult = sqlConnection2.Query(sql, dynamicParameters);

                                if (INVMBResult.Count() <= 0) throw new SystemException("此進貨品號【" + poDetail.MtlItemNo + "】資料錯誤!!");

                                foreach (var item2 in INVMBResult)
                                {
                                    if (item2.MB043 == "0")
                                    {
                                        QcStatus = "0";
                                        QcType = "N";
                                    }
                                    else
                                    {
                                        QcStatus = "1";
                                        QcType = "1";
                                    }
                                }
                                #endregion

                                #region //查詢目前庫存數
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT ISNULL(MC007, 0) MC007
                                        FROM INVMC
                                        WHERE MC001 = @MC001
                                        AND MC002 = @MC002";
                                dynamicParameters.Add("MC001", poDetail.MtlItemNo);
                                dynamicParameters.Add("MC002", poDetail.InventoryNo);

                                var INVMCResult = sqlConnection2.Query(sql, dynamicParameters);

                                double InventoryQty = 0;
                                foreach (var item2 in INVMCResult)
                                {
                                    InventoryQty = Convert.ToDouble(item2.MC007);
                                }
                                #endregion

                                #region //處理【原幣進貨金額】、【原幣未稅金額】、【原幣稅金額】、【本幣未稅金額】、【本幣稅金額】
                                double? AvailableQty = QcStatus == "0" ? poDetail.PoPriceQty - poDetail.SiPriceQty : 0; //計價數量
                                double? OrigUnitPrice = Math.Round(Convert.ToDouble(poDetail.PoUnitPrice), OriUnitDecimal); //原幣進貨單價
                                double? OrigDiscountAmt = 0; //原幣扣款金額
                                double? OrigAmount = Math.Round(Convert.ToDouble(AvailableQty * OrigUnitPrice), OriPriceDecimal); //原幣進貨金額

                                double OrigPreTaxAmt = 0; //原幣未稅金額
                                double OrigTaxAmt = 0; //原幣稅金額
                                double PreTaxAmt = 0; //本幣未稅金額
                                double TaxAmt = 0; //本幣稅金額

                                JObject calculateTaxAmtResult = JObject.Parse(CalculateTaxAmt(AvailableQty, OrigUnitPrice, OrigDiscountAmt, OrigAmount, goodsReceipt.TaxRate, goodsReceipt.TaxType, goodsReceipt.Exchange
                                                                            , OriUnitDecimal, OriPriceDecimal, UnitDecimal, PriceDecimal));

                                if (calculateTaxAmtResult["status"].ToString() == "success")
                                {
                                    OrigPreTaxAmt = Convert.ToDouble(calculateTaxAmtResult["OrigPreTaxAmt"]);
                                    OrigTaxAmt = Convert.ToDouble(calculateTaxAmtResult["OrigTaxAmt"]);
                                    PreTaxAmt = Convert.ToDouble(calculateTaxAmtResult["PreTaxAmt"]);
                                    TaxAmt = Convert.ToDouble(calculateTaxAmtResult["TaxAmt"]);
                                }
                                else
                                {
                                    throw new SystemException(calculateTaxAmtResult["msg"].ToString());
                                }
                                #endregion

                                #region //確認此進貨單身無相同採購單身
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT ISNULL(SUM(a.ReceiptQty), 0) TotalReceiptQty
                                        FROM SCM.GrDetail a 
                                        WHERE a.PoErpPrefix = @PoErpPrefix
                                        AND a.PoErpNo = @PoErpNo
                                        AND a.PoSeq = @PoSeq
                                        AND a.GrId = @GrId";
                                dynamicParameters.Add("PoErpPrefix", poDetail.PoErpPrefix);
                                dynamicParameters.Add("PoErpNo", poDetail.PoErpNo);
                                dynamicParameters.Add("PoSeq", poDetail.PoSeq);
                                dynamicParameters.Add("GrId", GrId);

                                var PoDetailResult = sqlConnection.Query(sql, dynamicParameters);

                                double? TotalReceiptQty = 0;
                                foreach (var item2 in PoDetailResult)
                                {
                                    TotalReceiptQty = item2.TotalReceiptQty;
                                }
                                #endregion

                                #region //INSERT SCM.GrDetail
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO SCM.GrDetail (GrId, GrErpPrefix, GrErpNo, GrSeq, MtlItemId, MtlItemNo, GrMtlItemName, GrMtlItemSpec
                                        , ReceiptQty, UomId, UomNo, InventoryId, InventoryNo, LotNumber, PoErpPrefix, PoErpNo, PoSeq, AcceptanceDate
                                        , AcceptQty, AvailableQty, ReturnQty, OrigUnitPrice, OrigAmount, OrigDiscountAmt, TrErpPrefix, TrErpNo, TrSeq
                                        , ReceiptExpense, DiscountDescription, PaymentHold, Overdue, QcStatus, ReturnCode, ConfirmStatus, CloseStatus
                                        , ReNewStatus, Remark, InventoryQty, SmallUnit, AvailableDate, ReCheckDate, ConfirmUser, ApvErpPrefix, ApvErpNo
                                        , ApvSeq, ProjectCode, ExpenseEntry, PremiumAmountFlag, OrigPreTaxAmt, OrigTaxAmt, PreTaxAmt, TaxAmt, ReceiptPackageQty
                                        , AcceptancePackageQty, ReturnPackageQty, PremiumAmount, PackageUnit, ReserveTaxCode, OrigPremiumAmount, AvailableUomId
                                        , AvailableUomNo, OrigCustomer, ApproveStatus, EbcErpPreNo, EbcEdition, ProductSeqAmount, MtlItemType, Loaction
                                        , TaxRate, MultipleFlag, GrErpStatus, TaxCode, DiscountRate, DiscountPrice, QcType, TransferStatus, TransferDate
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.GrDetailId
                                        VALUES (@GrId, @GrErpPrefix, @GrErpNo, @GrSeq, @MtlItemId, @MtlItemNo, @GrMtlItemName, @GrMtlItemSpec
                                        , @ReceiptQty, @UomId, @UomNo, @InventoryId, @InventoryNo, @LotNumber, @PoErpPrefix, @PoErpNo, @PoSeq, @AcceptanceDate
                                        , @AcceptQty, @AvailableQty, @ReturnQty, @OrigUnitPrice, @OrigAmount, @OrigDiscountAmt, @TrErpPrefix, @TrErpNo, @TrSeq
                                        , @ReceiptExpense, @DiscountDescription, @PaymentHold, @Overdue, @QcStatus, @ReturnCode, @ConfirmStatus, @CloseStatus
                                        , @ReNewStatus, @Remark, @InventoryQty, @SmallUnit, @AvailableDate, @ReCheckDate, @ConfirmUser, @ApvErpPrefix, @ApvErpNo
                                        , @ApvSeq, @ProjectCode, @ExpenseEntry, @PremiumAmountFlag, @OrigPreTaxAmt, @OrigTaxAmt, @PreTaxAmt, @TaxAmt, @ReceiptPackageQty
                                        , @AcceptancePackageQty, @ReturnPackageQty, @PremiumAmount, @PackageUnit, @ReserveTaxCode, @OrigPremiumAmount, @AvailableUomId
                                        , @AvailableUomNo, @OrigCustomer, @ApproveStatus, @EbcErpPreNo, @EbcEdition, @ProductSeqAmount, @MtlItemType, @Loaction
                                        , @TaxRate, @MultipleFlag, @GrErpStatus, @TaxCode, @DiscountRate, @DiscountPrice, @QcType, @TransferStatus, @TransferDate
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        GrId,
                                        goodsReceipt.GrErpPrefix,
                                        goodsReceipt.GrErpNo,
                                        GrSeq,
                                        poDetail.MtlItemId,
                                        poDetail.MtlItemNo,
                                        GrMtlItemName = poDetail.PoMtlItemName,
                                        GrMtlItemSpec = poDetail.PoMtlItemSpec,
                                        ReceiptQty = poDetail.Quantity - poDetail.SiQty - TotalReceiptQty > 0 ? poDetail.Quantity - poDetail.SiQty - TotalReceiptQty : 0,
                                        poDetail.UomId,
                                        poDetail.UomNo,
                                        poDetail.InventoryId,
                                        poDetail.InventoryNo,
                                        LotNumber = "",
                                        poDetail.PoErpPrefix,
                                        poDetail.PoErpNo,
                                        poDetail.PoSeq,
                                        AcceptanceDate = DateTime.Now.ToString("yyyyMMdd"),
                                        AcceptQty = QcStatus != "0" ? 0 : poDetail.Quantity - poDetail.SiQty,
                                        AvailableQty,
                                        ReturnQty = 0,
                                        OrigUnitPrice,
                                        OrigAmount,
                                        OrigDiscountAmt,
                                        TrErpPrefix = "",
                                        TrErpNo = "",
                                        TrSeq = "",
                                        ReceiptExpense = 0,
                                        DiscountDescription = "",
                                        PaymentHold = "N",
                                        Overdue = "N",
                                        QcStatus,
                                        ReturnCode = "N",
                                        ConfirmStatus = "N",
                                        CloseStatus = "N",
                                        ReNewStatus = "N",
                                        poDetail.Remark,
                                        InventoryQty,
                                        SmallUnit = "",
                                        AvailableDate = "",
                                        ReCheckDate = "",
                                        ConfirmUser = (int?)null,
                                        ApvErpPrefix = "",
                                        ApvErpNo = "",
                                        ApvSeq = "",
                                        ProjectCode = poDetail.Project,
                                        ExpenseEntry = "N",
                                        PremiumAmountFlag = "N",
                                        OrigPreTaxAmt,
                                        OrigTaxAmt,
                                        PreTaxAmt,
                                        TaxAmt,
                                        ReceiptPackageQty = 0,
                                        AcceptancePackageQty = 0,
                                        ReturnPackageQty = 0,
                                        PremiumAmount = 0,
                                        PackageUnit = "",
                                        ReserveTaxCode = "N",
                                        OrigPremiumAmount = "",
                                        AvailableUomId = poDetail.UomId,
                                        AvailableUomNo = poDetail.UomNo,
                                        OrigCustomer = "",
                                        ApproveStatus = "N",
                                        EbcErpPreNo = "",
                                        EbcEdition = "",
                                        ProductSeqAmount = 0,
                                        MtlItemType = "2",
                                        Loaction = "",
                                        goodsReceipt.TaxRate,
                                        MultipleFlag = "N",
                                        GrErpStatus = "0",
                                        goodsReceipt.TaxCode,
                                        poDetail.DiscountRate,
                                        DiscountPrice = poDetail.DiscountAmount,
                                        QcType,
                                        TransferStatus = "N",
                                        TransferDate = (DateTime?)null,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();

                                int GrDetailId = -1;
                                foreach (var item2 in insertResult)
                                {
                                    GrDetailId = item2.GrDetailId;
                                }
                                #endregion

                                #region //若品項為待驗品，自動新增進貨檢單據
                                if (QcStatus == "1")
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

                                    #region //取得進貨檢QcType
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 QcTypeId
                                            FROM QMS.QcType
                                            WHERE QcTypeNo = 'IQC'";

                                    var QcTypeResult = sqlConnection.Query(sql, dynamicParameters);

                                    int QcTypeId = -1;
                                    foreach (var item2 in QcTypeResult)
                                    {
                                        QcTypeId = item2.QcTypeId;
                                    }
                                    #endregion

                                    #region //INSERT MES.QcRecord
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO MES.QcRecord (QcTypeId, InputType, Remark, DefaultFileId, DefaultSpreadsheetData, ResolveFileJson
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.QcRecordId
                                            VALUES (@QcTypeId, @InputType, @Remark, @DefaultFileId, @DefaultSpreadsheetData, @ResolveFileJson
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            QcTypeId,
                                            InputType = "LotNumber",
                                            Remark = "由進貨單自動建立",
                                            DefaultFileId = -1,
                                            DefaultSpreadsheetData,
                                            ResolveFileJson = "",
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    var insertQcRecordResult = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertQcRecordResult.Count();

                                    int QcRecordId = -1;
                                    foreach (var item2 in insertQcRecordResult)
                                    {
                                        QcRecordId = item2.QcRecordId;
                                    }
                                    #endregion

                                    #region //INSERT MES.QcGoodsReceipt
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO MES.QcGoodsReceipt (QcRecordId, GrDetailId
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.QcGoodsReceiptId
                                            VALUES (@QcRecordId, @GrDetailId
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            QcRecordId,
                                            GrDetailId,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //建立量測單據Spreadsheet DATA
                                    #region //取得品號允收標準
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.MtlQcItemId, a.MtlItemId, a.QcItemId, a.DesignValue, a.UpperTolerance, a.LowerTolerance
                                            , a.QcItemDesc, a.SortNumber, a.QmmDetailId
                                            , b.QcItemNo, b.QcItemName, b.QcItemDesc OriQcItemDesc
                                            , c.MachineNumber
                                            , d.MachineDesc
                                            FROM PDM.MtlQcItem a
                                            INNER JOIN QMS.QcItem b ON a.QcItemId = b.QcItemId
                                            LEFT JOIN QMS.QmmDetail c ON a.QmmDetailId = c.QmmDetailId
                                            LEFT JOIN MES.Machine d ON c.MachineId = d.MachineId
                                            WHERE a.MtlItemId = @MtlItemId
                                            AND a.[Status] = 'A' 
                                            ORDER BY a.SortNumber";
                                    dynamicParameters.Add("MtlItemId", poDetail.MtlItemId);

                                    var MtlQcItemResult = sqlConnection.Query(sql, dynamicParameters);
                                    #endregion

                                    #region //設定單身量測標準
                                    int row = 2;
                                    foreach (var item2 in MtlQcItemResult)
                                    {
                                        #region //若有機台，整理序號格式
                                        string QcItemNo = item2.QcItemNo;
                                        if (item2.MachineNumber != null)
                                        {
                                            QcItemNo = item2.QcItemNo;
                                            string firstPart = QcItemNo.Substring(0, 3);
                                            string secondPart = QcItemNo.Substring(3, 4);
                                            QcItemNo = firstPart + item2.MachineNumber + secondPart;
                                        }
                                        #endregion

                                        string QcItemName = item2.QcItemName;

                                        #region //設定量測項目、備註、上下限
                                        data = new Data()
                                        {
                                            cell = "A" + row,
                                            css = "",
                                            format = "common",
                                            value = QcItemNo,
                                        };
                                        datas.Add(data);

                                        data = new Data()
                                        {
                                            cell = "B" + row,
                                            css = "",
                                            format = "common",
                                            value = QcItemName,
                                        };
                                        datas.Add(data);

                                        data = new Data()
                                        {
                                            cell = "C" + row,
                                            css = "",
                                            format = "common",
                                            value = item2.QcItemDesc,
                                        };
                                        datas.Add(data);

                                        data = new Data()
                                        {
                                            cell = "D" + row,
                                            css = "",
                                            format = "common",
                                            value = item2.MachineDesc != null ? item2.MachineDesc : "",
                                        };
                                        datas.Add(data);

                                        data = new Data()
                                        {
                                            cell = "E" + row,
                                            css = "",
                                            format = "common",
                                            value = item2.DesignValue != null ? item2.DesignValue.ToString() : "",
                                        };
                                        datas.Add(data);

                                        data = new Data()
                                        {
                                            cell = "F" + row,
                                            css = "",
                                            format = "common",
                                            value = item2.UpperTolerance != null ? item2.UpperTolerance.ToString() : "",
                                        };
                                        datas.Add(data);

                                        data = new Data()
                                        {
                                            cell = "G" + row,
                                            css = "",
                                            format = "common",
                                            value = item2.LowerTolerance != null ? item2.LowerTolerance.ToString() : "",
                                        };
                                        datas.Add(data);

                                        row++;
                                        #endregion
                                    }
                                    #endregion

                                    #region //整合Spreadsheet格式
                                    Sheets sheets = new Sheets()
                                    {
                                        name = "sheet1",
                                        data = datas
                                    };
                                    sheetss.Add(sheets);

                                    SpreadsheetModel spreadsheetModel = new SpreadsheetModel()
                                    {
                                        sheets = sheetss
                                    };

                                    string SpreadsheetData = JsonConvert.SerializeObject(spreadsheetModel);

                                    #region //更新至QcRecord SpreadsheetData
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE MES.QcRecord SET
                                            SpreadsheetData = @SpreadsheetData,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE QcRecordId = @QcRecordId";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            SpreadsheetData,
                                            LastModifiedDate,
                                            LastModifiedBy,
                                            QcRecordId
                                        });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                    #endregion
                                    #endregion
                                }
                                #endregion
                                #endregion
                            }

                            #region //重新計算進貨單資料
                            JObject updateGrDetailResult = JObject.Parse(UpdateGrTotal(GrId, goodsReceipt.Exchange, sqlConnection, sqlConnection2));
                            if (updateGrDetailResult["status"].ToString() != "success")
                            {
                                throw new SystemException(updateGrDetailResult["msg"].ToString());
                            }
                            #endregion

                            #region //更新進貨單頭交易條件
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.GoodsReceipt SET
                                    CurrencyCode = @CurrencyCode,
                                    Exchange = @Exchange,
                                    PaymentTerm = @PaymentTerm,
                                    TaxCode = @TaxCode,
                                    TaxType = @TaxType,
                                    TaxRate = @TaxRate,
                                    DeductType = @DeductType,
                                    InvoiceType = @InvoiceType,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE GrId = @GrId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    goodsReceipt.CurrencyCode,
                                    goodsReceipt.Exchange,
                                    goodsReceipt.PaymentTerm,
                                    goodsReceipt.TaxCode,
                                    goodsReceipt.TaxType,
                                    goodsReceipt.TaxRate,
                                    goodsReceipt.DeductType,
                                    goodsReceipt.InvoiceType,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    GrId
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

        #region //AddLotBindBarcode --新增綁定條碼 GPAI 20240412
        public string AddLotBindBarcode(int GrDetailId, int BarcodeQty, string BarcodeNo)
        {
            try
            {
                string companyNo = "";
                int rowsAffected = 0;
                int mtlItemId = 0;
                int receiptQty = 0;
                int lnBarcodeId = 0;
                int lnId = 0;

                int lnAllQty = 0;


                string lotNumber = "";


                #region //基本判斷
                if (GrDetailId <= 0) throw new SystemException("【單身ID】格式錯誤!!");

                if (BarcodeNo.Length <= 0) throw new SystemException("【條碼】格式錯誤!!");
                #endregion


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
                            companyNo = item.ErpNo;
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //確認進貨單身資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.GrId, a.GrDetailId, a.GrErpPrefix, a.GrErpNo, a.MtlItemId, a.ReceiptQty , a.LotNumber
                                    FROM SCM.GrDetail a
                                    WHERE a.GrDetailId = @GrDetailId";
                            dynamicParameters.Add("GrDetailId", GrDetailId);

                            var resultGrDetail = sqlConnection.Query(sql, dynamicParameters);

                            if (resultGrDetail.Count() <= 0) throw new SystemException("進貨單身資料錯誤!!");
                            foreach (var item in resultGrDetail)
                            {
                                lotNumber = item.LotNumber;
                                mtlItemId = item.MtlItemId;
                                receiptQty = item.ReceiptQty;
                            }

                            #endregion

                            #region //確認批號資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LotNumberId, LotNumberNo, MtlItemId 
                                    FROM SCM.LotNumber
                                    WHERE LotNumberNo = @LotNumberNo AND MtlItemId = @MtlItemId";
                            dynamicParameters.Add("LotNumberNo", lotNumber);
                            dynamicParameters.Add("MtlItemId", mtlItemId);


                            var resultLotNumber = sqlConnection.Query(sql, dynamicParameters);

                            if (resultLotNumber.Count() <= 0) throw new SystemException("批號資料錯誤!!");

                            foreach (var item in resultLotNumber)
                            {
                                lnId = item.LotNumberId;
                            }
                            #endregion

                            #region //確認條碼資料是否有綁定
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT * 
                                    FROM SCM.GrBarcode
                                    WHERE BarcodeNo = @BarcodeNo";
                            dynamicParameters.Add("BarcodeNo", BarcodeNo);

                            var resultLnBarcodeBind = sqlConnection.Query(sql, dynamicParameters);

                            if (resultLnBarcodeBind.Count() > 0) throw new SystemException("該條碼已綁定!");
                            
                            #endregion

                            #region //確認批號綁定條碼資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.GrBarcodeId, a.BarcodeNo, a.GrDetailId, b.BarcodeQty, c.MtlItemId, d.LotNumberNo
                                    from SCM.GrBarcode a
                                    INNER JOIN MES.Barcode b ON a.BarcodeNo = b.BarcodeNo
                                    INNER JOIN SCM.GrDetail c ON a.GrDetailId = c.GrDetailId
                                    INNER JOIN SCM.LotNumber d ON d.LotNumberNo = c.LotNumber AND d.MtlItemId = c.MtlItemId
                                    WHERE c.GrDetailId = @GrDetailId ";
                            dynamicParameters.Add("GrDetailId", GrDetailId);


                            var resultLnBarcode = sqlConnection.Query(sql, dynamicParameters);

                            if (resultLotNumber.Count() > 0)
                            {
                                foreach (var item in resultLnBarcode)
                                {
                                    if (BarcodeNo == item.BarcodeNo) throw new SystemException("該條碼已綁定在該批號");
                                    lnAllQty += item.BarcodeQty;
                                }
                            }

                            #endregion

                            #region //確認條碼資料是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT * 
                                    FROM MES.Barcode
                                    WHERE BarcodeNo = @BarcodeNo";
                            dynamicParameters.Add("BarcodeNo", BarcodeNo);

                            var resultBarcodeNo = sqlConnection.Query(sql, dynamicParameters);

                            if (resultBarcodeNo.Count() <= 0)
                            {
                                if (BarcodeQty <= 0) throw new SystemException("新條碼請先設定數量!");

                                if ((lnAllQty + BarcodeQty) > receiptQty) throw new SystemException("已超過進貨量，請重新確認輸入的條碼數量! 當前批號已綁定總量: " + lnAllQty);

                                #region //新條碼INSERT mes.barcode
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO MES.Barcode (BarcodeNo, CurrentMoProcessId, NextMoProcessId, BarcodeQty, BarcodeProcessId, CurrentProdStatus, BarcodeStatus
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            VALUES (@BarcodeNo, @CurrentMoProcessId, @NextMoProcessId, @BarcodeQty, @BarcodeProcessId, @CurrentProdStatus, @BarcodeStatus
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        BarcodeNo,
                                        CurrentMoProcessId = -1,
                                        NextMoProcessId = -1,
                                        //MoId = -1,//
                                        BarcodeQty,
                                        BarcodeProcessId = -1,
                                        CurrentProdStatus = "P",
                                        BarcodeStatus = "8",
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy ,
                                        LastModifiedBy
                                    });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                                #endregion

                            }
                            else {

                                

                                foreach (var item in resultBarcodeNo) {
                                    if (item.BarcodeStatus != "10") throw new SystemException("該條碼狀態並非已出貨! 請更換條碼! ");

                                    #region //是否有過站紀錄
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT * 
                                    FROM MES.BarcodeProcess
                                    WHERE BarcodeId = @BarcodeId";
                                    dynamicParameters.Add("BarcodeId", item.BarcodeId);

                                    var resultBarcodeProcess = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultBarcodeProcess.Count() > 0) throw new SystemException("該條碼有過站紀錄!");

                                    #endregion
                                    var barcodeQty = item.BarcodeQty;
                                    if ((lnAllQty + barcodeQty) > receiptQty) throw new SystemException("已超過進貨量，請重新確認輸入的條碼數量! 當前批號已綁定總量: " + lnAllQty);

                                    #region //UPDATE MES.Barcode
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE MES.Barcode SET
                                    BarcodeStatus = @BarcodeStatus,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE BarcodeId = @BarcodeId";
                                    dynamicParameters.AddDynamicParams(
                                      new
                                      {
                                          BarcodeStatus = "8",
                                          LastModifiedDate,
                                          LastModifiedBy,
                                          BarcodeId = item.BarcodeId
                                      });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                }
                            }
                            #endregion

                            #region //INSERT SCM.LnBarcode
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO SCM.GrBarcode (GrDetailId,BarcodeNo
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.GrBarcodeId
                                            VALUES (@GrDetailId, @BarcodeNo
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    GrDetailId = GrDetailId,
                                    BarcodeNo,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO SCM.LnBarcode (LotNumberId,BarcodeNo
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.LnBarcodeId
                                            VALUES (@LotNumberId, @BarcodeNo
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    LotNumberId = lnId,
                                    BarcodeNo = BarcodeNo,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
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

        #region //AddLotBindCreatBarcode --產出條碼並綁定批號 GPAI 20240412
        public string AddLotBindCreatBarcode(int GrDetailId, int BarcodeQty, string prefix, string suffix, int seq)
        {
            try
            {
                string companyNo = "";
                int rowsAffected = 0;
                int mtlItemId = 0;
                int receiptQty = 0;
                int lnBarcodeId = 0;
                int lnId = 0;

                int lnAllQty = 0;


                string lotNumber = "";
                string newBarcode = "";

                int div = 0;//整數
                int mod = 0;//餘數

                #region //基本判斷
                if (GrDetailId <= 0) throw new SystemException("【單身ID】格式錯誤!!");
                #endregion


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
                            companyNo = item.ErpNo;
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //確認進貨單身資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.GrId, a.GrDetailId, a.GrErpPrefix, a.GrErpNo, a.MtlItemId, a.ReceiptQty , a.LotNumber
                                    FROM SCM.GrDetail a
                                    WHERE a.GrDetailId = @GrDetailId";
                            dynamicParameters.Add("GrDetailId", GrDetailId);

                            var resultGrDetail = sqlConnection.Query(sql, dynamicParameters);

                            if (resultGrDetail.Count() <= 0) throw new SystemException("進貨單身資料錯誤!!");
                            foreach (var item in resultGrDetail)
                            {
                                lotNumber = item.LotNumber;
                                mtlItemId = item.MtlItemId;
                                receiptQty = item.ReceiptQty;
                                //div = BarcodeQty / receiptQty;
                                //mod = BarcodeQty % receiptQty;
                                //if (mod > 0 || div <= 0) throw new SystemException("條碼數量設定有誤!須能整除進貨量! 進貨量: " + receiptQty);
                                
                            }

                            #endregion

                            #region //確認批號資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LotNumberId, LotNumberNo, MtlItemId 
                                    FROM SCM.LotNumber
                                    WHERE LotNumberNo = @LotNumberNo AND MtlItemId = @MtlItemId";
                            dynamicParameters.Add("LotNumberNo", lotNumber);
                            dynamicParameters.Add("MtlItemId", mtlItemId);


                            var resultLotNumber = sqlConnection.Query(sql, dynamicParameters);

                            if (resultLotNumber.Count() <= 0) throw new SystemException("批號資料錯誤!!");

                            foreach (var item in resultLotNumber)
                            {
                                lnId = item.LotNumberId;
                            }
                            #endregion

                            #region //確認批號綁定條碼資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.GrBarcodeId, a.BarcodeNo, a.GrDetailId, b.BarcodeQty, c.MtlItemId, d.LotNumberNo
                                    from SCM.GrBarcode a
                                    INNER JOIN MES.Barcode b ON a.BarcodeNo = b.BarcodeNo
                                    INNER JOIN SCM.GrDetail c ON a.GrDetailId = c.GrDetailId
                                    INNER JOIN SCM.LotNumber d ON d.LotNumberNo = c.LotNumber AND d.MtlItemId = c.MtlItemId
                                    WHERE c.GrDetailId = @GrDetailId ";
                            dynamicParameters.Add("GrDetailId", GrDetailId);


                            var resultLnBarcode = sqlConnection.Query(sql, dynamicParameters);

                            if (resultLotNumber.Count() > 0)
                            {
                                foreach (var item in resultLnBarcode)
                                {
                                    lnAllQty += item.BarcodeQty;
                                }
                            }
                            div = (receiptQty - lnAllQty) / BarcodeQty;
                            mod = receiptQty % BarcodeQty;
                            if (mod > 0 || div <= 0) throw new SystemException("條碼數量設定有誤!須能整除進貨量! 進貨量: " + receiptQty);
                            #endregion

                            #region //找同前後綴的資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT * FROM MES.Barcode
                                    WHERE 1=1";
                            // dynamicParameters.Add("LotNumberId", lnId);

                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "prefix", @" AND BarcodeNo LIKE  @prefix + '%'", prefix);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "suffix", @" AND BarcodeNo LIKE '%' + @suffix", suffix);

                            var resultBarcodeCount = sqlConnection.Query(sql, dynamicParameters);
                            #endregion


                            #region //INSERT 

                            #region //新條碼INSERT mes.barcode
                            if ((lnAllQty + (BarcodeQty * div)) > receiptQty) throw new SystemException("已超過進貨量，請重新確認輸入的條碼數量! 當前批號已綁定總量: " + lnAllQty);

                            for (int i = 1; i <= div; i++)
                            {
                                int newNum = resultBarcodeCount.Count() + i;

                                newBarcode = prefix + newNum.ToString().PadLeft(seq, '0') + suffix;

                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO MES.Barcode (BarcodeNo, CurrentMoProcessId, NextMoProcessId, BarcodeQty, BarcodeProcessId, CurrentProdStatus, BarcodeStatus
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            VALUES (@BarcodeNo, @CurrentMoProcessId, @NextMoProcessId, @BarcodeQty, @BarcodeProcessId, @CurrentProdStatus, @BarcodeStatus
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        BarcodeNo = newBarcode,
                                        CurrentMoProcessId = -1,
                                        NextMoProcessId = -1,
                                        //MoId = -1,//
                                        BarcodeQty,
                                        BarcodeProcessId = -1,
                                        CurrentProdStatus = "P",
                                        BarcodeStatus = "8",
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);


                                //SCM.LnBarcode
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO SCM.GrBarcode (GrDetailId,BarcodeNo
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.GrBarcodeId
                                            VALUES (@GrDetailId, @BarcodeNo
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        GrDetailId = GrDetailId,
                                        BarcodeNo  = newBarcode,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO SCM.LnBarcode (LotNumberId,BarcodeNo
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.LnBarcodeId
                                            VALUES (@LotNumberId, @BarcodeNo
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        LotNumberId = lnId,
                                        BarcodeNo = newBarcode,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

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

        #region //IntoLotBindBarcode -- 帶入Lot BindBarode Excel -- GPAI 20240520
        public string IntoLotBindBarcode(List<LotBarcode> LotBarcodes, int GrDetailId, int FileId)
        {
            try
            {
                int rowsAffected = 0;
                int CompanyId = -1;

                string companyNo = "";
                int mtlItemId = 0;
                int receiptQty = 0;
                int lnBarcodeId = 0;
                int lnId = 0;

                int lnAllQty = 0;


                string lotNumber = "";

                using (TransactionScope transactionScope = new TransactionScope(TransactionScopeOption.Required, new TimeSpan(0, 180, 0)))
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認公司別DB
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CompanyId
                                FROM BAS.Company a
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var CompanyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (CompanyResult.Count() <= 0) throw new SystemException("公司別錯誤!");

                        foreach (var item in CompanyResult)
                        {
                            CompanyId = item.CompanyId;
                        }
                        #endregion

                        #region //確認進貨單身資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.GrId, a.GrDetailId, a.GrErpPrefix, a.GrErpNo, a.MtlItemId, a.ReceiptQty , a.LotNumber
                                    FROM SCM.GrDetail a
                                    WHERE a.GrDetailId = @GrDetailId";
                        dynamicParameters.Add("GrDetailId", GrDetailId);

                        var resultGrDetail = sqlConnection.Query(sql, dynamicParameters);

                        if (resultGrDetail.Count() <= 0) throw new SystemException("進貨單身資料錯誤!!");
                        foreach (var item in resultGrDetail)
                        {
                            lotNumber = item.LotNumber;
                            mtlItemId = item.MtlItemId;
                            receiptQty = item.ReceiptQty;
                        }

                        #endregion

                        #region //確認批號資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT LotNumberId, LotNumberNo, MtlItemId 
                                    FROM SCM.LotNumber
                                    WHERE LotNumberNo = @LotNumberNo AND MtlItemId = @MtlItemId";
                        dynamicParameters.Add("LotNumberNo", lotNumber);
                        dynamicParameters.Add("MtlItemId", mtlItemId);


                        var resultLotNumber = sqlConnection.Query(sql, dynamicParameters);

                        if (resultLotNumber.Count() <= 0) throw new SystemException("批號資料錯誤!!");

                        foreach (var item in resultLotNumber)
                        {
                            lnId = item.LotNumberId;
                        }
                        #endregion

                        #region //比對EXCEL項目是否存在現有項目，有則INSERT，無則UPDATE
                        foreach (var item in LotBarcodes)
                        {
                            #region //確認條碼資料是否有綁定
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT * 
                                    FROM SCM.GrBarcode
                                    WHERE BarcodeNo = @BarcodeNo";
                            dynamicParameters.Add("BarcodeNo", item.BarcodeNo);

                            var resultLnBarcodeBind = sqlConnection.Query(sql, dynamicParameters);

                            if (resultLnBarcodeBind.Count() > 0) throw new SystemException("該條碼已綁定!");

                            #endregion

                            #region //確認批號綁定條碼資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.GrBarcodeId, a.BarcodeNo, a.GrDetailId, b.BarcodeQty, c.MtlItemId, d.LotNumberNo
                                    from SCM.GrBarcode a
                                    INNER JOIN MES.Barcode b ON a.BarcodeNo = b.BarcodeNo
                                    INNER JOIN SCM.GrDetail c ON a.GrDetailId = c.GrDetailId
                                    INNER JOIN SCM.LotNumber d ON d.LotNumberNo = c.LotNumber AND d.MtlItemId = c.MtlItemId
                                    WHERE c.GrDetailId = @GrDetailId ";
                            dynamicParameters.Add("GrDetailId", GrDetailId);


                            var resultLnBarcode = sqlConnection.Query(sql, dynamicParameters);

                            if (resultLotNumber.Count() > 0)
                            {
                                foreach (var item2 in resultLnBarcode)
                                {
                                    if (item.BarcodeNo == item2.BarcodeNo) throw new SystemException("該條碼已綁定在該批號");
                                    lnAllQty += item.BarcodeCount;
                                }
                            }

                            #endregion

                            #region //確認條碼資料是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT * 
                                    FROM MES.Barcode
                                    WHERE BarcodeNo = @BarcodeNo";
                            dynamicParameters.Add("BarcodeNo", item.BarcodeNo);

                            var resultBarcodeNo = sqlConnection.Query(sql, dynamicParameters);

                            if (resultBarcodeNo.Count() <= 0)
                            {
                                if (item.BarcodeCount <= 0) throw new SystemException("新條碼請先設定數量!");

                                //if ((lnAllQty + item.BarcodeCount) > receiptQty) throw new SystemException("已超過進貨量，請重新確認輸入的條碼數量! 當前批號已綁定總量: " + lnAllQty);

                                #region //新條碼INSERT mes.barcode
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO MES.Barcode (BarcodeNo, CurrentMoProcessId, NextMoProcessId, BarcodeQty, BarcodeProcessId, CurrentProdStatus, BarcodeStatus
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            VALUES (@BarcodeNo, @CurrentMoProcessId, @NextMoProcessId, @BarcodeQty, @BarcodeProcessId, @CurrentProdStatus, @BarcodeStatus
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        item.BarcodeNo,
                                        CurrentMoProcessId = -1,
                                        NextMoProcessId = -1,
                                        //MoId = -1,//
                                        BarcodeQty = item.BarcodeCount,
                                        BarcodeProcessId = -1,
                                        CurrentProdStatus = "P",
                                        BarcodeStatus = "8",
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



                                foreach (var item2 in resultBarcodeNo)
                                {
                                    if (item2.BarcodeStatus != "10") throw new SystemException("該條碼狀態並非已出貨! 請更換條碼! ");

                                    #region //是否有過站紀錄
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT * 
                                    FROM MES.BarcodeProcess
                                    WHERE BarcodeId = @BarcodeId";
                                    dynamicParameters.Add("BarcodeId", item2.BarcodeId);

                                    var resultBarcodeProcess = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultBarcodeProcess.Count() > 0) throw new SystemException("該條碼有過站紀錄!");

                                    #endregion
                                    var barcodeQty = item2.BarcodeQty;
                                    //if ((lnAllQty + barcodeQty) > receiptQty) throw new SystemException("已超過進貨量，請重新確認輸入的條碼數量! 當前批號已綁定總量: " + lnAllQty);

                                    #region //UPDATE MES.Barcode
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE MES.Barcode SET
                                    BarcodeStatus = @BarcodeStatus,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE BarcodeId = @BarcodeId";
                                    dynamicParameters.AddDynamicParams(
                                      new
                                      {
                                          BarcodeStatus = "8",
                                          LastModifiedDate,
                                          LastModifiedBy,
                                          BarcodeId = item2.BarcodeId
                                      });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                }
                            }
                            #endregion

                            #region //INSERT SCM.LnBarcode
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO SCM.GrBarcode (GrDetailId,BarcodeNo
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.GrBarcodeId
                                            VALUES (@GrDetailId, @BarcodeNo
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    GrDetailId = GrDetailId,
                                    item.BarcodeNo,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO SCM.LnBarcode (LotNumberId,BarcodeNo
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.LnBarcodeId
                                            VALUES (@LotNumberId, @BarcodeNo
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    LotNumberId = lnId,
                                    BarcodeNo = item.BarcodeNo,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }
                        #endregion

                        #region //刪除檔案
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
        #region //UpdateGoodsReceiptSynchronize -- 進貨單資料同步 -- Ann 2024-02-26
        public string UpdateGoodsReceiptSynchronize(string CompanyNo, string UpdateDate)
        {
            try
            {
                List<GoodsReceipt> goodsReceipts = new List<GoodsReceipt>();
                List<GrDetail> grDetails = new List<GrDetail>();

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
                        #region //撈取ERP進貨單頭資料
                        sql = @"SELECT LTRIM(RTRIM(TG001)) GrErpPrefix, LTRIM(RTRIM(TG002)) GrErpNo, LTRIM(RTRIM(CREATOR)) CreateUserNo, LTRIM(RTRIM(MODIFIER)) LastModifiedUserNo
                                , CASE WHEN LEN(LTRIM(RTRIM(TG003))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TG003)) as date), 'yyyy-MM-dd') ELSE NULL END ReceiptDate
                                , LTRIM(RTRIM(TG005)) SupplierNo, LTRIM(RTRIM(TG006)) SupplierSo, LTRIM(RTRIM(TG007)) CurrencyCode, TG008 Exchange, LTRIM(RTRIM(TG009)) InvoiceType
                                , LTRIM(RTRIM(TG010)) TaxType, LTRIM(RTRIM(TG011)) InvoiceNo, LTRIM(RTRIM(TG012)) PrintCnt, LTRIM(RTRIM(TG013)) ConfirmStatus
                                , CASE WHEN LEN(LTRIM(RTRIM(TG014))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TG014)) as date), 'yyyy-MM-dd') ELSE NULL END DocDate
                                , LTRIM(RTRIM(TG015)) RenewFlag, LTRIM(RTRIM(TG016)) Remark, TG017 TotalAmount, TG018 DeductAmount, TG019 OrigTax, TG020 ReceiptAmount
                                , LTRIM(RTRIM(TG021)) SupplierName, LTRIM(RTRIM(TG022)) UiNo, LTRIM(RTRIM(TG023)) DeductType, LTRIM(RTRIM(TG024)) ObaccoAndLiquorFlag
                                , TG025 RowCnt, TG026 Quantity
                                , CASE WHEN LEN(LTRIM(RTRIM(TG027))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TG027)) as date), 'yyyy-MM-dd') ELSE NULL END InvoiceDate
                                , TG028 OrigPreTaxAmount, LTRIM(RTRIM(TG029)) ApplyYYMM, TG030 TaxRate, TG031 PreTaxAmount, TG032 TaxAmount, LTRIM(RTRIM(TG033)) PaymentTerm
                                , LTRIM(RTRIM(TG034)) PoErpPrefix, LTRIM(RTRIM(TG035)) PoErpNo, LTRIM(RTRIM(TG036)) PrepaidErpPrefix, LTRIM(RTRIM(TG037)) PrepaidErpNo
                                , TG038 OffsetAmount, TG039 TaxOffset, TG040 PackageQuantity, TG041 PremiumAmount, LTRIM(RTRIM(TG042)) ApproveStatus, LTRIM(RTRIM(TG043)) InvoiceFlag
                                , LTRIM(RTRIM(TG044)) ReserveTaxCode, TG045 SendCount, TG046 OrigPremiumAmount, LTRIM(RTRIM(TG047)) EbcErpPreNo, LTRIM(RTRIM(TG048)) EbcEdition
                                , LTRIM(RTRIM(TG049)) EbcFlag, LTRIM(RTRIM(TG050)) FromDocType, LTRIM(RTRIM(TG055)) NoticeFlag, LTRIM(RTRIM(TG058)) TaxCode
                                , LTRIM(RTRIM(TG059)) TradeTerm, LTRIM(RTRIM(TG063)) DetailMultiTax, TG064 TaxExchange, LTRIM(RTRIM(TG067)) ContactUser
                                , CASE WHEN LEN(LTRIM(RTRIM(CREATE_DATE))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_DATE)) as date), 'yyyy-MM-dd') ELSE NULL END TransferDate
                                , CASE WHEN LEN(LTRIM(RTRIM(CREATE_TIME))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_TIME)) as datetime), 'HH:mm:ss') ELSE NULL END TransferTime
                                FROM PURTG
                                WHERE 1=1";
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "UpdateDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME > @UpdateDate OR MODI_DATE + ' ' + MODI_TIME > @UpdateDate)", UpdateDate);
                        sql += @" ORDER BY TransferDate, TransferTime";

                        goodsReceipts = sqlConnection.Query<GoodsReceipt>(sql, dynamicParameters).ToList();
                        #endregion

                        #region //撈取ERP進貨單身資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT LTRIM(RTRIM(TH001)) GrErpPrefix, LTRIM(RTRIM(TH002)) GrErpNo, LTRIM(RTRIM(TH003)) GrSeq
                                , LTRIM(RTRIM(CREATOR)) CreateUserNo, LTRIM(RTRIM(MODIFIER)) LastModifiedUserNo
                                , LTRIM(RTRIM(TH004)) MtlItemNo, LTRIM(RTRIM(TH005)) GrMtlItemName, LTRIM(RTRIM(TH006)) GrMtlItemSpec, TH007 ReceiptQty
                                , LTRIM(RTRIM(TH008)) UomNo, LTRIM(RTRIM(TH009)) InventoryNo, LTRIM(RTRIM(TH010)) LotNumber
                                , LTRIM(RTRIM(TH011)) PoErpPrefix, LTRIM(RTRIM(TH012)) PoErpNo, LTRIM(RTRIM(TH013)) PoSeq
                                , CASE WHEN LEN(LTRIM(RTRIM(TH014))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TH014)) as date), 'yyyy-MM-dd') ELSE NULL END AcceptanceDate
                                , TH015 AcceptQty, TH016 AvailableQty, TH017 ReturnQty, TH018 OrigUnitPrice, TH019 OrigAmount, TH020 OrigDiscountAmt
                                , LTRIM(RTRIM(TH021)) TrErpPrefix, LTRIM(RTRIM(TH022)) TrErpNo, LTRIM(RTRIM(TH023)) TrSeq
                                , TH024 ReceiptExpense, LTRIM(RTRIM(TH025)) DiscountDescription, LTRIM(RTRIM(TH026)) PaymentHold, LTRIM(RTRIM(TH027)) Overdue
                                , LTRIM(RTRIM(TH028)) QcStatus, LTRIM(RTRIM(TH029)) ReturnCode, LTRIM(RTRIM(TH030)) ConfirmStatus, LTRIM(RTRIM(TH031)) CloseStatus
                                , LTRIM(RTRIM(TH032)) ReNewStatus, LTRIM(RTRIM(TH033)) Remark, TH034 InventoryQty, LTRIM(RTRIM(TH035)) SmallUnit
                                , CASE WHEN LEN(LTRIM(RTRIM(TH036))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TH036)) as date), 'yyyy-MM-dd') ELSE NULL END AvailableDate
                                , CASE WHEN LEN(LTRIM(RTRIM(TH037))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TH037)) as date), 'yyyy-MM-dd') ELSE NULL END ReCheckDate
                                , LTRIM(RTRIM(TH038)) ConfirmUserNo, LTRIM(RTRIM(TH039)) ApvErpPrefix, LTRIM(RTRIM(TH040)) ApvErpNo, LTRIM(RTRIM(TH041)) ApvSeq
                                , LTRIM(RTRIM(TH042)) ProjectCode, LTRIM(RTRIM(TH043)) ExpenseEntry, LTRIM(RTRIM(TH044)) PremiumAmountFlag, TH045 OrigPreTaxAmt
                                , TH046 OrigTaxAmt, TH047 PreTaxAmt, TH048 TaxAmt, TH049 ReceiptPackageQty, TH050 AcceptancePackageQty, TH051 ReturnPackageQty
                                , TH052 PremiumAmount, LTRIM(RTRIM(TH053)) PackageUnit, LTRIM(RTRIM(TH054)) ReserveTaxCode, TH055 OrigPremiumAmount
                                , LTRIM(RTRIM(TH056)) AvailableUomNo, LTRIM(RTRIM(TH057)) OrigCustomer, LTRIM(RTRIM(TH058)) ApproveStatus
                                , LTRIM(RTRIM(TH059)) EbcErpPreNo, LTRIM(RTRIM(TH060)) EbcEdition, TH061 ProductSeqAmount, LTRIM(RTRIM(TH062)) MtlItemType
                                , LTRIM(RTRIM(TH063)) Loaction, TH073 TaxRate, LTRIM(RTRIM(TH074)) MultipleFlag, LTRIM(RTRIM(TH082)) GrErpStatus, LTRIM(RTRIM(TH088)) TaxCode
                                , TH089 DiscountRate, TH090 DiscountPrice, LTRIM(RTRIM(TH091)) QcType
                                , CASE WHEN LEN(LTRIM(RTRIM(CREATE_DATE))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_DATE)) as date), 'yyyy-MM-dd') ELSE NULL END TransferDate
                                , CASE WHEN LEN(LTRIM(RTRIM(CREATE_TIME))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_TIME)) as datetime), 'HH:mm:ss') ELSE NULL END TransferTime
                                FROM PURTH
                                WHERE 1=1";
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "UpdateDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME > @UpdateDate OR MODI_DATE + ' ' + MODI_TIME > @UpdateDate)", UpdateDate);
                        sql += @" ORDER BY TransferDate, TransferTime";

                        grDetails = sqlConnection.Query<GrDetail>(sql, dynamicParameters).ToList();
                        #endregion
                    }

                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //取得MES供應商資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT SupplierId, SupplierNo
                                FROM SCM.Supplier
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<Supplier> resultSuppliers = sqlConnection.Query<Supplier>(sql, dynamicParameters).ToList();

                        goodsReceipts = goodsReceipts.Join(resultSuppliers, x => x.SupplierNo, y => y.SupplierNo, (x, y) => { x.SupplierId = y.SupplierId; return x; }).ToList();
                        #endregion

                        #region //撈取單位ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT UomId, UomNo
                                FROM PDM.UnitOfMeasure
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<Uom> resultSoPriceUomrNos = sqlConnection.Query<Uom>(sql, dynamicParameters).ToList();

                        grDetails = grDetails.Join(resultSoPriceUomrNos, x => x.UomNo.ToUpper(), y => y.UomNo, (x, y) => { x.UomId = y.UomId; return x; }).ToList();
                        grDetails = grDetails.Join(resultSoPriceUomrNos, x => x.AvailableUomNo.ToUpper(), y => y.UomNo, (x, y) => { x.AvailableUomId = y.UomId; return x; }).ToList();
                        #endregion

                        #region //撈取品號ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT MtlItemId, MtlItemNo
                                FROM PDM.MtlItem
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<MtlItem> resultMtlitems = sqlConnection.Query<MtlItem>(sql, dynamicParameters).ToList();

                        grDetails = grDetails.Join(resultMtlitems, x => x.MtlItemNo, y => y.MtlItemNo, (x, y) => { x.MtlItemId = y.MtlItemId; return x; }).Select(x => x).ToList();
                        #endregion

                        #region //撈取庫別ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT InventoryId, InventoryNo
                                FROM SCM.Inventory
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<Inventory> resultInventories = sqlConnection.Query<Inventory>(sql, dynamicParameters).ToList();

                        grDetails = grDetails.Join(resultInventories, x => x.InventoryNo, y => y.InventoryNo, (x, y) => { x.InventoryId = y.InventoryId; return x; }).ToList();
                        #endregion

                        #region //撈取使用者ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserId, a.UserNo 
                                FROM BAS.[User] a";

                        List<User> resultConfirmUsers = sqlConnection.Query<User>(sql, dynamicParameters).ToList();

                        grDetails = grDetails.Join(resultConfirmUsers, x => x.ConfirmUserNo, y => y.UserNo, (x, y) => { x.ConfirmUser = y.UserId; return x; }).ToList();
                        #endregion

                        #region //判斷SCM.GoodsReceipt是否有資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT GrId, GrErpPrefix, GrErpNo
                                FROM SCM.GoodsReceipt
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<GoodsReceipt> resultGoodsReceipt = sqlConnection.Query<GoodsReceipt>(sql, dynamicParameters).ToList();

                        goodsReceipts = goodsReceipts.GroupJoin(resultGoodsReceipt, x => new { x.GrErpPrefix, x.GrErpNo }, y => new { y.GrErpPrefix, y.GrErpNo }, (x, y) => { x.GrId = y.FirstOrDefault()?.GrId; return x; }).Select(x => x).ToList();
                        #endregion

                        #region //進貨單單頭(新增/修改)
                        List<GoodsReceipt> addGoodsReceipts = goodsReceipts.Where(x => x.GrId == null).ToList();
                        List<GoodsReceipt> updateGoodsReceipts = goodsReceipts.Where(x => x.GrId != null).ToList();

                        #region //新增
                        dynamicParameters = new DynamicParameters();
                        if (addGoodsReceipts.Count > 0)
                        {
                            addGoodsReceipts
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

                            sql = @"INSERT INTO SCM.GoodsReceipt (CompanyId, GrErpPrefix, GrErpNo, ReceiptDate, SupplierId, SupplierNo
                                    , SupplierSo, CurrencyCode, Exchange, InvoiceType, TaxType, InvoiceNo, PrintCnt, ConfirmStatus, DocDate
                                    , RenewFlag, Remark, TotalAmount, DeductAmount, OrigTax, ReceiptAmount, SupplierName, UiNo, DeductType
                                    , ObaccoAndLiquorFlag, RowCnt, Quantity, InvoiceDate, OrigPreTaxAmount, ApplyYYMM, TaxRate, PreTaxAmount, TaxAmount
                                    , PaymentTerm, PoId, PoErpPrefix, PoErpNo, PrepaidErpPrefix, PrepaidErpNo, OffsetAmount, TaxOffset, PackageQuantity
                                    , PremiumAmount, ApproveStatus, InvoiceFlag, ReserveTaxCode, SendCount, OrigPremiumAmount, EbcErpPreNo, EbcEdition, EbcFlag
                                    , FromDocType, NoticeFlag, TaxCode, TradeTerm, DetailMultiTax, TaxExchange, ContactUser, TransferStatus, TransferDate
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@CompanyId, @GrErpPrefix, @GrErpNo, @ReceiptDate, @SupplierId, @SupplierNo
                                    , @SupplierSo, @CurrencyCode, @Exchange, @InvoiceType, @TaxType, @InvoiceNo, @PrintCnt, @ConfirmStatus, @DocDate
                                    , @RenewFlag, @Remark, @TotalAmount, @DeductAmount, @OrigTax, @ReceiptAmount, @SupplierName, @UiNo, @DeductType
                                    , @ObaccoAndLiquorFlag, @RowCnt, @Quantity, @InvoiceDate, @OrigPreTaxAmount, @ApplyYYMM, @TaxRate, @PreTaxAmount, @TaxAmount
                                    , @PaymentTerm, @PoId, @PoErpPrefix, @PoErpNo, @PrepaidErpPrefix, @PrepaidErpNo, @OffsetAmount, @TaxOffset, @PackageQuantity
                                    , @PremiumAmount, @ApproveStatus, @InvoiceFlag, @ReserveTaxCode, @SendCount, @OrigPremiumAmount, @EbcErpPreNo, @EbcEdition, @EbcFlag
                                    , @FromDocType, @NoticeFlag, @TaxCode, @TradeTerm, @DetailMultiTax, @TaxExchange, @ContactUser, @TransferStatus, @TransferDate
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            rowsAffected += sqlConnection.Execute(sql, addGoodsReceipts);
                        }
                        #endregion

                        #region //修改
                        dynamicParameters = new DynamicParameters();
                        if (updateGoodsReceipts.Count > 0)
                        {
                            updateGoodsReceipts
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.LastModifiedDate = LastModifiedDate;
                                    x.LastModifiedBy = LastModifiedBy;
                                });

                            sql = @"UPDATE SCM.GoodsReceipt SET
                                    ReceiptDate = @ReceiptDate,
                                    SupplierId = @SupplierId,
                                    SupplierNo = @SupplierNo,
                                    CurrencyCode = @CurrencyCode,
                                    Exchange = @Exchange,
                                    InvoiceType = @InvoiceType,
                                    TaxType = @TaxType,
                                    InvoiceNo = @InvoiceNo,
                                    ConfirmStatus = @ConfirmStatus,
                                    DocDate = @DocDate,
                                    Remark = @Remark,
                                    TotalAmount = @TotalAmount,
                                    DeductAmount = @DeductAmount,
                                    OrigTax = @OrigTax,
                                    ReceiptAmount = @ReceiptAmount,
                                    SupplierName = @SupplierName,
                                    UiNo = @UiNo,
                                    DeductType = @DeductType,
                                    ObaccoAndLiquorFlag = @ObaccoAndLiquorFlag,
                                    Quantity = @Quantity,
                                    InvoiceDate = @InvoiceDate,
                                    OrigPreTaxAmount = @OrigPreTaxAmount,
                                    ApplyYYMM = @ApplyYYMM,
                                    TaxRate = @TaxRate,
                                    PreTaxAmount = @PreTaxAmount,
                                    TaxAmount = @TaxAmount,
                                    PaymentTerm = @PaymentTerm,
                                    OffsetAmount = @OffsetAmount,
                                    TaxOffset = @TaxOffset,
                                    PremiumAmount = @PremiumAmount,
                                    ApproveStatus = @ApproveStatus,
                                    OrigPremiumAmount = @OrigPremiumAmount,
                                    FromDocType = @FromDocType,
                                    TaxCode = @TaxCode,
                                    TradeTerm = @TradeTerm,
                                    DetailMultiTax = @DetailMultiTax,
                                    ContactUser = @ContactUser,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE GrId = @GrId";
                            rowsAffected += sqlConnection.Execute(sql, updateGoodsReceipts);
                        }
                        #endregion
                        #endregion

                        #region //進貨單單身(新增/修改)
                        #region //撈取進貨單單頭ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT GrId, GrErpPrefix, GrErpNo
                                FROM  SCM.GoodsReceipt
                                WHERE CompanyId = @CompanyId
                                ORDER BY GrId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        resultGoodsReceipt = sqlConnection.Query<GoodsReceipt>(sql, dynamicParameters).ToList();

                        grDetails = grDetails.Join(resultGoodsReceipt, x => new { x.GrErpPrefix, x.GrErpNo }, y => new { y.GrErpPrefix, y.GrErpNo }, (x, y) => { x.GrId = y.GrId; return x; }).Select(x => x).ToList();
                        #endregion

                        #region //判斷SCM.GrDetail是否有資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.GrDetailId, a.GrId, a.GrSeq
                                FROM SCM.GrDetail a
                                INNER JOIN SCM.GoodsReceipt b ON a.GrId = b.GrId
                                WHERE b.CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<GrDetail> resultGrDetail = sqlConnection.Query<GrDetail>(sql, dynamicParameters).ToList();
                        
                        grDetails = grDetails.GroupJoin(resultGrDetail, x => new { x.GrId, x.GrSeq }, y => new { y.GrId, y.GrSeq }, (x, y) => { x.GrDetailId = y.FirstOrDefault()?.GrDetailId; return x; }).ToList();
                        #endregion

                        List<GrDetail> addGrDetail = grDetails.Where(x => x.GrDetailId == null).ToList();
                        List<GrDetail> updateGrDetail = grDetails.Where(x => x.GrDetailId != null).ToList();

                        #region //新增
                        dynamicParameters = new DynamicParameters();
                        if (addGrDetail.Count > 0)
                        {
                            addGrDetail
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.TransferStatus = "Y";
                                    x.CreateDate = CreateDate;
                                    x.LastModifiedDate = LastModifiedDate;
                                    x.CreateBy = CreateBy;
                                    x.LastModifiedBy = LastModifiedBy;
                                });

                            sql = @"INSERT INTO SCM.GrDetail (GrId, GrErpPrefix, GrErpNo, GrSeq, MtlItemId, MtlItemNo, GrMtlItemName, GrMtlItemSpec
                                    , ReceiptQty, UomId, UomNo, InventoryId, InventoryNo, LotNumber, PoErpPrefix, PoErpNo, PoSeq, AcceptanceDate
                                    , AcceptQty, AvailableQty, ReturnQty, OrigUnitPrice, OrigAmount, OrigDiscountAmt, TrErpPrefix, TrErpNo, TrSeq
                                    , ReceiptExpense, DiscountDescription, PaymentHold, Overdue, QcStatus, ReturnCode, ConfirmStatus, CloseStatus
                                    , ReNewStatus, Remark, InventoryQty, SmallUnit, AvailableDate, ReCheckDate, ConfirmUser, ApvErpPrefix, ApvErpNo
                                    , ApvSeq, ProjectCode, ExpenseEntry, PremiumAmountFlag, OrigPreTaxAmt, OrigTaxAmt, PreTaxAmt, TaxAmt, ReceiptPackageQty
                                    , AcceptancePackageQty, ReturnPackageQty, PremiumAmount, PackageUnit, ReserveTaxCode, OrigPremiumAmount, AvailableUomId
                                    , AvailableUomNo, OrigCustomer, ApproveStatus, EbcErpPreNo, EbcEdition, ProductSeqAmount, MtlItemType, Loaction
                                    , TaxRate, MultipleFlag, GrErpStatus, TaxCode, DiscountRate, DiscountPrice, QcType, TransferStatus, TransferDate
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@GrId, @GrErpPrefix, @GrErpNo, @GrSeq, @MtlItemId, @MtlItemNo, @GrMtlItemName, @GrMtlItemSpec
                                    , @ReceiptQty, @UomId, @UomNo, @InventoryId, @InventoryNo, @LotNumber, @PoErpPrefix, @PoErpNo, @PoSeq, @AcceptanceDate
                                    , @AcceptQty, @AvailableQty, @ReturnQty, @OrigUnitPrice, @OrigAmount, @OrigDiscountAmt, @TrErpPrefix, @TrErpNo, @TrSeq
                                    , @ReceiptExpense, @DiscountDescription, @PaymentHold, @Overdue, @QcStatus, @ReturnCode, @ConfirmStatus, @CloseStatus
                                    , @ReNewStatus, @Remark, @InventoryQty, @SmallUnit, @AvailableDate, @ReCheckDate, @ConfirmUser, @ApvErpPrefix, @ApvErpNo
                                    , @ApvSeq, @ProjectCode, @ExpenseEntry, @PremiumAmountFlag, @OrigPreTaxAmt, @OrigTaxAmt, @PreTaxAmt, @TaxAmt, @ReceiptPackageQty
                                    , @AcceptancePackageQty, @ReturnPackageQty, @PremiumAmount, @PackageUnit, @ReserveTaxCode, @OrigPremiumAmount, @AvailableUomId
                                    , @AvailableUomNo, @OrigCustomer, @ApproveStatus, @EbcErpPreNo, @EbcEdition, @ProductSeqAmount, @MtlItemType, @Loaction
                                    , @TaxRate, @MultipleFlag, @GrErpStatus, @TaxCode, @DiscountRate, @DiscountPrice, @QcType, @TransferStatus, @TransferDate
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            rowsAffected += sqlConnection.Execute(sql, addGrDetail);
                        }
                        #endregion

                        #region //修改
                        dynamicParameters = new DynamicParameters();
                        if (updateGrDetail.Count > 0)
                        {
                            updateGrDetail
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.LastModifiedDate = LastModifiedDate;
                                    x.LastModifiedBy = LastModifiedBy;
                                });

                            sql = @"UPDATE SCM.GrDetail SET
                                    MtlItemId = @MtlItemId,
                                    MtlItemNo = @MtlItemNo,
                                    GrMtlItemName = @GrMtlItemName,
                                    GrMtlItemSpec = @GrMtlItemSpec,
                                    ReceiptQty = @ReceiptQty,
                                    UomId = @UomId,
                                    UomNo = @UomNo,
                                    InventoryId = @InventoryId,
                                    InventoryNo = @InventoryNo,
                                    LotNumber = @LotNumber,
                                    PoErpPrefix = @PoErpPrefix,
                                    PoErpNo = @PoErpNo,
                                    PoSeq = @PoSeq,
                                    AcceptanceDate = @AcceptanceDate,
                                    AcceptQty = @AcceptQty,
                                    AvailableQty = @AvailableQty,
                                    ReturnQty = @ReturnQty,
                                    OrigUnitPrice = @OrigUnitPrice,
                                    OrigAmount = @OrigAmount,
                                    OrigDiscountAmt = @OrigDiscountAmt,
                                    TrErpPrefix = @TrErpPrefix,
                                    TrErpNo = @TrErpNo,
                                    TrSeq = @TrSeq,
                                    ReceiptExpense = @ReceiptExpense,
                                    DiscountDescription = @DiscountDescription,
                                    Overdue = @Overdue,
                                    QcStatus = @QcStatus,
                                    ConfirmStatus = @ConfirmStatus,
                                    CloseStatus = @CloseStatus,
                                    Remark = @Remark,
                                    InventoryQty = @InventoryQty,
                                    SmallUnit = @SmallUnit,
                                    AvailableDate = @AvailableDate,
                                    ReCheckDate = @ReCheckDate,
                                    ConfirmUser = @ConfirmUser,
                                    ApvErpPrefix = @ApvErpPrefix,
                                    ApvErpNo = @ApvErpNo,
                                    ApvSeq = @ApvSeq,
                                    ProjectCode = @ProjectCode,
                                    PremiumAmountFlag = @PremiumAmountFlag,
                                    OrigPreTaxAmt = @OrigPreTaxAmt,
                                    OrigTaxAmt = @OrigTaxAmt,
                                    PreTaxAmt = @PreTaxAmt,
                                    TaxAmt = @TaxAmt,
                                    PremiumAmount = @PremiumAmount,
                                    PackageUnit = @PackageUnit,
                                    ReserveTaxCode = @ReserveTaxCode,
                                    OrigPremiumAmount = @OrigPremiumAmount,
                                    AvailableUomId = @AvailableUomId,
                                    AvailableUomNo = @AvailableUomNo,
                                    OrigCustomer = @OrigCustomer,
                                    ApproveStatus = @ApproveStatus,
                                    EbcErpPreNo = @EbcErpPreNo,
                                    EbcEdition = @EbcEdition,
                                    ProductSeqAmount = @ProductSeqAmount,
                                    MtlItemType = @MtlItemType,
                                    Loaction = @Loaction,
                                    TaxRate = @TaxRate,
                                    GrErpStatus = @GrErpStatus,
                                    TaxCode = @TaxCode,
                                    DiscountRate = @DiscountRate,
                                    DiscountPrice = @DiscountPrice,
                                    QcType = @QcType,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE GrDetailId = @GrDetailId";
                            rowsAffected += sqlConnection.Execute(sql, updateGrDetail);
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

        #region //UpdateGoodsReceiptManualSynchronize -- 進貨單資料手動同步 -- Ann 2024-02-29
        public string UpdateGoodsReceiptManualSynchronize(string GrErpFullNo, string SyncStartDate, string SyncEndDate
            , string NormalSync, string TranSync)
        {
            try
            {
                List<GoodsReceipt> goodsReceipts = new List<GoodsReceipt>();
                List<GrDetail> grDetails = new List<GrDetail>();

                int mainAffected = 0, detailAffected = 0, mainDelAffected = 0, detailDelAffected = 0, rowsAffected = 0;
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
                            #region //判斷ERP進貨單資料是否存在
                            if (GrErpFullNo.Length > 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM PURTG
                                        WHERE (LTRIM(RTRIM(TG001)) + '-' + LTRIM(RTRIM(TG002))) LIKE '%' + @GrErpFullNo + '%'";
                                dynamicParameters.Add("GrErpFullNo", GrErpFullNo);

                                var resultTsn = sqlConnection.Query(sql, dynamicParameters);
                                if (resultTsn.Count() <= 0) throw new SystemException("ERP進貨單資料錯誤");
                            }
                            #endregion

                            #region //撈取ERP進貨單單頭資料
                            sql = @"SELECT LTRIM(RTRIM(TG001)) GrErpPrefix, LTRIM(RTRIM(TG002)) GrErpNo
                                    , CASE WHEN LEN(LTRIM(RTRIM(TG003))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TG003)) as date), 'yyyy-MM-dd') ELSE NULL END ReceiptDate
                                    , LTRIM(RTRIM(TG005)) SupplierNo, LTRIM(RTRIM(TG006)) SupplierSo, LTRIM(RTRIM(TG007)) CurrencyCode, TG008 Exchange, LTRIM(RTRIM(TG009)) InvoiceType
                                    , LTRIM(RTRIM(TG010)) TaxType, LTRIM(RTRIM(TG011)) InvoiceNo, LTRIM(RTRIM(TG012)) PrintCnt, LTRIM(RTRIM(TG013)) ConfirmStatus
                                    , CASE WHEN LEN(LTRIM(RTRIM(TG014))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TG014)) as date), 'yyyy-MM-dd') ELSE NULL END DocDate
                                    , LTRIM(RTRIM(TG015)) RenewFlag, LTRIM(RTRIM(TG016)) Remark, TG017 TotalAmount, TG018 DeductAmount, TG019 OrigTax, TG020 ReceiptAmount
                                    , LTRIM(RTRIM(TG021)) SupplierName, LTRIM(RTRIM(TG022)) UiNo, LTRIM(RTRIM(TG023)) DeductType, LTRIM(RTRIM(TG024)) ObaccoAndLiquorFlag
                                    , TG025 RowCnt, TG026 Quantity
                                    , CASE WHEN LEN(LTRIM(RTRIM(TG027))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TG027)) as date), 'yyyy-MM-dd') ELSE NULL END InvoiceDate
                                    , TG028 OrigPreTaxAmount, LTRIM(RTRIM(TG029)) ApplyYYMM, TG030 TaxRate, TG031 PreTaxAmount, TG032 TaxAmount, LTRIM(RTRIM(TG033)) PaymentTerm
                                    , LTRIM(RTRIM(TG034)) PoErpPrefix, LTRIM(RTRIM(TG035)) PoErpNo, LTRIM(RTRIM(TG036)) PrepaidErpPrefix, LTRIM(RTRIM(TG037)) PrepaidErpNo
                                    , TG038 OffsetAmount, TG039 TaxOffset, TG040 PackageQuantity, TG041 PremiumAmount, LTRIM(RTRIM(TG042)) ApproveStatus, LTRIM(RTRIM(TG043)) InvoiceFlag
                                    , LTRIM(RTRIM(TG044)) ReserveTaxCode, TG045 SendCount, TG046 OrigPremiumAmount, LTRIM(RTRIM(TG047)) EbcErpPreNo, LTRIM(RTRIM(TG048)) EbcEdition
                                    , LTRIM(RTRIM(TG049)) EbcFlag, LTRIM(RTRIM(TG050)) FromDocType, LTRIM(RTRIM(TG055)) NoticeFlag, LTRIM(RTRIM(TG058)) TaxCode
                                    , LTRIM(RTRIM(TG059)) TradeTerm, LTRIM(RTRIM(TG063)) DetailMultiTax, TG064 TaxExchange, LTRIM(RTRIM(TG067)) ContactUser
                                    , CASE WHEN LEN(LTRIM(RTRIM(CREATE_DATE))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_DATE)) as date), 'yyyy-MM-dd') ELSE NULL END TransferDate
                                    , CASE WHEN LEN(LTRIM(RTRIM(CREATE_TIME))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_TIME)) as datetime), 'HH:mm:ss') ELSE NULL END TransferTime
                                    FROM PURTG
                                    WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "GrErpFullNo", @" AND (LTRIM(RTRIM(TG001)) + '-' + LTRIM(RTRIM(TG002))) LIKE '%' + @GrErpFullNo + '%'", GrErpFullNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SyncStartDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME >= @SyncStartDate OR MODI_DATE + ' ' + MODI_TIME >= @SyncStartDate)", SyncStartDate.Length > 0 ? Convert.ToDateTime(SyncStartDate).ToString("yyyyMMdd 00:00:00") : "");
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SyncEndDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME <= @SyncEndDate OR MODI_DATE + ' ' + MODI_TIME <= @SyncEndDate)", SyncEndDate.Length > 0 ? Convert.ToDateTime(SyncEndDate).ToString("yyyyMMdd 23:59:59") : ""); sql += @" ORDER BY TransferDate, TransferTime";

                            goodsReceipts = sqlConnection.Query<GoodsReceipt>(sql, dynamicParameters).ToList();
                            #endregion

                            #region //撈取ERP進貨單身資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(TH001)) GrErpPrefix, LTRIM(RTRIM(TH002)) GrErpNo, LTRIM(RTRIM(TH003)) GrSeq
                                    , LTRIM(RTRIM(TH004)) MtlItemNo, LTRIM(RTRIM(TH005)) GrMtlItemName, LTRIM(RTRIM(TH006)) GrMtlItemSpec, TH007 ReceiptQty
                                    , LTRIM(RTRIM(TH008)) UomNo, LTRIM(RTRIM(TH009)) InventoryNo, LTRIM(RTRIM(TH010)) LotNumber
                                    , LTRIM(RTRIM(TH011)) PoErpPrefix, LTRIM(RTRIM(TH012)) PoErpNo, LTRIM(RTRIM(TH013)) PoSeq
                                    , CASE WHEN LEN(LTRIM(RTRIM(TH014))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TH014)) as date), 'yyyy-MM-dd') ELSE NULL END AcceptanceDate
                                    , TH015 AcceptQty, TH016 AvailableQty, TH017 ReturnQty, TH018 OrigUnitPrice, TH019 OrigAmount, TH020 OrigDiscountAmt
                                    , LTRIM(RTRIM(TH021)) TrErpPrefix, LTRIM(RTRIM(TH022)) TrErpNo, LTRIM(RTRIM(TH023)) TrSeq
                                    , TH024 ReceiptExpense, LTRIM(RTRIM(TH025)) DiscountDescription, LTRIM(RTRIM(TH026)) PaymentHold, LTRIM(RTRIM(TH027)) Overdue
                                    , LTRIM(RTRIM(TH028)) QcStatus, LTRIM(RTRIM(TH029)) ReturnCode, LTRIM(RTRIM(TH030)) ConfirmStatus, LTRIM(RTRIM(TH031)) CloseStatus
                                    , LTRIM(RTRIM(TH032)) ReNewStatus, LTRIM(RTRIM(TH033)) Remark, TH034 InventoryQty, LTRIM(RTRIM(TH035)) SmallUnit
                                    , CASE WHEN LEN(LTRIM(RTRIM(TH036))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TH036)) as date), 'yyyy-MM-dd') ELSE NULL END AvailableDate
                                    , CASE WHEN LEN(LTRIM(RTRIM(TH037))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TH037)) as date), 'yyyy-MM-dd') ELSE NULL END ReCheckDate
                                    , LTRIM(RTRIM(TH038)) ConfirmUserNo, LTRIM(RTRIM(TH039)) ApvErpPrefix, LTRIM(RTRIM(TH040)) ApvErpNo, LTRIM(RTRIM(TH041)) ApvSeq
                                    , LTRIM(RTRIM(TH042)) ProjectCode, LTRIM(RTRIM(TH043)) ExpenseEntry, LTRIM(RTRIM(TH044)) PremiumAmountFlag, TH045 OrigPreTaxAmt
                                    , TH046 OrigTaxAmt, TH047 PreTaxAmt, TH048 TaxAmt, TH049 ReceiptPackageQty, TH050 AcceptancePackageQty, TH051 ReturnPackageQty
                                    , TH052 PremiumAmount, LTRIM(RTRIM(TH053)) PackageUnit, LTRIM(RTRIM(TH054)) ReserveTaxCode, TH055 OrigPremiumAmount
                                    , LTRIM(RTRIM(TH056)) AvailableUomNo, LTRIM(RTRIM(TH057)) OrigCustomer, LTRIM(RTRIM(TH058)) ApproveStatus
                                    , LTRIM(RTRIM(TH059)) EbcErpPreNo, LTRIM(RTRIM(TH060)) EbcEdition, TH061 ProductSeqAmount, LTRIM(RTRIM(TH062)) MtlItemType
                                    , LTRIM(RTRIM(TH063)) Loaction, TH073 TaxRate, LTRIM(RTRIM(TH074)) MultipleFlag, LTRIM(RTRIM(TH082)) GrErpStatus, LTRIM(RTRIM(TH088)) TaxCode
                                    , TH089 DiscountRate, TH090 DiscountPrice, LTRIM(RTRIM(TH091)) QcType
                                    , CASE WHEN LEN(LTRIM(RTRIM(CREATE_DATE))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_DATE)) as date), 'yyyy-MM-dd') ELSE NULL END TransferDate
                                    , CASE WHEN LEN(LTRIM(RTRIM(CREATE_TIME))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_TIME)) as datetime), 'HH:mm:ss') ELSE NULL END TransferTime
                                    FROM PURTH
                                    WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "GrErpFullNo", @" AND (LTRIM(RTRIM(TH001)) + '-' + LTRIM(RTRIM(TH002))) LIKE '%' + @GrErpFullNo + '%'", GrErpFullNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SyncStartDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME >= @SyncStartDate OR MODI_DATE + ' ' + MODI_TIME >= @SyncStartDate)", SyncStartDate.Length > 0 ? Convert.ToDateTime(SyncStartDate).ToString("yyyyMMdd 00:00:00") : "");
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SyncEndDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME <= @SyncEndDate OR MODI_DATE + ' ' + MODI_TIME <= @SyncEndDate)", SyncEndDate.Length > 0 ? Convert.ToDateTime(SyncEndDate).ToString("yyyyMMdd 23:59:59") : "");
                            sql += @" ORDER BY TransferDate, TransferTime";

                            grDetails = sqlConnection.Query<GrDetail>(sql, dynamicParameters).ToList();
                            #endregion
                        }

                        using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                        {
                            #region //取得MES供應商資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT SupplierId, SupplierNo
                                    FROM SCM.Supplier
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<Supplier> resultSuppliers = sqlConnection.Query<Supplier>(sql, dynamicParameters).ToList();

                            goodsReceipts = goodsReceipts.Join(resultSuppliers, x => x.SupplierNo, y => y.SupplierNo, (x, y) => { x.SupplierId = y.SupplierId; return x; }).ToList();
                            #endregion

                            #region //撈取單位ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT UomId, UomNo
                                    FROM PDM.UnitOfMeasure
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<Uom> resultSoPriceUomrNos = sqlConnection.Query<Uom>(sql, dynamicParameters).ToList();

                            grDetails = grDetails.Join(resultSoPriceUomrNos, x => x.UomNo.ToUpper(), y => y.UomNo, (x, y) => { x.UomId = y.UomId; return x; }).ToList();
                            grDetails = grDetails.Join(resultSoPriceUomrNos, x => x.AvailableUomNo.ToUpper(), y => y.UomNo, (x, y) => { x.AvailableUomId = y.UomId; return x; }).ToList();
                            #endregion

                            #region //撈取品號ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT MtlItemId, MtlItemNo
                                FROM PDM.MtlItem
                                WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<MtlItem> resultMtlitems = sqlConnection.Query<MtlItem>(sql, dynamicParameters).ToList();

                            grDetails = grDetails.Join(resultMtlitems, x => x.MtlItemNo, y => y.MtlItemNo, (x, y) => { x.MtlItemId = y.MtlItemId; return x; }).Select(x => x).ToList();
                            #endregion

                            #region //撈取庫別ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT InventoryId, InventoryNo
                                FROM SCM.Inventory
                                WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<Inventory> resultInventories = sqlConnection.Query<Inventory>(sql, dynamicParameters).ToList();

                            grDetails = grDetails.Join(resultInventories, x => x.InventoryNo, y => y.InventoryNo, (x, y) => { x.InventoryId = y.InventoryId; return x; }).ToList();
                            #endregion

                            #region //撈取確認者ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.UserId, a.UserNo 
                                    FROM BAS.[User] a";

                            List<User> resultConfirmUsers = sqlConnection.Query<User>(sql, dynamicParameters).ToList();

                            grDetails = grDetails.Join(resultConfirmUsers, x => x.ConfirmUserNo, y => y.UserNo, (x, y) => { x.ConfirmUser = y.UserId; return x; }).ToList();
                            #endregion

                            #region //判斷SCM.GoodsReceipt是否有資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT GrId, GrErpPrefix, GrErpNo
                                    FROM SCM.GoodsReceipt
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<GoodsReceipt> resultGoodsReceipt = sqlConnection.Query<GoodsReceipt>(sql, dynamicParameters).ToList();

                            goodsReceipts = goodsReceipts.GroupJoin(resultGoodsReceipt, x => new { x.GrErpPrefix, x.GrErpNo }, y => new { y.GrErpPrefix, y.GrErpNo }, (x, y) => { x.GrId = y.FirstOrDefault()?.GrId; return x; }).Select(x => x).ToList();
                            #endregion

                            #region //進貨單單頭(新增/修改)
                            List<GoodsReceipt> addGoodsReceipts = goodsReceipts.Where(x => x.GrId == null).ToList();
                            List<GoodsReceipt> updateGoodsReceipts = goodsReceipts.Where(x => x.GrId != null).ToList();

                            #region //新增
                            dynamicParameters = new DynamicParameters();
                            if (addGoodsReceipts.Count > 0)
                            {
                                addGoodsReceipts
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

                                sql = @"INSERT INTO SCM.GoodsReceipt (CompanyId, GrErpPrefix, GrErpNo, ReceiptDate, SupplierId, SupplierNo
                                        , SupplierSo, CurrencyCode, Exchange, InvoiceType, TaxType, InvoiceNo, PrintCnt, ConfirmStatus, DocDate
                                        , RenewFlag, Remark, TotalAmount, DeductAmount, OrigTax, ReceiptAmount, SupplierName, UiNo, DeductType
                                        , ObaccoAndLiquorFlag, RowCnt, Quantity, InvoiceDate, OrigPreTaxAmount, ApplyYYMM, TaxRate, PreTaxAmount, TaxAmount
                                        , PaymentTerm, PoId, PoErpPrefix, PoErpNo, PrepaidErpPrefix, PrepaidErpNo, OffsetAmount, TaxOffset, PackageQuantity
                                        , PremiumAmount, ApproveStatus, InvoiceFlag, ReserveTaxCode, SendCount, OrigPremiumAmount, EbcErpPreNo, EbcEdition, EbcFlag
                                        , FromDocType, NoticeFlag, TaxCode, TradeTerm, DetailMultiTax, TaxExchange, ContactUser, TransferStatus, TransferDate
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        VALUES (@CompanyId, @GrErpPrefix, @GrErpNo, @ReceiptDate, @SupplierId, @SupplierNo
                                        , @SupplierSo, @CurrencyCode, @Exchange, @InvoiceType, @TaxType, @InvoiceNo, @PrintCnt, @ConfirmStatus, @DocDate
                                        , @RenewFlag, @Remark, @TotalAmount, @DeductAmount, @OrigTax, @ReceiptAmount, @SupplierName, @UiNo, @DeductType
                                        , @ObaccoAndLiquorFlag, @RowCnt, @Quantity, @InvoiceDate, @OrigPreTaxAmount, @ApplyYYMM, @TaxRate, @PreTaxAmount, @TaxAmount
                                        , @PaymentTerm, @PoId, @PoErpPrefix, @PoErpNo, @PrepaidErpPrefix, @PrepaidErpNo, @OffsetAmount, @TaxOffset, @PackageQuantity
                                        , @PremiumAmount, @ApproveStatus, @InvoiceFlag, @ReserveTaxCode, @SendCount, @OrigPremiumAmount, @EbcErpPreNo, @EbcEdition, @EbcFlag
                                        , @FromDocType, @NoticeFlag, @TaxCode, @TradeTerm, @DetailMultiTax, @TaxExchange, @ContactUser, @TransferStatus, @TransferDate
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                mainAffected += sqlConnection.Execute(sql, addGoodsReceipts);
                            }
                            #endregion

                            #region //修改
                            dynamicParameters = new DynamicParameters();
                            if (updateGoodsReceipts.Count > 0)
                            {
                                updateGoodsReceipts
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.LastModifiedDate = LastModifiedDate;
                                        x.LastModifiedBy = LastModifiedBy;
                                    });

                                sql = @"UPDATE SCM.GoodsReceipt SET
                                        ReceiptDate = @ReceiptDate,
                                        SupplierId = @SupplierId,
                                        SupplierNo = @SupplierNo,
                                        CurrencyCode = @CurrencyCode,
                                        Exchange = @Exchange,
                                        InvoiceType = @InvoiceType,
                                        TaxType = @TaxType,
                                        InvoiceNo = @InvoiceNo,
                                        ConfirmStatus = @ConfirmStatus,
                                        DocDate = @DocDate,
                                        Remark = @Remark,
                                        TotalAmount = @TotalAmount,
                                        DeductAmount = @DeductAmount,
                                        OrigTax = @OrigTax,
                                        ReceiptAmount = @ReceiptAmount,
                                        SupplierName = @SupplierName,
                                        UiNo = @UiNo,
                                        DeductType = @DeductType,
                                        ObaccoAndLiquorFlag = @ObaccoAndLiquorFlag,
                                        Quantity = @Quantity,
                                        InvoiceDate = @InvoiceDate,
                                        OrigPreTaxAmount = @OrigPreTaxAmount,
                                        ApplyYYMM = @ApplyYYMM,
                                        TaxRate = @TaxRate,
                                        PreTaxAmount = @PreTaxAmount,
                                        TaxAmount = @TaxAmount,
                                        PaymentTerm = @PaymentTerm,
                                        OffsetAmount = @OffsetAmount,
                                        TaxOffset = @TaxOffset,
                                        PremiumAmount = @PremiumAmount,
                                        ApproveStatus = @ApproveStatus,
                                        OrigPremiumAmount = @OrigPremiumAmount,
                                        FromDocType = @FromDocType,
                                        TaxCode = @TaxCode,
                                        TradeTerm = @TradeTerm,
                                        DetailMultiTax = @DetailMultiTax,
                                        ContactUser = @ContactUser,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE GrId = @GrId";
                                mainAffected += sqlConnection.Execute(sql, updateGoodsReceipts);
                            }
                            #endregion
                            #endregion

                            #region //進貨單單身(新增/修改)
                            #region //撈取進貨單單頭ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT GrId, GrErpPrefix, GrErpNo
                                    FROM  SCM.GoodsReceipt
                                    WHERE CompanyId = @CompanyId
                                    ORDER BY GrId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            resultGoodsReceipt = sqlConnection.Query<GoodsReceipt>(sql, dynamicParameters).ToList();

                            grDetails = grDetails.Join(resultGoodsReceipt, x => new { x.GrErpPrefix, x.GrErpNo }, y => new { y.GrErpPrefix, y.GrErpNo }, (x, y) => { x.GrId = y.GrId; return x; }).Select(x => x).ToList();
                            #endregion

                            #region //判斷SCM.GrDetail是否有資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.GrDetailId, a.GrId, a.GrSeq
                                    FROM SCM.GrDetail a
                                    INNER JOIN SCM.GoodsReceipt b ON a.GrId = b.GrId
                                    WHERE b.CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<GrDetail> resultGrDetail = sqlConnection.Query<GrDetail>(sql, dynamicParameters).ToList();

                            grDetails = grDetails.GroupJoin(resultGrDetail, x => new { x.GrId, x.GrSeq }, y => new { y.GrId, y.GrSeq }, (x, y) => { x.GrDetailId = y.FirstOrDefault()?.GrDetailId; return x; }).ToList();
                            #endregion

                            List<GrDetail> addGrDetail = grDetails.Where(x => x.GrDetailId == null).ToList();
                            List<GrDetail> updateGrDetail = grDetails.Where(x => x.GrDetailId != null).ToList();

                            #region //新增
                            dynamicParameters = new DynamicParameters();
                            if (addGrDetail.Count > 0)
                            {
                                addGrDetail
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.TransferStatus = "Y";
                                        x.CreateDate = CreateDate;
                                        x.LastModifiedDate = LastModifiedDate;
                                        x.CreateBy = CreateBy;
                                        x.LastModifiedBy = LastModifiedBy;
                                    });

                                sql = @"INSERT INTO SCM.GrDetail (GrId, GrErpPrefix, GrErpNo, GrSeq, MtlItemId, MtlItemNo, GrMtlItemName, GrMtlItemSpec
                                        , ReceiptQty, UomId, UomNo, InventoryId, InventoryNo, LotNumber, PoErpPrefix, PoErpNo, PoSeq, AcceptanceDate
                                        , AcceptQty, AvailableQty, ReturnQty, OrigUnitPrice, OrigAmount, OrigDiscountAmt, TrErpPrefix, TrErpNo, TrSeq
                                        , ReceiptExpense, DiscountDescription, PaymentHold, Overdue, QcStatus, ReturnCode, ConfirmStatus, CloseStatus
                                        , ReNewStatus, Remark, InventoryQty, SmallUnit, AvailableDate, ReCheckDate, ConfirmUser, ApvErpPrefix, ApvErpNo
                                        , ApvSeq, ProjectCode, ExpenseEntry, PremiumAmountFlag, OrigPreTaxAmt, OrigTaxAmt, PreTaxAmt, TaxAmt, ReceiptPackageQty
                                        , AcceptancePackageQty, ReturnPackageQty, PremiumAmount, PackageUnit, ReserveTaxCode, OrigPremiumAmount, AvailableUomId
                                        , AvailableUomNo, OrigCustomer, ApproveStatus, EbcErpPreNo, EbcEdition, ProductSeqAmount, MtlItemType, Loaction
                                        , TaxRate, MultipleFlag, GrErpStatus, TaxCode, DiscountRate, DiscountPrice, QcType, TransferStatus, TransferDate
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        VALUES (@GrId, @GrErpPrefix, @GrErpNo, @GrSeq, @MtlItemId, @MtlItemNo, @GrMtlItemName, @GrMtlItemSpec
                                        , @ReceiptQty, @UomId, @UomNo, @InventoryId, @InventoryNo, @LotNumber, @PoErpPrefix, @PoErpNo, @PoSeq, @AcceptanceDate
                                        , @AcceptQty, @AvailableQty, @ReturnQty, @OrigUnitPrice, @OrigAmount, @OrigDiscountAmt, @TrErpPrefix, @TrErpNo, @TrSeq
                                        , @ReceiptExpense, @DiscountDescription, @PaymentHold, @Overdue, @QcStatus, @ReturnCode, @ConfirmStatus, @CloseStatus
                                        , @ReNewStatus, @Remark, @InventoryQty, @SmallUnit, @AvailableDate, @ReCheckDate, @ConfirmUser, @ApvErpPrefix, @ApvErpNo
                                        , @ApvSeq, @ProjectCode, @ExpenseEntry, @PremiumAmountFlag, @OrigPreTaxAmt, @OrigTaxAmt, @PreTaxAmt, @TaxAmt, @ReceiptPackageQty
                                        , @AcceptancePackageQty, @ReturnPackageQty, @PremiumAmount, @PackageUnit, @ReserveTaxCode, @OrigPremiumAmount, @AvailableUomId
                                        , @AvailableUomNo, @OrigCustomer, @ApproveStatus, @EbcErpPreNo, @EbcEdition, @ProductSeqAmount, @MtlItemType, @Loaction
                                        , @TaxRate, @MultipleFlag, @GrErpStatus, @TaxCode, @DiscountRate, @DiscountPrice, @QcType, @TransferStatus, @TransferDate
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                detailAffected += sqlConnection.Execute(sql, addGrDetail);
                            }
                            #endregion

                            #region //修改
                            dynamicParameters = new DynamicParameters();
                            if (updateGrDetail.Count > 0)
                            {
                                updateGrDetail
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.LastModifiedDate = LastModifiedDate;
                                        x.LastModifiedBy = LastModifiedBy;
                                    });

                                sql = @"UPDATE SCM.GrDetail SET
                                        MtlItemId = @MtlItemId,
                                        MtlItemNo = @MtlItemNo,
                                        GrMtlItemName = @GrMtlItemName,
                                        GrMtlItemSpec = @GrMtlItemSpec,
                                        ReceiptQty = @ReceiptQty,
                                        UomId = @UomId,
                                        UomNo = @UomNo,
                                        InventoryId = @InventoryId,
                                        InventoryNo = @InventoryNo,
                                        LotNumber = @LotNumber,
                                        PoErpPrefix = @PoErpPrefix,
                                        PoErpNo = @PoErpNo,
                                        PoSeq = @PoSeq,
                                        AcceptanceDate = @AcceptanceDate,
                                        AcceptQty = @AcceptQty,
                                        AvailableQty = @AvailableQty,
                                        ReturnQty = @ReturnQty,
                                        OrigUnitPrice = @OrigUnitPrice,
                                        OrigAmount = @OrigAmount,
                                        OrigDiscountAmt = @OrigDiscountAmt,
                                        TrErpPrefix = @TrErpPrefix,
                                        TrErpNo = @TrErpNo,
                                        TrSeq = @TrSeq,
                                        ReceiptExpense = @ReceiptExpense,
                                        DiscountDescription = @DiscountDescription,
                                        Overdue = @Overdue,
                                        QcStatus = @QcStatus,
                                        ConfirmStatus = @ConfirmStatus,
                                        CloseStatus = @CloseStatus,
                                        Remark = @Remark,
                                        InventoryQty = @InventoryQty,
                                        SmallUnit = @SmallUnit,
                                        AvailableDate = @AvailableDate,
                                        ReCheckDate = @ReCheckDate,
                                        ConfirmUser = @ConfirmUser,
                                        ApvErpPrefix = @ApvErpPrefix,
                                        ApvErpNo = @ApvErpNo,
                                        ApvSeq = @ApvSeq,
                                        ProjectCode = @ProjectCode,
                                        PremiumAmountFlag = @PremiumAmountFlag,
                                        OrigPreTaxAmt = @OrigPreTaxAmt,
                                        OrigTaxAmt = @OrigTaxAmt,
                                        PreTaxAmt = @PreTaxAmt,
                                        TaxAmt = @TaxAmt,
                                        PremiumAmount = @PremiumAmount,
                                        PackageUnit = @PackageUnit,
                                        ReserveTaxCode = @ReserveTaxCode,
                                        OrigPremiumAmount = @OrigPremiumAmount,
                                        AvailableUomId = @AvailableUomId,
                                        AvailableUomNo = @AvailableUomNo,
                                        OrigCustomer = @OrigCustomer,
                                        ApproveStatus = @ApproveStatus,
                                        EbcErpPreNo = @EbcErpPreNo,
                                        EbcEdition = @EbcEdition,
                                        ProductSeqAmount = @ProductSeqAmount,
                                        MtlItemType = @MtlItemType,
                                        Loaction = @Loaction,
                                        TaxRate = @TaxRate,
                                        GrErpStatus = @GrErpStatus,
                                        TaxCode = @TaxCode,
                                        DiscountRate = @DiscountRate,
                                        DiscountPrice = @DiscountPrice,
                                        QcType = @QcType,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE GrDetailId = @GrDetailId";
                                detailAffected += sqlConnection.Execute(sql, updateGrDetail);
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
                            #region //ERP進貨單頭資料
                            sql = @"SELECT TG001 GrErpPrefix, TG002 GrErpNo
                                    FROM PURTG
                                    WHERE 1=1
                                    ORDER BY TG001, TG002";
                            var resultErpGr = erpConnection.Query<GoodsReceipt>(sql, dynamicParameters);
                            #endregion

                            using (SqlConnection bmConnection = new SqlConnection(MainConnectionStrings))
                            {
                                #region //BM進貨單單頭資料
                                sql = @"SELECT GrErpPrefix, GrErpNo
                                        FROM SCM.GoodsReceipt
                                        WHERE CompanyId = @CompanyId
                                        ORDER BY GrErpPrefix, GrErpNo";
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                var resultBmGr = bmConnection.Query<GoodsReceipt>(sql, dynamicParameters);
                                #endregion

                                #region //篩選ERP不存在但BM還存在的進貨單單頭
                                var dictionaryErpGr = resultErpGr.ToDictionary(x => x.GrErpPrefix + "-" + x.GrErpNo, x => x.GrErpPrefix + "-" + x.GrErpNo);
                                var dictionaryBmGr = resultBmGr.ToDictionary(x => x.GrErpPrefix + "-" + x.GrErpNo, x => x.GrErpPrefix + "-" + x.GrErpNo);
                                var changeGr = dictionaryBmGr.Where(x => !dictionaryErpGr.ContainsKey(x.Key)).ToList();
                                var changeGrList = changeGr.Select(x => x.Value).ToList();
                                #endregion

                                #region //刪除異動進貨單單頭
                                if (changeGrList.Count > 0)
                                {
                                    #region //單身刪除
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE SCM.GrDetail
                                            WHERE GrId IN (
                                                SELECT GrId
                                                FROM SCM.GoodsReceipt
                                                WHERE GrErpPrefix + '-' + GrErpNo IN @GrErpFullNo
                                            )";
                                    dynamicParameters.Add("GrErpFullNo", changeGrList.Select(x => x).ToArray());
                                    mainDelAffected += bmConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //單頭刪除
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE SCM.GoodsReceipt
                                            WHERE GrErpPrefix + '-' + GrErpNo IN @GrErpFullNo";
                                    dynamicParameters.Add("GrErpFullNo", changeGrList.Select(x => x).ToArray());
                                    detailDelAffected += bmConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                                #endregion
                            }

                            #region //ERP進貨單單身資料
                            sql = @"SELECT TH001 GrErpPrefix, TH002 GrErpNo, TH003 GrSeq
                                    FROM PURTH
                                    WHERE 1=1
                                    ORDER BY TH001, TH002, TH003";
                            var resultErpGrDetail = erpConnection.Query<GrDetail>(sql, dynamicParameters);
                            #endregion

                            using (SqlConnection bmConnection = new SqlConnection(MainConnectionStrings))
                            {
                                #region //BM進貨單身資料
                                sql = @"SELECT b.GrErpPrefix, b.GrErpNo, a.GrSeq
                                        FROM SCM.GrDetail a
                                        INNER JOIN SCM.GoodsReceipt b ON a.GrId = b.GrId
                                        WHERE b.CompanyId = @CompanyId
                                        ORDER BY b.GrErpPrefix, b.GrErpNo, a.GrSeq";
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                var resultBmGrDetail = bmConnection.Query<GrDetail>(sql, dynamicParameters);
                                #endregion

                                #region //篩選ERP不存在但BM還存在的進貨單單身
                                var dictionaryErpGrDetail = resultErpGrDetail.ToDictionary(x => x.GrErpPrefix + "-" + x.GrErpNo + "-" + x.GrSeq, x => x.GrErpPrefix + "-" + x.GrErpNo + "-" + x.GrSeq);
                                var dictionaryBmGrDetail = resultBmGrDetail.ToDictionary(x => x.GrErpPrefix + "-" + x.GrErpNo + "-" + x.GrSeq, x => x.GrErpPrefix + "-" + x.GrErpNo + "-" + x.GrSeq);
                                var changeGrDetail = dictionaryBmGrDetail.Where(x => !dictionaryErpGrDetail.ContainsKey(x.Key)).ToList();
                                var changeGrDetailList = changeGrDetail.Select(x => x.Value).ToList();
                                #endregion

                                #region //刪除進貨單身
                                if (changeGrDetailList.Count > 0)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE a
                                            FROM SCM.GrDetail a
                                            INNER JOIN SCM.GoodsReceipt b ON a.GrId = b.GrId
                                            WHERE b.GrErpPrefix + '-' + b.GrErpNo + '-' + RIGHT('0000' + a.GrSeq, 4) IN @GrErpFullNo";
                                    dynamicParameters.Add("GrErpFullNo", changeGrDetailList.Select(x => x).ToArray());
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
                        data = "已更新資料<br />【" + mainAffected + "】筆單頭<br />【" + detailAffected + "】筆單身<br />刪除<br />【" + mainDelAffected + "】筆單頭<br />【" + detailDelAffected + "】筆單身"
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

        #region //UpdateGrTotal -- 更新進貨單頭統整資料 -- Ann 2024-03-07
        public string UpdateGrTotal(int GrId, double? Exchange, SqlConnection sqlConnection, SqlConnection sqlConnection2)
        {
            try
            {
                int rowsAffected = 0;
                int OriUnitDecimal = 0; //原幣單價取位
                int OriPriceDecimal = 0; //原幣金額取位
                int UnitDecimal = 0; //本幣單價取位
                int PriceDecimal = 0; //本幣金額取位

                double TotalAmount = 0;
                double DeductAmount = 0;
                double OrigTax = 0;
                double ReceiptAmount = 0;
                double Quantity = 0;
                double OrigPreTaxAmount = 0;
                double PreTaxAmount = 0;
                double TaxAmount = 0;

                #region //確認進貨單頭資料
                dynamicParameters = new DynamicParameters();
                sql = @"SELECT a.CurrencyCode, a.Exchange, a.TaxRate, a.TaxType
                        FROM SCM.GoodsReceipt a 
                        WHERE a.GrId = @GrId";
                dynamicParameters.Add("GrId", GrId);

                var GoodsReceiptResult = sqlConnection.Query(sql, dynamicParameters);

                if (GoodsReceiptResult.Count() <= 0) throw new SystemException("進貨單資料錯誤!!");

                string OriCurrencyCode = "";
                double OriExchange = -1;
                double TaxRate = -1;
                string TaxType = "";
                foreach (var item in GoodsReceiptResult)
                {
                    OriCurrencyCode = item.CurrencyCode;
                    OriExchange = item.Exchange;
                    TaxRate = item.TaxRate;
                    TaxType = item.TaxType;
                    if (Exchange < 0) Exchange = OriExchange;
                }
                #endregion

                #region //取得進貨單身資料
                dynamicParameters = new DynamicParameters();
                sql = @"SELECT a.GrDetailId, a.AvailableQty, a.OrigUnitPrice, a.OrigDiscountAmt, a.OrigAmount
                        FROM SCM.GrDetail a 
                        WHERE a.GrId = @GrId";
                dynamicParameters.Add("GrId", GrId);

                var GrDetailResult = sqlConnection.Query(sql, dynamicParameters);
                #endregion

                #region //原幣小數點取位
                dynamicParameters = new DynamicParameters();
                sql = @"SELECT a.MF003, a.MF004
                        FROM CMSMF a 
                        WHERE a.MF001 = @Currency";
                dynamicParameters.Add("Currency", OriCurrencyCode);

                var CMSMFResult = sqlConnection2.Query(sql, dynamicParameters);

                if (CMSMFResult.Count() <= 0) throw new SystemException("金額小數點取位資料錯誤!!");

                foreach (var item in CMSMFResult)
                {
                    OriUnitDecimal = Convert.ToInt32(item.MF003);
                    OriPriceDecimal = Convert.ToInt32(item.MF004);
                }
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

                #region //本幣小數點取位
                dynamicParameters = new DynamicParameters();
                sql = @"SELECT a.MF003, a.MF004
                        FROM CMSMF a 
                        WHERE a.MF001 = @Currency";
                dynamicParameters.Add("Currency", CurrencyCode);

                var CMSMFResult2 = sqlConnection2.Query(sql, dynamicParameters);

                if (CMSMFResult2.Count() <= 0) throw new SystemException("金額小數點取位資料錯誤!!");

                foreach (var item in CMSMFResult2)
                {
                    UnitDecimal = Convert.ToInt32(item.MF003);
                    PriceDecimal = Convert.ToInt32(item.MF004);
                }
                #endregion

                #region //更新單身金額
                foreach (var item in GrDetailResult)
                {
                    JObject calculateTaxAmtResult = JObject.Parse(CalculateTaxAmt(item.AvailableQty, item.OrigUnitPrice, item.OrigDiscountAmt, item.OrigAmount, TaxRate, TaxType, Exchange
                        , OriUnitDecimal, OriPriceDecimal, UnitDecimal, PriceDecimal));

                    if (calculateTaxAmtResult["status"].ToString() == "success")
                    {
                        double OrigPreTaxAmt = Convert.ToDouble(calculateTaxAmtResult["OrigPreTaxAmt"]);
                        double OrigTaxAmt = Convert.ToDouble(calculateTaxAmtResult["OrigTaxAmt"]);
                        double PreTaxAmt = Convert.ToDouble(calculateTaxAmtResult["PreTaxAmt"]);
                        double TaxAmt = Convert.ToDouble(calculateTaxAmtResult["TaxAmt"]);
                        double OrigAmount = Convert.ToDouble(calculateTaxAmtResult["OrigAmount"]);

                        #region //Update SCM.GrDetail
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.GrDetail SET
                                OrigPreTaxAmt = @OrigPreTaxAmt,
                                OrigTaxAmt = @OrigTaxAmt,
                                PreTaxAmt = @PreTaxAmt,
                                TaxAmt = @TaxAmt,
                                OrigAmount = @OrigAmount,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE GrDetailId = @GrDetailId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                OrigPreTaxAmt,
                                OrigTaxAmt,
                                PreTaxAmt,
                                TaxAmt,
                                OrigAmount,
                                LastModifiedDate,
                                LastModifiedBy,
                                item.GrDetailId
                            });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion
                    }
                    else
                    {
                        throw new SystemException(calculateTaxAmtResult["msg"].ToString());
                    }
                }
                #endregion

                dynamicParameters = new DynamicParameters();
                sql = @"SELECT SUM(a.OrigAmount) TotalAmount, SUM(a.OrigDiscountAmt) DeductAmount
                        , SUM(a.OrigTaxAmt) OrigTax, SUM(a.ReceiptExpense) ReceiptAmount
                        , SUM(a.AvailableQty) Quantity, SUM(a.OrigPreTaxAmt) OrigPreTaxAmount
                        , SUM(a.PreTaxAmt) PreTaxAmount, SUM(a.TaxAmt) TaxAmount
                        FROM SCM.GrDetail a 
                        WHERE a.GrId = @GrId";
                dynamicParameters.Add("GrId", GrId);

                var GrDetailQtyResult = sqlConnection.Query(sql, dynamicParameters);

                foreach (var item in GrDetailQtyResult)
                {
                    TotalAmount = Convert.ToDouble(item.TotalAmount);
                    DeductAmount = Convert.ToDouble(item.DeductAmount);
                    OrigTax = Convert.ToDouble(item.OrigTax);
                    ReceiptAmount = Convert.ToDouble(item.ReceiptAmount);
                    Quantity = Convert.ToDouble(item.Quantity);
                    OrigPreTaxAmount = Convert.ToDouble(item.OrigPreTaxAmount);
                    PreTaxAmount = Convert.ToDouble(item.PreTaxAmount);
                    TaxAmount = Convert.ToDouble(item.TaxAmount);
                }

                #region //Update SCM.GoodsReceipt
                dynamicParameters = new DynamicParameters();
                sql = @"UPDATE SCM.GoodsReceipt SET
                        Exchange = @Exchange,
                        TotalAmount = @TotalAmount,
                        DeductAmount = @DeductAmount,
                        OrigTax = @OrigTax,
                        ReceiptAmount = @ReceiptAmount,
                        Quantity = @Quantity,
                        OrigPreTaxAmount = @OrigPreTaxAmount,
                        PreTaxAmount = @PreTaxAmount,
                        TaxAmount = @TaxAmount,
                        LastModifiedDate = @LastModifiedDate,
                        LastModifiedBy = @LastModifiedBy
                        WHERE GrId = @GrId";
                dynamicParameters.AddDynamicParams(
                    new
                    {
                        Exchange,
                        TotalAmount = Math.Round(TotalAmount, OriPriceDecimal),
                        DeductAmount = Math.Round(DeductAmount, OriPriceDecimal),
                        OrigTax = Math.Round(OrigTax, OriPriceDecimal),
                        ReceiptAmount = Math.Round(ReceiptAmount, PriceDecimal),
                        Quantity,
                        OrigPreTaxAmount = Math.Round(OrigPreTaxAmount, OriPriceDecimal),
                        PreTaxAmount = Math.Round(PreTaxAmount, PriceDecimal),
                        TaxAmount = Math.Round(TaxAmount, PriceDecimal),
                        LastModifiedDate,
                        LastModifiedBy,
                        GrId
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
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //UpdateExchange -- 更新匯率 -- Ann 2024-03-11
        public string UpdateExchange(int GrId, double? Exchange, SqlConnection sqlConnection, SqlConnection sqlConnection2)
        {
            try
            {
                int rowsAffected = 0;
                int OriUnitDecimal = 0; //原幣單價取位
                int OriPriceDecimal = 0; //原幣金額取位
                int UnitDecimal = 0; //本幣單價取位
                int PriceDecimal = 0; //本幣金額取位

                #region //確認進貨單頭資料
                dynamicParameters = new DynamicParameters();
                sql = @"SELECT a.CurrencyCode, a.TaxType, a.TaxRate
                        FROM SCM.GoodsReceipt a 
                        WHERE a.GrId = @GrId";
                dynamicParameters.Add("GrId", GrId);

                var GoodsReceiptResult = sqlConnection.Query(sql, dynamicParameters);

                if (GoodsReceiptResult.Count() <= 0) throw new SystemException("進貨單資料錯誤!!");

                string TaxType = "";
                double TaxRate = 0;
                string OriCurrencyCode = "";
                foreach (var item in GoodsReceiptResult)
                {
                    TaxType = item.TaxType;
                    TaxRate = item.TaxRate;
                    OriCurrencyCode = item.CurrencyCode;
                }
                #endregion

                #region //取得進貨單身資料
                dynamicParameters = new DynamicParameters();
                sql = @"SELECT a.GrDetailId, a.AvailableQty, a.OrigUnitPrice, a.OrigDiscountAmt
                        , a.OrigAmount
                        FROM SCM.GrDetail a
                        WHERE a.GrId = @GrId";
                dynamicParameters.Add("GrId", GrId);

                List<GrDetail> grDetails = sqlConnection.Query<GrDetail>(sql, dynamicParameters).ToList();
                #endregion

                #region //原幣小數點取位
                dynamicParameters = new DynamicParameters();
                sql = @"SELECT a.MF003, a.MF004
                        FROM CMSMF a 
                        WHERE a.MF001 = @Currency";
                dynamicParameters.Add("Currency", OriCurrencyCode);

                var CMSMFResult = sqlConnection2.Query(sql, dynamicParameters);

                if (CMSMFResult.Count() <= 0) throw new SystemException("金額小數點取位資料錯誤!!");

                foreach (var item in CMSMFResult)
                {
                    OriUnitDecimal = Convert.ToInt32(item.MF003);
                    OriPriceDecimal = Convert.ToInt32(item.MF004);
                }
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

                #region //本幣小數點取位
                dynamicParameters = new DynamicParameters();
                sql = @"SELECT a.MF003, a.MF004
                        FROM CMSMF a 
                        WHERE a.MF001 = @Currency";
                dynamicParameters.Add("Currency", CurrencyCode);

                var CMSMFResult2 = sqlConnection2.Query(sql, dynamicParameters);

                if (CMSMFResult2.Count() <= 0) throw new SystemException("金額小數點取位資料錯誤!!");

                foreach (var item in CMSMFResult2)
                {
                    UnitDecimal = Convert.ToInt32(item.MF003);
                    PriceDecimal = Convert.ToInt32(item.MF004);
                }
                #endregion

                #region //更新單身資料
                foreach (var item in grDetails)
                {
                    #region //處理【原幣進貨金額】、【原幣未稅金額】、【原幣稅金額】、【本幣未稅金額】、【本幣稅金額】
                    double OrigPreTaxAmt = 0; //原幣未稅金額
                    double OrigTaxAmt = 0; //原幣稅金額
                    double PreTaxAmt = 0; //本幣未稅金額
                    double TaxAmt = 0; //本幣稅金額

                    JObject calculateTaxAmtResult = JObject.Parse(CalculateTaxAmt(item.AvailableQty, item.OrigUnitPrice, item.OrigDiscountAmt, item.OrigAmount, TaxRate, TaxType, Exchange
                                                                                , OriUnitDecimal, OriPriceDecimal, UnitDecimal, PriceDecimal));

                    if (calculateTaxAmtResult["status"].ToString() == "success")
                    {
                        OrigPreTaxAmt = Convert.ToDouble(calculateTaxAmtResult["OrigPreTaxAmt"]);
                        OrigTaxAmt = Convert.ToDouble(calculateTaxAmtResult["OrigTaxAmt"]);
                        PreTaxAmt = Convert.ToDouble(calculateTaxAmtResult["PreTaxAmt"]);
                        TaxAmt = Convert.ToDouble(calculateTaxAmtResult["TaxAmt"]);
                    }
                    else
                    {
                        throw new SystemException(calculateTaxAmtResult["msg"].ToString());
                    }
                    #endregion

                    #region //Update SCM.GrDetail
                    dynamicParameters = new DynamicParameters();
                    sql = @"UPDATE SCM.GrDetail SET
                            OrigPreTaxAmt = @OrigPreTaxAmt,
                            OrigTaxAmt = @OrigTaxAmt,
                            PreTaxAmt = @PreTaxAmt,
                            TaxAmt = @TaxAmt,
                            LastModifiedDate = @LastModifiedDate,
                            LastModifiedBy = @LastModifiedBy
                            WHERE GrDetailId = @GrDetailId";
                    dynamicParameters.AddDynamicParams(
                    new
                    {
                        OrigPreTaxAmt,
                        OrigTaxAmt,
                        PreTaxAmt,
                        TaxAmt,
                        LastModifiedDate,
                        LastModifiedBy,
                        item.GrDetailId
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
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //UpdateGoodsReceipt -- 更新進貨單資料 -- Ann 2024-03-08
        public string UpdateGoodsReceipt(int GrId, string GrErpPrefix, string DocDate, string ReceiptDate, int SupplierId, string Remark
            , string CurrencyCode, double Exchange, string PaymentTerm, int RowCnt
            , string TaxCode, string TaxType, string InvoiceType, double TaxRate, string UiNo, string InvoiceDate
            , string InvoiceNo, string ApplyYYMM, string DeductType, string ContactUser)
        {
            try
            {
                if (GrErpPrefix.Length <= 0) throw new SystemException("【進貨單別】不能為空!");
                if (!DateTime.TryParse(DocDate, out DateTime tempDate)) throw new SystemException("【單據日期】格式錯誤!");
                if (!DateTime.TryParse(ReceiptDate, out DateTime tempDate2)) throw new SystemException("【進貨日期】格式錯誤!");
                if (Exchange < 0) throw new SystemException("【匯率】格式錯誤!");
                if (PaymentTerm.Length <= 0) throw new SystemException("【交易方式】格式錯誤!");
                if (RowCnt < 0) throw new SystemException("【件數】格式錯誤!");
                if (TaxRate < 0) throw new SystemException("【營業稅率】格式錯誤!");
                if (ApplyYYMM.Length <= 0) throw new SystemException("【申報日期】格式錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    string companyNo = "";

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

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //比對ERP關帳日期(庫存關帳)
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MA012)) MA012
                                    FROM CMSMA";
                            var cmsmaResult = sqlConnection2.Query(sql, dynamicParameters);
                            if (cmsmaResult.Count() < 0) throw new SystemException("EPR關帳日期資料錯誤!!");

                            foreach (var item in cmsmaResult)
                            {
                                string eprDate = item.MA012;
                                string erpYear = eprDate.Substring(0, 4);
                                string erpMonth = eprDate.Substring(4, 2);
                                string erpFullDate = erpYear + "-" + erpMonth;
                                DateTime erpDateTime = Convert.ToDateTime(erpFullDate);
                                DateTime DocDateDateTime = Convert.ToDateTime(DocDate.Split('-')[0] + "-" + DocDate.Split('-')[1]);
                                int compare = DocDateDateTime.CompareTo(erpDateTime);
                                if (compare <= 0) throw new SystemException("ERP已關帳(" + erpFullDate + ")，無法開立此日期(" + DocDate + ")之單據!!");
                            }
                            #endregion

                            #region //確認進貨單資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.GrErpPrefix, a.GrErpNo, a.SupplierId
                                    , a.TransferStatus, a.ConfirmStatus
                                    , (
                                        SELECT TOP 1 1
                                        FROM SCM.GrDetail x 
                                        WHERE x.GrId = a.GrId
                                    ) GrDetail
                                    FROM SCM.GoodsReceipt a 
                                    WHERE a.GrId = @GrId";
                            dynamicParameters.Add("GrId", GrId);

                            var GoodsReceiptResult = sqlConnection.Query(sql, dynamicParameters);

                            if (GoodsReceiptResult.Count() <= 0) throw new SystemException("進貨單資料錯誤!!");

                            string GrErpNo = "";
                            foreach (var item in GoodsReceiptResult)
                            {
                                if (item.ConfirmStatus != "N") throw new SystemException("進貨單狀態不可更改!!");
                                if (GrErpPrefix != item.GrErpPrefix && item.TransferStatus == "Y") throw new SystemException("進貨單拋轉後便不可修改單別!!");
                                if (item.GrDetail != null && item.SupplierId != SupplierId) throw new SystemException("存在單身時無法修改供應商!!");
                                GrErpNo = item.GrErpNo;
                            }
                            #endregion

                            #region //確認ERP進貨單資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM PURTG 
                                    WHERE TG001 = @GrErpPrefix
                                    AND TG002 = @GrErpNo
                                    AND TG013 != 'N'";
                            dynamicParameters.Add("GrErpPrefix", GrErpPrefix);
                            dynamicParameters.Add("GrErpNo", GrErpNo);

                            var PURTGResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (PURTGResult.Count() > 0) throw new SystemException("ERP進貨單狀態不可更改!!");
                            #endregion

                            #region //判斷供應商資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 *
                                    FROM SCM.Supplier a 
                                    WHERE a.SupplierId = @SupplierId";
                            dynamicParameters.Add("SupplierId", SupplierId);

                            List<Supplier> resultSuppliers = sqlConnection.Query<Supplier>(sql, dynamicParameters).ToList();

                            if (resultSuppliers.Count() <= 0) throw new SystemException("【供應商】資料錯誤!");
                            #endregion

                            #region //確認幣別資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM CMSMF
                                    WHERE MF001 = @MF001";
                            dynamicParameters.Add("MF001", CurrencyCode);

                            var CMSMFResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (CMSMFResult.Count() <= 0) throw new SystemException("【幣別】資料錯誤!!");
                            #endregion

                            #region //確認付款條件資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(NA002)) PaymentTermNo, LTRIM(RTRIM(NA003)) PaymentTermName 
                                    FROM CMSNA
                                    WHERE NA002 = @NA002";
                            dynamicParameters.Add("NA002", PaymentTerm);

                            var CMSNAResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (CMSNAResult.Count() <= 0) throw new SystemException("【付款條件】資料錯誤!!");
                            #endregion

                            #region //確認稅別碼資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(NN001)) TaxNo, LTRIM(RTRIM(NN002)) TaxName, ROUND(LTRIM(RTRIM(NN004)),2) BusinessTaxRate, LTRIM(RTRIM(NN006)) Taxation
                                    FROM CMSNN 
                                    WHERE NN001 = @NN001";
                            dynamicParameters.Add("NN001", TaxCode);

                            var CMSNNResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (CMSNNResult.Count() <= 0) throw new SystemException("【課稅別】資料錯誤!!");
                            #endregion

                            #region //確認課稅別資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.[Type]
                                    WHERE TypeSchema = 'Customer.Taxation'
                                    AND TypeNo = @TaxType";
                            dynamicParameters.Add("TaxType", TaxType);

                            var TaxationResult = sqlConnection.Query(sql, dynamicParameters);

                            if (TaxationResult.Count() <= 0) throw new SystemException("【課稅別】資料錯誤!!");
                            #endregion

                            #region //確認發票聯數資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(NM001)) InvoiceCountNo, LTRIM(RTRIM(NM002)) InvoiceCountName 
                                    FROM CMSNM a
                                    INNER JOIN CMSNN b ON a.NM001 = b.NN005
                                    WHERE NM001 = @NM001";
                            dynamicParameters.Add("NM001", InvoiceType);

                            var CMSNMResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (CMSNMResult.Count() <= 0) throw new SystemException("【發票聯數】資料錯誤!!");
                            #endregion

                            #region //確認扣抵區分資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.[Type]
                                    WHERE TypeSchema = 'OspReceipt.DeductType'
                                    AND TypeNo = @DeductType";
                            dynamicParameters.Add("DeductType", DeductType);

                            var DeductTypeResult = sqlConnection.Query(sql, dynamicParameters);

                            if (DeductTypeResult.Count() <= 0) throw new SystemException("【扣抵區分】資料錯誤!!");
                            #endregion

                            #region //處理申報日期
                            if (ApplyYYMM.IndexOf("-") != -1)
                            {
                                string[] applyYYMMArray = ApplyYYMM.Split('-');
                                string year = applyYYMMArray[0];
                                string month = applyYYMMArray[1];
                                ApplyYYMM = year + month;
                            }
                            #endregion

                            #region //UPDATE SCM.GoodsReceipt
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.GoodsReceipt SET
                                    GrErpPrefix = @GrErpPrefix,
                                    ReceiptDate = @ReceiptDate,
                                    SupplierId = @SupplierId,
                                    SupplierNo = @SupplierNo,
                                    CurrencyCode = @CurrencyCode,
                                    Exchange = @Exchange,
                                    InvoiceType = @InvoiceType,
                                    TaxType = @TaxType,
                                    InvoiceNo = @InvoiceNo,
                                    DocDate = @DocDate,
                                    Remark = @Remark,
                                    SupplierName = @SupplierName,
                                    UiNo = @UiNo,
                                    DeductType = @DeductType,
                                    InvoiceDate = @InvoiceDate,
                                    ApplyYYMM = @ApplyYYMM,
                                    TaxRate = @TaxRate,
                                    PaymentTerm = @PaymentTerm,
                                    TaxCode = @TaxCode,
                                    ContactUser = @ContactUser,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE GrId = @GrId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    GrErpPrefix,
                                    ReceiptDate,
                                    SupplierId,
                                    resultSuppliers[0].SupplierNo,
                                    CurrencyCode,
                                    Exchange,
                                    InvoiceType,
                                    TaxType,
                                    InvoiceNo,
                                    DocDate,
                                    Remark,
                                    resultSuppliers[0].SupplierName,
                                    UiNo,
                                    DeductType,
                                    InvoiceDate,
                                    ApplyYYMM,
                                    TaxRate,
                                    PaymentTerm,
                                    TaxCode,
                                    ContactUser,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    GrId
                                });

                            int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //根據新匯率重新計算單身資料
                            JObject updateExchangeResult = JObject.Parse(UpdateExchange(GrId, Exchange, sqlConnection, sqlConnection2));
                            if (updateExchangeResult["status"].ToString() != "success")
                            {
                                throw new SystemException(updateExchangeResult["msg"].ToString());
                            }
                            #endregion

                            #region //重新計算進貨單資料
                            JObject updateGrDetailResult = JObject.Parse(UpdateGrTotal(GrId, Exchange, sqlConnection, sqlConnection2));
                            if (updateGrDetailResult["status"].ToString() != "success")
                            {
                                throw new SystemException(updateGrDetailResult["msg"].ToString());
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

        #region //UpdateGrDetail -- 更新進貨單單身 -- Ann 2024-03-08
        public string UpdateGrDetail(int GrDetailId, int GrId, int PoDetailId, int InventoryId, string AcceptanceDate, double ReceiptQty, double ReceiptExpense
            , double AcceptQty, double AvailableQty, double ReturnQty, int UomId, string QcStatus
            , double OrigUnitPrice, double OrigAmount, double OrigDiscountAmt, int? MtlItemId
            , string DiscountDescription, string Remark, string LotNumber)
        {
            try
            {
                string companyNo = "";
                int rowsAffected = 0;
                string MtlItemNo = "", MtlItemName = "", MtlItemSpec = "";
                double? Quantity = 0;
                double? SiQty = 0;
                string PoErpPrefix = "", PoErpNo = "", PoSeq = "";

                #region //基本判斷
                if (ReceiptQty < 0 || ReceiptExpense < 0 || AcceptQty < 0 || AvailableQty < 0 || ReturnQty < 0) throw new SystemException("【進貨數量】【驗收數量】【驗退數量】【計價數量】不可小於0!!");
                if ((AcceptQty + ReturnQty) != ReceiptQty && (AcceptQty != 0 && ReturnQty != 0)) throw new SystemException("【驗收數量】+【驗退數量】需等於【進貨數量】!!");
                if (AvailableQty > ReceiptQty) throw new SystemException("【計價數量】不可大於【進貨數量】!!");
                #endregion

                List<PoDetail> poDetails = new List<PoDetail>();

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
                            companyNo = item.ErpNo;
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //確認進貨單資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.GrId, a.CompanyId, a.GrErpPrefix, a.GrErpNo, GrErpPrefix + '-' + GrErpNo GrErpFullNo, a.ReceiptDate, a.SupplierId, a.SupplierNo
                                    , a.SupplierSo, a.CurrencyCode, a.Exchange, a.InvoiceType, a.TaxType, a.InvoiceNo, a.PrintCnt, a.ConfirmStatus, a.DocDate
                                    , a.RenewFlag, a.Remark, a.TotalAmount, a.DeductAmount, a.OrigTax, a.ReceiptAmount, a.SupplierName, a.UiNo, a.DeductType
                                    , a.ObaccoAndLiquorFlag, a.RowCnt, a.Quantity, FORMAT(a.InvoiceDate, 'yyyy-MM-dd') InvoiceDate, a.OrigPreTaxAmount, a.ApplyYYMM
                                    , a.TaxRate, a.PreTaxAmount, a.TaxAmount, a.PaymentTerm, a.PoId, a.PrepaidErpPrefix, a.PrepaidErpNo, a.OffsetAmount, a.TaxOffset, a.PackageQuantity
                                    , a.PremiumAmount, a.ApproveStatus, a.InvoiceFlag, a.ReserveTaxCode, a.SendCount, a.OrigPremiumAmount, a.EbcErpPreNo, a.EbcEdition, a.EbcFlag
                                    , a.FromDocType, a.NoticeFlag, a.TaxCode, a.TradeTerm, a.DetailMultiTax, a.TaxExchange, a.ContactUser, a.TransferStatus, a.TransferDate
                                    , a.CreateDate, a.LastModifiedDate, a.CreateBy, a.LastModifiedBy
                                    FROM SCM.GoodsReceipt a 
                                    WHERE a.GrId = @GrId";
                            dynamicParameters.Add("GrId", GrId);

                            List<GoodsReceipt> goodsReceipts = sqlConnection.Query<GoodsReceipt>(sql, dynamicParameters).ToList();

                            if (goodsReceipts.Count() <= 0) throw new SystemException("進貨單資料錯誤!!");

                            if (goodsReceipts[0].ConfirmStatus != "N") throw new SystemException("此進貨單已核准，無法更改!!");
                            #endregion

                            #region //確認ERP單據性質資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MQ019)) MQ019 
                                    FROM CMSMQ
                                    WHERE MQ001 = @GrErpPrefix";
                            dynamicParameters.Add("GrErpPrefix", goodsReceipts[0].GrErpPrefix);

                            var CMSMQResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (CMSMQResult.Count() <= 0) throw new SystemException("單據性質資料錯誤!!");

                            string MQ019 = "";
                            foreach (var item in CMSMQResult)
                            {
                                MQ019 = item.MQ019;

                                if (MQ019 == "Y")
                                {
                                    if (PoDetailId <= 0) throw new SystemException("此單據需【核對採購單】，採購單不能為空!!");
                                    //if (MtlItemId > 0) throw new SystemException("此單據需【核對採購單】，無法自行輸入品號!!");
                                }
                            }
                            #endregion

                            #region //確認品號資料是否正確
                            if (MtlItemId > 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.MtlItemNo, a.MtlItemName, a.MtlItemSpec
                                        FROM PDM.MtlItem a 
                                        WHERE a.MtlItemId = @MtlItemId";
                                dynamicParameters.Add("MtlItemId", MtlItemId);

                                var MtlItemResult = sqlConnection.Query(sql, dynamicParameters);

                                if (MtlItemResult.Count() <= 0) throw new SystemException("品號資料錯誤!!");

                                foreach (var item in MtlItemResult)
                                {
                                    MtlItemNo = item.MtlItemNo;
                                    MtlItemName = item.MtlItemName;
                                    MtlItemSpec = item.MtlItemSpec;
                                }
                            }
                            #endregion

                            #region //確認ERP進貨單資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM PURTG 
                                    WHERE TG001 = @GrErpPrefix
                                    AND TG002 = @GrErpNo
                                    AND TG013 != 'N'";
                            dynamicParameters.Add("GrErpPrefix", goodsReceipts[0].GrErpPrefix);
                            dynamicParameters.Add("GrErpNo", goodsReceipts[0].GrErpNo);

                            var PURTGResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (PURTGResult.Count() > 0) throw new SystemException("ERP進貨單狀態不可更改!!");
                            #endregion

                            #region //確認進貨單身資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM SCM.GrDetail 
                                    WHERE GrDetailId = @GrDetailId";
                            dynamicParameters.Add("GrDetailId", GrDetailId);

                            var GrDetailResult = sqlConnection.Query(sql, dynamicParameters);

                            if (GrDetailResult.Count() <= 0) throw new SystemException("進貨單身資料錯誤!!");
                            #endregion

                            #region //確認採購單身資料是否正確
                            if (PoDetailId > 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.PoDetailId, a.PoId, a.PoErpPrefix, a.PoErpNo, a.PoSeq, a.MtlItemId, a.PoMtlItemName, a.PoMtlItemSpec, a.InventoryId, a.Quantity
                                        , a.UomId, a.PoUnitPrice, a.PoPrice, a.PromiseDate, a.ReferencePrefix, a.Remark, a.SiQty, a.ClosureStatus, a.ConfirmStatus 
                                        , a.InventoryQty, a.SmallUnit, a.ReferenceNo, a.Project, a.ReferenceSeq, a.UrgentMtl, a.PrErpPrefix, a.PrErpNo, a.PrSequence 
                                        , a.FromDocType, a.MtlItemType, a.PartialSeq, a.OriPromiseDate, a.DeliveryDate, a.PoPriceQty, a.PoPriceUomId, a.SiPriceQty 
                                        , a.DiscountRate, a.DiscountAmount, a.CreateDate, a.LastModifiedDate, a.CreateBy, a.LastModifiedBy
                                        , b.UomNo
                                        , c.InventoryNo
                                        , d.MtlItemNo
                                        , e.TaxNo
                                        FROM SCM.PoDetail a
                                        INNER JOIN PDM.UnitOfMeasure b ON a.UomId = b.UomId
                                        INNER JOIN SCM.Inventory c ON a.InventoryId = c.InventoryId
                                        INNER JOIN PDM.MtlItem d ON a.MtlItemId = d.MtlItemId
                                        INNER JOIN SCM.PurchaseOrder e ON a.PoId = e.PoId
                                        WHERE a.PoDetailId = @PoDetailId";
                                dynamicParameters.Add("PoDetailId", PoDetailId);

                                poDetails = sqlConnection.Query<PoDetail>(sql, dynamicParameters).ToList();

                                #region //確認ERP採購單狀態是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT LTRIM(RTRIM(TC014)) TC014
                                        FROM PURTC 
                                        WHERE TC001 = @PoErpPrefix
                                        AND TC002 = @PoErpNo";
                                dynamicParameters.Add("PoErpPrefix", poDetails[0].PoErpPrefix);
                                dynamicParameters.Add("PoErpNo", poDetails[0].PoErpNo);

                                var PURTCResult = sqlConnection2.Query(sql, dynamicParameters);

                                if (PURTCResult.Count() <= 0) throw new SystemException("ERP採購單資料錯誤!!");

                                foreach (var item in PURTCResult)
                                {
                                    if (item.TC014 != "Y") throw new SystemException("ERP採購單狀態錯誤!!");
                                }
                                #endregion
                            }
                            #endregion

                            #region //原幣小數點取位
                            int OriUnitDecimal = 0;
                            int OriPriceDecimal = 0;
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.MF003, a.MF004
                                    FROM CMSMF a 
                                    WHERE a.MF001 = @Currency";
                            dynamicParameters.Add("Currency", goodsReceipts[0].CurrencyCode);

                            var CMSMFResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (CMSMFResult.Count() <= 0) throw new SystemException("金額小數點取位資料錯誤!!");

                            foreach (var item in CMSMFResult)
                            {
                                OriUnitDecimal = Convert.ToInt32(item.MF003);
                                OriPriceDecimal = Convert.ToInt32(item.MF004);
                            }
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

                            #region //本幣小數點取位
                            int UnitDecimal = 0;
                            int PriceDecimal = 0;
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.MF003, a.MF004
                                    FROM CMSMF a 
                                    WHERE a.MF001 = @Currency";
                            dynamicParameters.Add("Currency", CurrencyCode);

                            var CMSMFResult2 = sqlConnection2.Query(sql, dynamicParameters);

                            if (CMSMFResult2.Count() <= 0) throw new SystemException("金額小數點取位資料錯誤!!");

                            foreach (var item in CMSMFResult2)
                            {
                                UnitDecimal = Convert.ToInt32(item.MF003);
                                PriceDecimal = Convert.ToInt32(item.MF004);
                            }
                            #endregion

                            #region //確認庫別資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT InventoryNo
                                    FROM SCM.Inventory
                                    WHERE InventoryId = @InventoryId";
                            dynamicParameters.Add("InventoryId", InventoryId);

                            var InventoryResult = sqlConnection.Query(sql, dynamicParameters);

                            if (InventoryResult.Count() <= 0) throw new SystemException("庫別資料錯誤!!");

                            string InventoryNo = "";
                            foreach (var item in InventoryResult)
                            {
                                InventoryNo = item.InventoryNo;
                            }
                            #endregion

                            #region //確認進貨單位資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT UomNo
                                    FROM PDM.UnitOfMeasure
                                    WHERE UomId = @UomId";
                            dynamicParameters.Add("UomId", UomId);

                            var UnitOfMeasureResult = sqlConnection.Query(sql, dynamicParameters);

                            if (UnitOfMeasureResult.Count() <= 0) throw new SystemException("進貨單位資料錯誤!!");

                            string UomNo = "";
                            foreach (var item in UnitOfMeasureResult)
                            {
                                UomNo = item.UomNo;
                            }
                            #endregion

                            #region //採購單帶入資料模式
                            foreach (var item in poDetails)
                            {
                                if (MtlItemId <= 0)
                                {
                                    MtlItemId = item.MtlItemId;
                                    MtlItemNo = item.MtlItemNo;
                                    MtlItemName = item.PoMtlItemName;
                                    MtlItemSpec = item.PoMtlItemSpec;
                                }
                                Quantity = item.Quantity;
                                SiQty = item.SiQty;
                                PoErpPrefix = item.PoErpPrefix;
                                PoErpNo = item.PoErpNo;
                                PoSeq = item.PoSeq;

                                #region //確認ERP採購單身結案碼
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT LTRIM(RTRIM(TD016)) TD016
                                        , LTRIM(RTRIM(TD018)) TD018
                                        FROM PURTD
                                        WHERE TD001 = @PoErpPrefix
                                        AND TD002 = @PoErpNo
                                        AND TD003 = @PoSeq";
                                dynamicParameters.Add("PoErpPrefix", item.PoErpPrefix);
                                dynamicParameters.Add("PoErpNo", item.PoErpNo);
                                dynamicParameters.Add("PoSeq", item.PoSeq);

                                var PURTDResult = sqlConnection2.Query(sql, dynamicParameters);

                                if (PURTDResult.Count() <= 0) throw new SystemException("ERP採購單單身資料錯誤!!");

                                foreach (var item2 in PURTDResult)
                                {
                                    if (item2.TD016 != "N") throw new SystemException("ERP採購單單身結案碼狀態錯誤!!");
                                    if (item2.TD018 != "Y") throw new SystemException("ERP採購單單身狀態錯誤!!");
                                }
                                #endregion

                                #region //確認此進貨單身無相同採購單身
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT ISNULL(SUM(a.ReceiptQty), 0) TotalReceiptQty
                                        FROM SCM.GrDetail a 
                                        WHERE a.PoErpPrefix = @PoErpPrefix
                                        AND a.PoErpNo = @PoErpNo
                                        AND a.PoSeq = @PoSeq
                                        AND a.GrId = @GrId
                                        AND a.GrDetailId != @GrDetailId";
                                dynamicParameters.Add("PoErpPrefix", item.PoErpPrefix);
                                dynamicParameters.Add("PoErpNo", item.PoErpNo);
                                dynamicParameters.Add("PoSeq", item.PoSeq);
                                dynamicParameters.Add("GrId", GrId);
                                dynamicParameters.Add("GrDetailId", GrDetailId);

                                var PoDetailResult = sqlConnection.Query(sql, dynamicParameters);

                                foreach (var item2 in PoDetailResult)
                                {
                                    #region //確認數量沒有超過剩餘可進貨數量
                                    double poQty = Convert.ToDouble(item.Quantity);
                                    double siQty = Convert.ToDouble(item.SiQty);
                                    double TotalReceiptQty = item2.TotalReceiptQty;
                                    double allowQty = poQty - siQty - TotalReceiptQty > 0 ? poQty - siQty - TotalReceiptQty : 0;

                                    if (ReceiptQty > allowQty)
                                    {
                                        throw new SystemException("此採購單已超過剩餘可進貨數量: " + allowQty + "!!");
                                    }
                                    #endregion
                                }
                                #endregion
                            }
                            #endregion

                            #region //確認ERP品號相關資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MB004)) MB004
                                    , LTRIM(RTRIM(MB030)) MB030
                                    , LTRIM(RTRIM(MB031)) MB031
                                    , LTRIM(RTRIM(MB022)) MB022
                                    , LTRIM(RTRIM(MB043)) MB043
                                    , LTRIM(RTRIM(MB044)) MB044
                                    FROM INVMB
                                    WHERE MB001 = @MB001";
                            dynamicParameters.Add("MB001", MtlItemNo);

                            var INVMBResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (INVMBResult.Count() <= 0) throw new SystemException("ERP品號基本資料錯誤!!");

                            foreach (var item2 in INVMBResult)
                            {
                                #region //判斷ERP品號生效日與失效日
                                if (item2.MB030 != "" && item2.MB030 != null)
                                {
                                    #region //判斷單據日期需大於或等於生效日
                                    string EffectiveDate = item2.MB030;
                                    string effYear = EffectiveDate.Substring(0, 4);
                                    string effMonth = EffectiveDate.Substring(4, 2);
                                    string effDay = EffectiveDate.Substring(6, 2);
                                    DateTime effFullDate = Convert.ToDateTime(effYear + "-" + effMonth + "-" + effDay);
                                    int effresult = DateTime.Compare((DateTime)goodsReceipts[0].DocDate, effFullDate);
                                    if (effresult < 0) throw new SystemException("此品號尚未生效，無法使用!!");
                                    #endregion
                                }

                                if (item2.MB031 != "" && item2.MB031 != null)
                                {
                                    #region //判斷日期需小於或等於失效日
                                    string ExpirationDate = item2.MB031;
                                    string effYear = ExpirationDate.Substring(0, 4);
                                    string effMonth = ExpirationDate.Substring(4, 2);
                                    string effDay = ExpirationDate.Substring(6, 2);
                                    DateTime effFullDate = Convert.ToDateTime(effYear + "-" + effMonth + "-" + effDay);
                                    int effresult = DateTime.Compare((DateTime)goodsReceipts[0].DocDate, effFullDate);
                                    if (effresult > 0) throw new SystemException("此品號已失效，無法使用!!");
                                    #endregion
                                }
                                #endregion

                                #region //確認此品號是否需要批號管理
                                if (item2.MB022 == "Y" || item2.MB022 == "T")
                                {
                                    //if (LotNumber.Length <= 0) throw new SystemException("品號【" + MtlItemNo + "】需設定批號!!");
                                }
                                else
                                {
                                    if (LotNumber.Length > 0) throw new SystemException("品號【" + MtlItemNo + "】不需設定批號!!");
                                }
                                #endregion

                                #region //確認檢驗狀態
                                if (item2.MB043 == "0")
                                {
                                    if (QcStatus != "0") throw new SystemException("品號【" + MtlItemNo + "】設定為免檢，無法修改!!");
                                }
                                else
                                {
                                    if (QcStatus == "0") throw new SystemException("品號【" + MtlItemNo + "】非免檢類型，無法修改!!");
                                }
                                #endregion

                                #region //檢核數量是否超收
                                if (item2.MB044 == "Y" && poDetails.Count() > 0) //卡控超收
                                {
                                    if (ReceiptQty > (Quantity - SiQty)) throw new SystemException("此次數量已超過剩餘可進貨數量(" + (Quantity - SiQty).ToString() + ")!!");
                                }
                                #endregion
                            }
                            #endregion

                            #region //確認此進貨單是否有進貨檢驗量測單據
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.QcRecordId, a.QcStatus
                                    FROM MES.QcGoodsReceipt a 
                                    WHERE a.GrDetailId = @GrDetailId
                                    ORDER BY a.CreateDate DESC";
                            dynamicParameters.Add("GrDetailId", GrDetailId);

                            var QcGoodsReceiptResult = sqlConnection.Query(sql, dynamicParameters);

                            if (QcGoodsReceiptResult.Count() > 0)
                            {
                                bool qcFlag = false;
                                foreach (var item2 in QcGoodsReceiptResult)
                                {
                                    if (item2.QcStatus == "A" && (AcceptQty != 0 || ReturnQty != 0 || AvailableQty != 0))
                                    {
                                        throw new SystemException("此進貨單身已開立進貨檢驗單據【" + item2.QcRecordId + "】，但尚未量測完成!!");
                                    }

                                    qcFlag = true;

                                    #region //確認此量測單據是否有品異單
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 a.JudgeStatus, a.QcRecordId, a.ReleaseQty
                                            , b.AbnormalqualityNo
                                            FROM QMS.AqBarcode a 
                                            INNER JOIN QMS.Abnormalquality b ON a.AbnormalqualityId = b.AbnormalqualityId
                                            WHERE a.QcRecordId = @QcRecordId";
                                    dynamicParameters.Add("QcRecordId", item2.QcRecordId);

                                    var AqBarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                                    if (AqBarcodeResult.Count() > 0)
                                    {
                                        foreach (var item3 in AqBarcodeResult)
                                        {
                                            if (item3.JudgeStatus == null)
                                            {
                                                throw new SystemException("品異單據【" + item3.AbnormalqualityNo + "】尚未進行判定!!");
                                            }
                                            else if (item3.JudgeStatus == "AM")
                                            {
                                                qcFlag = false;
                                            }
                                            else if (item3.JudgeStatus == "R")
                                            {
                                                if (item3.ReleaseQty == null) throw new SystemException("品異單據【" + item3.AbnormalqualityNo + "】尚未維護特採數量!!");

                                                #region //確認驗收數量是否正確
                                                if (AcceptQty > item3.ReleaseQty) throw new SystemException("驗收數量大於特採數量!!");
                                                #endregion
                                            }
                                            else if (item3.JudgeStatus == "S")
                                            {
                                                #region //若判定為報廢，驗收數量需為0
                                                if (AcceptQty != 0) throw new SystemException("此進貨單身判定為報廢，驗收數量需為0!!");
                                                #endregion
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (item2.QcStatus != "Y" && (AcceptQty != 0 || ReturnQty != 0 || AvailableQty != 0))
                                        {
                                            throw new SystemException("進貨檢驗單號【" + item2.QcRecordId + "】狀態非合格，但查無品異單據!!");
                                        }
                                    }
                                    #endregion
                                }

                                if (qcFlag == false)
                                {
                                    throw new SystemException("此進貨單身已有開立品異單判定為【重新量測】，但尚未有新量測單據!!");
                                }
                            }
                            else
                            {
                                if (QcStatus != "0" && QcStatus != "1") throw new SystemException("品號【" + MtlItemNo + "】非免檢類型，需要進行進貨檢驗!!");
                            }
                            #endregion

                            #region //查詢單身序號
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT ISNULL(MAX(a.GrSeq), 0) MaxGrSeq
                                        FROM SCM.GrDetail a 
                                        WHERE a.GrId = @GrId";
                            dynamicParameters.Add("GrId", GrId);

                            var GrSeqResult = sqlConnection.Query(sql, dynamicParameters);

                            int MaxGrSeq = 0;
                            foreach (var item2 in GrSeqResult)
                            {
                                MaxGrSeq = Convert.ToInt32(item2.MaxGrSeq);
                            }

                            string GrSeq = (MaxGrSeq + 1).ToString("D4");
                            #endregion

                            #region //查詢目前庫存數
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT ISNULL(MC007, 0) MC007
                                    FROM INVMC
                                    WHERE MC001 = @MC001
                                    AND MC002 = @MC002";
                            dynamicParameters.Add("MC001", MtlItemNo);
                            dynamicParameters.Add("MC002", InventoryNo);

                            var INVMCResult = sqlConnection2.Query(sql, dynamicParameters);

                            double InventoryQty = 0;
                            foreach (var item2 in INVMCResult)
                            {
                                InventoryQty = Convert.ToDouble(item2.MC007);
                            }
                            #endregion

                            #region //處理【原幣進貨金額】、【原幣未稅金額】、【原幣稅金額】、【本幣未稅金額】、【本幣稅金額】
                            double OrigPreTaxAmt = 0; //原幣未稅金額
                            double OrigTaxAmt = 0; //原幣稅金額
                            double PreTaxAmt = 0; //本幣未稅金額
                            double TaxAmt = 0; //本幣稅金額

                            JObject calculateTaxAmtResult = JObject.Parse(CalculateTaxAmt(AvailableQty, OrigUnitPrice, OrigDiscountAmt, OrigAmount, goodsReceipts[0].TaxRate, goodsReceipts[0].TaxType, goodsReceipts[0].Exchange
                                                                            , OriUnitDecimal, OriPriceDecimal, UnitDecimal, PriceDecimal));

                            if (calculateTaxAmtResult["status"].ToString() == "success")
                            {
                                OrigPreTaxAmt = Convert.ToDouble(calculateTaxAmtResult["OrigPreTaxAmt"]);
                                OrigTaxAmt = Convert.ToDouble(calculateTaxAmtResult["OrigTaxAmt"]);
                                PreTaxAmt = Convert.ToDouble(calculateTaxAmtResult["PreTaxAmt"]);
                                TaxAmt = Convert.ToDouble(calculateTaxAmtResult["TaxAmt"]);
                            }
                            else
                            {
                                throw new SystemException(calculateTaxAmtResult["msg"].ToString());
                            }
                            #endregion

                            #region //UPDATE SCM.GrDetail
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.GrDetail SET
                                    MtlItemId = @MtlItemId,
                                    MtlItemNo = @MtlItemNo,
                                    GrMtlItemName = @GrMtlItemName,
                                    GrMtlItemSpec = @GrMtlItemSpec,
                                    ReceiptQty = @ReceiptQty,
                                    LotNumber = @LotNumber,
                                    UomId = @UomId,
                                    UomNo = @UomNo,
                                    InventoryId = @InventoryId,
                                    InventoryNo = @InventoryNo,
                                    PoErpPrefix = @PoErpPrefix,
                                    PoErpNo = @PoErpNo,
                                    PoSeq = @PoSeq,
                                    AcceptanceDate = @AcceptanceDate,
                                    AcceptQty = @AcceptQty,
                                    AvailableQty = @AvailableQty,
                                    ReturnQty = @ReturnQty,
                                    OrigUnitPrice = @OrigUnitPrice,
                                    OrigAmount = @OrigAmount,
                                    OrigDiscountAmt = @OrigDiscountAmt,
                                    ReceiptExpense = @ReceiptExpense,
                                    DiscountDescription = @DiscountDescription,
                                    QcStatus = @QcStatus,
                                    Remark = @Remark,
                                    InventoryQty = @InventoryQty,
                                    OrigPreTaxAmt = @OrigPreTaxAmt,
                                    OrigTaxAmt = @OrigTaxAmt,
                                    PreTaxAmt = @PreTaxAmt,
                                    TaxAmt = @TaxAmt,
                                    AvailableUomId = @AvailableUomId,
                                    AvailableUomNo = @AvailableUomNo,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE GrDetailId = @GrDetailId";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                MtlItemId,
                                MtlItemNo,
                                GrMtlItemName = MtlItemName,
                                GrMtlItemSpec = MtlItemSpec,
                                ReceiptQty,
                                LotNumber,
                                UomId,
                                UomNo,
                                InventoryId,
                                InventoryNo,
                                PoErpPrefix,
                                PoErpNo,
                                PoSeq,
                                AcceptanceDate,
                                AcceptQty,
                                AvailableQty,
                                ReturnQty,
                                OrigUnitPrice,
                                OrigAmount,
                                OrigDiscountAmt,
                                ReceiptExpense,
                                DiscountDescription,
                                QcStatus,
                                Remark,
                                InventoryQty,
                                OrigPreTaxAmt,
                                OrigTaxAmt,
                                PreTaxAmt,
                                TaxAmt,
                                AvailableUomId = UomId,
                                AvailableUomNo = UomNo,
                                LastModifiedDate,
                                LastModifiedBy,
                                GrDetailId
                            });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //重新計算進貨單資料
                            JObject updateGrDetailResult = JObject.Parse(UpdateGrTotal(GrId, goodsReceipts[0].Exchange, sqlConnection, sqlConnection2));
                            if (updateGrDetailResult["status"].ToString() != "success")
                            {
                                throw new SystemException(updateGrDetailResult["msg"].ToString());
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

        #region //UpdateTransferGoodsReceipt -- 拋轉進貨單據 -- Ann 2024-03-11
        public string UpdateTransferGoodsReceipt(int GrId)
        {
            try
            {
                string companyNo = "";
                int rowsAffected = 0;
                string supplierNo = "";

                List<GoodsReceipt> goodsReceipts = new List<GoodsReceipt>();
                List<GrDetail> grDetails = new List<GrDetail>();

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
                            companyNo = item.ErpNo;
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //更新BM進貨單資料
                            JObject updateGrDetailResult = JObject.Parse(UpdateGrTotal(GrId, -1, sqlConnection, sqlConnection2));
                            if (updateGrDetailResult["status"].ToString() != "success")
                            {
                                throw new SystemException(updateGrDetailResult["msg"].ToString());
                            }
                            #endregion

                            #region //取得BM進貨單頭資料
                            sql = @"SELECT a.GrId, a.CompanyId, a.GrErpPrefix, a.GrErpNo, GrErpPrefix + '-' + GrErpNo GrErpFullNo
                                    , CASE WHEN LEN(a.ReceiptDate) > 0 THEN FORMAT(CAST(a.ReceiptDate as date), 'yyyy-MM-dd') ELSE null END AS ReceiptDate
                                    , CASE WHEN LEN(a.ReceiptDate) > 0 THEN FORMAT(CAST(a.ReceiptDate as date), 'yyyyMMdd') ELSE '' END AS ErpReceiptDate
                                    , a.SupplierId, a.SupplierNo, a.SupplierSo, a.CurrencyCode, a.Exchange, a.InvoiceType, a.TaxType, a.InvoiceNo, a.PrintCnt, a.ConfirmStatus
                                    , a.DocDate, CASE WHEN LEN(a.DocDate) > 0 THEN FORMAT(CAST(a.DocDate as date), 'yyyyMMdd') ELSE '' END AS ErpDocDate
                                    , a.RenewFlag, a.Remark, a.TotalAmount, a.DeductAmount, a.OrigTax, a.ReceiptAmount, a.SupplierName, a.UiNo, a.DeductType
                                    , a.ObaccoAndLiquorFlag, a.RowCnt, a.Quantity
                                    , CASE WHEN FORMAT(a.InvoiceDate, 'yyyy-MM-dd') = '1900-01-01' THEN null ELSE FORMAT(CAST(a.InvoiceDate as date), 'yyyy-MM-dd') END AS InvoiceDate
                                    , CASE WHEN LEN(a.InvoiceDate) > 0 THEN FORMAT(CAST(a.InvoiceDate as date), 'yyyyMMdd') ELSE '' END AS ErpInvoiceDate
                                    , a.OrigPreTaxAmount, a.ApplyYYMM, a.TaxRate, a.PreTaxAmount, a.TaxAmount, a.PaymentTerm, a.PoId, a.PrepaidErpPrefix, a.PrepaidErpNo, a.OffsetAmount
                                    , a.TaxOffset, a.PackageQuantity, a.PremiumAmount, a.ApproveStatus, a.InvoiceFlag, a.ReserveTaxCode, a.SendCount, a.OrigPremiumAmount, a.EbcErpPreNo
                                    , a.EbcEdition, a.EbcFlag, a.FromDocType, a.NoticeFlag, a.TaxCode, a.TradeTerm, a.DetailMultiTax, a.TaxExchange, a.ContactUser, a.TransferStatus, a.TransferDate
                                    , a.CreateDate, a.LastModifiedDate, a.CreateBy, a.LastModifiedBy
                                    FROM SCM.GoodsReceipt a 
                                    WHERE a.CompanyId = @CompanyId
                                    AND a.GrId = @GrId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("GrId", GrId);

                            goodsReceipts = sqlConnection.Query<GoodsReceipt>(sql, dynamicParameters).ToList();

                            if (goodsReceipts[0].ConfirmStatus != "N") throw new SystemException("進貨單目前狀態不可拋轉!!");
                            supplierNo = goodsReceipts[0].SupplierNo;

                            #region //整理單頭日期格式
                            string DocDate = ((DateTime)goodsReceipts.FirstOrDefault().DocDate).ToString("yyyy-MM-dd");
                            string ReceiptDate = ((DateTime)goodsReceipts.FirstOrDefault().ReceiptDate).ToString("yyyy-MM-dd");
                            string InvoiceDate = goodsReceipts.FirstOrDefault().InvoiceDate == null ? "" : ((DateTime)goodsReceipts.FirstOrDefault().InvoiceDate).ToString("yyyyMMdd");
                            #endregion
                            #endregion

                            #region //取得BM進貨單身資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.GrDetailId, a.GrId, a.GrErpPrefix, a.GrErpNo, a.GrSeq, a.MtlItemId, a.MtlItemNo, a.GrMtlItemName, a.GrMtlItemSpec
                                    , a.ReceiptQty, a.UomId, a.UomNo, a.InventoryId, a.InventoryNo, a.LotNumber, a.PoErpPrefix, a.PoErpNo, a.PoSeq
                                    , a.AcceptanceDate, CASE WHEN LEN(a.AcceptanceDate) > 0 THEN FORMAT(CAST(a.AcceptanceDate as date), 'yyyyMMdd') ELSE '' END AS ErpAcceptanceDate
                                    , a.AcceptQty, a.AvailableQty, a.ReturnQty, a.OrigUnitPrice, a.OrigAmount, a.OrigDiscountAmt, a.TrErpPrefix, a.TrErpNo, a.TrSeq
                                    , a.ReceiptExpense, a.DiscountDescription, a.PaymentHold, a.Overdue, a.QcStatus, a.ReturnCode, a.ConfirmStatus, a.CloseStatus
                                    , a.ReNewStatus, a.Remark, a.InventoryQty, a.SmallUnit
                                    , CASE WHEN FORMAT(a.AvailableDate, 'yyyy-MM-dd') = '1900-01-01' THEN null ELSE FORMAT(CAST(a.AvailableDate as date), 'yyyy-MM-dd') END AS AvailableDate
                                    , CASE WHEN LEN(a.AvailableDate) > 0 THEN FORMAT(CAST(a.AvailableDate as date), 'yyyyMMdd') ELSE '' END AS ErpAvailableDate
                                    , CASE WHEN FORMAT(a.ReCheckDate, 'yyyy-MM-dd') = '1900-01-01' THEN null ELSE FORMAT(CAST(a.ReCheckDate as date), 'yyyy-MM-dd') END AS ReCheckDate
                                    , CASE WHEN LEN(a.ReCheckDate) > 0 THEN FORMAT(CAST(a.ReCheckDate as date), 'yyyyMMdd') ELSE '' END AS ErpReCheckDate
                                    , a.ConfirmUser, a.ApvErpPrefix, a.ApvErpNo
                                    , a.ApvSeq, a.ProjectCode, a.ExpenseEntry, a.PremiumAmountFlag, a.OrigPreTaxAmt, a.OrigTaxAmt, a.PreTaxAmt, a.TaxAmt, a.ReceiptPackageQty
                                    , a.AcceptancePackageQty, a.ReturnPackageQty, a.PremiumAmount, a.PackageUnit, a.ReserveTaxCode, a.OrigPremiumAmount, a.AvailableUomId
                                    , a.AvailableUomNo, a.OrigCustomer, a.ApproveStatus, a.EbcErpPreNo, a.EbcEdition, a.ProductSeqAmount, a.MtlItemType, a.Loaction
                                    , a.TaxRate, a.MultipleFlag, a.GrErpStatus, a.TaxCode, a.DiscountRate, a.DiscountPrice, a.QcType, a.TransferStatus, a.TransferDate
                                    , a.CreateDate, a.LastModifiedDate, a.CreateBy, a.LastModifiedBy
                                    FROM SCM.GrDetail a
                                    WHERE a.GrId = @GrId";
                            dynamicParameters.Add("GrId", GrId);

                            grDetails = sqlConnection.Query<GrDetail>(sql, dynamicParameters).ToList();
                            #endregion

                            #region //確認ERP進貨單單頭資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(TG013)) TG013
                                    FROM PURTG
                                    WHERE TG001 = @GrErpPrefix
                                    AND TG002 = @GrErpNo";
                            dynamicParameters.Add("GrErpPrefix", goodsReceipts[0].GrErpPrefix);
                            dynamicParameters.Add("GrErpNo", goodsReceipts[0].GrErpNo);

                            var PURTGResult = sqlConnection2.Query(sql, dynamicParameters);

                            foreach (var item in PURTGResult)
                            {
                                if (item.TG013 != "N") throw new SystemException("進貨單目前狀態不可拋轉!!");
                            }
                            #endregion

                            #region //比對ERP關帳日期(庫存關帳)
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MA012)) MA012
                                    FROM CMSMA";
                            var cmsmaResult = sqlConnection2.Query(sql, dynamicParameters);
                            if (cmsmaResult.Count() < 0) throw new SystemException("EPR關帳日期資料錯誤!!");

                            foreach (var item in cmsmaResult)
                            {
                                string eprDate = item.MA012;
                                string erpYear = eprDate.Substring(0, 4);
                                string erpMonth = eprDate.Substring(4, 2);
                                string erpFullDate = erpYear + "-" + erpMonth;
                                DateTime erpDateTime = Convert.ToDateTime(erpFullDate);
                                DateTime DocDateDateTime = Convert.ToDateTime(DocDate.ToString().Split('-')[0] + "-" + DocDate.ToString().Split('-')[1]);
                                int compare = DocDateDateTime.CompareTo(erpDateTime);
                                if (compare <= 0) throw new SystemException("ERP已關帳(" + erpFullDate + ")，無法開立此日期(" + goodsReceipts[0].DocDate.ToString() + ")之單據!!");
                            }
                            #endregion

                            #region //取得目前使用者資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT UserNo
                                    FROM BAS.[User]
                                    WHERE UserId = @UserId";
                            dynamicParameters.Add("UserId", CurrentUser);

                            var UserResult = sqlConnection.Query(sql, dynamicParameters);

                            string UserNo = "";
                            foreach (var item in UserResult)
                            {
                                UserNo = item.UserNo;
                            }
                            #endregion

                            #region //審核ERP權限
                            string USR_GROUP = BaseHelper.CheckErpAuthority(UserNo, sqlConnection2, "", "");
                            #endregion

                            #region //查詢廠別
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MB001)) MB001 FROM CMSMB";
                            var CMSMBResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (CMSMBResult.Count() <= 0) throw new SystemException("找不到廠別!!");
                            if (CMSMBResult.Count() > 1) throw new SystemException("廠別數有多個，請與資訊人員確認!!");

                            string factoryNo = "";
                            foreach (var item in CMSMBResult)
                            {
                                factoryNo = item.MB001;
                            }
                            #endregion

                            #region //新增/修改單頭PURTG資料
                            if (PURTGResult.Count() > 0)
                            {
                                #region //UPDATE PURTG 進貨單單頭
                                foreach (var item in goodsReceipts)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE PURTG SET
                                            MODIFIER = @MODIFIER,
                                            MODI_DATE = @MODI_DATE,
                                            FLAG += 1,
                                            MODI_TIME = @MODI_TIME,
                                            MODI_AP = @MODI_AP,
                                            MODI_PRID = @MODI_PRID,
                                            TG003 = @ReceiptDate,
                                            TG005 = @SupplierNo,
                                            TG007 = @CurrencyCode,
                                            TG008 = @Exchange,
                                            TG009 = @InvoiceType,
                                            TG010 = @TaxType,
                                            TG011 = @InvoiceNo,
                                            TG014 = @DocDate,
                                            TG016 = @Remark,
                                            TG017 = @TotalAmount,
                                            TG018 = @DeductAmount,
                                            TG019 = @OrigTax,
                                            TG020 = @ReceiptAmount,
                                            TG021 = @SupplierName,
                                            TG023 = @DeductType,
                                            TG026 = @Quantity,
                                            TG027 = @InvoiceDate,
                                            TG028 = @OrigPreTaxAmount,
                                            TG029 = @ApplyYYMM,
                                            TG030 = @TaxRate,
                                            TG031 = @PreTaxAmount,
                                            TG032 = @TaxAmount,
                                            TG033 = @PaymentTerm,
                                            TG034 = @PoErpPrefix,
                                            TG035 = @PoErpNo,
                                            TG041 = @PremiumAmount,
                                            TG058 = @TaxCode,
                                            TG059 = @TradeTerm,
                                            TG067 = @ContactUser
                                            WHERE TG001 = @GrErpPrefix
                                            AND TG002 = @GrErpNo";
                                    dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        MODIFIER = UserNo,
                                        MODI_DATE = DateTime.Now.ToString("yyyyMMdd"),
                                        MODI_TIME = DateTime.Now.ToString("HH:mm:ss"),
                                        MODI_AP = UserNo + "PC",
                                        MODI_PRID = "BM",
                                        ReceiptDate = item.ErpReceiptDate,
                                        item.SupplierNo,
                                        item.CurrencyCode,
                                        item.Exchange,
                                        item.InvoiceType,
                                        item.TaxType,
                                        item.InvoiceNo,
                                        DocDate = item.ErpDocDate,
                                        item.Remark,
                                        item.TotalAmount,
                                        item.DeductAmount,
                                        item.OrigTax,
                                        item.ReceiptAmount,
                                        item.SupplierName,
                                        item.DeductType,
                                        item.Quantity,
                                        InvoiceDate = item.ErpInvoiceDate,
                                        item.OrigPreTaxAmount,
                                        item.ApplyYYMM,
                                        item.TaxRate,
                                        item.PreTaxAmount,
                                        item.TaxAmount,
                                        item.PaymentTerm,
                                        item.PoErpPrefix,
                                        item.PoErpNo,
                                        item.PremiumAmount,
                                        item.TaxCode,
                                        item.TradeTerm,
                                        item.ContactUser,
                                        item.GrErpPrefix,
                                        item.GrErpNo,
                                    });

                                    rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);
                                }
                                #endregion

                                #region //刪除原本單身
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE PURTH
                                        WHERE TH001 = @GrErpPrefix
                                        AND TH002 = @GrErpNo";
                                dynamicParameters.Add("GrErpPrefix", goodsReceipts[0].GrErpPrefix);
                                dynamicParameters.Add("GrErpNo", goodsReceipts[0].GrErpNo);

                                rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);
                                #endregion
                            }
                            else
                            {
                                #region //取號流程
                                string GrErpNo = "";
                                DateTime referenceTime = DateTime.ParseExact(goodsReceipts[0].ErpDocDate, "yyyyMMdd", CultureInfo.InvariantCulture);

                                #region //單據設定
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.MQ004, a.MQ005, a.MQ006, a.MQ044
                                        FROM CMSMQ a
                                        WHERE a.COMPANY = @CompanyNo
                                        AND a.MQ001 = @ErpPrefix";
                                dynamicParameters.Add("CompanyNo", companyNo);
                                dynamicParameters.Add("ErpPrefix", goodsReceipts[0].GrErpPrefix);

                                var resultDocSetting = sqlConnection2.Query(sql, dynamicParameters);
                                if (resultDocSetting.Count() <= 0) throw new SystemException("ERP單據設定資料不存在!");

                                string encode = "";
                                int yearLength = -1;
                                int lineLength = -1;
                                foreach (var item in resultDocSetting)
                                {
                                    encode = item.MQ004; //編碼方式
                                    yearLength = Convert.ToInt32(item.MQ005); //年碼數
                                    lineLength = Convert.ToInt32(item.MQ006); //流水號碼數
                                }
                                #endregion

                                #region //單號取號
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT CONVERT(int, RIGHT(ISNULL(MAX(RTRIM(LTRIM(TG002))), '" + new string('0', lineLength) + @"'), " + lineLength + @")) + 1 CurrentNum
                                        FROM PURTG
                                        WHERE TG001 = @ErpPrefix";
                                dynamicParameters.Add("ErpPrefix", goodsReceipts[0].GrErpPrefix);

                                #region //編碼方式
                                string dateFormat = "";
                                switch (encode)
                                {
                                    case "1": //日編
                                        dateFormat = new string('y', yearLength) + "MMdd";
                                        sql += @" AND RTRIM(LTRIM(TG002)) LIKE @ReferenceTime + '" + new string('_', lineLength) + @"'";
                                        dynamicParameters.Add("ReferenceTime", referenceTime.ToString(dateFormat));
                                        GrErpNo = referenceTime.ToString(dateFormat);
                                        break;
                                    case "2": //月編
                                        dateFormat = new string('y', yearLength) + "MM";
                                        sql += @" AND RTRIM(LTRIM(TG002)) LIKE @ReferenceTime + '" + new string('_', lineLength) + @"'";
                                        dynamicParameters.Add("ReferenceTime", referenceTime.ToString(dateFormat));
                                        GrErpNo = referenceTime.ToString(dateFormat);
                                        break;
                                    case "3": //流水號
                                        break;
                                    case "4": //手動編號
                                        break;
                                    default:
                                        throw new SystemException("編碼方式錯誤!");
                                }
                                #endregion

                                var currentNum = sqlConnection2.QueryFirstOrDefault(sql, dynamicParameters).CurrentNum;
                                GrErpNo += string.Format("{0:" + new string('0', lineLength) + "}", currentNum);
                                goodsReceipts[0].GrErpNo = GrErpNo;
                                #endregion
                                #endregion

                                #region //INSERT PURTG 進貨單單頭
                                foreach (var item in goodsReceipts)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO PURTG (COMPANY, CREATOR, USR_GROUP, CREATE_DATE, FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID
                                            , TG001, TG002, TG003, TG004, TG005, TG006, TG007, TG008, TG009, TG010, TG011, TG012, TG013, TG014, TG015, TG016, TG017
                                            , TG018, TG019, TG020, TG021, TG022, TG023, TG024, TG025, TG026, TG027, TG028, TG029, TG030, TG031, TG032, TG033, TG034
                                            , TG035, TG036, TG037, TG038, TG039, TG040, TG041, TG042, TG043, TG044, TG045, TG046, TG047, TG048, TG049, TG050, TG051
                                            , TG052, TG053, TG054, TG055, TG056, TG057, TG058, TG059, TG060, TG061, TG062, TG063, TG064, TG065, TG066, TG067, TG068
                                            , TG500, TG550, TG069, TG070, TG071, TG072, TG073, TG074, TG075, TG076, TG077)
                                            VALUES (@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID
                                            , @TG001, @TG002, @TG003, @TG004, @TG005, @TG006, @TG007, @TG008, @TG009, @TG010, @TG011, @TG012, @TG013, @TG014, @TG015, @TG016, @TG017
                                            , @TG018, @TG019, @TG020, @TG021, @TG022, @TG023, @TG024, @TG025, @TG026, @TG027, @TG028, @TG029, @TG030, @TG031, @TG032, @TG033, @TG034
                                            , @TG035, @TG036, @TG037, @TG038, @TG039, @TG040, @TG041, @TG042, @TG043, @TG044, @TG045, @TG046, @TG047, @TG048, @TG049, @TG050, @TG051
                                            , @TG052, @TG053, @TG054, @TG055, @TG056, @TG057, @TG058, @TG059, @TG060, @TG061, @TG062, @TG063, @TG064, @TG065, @TG066, @TG067, @TG068
                                            , @TG500, @TG550, @TG069, @TG070, @TG071, @TG072, @TG073, @TG074, @TG075, @TG076, @TG077)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            COMPANY = companyNo,
                                            CREATOR = UserNo,
                                            USR_GROUP,
                                            CREATE_DATE = DateTime.Now.ToString("yyyyMMdd"),
                                            FLAG = 1,
                                            CREATE_TIME = DateTime.Now.ToString("HH:mm:ss"),
                                            CREATE_AP = UserNo + "PC",
                                            CREATE_PRID = "BM",
                                            TG001 = item.GrErpPrefix,
                                            TG002 = item.GrErpNo,
                                            TG003 = item.ErpReceiptDate,
                                            TG004 = factoryNo,
                                            TG005 = item.SupplierNo,
                                            TG006 = item.SupplierSo,
                                            TG007 = item.CurrencyCode,
                                            TG008 = item.Exchange,
                                            TG009 = item.InvoiceType,
                                            TG010 = item.TaxType,
                                            TG011 = item.InvoiceNo,
                                            TG012 = item.PrintCnt,
                                            TG013 = item.ConfirmStatus,
                                            TG014 = item.ErpDocDate,
                                            TG015 = item.RenewFlag,
                                            TG016 = item.Remark,
                                            TG017 = item.TotalAmount,
                                            TG018 = item.DeductAmount,
                                            TG019 = item.OrigTax,
                                            TG020 = item.ReceiptAmount,
                                            TG021 = item.SupplierName,
                                            TG022 = item.UiNo,
                                            TG023 = item.DeductType,
                                            TG024 = item.ObaccoAndLiquorFlag,
                                            TG025 = item.RowCnt,
                                            TG026 = item.Quantity,
                                            TG027 = InvoiceDate,
                                            TG028 = item.OrigPreTaxAmount,
                                            TG029 = item.ApplyYYMM,
                                            TG030 = item.TaxRate,
                                            TG031 = item.PreTaxAmount,
                                            TG032 = item.TaxAmount,
                                            TG033 = item.PaymentTerm,
                                            TG034 = item.PoErpPrefix,
                                            TG035 = item.PoErpNo,
                                            TG036 = item.PrepaidErpPrefix,
                                            TG037 = item.PrepaidErpNo,
                                            TG038 = item.OffsetAmount,
                                            TG039 = item.TaxOffset,
                                            TG040 = item.PackageQuantity,
                                            TG041 = item.PremiumAmount,
                                            TG042 = item.ApproveStatus,
                                            TG043 = item.InvoiceFlag,
                                            TG044 = item.ReserveTaxCode,
                                            TG045 = item.SendCount,
                                            TG046 = item.OrigPremiumAmount,
                                            TG047 = item.EbcErpPreNo,
                                            TG048 = item.EbcEdition,
                                            TG049 = item.EbcFlag,
                                            TG050 = item.FromDocType,
                                            TG051 = "",
                                            TG052 = "",
                                            TG053 = 0.000000,
                                            TG054 = 0.000000,
                                            TG055 = item.NoticeFlag,
                                            TG056 = "",
                                            TG057 = "",
                                            TG058 = item.TaxCode,
                                            TG059 = item.TradeTerm,
                                            TG060 = "",
                                            TG061 = "",
                                            TG062 = "",
                                            TG063 = item.DetailMultiTax,
                                            TG064 = item.TaxExchange,
                                            TG065 = 0,
                                            TG066 = 0,
                                            TG067 = item.ContactUser,
                                            TG068 = "",
                                            TG500 = "",
                                            TG550 = "",
                                            TG069 = "",
                                            TG070 = "",
                                            TG071 = "",
                                            TG072 = "",
                                            TG073 = "",
                                            TG074 = "",
                                            TG075 = "",
                                            TG076 = "",
                                            TG077 = ""
                                        });
                                    rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);
                                }
                                #endregion
                            }
                            #endregion

                            #region //新增單身PURTH資料
                            foreach (var item in grDetails)
                            {
                                #region //確認ERP採購單身結案碼
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT LTRIM(RTRIM(a.TD016)) TD016
                                        , LTRIM(RTRIM(a.TD018)) TD018
                                        , LTRIM(RTRIM(b.TC004)) TC004
                                        FROM PURTD a 
                                        INNER JOIN PURTC b ON a.TD001 = b.TC001 AND a.TD002 = b.TC002
                                        WHERE a.TD001 = @PoErpPrefix
                                        AND a.TD002 = @PoErpNo
                                        AND a.TD003 = @PoSeq";
                                dynamicParameters.Add("PoErpPrefix", item.PoErpPrefix);
                                dynamicParameters.Add("PoErpNo", item.PoErpNo);
                                dynamicParameters.Add("PoSeq", item.PoSeq);

                                var PURTDResult = sqlConnection2.Query(sql, dynamicParameters);

                                if (PURTDResult.Count() <= 0) throw new SystemException("ERP採購單單身【" + item.PoErpPrefix + "-" + item.PoErpNo + "(" + item.PoSeq + ")】資料錯誤!!");

                                foreach (var item2 in PURTDResult)
                                {
                                    if (item.ConfirmStatus != "Y") {
                                        if (item2.TD016 != "N") throw new SystemException("ERP採購單單身【" + item.PoErpPrefix + "-" + item.PoErpNo + "(" + item.PoSeq + ")】已結案!!");
                                    }
                                    if (item2.TD018 != "Y") throw new SystemException("ERP採購單單身【" + item.PoErpPrefix + "-" + item.PoErpNo + "(" + item.PoSeq + ")】狀態錯誤!!");
                                    if (item2.TC004 != supplierNo) throw new SystemException("ERP採購單單身【" + item.PoErpPrefix + "-" + item.PoErpNo + "(" + item.PoSeq + ")】供應商與進貨單供應商不同!!");
                                }
                                #endregion

                                #region //確認ERP品號相關資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT LTRIM(RTRIM(MB004)) MB004
                                        , LTRIM(RTRIM(MB030)) MB030
                                        , LTRIM(RTRIM(MB031)) MB031
                                        , LTRIM(RTRIM(MB022)) MB022
                                        FROM INVMB
                                        WHERE MB001 = @MB001";
                                dynamicParameters.Add("MB001", item.MtlItemNo);

                                var INVMBResult = sqlConnection2.Query(sql, dynamicParameters);

                                if (INVMBResult.Count() <= 0) throw new SystemException("ERP品號基本資料錯誤!!");

                                foreach (var item2 in INVMBResult)
                                {
                                    #region //判斷ERP品號生效日與失效日
                                    if (item2.MB030 != "" && item2.MB030 != null)
                                    {
                                        #region //判斷單據日期需大於或等於生效日
                                        string EffectiveDate = item2.MB030;
                                        string effYear = EffectiveDate.Substring(0, 4);
                                        string effMonth = EffectiveDate.Substring(4, 2);
                                        string effDay = EffectiveDate.Substring(6, 2);
                                        DateTime effFullDate = Convert.ToDateTime(effYear + "-" + effMonth + "-" + effDay);
                                        int effresult = DateTime.Compare((DateTime)goodsReceipts[0].DocDate, effFullDate);
                                        if (effresult < 0) throw new SystemException("此品號尚未生效，無法使用!!");
                                        #endregion
                                    }

                                    if (item2.MB031 != "" && item2.MB031 != null)
                                    {
                                        #region //判斷日期需小於或等於失效日
                                        string ExpirationDate = item2.MB031;
                                        string effYear = ExpirationDate.Substring(0, 4);
                                        string effMonth = ExpirationDate.Substring(4, 2);
                                        string effDay = ExpirationDate.Substring(6, 2);
                                        DateTime effFullDate = Convert.ToDateTime(effYear + "-" + effMonth + "-" + effDay);
                                        int effresult = DateTime.Compare((DateTime)goodsReceipts[0].DocDate, effFullDate);
                                        if (effresult > 0) throw new SystemException("此品號已失效，無法使用!!");
                                        #endregion
                                    }
                                    #endregion

                                    #region //確認此品號是否需要批號管理
                                    if (item2.MB022 == "Y" || item2.MB022 == "T")
                                    {
                                        if (item.LotNumber.Length <= 0) throw new SystemException("品號【"+item.MtlItemNo+"】需設定批號!!");
                                    }
                                    else
                                    {
                                        if (item.LotNumber.Length > 0) throw new SystemException("品號【" + item.MtlItemNo + "】不需設定批號!!");
                                    }
                                    #endregion
                                }
                                #endregion

                                #region //整理單身日期格式
                                string AcceptanceDate = ((DateTime)item.AcceptanceDate).ToString("yyyy-MM-dd");
                                string AvailableDate = item.AvailableDate == null ? "" : ((DateTime)item.AvailableDate).ToString("yyyyMMdd");
                                string ReCheckDate = item.ReCheckDate == null ? "" : ((DateTime)item.ReCheckDate).ToString("yyyyMMdd");
                                #endregion

                                #region //確認此進貨單是否有進貨檢驗量測單據
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 a.QcRecordId, a.QcStatus
                                        FROM MES.QcGoodsReceipt a 
                                        WHERE a.GrDetailId = @GrDetailId
                                        ORDER BY a.CreateDate DESC";
                                dynamicParameters.Add("GrDetailId", item.GrDetailId);

                                var QcGoodsReceiptResult = sqlConnection.Query(sql, dynamicParameters);

                                if (QcGoodsReceiptResult.Count() > 0)
                                {
                                    bool qcFlag = false;
                                    foreach (var item2 in QcGoodsReceiptResult)
                                    {
                                        if (item2.QcStatus == "A")
                                        {
                                            //throw new SystemException("此進貨單身已開立進貨檢驗單據【" + item2.QcRecordId + "】，但尚未量測完成!!");
                                        }

                                        qcFlag = true;

                                        #region //確認此量測單據是否有品異單
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT TOP 1 a.JudgeStatus, a.QcRecordId, a.ReleaseQty
                                                , b.AbnormalqualityNo
                                                FROM QMS.AqBarcode a 
                                                INNER JOIN QMS.Abnormalquality b ON a.AbnormalqualityId = b.AbnormalqualityId
                                                WHERE a.QcRecordId = @QcRecordId";
                                        dynamicParameters.Add("QcRecordId", item2.QcRecordId);

                                        var AqBarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                                        if (AqBarcodeResult.Count() > 0)
                                        {
                                            foreach (var item3 in AqBarcodeResult)
                                            {
                                                if (item3.JudgeStatus == null)
                                                {
                                                    throw new SystemException("品異單據【" + item3.AbnormalqualityNo + "】尚未進行判定!!");
                                                }
                                                else if (item3.JudgeStatus == "AM")
                                                {
                                                    qcFlag = false;
                                                }
                                                else if (item3.JudgeStatus == "R")
                                                {
                                                    if (item3.ReleaseQty == null) throw new SystemException("品異單據【" + item3.AbnormalqualityNo + "】尚未維護特採數量!!");

                                                    #region //確認驗收數量是否正確
                                                    if (item.AcceptQty > item3.ReleaseQty) throw new SystemException("驗收數量大於特採數量!!");
                                                    #endregion
                                                }
                                                else if (item3.JudgeStatus == "S")
                                                {
                                                    #region //若判定為報廢，驗收數量需為0
                                                    if (item.AcceptQty != 0) throw new SystemException("此進貨單身判定為報廢，驗收數量需為0!!");
                                                    #endregion
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (item2.QcStatus != "Y")
                                            {
                                                //throw new SystemException("進貨檢驗單號【" + item2.QcRecordId + "】狀態非合格，但查無品異單據!!");
                                            }
                                        }
                                        #endregion
                                    }

                                    if (qcFlag == false)
                                    {
                                        //throw new SystemException("此進貨單身已有開立品異單判定為【重新量測】，但尚未有新量測單據!!");
                                    }
                                }
                                else
                                {
                                    if (item.QcStatus == "1") throw new SystemException("品號【" + item.MtlItemNo + "】為待驗狀態，拋轉前需至少建立一張進貨檢驗單據!!");
                                }
                                #endregion

                                #region //確認是否已有核單者
                                string confirmUserNo = "";
                                if (item.ConfirmUser != null && item.ConfirmUser > 0)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT UserNo FROM BAS.[User]
                                            WHERE UserId = @UserId";
                                    dynamicParameters.Add("UserId", item.ConfirmUser);

                                    var ConfirmUserResult = sqlConnection.Query(sql, dynamicParameters);

                                    if (ConfirmUserResult.Count() <= 0)
                                    {
                                        throw new SystemException($"單身【{item.GrSeq}】: 查詢核單者資料時錯誤!");
                                    }

                                    foreach (var item2 in ConfirmUserResult)
                                    {
                                        confirmUserNo = item2.UserNo;
                                    }
                                }
                                #endregion

                                #region //INSERT PURTH
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO PURTH (COMPANY, CREATOR, USR_GROUP, CREATE_DATE, FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID
                                        , TH001, TH002, TH003, TH004, TH005, TH006, TH007, TH008, TH009, TH010, TH011, TH012, TH013, TH014, TH015
                                        , TH016, TH017, TH018, TH019, TH020, TH021, TH022, TH023, TH024, TH025, TH026, TH027, TH028, TH029, TH030
                                        , TH031, TH032, TH033, TH034, TH035, TH036, TH037, TH038, TH039, TH040, TH041, TH042, TH043, TH044, TH045
                                        , TH046, TH047, TH048, TH049, TH050, TH051, TH052, TH053, TH054, TH055, TH056, TH057, TH058, TH059, TH060
                                        , TH061, TH062, TH063, TH064, TH065, TH066, TH067, TH068, TH069, TH070, TH071, TH072, TH073, TH074, TH500
                                        , TH501, TH502, TH503, TH550, TH551, TH552, TH553, TH554, TH555, TH556, TH557, TH558, TH075, TH076, TH077
                                        , TH078, TH079, TH080, TH081, TH082, TH083, TH084, TH085, TH086, TH087, TH088, TH089, TH090, TH091, TH092)
                                        VALUES (@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID
                                        , @TH001, @TH002, @TH003, @TH004, @TH005, @TH006, @TH007, @TH008, @TH009, @TH010, @TH011, @TH012, @TH013, @TH014, @TH015
                                        , @TH016, @TH017, @TH018, @TH019, @TH020, @TH021, @TH022, @TH023, @TH024, @TH025, @TH026, @TH027, @TH028, @TH029, @TH030
                                        , @TH031, @TH032, @TH033, @TH034, @TH035, @TH036, @TH037, @TH038, @TH039, @TH040, @TH041, @TH042, @TH043, @TH044, @TH045
                                        , @TH046, @TH047, @TH048, @TH049, @TH050, @TH051, @TH052, @TH053, @TH054, @TH055, @TH056, @TH057, @TH058, @TH059, @TH060
                                        , @TH061, @TH062, @TH063, @TH064, @TH065, @TH066, @TH067, @TH068, @TH069, @TH070, @TH071, @TH072, @TH073, @TH074, @TH500
                                        , @TH501, @TH502, @TH503, @TH550, @TH551, @TH552, @TH553, @TH554, @TH555, @TH556, @TH557, @TH558, @TH075, @TH076, @TH077
                                        , @TH078, @TH079, @TH080, @TH081, @TH082, @TH083, @TH084, @TH085, @TH086, @TH087, @TH088, @TH089, @TH090, @TH091, @TH092)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        COMPANY = companyNo,
                                        CREATOR = UserNo,
                                        USR_GROUP,
                                        CREATE_DATE = DateTime.Now.ToString("yyyyMMdd"),
                                        FLAG = 1,
                                        CREATE_TIME = DateTime.Now.ToString("HH:mm:ss"),
                                        CREATE_AP = UserNo + "PC",
                                        CREATE_PRID = "BM",
                                        TH001 = goodsReceipts[0].GrErpPrefix,
                                        TH002 = goodsReceipts[0].GrErpNo,
                                        TH003 = item.GrSeq,
                                        TH004 = item.MtlItemNo,
                                        TH005 = item.GrMtlItemName,
                                        TH006 = item.GrMtlItemSpec,
                                        TH007 = item.ReceiptQty,
                                        TH008 = item.UomNo,
                                        TH009 = item.InventoryNo,
                                        TH010 = item.LotNumber,
                                        TH011 = item.PoErpPrefix,
                                        TH012 = item.PoErpNo,
                                        TH013 = item.PoSeq,
                                        TH014 = item.ErpAcceptanceDate,
                                        TH015 = item.AcceptQty,
                                        TH016 = item.AvailableQty,
                                        TH017 = item.ReturnQty,
                                        TH018 = item.OrigUnitPrice,
                                        TH019 = item.OrigAmount,
                                        TH020 = item.OrigDiscountAmt,
                                        TH021 = item.TrErpPrefix,
                                        TH022 = item.TrErpNo,
                                        TH023 = item.TrSeq,
                                        TH024 = item.ReceiptExpense,
                                        TH025 = item.DiscountDescription,
                                        TH026 = item.PaymentHold,
                                        TH027 = item.Overdue,
                                        TH028 = item.QcStatus,
                                        TH029 = item.ReturnCode,
                                        TH030 = item.ConfirmStatus,
                                        TH031 = item.CloseStatus,
                                        TH032 = item.ReNewStatus,
                                        TH033 = item.Remark,
                                        TH034 = item.InventoryQty,
                                        TH035 = item.SmallUnit,
                                        TH036 = AvailableDate,
                                        TH037 = ReCheckDate,
                                        TH038 = item.ConfirmUser,
                                        TH039 = item.ApvErpPrefix,
                                        TH040 = item.ApvErpNo,
                                        TH041 = item.ApvSeq,
                                        TH042 = item.ProjectCode,
                                        TH043 = item.ExpenseEntry,
                                        TH044 = item.PremiumAmountFlag,
                                        TH045 = item.OrigPreTaxAmt,
                                        TH046 = item.OrigTaxAmt,
                                        TH047 = item.PreTaxAmt,
                                        TH048 = item.TaxAmt,
                                        TH049 = item.ReceiptPackageQty,
                                        TH050 = item.AcceptancePackageQty,
                                        TH051 = item.ReturnPackageQty,
                                        TH052 = item.PremiumAmount,
                                        TH053 = item.PackageUnit,
                                        TH054 = item.ReserveTaxCode,
                                        TH055 = item.OrigPremiumAmount,
                                        TH056 = item.AvailableUomNo,
                                        TH057 = item.OrigCustomer,
                                        TH058 = item.ApproveStatus,
                                        TH059 = item.EbcErpPreNo,
                                        TH060 = item.EbcEdition,
                                        TH061 = item.ProductSeqAmount,
                                        TH062 = item.MtlItemType,
                                        TH063 = item.Loaction,
                                        TH064 = "",
                                        TH065 = "",
                                        TH066 = "",
                                        TH067 = 0.000000,
                                        TH068 = 0.000000,
                                        TH069 = "",
                                        TH070 = "",
                                        TH071 = "",
                                        TH072 = "",
                                        TH073 = item.TaxRate,
                                        TH074 = item.MultipleFlag,
                                        TH500 = "",
                                        TH501 = "",
                                        TH502 = "",
                                        TH503 = "",
                                        TH550 = "",
                                        TH551 = "",
                                        TH552 = "",
                                        TH553 = "",
                                        TH554 = "",
                                        TH555 = "",
                                        TH556 = "",
                                        TH557 = "",
                                        TH558 = "",
                                        TH075 = "",
                                        TH076 = "",
                                        TH077 = "",
                                        TH078 = "",
                                        TH079 = "",
                                        TH080 = "",
                                        TH081 = "",
                                        TH082 = item.GrErpStatus,
                                        TH083 = "",
                                        TH084 = "",
                                        TH085 = "",
                                        TH086 = "",
                                        TH087 = "",
                                        TH088 = item.TaxCode,
                                        TH089 = item.DiscountRate,
                                        TH090 = item.DiscountPrice,
                                        TH091 = item.QcType,
                                        TH092 = ""
                                    });
                                rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);
                                #endregion
                            }
                            #endregion

                            #region //UPDATE SCM.GoodsReceipt
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.GoodsReceipt SET
                                    GrErpNo = @GrErpNo,
                                    TransferStatus = 'Y',
                                    TransferDate = @TransferDate,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE GrId = @GrId";
                            dynamicParameters.AddDynamicParams(
                              new
                              {
                                  goodsReceipts[0].GrErpNo,
                                  TransferDate = DateTime.Now,
                                  LastModifiedDate,
                                  LastModifiedBy,
                                  GrId
                              });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //UPDATE SCM.GrDetail
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.GrDetail SET
                                    GrErpNo = @GrErpNo,
                                    TransferStatus = 'Y',
                                    TransferDate = @TransferDate,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE GrId = @GrId";
                            dynamicParameters.AddDynamicParams(
                              new
                              {
                                  goodsReceipts[0].GrErpNo,
                                  TransferDate = DateTime.Now,
                                  LastModifiedDate,
                                  LastModifiedBy,
                                  GrId
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

        #region //UpdateGrLotNumber -- 更新進貨單身批號資料 -- Ann 2024-05-30
        public string UpdateGrLotNumber(int GrDetailId, string LotNumber)
        {
            try
            {
                int rowsAffected = 0;
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認進貨單身資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.GrId, a.ConfirmStatus
                                FROM SCM.GrDetail a 
                                WHERE a.GrDetailId = @GrDetailId";
                        dynamicParameters.Add("GrDetailId", GrDetailId);

                        var GrDetailResult = sqlConnection.Query(sql, dynamicParameters);

                        if (GrDetailResult.Count() <= 0) throw new SystemException("進貨單身資料錯誤!!");

                        foreach (var item in GrDetailResult)
                        {
                            if (item.ConfirmStatus != "N")
                            {
                                throw new SystemException("進貨單狀態無法更改批號!!");
                            }
                        }
                        #endregion

                        #region //UPDATE SCM.GrDetail LotNumber
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.GrDetail SET
                                LotNumber = @LotNumber,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE GrDetailId = @GrDetailId";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            LotNumber,
                            LastModifiedDate,
                            LastModifiedBy,
                            GrDetailId
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

        #region //UpdateReceiptQty -- 更新進貨數量 -- Ann 2024-11-06
        public string UpdateReceiptQty(string UploadJsonString)
        {
            try
            {
                if (UploadJsonString.Length <= 0) throw new SystemException("上傳JSON不能為空!!");
                int rowsAffected = 0;
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
                            JObject uploadJson = JObject.Parse(UploadJsonString);
                            foreach (var item in uploadJson["uploadInfo"])
                            {
                                #region //確認進貨單身資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.GrId, a.GrSeq, a.ReceiptQty, a.PoErpPrefix, a.PoErpNo, a.PoSeq
                                    , a.ConfirmStatus, a.AcceptQty, a.AvailableQty, a.QcStatus, a.LotNumber
                                    , a.OrigUnitPrice, a.OrigDiscountAmt, a.OrigAmount
                                    , b.Quantity, b.SiQty
                                    , c.LotManagement
                                    , d.DocDate, d.TaxRate, d.TaxType, d.Exchange, d.CurrencyCode
                                    FROM SCM.GrDetail a 
                                    INNER JOIN SCM.PoDetail b ON a.PoErpPrefix = b.PoErpPrefix AND a.PoErpNo = b.PoErpNo AND a.PoSeq = b.PoSeq
                                    INNER JOIN PDM.MtlItem c ON a.MtlItemId = c.MtlItemId
                                    INNER JOIN SCM.GoodsReceipt d ON a.GrId = d.GrId
                                    WHERE a.GrDetailId = @GrDetailId";
                                dynamicParameters.Add("GrDetailId", Convert.ToInt32(item["GrDetailId"]));

                                var GrDetailResult = sqlConnection.Query(sql, dynamicParameters);

                                if (GrDetailResult.Count() <= 0) throw new SystemException("進貨單身資料錯誤!!");

                                int ReceiptQty = 0;
                                int AcceptQty = 0;
                                int AvailableQty = 0;
                                int GrId = 0;
                                string GrSeq = "";
                                double OrigUnitPrice = 0;
                                double OrigDiscountAmt = 0;
                                double OrigAmount = 0;
                                double TaxRate = 0;
                                string TaxType = "";
                                double Exchange = 0;
                                string QcStatus = "";
                                string OriCurrencyCode = "";
                                DateTime DocDate = new DateTime();
                                foreach (var item2 in GrDetailResult)
                                {
                                    GrId = item2.GrId;
                                    GrSeq = item2.GrSeq;
                                    ReceiptQty = item2.ReceiptQty;
                                    AcceptQty = item2.AcceptQty;
                                    AvailableQty = item2.AvailableQty;
                                    OrigUnitPrice = item2.OrigUnitPrice;
                                    OrigDiscountAmt = item2.OrigDiscountAmt;
                                    OrigAmount = item2.OrigAmount;
                                    TaxRate = item2.TaxRate;
                                    TaxType = item2.TaxType;
                                    Exchange = item2.Exchange;
                                    QcStatus = item2.QcStatus;
                                    OriCurrencyCode = item2.CurrencyCode;
                                    DocDate = item2.DocDate;

                                    #region //卡控檢核
                                    if (item2.ConfirmStatus != "N")
                                    {
                                        throw new SystemException("進貨序號【" + GrSeq + "】<br>進貨單狀態無法更改批號!!");
                                    }

                                    if (item2.QcStatus != "0" && (item2.AcceptQty > 0 || item2.AvailableQty > 0))
                                    {
                                        throw new SystemException("進貨序號【" + GrSeq + "】<br>進貨單已有驗收或計價數量，無法更改進貨數量!!");
                                    }

                                    if (item2.QcStatus != "0" && item2.QcStatus != "1")
                                    {
                                        throw new SystemException("進貨序號【" + GrSeq + "】<br>進貨單身檢驗狀態錯誤，無法更改進貨數量!!");
                                    }

                                    if (item2.LotManagement != "T" && item["LotNumber"].ToString() != "")
                                    {
                                        throw new SystemException("進貨序號【" + GrSeq + "】<br>此單身品號不需維護批號!!");
                                    }
                                    #endregion

                                    #region //若免檢則將驗收及計價同步進貨數量
                                    if (item2.QcStatus == "0")
                                    {
                                        AcceptQty = Convert.ToInt32(item["ReceiptQty"]);
                                        AvailableQty = Convert.ToInt32(item["ReceiptQty"]);
                                    }
                                    #endregion

                                    #region //確認此進貨單身無相同採購單身
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT ISNULL(SUM(a.ReceiptQty), 0) TotalReceiptQty
                                            FROM SCM.GrDetail a 
                                            WHERE a.PoErpPrefix = @PoErpPrefix
                                            AND a.PoErpNo = @PoErpNo
                                            AND a.PoSeq = @PoSeq
                                            AND a.GrId = @GrId
                                            AND a.GrDetailId != @GrDetailId";
                                    dynamicParameters.Add("PoErpPrefix", item2.PoErpPrefix);
                                    dynamicParameters.Add("PoErpNo", item2.PoErpNo);
                                    dynamicParameters.Add("PoSeq", item2.PoSeq);
                                    dynamicParameters.Add("GrId", item2.GrId);
                                    dynamicParameters.Add("GrDetailId", Convert.ToInt32(item["GrDetailId"]));

                                    var PoDetailResult = sqlConnection.Query(sql, dynamicParameters);

                                    foreach (var item3 in PoDetailResult)
                                    {
                                        #region //確認數量沒有超過剩餘可進貨數量
                                        double poQty = Convert.ToDouble(item2.Quantity);
                                        double siQty = Convert.ToDouble(item2.SiQty);
                                        double TotalReceiptQty = item3.TotalReceiptQty;
                                        double allowQty = poQty - siQty - TotalReceiptQty > 0 ? poQty - siQty - TotalReceiptQty : 0;

                                        if (Convert.ToDouble(item["ReceiptQty"]) > allowQty)
                                        {
                                            throw new SystemException("進貨序號【" + GrSeq + "】<br>此次修改進貨數量已超過剩餘可進貨數量: " + allowQty + "!!");
                                        }
                                        #endregion
                                    }
                                    #endregion
                                }
                                #endregion

                                #region //原幣小數點取位
                                int OriUnitDecimal = 0;
                                int OriPriceDecimal = 0;
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.MF003, a.MF004
                                        FROM CMSMF a 
                                        WHERE a.MF001 = @Currency";
                                dynamicParameters.Add("Currency", OriCurrencyCode);

                                var CMSMFResult = sqlConnection2.Query(sql, dynamicParameters);

                                if (CMSMFResult.Count() <= 0) throw new SystemException("金額小數點取位資料錯誤!!");

                                foreach (var item2 in CMSMFResult)
                                {
                                    OriUnitDecimal = Convert.ToInt32(item2.MF003);
                                    OriPriceDecimal = Convert.ToInt32(item2.MF004);
                                }
                                #endregion

                                #region //取得本幣幣別
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 LTRIM(RTRIM(MA003)) CurrencyCode
                                        FROM CMSMA
                                        WHERE 1=1";

                                var CMSMAResult = sqlConnection2.Query(sql, dynamicParameters);

                                string CurrencyCode = "";
                                foreach (var item2 in CMSMAResult)
                                {
                                    CurrencyCode = item2.CurrencyCode;
                                }
                                #endregion

                                #region //本幣小數點取位
                                int UnitDecimal = 0;
                                int PriceDecimal = 0;
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.MF003, a.MF004
                                        FROM CMSMF a 
                                        WHERE a.MF001 = @Currency";
                                dynamicParameters.Add("Currency", CurrencyCode);

                                var CMSMFResult2 = sqlConnection2.Query(sql, dynamicParameters);

                                if (CMSMFResult2.Count() <= 0) throw new SystemException("金額小數點取位資料錯誤!!");

                                foreach (var item2 in CMSMFResult2)
                                {
                                    UnitDecimal = Convert.ToInt32(item2.MF003);
                                    PriceDecimal = Convert.ToInt32(item2.MF004);
                                }
                                #endregion

                                #region //處理【原幣進貨金額】、【原幣未稅金額】、【原幣稅金額】、【本幣未稅金額】、【本幣稅金額】
                                double OrigPreTaxAmt = 0; //原幣未稅金額
                                double OrigTaxAmt = 0; //原幣稅金額
                                double PreTaxAmt = 0; //本幣未稅金額
                                double TaxAmt = 0; //本幣稅金額

                                JObject calculateTaxAmtResult = JObject.Parse(CalculateTaxAmt(AvailableQty, OrigUnitPrice, OrigDiscountAmt, OrigAmount, TaxRate, TaxType, Exchange
                                                                                , OriUnitDecimal, OriPriceDecimal, UnitDecimal, PriceDecimal));

                                if (calculateTaxAmtResult["status"].ToString() == "success")
                                {
                                    OrigPreTaxAmt = Convert.ToDouble(calculateTaxAmtResult["OrigPreTaxAmt"]);
                                    OrigTaxAmt = Convert.ToDouble(calculateTaxAmtResult["OrigTaxAmt"]);
                                    PreTaxAmt = Convert.ToDouble(calculateTaxAmtResult["PreTaxAmt"]);
                                    TaxAmt = Convert.ToDouble(calculateTaxAmtResult["TaxAmt"]);
                                }
                                else
                                {
                                    throw new SystemException(calculateTaxAmtResult["msg"].ToString());
                                }
                                #endregion

                                #region //確認庫別資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.InventoryNo
                                        FROM SCM.Inventory a 
                                        WHERE a.InventoryId = @InventoryId";
                                dynamicParameters.Add("InventoryId", Convert.ToInt32(item["InventoryId"]));

                                var InventoryResult = sqlConnection.Query(sql, dynamicParameters);

                                if (InventoryResult.Count() <= 0) throw new SystemException("進貨序號【" + GrSeq + "】<br>庫別資料錯誤!");

                                string InventoryNo = "";
                                foreach (var item2 in InventoryResult)
                                {
                                    InventoryNo = item2.InventoryNo;
                                }
                                #endregion

                                #region //確認驗收日期
                                DateTime dateValue;
                                if (DateTime.TryParse(item["AcceptanceDate"].ToString(), out dateValue))
                                {
                                    if (DocDate < dateValue)
                                    {
                                        throw new SystemException("進貨序號【" + GrSeq + "】<br>單據日期不可小於驗收日期!!");
                                    }
                                }
                                else
                                {
                                    throw new SystemException("進貨序號【" + GrSeq + "】<br>驗收日期格式錯誤!!");
                                }
                                #endregion

                                #region //UPDATE SCM.GrDetail ReceiptQty
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE SCM.GrDetail SET
                                        ReceiptQty = @ReceiptQty,
                                        AcceptQty = @AcceptQty,
                                        AvailableQty = @AvailableQty,
                                        OrigPreTaxAmt = @OrigPreTaxAmt,
                                        OrigTaxAmt = @OrigTaxAmt,
                                        PreTaxAmt = @PreTaxAmt,
                                        TaxAmt = @TaxAmt,
                                        InventoryId = @InventoryId,
                                        InventoryNo = @InventoryNo,
                                        LotNumber = @LotNumber,
                                        AcceptanceDate = @AcceptanceDate,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE GrDetailId = @GrDetailId";
                                dynamicParameters.AddDynamicParams(
                                new
                                {
                                    ReceiptQty = Convert.ToInt32(item["ReceiptQty"]),
                                    AcceptQty,
                                    AvailableQty,
                                    OrigPreTaxAmt,
                                    OrigTaxAmt,
                                    PreTaxAmt,
                                    TaxAmt,
                                    InventoryId = Convert.ToInt32(item["InventoryId"]),
                                    InventoryNo,
                                    LotNumber = item["LotNumber"].ToString(),
                                    AcceptanceDate = item["AcceptanceDate"].ToString(),
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    GrDetailId = Convert.ToInt32(item["GrDetailId"])
                                });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region //重新計算進貨單資料
                                JObject updateGrDetailResult = JObject.Parse(UpdateGrTotal(GrId, Exchange, sqlConnection, sqlConnection2));
                                if (updateGrDetailResult["status"].ToString() != "success")
                                {
                                    throw new SystemException(updateGrDetailResult["msg"].ToString());
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

        #region //UpdateGrConfirm -- 確認進貨單據 -- Shintokuro 2024-04-29
        public string UpdateGrConfirm(int GrId)
        {
            try
            {
                #region //Ann 2024-10-11更新: 確認前先進行拋轉
                var TransferResult = UpdateTransferGoodsReceipt(GrId);
                JObject transferResultJson = JObject.Parse(TransferResult);
                if (transferResultJson["status"].ToString() != "success")
                {
                    throw new SystemException(transferResultJson["msg"].ToString());
                }
                #endregion

                string dateNow = DateTime.Now.ToString("yyyyMMdd");
                string timeNow = DateTime.Now.ToString("HH:mm:ss");
                string dateNowMonth = DateTime.Now.ToString("yyyyMM");

                string UserNo = "";

                int rowsAffected = 0;
                var ErpNo = "";
                string ErpDbName = "";
                string USR_GROUP = "";
                int UserId = CreateBy;

                string GrErpPrefix = "";
                string GrErpNo = "";
                string ReceiptDate = "";
                int CompanyIdBase = -1;
                int GrDetailId = -1;

                string MaxReceiptDate = "";

                using (TransactionScope transactionScope = new TransactionScope())
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

                        var resultCompany = sqlConnection.Query(sql, dynamicParameters);

                        if (resultCompany.Count() <= 0) throw new SystemException("公司別錯誤!");

                        foreach (var item in resultCompany)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            ErpDbName = ErpConnectionStrings.Split('=')[2].Split(';')[0];
                            ErpNo = item.ErpNo;
                        }
                        #endregion

                        #region //撈取使用者No
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
                        dynamicParameters.Add("FunctionCode", "GoodsReceiptManagement");
                        dynamicParameters.Add("DetailCode", "confirm");

                        var resultUser = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUser.Count() <= 0) throw new SystemException("使用者資料錯誤!");

                        foreach (var item in resultUser)
                        {
                            UserNo = item.UserNo;
                        }
                        #endregion

                        #region /撈取單頭資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.GrErpPrefix,a.GrErpNo, a.ConfirmStatus,a.CompanyId
                                ,FORMAT(a.ReceiptDate, 'yyyyMMdd') ReceiptDate,a.TransferStatus
                                ,b.GrDetailId,b.CloseStatus,x.AcceptanceDateCount, b.AcceptQty, b.QcStatus, b.MtlItemNo
                                FROM SCM.GoodsReceipt a
                                INNER JOIN SCM.GrDetail b on a.GrId = b.GrId
                                OUTER APPLY(
                                    SELECT CASE WHEN COUNT(DISTINCT FORMAT(x1.AcceptanceDate, 'yyyyMM'))  >1 
                                    THEN 'N' ELSE 'Y' 
                                    END AS AcceptanceDateCount
                                    FROM SCM.GrDetail x1
                                    WHERE x1.GrId = a.GrId
                                ) x
                                WHERE a.GrId = @GrId";
                        dynamicParameters.Add("GrId", GrId);

                        var resultRoMes = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in resultRoMes)
                        {
                            if (item.TransferStatus != "Y") throw new SystemException("該進貨單單據尚未拋轉,請重新拋轉");
                            if (item.ConfirmStatus == "Y") throw new SystemException("該進貨單單據已經核單,不能操作");
                            if (item.ConfirmStatus == "V") throw new SystemException("該進貨單單據已經作廢,不能操作");
                            if (item.CompanyId != CurrentCompany) throw new SystemException("該進貨單單據公司別與後端公司別不一致,不能操作");
                            if (item.CloseStatus == "Y") throw new SystemException("該進貨單已經結帳,不能操作");
                            if (item.AcceptanceDateCount == "N") throw new SystemException("進貨單單身驗收日須同一天,才能操作");
                            CompanyIdBase = item.CompanyId;
                            GrErpPrefix = item.GrErpPrefix;
                            GrErpNo = item.GrErpNo;
                            ReceiptDate = item.ReceiptDate;
                            GrDetailId = item.GrDetailId;

                            #region //Ann 2024-10-11補充: 加上判斷進貨檢驗單據條件
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.QcRecordId, a.QcStatus
                                    , b.GrSeq
                                    FROM MES.QcGoodsReceipt a 
                                    INNER JOIN SCM.GrDetail b ON a.GrDetailId = b.GrDetailId
                                    LEFT JOIN MES.QcGoodsReceiptLog c ON a.QcGoodsReceiptId = c.QcGoodsReceiptId
                                    WHERE a.GrDetailId = @GrDetailId
                                    AND c.LogId IS NULL
                                    ORDER BY a.CreateDate DESC";
                            dynamicParameters.Add("GrDetailId", GrDetailId);

                            var QcGoodsReceiptResult = sqlConnection.Query(sql, dynamicParameters);

                            if (QcGoodsReceiptResult.Count() > 0)
                            {
                                bool qcFlag = false;
                                string GrSeq = "";
                                foreach (var item2 in QcGoodsReceiptResult)
                                {
                                    GrSeq = item2.GrSeq;

                                    if (item2.QcStatus == "A")
                                    {
                                        throw new SystemException("單身序號【" + item2.GrSeq + "】已開立進貨檢驗單據【" + item2.QcRecordId + "】，但尚未量測完成!!");
                                    }

                                    qcFlag = true;

                                    #region //確認此量測單據是否有品異單
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 a.JudgeStatus, a.QcRecordId, a.ReleaseQty
                                                , b.AbnormalqualityNo
                                                FROM QMS.AqBarcode a 
                                                INNER JOIN QMS.Abnormalquality b ON a.AbnormalqualityId = b.AbnormalqualityId
                                                WHERE a.QcRecordId = @QcRecordId";
                                    dynamicParameters.Add("QcRecordId", item2.QcRecordId);

                                    var AqBarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                                    if (AqBarcodeResult.Count() > 0)
                                    {
                                        foreach (var item3 in AqBarcodeResult)
                                        {
                                            if (item3.JudgeStatus == null)
                                            {
                                                throw new SystemException("品異單據【" + item3.AbnormalqualityNo + "】尚未進行判定!!");
                                            }
                                            else if (item3.JudgeStatus == "AM")
                                            {
                                                qcFlag = false;
                                            }
                                            else if (item3.JudgeStatus == "R")
                                            {
                                                if (item3.ReleaseQty == null) throw new SystemException("品異單據【" + item3.AbnormalqualityNo + "】尚未維護特採數量!!");

                                                #region //確認驗收數量是否正確
                                                if (item.AcceptQty > item3.ReleaseQty) throw new SystemException("驗收數量大於特採數量!!");
                                                #endregion
                                            }
                                            else if (item3.JudgeStatus == "S")
                                            {
                                                #region //若判定為報廢，驗收數量需為0
                                                if (item.AcceptQty != 0) throw new SystemException("單身序號【" + item2.GrSeq + "】判定為報廢，驗收數量需為0!!");
                                                #endregion
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (item2.QcStatus != "Y")
                                        {
                                            throw new SystemException("進貨檢驗單號【" + item2.QcRecordId + "】狀態非合格，但查無品異單據!!");
                                        }
                                    }
                                    #endregion
                                }

                                if (qcFlag == false)
                                {
                                    throw new SystemException("單身序號【" + GrSeq + "】已有開立品異單判定為【重新量測】，但尚未有新量測單據!!");
                                }
                            }
                            else
                            {
                                if (item.QcStatus == "1") throw new SystemException("品號【" + item.MtlItemNo + "】為待驗狀態，需進行檢驗!!");
                            }
                            #endregion
                        }
                        #endregion
                    }
                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //審核ERP權限
                        USR_GROUP = BaseHelper.CheckErpAuthority(UserNo, sqlConnection, "PURI09", "CONFIRM");
                        #endregion

                        #region //判斷開立日期是否超過結帳日
                        string CloseDateBase = "";
                        string ReceiveDateBase = ReceiptDate;
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT MA011,MA012,MA013
                                FROM CMSMA ";
                        var resultCloseDate = sqlConnection.Query(sql, dynamicParameters);
                        foreach (var item in resultCloseDate)
                        {
                            CloseDateBase = item.MA013;
                        }

                        DateTime docDateBase;
                        DateTime closeDateBase;

                        if (DateTime.TryParseExact(ReceiveDateBase, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out docDateBase) &&
                            DateTime.TryParseExact(CloseDateBase, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out closeDateBase))
                        {
                            int resultDate = DateTime.Compare(docDateBase, closeDateBase);

                            if (resultDate <= 0)
                            {
                                throw new SystemException("【進貨單】ERP已經結帳,不可以在新增【" + CloseDateBase + "】之前的單據");
                            }
                        }
                        else
                        {
                            throw new SystemException("日期字符串无效，无法比较");
                        }
                        #endregion

                        #region //取出進貨單單據資料 + 異動更新相關資料表
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.TG003,a.TG005,a.TG007,a.TG013,a.TG014
                                ,b.TH003,b.TH004,b.TH007,b.TH009,b.TH010
                                ,b.TH011,b.TH012,b.TH013
                                ,b.TH014,b.TH015,b.TH016,b.TH017,b.TH030,b.TH031,b.TH063
                                ,b.TH021,b.TH022,b.TH023
                                ,b.TH036,b.TH037,b.TH047
                                ,c.MB022
                                ,d.MQ019,md3.MF005
                                ,ISNULL(md1.GrRate,1) GrRate
                                ,ISNULL(md2.GrAvailableRate,1) GrAvailableRate
                                ,(b.TH015*ISNULL(md1.GrRate,1)) ConversionQty
                                ,(b.TH016*ISNULL(md2.GrAvailableRate,1)) ConversionAvailableQty
                                ,CASE 
                                    WHEN b.TH016 != 0 THEN ROUND((b.TH047 / (b.TH016*ISNULL(md2.GrAvailableRate,1))),CAST(md3.MF005 AS INT)) 
                                    ELSE 0
                                END AS UnitCost
                                FROM PURTG a
                                INNER JOIN PURTH b on a.TG001 = b.TH001 AND a.TG002 = b.TH002
                                INNER JOIN INVMB c on b.TH004 = c.MB001
                                INNER JOIN CMSMQ d ON d.MQ001=a.TG001 AND d.MQ003='34'
                                INNER JOIN CMSMC e ON e.MC001=b.TH009
                                LEFT JOIN INVMD f ON f.MD001=b.TH004
                                OUTER APPLY(
                                    SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) GrRate
                                    FROM INVMD x1
                                    WHERE x1.MD001=b.TH004
                                    AND  x1.MD002 = b.TH008
                                ) md1
                                OUTER APPLY(
                                    SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) GrAvailableRate
                                    FROM INVMD x1
                                    WHERE x1.MD001=b.TH004
                                    AND  x1.MD002 = b.TH056
                                ) md2
                                OUTER APPLY(
                                    SELECT x1.MF003,x1.MF004,x1.MF005,x1.MF006
                                    FROM CMSMF x1
                                    INNER JOIN CMSMA x2 on x1.MF001 = x2.MA003 
                                ) md3
                                WHERE a.TG001 = @GrErpPrefix
                                AND a.TG002 = @GrErpNo";
                        dynamicParameters.Add("GrErpPrefix", GrErpPrefix);
                        dynamicParameters.Add("GrErpNo", GrErpNo);
                        var resultRoErp = sqlConnection.Query(sql, dynamicParameters);
                        if (resultRoErp.Count() <= 0) throw new SystemException("找不到ERP進貨單單據!");
                        foreach (var item in resultRoErp)
                        {
                            if (item.TG013 == "Y") throw new SystemException("該進貨單單據已經核單,不能操作");
                            if (item.TG013 == "V") throw new SystemException("該進貨單單據已經作廢,不能操作");
                            //if (item.TH030 == "Y") throw new SystemException("該進貨單單據單身已經核單,不能操作");
                            if (item.TH031 == "Y") throw new SystemException("該進貨單單據已經結帳,不能操作");
                            if (item.TH021 != "") throw new SystemException("來源暫入單模式尚未開發,不能操作");
                            if (item.MQ019 == "N") throw new SystemException("不核對採購單模式尚未開發,不能操作");

                            #region //欄位參數撈取
                            string ReceiveDateErp = item.TG003; //TG003 進貨日期
                            string SupplierNo = item.TG005; //TG005 供應商
                            string CurrencyCode = item.TG007; //TG007 幣別
                            Double Exchange = Convert.ToDouble(item.TG008); //TG008 匯率
                            string ConfirmStatus = item.TG013; //TG013 單頭確認
                            string DocDate = item.TG014; //TG013 單頭確認

                            string GrSeq = item.TH003; //TH003 進貨單序號
                            string GrFull = GrErpPrefix + '-' + GrErpNo + '-' + GrSeq;
                            string MtlItemNo = item.TH004; //TH004 品號
                            string InventoryNo = item.TH009; //TH009 庫別
                            string LotNumber = item.TH010; //TH010 批號
                            string PoErpPrefix = item.TH011; //TH011 採購單別 單號 序號
                            string PoErpNo = item.TH012; //TH012 
                            string PoSeq = item.TH013; //TH013 
                            string PoFull = item.TH011 + '-' + item.TH012 + '-' + item.TH013;
                            string TH014 = item.TH014; //TH014 驗收日期
                            Double ReceiptQty = Convert.ToDouble(item.TH007); //TH014 進貨數量
                            Double AcceptQty = Convert.ToDouble(item.TH015); //TH014 驗收數量
                            Double AvailableQty = Convert.ToDouble(item.TH016); //TH014 計價數量 
                            Double ReturnQty = Convert.ToDouble(item.TH017); //TH014 驗退數量
                            string TH030 = item.TH030; //TH030 確認碼
                            string CloseStatus = item.TH031; //TH031 結帳碼
                            string AvailableDate = item.TH036; //TH036 有效日期
                            string ReCheckDate = item.TH037; //TH037 複檢日期
                            string Location = item.TH063; //TH060 儲位
                            string TrErpPrefix = item.TH021; //TH021 暫入單別 / TH022 暫入單號 / TH023 暫入序號
                            string TrErpNo = item.TH022;
                            string TrSeq = item.TH023;
                            string TsnFull = item.TH021 + '-' + item.TH022 + '-' + item.TH023;
                            string LotManagement = item.MB022; //品號批號管理

                            Double GrRate = Convert.ToDouble(item.GrRate); //單位換算率
                            Double ConversionQty = Convert.ToDouble(item.ConversionQty); //單位換算後數量
                            Double ConversionAvailableQty = Convert.ToDouble(item.ConversionAvailableQty); //單位換算後計價數量
                            Double UnitCost = Convert.ToDouble(item.UnitCost); //單位成本
                            Double TH047 = Convert.ToDouble(item.TH047); //本幣未稅金額
                            Double TH018 = Convert.ToDouble(item.TH018); //原幣單價
                            Double PoRate = 0; //採購單位換算率
                            Double PoAvailableRate = 0; //採購計價單位換算率
                            Double PoConversionQty = 0; //採購單位換算後數量
                            Double PoConversionAvailableQty = 0; //採購單位換算後計價數量

                            DateTime dateC = DateTime.ParseExact(DocDate, "yyyyMMdd", null);
                            DateTime dateD = DateTime.ParseExact(TH014, "yyyyMMdd", null);
                            if (dateC > dateD)
                            {
                                throw new SystemException("資料異常:單身驗收日不可以小於單據日!!");
                            }
                            if (MaxReceiptDate != "")
                            {
                                DateTime dateA = DateTime.ParseExact(item.TH014, "yyyyMMdd", null);
                                DateTime dateB = DateTime.ParseExact(MaxReceiptDate, "yyyyMMdd", null);
                                //MaxReceiptDate = date.ToString("yyyy-MM-dd"); // yyyy-MM-dd 格式


                                //DateTime dateA = DateTime.Parse(item.TH014);
                                //DateTime dateB = DateTime.Parse(MaxReceiptDate);
                                if (dateA > dateB)
                                {
                                    MaxReceiptDate = item.TH014;
                                }
                            }
                            else
                            {
                                MaxReceiptDate = item.TH014;
                            }

                            ReceiveDateErp = MaxReceiptDate;
                            #endregion

                            #region //檢查段

                            #region //數量查驗
                            if (ReceiptQty < 0) throw new SystemException("【進貨單:" + GrFull + "】進貨數量不可小於0!!!");
                            if (AcceptQty < 0) throw new SystemException("【進貨單:" + GrFull + "】驗收數量不可小於0!!!");
                            if (AvailableQty < 0) throw new SystemException("【進貨單:" + GrFull + "】計價數量不可小於0!!!");
                            if (ReceiptQty < AcceptQty) throw new SystemException("【進貨單:" + GrFull + "】驗收數量不可大於進貨數量!!!");
                            if (AcceptQty < AvailableQty) throw new SystemException("【進貨單:" + GrFull + "】計價數量不可大於驗收數量!!!");
                            #endregion

                            #region //PURTD 採購單
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 ISNULL(a.TD008,0) TD008,ISNULL(a.TD015,0) TD015,a.TD016,a.TD018,
                                    ISNULL(md1.PoRate,1) PoRate, 
                                    ISNULL(md2.PoAvailableRate,1) PoAvailableRate, 
                                    (a.TD008*ISNULL(md1.PoRate,1)) PoConversionQty,
                                    (a.TD015*ISNULL(md1.PoRate,1)) PoConversionSiQty,
                                    (a.TD058*ISNULL(md2.PoAvailableRate,1)) PoConversionAvailableQty,
                                    (a.TD060*ISNULL(md2.PoAvailableRate,1)) PoConversionAvailableSiQty
                                    FROM PURTD a
                                    OUTER APPLY(
                                        SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) PoRate
                                        FROM INVMD x1
                                        WHERE x1.MD001 = a.TD004
                                        AND  x1.MD002 = a.TD009
                                    ) md1
                                    OUTER APPLY(
                                        SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) PoAvailableRate
                                        FROM INVMD x1
                                        WHERE x1.MD001 = a.TD004
                                        AND  x1.MD002 = a.TD059
                                    ) md2
                                    WHERE a.TD001 = @TD001
                                    AND a.TD002 = @TD002
                                    AND a.TD003 = @TD003";
                            dynamicParameters.Add("TD001", PoErpPrefix);
                            dynamicParameters.Add("TD002", PoErpNo);
                            dynamicParameters.Add("TD003", PoSeq);
                            var resultCOPTD = sqlConnection.Query(sql, dynamicParameters);
                            if (resultCOPTD.Count() <= 0) throw new SystemException("採購單找不到,請重新確認");
                            foreach (var item1 in resultCOPTD)
                            {
                                if (item1.TD018 != "Y") throw new SystemException("【採購單:" + PoFull + "】非確認狀態,不可操作!!!");
                                if (item.TH030 != "Y") {
                                    if (item1.TD016 == "Y") throw new SystemException("【採購單:" + PoFull + "】已經結案!!!"); 

                                }
                                Double TD008 = Convert.ToDouble(item1.TD008);
                                Double TD015 = Convert.ToDouble(item1.TD015);
                                PoRate = Convert.ToDouble(item1.PoRate);
                                PoAvailableRate = Convert.ToDouble(item1.PoAvailableRate);
                                PoConversionQty = Convert.ToDouble(item1.PoConversionQty);
                                Double PoConversionSiQty = Convert.ToDouble(item1.PoConversionSiQty);
                                PoConversionAvailableQty = Convert.ToDouble(item1.PoConversionAvailableQty);
                                Double PoConversionAvailableSiQty = Convert.ToDouble(item1.PoConversionAvailableSiQty);
                                if (item.TH030 != "Y")
                                {
                                    if (PoConversionQty == PoConversionSiQty + ConversionQty)
                                    {
                                        CloseStatus = "Y";
                                    }
                                    else if (PoConversionQty < PoConversionSiQty + ConversionQty)
                                    {
                                        throw new SystemException("【採購單:" + PoFull + "】已超過採購數量,目前已交數量:" + TD015 + ",本次進貨驗收數量:" + AcceptQty + ",不可操作!!!");
                                    }
                                }
                            }
                            #endregion
                            #endregion

                            #region //確認段

                            #region //PURTG 進貨單單頭檔
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE PURTG SET
                                    MODIFIER =@MODIFIER,
                                    MODI_DATE= @MODI_DATE,
                                    MODI_TIME= @MODI_TIME,
                                    MODI_AP= @MODI_AP,
                                    MODI_PRID= @MODI_PRID,
                                    FLAG= FLAG + 1,
                                    TG003 = @TG003,
                                    TG013 = 'Y'
                                    WHERE TG001 = @GrErpPrefix
                                    AND TG002 = @GrErpNo";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                MODIFIER = UserNo,
                                MODI_DATE = dateNow,
                                MODI_TIME = timeNow,
                                MODI_AP = BaseHelper.ClientComputer(),
                                MODI_PRID = "BM",
                                GrErpPrefix,
                                GrErpNo,
                                TG003 = MaxReceiptDate
                            });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);


                            #endregion

                            if (item.TH030 != "Y")
                            {
                                #region //固定更新
                                #region //PURMA 廠商基本資料檔
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE PURMA SET
                                    MODIFIER = @MODIFIER,
                                    MODI_DATE = @MODI_DATE,
                                    MODI_TIME = @MODI_TIME,
                                    MODI_AP = @MODI_AP,
                                    MODI_PRID = @MODI_PRID,
                                    FLAG = CASE 
                                      WHEN FLAG >= 999 THEN 1
                                      ELSE FLAG + 1
                                    END,
                                    MA023 = @MA023
                                    WHERE MA001 = @SupplierNo";
                                dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MODIFIER = UserNo,
                                    MODI_DATE = dateNow,
                                    MODI_TIME = timeNow,
                                    MODI_AP = BaseHelper.ClientComputer(),
                                    MODI_PRID = "BM",
                                    MA023 = ReceiveDateErp,
                                    SupplierNo
                                });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region //PURMB 品號廠商單頭檔
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE PURMB SET
                                    MODIFIER = @MODIFIER,
                                    MODI_DATE = @MODI_DATE,
                                    MODI_TIME = @MODI_TIME,
                                    MODI_AP = @MODI_AP,
                                    MODI_PRID = @MODI_PRID,
                                    FLAG = FLAG + 1,
                                    MB009 = @MB009
                                    WHERE MB001 = @MB001
                                    AND MB002 = @MB002";
                                dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MODIFIER = UserNo,
                                    MODI_DATE = dateNow,
                                    MODI_TIME = timeNow,
                                    MODI_AP = BaseHelper.ClientComputer(),
                                    MODI_PRID = "BM",
                                    MB009 = ReceiveDateErp,
                                    MB001 = MtlItemNo,
                                    MB002 = SupplierNo
                                });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region //PURTD 採購單單身資料檔
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE PURTD SET
                                    MODIFIER =@MODIFIER,
                                    MODI_DATE= @MODI_DATE,
                                    MODI_TIME= @MODI_TIME,
                                    MODI_AP= @MODI_AP,
                                    MODI_PRID= @MODI_PRID,
                                    FLAG= FLAG + 1,
                                    TD015 = TD015 + @TD015,
                                    TD016 = @TD016,
                                    TD060 = TD060 + @TD060
                                    WHERE TD001 = @TD001
                                    AND TD002 = @TD002
                                    AND TD003 = @TD003";
                                dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MODIFIER = UserNo,
                                    MODI_DATE = dateNow,
                                    MODI_TIME = timeNow,
                                    MODI_AP = BaseHelper.ClientComputer(),
                                    MODI_PRID = "BM",
                                    TD015 = Math.Round(ConversionQty / PoRate, 3, MidpointRounding.AwayFromZero),
                                    TD016 = CloseStatus,
                                    TD060 = Math.Round(ConversionAvailableQty / PoAvailableRate, 3, MidpointRounding.AwayFromZero),
                                    TD001 = PoErpPrefix,
                                    TD002 = PoErpNo,
                                    TD003 = PoSeq
                                });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region //PURTH 進貨單單身檔
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE PURTH SET
                                    MODIFIER =@MODIFIER,
                                    MODI_DATE= @MODI_DATE,
                                    MODI_TIME= @MODI_TIME,
                                    MODI_AP= @MODI_AP,
                                    MODI_PRID= @MODI_PRID,
                                    FLAG= FLAG + 1,
                                    TH030 = 'Y',
                                    TH038 = @UserNo
                                    WHERE TH001 = @GrErpPrefix
                                    AND TH002 = @GrErpNo
                                    AND TH003 = @GrSeq";
                                dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MODIFIER = UserNo,
                                    MODI_DATE = dateNow,
                                    MODI_TIME = timeNow,
                                    MODI_AP = BaseHelper.ClientComputer(),
                                    MODI_PRID = "BM",
                                    GrErpPrefix,
                                    GrErpNo,
                                    GrSeq,
                                    UserNo
                                });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                                #endregion

                                #region //INVMB 品號基本資料檔 庫存數量 庫存金額
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE INVMB SET
                                    MODIFIER =@MODIFIER,
                                    MODI_DATE= @MODI_DATE,
                                    MODI_TIME= @MODI_TIME,
                                    MODI_AP= @MODI_AP,
                                    MODI_PRID= @MODI_PRID,
                                    FLAG= FLAG + 1,
                                    MB064  = MB064 + @MB064,
                                    MB065  = MB065 + @MB065
                                    WHERE MB001 = @MB001";
                                dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MODIFIER = UserNo,
                                    MODI_DATE = dateNow,
                                    MODI_TIME = timeNow,
                                    MODI_AP = BaseHelper.ClientComputer(),
                                    MODI_PRID = "BM",
                                    MB048 = CurrencyCode,
                                    MB049 = Math.Round(TH018 / GrRate, 4, MidpointRounding.AwayFromZero),
                                    MB050 = Math.Round(TH018 * Exchange / GrRate, 4, MidpointRounding.AwayFromZero),
                                    MB064 = ConversionQty,
                                    MB065 = TH047,
                                    MB001 = MtlItemNo
                                });
                                rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                                #endregion

                                #region //固定新增
                                #region //INVLA 異動明細資料檔
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO INVLA (COMPANY, CREATOR, USR_GROUP, 
                                    CREATE_DATE,MODIFIER, MODI_DATE, FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID,
                                    UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10,
                                    LA001, LA004, LA005, LA006, LA007,
                                    LA008, LA009, LA010, LA011, LA012,
                                    LA013, LA014, LA015, LA016, LA017,
                                    LA018, LA019, LA020, LA021, LA022,
                                    LA023, LA024, LA025, LA026, LA027,
                                    LA028, LA029, LA030, LA031, LA032,
                                    LA033, LA034)
                                    OUTPUT INSERTED.LA001
                                    VALUES (@COMPANY, @CREATOR, @USR_GROUP,
                                    @CREATE_DATE,@MODIFIER,@MODI_DATE,@FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID,
                                    '','','','','',0,0,0,0,0,
                                    @LA001, @LA004, @LA005, @LA006, @LA007,
                                    @LA008, @LA009, @LA010, @LA011, @LA012,
                                    @LA013, @LA014, @LA015, @LA016, @LA017,
                                    @LA018, @LA019, @LA020, @LA021, @LA022,
                                    0, 0, '', '', '',
                                    '', '','','','',
                                    '', '')";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        COMPANY = ErpNo,
                                        CREATOR = UserNo,
                                        USR_GROUP,
                                        CREATE_DATE = dateNow,
                                        MODIFIER = "",
                                        MODI_DATE = "",
                                        FLAG = 1,
                                        CREATE_TIME = timeNow,
                                        CREATE_AP = BaseHelper.ClientComputer(),
                                        CREATE_PRID = "BM",
                                        LA001 = MtlItemNo,
                                        LA004 = MaxReceiptDate,
                                        LA005 = 1,
                                        LA006 = GrErpPrefix,
                                        LA007 = GrErpNo,
                                        LA008 = GrSeq,
                                        LA009 = InventoryNo,
                                        LA010 = SupplierNo,
                                        LA011 = ConversionQty,
                                        LA012 = UnitCost,
                                        LA013 = TH047,
                                        LA014 = "1",
                                        LA015 = "Y",
                                        LA016 = LotNumber,
                                        LA017 = TH047, //金額-材料 暫未確認 
                                        LA018 = 0,
                                        LA019 = 0,
                                        LA020 = 0,
                                        LA021 = 0,
                                        LA022 = Location != "" ? Location : "" //儲位尚未給
                                    });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                                #endregion

                                #region //第一次新增後續更新

                                #region //INVMC 品號庫別檔異動
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT MC012 FROM INVMC
                                    WHERE MC001 = @MC001
                                      AND MC002 = @MC002";
                                dynamicParameters.Add("MC001", MtlItemNo);
                                dynamicParameters.Add("MC002", InventoryNo);
                                var resultINVMC = sqlConnection.Query(sql, dynamicParameters);

                                if (resultINVMC.Count() > 0)
                                {
                                    #region //後續更新流程
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE INVMC SET
                                        MODIFIER =@MODIFIER,
                                        MODI_DATE= @MODI_DATE,
                                        MODI_TIME= @MODI_TIME,
                                        MODI_AP= @MODI_AP,
                                        MODI_PRID= @MODI_PRID,
                                        FLAG= FLAG + 1,
                                        MC007  = MC007 + @MC007,
                                        MC008  = MC008 + @MC008,
                                        MC012  = @MC012 
                                        WHERE MC001 = @MC001
                                        AND MC002 = @MC002";
                                    dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        MODIFIER = UserNo,
                                        MODI_DATE = dateNow,
                                        MODI_TIME = timeNow,
                                        MODI_AP = BaseHelper.ClientComputer(),
                                        MODI_PRID = "BM",
                                        MC007 = ConversionQty,
                                        MC008 = TH047,
                                        MC012 = ReceiveDateErp,
                                        MC001 = MtlItemNo,
                                        MC002 = InventoryNo
                                    });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                                else
                                {
                                    #region //第一次新增流程
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO INVMC (COMPANY, CREATOR, USR_GROUP, 
                                        CREATE_DATE,MODIFIER, MODI_DATE, FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID,
                                        MC001, MC002, MC003, MC004, MC005, MC006, MC007, MC008, MC009, MC010, MC011,
                                        MC012, MC013, MC014, MC015)
                                        VALUES (@COMPANY, @CREATOR, @USR_GROUP,
                                        @CREATE_DATE,@MODIFIER,@MODI_DATE,@FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID,
                                        @MC001, @MC002, @MC003, @MC004, @MC005, @MC006, @MC007, @MC008, @MC009, @MC010, @MC011,
                                        @MC012, @MC013, @MC014, @MC015)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            COMPANY = ErpNo,
                                            CREATOR = UserNo,
                                            USR_GROUP,
                                            CREATE_DATE = dateNow,
                                            MODIFIER = "",
                                            MODI_DATE = "",
                                            FLAG = 1,
                                            CREATE_TIME = timeNow,
                                            CREATE_AP = BaseHelper.ClientComputer(),
                                            CREATE_PRID = "BM",
                                            MC001 = MtlItemNo,
                                            MC002 = InventoryNo,
                                            MC003 = "",
                                            MC004 = 0.000,
                                            MC005 = 0.000,
                                            MC006 = 0.000,
                                            MC007 = ConversionQty,
                                            MC008 = TH047,
                                            MC009 = 0.000,
                                            MC010 = 0.00000,
                                            MC011 = "",
                                            MC012 = ReceiveDateErp,
                                            MC013 = "",
                                            MC014 = 0.000,
                                            MC015 = ""
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                }
                                #endregion

                                #endregion

                                #region //品號有批號管理觸發
                                if (LotManagement != "N")
                                {
                                    #region //固定新增
                                    #region //INVLF 儲位批號異動明細資料檔
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO INVLF (COMPANY, CREATOR, USR_GROUP, 
                                        CREATE_DATE,MODIFIER, MODI_DATE, FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID,
                                        UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10,
                                        LF001, LF002, LF003, LF004, LF005,
                                        LF006, LF007, LF008, LF009, LF010,
                                        LF011, LF012, LF013)
                                VALUES (@COMPANY, @CREATOR, @USR_GROUP,
                                        @CREATE_DATE,@MODIFIER,@MODI_DATE,@FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID,
                                        '','','','','',0,0,0,0,0,
                                        @LF001, @LF002, @LF003, @LF004, @LF005,
                                        @LF006, @LF007, @LF008, @LF009, @LF010,
                                        @LF011, @LF012, @LF013)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            COMPANY = ErpNo,
                                            CREATOR = UserNo,
                                            USR_GROUP,
                                            CREATE_DATE = dateNow,
                                            MODIFIER = "",
                                            MODI_DATE = "",
                                            FLAG = 1,
                                            CREATE_TIME = timeNow,
                                            CREATE_AP = BaseHelper.ClientComputer(),
                                            CREATE_PRID = "BM",
                                            LF001 = GrErpPrefix,
                                            LF002 = GrErpNo,
                                            LF003 = GrSeq,
                                            LF004 = MtlItemNo,
                                            LF005 = InventoryNo,
                                            LF006 = Location.Length <= 0 ? "##########" : Location,
                                            LF007 = LotNumber.Length <= 0 ? "####################" : LotNumber,
                                            LF008 = 1,
                                            LF009 = ReceiveDateErp,
                                            LF010 = 1,
                                            LF011 = ConversionQty,
                                            LF012 = 0,
                                            LF013 = SupplierNo
                                        });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //INVMF 品號批號資料單身
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO INVMF (COMPANY, CREATOR, USR_GROUP, 
                                        CREATE_DATE,MODIFIER, MODI_DATE, FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID,
                                        UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10,
                                        MF001, MF002, MF003, MF004, MF005, 
                                        MF006, MF007, MF008, MF009, MF010, 
                                        MF013, MF014,
                                        MF016, MF017, MF502, MF503, MF504)
                                 VALUES(@COMPANY, @CREATOR, @USR_GROUP,
                                        @CREATE_DATE,@MODIFIER,@MODI_DATE,@FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID,
                                        '','','','','',0,0,0,0,0,
                                        @MF001, @MF002, @MF003, @MF004, @MF005,
                                        @MF006, @MF007, @MF008, @MF009, @MF010,
                                        @MF013, @MF014,
                                        0, 0, 0, 0, 0)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            COMPANY = ErpNo,
                                            CREATOR = UserNo,
                                            USR_GROUP,
                                            CREATE_DATE = dateNow,
                                            MODIFIER = "",
                                            MODI_DATE = "",
                                            FLAG = 1,
                                            CREATE_TIME = timeNow,
                                            CREATE_AP = BaseHelper.ClientComputer(),
                                            CREATE_PRID = "BM",
                                            MF001 = MtlItemNo,
                                            MF002 = LotNumber,
                                            MF003 = ReceiveDateErp,
                                            MF004 = GrErpPrefix,
                                            MF005 = GrErpNo,
                                            MF006 = GrSeq,
                                            MF007 = InventoryNo,
                                            MF008 = 1,
                                            MF009 = "1",
                                            MF010 = ConversionQty,
                                            MF013 = "", //備註
                                            MF014 = 0 //包裝數量
                                        });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #endregion

                                    #region //只新增一次
                                    #region //INVME 品號批號資料單頭
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT ME001,ME007 FROM INVME
                                        WHERE ME001 = @ME001
                                          AND ME002 = @ME002";
                                    dynamicParameters.Add("ME001", MtlItemNo);
                                    dynamicParameters.Add("ME002", LotNumber);
                                    var resultINVME = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultINVME.Count() <= 0)
                                    {
                                        #region //INVME 新增
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO INVME (COMPANY, CREATOR, USR_GROUP, 
                                       CREATE_DATE,MODIFIER, MODI_DATE, FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID,
                                       UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10,
                                       ME001, ME002, ME003, ME004, ME005, 
                                       ME006, ME007, ME008, ME009, ME010,
                                       ME011, ME012, ME503, ME504, ME505 )
                                VALUES(@COMPANY, @CREATOR, @USR_GROUP,
                                       @CREATE_DATE,@MODIFIER,@MODI_DATE,@FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID,
                                       '','','','','',0,0,0,0,0,
                                       @ME001, @ME002, @ME003, @ME004, @ME005, 
                                       @ME006, @ME007, @ME008, @ME009, @ME010,
                                       0, 0, 0, 0, 0 )";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                COMPANY = ErpNo,
                                                CREATOR = UserNo,
                                                USR_GROUP,
                                                CREATE_DATE = dateNow,
                                                MODIFIER = "",
                                                MODI_DATE = "",
                                                FLAG = 1,
                                                CREATE_TIME = timeNow,
                                                CREATE_AP = BaseHelper.ClientComputer(),
                                                CREATE_PRID = "BM",
                                                ME001 = MtlItemNo,
                                                ME002 = LotNumber,
                                                ME003 = ReceiveDateErp,
                                                ME004 = "",
                                                ME005 = GrErpPrefix,
                                                ME006 = GrErpNo,
                                                ME007 = "N",
                                                ME008 = "",
                                                ME009 = AvailableDate,
                                                ME010 = ReCheckDate
                                            });
                                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                        #endregion

                                    }
                                    else
                                    {
                                        foreach (var item1 in resultINVME)
                                        {
                                            if (item1.ME007 != "N") throw new SystemException("品號:" + MtlItemNo + "批號:" + LotNumber + "已經結案了!!!");
                                        }
                                    }
                                    #endregion

                                    #endregion

                                    #region //第一次新增後續更新
                                    #region //INVMM 品號庫別儲位批號檔
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT MM001 FROM INVMM
                                        WHERE MM001 = @MM001
                                          AND MM002 = @MM002
                                          AND MM003 = @MM003
                                          AND MM004 = @MM004";
                                    dynamicParameters.Add("MM001", MtlItemNo);
                                    dynamicParameters.Add("MM002", InventoryNo);
                                    dynamicParameters.Add("MM003", Location != "" ? Location : "##########");
                                    dynamicParameters.Add("MM004", LotNumber);
                                    var resultINVMM = sqlConnection.Query(sql, dynamicParameters);

                                    if (resultINVMM.Count() > 0)
                                    {
                                        #region //後續更新流程
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"UPDATE INVMM SET
                                                  MODIFIER = @MODIFIER,
                                                  MODI_DATE = @MODI_DATE,
                                                  MODI_TIME = @MODI_TIME,
                                                  MODI_AP = @MODI_AP,
                                                  MODI_PRID = @MODI_PRID,
                                                  FLAG = FLAG + 1,
                                                  MM005 = MM005 + @MM005,
                                                  MM008 = @MM008
                                            WHERE MM001 = @MM001
                                              AND MM002 = @MM002
                                              AND MM003 = @MM003
                                              AND MM004 = @MM004";
                                        dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            MODIFIER = UserNo,
                                            MODI_DATE = dateNow,
                                            MODI_TIME = timeNow,
                                            MODI_AP = BaseHelper.ClientComputer(),
                                            MODI_PRID = "BM",
                                            MM005 = ConversionQty,
                                            MM008 = ReceiveDateErp,
                                            MM001 = MtlItemNo,
                                            MM002 = InventoryNo,
                                            MM003 = Location != "" ? Location : "##########",
                                            MM004 = LotNumber,
                                        });
                                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                        #endregion
                                    }
                                    else
                                    {
                                        #region //第一次新增流程
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO INVMM (COMPANY, CREATOR, USR_GROUP, 
                                            CREATE_DATE,MODIFIER, MODI_DATE, FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID,
                                            UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10,
                                            MM001, MM002, MM003, MM004, MM005, 
                                            MM006, MM007, MM008, MM009, MM010, 
                                            MM011, MM012, MM013, MM014)
                                    VALUES (@COMPANY, @CREATOR, @USR_GROUP,
                                            @CREATE_DATE,@MODIFIER,@MODI_DATE,@FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID,
                                            '','','','','',0,0,0,0,0,
                                            @MM001, @MM002, @MM003, @MM004, @MM005, 
                                            @MM006, @MM007, @MM008, @MM009, @MM010, 
                                            @MM011, @MM012, @MM013, @MM014)";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                COMPANY = ErpNo,
                                                CREATOR = UserNo,
                                                USR_GROUP,
                                                CREATE_DATE = dateNow,
                                                MODIFIER = "",
                                                MODI_DATE = "",
                                                FLAG = 1,
                                                CREATE_TIME = timeNow,
                                                CREATE_AP = BaseHelper.ClientComputer(),
                                                CREATE_PRID = "BM",
                                                MM001 = MtlItemNo,
                                                MM002 = InventoryNo,
                                                MM003 = Location.Length > 0 ? Location : "##########",
                                                MM004 = LotNumber.Length > 0 ? LotNumber : "####################",
                                                MM005 = ConversionQty,
                                                MM006 = 0, //庫存包裝數量,
                                                MM007 = "",
                                                MM008 = ReceiveDateErp,
                                                MM009 = "",
                                                MM010 = "", //備註
                                                MM011 = 0,
                                                MM012 = "",
                                                MM013 = 0,
                                                MM014 = 0
                                            });

                                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                        #endregion
                                    }
                                    #endregion

                                    #endregion
                                }
                                #endregion

                                #region //進貨單來源為暫入單觸發
                                if (TrErpPrefix != "")
                                {
                                    //#region //INVTF 暫出入轉撥單頭檔更新 轉銷日更新
                                    //dynamicParameters = new DynamicParameters();
                                    //sql = @"UPDATE INVTF 
                                    //        SET MODIFIER =@MODIFIER,
                                    //       MODI_DATE= @MODI_DATE,
                                    //       MODI_TIME= @MODI_TIME,
                                    //       MODI_AP= @MODI_AP,
                                    //       MODI_PRID= @MODI_PRID,
                                    //       FLAG= FLAG + 1,
                                    //        TF042  = @TF042
                                    //        WHERE TF001 = @TF001
                                    //        AND TF002 = @TF002";
                                    //dynamicParameters.AddDynamicParams(
                                    //new
                                    //{
                                    //    TF042 = ReceiveDateErp,
                                    //    TF001 = TH032,
                                    //    TF002 = TH033
                                    //});
                                    //rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    //#endregion

                                    //#region //INVTG 暫出入轉撥單身檔 轉銷數量
                                    //dynamicParameters = new DynamicParameters();
                                    //sql = @"UPDATE INVTG SET
                                    //        TG020  = TG020 + @TG020,
                                    //        TG046  = TG046 + @TG046,
                                    //        TG054  = TG054 + @TG054,
                                    //        TG024  = @TG024
                                    //        WHERE TG001 = @TG001
                                    //        AND TG002 = @TG002
                                    //        AND TG003 = @TG003";
                                    //dynamicParameters.AddDynamicParams(
                                    //new
                                    //{
                                    //    TG020 = Quantity,
                                    //    TG046 = FreeSpareQty,
                                    //    TG054 = AvailableQty,
                                    //    TG001 = TH032,
                                    //    TG002 = TH033,
                                    //    TG003 = TH034,
                                    //    TG024 = ClosureStatusINVTG
                                    //});
                                    //rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    //#endregion
                                }
                                #endregion

                            }
                        

                            #endregion

                        }
                        #endregion
                    }
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //若成功則將相關資料寫回MES
                        DateTime date = DateTime.ParseExact(MaxReceiptDate, "yyyyMMdd", null);
                        MaxReceiptDate = date.ToString("yyyy-MM-dd"); // yyyy-MM-dd 格式

                        #region //更新 - 進貨單單頭
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.GoodsReceipt SET
                                ReceiptDate = @ReceiptDate,
                                ConfirmStatus = @ConfirmStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE GrId = @GrId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ReceiptDate = MaxReceiptDate,
                                ConfirmStatus = "Y",
                                LastModifiedDate,
                                LastModifiedBy,
                                GrId
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //更新 - 進貨單單單身
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.GrDetail SET
                                ConfirmStatus = @ConfirmStatus,
                                ConfirmUser = @ConfirmUser,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE GrId = @GrId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmStatus = "Y",
                                ConfirmUser = LastModifiedBy,
                                LastModifiedDate,
                                LastModifiedBy,
                                GrId
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region /撈取進貨單單身資料=>更新採購單,暫入單
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.AcceptQty,a.AvailableQty,a.PoDetailId, a.ConfirmStatus
                                FROM SCM.GrDetail a
                                WHERE a.GrId = @GrId";
                        dynamicParameters.Add("GrId", GrId);

                        var resultGrDetail = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in resultGrDetail)
                        {
                            if (item.ConfirmStatus != "Y")
                            {
                                #region //更新 - 採購單 已交數量,已交計價數量
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE SCM.PoDetail SET
                                    SiQty = SiQty + @SiQty,
                                    SiPriceQty = SiPriceQty + @SiPriceQty,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE PoDetailId = @PoDetailId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        SiQty = item.AcceptQty,
                                        SiPriceQty = item.AvailableQty,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        PoDetailId = item.PoDetailId
                                    });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region //更新 - 暫入單 轉進銷量,轉銷贈/備品量
                                //if (item.TsnDetailId > 0)
                                //{
                                //    dynamicParameters = new DynamicParameters();
                                //    sql = @"UPDATE SCM.TsnDetail SET
                                //            SaleQty = SaleQty + @SaleQty,
                                //            SaleFSQty = SaleFSQty + @SaleFSQty,
                                //            LastModifiedDate = @LastModifiedDate,
                                //            LastModifiedBy = @LastModifiedBy
                                //            WHERE TsnDetailId = @TsnDetailId";
                                //    dynamicParameters.AddDynamicParams(
                                //        new
                                //        {
                                //            SaleQty = item.Quantity,
                                //            SaleFSQty = item.FreeSpareQty,
                                //            LastModifiedDate,
                                //            LastModifiedBy,
                                //            TsnDetailId = item.TsnDetailId
                                //        });
                                //    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                //}
                                #endregion
                            }
                        
                        }
                        #endregion
                        #endregion

                        #region //批號相關異動
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.GrSeq,a.LotNumber,a.MtlItemId,b.LotManagement,a.InventoryId,a.AcceptQty, a.ConfirmStatus
                                FROM SCM.GrDetail a
                                INNER JOIN PDM.MtlItem b on a.MtlItemId = b.MtlItemId
                                WHERE a.GrId = @GrId";
                        dynamicParameters.Add("GrId", GrId);

                        var resultRoMes = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in resultRoMes)
                        {
                            string GrSeq = item.GrSeq;
                            int MtlItemId = item.MtlItemId;
                            string LotNumber = item.LotNumber;
                            string LotManagement = item.LotManagement;
                            int InventoryId = item.InventoryId;
                            double AcceptQty = Convert.ToDouble(item.AcceptQty);
                            if (LotManagement != "N")
                            {
                                if (item.ConfirmStatus != "Y")
                                {
                                    #region //確認品號資料是否正確
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.MtlItemNo, a.LotManagement
                                    FROM PDM.MtlItem a 
                                    WHERE a.MtlItemId = @MtlItemId";
                                    dynamicParameters.Add("MtlItemId", MtlItemId);

                                    var MtlItemResult = sqlConnection.Query(sql, dynamicParameters);

                                    if (MtlItemResult.Count() <= 0) throw new SystemException("品號資料錯誤!!");

                                    string MtlItemNo = "";
                                    foreach (var item1 in MtlItemResult)
                                    {
                                        if (item.LotManagement != "T") throw new SystemException("品號批號管控參數錯誤!!");
                                        MtlItemNo = item.MtlItemNo;
                                    }
                                    #endregion

                                    #region //確認此品號是否已存在此批號
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 LotNumberId
                                    FROM SCM.LotNumber a 
                                    WHERE a.MtlItemId = @MtlItemId
                                    AND a.LotNumberNo = @LotNumberNo";
                                    dynamicParameters.Add("MtlItemId", MtlItemId);
                                    dynamicParameters.Add("LotNumberNo", LotNumber);

                                    var LotNumberResult = sqlConnection.Query(sql, dynamicParameters);

                                    if (LotNumberResult.Count() <= 0)
                                    {
                                        #region //INSERT SCM.GoodsReceipt
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO SCM.LotNumber (MtlItemId, LotNumberNo, Remark
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.LotNumberId
                                    VALUES (@MtlItemId, @LotNumberNo, @Remark
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                MtlItemId,
                                                LotNumberNo = LotNumber,
                                                Remark = "",
                                                CreateDate,
                                                LastModifiedDate,
                                                CreateBy,
                                                LastModifiedBy
                                            });
                                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                        rowsAffected += insertResult.Count();
                                        #endregion

                                        int LotNumberId = -1;
                                        foreach (var item1 in insertResult)
                                        {
                                            LotNumberId = item1.LotNumberId;
                                        }

                                        #region //INSERT SCM.LnDetail
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO SCM.LnDetail (LotNumberId, TransactionDate, FromErpPrefix,FromErpNo,FromSeq
                                    ,InventoryId,TransactionType,DocType,Quantity,Remark
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@LotNumberId, @TransactionDate, @FromErpPrefix,@FromErpNo,@FromSeq
                                    ,@InventoryId,@TransactionType,@DocType,@Quantity,@Remark
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                LotNumberId,
                                                TransactionDate = ReceiptDate,
                                                FromErpPrefix = GrErpPrefix,
                                                FromErpNo = GrErpNo,
                                                FromSeq = GrSeq,
                                                InventoryId,
                                                TransactionType = 1,
                                                DocType = 1,
                                                Quantity = AcceptQty,
                                                Remark = "",
                                                CreateDate,
                                                LastModifiedDate,
                                                CreateBy,
                                                LastModifiedBy
                                            });
                                        insertResult = sqlConnection.Query(sql, dynamicParameters);

                                        rowsAffected += insertResult.Count();
                                        #endregion
                                    }
                                    else
                                    {
                                        int LotNumberId = -1;
                                        foreach (var item1 in LotNumberResult)
                                        {
                                            LotNumberId = item1.LotNumberId;
                                        }

                                        #region //INSERT SCM.LnDetail
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO SCM.LnDetail (LotNumberId, TransactionDate, FromErpPrefix,FromErpNo,FromSeq
                                    ,InventoryId,TransactionType,DocType,Quantity,Remark
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@LotNumberId, @TransactionDate, @FromErpPrefix,@FromErpNo,@FromSeq
                                    ,@InventoryId,@TransactionType,@DocType,@Quantity,@Remark
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                LotNumberId,
                                                TransactionDate = ReceiptDate,
                                                FromErpPrefix = GrErpPrefix,
                                                FromErpNo = GrErpNo,
                                                FromSeq = GrSeq,
                                                InventoryId,
                                                TransactionType = 1,
                                                DocType = 1,
                                                Quantity = AcceptQty,
                                                Remark = "",
                                                CreateDate,
                                                LastModifiedDate,
                                                CreateBy,
                                                LastModifiedBy
                                            });
                                        var insertResult1 = sqlConnection.Query(sql, dynamicParameters);

                                        rowsAffected += insertResult1.Count();
                                        #endregion
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

        #region //UpdateGrReConfirm -- 反確認進貨單據 -- Shintokuro 2024-04-29
        public string UpdateGrReConfirm(int GrId)
        {
            try
            {
                string dateNow = DateTime.Now.ToString("yyyyMMdd");
                string timeNow = DateTime.Now.ToString("HH:mm:ss");
                string dateNowMonth = DateTime.Now.ToString("yyyyMM");

                string UserNo = "";

                int rowsAffected = 0;
                var ErpNo = "";
                string ErpDbName = "";
                string USR_GROUP = "";
                int UserId = CreateBy;

                string GrErpPrefix = "";
                string GrErpNo = "";
                string ReceiptDate = "";
                int CompanyIdBase = -1;
                int GrDetailId = -1;

                using (TransactionScope transactionScope = new TransactionScope())
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

                        var resultCompany = sqlConnection.Query(sql, dynamicParameters);

                        if (resultCompany.Count() <= 0) throw new SystemException("公司別錯誤!");

                        foreach (var item in resultCompany)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            ErpDbName = ErpConnectionStrings.Split('=')[2].Split(';')[0];
                            ErpNo = item.ErpNo;
                        }
                        #endregion

                        #region //撈取使用者No
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
                        dynamicParameters.Add("FunctionCode", "GoodsReceiptManagement");
                        dynamicParameters.Add("DetailCode", "confirm");

                        var resultUser = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUser.Count() <= 0) throw new SystemException("使用者資料錯誤!");

                        foreach (var item in resultUser)
                        {
                            UserNo = item.UserNo;
                        }
                        #endregion

                        #region /撈取單頭資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.GrErpPrefix,a.GrErpNo, a.ConfirmStatus,a.CompanyId
                                ,FORMAT(a.ReceiptDate, 'yyyyMMdd') ReceiptDate
                                ,b.GrDetailId,b.CloseStatus
                                FROM SCM.GoodsReceipt a
                                INNER JOIN SCM.GrDetail b on a.GrId = b.GrId
                                WHERE a.GrId = @GrId";
                        dynamicParameters.Add("GrId", GrId);

                        var resultRoMes = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in resultRoMes)
                        {
                            if (item.ConfirmStatus == "V") throw new SystemException("該進貨單單據已經作廢,不能操作");
                            if (item.CloseStatus == "Y") throw new SystemException("該進貨單已經結帳,不能操作");
                            if (item.ConfirmStatus != "Y") throw new SystemException("該進貨單單據須處於核單狀態,才能操作");
                            if (item.CompanyId != CurrentCompany) throw new SystemException("該進貨單單據公司別與後端公司別不一致,不能操作");
                            CompanyIdBase = item.CompanyId;
                            GrErpPrefix = item.GrErpPrefix;
                            GrErpNo = item.GrErpNo;
                            ReceiptDate = item.ReceiptDate;
                            GrDetailId = item.GrDetailId;
                        }
                        #endregion
                    }
                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //審核ERP權限
                        USR_GROUP = BaseHelper.CheckErpAuthority(UserNo, sqlConnection, "PURI09", "CONFIRM");
                        #endregion

                        #region //判斷開立日期是否超過結帳日
                        string CloseDateBase = "";
                        string ReceiveDateBase = ReceiptDate;
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT MA011,MA012,MA013
                                FROM CMSMA ";
                        var resultCloseDate = sqlConnection.Query(sql, dynamicParameters);
                        foreach (var item in resultCloseDate)
                        {
                            CloseDateBase = item.MA013;
                        }

                        DateTime docDateBase;
                        DateTime closeDateBase;

                        if (DateTime.TryParseExact(ReceiveDateBase, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out docDateBase) &&
                            DateTime.TryParseExact(CloseDateBase, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out closeDateBase))
                        {
                            int resultDate = DateTime.Compare(docDateBase, closeDateBase);

                            if (resultDate <= 0)
                            {
                                throw new SystemException("【進貨單】ERP已經結帳,不可以在異動【" + CloseDateBase + "】之前的單據");
                            }
                        }
                        else
                        {
                            throw new SystemException("日期字符串无效，无法比较");
                        }
                        #endregion

                        #region //取出進貨單單據資料 + 異動更新相關資料表
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.TG003,a.TG005,a.TG007,a.TG013
                                ,b.TH003,b.TH004,b.TH007,b.TH009,b.TH010
                                ,b.TH011,b.TH012,b.TH013
                                ,b.TH014,b.TH015,b.TH016,b.TH017,b.TH030,b.TH031,b.TH063
                                ,b.TH021,b.TH022,b.TH023
                                ,b.TH036,b.TH037,b.TH047
                                ,c.MB022
                                ,d.MQ019
                                ,ISNULL(md1.GrRate,1) GrRate
                                ,ISNULL(md2.GrAvailableRate,1) GrAvailableRate
                                ,(b.TH015*ISNULL(md1.GrRate,1)) ConversionQty
                                ,(b.TH016*ISNULL(md2.GrAvailableRate,1)) ConversionAvailableQty
                                ,CASE 
                                    WHEN b.TH016 != 0 THEN ROUND((b.TH047 / (b.TH016*ISNULL(md2.GrAvailableRate,1))),CAST(md3.MF005 AS INT)) 
                                    ELSE 0
                                END AS UnitCost
                                FROM PURTG a
                                INNER JOIN PURTH b on a.TG001 = b.TH001 AND a.TG002 = b.TH002
                                INNER JOIN INVMB c on b.TH004 = c.MB001
                                INNER JOIN CMSMQ d ON d.MQ001=a.TG001 AND d.MQ003='34'
                                INNER JOIN CMSMC e ON e.MC001=b.TH009
                                LEFT JOIN INVMD f ON f.MD001=b.TH004
                                OUTER APPLY(
                                    SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) GrRate
                                    FROM INVMD x1
                                    WHERE x1.MD001=b.TH004
                                    AND  x1.MD002 = b.TH008
                                ) md1
                                OUTER APPLY(
                                    SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) GrAvailableRate
                                    FROM INVMD x1
                                    WHERE x1.MD001=b.TH004
                                    AND  x1.MD002 = b.TH056
                                ) md2
                                OUTER APPLY(
                                    SELECT x1.MF003,x1.MF004,x1.MF005,x1.MF006
                                    FROM CMSMF x1
                                    INNER JOIN CMSMA x2 on x1.MF001 = x2.MA003 
                                ) md3
                                WHERE a.TG001 = @GrErpPrefix
                                AND a.TG002 = @GrErpNo";
                        dynamicParameters.Add("GrErpPrefix", GrErpPrefix);
                        dynamicParameters.Add("GrErpNo", GrErpNo);
                        var resultRoErp = sqlConnection.Query(sql, dynamicParameters);
                        if (resultRoErp.Count() <= 0) throw new SystemException("找不到ERP進貨單單據!");
                        foreach (var item in resultRoErp)
                        {
                            if (item.TG013 == "V") throw new SystemException("該進貨單單據已經作廢,不能操作");
                            if (item.TG013 != "Y") throw new SystemException("該進貨單單據須處於核單,才能操作");
                            if (item.TH030 != "Y") throw new SystemException("該進貨單單據單身須處於核單,才能操作");
                            if (item.TH021 != "") throw new SystemException("來源暫入單模式尚未開發,不能操作");
                            if (item.MQ019 != "Y") throw new SystemException("不核對採購單模式尚未開發,不能操作");

                            #region //欄位參數撈取
                            string ReceiveDateErp = item.TG003; //TG003 進貨日期
                            string SupplierNo = item.TG005; //TG005 供應商
                            string CurrencyCode = item.TG007; //TG007 幣別
                            Double Exchange = Convert.ToDouble(item.TG008); //TG008 匯率
                            string ConfirmStatus = item.TG013; //TG013 單頭確認

                            string GrSeq = item.TH003; //TH003 進貨單序號
                            string GrFull = GrErpPrefix + '-' + GrErpNo + '-' + GrSeq;
                            string MtlItemNo = item.TH004; //TH004 品號
                            string InventoryNo = item.TH009; //TH009 庫別
                            string LotNumber = item.TH010; //TH010 批號
                            string PoErpPrefix = item.TH011; //TH011 採購單別 單號 序號
                            string PoErpNo = item.TH012; //TH012 
                            string PoSeq = item.TH013; //TH013 
                            string PoFull = item.TH011 + '-' + item.TH012 + '-' + item.TH013;
                            string TH014 = item.TH014; //TH014 驗收日期
                            Double ReceiptQty = Convert.ToDouble(item.TH007); //TH014 進貨數量
                            Double AcceptQty = Convert.ToDouble(item.TH015); //TH014 驗收數量
                            Double AvailableQty = Convert.ToDouble(item.TH016); //TH014 計價數量 
                            Double ReturnQty = Convert.ToDouble(item.TH017); //TH014 驗退數量
                            string TH030 = item.TH030; //TH030 確認碼
                            string CloseStatus = item.TH031; //TH031 結帳碼
                            string AvailableDate = item.TH036; //TH036 有效日期
                            string ReCheckDate = item.TH037; //TH037 複檢日期
                            string Location = item.TH063; //TH060 儲位
                            string TrErpPrefix = item.TH021; //TH021 暫入單別 / TH022 暫入單號 / TH023 暫入序號
                            string TrErpNo = item.TH022;
                            string TrSeq = item.TH023;
                            string TsnFull = item.TH021 + '-' + item.TH022 + '-' + item.TH023;
                            string LotManagement = item.MB022; //品號批號管理

                            Double GrRate = Convert.ToDouble(item.GrRate); //單位換算率
                            Double ConversionQty = Convert.ToDouble(item.ConversionQty); //單位換算後數量
                            Double ConversionAvailableQty = Convert.ToDouble(item.ConversionAvailableQty); //單位換算後計價數量
                            Double UnitCost = Convert.ToDouble(item.UnitCost); //單位成本
                            Double TH047 = Convert.ToDouble(item.TH047); //本幣未稅金額
                            Double TH018 = Convert.ToDouble(item.TH018); //原幣單價
                            Double PoRate = 0; //採購單位換算率
                            Double PoAvailableRate = 0; //採購計價單位換算率
                            Double PoConversionQty = 0; //採購單位換算後數量
                            Double PoConversionAvailableQty = 0; //採購單位換算後計價數量
                            #endregion

                            #region //檢查段

                            #region //數量查驗
                            if (ReceiptQty < 0) throw new SystemException("【進貨單:" + GrFull + "】進貨數量不可小於0!!!");
                            if (AcceptQty < 0) throw new SystemException("【進貨單:" + GrFull + "】驗收數量不可小於0!!!");
                            if (AvailableQty < 0) throw new SystemException("【進貨單:" + GrFull + "】計價數量不可小於0!!!");
                            if (ReceiptQty < AcceptQty) throw new SystemException("【進貨單:" + GrFull + "】驗收數量不可大於進貨數量!!!");
                            if (AcceptQty < AvailableQty) throw new SystemException("【進貨單:" + GrFull + "】計價數量不可大於驗收數量!!!");


                            #region //判斷該品號批號庫存是否足夠
                            if (LotManagement != "N")
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT MM005
                                    FROM INVMM a
                                    WHERE a.MM001 = @MM001
                                    AND a.MM002 = @MM002
                                    AND a.MM003 = @MM003
                                    AND a.MM004 = @MM004
                                    ";
                                dynamicParameters.Add("MM001", MtlItemNo);
                                dynamicParameters.Add("MM002", InventoryNo);
                                dynamicParameters.Add("MM003", Location != "" ? Location : "##########");
                                dynamicParameters.Add("MM004", LotNumber);
                                var INVMMResult = sqlConnection.Query(sql, dynamicParameters);
                                if (INVMMResult.Count() > 0)
                                {
                                    foreach (var item1 in INVMMResult)
                                    {
                                        if (Convert.ToInt32(item1.MM005) < ConversionQty) throw new SystemException("批號庫存不足不能執行反確認,請重新確認");
                                    }
                                }
                                else
                                {
                                    throw new SystemException("批號庫存不足不能執行反確認,請重新確認");
                                }
                            }
                            #endregion

                            #region //判斷該品號庫存是否足夠
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT MC007
                                    FROM INVMC a
                                    WHERE a.MC001 = @MC001
                                    AND a.MC002 = @MC002
                                    ";
                            dynamicParameters.Add("MC001", MtlItemNo);
                            dynamicParameters.Add("MC002", InventoryNo);
                            var INVMCResult = sqlConnection.Query(sql, dynamicParameters);
                            if (INVMCResult.Count() > 0)
                            {
                                foreach (var item1 in INVMCResult)
                                {
                                    if (Convert.ToInt32(item1.MC007) < ConversionQty) throw new SystemException("庫存不足不能執行反確認,請重新確認");
                                }
                            }
                            else
                            {
                                throw new SystemException("庫存不足不能執行反確認,請重新確認");
                            }
                            #endregion
                            

                            #endregion

                            #region //PURTD 採購單
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 ISNULL(a.TD008,0) TD008,ISNULL(a.TD015,0) TD015,a.TD016,a.TD018,
                                    ISNULL(md1.PoRate,1) PoRate, 
                                    ISNULL(md2.PoAvailableRate,1) PoAvailableRate, 
                                    (a.TD008*ISNULL(md1.PoRate,1)) PoConversionQty,
                                    (a.TD060*ISNULL(md2.PoAvailableRate,1)) PoConversionAvailableQty
                                    FROM PURTD a
                                    OUTER APPLY(
                                        SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) PoRate
                                        FROM INVMD x1
                                        WHERE x1.MD001 = a.TD004
                                        AND  x1.MD002 = a.TD009
                                    ) md1
                                    OUTER APPLY(
                                        SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) PoAvailableRate
                                        FROM INVMD x1
                                        WHERE x1.MD001 = a.TD004
                                        AND  x1.MD002 = a.TD059
                                    ) md2
                                    WHERE a.TD001 = @TD001
                                    AND a.TD002 = @TD002
                                    AND a.TD003 = @TD003";
                            dynamicParameters.Add("TD001", PoErpPrefix);
                            dynamicParameters.Add("TD002", PoErpNo);
                            dynamicParameters.Add("TD003", PoSeq);
                            var resultCOPTD = sqlConnection.Query(sql, dynamicParameters);
                            if (resultCOPTD.Count() <= 0) throw new SystemException("採購單找不到,請重新確認");
                            foreach (var item1 in resultCOPTD)
                            {
                                if (item1.TD018 != "Y") throw new SystemException("【採購單:" + PoFull + "】非確認狀態,不可操作!!!");
                                if (Convert.ToDouble(item1.TD008) == Convert.ToDouble(item1.TD015) - AcceptQty)
                                {
                                    CloseStatus = "Y";
                                }
                                else if (Convert.ToDouble(item1.TD008) < Convert.ToDouble(item1.TD015) - AcceptQty)
                                {
                                    throw new SystemException("【採購單:" + PoFull + "】已超過採購數量,目前已交數量:" + Convert.ToDouble(item1.TD015) + ",本次反確認進貨驗收數量:" + AcceptQty + ",不可操作!!!");
                                }
                                PoRate = Convert.ToDouble(item1.PoRate);
                                PoAvailableRate = Convert.ToDouble(item1.PoAvailableRate);
                                PoConversionQty = Convert.ToDouble(item1.PoConversionQty);
                                PoConversionAvailableQty = Convert.ToDouble(item1.PoConversionAvailableQty);
                            }
                            #endregion
                            #endregion

                            #region //反確認段
                            #region //固定更新

                            #region //PURTD 採購單單身資料檔
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE PURTD SET
                                    MODIFIER =@MODIFIER,
                                    MODI_DATE= @MODI_DATE,
                                    MODI_TIME= @MODI_TIME,
                                    MODI_AP= @MODI_AP,
                                    MODI_PRID= @MODI_PRID,
                                    FLAG= FLAG + 1,
                                    TD015 = TD015 - @TD015,
                                    TD016 = @TD016,
                                    TD060 = TD060 - @TD060
                                    WHERE TD001 = @TD001
                                    AND TD002 = @TD002
                                    AND TD003 = @TD003";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                MODIFIER = UserNo,
                                MODI_DATE = dateNow,
                                MODI_TIME = timeNow,
                                MODI_AP = BaseHelper.ClientComputer(),
                                MODI_PRID = "BM",
                                TD015 = Math.Round(ConversionQty / PoRate, 3, MidpointRounding.AwayFromZero),
                                TD016 = CloseStatus,
                                TD060 = Math.Round(ConversionAvailableQty / PoAvailableRate, 3, MidpointRounding.AwayFromZero),
                                TD001 = PoErpPrefix,
                                TD002 = PoErpNo,
                                TD003 = PoSeq
                            });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //PURTG 進貨單單頭檔
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE PURTG SET
                                    MODIFIER =@MODIFIER,
                                    MODI_DATE= @MODI_DATE,
                                    MODI_TIME= @MODI_TIME,
                                    MODI_AP= @MODI_AP,
                                    MODI_PRID= @MODI_PRID,
                                    FLAG= FLAG + 1,
                                    TG013 = 'N'
                                    WHERE TG001 = @GrErpPrefix
                                    AND TG002 = @GrErpNo";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                MODIFIER = UserNo,
                                MODI_DATE = dateNow,
                                MODI_TIME = timeNow,
                                MODI_AP = BaseHelper.ClientComputer(),
                                MODI_PRID = "BM",
                                GrErpPrefix,
                                GrErpNo
                            });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);


                            #endregion

                            #region //PURTH 進貨單單身檔
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE PURTH SET
                                    MODIFIER =@MODIFIER,
                                    MODI_DATE= @MODI_DATE,
                                    MODI_TIME= @MODI_TIME,
                                    MODI_AP= @MODI_AP,
                                    MODI_PRID= @MODI_PRID,
                                    FLAG= FLAG + 1,
                                    TH030 = 'N',
                                    TH031 = 'N',
                                    TH038 = @UserNo
                                    WHERE TH001 = @GrErpPrefix
                                    AND TH002 = @GrErpNo
                                    AND TH003 = @GrSeq";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                MODIFIER = UserNo,
                                MODI_DATE = dateNow,
                                MODI_TIME = timeNow,
                                MODI_AP = BaseHelper.ClientComputer(),
                                MODI_PRID = "BM",
                                GrErpPrefix,
                                GrErpNo,
                                GrSeq,
                                UserNo
                            });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                            #endregion

                            #region //INVMB 品號基本資料檔 庫存數量 庫存金額
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE INVMB SET
                                    MODIFIER =@MODIFIER,
                                    MODI_DATE= @MODI_DATE,
                                    MODI_TIME= @MODI_TIME,
                                    MODI_AP= @MODI_AP,
                                    MODI_PRID= @MODI_PRID,
                                    FLAG= FLAG + 1,
                                    MB064  = MB064 - @MB064,
                                    MB065  = MB065 - @MB065
                                    WHERE MB001 = @MB001";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                MODIFIER = UserNo,
                                MODI_DATE = dateNow,
                                MODI_TIME = timeNow,
                                MODI_AP = BaseHelper.ClientComputer(),
                                MODI_PRID = "BM",
                                MB064 = ConversionQty,
                                MB065 = TH047,
                                MB001 = MtlItemNo
                            });
                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //INVMC 品號庫別檔
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE INVMC SET
                                    MODIFIER =@MODIFIER,
                                    MODI_DATE= @MODI_DATE,
                                    MODI_TIME= @MODI_TIME,
                                    MODI_AP= @MODI_AP,
                                    MODI_PRID= @MODI_PRID,
                                    FLAG= FLAG + 1,
                                    MC007  = MC007 - @MC007,
                                    MC008  = MC008 - @MC008
                                    WHERE MC001 = @MC001
                                    AND MC002 = @MC002";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                MODIFIER = UserNo,
                                MODI_DATE = dateNow,
                                MODI_TIME = timeNow,
                                MODI_AP = BaseHelper.ClientComputer(),
                                MODI_PRID = "BM",
                                MC007 = ConversionQty,
                                MC008 = TH047,
                                MC013 = ReceiveDateErp,
                                MC001 = MtlItemNo,
                                MC002 = InventoryNo
                            });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #endregion

                            #region //固定刪除
                            #region //INVLA 異動明細資料檔
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE INVLA
                                    WHERE LA001 = @LA001
                                      AND LA004 = @LA004
                                      AND LA005 = @LA005
                                      AND LA006 = @LA006
                                       AND LA008 = @LA008";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                LA001 = MtlItemNo,
                                LA004 = ReceiveDateErp,
                                LA005 = 1,
                                LA006 = GrErpPrefix,
                                LA007 = GrErpNo,
                                LA008 = GrSeq
                            });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                            #endregion

                            #region //品號有批號管理觸發
                            if (LotManagement != "N")
                            {
                                #region //固定刪除

                                #region //INVMF 品號批號資料單身
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE INVMF
                                         WHERE MF001 = @MF001
                                           AND MF002 = @MF002
                                           AND MF003 = @MF003
                                           AND MF004 = @MF004
                                           AND MF005 = @MF005
                                           AND MF006 = @MF006";
                                dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MF001 = MtlItemNo,
                                    MF002 = LotNumber,
                                    MF003 = ReceiveDateErp,
                                    MF004 = GrErpPrefix,
                                    MF005 = GrErpNo,
                                    MF006 = GrSeq
                                });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region //INVLF 儲位批號異動明細資料檔
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE INVLF
                                        WHERE LF001 = @LF001
                                          AND LF002 = @LF002
                                          AND LF003 = @LF003
                                          AND LF004 = @LF004
                                          AND LF005 = @LF005
                                          AND LF006 = @LF006
                                          AND LF007 = @LF007
                                          AND LF008 = @LF008";
                                dynamicParameters.AddDynamicParams(
                                new
                                {
                                    LF001 = GrErpPrefix,
                                    LF002 = GrErpNo,
                                    LF003 = GrSeq,
                                    LF004 = MtlItemNo,
                                    LF005 = InventoryNo,
                                    LF006 = Location != "" ? Location : "##########",
                                    LF007 = LotNumber,
                                    LF008 = 1
                                });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #endregion

                                #region //固定更新
                                #region //INVMM 品號庫別儲位批號檔
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE INVMM SET
                                             MM005 = MM005 - @MM005
                                       WHERE MM001 = @MM001
                                         AND MM002 = @MM002
                                         AND MM003 = @MM003
                                         AND MM004 = @MM004";
                                dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MM005 = ConversionQty,
                                    MM001 = MtlItemNo,
                                    MM002 = InventoryNo,
                                    MM003 = Location != "" ? Location : "##########",
                                    MM004 = LotNumber
                                });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                                #endregion

                                #region //無數量就刪除
                                #region //INVME 品號批號資料單頭
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1 FROM INVMF
                                        WHERE MF001 = @MF001
                                          AND MF002 = @MF002";
                                dynamicParameters.Add("MF001", MtlItemNo);
                                dynamicParameters.Add("MF002", LotNumber);
                                var resultINVME = sqlConnection.Query(sql, dynamicParameters);
                                if (resultINVME.Count() <= 0)
                                {
                                    #region //INVME 刪除
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE INVME
                                         WHERE ME001 = @ME001
                                           AND ME002 = @ME002";
                                    dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        ME001 = MtlItemNo,
                                        ME002 = LotNumber
                                    });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                                #endregion

                                #endregion
                            }
                            #endregion

                            #region //進貨單來源為暫入單觸發
                            if (TrErpPrefix != "")
                            {
                                //#region //INVTF 暫出入轉撥單頭檔更新 轉銷日更新
                                //dynamicParameters = new DynamicParameters();
                                //sql = @"UPDATE INVTF 
                                //        SET MODIFIER =@MODIFIER,
                                //       MODI_DATE= @MODI_DATE,
                                //       MODI_TIME= @MODI_TIME,
                                //       MODI_AP= @MODI_AP,
                                //       MODI_PRID= @MODI_PRID,
                                //       FLAG= FLAG + 1,
                                //        TF042  = @TF042
                                //        WHERE TF001 = @TF001
                                //        AND TF002 = @TF002";
                                //dynamicParameters.AddDynamicParams(
                                //new
                                //{
                                //    TF042 = ReceiveDateErp,
                                //    TF001 = TH032,
                                //    TF002 = TH033
                                //});
                                //rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                //#endregion

                                //#region //INVTG 暫出入轉撥單身檔 轉銷數量
                                //dynamicParameters = new DynamicParameters();
                                //sql = @"UPDATE INVTG SET
                                //        TG020  = TG020 + @TG020,
                                //        TG046  = TG046 + @TG046,
                                //        TG054  = TG054 + @TG054,
                                //        TG024  = @TG024
                                //        WHERE TG001 = @TG001
                                //        AND TG002 = @TG002
                                //        AND TG003 = @TG003";
                                //dynamicParameters.AddDynamicParams(
                                //new
                                //{
                                //    TG020 = Quantity,
                                //    TG046 = FreeSpareQty,
                                //    TG054 = AvailableQty,
                                //    TG001 = TH032,
                                //    TG002 = TH033,
                                //    TG003 = TH034,
                                //    TG024 = ClosureStatusINVTG
                                //});
                                //rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                //#endregion
                            }
                            #endregion
                            #endregion
                        }
                        #endregion
                    }
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //若成功則將相關資料寫回MES
                        #region //更新 - 進貨單單頭
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.GoodsReceipt SET
                                ConfirmStatus = @ConfirmStatus,
                                TransferStatus = @TransferStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE GrId = @GrId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmStatus = "N",
                                TransferStatus = "R",
                                ConfirmUserId = LastModifiedBy,
                                LastModifiedDate,
                                LastModifiedBy,
                                GrId
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //更新 - 進貨單單單身
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.GrDetail SET
                                ConfirmStatus = @ConfirmStatus,
                                ConfirmUser = @ConfirmUser,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE GrId = @GrId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmStatus = "N",
                                ConfirmUser = LastModifiedBy,
                                LastModifiedDate,
                                LastModifiedBy,
                                GrId
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region /撈取進貨單單身資料=>更新採購單,暫入單
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.AcceptQty,a.AvailableQty,a.PoDetailId
                                FROM SCM.GrDetail a
                                WHERE a.GrId = @GrId";
                        dynamicParameters.Add("GrId", GrId);

                        var resultGrDetail = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in resultGrDetail)
                        {
                            #region //更新 - 採購單 已交數量,已交計價數量
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.PoDetail SET
                                    SiQty = SiQty - @SiQty,
                                    SiPriceQty = SiPriceQty - @SiPriceQty,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE PoDetailId = @PoDetailId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    SiQty = item.AcceptQty,
                                    SiPriceQty = item.AvailableQty,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    PoDetailId = item.PoDetailId
                                });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //更新 - 暫入單 轉進銷量,轉銷贈/備品量
                            //if (item.TsnDetailId > 0)
                            //{
                            //    dynamicParameters = new DynamicParameters();
                            //    sql = @"UPDATE SCM.TsnDetail SET
                            //            SaleQty = SaleQty + @SaleQty,
                            //            SaleFSQty = SaleFSQty + @SaleFSQty,
                            //            LastModifiedDate = @LastModifiedDate,
                            //            LastModifiedBy = @LastModifiedBy
                            //            WHERE TsnDetailId = @TsnDetailId";
                            //    dynamicParameters.AddDynamicParams(
                            //        new
                            //        {
                            //            SaleQty = item.Quantity,
                            //            SaleFSQty = item.FreeSpareQty,
                            //            LastModifiedDate,
                            //            LastModifiedBy,
                            //            TsnDetailId = item.TsnDetailId
                            //        });
                            //    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            //}
                            #endregion

                        }
                        #endregion
                        #endregion

                        #region //批號相關異動
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.GrSeq, a.LotNumber,a.MtlItemId,b.LotManagement,a.InventoryId,a.AcceptQty
                                FROM SCM.GrDetail a
                                INNER JOIN PDM.MtlItem b on a.MtlItemId = b.MtlItemId
                                WHERE a.GrId = @GrId";
                        dynamicParameters.Add("GrId", GrId);

                        var resultRoMes = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in resultRoMes)
                        {
                            string GrSeq = item.GrSeq;
                            int MtlItemId = item.MtlItemId;
                            string LotNumber = item.LotNumber;
                            string LotManagement = item.LotManagement;
                            int InventoryId = item.InventoryId;
                            double AcceptQty = Convert.ToDouble(item.AcceptQty);
                            if (LotManagement != "N")
                            {
                                #region //確認品號資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.MtlItemNo, a.LotManagement
                                    FROM PDM.MtlItem a 
                                    WHERE a.MtlItemId = @MtlItemId";
                                dynamicParameters.Add("MtlItemId", MtlItemId);

                                var MtlItemResult = sqlConnection.Query(sql, dynamicParameters);

                                if (MtlItemResult.Count() <= 0) throw new SystemException("品號資料錯誤!!");

                                string MtlItemNo = "";
                                foreach (var item1 in MtlItemResult)
                                {
                                    if (item.LotManagement != "T") throw new SystemException("品號批號管控參數錯誤!!");
                                    MtlItemNo = item.MtlItemNo;
                                }
                                #endregion

                                #region //確認此品號是否已存在此批號
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 LotNumberId
                                    FROM SCM.LotNumber a 
                                    WHERE a.MtlItemId = @MtlItemId
                                    AND a.LotNumberNo = @LotNumberNo";
                                dynamicParameters.Add("MtlItemId", MtlItemId);
                                dynamicParameters.Add("LotNumberNo", LotNumber);

                                var LotNumberResult = sqlConnection.Query(sql, dynamicParameters);

                                if (LotNumberResult.Count() > 0)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 LotNumberId
                                            FROM SCM.LnDetail a 
                                            WHERE 1=1
                                            AND a.TransactionDate = @TransactionDate
                                            AND a.FromErpPrefix = @FromErpPrefix
                                            AND a.FromErpNo = @FromErpNo
                                            AND a.FromSeq = @FromSeq
                                            AND a.InventoryId = @InventoryId
                                            AND a.TransactionType = @TransactionType
                                            AND a.DocType = @DocType";
                                    dynamicParameters.Add("TransactionDate", ReceiptDate);
                                    dynamicParameters.Add("FromErpPrefix", GrErpPrefix);
                                    dynamicParameters.Add("FromErpNo", GrErpNo);
                                    dynamicParameters.Add("FromSeq", GrSeq);
                                    dynamicParameters.Add("InventoryId", InventoryId);
                                    dynamicParameters.Add("TransactionType", 1);
                                    dynamicParameters.Add("DocType", 1);

                                    var LnDetailResult = sqlConnection.Query(sql, dynamicParameters);
                                    if (LotNumberResult.Count() > 0)
                                    {
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"DELETE SCM.LnDetail
                                                WHERE 1=1
                                                AND TransactionDate = @TransactionDate
                                                AND FromErpPrefix = @FromErpPrefix
                                                AND FromErpNo = @FromErpNo
                                                AND FromSeq = @FromSeq
                                                AND InventoryId = @InventoryId
                                                AND TransactionType = @TransactionType
                                                AND DocType = @DocType";
                                        dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            TransactionDate = ReceiptDate,
                                            FromErpPrefix = GrErpPrefix,
                                            FromErpNo = GrErpNo,
                                            FromSeq = GrSeq,
                                            InventoryId,
                                            TransactionType = 1,
                                            DocType = 1,
                                        });
                                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    }
                                }

                                foreach (var item1 in LotNumberResult)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 LotNumberId
                                            FROM MES.VehicleDetail a 
                                            WHERE 1=1
                                            AND a.LotNumberId = @LotNumberId";
                                    dynamicParameters.Add("LotNumberId", item1.LotNumberId);

                                    var VehicleResult = sqlConnection.Query(sql, dynamicParameters);
                                    if(VehicleResult.Count() > 0) throw new SystemException("批號已經被使用,不可以執行反確認!!");


                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 LotNumberId
                                            FROM SCM.LnDetail a 
                                            WHERE 1=1
                                            AND a.LotNumberId = @LotNumberId";
                                    dynamicParameters.Add("LotNumberId", item1.LotNumberId);

                                    var LnDetailResult = sqlConnection.Query(sql, dynamicParameters);
                                    if (LnDetailResult.Count() <= 0)
                                    {
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"DELETE SCM.LotNumber
                                                WHERE 1=1
                                                AND LotNumberId = @LotNumberId
                                                ";
                                        dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            LotNumberId = item1.LotNumberId,
                                        });
                                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    }
                                }

                                #endregion
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

        #region //編輯批號綁定條碼 GPAI 20240412 **
        public string UpdateLnBarcode(int LnBarcodeId, int BarcidrQty)
        {
            try
            {
                int rowsAffected = 0;

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

                            #region //確認data
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.*
                                    FROM SCM.LnBarcode a 
                                    WHERE a.LnBarcodeId = @LnBarcodeId";
                            dynamicParameters.Add("LnBarcodeId", LnBarcodeId);

                            var LnBarcodeResult = sqlConnection.Query(sql, dynamicParameters);
                            if (LnBarcodeResult.Count() <= 0) throw new SystemException("無此條碼綁定資訊!");


                            #endregion

                            #region //刪除
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE SCM.LnBarcode
                                        WHERE LnBarcodeId = @LnBarcodeId";
                            dynamicParameters.Add("LnBarcodeId", LnBarcodeId);

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

        #region //UpdateGrdetailConfirm -- 確認進貨單單筆單身 -- GPAI 2025-02-13 
        public string UpdateGrdetailConfirm(int GrDetailId)
        {
            try
            {
                #region //確認前先進行拋轉
                //var TransferResult = UpdateTransferGoodsReceipt(GrId);
                //JObject transferResultJson = JObject.Parse(TransferResult);
                //if (transferResultJson["status"].ToString() != "success")
                //{
                //    throw new SystemException(transferResultJson["msg"].ToString());
                //}
                #endregion

                #region //參數
                string dateNow = DateTime.Now.ToString("yyyyMMdd");
                string timeNow = DateTime.Now.ToString("HH:mm:ss");
                string dateNowMonth = DateTime.Now.ToString("yyyyMM");

                string UserNo = "";

                int rowsAffected = 0;
                var ErpNo = "";
                string ErpDbName = "";
                string USR_GROUP = "";
                int UserId = CreateBy;

                string GrErpPrefix = "";
                string GrErpNo = "";
                string ReceiptDate = "";
                int CompanyIdBase = -1;
                string GrSeq = "";

                int grId = -1;

                string MaxReceiptDate = "";

                var TH001 = "";
                var TH002 = "";
                var TH003 = "";
                var TH004 = "";
                var TH005 = "";
                var TH006 = "";
                var TH007 = 0.000000;
                var TH008 = "";
                var TH009 = "";
                var TH010 = "";
                var TH011 = "";
                var TH012 = "";
                var TH013 = "";
                var TH014 = "";
                var TH015 = 0.000000;
                var TH016 = 0.000000;
                var TH017 = 0.000000;
                var TH018 = 0.000000;
                var TH019 = 0.000000;
                var TH020 = 0.000000;
                var TH021 = "";
                var TH022 = "";
                var TH023 = "";
                var TH024 = 0.000000;
                var TH025 = "";
                var TH026 = "";
                var TH027 = "";
                var TH028 = "";
                var TH029 = "";
                var TH030 = "";
                var TH031 = "";
                var TH032 = "";
                var TH033 = "";
                var TH034 = 0.000000;
                var TH035 = "";
                var TH036 = "";
                var TH037 = "";
                var TH038 = "";
                var TH039 = "";
                var TH040 = "";
                var TH041 = "";
                var TH042 ="";
                var TH043 = "";
                var TH044 = "";
                var TH045 = 0.000000;
                var TH046 = 0.000000;
                var TH047 = 0.000000;
                var TH048 = 0.000000;
                var TH049 = 0.000000;
                var TH050 = 0.000000;
                var TH051 = 0.000000;
                var TH052 = 0.000000;
                var TH053 = "";
                var TH054 = "";
                var TH055 = 0.000000;
                var TH056 = "";
                var TH057 = "";
                var TH058 = "";
                var TH059 = "";
                var TH060 = "";
                var TH061 = 0.000000;
                var TH062 = "";
                var TH063 = "";
                var TH064 = "";
                var TH065 = "";
                var TH066 = "";
                var TH067 = 0.000000;
                var TH068 = 0.000000;
                var TH069 = "";
                var TH070 = "";
                var TH071 = "";
                var TH072 = "";
                var TH073 = 0.000000;
                var TH074 = "";
                var TH500 = "";
                var TH501 = "";
                var TH502 = "";
                var TH503 = "";
                var TH550 = "";
                var TH551 = "";
                var TH552 = "";
                var TH553 = "";
                var TH554 = "";
                var TH555 = "";
                var TH556 = "";
                var TH557 = "";
                var TH558 = "";
                var TH075 = "";
                var TH076 = "";
                var TH077 = "";
                var TH078 = "";
                var TH079 = "";
                var TH080 = "";
                var TH081 = "";
                var TH082 = "";
                var TH083 = "";
                var TH084 = "";
                var TH085 = "";
                var TH086 = "";
                var TH087 = "";
                var TH088 = "";
                var TH089 = 0.000000;
                var TH090 = 0.000000;
                var TH091 = "";
                var TH092 = "";

                string CloseStatus = "";
                string PoErpPrefix = ""; //TH011 採購單別 單號 序號
                string PoErpNo = ""; //TH012 
                string PoSeq = ""; //TH013 
                #endregion

                using (TransactionScope transactionScope = new TransactionScope())
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

                        var resultCompany = sqlConnection.Query(sql, dynamicParameters);

                        if (resultCompany.Count() <= 0) throw new SystemException("公司別錯誤!");

                        foreach (var item in resultCompany)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            ErpDbName = ErpConnectionStrings.Split('=')[2].Split(';')[0];
                            ErpNo = item.ErpNo;
                        }
                        #endregion

                        #region //撈取使用者No
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
                        dynamicParameters.Add("FunctionCode", "GoodsReceiptManagement");
                        dynamicParameters.Add("DetailCode", "confirm");

                        var resultUser = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUser.Count() <= 0) throw new SystemException("使用者資料錯誤!");

                        foreach (var item in resultUser)
                        {
                            UserNo = item.UserNo;
                        }
                        #endregion

                        #region /撈取單身資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.TransferStatus ,b.GrDetailId, b.GrId, b.GrErpPrefix, b.GrErpNo, b.GrSeq, b.MtlItemId, b.MtlItemNo, b.GrMtlItemName, b.GrMtlItemSpec
                                , b.ReceiptQty, b.UomId, b.UomNo, b.InventoryId, b.InventoryNo, b.LotNumber, b.PoErpPrefix, b.PoErpNo, b.PoSeq
                                , b.AcceptanceDate, CASE WHEN LEN(b.AcceptanceDate) > 0 THEN FORMAT(CAST(b.AcceptanceDate as date), 'yyyyMMdd') ELSE '' END AS ErpAcceptanceDate
                                , b.AcceptQty, b.AvailableQty, b.ReturnQty, b.OrigUnitPrice, b.OrigAmount, b.OrigDiscountAmt, b.TrErpPrefix, b.TrErpNo, b.TrSeq
                                , b.ReceiptExpense, b.DiscountDescription, b.PaymentHold, b.Overdue, b.QcStatus, b.ReturnCode, b.ConfirmStatus, b.CloseStatus
                                , b.ReNewStatus, b.Remark, b.InventoryQty, b.SmallUnit
                                , CASE WHEN FORMAT(b.AvailableDate, 'yyyy-MM-dd') = '1900-01-01' THEN null ELSE FORMAT(CAST(b.AvailableDate as date), 'yyyy-MM-dd') END AS AvailableDate
                                , CASE WHEN LEN(b.AvailableDate) > 0 THEN FORMAT(CAST(b.AvailableDate as date), 'yyyyMMdd') ELSE '' END AS ErpAvailableDate
                                , CASE WHEN FORMAT(b.ReCheckDate, 'yyyy-MM-dd') = '1900-01-01' THEN null ELSE FORMAT(CAST(b.ReCheckDate as date), 'yyyy-MM-dd') END AS ReCheckDate
                                , CASE WHEN LEN(b.ReCheckDate) > 0 THEN FORMAT(CAST(b.ReCheckDate as date), 'yyyyMMdd') ELSE '' END AS ErpReCheckDate
                                , b.ConfirmUser, b.ApvErpPrefix, b.ApvErpNo
                                , b.ApvSeq, b.ProjectCode, b.ExpenseEntry, b.PremiumAmountFlag, b.OrigPreTaxAmt, b.OrigTaxAmt, b.PreTaxAmt, b.TaxAmt, b.ReceiptPackageQty
                                , b.AcceptancePackageQty, b.ReturnPackageQty, b.PremiumAmount, b.PackageUnit, b.ReserveTaxCode, b.OrigPremiumAmount, b.AvailableUomId
                                , b.AvailableUomNo, b.OrigCustomer, b.ApproveStatus, b.EbcErpPreNo, b.EbcEdition, b.ProductSeqAmount, b.MtlItemType, b.Loaction
                                , b.TaxRate, b.MultipleFlag, b.GrErpStatus, b.TaxCode, b.DiscountRate, b.DiscountPrice, b.QcType, b.TransferStatus, b.TransferDate
                                , b.CreateDate, b.LastModifiedDate, b.CreateBy, b.LastModifiedBy
                                , a.ConfirmStatus,a.CompanyId
                                ,FORMAT(a.ReceiptDate, 'yyyyMMdd') ReceiptDate,a.TransferStatus
								,x.AcceptanceDateCount,a.GrId
                                FROM SCM.GoodsReceipt a
                                INNER JOIN SCM.GrDetail b on a.GrId = b.GrId
                                OUTER APPLY(
                                    SELECT CASE WHEN COUNT(DISTINCT FORMAT(x1.AcceptanceDate, 'yyyyMM'))  >1 
                                    THEN 'N' ELSE 'Y' 
                                    END AS AcceptanceDateCount
                                    FROM SCM.GrDetail x1
                                    WHERE x1.GrId = a.GrId
                                ) x
                                WHERE b.GrDetailId = @GrDetailId";
                        dynamicParameters.Add("GrDetailId", GrDetailId);

                        var resultRoMes = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in resultRoMes)
                        {
                            //if (item.TransferStatus != "Y") throw new SystemException("該進貨單單據尚未拋轉,請重新拋轉");
                            if (item.ConfirmStatus == "Y") throw new SystemException("該進貨單單據已經核單,不能操作");
                            if (item.ConfirmStatus == "V") throw new SystemException("該進貨單單據已經作廢,不能操作");
                            if (item.CompanyId != CurrentCompany) throw new SystemException("該進貨單單據公司別與後端公司別不一致,不能操作");
                            if (item.CloseStatus == "Y") throw new SystemException("該進貨單已經結帳,不能操作");
                            if (item.AcceptanceDateCount == "N") throw new SystemException("進貨單單身驗收日須同一天,才能操作");

                            #region //參數
                            CompanyIdBase = item.CompanyId;
                            GrErpPrefix = item.GrErpPrefix;
                            GrErpNo = item.GrErpNo;
                            ReceiptDate = item.ReceiptDate;
                            string AvailableDate = item.AvailableDate == null ? "" : ((DateTime)item.AvailableDate).ToString("yyyyMMdd");
                            string ReCheckDate = item.ReCheckDate == null ? "" : ((DateTime)item.ReCheckDate).ToString("yyyyMMdd");

                            TH001 = item.GrErpPrefix;
                            TH002 = item.GrErpNo;
                            TH003 = item.GrSeq;
                            TH004 = item.MtlItemNo;
                            TH005 = item.GrMtlItemName;
                            TH006 = item.GrMtlItemSpec;
                            TH007 = item.ReceiptQty;
                            TH008 = item.UomNo;
                            TH009 = item.InventoryNo;
                            TH010 = item.LotNumber;
                            TH011 = item.PoErpPrefix;
                            TH012 = item.PoErpNo;
                            TH013 = item.PoSeq;
                            TH014 = item.ErpAcceptanceDate != null ? item.ErpAcceptanceDate : "";
                            TH015 = item.AcceptQty;
                            TH016 = item.AvailableQty;
                            TH017 = item.ReturnQty;
                            TH018 = item.OrigUnitPrice;
                            TH019 = item.OrigAmount;
                            TH020 = item.OrigDiscountAmt;
                            TH021 = item.TrErpPrefix;
                            TH022 = item.TrErpNo;
                            TH023 = item.TrSeq;
                            TH024 = item.ReceiptExpense;
                            TH025 = item.DiscountDescription;
                            TH026 = item.PaymentHold;
                            TH027 = item.Overdue;
                            TH028 = item.QcStatus;
                            TH029 = item.ReturnCode;
                            TH030 = item.ConfirmStatus;
                            TH031 = item.CloseStatus;
                            TH032 = item.ReNewStatus;
                            TH033 = item.Remark;
                            TH034 = item.InventoryQty;
                            TH035 = item.SmallUnit;
                            TH036 = AvailableDate;
                            TH037 = ReCheckDate;
                            //TH038 = item.ConfirmUser;
                            TH039 = item.ApvErpPrefix;
                            TH040 = item.ApvErpNo;
                            TH041 = item.ApvSeq;
                            TH042 = item.ProjectCode;
                            TH043 = item.ExpenseEntry;
                            TH044 = item.PremiumAmountFlag;
                            TH045 = item.OrigPreTaxAmt;
                            TH046 = item.OrigTaxAmt;
                            TH047 = item.PreTaxAmt;
                            TH048 = item.TaxAmt;
                            TH049 = item.ReceiptPackageQty;
                            TH050 = item.AcceptancePackageQty;
                            TH051 = item.ReturnPackageQty;
                            TH052 = item.PremiumAmount;
                            TH053 = item.PackageUnit;
                            TH054 = item.ReserveTaxCode;
                            TH055 = item.OrigPremiumAmount;
                            TH056 = item.AvailableUomNo;
                            TH057 = item.OrigCustomer;
                            TH058 = item.ApproveStatus;
                            TH059 = item.EbcErpPreNo;
                            TH060 = item.EbcEdition;
                            TH061 = item.ProductSeqAmount;
                            TH062 = item.MtlItemType;
                            TH063 = item.Loaction;
                            TH064 = "";
                            TH065 = "";
                            TH066 = "";
                            TH067 = 0.000000;
                            TH068 = 0.000000;
                            TH069 = "";
                            TH070 = "";
                            TH071 = "";
                            TH072 = "";
                            TH073 = item.TaxRate;
                            TH074 = item.MultipleFlag;
                            TH500 = "";
                            TH501 = "";
                            TH502 = "";
                            TH503 = "";
                            TH550 = "";
                            TH551 = "";
                            TH552 = "";
                            TH553 = "";
                            TH554 = "";
                            TH555 = "";
                            TH556 = "";
                            TH557 = "";
                            TH558 = "";
                            TH075 = "";
                            TH076 = "";
                            TH077 = "";
                            TH078 = "";
                            TH079 = "";
                            TH080 = "";
                            TH081 = "";
                            TH082 = item.GrErpStatus;
                            TH083 = "";
                            TH084 = "";
                            TH085 = "";
                            TH086 = "";
                            TH087 = "";
                            TH088 = item.TaxCode;
                            TH089 = item.DiscountRate;
                            TH090 = item.DiscountPrice;
                            TH091 = item.QcType;
                            TH092 = "";
                            //GrDetailId = item.GrDetailId;
                            grId = item.GrId;
                            #endregion

                            #region //判斷進貨檢驗單據條件
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.QcRecordId, a.QcStatus
                                    , b.GrSeq
                                    FROM MES.QcGoodsReceipt a 
                                    INNER JOIN SCM.GrDetail b ON a.GrDetailId = b.GrDetailId
                                    WHERE a.GrDetailId = @GrDetailId
                                    ORDER BY a.CreateDate DESC";
                            dynamicParameters.Add("GrDetailId", GrDetailId);

                            var QcGoodsReceiptResult = sqlConnection.Query(sql, dynamicParameters);

                            if (QcGoodsReceiptResult.Count() > 0)
                            {
                                bool qcFlag = false;
                                foreach (var item2 in QcGoodsReceiptResult)
                                {
                                    GrSeq = item2.GrSeq;

                                    if (item2.QcStatus == "A")
                                    {
                                        throw new SystemException("單身序號【" + item2.GrSeq + "】已開立進貨檢驗單據【" + item2.QcRecordId + "】，但尚未量測完成!!");
                                    }

                                    qcFlag = true;

                                    #region //確認此量測單據是否有品異單
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 a.JudgeStatus, a.QcRecordId, a.ReleaseQty
                                                , b.AbnormalqualityNo
                                                FROM QMS.AqBarcode a 
                                                INNER JOIN QMS.Abnormalquality b ON a.AbnormalqualityId = b.AbnormalqualityId
                                                WHERE a.QcRecordId = @QcRecordId";
                                    dynamicParameters.Add("QcRecordId", item2.QcRecordId);

                                    var AqBarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                                    if (AqBarcodeResult.Count() > 0)
                                    {
                                        foreach (var item3 in AqBarcodeResult)
                                        {
                                            if (item3.JudgeStatus == null)
                                            {
                                                throw new SystemException("品異單據【" + item3.AbnormalqualityNo + "】尚未進行判定!!");
                                            }
                                            else if (item3.JudgeStatus == "AM")
                                            {
                                                qcFlag = false;
                                            }
                                            else if (item3.JudgeStatus == "R")
                                            {
                                                if (item3.ReleaseQty == null) throw new SystemException("品異單據【" + item3.AbnormalqualityNo + "】尚未維護特採數量!!");

                                                #region //確認驗收數量是否正確
                                                if (item.AcceptQty > item3.ReleaseQty) throw new SystemException("驗收數量大於特採數量!!");
                                                #endregion
                                            }
                                            else if (item3.JudgeStatus == "S")
                                            {
                                                #region //若判定為報廢，驗收數量需為0
                                                if (item.AcceptQty != 0) throw new SystemException("單身序號【" + item2.GrSeq + "】判定為報廢，驗收數量需為0!!");
                                                #endregion
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (item2.QcStatus != "Y")
                                        {
                                            throw new SystemException("進貨檢驗單號【" + item2.QcRecordId + "】狀態非合格，但查無品異單據!!");
                                        }
                                    }
                                    #endregion
                                }

                                if (qcFlag == false)
                                {
                                    throw new SystemException("單身序號【" + GrSeq + "】已有開立品異單判定為【重新量測】，但尚未有新量測單據!!");
                                }
                            }
                            else
                            {
                                if (item.QcStatus == "1") throw new SystemException("品號【" + item.MtlItemNo + "】為待驗狀態，需進行檢驗!!");
                            }
                            #endregion
                        }
                        #endregion
                    }
                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //審核ERP權限
                        USR_GROUP = BaseHelper.CheckErpAuthority(UserNo, sqlConnection, "PURI13", "CONFIRM");
                        #endregion

                        #region //判斷開立日期是否超過結帳日
                        string CloseDateBase = "";
                        string ReceiveDateBase = ReceiptDate;
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT MA011,MA012,MA013
                                FROM CMSMA ";
                        var resultCloseDate = sqlConnection.Query(sql, dynamicParameters);
                        foreach (var item in resultCloseDate)
                        {
                            CloseDateBase = item.MA013;
                        }

                        DateTime docDateBase;
                        DateTime closeDateBase;

                        if (DateTime.TryParseExact(ReceiveDateBase, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out docDateBase) &&
                            DateTime.TryParseExact(CloseDateBase, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out closeDateBase))
                        {
                            int resultDate = DateTime.Compare(docDateBase, closeDateBase);

                            if (resultDate <= 0)
                            {
                                throw new SystemException("【進貨單】ERP已經結帳,不可以在新增【" + CloseDateBase + "】之前的單據");
                            }
                        }
                        else
                        {
                            throw new SystemException("日期字符串无效，无法比较");
                        }
                        #endregion

                        #region //取出進貨單單據資料 + 異動更新相關資料表
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.TG003,a.TG005,a.TG007,a.TG013,a.TG014
                                ,b.TH003,b.TH004,b.TH007,b.TH009,b.TH010
                                ,b.TH011,b.TH012,b.TH013
                                ,b.TH014,b.TH015,b.TH016,b.TH017,b.TH030,b.TH031,b.TH063
                                ,b.TH021,b.TH022,b.TH023
                                ,b.TH036,b.TH037,b.TH047
                                ,c.MB022
                                ,d.MQ019,md3.MF005
                                ,ISNULL(md1.GrRate,1) GrRate
                                ,ISNULL(md2.GrAvailableRate,1) GrAvailableRate
                                ,(b.TH015*ISNULL(md1.GrRate,1)) ConversionQty
                                ,(b.TH016*ISNULL(md2.GrAvailableRate,1)) ConversionAvailableQty
                                ,CASE 
                                    WHEN b.TH016 != 0 THEN ROUND((b.TH047 / (b.TH016*ISNULL(md2.GrAvailableRate,1))),CAST(md3.MF005 AS INT)) 
                                    ELSE 0
                                END AS UnitCost
                                FROM PURTG a
                                INNER JOIN PURTH b on a.TG001 = b.TH001 AND a.TG002 = b.TH002
                                INNER JOIN INVMB c on b.TH004 = c.MB001
                                INNER JOIN CMSMQ d ON d.MQ001=a.TG001 AND d.MQ003='34'
                                INNER JOIN CMSMC e ON e.MC001=b.TH009
                                LEFT JOIN INVMD f ON f.MD001=b.TH004
                                OUTER APPLY(
                                    SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) GrRate
                                    FROM INVMD x1
                                    WHERE x1.MD001=b.TH004
                                    AND  x1.MD002 = b.TH008
                                ) md1
                                OUTER APPLY(
                                    SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) GrAvailableRate
                                    FROM INVMD x1
                                    WHERE x1.MD001=b.TH004
                                    AND  x1.MD002 = b.TH056
                                ) md2
                                OUTER APPLY(
                                    SELECT x1.MF003,x1.MF004,x1.MF005,x1.MF006
                                    FROM CMSMF x1
                                    INNER JOIN CMSMA x2 on x1.MF001 = x2.MA003 
                                ) md3
                                WHERE b.TH001 = @GrErpPrefix
                                AND b.TH002 = @GrErpNo
                                AND b.TH003 = @GrSeq";
                        dynamicParameters.Add("GrErpPrefix", GrErpPrefix);
                        dynamicParameters.Add("GrErpNo", GrErpNo);
                        dynamicParameters.Add("GrSeq", GrSeq);

                        var resultRoErp = sqlConnection.Query(sql, dynamicParameters);
                        if (resultRoErp.Count() <= 0) throw new SystemException("找不到ERP進貨單單據!");
                        foreach (var item in resultRoErp)
                        {
                            if (item.TG013 == "Y") throw new SystemException("該進貨單單據已經核單,不能操作");
                            if (item.TG013 == "V") throw new SystemException("該進貨單單據已經作廢,不能操作");
                            if (item.TH030 == "Y") throw new SystemException("該進貨單單據單身已經核單,不能操作");
                            if (item.TH031 == "Y") throw new SystemException("該進貨單單據已經結帳,不能操作");
                            if (item.TH021 != "") throw new SystemException("來源暫入單模式尚未開發,不能操作");
                            if (item.MQ019 == "N") throw new SystemException("不核對採購單模式尚未開發,不能操作");

                            #region //欄位參數撈取
                            string ReceiveDateErp = item.TG003; //TG003 進貨日期
                            string SupplierNo = item.TG005; //TG005 供應商
                            string CurrencyCode = item.TG007; //TG007 幣別
                            Double Exchange = Convert.ToDouble(item.TG008); //TG008 匯率
                            string ConfirmStatus = item.TG013; //TG013 單頭確認
                            string DocDate = item.TG014; //TG013 單頭確認

                            //string GrSeq = item.TH003; //TH003 進貨單序號
                            string GrFull = GrErpPrefix + '-' + GrErpNo + '-' + GrSeq;
                            string MtlItemNo = item.TH004; //TH004 品號
                            string InventoryNo = item.TH009; //TH009 庫別
                            string LotNumber = item.TH010; //TH010 批號
                            PoErpPrefix = item.TH011; //TH011 採購單別 單號 序號
                            PoErpNo = item.TH012; //TH012 
                            PoSeq = item.TH013; //TH013 
                            string PoFull = item.TH011 + '-' + item.TH012 + '-' + item.TH013;
                            //string TH014 = item.TH014; //TH014 驗收日期
                            Double ReceiptQty = Convert.ToDouble(item.TH007); //TH014 進貨數量
                            Double AcceptQty = Convert.ToDouble(item.TH015); //TH014 驗收數量
                            Double AvailableQty = Convert.ToDouble(item.TH016); //TH014 計價數量 
                            Double ReturnQty = Convert.ToDouble(item.TH017); //TH014 驗退數量
                            //string TH030 = item.TH030; //TH030 確認碼
                            CloseStatus = item.TH031; //TH031 結帳碼
                            string AvailableDate = item.TH036; //TH036 有效日期
                            string ReCheckDate = item.TH037; //TH037 複檢日期
                            string Location = item.TH063; //TH060 儲位
                            string TrErpPrefix = item.TH021; //TH021 暫入單別 / TH022 暫入單號 / TH023 暫入序號
                            string TrErpNo = item.TH022;
                            string TrSeq = item.TH023;
                            string TsnFull = item.TH021 + '-' + item.TH022 + '-' + item.TH023;
                            string LotManagement = item.MB022; //品號批號管理

                            Double GrRate = Convert.ToDouble(item.GrRate); //單位換算率
                            Double ConversionQty = Convert.ToDouble(TH015*item.GrRate); //單位換算後數量
                            Double ConversionAvailableQty = Convert.ToDouble(TH016*item.GrAvailableRate); //單位換算後計價數量
                            Double UnitCost = Convert.ToDouble(TH047 / (TH016 * item.GrAvailableRate)); //單位成本
                            //Double TH047 = Convert.ToDouble(item.TH047); //本幣未稅金額
                           // Double TH018 = Convert.ToDouble(item.TH018); //原幣單價
                            Double PoRate = 0; //採購單位換算率
                            Double PoAvailableRate = 0; //採購計價單位換算率
                            Double PoConversionQty = 0; //採購單位換算後數量
                            Double PoConversionAvailableQty = 0; //採購單位換算後計價數量

                            DateTime dateC = DateTime.ParseExact(DocDate, "yyyyMMdd", null);
                            DateTime dateD = DateTime.ParseExact(TH014, "yyyyMMdd", null);
                            if (dateC > dateD)
                            {
                                throw new SystemException("資料異常:單身驗收日不可以小於單據日!!");
                            }
                            if (MaxReceiptDate != "")
                            {
                                DateTime dateA = DateTime.ParseExact(item.TH014, "yyyyMMdd", null);
                                DateTime dateB = DateTime.ParseExact(MaxReceiptDate, "yyyyMMdd", null);
                                //MaxReceiptDate = date.ToString("yyyy-MM-dd"); // yyyy-MM-dd 格式


                                //DateTime dateA = DateTime.Parse(item.TH014);
                                //DateTime dateB = DateTime.Parse(MaxReceiptDate);
                                if (dateA > dateB)
                                {
                                    MaxReceiptDate = item.TH014;
                                }
                            }
                            else
                            {
                                MaxReceiptDate = item.TH014;
                            }
                            #endregion

                            #region //檢查段

                            #region //數量查驗
                            if (ReceiptQty < 0) throw new SystemException("【進貨單:" + GrFull + "】進貨數量不可小於0!!!");
                            if (AcceptQty < 0) throw new SystemException("【進貨單:" + GrFull + "】驗收數量不可小於0!!!");
                            if (AvailableQty < 0) throw new SystemException("【進貨單:" + GrFull + "】計價數量不可小於0!!!");
                            if (ReceiptQty < AcceptQty) throw new SystemException("【進貨單:" + GrFull + "】驗收數量不可大於進貨數量!!!");
                            if (AcceptQty < AvailableQty) throw new SystemException("【進貨單:" + GrFull + "】計價數量不可大於驗收數量!!!");
                            #endregion

                            #region //PURTD 採購單
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 ISNULL(a.TD008,0) TD008,ISNULL(a.TD015,0) TD015,a.TD016,a.TD018,
                                    ISNULL(md1.PoRate,1) PoRate, 
                                    ISNULL(md2.PoAvailableRate,1) PoAvailableRate, 
                                    (a.TD008*ISNULL(md1.PoRate,1)) PoConversionQty,
                                    (a.TD015*ISNULL(md1.PoRate,1)) PoConversionSiQty,
                                    (a.TD058*ISNULL(md2.PoAvailableRate,1)) PoConversionAvailableQty,
                                    (a.TD060*ISNULL(md2.PoAvailableRate,1)) PoConversionAvailableSiQty
                                    FROM PURTD a
                                    OUTER APPLY(
                                        SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) PoRate
                                        FROM INVMD x1
                                        WHERE x1.MD001 = a.TD004
                                        AND  x1.MD002 = a.TD009
                                    ) md1
                                    OUTER APPLY(
                                        SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) PoAvailableRate
                                        FROM INVMD x1
                                        WHERE x1.MD001 = a.TD004
                                        AND  x1.MD002 = a.TD059
                                    ) md2
                                    WHERE a.TD001 = @TD001
                                    AND a.TD002 = @TD002
                                    AND a.TD003 = @TD003";
                            dynamicParameters.Add("TD001", PoErpPrefix);
                            dynamicParameters.Add("TD002", PoErpNo);
                            dynamicParameters.Add("TD003", PoSeq);
                            var resultCOPTD = sqlConnection.Query(sql, dynamicParameters);
                            if (resultCOPTD.Count() <= 0) throw new SystemException("採購單找不到,請重新確認");
                            foreach (var item1 in resultCOPTD)
                            {
                                if (item1.TD018 != "Y") throw new SystemException("【採購單:" + PoFull + "】非確認狀態,不可操作!!!");
                                if (item1.TD016 == "Y") throw new SystemException("【採購單:" + PoFull + "】已經結案!!!");
                                Double TD008 = Convert.ToDouble(item1.TD008);
                                Double TD015 = Convert.ToDouble(item1.TD015);
                                PoRate = Convert.ToDouble(item1.PoRate);
                                PoAvailableRate = Convert.ToDouble(item1.PoAvailableRate);
                                PoConversionQty = Convert.ToDouble(item1.PoConversionQty);
                                Double PoConversionSiQty = Convert.ToDouble(item1.PoConversionSiQty);
                                PoConversionAvailableQty = Convert.ToDouble(item1.PoConversionAvailableQty);
                                Double PoConversionAvailableSiQty = Convert.ToDouble(item1.PoConversionAvailableSiQty);

                                if (PoConversionQty == PoConversionSiQty + ConversionQty)
                                {
                                    CloseStatus = "Y";
                                }
                                else if (PoConversionQty < PoConversionSiQty + ConversionQty)
                                {
                                    throw new SystemException("【採購單:" + PoFull + "】已超過採購數量,目前已交數量:" + TD015 + ",本次進貨驗收數量:" + AcceptQty + ",不可操作!!!");
                                }
                            }
                            #endregion
                            #endregion

                            #region //確認段
                            #region //固定更新
                            #region //PURMA 廠商基本資料檔
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE PURMA SET
                                    MODIFIER = @MODIFIER,
                                    MODI_DATE = @MODI_DATE,
                                    MODI_TIME = @MODI_TIME,
                                    MODI_AP = @MODI_AP,
                                    MODI_PRID = @MODI_PRID,
                                    FLAG = FLAG + 1,
                                    MA023 = @MA023
                                    WHERE MA001 = @SupplierNo";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                MODIFIER = UserNo,
                                MODI_DATE = dateNow,
                                MODI_TIME = timeNow,
                                MODI_AP = BaseHelper.ClientComputer(),
                                MODI_PRID = "BM",
                                MA023 = ReceiveDateErp,
                                SupplierNo
                            });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //PURMB 品號廠商單頭檔
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE PURMB SET
                                    MODIFIER = @MODIFIER,
                                    MODI_DATE = @MODI_DATE,
                                    MODI_TIME = @MODI_TIME,
                                    MODI_AP = @MODI_AP,
                                    MODI_PRID = @MODI_PRID,
                                    FLAG = FLAG + 1,
                                    MB009 = @MB009
                                    WHERE MB001 = @MB001
                                    AND MB002 = @MB002";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                MODIFIER = UserNo,
                                MODI_DATE = dateNow,
                                MODI_TIME = timeNow,
                                MODI_AP = BaseHelper.ClientComputer(),
                                MODI_PRID = "BM",
                                MB009 = ReceiveDateErp,
                                MB001 = MtlItemNo,
                                MB002 = SupplierNo
                            });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //PURTD 採購單單身資料檔
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE PURTD SET
                                    MODIFIER =@MODIFIER,
                                    MODI_DATE= @MODI_DATE,
                                    MODI_TIME= @MODI_TIME,
                                    MODI_AP= @MODI_AP,
                                    MODI_PRID= @MODI_PRID,
                                    FLAG= FLAG + 1,
                                    TD015 = TD015 + @TD015,
                                    TD016 = @TD016,
                                    TD060 = TD060 + @TD060
                                    WHERE TD001 = @TD001
                                    AND TD002 = @TD002
                                    AND TD003 = @TD003";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                MODIFIER = UserNo,
                                MODI_DATE = dateNow,
                                MODI_TIME = timeNow,
                                MODI_AP = BaseHelper.ClientComputer(),
                                MODI_PRID = "BM",
                                TD015 = Math.Round(ConversionQty / PoRate, 3, MidpointRounding.AwayFromZero),
                                TD016 = CloseStatus,
                                TD060 = Math.Round(ConversionAvailableQty / PoAvailableRate, 3, MidpointRounding.AwayFromZero),
                                TD001 = PoErpPrefix,
                                TD002 = PoErpNo,
                                TD003 = PoSeq
                            });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //PURTH 進貨單單身檔
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE PURTH SET
                                    MODIFIER =@MODIFIER,
                                    MODI_DATE= @MODI_DATE,
                                    MODI_TIME= @MODI_TIME,
                                    MODI_AP= @MODI_AP,
                                    MODI_PRID= @MODI_PRID,
                                    FLAG= FLAG + 1,
                                    TH030 = 'Y',
                                    TH038 = @UserNo,
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
                                    TH031 = @TH031,
                                    TH032 = @TH032,
                                    TH033 = @TH033,
                                    TH034 = @TH034,
                                    TH035 = @TH035,
                                    TH036 = @TH036,
                                    TH037 = @TH037,
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
                                    TH051 = @TH051,
                                    TH052 = @TH052,
                                    TH053 = @TH053,
                                    TH054 = @TH054,
                                    TH055 = @TH055,
                                    TH056 = @TH056,
                                    TH057 = @TH057,
                                    TH058 = @TH058,
                                    TH059 = @TH059,
                                    TH060 = @TH060,
                                    TH061 = @TH061,
                                    TH062 = @TH062,
                                    TH063 = @TH063,
                                    TH064 = @TH064,
                                    TH065 = @TH065,
                                    TH066 = @TH066,
                                    TH067 = @TH067,
                                    TH068 = @TH068,
                                    TH069 = @TH069,
                                    TH070 = @TH070,
                                    TH071 = @TH071,
                                    TH072 = @TH072,
                                    TH073 = @TH073,
                                    TH074 = @TH074,
                                    TH500 = @TH500,
                                    TH501 = @TH501,
                                    TH502 = @TH502,
                                    TH503 = @TH503,
                                    TH550 = @TH550,
                                    TH551 = @TH551,
                                    TH552 = @TH552,
                                    TH553 = @TH553,
                                    TH554 = @TH554,
                                    TH555 = @TH555,
                                    TH556 = @TH556,
                                    TH557 = @TH557,
                                    TH558 = @TH558,
                                    TH075 = @TH075,
                                    TH076 = @TH076,
                                    TH077 = @TH077,
                                    TH078 = @TH078,
                                    TH079 = @TH079,
                                    TH080 = @TH080,
                                    TH081 = @TH081,
                                    TH082 = @TH082,
                                    TH083 = @TH083,
                                    TH084 = @TH084,
                                    TH085 = @TH085,
                                    TH086 = @TH086,
                                    TH087 = @TH087,
                                    TH088 = @TH088,
                                    TH089 = @TH089,
                                    TH090 = @TH090,
                                    TH091 = @TH091,
                                    TH092 = @TH092            
                                    WHERE TH001 = @GrErpPrefix
                                    AND TH002 = @GrErpNo
                                    AND TH003 = @GrSeq";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                MODIFIER = UserNo,
                                MODI_DATE = dateNow,
                                MODI_TIME = timeNow,
                                MODI_AP = BaseHelper.ClientComputer(),
                                MODI_PRID = "BM",
                                GrErpPrefix,
                                GrErpNo,
                                GrSeq,
                                UserNo,
                                TH004,
                                TH005,
                                TH006,
                                TH007,
                                TH008,
                                TH009,
                                TH010,
                                TH011,
                                TH012,
                                TH013,
                                TH014,
                                TH015,
                                TH016,
                                TH017,
                                TH018,
                                TH019,
                                TH020,
                                TH021,
                                TH022,
                                TH023,
                                TH024,
                                TH025,
                                TH026,
                                TH027,
                                TH028,
                                TH029,
                                TH031,
                                TH032,
                                TH033,
                                TH034,
                                TH035,
                                TH036,
                                TH037,
                                TH039,
                                TH040,
                                TH041,
                                TH042,
                                TH043,
                                TH044,
                                TH045,
                                TH046,
                                TH047,
                                TH048,
                                TH049,
                                TH050,
                                TH051,
                                TH052,
                                TH053,
                                TH054,
                                TH055,
                                TH056,
                                TH057,
                                TH058,
                                TH059,
                                TH060,
                                TH061,
                                TH062,
                                TH063,
                                TH064,
                                TH065,
                                TH066,
                                TH067,
                                TH068,
                                TH069,
                                TH070,
                                TH071,
                                TH072,
                                TH073,
                                TH074,
                                TH500,
                                TH501,
                                TH502,
                                TH503,
                                TH550,
                                TH551,
                                TH552,
                                TH553,
                                TH554,
                                TH555,
                                TH556,
                                TH557,
                                TH558,
                                TH075,
                                TH076,
                                TH077,
                                TH078,
                                TH079,
                                TH080,
                                TH081,
                                TH082,
                                TH083,
                                TH084,
                                TH085,
                                TH086,
                                TH087,
                                TH088,
                                TH089,
                                TH090,
                                TH091,
                                TH092
                            });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                            #endregion

                            #region //找是否有未確認單身，若無則更新單頭
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT (a.TH001 + '-' + a.TH002) GrErpFullNo, a.TH003 , a.TH030 detailConfirmStatus
                                    , b.TG013 headerConfirm
                                    FROM PURTH a
                                    INNER JOIN PURTG b ON (a.TH001 + '-' + a.TH002) = (b.TG001 + '-' + b.TG002)
                                    WHERE a.TH001 = @GrErpPrefix AND a.TH002 = @GrErpNo
                                    AND a.TH030 = 'N' 
                                    AND a.TH003 != @GrSeq";
                            dynamicParameters.Add("GrErpPrefix", GrErpPrefix);
                            dynamicParameters.Add("GrErpNo", GrErpNo);
                            dynamicParameters.Add("GrSeq", GrSeq);

                            var unConfirmerpresult = sqlConnection.Query(sql, dynamicParameters);
                            #endregion

                            #region //更新 - 進貨單單頭
                            if (unConfirmerpresult.Count() <= 0)
                            {
                                #region //PURTG 進貨單單頭檔
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE PURTG SET
                                        MODIFIER =@MODIFIER,
                                        MODI_DATE= @MODI_DATE,
                                        MODI_TIME= @MODI_TIME,
                                        MODI_AP= @MODI_AP,
                                        MODI_PRID= @MODI_PRID,
                                        FLAG= FLAG + 1,
                                        TG003 = @TG003,
                                        TG013 = 'Y'
                                        WHERE TG001 = @GrErpPrefix
                                        AND TG002 = @GrErpNo";
                                dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MODIFIER = UserNo,
                                    MODI_DATE = dateNow,
                                    MODI_TIME = timeNow,
                                    MODI_AP = BaseHelper.ClientComputer(),
                                    MODI_PRID = "BM",
                                    GrErpPrefix,
                                    GrErpNo,
                                    TG003 = MaxReceiptDate
                                });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                            }

                            #endregion

                            #region //INVMB 品號基本資料檔 庫存數量 庫存金額
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE INVMB SET
                                    MODIFIER =@MODIFIER,
                                    MODI_DATE= @MODI_DATE,
                                    MODI_TIME= @MODI_TIME,
                                    MODI_AP= @MODI_AP,
                                    MODI_PRID= @MODI_PRID,
                                    FLAG= FLAG + 1,
                                    MB064  = MB064 + @MB064,
                                    MB065  = MB065 + @MB065
                                    WHERE MB001 = @MB001";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                MODIFIER = UserNo,
                                MODI_DATE = dateNow,
                                MODI_TIME = timeNow,
                                MODI_AP = BaseHelper.ClientComputer(),
                                MODI_PRID = "BM",
                                MB048 = CurrencyCode,
                                MB049 = Math.Round(TH018 / GrRate, 4, MidpointRounding.AwayFromZero),
                                MB050 = Math.Round(TH018 * Exchange / GrRate, 4, MidpointRounding.AwayFromZero),
                                MB064 = ConversionQty,
                                MB065 = TH047,
                                MB001 = MtlItemNo
                            });
                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                            #endregion

                            #region //固定新增
                            #region //INVLA 異動明細資料檔
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO INVLA (COMPANY, CREATOR, USR_GROUP, 
                                    CREATE_DATE,MODIFIER, MODI_DATE, FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID,
                                    UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10,
                                    LA001, LA004, LA005, LA006, LA007,
                                    LA008, LA009, LA010, LA011, LA012,
                                    LA013, LA014, LA015, LA016, LA017,
                                    LA018, LA019, LA020, LA021, LA022,
                                    LA023, LA024, LA025, LA026, LA027,
                                    LA028, LA029, LA030, LA031, LA032,
                                    LA033, LA034)
                                    OUTPUT INSERTED.LA001
                                    VALUES (@COMPANY, @CREATOR, @USR_GROUP,
                                    @CREATE_DATE,@MODIFIER,@MODI_DATE,@FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID,
                                    '','','','','',0,0,0,0,0,
                                    @LA001, @LA004, @LA005, @LA006, @LA007,
                                    @LA008, @LA009, @LA010, @LA011, @LA012,
                                    @LA013, @LA014, @LA015, @LA016, @LA017,
                                    @LA018, @LA019, @LA020, @LA021, @LA022,
                                    0, 0, '', '', '',
                                    '', '','','','',
                                    '', '')";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    COMPANY = ErpNo,
                                    CREATOR = UserNo,
                                    USR_GROUP,
                                    CREATE_DATE = dateNow,
                                    MODIFIER = "",
                                    MODI_DATE = "",
                                    FLAG = 1,
                                    CREATE_TIME = timeNow,
                                    CREATE_AP = BaseHelper.ClientComputer(),
                                    CREATE_PRID = "BM",
                                    LA001 = MtlItemNo,
                                    LA004 = MaxReceiptDate,
                                    LA005 = 1,
                                    LA006 = GrErpPrefix,
                                    LA007 = GrErpNo,
                                    LA008 = GrSeq,
                                    LA009 = InventoryNo,
                                    LA010 = SupplierNo,
                                    LA011 = ConversionQty,
                                    LA012 = UnitCost,
                                    LA013 = TH047,
                                    LA014 = "1",
                                    LA015 = "Y",
                                    LA016 = LotNumber,
                                    LA017 = TH047, //金額-材料 暫未確認 
                                    LA018 = 0,
                                    LA019 = 0,
                                    LA020 = 0,
                                    LA021 = 0,
                                    LA022 = Location != "" ? Location : "" //儲位尚未給
                                });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                            #endregion

                            #region //第一次新增後續更新

                            #region //INVMC 品號庫別檔異動
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT MC012 FROM INVMC
                                    WHERE MC001 = @MC001
                                      AND MC002 = @MC002";
                            dynamicParameters.Add("MC001", MtlItemNo);
                            dynamicParameters.Add("MC002", InventoryNo);
                            var resultINVMC = sqlConnection.Query(sql, dynamicParameters);

                            if (resultINVMC.Count() > 0)
                            {
                                #region //後續更新流程
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE INVMC SET
                                        MODIFIER =@MODIFIER,
                                        MODI_DATE= @MODI_DATE,
                                        MODI_TIME= @MODI_TIME,
                                        MODI_AP= @MODI_AP,
                                        MODI_PRID= @MODI_PRID,
                                        FLAG= FLAG + 1,
                                        MC007  = MC007 + @MC007,
                                        MC008  = MC008 + @MC008,
                                        MC012  = @MC012 
                                        WHERE MC001 = @MC001
                                        AND MC002 = @MC002";
                                dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MODIFIER = UserNo,
                                    MODI_DATE = dateNow,
                                    MODI_TIME = timeNow,
                                    MODI_AP = BaseHelper.ClientComputer(),
                                    MODI_PRID = "BM",
                                    MC007 = ConversionQty,
                                    MC008 = TH047,
                                    MC012 = ReceiveDateErp,
                                    MC001 = MtlItemNo,
                                    MC002 = InventoryNo
                                });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }
                            else
                            {
                                #region //第一次新增流程
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO INVMC (COMPANY, CREATOR, USR_GROUP, 
                                        CREATE_DATE,MODIFIER, MODI_DATE, FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID,
                                        MC001, MC002, MC003, MC004, MC005, MC006, MC007, MC008, MC009, MC010, MC011,
                                        MC012, MC013, MC014, MC015)
                                        VALUES (@COMPANY, @CREATOR, @USR_GROUP,
                                        @CREATE_DATE,@MODIFIER,@MODI_DATE,@FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID,
                                        @MC001, @MC002, @MC003, @MC004, @MC005, @MC006, @MC007, @MC008, @MC009, @MC010, @MC011,
                                        @MC012, @MC013, @MC014, @MC015)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        COMPANY = ErpNo,
                                        CREATOR = UserNo,
                                        USR_GROUP,
                                        CREATE_DATE = dateNow,
                                        MODIFIER = "",
                                        MODI_DATE = "",
                                        FLAG = 1,
                                        CREATE_TIME = timeNow,
                                        CREATE_AP = BaseHelper.ClientComputer(),
                                        CREATE_PRID = "BM",
                                        MC001 = MtlItemNo,
                                        MC002 = InventoryNo,
                                        MC003 = "",
                                        MC004 = 0.000,
                                        MC005 = 0.000,
                                        MC006 = 0.000,
                                        MC007 = ConversionQty,
                                        MC008 = TH047,
                                        MC009 = 0.000,
                                        MC010 = 0.00000,
                                        MC011 = "",
                                        MC012 = ReceiveDateErp,
                                        MC013 = "",
                                        MC014 = 0.000,
                                        MC015 = ""
                                    });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                            }
                            #endregion

                            #endregion

                            #region //品號有批號管理觸發
                            if (LotManagement != "N")
                            {
                                #region //固定新增
                                #region //INVLF 儲位批號異動明細資料檔
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO INVLF (COMPANY, CREATOR, USR_GROUP, 
                                        CREATE_DATE,MODIFIER, MODI_DATE, FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID,
                                        UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10,
                                        LF001, LF002, LF003, LF004, LF005,
                                        LF006, LF007, LF008, LF009, LF010,
                                        LF011, LF012, LF013)
                                VALUES (@COMPANY, @CREATOR, @USR_GROUP,
                                        @CREATE_DATE,@MODIFIER,@MODI_DATE,@FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID,
                                        '','','','','',0,0,0,0,0,
                                        @LF001, @LF002, @LF003, @LF004, @LF005,
                                        @LF006, @LF007, @LF008, @LF009, @LF010,
                                        @LF011, @LF012, @LF013)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        COMPANY = ErpNo,
                                        CREATOR = UserNo,
                                        USR_GROUP,
                                        CREATE_DATE = dateNow,
                                        MODIFIER = "",
                                        MODI_DATE = "",
                                        FLAG = 1,
                                        CREATE_TIME = timeNow,
                                        CREATE_AP = BaseHelper.ClientComputer(),
                                        CREATE_PRID = "BM",
                                        LF001 = GrErpPrefix,
                                        LF002 = GrErpNo,
                                        LF003 = GrSeq,
                                        LF004 = MtlItemNo,
                                        LF005 = InventoryNo,
                                        LF006 = Location.Length <= 0 ? "##########" : Location,
                                        LF007 = LotNumber.Length <= 0 ? "####################" : LotNumber,
                                        LF008 = 1,
                                        LF009 = ReceiveDateErp,
                                        LF010 = 1,
                                        LF011 = ConversionQty,
                                        LF012 = 0,
                                        LF013 = SupplierNo
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region //INVMF 品號批號資料單身
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO INVMF (COMPANY, CREATOR, USR_GROUP, 
                                        CREATE_DATE,MODIFIER, MODI_DATE, FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID,
                                        UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10,
                                        MF001, MF002, MF003, MF004, MF005, 
                                        MF006, MF007, MF008, MF009, MF010, 
                                        MF013, MF014,
                                        MF016, MF017, MF502, MF503, MF504)
                                 VALUES(@COMPANY, @CREATOR, @USR_GROUP,
                                        @CREATE_DATE,@MODIFIER,@MODI_DATE,@FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID,
                                        '','','','','',0,0,0,0,0,
                                        @MF001, @MF002, @MF003, @MF004, @MF005,
                                        @MF006, @MF007, @MF008, @MF009, @MF010,
                                        @MF013, @MF014,
                                        0, 0, 0, 0, 0)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        COMPANY = ErpNo,
                                        CREATOR = UserNo,
                                        USR_GROUP,
                                        CREATE_DATE = dateNow,
                                        MODIFIER = "",
                                        MODI_DATE = "",
                                        FLAG = 1,
                                        CREATE_TIME = timeNow,
                                        CREATE_AP = BaseHelper.ClientComputer(),
                                        CREATE_PRID = "BM",
                                        MF001 = MtlItemNo,
                                        MF002 = LotNumber,
                                        MF003 = ReceiveDateErp,
                                        MF004 = GrErpPrefix,
                                        MF005 = GrErpNo,
                                        MF006 = GrSeq,
                                        MF007 = InventoryNo,
                                        MF008 = 1,
                                        MF009 = "1",
                                        MF010 = ConversionQty,
                                        MF013 = "", //備註
                                        MF014 = 0 //包裝數量
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #endregion

                                #region //只新增一次
                                #region //INVME 品號批號資料單頭
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT ME001,ME007 FROM INVME
                                        WHERE ME001 = @ME001
                                          AND ME002 = @ME002";
                                dynamicParameters.Add("ME001", MtlItemNo);
                                dynamicParameters.Add("ME002", LotNumber);
                                var resultINVME = sqlConnection.Query(sql, dynamicParameters);
                                if (resultINVME.Count() <= 0)
                                {
                                    #region //INVME 新增
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO INVME (COMPANY, CREATOR, USR_GROUP, 
                                       CREATE_DATE,MODIFIER, MODI_DATE, FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID,
                                       UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10,
                                       ME001, ME002, ME003, ME004, ME005, 
                                       ME006, ME007, ME008, ME009, ME010,
                                       ME011, ME012, ME503, ME504, ME505 )
                                VALUES(@COMPANY, @CREATOR, @USR_GROUP,
                                       @CREATE_DATE,@MODIFIER,@MODI_DATE,@FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID,
                                       '','','','','',0,0,0,0,0,
                                       @ME001, @ME002, @ME003, @ME004, @ME005, 
                                       @ME006, @ME007, @ME008, @ME009, @ME010,
                                       0, 0, 0, 0, 0 )";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            COMPANY = ErpNo,
                                            CREATOR = UserNo,
                                            USR_GROUP,
                                            CREATE_DATE = dateNow,
                                            MODIFIER = "",
                                            MODI_DATE = "",
                                            FLAG = 1,
                                            CREATE_TIME = timeNow,
                                            CREATE_AP = BaseHelper.ClientComputer(),
                                            CREATE_PRID = "BM",
                                            ME001 = MtlItemNo,
                                            ME002 = LotNumber,
                                            ME003 = ReceiveDateErp,
                                            ME004 = "",
                                            ME005 = GrErpPrefix,
                                            ME006 = GrErpNo,
                                            ME007 = "N",
                                            ME008 = "",
                                            ME009 = AvailableDate,
                                            ME010 = ReCheckDate
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                }
                                else
                                {
                                    foreach (var item1 in resultINVME)
                                    {
                                        if (item1.ME007 != "N") throw new SystemException("品號:" + MtlItemNo + "批號:" + LotNumber + "已經結案了!!!");
                                    }
                                }
                                #endregion

                                #endregion

                                #region //第一次新增後續更新
                                #region //INVMM 品號庫別儲位批號檔
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT MM001 FROM INVMM
                                        WHERE MM001 = @MM001
                                          AND MM002 = @MM002
                                          AND MM003 = @MM003
                                          AND MM004 = @MM004";
                                dynamicParameters.Add("MM001", MtlItemNo);
                                dynamicParameters.Add("MM002", InventoryNo);
                                dynamicParameters.Add("MM003", Location != "" ? Location : "##########");
                                dynamicParameters.Add("MM004", LotNumber);
                                var resultINVMM = sqlConnection.Query(sql, dynamicParameters);

                                if (resultINVMM.Count() > 0)
                                {
                                    #region //後續更新流程
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE INVMM SET
                                                  MODIFIER = @MODIFIER,
                                                  MODI_DATE = @MODI_DATE,
                                                  MODI_TIME = @MODI_TIME,
                                                  MODI_AP = @MODI_AP,
                                                  MODI_PRID = @MODI_PRID,
                                                  FLAG = FLAG + 1,
                                                  MM005 = MM005 + @MM005,
                                                  MM008 = @MM008
                                            WHERE MM001 = @MM001
                                              AND MM002 = @MM002
                                              AND MM003 = @MM003
                                              AND MM004 = @MM004";
                                    dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        MODIFIER = UserNo,
                                        MODI_DATE = dateNow,
                                        MODI_TIME = timeNow,
                                        MODI_AP = BaseHelper.ClientComputer(),
                                        MODI_PRID = "BM",
                                        MM005 = ConversionQty,
                                        MM008 = ReceiveDateErp,
                                        MM001 = MtlItemNo,
                                        MM002 = InventoryNo,
                                        MM003 = Location != "" ? Location : "##########",
                                        MM004 = LotNumber,
                                    });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                                else
                                {
                                    #region //第一次新增流程
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO INVMM (COMPANY, CREATOR, USR_GROUP, 
                                            CREATE_DATE,MODIFIER, MODI_DATE, FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID,
                                            UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10,
                                            MM001, MM002, MM003, MM004, MM005, 
                                            MM006, MM007, MM008, MM009, MM010, 
                                            MM011, MM012, MM013, MM014)
                                    VALUES (@COMPANY, @CREATOR, @USR_GROUP,
                                            @CREATE_DATE,@MODIFIER,@MODI_DATE,@FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID,
                                            '','','','','',0,0,0,0,0,
                                            @MM001, @MM002, @MM003, @MM004, @MM005, 
                                            @MM006, @MM007, @MM008, @MM009, @MM010, 
                                            @MM011, @MM012, @MM013, @MM014)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            COMPANY = ErpNo,
                                            CREATOR = UserNo,
                                            USR_GROUP,
                                            CREATE_DATE = dateNow,
                                            MODIFIER = "",
                                            MODI_DATE = "",
                                            FLAG = 1,
                                            CREATE_TIME = timeNow,
                                            CREATE_AP = BaseHelper.ClientComputer(),
                                            CREATE_PRID = "BM",
                                            MM001 = MtlItemNo,
                                            MM002 = InventoryNo,
                                            MM003 = Location.Length > 0 ? Location : "##########",
                                            MM004 = LotNumber.Length > 0 ? LotNumber : "####################",
                                            MM005 = ConversionQty,
                                            MM006 = 0, //庫存包裝數量,
                                            MM007 = "",
                                            MM008 = ReceiveDateErp,
                                            MM009 = "",
                                            MM010 = "", //備註
                                            MM011 = 0,
                                            MM012 = "",
                                            MM013 = 0,
                                            MM014 = 0
                                        });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                                #endregion

                                #endregion
                            }
                            #endregion

                            #region //進貨單來源為暫入單觸發
                            if (TrErpPrefix != "")
                            {
                                //#region //INVTF 暫出入轉撥單頭檔更新 轉銷日更新
                                //dynamicParameters = new DynamicParameters();
                                //sql = @"UPDATE INVTF 
                                //        SET MODIFIER =@MODIFIER,
                                //       MODI_DATE= @MODI_DATE,
                                //       MODI_TIME= @MODI_TIME,
                                //       MODI_AP= @MODI_AP,
                                //       MODI_PRID= @MODI_PRID,
                                //       FLAG= FLAG + 1,
                                //        TF042  = @TF042
                                //        WHERE TF001 = @TF001
                                //        AND TF002 = @TF002";
                                //dynamicParameters.AddDynamicParams(
                                //new
                                //{
                                //    TF042 = ReceiveDateErp,
                                //    TF001 = TH032,
                                //    TF002 = TH033
                                //});
                                //rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                //#endregion

                                //#region //INVTG 暫出入轉撥單身檔 轉銷數量
                                //dynamicParameters = new DynamicParameters();
                                //sql = @"UPDATE INVTG SET
                                //        TG020  = TG020 + @TG020,
                                //        TG046  = TG046 + @TG046,
                                //        TG054  = TG054 + @TG054,
                                //        TG024  = @TG024
                                //        WHERE TG001 = @TG001
                                //        AND TG002 = @TG002
                                //        AND TG003 = @TG003";
                                //dynamicParameters.AddDynamicParams(
                                //new
                                //{
                                //    TG020 = Quantity,
                                //    TG046 = FreeSpareQty,
                                //    TG054 = AvailableQty,
                                //    TG001 = TH032,
                                //    TG002 = TH033,
                                //    TG003 = TH034,
                                //    TG024 = ClosureStatusINVTG
                                //});
                                //rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                //#endregion
                            }
                            #endregion

                            #endregion

                        }
                        #endregion
                    }
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //若成功則將相關資料寫回MES
                        DateTime date = DateTime.ParseExact(MaxReceiptDate, "yyyyMMdd", null);
                        MaxReceiptDate = date.ToString("yyyy-MM-dd"); // yyyy-MM-dd 格式

                        #region //更新 - 進貨單單單身
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.GrDetail SET
                                ConfirmStatus = @ConfirmStatus,
                                ConfirmUser = @ConfirmUser,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE GrDetailId = @GrDetailId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmStatus = "Y",
                                ConfirmUser = LastModifiedBy,
                                LastModifiedDate,
                                LastModifiedBy,
                                GrDetailId
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //找是否有未確認單身，若無則更新單頭
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT (a.GrErpPrefix + '-' + a.GrErpNo) GrErpFullNo , a.GrDetailId, a.ConfirmStatus detailConfirmStatus, a.GrSeq
                                , b.GrId, b.ConfirmStatus headerConfirm
                                FROM SCM.GrDetail a
                                INNER JOIN SCM.GoodsReceipt b ON a.GrId = b.GrId
                                WHERE a.GrId = @GrId 
                                and a.ConfirmStatus = 'N' 
                                and a.GrDetailId !=  @GrDetailId";
                        dynamicParameters.Add("GrDetailId", GrDetailId);
                        dynamicParameters.Add("GrId", grId);


                        var unConfirmresult = sqlConnection.Query(sql, dynamicParameters);
                        #endregion

                        #region //更新 - 進貨單單頭
                        if (unConfirmresult.Count() <= 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.GoodsReceipt SET
                                    ReceiptDate = @ReceiptDate,
                                    ConfirmStatus = @ConfirmStatus,
                                    TransferStatus = 'Y',
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE GrId = @GrId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    ReceiptDate = MaxReceiptDate,
                                    ConfirmStatus = "Y",
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    grId
                                });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        }

                        #endregion

                        #region /撈取進貨單單身資料=>更新採購單,暫入單
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.AcceptQty,a.AvailableQty,a.PoDetailId
                                FROM SCM.GrDetail a
                                WHERE a.GrDetailId = @GrDetailId";
                        dynamicParameters.Add("GrDetailId", GrDetailId);

                        var resultGrDetail = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in resultGrDetail)
                        {
                            #region //更新 - 採購單 已交數量,已交計價數量
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.PoDetail SET
                                    SiQty = SiQty + @SiQty,
                                    SiPriceQty = SiPriceQty + @SiPriceQty,
                                    ClosureStatus = @ClosureStatus,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE PoErpPrefix = @PoErpPrefix AND PoErpNo = @PoErpNo AND PoSeq = @PoSeq";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    SiQty = item.AcceptQty,
                                    SiPriceQty = item.AvailableQty,
                                    ClosureStatus = CloseStatus,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    PoErpPrefix,
                                    PoErpNo,
                                    PoSeq
                                });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //更新 - 暫入單 轉進銷量,轉銷贈/備品量
                            //if (item.TsnDetailId > 0)
                            //{
                            //    dynamicParameters = new DynamicParameters();
                            //    sql = @"UPDATE SCM.TsnDetail SET
                            //            SaleQty = SaleQty + @SaleQty,
                            //            SaleFSQty = SaleFSQty + @SaleFSQty,
                            //            LastModifiedDate = @LastModifiedDate,
                            //            LastModifiedBy = @LastModifiedBy
                            //            WHERE TsnDetailId = @TsnDetailId";
                            //    dynamicParameters.AddDynamicParams(
                            //        new
                            //        {
                            //            SaleQty = item.Quantity,
                            //            SaleFSQty = item.FreeSpareQty,
                            //            LastModifiedDate,
                            //            LastModifiedBy,
                            //            TsnDetailId = item.TsnDetailId
                            //        });
                            //    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            //}
                            #endregion

                        }
                        #endregion
                        #endregion

                        #region //批號相關異動
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.GrSeq,a.LotNumber,a.MtlItemId,b.LotManagement,a.InventoryId,a.AcceptQty
                                FROM SCM.GrDetail a
                                INNER JOIN PDM.MtlItem b on a.MtlItemId = b.MtlItemId
                                WHERE a.GrDetailId = @GrDetailId";
                        dynamicParameters.Add("GrDetailId", GrDetailId);

                        var resultRoMes = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in resultRoMes)
                        {
                            //string GrSeq = item.GrSeq;
                            int MtlItemId = item.MtlItemId;
                            string LotNumber = item.LotNumber;
                            string LotManagement = item.LotManagement;
                            int InventoryId = item.InventoryId;
                            double AcceptQty = Convert.ToDouble(item.AcceptQty);
                            if (LotManagement != "N")
                            {
                                #region //確認品號資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.MtlItemNo, a.LotManagement
                                    FROM PDM.MtlItem a 
                                    WHERE a.MtlItemId = @MtlItemId";
                                dynamicParameters.Add("MtlItemId", MtlItemId);

                                var MtlItemResult = sqlConnection.Query(sql, dynamicParameters);

                                if (MtlItemResult.Count() <= 0) throw new SystemException("品號資料錯誤!!");

                                string MtlItemNo = "";
                                foreach (var item1 in MtlItemResult)
                                {
                                    if (item1.LotManagement != "T") throw new SystemException("品號批號管控參數錯誤!!");
                                    MtlItemNo = item.MtlItemNo;
                                }
                                #endregion

                                #region //確認此品號是否已存在此批號
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 LotNumberId
                                    FROM SCM.LotNumber a 
                                    WHERE a.MtlItemId = @MtlItemId
                                    AND a.LotNumberNo = @LotNumberNo";
                                dynamicParameters.Add("MtlItemId", MtlItemId);
                                dynamicParameters.Add("LotNumberNo", LotNumber);

                                var LotNumberResult = sqlConnection.Query(sql, dynamicParameters);

                                if (LotNumberResult.Count() <= 0)
                                {
                                    #region //INSERT SCM.GoodsReceipt
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO SCM.LotNumber (MtlItemId, LotNumberNo, Remark
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.LotNumberId
                                    VALUES (@MtlItemId, @LotNumberNo, @Remark
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            MtlItemId,
                                            LotNumberNo = LotNumber,
                                            Remark = "",
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });
                                    var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult.Count();
                                    #endregion

                                    int LotNumberId = -1;
                                    foreach (var item1 in insertResult)
                                    {
                                        LotNumberId = item1.LotNumberId;
                                    }

                                    #region //INSERT SCM.LnDetail
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO SCM.LnDetail (LotNumberId, TransactionDate, FromErpPrefix,FromErpNo,FromSeq
                                    ,InventoryId,TransactionType,DocType,Quantity,Remark
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@LotNumberId, @TransactionDate, @FromErpPrefix,@FromErpNo,@FromSeq
                                    ,@InventoryId,@TransactionType,@DocType,@Quantity,@Remark
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            LotNumberId,
                                            TransactionDate = ReceiptDate,
                                            FromErpPrefix = GrErpPrefix,
                                            FromErpNo = GrErpNo,
                                            FromSeq = GrSeq,
                                            InventoryId,
                                            TransactionType = 1,
                                            DocType = 1,
                                            Quantity = AcceptQty,
                                            Remark = "",
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });
                                    insertResult = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult.Count();
                                    #endregion
                                }
                                else
                                {
                                    int LotNumberId = -1;
                                    foreach (var item1 in LotNumberResult)
                                    {
                                        LotNumberId = item1.LotNumberId;
                                    }

                                    #region //INSERT SCM.LnDetail
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO SCM.LnDetail (LotNumberId, TransactionDate, FromErpPrefix,FromErpNo,FromSeq
                                    ,InventoryId,TransactionType,DocType,Quantity,Remark
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@LotNumberId, @TransactionDate, @FromErpPrefix,@FromErpNo,@FromSeq
                                    ,@InventoryId,@TransactionType,@DocType,@Quantity,@Remark
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            LotNumberId,
                                            TransactionDate = ReceiptDate,
                                            FromErpPrefix = GrErpPrefix,
                                            FromErpNo = GrErpNo,
                                            FromSeq = GrSeq,
                                            InventoryId,
                                            TransactionType = 1,
                                            DocType = 1,
                                            Quantity = AcceptQty,
                                            Remark = "",
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });
                                    var insertResult1 = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult1.Count();
                                    #endregion
                                }
                                #endregion
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

        #region //UpdateGrDetailReConfirm -- 反確認進貨單身 -- GPAI 2025-02-23 
        public string UpdateGrDetailReConfirm(int GrDetailId)
        {
            try
            {
                string dateNow = DateTime.Now.ToString("yyyyMMdd");
                string timeNow = DateTime.Now.ToString("HH:mm:ss");
                string dateNowMonth = DateTime.Now.ToString("yyyyMM");

                string UserNo = "";

                int rowsAffected = 0;
                var ErpNo = "";
                string ErpDbName = "";
                string USR_GROUP = "";
                int UserId = CreateBy;

                string GrErpPrefix = "";
                string GrErpNo = "";
                string ReceiptDate = "";
                int CompanyIdBase = -1;
                string GrSeq = "";

                int grId = -1;

                string PoErpPrefix = ""; //TH011 採購單別 單號 序號
                string PoErpNo = ""; //TH012 
                string PoSeq =""; //TH013 

                using (TransactionScope transactionScope = new TransactionScope())
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

                        var resultCompany = sqlConnection.Query(sql, dynamicParameters);

                        if (resultCompany.Count() <= 0) throw new SystemException("公司別錯誤!");

                        foreach (var item in resultCompany)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            ErpDbName = ErpConnectionStrings.Split('=')[2].Split(';')[0];
                            ErpNo = item.ErpNo;
                        }
                        #endregion

                        #region //撈取使用者No
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
                        dynamicParameters.Add("FunctionCode", "GoodsReceiptManagement");
                        dynamicParameters.Add("DetailCode", "confirm");

                        var resultUser = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUser.Count() <= 0) throw new SystemException("使用者資料錯誤!");

                        foreach (var item in resultUser)
                        {
                            UserNo = item.UserNo;
                        }
                        #endregion

                        #region /撈取單頭資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.GrErpPrefix,a.GrErpNo, a.ConfirmStatus,a.CompanyId
                                ,FORMAT(a.ReceiptDate, 'yyyyMMdd') ReceiptDate
                                ,b.GrDetailId,b.CloseStatus, b.GrSeq, a.GrId
                                FROM SCM.GoodsReceipt a
                                INNER JOIN SCM.GrDetail b on a.GrId = b.GrId
                                WHERE b.GrDetailId = @GrDetailId";
                        dynamicParameters.Add("GrDetailId", GrDetailId);

                        var resultRoMes = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in resultRoMes)
                        {
                            if (item.ConfirmStatus == "V") throw new SystemException("該進貨單單據已經作廢,不能操作");
                            if (item.CloseStatus == "Y") throw new SystemException("該進貨單已經結帳,不能操作");
                            //if (item.ConfirmStatus != "Y") throw new SystemException("該進貨單單據須處於核單狀態,才能操作");
                            if (item.CompanyId != CurrentCompany) throw new SystemException("該進貨單單據公司別與後端公司別不一致,不能操作");
                            CompanyIdBase = item.CompanyId;
                            GrErpPrefix = item.GrErpPrefix;
                            GrErpNo = item.GrErpNo;
                            ReceiptDate = item.ReceiptDate;
                            GrSeq = item.GrSeq;
                            //GrDetailId = item.GrDetailId;
                            grId = item.GrId;
                        }
                        #endregion
                    }
                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //審核ERP權限
                        USR_GROUP = BaseHelper.CheckErpAuthority(UserNo, sqlConnection, "PURI13", "CONFIRM");
                        #endregion

                        #region //判斷開立日期是否超過結帳日
                        string CloseDateBase = "";
                        string ReceiveDateBase = ReceiptDate;
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT MA011,MA012,MA013
                                FROM CMSMA ";
                        var resultCloseDate = sqlConnection.Query(sql, dynamicParameters);
                        foreach (var item in resultCloseDate)
                        {
                            CloseDateBase = item.MA013;
                        }

                        DateTime docDateBase;
                        DateTime closeDateBase;

                        if (DateTime.TryParseExact(ReceiveDateBase, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out docDateBase) &&
                            DateTime.TryParseExact(CloseDateBase, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out closeDateBase))
                        {
                            int resultDate = DateTime.Compare(docDateBase, closeDateBase);

                            if (resultDate <= 0)
                            {
                                throw new SystemException("【進貨單】ERP已經結帳,不可以在異動【" + CloseDateBase + "】之前的單據");
                            }
                        }
                        else
                        {
                            throw new SystemException("日期字符串无效，无法比较");
                        }
                        #endregion

                        #region //取出進貨單單據資料 + 異動更新相關資料表
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.TG003,a.TG005,a.TG007,a.TG013
                                ,b.TH003,b.TH004,b.TH007,b.TH009,b.TH010
                                ,b.TH011,b.TH012,b.TH013
                                ,b.TH014,b.TH015,b.TH016,b.TH017,b.TH030,b.TH031,b.TH063
                                ,b.TH021,b.TH022,b.TH023
                                ,b.TH036,b.TH037,b.TH047
                                ,c.MB022
                                ,d.MQ019
                                ,ISNULL(md1.GrRate,1) GrRate
                                ,ISNULL(md2.GrAvailableRate,1) GrAvailableRate
                                ,(b.TH015*ISNULL(md1.GrRate,1)) ConversionQty
                                ,(b.TH016*ISNULL(md2.GrAvailableRate,1)) ConversionAvailableQty
                                ,CASE 
                                    WHEN b.TH016 != 0 THEN ROUND((b.TH047 / (b.TH016*ISNULL(md2.GrAvailableRate,1))),CAST(md3.MF005 AS INT)) 
                                    ELSE 0
                                END AS UnitCost
                                FROM PURTG a
                                INNER JOIN PURTH b on a.TG001 = b.TH001 AND a.TG002 = b.TH002
                                INNER JOIN INVMB c on b.TH004 = c.MB001
                                INNER JOIN CMSMQ d ON d.MQ001=a.TG001 AND d.MQ003='34'
                                INNER JOIN CMSMC e ON e.MC001=b.TH009
                                LEFT JOIN INVMD f ON f.MD001=b.TH004
                                OUTER APPLY(
                                    SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) GrRate
                                    FROM INVMD x1
                                    WHERE x1.MD001=b.TH004
                                    AND  x1.MD002 = b.TH008
                                ) md1
                                OUTER APPLY(
                                    SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) GrAvailableRate
                                    FROM INVMD x1
                                    WHERE x1.MD001=b.TH004
                                    AND  x1.MD002 = b.TH056
                                ) md2
                                OUTER APPLY(
                                    SELECT x1.MF003,x1.MF004,x1.MF005,x1.MF006
                                    FROM CMSMF x1
                                    INNER JOIN CMSMA x2 on x1.MF001 = x2.MA003 
                                ) md3
                                WHERE b.TH001 = @GrErpPrefix
                                AND b.TH002 = @GrErpNo
                                AND b.TH003 = @GrSeq";
                        dynamicParameters.Add("GrErpPrefix", GrErpPrefix);
                        dynamicParameters.Add("GrErpNo", GrErpNo);
                        dynamicParameters.Add("GrSeq", GrSeq);

                        var resultRoErp = sqlConnection.Query(sql, dynamicParameters);
                        if (resultRoErp.Count() <= 0) throw new SystemException("找不到ERP進貨單單據!");
                        foreach (var item in resultRoErp)
                        {
                            if (item.TG013 == "V") throw new SystemException("該進貨單單據已經作廢,不能操作");
                           // if (item.TG013 != "Y") throw new SystemException("該進貨單單據須處於核單,才能操作");
                            if (item.TH030 != "Y") throw new SystemException("該進貨單單據單身須處於核單,才能操作");
                            if (item.TH021 != "") throw new SystemException("來源暫入單模式尚未開發,不能操作");
                            if (item.MQ019 != "Y") throw new SystemException("不核對採購單模式尚未開發,不能操作");

                            #region //欄位參數撈取
                            string ReceiveDateErp = item.TG003; //TG003 進貨日期
                            string SupplierNo = item.TG005; //TG005 供應商
                            string CurrencyCode = item.TG007; //TG007 幣別
                            Double Exchange = Convert.ToDouble(item.TG008); //TG008 匯率
                            string ConfirmStatus = item.TG013; //TG013 單頭確認

                           // string GrSeq = item.TH003; //TH003 進貨單序號
                            string GrFull = GrErpPrefix + '-' + GrErpNo + '-' + GrSeq;
                            string MtlItemNo = item.TH004; //TH004 品號
                            string InventoryNo = item.TH009; //TH009 庫別
                            string LotNumber = item.TH010; //TH010 批號
                            PoErpPrefix = item.TH011; //TH011 採購單別 單號 序號
                            PoErpNo = item.TH012; //TH012 
                            PoSeq = item.TH013; //TH013 
                            string PoFull = item.TH011 + '-' + item.TH012 + '-' + item.TH013;
                            string TH014 = item.TH014; //TH014 驗收日期
                            Double ReceiptQty = Convert.ToDouble(item.TH007); //TH014 進貨數量
                            Double AcceptQty = Convert.ToDouble(item.TH015); //TH014 驗收數量
                            Double AvailableQty = Convert.ToDouble(item.TH016); //TH014 計價數量 
                            Double ReturnQty = Convert.ToDouble(item.TH017); //TH014 驗退數量
                            string TH030 = item.TH030; //TH030 確認碼
                            string CloseStatus = item.TH031; //TH031 結帳碼
                            string AvailableDate = item.TH036; //TH036 有效日期
                            string ReCheckDate = item.TH037; //TH037 複檢日期
                            string Location = item.TH063; //TH060 儲位
                            string TrErpPrefix = item.TH021; //TH021 暫入單別 / TH022 暫入單號 / TH023 暫入序號
                            string TrErpNo = item.TH022;
                            string TrSeq = item.TH023;
                            string TsnFull = item.TH021 + '-' + item.TH022 + '-' + item.TH023;
                            string LotManagement = item.MB022; //品號批號管理

                            Double GrRate = Convert.ToDouble(item.GrRate); //單位換算率
                            Double ConversionQty = Convert.ToDouble(item.ConversionQty); //單位換算後數量
                            Double ConversionAvailableQty = Convert.ToDouble(item.ConversionAvailableQty); //單位換算後計價數量
                            Double UnitCost = Convert.ToDouble(item.UnitCost); //單位成本
                            Double TH047 = Convert.ToDouble(item.TH047); //本幣未稅金額
                            Double TH018 = Convert.ToDouble(item.TH018); //原幣單價
                            Double PoRate = 0; //採購單位換算率
                            Double PoAvailableRate = 0; //採購計價單位換算率
                            Double PoConversionQty = 0; //採購單位換算後數量
                            Double PoConversionAvailableQty = 0; //採購單位換算後計價數量
                            #endregion

                            #region //檢查段

                            #region //數量查驗
                            if (ReceiptQty < 0) throw new SystemException("【進貨單:" + GrFull + "】進貨數量不可小於0!!!");
                            if (AcceptQty < 0) throw new SystemException("【進貨單:" + GrFull + "】驗收數量不可小於0!!!");
                            if (AvailableQty < 0) throw new SystemException("【進貨單:" + GrFull + "】計價數量不可小於0!!!");
                            if (ReceiptQty < AcceptQty) throw new SystemException("【進貨單:" + GrFull + "】驗收數量不可大於進貨數量!!!");
                            if (AcceptQty < AvailableQty) throw new SystemException("【進貨單:" + GrFull + "】計價數量不可大於驗收數量!!!");


                            #region //判斷該品號批號庫存是否足夠
                            if (LotManagement != "N")
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT MM005
                                    FROM INVMM a
                                    WHERE a.MM001 = @MM001
                                    AND a.MM002 = @MM002
                                    AND a.MM003 = @MM003
                                    AND a.MM004 = @MM004
                                    ";
                                dynamicParameters.Add("MM001", MtlItemNo);
                                dynamicParameters.Add("MM002", InventoryNo);
                                dynamicParameters.Add("MM003", Location != "" ? Location : "##########");
                                dynamicParameters.Add("MM004", LotNumber);
                                var INVMMResult = sqlConnection.Query(sql, dynamicParameters);
                                if (INVMMResult.Count() > 0)
                                {
                                    foreach (var item1 in INVMMResult)
                                    {
                                        if (Convert.ToInt32(item1.MM005) < ConversionQty) throw new SystemException("批號庫存不足不能執行反確認,請重新確認");
                                    }
                                }
                                else
                                {
                                    throw new SystemException("批號庫存不足不能執行反確認,請重新確認");
                                }
                            }
                            #endregion

                            #region //判斷該品號庫存是否足夠
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT MC007
                                    FROM INVMC a
                                    WHERE a.MC001 = @MC001
                                    AND a.MC002 = @MC002
                                    ";
                            dynamicParameters.Add("MC001", MtlItemNo);
                            dynamicParameters.Add("MC002", InventoryNo);
                            var INVMCResult = sqlConnection.Query(sql, dynamicParameters);
                            if (INVMCResult.Count() > 0)
                            {
                                foreach (var item1 in INVMCResult)
                                {
                                    if (Convert.ToInt32(item1.MC007) < ConversionQty) throw new SystemException("庫存不足不能執行反確認,請重新確認");
                                }
                            }
                            else
                            {
                                throw new SystemException("庫存不足不能執行反確認,請重新確認");
                            }
                            #endregion


                            #endregion

                            #region //PURTD 採購單
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 ISNULL(a.TD008,0) TD008,ISNULL(a.TD015,0) TD015,a.TD016,a.TD018,
                                    ISNULL(md1.PoRate,1) PoRate, 
                                    ISNULL(md2.PoAvailableRate,1) PoAvailableRate, 
                                    (a.TD008*ISNULL(md1.PoRate,1)) PoConversionQty,
                                    (a.TD060*ISNULL(md2.PoAvailableRate,1)) PoConversionAvailableQty
                                    FROM PURTD a
                                    OUTER APPLY(
                                        SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) PoRate
                                        FROM INVMD x1
                                        WHERE x1.MD001 = a.TD004
                                        AND  x1.MD002 = a.TD009
                                    ) md1
                                    OUTER APPLY(
                                        SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) PoAvailableRate
                                        FROM INVMD x1
                                        WHERE x1.MD001 = a.TD004
                                        AND  x1.MD002 = a.TD059
                                    ) md2
                                    WHERE a.TD001 = @TD001
                                    AND a.TD002 = @TD002
                                    AND a.TD003 = @TD003";
                            dynamicParameters.Add("TD001", PoErpPrefix);
                            dynamicParameters.Add("TD002", PoErpNo);
                            dynamicParameters.Add("TD003", PoSeq);
                            var resultCOPTD = sqlConnection.Query(sql, dynamicParameters);
                            if (resultCOPTD.Count() <= 0) throw new SystemException("採購單找不到,請重新確認");
                            foreach (var item1 in resultCOPTD)
                            {
                                if (item1.TD018 != "Y") throw new SystemException("【採購單:" + PoFull + "】非確認狀態,不可操作!!!");
                                if (Convert.ToDouble(item1.TD008) == Convert.ToDouble(item1.TD015) - AcceptQty)
                                {
                                    CloseStatus = "Y";
                                }
                                else if (Convert.ToDouble(item1.TD008) < Convert.ToDouble(item1.TD015) - AcceptQty)
                                {
                                    throw new SystemException("【採購單:" + PoFull + "】已超過採購數量,目前已交數量:" + Convert.ToDouble(item1.TD015) + ",本次反確認進貨驗收數量:" + AcceptQty + ",不可操作!!!");
                                }
                                PoRate = Convert.ToDouble(item1.PoRate);
                                PoAvailableRate = Convert.ToDouble(item1.PoAvailableRate);
                                PoConversionQty = Convert.ToDouble(item1.PoConversionQty);
                                PoConversionAvailableQty = Convert.ToDouble(item1.PoConversionAvailableQty);
                            }
                            #endregion
                            #endregion

                            #region //反確認段
                            #region //固定更新

                            #region //PURTD 採購單單身資料檔
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE PURTD SET
                                    MODIFIER =@MODIFIER,
                                    MODI_DATE= @MODI_DATE,
                                    MODI_TIME= @MODI_TIME,
                                    MODI_AP= @MODI_AP,
                                    MODI_PRID= @MODI_PRID,
                                    FLAG= FLAG + 1,
                                    TD015 = TD015 - @TD015,
                                    TD016 = @TD016,
                                    TD060 = TD060 - @TD060
                                    WHERE TD001 = @TD001
                                    AND TD002 = @TD002
                                    AND TD003 = @TD003";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                MODIFIER = UserNo,
                                MODI_DATE = dateNow,
                                MODI_TIME = timeNow,
                                MODI_AP = BaseHelper.ClientComputer(),
                                MODI_PRID = "BM",
                                TD015 = Math.Round(ConversionQty / PoRate, 3, MidpointRounding.AwayFromZero),
                                TD016 = CloseStatus,
                                TD060 = Math.Round(ConversionAvailableQty / PoAvailableRate, 3, MidpointRounding.AwayFromZero),
                                TD001 = PoErpPrefix,
                                TD002 = PoErpNo,
                                TD003 = PoSeq
                            });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //PURTG 進貨單單頭檔***
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE PURTG SET
                                    MODIFIER =@MODIFIER,
                                    MODI_DATE= @MODI_DATE,
                                    MODI_TIME= @MODI_TIME,
                                    MODI_AP= @MODI_AP,
                                    MODI_PRID= @MODI_PRID,
                                    FLAG= FLAG + 1,
                                    TG013 = 'N'
                                    WHERE TG001 = @GrErpPrefix
                                    AND TG002 = @GrErpNo";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                MODIFIER = UserNo,
                                MODI_DATE = dateNow,
                                MODI_TIME = timeNow,
                                MODI_AP = BaseHelper.ClientComputer(),
                                MODI_PRID = "BM",
                                GrErpPrefix,
                                GrErpNo
                            });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);


                            #endregion

                            #region //PURTH 進貨單單身檔
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE PURTH SET
                                    MODIFIER =@MODIFIER,
                                    MODI_DATE= @MODI_DATE,
                                    MODI_TIME= @MODI_TIME,
                                    MODI_AP= @MODI_AP,
                                    MODI_PRID= @MODI_PRID,
                                    FLAG= FLAG + 1,
                                    TH030 = 'N',
                                    TH031 = 'N',
                                    TH038 = @UserNo
                                    WHERE TH001 = @GrErpPrefix
                                    AND TH002 = @GrErpNo
                                    AND TH003 = @GrSeq";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                MODIFIER = UserNo,
                                MODI_DATE = dateNow,
                                MODI_TIME = timeNow,
                                MODI_AP = BaseHelper.ClientComputer(),
                                MODI_PRID = "BM",
                                GrErpPrefix,
                                GrErpNo,
                                GrSeq,
                                UserNo
                            });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                            #endregion

                            #region //INVMB 品號基本資料檔 庫存數量 庫存金額
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE INVMB SET
                                    MODIFIER =@MODIFIER,
                                    MODI_DATE= @MODI_DATE,
                                    MODI_TIME= @MODI_TIME,
                                    MODI_AP= @MODI_AP,
                                    MODI_PRID= @MODI_PRID,
                                    FLAG= FLAG + 1,
                                    MB064  = MB064 - @MB064,
                                    MB065  = MB065 - @MB065
                                    WHERE MB001 = @MB001";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                MODIFIER = UserNo,
                                MODI_DATE = dateNow,
                                MODI_TIME = timeNow,
                                MODI_AP = BaseHelper.ClientComputer(),
                                MODI_PRID = "BM",
                                MB064 = ConversionQty,
                                MB065 = TH047,
                                MB001 = MtlItemNo
                            });
                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //INVMC 品號庫別檔
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE INVMC SET
                                    MODIFIER =@MODIFIER,
                                    MODI_DATE= @MODI_DATE,
                                    MODI_TIME= @MODI_TIME,
                                    MODI_AP= @MODI_AP,
                                    MODI_PRID= @MODI_PRID,
                                    FLAG= FLAG + 1,
                                    MC007  = MC007 - @MC007,
                                    MC008  = MC008 - @MC008
                                    WHERE MC001 = @MC001
                                    AND MC002 = @MC002";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                MODIFIER = UserNo,
                                MODI_DATE = dateNow,
                                MODI_TIME = timeNow,
                                MODI_AP = BaseHelper.ClientComputer(),
                                MODI_PRID = "BM",
                                MC007 = ConversionQty,
                                MC008 = TH047,
                                MC013 = ReceiveDateErp,
                                MC001 = MtlItemNo,
                                MC002 = InventoryNo
                            });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #endregion

                            #region //固定刪除
                            #region //INVLA 異動明細資料檔
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE INVLA
                                    WHERE LA001 = @LA001
                                      AND LA004 = @LA004
                                      AND LA005 = @LA005
                                      AND LA006 = @LA006
                                      AND LA007 = @LA007
                                      AND LA008 = @LA008";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                LA001 = MtlItemNo,
                                LA004 = ReceiveDateErp,
                                LA005 = 1,
                                LA006 = GrErpPrefix,
                                LA007 = GrErpNo,
                                LA008 = GrSeq
                            });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                            #endregion

                            #region //品號有批號管理觸發
                            if (LotManagement != "N")
                            {
                                #region //固定刪除

                                #region //INVMF 品號批號資料單身
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE INVMF
                                         WHERE MF001 = @MF001
                                           AND MF002 = @MF002
                                           AND MF003 = @MF003
                                           AND MF004 = @MF004
                                           AND MF005 = @MF005
                                           AND MF006 = @MF006";
                                dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MF001 = MtlItemNo,
                                    MF002 = LotNumber,
                                    MF003 = ReceiveDateErp,
                                    MF004 = GrErpPrefix,
                                    MF005 = GrErpNo,
                                    MF006 = GrSeq
                                });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region //INVLF 儲位批號異動明細資料檔
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE INVLF
                                        WHERE LF001 = @LF001
                                          AND LF002 = @LF002
                                          AND LF003 = @LF003
                                          AND LF004 = @LF004
                                          AND LF005 = @LF005
                                          AND LF006 = @LF006
                                          AND LF007 = @LF007
                                          AND LF008 = @LF008";
                                dynamicParameters.AddDynamicParams(
                                new
                                {
                                    LF001 = GrErpPrefix,
                                    LF002 = GrErpNo,
                                    LF003 = GrSeq,
                                    LF004 = MtlItemNo,
                                    LF005 = InventoryNo,
                                    LF006 = Location != "" ? Location : "##########",
                                    LF007 = LotNumber,
                                    LF008 = 1
                                });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #endregion

                                #region //固定更新
                                #region //INVMM 品號庫別儲位批號檔
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE INVMM SET
                                             MM005 = MM005 - @MM005
                                       WHERE MM001 = @MM001
                                         AND MM002 = @MM002
                                         AND MM003 = @MM003
                                         AND MM004 = @MM004";
                                dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MM005 = ConversionQty,
                                    MM001 = MtlItemNo,
                                    MM002 = InventoryNo,
                                    MM003 = Location != "" ? Location : "##########",
                                    MM004 = LotNumber
                                });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                                #endregion

                                #region //無數量就刪除
                                #region //INVME 品號批號資料單頭
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1 FROM INVMF
                                        WHERE MF001 = @MF001
                                          AND MF002 = @MF002";
                                dynamicParameters.Add("MF001", MtlItemNo);
                                dynamicParameters.Add("MF002", LotNumber);
                                var resultINVME = sqlConnection.Query(sql, dynamicParameters);
                                if (resultINVME.Count() <= 0)
                                {
                                    #region //INVME 刪除
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE INVME
                                         WHERE ME001 = @ME001
                                           AND ME002 = @ME002";
                                    dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        ME001 = MtlItemNo,
                                        ME002 = LotNumber
                                    });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                                #endregion

                                #endregion
                            }
                            #endregion

                            #region //進貨單來源為暫入單觸發
                            if (TrErpPrefix != "")
                            {
                                //#region //INVTF 暫出入轉撥單頭檔更新 轉銷日更新
                                //dynamicParameters = new DynamicParameters();
                                //sql = @"UPDATE INVTF 
                                //        SET MODIFIER =@MODIFIER,
                                //       MODI_DATE= @MODI_DATE,
                                //       MODI_TIME= @MODI_TIME,
                                //       MODI_AP= @MODI_AP,
                                //       MODI_PRID= @MODI_PRID,
                                //       FLAG= FLAG + 1,
                                //        TF042  = @TF042
                                //        WHERE TF001 = @TF001
                                //        AND TF002 = @TF002";
                                //dynamicParameters.AddDynamicParams(
                                //new
                                //{
                                //    TF042 = ReceiveDateErp,
                                //    TF001 = TH032,
                                //    TF002 = TH033
                                //});
                                //rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                //#endregion

                                //#region //INVTG 暫出入轉撥單身檔 轉銷數量
                                //dynamicParameters = new DynamicParameters();
                                //sql = @"UPDATE INVTG SET
                                //        TG020  = TG020 + @TG020,
                                //        TG046  = TG046 + @TG046,
                                //        TG054  = TG054 + @TG054,
                                //        TG024  = @TG024
                                //        WHERE TG001 = @TG001
                                //        AND TG002 = @TG002
                                //        AND TG003 = @TG003";
                                //dynamicParameters.AddDynamicParams(
                                //new
                                //{
                                //    TG020 = Quantity,
                                //    TG046 = FreeSpareQty,
                                //    TG054 = AvailableQty,
                                //    TG001 = TH032,
                                //    TG002 = TH033,
                                //    TG003 = TH034,
                                //    TG024 = ClosureStatusINVTG
                                //});
                                //rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                //#endregion
                            }
                            #endregion
                            #endregion
                        }
                        #endregion
                    }
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //若成功則將相關資料寫回MES
                        #region //更新 - 進貨單單頭
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.GoodsReceipt SET
                                ConfirmStatus = @ConfirmStatus,
                                TransferStatus = @TransferStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE GrId = @GrId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmStatus = "N",
                                TransferStatus = "R",
                                ConfirmUserId = LastModifiedBy,
                                LastModifiedDate,
                                LastModifiedBy,
                                grId
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //更新 - 進貨單單單身
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.GrDetail SET
                                ConfirmStatus = @ConfirmStatus,
                                ConfirmUser = @ConfirmUser,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE GrDetailId = @GrDetailId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmStatus = "N",
                                ConfirmUser = LastModifiedBy,
                                LastModifiedDate,
                                LastModifiedBy,
                                GrDetailId
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region /撈取進貨單單身資料=>更新採購單,暫入單
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.AcceptQty,a.AvailableQty,a.PoDetailId
                                FROM SCM.GrDetail a
                                WHERE a.GrDetailId = @GrDetailId";
                        dynamicParameters.Add("GrDetailId", GrDetailId);

                        var resultGrDetail = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in resultGrDetail)
                        {
                            #region //更新 - 採購單 已交數量,已交計價數量
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.PoDetail SET
                                    SiQty = SiQty - @SiQty,
                                    SiPriceQty = SiPriceQty - @SiPriceQty,
                                    ClosureStatus = @ClosureStatus,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE PoErpPrefix = @PoErpPrefix AND PoErpNo = @PoErpNo AND PoSeq = @PoSeq";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    SiQty = item.AcceptQty,
                                    SiPriceQty = item.AvailableQty,
                                    ClosureStatus = "N",
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    PoErpPrefix ,
                                    PoErpNo,
                                    PoSeq
                                });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //更新 - 暫入單 轉進銷量,轉銷贈/備品量
                            //if (item.TsnDetailId > 0)
                            //{
                            //    dynamicParameters = new DynamicParameters();
                            //    sql = @"UPDATE SCM.TsnDetail SET
                            //            SaleQty = SaleQty + @SaleQty,
                            //            SaleFSQty = SaleFSQty + @SaleFSQty,
                            //            LastModifiedDate = @LastModifiedDate,
                            //            LastModifiedBy = @LastModifiedBy
                            //            WHERE TsnDetailId = @TsnDetailId";
                            //    dynamicParameters.AddDynamicParams(
                            //        new
                            //        {
                            //            SaleQty = item.Quantity,
                            //            SaleFSQty = item.FreeSpareQty,
                            //            LastModifiedDate,
                            //            LastModifiedBy,
                            //            TsnDetailId = item.TsnDetailId
                            //        });
                            //    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            //}
                            #endregion

                        }
                        #endregion
                        #endregion

                        #region //批號相關異動
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.GrSeq, a.LotNumber,a.MtlItemId,b.LotManagement,a.InventoryId,a.AcceptQty
                                FROM SCM.GrDetail a
                                INNER JOIN PDM.MtlItem b on a.MtlItemId = b.MtlItemId
                                WHERE a.GrDetailId = @GrDetailId";
                        dynamicParameters.Add("GrDetailId", GrDetailId);

                        var resultRoMes = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in resultRoMes)
                        {
                            //string GrSeq = item.GrSeq;
                            int MtlItemId = item.MtlItemId;
                            string LotNumber = item.LotNumber;
                            string LotManagement = item.LotManagement;
                            int InventoryId = item.InventoryId;
                            double AcceptQty = Convert.ToDouble(item.AcceptQty);
                            if (LotManagement != "N")
                            {
                                #region //確認品號資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.MtlItemNo, a.LotManagement
                                    FROM PDM.MtlItem a 
                                    WHERE a.MtlItemId = @MtlItemId";
                                dynamicParameters.Add("MtlItemId", MtlItemId);

                                var MtlItemResult = sqlConnection.Query(sql, dynamicParameters);

                                if (MtlItemResult.Count() <= 0) throw new SystemException("品號資料錯誤!!");

                                string MtlItemNo = "";
                                foreach (var item1 in MtlItemResult)
                                {
                                    if (item.LotManagement != "T") throw new SystemException("品號批號管控參數錯誤!!");
                                    MtlItemNo = item.MtlItemNo;
                                }
                                #endregion

                                #region //確認此品號是否已存在此批號
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 LotNumberId
                                    FROM SCM.LotNumber a 
                                    WHERE a.MtlItemId = @MtlItemId
                                    AND a.LotNumberNo = @LotNumberNo";
                                dynamicParameters.Add("MtlItemId", MtlItemId);
                                dynamicParameters.Add("LotNumberNo", LotNumber);

                                var LotNumberResult = sqlConnection.Query(sql, dynamicParameters);

                                if (LotNumberResult.Count() > 0)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 LotNumberId
                                            FROM SCM.LnDetail a 
                                            WHERE 1=1
                                            AND a.TransactionDate = @TransactionDate
                                            AND a.FromErpPrefix = @FromErpPrefix
                                            AND a.FromErpNo = @FromErpNo
                                            AND a.FromSeq = @FromSeq
                                            AND a.InventoryId = @InventoryId
                                            AND a.TransactionType = @TransactionType
                                            AND a.DocType = @DocType";
                                    dynamicParameters.Add("TransactionDate", ReceiptDate);
                                    dynamicParameters.Add("FromErpPrefix", GrErpPrefix);
                                    dynamicParameters.Add("FromErpNo", GrErpNo);
                                    dynamicParameters.Add("FromSeq", GrSeq);
                                    dynamicParameters.Add("InventoryId", InventoryId);
                                    dynamicParameters.Add("TransactionType", 1);
                                    dynamicParameters.Add("DocType", 1);

                                    var LnDetailResult = sqlConnection.Query(sql, dynamicParameters);
                                    if (LotNumberResult.Count() > 0)
                                    {
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"DELETE SCM.LnDetail
                                                WHERE 1=1
                                                AND TransactionDate = @TransactionDate
                                                AND FromErpPrefix = @FromErpPrefix
                                                AND FromErpNo = @FromErpNo
                                                AND FromSeq = @FromSeq
                                                AND InventoryId = @InventoryId
                                                AND TransactionType = @TransactionType
                                                AND DocType = @DocType";
                                        dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            TransactionDate = ReceiptDate,
                                            FromErpPrefix = GrErpPrefix,
                                            FromErpNo = GrErpNo,
                                            FromSeq = GrSeq,
                                            InventoryId,
                                            TransactionType = 1,
                                            DocType = 1,
                                        });
                                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    }
                                }

                                foreach (var item1 in LotNumberResult)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 LotNumberId
                                            FROM MES.VehicleDetail a 
                                            WHERE 1=1
                                            AND a.LotNumberId = @LotNumberId";
                                    dynamicParameters.Add("LotNumberId", item1.LotNumberId);

                                    var VehicleResult = sqlConnection.Query(sql, dynamicParameters);
                                    if (VehicleResult.Count() > 0) throw new SystemException("批號已經被使用,不可以執行反確認!!");


                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 LotNumberId
                                            FROM SCM.LnDetail a 
                                            WHERE 1=1
                                            AND a.LotNumberId = @LotNumberId";
                                    dynamicParameters.Add("LotNumberId", item1.LotNumberId);

                                    var LnDetailResult = sqlConnection.Query(sql, dynamicParameters);
                                    if (LnDetailResult.Count() <= 0)
                                    {
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"DELETE SCM.LotNumber
                                                WHERE 1=1
                                                AND LotNumberId = @LotNumberId
                                                ";
                                        dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            LotNumberId = item1.LotNumberId,
                                        });
                                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    }
                                }

                                #endregion
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

        #endregion

        #region //Delete
        #region //DeleteGoodsReceipt -- 進貨單資料刪除 -- Ann 2024-03-05
        public string DeleteGoodsReceipt(int GrId)
        {
            try
            {
                int rowsAffected = 0;

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷進貨單資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 ConfirmStatus, TransferStatus
                                FROM SCM.GoodsReceipt
                                WHERE GrId = @GrId
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("GrId", GrId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var GoodsReceiptResult = sqlConnection.Query(sql, dynamicParameters);
                        if (GoodsReceiptResult.Count() <= 0) throw new SystemException("進貨單資料錯誤!");

                        foreach (var item in GoodsReceiptResult)
                        {
                            if (item.TransferStatus != "N") throw new SystemException("進貨單已拋轉，無法刪除!!");
                            if (item.ConfirmStatus != "N") throw new SystemException("進貨單狀態無法刪除!!");
                        }
                        #endregion

                        #region //確認進貨單身是否有綁定進貨檢單據
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.QcRecordId, a.QcStatus 
                                , c.GrErpPrefix, c.GrErpNo, b.GrSeq
                                FROM MES.QcGoodsReceipt a 
                                INNER JOIN SCM.GrDetail b ON a.GrDetailId = b.GrDetailId
                                INNER JOIN SCM.GoodsReceipt c ON b.GrId = c.GrId
                                WHERE  c.GrId = @GrId";
                        dynamicParameters.Add("GrId", GrId);

                        var QcGoodsReceiptResult = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in QcGoodsReceiptResult)
                        {
                            if (item.QcStatus != "A")
                            {
                                throw new SystemException("進貨單身【" + item.GrErpPrefix + "-" + item.GrErpNo + "(" + item.GrSeq + ")】已綁定進貨檢單據無法刪除!!");
                            }

                            #region //進貨檢單據刪除
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE MES.QcGoodsReceipt
                                    WHERE QcRecordId = @QcRecordId
                                    AND QcStatus = 'A'";
                            dynamicParameters.Add("QcRecordId", item.QcRecordId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE MES.QcRecord
                                    WHERE QcRecordId = @QcRecordId";
                            dynamicParameters.Add("QcRecordId", item.QcRecordId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }
                        #endregion

                        #region //確認是否有註冊進貨條碼，且條碼尚未被其它單據使用
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 c.GrErpPrefix, c.GrErpNo, c.GrSeq
                                FROM SCM.GrBarcode a 
                                INNER JOIN MES.Barcode b ON a.BarcodeNo = b.BarcodeNo
                                INNER JOIN SCM.GrDetail c ON a.GrDetailId = c.GrDetailId
                                WHERE c.GrId = @GrId
                                AND (b.MoId IS NOT NULL OR b.BarcodeStatus != '8')";
                        dynamicParameters.Add("GrId", GrId);

                        var GrBarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in GrBarcodeResult)
                        {
                            throw new SystemException("進貨單身【" + item.GrErpPrefix + "-" + item.GrErpNo + "(" + item.GrSeq + ")】已註冊進貨條碼，且條碼已被使用，無法刪除!!");
                        }
                        #endregion

                        #region //刪除關聯table
                        #region //進貨條碼刪除
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a 
                                FROM SCM.GrBarcode a 
                                INNER JOIN SCM.GrDetail b ON a.GrDetailId = b.GrDetailId
                                WHERE b.GrId = @GrId";
                        dynamicParameters.Add("GrId", GrId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //進貨單身刪除
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.GrDetail
                                WHERE GrId = @GrId";
                        dynamicParameters.Add("GrId", GrId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.GoodsReceipt
                                WHERE GrId = @GrId";
                        dynamicParameters.Add("GrId", GrId);

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

        #region //DeleteGrDetail -- 進貨單單身資料刪除 -- Ann 2024-03-05
        public string DeleteGrDetail(int GrDetailId)
        {
            try
            {
                int rowsAffected = 0;

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
                            #region //判斷進貨單單身資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.GrId
                                    , b.ConfirmStatus, b.TransferStatus, b.Exchange
                                    FROM SCM.GrDetail a 
                                    INNER JOIN SCM.GoodsReceipt b ON a.GrId = b.GrId
                                    WHERE a.GrDetailId = @GrDetailId";
                            dynamicParameters.Add("GrDetailId", GrDetailId);

                            var GrDetailResult = sqlConnection.Query(sql, dynamicParameters);
                            if (GrDetailResult.Count() <= 0) throw new SystemException("進貨單資料錯誤!");

                            int GrId = -1;
                            double? Exchange = 0;
                            foreach (var item in GrDetailResult)
                            {
                                if (item.ConfirmStatus != "N") throw new SystemException("進貨單狀態無法刪除!!");
                                GrId = item.GrId;
                                Exchange = item.Exchange;
                            }
                            #endregion

                            #region //確認此品項是否有綁定進貨檢單據
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.QcRecordId, a.QcStatus
                                    FROM MES.QcGoodsReceipt a 
                                    WHERE a.GrDetailId = @GrDetailId";
                            dynamicParameters.Add("GrDetailId", GrDetailId);

                            var QcGoodsReceiptResult = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item in QcGoodsReceiptResult)
                            {
                                if (item.QcStatus != "A")
                                {
                                    throw new SystemException("此進貨單身已有進貨檢紀錄，無法刪除!!");
                                }

                                #region //進貨檢單據刪除
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE MES.QcGoodsReceipt
                                        WHERE QcRecordId = @QcRecordId
                                        AND QcStatus = 'A'";
                                dynamicParameters.Add("QcRecordId", item.QcRecordId);

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE MES.QcRecord
                                        WHERE QcRecordId = @QcRecordId";
                                dynamicParameters.Add("QcRecordId", item.QcRecordId);

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }
                            #endregion

                            #region //確認是否有註冊進貨條碼，且條碼尚未被其它單據使用
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM SCM.GrBarcode a 
                                    INNER JOIN MES.Barcode b ON a.BarcodeNo = b.BarcodeNo
                                    WHERE a.GrDetailId = @GrDetailId
                                    AND (b.MoId IS NOT NULL OR b.BarcodeStatus != '8')";
                            dynamicParameters.Add("GrDetailId", GrDetailId);

                            var GrBarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item in GrBarcodeResult)
                            {
                                throw new SystemException("此進貨單身已註冊進貨條碼，且條碼已被使用，無法刪除!!");
                            }
                            #endregion

                            #region //進貨條碼刪除
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE SCM.GrBarcode
                                    WHERE GrDetailId = @GrDetailId";
                            dynamicParameters.Add("GrDetailId", GrDetailId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //進貨單身刪除
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE SCM.GrDetail
                                    WHERE GrDetailId = @GrDetailId";
                            dynamicParameters.Add("GrDetailId", GrDetailId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //重新計算進貨單資料
                            JObject updateGrDetailResult = JObject.Parse(UpdateGrTotal(GrId, Exchange, sqlConnection, sqlConnection2));
                            if (updateGrDetailResult["status"].ToString() != "success")
                            {
                                throw new SystemException(updateGrDetailResult["msg"].ToString());
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

        #region //BatchDeleteGrDetail -- 批量刪除進貨單詳細資料 -- Ann 2024-11-13
        public string BatchDeleteGrDetail(string GrDetailList)
        {
            try
            {
                if (GrDetailList.Length <= 0) throw new SystemException("進貨單身清單不能為空!!");

                int rowsAffected = 0;
                int GrId = -1;
                double? Exchange = 0;
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

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            string[] GrDetailIdList = GrDetailList.Split(',');
                            foreach (var GrDetailId in GrDetailIdList)
                            {
                                #region //判斷進貨單單身資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 a.GrId, a.GrSeq
                                        , b.ConfirmStatus, b.TransferStatus, b.Exchange
                                        FROM SCM.GrDetail a 
                                        INNER JOIN SCM.GoodsReceipt b ON a.GrId = b.GrId
                                        WHERE a.GrDetailId = @GrDetailId";
                                dynamicParameters.Add("GrDetailId", GrDetailId);

                                var GrDetailResult = sqlConnection.Query(sql, dynamicParameters);
                                if (GrDetailResult.Count() <= 0) throw new SystemException("進貨單資料錯誤!");

                                string GrSeq = "";
                                foreach (var item in GrDetailResult)
                                {
                                    if (item.ConfirmStatus != "N") throw new SystemException("進貨單狀態無法刪除!!");
                                    GrId = item.GrId;
                                    GrSeq = item.GrSeq;
                                    Exchange = item.Exchange;
                                }
                                #endregion

                                #region //確認此品項是否有綁定進貨檢單據
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.QcRecordId, a.QcStatus
                                        FROM MES.QcGoodsReceipt a 
                                        WHERE a.GrDetailId = @GrDetailId";
                                dynamicParameters.Add("GrDetailId", GrDetailId);

                                var QcGoodsReceiptResult = sqlConnection.Query(sql, dynamicParameters);

                                foreach (var item in QcGoodsReceiptResult)
                                {
                                    if (item.QcStatus != "A")
                                    {
                                        throw new SystemException("此進貨單身已有進貨檢紀錄，無法刪除!!");
                                    }

                                    #region //進貨檢單據刪除
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE MES.QcGoodsReceipt
                                            WHERE QcRecordId = @QcRecordId
                                            AND QcStatus = 'A'";
                                    dynamicParameters.Add("QcRecordId", item.QcRecordId);

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE MES.QcRecord
                                            WHERE QcRecordId = @QcRecordId";
                                    dynamicParameters.Add("QcRecordId", item.QcRecordId);

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                                #endregion

                                #region //確認是否有註冊進貨條碼，且條碼尚未被其它單據使用
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM SCM.GrBarcode a 
                                        INNER JOIN MES.Barcode b ON a.BarcodeNo = b.BarcodeNo
                                        WHERE a.GrDetailId = @GrDetailId
                                        AND (b.MoId IS NOT NULL OR b.BarcodeStatus != '8')";
                                dynamicParameters.Add("GrDetailId", GrDetailId);

                                var GrBarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                                foreach (var item in GrBarcodeResult)
                                {
                                    throw new SystemException("此進貨單身【" + GrSeq + "】已註冊進貨條碼，且條碼已被使用，無法刪除!!");
                                }
                                #endregion

                                #region //進貨條碼刪除
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE SCM.GrBarcode
                                        WHERE GrDetailId = @GrDetailId";
                                dynamicParameters.Add("GrDetailId", GrDetailId);

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region //進貨單身刪除
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE SCM.GrDetail
                                        WHERE GrDetailId = @GrDetailId";
                                dynamicParameters.Add("GrDetailId", GrDetailId);

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }

                            #region //重新計算進貨單資料
                            JObject updateGrDetailResult = JObject.Parse(UpdateGrTotal(GrId, Exchange, sqlConnection, sqlConnection2));
                            if (updateGrDetailResult["status"].ToString() != "success")
                            {
                                throw new SystemException(updateGrDetailResult["msg"].ToString());
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

        #region //DeleteLnBarcode 刪除批號綁定條碼 GPAI 20240412
        public string DeleteLnBarcode(int LnBarcodeId)
        {
            try
            {
                int rowsAffected = 0;

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

                            #region //確認data
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.*
                                    FROM SCM.GrBarcode a 
                                    WHERE a.GrBarcodeId = @GrBarcodeId";
                            dynamicParameters.Add("GrBarcodeId", LnBarcodeId);

                            var LnBarcodeResult = sqlConnection.Query(sql, dynamicParameters);
                            if (LnBarcodeResult.Count() <= 0) throw new SystemException("無此條碼綁定資訊!");


                            #endregion

                            #region //確認條碼資料是否存在
                            int deleteBarcodeId = 0;
                            foreach (var item in LnBarcodeResult) {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT * 
                                    FROM MES.Barcode
                                    WHERE BarcodeNo = @BarcodeNo";
                                dynamicParameters.Add("BarcodeNo", item.BarcodeNo);

                                var resultBarcodeNo = sqlConnection.Query(sql, dynamicParameters);
                                if (resultBarcodeNo.Count() <= 0) throw new SystemException("無此條碼資訊!");

                                #region //是否有過站紀錄
                                foreach (var item2 in resultBarcodeNo) {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT * 
                                    FROM MES.BarcodeProcess
                                    WHERE BarcodeId = @BarcodeId";
                                    dynamicParameters.Add("BarcodeId", item.BarcodeId);

                                    var resultBarcodeProcess = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultBarcodeProcess.Count() > 0) throw new SystemException("該條碼有過站紀錄!");

                                    deleteBarcodeId = item2.BarcodeId;
                                }

                                
                                #endregion
                            }

                            #endregion

                            #region //刪除
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE SCM.LnBarcode
                                        WHERE LnBarcodeId = @LnBarcodeId";
                            dynamicParameters.Add("LnBarcodeId", LnBarcodeId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);


                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE SCM.GrBarcode
                                        WHERE GrBarcodeId = @GrBarcodeId";
                            dynamicParameters.Add("GrBarcodeId", LnBarcodeId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                            //dynamicParameters = new DynamicParameters();
                            //sql = @"DELETE MES.Barcode
                            //            WHERE BarcodeId = @BarcodeId";
                            //dynamicParameters.Add("BarcodeId", deleteBarcodeId);

                            //rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

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

        #region //DeleteAllBind 刪除批號全部綁定條碼 GPAI 20240412
        public string DeleteAllBind(int GrDetailId)
        {
            try
            {
                int rowsAffected = 0;

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

                            #region //確認data
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.GrBarcodeId, a.GrDetailId, a.BarcodeNo
                                    , b.LotNumber, b.MtlItemId
                                    FROM SCM.GrBarcode a
                                    INNER JOIN SCM.GrDetail b ON a.GrDetailId = b.GrDetailId
                                    WHERE a.GrDetailId = @GrDetailId";
                            dynamicParameters.Add("GrDetailId", GrDetailId);

                            var GrBarcode = sqlConnection.Query(sql, dynamicParameters);
                            if (GrBarcode.Count() <= 0) throw new SystemException("無此條碼綁定資訊!");
                            var mtlItemId = GrBarcode.FirstOrDefault().MtlItemId;
                            var lotNumber = GrBarcode.FirstOrDefault().LotNumber;

                            #endregion

                            #region //確認data
                            dynamicParameters = new DynamicParameters();
                            sql = @"select a.LotNumberId, a.BarcodeNo, b.LotNumberNo
                                    from SCM.LnBarcode a
                                    INNER JOIN SCM.LotNumber b ON a.LotNumberId = b.LotNumberId
                                    WHERE b.MtlItemId = @MtlItemId AND b.LotNumberNo = @LotNumberNo";
                            dynamicParameters.Add("MtlItemId", mtlItemId);
                            dynamicParameters.Add("LotNumberNo", lotNumber);


                            var LnBarcode = sqlConnection.Query(sql, dynamicParameters);
                            if (LnBarcode.Count() <= 0) throw new SystemException("無此條碼綁定資訊!");
                            var lotNumberId = LnBarcode.FirstOrDefault().LotNumberId;
                            #endregion

                            #region //確認條碼資料是否存在
                            int deleteBarcodeId = 0;
                            foreach (var item in GrBarcode)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT * 
                                    FROM MES.Barcode
                                    WHERE BarcodeNo = @BarcodeNo";
                                dynamicParameters.Add("BarcodeNo", item.BarcodeNo);

                                var resultBarcodeNo = sqlConnection.Query(sql, dynamicParameters);
                                if (resultBarcodeNo.Count() <= 0) throw new SystemException("無此條碼資訊!");

                                #region //是否有過站紀錄
                                foreach (var item2 in resultBarcodeNo)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT * 
                                    FROM MES.BarcodeProcess
                                    WHERE BarcodeId = @BarcodeId";
                                    dynamicParameters.Add("BarcodeId", item.BarcodeId);

                                    var resultBarcodeProcess = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultBarcodeProcess.Count() > 0) throw new SystemException("該條碼有過站紀錄!");

                                    deleteBarcodeId = item2.BarcodeId;
                                }


                                #endregion
                            }

                            #endregion

                            #region //刪除
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE SCM.LnBarcode
                                        WHERE LotNumberId = @LotNumberId";
                            dynamicParameters.Add("LotNumberId", lotNumberId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);


                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE SCM.GrBarcode
                                        WHERE GrDetailId = @GrDetailId";
                            dynamicParameters.Add("GrDetailId", GrDetailId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                            //dynamicParameters = new DynamicParameters();
                            //sql = @"DELETE MES.Barcode
                            //            WHERE BarcodeId = @BarcodeId";
                            //dynamicParameters.Add("BarcodeId", deleteBarcodeId);

                            //rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

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


        #endregion

        #region //API
        #region //CalculateTaxAmt -- 計算稅額 -- Ann 2024-03-14
        public string CalculateTaxAmt(double? AvailableQty, double? OrigUnitPrice, double? OrigDiscountAmt, double? OrigAmount, double? TaxRate, string TaxType, double? Exchange
            , int OriUnitDecimal, int OriPriceDecimal, int UnitDecimal, int PriceDecimal)
        {
            try
            {
                if (OriUnitDecimal < 0) throw new SystemException("【原幣單價小數點進位數】格式錯誤!");
                if (OriPriceDecimal < 0) throw new SystemException("【原幣金額小數點進位數】格式錯誤!");
                if (UnitDecimal < 0) throw new SystemException("【本幣單價小數點進位數】格式錯誤!");
                if (PriceDecimal < 0) throw new SystemException("【本幣金額小數點進位數】格式錯誤!");
                if (TaxType.Length <= 0) throw new SystemException("【課稅別】格式錯誤!");

                double OrigPreTaxAmt = 0; //原幣未稅金額
                double OrigTaxAmt = 0; //原幣稅金額
                double PreTaxAmt = 0; //本幣未稅金額
                double TaxAmt = 0; //本幣稅金額
                OrigAmount = Math.Round(Convert.ToDouble(AvailableQty * OrigUnitPrice), OriPriceDecimal);

                switch (TaxType)
                {
                    #region //應稅內含
                    case "1":
                        OrigPreTaxAmt = Math.Round(Convert.ToDouble((AvailableQty * OrigUnitPrice - OrigDiscountAmt) / (1 + TaxRate)), OriPriceDecimal);
                        OrigTaxAmt = Math.Round(Convert.ToDouble(OrigAmount - OrigPreTaxAmt), OriPriceDecimal);
                        PreTaxAmt = Math.Round(Convert.ToDouble(OrigPreTaxAmt * Exchange), PriceDecimal);
                        TaxAmt = Math.Round(Convert.ToDouble(OrigTaxAmt * Exchange), PriceDecimal);
                        break;
                    #endregion
                    #region //應稅外加
                    case "2":
                        OrigPreTaxAmt = Math.Round(Convert.ToDouble(AvailableQty * OrigUnitPrice - OrigDiscountAmt), OriPriceDecimal);
                        OrigTaxAmt = Math.Round(Convert.ToDouble((AvailableQty * OrigUnitPrice - OrigDiscountAmt) * TaxRate), OriPriceDecimal);
                        PreTaxAmt = Math.Round(Convert.ToDouble(OrigPreTaxAmt * Exchange), PriceDecimal);
                        TaxAmt = Math.Round(Convert.ToDouble(OrigTaxAmt * Exchange), PriceDecimal);
                        break;
                    #endregion
                    #region //免稅
                    case "3":
                    case "4":
                    case "9":
                        OrigPreTaxAmt = Math.Round(Convert.ToDouble(AvailableQty * OrigUnitPrice - OrigDiscountAmt), OriPriceDecimal);
                        OrigTaxAmt = 0;
                        PreTaxAmt = Math.Round(Convert.ToDouble(OrigPreTaxAmt * Exchange), PriceDecimal);
                        TaxAmt = 0;
                        break;
                    #endregion
                    #region //例外狀況
                    default:
                        throw new SystemException("稅別碼資料錯誤!!");
                        #endregion
                }

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    OrigPreTaxAmt,
                    OrigTaxAmt,
                    PreTaxAmt,
                    TaxAmt,
                    OrigAmount
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
        #endregion
    }
}
