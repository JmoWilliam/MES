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
using System.Text.RegularExpressions;
using System.Transactions;
using System.Web;

namespace SCMDA
{
    public class ReceiveOrderDA
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

        public ReceiveOrderDA()
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
        #region //GetReceiveOrder -- 取得銷貨單資料 -- Shintokuro 2024.02.01
        public string GetReceiveOrder(int RoId, string RoErpPrefix, string RoErpFullNo, int CustomerId, int SalesmenId
            , string StartDocDate, string EndDocDate, string ConfirmStatus, string TransferStatusMES
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.RoId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" ,a.CompanyId, a.RoErpPrefix, a.RoErpNo, a.CustomerId, a.DepartmentId, a.SalesmenId, 
                            a.CustomerFullName, a.CustomerAddressFirst, a.CustomerAddressSecond, a.Factory, a.Currency, a.ExchangeRate, 
                            a.OriginalAmount, a.InvNumEnd, a.UiNo, a.InvoiceType, a.TaxType, a.InvoiceAddressFirst, a.InvoiceAddressSecond, a.Remark, 
                            a.PrintCount, a.ConfirmStatus, a.UpdateCode, a.OriginalTaxAmount, a.CollectionSalesmenId, a.Remark1, a.Remark2, 
                            a.Remark3, a.InvoicesVoid, a.CustomsClearance, a.RowCnt, a.TotalQuantity, a.CashSales, a.StaffId, a.RevenueJournalCode, 
                            a.CostJournalCode, a.ApplyYYMM, a.LCNO, a.INVOICENO, a.InvoicesPrintCount, a.ConfirmUserId, a.TaxRate, 
                            a.PretaxAmount, a.TaxAmount, a.PaymentTerm, a.SoDetailId, a.AdvOrderPrefix, a.AdvOrderNo, a.OffsetAmount, a.OffsetTaxAmount, 
                            a.TotalPackages, a.SignatureStatus, a.ChangeInvCode, a.NewRoPrefix, a.NewRoNo, a.TransferStatusERP, a.ProcessCode, 
                            a.AttachInvWithShip, a.BondCode, a.TransmissionCount, a.Invoicer, a.InvCode, a.ContactPerson, a.Courier, a.SiteCommCalcMethod, 
                            a.SiteCommRate, a.CommCalcInclTax, a.TotalCommAmount, a.TransportMethod, a.DispatchOrderPrefix, a.DispatchOrderNo, 
                            a.DeclarationNumber, a.FullNameOfDelivCustomer, a.RoPriceType, a.TelephoneNumber, a.FaxNumber, a.ShipNoticePrefix, 
                            a.ShipNoticeNo, a.TradingTerms, a.CustomerEgFullName, a.InvNumGenMethod, a.DocSourceCode, a.NoCredLimitControl, 
                            a.InstallmentSettlement, a.InstallmentCount, a.AutoApportionByInstallment, a.StartYearMonth, a.TaxCode, a.CustomsManual, 
                            a.RemarkCode, a.MultipleInvoices, a.InvNumStart, a.NumberOfInvoices, a.MultiTaxRate, a.TaxCurrencyRate, a.VoidDate, 
                            a.VoidApprovalDocNum, a.VoidReason, a.Source, a.IncomeDraftID, a.IncomeDraftSeq, a.IncomeVoucherType, 
                            a.IncomeVoucherNumber, a.Status, a.ZeroTaxForBuyer, a.GenLedgerAcctType, a.InvoiceTime, a.InvCode2, a.InvSymbol, 
                            a.DeliveryCountry, a.VehicleIDshow, a.VehicleTypeNumber, a.VehicleIDhide, a.InvDonationRecipient, a.InvRandomCode, 
                            a.ReservedField, a.CreditCard4No, a.ContactEmail, a.ExpectedReceiptDate, a.OrigInvNumber, a.TransferStatusMES,
                            (a.RoErpPrefix + '-' + a.RoErpNo) RoErpFullNo, FORMAT(a.ReceiveDate, 'yyyy-MM-dd') ReceiveDate,
                            FORMAT(a.DocDate, 'yyyy-MM-dd') DocDate, FORMAT(a.InvoiceDate, 'yyyy-MM-dd') InvoiceDate, SourceType, PriceSourceTypeMain,
                            a.SourceOrderId, a.SourceFull, b.UserName SalesmenName, c.UserName ConfirmUserName, d.CustomerName
                            , (
                                SELECT 
                                  ISNULL(x1.RoSequence, '') RoSequence, ISNULL(x2.MtlItemNo, '') MtlItemNo, ISNULL(x2.MtlItemName, '') MtlItemName
                                , ISNULL(x3.InventoryNo, '') InventoryNo, ISNULL(x1.Quantity, '') Quantity, ISNULL(x1.Type, '') Type
                                , ISNULL(x1.FreeSpareQty, '') FreeSpareQty, ISNULL(x1.AvailableQty, '') AvailableQty, ISNULL(x1.UnitPrice, '') UnitPrice
                                , ISNULL(x1.Amount, '') Amount
                                , ISNULL(x5.SoErpPrefix+'-'+x5.SoErpNo+'-'+x4.SoSequence, '') SoFullNo
                                , ISNULL(x7.TsnErpPrefix+'-'+x7.TsnErpNo+'-'+x6.TsnSequence, '') TsnFullNo
                                , ISNULL(x1.LotNumber, '') LotNumber
                                FROM SCM.RoDetail x1
                                LEFT JOIN PDM.MtlItem x2 ON x1.MtlItemId = x2.MtlItemId
                                LEFT JOIN SCM.Inventory x3 ON x1.InventoryId = x3.InventoryId
                                LEFT JOIN SCM.SoDetail x4 ON x1.SoDetailId = x4.SoDetailId
                                LEFT JOIN SCM.SaleOrder x5 ON x4.SoId = x5.SoId
                                LEFT JOIN SCM.TsnDetail x6 ON x1.TsnDetailId = x6.TsnDetailId
                                LEFT JOIN SCM.TempShippingNote x7 ON x6.TsnId = x7.TsnId
                                WHERE x1.RoId = a.RoId
                                FOR JSON PATH, ROOT('data')
                            ) RoDetail
                            ";
                    sqlQuery.mainTables =
                        @"FROM SCM.ReceiveOrder a
                          LEFT JOIN BAS.[User] b on a.SalesmenId = b.UserId
                          LEFT JOIN BAS.[User] c on a.ConfirmUserId = c.UserId
                          LEFT JOIN SCM.Customer d on a.CustomerId = d.CustomerId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RoId", @" AND a.RoId = @RoId", RoId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RoErpPrefix", @" AND a.RoErpPrefix = @RoErpPrefix", RoErpPrefix);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RoErpFullNo", @" AND (a.RoErpPrefix + '-' + a.RoErpNo ) LIKE '%' + @RoErpFullNo + '%'", RoErpFullNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerId", @" AND a.CustomerId = @CustomerId", CustomerId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SalesmenId", @" AND a.SalesmenId = @SalesmenId", SalesmenId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SalesmenId", @" AND a.SalesmenId = @SalesmenId", SalesmenId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ConfirmStatus", @" AND a.ConfirmStatus = @ConfirmStatus", ConfirmStatus);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TransferStatusMES", @" AND a.TransferStatusMES = @TransferStatusMES", TransferStatusMES);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.DocDate >= @StartDate ", StartDocDate.Length > 0 ? Convert.ToDateTime(StartDocDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.DocDate <= @EndDate ", EndDocDate.Length > 0 ? Convert.ToDateTime(EndDocDate).ToString("yyyy-MM-dd 23:59:59") : "");

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.RoId DESC";
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

        #region //GetRoDetail -- 取得銷貨單單身資料 -- Shintokuro 2024.02.06
        public string GetRoDetail(int RoId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.RoDetailId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" ,a.RoId, a.RoSequence, a.MtlItemId, a.InventoryId, a.Quantity, a.UomId, a.UnitPrice, a.Amount, 
                            a.SoDetailId, a.LotNumber, a.Remarks, a.CustomerMtlItemNo, a.ConfirmationCode, a.UpdateCode, a.FreeSpareQty, a.DiscountRate, 
                            a.CheckOutCode, a.CheckOutPrefix, a.CheckOutNo, a.CheckOutSequence, a.ProjectCode, a.Type, a.TsnDetailId, a.OrigPreTaxAmt, 
                            a.OrigTaxAmt, a.PreTaxAmt, a.TaxAmt, a.PackageQty, a.PackageFreeSpareQty, a.PackageUomId, a.BondedCode, a.SrQty, 
                            a.SrPackageQty, a.SrOrigPreTaxAmt, a.SrOrigTaxAmt, a.SrPreTaxAmt, a.SrTaxAmt, a.CommissionRate, a.CommissionAmount, 
                            a.OriginalCustomer, a.SrFreeSpareQty, a.SrPackageFreeSpareQty, a.NotPayTemp, a.ProductSerialNumberQty, a.ForecastCode, 
                            a.ForecastSequence, a.Location, a.AvailableQty, a.AvailableUomId, a.MultiBatch, a.FreeSpareRate, a.FinalCustomerCode, 
                            a.ReferenceQty, a.ReferencePackageQty, a.TaxRate, a.CRMSource, a.CRMPrefix, a.CRMNo, a.CRMSequence, a.CRMContractCode, 
                            a.CRMAllowDeduction, a.CRMDeductionQty, a.CRMDeductionUnit, a.DebitAccount, a.CreditAccount, a.TaxAmountAccount, 
                            a.BusinessItemNumber, a.TaxCode, a.DiscountAmount, a.K2NO, a.MarkingBINRecord, a.MarkingManagement, 
                            a.BillingUnitInPackage, a.DATECODE,
                            b.MtlItemNo,b.MtlItemName,b.MtlItemSpec,b.LotManagement,
                            c.InventoryNo,c.InventoryName,
                            b.MtlItemSpec,d.UomNo ,d1.UomNo AvailableUomNo,
                            (e1.SoErpPrefix + '-' + e1.SoErpNo + '-' + e.SoSequence ) SoFullNo,
                            (f1.TsnErpPrefix + '-' + f1.TsnErpNo + '-' + f.TsnSequence ) ToFullNo";
                    sqlQuery.mainTables =
                        @"FROM SCM.RoDetail a
                          INNER JOIN SCM.ReceiveOrder a1 on a.RoId = a1.RoId
                          INNER JOIN PDM.MtlItem b on a.MtlItemId = b.MtlItemId
                          INNER JOIN SCM.Inventory c on a.InventoryId = c.InventoryId
                          INNER JOIN PDM.UnitOfMeasure d on a.UomId = d.UomId
                          INNER JOIN PDM.UnitOfMeasure d1 on a.AvailableUomId = d1.UomId
                          LEFT JOIN SCM.SoDetail e on a.SoDetailId = e.SoDetailId
                          LEFT JOIN SCM.SaleOrder e1 on e.SoId = e1.SoId
                          LEFT JOIN SCM.TsnDetail f on a.TsnDetailId = f.TsnDetailId
                          LEFT JOIN SCM.TempShippingNote f1 on f.TsnId = f1.TsnId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND a1.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RoId", @" AND a.RoId = @RoId", RoId);

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.RoDetailId ASC";
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

        #region //GetRoErpPrefixSetting -- 取得銷貨單單別設定 -- Shintokuro 2024.01.30
        public string GetRoErpPrefixSetting(string RoErpPrefix)
        {
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
                }

                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT LTRIM(RTRIM(a.MQ067)) InvNumGenMethod 
                            FROM CMSMQ a
                            WHERE 1=1
                            AND a.MQ001 =@RoErpPrefix
                            ";
                    dynamicParameters.Add("RoErpPrefix", RoErpPrefix);

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

        #region //GetExchangeRate -- 取得ERP匯率資料 -- Shintokuro 2024.02.05
        public string GetExchangeRate(string Condition, string ErpPrefix, string OrderBy)
        {
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
                }

                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                {
                    //【銀行買進匯率】（報價單、客戶訂單、銷貨單、銷退單、結帳單、收款單）
                    //【銀行賣出匯率】（核價單、採購單、進貨單、退貨單、應付憑單、付款單）
                    //【報關買進匯率】應用出口系統『出貨通知單建立作業』
                    //【報關賣出匯率】應用進口系統『報關 / 贖單資料建立作業』所用的預設匯率。
                    dynamicParameters = new DynamicParameters();
                    string Today = DateTime.Now.ToString("yyyyMMdd");
                    sql = @"SELECT TOP 1
                                a.MG002 StartDate
                            , ROUND(LTRIM(RTRIM(a.MG003)),4) ExchangeRateInBank
                            , ROUND(LTRIM(RTRIM(a.MG004)),4) ExchangeRateOutBank
                            , ROUND(LTRIM(RTRIM(a.MG005)),4) ExchangeRateIn
                            , ROUND(LTRIM(RTRIM(a.MG006)),4) ExchangeRateOut
                            , b.MF001 Currency,b.MF003 decimalOrPrice ,b.MF004 decimalOrAmoute
                            , x.*
                            , CASE 
                                WHEN x1.ErpPrefix = 'I' THEN ROUND(LTRIM(RTRIM(a.MG003)),4) 
                                WHEN x1.ErpPrefix = 'O' THEN ROUND(LTRIM(RTRIM(a.MG004)),4) 
                                WHEN x1.ErpPrefix = 'E' THEN ROUND(LTRIM(RTRIM(a.MG005)),4) 
                                WHEN x1.ErpPrefix = 'W' THEN ROUND(LTRIM(RTRIM(a.MG006)),4) 
                               END AS ExchangeRate
                            FROM CMSMG a
                            INNER JOIN CMSMF b on a.MG001 = b.MF001
                            OUTER APPLY(
                                SELECT x1.MA003 CurrencyLocal,x2.MF003 decimalLoPrice,x2.MF004 decimalLoAmoute
                                FROM CMSMA x1
                                INNER JOIN CMSMF x2 on x1.MA003 = x2.MF001
                            ) x
                            OUTER APPLY(
                                SELECT MQ044 ErpPrefix 
                                FROM CMSMQ 
                                WHERE MQ001 = @ErpPrefix
                            ) x1
                            WHERE 1=1
                            ";
                    dynamicParameters.Add("ErpPrefix", ErpPrefix);

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "Condition", @" AND a.MG001 = @Condition", Condition);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "Today", @" AND a.MG002 <= @Today", Today);
                    OrderBy = OrderBy.Length > 0 ? OrderBy : "a.MG002 DESC";

                    sql += OrderBy.Length > 0 ? string.Format(@" ORDER BY {0}", OrderBy) : "";
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

        #region //GetCurrencyDouble -- 取得幣別小數取位 -- Shintokuro 2024.05.23
        public string GetCurrencyDouble(string Currency)
        {
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
                }

                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    string Today = DateTime.Now.ToString("yyyyMMdd");
                    sql = @"SELECT a.MF003 PriceDouble ,a.MF004 AmountDouble
                            FROM CMSMF a
                            WHERE 1=1
                            AND a.MF001 = @Currency";
                    dynamicParameters.Add("Currency", Currency);

                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0) throw new SystemException("【幣別】資料錯誤!找不到!!");

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


        #region //GetInventory -- 取得ERP庫別資料 -- Shintokuro 2024.02.05
        public string GetInventory(int ViewCompanyId, string Table,string InventoryNo)
        {
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
                }

                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();

                    switch (Table)
                    {
                        case "CMSMC":
                            sql = @"SELECT LTRIM(RTRIM(MC001)) InventoryNo, LTRIM(RTRIM(MC002)) InventoryName
                                    FROM CMSMC 
                                    WHERE 1=1";
                            break;
                        case "CMSNL":
                            sql = @"SELECT LTRIM(RTRIM(NL002)) Location,LTRIM(RTRIM(NL003)) LocationName
                                    FROM CMSNL 
                                    WHERE 1=1
                                    AND NL001=@InventoryNo";
                            dynamicParameters.Add("InventoryNo", InventoryNo);
                            break;
                    }
                    
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

        #region //GetSourceOrderData-- 取得來源單據資料 -- Shintokuro 2024.01.30
        public string GetSourceOrderData(string SourceType, string SourcePrefix, string SourceNo, string SourceSequence, string DoDetailList)
        {
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

                    if(SourceType == "SoDetail" && DoDetailList.Length > 0)
                    {
                        string SoSequenceList = "";
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT b.SoSequence
                                FROM SCM.DoDetail a
                                INNER JOIN SCM.SoDetail b on a.SoDetailId = b.SoDetailId
                                WHERE a.DoDetailId in @SoDetailId";
                        dynamicParameters.Add("SoDetailId", DoDetailList.Split(','));

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【訂單】資料找不到請重新確認錯誤!");

                        foreach (var item in result)
                        {
                            SoSequenceList += ",'"+ item.SoSequence +"'";
                        }
                        SourceSequence = "(" + SoSequenceList.Substring(1) + ")";
                    }
                }

                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();

                    switch (SourceType) {
                        #region //訂單單頭
                        case "So":
                            sql = @"SELECT 
                                    LTRIM(RTRIM(a.TC004)) Customer,
                                    LTRIM(RTRIM(a.TC005)) Department, LTRIM(RTRIM(a.TC006)) Salesmen,
                                    LTRIM(RTRIM(a.TC007)) Factory, LTRIM(RTRIM(a.TC016)) TaxType, LTRIM(RTRIM(a.TC008)) Currency,
                                    LTRIM(RTRIM(a.TC009)) ExchangeRate, LTRIM(RTRIM(0)) RowCnt, LTRIM(RTRIM(a.TC015)) Remark,
                                    LTRIM(RTRIM(a.TC053)) CustomerFullName, LTRIM(RTRIM(a.TC010)) CustomerAddressFirst, LTRIM(RTRIM(a.TC011)) CustomerAddressSecond,
                                    LTRIM(RTRIM(a.TC031)) TotalQty, LTRIM(RTRIM(a.TC029)) TotalAmount, LTRIM(RTRIM(a.TC041)) TaxRate,
                                    LTRIM(RTRIM(a.TC030)) TaxAmt, LTRIM(RTRIM(a.TC019)) TransportMethod, '' DispatchOrderPrefix,
                                    '' DispatchOrderNo, LTRIM(RTRIM(a.TC078)) TaxCode, LTRIM(RTRIM(a.TC091)) MultiTaxRate,
                                    LTRIM(RTRIM(a.TC077)) NoCredLimitControl, LTRIM(RTRIM(a.TC017)) LCNO, LTRIM(RTRIM(a.TC037)) InvoiceNO,
                                    LTRIM(RTRIM(a.TC042)) PaymentTerm , LTRIM(RTRIM(a.TC068)) TradingTerms, LTRIM(RTRIM(a.TC053)) CustomerName, 
                                    LTRIM(RTRIM(a.TC018)) ContactPerson, LTRIM(RTRIM(a.TC066)) TelephoneNumber , LTRIM(RTRIM(a.TC067)) FaxNumber,
                                    LTRIM(RTRIM(a.TC063)) InvoiceAddressFirst, LTRIM(RTRIM(a.TC071)) CustomerEgFullName
                                    ,c.MF003 PriceDouble ,c.MF004 AmountDouble
                                    FROM COPTC a
                                    INNER JOIN COPTD b on a.TC001 = b.TD001 AND a.TC002 = b.TD002
                                    INNER JOIN CMSMF c on a.TC008 = c.MF001
                                    WHERE 1=1
                                    AND a.TC001 =@SourcePrefix
                                    AND a.TC002 =@SourceNo
                                    ";
                            break;
                        #endregion
                        #region //訂單單身
                        case "SoDetail":
                            sql = @"SELECT 
                                      LTRIM(RTRIM(a.TD001)) SoErpPrefix     , LTRIM(RTRIM(a.TD002)) SoErpNo             , LTRIM(RTRIM(a.TD003)) SoSequence
                                    , LTRIM(RTRIM(a.TD004)) MtlItemNo       , LTRIM(RTRIM(a.TD007)) InventoryNo         , LTRIM(RTRIM(a.TD008)) Quantity
                                    , LTRIM(RTRIM(a.TD010)) UomNo           , LTRIM(RTRIM(a.TD014)) CustomerMtlItemNo   , LTRIM(RTRIM(a.TD016)) CheckOutCode    
                                    , LTRIM(RTRIM(a.TD020)) Remarks         , LTRIM(RTRIM(a.TD021)) ConfirmStatus       , LTRIM(RTRIM(a.TD024)) FreeSpareQty        
                                    , LTRIM(RTRIM(a.TD026)) DiscountRate    , LTRIM(RTRIM(a.TD027)) ProjectNo           , LTRIM(RTRIM(a.TD049)) Type                
                                    , LTRIM(RTRIM(a.TD050)) SparePartsQty   , LTRIM(RTRIM(a.TD066)) LotNumber           , LTRIM(RTRIM(a.TD070)) TaxRate     
                                    , LTRIM(RTRIM(a.TD076)) AvailableQty    , LTRIM(RTRIM(a.TD077)) AvailableUomNo  
                                    , LTRIM(RTRIM(c.MB002)) MtlItemName     , LTRIM(RTRIM(c.MB003)) MtlItemSpec         , LTRIM(RTRIM(c.MB022)) LotManagement
                                    , LTRIM(RTRIM(d.MC002)) InventoryName   , LTRIM(RTRIM(b.TC012)) CustomerPurchaseOrder
                                    , ROUND(LTRIM(RTRIM(a.TD011)) , CAST(e.MF003 AS INT)) UnitPrice
                                    , ROUND(LTRIM(RTRIM(a.TD012)) , CAST(e.MF004 AS INT)) Amount
                                    , ROUND(LTRIM(RTRIM(a.TD012)) , CAST(e.MF004 AS INT)) OriginalAmount
                                    , ROUND(LTRIM(RTRIM(a.TD012 * a.TD070)) , CAST(e.MF004 AS INT)) OriginalTaxAmount
                                    , ROUND(LTRIM(RTRIM(a.TD012 * b.TC009)) , CAST(e.MF004 AS INT)) PretaxAmount
                                    , ROUND(LTRIM(RTRIM(a.TD012 * a.TD070 * b.TC009)) , CAST(e.MF004 AS INT)) TaxAmount
                                    , (ISNULL(a.TD008,0) -  ISNULL(a.TD009,0) - ISNULL(x.UnconfirmedQty,0)) CanUseQty
                                    , (ISNULL(a.TD024,0) -  ISNULL(a.TD025,0) - ISNULL(x.UnconfirmedFreeSpareQty,0)) CanUseFreebieQty
                                    , (ISNULL(a.TD050,0) -  ISNULL(a.TD051,0) - ISNULL(x.UnconfirmedFreeSpareQty,0)) CanUseSpareQty
                                    , (ISNULL(a.TD076,0) -  ISNULL(a.TD078,0) - ISNULL(x.UnconfirmedAvailableQty,0)) CanUseAvailableQty,
                                    ISNULL(x.UnconfirmedQty,0) UnconfirmedQty,
                                    ISNULL(x.UnconfirmedFreeSpareQty,0) UnconfirmedFreeSpareQty, 
                                    ISNULL(x.UnconfirmedAvailableQty,0) UnconfirmedAvailableQty
                                    FROM COPTD a
                                    INNER JOIN COPTC b on a.TD001 = b.TC001 AND a.TD002 = b.TC002
                                    INNER JOIN INVMB c on a.TD004 = c.MB001
                                    INNER JOIN CMSMC d on a.TD007 = d.MC001
                                    INNER JOIN CMSMF e on b.TC008 = e.MF001
                                    OUTER APPLY(
                                        SELECT SUM(ISNULL(x1.TH008,0)) UnconfirmedQty 
                                              ,SUM(ISNULL(x1.TH024,0)) UnconfirmedFreeSpareQty 
                                              ,SUM(ISNULL(x1.TH061,0)) UnconfirmedAvailableQty
                                        FROM COPTH x1
                                        WHERE x1.TH014 = a.TD001
                                        AND x1.TH015 = a.TD002
                                        AND x1.TH016 = a.TD003
                                        AND x1.TH020 ='N'
                                    ) x
                                    WHERE a.TD001 = @SourcePrefix
                                    AND a.TD002 = @SourceNo
                                    AND b.TC027 = 'Y'
                                    AND a.TD016 = 'N'
                                    ";
                            if (SourceSequence.Length > 0)
                            {
                                sql += "AND a.TD003 in " + SourceSequence;
                                dynamicParameters.Add("SourceSequence", SourceSequence);
                                //sql += "AND a.TG003 = '" + SourceSequence + @"'";
                            }
                            break;
                        #endregion

                        #region //暫出單單頭
                        case "To":
                            sql = @"SELECT 
                                     LTRIM(RTRIM(a.TF004)) ToObject             ,LTRIM(RTRIM(a.TF005)) Customer                 ,LTRIM(RTRIM(a.TF006)) CustomerNameTsn
                                    ,LTRIM(RTRIM(a.TF007)) Department           ,LTRIM(RTRIM(a.TF008)) UserNo                   ,LTRIM(RTRIM(a.TF009)) Factory
                                    ,LTRIM(RTRIM(a.TF010)) TaxType              ,LTRIM(RTRIM(a.TF011)) Currency                 ,LTRIM(RTRIM(a.TF012)) ExchangeRate
                                    ,LTRIM(RTRIM(a.TF013)) RowCnt               ,LTRIM(RTRIM(a.TF014)) Remark                   ,LTRIM(RTRIM(a.TF015)) CustomerFullName
                                    ,LTRIM(RTRIM(a.TF016)) CustomerAddressFirst ,LTRIM(RTRIM(a.TF017)) CustomerAddressSecond    ,LTRIM(RTRIM(a.TF022)) TotalQty
                                    ,LTRIM(RTRIM(a.TF023)) TotalAmount          ,LTRIM(RTRIM(a.TF026)) TaxRate                  ,LTRIM(RTRIM(a.TF027)) TaxAmt
                                    ,LTRIM(RTRIM(a.TF031)) TransportMethod      ,LTRIM(RTRIM(a.TF032)) DispatchOrderPrefix      ,LTRIM(RTRIM(a.TF033)) DispatchOrderNo
                                    ,LTRIM(RTRIM(a.TF039)) TaxCode              ,LTRIM(RTRIM(a.TF040)) MultiTaxRate             ,LTRIM(RTRIM(a.TF044)) NoCredLimitControl
                                    ,LTRIM(RTRIM(a.TF045)) LCNO                 ,LTRIM(RTRIM(a.TF046)) InvoiceNO
                                    ,x.PaymentTerm,x.TradingTerms,x.CustomerName,x.ContactPerson,x.TelephoneNumber,x.FaxNumber
                                    ,x.InvoiceAddressFirst,x.CustomerEgFullName
                                    ,b.MF003 PriceDouble ,b.MF004 AmountDouble
                                    FROM INVTF a
                                    INNER JOIN CMSMF b on a.TF011 = b.MF001
                                    OUTER APPLY(
                                        SELECT TOP 1 LTRIM(RTRIM(x2.TC042)) PaymentTerm , LTRIM(RTRIM(x2.TC068)) TradingTerms, LTRIM(RTRIM(x2.TC053)) CustomerName,  
                                        LTRIM(RTRIM(x2.TC018)) ContactPerson, LTRIM(RTRIM(x2.TC066)) TelephoneNumber , LTRIM(RTRIM(x2.TC067)) FaxNumber, 
                                        LTRIM(RTRIM(x2.TC063)) InvoiceAddressFirst, LTRIM(RTRIM(x2.TC071)) CustomerEgFullName
                                        FROM INVTG x1
                                        INNER JOIN  COPTC x2 on x1.TG014 = x2.TC001 AND x1.TG015 = x2.TC002
                                        WHERE a.TF001= x1.TG001 AND a.TF002 = x1.TG002
                                        ORDER BY x1.TG003 ASC
                                    ) x
                                    WHERE 1=1
                                    AND a.TF001 =@SourcePrefix
                                    AND a.TF002 =@SourceNo
                                    ";
                            break;
                        #endregion
                        #region //暫出單單身
                        case "ToDetail":
                            sql = @"SELECT 
                                     LTRIM(RTRIM(a.TG001)) TsnErpPrefix             ,LTRIM(RTRIM(a.TG002)) TsnErpNo             ,LTRIM(RTRIM(a.TG003)) TsnSequence
                                    ,LTRIM(RTRIM(a.TG004)) MtlItemNo                ,LTRIM(RTRIM(c.MB002)) MtlItemName          ,LTRIM(RTRIM(c.MB003)) MtlItemSpec 
                                    ,LTRIM(RTRIM(a.TG008)) InventoryNo              ,LTRIM(RTRIM(d.MC002)) InventoryName        ,ISNULL(LTRIM(RTRIM(e.NL002)),'') Location
                                    ,LTRIM(RTRIM(f.MC003)) SaveLocation             ,LTRIM(RTRIM(a.TG009)) Quantity             ,LTRIM(RTRIM(a.TG021)) ReQuantity
                                    ,LTRIM(RTRIM(a.TG025)) AvailableDate            ,LTRIM(RTRIM(a.TG026)) ReCheckDate
                                    ,LTRIM(RTRIM(a.TG043)) Type                     ,LTRIM(RTRIM(a.TG044)) FreeSpareQty         ,LTRIM(RTRIM(a.TG048)) ReFreeSpareQty
                                    ,LTRIM(RTRIM(a.TG010)) UomNo                    ,LTRIM(RTRIM(a.TG052)) AvailableQty         ,LTRIM(RTRIM(a.TG055)) ReAvailableQty
                                    ,LTRIM(RTRIM(a.TG053)) AvailableUomNo           ,LTRIM(RTRIM(a.TG042)) TaxRate              ,LTRIM(RTRIM(a.TG014)) FromOrderPrefix
                                    ,LTRIM(RTRIM(a.TG015)) FromOrderNo              ,LTRIM(RTRIM(a.TG016)) FromOrderSequence    ,LTRIM(RTRIM(a.TG017)) LotNumber
                                    ,'' CustomerMtlItemNo                           ,LTRIM(RTRIM(a.TG018)) ProjectNo            ,LTRIM(RTRIM(a.TG050)) BusinessItemNumber
                                    ,LTRIM(RTRIM(h.TC012)) CustomerPurchaseOrder    ,LTRIM(RTRIM(a.TG019)) Remarks              ,LTRIM(RTRIM(a.TG024)) CheckOutCode
                                    ,LTRIM(RTRIM(g.TD026)) DiscountRate             ,LTRIM(RTRIM(c.MB022)) LotManagement
                                    , (LTRIM(RTRIM(a.TG001)) + '-' + LTRIM(RTRIM(a.TG002)) + '-' + LTRIM(RTRIM(a.TG003))) TsnFullNo
                                    ,ROUND(LTRIM(RTRIM(a.TG012)) , CAST(i.MF003 AS INT))                        UnitPrice
                                    ,ROUND(LTRIM(RTRIM(a.TG013)) , CAST(i.MF004 AS INT))                        Amount
                                    ,ROUND(LTRIM(RTRIM(a.TG013)) , CAST(i.MF004 AS INT))                        OriginalAmount
                                    ,ROUND(LTRIM(RTRIM(a.TG013* a.TG042)) , CAST(i.MF004 AS INT))               OriginalTaxAmount
                                    ,ROUND(LTRIM(RTRIM(a.TG013 * b.TF012)) , CAST(i.MF004 AS INT))              PretaxAmount
                                    ,ROUND(LTRIM(RTRIM(a.TG013 * a.TG042 * b.TF012)) , CAST(i.MF004 AS INT))    TaxAmount
                                    ,(ISNULL(j.CanUseQty,0) - ISNULL(j1.UnconfirmedQty,0) - ISNULL(j2.UnconfirmedTsrnQty,0))                        CanUseQty
                                    ,(ISNULL(j.CanUseFreeSpareQty,0) - ISNULL(j1.UnconfirmedFreebieQty,0) - ISNULL(j2.UnconfirmedTsrnFsQty,0))      CanUseFreeSpareQty
                                    ,(ISNULL(j.CanUseAvailableQty,0) - ISNULL(j1.UnconfirmedAvailableQty,0) - ISNULL(j2.UnconfirmedTsrnAvQty,0))    CanUseAvailableQty
                                    ,ISNULL(j1.UnconfirmedQty,0)                                                 UnconfirmedQty
                                    ,ISNULL(j1.UnconfirmedFreebieQty,0)                                          UnconfirmedFreebieQty
                                    ,ISNULL(j1.UnconfirmedAvailableQty,0)                                        UnconfirmedAvailableQty
                                    ,ISNULL(j2.UnconfirmedTsrnQty,0)                                             UnconfirmedTsrnQty
                                    ,ISNULL(j2.UnconfirmedTsrnFsQty,0)                                           UnconfirmedTsrnFsQty
                                    ,ISNULL(j2.UnconfirmedTsrnAvQty,0)                                           UnconfirmedTsrnAvQty
                                    FROM INVTG a
                                    INNER JOIN INVTF b on a.TG001 = b.TF001 AND a.TG002 = b.TF002
                                    INNER JOIN INVMB c on a.TG004 = c.MB001
                                    INNER JOIN CMSMC d on a.TG008 = d.MC001
                                    LEFT JOIN  CMSNL e on d.MC001 = e.NL001 AND a.TG036 = e.NL002
                                    LEFT JOIN  INVMC f on a.TG004 = f.MC001 AND a.TG008 = f.MC002
                                    LEFT JOIN  COPTD g on a.TG014 = g.TD001 AND a.TG015 = g.TD002 AND a.TG016 = g.TD003
                                    LEFT JOIN  COPTC h on a.TG014 = h.TC001 AND a.TG015 = h.TC002
                                    INNER JOIN CMSMF i on b.TF011 = i.MF001
                                    OUTER APPLY(
                                        SELECT 
                                         (ISNULL(x1.TG009,0) - ISNULL(x1.TG020,0) - ISNULL(x1.TG021,0)) CanUseQty
                                        ,(ISNULL(x1.TG044,0) - ISNULL(x1.TG046,0) - ISNULL(x1.TG048,0)) CanUseFreeSpareQty 
                                        ,(ISNULL(x1.TG052,0) - ISNULL(x1.TG054,0) - ISNULL(x1.TG055,0)) CanUseAvailableQty
                                        FROM INVTG x1
                                        WHERE x1.TG001 = a.TG001
                                        AND x1.TG002 = a.TG002
                                        AND x1.TG003 = a.TG003
                                        AND x1.TG022 ='Y'
                                    ) j
                                    OUTER APPLY(
                                        SELECT SUM(ISNULL(x1.TH008,0)) UnconfirmedQty 
                                              ,SUM(ISNULL(x1.TH024,0)) UnconfirmedFreebieQty 
                                              ,SUM(ISNULL(x1.TH061,0)) UnconfirmedAvailableQty
                                        FROM COPTH x1
                                        WHERE x1.TH032 = a.TG001
                                        AND x1.TH033 = a.TG002
                                        AND x1.TH034 = a.TG003
                                        AND x1.TH020 ='N'
                                    ) j1
                                    OUTER APPLY(
                                        SELECT SUM(ISNULL(x1.TI009,0)) UnconfirmedTsrnQty 
                                                ,SUM(ISNULL(x1.TI035,0)) UnconfirmedTsrnFsQty 
                                                ,SUM(ISNULL(x1.TI038,0)) UnconfirmedTsrnAvQty 
                                        FROM INVTI x1
                                        WHERE 1=1
                                        AND x1.TI014 = a.TG001
                                        AND x1.TI015 = a.TG002
                                        AND x1.TI016 = a.TG003
                                        AND x1.TI022 ='N'
                                    ) j2
                                    WHERE a.TG001 = @SourcePrefix
                                    AND a.TG002 = @SourceNo
                                    AND b.TF020 = 'Y'
                                    AND a.TG024 = 'N'
                                    ";
                            if (SourceSequence.Length > 0)
                            {
                                sql += "AND a.TG003 = @SourceSequence";
                                dynamicParameters.Add("SourceSequence", SourceSequence);
                                //sql += "AND a.TG003 = '" + SourceSequence + @"'";
                            }
                            break;
                        #endregion

                        #region //銷貨單單頭
                        case "Ro":
                            sql = @"SELECT 
                                      LTRIM(RTRIM(TG005)) DepartmentNo          , LTRIM(RTRIM(TG006)) SalesmenNo 
                                    , LTRIM(RTRIM(TG035)) StaffNo               , LTRIM(RTRIM(TG026)) CollectionSalesmenNo
                                    , LTRIM(RTRIM(TG010)) Factory               , LTRIM(RTRIM(TG047)) PaymentTerm          
                                    , LTRIM(RTRIM(TG082)) TradingTerms          , LTRIM(RTRIM(TG011)) Currency
                                    ,     ISNULL(TG012,0) ExchangeRate          , LTRIM(RTRIM(TG062)) BondCode
                                    , LTRIM(RTRIM(TG020)) Remark                ,     ISNULL(TG032,0) RowCnt
                                    , LTRIM(RTRIM(TG094)) TaxCode               , LTRIM(RTRIM(TG016)) InvoiceType 
                                    , LTRIM(RTRIM(TG017)) TaxType               , LTRIM(RTRIM(TG015)) UiNo
                                    , LTRIM(RTRIM(TG098)) InvNumStart           , LTRIM(RTRIM(TG014)) InvNumEnd    
                                    , LTRIM(RTRIM(TG038)) ApplyYYMM             , LTRIM(RTRIM(TG031)) CustomsClearance    
                                    , LTRIM(RTRIM(TG007)) CustomerFullName      , LTRIM(RTRIM(TG083)) CustomerEgFullName 
                                    , LTRIM(RTRIM(TG097)) MultipleInvoices      , LTRIM(RTRIM(TG100)) MultiTaxRate  
                                    , LTRIM(RTRIM(TG066)) ContactPerson         , LTRIM(RTRIM(TG131)) ContactEmail
                                    , LTRIM(RTRIM(TG008)) CustomerAddressFirst  , LTRIM(RTRIM(TG009)) CustomerAddressSecond
                                    , LTRIM(RTRIM(TG027)) Remark1               , LTRIM(RTRIM(TG028)) Remark2
                                    , LTRIM(RTRIM(TG029)) Remark3               ,     ISNULL(TG044,0) TaxRate
                                    , CASE WHEN LEN(LTRIM(RTRIM(TG021))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TG021)) as date), 'yyyy-MM-dd') ELSE NULL END InvoiceDate
                                    , b.MF003 PriceDouble , b.MF004 AmountDouble
                                    FROM COPTG a
                                    INNER JOIN CMSMF b on a.TG011 = b.MF001
                                    WHERE 1=1
                                    AND a.TG001 =@SourcePrefix
                                    AND a.TG002 =@SourceNo
                                    AND a.TG023 = 'Y'
                                    ";
                            break;
                        #endregion
                        #region //銷貨單單身
                        case "RoDetail":
                            sql = @"SELECT 
                                      LTRIM(RTRIM(a.TH004)) MtlItemNo             , LTRIM(RTRIM(c.MB002)) MtlItemName 
                                    , LTRIM(RTRIM(c.MB003)) MtlItemSpec           , LTRIM(RTRIM(a.TH008)) Quantity
                                    , LTRIM(RTRIM(a.TH031)) FreeSpareType         , LTRIM(RTRIM(a.TH024)) FreeSpareQty
                                    , LTRIM(RTRIM(a.TH009)) UomNo                 , LTRIM(RTRIM(a.TH061)) AvailableQty
                                    , LTRIM(RTRIM(a.TH062)) AvailableUomNo        
                                    , LTRIM(RTRIM(a.TH073)) TaxRate               , LTRIM(RTRIM(a.TH025)) RoDiscountRate 
                                    , LTRIM(RTRIM(b.TD026)) SoDiscountRate        
                                    , LTRIM(RTRIM(a.TH007)) InventoryNo           , LTRIM(RTRIM(a.TH060)) [Location]
                                    , LTRIM(RTRIM(a.TH017)) LotNumber 
                                    , LTRIM(RTRIM(a.TH001)) RoErpPrefix           , LTRIM(RTRIM(a.TH002)) RoErpNo
                                    , LTRIM(RTRIM(a.TH003)) RoSequence    
                                    , LTRIM(RTRIM(a.TH014)) SoErpPrefix           , LTRIM(RTRIM(a.TH015)) SoErpNo 
                                    , LTRIM(RTRIM(a.TH016)) SoSequence
                                    , LTRIM(RTRIM(a.TH030)) ProjectCode           , LTRIM(RTRIM(a.TH030)) ProjectCode
                                    , LTRIM(RTRIM(a.TH018)) Remarks
                                    , LTRIM(RTRIM(a.TH027)) CheckOutPrefix        , LTRIM(RTRIM(a.TH028)) CheckOutNo
                                    , LTRIM(RTRIM(a.TH029)) CheckOutSequence  
                                    , LTRIM(RTRIM(a.TH042)) BondedCode            , LTRIM(RTRIM(a.TH019)) CustomerMtlItemNo    
                                    , ROUND(LTRIM(RTRIM(a.TH012)) , CAST(e.MF003 AS INT)) UnitPrice
                                    , ROUND(LTRIM(RTRIM(a.TH013)) , CAST(e.MF004 AS INT)) Amount
                                    , ROUND(LTRIM(RTRIM(a.TH035)) , CAST(e.MF004 AS INT)) OrigPreTaxAmt
                                    , ROUND(LTRIM(RTRIM(a.TH036)) , CAST(e.MF004 AS INT)) OrigTaxAmt
                                    , ROUND(LTRIM(RTRIM(a.TH037)) , CAST(e.MF004 AS INT)) PreTaxAmt
                                    , ROUND(LTRIM(RTRIM(a.TH038)) , CAST(e.MF004 AS INT)) TaxAmt
                                    , (LTRIM(RTRIM(a.TH001)) + '-' + LTRIM(RTRIM(a.TH002)) + '-' + LTRIM(RTRIM(a.TH003))) RoFullNo
                                    , (LTRIM(RTRIM(a.TH014)) + '-' + LTRIM(RTRIM(a.TH015)) + '-' + LTRIM(RTRIM(a.TH016))) SoFullNo
                                    , (LTRIM(RTRIM(a.TH027)) + '-' + LTRIM(RTRIM(a.TH028)) + '-' + LTRIM(RTRIM(a.TH029))) CoFullNo
                                    , c.MB022 LotManagement
                                    , (ISNULL(a.TH008,0) - ISNULL(a.TH043,0) - ISNULL(x.UnconfirmedQty,0)) CanUseQty
                                    , (ISNULL(a.TH024,0) - ISNULL(a.TH052,0) - ISNULL(x.UnconfirmedFreebieQty,0)) CanUseFreeSpareQty
                                    FROM COPTH a
                                    INNER JOIN COPTG a1 on a.TH001 = a1.TG001 AND a.TH002 = a1.TG002
                                    INNER JOIN COPTD b on a.TH014 = b.TD001 AND a.TH015 = b.TD002 AND a.TH016 = b.TD003
                                    INNER JOIN INVMB c on a.TH004 = c.MB001
                                    INNER JOIN CMSMC d on a.TH007 = d.MC001
                                    INNER JOIN CMSMF e on a1.TG011 = e.MF001
                                    OUTER APPLY(
                                        SELECT SUM(ISNULL(x1.TJ007,0)) UnconfirmedQty
                                              ,SUM(ISNULL(x1.TJ042,0)) UnconfirmedFreebieQty 
                                        FROM COPTJ x1
                                        WHERE x1.TJ015 = a.TH001
                                        AND x1.TJ016 = a.TH002
                                        AND x1.TJ017 = a.TH003
                                        AND x1.TJ021 ='N'
                                    ) x
                                    WHERE a.TH001 = @SourcePrefix
                                    AND a.TH002 = @SourceNo
                                    AND a.TH020 = 'Y'
                                    ";
                            break;
                            #endregion

                    }
                    dynamicParameters.Add("SourcePrefix", SourcePrefix);
                    dynamicParameters.Add("SourceNo", SourceNo);
                    
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

        #region //GetMtlItemUom -- 取得品號單位 -- Shintokuro 2024.02.07
        public string GetMtlItemUom(string MtlItemNo)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    #region //判斷品號是否存在
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT b.UomNo OreUomNo, 1 SwapDenominator , b.UomNo ChaUomNo, 1 SwapNumerator
                            FROM PDM.MtlItem a
                            INNER JOIN PDM.UnitOfMeasure b on a.InventoryUomId = b.UomId
                            WHERE a.MtlItemNo = @MtlItemNo
                            UNION
                            SELECT b1.UomNo OreUomNo,SwapDenominator,b2.UomNo ChaUomNo, SwapNumerator
                            FROM PDM.MtlItemUomCalculate a
                            INNER JOIN PDM.MtlItem a1 on a.MtlItemId = a1.MtlItemId
                            INNER JOIN PDM.UnitOfMeasure b1 on a.UomId = b1.UomId
                            INNER JOIN PDM.UnitOfMeasure b2 on a.ConvertUomId = b2.UomId
                            WHERE a1.MtlItemNo = @MtlItemNo
                            ";
                    dynamicParameters.Add("MtlItemNo", MtlItemNo);
                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0) throw new SystemException("【品號單位換算】－品號不存在，請重新輸入!");
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

        #region //GetReturnReceiveOrder -- 取得銷貨單資料 -- Shintokuro 2024.05.07
        public string GetReturnReceiveOrder(int RtId, string RtErpPrefix, string RtErpFullNo, int CustomerId, int SalesmenId
            , string StartDocDate, string EndDocDate, string ConfirmStatus, string TransferStatusMES
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.RtId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @"  ,FORMAT(a.DocDate, 'yyyy-MM-dd') DocDate, 
                             FORMAT(a.ReturnDate, 'yyyy-MM-dd') ReturnDate,
                             FORMAT(a.InvoiceDate, 'yyyy-MM-dd') InvoiceDate,
                            a.CompanyId, a.RtErpPrefix, a.RtErpNo, a.CustomerId,a.ProcessCode,a.RevenueJournalCode,a.CostJournalCode, 
                            a.OriginalAmount, a.OriginalTaxAmount, a.PretaxAmount, a.TaxAmount, a.TotalQuantity,
                            a.DepartmentId, a.SalesmenId, a.StaffId, a.CollectionSalesmenId,a.Factory,
                            a.PaymentTerm, a.TradingTerms, a.Currency, a.ExchangeRate,a.BondCode,
                            a.RowCnt, a.ConfirmStatus, a.ConfirmUserId, a.Remark,
                            a.TaxCode, a.InvoiceType, a.TaxType, a.TaxRate,a.UiNo,
                            a.InvNumStart, a.InvNumEnd, a.ApplyYYMM, a.CustomsClearance,a.CustomerFullName,
                            a.CustomerEgFullName, a.MultipleInvoices, a.MultiTaxRate, a.DebitNote,
                            a.ContactPerson, a.ContactEmail, a.CustomerAddressFirst, a.CustomerAddressSecond,a.Remark1,
                            a.Remark2, a.Remark3,a.TransferStatusMES,
                            (a.RtErpPrefix + '-' + a.RtErpNo) RtFullNo,
                            b.UserName SalesmenName,c.UserName ConfirmUserName,
                            b.UserName SalesmenName,c.UserName ConfirmUserName,
                            d.CustomerName
                            , (
                                SELECT 
                                  ISNULL(x1.RtSequence, '') RtSequence, ISNULL(x2.MtlItemNo, '') MtlItemNo, ISNULL(x2.MtlItemName, '') MtlItemName
                                , ISNULL(x3.InventoryNo, '') InventoryNo, ISNULL(x1.Quantity, '') Quantity, ISNULL(x1.Type, '') Type
                                , ISNULL(x1.FreeSpareQty, '') FreeSpareQty, ISNULL(x1.AvailableQty, '') AvailableQty, ISNULL(x1.UnitPrice, '') UnitPrice
                                , ISNULL(x1.Amount, '') Amount
                                , ISNULL(x5.SoErpPrefix+'-'+x5.SoErpNo+'-'+x4.SoSequence, '') SoFullNo
                                , ISNULL(x7.RoErpPrefix+'-'+x7.RoErpNo+'-'+x6.RoSequence, '') RoFullNo
                                , ISNULL(x1.LotNumber, '') LotNumber
                                FROM SCM.RtDetail x1
                                LEFT JOIN PDM.MtlItem x2 ON x1.MtlItemId = x2.MtlItemId
                                LEFT JOIN SCM.Inventory x3 ON x1.InventoryId = x3.InventoryId
                                LEFT JOIN SCM.SoDetail x4 ON x1.SoDetailId = x4.SoDetailId
                                LEFT JOIN SCM.SaleOrder x5 ON x4.SoId = x5.SoId
                                LEFT JOIN SCM.RoDetail x6 ON x1.RoDetailId = x6.RoDetailId
                                LEFT JOIN SCM.ReceiveOrder x7 ON x6.RoId = x7.RoId
                                WHERE x1.RtId = a.RtId
                                FOR JSON PATH, ROOT('data')
                            ) RoDetail
                            ";
                    sqlQuery.mainTables =
                        @"FROM SCM.ReturnReceiveOrder a
                          LEFT JOIN BAS.[User] b on a.SalesmenId = b.UserId
                          LEFT JOIN BAS.[User] c on a.ConfirmUserId = c.UserId
                          LEFT JOIN SCM.Customer d on a.CustomerId = d.CustomerId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RtId", @" AND a.RtId = @RtId", RtId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RtErpPrefix", @" AND a.RtErpPrefix = @RtErpPrefix", RtErpPrefix);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RtErpFullNo", @" AND (a.RtErpPrefix +'-' + a.RtErpNo ) LIKE '%' + @RtErpFullNo + '%'", RtErpFullNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerId", @" AND a.CustomerId = @CustomerId", CustomerId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SalesmenId", @" AND a.SalesmenId = @SalesmenId", SalesmenId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SalesmenId", @" AND a.SalesmenId = @SalesmenId", SalesmenId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ConfirmStatus", @" AND a.ConfirmStatus = @ConfirmStatus", ConfirmStatus);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TransferStatusMES", @" AND a.TransferStatusMES = @TransferStatusMES", TransferStatusMES);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.DocDate >= @StartDate ", StartDocDate.Length > 0 ? Convert.ToDateTime(StartDocDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.DocDate <= @EndDate ", ConfirmStatus.Length > 0 ? Convert.ToDateTime(ConfirmStatus).ToString("yyyy-MM-dd 23:59:59") : "");

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.RtId DESC";
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

        #region //GetRtDetail -- 取得銷退單單身資料 -- Shintokuro 2024.05.09
        public string GetRtDetail(int RtId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.RtDetailId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" ,a.RtDetailId, a.RtId, a.RtSequence, a.MtlItemId, a.Quantity, a.UomId,
                            a.UnitPrice, a.Amount, a.InventoryId, a.LotNumber, a.RoDetailId, a.SoDetailId,
                            a.ConfirmStatus, a.UpdateCode, a.Remarks, a.CheckOutCode, a.CheckOutPrefix, a.CheckOutNo,
                            a.CheckOutSeq, a.ProjectCode, a.CustomerMtlItemNo, a.Type, a.OrigPreTaxAmt, a.OrigTaxAmt,
                            a.PreTaxAmt, a.TaxAmt, a.PackageQty, a.PackageUomId, a.BondedCode, a.CommissionRate, a.CommissionAmount,
                            a.OriginalCustomer, a.FreeSpareType, a.FreeSpareQty, a.PackageFreeSpareQty, a.NotPayTemp, a.Location,
                            a.AvailableQty, a.AvailableUomId, a.TaxRate, a.TaxCode, a.DiscountAmount, a.TransferStatusMES,
                            b.MtlItemNo,b.MtlItemName,b.MtlItemSpec,b.LotManagement,
                            b.MtlItemSpec,
                            d.UomNo ,d1.UomNo AvailableUomNo,
                            c.InventoryNo,c.InventoryName,
                            e.DiscountRate SoDiscountRate,
                            f.DiscountRate RoDiscountRate,
                            (e1.SoErpPrefix + '-' + e1.SoErpNo + '-' + e.SoSequence ) SoFullNo,
                            (f1.RoErpPrefix + '-' + f1.RoErpNo + '-' + f.RoSequence ) RoFullNo";
                    sqlQuery.mainTables =
                        @"FROM SCM.RtDetail a
                          INNER JOIN SCM.ReturnReceiveOrder a1 on a.RtId = a1.RtId
                          INNER JOIN PDM.MtlItem b on a.MtlItemId = b.MtlItemId
                          INNER JOIN SCM.Inventory c on a.InventoryId = c.InventoryId
                          INNER JOIN PDM.UnitOfMeasure d on a.UomId = d.UomId
                          INNER JOIN PDM.UnitOfMeasure d1 on a.AvailableUomId = d1.UomId
                          LEFT JOIN SCM.SoDetail e on a.SoDetailId = e.SoDetailId
                          LEFT JOIN SCM.SaleOrder e1 on e.SoId = e1.SoId
                          LEFT JOIN SCM.RoDetail f on a.RoDetailId = f.RoDetailId
                          LEFT JOIN SCM.ReceiveOrder f1 on f.RoId = f1.RoId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND a1.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RtId", @" AND a.RtId = @RtId", RtId);

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.RtDetailId ASC";
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

        #region //GetReceiveOrder -- 取得銷貨單資料 -- Shintokuro 2024.02.01
        public string GetDeliveryOrder(int CustomerId, string SearchKey, string StartDate, string EndDate
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.DoDetailId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @"  , FORMAT(a1.DoDate, 'yyyy-MM-dd') DoDate,c.SoErpPrefix + '-' + c.SoErpNo +'-'+b.SoSequence SoFullNo,d.InventoryNo,e.MtlItemNo,e.MtlItemName,e.MtlItemSpec
                            ,a.DoQty ,a.FreebieQty DoFreebieQty, a.SpareQty DoSpareQty
                            ,b.SoQty,b.FreebieQty,b.SpareQty
                            ,x.RoQty, x.RoFreeSpareQty
                            ,y.PickQty
                            ,f.CustomerNo,f.CustomerShortName,g.UserNo,g.UserName
                            ";
                    sqlQuery.mainTables =
                        @"FROM SCM.DoDetail a
                          INNER JOIN SCM.DeliveryOrder a1 ON a.DoId = a1.DoId
                          INNER JOIN SCM.SoDetail b on a.SoDetailId = b.SoDetailId
                          INNER JOIN SCM.SaleOrder c on b.SoId = c.SoId
                          INNER JOIN SCM.Inventory d on b.InventoryId = d.InventoryId
                          INNER JOIN PDM.MtlItem e on b.MtlItemId = e.MtlItemId
                          INNER JOIN SCM.Customer f on c.CustomerId = f.CustomerId
                          INNER JOIN BAS.[User] g on c.SalesmenId = g.UserId
                          OUTER APPLY(
                              SELECT SUM(Quantity) RoQty,SUM(FreeSpareQty) RoFreeSpareQty
                              FROM SCM.RoDetail x1
                              WHERE b.SoDetailId = x1.SoDetailId 
                              AND x1.ConfirmationCode = 'Y'
                          ) x
                          OUTER APPLY(
                              SELECT SUM(ItemQty) PickQty
                              FROM SCM.PickingItem x1
                              WHERE a.DoId = x1.DoId
                          ) y
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND c.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey",
                        @" AND (g.MtlItemNo LIKE '%' + @SearchKey + '%' 
                            OR g.MtlItemName LIKE '%' + @SearchKey + '%'
                            OR f.SoErpPrefix + '-' + f.SoErpNo LIKE '%' + @SearchKey + '%')", SearchKey);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerId", @" AND c.CustomerId = @CustomerId", CustomerId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND b.DoDate >= @StartDate ", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND b.DoDate <= @EndDate ", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "FORMAT(a1.DoDate, 'yyyy-MM-dd') DESC, b.DcId, c.SoErpPrefix, c.SoErpNo, b.SoSequence";
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

        #region //Add
        #region //AddReceiveOrder -- 銷貨單新增 -- Shintokuro 2024.02.01
        public string AddReceiveOrder(int ViewCompanyId
            , string DocDate
            , string ReceiveDate, string RoErpPrefix, string RoErpNo, int CustomerId, string ProcessCode
            , string RevenueJournalCode, string CostJournalCode, string NoCredLimitControl, string CashSales
            , string SourceType, string PriceSourceTypeMain, int SourceOrderId, string SourceFull
            , int DepartmentId
            , double OriginalAmount, double OriginalTaxAmount, double TotalQuantity, double PretaxAmount, double TaxAmount
            , double TaxCurrencyRate
            , int SalesmenId, string Factory, string PaymentTerm, string TradingTerms, string Currency
            , double ExchangeRate, string BondCode, int RowCnt, string Remark
            , string CustomerFullName, string ContactPerson, string CustomerAddressFirst, string TelephoneNumber
            , string CustomerAddressSecond, string FaxNumber
            , string TaxCode
            , string InvoiceType, string TaxType, double TaxRate, string InvNumGenMethod, string UiNo
            , string InvoiceDate, string InvoiceTime, string InvNumStart, string InvNumEnd, string ApplyYYMM, string CustomsClearance
            , string CustomerFullNameOre, string InvoiceAddressFirst, string CustomerEgFullName, string InvoiceAddressSecond
            , string MultipleInvoices, string MultiTaxRate, string AttachInvWithShip, string InvoicesVoid
            , string VehicleTypeNumber
            , string InvDonationRecipient, string VehicleIDshow, string VehicleIDhide, string CreditCard4No, string InvRandomCode, string ContactEmail
            , int StaffId
            , int CollectionSalesmenId, string Remark1, string Remark2, string Remark3, string LCNO, string INVOICENO
            , string DeclarationNumber, string NewRoFull, string ShipNoticeFull, string ChangeInvCode
            , string DepositBatches
            , string SoFull, string AdvOrderFull, double OffsetAmount, double OffsetTaxAmount, string TransportMethod
            , string DispatchOrderFull, string CarNumber, string TrainNumber, string DeliveryUser
            , string Courier, string SiteCommCalcMethod, string SiteCommRate, string TotalCommAmount
            , string RoDetailData
            )
        {
            try
            {
                if (RoErpPrefix.Length <= 0) throw new SystemException("【銷貨單別】不能為空!");
                if (ReceiveDate.Length <= 0) throw new SystemException("【銷貨日期】不能為空!");
                if (CustomerId <= 0) throw new SystemException("【客戶】不能為空!");
                if (Factory.Length <= 0) throw new SystemException("【廠別】不能為空!");
                if (Currency.Length <= 0) throw new SystemException("【幣別】不能為空!");
                if (ExchangeRate < 0) throw new SystemException("【匯率】不能為負!");
                if (ApplyYYMM.Length < 0) throw new SystemException("【申報年月】不能為空!");
                if (DocDate.Length < 0) throw new SystemException("【單據日期】不能為空!");
                if (TaxRate < 0) throw new SystemException("【營業稅率】不能為負!");
                ApplyYYMM = ApplyYYMM.Replace("-", "");
                if (CurrentCompany != ViewCompanyId) throw new SystemException("頁面的公司別與後端公司別不同，請嘗試重新登入!!");

                List<Dictionary<string, string>> RoDetailJsonList = JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(RoDetailData);
                int? nullData = null;
                int rowsAffected = 0;
                int RoId = 0;
                string SoErpPrefix = "";
                string SoErpNo = "";
                string SoSequence = "";
                string MESCompanyNo = "";
                string userNo = "";
                string userName = "";
                string departmentNo = "";
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷公司資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 CompanyId, CompanyName, ErpDb,ErpNo
                            FROM BAS.Company
                            WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【公司】資料錯誤!");

                        foreach (var item in result)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            MESCompanyNo = item.ErpNo;
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
                        dynamicParameters.Add("FunctionCode", "ReceiveOrderManagment");
                        dynamicParameters.Add("DetailCode", "add");

                        var resultUser = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUser.Count() <= 0) throw new SystemException("使用者資料錯誤!");

                        foreach (var item in resultUser)
                        {
                            userNo = item.UserNo;
                            userName = item.UserName;
                            departmentNo = item.DepartmentNo;
                        }
                        #endregion

                        #region //判斷客戶資料是否存在
                        sql = @"SELECT TOP 1 1
                                FROM SCM.Customer
                                WHERE CompanyId = @CompanyId
                                AND CustomerId = @CustomerId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("CustomerId", CustomerId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("客戶資料找不到，請重新確認!!");
                        #endregion

                        #region //判斷部門資料是否存在
                        if (DepartmentId > 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.Department
                                    WHERE CompanyId = @CompanyId
                                    AND DepartmentId = @DepartmentId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("DepartmentId", DepartmentId);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("部門資料找不到，請重新確認!!");
                        }
                        #endregion

                        #region //判斷業務資料是否存在
                        if (SalesmenId > 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.[User]
                                    WHERE UserId = @UserId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("UserId", SalesmenId);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("業務資料找不到，請重新確認!!");
                        }
                        #endregion

                        #region //判斷收款業務資料是否存在
                        if (CollectionSalesmenId > 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                FROM BAS.[User]
                                WHERE UserId = @UserId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("UserId", CollectionSalesmenId);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("收款業務資料找不到，請重新確認!!");
                        }
                        #endregion

                        #region //判斷員工資料是否存在
                        if (StaffId > 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                FROM BAS.[User]
                                WHERE UserId = @UserId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("UserId", StaffId);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("員工資料找不到，請重新確認!!");
                        }
                        #endregion

                        #region //判斷銷貨單單別+單號是否重複
                        RoErpNo = BaseHelper.RandomCode(11);
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.ReceiveOrder
                                WHERE RoErpPrefix = @RoErpPrefix
                                AND RoErpNo = @RoErpNo
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("RoErpPrefix", RoErpPrefix);
                        dynamicParameters.Add("RoErpNo", RoErpNo);
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【銷貨單新增】－【單別+單號】重複，請重新輸入!");
                        #endregion

                        #region //判斷來源單據是否存在
                        switch (SourceType)
                        {
                            case "So":
                                if(CurrentCompany != 3) throw new SystemException("非紘立公司不可以使用訂單轉銷貨");
                                #region //判斷訂單
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM SCM.SaleOrder
                                        WHERE CompanyId = @CompanyId
                                        AND SoId = @SourceOrderId";
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                dynamicParameters.Add("SourceOrderId", SourceOrderId);
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("【銷貨單新增】－找不到來源訂單單據，請重新確認!");
                                #endregion
                                break;
                            case "To":
                                #region //判斷暫出單
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM SCM.TempShippingNote
                                        WHERE CompanyId = @CompanyId
                                        AND TsnId = @SourceOrderId";
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                dynamicParameters.Add("SourceOrderId", SourceOrderId);
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("【銷貨單新增】－找不到來源暫出單單據，請重新確認!");
                                #endregion
                                break;
                        }
                        #endregion

                        #region //新增單頭資料表
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.ReceiveOrder (CompanyId, RoErpPrefix, RoErpNo, ReceiveDate, CustomerId, DepartmentId, SalesmenId, 
                                CustomerFullName, CustomerAddressFirst, CustomerAddressSecond, Factory, Currency, ExchangeRate, 
                                OriginalAmount, InvNumEnd, UiNo, InvoiceType, TaxType, InvoiceAddressFirst, InvoiceAddressSecond, Remark, 
                                InvoiceDate, PrintCount, ConfirmStatus, UpdateCode, OriginalTaxAmount, CollectionSalesmenId, Remark1, Remark2, Remark3, InvoicesVoid, 
                                CustomsClearance, RowCnt, TotalQuantity, CashSales, StaffId, RevenueJournalCode, CostJournalCode, ApplyYYMM, 
                                LCNO, INVOICENO, InvoicesPrintCount, DocDate, ConfirmUserId, TaxRate, PretaxAmount, TaxAmount, PaymentTerm, SoDetailId, 
                                AdvOrderPrefix, AdvOrderNo, OffsetAmount, OffsetTaxAmount, TotalPackages, SignatureStatus, ChangeInvCode, 
                                NewRoPrefix, NewRoNo, TransferStatusERP, ProcessCode, AttachInvWithShip, BondCode, TransmissionCount, 
                                Invoicer, InvCode, ContactPerson, Courier, SiteCommCalcMethod, SiteCommRate, CommCalcInclTax, 
                                TotalCommAmount, TransportMethod, DispatchOrderPrefix, DispatchOrderNo, DeclarationNumber, 
                                FullNameOfDelivCustomer, RoPriceType, TelephoneNumber, FaxNumber, ShipNoticePrefix, ShipNoticeNo, 
                                TradingTerms, CustomerEgFullName, InvNumGenMethod, DocSourceCode, NoCredLimitControl, 
                                InstallmentSettlement, InstallmentCount, AutoApportionByInstallment, StartYearMonth, TaxCode, CustomsManual, 
                                RemarkCode, MultipleInvoices, InvNumStart, NumberOfInvoices, MultiTaxRate, TaxCurrencyRate, 
                                VoidDate, VoidApprovalDocNum, VoidReason, Source, IncomeDraftID, IncomeDraftSeq, IncomeVoucherType, 
                                IncomeVoucherNumber, Status, ZeroTaxForBuyer, GenLedgerAcctType, InvoiceTime, InvCode2, InvSymbol, 
                                DeliveryCountry, VehicleIDshow, VehicleTypeNumber, VehicleIDhide, InvDonationRecipient, InvRandomCode, 
                                ReservedField, CreditCard4No, ContactEmail, ExpectedReceiptDate, OrigInvNumber, TransferStatusMES, 
                                TransferTime, ConfirmTime, SourceType, PriceSourceTypeMain, SourceOrderId, SourceFull,
                                CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.RoId
                                VALUES ( @CompanyId, @RoErpPrefix, @RoErpNo, @ReceiveDate, @CustomerId, @DepartmentId, @SalesmenId, 
                                @CustomerFullName, @CustomerAddressFirst, @CustomerAddressSecond, @Factory, @Currency, @ExchangeRate, 
                                @OriginalAmount, @InvNumEnd, @UiNo, @InvoiceType, @TaxType, @InvoiceAddressFirst, @InvoiceAddressSecond, @Remark, 
                                @InvoiceDate, @PrintCount, @ConfirmStatus, @UpdateCode, @OriginalTaxAmount, @CollectionSalesmenId, @Remark1, @Remark2, @Remark3, @InvoicesVoid, 
                                @CustomsClearance, @RowCnt, @TotalQuantity, @CashSales, @StaffId, @RevenueJournalCode, @CostJournalCode, @ApplyYYMM, 
                                @LCNO, @INVOICENO, @InvoicesPrintCount, @DocDate, @ConfirmUserId, @TaxRate, @PretaxAmount, @TaxAmount, @PaymentTerm, @SoDetailId, 
                                @AdvOrderPrefix, @AdvOrderNo, @OffsetAmount, @OffsetTaxAmount, @TotalPackages, @SignatureStatus, @ChangeInvCode, 
                                @NewRoPrefix, @NewRoNo, @TransferStatusERP, @ProcessCode, @AttachInvWithShip, @BondCode, @TransmissionCount, 
                                @Invoicer, @InvCode, @ContactPerson, @Courier, @SiteCommCalcMethod, @SiteCommRate, @CommCalcInclTax, 
                                @TotalCommAmount, @TransportMethod, @DispatchOrderPrefix, @DispatchOrderNo, @DeclarationNumber, 
                                @FullNameOfDelivCustomer, @RoPriceType, @TelephoneNumber, @FaxNumber, @ShipNoticePrefix, @ShipNoticeNo, 
                                @TradingTerms, @CustomerEgFullName, @InvNumGenMethod, @DocSourceCode, @NoCredLimitControl, 
                                @InstallmentSettlement, @InstallmentCount, @AutoApportionByInstallment, @StartYearMonth, @TaxCode, @CustomsManual, 
                                @RemarkCode, @MultipleInvoices, @InvNumStart, @NumberOfInvoices, @MultiTaxRate, @TaxCurrencyRate, 
                                @VoidDate, @VoidApprovalDocNum, @VoidReason, @Source, @IncomeDraftID, @IncomeDraftSeq, @IncomeVoucherType, 
                                @IncomeVoucherNumber, @Status, @ZeroTaxForBuyer, @GenLedgerAcctType, @InvoiceTime, @InvCode2, @InvSymbol, 
                                @DeliveryCountry, @VehicleIDshow, @VehicleTypeNumber, @VehicleIDhide, @InvDonationRecipient, @InvRandomCode, 
                                @ReservedField, @CreditCard4No, @ContactEmail, @ExpectedReceiptDate, @OrigInvNumber, @TransferStatusMES, 
                                @TransferTime, @ConfirmTime, @SourceType, @PriceSourceTypeMain, @SourceOrderId, @SourceFull, 
                                @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = ViewCompanyId,
                                RoErpPrefix,
                                RoErpNo,
                                ReceiveDate,
                                CustomerId,
                                DepartmentId,
                                SalesmenId,
                                CustomerFullName,
                                CustomerAddressFirst,
                                CustomerAddressSecond,
                                Factory,
                                Currency,
                                ExchangeRate,
                                OriginalAmount,
                                UiNo,
                                InvoiceType,
                                TaxType,
                                InvoiceAddressFirst,
                                InvoiceAddressSecond,
                                Remark,
                                InvoiceDate = InvoiceDate != "" ? InvoiceDate : null,
                                PrintCount = 0,
                                ConfirmStatus = "N",
                                UpdateCode = "N",
                                OriginalTaxAmount,
                                CollectionSalesmenId,
                                Remark1,
                                Remark2,
                                Remark3,
                                InvoicesVoid,
                                CustomsClearance,
                                RowCnt,
                                TotalQuantity,
                                CashSales,
                                StaffId,
                                RevenueJournalCode,
                                CostJournalCode,
                                ApplyYYMM,
                                LCNO,
                                INVOICENO,
                                InvoicesPrintCount = 0,
                                DocDate,
                                ConfirmUserId = nullData,
                                TaxRate,
                                PretaxAmount,
                                TaxAmount,
                                PaymentTerm,
                                SoDetailId = nullData,
                                AdvOrderPrefix = "",
                                AdvOrderNo = "",
                                OffsetAmount,
                                OffsetTaxAmount,
                                TotalPackages = 0,
                                SignatureStatus = "N",
                                ChangeInvCode,
                                NewRoPrefix = "",
                                NewRoNo = "",
                                TransferStatusERP = "N",
                                ProcessCode,
                                AttachInvWithShip,
                                BondCode,
                                TransmissionCount = 0,
                                Invoicer = "",
                                InvCode = "",
                                ContactPerson,
                                Courier,
                                SiteCommCalcMethod = 1,
                                SiteCommRate,
                                CommCalcInclTax = "N",
                                TotalCommAmount,
                                TransportMethod,
                                DispatchOrderPrefix = "",
                                DispatchOrderNo = "",
                                DeclarationNumber,
                                FullNameOfDelivCustomer = CustomerFullName,
                                RoPriceType = "",
                                TelephoneNumber,
                                FaxNumber,
                                ShipNoticePrefix = "",
                                ShipNoticeNo = "",
                                TradingTerms,
                                CustomerEgFullName,
                                InvNumGenMethod,
                                DocSourceCode = "",
                                NoCredLimitControl,
                                InstallmentSettlement = "",
                                InstallmentCount = 0,
                                AutoApportionByInstallment = "",
                                StartYearMonth = "",
                                TaxCode,
                                CustomsManual = "",
                                RemarkCode = "",
                                MultipleInvoices,
                                InvNumStart,
                                InvNumEnd,
                                NumberOfInvoices = 0,
                                MultiTaxRate,
                                TaxCurrencyRate = 1,
                                VoidDate = "",
                                VoidApprovalDocNum = "",
                                VoidReason = "",
                                Source = "",
                                IncomeDraftID = "",
                                IncomeDraftSeq = "",
                                IncomeVoucherType = "",
                                IncomeVoucherNumber = "",
                                Status = "0",
                                ZeroTaxForBuyer = "0",
                                GenLedgerAcctType = "",
                                InvoiceTime,
                                InvCode2 = "",
                                InvSymbol = "",
                                DeliveryCountry = "",
                                VehicleIDshow,
                                VehicleTypeNumber,
                                VehicleIDhide,
                                InvDonationRecipient,
                                InvRandomCode,
                                ReservedField = "",
                                CreditCard4No,
                                ContactEmail,
                                ExpectedReceiptDate = "",
                                OrigInvNumber = "",
                                TransferStatusMES = "N",
                                TransferTime = nullData,
                                ConfirmTime = nullData,
                                SourceType,
                                PriceSourceTypeMain,
                                SourceOrderId,
                                SourceFull,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);
                        rowsAffected += insertResult.Count();
                        #endregion

                        foreach (var item in insertResult)
                        {
                            RoId = item.RoId;
                        }

                        if (RoDetailJsonList.Count() > 0)
                        {
                            foreach (var item in RoDetailJsonList)
                            {
                                string MtlItemNo = item["MtlItemNo"] != null ? item["MtlItemNo"].ToString() : throw new SystemException("【資料維護不完整】品號欄位資料不可以為空,請重新確認~~");
                                string InventoryNo = item["InventoryNo"] != null ? item["InventoryNo"].ToString() : throw new SystemException("【資料維護不完整】庫別欄位資料不可以為空,請重新確認~~");
                                string Location = item["Location"] != null ? item["Location"].ToString() : "";
                                Double Quantity = item["Quantity"] != null ? Convert.ToDouble(item["Quantity"].ToString())  : 0;
                                string Type = item["Type"] != null ? item["Type"].ToString() : throw new SystemException("【資料維護不完整】類別欄位資料不可以為空,請重新確認~~");
                                Double FreeSpareQty = item["FreeSpareQty"] != null ? Convert.ToDouble(item["FreeSpareQty"].ToString()) : 0;
                                string UomNo = item["UomNo"] != null ? item["UomNo"].ToString() : throw new SystemException("【資料維護不完整】單位欄位資料不可以為空,請重新確認~~");
                                Double AvailableQty = item["AvailableQty"] != null ? Convert.ToDouble(item["AvailableQty"].ToString()) : 0;
                                string AvailableUomNo = item["AvailableUomNo"] != null ? item["AvailableUomNo"].ToString() : throw new SystemException("【資料維護不完整】計價單位欄位資料不可以為空,請重新確認~~");
                                Double UnitPrice = item["UnitPrice"] != null ? Convert.ToDouble(item["UnitPrice"].ToString()) : 0;
                                Double TaxRateDetail = item["TaxRate"] != null ? Convert.ToDouble(item["TaxRate"].ToString()) : 0;
                                Double DiscountRate = item["DiscountRate"] != null ? Convert.ToDouble(item["DiscountRate"].ToString()) : 0;
                                Double Amount = item["Amount"] != null ? Convert.ToDouble(item["Amount"].ToString()) : 0;
                                Double OrigPreTaxAmt = item["OrigPreTaxAmt"] != null ? Convert.ToDouble(item["OrigPreTaxAmt"].ToString()) : 0;
                                Double OrigTaxAmt = item["OrigTaxAmt"] != null ? Convert.ToDouble(item["OrigTaxAmt"].ToString()) : 0;
                                Double PreTaxAmt = item["PreTaxAmt"] != null ? Convert.ToDouble(item["PreTaxAmt"].ToString()) : 0;
                                Double TaxAmt = item["TaxAmt"] != null ? Convert.ToDouble(item["TaxAmt"].ToString()) : 0;
                                string SoFullDetail = item["SoFull"] != null ? item["SoFull"].ToString() : "";
                                string LotNumber = item["LotNumber"] != null ? item["LotNumber"].ToString() : "";
                                string CustomerMtlItemNo = item["CustomerMtlItemNo"] != null ? item["CustomerMtlItemNo"].ToString() : "";
                                string ProjectCode = item["ProjectCode"] != null ? item["ProjectCode"].ToString() : "";
                                string Remarks = item["Remarks"] != null ? item["Remarks"].ToString() : "";
                                string TsnFull = item["TsnFull"] != null ? item["TsnFull"].ToString() : "";
                                string BondedCode = item["BondedCode"] != null ? item["BondedCode"].ToString() : "";
                                string CustomerPurchaseOrder = item["CustomerPurchaseOrder"] != null ? item["CustomerPurchaseOrder"].ToString() : "";
                                string ForecastFull = item["ForecastFull"] != null ? item["ForecastFull"].ToString() : "";

                                #region //匯率等於1時判定
                                if (ExchangeRate == 1)
                                {
                                    if (OrigPreTaxAmt != PreTaxAmt) throw new SystemException("單據幣別與本國貨幣相同時,原幣金額與本幣金額須相同!!");
                                    if (OrigTaxAmt != TaxAmt) throw new SystemException("單據幣別與本國貨幣相同時,原幣稅額與本幣稅額須相同!!");
                                }
                                #endregion

                                #region //判斷品號是否存在
                                int MtlItemId = -1;
                                string OverDeliveryManagement = "";
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 MtlItemId,OverDeliveryManagement,LotManagement
                                        FROM PDM.MtlItem
                                        WHERE CompanyId = @CompanyId
                                        AND MtlItemNo = @MtlItemNo";
                                dynamicParameters.Add("MtlItemNo", MtlItemNo);
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("【品號:"+ MtlItemNo + "】找不到，請重新輸入!");
                                foreach(var item2 in result)
                                {
                                    MtlItemId = item2.MtlItemId;
                                    OverDeliveryManagement = item2.OverDeliveryManagement;
                                    if(item2.LotManagement != "N" && LotNumber == "") throw new SystemException("【品號:" + MtlItemNo + "】須維護批號，請重新輸入!");
                                }
                                #endregion

                                #region //判斷庫別是否存在
                                int InventoryId = -1;
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 InventoryId
                                        FROM SCM.Inventory
                                        WHERE CompanyId = @CompanyId
                                        AND InventoryNo = @InventoryNo";
                                dynamicParameters.Add("InventoryNo", InventoryNo.Split(':')[0]);
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("【庫別:" + InventoryNo + "】找不到，請重新輸入!");
                                foreach (var item2 in result)
                                {
                                    InventoryId = item2.InventoryId;
                                }
                                #endregion

                                #region //判斷單位是否存在
                                int UomId = -1;
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 UomId
                                        FROM PDM.UnitOfMeasure
                                        WHERE CompanyId = @CompanyId
                                        AND UomNo = @UomNo";
                                dynamicParameters.Add("UomNo", UomNo);
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("【單位:" + UomNo + "】找不到，請重新輸入!");
                                foreach (var item2 in result)
                                {
                                    UomId = item2.UomId;
                                }
                                #endregion

                                #region //判斷計價單位是否存在
                                int AvailableUomId = -1;
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 UomId
                                        FROM PDM.UnitOfMeasure
                                        WHERE CompanyId = @CompanyId
                                        AND UomNo = @UomNo";
                                dynamicParameters.Add("UomNo", AvailableUomNo);
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("【計價單位:" + AvailableUomNo + "】找不到，請重新輸入!");
                                foreach (var item2 in result)
                                {
                                    AvailableUomId = item2.UomId;
                                }
                                #endregion

                                #region //判斷訂單單身是否存在
                                int? SoDetailId = null;
                                string ProductType = "";
                                Double FreebieQty = 0;
                                Double SpareQty = 0;
                                if ( SoFullDetail.Length > 0)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 a.SoDetailId, a.ProductType, a.FreebieQty, a.SpareQty
                                            ,b.SoErpPrefix, b.SoErpNo, a.SoSequence, a.ClosureStatus
                                            FROM SCM.SoDetail a
                                            INNER JOIN SCM.SaleOrder b on a.SoId = b.SoId
                                            WHERE b.CompanyId = @CompanyId
                                            AND b.SoErpPrefix = @SoErpPrefix
                                            AND b.SoErpNo = @SoErpNo
                                            AND a.SoSequence = @SoSequence";
                                    dynamicParameters.Add("CompanyId", CurrentCompany);
                                    dynamicParameters.Add("SoErpPrefix", SoFullDetail.Split('-')[0]);
                                    dynamicParameters.Add("SoErpNo", SoFullDetail.Split('-')[1]);
                                    dynamicParameters.Add("SoSequence", SoFullDetail.Split('-')[2]);
                                    result = sqlConnection.Query(sql, dynamicParameters);
                                    if (result.Count() <= 0) throw new SystemException("【訂單:" + SoFullDetail + "】找不到，請重新輸入!");
                                    foreach (var item2 in result)
                                    {
                                        if(item2.ClosureStatus == "Y") throw new SystemException("【訂單:" + SoFullDetail + "】已經結案!!!");
                                        SoErpPrefix = item2.SoErpPrefix;
                                        SoErpNo = item2.SoErpNo;
                                        SoSequence = item2.SoSequence;
                                        SoDetailId = item2.SoDetailId;
                                        ProductType = item2.ProductType;
                                        FreebieQty = item2.FreebieQty;
                                        SpareQty = item2.SpareQty;

                                    }
                                }
                                #endregion

                                #region //判斷暫出單身是否存在
                                int? TsnDetailId = null;
                                string TsnProductType = "";
                                Double FreebieOrSpareQty = 0;
                                if(TsnFull.Length > 0)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 b.TsnDetailId,b.ProductType,b.FreebieOrSpareQty,b.ClosureStatus
                                            FROM SCM.TempShippingNote a
                                            INNER JOIN SCM.TsnDetail b on a.TsnId = b.TsnId
                                            WHERE CompanyId = @CompanyId
                                            AND a.TsnErpPrefix = @TsnErpPrefix
                                            AND a.TsnErpNo = @TsnErpNo
                                            AND b.TsnSequence = @TsnSequence
                                            AND a.ConfirmStatus = 'Y'";
                                    dynamicParameters.Add("CompanyId", CurrentCompany);
                                    dynamicParameters.Add("TsnErpPrefix", TsnFull.Split('-')[0]);
                                    dynamicParameters.Add("TsnErpNo", TsnFull.Split('-')[1]);
                                    dynamicParameters.Add("TsnSequence", TsnFull.Split('-')[2]);
                                    result = sqlConnection.Query(sql, dynamicParameters);
                                    if (result.Count() <= 0) throw new SystemException("【暫出單:" + TsnFull + "】找不到，請重新輸入!");
                                    foreach (var item2 in result)
                                    {
                                        if(item2.ClosureStatus == "Y") throw new SystemException("【暫出單:" + TsnFull + "】已經結案!!!");
                                        TsnDetailId = item2.TsnDetailId;
                                        TsnProductType = item2.ProductType;
                                        FreebieOrSpareQty = item2.FreebieOrSpareQty;
                                    }
                                }
                                #endregion

                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO SCM.RoDetail (RoId, RoSequence, MtlItemId, InventoryId, Quantity, UomId, UnitPrice, Amount, 
                                        SoDetailId, LotNumber, Remarks, CustomerMtlItemNo, ConfirmationCode, UpdateCode, FreeSpareQty, DiscountRate, 
                                        CheckOutCode, CheckOutPrefix, CheckOutNo, CheckOutSequence, ProjectCode, Type, TsnDetailId, OrigPreTaxAmt, 
                                        OrigTaxAmt, PreTaxAmt, TaxAmt, PackageQty, PackageFreeSpareQty, PackageUomId, BondedCode, SrQty, 
                                        SrPackageQty, SrOrigPreTaxAmt, SrOrigTaxAmt, SrPreTaxAmt, SrTaxAmt, CommissionRate, CommissionAmount, 
                                        OriginalCustomer, SrFreeSpareQty, SrPackageFreeSpareQty, NotPayTemp, ProductSerialNumberQty, ForecastCode, 
                                        ForecastSequence, Location, AvailableQty, AvailableUomId, MultiBatch, FreeSpareRate, FinalCustomerCode, 
                                        ReferenceQty, ReferencePackageQty, TaxRate, CRMSource, CRMPrefix, CRMNo, CRMSequence, CRMContractCode, 
                                        CRMAllowDeduction, CRMDeductionQty, CRMDeductionUnit, DebitAccount, CreditAccount, TaxAmountAccount, 
                                        BusinessItemNumber, TaxCode, DiscountAmount, K2NO, MarkingBINRecord, MarkingManagement, 
                                        BillingUnitInPackage, DATECODE, TransferStatusMES, CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        VALUES ( @RoId, @RoSequence, @MtlItemId, @InventoryId, @Quantity, @UomId, @UnitPrice, @Amount, 
                                        @SoDetailId, @LotNumber, @Remarks, @CustomerMtlItemNo, @ConfirmationCode, @UpdateCode, @FreeSpareQty, @DiscountRate, 
                                        @CheckOutCode, @CheckOutPrefix, @CheckOutNo, @CheckOutSequence, @ProjectCode, @Type, @TsnDetailId, @OrigPreTaxAmt, 
                                        @OrigTaxAmt, @PreTaxAmt, @TaxAmt, @PackageQty, @PackageFreeSpareQty, @PackageUomId, @BondedCode, @SrQty, 
                                        @SrPackageQty, @SrOrigPreTaxAmt, @SrOrigTaxAmt, @SrPreTaxAmt, @SrTaxAmt, @CommissionRate, @CommissionAmount, 
                                        @OriginalCustomer, @SrFreeSpareQty, @SrPackageFreeSpareQty, @NotPayTemp, @ProductSerialNumberQty, @ForecastCode, 
                                        @ForecastSequence, @Location, @AvailableQty, @AvailableUomId, @MultiBatch, @FreeSpareRate, @FinalCustomerCode, 
                                        @ReferenceQty, @ReferencePackageQty, @TaxRate, @CRMSource, @CRMPrefix, @CRMNo, @CRMSequence, @CRMContractCode, 
                                        @CRMAllowDeduction, @CRMDeductionQty, @CRMDeductionUnit, @DebitAccount, @CreditAccount, @TaxAmountAccount, 
                                        @BusinessItemNumber, @TaxCode, @DiscountAmount, @K2NO, @MarkingBINRecord, @MarkingManagement, 
                                        @BillingUnitInPackage, @DATECODE, @TransferStatusMES, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy
                                        )";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        RoId,
                                        RoSequence = "",
                                        MtlItemId,
                                        InventoryId,
                                        Quantity,
                                        UomId,
                                        UnitPrice,
                                        Amount,
                                        SoDetailId,
                                        LotNumber,
                                        Remarks,
                                        CustomerMtlItemNo,
                                        ConfirmationCode = "N",
                                        UpdateCode = "N",
                                        FreeSpareQty,
                                        DiscountRate,
                                        CheckOutCode = "N",
                                        CheckOutPrefix = "",
                                        CheckOutNo = "",
                                        CheckOutSequence = "",
                                        ProjectCode,
                                        Type = Type.Split('.')[0],
                                        TsnDetailId,
                                        OrigPreTaxAmt,
                                        OrigTaxAmt,
                                        PreTaxAmt,
                                        TaxAmt,
                                        PackageQty = 0,
                                        PackageFreeSpareQty = 0,
                                        PackageUomId = "",
                                        BondedCode,
                                        SrQty = 0,
                                        SrPackageQty = 0,
                                        SrOrigPreTaxAmt = 0,
                                        SrOrigTaxAmt = 0,
                                        SrPreTaxAmt = 0,
                                        SrTaxAmt = 0,
                                        CommissionRate = 0,
                                        CommissionAmount = 0,
                                        OriginalCustomer = "",
                                        SrFreeSpareQty =0,
                                        SrPackageFreeSpareQty = 0,
                                        NotPayTemp = "N",
                                        ProductSerialNumberQty = 0,
                                        ForecastCode = "",
                                        ForecastSequence = "",
                                        Location = Location.Split(':')[0],
                                        AvailableQty,
                                        AvailableUomId,
                                        MultiBatch = "N",
                                        FreeSpareRate =0,
                                        FinalCustomerCode = "",
                                        ReferenceQty = 0,
                                        ReferencePackageQty =0,
                                        TaxRate,
                                        CRMSource = "",
                                        CRMPrefix = "",
                                        CRMNo = "",
                                        CRMSequence = "",
                                        CRMContractCode = "",
                                        CRMAllowDeduction = "",
                                        CRMDeductionQty = 0,
                                        CRMDeductionUnit = "",
                                        DebitAccount = "",
                                        CreditAccount = "",
                                        TaxAmountAccount = "",
                                        BusinessItemNumber = "",
                                        TaxCode = TaxCode,
                                        DiscountAmount =0,
                                        K2NO = "",
                                        MarkingBINRecord = "",
                                        MarkingManagement = "",
                                        BillingUnitInPackage = "",
                                        DATECODE = "",
                                        TransferStatusMES = "N",
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();
                            }
                        }
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        #region //ERP公司代碼取得
                        string ERPCompanyNo = "";
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 ML001  
                                         FROM CMSML";
                        var resultCompanyNo = sqlConnection.Query(sql, dynamicParameters);
                        foreach (var item in resultCompanyNo)
                        {
                            ERPCompanyNo = item.ML001;
                        }
                        if (MESCompanyNo != ERPCompanyNo) throw new SystemException("ERP的公司代碼與MES的公司代碼不同，請嘗試重新登入!!");
                        #endregion

                        #region //判斷ERP使用者是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 MF005
                                FROM ADMMF
                                WHERE COMPANY = @CompanyNo
                                AND MF001 = @userNo";
                        dynamicParameters.Add("CompanyNo", ERPCompanyNo);
                        dynamicParameters.Add("userNo", userNo);
                        var resultUserExist = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUserExist.Count() <= 0) throw new SystemException("【ERP使用者】不存在!");
                        #endregion

                        #region //審核ERP權限-建立
                        var USR_GROUP = BaseHelper.CheckErpAuthority(userNo, sqlConnection, "COPI08", "CREATE");
                        #endregion

                        #region //判斷開立日期是否超過結帳日
                        string CloseDateBase = "";
                        DateTime date = DateTime.Parse(ReceiveDate);
                        string ReceiveDateBase = date.ToString("yyyyMMdd");
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
                                throw new SystemException("【銷貨單】ERP已經結帳,不可以在新增【" + CloseDateBase + "】之前的單據");
                            }
                        }
                        else
                        {
                            throw new SystemException("日期字符串无效，无法比较");
                        }
                        #endregion

                        switch (SourceType)
                        {
                            case "So":
                                foreach (var item in RoDetailJsonList)
                                {
                                    string MtlItemNo = item["MtlItemNo"];
                                    string SoFullDetail = item["SoFull"] != null ? item["SoFull"].ToString() : "";
                                    Double Quantity = item["Quantity"] != null ? Convert.ToDouble(item["Quantity"].ToString()) : 0;
                                    Double FreeSpareQty = item["FreeSpareQty"] != null ? Convert.ToDouble(item["FreeSpareQty"].ToString()) : 0;
                                    Double AvailableQty = item["AvailableQty"] != null ? Convert.ToDouble(item["AvailableQty"].ToString()) : 0;
                                    SoErpPrefix = SoFullDetail.Split('-')[0];
                                    SoErpNo = SoFullDetail.Split('-')[1];
                                    SoSequence = SoFullDetail.Split('-')[2];

                                    #region //撈取訂單品號的數量+贈/備品數量
                                    string ProductType = "";
                                    Double CanUseQty = 0;
                                    Double CanUseFreebieQty = 0;
                                    Double CanUseSpareQty = 0;
                                    Double CanUseAvailableQty = 0;
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.TD003 SoSequence, a.TD049 ProductType
                                            ,(ISNULL(a.TD008,0) -  ISNULL(a.TD009,0) - ISNULL(x.UnconfirmedQty,0)) CanUseQty 
                                            ,(ISNULL(a.TD024,0) -  ISNULL(a.TD025,0) - ISNULL(x.UnconfirmedFreeSpareQty,0)) CanUseFreebieQty 
                                            ,(ISNULL(a.TD050,0) -  ISNULL(a.TD051,0) - ISNULL(x.UnconfirmedFreeSpareQty,0)) CanUseSpareQty 
                                            ,(ISNULL(a.TD076,0) -  ISNULL(a.TD078,0) - ISNULL(x.UnconfirmedAvailableQty,0)) CanUseAvailableQty
                                            ,ISNULL(x.UnconfirmedQty,0) UnconfirmedQty
                                            ,ISNULL(x.UnconfirmedFreeSpareQty,0) UnconfirmedFreeSpareQty 
                                            ,ISNULL(x.UnconfirmedAvailableQty,0) UnconfirmedAvailableQty
                                            FROM COPTD a
                                            OUTER APPLY(
                                                SELECT SUM(ISNULL(x1.TH008,0)) UnconfirmedQty 
                                                      ,SUM(ISNULL(x1.TH024,0)) UnconfirmedFreeSpareQty 
                                                      ,SUM(ISNULL(x1.TH061,0)) UnconfirmedAvailableQty
                                                FROM COPTH x1
                                                WHERE x1.TH014 = a.TD001
                                                AND x1.TH015 = a.TD002
                                                AND x1.TH016 = a.TD003
                                                AND x1.TH020 ='N'
                                            ) x
                                            WHERE a.TD001 = @TG014 
                                            AND a.TD002 = @TG015  
                                            AND a.TD003 = @TG016
                                            AND a.TD021 = 'Y'
                                            ";
                                    dynamicParameters.Add("TG014", SoErpPrefix);
                                    dynamicParameters.Add("TG015", SoErpNo);
                                    dynamicParameters.Add("TG016", SoSequence);
                                    var resultCOPTD = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultCOPTD.Count() > 0)
                                    {
                                        foreach (var item2 in resultCOPTD)
                                        {
                                            ProductType = item2.ProductType;
                                            CanUseQty = Convert.ToDouble(item2.CanUseQty);
                                            CanUseFreebieQty = Convert.ToDouble(item2.CanUseFreebieQty);
                                            CanUseSpareQty = Convert.ToDouble(item2.CanUseSpareQty);
                                            CanUseAvailableQty = Convert.ToDouble(item2.CanUseAvailableQty);
                                        }
                                    }
                                    if(Quantity > CanUseQty) throw new SystemException("【銷貨量超交】:品號『" + MtlItemNo + "』本次最大可銷貨數量為" + CanUseQty + "，請重新輸入!");
                                    if(Quantity > CanUseAvailableQty) throw new SystemException("【銷貨量超交】:品號『" + MtlItemNo + "』本次最大可銷貨計價數量為" + CanUseAvailableQty + "，請重新輸入!");
                                    if(ProductType == "1")
                                    {
                                        if (FreeSpareQty > CanUseFreebieQty) throw new SystemException("【贈品量超交】:品號『" + MtlItemNo + "』本次最大可銷貨贈品數量為" + CanUseFreebieQty + "】，請重新輸入!");
                                    }
                                    else if(ProductType == "2")
                                    {
                                        if (FreeSpareQty > CanUseSpareQty) throw new SystemException("【備品量超交】:品號『" + MtlItemNo + "』本次最大可銷貨備品數量為" + CanUseSpareQty + "】，請重新輸入!");
                                    }
                                    else
                                    {
                                        throw new SystemException("贈備品類別異常,請重新確認");
                                    }
                                    #endregion
                                }
                                break;
                            case "To":
                                foreach (var item in RoDetailJsonList)
                                {
                                    string MtlItemNo = item["MtlItemNo"];
                                    string TsnFull = item["TsnFull"] != null ? item["TsnFull"].ToString() : "";
                                    Double Quantity = item["Quantity"] != null ? Convert.ToDouble(item["Quantity"].ToString()) : 0;
                                    Double FreeSpareQty = item["FreeSpareQty"] != null ? Convert.ToDouble(item["FreeSpareQty"].ToString()) : 0;
                                    string TsnErpPrefix = TsnFull.Split('-')[0];
                                    string TsnErpNo = TsnFull.Split('-')[1];
                                    string TsnSequence = TsnFull.Split('-')[2];


                                    #region //撈取暫出單的可銷貨數量+可銷貨贈/備品數量
                                    Double CanUseQty = 0;
                                    Double CanUseFreeSpareQty = 0;
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT 
                                             (ISNULL(a.TG009,0) - ISNULL(a.TG020,0) - ISNULL(a.TG021,0) - ISNULL(b.UnconfirmedQty,0)) CanUseQty
                                            ,(ISNULL(a.TG044,0) - ISNULL(a.TG046,0) - ISNULL(a.TG048,0) - ISNULL(b.UnconfirmedFreebieQty,0)) CanUseFreeSpareQty 
                                            ,(ISNULL(a.TG052,0) - ISNULL(a.TG054,0) - ISNULL(a.TG055,0) - ISNULL(b.UnconfirmedAvailableQty,0)) CanUseAvailableQty 
                                            FROM INVTG a
                                            OUTER APPLY(
                                                SELECT SUM(ISNULL(x1.TH008,0)) UnconfirmedQty 
                                                      ,SUM(ISNULL(x1.TH024,0)) UnconfirmedFreebieQty 
                                                      ,SUM(ISNULL(x1.TH061,0)) UnconfirmedAvailableQty
                                                FROM COPTH x1
                                                WHERE x1.TH032 = a.TG001
                                                AND x1.TH033 = a.TG002
                                                AND x1.TH034 = a.TG003
                                                AND x1.TH020 ='N'
                                            ) b
                                            WHERE a.TG001 = @TG001
                                            AND a.TG002 = @TG002
                                            AND a.TG003 = @TG003
                                            AND a.TG022 ='Y'
                                            ";
                                    dynamicParameters.Add("TG001", TsnErpPrefix);
                                    dynamicParameters.Add("TG002", TsnErpNo);
                                    dynamicParameters.Add("TG003", TsnSequence);
                                    var resultINVTG = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultINVTG.Count() > 0)
                                    {
                                        foreach (var item2 in resultINVTG)
                                        {
                                            CanUseQty = Convert.ToDouble(item2.CanUseQty);
                                            CanUseFreeSpareQty = Convert.ToDouble(item2.CanUseFreeSpareQty);
                                        }
                                    }
                                    #endregion

                                    #region //判斷可銷貨數量是否有超收

                                    if (CanUseQty < Quantity) throw new SystemException("【銷貨量超交】:品號『" + MtlItemNo + "』本次最大可銷貨量為" + (CanUseQty) + "，請重新輸入!");
                                    if (CanUseFreeSpareQty < FreeSpareQty) throw new SystemException("【贈品量超交】:品號『" + MtlItemNo + "』本次最大可贈品量為" + (CanUseFreeSpareQty) + "，請重新輸入!");
                                    #endregion
                                }
                                break;
                        }

                        var result = sqlConnection.Query(sql, dynamicParameters);

                        int[] RoIdResult = { RoId };


                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            data = RoIdResult
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

        #region //AddReturnReceiveOrder -- 銷退單新增 -- Shintokuro 2024.05.10
        public string AddReturnReceiveOrder(int ViewCompanyId
            , string DocDate
            , string ReturnDate, string RtErpPrefix, string RtErpNo, int CustomerId, string ProcessCode
            , string RevenueJournalCode, string CostJournalCode
            , double OriginalAmount
            , double OriginalTaxAmount, double PretaxAmount, double TaxAmount, double TotalQuantity
            , double TaxCurrencyRate
            , int DepartmentId
            , int SalesmenId, int StaffId, int CollectionSalesmenId, string Factory, string PaymentTerm, string TradingTerms
            , string Currency, double ExchangeRate, string BondCode, int RowCnt, string Remark
            , string TaxCode
            , string InvoiceType, string TaxType, double TaxRate, string UiNo, string InvoiceDate, string InvNumStart, string InvNumEnd
            , string ApplyYYMM, string CustomsClearance
            , string CustomerFullName, string CustomerEgFullName, string MultipleInvoices, string DebitNote
            , string ContactPerson
            , string ContactEmail, string CustomerAddressFirst, string CustomerAddressSecond, string Remark1, string Remark2, string Remark3
            , string RtDetailData
            )
        {
            try
            {
                if (RtErpPrefix.Length <= 0) throw new SystemException("【銷退單別】不能為空!");
                if (ReturnDate.Length <= 0) throw new SystemException("【銷退日期】不能為空!");
                if (CustomerId <= 0) throw new SystemException("【客戶】不能為空!");
                if (Factory.Length <= 0) throw new SystemException("【廠別】不能為空!");
                if (Currency.Length <= 0) throw new SystemException("【幣別】不能為空!");
                if (ExchangeRate < 0) throw new SystemException("【匯率】不能為負!");
                if (ApplyYYMM.Length < 0) throw new SystemException("【申報年月】不能為空!");
                if (DocDate.Length < 0) throw new SystemException("【單據日期】不能為空!");
                if (TaxRate < 0) throw new SystemException("【營業稅率】不能為負!");
                ApplyYYMM = ApplyYYMM.Replace("-", "");
                if (CurrentCompany != ViewCompanyId) throw new SystemException("頁面的公司別與後端公司別不同，請嘗試重新登入!!");

                List<Dictionary<string, string>> RtDetailJsonList = JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(RtDetailData);
                int? nullData = null;
                int rowsAffected = 0;
                int RtId = 0;
                string SoErpPrefix = "";
                string SoErpNo = "";
                string SoSequence = "";
                string MESCompanyNo = "";
                string userNo = "";
                string userName = "";
                string departmentNo = "";
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷公司資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 CompanyId, CompanyName, ErpDb,ErpNo
                            FROM BAS.Company
                            WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【公司】資料錯誤!");

                        foreach (var item in result)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            MESCompanyNo = item.ErpNo;
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
                        dynamicParameters.Add("FunctionCode", "ReturnReceiveOrderManagment");
                        dynamicParameters.Add("DetailCode", "add");

                        var resultUser = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUser.Count() <= 0) throw new SystemException("使用者資料錯誤!");

                        foreach (var item in resultUser)
                        {
                            userNo = item.UserNo;
                            userName = item.UserName;
                            departmentNo = item.DepartmentNo;
                        }
                        #endregion

                        #region //判斷客戶資料是否存在
                        sql = @"SELECT TOP 1 1
                                FROM SCM.Customer
                                WHERE CompanyId = @CompanyId
                                AND CustomerId = @CustomerId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("CustomerId", CustomerId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("客戶資料找不到，請重新確認!!");
                        #endregion

                        #region //判斷部門資料是否存在
                        if (DepartmentId > 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.Department
                                    WHERE CompanyId = @CompanyId
                                    AND DepartmentId = @DepartmentId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("DepartmentId", DepartmentId);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("部門資料找不到，請重新確認!!");
                        }
                        #endregion

                        #region //判斷業務資料是否存在
                        if (SalesmenId > 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.[User]
                                    WHERE UserId = @UserId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("UserId", SalesmenId);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("業務資料找不到，請重新確認!!");
                        }
                        #endregion

                        #region //判斷收款業務資料是否存在
                        if (CollectionSalesmenId > 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.[User]
                                    WHERE UserId = @UserId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("UserId", CollectionSalesmenId);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("收款業務資料找不到，請重新確認!!");
                        }
                        #endregion

                        #region //判斷員工資料是否存在
                        if (StaffId > 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                FROM BAS.[User]
                                WHERE UserId = @UserId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("UserId", StaffId);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("員工資料找不到，請重新確認!!");
                        }
                        #endregion

                        #region //判斷銷退單單別+單號是否重複
                        RtErpNo = BaseHelper.RandomCode(11);
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.ReturnReceiveOrder
                                WHERE RtErpPrefix = @RtErpPrefix
                                AND RtErpNo = @RtErpNo
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("RtErpPrefix", RtErpPrefix);
                        dynamicParameters.Add("RtErpNo", RtErpNo);
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【銷退單新增】－【單別+單號】重複，請重新輸入!");
                        #endregion

                        #region //新增單頭資料表
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.ReturnReceiveOrder (CompanyId, RtErpPrefix, RtErpNo, ReturnDate, CustomerId, DepartmentId, 
                                        SalesmenId, Factory, Currency, ExchangeRate, OriginalAmount, OriginalTaxAmount, InvoiceType, TaxType, 
                                        InvNumEnd, UiNo, PrintCount, InvoiceDate, UpdateCode, ConfirmStatus, Remark, CustomerFullName, 
                                        CollectionSalesmenId, Remark1, Remark2, Remark3, DeductionDistinction, CustomsClearance, RowCnt, 
                                        TotalQuantity, StaffId, RevenueJournalCode, CostJournalCode, ApplyYYMM, DocDate, ConfirmUserId, TaxRate, 
                                        PretaxAmount, TaxAmount, PaymentTerm, TotalPackages, SignatureStatus, ProcessCode, TransferStatusERP, 
                                        BondCode, TransmissionCount, CustomerAddressFirst, CustomerAddressSecond, ContactPerson, Courier, 
                                        SiteCommCalcMethod, SiteCommRate, CommCalcInclTax, TotalCommAmount, TradingTerms, CustomerEgFullName, 
                                        DebitNote, TaxCode, RemarkCode, MultipleInvoices, InvNumStart, NumberOfInvoices, MultiTaxRate, 
                                        TaxCurrencyRate, VoidDate, VoidTime, VoidRemark, IncomeDraftID, IncomeDraftSeq, IncomeVoucherType, 
                                        IncomeVoucherNumber, Status, GenLedgerAcctType, InvCode2, InvSymbol, ContactEmail, OrigInvNumber, 
                                        TransferStatusMES, CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.RtId
                                        VALUES ( @CompanyId, @RtErpPrefix, @RtErpNo, @ReturnDate, @CustomerId, @DepartmentId, 
                                        @SalesmenId, @Factory, @Currency, @ExchangeRate, @OriginalAmount, @OriginalTaxAmount, @InvoiceType, @TaxType, 
                                        @InvNumEnd, @UiNo, @PrintCount, @InvoiceDate, @UpdateCode, @ConfirmStatus, @Remark, @CustomerFullName, 
                                        @CollectionSalesmenId, @Remark1, @Remark2, @Remark3, @DeductionDistinction, @CustomsClearance, @RowCnt, 
                                        @TotalQuantity, @StaffId, @RevenueJournalCode, @CostJournalCode, @ApplyYYMM, @DocDate, @ConfirmUserId, @TaxRate, 
                                        @PretaxAmount, @TaxAmount, @PaymentTerm, @TotalPackages, @SignatureStatus, @ProcessCode, @TransferStatusERP, 
                                        @BondCode, @TransmissionCount, @CustomerAddressFirst, @CustomerAddressSecond, @ContactPerson, @Courier, 
                                        @SiteCommCalcMethod, @SiteCommRate, @CommCalcInclTax, @TotalCommAmount, @TradingTerms, @CustomerEgFullName, 
                                        @DebitNote, @TaxCode, @RemarkCode, @MultipleInvoices, @InvNumStart, @NumberOfInvoices, @MultiTaxRate, 
                                        @TaxCurrencyRate, @VoidDate, @VoidTime, @VoidRemark, @IncomeDraftID, @IncomeDraftSeq, @IncomeVoucherType, 
                                        @IncomeVoucherNumber, @Status, @GenLedgerAcctType, @InvCode2, @InvSymbol, @ContactEmail, @OrigInvNumber, 
                                        @TransferStatusMES, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = ViewCompanyId,
                                RtErpPrefix,
                                RtErpNo,
                                ReturnDate,
                                CustomerId,
                                DepartmentId,
                                SalesmenId,
                                Factory,
                                Currency,
                                ExchangeRate,
                                OriginalAmount,
                                OriginalTaxAmount,
                                InvoiceType,
                                TaxType,
                                InvNumEnd,
                                UiNo,
                                PrintCount = 0,
                                InvoiceDate,
                                UpdateCode = "N",
                                ConfirmStatus = "N",
                                Remark,
                                CustomerFullName,
                                CollectionSalesmenId,
                                Remark1,
                                Remark2,
                                Remark3,
                                DeductionDistinction = "3",
                                CustomsClearance,
                                RowCnt,
                                TotalQuantity,
                                StaffId,
                                RevenueJournalCode,
                                CostJournalCode,
                                ApplyYYMM,
                                DocDate,
                                ConfirmUserId =(int?)null,
                                TaxRate,
                                PretaxAmount,
                                TaxAmount,
                                PaymentTerm,
                                TotalPackages = 0,
                                SignatureStatus = "N",
                                ProcessCode,
                                TransferStatusERP = "N",
                                BondCode,
                                TransmissionCount = 0,
                                CustomerAddressFirst,
                                CustomerAddressSecond,
                                ContactPerson,
                                Courier = "",
                                SiteCommCalcMethod = "1",
                                SiteCommRate = 0,
                                CommCalcInclTax = "N",
                                TotalCommAmount = 0,
                                TradingTerms,
                                CustomerEgFullName,
                                DebitNote,
                                TaxCode,
                                RemarkCode = "",
                                MultipleInvoices,
                                InvNumStart,
                                NumberOfInvoices = 0,
                                MultiTaxRate = "N",
                                TaxCurrencyRate = 1,
                                VoidDate = "",
                                VoidTime = "",
                                VoidRemark = "",
                                IncomeDraftID = "",
                                IncomeDraftSeq = "",
                                IncomeVoucherType = "",
                                IncomeVoucherNumber = "",
                                Status = "0",
                                GenLedgerAcctType = "",
                                InvCode2 = "",
                                InvSymbol = "",
                                ContactEmail,
                                OrigInvNumber = "",
                                TransferStatusMES = "N",
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);
                        rowsAffected += insertResult.Count();
                        #endregion

                        foreach (var item in insertResult)
                        {
                            RtId = item.RtId;
                        }

                        if (RtDetailJsonList.Count() > 0)
                        {
                            foreach (var item in RtDetailJsonList)
                            {
                                string MtlItemNo = item["MtlItemNo"] != null ? item["MtlItemNo"].ToString() : throw new SystemException("【資料維護不完整】品號欄位資料不可以為空,請重新確認~~");
                                string Type = item["Type"] != "" ? item["Type"].ToString().Substring(0, 1) : throw new SystemException("【資料維護不完整】類別欄位資料不可以為空,請重新確認~~");
                                Double Quantity = item["Quantity"] != null ? Convert.ToDouble(item["Quantity"].ToString()) : 0;
                                string FreeSpareType = item["Type"] != null ? item["FreeSpareType"].ToString() : throw new SystemException("【資料維護不完整】贈備品類別欄位資料不可以為空,請重新確認~~");
                                Double FreeSpareQty = item["FreeSpareQty"] != null ? Convert.ToDouble(item["FreeSpareQty"].ToString()) : 0;
                                string UomNo = item["UomNo"] != null ? item["UomNo"].ToString() : throw new SystemException("【資料維護不完整】單位欄位資料不可以為空,請重新確認~~");
                                Double AvailableQty = item["AvailableQty"] != null ? Convert.ToDouble(item["AvailableQty"].ToString()) : 0;
                                string AvailableUomNo = item["AvailableUomNo"] != null ? item["AvailableUomNo"].ToString() : throw new SystemException("【資料維護不完整】計價單位欄位資料不可以為空,請重新確認~~");
                                Double UnitPrice = item["UnitPrice"] != null ? Convert.ToDouble(item["UnitPrice"].ToString()) : 0;
                                Double TaxRateDetail = item["TaxRate"] != null ? Convert.ToDouble(item["TaxRate"].ToString()) : 0;
                                Double RoDiscountRate = item["RoDiscountRate"] != "" ? Convert.ToDouble(item["RoDiscountRate"].ToString()) : 0;
                                Double SoDiscountRate = item["SoDiscountRate"] != "" ? Convert.ToDouble(item["SoDiscountRate"].ToString()) : 0;
                                Double Amount = item["Amount"] != null ? Convert.ToDouble(item["Amount"].ToString()) : 0;
                                Double OrigPreTaxAmt = item["OrigPreTaxAmt"] != null ? Convert.ToDouble(item["OrigPreTaxAmt"].ToString()) : 0;
                                Double OrigTaxAmt = item["OrigTaxAmt"] != null ? Convert.ToDouble(item["OrigTaxAmt"].ToString()) : 0;
                                Double PreTaxAmt = item["PreTaxAmt"] != null ? Convert.ToDouble(item["PreTaxAmt"].ToString()) : 0;
                                Double TaxAmt = item["TaxAmt"] != null ? Convert.ToDouble(item["TaxAmt"].ToString()) : 0;
                                string InventoryNo = item["InventoryNo"] != null ? item["InventoryNo"].ToString() : throw new SystemException("【資料維護不完整】庫別欄位資料不可以為空,請重新確認~~");
                                string Location = item["Location"] != null ? item["Location"].ToString() : "";
                                string SaveLocation = item["SaveLocation"] != null ? item["SaveLocation"].ToString() : "";
                                string LotNumber = item["LotNumber"] != null ? item["LotNumber"].ToString() : "";
                                string RoFullNo = item["RoFullNo"] != null ? item["RoFullNo"].ToString() : "";
                                string SoFullNo = item["SoFullNo"] != null ? item["SoFullNo"].ToString() : "";
                                string CustomerMtlItemNo = item["CustomerMtlItemNo"] != null ? item["CustomerMtlItemNo"].ToString() : "";
                                string ProjectCode = item["ProjectCode"] != null ? item["ProjectCode"].ToString() : "";
                                string Remarks = item["ReturnRemarks"] != null ? item["ReturnRemarks"].ToString() : "";
                                string BondedCode = item["BondedCode"] != null ? item["BondedCode"].ToString() : "";
                                string ReturnNo = item["ReturnNo"] != null ? item["ReturnNo"].ToString() : "";

                                #region //匯率等於1時判定
                                if (ExchangeRate == 1)
                                {
                                    if (OrigPreTaxAmt != PreTaxAmt) throw new SystemException("單據幣別與本國貨幣相同時,原幣金額與本幣金額須相同!!");
                                    if (OrigTaxAmt != TaxAmt) throw new SystemException("單據幣別與本國貨幣相同時,原幣稅額與本幣稅額須相同!!");
                                }
                                #endregion

                                #region //判斷品號是否存在
                                int MtlItemId = -1;
                                string OverDeliveryManagement = "";
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 MtlItemId,OverDeliveryManagement,LotManagement
                                        FROM PDM.MtlItem
                                        WHERE CompanyId = @CompanyId
                                        AND MtlItemNo = @MtlItemNo";
                                dynamicParameters.Add("MtlItemNo", MtlItemNo);
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("【品號:" + MtlItemNo + "】找不到，請重新輸入!");
                                foreach (var item2 in result)
                                {
                                    MtlItemId = item2.MtlItemId;
                                    OverDeliveryManagement = item2.OverDeliveryManagement;
                                    if (item2.LotManagement != "N" && LotNumber == "") throw new SystemException("【品號:" + MtlItemNo + "】須維護批號，請重新輸入!");
                                }
                                #endregion

                                #region //判斷庫別是否存在
                                int InventoryId = -1;
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 InventoryId
                                        FROM SCM.Inventory
                                        WHERE CompanyId = @CompanyId
                                        AND InventoryNo = @InventoryNo";
                                dynamicParameters.Add("InventoryNo", InventoryNo.Split(':')[0]);
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("【庫別:" + InventoryNo + "】找不到，請重新輸入!");
                                foreach (var item2 in result)
                                {
                                    InventoryId = item2.InventoryId;
                                }
                                #endregion

                                #region //判斷單位是否存在
                                int UomId = -1;
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 UomId
                                        FROM PDM.UnitOfMeasure
                                        WHERE CompanyId = @CompanyId
                                        AND UomNo = @UomNo";
                                dynamicParameters.Add("UomNo", UomNo);
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("【單位:" + UomNo + "】找不到，請重新輸入!");
                                foreach (var item2 in result)
                                {
                                    UomId = item2.UomId;
                                }
                                #endregion

                                #region //判斷計價單位是否存在
                                int AvailableUomId = -1;
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 UomId
                                        FROM PDM.UnitOfMeasure
                                        WHERE CompanyId = @CompanyId
                                        AND UomNo = @UomNo";
                                dynamicParameters.Add("UomNo", AvailableUomNo);
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("【計價單位:" + AvailableUomNo + "】找不到，請重新輸入!");
                                foreach (var item2 in result)
                                {
                                    AvailableUomId = item2.UomId;
                                }
                                #endregion

                                #region //判斷訂單單身是否存在
                                int? SoDetailId = null;
                                string ProductType = "";
                                Double FreebieQty = 0;
                                Double SpareQty = 0;
                                if (SoFullNo.Length > 0)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 a.SoDetailId, a.ProductType, a.FreebieQty, a.SpareQty
                                            ,b.SoErpPrefix, b.SoErpNo, a.SoSequence, a.ClosureStatus
                                            FROM SCM.SoDetail a
                                            INNER JOIN SCM.SaleOrder b on a.SoId = b.SoId
                                            WHERE b.CompanyId = @CompanyId
                                            AND b.SoErpPrefix = @SoErpPrefix
                                            AND b.SoErpNo = @SoErpNo
                                            AND a.SoSequence = @SoSequence";
                                    dynamicParameters.Add("CompanyId", CurrentCompany);
                                    dynamicParameters.Add("SoErpPrefix", SoFullNo.Split('-')[0]);
                                    dynamicParameters.Add("SoErpNo", SoFullNo.Split('-')[1]);
                                    dynamicParameters.Add("SoSequence", SoFullNo.Split('-')[2]);
                                    result = sqlConnection.Query(sql, dynamicParameters);
                                    if (result.Count() <= 0) throw new SystemException("【訂單:" + SoFullNo + "】找不到，請重新輸入!");
                                    foreach (var item2 in result)
                                    {
                                        SoErpPrefix = item2.SoErpPrefix;
                                        SoErpNo = item2.SoErpNo;
                                        SoSequence = item2.SoSequence;
                                        SoDetailId = item2.SoDetailId;
                                        ProductType = item2.ProductType;
                                        FreebieQty = item2.FreebieQty;
                                        SpareQty = item2.SpareQty;

                                    }
                                }
                                #endregion

                                #region //判斷銷貨單身是否存在
                                int? RoDetailId = null;
                                string RoProductType = "";
                                Double FreebieOrSpareQty = 0;
                                if (RoFullNo.Length > 0)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 b.RoDetailId,b.Type,b.FreeSpareQty,b.CheckOutCode
                                            FROM SCM.ReceiveOrder a
                                            INNER JOIN SCM.RoDetail b on a.RoId = b.RoId
                                            WHERE CompanyId = @CompanyId
                                            AND a.RoErpPrefix = @RoErpPrefix
                                            AND a.RoErpNo = @RoErpNo
                                            AND b.RoSequence = @RoSequence
                                            AND a.ConfirmStatus = 'Y'";
                                    dynamicParameters.Add("CompanyId", CurrentCompany);
                                    dynamicParameters.Add("RoErpPrefix", RoFullNo.Split('-')[0]);
                                    dynamicParameters.Add("RoErpNo", RoFullNo.Split('-')[1]);
                                    dynamicParameters.Add("RoSequence", RoFullNo.Split('-')[2]);
                                    result = sqlConnection.Query(sql, dynamicParameters);
                                    if (result.Count() <= 0) throw new SystemException("【銷貨單:" + RoFullNo + "】找不到，請重新輸入!");
                                    foreach (var item2 in result)
                                    {
                                        if (item2.CheckOutCode == "Y") throw new SystemException("【銷貨單:" + RoFullNo + "】已經結案!!!");
                                        RoDetailId = item2.RoDetailId;
                                        RoProductType = item2.Type;
                                        FreebieOrSpareQty = item2.FreeSpareQty;
                                    }
                                }
                                #endregion

                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO SCM.RtDetail (RtId, RtSequence, MtlItemId, Quantity, UomId, UnitPrice, Amount, InventoryId, LotNumber, 
                                        RoDetailId, SoDetailId, ConfirmStatus, UpdateCode, Remarks, CheckOutCode, CheckOutPrefix, CheckOutNo, 
                                        CheckOutSeq, ProjectCode, CustomerMtlItemNo, Type, OrigPreTaxAmt, OrigTaxAmt, PreTaxAmt, TaxAmt, PackageQty, 
                                        PackageUomId, BondedCode, CommissionRate, CommissionAmount, OriginalCustomer, FreeSpareType, 
                                        FreeSpareQty, PackageFreeSpareQty, NotPayTemp, Location, AvailableQty, AvailableUomId, TaxRate, TaxCode, 
                                        DiscountAmount, TransferStatusMES, CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        VALUES ( @RtId, @RtSequence, @MtlItemId, @Quantity, @UomId, @UnitPrice, @Amount, @InventoryId, @LotNumber, 
                                        @RoDetailId, @SoDetailId, @ConfirmStatus, @UpdateCode, @Remarks, @CheckOutCode, @CheckOutPrefix, @CheckOutNo, 
                                        @CheckOutSeq, @ProjectCode, @CustomerMtlItemNo, @Type, @OrigPreTaxAmt, @OrigTaxAmt, @PreTaxAmt, @TaxAmt, @PackageQty, 
                                        @PackageUomId, @BondedCode, @CommissionRate, @CommissionAmount, @OriginalCustomer, @FreeSpareType, 
                                        @FreeSpareQty, @PackageFreeSpareQty, @NotPayTemp, @Location, @AvailableQty, @AvailableUomId, @TaxRate, @TaxCode, 
                                        @DiscountAmount, @TransferStatusMES, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy
                                        )";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        RtId,
                                        RtSequence = "",
                                        MtlItemId,
                                        Quantity,
                                        UomId,
                                        UnitPrice,
                                        Amount,
                                        InventoryId,
                                        LotNumber,
                                        RoDetailId,
                                        SoDetailId,
                                        ConfirmStatus = "N",
                                        UpdateCode = "N",
                                        Remarks,
                                        CheckOutCode = "N",
                                        CheckOutPrefix = "",
                                        CheckOutNo = "",
                                        CheckOutSeq = "",
                                        ProjectCode,
                                        CustomerMtlItemNo,
                                        Type,
                                        OrigPreTaxAmt,
                                        OrigTaxAmt,
                                        PreTaxAmt,
                                        TaxAmt,
                                        PackageQty = 0,
                                        PackageUomId = "",
                                        BondedCode,
                                        CommissionRate = 0,
                                        CommissionAmount = 0,
                                        OriginalCustomer = "",
                                        FreeSpareType = RoProductType,
                                        FreeSpareQty,
                                        PackageFreeSpareQty = 0,
                                        NotPayTemp = "N",
                                        Location,
                                        AvailableQty,
                                        AvailableUomId,
                                        TaxRate,
                                        TaxCode,
                                        DiscountAmount = 0,
                                        TransferStatusMES = "N",
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();
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

                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        #region //ERP公司代碼取得
                        string ERPCompanyNo = "";
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 ML001  
                                         FROM CMSML";
                        var resultCompanyNo = sqlConnection.Query(sql, dynamicParameters);
                        foreach (var item in resultCompanyNo)
                        {
                            ERPCompanyNo = item.ML001;
                        }
                        if (MESCompanyNo != ERPCompanyNo) throw new SystemException("ERP的公司代碼與MES的公司代碼不同，請嘗試重新登入!!");
                        #endregion

                        #region //判斷ERP使用者是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 MF005
                                FROM ADMMF
                                WHERE COMPANY = @CompanyNo
                                AND MF001 = @userNo";
                        dynamicParameters.Add("CompanyNo", ERPCompanyNo);
                        dynamicParameters.Add("userNo", userNo);
                        var resultUserExist = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUserExist.Count() <= 0) throw new SystemException("【ERP使用者】不存在!");
                        #endregion

                        #region //審核ERP權限-建立
                        var USR_GROUP = BaseHelper.CheckErpAuthority(userNo, sqlConnection, "COPI08", "CREATE");
                        #endregion

                        #region //判斷開立日期是否超過結帳日
                        string CloseDateBase = "";
                        DateTime date = DateTime.Parse(ReturnDate);
                        string ReceiveDateBase = date.ToString("yyyyMMdd");
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
                                throw new SystemException("【銷貨單】ERP已經結帳,不可以在新增【" + CloseDateBase + "】之前的單據");
                            }
                        }
                        else
                        {
                            throw new SystemException("日期字符串无效，无法比较");
                        }
                        #endregion

                        foreach (var item in RtDetailJsonList)
                        {
                            string MtlItemNo = item["MtlItemNo"];
                            string SoFullDetail = item["SoFullNo"] != null ? item["SoFullNo"].ToString() : "";
                            Double Quantity = item["Quantity"] != null ? Convert.ToDouble(item["Quantity"].ToString()) : 0;
                            Double FreeSpareQty = item["FreeSpareQty"] != null ? Convert.ToDouble(item["FreeSpareQty"].ToString()) : 0;
                            Double AvailableQty = item["AvailableQty"] != null ? Convert.ToDouble(item["AvailableQty"].ToString()) : 0;
                            SoErpPrefix = SoFullDetail.Split('-')[0];
                            SoErpNo = SoFullDetail.Split('-')[1];
                            SoSequence = SoFullDetail.Split('-')[2];

                            #region //撈取訂單品號的數量+贈/備品數量
                            string ProductType = "";
                            Double CanUseQty = 0;
                            Double CanUseFreebieQty = 0;
                            Double CanUseSpareQty = 0;
                            Double CanUseAvailableQty = 0;
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.TD003 SoSequence, a.TD049 ProductType
                                            ,(ISNULL(a.TD008,0) -  ISNULL(a.TD009,0) - ISNULL(x.UnconfirmedQty,0)) CanUseQty 
                                            ,(ISNULL(a.TD024,0) -  ISNULL(a.TD025,0) - ISNULL(x.UnconfirmedFreeSpareQty,0)) CanUseFreebieQty 
                                            ,(ISNULL(a.TD050,0) -  ISNULL(a.TD051,0) - ISNULL(x.UnconfirmedFreeSpareQty,0)) CanUseSpareQty 
                                            ,(ISNULL(a.TD076,0) -  ISNULL(a.TD078,0) - ISNULL(x.UnconfirmedAvailableQty,0)) CanUseAvailableQty
                                            ,ISNULL(x.UnconfirmedQty,0) UnconfirmedQty
                                            ,ISNULL(x.UnconfirmedFreeSpareQty,0) UnconfirmedFreeSpareQty 
                                            ,ISNULL(x.UnconfirmedAvailableQty,0) UnconfirmedAvailableQty
                                            FROM COPTD a
                                            OUTER APPLY(
                                                SELECT SUM(ISNULL(x1.TH008,0)) UnconfirmedQty 
                                                      ,SUM(ISNULL(x1.TH024,0)) UnconfirmedFreeSpareQty 
                                                      ,SUM(ISNULL(x1.TH061,0)) UnconfirmedAvailableQty
                                                FROM COPTH x1
                                                WHERE x1.TH014 = a.TD001
                                                AND x1.TH015 = a.TD002
                                                AND x1.TH016 = a.TD003
                                                AND x1.TH020 ='N'
                                            ) x
                                            WHERE a.TD001 = @TG014 
                                            AND a.TD002 = @TG015  
                                            AND a.TD003 = @TG016
                                            AND a.TD021 = 'Y'
                                            ";
                            dynamicParameters.Add("TG014", SoErpPrefix);
                            dynamicParameters.Add("TG015", SoErpNo);
                            dynamicParameters.Add("TG016", SoSequence);
                            var resultCOPTD = sqlConnection.Query(sql, dynamicParameters);
                            if (resultCOPTD.Count() > 0)
                            {
                                foreach (var item2 in resultCOPTD)
                                {
                                    ProductType = item2.ProductType;
                                    CanUseQty = Convert.ToDouble(item2.CanUseQty);
                                    CanUseFreebieQty = Convert.ToDouble(item2.CanUseFreebieQty);
                                    CanUseSpareQty = Convert.ToDouble(item2.CanUseSpareQty);
                                    CanUseAvailableQty = Convert.ToDouble(item2.CanUseAvailableQty);
                                }
                            }
                            //if (Quantity > CanUseQty) throw new SystemException("【銷貨量超交】:品號『" + MtlItemNo + "』本次最大可銷貨數量為" + CanUseQty + "，請重新輸入!");
                            //if (Quantity > CanUseAvailableQty) throw new SystemException("【銷貨量超交】:品號『" + MtlItemNo + "』本次最大可銷貨計價數量為" + CanUseAvailableQty + "，請重新輸入!");
                            //if (ProductType == "1")
                            //{
                            //    if (FreeSpareQty > CanUseFreebieQty) throw new SystemException("【贈品量超交】:品號『" + MtlItemNo + "』本次最大可銷貨贈品數量為" + CanUseFreebieQty + "】，請重新輸入!");
                            //}
                            //else if (ProductType == "2")
                            //{
                            //    if (FreeSpareQty > CanUseSpareQty) throw new SystemException("【備品量超交】:品號『" + MtlItemNo + "』本次最大可銷貨備品數量為" + CanUseSpareQty + "】，請重新輸入!");
                            //}
                            //else
                            //{
                            //    throw new SystemException("贈備品類別異常,請重新確認");
                            //}
                            #endregion
                        }

                        var result = sqlConnection.Query(sql, dynamicParameters);

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
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

        #region //AddReceiveOrderToErp -- 出貨單轉銷貨 -- Shintokuro 2024.09.03
        public string AddReceiveOrderToErp(int ViewCompanyId, string DocDate, string RoErpPrefix
            , int CustomerId, int SalesmenId, int CollectionSalesmenId, int DepartmentId, string InvNumStart, string Remark, string DoDetails
            )
        {
            try
            {
                throw new SystemException("不開放使用囉");
                if (RoErpPrefix.Length <= 0) throw new SystemException("【銷貨單別】不能為空!");
                //if (CustomerId <= 0) throw new SystemException("【客戶】不能為空!");
                //if (SalesmenId <= 0) throw new SystemException("【業務】不能為空!");
                //if (DepartmentId <= 0) throw new SystemException("【部門】不能為空!");
                //if (DocDate.Length < 0) throw new SystemException("【單據日期】不能為空!");
                if (CurrentCompany != ViewCompanyId) throw new SystemException("頁面的公司別與後端公司別不同，請嘗試重新登入!!");
                if (!DoDetails.TryParseJson(out JObject tempJObject)) throw new SystemException("出貨資料格式錯誤");
                JObject doJson = JObject.Parse(DoDetails);
                if (!doJson.ContainsKey("data")) throw new SystemException("出貨資料錯誤");
                JToken doData = doJson["data"];
                if (doData.Count() < 0) throw new SystemException("查無出貨資料內容");

                string dateNow = CreateDate.ToString("yyyyMMdd"), timeNow = CreateDate.ToString("HH:mm:ss"), errorMsg = "";
                string ApplyYYMM = CreateDate.ToString("yyyyMM");
                int? nullData = null;
                int rowsAffected = 0;
                int RoId = 0;
                string RoErpNo = "";

                string SoErpPrefix = "";
                string SoErpNo = "";
                string SoSequence = "";

                string MESCompanyNo = "";
                string userNo = "";
                string userName = "";
                string departmentNo = "";
                string collectionSalesmenNo = "";

                List<ReceiveOrder> ReceiveOrders = new List<ReceiveOrder>();
                List<RoDetail> RoDetails = new List<RoDetail>();

                List<DoDetail> doDetailList = new List<DoDetail>();

                List<COPTG> ReceiveOrderErp = new List<COPTG>();
                List<COPTH> RoDetailErp = new List<COPTH>();

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷公司資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 CompanyId, CompanyName, ErpDb,ErpNo
                            FROM BAS.Company
                            WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【公司】資料錯誤!");

                        foreach (var item in result)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            MESCompanyNo = item.ErpNo;
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
                        dynamicParameters.Add("FunctionCode", "ReceiveOrderManagment");
                        dynamicParameters.Add("DetailCode", "add");

                        var resultUser = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUser.Count() <= 0) throw new SystemException("使用者資料錯誤!");

                        foreach (var item in resultUser)
                        {
                            userNo = item.UserNo;
                            userName = item.UserName;
                            departmentNo = item.DepartmentNo;
                        }
                        #endregion

                        //#region //判斷客戶資料是否存在
                        //sql = @"SELECT TOP 1 1
                        //        FROM SCM.Customer
                        //        WHERE CompanyId = @CompanyId
                        //        AND CustomerId = @CustomerId";
                        //dynamicParameters.Add("CompanyId", CurrentCompany);
                        //dynamicParameters.Add("CustomerId", CustomerId);
                        //result = sqlConnection.Query(sql, dynamicParameters);
                        //if (result.Count() <= 0) throw new SystemException("客戶資料找不到，請重新確認!!");
                        //#endregion

                        //#region //判斷部門資料是否存在
                        //dynamicParameters = new DynamicParameters();
                        //sql = @"SELECT TOP 1 1
                        //            FROM BAS.Department
                        //            WHERE CompanyId = @CompanyId
                        //            AND DepartmentId = @DepartmentId";
                        //dynamicParameters.Add("CompanyId", CurrentCompany);
                        //dynamicParameters.Add("DepartmentId", DepartmentId);
                        //result = sqlConnection.Query(sql, dynamicParameters);
                        //if (result.Count() <= 0) throw new SystemException("部門資料找不到，請重新確認!!");
                        //#endregion

                        //#region //判斷業務資料是否存在
                        //dynamicParameters = new DynamicParameters();
                        //sql = @"SELECT TOP 1 1
                        //            FROM BAS.[User]
                        //            WHERE UserId = @UserId";
                        //dynamicParameters.Add("CompanyId", CurrentCompany);
                        //dynamicParameters.Add("UserId", SalesmenId);
                        //result = sqlConnection.Query(sql, dynamicParameters);
                        //if (result.Count() <= 0) throw new SystemException("業務資料找不到，請重新確認!!");
                        //#endregion

                        //#region //判斷收款業務資料是否存在
                        //dynamicParameters = new DynamicParameters();
                        //sql = @"SELECT TOP 1 UserNo CollectionSalesmenNo
                        //        FROM BAS.[User]
                        //        WHERE UserId = @UserId";
                        //dynamicParameters.Add("CompanyId", CurrentCompany);
                        //dynamicParameters.Add("UserId", CollectionSalesmenId);
                        //result = sqlConnection.Query(sql, dynamicParameters);
                        //if (result.Count() <= 0) throw new SystemException("收款業務資料找不到，請重新確認!!");
                        //foreach(var item in result)
                        //{
                        //    collectionSalesmenNo = item.CollectionSalesmenNo;
                        //}
                        //#endregion

                        #region //判斷出貨日要同一天
                        List<int> doDetail = new List<int>();

                        for (int i = 0; i < doData.Count(); i++)
                        {
                            int doDetailId = Convert.ToInt32(doData[i]["doDetailId"]);
                            string inventory = doData[i]["Inventory"].ToString();
                            doDetailList.Add(new DoDetail(doDetailId, inventory,"",0,0,0));
                            doDetail.Add(doDetailId);
                        }

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.DoDetailId,a2.SoErpPrefix,a2.SoErpNo,a2.SoErpPrefix +'-'+a2.SoErpNo+'-'+a1.SoSequence SoFullNo,xx.DataNum,xy.DateNum
                                ,ISNULL(zx.Qty,0) Qty,ISNULL(zy.FreeQty,0) FreeQty,ISNULL(zz.SpareQty,0) SpareQty,a.UnitPrice
                                FROM SCM.DoDetail a
                                INNER JOIN SCM.SoDetail a1 ON a.SoDetailId = a1.SoDetailId
                                INNER JOIN SCM.SaleOrder a2 ON a1.SoId = a2.SoId
                                OUTER APPLY(
                                    SELECT COUNT(*) AS DataNum
                                    FROM (
                                        SELECT DISTINCT CustomerId, DepartmentId, Currency, PaymentTerm, Taxation
                                        FROM SCM.DoDetail x1
                                        INNER JOIN SCM.SoDetail x2 ON x1.SoDetailId = x2.SoDetailId
                                        INNER JOIN SCM.SaleOrder x3 ON x2.SoId = x3.SoId
                                        WHERE x1.DoDetailId IN @DoDetail
                                    ) AS xx
                                ) xx
                                OUTER APPLY(
                                    SELECT COUNT(*) AS DateNum
                                    FROM(
                                        SELECT DISTINCT x2.DoDate
                                        FROM SCM.DoDetail x1
                                        INNER JOIN SCM.DeliveryOrder x2 ON x1.DoId = x2.DoId
                                        WHERE x1.DoDetailId IN @DoDetail
                                        AND x2.[Status] = 'S'
                                    ) AS xy
                                ) xy
                                OUTER APPLY(
                                    SELECT SUM(x1.ItemQty) Qty
                                    FROM SCM.PickingItem x1 
                                    WHERE a.SoDetailId = x1.SoDetailId
                                    AND ItemType = '1'
                                    AND x1.DoId =a.DoId
                                ) zx
                                OUTER APPLY(
                                    SELECT SUM(x1.ItemQty) FreeQty 
                                    FROM SCM.PickingItem x1 
                                    WHERE a.SoDetailId = x1.SoDetailId
                                    AND ItemType = '2'
                                    AND x1.DoId =a.DoId
                                ) zy
                                OUTER APPLY(
                                    SELECT SUM(x1.ItemQty) SpareQty
                                    FROM SCM.PickingItem x1 
                                    WHERE a.SoDetailId = x1.SoDetailId
                                    AND ItemType = '3'
                                    AND x1.DoId =a.DoId
                                ) zz
                                WHERE a.DoDetailId IN @DoDetail";
                        dynamicParameters.Add("DoDetail", doDetail.ToArray());

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("資料異常找不到,請重新確認!");
                        foreach (var item in result)
                        {
                            if (Convert.ToInt32(item.DateNum) > 1) throw new SystemException("出貨日期不同!");
                            if(Convert.ToInt32(item.DataNum) > 1) throw new SystemException("選取的訂單單頭資料不一致,不可以於同一張銷貨單開立(檢核條件【客戶代號】【部門代號】【交易幣別】【付款條件】【課稅別】)");
                            SoErpPrefix = item.SoErpPrefix;
                            SoErpNo = item.SoErpNo;
                            foreach (DoDetail item1 in doDetailList)
                            {
                                if ( item1.DoDetailId == item.DoDetailId)
                                {
                                    item1.SoFullNo = item.SoFullNo;
                                    item1.Qty = item.Qty;
                                    item1.FreeQty = item.FreeQty;
                                    item1.SpareQty = item.SpareQty;
                                    break;
                                }
                            }
                        }
                        #endregion

                    }

                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        //TG013-原幣銷貨金額
                        //TG025-原幣銷貨稅額
                        //TG045-本幣銷貨金額
                        //TG046-本幣銷貨稅額
                        //TG033-總數量

                        string CustomsClearance = "";
                        string TaxCode = "";
                        string TaxType = "";
                        string InvoiceCountNo = "";
                        int TotalQuantity = 0;
                        double TotalOriginalAmount = 0;
                        double TotalOriginalTaxAmount = 0;
                        double TotalPretaxAmount = 0;
                        double TotalTaxAmount = 0;
                        double ExchangeRate = 0;
                        int decimalOrAmoute = 0;
                        int decimalLoAmoute = 0;
                        int OraExponent = 0;
                        int LoaExponent = 0;

                        #region //ERP公司代碼取得
                        string ERPCompanyNo = "";
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 ML001  
                                         FROM CMSML";
                        var resultCompanyNo = sqlConnection.Query(sql, dynamicParameters);
                        foreach (var item in resultCompanyNo)
                        {
                            ERPCompanyNo = item.ML001;
                        }
                        if (MESCompanyNo != ERPCompanyNo) throw new SystemException("ERP的公司代碼與MES的公司代碼不同，請嘗試重新登入!!");
                        #endregion

                        #region //判斷ERP使用者是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 MF005
                                FROM ADMMF
                                WHERE COMPANY = @CompanyNo
                                AND MF001 = @userNo";
                        dynamicParameters.Add("CompanyNo", ERPCompanyNo);
                        dynamicParameters.Add("userNo", userNo);
                        var resultUserExist = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUserExist.Count() <= 0) throw new SystemException("【ERP使用者】不存在!");
                        #endregion

                        #region //審核ERP權限-建立
                        var USR_GROUP = BaseHelper.CheckErpAuthority(userNo, sqlConnection, "COPI08", "CREATE");
                        #endregion

                        #region //判斷開立日期是否超過結帳日
                        string CloseDateBase = "";
                        DateTime date = DateTime.Parse(DocDate);
                        string ReceiveDateBase = date.ToString("yyyyMMdd");
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
                                throw new SystemException("【銷貨單】ERP已經結帳,不可以在新增【" + CloseDateBase + "】之前的單據");
                            }
                        }
                        else
                        {
                            throw new SystemException("日期字符串无效，无法比较");
                        }
                        #endregion

                        #region //編碼設定

                        #region //撈取單據編號編碼格式設定
                        string WoType = "";
                        string encode = ""; // 編碼格式
                        int yearLength = 0; // 年碼數
                        int lineLength = 0; // 流水碼數
                        DateTime referenceTime = default(DateTime);
                        referenceTime = DateTime.ParseExact(dateNow, "yyyyMMdd", CultureInfo.InvariantCulture);
                        #endregion

                        #region //撈取ERP單據設定
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MQ004, a.MQ005, a.MQ006, a.MQ044, a.MQ017
                                            FROM CMSMQ a
                                            WHERE MQ001 = @RoErpPrefix";
                        dynamicParameters.Add("RoErpPrefix", RoErpPrefix);
                        var resultReceiptNo = sqlConnection.Query(sql, dynamicParameters);
                        foreach (var item in resultReceiptNo)
                        {
                            encode = item.MQ004; //編碼方式
                            yearLength = Convert.ToInt32(item.MQ005); //年碼數
                            lineLength = Convert.ToInt32(item.MQ006); //流水號碼數
                            WoType = item.MQ017;
                        }
                        #endregion

                        #region //單號自動取號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT CONVERT(int, RIGHT(ISNULL(MAX(RTRIM(LTRIM(TG002))), '" + new string('0', lineLength) + @"'), " + lineLength + @")) + 1 CurrentNum
                                FROM COPTG
                                WHERE TG001 = @RoErpPrefix";
                        dynamicParameters.Add("RoErpPrefix", RoErpPrefix);
                        #endregion

                        #region //編碼格式相關
                        string dateFormat = "";
                        switch (encode)
                        {
                            case "1": //日編
                                if (yearLength < 0 || yearLength > 4) throw new SystemException("【ERP單據性質】年碼數不能小於等於0或大於4");
                                if ((lineLength + yearLength + 4) > 11) throw new SystemException("【ERP單據性質】編碼格式碼數不能大於11碼");
                                dateFormat = new string('y', yearLength) + "MMdd";
                                sql += @" AND RTRIM(LTRIM(TG002)) LIKE @ReferenceTime  + '" + new string('_', lineLength) + @"'";
                                //sql += @" AND TA002 LIKE '%' + @ReferenceTime + '%' + '" + new string('_', lineLength) + @"'";
                                dynamicParameters.Add("ReferenceTime", referenceTime.ToString(dateFormat));

                                string tedstNo = referenceTime.ToString(dateFormat);
                                RoErpNo = referenceTime.ToString(dateFormat);
                                break;
                            case "2": //月編
                                if (yearLength < 0 || yearLength > 4) throw new SystemException("【ERP單據性質】年碼數不能小於等於0或大於4");
                                if ((lineLength + yearLength + 2) > 11) throw new SystemException("【ERP單據性質】編碼格式碼數不能大於11碼");
                                dateFormat = new string('y', yearLength) + "MM";
                                sql += @" AND RTRIM(LTRIM(TG002))  LIKE @ReferenceTime + '" + new string('_', lineLength) + @"'";
                                dynamicParameters.Add("ReferenceTime", referenceTime.ToString(dateFormat));
                                RoErpNo = referenceTime.ToString(dateFormat);
                                break;
                            case "3": //流水號
                                if (yearLength == 0) throw new SystemException("【ERP單據性質】年碼數必須等於0");
                                if (lineLength <= 0 || lineLength > 11) throw new SystemException("【ERP單據性質】流水編碼碼數必須大於0小於等於11");
                                break;
                            case "4": //手動編號
                                break;
                            default:
                                throw new SystemException("編碼方式錯誤!");
                        }

                        int currentNum = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).CurrentNum;
                        RoErpNo += string.Format("{0:" + new string('0', lineLength) + "}", currentNum);
                        #endregion

                        #endregion

                        #region //判斷ERP銷貨單單號是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM COPTG
                                WHERE TG001 = @TG001
                                AND TG002 = @TG002";
                        dynamicParameters.Add("TG001", RoErpPrefix);
                        dynamicParameters.Add("TG002", RoErpNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【銷貨單單號】重複，請重新取號!");
                        string CFIELD01 = RoErpPrefix + "-" + RoErpNo;
                        #endregion

                        foreach(var item in doDetailList)
                        {
                            TotalQuantity += item.Qty;
                            TotalQuantity += item.FreeQty;
                            TotalQuantity += item.SpareQty;
                        }

                        #region //單頭資料撈取
                        sql = @"SELECT 
                                LTRIM(RTRIM(a.TC004)) TG004, LTRIM(RTRIM(a.TC005)) TG005, LTRIM(RTRIM(a.TC006)) TG006,
                                LTRIM(RTRIM(a.TC053)) TG007, LTRIM(RTRIM(a.TC010)) TG008, LTRIM(RTRIM(a.TC011)) TG009,
                                LTRIM(RTRIM(a.TC007)) TG010, LTRIM(RTRIM(a.TC008)) TG011, LTRIM(RTRIM(a.TC016)) TG017,
                                LTRIM(RTRIM(a.TC063)) TG018, LTRIM(RTRIM(a.TC015)) TG020, LTRIM(RTRIM(a.TC041)) TG044,
                                LTRIM(RTRIM(a.TC042)) TG047, LTRIM(RTRIM(a.TC018)) TG066, LTRIM(RTRIM(a.TC019)) TG072,
                                LTRIM(RTRIM(a.TC065)) TG076, LTRIM(RTRIM(a.TC066)) TG078, LTRIM(RTRIM(a.TC067)) TG079, 
                                LTRIM(RTRIM(a.TC068)) TG082, LTRIM(RTRIM(a.TC071)) TG083, LTRIM(RTRIM(a.TC078)) TG094,
                                LTRIM(RTRIM(x.MG003)) TG012,
                                x.decimalOrPrice,x.decimalOrAmoute,
                                y.decimalLoPrice,y.decimalLoAmoute
                                FROM COPTC a
                                OUTER APPLY(
                                    SELECT TOP 1 x2.MG003 ,x1.MF003 decimalOrPrice,x1.MF004 decimalOrAmoute
                                    FROM CMSMF x1
                                    INNER JOIN CMSMG x2 on x1.MF001 = x2.MG001
                                    WHERE 1=1
                                    AND x1.MF001=a.TC008
                                    ORDER BY x2.MG002 DESC
                                ) x
                                OUTER APPLY(
                                    SELECT x1.MA003 CurrencyLocal,x2.MF003 decimalLoPrice,x2.MF004 decimalLoAmoute
                                    FROM CMSMA x1
                                    INNER JOIN CMSMF x2 on x1.MA003 = x2.MF001
                                ) y
                                WHERE 1=1
                                AND a.TC001 =@SoErpPrefix
                                AND a.TC002 =@SoErpNo
                                ";
                        dynamicParameters.Add("SoErpPrefix", SoErpPrefix);
                        dynamicParameters.Add("SoErpNo", SoErpNo);
                        ReceiveOrderErp = sqlConnection.Query<COPTG>(sql, dynamicParameters).ToList();
                        result = sqlConnection.Query(sql, dynamicParameters);
                        foreach (var item in result)
                        {
                            TaxType = item.TG017;
                            ExchangeRate = Convert.ToDouble(item.TG012);
                            decimalOrAmoute =Convert.ToInt32(item.decimalOrPrice);
                            decimalLoAmoute = Convert.ToInt32(item.decimalLoAmoute != "0" ? item.decimalLoAmoute : 1);
                            OraExponent = Convert.ToInt32(Math.Log(decimalOrAmoute) / Math.Log(10));
                            LoaExponent = decimalLoAmoute > 0 ? Convert.ToInt32(Math.Log(decimalLoAmoute) / Math.Log(10)) : 0;
                            if (TaxType == "3")
                            {
                                CustomsClearance = "2";
                            }
                            else
                            {
                                CustomsClearance = "1";
                            }
                            TaxCode = item.TG094;
                        }
                        #endregion

                        #region //單身資料撈取
                        sql = @"SELECT 
                                LTRIM(RTRIM(a.TD004)) TH004, LTRIM(RTRIM(a.TD005)) TH005, LTRIM(RTRIM(a.TD006)) TH006,
                                LTRIM(RTRIM(a.TD010)) TH009, LTRIM(RTRIM(a.TD011)) TH012, LTRIM(RTRIM(a.TD001)) TH014,
                                LTRIM(RTRIM(a.TD002)) TH015, LTRIM(RTRIM(a.TD003)) TH016, LTRIM(RTRIM(a.TD049)) TH031,
                                LTRIM(RTRIM(a.TD010)) TH062, LTRIM(RTRIM(a.TD070)) TH073, LTRIM(RTRIM(a.TD079)) TH090,
                                 0 TH010,  '' TH011,  '' TH017,
                                '' TH018, 'N' TH020, 'N' TH021,
                                '' TH022,  '' TH023,   1 TH025,
                                'N' TH026, '' TH027, '' TH028,
                                '' TH029, '' TH030, '' TH032,
                                '' TH032, '' TH034, 0 TH039,
                                0 TH040, '' TH041, 0 TH043,
                                0 TH044, 0 TH045, 0 TH046,
                                0 TH047, 0 TH048, 0 TH049,
                                0 TH050, '' TH051, 0 TH052, 0 TH053,
                                'N' TH054, '' TH055, '' TH056,
                                0 TH057, '' TH058, '' TH059,
                                '' TH060, 0 TH063, 0 TH064,
                                '' TH065, '' TH066, '' TH067,
                                'N' TH068, 0 TH069, '' TH070, 0 TH071,
                                0 TH072, '' TH074, '' TH075,
                                '' TH076, '' TH077, '' TH078,
                                '' TH079, 0 TH080, '' TH081,
                                '' TH082, '' TH083, '' TH084,
                                '' TH085, '' TH086, '' TH087,
                                '' TH088, '' TH089, 0 TH091,
                                '' TH092, '' TH093, '' TH094,
                                '' TH500, '' TH501, '' TH502,
                                '' TH503
                                FROM COPTD a
                                WHERE TD001 = @SoErpPrefix
                                AND TD002 = @SoErpNo
                                ";
                        dynamicParameters.Add("SoErpPrefix", SoErpPrefix);
                        dynamicParameters.Add("SoErpNo", SoErpNo);
                        RoDetailErp = sqlConnection.Query<COPTH>(sql, dynamicParameters).ToList();

                        #endregion

                        #region //撈取發票聯數
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT LTRIM(RTRIM(NM001)) InvoiceCountNo, LTRIM(RTRIM(NM002)) InvoiceCountName 
                                FROM CMSNM a
                                INNER JOIN CMSNN b ON a.NM001 = b.NN005
                                WHERE 1=1
                                AND b.NN001 = @Condition";
                        dynamicParameters.Add("Condition", TaxCode);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("收款業務資料找不到，請重新確認!!");
                        foreach (var item in result)
                        {
                            InvoiceCountNo = item.InvoiceCountNo;
                        }
                        #endregion

                        #region //賦值段
                        #region //銷貨單身賦值
                        int num = 1;

                        RoDetailErp
                            .ToList()
                            .ForEach(x =>
                            {
                                double Qty = doDetailList[num - 1].Qty;
                                double FreeQty = doDetailList[num - 1].FreeQty;
                                double SpareQty = doDetailList[num - 1].SpareQty;
                                string Inventory = doDetailList[num - 1].Inventory;
                                double Amount = Qty * Convert.ToDouble(x.TH012);
                                double TaxRate = x.TH073;
                                double OriginalAmount = 0;
                                double OriginalTaxAmount = 0;
                                double PretaxAmount = 0;
                                double TaxAmount = 0;
                                

                                switch (TaxType)
                                {
                                    case "1": //內含
                                        //原幣未稅金額
                                        OriginalAmount = Math.Round((Amount / (1 + TaxRate) * decimalOrAmoute / decimalOrAmoute), OraExponent);
                                        //原幣稅額
                                        OriginalTaxAmount = Math.Round((Amount - OriginalAmount), OraExponent);
                                        //本幣未稅金額
                                        PretaxAmount = Math.Round((OriginalAmount * ExchangeRate * decimalLoAmoute / decimalLoAmoute), LoaExponent);
                                        //本幣稅額
                                        TaxAmount = Math.Round((Amount * ExchangeRate - PretaxAmount), LoaExponent);
                                        break;
                                    case "2": //外加
                                        //原幣未稅金額
                                        OriginalAmount = Math.Round(Amount, LoaExponent);
                                        //原幣稅額
                                        OriginalTaxAmount = Math.Round((Math.Round(Amount * TaxRate * decimalOrAmoute) / decimalOrAmoute), LoaExponent);
                                        //本幣未稅金額
                                        PretaxAmount = Math.Round((OriginalAmount * ExchangeRate * decimalLoAmoute / decimalLoAmoute), LoaExponent);
                                        //本幣稅額
                                        TaxAmount = Math.Round((OriginalTaxAmount * ExchangeRate * decimalLoAmoute / decimalLoAmoute), LoaExponent);
                                        break;
                                    default:
                                        //原幣未稅金額
                                        OriginalAmount = Math.Round(Amount, LoaExponent);
                                        OriginalTaxAmount = 0;
                                        PretaxAmount = Math.Round((OriginalAmount * ExchangeRate), LoaExponent);
                                        TaxAmount = 0;
                                        break;
                                }

                                TotalOriginalAmount += OriginalAmount;
                                TotalOriginalTaxAmount += OriginalTaxAmount;
                                TotalPretaxAmount += PretaxAmount;
                                TotalTaxAmount += TaxAmount;

                                x.COMPANY = MESCompanyNo;
                                x.USR_GROUP = USR_GROUP;
                                x.FLAG = "1";
                                x.CREATOR = userNo;
                                x.CREATE_DATE = dateNow;
                                x.CREATE_TIME = timeNow;
                                x.CREATE_AP = BaseHelper.ClientComputer();
                                x.CREATE_PRID = "BM";
                                x.MODIFIER = "";
                                x.MODI_DATE = "";
                                x.MODI_TIME = "";
                                x.MODI_AP = "";
                                x.MODI_PRID = "";
                                x.UDF01 = "";
                                x.UDF02 = "";
                                x.UDF03 = "";
                                x.UDF04 = "";
                                x.UDF05 = "";
                                x.UDF06 = 0.0;
                                x.UDF07 = 0.0;
                                x.UDF08 = 0.0;
                                x.UDF09 = 0.0;
                                x.UDF10 = 0.0;
                                x.TH001 = RoErpPrefix;
                                x.TH002 = RoErpNo;
                                x.TH003 = string.Format("{0:0000}", num); //序號
                                x.TH007 = Inventory; //庫別
                                x.TH008 = Qty; //數量
                                x.TH013 = Amount; //金額
                                x.TH019 = ""; //客戶品號
                                x.TH024 = x.TH031 == "1" ? FreeQty : SpareQty; //贈/備品數量
                                x.TH035 = OriginalAmount; //原幣未稅金額
                                x.TH036 = OriginalTaxAmount; //原幣稅額
                                x.TH037 = PretaxAmount; //本幣未稅金額
                                x.TH038 = TaxAmount; //本幣稅額
                                x.TH042 = ""; //保稅碼
                                x.TH061 = Qty; //計價數量

                                num++;
                            });
                        #endregion

                        #region //銷貨單頭賦值
                        ReceiveOrderErp
                            .ToList()
                            .ForEach(x =>
                            {
                                x.COMPANY = MESCompanyNo;
                                x.USR_GROUP = USR_GROUP;
                                x.FLAG = "1";
                                x.CREATOR = userNo;
                                x.CREATE_DATE = dateNow;
                                x.CREATE_TIME = timeNow;
                                x.CREATE_AP = BaseHelper.ClientComputer();
                                x.CREATE_PRID = "BM";
                                x.MODIFIER = "";
                                x.MODI_DATE = "";
                                x.MODI_TIME = "";
                                x.MODI_AP = "";
                                x.MODI_PRID = "";
                                x.TG021 = x.TG021 != null ? x.TG021 : "";
                                x.CFIELD01 = CFIELD01;
                                x.UDF01 = "";
                                x.UDF02 = "";
                                x.UDF03 = "";
                                x.UDF04 = "";
                                x.UDF05 = "";
                                x.UDF06 = 0.0;
                                x.UDF07 = 0.0;
                                x.UDF08 = 0.0;
                                x.UDF09 = 0.0;
                                x.UDF10 = 0.0;
                                x.TG001 = RoErpPrefix;
                                x.TG002 = RoErpNo;
                                x.TG003 = dateNow;
                                x.TG013 = TotalOriginalAmount;
                                x.TG014 = "";
                                x.TG015 = "";
                                x.TG016 = InvoiceCountNo;
                                x.TG019 = "";
                                x.TG021 = "";
                                x.TG022 = 0;
                                x.TG023 = "N";
                                x.TG024 = "N";
                                x.TG025 = TotalOriginalTaxAmount;
                                x.TG026 = collectionSalesmenNo;
                                x.TG027 = "";
                                x.TG028 = "";
                                x.TG029 = "";
                                x.TG030 = "N";
                                x.TG031 = CustomsClearance;
                                x.TG032 = 0;
                                x.TG033 = TotalQuantity;
                                x.TG034 = "N";
                                x.TG035 = "";
                                x.TG036 = "N";
                                x.TG037 = "N";
                                x.TG038 = ApplyYYMM;
                                x.TG039 = "";
                                x.TG040 = "";
                                x.TG041 = 0;
                                x.TG042 = dateNow;
                                x.TG043 = "";
                                x.TG045 = TotalPretaxAmount;
                                x.TG046 = TotalTaxAmount;
                                x.TG048 = "";
                                x.TG049 = "";
                                x.TG050 = "";
                                x.TG051 = "";
                                x.TG052 = 0;
                                x.TG053 = 0;
                                x.TG054 = 0;
                                x.TG055 = "N";
                                x.TG056 = "N";
                                x.TG057 = "";
                                x.TG058 = "";
                                x.TG059 = "N";
                                x.TG060 = "";
                                x.TG061 = "N";
                                x.TG062 = "0";
                                x.TG063 = 0;
                                x.TG064 = "";
                                x.TG065 = "";
                                x.TG067 = "";
                                x.TG068 = "1";
                                x.TG069 = 0;
                                x.TG070 = "N";
                                x.TG071 = 0;
                                x.TG073 = "";
                                x.TG074 = "";
                                x.TG075 = "";
                                x.TG077 = "";
                                x.TG080 = "";
                                x.TG081 = "";
                                x.TG084 = 0;
                                x.TG085 = 0;
                                x.TG086 = "1";
                                x.TG087 = "";
                                x.TG088 = "";
                                x.TG089 = "N";
                                x.TG090 = "";
                                x.TG091 = 0;
                                x.TG092 = "";
                                x.TG093 = "";
                                x.TG095 = "";
                                x.TG096 = "";
                                x.TG097 = "N";
                                x.TG098 = InvNumStart;
                                x.TG099 = 0;
                                x.TG100 = "N";
                                x.TG101 = 1;
                                x.TG102 = "";
                                x.TG103 = "";
                                x.TG104 = "";
                                x.TG105 = "";
                                x.TG106 = "";
                                x.TG107 = "";
                                x.TG108 = "";
                                x.TG109 = "";
                                x.TG110 = "";
                                x.TG111 = "0";
                                x.TG112 = "0";
                                x.TG113 = "";
                                x.TG114 = "";
                                x.TG115 = "";
                                x.TG116 = "";
                                x.TG117 = "";
                                x.TG118 = "";
                                x.TG119 = "";
                                x.TG120 = "";
                                x.TG121 = "";
                                x.TG122 = "";
                                x.TG123 = "";
                                x.TG124 = "";
                                x.TG125 = "";
                                x.TG126 = "";
                                x.TG127 = "";
                                x.TG128 = "";
                                x.TG129 = "";
                                x.TG130 = "";
                                x.TG131 = "";
                                x.TG132 = "";
                                x.TG133 = "";
                            });
                        #endregion
                        #endregion


                        #region //新增段
                        #region //新增銷貨單頭 COPTG
                        sql = @"INSERT INTO COPTG (COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE, FLAG, CREATE_TIME, 
                                CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID, TG001, TG002, TG003, TG004, TG005, TG006, 
                                TG007, TG008, TG009, TG010, TG011, TG012, TG013, TG014, TG015, TG016, TG017, TG018, TG019, TG020, TG021, 
                                TG022, TG023, TG024, TG025, TG026, TG027, TG028, TG029, TG030, TG031, TG032, TG033, TG034, TG035, TG036, 
                                TG037, TG038, TG039, TG040, TG041, TG042, TG043, TG044, TG045, TG046, TG047, TG048, TG049, TG050, TG051, 
                                TG052, TG053, TG054, TG055, TG056, TG057, TG058, TG059, TG060, TG061, TG062, TG063, TG064, TG065, TG066, 
                                TG067, TG068, TG069, TG070, TG071, TG072, TG073, TG074, TG075, TG076, TG077, TG078, TG079, TG080, TG081, 
                                TG082, TG083, TG084, TG085, TG086, TG087, TG088, TG089, TG090, TG091, TG092, TG093, TG094, TG095, TG096, 
                                TG097, TG098, TG099, TG100, TG101, TG102, TG103, TG104, TG105, TG106, TG107, TG108, TG109, TG110, TG111, 
                                TG112, TG113, TG114, TG115, TG116, UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, 
                                UDF10, TG117, TG118, TG119, TG120, TG121, TG124, TG125, TG126, TG127, TG128, TG122, TG123, 
                                TG129, TG130, TG131, TG132, TG133)
                                VALUES (@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE, @FLAG, @CREATE_TIME,
                                @CREATE_AP, @CREATE_PRID, @MODI_TIME, @MODI_AP, @MODI_PRID, @TG001, @TG002, @TG003, @TG004, @TG005, @TG006,
                                @TG007, @TG008, @TG009, @TG010, @TG011, @TG012, @TG013, @TG014, @TG015, @TG016, @TG017, @TG018, @TG019, @TG020, @TG021,
                                @TG022, @TG023, @TG024, @TG025, @TG026, @TG027, @TG028, @TG029, @TG030, @TG031, @TG032, @TG033, @TG034, @TG035, @TG036,
                                @TG037, @TG038, @TG039, @TG040, @TG041, @TG042, @TG043, @TG044, @TG045, @TG046, @TG047, @TG048, @TG049, @TG050, @TG051,
                                @TG052, @TG053, @TG054, @TG055, @TG056, @TG057, @TG058, @TG059, @TG060, @TG061, @TG062, @TG063, @TG064, @TG065, @TG066,
                                @TG067, @TG068, @TG069, @TG070, @TG071, @TG072, @TG073, @TG074, @TG075, @TG076, @TG077, @TG078, @TG079, @TG080, @TG081,
                                @TG082, @TG083, @TG084, @TG085, @TG086, @TG087, @TG088, @TG089, @TG090, @TG091, @TG092, @TG093, @TG094, @TG095, @TG096,
                                @TG097, @TG098, @TG099, @TG100, @TG101, @TG102, @TG103, @TG104, @TG105, @TG106, @TG107, @TG108, @TG109, @TG110, @TG111,
                                @TG112, @TG113, @TG114, @TG115, @TG116, @UDF01, @UDF02, @UDF03, @UDF04, @UDF05, @UDF06, @UDF07, @UDF08, @UDF09,
                                @UDF10, @TG117, @TG118, @TG119, @TG120, @TG121, @TG124, @TG125, @TG126, @TG127, @TG128, @TG122, @TG123,
                                @TG129, @TG130, @TG131, @TG132, @TG133)";
                        rowsAffected += sqlConnection.Execute(sql, ReceiveOrderErp);
                        #endregion

                        #region //新增銷貨單身 COPTG
                        sql = @"INSERT INTO COPTH(COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE, 
                                FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID, 
                                TH001, TH002, TH003, TH004, TH005, TH006, TH007, TH008, TH009, TH010,
                                TH011, TH012, TH013, TH014, TH015, TH016, TH017, TH018, TH019, TH020,
                                TH021, TH022, TH023, TH024, TH025, TH026, TH027, TH028, TH029, TH030,
                                TH031, TH032, TH033, TH034, TH035, TH036, TH037, TH038, TH039, TH040,
                                TH041, TH042, TH043, TH044, TH045, TH046, TH047, TH048, TH049, TH050,
                                TH051, TH052, TH053, TH054, TH055, TH056, TH057, TH058, TH059, TH060,
                                TH061, TH062, TH063, TH064, TH065, TH066, TH067, TH068, TH069, TH070,
                                TH071, TH072, TH073, TH074, TH075, TH076, TH077, TH078, TH079, TH080,
                                TH081, TH500, TH501, TH502, TH503, TH082, TH083, TH084, TH085, TH086,
                                TH087, TH088, TH089, UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07,
                                UDF08, UDF09, UDF10, TH090, TH091, TH092, TH093, TH094)
                                VALUES(@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE, 
                                @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID,@MODI_TIME, @MODI_AP, @MODI_PRID, 
                                @TH001, @TH002, @TH003, @TH004, @TH005, @TH006, @TH007, @TH008, @TH009, @TH010,
                                @TH011, @TH012, @TH013, @TH014, @TH015, @TH016, @TH017, @TH018, @TH019, @TH020,
                                @TH021, @TH022, @TH023, @TH024, @TH025, @TH026, @TH027, @TH028, @TH029, @TH030,
                                @TH031, @TH032, @TH033, @TH034, @TH035, @TH036, @TH037, @TH038, @TH039, @TH040,
                                @TH041, @TH042, @TH043, @TH044, @TH045, @TH046, @TH047, @TH048, @TH049, @TH050,
                                @TH051, @TH052, @TH053, @TH054, @TH055, @TH056, @TH057, @TH058, @TH059, @TH060,
                                @TH061, @TH062, @TH063, @TH064, @TH065, @TH066, @TH067, @TH068, @TH069, @TH070,
                                @TH071, @TH072, @TH073, @TH074, @TH075, @TH076, @TH077, @TH078, @TH079, @TH080,
                                @TH081, @TH500, @TH501, @TH502, @TH503, @TH082, @TH083, @TH084, @TH085, @TH086,
                                @TH087, @TH088, @TH089, @UDF01, @UDF02, @UDF03, @UDF04, @UDF05, @UDF06, @UDF07,
                                @UDF08, @UDF09, @UDF10, @TH090, @TH091, @TH092, @TH093, @TH094)";
                        rowsAffected += sqlConnection.Execute(sql, RoDetailErp);
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
        #endregion

        #region //Update

        #region //UpdateReceiveOrder -- 銷貨單更新 -- Shintokuro 2024.02.20
        public string UpdateReceiveOrder(int RoId, int ViewCompanyId
            , string DocDate
            , string ReceiveDate, int CustomerId, string ProcessCode
            , string RevenueJournalCode, string CostJournalCode, string NoCredLimitControl, string CashSales
            , string SourceType, string PriceSourceTypeMain, int SourceOrderId, string SourceFull
            , int DepartmentId
            , double OriginalAmount, double OriginalTaxAmount, double TotalQuantity, double PretaxAmount, double TaxAmount
            , double TaxCurrencyRate
            , int SalesmenId, string Factory, string PaymentTerm, string TradingTerms, string Currency
            , double ExchangeRate, string BondCode, int RowCnt, string Remark
            , string CustomerFullName, string ContactPerson, string CustomerAddressFirst, string TelephoneNumber
            , string CustomerAddressSecond, string FaxNumber
            , string TaxCode
            , string InvoiceType, string TaxType, double TaxRate, string InvNumGenMethod, string UiNo
            , string InvoiceDate, string InvoiceTime, string InvNumStart, string InvNumEnd, string ApplyYYMM, string CustomsClearance
            , string CustomerFullNameOre, string InvoiceAddressFirst, string CustomerEgFullName, string InvoiceAddressSecond
            , string MultipleInvoices, string MultiTaxRate, string AttachInvWithShip, string InvoicesVoid
            , string VehicleTypeNumber
            , string InvDonationRecipient, string VehicleIDshow, string VehicleIDhide, string CreditCard4No, string InvRandomCode, string ContactEmail
            , int StaffId
            , int CollectionSalesmenId, string Remark1, string Remark2, string Remark3, string LCNO, string INVOICENO
            , string DeclarationNumber, string NewRoFull, string ShipNoticeFull, string ChangeInvCode
            , string DepositBatches
            , string SoFull, string AdvOrderFull, double OffsetAmount, double OffsetTaxAmount, string TransportMethod
            , string DispatchOrderFull, string CarNumber, string TrainNumber, string DeliveryUser
            , string Courier, string SiteCommCalcMethod, string SiteCommRate, string TotalCommAmount
            , string RoDetailData
            )
        {
            try
            {
                if (ReceiveDate.Length <= 0) throw new SystemException("【銷貨日期】不能為空!");
                if (CustomerId <= 0) throw new SystemException("【客戶】不能為空!");
                if (Factory.Length <= 0) throw new SystemException("【廠別】不能為空!");
                if (Currency.Length <= 0) throw new SystemException("【幣別】不能為空!");
                if (ExchangeRate < 0) throw new SystemException("【匯率】不能為負!");
                if (ApplyYYMM.Length < 0) throw new SystemException("【申報年月】不能為空!");
                if (DocDate.Length < 0) throw new SystemException("【單據日期】不能為空!");
                if (TaxRate < 0) throw new SystemException("【營業稅率】不能為負!");
                ApplyYYMM = ApplyYYMM.Replace("-", "");

                List<Dictionary<string, string>> RoDetailJsonList = JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(RoDetailData);
                int? nullData = null;
                int rowsAffected = 0;
                string SoErpPrefix = "";
                string SoErpNo = "";
                string SoSequence = "";
                string TransferStatusMES = "";
                string RoErpPrefix = "";
                string RoErpNo = "";
                string userNo = "";
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

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
                        dynamicParameters.Add("FunctionCode", "ReceiveOrderManagment");
                        dynamicParameters.Add("DetailCode", "update");

                        var resultUser = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUser.Count() <= 0) throw new SystemException("使用者資料錯誤!");

                        foreach (var item in resultUser)
                        {
                            userNo = item.UserNo;
                        }
                        #endregion

                        #region //撈取銷貨單單頭資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CompanyId,　a.TransferStatusMES,  a.ConfirmStatus, a.RoErpPrefix, a.RoErpNo
                                FROM SCM.ReceiveOrder a
                                WHERE a.RoId = @RoId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("RoId", RoId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("銷貨單資料找不到，請重新確認");
                        foreach (var item in result)
                        {
                            if (item.TransferStatusMES == "Y") throw new SystemException("銷貨單已經拋轉ERP，不可以修改");
                            if (item.ConfirmStatus == "Y") throw new SystemException("銷貨單已經核單，不可以修改");
                            if (item.CompanyId != ViewCompanyId) throw new SystemException("頁面的公司別與單據公司別不同，請嘗試重新登入!!");
                            TransferStatusMES = item.TransferStatusMES;
                            RoErpPrefix = item.RoErpPrefix;
                            RoErpNo = item.RoErpNo;
                        }
                        #endregion

                        #region //判斷客戶資料是否存在
                        sql = @"SELECT TOP 1 1
                                FROM SCM.Customer
                                WHERE CompanyId = @CompanyId
                                AND CustomerId = @CustomerId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("CustomerId", CustomerId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("客戶資料找不到，請重新確認!!");
                        #endregion

                        #region //判斷部門資料是否存在
                        if (DepartmentId > 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.Department
                                    WHERE CompanyId = @CompanyId
                                    AND DepartmentId = @DepartmentId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("DepartmentId", DepartmentId);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("部門資料找不到，請重新確認!!");
                        }
                        #endregion

                        #region //判斷業務資料是否存在
                        if (SalesmenId > 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.[User]
                                    WHERE UserId = @UserId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("UserId", SalesmenId);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("業務資料找不到，請重新確認!!");
                        }
                        #endregion

                        #region //判斷收款業務資料是否存在
                        if (CollectionSalesmenId > 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.[User]
                                    WHERE UserId = @UserId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("UserId", CollectionSalesmenId);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("收款業務資料找不到，請重新確認!!");
                        }
                        #endregion

                        #region //判斷員工資料是否存在
                        if (StaffId > 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.[User]
                                    WHERE UserId = @UserId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("UserId", StaffId);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("員工資料找不到，請重新確認!!");
                        }
                        #endregion

                        #region //判斷來源單據是否存在
                        switch (SourceType)
                        {
                            case "So":
                                #region //判斷訂單
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM SCM.SaleOrder
                                        WHERE CompanyId = @CompanyId
                                        AND SoId = @SourceOrderId";
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                dynamicParameters.Add("SourceOrderId", SourceOrderId);
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("【銷貨單新增】－找不到來源訂單單據，請重新確認!");
                                #endregion
                                break;
                            case "To":
                                #region //判斷暫出單
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM SCM.TempShippingNote
                                        WHERE CompanyId = @CompanyId
                                        AND TsnId = @SourceOrderId";
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                dynamicParameters.Add("SourceOrderId", SourceOrderId);
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("【銷貨單新增】－找不到來源暫出單單據，請重新確認!");
                                #endregion
                                break;
                        }
                        #endregion

                        #region //更新單頭資料表
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.ReceiveOrder SET
                                ReceiveDate = @ReceiveDate,
                                CustomerId = @CustomerId,
                                DepartmentId = @DepartmentId,
                                SalesmenId = @SalesmenId,
                                CustomerFullName = @CustomerFullName,
                                CustomerAddressFirst = @CustomerAddressFirst,
                                CustomerAddressSecond = @CustomerAddressSecond,
                                Factory = @Factory,
                                Currency = @Currency,
                                ExchangeRate = @ExchangeRate,
                                OriginalAmount = @OriginalAmount,
                                InvNumEnd = @InvNumEnd,
                                UiNo = @UiNo,
                                InvoiceType = @InvoiceType,
                                TaxType = @TaxType,
                                InvoiceAddressFirst = @InvoiceAddressFirst,
                                InvoiceAddressSecond = @InvoiceAddressSecond,
                                Remark = @Remark,
                                InvoiceDate = @InvoiceDate,
                                OriginalTaxAmount = @OriginalTaxAmount,
                                CollectionSalesmenId = @CollectionSalesmenId,
                                Remark1 = @Remark1,
                                Remark2 = @Remark2,
                                Remark3 = @Remark3,
                                InvoicesVoid = @InvoicesVoid,
                                CustomsClearance = @CustomsClearance,
                                RowCnt = @RowCnt,
                                TotalQuantity = @TotalQuantity,
                                CashSales = @CashSales,
                                StaffId = @StaffId,
                                RevenueJournalCode = @RevenueJournalCode,
                                CostJournalCode = @CostJournalCode,
                                ApplyYYMM = @ApplyYYMM,
                                LCNO = @LCNO,
                                INVOICENO = @INVOICENO,
                                DocDate = @DocDate,
                                TaxRate = @TaxRate,
                                PretaxAmount = @PretaxAmount,
                                TaxAmount = @TaxAmount,
                                PaymentTerm = @PaymentTerm,
                                OffsetAmount = @OffsetAmount,
                                OffsetTaxAmount = @OffsetTaxAmount,
                                ChangeInvCode = @ChangeInvCode,
                                ProcessCode = @ProcessCode,
                                AttachInvWithShip = @AttachInvWithShip,
                                BondCode = @BondCode,
                                ContactPerson = @ContactPerson,
                                Courier = @Courier,
                                SiteCommCalcMethod = @SiteCommCalcMethod,
                                SiteCommRate = @SiteCommRate,
                                TotalCommAmount = @TotalCommAmount,
                                TransportMethod = @TransportMethod,
                                DeclarationNumber = @DeclarationNumber,
                                FullNameOfDelivCustomer = @FullNameOfDelivCustomer,
                                TelephoneNumber = @TelephoneNumber,
                                FaxNumber = @FaxNumber,
                                TradingTerms = @TradingTerms,
                                CustomerEgFullName = @CustomerEgFullName,
                                InvNumGenMethod = @InvNumGenMethod,
                                NoCredLimitControl = @NoCredLimitControl,
                                TaxCode = @TaxCode,
                                MultipleInvoices = @MultipleInvoices,
                                InvNumStart = @InvNumStart,
                                MultiTaxRate = @MultiTaxRate,
                                TaxCurrencyRate = @TaxCurrencyRate,
                                InvoiceTime = @InvoiceTime,
                                VehicleIDshow = @VehicleIDshow,
                                VehicleTypeNumber = @VehicleTypeNumber,
                                VehicleIDhide = @VehicleIDhide,
                                InvDonationRecipient = @InvDonationRecipient,
                                InvRandomCode = @InvRandomCode,
                                CreditCard4No = @CreditCard4No,
                                ContactEmail = @ContactEmail,
                                SourceType = @SourceType,
                                PriceSourceTypeMain = @PriceSourceTypeMain,
                                SourceOrderId = @SourceOrderId,
                                SourceFull = @SourceFull,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RoId = @RoId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ReceiveDate,
                                CustomerId,
                                DepartmentId,
                                SalesmenId,
                                CustomerFullName,
                                CustomerAddressFirst,
                                CustomerAddressSecond,
                                Factory,
                                Currency,
                                ExchangeRate,
                                OriginalAmount,
                                InvNumEnd,
                                UiNo,
                                InvoiceType,
                                TaxType,
                                InvoiceAddressFirst,
                                InvoiceAddressSecond,
                                Remark,
                                InvoiceDate = InvoiceDate != "" ? InvoiceDate : null,
                                OriginalTaxAmount,
                                CollectionSalesmenId,
                                Remark1,
                                Remark2,
                                Remark3,
                                InvoicesVoid,
                                CustomsClearance,
                                RowCnt,
                                TotalQuantity,
                                CashSales,
                                StaffId,
                                RevenueJournalCode,
                                CostJournalCode,
                                ApplyYYMM,
                                LCNO,
                                INVOICENO,
                                DocDate,
                                TaxRate,
                                PretaxAmount,
                                TaxAmount,
                                PaymentTerm,
                                OffsetAmount,
                                OffsetTaxAmount,
                                ChangeInvCode,
                                ProcessCode,
                                AttachInvWithShip,
                                BondCode,
                                ContactPerson,
                                Courier,
                                SiteCommCalcMethod,
                                SiteCommRate,
                                TotalCommAmount,
                                TransportMethod,
                                DeclarationNumber,
                                FullNameOfDelivCustomer = CustomerFullName,
                                TelephoneNumber,
                                FaxNumber,
                                TradingTerms,
                                CustomerEgFullName,
                                InvNumGenMethod,
                                NoCredLimitControl,
                                TaxCode,
                                MultipleInvoices,
                                InvNumStart,
                                MultiTaxRate,
                                TaxCurrencyRate = 1,
                                InvoiceTime,
                                VehicleIDshow,
                                VehicleTypeNumber,
                                VehicleIDhide,
                                InvDonationRecipient,
                                InvRandomCode,
                                CreditCard4No,
                                ContactEmail,
                                SourceType,
                                PriceSourceTypeMain,
                                SourceOrderId,
                                SourceFull,
                                LastModifiedDate,
                                LastModifiedBy,
                                RoId,
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //清空銷貨單單身
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.RoDetail
                                WHERE RoId = @RoId";
                        dynamicParameters.Add("RoId", RoId);
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        if (RoDetailJsonList.Count() > 0)
                        {
                            foreach (var item in RoDetailJsonList)
                            {
                                string MtlItemNo = item["MtlItemNo"] != null ? item["MtlItemNo"].ToString() : throw new SystemException("【資料維護不完整】品號欄位資料不可以為空,請重新確認~~");
                                string InventoryNo = item["InventoryNo"] != null ? item["InventoryNo"].ToString() : throw new SystemException("【資料維護不完整】庫別欄位資料不可以為空,請重新確認~~");
                                string Location = item["Location"] != null ? item["Location"].ToString() : "";
                                Double Quantity = item["Quantity"] != null ? Convert.ToDouble(item["Quantity"].ToString()) : 0;
                                string Type = item["Type"] != null ? item["Type"].ToString() : throw new SystemException("【資料維護不完整】類別欄位資料不可以為空,請重新確認~~");
                                Double FreeSpareQty = item["FreeSpareQty"] != null ? Convert.ToDouble(item["FreeSpareQty"].ToString()) : 0;
                                string UomNo = item["UomNo"] != null ? item["UomNo"].ToString() : throw new SystemException("【資料維護不完整】單位欄位資料不可以為空,請重新確認~~");
                                Double AvailableQty = item["AvailableQty"] != null ? Convert.ToDouble(item["AvailableQty"].ToString()) : 0;
                                string AvailableUomNo = item["AvailableUomNo"] != null ? item["AvailableUomNo"].ToString() : throw new SystemException("【資料維護不完整】計價單位欄位資料不可以為空,請重新確認~~");
                                Double UnitPrice = item["UnitPrice"] != null ? Convert.ToDouble(item["UnitPrice"].ToString()) : 0;
                                Double TaxRateDetail = item["TaxRate"] != null ? Convert.ToDouble(item["TaxRate"].ToString()) : 0;
                                Double DiscountRate = item["DiscountRate"] != null ? Convert.ToDouble(item["DiscountRate"].ToString()) : 0;
                                Double Amount = item["Amount"] != null ? Convert.ToDouble(item["Amount"].ToString()) : 0;
                                Double OrigPreTaxAmt = item["OrigPreTaxAmt"] != null ? Convert.ToDouble(item["OrigPreTaxAmt"].ToString()) : 0;
                                Double OrigTaxAmt = item["OrigTaxAmt"] != null ? Convert.ToDouble(item["OrigTaxAmt"].ToString()) : 0;
                                Double PreTaxAmt = item["PreTaxAmt"] != null ? Convert.ToDouble(item["PreTaxAmt"].ToString()) : 0;
                                Double TaxAmt = item["TaxAmt"] != null ? Convert.ToDouble(item["TaxAmt"].ToString()) : 0;
                                string SoFullDetail = item["SoFull"] != null ? item["SoFull"].ToString() : "";
                                string LotNumber = item["LotNumber"] != null ? item["LotNumber"].ToString() : "";
                                string CustomerMtlItemNo = item["CustomerMtlItemNo"] != null ? item["CustomerMtlItemNo"].ToString() : "";
                                string ProjectCode = item["ProjectCode"] != null ? item["ProjectCode"].ToString() : "";
                                string Remarks = item["Remarks"] != null ? item["Remarks"].ToString() : "";
                                string TsnFull = item["TsnFull"] != null ? item["TsnFull"].ToString() : "";
                                string BondedCode = item["BondedCode"] != null ? item["BondedCode"].ToString() : "";
                                string CustomerPurchaseOrder = item["CustomerPurchaseOrder"] != null ? item["CustomerPurchaseOrder"].ToString() : "";
                                string ForecastFull = item["ForecastFull"] != null ? item["ForecastFull"].ToString() : "";

                                #region //匯率等於1時判定
                                if (ExchangeRate == 1)
                                {
                                    if (OrigPreTaxAmt != PreTaxAmt) throw new SystemException("單據幣別與本國貨幣相同時,原幣金額與本幣金額須相同!!");
                                    if (OrigTaxAmt != TaxAmt) throw new SystemException("單據幣別與本國貨幣相同時,原幣稅額與本幣稅額須相同!!");
                                }
                                #endregion

                                #region //判斷品號是否存在
                                int MtlItemId = -1;
                                string OverDeliveryManagement = "";
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 MtlItemId,OverDeliveryManagement
                                        FROM PDM.MtlItem
                                        WHERE CompanyId = @CompanyId
                                        AND MtlItemNo = @MtlItemNo";
                                dynamicParameters.Add("MtlItemNo", MtlItemNo);
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("【品號:" + MtlItemNo + "】找不到，請重新輸入!");
                                foreach (var item2 in result)
                                {
                                    MtlItemId = item2.MtlItemId;
                                    OverDeliveryManagement = item2.OverDeliveryManagement;
                                }
                                #endregion

                                #region //判斷庫別是否存在
                                int InventoryId = -1;
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 InventoryId
                                        FROM SCM.Inventory
                                        WHERE CompanyId = @CompanyId
                                        AND InventoryNo = @InventoryNo";
                                dynamicParameters.Add("InventoryNo", InventoryNo.Split(':')[0]);
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("【庫別:" + InventoryNo + "】找不到，請重新輸入!");
                                foreach (var item2 in result)
                                {
                                    InventoryId = item2.InventoryId;
                                }
                                #endregion

                                #region //判斷單位是否存在
                                int UomId = -1;
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 UomId
                                        FROM PDM.UnitOfMeasure
                                        WHERE CompanyId = @CompanyId
                                        AND UomNo = @UomNo";
                                dynamicParameters.Add("UomNo", UomNo);
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("【單位:" + UomNo + "】找不到，請重新輸入!");
                                foreach (var item2 in result)
                                {
                                    UomId = item2.UomId;
                                }
                                #endregion

                                #region //判斷計價單位是否存在
                                int AvailableUomId = -1;
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 UomId
                                        FROM PDM.UnitOfMeasure
                                        WHERE CompanyId = @CompanyId
                                        AND UomNo = @UomNo";
                                dynamicParameters.Add("UomNo", AvailableUomNo);
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("【計價單位:" + AvailableUomNo + "】找不到，請重新輸入!");
                                foreach (var item2 in result)
                                {
                                    AvailableUomId = item2.UomId;
                                }
                                #endregion

                                #region //判斷訂單單身是否存在
                                int? SoDetailId = null;
                                string ProductType = "";
                                Double FreebieQty = 0;
                                Double SpareQty = 0;
                                if (SoFullDetail.Length > 0)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 a.SoDetailId, a.ProductType, a.FreebieQty, a.SpareQty
                                            ,b.SoErpPrefix, b.SoErpNo, a.SoSequence, a.ClosureStatus
                                            FROM SCM.SoDetail a
                                            INNER JOIN SCM.SaleOrder b on a.SoId = b.SoId
                                            WHERE b.CompanyId = @CompanyId
                                            AND b.SoErpPrefix = @SoErpPrefix
                                            AND b.SoErpNo = @SoErpNo
                                            AND a.SoSequence = @SoSequence";
                                    dynamicParameters.Add("CompanyId", CurrentCompany);
                                    dynamicParameters.Add("SoErpPrefix", SoFullDetail.Split('-')[0]);
                                    dynamicParameters.Add("SoErpNo", SoFullDetail.Split('-')[1]);
                                    dynamicParameters.Add("SoSequence", SoFullDetail.Split('-')[2]);
                                    result = sqlConnection.Query(sql, dynamicParameters);
                                    if (result.Count() <= 0) throw new SystemException("【訂單:" + SoFullDetail + "】找不到，請重新輸入!");
                                    foreach (var item2 in result)
                                    {
                                        if(item2.ClosureStatus == "Y") throw new SystemException("【訂單:" + SoFullDetail + "】已經結案!!!");
                                        SoErpPrefix = item2.SoErpPrefix;
                                        SoErpNo = item2.SoErpNo;
                                        SoSequence = item2.SoSequence;
                                        SoDetailId = item2.SoDetailId;
                                        ProductType = item2.ProductType;
                                        FreebieQty = item2.FreebieQty;
                                        SpareQty = item2.SpareQty;
                                    }
                                }
                                #endregion

                                #region //判斷暫出單身是否存在
                                int? TsnDetailId = null;
                                string TsnProductType = "";
                                Double FreebieOrSpareQty = 0;
                                if (TsnFull.Length > 0)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 b.TsnDetailId,b.ProductType,b.FreebieOrSpareQty,b.ClosureStatus
                                            FROM SCM.TempShippingNote a
                                            INNER JOIN SCM.TsnDetail b on a.TsnId = b.TsnId
                                            WHERE CompanyId = @CompanyId
                                            AND a.TsnErpPrefix = @TsnErpPrefix
                                            AND a.TsnErpNo = @TsnErpNo
                                            AND b.TsnSequence = @TsnSequence
                                            AND a.ConfirmStatus = 'Y'";
                                    dynamicParameters.Add("CompanyId", CurrentCompany);
                                    dynamicParameters.Add("TsnErpPrefix", TsnFull.Split('-')[0]);
                                    dynamicParameters.Add("TsnErpNo", TsnFull.Split('-')[1]);
                                    dynamicParameters.Add("TsnSequence", TsnFull.Split('-')[2]);
                                    result = sqlConnection.Query(sql, dynamicParameters);
                                    if (result.Count() <= 0) throw new SystemException("【暫出單:" + TsnFull + "】找不到，請重新輸入!");
                                    foreach (var item2 in result)
                                    {
                                        if(item2.ClosureStatus == "Y") throw new SystemException("【暫出單:" + TsnFull + "】已經結案!!!");
                                        TsnDetailId = item2.TsnDetailId;
                                        TsnProductType = item2.ProductType;
                                        FreebieOrSpareQty = item2.FreebieOrSpareQty;
                                    }
                                }
                                #endregion

                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO SCM.RoDetail (RoId, RoSequence, MtlItemId, InventoryId, Quantity, UomId, UnitPrice, Amount, 
                                        SoDetailId, LotNumber, Remarks, CustomerMtlItemNo, ConfirmationCode, UpdateCode, FreeSpareQty, DiscountRate, 
                                        CheckOutCode, CheckOutPrefix, CheckOutNo, CheckOutSequence, ProjectCode, Type, TsnDetailId, OrigPreTaxAmt, 
                                        OrigTaxAmt, PreTaxAmt, TaxAmt, PackageQty, PackageFreeSpareQty, PackageUomId, BondedCode, SrQty, 
                                        SrPackageQty, SrOrigPreTaxAmt, SrOrigTaxAmt, SrPreTaxAmt, SrTaxAmt, CommissionRate, CommissionAmount, 
                                        OriginalCustomer, SrFreeSpareQty, SrPackageFreeSpareQty, NotPayTemp, ProductSerialNumberQty, ForecastCode, 
                                        ForecastSequence, Location, AvailableQty, AvailableUomId, MultiBatch, FreeSpareRate, FinalCustomerCode, 
                                        ReferenceQty, ReferencePackageQty, TaxRate, CRMSource, CRMPrefix, CRMNo, CRMSequence, CRMContractCode, 
                                        CRMAllowDeduction, CRMDeductionQty, CRMDeductionUnit, DebitAccount, CreditAccount, TaxAmountAccount, 
                                        BusinessItemNumber, TaxCode, DiscountAmount, K2NO, MarkingBINRecord, MarkingManagement, 
                                        BillingUnitInPackage, DATECODE, TransferStatusMES, CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        VALUES ( @RoId, @RoSequence, @MtlItemId, @InventoryId, @Quantity, @UomId, @UnitPrice, @Amount, 
                                        @SoDetailId, @LotNumber, @Remarks, @CustomerMtlItemNo, @ConfirmationCode, @UpdateCode, @FreeSpareQty, @DiscountRate, 
                                        @CheckOutCode, @CheckOutPrefix, @CheckOutNo, @CheckOutSequence, @ProjectCode, @Type, @TsnDetailId, @OrigPreTaxAmt, 
                                        @OrigTaxAmt, @PreTaxAmt, @TaxAmt, @PackageQty, @PackageFreeSpareQty, @PackageUomId, @BondedCode, @SrQty, 
                                        @SrPackageQty, @SrOrigPreTaxAmt, @SrOrigTaxAmt, @SrPreTaxAmt, @SrTaxAmt, @CommissionRate, @CommissionAmount, 
                                        @OriginalCustomer, @SrFreeSpareQty, @SrPackageFreeSpareQty, @NotPayTemp, @ProductSerialNumberQty, @ForecastCode, 
                                        @ForecastSequence, @Location, @AvailableQty, @AvailableUomId, @MultiBatch, @FreeSpareRate, @FinalCustomerCode, 
                                        @ReferenceQty, @ReferencePackageQty, @TaxRate, @CRMSource, @CRMPrefix, @CRMNo, @CRMSequence, @CRMContractCode, 
                                        @CRMAllowDeduction, @CRMDeductionQty, @CRMDeductionUnit, @DebitAccount, @CreditAccount, @TaxAmountAccount, 
                                        @BusinessItemNumber, @TaxCode, @DiscountAmount, @K2NO, @MarkingBINRecord, @MarkingManagement, 
                                        @BillingUnitInPackage, @DATECODE, @TransferStatusMES, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy
                                        )";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        RoId,
                                        RoSequence = "",
                                        MtlItemId,
                                        InventoryId,
                                        Quantity,
                                        UomId,
                                        UnitPrice,
                                        Amount,
                                        SoDetailId,
                                        LotNumber,
                                        Remarks,
                                        CustomerMtlItemNo,
                                        ConfirmationCode = "N",
                                        UpdateCode = "N",
                                        FreeSpareQty,
                                        DiscountRate,
                                        CheckOutCode = "N",
                                        CheckOutPrefix = "",
                                        CheckOutNo = "",
                                        CheckOutSequence = "",
                                        ProjectCode,
                                        Type = Type.Split('.')[0],
                                        TsnDetailId,
                                        OrigPreTaxAmt,
                                        OrigTaxAmt,
                                        PreTaxAmt,
                                        TaxAmt,
                                        PackageQty = 0,
                                        PackageFreeSpareQty = 0,
                                        PackageUomId = "",
                                        BondedCode,
                                        SrQty = 0,
                                        SrPackageQty = 0,
                                        SrOrigPreTaxAmt = 0,
                                        SrOrigTaxAmt = 0,
                                        SrPreTaxAmt = 0,
                                        SrTaxAmt = 0,
                                        CommissionRate = 0,
                                        CommissionAmount = 0,
                                        OriginalCustomer = "",
                                        SrFreeSpareQty = 0,
                                        SrPackageFreeSpareQty = 0,
                                        NotPayTemp = "N",
                                        ProductSerialNumberQty = 0,
                                        ForecastCode = "",
                                        ForecastSequence = "",
                                        Location = Location.Split(':')[0],
                                        AvailableQty,
                                        AvailableUomId,
                                        MultiBatch = "N",
                                        FreeSpareRate = 0,
                                        FinalCustomerCode = "",
                                        ReferenceQty = 0,
                                        ReferencePackageQty = 0,
                                        TaxRate,
                                        CRMSource = "",
                                        CRMPrefix = "",
                                        CRMNo = "",
                                        CRMSequence = "",
                                        CRMContractCode = "",
                                        CRMAllowDeduction = "",
                                        CRMDeductionQty = 0,
                                        CRMDeductionUnit = "",
                                        DebitAccount = "",
                                        CreditAccount = "",
                                        TaxAmountAccount = "",
                                        BusinessItemNumber = "",
                                        TaxCode = "",
                                        DiscountAmount = 0,
                                        K2NO = "",
                                        MarkingBINRecord = "",
                                        MarkingManagement = "",
                                        BillingUnitInPackage = "",
                                        DATECODE = "",
                                        TransferStatusMES = "Y",
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        IEnumerable<dynamic> result = null;

                        switch (SourceType)
                        {
                            case "So":
                                foreach (var item in RoDetailJsonList)
                                {
                                    string MtlItemNo = item["MtlItemNo"];
                                    string SoFullDetail = item["SoFull"] != null ? item["SoFull"].ToString() : "";
                                    Double Quantity = item["Quantity"] != null ? Convert.ToDouble(item["Quantity"].ToString()) : 0;
                                    Double FreeSpareQty = item["FreeSpareQty"] != null ? Convert.ToDouble(item["FreeSpareQty"].ToString()) : 0;
                                    Double AvailableQty = item["AvailableQty"] != null ? Convert.ToDouble(item["AvailableQty"].ToString()) : 0;
                                    SoErpPrefix = SoFullDetail.Split('-')[0];
                                    SoErpNo = SoFullDetail.Split('-')[1];
                                    SoSequence = SoFullDetail.Split('-')[2];

                                    #region //撈取訂單品號的數量+贈/備品數量
                                    string ReEditCondition = "";
                                    if (TransferStatusMES == "R")
                                    {
                                        string RoErpFullNo = RoErpPrefix + "-" + RoErpNo;
                                        ReEditCondition = @" AND (x1.TH001 + '-' + x1.TH002) != '"+ RoErpFullNo + @"'";
                                    }
                                    #endregion

                                    #region //撈取訂單品號的數量+贈/備品數量
                                    string ProductType = "";
                                    Double CanUseQty = 0;
                                    Double CanUseFreebieQty = 0;
                                    Double CanUseSpareQty = 0;
                                    Double CanUseAvailableQty = 0;
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.TD003 SoSequence, a.TD049 ProductType
                                            ,(ISNULL(a.TD008,0) -  ISNULL(a.TD009,0) - ISNULL(x.UnconfirmedQty,0)) CanUseQty 
                                            ,(ISNULL(a.TD024,0) -  ISNULL(a.TD025,0) - ISNULL(x.UnconfirmedFreeSpareQty,0)) CanUseFreebieQty 
                                            ,(ISNULL(a.TD050,0) -  ISNULL(a.TD051,0) - ISNULL(x.UnconfirmedFreeSpareQty,0)) CanUseSpareQty 
                                            ,(ISNULL(a.TD076,0) -  ISNULL(a.TD078,0) - ISNULL(x.UnconfirmedAvailableQty,0)) CanUseAvailableQty
                                            ,ISNULL(x.UnconfirmedQty,0) UnconfirmedQty
                                            ,ISNULL(x.UnconfirmedFreeSpareQty,0) UnconfirmedFreeSpareQty 
                                            ,ISNULL(x.UnconfirmedAvailableQty,0) UnconfirmedAvailableQty
                                            FROM COPTD a
                                            OUTER APPLY(
                                                SELECT SUM(ISNULL(x1.TH008,0)) UnconfirmedQty 
                                                      ,SUM(ISNULL(x1.TH024,0)) UnconfirmedFreeSpareQty 
                                                      ,SUM(ISNULL(x1.TH061,0)) UnconfirmedAvailableQty
                                                FROM COPTH x1
                                                WHERE x1.TH014 = a.TD001
                                                AND x1.TH015 = a.TD002
                                                AND x1.TH016 = a.TD003
                                                AND x1.TH020 ='N'"+
                                                ReEditCondition
                                                + @"
                                            ) x
                                            WHERE a.TD001 = @TG014 
                                            AND a.TD002 = @TG015  
                                            AND a.TD003 = @TG016
                                            AND a.TD021 = 'Y'
                                            ";
                                    dynamicParameters.Add("TG014", SoErpPrefix);
                                    dynamicParameters.Add("TG015", SoErpNo);
                                    dynamicParameters.Add("TG016", SoSequence);
                                    var resultCOPTD = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultCOPTD.Count() > 0)
                                    {
                                        foreach (var item2 in resultCOPTD)
                                        {
                                            ProductType = item2.ProductType;
                                            CanUseQty = Convert.ToDouble(item2.CanUseQty);
                                            CanUseFreebieQty = Convert.ToDouble(item2.CanUseFreebieQty);
                                            CanUseSpareQty = Convert.ToDouble(item2.CanUseSpareQty);
                                            CanUseAvailableQty = Convert.ToDouble(item2.CanUseAvailableQty);
                                        }
                                    }
                                    if(Quantity > CanUseQty) throw new SystemException("【銷貨量超交】:品號『" + MtlItemNo + "』本次最大可銷貨數量為" + CanUseQty + "，請重新輸入!");
                                    if(AvailableQty > CanUseAvailableQty) throw new SystemException("【銷貨量超交】:品號『" + MtlItemNo + "』本次最大可銷貨計價數量為" + CanUseAvailableQty + "，請重新輸入!");
                                    if (ProductType == "1")
                                    {
                                        if (FreeSpareQty > CanUseFreebieQty) throw new SystemException("【贈品量超交】:品號『" + MtlItemNo + "』本次最大可銷貨贈品數量為" + CanUseFreebieQty + "】，請重新輸入!");
                                    }
                                    else if (ProductType == "2")
                                    {
                                        if (FreeSpareQty > CanUseSpareQty) throw new SystemException("【備品量超交】:品號『" + MtlItemNo + "』本次最大可銷貨備品數量為" + CanUseSpareQty + "】，請重新輸入!");
                                    }
                                    else
                                    {
                                        throw new SystemException("贈備品類別異常,請重新確認");
                                    }
                                    #endregion
                                }
                                result = sqlConnection.Query(sql, dynamicParameters);
                                break;
                            case "To":
                                foreach (var item in RoDetailJsonList)
                                {
                                    string MtlItemNo = item["MtlItemNo"];
                                    string TsnFull = item["TsnFull"] != null ? item["TsnFull"].ToString() : "";
                                    Double Quantity = item["Quantity"] != null ? Convert.ToDouble(item["Quantity"].ToString()) : 0;
                                    Double FreeSpareQty = item["FreeSpareQty"] != null ? Convert.ToDouble(item["FreeSpareQty"].ToString()) : 0;
                                    string TsnErpPrefix = TsnFull.Split('-')[0];
                                    string TsnErpNo = TsnFull.Split('-')[1];
                                    string TsnSequence = TsnFull.Split('-')[2];

                                    #region //撈取訂單品號的數量+贈/備品數量
                                    string ReEditCondition = "";
                                    if (TransferStatusMES == "R")
                                    {
                                        string RoErpFullNo = RoErpPrefix + "-" + RoErpNo;
                                        ReEditCondition = @" AND (x1.TH001 + '-' + x1.TH002) != '" + RoErpFullNo + @"'";
                                    }
                                    #endregion

                                    #region //撈取暫出單的可銷貨數量+可銷貨贈/備品數量
                                    Double CanUseQty = 0;
                                    Double CanUseFreeSpareQty = 0;
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT 
                                             (ISNULL(a.TG009,0) - ISNULL(a.TG020,0) - ISNULL(a.TG021,0) - ISNULL(b.UnconfirmedQty,0)) CanUseQty
                                            ,(ISNULL(a.TG044,0) - ISNULL(a.TG046,0) - ISNULL(a.TG048,0) - ISNULL(b.UnconfirmedFreebieQty,0)) CanUseFreeSpareQty 
                                            ,(ISNULL(a.TG052,0) - ISNULL(a.TG054,0) - ISNULL(a.TG055,0) - ISNULL(b.UnconfirmedAvailableQty,0)) CanUseAvailableQty 
                                            FROM INVTG a
                                            OUTER APPLY(
                                                SELECT SUM(ISNULL(x1.TH008,0)) UnconfirmedQty 
                                                      ,SUM(ISNULL(x1.TH024,0)) UnconfirmedFreebieQty 
                                                      ,SUM(ISNULL(x1.TH061,0)) UnconfirmedAvailableQty
                                                FROM COPTH x1
                                                WHERE x1.TH032 = a.TG001
                                                AND x1.TH033 = a.TG002
                                                AND x1.TH034 = a.TG003
                                                AND x1.TH020 ='N'" +
                                                ReEditCondition
                                                + @"
                                            ) b
                                            WHERE a.TG001 = @TG001
                                            AND a.TG002 = @TG002
                                            AND a.TG003 = @TG003
                                            AND a.TG022 ='Y'
                                            ";
                                    dynamicParameters.Add("TG001", TsnErpPrefix);
                                    dynamicParameters.Add("TG002", TsnErpNo);
                                    dynamicParameters.Add("TG003", TsnSequence);
                                    var resultINVTG = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultINVTG.Count() > 0)
                                    {
                                        foreach (var item2 in resultINVTG)
                                        {
                                            CanUseQty = Convert.ToDouble(item2.CanUseQty);
                                            CanUseFreeSpareQty = Convert.ToDouble(item2.CanUseFreeSpareQty);
                                        }
                                    }
                                    #endregion

                                    #region //判斷可銷貨數量是否有超收

                                    if (CanUseQty < Quantity) throw new SystemException("【銷貨量超交】:品號『" + MtlItemNo + "』本次最大可銷貨量為" + (CanUseQty) + "，請重新輸入!");
                                    if (CanUseFreeSpareQty < FreeSpareQty) throw new SystemException("【贈品量超交】:品號『" + MtlItemNo + "』本次最大可贈備品量為" + (CanUseFreeSpareQty) + "】，請重新輸入!");
                                    #endregion
                                }
                                result = sqlConnection.Query(sql, dynamicParameters);
                                break;
                        }


                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
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

        #region //UpdateReceiveOrderToERP -- 銷貨單拋轉ERP -- Shintokuro 2024-02-19
        public string UpdateReceiveOrderToERP(int RoId)
        {
            try
            {
                if (RoId < 0) throw new SystemException("銷貨單不可為空,請重新確認!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    string dateNow = CreateDate.ToString("yyyyMMdd"), timeNow = CreateDate.ToString("HH:mm:ss"), errorMsg = "";
                    string uomNo = "", mtlItemNo = "", departmentNo = "", MESCompanyNo = "", userNo = "", userName = "", userGroup = "";
                    string RoErpPrefix ="", RoErpNo = "", TransferStatusMES = "", ConfirmStatus = "", SourceType ="", ReceiveDate="";
                    int OrderCompanyIdBase = -1;

                    List<COPTG> ReceiveOrders = new List<COPTG>();
                    List<COPTH> RoDetails = new List<COPTH>();


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
                            MESCompanyNo = item.ErpNo;
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
                        dynamicParameters.Add("FunctionCode", "ReceiveOrderManagment");
                        dynamicParameters.Add("DetailCode", "import");

                        var resultUser = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUser.Count() <= 0) throw new SystemException("使用者資料錯誤!");

                        foreach (var item in resultUser)
                        {
                            userNo = item.UserNo;
                            userName = item.UserName;
                            departmentNo = item.DepartmentNo;
                        }
                        #endregion

                        #region //撈取銷貨單單頭資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.TransferStatusMES, a.CompanyId, a.RoErpPrefix TG001, a.RoErpNo TG002, FORMAT(a.ReceiveDate, 'yyyyMMdd') TG003, b.CustomerNo TG004 , ISNULL( c.DepartmentNo,'' ) TG005
                                , ISNULL( d.UserNo,'' ) TG006, a.CustomerFullName TG007, a.CustomerAddressFirst TG008, a.CustomerAddressSecond TG009, a.Factory TG010
                                , a.Currency TG011, a.ExchangeRate TG012, a.OriginalAmount TG013, a.InvNumEnd TG014, a.UiNo TG015
                                , a.InvoiceType TG016, a.TaxType TG017, a.InvoiceAddressFirst TG018, a.InvoiceAddressSecond TG019, a.Remark TG020
                                , FORMAT(a.InvoiceDate, 'yyyyMMdd') TG021, a.PrintCount TG022, a.ConfirmStatus TG023, a.UpdateCode TG024, a.OriginalTaxAmount TG025
                                , ISNULL( d1.UserNo,'' ) TG026, a.Remark1 TG027, a.Remark2 TG028, a.Remark3 TG029, a.InvoicesVoid TG030
                                , a.CustomsClearance TG031, a.RowCnt TG032, a.TotalQuantity TG033, a.CashSales TG034, ISNULL( d2.UserNo,'' ) TG035
                                , a.RevenueJournalCode TG036, a.CostJournalCode TG037,a.ApplyYYMM TG038, a.LCNO TG039, a.INVOICENO TG040
                                , a.InvoicesPrintCount TG041, FORMAT(a.DocDate, 'yyyyMMdd') TG042, ISNULL( d3.UserNo,'' ) TG043, a.TaxRate TG044, a.PretaxAmount TG045
                                , a.TaxAmount TG046, a.PaymentTerm TG047, ISNULL( e2.SoErpPrefix,'' ) TG048, ISNULL( e2.SoErpNo,'' ) TG049 , a.AdvOrderPrefix TG050
                                , a.AdvOrderNo TG051, a.OffsetAmount TG052, a.OffsetTaxAmount TG053, a.TotalPackages TG054, a.SignatureStatus TG055
                                , a.ChangeInvCode TG056, a.NewRoPrefix TG057, a.NewRoNo TG058, a.TransferStatusERP TG059, a.ProcessCode TG060
                                , a.AttachInvWithShip TG061, a.BondCode TG062, a.TransmissionCount TG063, a.Invoicer TG064, a.InvCode TG065
                                , a.ContactPerson TG066, a.Courier TG067, a.SiteCommCalcMethod TG068, a.SiteCommRate TG069, a.CommCalcInclTax TG070
                                , a.TotalCommAmount TG071, a.TransportMethod TG072, a.DispatchOrderPrefix TG073, a.DispatchOrderNo TG074, a.DeclarationNumber TG075
                                , a.FullNameOfDelivCustomer TG076, a.RoPriceType TG077, a.TelephoneNumber TG078, a.FaxNumber TG079, a.ShipNoticePrefix TG080
                                , a.ShipNoticeNo TG081, a.TradingTerms TG082, a.CustomerEgFullName TG083, 0 TG084, 0 TG085
                                , a.InvNumGenMethod TG086, a.DocSourceCode TG087, '' TG088, a.NoCredLimitControl TG089, a.InstallmentSettlement TG090
                                , a.InstallmentCount TG091, a.AutoApportionByInstallment TG092, a.StartYearMonth TG093, a.TaxCode TG094, '' TG095
                                , '' TG096, a.MultipleInvoices TG097, a.InvNumStart TG098, a.NumberOfInvoices TG099, a.MultiTaxRate TG100
                                , a.TaxCurrencyRate TG101, ''  TG102, a.VoidApprovalDocNum TG103, a.VoidReason TG104, a.VoidReason TG105
                                , a.[Source] TG106, a.IncomeDraftID TG107, a.IncomeDraftSeq TG108, a.IncomeVoucherType TG109, a.IncomeVoucherNumber TG110
                                , a.[Status] TG111, a.ZeroTaxForBuyer TG112, '' TG113, '' TG114, '' TG115
                                , a.GenLedgerAcctType TG116, a.InvoiceTime TG117, a.InvCode2 TG118, a.InvCode2 TG119, a.DeliveryCountry TG120
                                , '' TG121, '' TG122, '' TG123, a.VehicleIDshow TG124, a.VehicleTypeNumber TG125
                                , a.VehicleIDhide TG126, a.InvDonationRecipient TG127, a.InvRandomCode TG128, a.ReservedField TG129, a.CreditCard4No TG130
                                , a.ContactEmail TG131,'' TG132, a.OrigInvNumber TG133
                                ,a.DepartmentId, a.SoDetailId ,a.SourceType
                                FROM SCM.ReceiveOrder a
                                INNER JOIN SCM.Customer b on a.CustomerId = b.CustomerId
                                LEFT JOIN BAS.Department c on a.DepartmentId = c.DepartmentId
                                LEFT JOIN BAS.[User] d on a.SalesmenId = d.UserId
                                LEFT JOIN BAS.[User] d1 on a.CollectionSalesmenId = d1.UserId
                                LEFT JOIN BAS.[User] d2 on a.StaffId = d2.UserId
                                LEFT JOIN BAS.[User] d3 on a.ConfirmUserId = d3.UserId
                                LEFT JOIN SCM.SoDetail e1 on a.SoDetailId = e1.SoDetailId
                                LEFT JOIN SCM.SaleOrder e2 on e1.SoId = e2.SoId
                                WHERE a.CompanyId = @CompanyId
                                AND a.RoId = @RoId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("RoId", RoId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count()<=0) throw new SystemException("銷貨單單頭資料找不到，請重新確認");
                        foreach (var item in result)
                        {
                            SourceType = item.SourceType;
                            TransferStatusMES = item.TransferStatusMES;
                            ConfirmStatus = item.ConfirmStatus;
                            OrderCompanyIdBase = item.CompanyId;
                            RoErpPrefix = item.TG001;
                            RoErpNo = item.TG002;
                            ReceiveDate = item.TG003;
                        }
                        if (ConfirmStatus == "Y") throw new SystemException("銷貨單處於確認狀態不可以拋轉");
                        if (ConfirmStatus == "V") throw new SystemException("銷貨單處於作廢狀態不可以拋轉");
                        if (OrderCompanyIdBase != CurrentCompany) throw new SystemException("單據的公司別與後端公司別不同，請嘗試重新登入!!");
                        ReceiveOrders = sqlConnection.Query<COPTG>(sql, dynamicParameters).ToList();
                        #endregion

                        #region //撈取銷貨單單身資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.TransferStatusMES,a.RoDetailId, b.RoErpPrefix TH001,b.RoErpNo TH002,FORMAT(ROW_NUMBER() OVER(ORDER BY a.RoDetailId),'0000') TH003, c.MtlItemNo TH004,c.MtlItemName TH005
                                , c.MtlItemSpec TH006, d.InventoryNo TH007 ,a.Quantity TH008, e.UomNo TH009, 0 TH010
                                , '' TH011, a.UnitPrice TH012, a.Amount TH013, ISNULL( f2.SoErpPrefix,'') TH014, ISNULL( f2.SoErpNo,'') TH015 
                                , ISNULL( f1.SoSequence,'') TH016 , a.LotNumber TH017, a.Remarks TH018, a.CustomerMtlItemNo TH019, a.ConfirmationCode TH020
                                , a.UpdateCode TH021, '' TH022, '' TH023, a.FreeSpareQty TH024, a.DiscountRate TH025
                                , a.CheckOutCode TH026, a.CheckOutPrefix TH027, a.CheckOutNo TH028, a.CheckOutSequence TH029, a.ProjectCode TH030
                                , a.[Type] TH031, ISNULL( g2.TsnErpPrefix,'') TH032, ISNULL( g2.TsnErpNo,'') TH033, ISNULL( g1.TsnSequence,'') TH034, a.OrigPreTaxAmt TH035 
                                , a.OrigTaxAmt TH036, a.PreTaxAmt TH037, a.TaxAmt TH038, a.PackageQty TH039, a.PackageFreeSpareQty TH040
                                , ISNULL(e2.UomNo,'') TH041, a.BondedCode TH042, a.SrQty TH043, a.SrPackageQty TH044, a.SrOrigPreTaxAmt TH045
                                , a.SrOrigTaxAmt TH046, a.SrPreTaxAmt TH047, a.SrTaxAmt TH048, a.CommissionRate TH049, a.CommissionAmount TH050
                                , a.OriginalCustomer TH051, a.SrFreeSpareQty TH052, a.SrPackageFreeSpareQty TH053, a.NotPayTemp TH054,'' TH055
                                , '' TH056, a.ProductSerialNumberQty TH057, a.ForecastCode TH058, a.ForecastSequence TH059, a.[Location] TH060
                                , a.AvailableQty TH061, e3.UomNo TH062, 0 TH063, 0 TH064, '' TH065
                                , '' TH066, '' TH067, a.MultiBatch TH068, a.FreeSpareRate TH069, a.FinalCustomerCode TH070
                                , a.ReferenceQty TH071, a.ReferencePackageQty TH072, a.TaxRate TH073, a.CRMSource TH074, a.CRMPrefix TH075
                                , a.CRMNo TH076, a.CRMSequence TH077, a.CRMContractCode TH078, a.CRMAllowDeduction TH079, a.CRMDeductionQty TH080
                                , a.CRMDeductionUnit TH081, a.MarkingBINRecord TH500, a.MarkingManagement TH501, a.BillingUnitInPackage TH502, a.DATECODE TH503
                                , a.DebitAccount TH082, a.CreditAccount TH083, a.TaxAmountAccount TH084, '' TH085
                                , '' TH086, '' TH087, '' TH088, a.BusinessItemNumber TH089, a.TaxCode TH090
                                , a.DiscountAmount TH091, '' TH092, '' TH093, a.K2NO TH094
                                FROM SCM.RoDetail a
                                INNER JOIN SCM.ReceiveOrder b on a.RoId = b.RoId
                                INNER JOIN PDM.MtlItem c on a.MtlItemId = c.MtlItemId
                                INNER JOIN SCM.Inventory d on a.InventoryId = d.InventoryId
                                INNER JOIN PDM.UnitOfMeasure e on a.UomId = e.UomId
                                LEFT JOIN PDM.UnitOfMeasure e2 on a.PackageUomId = e2.UomId
                                LEFT JOIN PDM.UnitOfMeasure e3 on a.AvailableUomId = e3.UomId
                                LEFT JOIN SCM.SoDetail f1 on a.SoDetailId = f1.SoDetailId
                                LEFT JOIN SCM.SaleOrder f2 on f1.SoId = f2.SoId
                                LEFT JOIN SCM.TsnDetail g1 on a.TsnDetailId = g1.TsnDetailId
                                LEFT JOIN SCM.TempShippingNote g2 on g1.TsnId = g2.TsnId
                                WHERE b.CompanyId = @CompanyId
                                AND b.RoId = @RoId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("RoId", RoId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("銷貨單單身資料找不到，請重新確認");
                        foreach (var item in result)
                        {
                            if (item.ConfirmationCode == "Y") throw new SystemException("銷貨單單身處於確認狀態不可以拋轉");
                            if (item.ConfirmationCode == "V") throw new SystemException("銷貨單單身處於作廢狀態不可以拋轉");
                            if (item.CheckOutCode == "Y") throw new SystemException("銷貨單單身已經結帳狀態不可以拋轉");
                        }
                        RoDetails = sqlConnection.Query<COPTH>(sql, dynamicParameters).ToList();
                        #endregion
                    }


                    if (ReceiveOrders.Select(x => x.TG023).FirstOrDefault() == "N") //銷貨單沒有確認
                    {
                        if(ReceiveOrders.Count() > 0)
                        {
                            using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                            {
                                DynamicParameters dynamicParameters = new DynamicParameters();

                                //單頭資料
                                List<COPTG> addReceiveOrders = ReceiveOrders.Where(x => x.TransferStatusMES == "N").ToList();
                                List<COPTG> updateReceiveOrders = ReceiveOrders.Where(x => x.TransferStatusMES == "R").ToList();
                                //單身清單
                                List<COPTH> addRoDetails = new List<COPTH>();
                                List<COPTH> updateRoDetails = new List<COPTH>();
                                addRoDetails = RoDetails.Where(x => x.TH020 == "N").ToList();

                                #region //ERP公司代碼取得
                                string ERPCompanyNo = "";
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 ML001  
                                         FROM CMSML";
                                var resultCompanyNo = sqlConnection.Query(sql, dynamicParameters);
                                foreach (var item in resultCompanyNo)
                                {
                                    ERPCompanyNo = item.ML001;
                                }
                                if (MESCompanyNo != ERPCompanyNo) throw new SystemException("ERP的公司代碼與MES的公司代碼不同，請嘗試重新登入!!");
                                #endregion

                                #region //判斷ERP使用者是否存在
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 MF005
                                        FROM ADMMF
                                        WHERE COMPANY = @CompanyNo
                                        AND MF001 = @userNo";
                                dynamicParameters.Add("CompanyNo", ERPCompanyNo);
                                dynamicParameters.Add("userNo", userNo);
                                var resultUserExist = sqlConnection.Query(sql, dynamicParameters);
                                if (resultUserExist.Count() <= 0) throw new SystemException("【ERP使用者】不存在!");
                                #endregion

                                #region //審核ERP權限-建立
                                var USR_GROUP = BaseHelper.CheckErpAuthority(userNo, sqlConnection, "COPI08", "CREATE");
                                #endregion

                                #region //判斷開立日期是否超過結帳日
                                string CloseDateBase = "";
                                string ReceiveDateBase = ReceiveDate;
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
                                        throw new SystemException("【銷貨單】ERP已經結帳,不可以在新增【" + CloseDateBase + "】之前的單據");
                                    }
                                }
                                else
                                {
                                    throw new SystemException("日期字符串无效，无法比较");
                                }
                                #endregion

                                if (ReceiveOrders.Select(x => x.TransferStatusMES).FirstOrDefault() == "N") //沒有拋轉過
                                {
                                    #region //撈取單據編號編碼格式設定
                                    RoErpPrefix = ReceiveOrders.Select(x => x.TG001).FirstOrDefault();
                                    string DocDate = ReceiveOrders.Select(x => x.TG042).FirstOrDefault();

                                    string WoType = "";
                                    string encode = ""; // 編碼格式
                                    int yearLength = 0; // 年碼數
                                    int lineLength = 0; // 流水碼數
                                    DateTime referenceTime = default(DateTime);
                                    string docDate = addReceiveOrders.Select(x => x.TG042).FirstOrDefault();
                                    referenceTime = DateTime.ParseExact(docDate, "yyyyMMdd", CultureInfo.InvariantCulture);
                                    #endregion

                                    #region //撈取ERP單據設定
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.MQ004, a.MQ005, a.MQ006, a.MQ044, a.MQ017
                                            FROM CMSMQ a
                                            WHERE MQ001 = @RoErpPrefix";
                                    dynamicParameters.Add("RoErpPrefix", RoErpPrefix);
                                    var resultReceiptNo = sqlConnection.Query(sql, dynamicParameters);
                                    foreach (var item in resultReceiptNo)
                                    {
                                        encode = item.MQ004; //編碼方式
                                        yearLength = Convert.ToInt32(item.MQ005); //年碼數
                                        lineLength = Convert.ToInt32(item.MQ006); //流水號碼數
                                        WoType = item.MQ017;
                                    }
                                    #endregion

                                    #region //單號自動取號
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT CONVERT(int, RIGHT(ISNULL(MAX(RTRIM(LTRIM(TG002))), '" + new string('0', lineLength) + @"'), " + lineLength + @")) + 1 CurrentNum
                                            FROM COPTG
                                            WHERE TG001 = @RoErpPrefix";
                                    dynamicParameters.Add("RoErpPrefix", RoErpPrefix);
                                    #endregion

                                    #region //編碼格式相關
                                    string dateFormat = "";
                                    switch (encode)
                                    {
                                        case "1": //日編
                                            if (yearLength < 0 || yearLength > 4) throw new SystemException("【ERP單據性質】年碼數不能小於等於0或大於4");
                                            if ((lineLength + yearLength + 4) > 11) throw new SystemException("【ERP單據性質】編碼格式碼數不能大於11碼");
                                            dateFormat = new string('y', yearLength) + "MMdd";
                                            sql += @" AND RTRIM(LTRIM(TG002)) LIKE @ReferenceTime  + '" + new string('_', lineLength) + @"'";
                                            //sql += @" AND TA002 LIKE '%' + @ReferenceTime + '%' + '" + new string('_', lineLength) + @"'";
                                            dynamicParameters.Add("ReferenceTime", referenceTime.ToString(dateFormat));

                                            string tedstNo = referenceTime.ToString(dateFormat);
                                            RoErpNo = referenceTime.ToString(dateFormat);
                                            break;
                                        case "2": //月編
                                            if (yearLength < 0 || yearLength > 4) throw new SystemException("【ERP單據性質】年碼數不能小於等於0或大於4");
                                            if ((lineLength + yearLength + 2) > 11) throw new SystemException("【ERP單據性質】編碼格式碼數不能大於11碼");
                                            dateFormat = new string('y', yearLength) + "MM";
                                            sql += @" AND RTRIM(LTRIM(TG002))  LIKE @ReferenceTime + '" + new string('_', lineLength) + @"'";
                                            dynamicParameters.Add("ReferenceTime", referenceTime.ToString(dateFormat));
                                            RoErpNo = referenceTime.ToString(dateFormat);
                                            break;
                                        case "3": //流水號
                                            if (yearLength == 0) throw new SystemException("【ERP單據性質】年碼數必須等於0");
                                            if (lineLength <= 0 || lineLength > 11) throw new SystemException("【ERP單據性質】流水編碼碼數必須大於0小於等於11");
                                            break;
                                        case "4": //手動編號
                                            break;
                                        default:
                                            throw new SystemException("編碼方式錯誤!");
                                    }

                                    int currentNum = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).CurrentNum;
                                    RoErpNo += string.Format("{0:" + new string('0', lineLength) + "}", currentNum);
                                    #endregion

                                    #region //判斷ERP銷貨單單號是否重複
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM COPTG
                                            WHERE TG001 = @TG001
                                            AND TG002 = @TG002";
                                    dynamicParameters.Add("TG001", RoErpPrefix);
                                    dynamicParameters.Add("TG002", RoErpNo);

                                    var result = sqlConnection.Query(sql, dynamicParameters);
                                    if (result.Count() > 0) throw new SystemException("【銷貨單單號】重複，請重新取號!");
                                    string CFIELD01 = RoErpPrefix + "-" + RoErpNo;
                                    #endregion

                                    #region //新增
                                    #region //單頭
                                    if (addReceiveOrders.Count > 0)
                                    {
                                        #region //賦值
                                        addReceiveOrders
                                            .ToList()
                                            .ForEach(x =>
                                            {
                                                x.COMPANY = MESCompanyNo;
                                                x.USR_GROUP = USR_GROUP;
                                                x.FLAG = "1";
                                                x.CREATOR = userNo;
                                                x.CREATE_DATE = dateNow;
                                                x.CREATE_TIME = timeNow;
                                                x.CREATE_AP = BaseHelper.ClientComputer();
                                                x.CREATE_PRID = "BM";
                                                x.MODIFIER = "";
                                                x.MODI_DATE = "";
                                                x.MODI_TIME = "";
                                                x.MODI_AP = "";
                                                x.MODI_PRID = "";
                                                x.TG002 = RoErpNo;
                                                x.TG021 = x.TG021 != null ? x.TG021 : "";
                                                x.CFIELD01 = CFIELD01;
                                                x.UDF01 = "";
                                                x.UDF02 = "";
                                                x.UDF03 = "";
                                                x.UDF04 = "";
                                                x.UDF05 = "";
                                                x.UDF06 = 0.0;
                                                x.UDF07 = 0.0;
                                                x.UDF08 = 0.0;
                                                x.UDF09 = 0.0;
                                                x.UDF10 = 0.0;
                                            });
                                        #endregion

                                        #region //新增銷貨單頭 COPTG
                                        sql = @"INSERT INTO COPTG (COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE, FLAG, CREATE_TIME, 
                                                CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID, TG001, TG002, TG003, TG004, TG005, TG006, 
                                                TG007, TG008, TG009, TG010, TG011, TG012, TG013, TG014, TG015, TG016, TG017, TG018, TG019, TG020, TG021, 
                                                TG022, TG023, TG024, TG025, TG026, TG027, TG028, TG029, TG030, TG031, TG032, TG033, TG034, TG035, TG036, 
                                                TG037, TG038, TG039, TG040, TG041, TG042, TG043, TG044, TG045, TG046, TG047, TG048, TG049, TG050, TG051, 
                                                TG052, TG053, TG054, TG055, TG056, TG057, TG058, TG059, TG060, TG061, TG062, TG063, TG064, TG065, TG066, 
                                                TG067, TG068, TG069, TG070, TG071, TG072, TG073, TG074, TG075, TG076, TG077, TG078, TG079, TG080, TG081, 
                                                TG082, TG083, TG084, TG085, TG086, TG087, TG088, TG089, TG090, TG091, TG092, TG093, TG094, TG095, TG096, 
                                                TG097, TG098, TG099, TG100, TG101, TG102, TG103, TG104, TG105, TG106, TG107, TG108, TG109, TG110, TG111, 
                                                TG112, TG113, TG114, TG115, TG116, UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, 
                                                UDF10, TG117, TG118, TG119, TG120, TG121, TG124, TG125, TG126, TG127, TG128, TG122, TG123, 
                                                TG129, TG130, TG131, TG132, TG133)
                                                VALUES (@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE, @FLAG, @CREATE_TIME,
                                                @CREATE_AP, @CREATE_PRID, @MODI_TIME, @MODI_AP, @MODI_PRID, @TG001, @TG002, @TG003, @TG004, @TG005, @TG006,
                                                @TG007, @TG008, @TG009, @TG010, @TG011, @TG012, @TG013, @TG014, @TG015, @TG016, @TG017, @TG018, @TG019, @TG020, @TG021,
                                                @TG022, @TG023, @TG024, @TG025, @TG026, @TG027, @TG028, @TG029, @TG030, @TG031, @TG032, @TG033, @TG034, @TG035, @TG036,
                                                @TG037, @TG038, @TG039, @TG040, @TG041, @TG042, @TG043, @TG044, @TG045, @TG046, @TG047, @TG048, @TG049, @TG050, @TG051,
                                                @TG052, @TG053, @TG054, @TG055, @TG056, @TG057, @TG058, @TG059, @TG060, @TG061, @TG062, @TG063, @TG064, @TG065, @TG066,
                                                @TG067, @TG068, @TG069, @TG070, @TG071, @TG072, @TG073, @TG074, @TG075, @TG076, @TG077, @TG078, @TG079, @TG080, @TG081,
                                                @TG082, @TG083, @TG084, @TG085, @TG086, @TG087, @TG088, @TG089, @TG090, @TG091, @TG092, @TG093, @TG094, @TG095, @TG096,
                                                @TG097, @TG098, @TG099, @TG100, @TG101, @TG102, @TG103, @TG104, @TG105, @TG106, @TG107, @TG108, @TG109, @TG110, @TG111,
                                                @TG112, @TG113, @TG114, @TG115, @TG116, @UDF01, @UDF02, @UDF03, @UDF04, @UDF05, @UDF06, @UDF07, @UDF08, @UDF09,
                                                @UDF10, @TG117, @TG118, @TG119, @TG120, @TG121, @TG124, @TG125, @TG126, @TG127, @TG128, @TG122, @TG123,
                                                @TG129, @TG130, @TG131, @TG132, @TG133)";
                                        rowsAffected += sqlConnection.Execute(sql, addReceiveOrders);
                                        #endregion
                                    }
                                    #endregion

                                    #region //單身                                    
                                    if (addRoDetails.Count > 0)
                                    {
                                        #region //銷貨單單身資料驗證                                    
                                        string AllOverMistake = "";
                                        foreach (var item in addRoDetails)
                                        {
                                            string OverMtlItemNo = "";
                                            string OverMistake = "";
                                            string TH004 = item.TH004; //品號
                                            string TH007 = item.TH007; //庫別
                                            string TH014 = item.TH014; //訂單單別 單號 序號
                                            string TH015 = item.TH015;
                                            string TH016 = item.TH016;
                                            string SoFull = item.TH014 + '-'+ item.TH015 + '-' + item.TH016;
                                            string TH032 = item.TH032; //暫出單單別 單號 序號
                                            string TH033 = item.TH033;
                                            string TH034 = item.TH034;
                                            string TsnFull = item.TH032 + '-'+ item.TH033 + '-' + item.TH034;
                                            string TH017 = item.TH017; //批號
                                            Double TH008 = item.TH008; //數量
                                            Double TH024 = item.TH024; //贈/備品數量
                                            string TD049 = "";
                                            Double TD008 = 0;
                                            Double TD024 = 0;
                                            Double TD050 = 0;
                                            Double TG009 = 0;
                                            Double TG044 = 0;
                                            Double TG052 = 0;

                                            double CanUseQty = 0;
                                            double CanUseFreebieQty = 0;
                                            double CanUseSpareQty = 0;
                                            double CanUseFreeSpareQty = 0;
                                            double CanUseAvailableQty = 0;

                                            #region //品號批號管理檢查
                                            string MB022 = "";
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"SELECT TOP 1 MB022
                                                    FROM INVMB
                                                    WHERE MB001 = @TH004";
                                            dynamicParameters.Add("TH004", TH004);
                                            result = sqlConnection.Query(sql, dynamicParameters);
                                            if (result.Count() <= 0) throw new SystemException("品號找不到,請重新確認");
                                            foreach (var item1 in result)
                                            {
                                                MB022 = item1.MB022;
                                            }
                                            if(MB022 != "N")
                                            {
                                                #region //判斷品號批號庫存是否足夠
                                                dynamicParameters = new DynamicParameters();
                                                sql = @"SELECT a.ME001, a.ME002,a.ME003,a.ME009,b.num,b.MF007
                                                        FROM INVME a
                                                        OUTER APPLY(
                                                              SELECT SUM(x.MF010) num,x.MF007 FROM INVMF x
                                                              WHERE x.MF001 = a.ME001 AND x.MF002 = a.ME002
                                                              GROUP BY x.MF007
                                                          ) b
                                                        WHERE 1 = 1
                                                        AND b.MF007 = @InventoryNo
                                                        AND a.ME001 = @MtlItemNo
                                                        AND a.ME002 = @LotNumber
                                                        ORDER BY a.ME007 DESC   
                                                        ";
                                                dynamicParameters.Add("InventoryNo", TH007);
                                                dynamicParameters.Add("MtlItemNo", TH004);
                                                dynamicParameters.Add("LotNumber", TH017);
                                                result = sqlConnection.Query(sql, dynamicParameters);
                                                foreach(var item1 in result)
                                                {
                                                    if(Convert.ToDouble(item1.num) < (TH008+TH024)) throw new SystemException("該品號庫存批號不足,請重新確認");
                                                }
                                                #endregion
                                            }
                                            #endregion

                                            #region //單據來源判斷其數量卡控走向
                                            switch (SourceType)
                                            {
                                                case "So":
                                                    #region //撈取訂單數量資料
                                                    dynamicParameters = new DynamicParameters();
                                                    sql = @"SELECT TOP 1 a.TD049,a.TD016,
                                                            ISNULL(a.TD008,0) TD008,
                                                            ISNULL(a.TD024,0) TD024,
                                                            ISNULL(a.TD050,0) TD050,
                                                            (ISNULL(a.TD008,0) -  ISNULL(a.TD009,0) - ISNULL(x.UnconfirmedQty,0)) CanUseQty,
                                                            (ISNULL(a.TD024,0) -  ISNULL(a.TD025,0) - ISNULL(x.UnconfirmedFreeSpareQty,0)) CanUseFreebieQty,
                                                            (ISNULL(a.TD050,0) -  ISNULL(a.TD051,0) - ISNULL(x.UnconfirmedFreeSpareQty,0)) CanUseSpareQty,
                                                            (ISNULL(a.TD076,0) -  ISNULL(a.TD078,0) - ISNULL(x.UnconfirmedAvailableQty,0)) CanUseAvailableQty,
                                                            ISNULL(x.UnconfirmedQty,0) UnconfirmedQty,
                                                            ISNULL(x.UnconfirmedFreeSpareQty,0) UnconfirmedFreeSpareQty, 
                                                            ISNULL(x.UnconfirmedAvailableQty,0) UnconfirmedAvailableQty
                                                            FROM COPTD a
                                                            OUTER APPLY(
                                                                SELECT SUM(ISNULL(x1.TH008,0)) UnconfirmedQty 
                                                                        ,SUM(ISNULL(x1.TH024,0)) UnconfirmedFreeSpareQty 
                                                                        ,SUM(ISNULL(x1.TH061,0)) UnconfirmedAvailableQty
                                                                FROM COPTH x1
                                                                WHERE x1.TH014 = a.TD001
                                                                AND x1.TH015 = a.TD002
                                                                AND x1.TH016 = a.TD003
                                                                AND x1.TH020 ='N'
                                                            ) x
                                                            WHERE a.TD001 = @TH014
                                                            AND a.TD002 = @TH015
                                                            AND a.TD003 = @TH016
                                                            AND a.TD021 = 'Y'";
                                                    dynamicParameters.Add("TH014", TH014);
                                                    dynamicParameters.Add("TH015", TH015);
                                                    dynamicParameters.Add("TH016", TH016);
                                                    result = sqlConnection.Query(sql, dynamicParameters);
                                                    if (result.Count() <= 0) throw new SystemException("訂單找不到,請重新確認");
                                                    foreach (var item1 in result)
                                                    {
                                                        if(item1.TD016 == "Y") throw new SystemException("【訂單:" + SoFull + "】已經結案!!!");
                                                        TD049 = item1.TD049;
                                                        TD008 = Convert.ToDouble(item1.TD008);
                                                        TD024 = Convert.ToDouble(item1.TD024);
                                                        TD050 = Convert.ToDouble(item1.TD050);
                                                        CanUseQty = Convert.ToDouble(item1.CanUseQty);
                                                        CanUseFreebieQty = Convert.ToDouble(item1.CanUseFreebieQty);
                                                        CanUseSpareQty = Convert.ToDouble(item1.CanUseSpareQty);
                                                        CanUseAvailableQty = Convert.ToDouble(item1.CanUseAvailableQty);
                                                    }
                                                    #endregion

                                                    #region //判斷是否有超過可銷貨數
                                                    OverMtlItemNo = "銷貨單單身品號:" + TH004;
                                                    if ((CanUseQty - TH008) < 0)
                                                    {
                                                        OverMistake += "<br>已經超過可銷貨數量<br>訂單數量:" + TD008 + ",目前可銷貨最大數:" + CanUseQty + "<br>欲新增銷貨數量:" + TH008 + ",請重新確認";
                                                    }
                                                    if (TD049 == "1")
                                                    {
                                                        if ((CanUseFreebieQty - TH024) < 0)
                                                        {
                                                            OverMistake += "<br>已經超過可銷貨贈品數量,訂單贈品數量:" + TD024 + ",目前可銷貨贈品最大數量:" + CanUseFreebieQty + "<br>欲新增銷貨贈品數量:" + TH024 + ",請重新確認<br>";
                                                        }
                                                    }
                                                    else if (TD049 == "2")
                                                    {
                                                        if ((CanUseSpareQty - TH024) < 0)
                                                        {
                                                            OverMistake += "<br>已經超過可銷貨備品數量,訂單備品數量:" + TD050 + ",目前可銷貨貨品最大數量:" + CanUseSpareQty + "<br>欲新增銷貨品數量:" + TH024 + ",請重新確認<br>";
                                                        }
                                                    }
                                                    if (OverMistake.Length > 0)
                                                    {
                                                        AllOverMistake += OverMtlItemNo + OverMistake;
                                                    }
                                                    #endregion
                                                    break;
                                                case "To":
                                                    #region //撈取訂單數量資料
                                                    dynamicParameters = new DynamicParameters();
                                                    sql = @"SELECT a.TG024,ISNULL(a.TG009,0) TG009,ISNULL(a.TG044,0) TG044,ISNULL(a.TG052,0) TG052,
                                                            (ISNULL(a.TG009,0) - ISNULL(a.TG020,0) - ISNULL(a.TG021,0) - ISNULL(b.UnconfirmedQty,0)) CanUseQty
                                                            ,(ISNULL(a.TG044,0) - ISNULL(a.TG046,0) - ISNULL(a.TG048,0) - ISNULL(b.UnconfirmedFreebieQty,0)) CanUseFreeSpareQty 
                                                            ,(ISNULL(a.TG052,0) - ISNULL(a.TG054,0) - ISNULL(a.TG055,0) - ISNULL(b.UnconfirmedAvailableQty,0)) CanUseAvailableQty 
                                                            FROM INVTG a
                                                            OUTER APPLY(
                                                                SELECT SUM(ISNULL(x1.TH008,0)) UnconfirmedQty 
                                                                        ,SUM(ISNULL(x1.TH024,0)) UnconfirmedFreebieQty 
                                                                        ,SUM(ISNULL(x1.TH061,0)) UnconfirmedAvailableQty
                                                                FROM COPTH x1
                                                                WHERE x1.TH032 = a.TG001
                                                                AND x1.TH033 = a.TG002
                                                                AND x1.TH034 = a.TG003
                                                                AND x1.TH020 ='N'
                                                            ) b
                                                            WHERE a.TG001 = @TG001
                                                            AND a.TG002 = @TG002
                                                            AND a.TG003 = @TG003
                                                            AND a.TG022 ='Y'
                                                            ";
                                                    dynamicParameters.Add("TG001", TH032);
                                                    dynamicParameters.Add("TG002", TH033);
                                                    dynamicParameters.Add("TG003", TH034);
                                                    result = sqlConnection.Query(sql, dynamicParameters);
                                                    if (result.Count() <= 0) throw new SystemException("暫出單找不到,請重新確認");
                                                    foreach (var item1 in result)
                                                    {
                                                        if(item1.TG024 == "Y") throw new SystemException("【訂單:" + TsnFull + "】已經結案!!!");
                                                        TG009 = Convert.ToDouble(item1.TG009);
                                                        TG044 = Convert.ToDouble(item1.TG044);
                                                        TG052 = Convert.ToDouble(item1.TG052);
                                                        CanUseQty = Convert.ToDouble(item1.CanUseQty);
                                                        CanUseFreeSpareQty = Convert.ToDouble(item1.CanUseFreeSpareQty);
                                                        CanUseAvailableQty = Convert.ToDouble(item1.CanUseAvailableQty);
                                                    }
                                                    #endregion

                                                    #region //判斷是否有超過可銷貨數
                                                    OverMtlItemNo = "銷貨單單身品號:" + TH004;
                                                    if ((CanUseQty - TH008) < 0)
                                                    {
                                                        OverMistake += "<br>已經超過可銷貨數量<br>暫出單數量:" + TG009 + ",目前可銷貨最大數:" + CanUseQty + "<br>欲新增銷貨數量:" + TH008 + ",請重新確認";
                                                    }
                                                    if ((CanUseFreeSpareQty - TH024) < 0)
                                                    {
                                                        OverMistake += "<br>已經超過可銷貨贈備品數量,暫出單贈備品數量:" + TG044 + ",目前可銷貨贈備品最大數量:" + CanUseFreeSpareQty + "<br>欲新增銷貨贈備品數量:" + TH024 + ",請重新確認<br>";
                                                    }
                                                    
                                                    if (OverMistake.Length > 0)
                                                    {
                                                        AllOverMistake += OverMtlItemNo + OverMistake;
                                                    }
                                                    #endregion


                                                    break;
                                            }
                                            #endregion

                                        }

                                        if (AllOverMistake.Length > 0)
                                        {
                                            throw new SystemException(AllOverMistake);
                                        }
                                        #endregion

                                        #region //銷貨單身賦值
                                        addRoDetails
                                            .ToList()
                                            .ForEach(x =>
                                            {
                                                x.COMPANY = MESCompanyNo;
                                                x.USR_GROUP = USR_GROUP;
                                                x.FLAG = "1";
                                                x.CREATOR = userNo;
                                                x.CREATE_DATE = dateNow;
                                                x.CREATE_TIME = timeNow;
                                                x.CREATE_AP = BaseHelper.ClientComputer();
                                                x.CREATE_PRID = "BM";
                                                x.MODIFIER = "";
                                                x.MODI_DATE = "";
                                                x.MODI_TIME = "";
                                                x.MODI_AP = "";
                                                x.MODI_PRID = "";
                                                x.TH002 = RoErpNo;
                                                x.UDF01 = "";
                                                x.UDF02 = "";
                                                x.UDF03 = "";
                                                x.UDF04 = "";
                                                x.UDF05 = "";
                                                x.UDF06 = 0.0;
                                                x.UDF07 = 0.0;
                                                x.UDF08 = 0.0;
                                                x.UDF09 = 0.0;
                                                x.UDF10 = 0.0;
                                            });
                                        #endregion

                                        #region //新增銷貨單身 COPTH                                    
                                        sql = @"INSERT INTO COPTH(COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE, 
                                                FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID, 
                                                TH001, TH002, TH003, TH004, TH005, TH006, TH007, TH008, TH009, TH010,
                                                TH011, TH012, TH013, TH014, TH015, TH016, TH017, TH018, TH019, TH020,
                                                TH021, TH022, TH023, TH024, TH025, TH026, TH027, TH028, TH029, TH030,
                                                TH031, TH032, TH033, TH034, TH035, TH036, TH037, TH038, TH039, TH040,
                                                TH041, TH042, TH043, TH044, TH045, TH046, TH047, TH048, TH049, TH050,
                                                TH051, TH052, TH053, TH054, TH055, TH056, TH057, TH058, TH059, TH060,
                                                TH061, TH062, TH063, TH064, TH065, TH066, TH067, TH068, TH069, TH070,
                                                TH071, TH072, TH073, TH074, TH075, TH076, TH077, TH078, TH079, TH080,
                                                TH081, TH500, TH501, TH502, TH503, TH082, TH083, TH084, TH085, TH086,
                                                TH087, TH088, TH089, UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07,
                                                UDF08, UDF09, UDF10, TH090, TH091, TH092, TH093, TH094)
                                                VALUES(@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE, 
                                                @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID,@MODI_TIME, @MODI_AP, @MODI_PRID, 
                                                @TH001, @TH002, @TH003, @TH004, @TH005, @TH006, @TH007, @TH008, @TH009, @TH010,
                                                @TH011, @TH012, @TH013, @TH014, @TH015, @TH016, @TH017, @TH018, @TH019, @TH020,
                                                @TH021, @TH022, @TH023, @TH024, @TH025, @TH026, @TH027, @TH028, @TH029, @TH030,
                                                @TH031, @TH032, @TH033, @TH034, @TH035, @TH036, @TH037, @TH038, @TH039, @TH040,
                                                @TH041, @TH042, @TH043, @TH044, @TH045, @TH046, @TH047, @TH048, @TH049, @TH050,
                                                @TH051, @TH052, @TH053, @TH054, @TH055, @TH056, @TH057, @TH058, @TH059, @TH060,
                                                @TH061, @TH062, @TH063, @TH064, @TH065, @TH066, @TH067, @TH068, @TH069, @TH070,
                                                @TH071, @TH072, @TH073, @TH074, @TH075, @TH076, @TH077, @TH078, @TH079, @TH080,
                                                @TH081, @TH500, @TH501, @TH502, @TH503, @TH082, @TH083, @TH084, @TH085, @TH086,
                                                @TH087, @TH088, @TH089, @UDF01, @UDF02, @UDF03, @UDF04, @UDF05, @UDF06, @UDF07,
                                                @UDF08, @UDF09, @UDF10, @TH090, @TH091, @TH092, @TH093, @TH094)";
                                        rowsAffected += sqlConnection.Execute(sql, addRoDetails);
                                        #endregion
                                    }
                                    #endregion
                                    #endregion

                                }
                                else if (ReceiveOrders.Select(x => x.TransferStatusMES).FirstOrDefault() == "R") //有拋轉過
                                {
                                    #region //判斷銷貨單是否存在+可不可以拋轉
                                    string erpConfirmStatus = "";
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TG023
                                            FROM COPTG
                                            WHERE TG001 =@RoErpPrefix
                                            AND TG002 = @RoErpNo";
                                    dynamicParameters.Add("RoErpPrefix", RoErpPrefix);
                                    dynamicParameters.Add("RoErpNo", RoErpNo);
                                    var result = sqlConnection.Query(sql, dynamicParameters);
                                    if (result.Count() <= 0) throw new SystemException("銷貨單不存在,請重新確認!!!");

                                    foreach (var item in result)
                                    {
                                        erpConfirmStatus = item.TG023;
                                    }
                                    if (erpConfirmStatus == "Y") throw new SystemException("ERP銷貨單處於確認狀態不可以拋轉!!!");
                                    if (erpConfirmStatus == "V") throw new SystemException("ERP銷貨單處於作廢狀態不可以拋轉!!!");
                                    #endregion

                                    #region //單頭修改
                                    if (updateReceiveOrders.Count > 0)
                                    {
                                        updateReceiveOrders
                                            .ToList()
                                            .ForEach(x =>
                                            {
                                                x.MODIFIER = userNo;
                                                x.MODI_DATE = dateNow;
                                                x.MODI_TIME = timeNow;
                                                x.MODI_AP = BaseHelper.ClientComputer();
                                                x.MODI_PRID = "BM";
                                            });

                                        sql = @"UPDATE COPTG SET
                                                TG003 = @TG003,  
                                                TG004 = @TG004,  
                                                TG005 = @TG005,  
                                                TG006 = @TG006,  
                                                TG007 = @TG007,  
                                                TG008 = @TG008,  
                                                TG009 = @TG009,  
                                                TG010 = @TG010,  
                                                TG011 = @TG011,  
                                                TG012 = @TG012,  
                                                TG013 = @TG013,  
                                                TG014 = @TG014,  
                                                TG015 = @TG015,  
                                                TG016 = @TG016,  
                                                TG017 = @TG017,  
                                                TG018 = @TG018,  
                                                TG019 = @TG019,  
                                                TG020 = @TG020,  
                                                TG021 = @TG021,  
                                                TG022 = @TG022,  
                                                TG024 = @TG024,  
                                                TG025 = @TG025,  
                                                TG026 = @TG026,  
                                                TG027 = @TG027,  
                                                TG028 = @TG028,  
                                                TG029 = @TG029,  
                                                TG030 = @TG030,  
                                                TG031 = @TG031,  
                                                TG032 = @TG032,  
                                                TG033 = @TG033,  
                                                TG034 = @TG034,  
                                                TG035 = @TG035,  
                                                TG036 = @TG036,  
                                                TG037 = @TG037,  
                                                TG038 = @TG038,  
                                                TG039 = @TG039,  
                                                TG040 = @TG040,  
                                                TG041 = @TG041,  
                                                TG042 = @TG042,  
                                                TG044 = @TG044,  
                                                TG045 = @TG045,  
                                                TG046 = @TG046,  
                                                TG047 = @TG047,  
                                                TG048 = @TG048,  
                                                TG049 = @TG049,  
                                                TG050 = @TG050,  
                                                TG051 = @TG051,  
                                                TG052 = @TG052,  
                                                TG053 = @TG053,  
                                                TG054 = @TG054,  
                                                TG055 = @TG055,  
                                                TG056 = @TG056,  
                                                TG057 = @TG057,  
                                                TG058 = @TG058,  
                                                TG059 = @TG059,  
                                                TG060 = @TG060,  
                                                TG061 = @TG061,  
                                                TG062 = @TG062,  
                                                TG063 = @TG063,  
                                                TG064 = @TG064,  
                                                TG065 = @TG065,  
                                                TG066 = @TG066,  
                                                TG067 = @TG067,  
                                                TG068 = @TG068,  
                                                TG069 = @TG069,  
                                                TG070 = @TG070,  
                                                TG071 = @TG071,  
                                                TG072 = @TG072,  
                                                TG073 = @TG073,  
                                                TG074 = @TG074,  
                                                TG075 = @TG075,  
                                                TG076 = @TG076,  
                                                TG077 = @TG077,  
                                                TG078 = @TG078,  
                                                TG079 = @TG079,  
                                                TG080 = @TG080,  
                                                TG081 = @TG081,  
                                                TG082 = @TG082,  
                                                TG083 = @TG083,  
                                                TG084 = @TG084,  
                                                TG085 = @TG085,  
                                                TG086 = @TG086,  
                                                TG087 = @TG087,  
                                                TG088 = @TG088,  
                                                TG089 = @TG089,  
                                                TG090 = @TG090,  
                                                TG091 = @TG091,  
                                                TG092 = @TG092,  
                                                TG093 = @TG093,  
                                                TG094 = @TG094,  
                                                TG095 = @TG095,  
                                                TG096 = @TG096,  
                                                TG097 = @TG097,  
                                                TG098 = @TG098,  
                                                TG099 = @TG099,  
                                                TG100 = @TG100,  
                                                TG101 = @TG101,  
                                                TG102 = @TG102,  
                                                TG103 = @TG103,  
                                                TG104 = @TG104,  
                                                TG105 = @TG105,  
                                                TG106 = @TG106,  
                                                TG107 = @TG107,  
                                                TG108 = @TG108,  
                                                TG109 = @TG109,  
                                                TG110 = @TG110,  
                                                TG111 = @TG111,  
                                                TG112 = @TG112,  
                                                TG113 = @TG113,  
                                                TG114 = @TG114,  
                                                TG115 = @TG115,  
                                                TG116 = @TG116,  
                                                TG117 = @TG117,  
                                                TG118 = @TG118,  
                                                TG119 = @TG119,  
                                                TG120 = @TG120,  
                                                TG121 = @TG121,  
                                                TG124 = @TG124,  
                                                TG125 = @TG125,  
                                                TG126 = @TG126,  
                                                TG127 = @TG127,  
                                                TG128 = @TG128,  
                                                TG122 = @TG122,  
                                                TG123 = @TG123,  
                                                TG129 = @TG129,  
                                                TG130 = @TG130,  
                                                TG131 = @TG131,  
                                                TG132 = @TG132,  
                                                TG133 = @TG133  
                                                WHERE TG001 = @TG001
                                                AND TG002 = @TG002";
                                        rowsAffected += sqlConnection.Execute(sql, updateReceiveOrders);
                                    }
                                    #endregion

                                    #region //單身依MES最新版重新拋轉
                                    #region //刪除單身
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE COPTH
                                            WHERE TH001 = @RoErpPrefix
                                            AND TH002 = @RoErpNo";
                                    dynamicParameters.Add("RoErpPrefix", RoErpPrefix);
                                    dynamicParameters.Add("RoErpNo", RoErpNo);

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //新增單身                                    
                                    if (addRoDetails.Count > 0)
                                    {
                                        #region //銷貨單單身資料驗證                                    
                                        string AllOverMistake = "";
                                        foreach (var item in addRoDetails)
                                        {
                                            string OverMtlItemNo = "";
                                            string OverMistake = "";
                                            string TH004 = item.TH004; //品號
                                            string TH007 = item.TH007; //庫別
                                            string TH014 = item.TH014; //訂單單別 單號 序號
                                            string TH015 = item.TH015;
                                            string TH016 = item.TH016;
                                            string SoFull = item.TH014 + '-' + item.TH015 + '-' + item.TH016;
                                            string TH032 = item.TH032; //暫出單單別 單號 序號
                                            string TH033 = item.TH033;
                                            string TH034 = item.TH034;
                                            string TsnFull = item.TH032 + '-' + item.TH033 + '-' + item.TH034;
                                            string TH017 = item.TH017; //批號
                                            Double TH008 = item.TH008; //數量
                                            Double TH024 = item.TH024; //贈/備品數量
                                            string TD049 = "";
                                            Double TD008 = 0;
                                            Double TD024 = 0;
                                            Double TD050 = 0;
                                            Double TG009 = 0;
                                            Double TG044 = 0;
                                            Double TG052 = 0;

                                            double CanUseQty = 0;
                                            double CanUseFreebieQty = 0;
                                            double CanUseSpareQty = 0;
                                            double CanUseFreeSpareQty = 0;
                                            double CanUseAvailableQty = 0;

                                            #region //品號批號管理檢查
                                            string MB022 = "";
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"SELECT TOP 1 MB022
                                                    FROM INVMB
                                                    WHERE MB001 = @TH004";
                                            dynamicParameters.Add("TH004", TH004);
                                            result = sqlConnection.Query(sql, dynamicParameters);
                                            if (result.Count() <= 0) throw new SystemException("品號找不到,請重新確認");
                                            foreach (var item1 in result)
                                            {
                                                MB022 = item1.MB022;
                                            }
                                            if (MB022 != "N")
                                            {
                                                #region //判斷品號批號庫存是否足夠
                                                dynamicParameters = new DynamicParameters();
                                                sql = @"SELECT a.ME001, a.ME002,a.ME003,a.ME009,b.num,b.MF007
                                                        FROM INVME a
                                                        OUTER APPLY(
                                                              SELECT SUM(x.MF010) num,x.MF007 FROM INVMF x
                                                              WHERE x.MF001 = a.ME001 AND x.MF002 = a.ME002
                                                              GROUP BY x.MF007
                                                          ) b
                                                        WHERE 1 = 1
                                                        AND b.MF007 = @InventoryNo
                                                        AND a.ME001 = @MtlItemNo
                                                        AND a.ME002 = @LotNumber
                                                        ORDER BY a.ME007 DESC   
                                                        ";
                                                dynamicParameters.Add("InventoryNo", TH007);
                                                dynamicParameters.Add("MtlItemNo", TH004);
                                                dynamicParameters.Add("LotNumber", TH017);
                                                result = sqlConnection.Query(sql, dynamicParameters);
                                                foreach (var item1 in result)
                                                {
                                                    if (Convert.ToDouble(item1.num) < (TH008 + TH024)) throw new SystemException("該品號庫存批號不足,請重新確認");
                                                }
                                                #endregion
                                            }
                                            #endregion

                                            #region //單據來源判斷其數量卡控走向
                                            switch (SourceType)
                                            {
                                                case "So":
                                                    #region //撈取訂單數量資料
                                                    dynamicParameters = new DynamicParameters();
                                                    sql = @"SELECT TOP 1 a.TD049,a.TD016,
                                                            ISNULL(a.TD008,0) TD008,
                                                            ISNULL(a.TD024,0) TD024,
                                                            ISNULL(a.TD050,0) TD050,
                                                            (ISNULL(a.TD008,0) -  ISNULL(a.TD009,0) - ISNULL(x.UnconfirmedQty,0)) CanUseQty,
                                                            (ISNULL(a.TD024,0) -  ISNULL(a.TD025,0) - ISNULL(x.UnconfirmedFreeSpareQty,0)) CanUseFreebieQty,
                                                            (ISNULL(a.TD050,0) -  ISNULL(a.TD051,0) - ISNULL(x.UnconfirmedFreeSpareQty,0)) CanUseSpareQty,
                                                            (ISNULL(a.TD076,0) -  ISNULL(a.TD078,0) - ISNULL(x.UnconfirmedAvailableQty,0)) CanUseAvailableQty,
                                                            ISNULL(x.UnconfirmedQty,0) UnconfirmedQty,
                                                            ISNULL(x.UnconfirmedFreeSpareQty,0) UnconfirmedFreeSpareQty, 
                                                            ISNULL(x.UnconfirmedAvailableQty,0) UnconfirmedAvailableQty
                                                            FROM COPTD a
                                                            OUTER APPLY(
                                                                SELECT SUM(ISNULL(x1.TH008,0)) UnconfirmedQty 
                                                                        ,SUM(ISNULL(x1.TH024,0)) UnconfirmedFreeSpareQty 
                                                                        ,SUM(ISNULL(x1.TH061,0)) UnconfirmedAvailableQty
                                                                FROM COPTH x1
                                                                WHERE x1.TH014 = a.TD001
                                                                AND x1.TH015 = a.TD002
                                                                AND x1.TH016 = a.TD003
                                                                AND x1.TH020 ='N'
                                                            ) x
                                                            WHERE a.TD001 = @TH014
                                                            AND a.TD002 = @TH015
                                                            AND a.TD003 = @TH016
                                                            AND a.TD021 = 'Y'";
                                                    dynamicParameters.Add("TH014", TH014);
                                                    dynamicParameters.Add("TH015", TH015);
                                                    dynamicParameters.Add("TH016", TH016);
                                                    result = sqlConnection.Query(sql, dynamicParameters);
                                                    if (result.Count() <= 0) throw new SystemException("訂單找不到,請重新確認");
                                                    foreach (var item1 in result)
                                                    {
                                                        if (item1.TD016 == "Y") throw new SystemException("【訂單:" + SoFull + "】已經結案!!!");
                                                        TD049 = item1.TD049;
                                                        TD008 = Convert.ToDouble(item1.TD008);
                                                        TD024 = Convert.ToDouble(item1.TD024);
                                                        TD050 = Convert.ToDouble(item1.TD050);
                                                        CanUseQty = Convert.ToDouble(item1.CanUseQty);
                                                        CanUseFreebieQty = Convert.ToDouble(item1.CanUseFreebieQty);
                                                        CanUseSpareQty = Convert.ToDouble(item1.CanUseSpareQty);
                                                        CanUseAvailableQty = Convert.ToDouble(item1.CanUseAvailableQty);
                                                    }
                                                    #endregion

                                                    #region //判斷是否有超過可銷貨數
                                                    OverMtlItemNo = "銷貨單單身品號:" + TH004;
                                                    if ((CanUseQty - TH008) < 0)
                                                    {
                                                        OverMistake += "<br>已經超過可銷貨數量<br>訂單數量:" + TD008 + ",目前可銷貨最大數:" + CanUseQty + "<br>欲新增銷貨數量:" + TH008 + ",請重新確認";
                                                    }
                                                    if (TD049 == "1")
                                                    {
                                                        if ((CanUseFreebieQty - TH024) < 0)
                                                        {
                                                            OverMistake += "<br>已經超過可銷貨贈品數量,訂單贈品數量:" + TD024 + ",目前可銷貨贈品最大數量:" + CanUseFreebieQty + "<br>欲新增銷貨贈品數量:" + TH024 + ",請重新確認<br>";
                                                        }
                                                    }
                                                    else if (TD049 == "2")
                                                    {
                                                        if ((CanUseSpareQty - TH024) < 0)
                                                        {
                                                            OverMistake += "<br>已經超過可銷貨備品數量,訂單備品數量:" + TD050 + ",目前可銷貨貨品最大數量:" + CanUseSpareQty + "<br>欲新增銷貨品數量:" + TH024 + ",請重新確認<br>";
                                                        }
                                                    }
                                                    if (OverMistake.Length > 0)
                                                    {
                                                        AllOverMistake += OverMtlItemNo + OverMistake;
                                                    }
                                                    #endregion
                                                    break;
                                                case "To":
                                                    #region //撈取訂單數量資料
                                                    dynamicParameters = new DynamicParameters();
                                                    sql = @"SELECT a.TG024,ISNULL(a.TG009,0) TG009,ISNULL(a.TG044,0) TG044,ISNULL(a.TG052,0) TG052,
                                                            (ISNULL(a.TG009,0) - ISNULL(a.TG020,0) - ISNULL(a.TG021,0) - ISNULL(b.UnconfirmedQty,0)) CanUseQty
                                                            ,(ISNULL(a.TG044,0) - ISNULL(a.TG046,0) - ISNULL(a.TG048,0) - ISNULL(b.UnconfirmedFreebieQty,0)) CanUseFreeSpareQty 
                                                            ,(ISNULL(a.TG052,0) - ISNULL(a.TG054,0) - ISNULL(a.TG055,0) - ISNULL(b.UnconfirmedAvailableQty,0)) CanUseAvailableQty 
                                                            FROM INVTG a
                                                            OUTER APPLY(
                                                                SELECT SUM(ISNULL(x1.TH008,0)) UnconfirmedQty 
                                                                        ,SUM(ISNULL(x1.TH024,0)) UnconfirmedFreebieQty 
                                                                        ,SUM(ISNULL(x1.TH061,0)) UnconfirmedAvailableQty
                                                                FROM COPTH x1
                                                                WHERE x1.TH032 = a.TG001
                                                                AND x1.TH033 = a.TG002
                                                                AND x1.TH034 = a.TG003
                                                                AND x1.TH020 ='N'
                                                            ) b
                                                            WHERE a.TG001 = @TG001
                                                            AND a.TG002 = @TG002
                                                            AND a.TG003 = @TG003
                                                            AND a.TG022 ='Y'
                                                            ";
                                                    dynamicParameters.Add("TG001", TH032);
                                                    dynamicParameters.Add("TG002", TH033);
                                                    dynamicParameters.Add("TG003", TH034);
                                                    result = sqlConnection.Query(sql, dynamicParameters);
                                                    if (result.Count() <= 0) throw new SystemException("暫出單找不到,請重新確認");
                                                    foreach (var item1 in result)
                                                    {
                                                        if (item1.TG024 == "Y") throw new SystemException("【訂單:" + TsnFull + "】已經結案!!!");
                                                        TG009 = Convert.ToDouble(item1.TG009);
                                                        TG044 = Convert.ToDouble(item1.TG044);
                                                        TG052 = Convert.ToDouble(item1.TG052);
                                                        CanUseQty = Convert.ToDouble(item1.CanUseQty);
                                                        CanUseFreeSpareQty = Convert.ToDouble(item1.CanUseFreeSpareQty);
                                                        CanUseAvailableQty = Convert.ToDouble(item1.CanUseAvailableQty);
                                                    }
                                                    #endregion

                                                    #region //判斷是否有超過可銷貨數
                                                    OverMtlItemNo = "銷貨單單身品號:" + TH004;
                                                    if ((CanUseQty - TH008) < 0)
                                                    {
                                                        OverMistake += "<br>已經超過可銷貨數量<br>暫出單數量:" + TG009 + ",目前可銷貨最大數:" + CanUseQty + "<br>欲新增銷貨數量:" + TH008 + ",請重新確認";
                                                    }
                                                    if ((CanUseFreeSpareQty - TH024) < 0)
                                                    {
                                                        OverMistake += "<br>已經超過可銷貨贈備品數量,暫出單贈備品數量:" + TG044 + ",目前可銷貨贈備品最大數量:" + CanUseFreeSpareQty + "<br>欲新增銷貨贈備品數量:" + TH024 + ",請重新確認<br>";
                                                    }

                                                    if (OverMistake.Length > 0)
                                                    {
                                                        AllOverMistake += OverMtlItemNo + OverMistake;
                                                    }
                                                    #endregion


                                                    break;
                                            }
                                            #endregion

                                        }

                                        if (AllOverMistake.Length > 0)
                                        {
                                            throw new SystemException(AllOverMistake);
                                        }
                                        #endregion

                                        #region //銷貨單身賦值
                                        addRoDetails
                                            .ToList()
                                            .ForEach(x =>
                                            {
                                                x.COMPANY = MESCompanyNo;
                                                x.USR_GROUP = USR_GROUP;
                                                x.FLAG = "1";
                                                x.CREATOR = userNo;
                                                x.CREATE_DATE = dateNow;
                                                x.CREATE_TIME = timeNow;
                                                x.CREATE_AP = BaseHelper.ClientComputer();
                                                x.CREATE_PRID = "BM";
                                                x.MODIFIER = "";
                                                x.MODI_DATE = "";
                                                x.MODI_TIME = "";
                                                x.MODI_AP = "";
                                                x.MODI_PRID = "";
                                                x.TH002 = RoErpNo;
                                                x.UDF01 = "";
                                                x.UDF02 = "";
                                                x.UDF03 = "";
                                                x.UDF04 = "";
                                                x.UDF05 = "";
                                                x.UDF06 = 0.0;
                                                x.UDF07 = 0.0;
                                                x.UDF08 = 0.0;
                                                x.UDF09 = 0.0;
                                                x.UDF10 = 0.0;
                                            });
                                        #endregion

                                        #region //新增銷貨單身 COPTH                                    
                                        sql = @"INSERT INTO COPTH(COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE, 
                                                FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID, 
                                                TH001, TH002, TH003, TH004, TH005, TH006, TH007, TH008, TH009, TH010,
                                                TH011, TH012, TH013, TH014, TH015, TH016, TH017, TH018, TH019, TH020,
                                                TH021, TH022, TH023, TH024, TH025, TH026, TH027, TH028, TH029, TH030,
                                                TH031, TH032, TH033, TH034, TH035, TH036, TH037, TH038, TH039, TH040,
                                                TH041, TH042, TH043, TH044, TH045, TH046, TH047, TH048, TH049, TH050,
                                                TH051, TH052, TH053, TH054, TH055, TH056, TH057, TH058, TH059, TH060,
                                                TH061, TH062, TH063, TH064, TH065, TH066, TH067, TH068, TH069, TH070,
                                                TH071, TH072, TH073, TH074, TH075, TH076, TH077, TH078, TH079, TH080,
                                                TH081, TH500, TH501, TH502, TH503, TH082, TH083, TH084, TH085, TH086,
                                                TH087, TH088, TH089, UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07,
                                                UDF08, UDF09, UDF10, TH090, TH091, TH092, TH093, TH094)
                                                VALUES(@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE, 
                                                @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID,@MODI_TIME, @MODI_AP, @MODI_PRID, 
                                                @TH001, @TH002, @TH003, @TH004, @TH005, @TH006, @TH007, @TH008, @TH009, @TH010,
                                                @TH011, @TH012, @TH013, @TH014, @TH015, @TH016, @TH017, @TH018, @TH019, @TH020,
                                                @TH021, @TH022, @TH023, @TH024, @TH025, @TH026, @TH027, @TH028, @TH029, @TH030,
                                                @TH031, @TH032, @TH033, @TH034, @TH035, @TH036, @TH037, @TH038, @TH039, @TH040,
                                                @TH041, @TH042, @TH043, @TH044, @TH045, @TH046, @TH047, @TH048, @TH049, @TH050,
                                                @TH051, @TH052, @TH053, @TH054, @TH055, @TH056, @TH057, @TH058, @TH059, @TH060,
                                                @TH061, @TH062, @TH063, @TH064, @TH065, @TH066, @TH067, @TH068, @TH069, @TH070,
                                                @TH071, @TH072, @TH073, @TH074, @TH075, @TH076, @TH077, @TH078, @TH079, @TH080,
                                                @TH081, @TH500, @TH501, @TH502, @TH503, @TH082, @TH083, @TH084, @TH085, @TH086,
                                                @TH087, @TH088, @TH089, @UDF01, @UDF02, @UDF03, @UDF04, @UDF05, @UDF06, @UDF07,
                                                @UDF08, @UDF09, @UDF10, @TH090, @TH091, @TH092, @TH093, @TH094)";
                                        rowsAffected += sqlConnection.Execute(sql, addRoDetails);
                                        #endregion
                                    }
                                    #endregion
                                    #endregion
                                }
                            }

                            using (SqlConnection sqlConnections = new SqlConnection(MainConnectionStrings))
                            {
                                DynamicParameters dynamicParameters = new DynamicParameters();
                                if (ReceiveOrders.Select(x => x.TransferStatusMES).FirstOrDefault() == "N") //沒有拋轉過
                                {
                                    #region //單頭
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE SCM.ReceiveOrder SET
                                            RoErpNo = @RoErpNo,
                                            TransferStatusMES = @TransferStatusMES,
                                            TransferTime = @LastModifiedDate,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE RoId = @RoId";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            RoErpNo = ReceiveOrders.Select(x => x.TG002).FirstOrDefault(),
                                            TransferStatusMES = "Y",
                                            LastModifiedDate,
                                            LastModifiedBy,
                                            RoId
                                        });
                                    var insertResult = sqlConnections.Query(sql, dynamicParameters);

                                    rowsAffected += sqlConnections.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //單身
                                    dynamicParameters = new DynamicParameters();
                                    RoDetails
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        sql = @"UPDATE SCM.RoDetail SET
                                                RoSequence = @RoSequence,
                                                TransferStatusMES = @TransferStatusMES,
                                                LastModifiedDate = @LastModifiedDate,
                                                LastModifiedBy = @LastModifiedBy
                                                WHERE RoDetailId = @RoDetailId";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                RoSequence = x.TH003,
                                                TransferStatusMES = "Y",
                                                LastModifiedDate,
                                                LastModifiedBy,
                                                RoDetailId = x.RoDetailId
                                            });
                                        insertResult = sqlConnections.Query(sql, dynamicParameters);

                                        rowsAffected += sqlConnections.Execute(sql, dynamicParameters);
                                    });
                                    
                                    #endregion
                                }
                                else if (ReceiveOrders.Select(x => x.TransferStatusMES).FirstOrDefault() == "R") //有拋轉過
                                {
                                    #region //單頭
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE SCM.ReceiveOrder SET
                                            TransferStatusMES = @TransferStatusMES,
                                            TransferTime = @LastModifiedDate,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE RoId = @RoId";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            TransferStatusMES = "Y",
                                            LastModifiedDate,
                                            LastModifiedBy,
                                            RoId
                                        });

                                    rowsAffected += sqlConnections.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //單身
                                    dynamicParameters = new DynamicParameters();
                                    RoDetails
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        sql = @"UPDATE SCM.RoDetail SET
                                                RoSequence = @RoSequence,
                                                TransferStatusMES = @TransferStatusMES,
                                                LastModifiedDate = @LastModifiedDate,
                                                LastModifiedBy = @LastModifiedBy
                                                WHERE RoDetailId = @RoDetailId";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                RoSequence = x.TH003,
                                                TransferStatusMES = "Y",
                                                LastModifiedDate,
                                                LastModifiedBy,
                                                RoDetailId = x.RoDetailId
                                            });

                                        rowsAffected += sqlConnections.Execute(sql, dynamicParameters);
                                    });

                                    #endregion
                                }

                            }

                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                msg = "拋轉成功"
                            });
                            #endregion

                            transactionScope.Complete();
                        }
                    }
                    else
                    {
                        throw new SystemException("銷貨單已拋轉ERP!!!");
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

        #region //UpdateReceiveOrderReviseMES -- 銷貨單回歸MES重新編輯 -- Shintokuro 2024-02-20
        public string UpdateReceiveOrderReviseMES(int RoId)
        {
            try
            {
                if (RoId < 0) throw new SystemException("銷貨單不可為空,請重新確認!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;
                        string dateNow = CreateDate.ToString("yyyyMMdd"), timeNow = CreateDate.ToString("HH:mm:ss"), errorMsg = "";
                        string MESCompanyNo = "", departmentNo = "", userNo = "", userName = "", userGroup = "";
                        string RoErpPrefix = "", RoErpNo = "", TransferStatusMES = "", ConfirmStatus = "";
                        int OrderCompanyIdBase = -1;


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
                            MESCompanyNo = item.ErpNo;
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
                        dynamicParameters.Add("FunctionCode", "ReceiveOrderManagment");
                        dynamicParameters.Add("DetailCode", "import");

                        var resultUser = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUser.Count() <= 0) throw new SystemException("使用者資料錯誤!");

                        foreach (var item in resultUser)
                        {
                            userNo = item.UserNo;
                            userName = item.UserName;
                            departmentNo = item.DepartmentNo;
                        }
                        #endregion

                        #region //確認銷貨單是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.TransferStatusMES, a.ConfirmStatus, a.CompanyId 
                                FROM SCM.ReceiveOrder a
                                WHERE a.CompanyId = @CompanyId
                                AND a.RoId = @RoId
                                ";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("RoId", RoId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("銷貨單單頭資料找不到，請重新確認");
                        foreach (var item in result)
                        {
                            TransferStatusMES = item.TransferStatusMES;
                            ConfirmStatus = item.ConfirmStatus;
                            OrderCompanyIdBase = item.CompanyId;
                        }
                        if (TransferStatusMES != "Y") throw new SystemException("銷貨單未處於拋轉ERP狀態不可以操作");
                        if (ConfirmStatus == "Y") throw new SystemException("銷貨單處於確認狀態不可以操作");
                        if (ConfirmStatus == "V") throw new SystemException("銷貨單處於作廢狀態不可以操作");
                        if (OrderCompanyIdBase != CurrentCompany) throw new SystemException("單據的公司別與目前公司別不同，請嘗試重新登入!!");
                        #endregion

                        #region //更新單頭
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.ReceiveOrder SET
                                TransferStatusMES = @TransferStatusMES,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RoId = @RoId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TransferStatusMES = "R",
                                LastModifiedDate,
                                LastModifiedBy,
                                RoId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //更新單身
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.RoDetail SET
                                TransferStatusMES = @TransferStatusMES,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RoId = @RoId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TransferStatusMES = "R",
                                LastModifiedDate,
                                LastModifiedBy,
                                RoId
                            });
                        insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "可以重新編輯單據了!"
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

        #region //UpdateRoConfirm -- MES銷貨單確認單據 -- Shintokuro 2024-04-17
        public string UpdateRoConfirm(int RoId)
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
                int CompanyIdBase = -1;
                int SoDetailId = -1;

                string RoErpPrefix = "";
                string RoErpNo = "";
                string ReceiveDate = "";
                string SourceType = "";
                string PriceSourceTypeMain = "";
                

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        string CustomerNo = "";
                        string Currency = "";
                        decimal TotalAmount = 0;
                        string CompanyNo = "";

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
                        dynamicParameters.Add("FunctionCode", "ReceiveOrderManagment");
                        dynamicParameters.Add("DetailCode", "confirm");

                        var resultUser = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUser.Count() <= 0) throw new SystemException("使用者資料錯誤!");

                        foreach (var item in resultUser)
                        {
                            UserNo = item.UserNo;
                        }
                        #endregion

                        #region //撈取單頭資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.RoErpPrefix,a.RoErpNo, a.ConfirmStatus,a.CompanyId
                                ,FORMAT(a.ReceiveDate, 'yyyyMMdd') ReceiveDate
                                ,a.SourceType,a.PriceSourceTypeMain
                                ,b.SoDetailId
                                ,c.CustomerNo, a.Currency, a.PretaxAmount ,d.CompanyNo
                                FROM SCM.ReceiveOrder a
                                INNER JOIN SCM.RoDetail b on a.RoId = b.RoId
                                INNER JOIN SCM.Customer c on a.CustomerId = c.CustomerId
                                INNER JOIN BAS.Company d on a.CompanyId = d.CompanyId
                                WHERE a.RoId = @RoId";
                        dynamicParameters.Add("RoId", RoId);

                        var resultRoMes = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in resultRoMes)
                        {
                            if (item.ConfirmStatus == "Y") throw new SystemException("該銷貨單單據已經核單,不能確認");
                            if (item.ConfirmStatus == "V") throw new SystemException("該銷貨單單據已經作廢,不能確認");
                            if (item.CompanyId != CurrentCompany) throw new SystemException("該銷貨單單據公司別與後端公司別不一致,不能確認");
                            CompanyIdBase = item.CompanyId;
                            RoErpPrefix = item.RoErpPrefix;
                            RoErpNo = item.RoErpNo;
                            ReceiveDate = item.ReceiveDate;
                            SoDetailId = item.SoDetailId;
                            SourceType = item.SourceType;
                            PriceSourceTypeMain = item.PriceSourceTypeMain;
                            CustomerNo = item.CustomerNo;
                            Currency = item.Currency;
                            TotalAmount = Convert.ToDecimal(item.PretaxAmount);
                            CompanyNo = item.CompanyNo;
                        }
                        //CheckCreditLimit(CustomerNo, Currency, TotalAmount, "ShippingOrder", 0, CompanyNo);
                        #endregion

                        #region //信用額度檢核
                        //string domainUrl = "http://192.168.20.208:16668/";
                        string domainUrl = "https://bm.zy-tech.com.tw/";


                        string targetDate = DateTime.Now.AddDays(-2).ToString("yyyyMMdd HH:mm:ss"),
                            targetUrl = string.Format("{0}{1}", domainUrl, "api/BM/CheckCreditLimit");

                        var postData = new List<KeyValuePair<string, string>>
                        {
                            new KeyValuePair<string, string>("CustomerNo", CustomerNo),
                            new KeyValuePair<string, string>("Currency", Currency),
                            new KeyValuePair<string, string>("TotalAmount", TotalAmount.ToString()),
                            new KeyValuePair<string, string>("DocType", "ShippingOrder"),
                            new KeyValuePair<string, string>("Amount", "0"),
                            new KeyValuePair<string, string>("CompanyNo", CompanyNo)
                        };

                        string response = BaseHelper.PostWebRequest(targetUrl, postData);

                        if (response.TryParseJson(out JObject tempJObject))
                        {
                            JObject resultJson = JObject.Parse(response);

                            Console.WriteLine("狀態：" + resultJson["status"].ToString());
                            Console.WriteLine("回傳訊息：" + resultJson["msg"].ToString());

                            if (resultJson["status"].ToString() != "ok")
                            {
                                throw new SystemException(resultJson["msg"].ToString());
                            }
                        }
                        else
                        {
                            logger.Error(response);
                        }
                        #endregion
                    }
                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //審核ERP權限
                        USR_GROUP = BaseHelper.CheckErpAuthority(UserNo, sqlConnection, "COPI08", "CONFIRM");
                        #endregion

                        #region //判斷開立日期是否超過結帳日
                        string CloseDateBase = "";
                        string ReceiveDateBase = ReceiveDate;
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
                                throw new SystemException("【銷貨單】ERP已經結帳,不可以在新增【" + CloseDateBase + "】之前的單據");
                            }
                        }
                        else
                        {
                            throw new SystemException("日期字符串无效，无法比较");
                        }
                        #endregion

                        #region //取出銷貨單單據資料 + 異動更新相關資料表
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.TG003,a.TG023 ,a.TG007,
                                b.TH003,b.TH004,b.TH031,
                                b.TH008,b.TH007,b.TH017,b.TH024,b.TH061,
                                b.TH014,b.TH015,b.TH016,
                                b.TH032,b.TH033,b.TH034,b.TH060,
                                b.TH012,b.TH037,
                                c.MB022,c.MB004,
                                ISNULL(md1.RoRate,1) RoRate, 
                                ISNULL(md2.RoAvailableRate,1) RoAvailableRate, 
                                (b.TH008*ISNULL(md1.RoRate,1)) ConversionQty,
                                (b.TH024*ISNULL(md1.RoRate,1)) ConversionFreeSpareQty,
                                (b.TH061*ISNULL(md2.RoAvailableRate,1)) ConversionAvailableQty,
                                ROUND((b.TH037 / (b.TH061*ISNULL(md2.RoAvailableRate,1))),CAST(md3.MF005 AS INT)) UnitCost
                                FROM COPTG a
                                INNER JOIN COPTH b on a.TG001 = b.TH001 AND a.TG002 = b.TH002
                                INNER JOIN INVMB c on b.TH004 = c.MB001
                                OUTER APPLY(
                                    SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) RoRate
                                    FROM INVMD x1
                                    WHERE x1.MD001 = b.TH004
                                    AND  x1.MD002 = b.TH009
                                ) md1
                                OUTER APPLY(
                                    SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) RoAvailableRate
                                    FROM INVMD x1
                                    WHERE x1.MD001 = b.TH004
                                    AND  x1.MD002 = b.TH062
                                ) md2
                                OUTER APPLY(
                                    SELECT x1.MF003,x1.MF004,x1.MF005,x1.MF006
                                    FROM CMSMF x1
                                    INNER JOIN CMSMA x2 on x1.MF001 = x2.MA003 
                                ) md3
                                WHERE a.TG001 = @RoErpPrefix
                                AND a.TG002 = @RoErpNo";
                        dynamicParameters.Add("RoErpPrefix", RoErpPrefix);
                        dynamicParameters.Add("RoErpNo", RoErpNo);
                        var resultRoErp = sqlConnection.Query(sql, dynamicParameters);
                        if (resultRoErp.Count() <= 0) throw new SystemException("找不到ERP銷貨單單據!");
                        foreach (var item in resultRoErp)
                        {
                            if (item.TG023 == "Y") throw new SystemException("該銷貨單單據已經核單,不能確認");
                            if (item.TG023 == "V") throw new SystemException("該銷貨單單據已經作廢,不能確認");

                            #region //欄位參數撈取
                            string ReceiveDateErp = item.TG003; //TG003 單據日
                            string ConfirmStatus = item.TG023; //TG023 單頭確認
                            string CustomerFullName = item.TG007; //TG007 客戶全名
                            string RoSequence = item.TH003; //TH003 序號
                            string MtlItemNo = item.TH004; //TH004 品號
                            string TypeErp = item.TH031; //TH031 類型(1.贈品/2.備品)
                            string InventoryNo = item.TH007; //TH007 庫別
                            string LotNumber = item.TH017; //TH017 批號
                            string Location = item.TH060; //TH060 儲位
                            double Quantity = Convert.ToDouble(item.TH008); //TH008 數量
                            double FreeSpareQty = Convert.ToDouble(item.TH024); //TH024 贈/備品數量
                            double AvailableQty = Convert.ToDouble(item.TH061); //TH061 計價數量
                            string TH014 = item.TH014; //TH014 訂單單別　/ TH015 訂單單號 / TH016 訂單序號
                            string TH015 = item.TH015;
                            string TH016 = item.TH016;
                            string SoFull = item.TH014 + '-'+ item.TH015 + '-' + item.TH016;
                            string TH032 = item.TH032; //TH032 暫出單別 / TH033 暫出單號 / TH034 暫出序號
                            string TH033 = item.TH033;
                            string TH034 = item.TH034;
                            string TsnFull = item.TH032 + '-'+ item.TH033 + '-' + item.TH034;
                            string LotManagement = item.MB022; //品號批號管理
                            string InventoryUomNo = item.MB004; //庫存單位

                            string AllOverMistake = "";
                            string OverMtlItemNo = "";
                            string OverMistake = "";
                            string ClosureStatusCOPTD = "N";
                            string ClosureStatusINVTG = "N";
                            string TD049 = "";
                            double TD008 = 0;
                            double TD024 = 0;
                            double TD050 = 0;
                            double TG009 = 0;
                            double TG044 = 0;
                            double TG052 = 0;

                            double CanUseQty = 0;
                            double CanUseFreebieQty = 0;
                            double CanUseSpareQty = 0;
                            double CanUseFreeSpareQty = 0;
                            double CanUseAvailableQty = 0;

                            double RoRate = Convert.ToDouble(item.RoRate); //銷貨-單位換算率
                            double RoAvailableRate = Convert.ToDouble(item.RoAvailableRate); //銷貨-計價單位換算率
                            double ConversionQty = Convert.ToDouble(item.ConversionQty); //銷貨-單位換算後數量
                            double ConversionFreeSpareQty = Convert.ToDouble(item.ConversionFreeSpareQty); //銷貨-單位換算後贈備品數量
                            double ConversionAvailableQty = Convert.ToDouble(item.ConversionAvailableQty); //銷貨-單位換算後計價數量
                            double UnitCost = Convert.ToDouble(item.UnitCost); //單位成本
                            double PreTaxAmt = Convert.ToDouble(item.TH037); //本幣未稅金額
                            double UnitPrice = Convert.ToDouble(item.TH012); //原幣單價

                            double SoRate = 0;
                            double SoAvailableRate = 0;
                            double TsnRate = 0;
                            double TsnAvailableRate = 0;
                            #endregion

                            #region //檢查段
                            #region //數量是否有超收檢查
                            #region //訂單部分
                            #region //撈取訂單數量資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 TD049,a.TD016,TD010,TD077,TD008,TD009,TD076,TD078,
                                    ISNULL(a.TD008,0) TD008,
                                    ISNULL(a.TD024,0) TD024,
                                    ISNULL(a.TD050,0) TD050,
                                    (ISNULL(a.TD008,0) * ISNULL(md1.SoRate,1)) - (ISNULL(a.TD009,0) * ISNULL(md1.SoRate,1)) CanUseQty,
                                    (ISNULL(a.TD024,0) * ISNULL(md1.SoRate,1)) - (ISNULL(a.TD025,0) * ISNULL(md1.SoRate,1)) CanUseFreebieQty,
                                    (ISNULL(a.TD050,0) * ISNULL(md1.SoRate,1)) - (ISNULL(a.TD051,0) * ISNULL(md1.SoRate,1)) CanUseSpareQty,
                                    (ISNULL(a.TD076,0) * ISNULL(md2.SoAvailableRate,1)) - (ISNULL(a.TD078,0) * ISNULL(md2.SoAvailableRate,1)) CanUseAvailableQty,
                                    ISNULL(md1.SoRate,1) SoRate, 
                                    ISNULL(md2.SoAvailableRate,1) SoAvailableRate, 
                                    (ISNULL(a.TD008,0) * ISNULL(md1.SoRate,1)) SoConversionQty,
                                    (ISNULL(a.TD009,0) * ISNULL(md1.SoRate,1)) SoConversionSiQty,
                                    (ISNULL(a.TD024,0) * ISNULL(md1.SoRate,1)) SoConversionFreebieQty,
                                    (ISNULL(a.TD025,0) * ISNULL(md1.SoRate,1)) SoConversionFreebieSiQty,
                                    (ISNULL(a.TD050,0) * ISNULL(md1.SoRate,1)) SoConversionSpareQty,
                                    (ISNULL(a.TD051,0) * ISNULL(md1.SoRate,1)) SoConversionSpareSiQty,
                                    (ISNULL(a.TD076,0) * ISNULL(md2.SoAvailableRate,1)) SoConversionSoPriceQty,
                                    (ISNULL(a.TD078,0) * ISNULL(md2.SoAvailableRate,1)) SoConversionSoPriceSiQty
                                    FROM COPTD a
                                    OUTER APPLY(
                                        SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) SoRate
                                        FROM INVMD x1
                                        WHERE x1.MD001= a.TD004
                                        AND  x1.MD002 = a.TD010
                                    ) md1
                                    OUTER APPLY(
                                        SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) SoAvailableRate
                                        FROM INVMD x1
                                        WHERE x1.MD001= a.TD004
                                        AND  x1.MD002 = a.TD077
                                    ) md2
                                    WHERE a.TD001 = @TH014
                                    AND a.TD002 = @TH015
                                    AND a.TD003 = @TH016
                                    AND a.TD021 = 'Y'";
                            dynamicParameters.Add("TH014", TH014);
                            dynamicParameters.Add("TH015", TH015);
                            dynamicParameters.Add("TH016", TH016);
                            var resultCOPTD = sqlConnection.Query(sql, dynamicParameters);
                            if (resultCOPTD.Count() <= 0) throw new SystemException("訂單找不到,請重新確認");
                            foreach (var item1 in resultCOPTD)
                            {
                                if(item1.TD016 == "Y") throw new SystemException("【訂單:" + SoFull + "】已經結案!!!");
                                TD049 = item1.TD049;
                                TD008 = Convert.ToDouble(item1.TD008);
                                TD024 = Convert.ToDouble(item1.TD024);
                                TD050 = Convert.ToDouble(item1.TD050);
                                CanUseQty = Convert.ToDouble(item1.CanUseQty); //可使用訂單數量
                                CanUseFreebieQty = Convert.ToDouble(item1.CanUseFreebieQty); //可使用贈品數量
                                CanUseSpareQty = Convert.ToDouble(item1.CanUseSpareQty); //可使用備品數量
                                CanUseAvailableQty = Convert.ToDouble(item1.CanUseAvailableQty); //可使用計價數量
                                SoRate = Convert.ToDouble(item1.SoRate);//數量單位轉換成庫存單位的換算率
                                SoAvailableRate = Convert.ToDouble(item1.SoAvailableRate);//計價單位轉換成庫存單位的換算率
                                double SoConversionQty = Convert.ToDouble(item1.SoConversionQty); //轉換成庫存單位的訂單數量
                                double SoConversionSiQty = Convert.ToDouble(item1.SoConversionSiQty);//轉換成庫存單位的已交訂單數量
                                double SoConversionFreebieQty = Convert.ToDouble(item1.SoConversionFreebieQty);//轉換成庫存單位的贈品數量
                                double SoConversionFreebieSiQty = Convert.ToDouble(item1.SoConversionFreebieSiQty);//轉換成庫存單位的已交贈品數量
                                double SoConversionSpareQty = Convert.ToDouble(item1.SoConversionSpareQty);//轉換成庫存單位的備品數量
                                double SoConversionSpareSiQty = Convert.ToDouble(item1.SoConversionSpareSiQty);//轉換成庫存單位的已交備品數量
                                double SoConversionSoPriceQty = Convert.ToDouble(item1.SoConversionSoPriceQty);//轉換成庫存單位的計價數量
                                double SoConversionSoPriceSiQty = Convert.ToDouble(item1.SoConversionSoPriceSiQty);//轉換成庫存單位的已交計價數量
                                
                                #region //判斷是否有超過可銷貨數
                                OverMtlItemNo = "銷貨單單身品號:" + MtlItemNo;
                                if ((CanUseQty - ConversionQty) < 0)
                                {
                                    OverMistake += "<br>已經超過可銷貨數量<br>訂單數量:" + SoConversionQty + InventoryUomNo + ",目前可銷貨最大數:" + CanUseQty + InventoryUomNo + "<br>欲新增銷貨數量:" + ConversionQty + InventoryUomNo + ",請重新確認";
                                }
                                if (TD049 == "1")
                                {
                                    if ((CanUseFreebieQty - ConversionFreeSpareQty) < 0)
                                    {
                                        OverMistake += "<br>已經超過可銷貨贈品數量,訂單贈品數量:" + SoConversionFreebieQty + InventoryUomNo + ",目前可銷貨贈品最大數量:" + CanUseFreebieQty + InventoryUomNo + "<br>欲新增銷貨贈品數量:" + ConversionFreeSpareQty + InventoryUomNo + ",請重新確認<br>";
                                    }
                                    if (CanUseQty - Quantity == 0 && CanUseFreebieQty - FreeSpareQty == 0)
                                    {
                                        ClosureStatusCOPTD = "Y";
                                    }
                                }
                                else if (TD049 == "2")
                                {
                                    if ((CanUseSpareQty - ConversionFreeSpareQty) < 0)
                                    {
                                        OverMistake += "<br>已經超過可銷貨備品數量,訂單備品數量:" + SoConversionSpareQty + InventoryUomNo + ",目前可銷貨貨品最大數量:" + CanUseSpareQty + InventoryUomNo + "<br>欲新增銷貨品數量:" + ConversionFreeSpareQty + InventoryUomNo + ",請重新確認<br>";
                                    }
                                    if (CanUseQty - ConversionQty == 0 && CanUseSpareQty - ConversionFreeSpareQty == 0)
                                    {
                                        ClosureStatusCOPTD = "Y";
                                    }
                                }
                                if (OverMistake.Length > 0)
                                {
                                    AllOverMistake += OverMtlItemNo + OverMistake;
                                }

                                if (AllOverMistake.Length > 0)
                                {
                                    throw new SystemException(AllOverMistake);
                                }
                                #endregion

                            }
                            #endregion

                            #endregion

                            #region //暫出單部分
                            if (SourceType == "To" || TH032 != "")
                            {
                                CanUseQty = 0;
                                CanUseFreeSpareQty = 0;
                                CanUseAvailableQty = 0;
                                AllOverMistake = "";
                                OverMtlItemNo = "";
                                OverMistake = "";
                                #region //撈取暫出單數量資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.TG024,ISNULL(a.TG009,0) TG009,ISNULL(a.TG044,0) TG044,ISNULL(a.TG052,0) TG052,
                                        ((ISNULL(a.TG009,0) * ISNULL(md1.TsnRate,1)) - (ISNULL(a.TG020,0) * ISNULL(md1.TsnRate,1)) - (ISNULL(a.TG021,0) * ISNULL(md1.TsnRate,1))) CanUseQty,
                                        ((ISNULL(a.TG044,0) * ISNULL(md1.TsnRate,1)) - (ISNULL(a.TG046,0) * ISNULL(md1.TsnRate,1)) - (ISNULL(a.TG048,0) * ISNULL(md1.TsnRate,1))) CanUseFreeSpareQty,
                                        ((ISNULL(a.TG052,0) * ISNULL(md2.TsnAvailableRate,1)) - (ISNULL(a.TG054,0) * ISNULL(md2.TsnAvailableRate,1)) - (ISNULL(a.TG055,0) * ISNULL(md2.TsnAvailableRate,1))) CanUseAvailableQty,
                                        ISNULL(md1.TsnRate,1) TsnRate, 
                                        ISNULL(md2.TsnAvailableRate,1) TsnAvailableRate,
                                        (ISNULL(a.TG009,0) * ISNULL(md1.TsnRate,1)) TsnConversionQty,
                                        (ISNULL(a.TG044,0) * ISNULL(md1.TsnRate,1)) TsnConversionFreeSpareQty,
                                        (ISNULL(a.TG052,0) * ISNULL(md2.TsnAvailableRate,1)) TsnConversionAvailableQty
                                        FROM INVTG a
                                        OUTER APPLY(
                                            SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) TsnRate
                                            FROM INVMD x1
                                            WHERE x1.MD001= a.TG004
                                            AND  x1.MD002 = a.TG010
                                        ) md1
                                        OUTER APPLY(
                                            SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) TsnAvailableRate
                                            FROM INVMD x1
                                            WHERE x1.MD001= a.TG004
                                            AND  x1.MD002 = a.TG053
                                        ) md2
                                        WHERE a.TG001 = @TG001
                                        AND a.TG002 = @TG002
                                        AND a.TG003 = @TG003
                                        AND a.TG022 ='Y'
                                        ";
                                dynamicParameters.Add("TG001", TH032);
                                dynamicParameters.Add("TG002", TH033);
                                dynamicParameters.Add("TG003", TH034);
                                var resultINVTG = sqlConnection.Query(sql, dynamicParameters);
                                if (resultINVTG.Count() <= 0) throw new SystemException("暫出單找不到,請重新確認");
                                foreach (var item1 in resultINVTG)
                                {
                                    if(item1.TG024 == "Y") throw new SystemException("【訂單:" + TsnFull + "】已經結案!!!");
                                    TG009 = Convert.ToDouble(item1.TG009);
                                    TG044 = Convert.ToDouble(item1.TG044);
                                    TG052 = Convert.ToDouble(item1.TG052);
                                    CanUseQty = Convert.ToDouble(item1.CanUseQty);
                                    CanUseFreeSpareQty = Convert.ToDouble(item1.CanUseFreeSpareQty);
                                    CanUseAvailableQty = Convert.ToDouble(item1.CanUseAvailableQty);
                                    TsnRate = Convert.ToDouble(item1.TsnRate);
                                    TsnAvailableRate = Convert.ToDouble(item1.TsnAvailableRate);
                                    double TsnConversionQty = Convert.ToDouble(item1.TsnConversionQty);
                                    double TsnConversionFreeSpareQty = Convert.ToDouble(item1.TsnConversionFreeSpareQty);
                                    double TsnConversionAvailableQty = Convert.ToDouble(item1.TsnConversionAvailableQty);

                                    #region //判斷是否有超過可銷貨數
                                    OverMtlItemNo = "銷貨單單身品號:" + MtlItemNo;
                                    if ((CanUseQty - ConversionQty) < 0)
                                    {
                                        OverMistake += "<br>已經超過可銷貨數量<br>暫出單數量:" + TsnConversionQty + InventoryUomNo + ",目前可銷貨最大數:" + CanUseQty + InventoryUomNo + "<br>欲新增銷貨數量:" + ConversionQty + InventoryUomNo + ",請重新確認";
                                    }
                                    if ((CanUseFreeSpareQty - ConversionFreeSpareQty) < 0)
                                    {
                                        OverMistake += "<br>已經超過可銷貨贈備品數量,暫出單贈備品數量:" + TsnConversionFreeSpareQty + InventoryUomNo + ",目前可銷貨贈備品最大數量:" + CanUseFreeSpareQty + InventoryUomNo + "<br>欲新增銷貨贈備品數量:" + ConversionFreeSpareQty + InventoryUomNo + ",請重新確認<br>";
                                    }

                                    if (OverMistake.Length > 0)
                                    {
                                        AllOverMistake += OverMtlItemNo + OverMistake;
                                    }
                                    if (CanUseQty - ConversionQty == 0 && CanUseFreeSpareQty - ConversionFreeSpareQty == 0)
                                    {
                                        ClosureStatusINVTG = "Y";
                                    }

                                    if (AllOverMistake.Length > 0)
                                    {
                                        throw new SystemException(AllOverMistake);
                                    }
                                    #endregion

                                }
                                #endregion
                            }
                            #endregion
                            #endregion

                            #region //INVMC 品號庫別檔 搜尋該品號是否有庫別庫存
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT MC007
                                    FROM INVMC
                                    WHERE MC001 = @MtlItemNo
                                    AND MC002 = @InventoryNo";
                            dynamicParameters.Add("MtlItemNo", MtlItemNo);
                            dynamicParameters.Add("InventoryNo", InventoryNo);
                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【品號：" + MtlItemNo + "庫別:" + InventoryNo + "】找不到庫存資料,請重新確認!!");
                            foreach(var item1 in result)
                            {
                                if(Convert.ToDouble(item1.MC007)< ConversionQty + ConversionFreeSpareQty) throw new SystemException("【品號：" + MtlItemNo + "庫別:"+ InventoryNo + "】庫存不足請重新確認");
                            }
                            #endregion

                            #region //品號有批號管理觸發
                            if (LotManagement != "N")
                            {
                                #region //INVMF 品號批號資料單身 判斷該品號批號庫存是否足夠
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT SUM(MF008*MF010)  AS LotQty, SUM(MF008*MF014) AS K_QTY
                                        FROM INVMF
                                        WHERE MF001 = @MtlItemNo 
                                        AND MF002 = @LotNumber 
                                        AND MF007 = @InventoryNo";
                                dynamicParameters.Add("MtlItemNo", MtlItemNo);
                                dynamicParameters.Add("LotNumber", LotNumber);
                                dynamicParameters.Add("InventoryNo", InventoryNo);

                                var resultCheckLot = sqlConnection.Query(sql, dynamicParameters);
                                if (resultCheckLot.Count() <= 0) throw new SystemException("找不到庫存資訊!!!");
                                foreach(var item1 in resultCheckLot)
                                {
                                    if (Convert.ToDouble(item1.LotQty)< ConversionQty + ConversionFreeSpareQty) throw new SystemException("【品號：" + MtlItemNo + "庫別:"+ InventoryNo + "批號:" + LotNumber + "】庫存不足請重新確認");
                                }
                                #endregion

                                #region //INVLF 品號庫別儲位批號檔 判斷改品號批號庫別儲位庫存數量是否足夠
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT SUM(LF008*LF011)  AS LotQty, SUM(LF008*LF012) AS K_QTY
                                        FROM INVLF
                                        WHERE LF004 = @MtlItemNo 
                                        AND LF005 = @InventoryNo 
                                        AND LF006 = @Location
                                        AND LF007 = @LotNumber";
                                dynamicParameters.Add("MtlItemNo", MtlItemNo);
                                dynamicParameters.Add("InventoryNo", InventoryNo);
                                dynamicParameters.Add("Location", Location != "" ? Location : "##########");
                                dynamicParameters.Add("LotNumber", LotNumber);

                                resultCheckLot = sqlConnection.Query(sql, dynamicParameters);
                                if (resultCheckLot.Count() <= 0) throw new SystemException("找不到庫存資訊!!!");
                                foreach (var item1 in resultCheckLot)
                                {
                                    if (Convert.ToDouble(item1.LotQty) < ConversionQty + ConversionFreeSpareQty) throw new SystemException("【品號：" + MtlItemNo + "庫別:" + InventoryNo + "儲位:" + Location != "" ? Location : "##########" + "批號:" + LotNumber + "】庫存不足請重新確認");
                                }
                                #endregion
                            }
                            #endregion
                            #endregion

                            #region //異動更新段-銷貨單確認後
                            #region //COPTD 銷貨單確認後回寫訂單單身 已交數量
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE COPTD SET
                                    MODIFIER = @MODIFIER,
                                    MODI_DATE = @MODI_DATE,
                                    MODI_TIME = @MODI_TIME,
                                    MODI_AP = @MODI_AP,
                                    MODI_PRID = @MODI_PRID,
                                    FLAG = FLAG + 1,
                                    TD009  = TD009 + @TD009,
                                    TD025  = TD025 + @TD025,
                                    TD051  = TD051 + @TD051,
                                    TD078  = TD078 + @TD078,
                                    TD016  = @TD016 
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
                                TD009 = Math.Round(ConversionQty / SoRate, 3, MidpointRounding.AwayFromZero),
                                TD025 = TypeErp == "1" ? Math.Round(ConversionFreeSpareQty / SoRate, 3, MidpointRounding.AwayFromZero) : 0,
                                TD051 = TypeErp == "2" ? Math.Round(ConversionFreeSpareQty / SoRate, 3, MidpointRounding.AwayFromZero) : 0,
                                TD078 = Math.Round(ConversionAvailableQty / SoAvailableRate, 3, MidpointRounding.AwayFromZero),
                                TD016 = ClosureStatusCOPTD,
                                TD001 = TH014,
                                TD002 = TH015,
                                TD003 = TH016
                            });
                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //銷貨單來源為暫出單觸發
                            if (SourceType == "To" || TH032 != "")
                            {
                                #region //INVTF 暫出入轉撥單頭檔更新 轉銷日更新
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE INVTF SET
                                        MODIFIER = @MODIFIER,
                                        MODI_DATE = @MODI_DATE,
                                        MODI_TIME = @MODI_TIME,
                                        MODI_AP = @MODI_AP,
                                        MODI_PRID = @MODI_PRID,
                                        FLAG = FLAG + 1,
                                        TF042  = @TF042
                                        WHERE TF001 = @TF001
                                        AND TF002 = @TF002";
                                dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MODIFIER = UserNo,
                                    MODI_DATE = dateNow,
                                    MODI_TIME = timeNow,
                                    MODI_AP = BaseHelper.ClientComputer(),
                                    MODI_PRID = "BM",
                                    TF042 = ReceiveDateErp,
                                    TF001 = TH032,
                                    TF002 = TH033
                                });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region //INVTG 暫出入轉撥單身檔 轉銷數量
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE INVTG SET
                                        MODIFIER = @MODIFIER,
                                        MODI_DATE = @MODI_DATE,
                                        MODI_TIME = @MODI_TIME,
                                        MODI_AP = @MODI_AP,
                                        MODI_PRID = @MODI_PRID,
                                        FLAG = FLAG + 1,
                                        TG020  = TG020 + @TG020,
                                        TG046  = TG046 + @TG046,
                                        TG054  = TG054 + @TG054,
                                        TG024  = @TG024
                                        WHERE TG001 = @TG001
                                        AND TG002 = @TG002
                                        AND TG003 = @TG003";
                                dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MODIFIER = UserNo,
                                    MODI_DATE = dateNow,
                                    MODI_TIME = timeNow,
                                    MODI_AP = BaseHelper.ClientComputer(),
                                    MODI_PRID = "BM",
                                    TG020 = Math.Round(ConversionQty / TsnRate, 3, MidpointRounding.AwayFromZero),
                                    TG046 = Math.Round(ConversionFreeSpareQty / TsnRate, 3, MidpointRounding.AwayFromZero),
                                    TG054 = Math.Round(ConversionAvailableQty / TsnAvailableRate, 3, MidpointRounding.AwayFromZero),
                                    TG001 = TH032,
                                    TG002 = TH033,
                                    TG003 = TH034,
                                    TG024 = ClosureStatusINVTG
                                });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }
                            #endregion

                            #region //INVLA 交易紀錄建立
                            #region //INVLB 取得月初總數量,月初總成本, 單據單位成本
                            double LB010 = 0;
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 LB003,LB004,LB010
                                    FROM INVLB 
                                    WHERE LB001 =@MtlItemNo
                                    AND LB010 != 0
                                    ORDER BY LB002 DESC";
                            dynamicParameters.Add("MtlItemNo", MtlItemNo);
                            var resultINVLB = sqlConnection.Query(sql, dynamicParameters);

                            if (resultINVLB.Count() > 0)
                            {
                                foreach(var item1 in resultINVLB)
                                {
                                    LB010 = Convert.ToDouble(item1.LB010);
                                }
                            }
                            #endregion

                            #region //金額小數位數取位
                            #region //取得本國幣別
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 LTRIM(RTRIM(MA003)) MA003 FROM CMSMA";

                            var resultCMSMA = sqlConnection.Query(sql, dynamicParameters);

                            if (resultCMSMA.Count() <= 0) throw new SystemException("共用參數檔案資料錯誤!!");

                            string Currency = "";
                            foreach (var item1 in resultCMSMA)
                            {
                                Currency = item1.MA003;
                            }
                            #endregion

                            #region //小數點後取位
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.MF004,a.MF005,a.MF006
                                    FROM CMSMF a 
                                    WHERE a.MF001 = @Currency";
                            dynamicParameters.Add("Currency", Currency);

                            var resultCMSMF = sqlConnection.Query(sql, dynamicParameters);

                            if (resultCMSMF.Count() <= 0) throw new SystemException("金額小數點取位資料錯誤!!");

                            int AmountRounding = 0; //金額取位
                            foreach (var item1 in resultCMSMF)
                            {
                                AmountRounding = Convert.ToInt32(item1.MF00ˋ);
                            }
                            #endregion
                            #endregion

                            #region //新增INVLA
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
                                    LA004 = ReceiveDateErp,
                                    LA005 = -1,
                                    LA006 = RoErpPrefix,
                                    LA007 = RoErpNo,
                                    LA008 = RoSequence,
                                    LA009 = InventoryNo,
                                    LA010 = CustomerFullName,
                                    LA011 = ConversionQty + ConversionFreeSpareQty,
                                    LA012 = 0,
                                    LA013 = 0,
                                    LA014 = "2",
                                    LA015 = "N",
                                    LA016 = LotNumber,
                                    LA017 = 0,
                                    LA018 = 0,
                                    LA019 = 0,
                                    LA020 = 0,
                                    LA021 = 0,
                                    LA022 = Location != "" ? Location: ""
                                });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                            #endregion

                            #region //INVMB 品號基本資料檔 品號庫存更新
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE INVMB SET
                                    MODIFIER = @MODIFIER,
                                    MODI_DATE = @MODI_DATE,
                                    MODI_TIME = @MODI_TIME,
                                    MODI_AP = @MODI_AP,
                                    MODI_PRID = @MODI_PRID,
                                    FLAG = FLAG + 1,
                                    MB064  = MB064 - @MB064
                                    WHERE MB001 = @MB001";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                MODIFIER = UserNo,
                                MODI_DATE = dateNow,
                                MODI_TIME = timeNow,
                                MODI_AP = BaseHelper.ClientComputer(),
                                MODI_PRID = "BM",
                                MB064 = ConversionQty + ConversionFreeSpareQty,
                                MB001 = MtlItemNo
                            });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //INVMC 品號庫別檔 品號庫別庫存更新
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE INVMC SET
                                    MODIFIER = @MODIFIER,
                                    MODI_DATE = @MODI_DATE,
                                    MODI_TIME = @MODI_TIME,
                                    MODI_AP = @MODI_AP,
                                    MODI_PRID = @MODI_PRID,
                                    FLAG = FLAG + 1,
                                    MC007  = MC007 - @MC007,
                                    MC013  = @MC013 
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
                                MC007 = ConversionQty + ConversionFreeSpareQty,
                                MC013 = ReceiveDateErp,
                                MC001 = MtlItemNo,
                                MC002 = InventoryNo
                            });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //品號有批號管理觸發
                            if (LotManagement != "N")
                            {
                                #region //INVMF 品號批號資料單身 資料建立
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO INVMF (COMPANY, CREATOR, USR_GROUP, CREATE_DATE,
                                       MODIFIER, MODI_DATE, FLAG, UDF01, UDF02,
                                       UDF03, UDF04, UDF05, UDF06, UDF07,
                                       UDF08, UDF09, UDF10,
                                       MF001, MF002, MF003, MF004, MF005, 
                                       MF006, MF007, MF008, MF009, MF010, 
                                       MF014)
                                VALUES(@COMPANY, @CREATOR, @USR_GROUP,
                                       CONVERT(varchar(100),GETDATE(),112),@CREATOR,CONVERT(varchar(100),GETDATE(),112),0,
                                       '','','','','',0,0,0,0,0,
                                       @MF001 , @MF002 , @MF003 , @MF004 , @MF005,
                                       @MF006 , @MF007 , @MF008 , @MF009 , @MF010,
                                       @MF014)";
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
                                        MF004 = RoErpPrefix,
                                        MF005 = RoErpNo,
                                        MF006 = RoSequence,
                                        MF007 = InventoryNo,
                                        MF008 = -1,
                                        MF009 = "2",
                                        MF010 = ConversionQty + ConversionFreeSpareQty,
                                        MF014 = 0
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region //INVLF 儲位批號異動明細資料檔 資料建立
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
                                       @LF011, @LF012 , @LF013)";
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
                                        LF001 = RoErpPrefix,
                                        LF002 = RoErpNo,
                                        LF003 = RoSequence,
                                        LF004 = MtlItemNo,
                                        LF005 = InventoryNo,
                                        LF006 = Location != "" ? Location : "##########",
                                        LF007 = LotNumber,
                                        LF008 = -1,
                                        LF009 = ReceiveDateErp,
                                        LF010 = "2",
                                        LF011 = "2",
                                        LF012 = ConversionQty + ConversionFreeSpareQty,
                                        LF013 = 0
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region //INVMM 品號庫別儲位批號檔 庫存數量更新
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE INVMM SET
                                        MODIFIER = @MODIFIER,
                                        MODI_DATE = @MODI_DATE,
                                        MODI_TIME = @MODI_TIME,
                                        MODI_AP = @MODI_AP,
                                        MODI_PRID = @MODI_PRID,
                                        FLAG = FLAG + 1,
                                        MM005  = MM005 - @MM005,
                                        MM009  = @MM009
                                        WHERE MM001 = @MM001
                                        AND MM002 = @MM002
                                        AND MM003 = @MM003";
                                dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MODIFIER = UserNo,
                                    MODI_DATE = dateNow,
                                    MODI_TIME = timeNow,
                                    MODI_AP = BaseHelper.ClientComputer(),
                                    MODI_PRID = "BM",
                                    MM005 = ConversionQty + ConversionFreeSpareQty,
                                    MM009 = ReceiveDateErp,
                                    MM001 = MtlItemNo,
                                    MM002 = InventoryNo,
                                    MM003 = Location != "" ? Location : "##########"
                                });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }
                            #endregion
                            #endregion

                        }
                        #endregion

                        #region //銷貨單單據確認
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE COPTG SET 
                                MODIFIER = @MODIFIER,
                                MODI_DATE = @MODI_DATE,
                                MODI_TIME = @MODI_TIME,
                                MODI_AP = @MODI_AP,
                                MODI_PRID = @MODI_PRID,
                                FLAG = FLAG + 1,
                                TG023 = 'Y'
                                WHERE TG001 = @RoErpPrefix
                                AND TG002 = @RoErpNo";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            MODIFIER = UserNo,
                            MODI_DATE = dateNow,
                            MODI_TIME = timeNow,
                            MODI_AP = BaseHelper.ClientComputer(),
                            MODI_PRID = "BM",
                            RoErpPrefix,
                            RoErpNo
                        });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE COPTH SET
                                MODIFIER = @MODIFIER,
                                MODI_DATE = @MODI_DATE,
                                MODI_TIME = @MODI_TIME,
                                MODI_AP = @MODI_AP,
                                MODI_PRID = @MODI_PRID,
                                FLAG = FLAG + 1,
                                TH020 = 'Y'
                                WHERE TH001 = @RoErpPrefix
                                AND TH002 = @RoErpNo";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            MODIFIER = UserNo,
                            MODI_DATE = dateNow,
                            MODI_TIME = timeNow,
                            MODI_AP = BaseHelper.ClientComputer(),
                            MODI_PRID = "BM",
                            RoErpPrefix,
                            RoErpNo
                        });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                    }
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //若成功則將相關資料寫回MES
                        #region //更新 - 銷貨單單頭
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.ReceiveOrder SET
                                ConfirmStatus = @ConfirmStatus,
                                ConfirmUserId = @ConfirmUserId,
                                ConfirmTime = @ConfirmTime,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RoId = @RoId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmStatus = "Y",
                                ConfirmUserId = LastModifiedBy,
                                ConfirmTime = LastModifiedDate,
                                LastModifiedDate,
                                LastModifiedBy,
                                RoId
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //更新 - 銷貨單單單身
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.RoDetail SET
                                ConfirmationCode = 'Y',
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RoId = @RoId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                LastModifiedDate,
                                LastModifiedBy,
                                RoId
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region /撈取銷貨單單身資料=>更新訂單,暫出單
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.[Type],a.FreeSpareQty,a.Quantity,a.AvailableQty,a.SoDetailId,a.TsnDetailId
                                FROM SCM.RoDetail a
                                WHERE a.RoId = @RoId";
                        dynamicParameters.Add("RoId", RoId);

                        var resultRoDetail = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in resultRoDetail)
                        {
                            #region //更新 - 訂單 已交數量,贈品已交數量,備品已交量
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.SoDetail SET
                                    SiQty = SiQty + @SiQty,
                                    FreebieSiQty = FreebieSiQty + @FreebieSiQty,
                                    SpareSiQty = SpareSiQty + @SpareSiQty,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE SoDetailId = @SoDetailId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    SiQty = item.Quantity,
                                    FreebieSiQty = item.Type == "1"? item.FreeSpareQty : 0,
                                    SpareSiQty = item.Type == "2" ? item.FreeSpareQty : 0,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    SoDetailId = item.SoDetailId
                                });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //更新 - 暫出單 轉進銷量,轉銷贈/備品量
                            if(item.TsnDetailId > 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE SCM.TsnDetail SET
                                        SaleQty = SaleQty + @SaleQty,
                                        SaleFSQty = SaleFSQty + @SaleFSQty,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE TsnDetailId = @TsnDetailId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        SaleQty = item.Quantity,
                                        SaleFSQty = item.FreeSpareQty,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        TsnDetailId = item.TsnDetailId
                                    });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            }
                            #endregion

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

        #region //UpdateRoReconfirm -- MES銷貨單反確認單據 -- Shintokuro 2024-04-24
        public string UpdateRoReconfirm(int RoId)
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
                int CompanyIdBase = -1;
                int SoDetailId = -1;

                string RoErpPrefix = "";
                string RoErpNo = "";
                string ReceiveDate = "";
                string SourceType = "";
                string PriceSourceTypeMain = "";

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
                        dynamicParameters.Add("FunctionCode", "ReceiveOrderManagment");
                        dynamicParameters.Add("DetailCode", "reconfirm");

                        var resultUser = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUser.Count() <= 0) throw new SystemException("使用者資料錯誤!");

                        foreach (var item in resultUser)
                        {
                            UserNo = item.UserNo;
                        }
                        #endregion

                        #region /撈取單頭資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.RoErpPrefix,a.RoErpNo, a.ConfirmStatus,a.CompanyId
                                ,FORMAT(a.ReceiveDate, 'yyyyMMdd') ReceiveDate
                                ,a.SourceType,a.PriceSourceTypeMain
                                ,b.SoDetailId
                                FROM SCM.ReceiveOrder a
                                INNER JOIN SCM.RoDetail b on a.RoId = b.RoId
                                WHERE a.RoId = @RoId";
                        dynamicParameters.Add("RoId", RoId);

                        var resultRoMes = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in resultRoMes)
                        {
                            if (item.ConfirmStatus == "N") throw new SystemException("該銷貨單單據未核單,不能反確認");
                            if (item.ConfirmStatus == "V") throw new SystemException("該銷貨單單據已經作廢,不能反確認");
                            if (item.CompanyId != CurrentCompany) throw new SystemException("該銷貨單單據公司別與後端公司別不一致,不能確認");
                            CompanyIdBase = item.CompanyId;
                            RoErpPrefix = item.RoErpPrefix;
                            RoErpNo = item.RoErpNo;
                            ReceiveDate = item.ReceiveDate;
                            SoDetailId = item.SoDetailId;
                            SourceType = item.SourceType;
                            PriceSourceTypeMain = item.PriceSourceTypeMain;
                        }
                        #endregion
                    }
                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //審核ERP權限
                        USR_GROUP = BaseHelper.CheckErpAuthority(UserNo, sqlConnection, "COPI08", "CONFIRM");
                        #endregion

                        #region //判斷開立日期是否超過結帳日
                        string CloseDateBase = "";
                        string ReceiveDateBase = ReceiveDate;
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
                                throw new SystemException("【銷貨單】ERP已經結帳,不可以在異動【" + CloseDateBase + "】之前的單據");
                            }
                        }
                        else
                        {
                            throw new SystemException("日期字符串无效，无法比较");
                        }
                        #endregion

                        #region //取出銷貨單單據資料 + 異動更新相關資料表
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.TG003,a.TG023 ,a.TG007,
                                b.TH003,b.TH004,b.TH031,
                                b.TH008,b.TH007,b.TH017,b.TH024,b.TH061,
                                b.TH014,b.TH015,b.TH016,
                                b.TH032,b.TH033,b.TH034,b.TH060,
                                b.TH012,b.TH037,
                                c.MB022,c.MB004,
                                ISNULL(md1.RoRate,1) RoRate, 
                                ISNULL(md2.RoAvailableRate,1) RoAvailableRate, 
                                (b.TH008*ISNULL(md1.RoRate,1)) ConversionQty,
                                (b.TH024*ISNULL(md1.RoRate,1)) ConversionFreeSpareQty,
                                (b.TH061*ISNULL(md2.RoAvailableRate,1)) ConversionAvailableQty,
                                ROUND((b.TH037 / (b.TH061*ISNULL(md2.RoAvailableRate,1))),CAST(md3.MF005 AS INT)) UnitCost
                                FROM COPTG a
                                INNER JOIN COPTH b on a.TG001 = b.TH001 AND a.TG002 = b.TH002
                                INNER JOIN INVMB c on b.TH004 = c.MB001
                                OUTER APPLY(
                                    SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) RoRate
                                    FROM INVMD x1
                                    WHERE x1.MD001 = b.TH004
                                    AND  x1.MD002 = b.TH009
                                ) md1
                                OUTER APPLY(
                                    SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) RoAvailableRate
                                    FROM INVMD x1
                                    WHERE x1.MD001 = b.TH004
                                    AND  x1.MD002 = b.TH062
                                ) md2
                                OUTER APPLY(
                                    SELECT x1.MF003,x1.MF004,x1.MF005,x1.MF006
                                    FROM CMSMF x1
                                    INNER JOIN CMSMA x2 on x1.MF001 = x2.MA003 
                                ) md3
                                WHERE a.TG001 = @RoErpPrefix
                                AND a.TG002 = @RoErpNo";
                        dynamicParameters.Add("RoErpPrefix", RoErpPrefix);
                        dynamicParameters.Add("RoErpNo", RoErpNo);
                        var resultRoErp = sqlConnection.Query(sql, dynamicParameters);
                        if (resultRoErp.Count() <= 0) throw new SystemException("找不到ERP銷貨單單據!");
                        foreach (var item in resultRoErp)
                        {
                            if (item.TG023 == "N") throw new SystemException("該銷貨單單據未核單,不能反確認");
                            if (item.TG023 == "V") throw new SystemException("該銷貨單單據已經作廢,不能反確認");

                            #region //欄位參數撈取
                            string ReceiveDateErp = item.TG003; //TG003 單據日
                            string ConfirmStatus = item.TG023; //TG023 單頭確認
                            string CustomerFullName = item.TG007; //TG007 客戶全名
                            string RoSequence = item.TH003; //TH003 序號
                            string MtlItemNo = item.TH004; //TH004 品號
                            string TypeErp = item.TH031; //TH031 類型(1.贈品/2.備品)
                            string InventoryNo = item.TH007; //TH007 庫別
                            string LotNumber = item.TH017; //TH017 批號
                            string Location = item.TH060; //TH060 儲位
                            double Quantity = Convert.ToDouble(item.TH008); //TH008 數量
                            double FreeSpareQty = Convert.ToDouble(item.TH024); //TH024 贈/備品數量
                            double AvailableQty = Convert.ToDouble(item.TH061); //TH061 計價數量
                            string TH014 = item.TH014; //TH014 訂單單別　/ TH015 訂單單號 / TH016 訂單序號
                            string TH015 = item.TH015;
                            string TH016 = item.TH016;
                            string SoFull = item.TH014 + '-' + item.TH015 + '-' + item.TH016;
                            string TH032 = item.TH032; //TH032 暫出單別 / TH033 暫出單號 / TH034 暫出序號
                            string TH033 = item.TH033;
                            string TH034 = item.TH034;
                            string TsnFull = item.TH032 + '-' + item.TH033 + '-' + item.TH034;
                            string LotManagement = item.MB022; //品號批號管理
                            string InventoryUomNo = item.MB004; //庫存單位

                            string AllOverMistake = "";
                            string OverMtlItemNo = "";
                            string OverMistake = "";
                            string ClosureStatusCOPTD = "N";
                            string ClosureStatusINVTG = "N";
                            string TD049 = "";
                            double TD008 = 0;
                            double TD024 = 0;
                            double TD050 = 0;
                            double TG009 = 0;
                            double TG044 = 0;
                            double TG052 = 0;

                            double CanUseQty = 0;
                            double CanUseFreebieQty = 0;
                            double CanUseSpareQty = 0;
                            double CanUseFreeSpareQty = 0;
                            double CanUseAvailableQty = 0;

                            double RoRate = Convert.ToDouble(item.RoRate); //單位換算率
                            double RoAvailableRate = Convert.ToDouble(item.RoAvailableRate); //計價單位換算率
                            double ConversionQty = Convert.ToDouble(item.ConversionQty); //單位換算後數量
                            double ConversionFreeSpareQty = Convert.ToDouble(item.ConversionFreeSpareQty); //單位換算後贈備品數量
                            double ConversionAvailableQty = Convert.ToDouble(item.ConversionAvailableQty); //單位換算後計價數量
                            double UnitCost = Convert.ToDouble(item.UnitCost); //單位成本
                            double PreTaxAmt = Convert.ToDouble(item.TH037); //本幣未稅金額
                            double UnitPrice = Convert.ToDouble(item.TH012); //原幣單價

                            double SoRate = 0;
                            double SoAvailableRate = 0;
                            double TsnRate = 0;
                            double TsnAvailableRate = 0;
                            #endregion

                            #region //檢查段
                            #region //訂單部分
                            #region //撈取訂單數量資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 TD049,a.TD016,TD010,TD077,TD008,TD009,TD076,TD078,
                                    ISNULL(a.TD008,0) TD008,
                                    ISNULL(a.TD024,0) TD024,
                                    ISNULL(a.TD050,0) TD050,
                                    (ISNULL(a.TD008,0) * ISNULL(md1.SoRate,1)) - (ISNULL(a.TD009,0) * ISNULL(md1.SoRate,1)) CanUseQty,
                                    (ISNULL(a.TD024,0) * ISNULL(md1.SoRate,1)) - (ISNULL(a.TD025,0) * ISNULL(md1.SoRate,1)) CanUseFreebieQty,
                                    (ISNULL(a.TD050,0) * ISNULL(md1.SoRate,1)) - (ISNULL(a.TD051,0) * ISNULL(md1.SoRate,1)) CanUseSpareQty,
                                    (ISNULL(a.TD076,0) * ISNULL(md2.SoAvailableRate,1)) - (ISNULL(a.TD078,0) * ISNULL(md2.SoAvailableRate,1)) CanUseAvailableQty,
                                    ISNULL(md1.SoRate,1) SoRate, 
                                    ISNULL(md2.SoAvailableRate,1) SoAvailableRate, 
                                    (ISNULL(a.TD008,0) * ISNULL(md1.SoRate,1)) SoConversionQty,
                                    (ISNULL(a.TD009,0) * ISNULL(md1.SoRate,1)) SoConversionSiQty,
                                    (ISNULL(a.TD024,0) * ISNULL(md1.SoRate,1)) SoConversionFreebieQty,
                                    (ISNULL(a.TD025,0) * ISNULL(md1.SoRate,1)) SoConversionFreebieSiQty,
                                    (ISNULL(a.TD050,0) * ISNULL(md1.SoRate,1)) SoConversionSpareQty,
                                    (ISNULL(a.TD051,0) * ISNULL(md1.SoRate,1)) SoConversionSpareSiQty,
                                    (ISNULL(a.TD076,0) * ISNULL(md2.SoAvailableRate,1)) SoConversionSoPriceQty,
                                    (ISNULL(a.TD078,0) * ISNULL(md2.SoAvailableRate,1)) SoConversionSoPriceSiQty
                                    FROM COPTD a
                                    OUTER APPLY(
                                        SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) SoRate
                                        FROM INVMD x1
                                        WHERE x1.MD001= a.TD004
                                        AND  x1.MD002 = a.TD010
                                    ) md1
                                    OUTER APPLY(
                                        SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) SoAvailableRate
                                        FROM INVMD x1
                                        WHERE x1.MD001= a.TD004
                                        AND  x1.MD002 = a.TD077
                                    ) md2
                                    WHERE a.TD001 = @TH014
                                    AND a.TD002 = @TH015
                                    AND a.TD003 = @TH016
                                    AND a.TD021 = 'Y'";
                            dynamicParameters.Add("TH014", TH014);
                            dynamicParameters.Add("TH015", TH015);
                            dynamicParameters.Add("TH016", TH016);
                            var resultCOPTD = sqlConnection.Query(sql, dynamicParameters);
                            if (resultCOPTD.Count() <= 0) throw new SystemException("訂單找不到,請重新確認");
                            foreach (var item1 in resultCOPTD)
                            {
                                TD049 = item1.TD049;
                                TD008 = Convert.ToDouble(item1.TD008);
                                TD024 = Convert.ToDouble(item1.TD024);
                                TD050 = Convert.ToDouble(item1.TD050);
                                CanUseQty = Convert.ToDouble(item1.CanUseQty); //可使用訂單數量
                                CanUseFreebieQty = Convert.ToDouble(item1.CanUseFreebieQty); //可使用贈品數量
                                CanUseSpareQty = Convert.ToDouble(item1.CanUseSpareQty); //可使用備品數量
                                CanUseAvailableQty = Convert.ToDouble(item1.CanUseAvailableQty); //可使用計價數量
                                SoRate = Convert.ToDouble(item1.SoRate);//數量單位轉換成庫存單位的換算率
                                SoAvailableRate = Convert.ToDouble(item1.SoAvailableRate);//計價單位轉換成庫存單位的換算率
                                double SoConversionQty = Convert.ToDouble(item1.SoConversionQty); //轉換成庫存單位的訂單數量
                                double SoConversionSiQty = Convert.ToDouble(item1.SoConversionSiQty);//轉換成庫存單位的已交訂單數量
                                double SoConversionFreebieQty = Convert.ToDouble(item1.SoConversionFreebieQty);//轉換成庫存單位的贈品數量
                                double SoConversionFreebieSiQty = Convert.ToDouble(item1.SoConversionFreebieSiQty);//轉換成庫存單位的已交贈品數量
                                double SoConversionSpareQty = Convert.ToDouble(item1.SoConversionSpareQty);//轉換成庫存單位的備品數量
                                double SoConversionSpareSiQty = Convert.ToDouble(item1.SoConversionSpareSiQty);//轉換成庫存單位的已交備品數量
                                double SoConversionSoPriceQty = Convert.ToDouble(item1.SoConversionSoPriceQty);//轉換成庫存單位的計價數量
                                double SoConversionSoPriceSiQty = Convert.ToDouble(item1.SoConversionSoPriceSiQty);//轉換成庫存單位的已交計價數量
                            }
                            #endregion

                            #endregion

                            #region //暫出單部分
                            if (SourceType == "To" || TH032 != "")
                            {
                                CanUseQty = 0;
                                CanUseFreeSpareQty = 0;
                                CanUseAvailableQty = 0;
                                AllOverMistake = "";
                                OverMtlItemNo = "";
                                OverMistake = "";
                                #region //撈取暫出單數量資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.TG024,ISNULL(a.TG009,0) TG009,ISNULL(a.TG044,0) TG044,ISNULL(a.TG052,0) TG052,
                                        ((ISNULL(a.TG009,0) * ISNULL(md1.TsnRate,1)) - (ISNULL(a.TG020,0) * ISNULL(md1.TsnRate,1)) - (ISNULL(a.TG021,0) * ISNULL(md1.TsnRate,1))) CanUseQty,
                                        ((ISNULL(a.TG044,0) * ISNULL(md1.TsnRate,1)) - (ISNULL(a.TG046,0) * ISNULL(md1.TsnRate,1)) - (ISNULL(a.TG048,0) * ISNULL(md1.TsnRate,1))) CanUseFreeSpareQty,
                                        ((ISNULL(a.TG052,0) * ISNULL(md2.TsnAvailableRate,1)) - (ISNULL(a.TG054,0) * ISNULL(md2.TsnAvailableRate,1)) - (ISNULL(a.TG055,0) * ISNULL(md2.TsnAvailableRate,1))) CanUseAvailableQty,
                                        ISNULL(md1.TsnRate,1) TsnRate, 
                                        ISNULL(md2.TsnAvailableRate,1) TsnAvailableRate,
                                        (ISNULL(a.TG009,0) * ISNULL(md1.TsnRate,1)) TsnConversionQty,
                                        (ISNULL(a.TG044,0) * ISNULL(md1.TsnRate,1)) TsnConversionFreeSpareQty,
                                        (ISNULL(a.TG052,0) * ISNULL(md2.TsnAvailableRate,1)) TsnConversionAvailableQty
                                        FROM INVTG a
                                        OUTER APPLY(
                                            SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) TsnRate
                                            FROM INVMD x1
                                            WHERE x1.MD001= a.TG004
                                            AND  x1.MD002 = a.TG010
                                        ) md1
                                        OUTER APPLY(
                                            SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) TsnAvailableRate
                                            FROM INVMD x1
                                            WHERE x1.MD001= a.TG004
                                            AND  x1.MD002 = a.TG053
                                        ) md2
                                        WHERE a.TG001 = @TG001
                                        AND a.TG002 = @TG002
                                        AND a.TG003 = @TG003
                                        AND a.TG022 ='Y'
                                        ";
                                dynamicParameters.Add("TG001", TH032);
                                dynamicParameters.Add("TG002", TH033);
                                dynamicParameters.Add("TG003", TH034);
                                var resultINVTG = sqlConnection.Query(sql, dynamicParameters);
                                if (resultINVTG.Count() <= 0) throw new SystemException("暫出單找不到,請重新確認");
                                foreach (var item1 in resultINVTG)
                                {
                                    TG009 = Convert.ToDouble(item1.TG009);
                                    TG044 = Convert.ToDouble(item1.TG044);
                                    TG052 = Convert.ToDouble(item1.TG052);
                                    CanUseQty = Convert.ToDouble(item1.CanUseQty);
                                    CanUseFreeSpareQty = Convert.ToDouble(item1.CanUseFreeSpareQty);
                                    CanUseAvailableQty = Convert.ToDouble(item1.CanUseAvailableQty);
                                    TsnRate = Convert.ToDouble(item1.TsnRate);
                                    TsnAvailableRate = Convert.ToDouble(item1.TsnAvailableRate);
                                    double TsnConversionQty = Convert.ToDouble(item1.TsnConversionQty);
                                    double TsnConversionFreeSpareQty = Convert.ToDouble(item1.TsnConversionFreeSpareQty);
                                    double TsnConversionAvailableQty = Convert.ToDouble(item1.TsnConversionAvailableQty);
                                }
                                #endregion
                            }
                            #endregion
                            #endregion

                            #region //異動更新段-銷貨單確認後
                            #region //COPTD 銷貨單確認後回寫訂單單身 已交數量
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE COPTD SET
                                    MODIFIER = @MODIFIER,
                                    MODI_DATE = @MODI_DATE,
                                    MODI_TIME = @MODI_TIME,
                                    MODI_AP = @MODI_AP,
                                    MODI_PRID = @MODI_PRID,
                                    FLAG = FLAG + 1,
                                    TD009  = TD009 - @TD009,
                                    TD025  = TD025 - @TD025,
                                    TD051  = TD051 - @TD051,
                                    TD078  = TD078 - @TD078,
                                    TD016  = @TD016 
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
                                TD009 = Math.Round(ConversionQty / SoRate, 3, MidpointRounding.AwayFromZero),
                                TD025 = TypeErp == "1" ? Math.Round(ConversionFreeSpareQty / SoRate, 3, MidpointRounding.AwayFromZero) : 0,
                                TD051 = TypeErp == "2" ? Math.Round(ConversionFreeSpareQty / SoRate, 3, MidpointRounding.AwayFromZero) : 0,
                                TD078 = Math.Round(ConversionAvailableQty / SoAvailableRate, 3, MidpointRounding.AwayFromZero),
                                TD016 = "N",
                                TD001 = TH014,
                                TD002 = TH015,
                                TD003 = TH016
                            });
                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //銷貨單來源為暫出單觸發
                            if (SourceType == "To" || TH032 != "")
                            {
                                #region //INVTF 暫出入轉撥單頭檔更新 轉銷日更新
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE INVTF SET
                                        MODIFIER = @MODIFIER,
                                        MODI_DATE = @MODI_DATE,
                                        MODI_TIME = @MODI_TIME,
                                        MODI_AP = @MODI_AP,
                                        MODI_PRID = @MODI_PRID,
                                        FLAG = FLAG + 1,
                                        TF042  = @TF042
                                        WHERE TF001 = @TF001
                                        AND TF002 = @TF002";
                                dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MODIFIER = UserNo,
                                    MODI_DATE = dateNow,
                                    MODI_TIME = timeNow,
                                    MODI_AP = BaseHelper.ClientComputer(),
                                    MODI_PRID = "BM",
                                    TF042 = ReceiveDateErp,
                                    TF001 = TH032,
                                    TF002 = TH033
                                });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region //INVTG 暫出入轉撥單身檔 轉銷數量
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE INVTG SET
                                        MODIFIER = @MODIFIER,
                                        MODI_DATE = @MODI_DATE,
                                        MODI_TIME = @MODI_TIME,
                                        MODI_AP = @MODI_AP,
                                        MODI_PRID = @MODI_PRID,
                                        FLAG = FLAG + 1,
                                        TG020  = TG020 - @TG020,
                                        TG046  = TG046 - @TG046,
                                        TG054  = TG054 - @TG054,
                                        TG024  = @TG024 
                                        WHERE TG001 = @TG001
                                        AND TG002 = @TG002
                                        AND TG003 = @TG003";
                                dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MODIFIER = UserNo,
                                    MODI_DATE = dateNow,
                                    MODI_TIME = timeNow,
                                    MODI_AP = BaseHelper.ClientComputer(),
                                    MODI_PRID = "BM",
                                    TG020 = Math.Round(ConversionQty / TsnRate, 3, MidpointRounding.AwayFromZero),
                                    TG046 = Math.Round(ConversionFreeSpareQty / TsnRate, 3, MidpointRounding.AwayFromZero),
                                    TG054 = Math.Round(ConversionAvailableQty / TsnAvailableRate, 3, MidpointRounding.AwayFromZero),
                                    TG001 = TH032,
                                    TG002 = TH033,
                                    TG003 = TH034,
                                    TG024 = "N"
                                });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }
                            #endregion

                            #region //INVLA 交易紀錄刪除
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
                                LA005 = -1,
                                LA006 = RoErpPrefix,
                                LA007 = RoErpNo,
                                LA008 = RoSequence
                            });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //INVMB 品號基本資料檔 品號庫存更新
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE INVMB SET
                                    MODIFIER = @MODIFIER,
                                    MODI_DATE = @MODI_DATE,
                                    MODI_TIME = @MODI_TIME,
                                    MODI_AP = @MODI_AP,
                                    MODI_PRID = @MODI_PRID,
                                    FLAG = FLAG + 1,
                                    MB064  = MB064 + @MB064
                                    WHERE MB001 = @MB001";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                MODIFIER = UserNo,
                                MODI_DATE = dateNow,
                                MODI_TIME = timeNow,
                                MODI_AP = BaseHelper.ClientComputer(),
                                MODI_PRID = "BM",
                                MB064 = ConversionQty + ConversionFreeSpareQty,
                                MB001 = MtlItemNo
                            });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //INVMC 品號庫別檔 品號庫別庫存更新
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE INVMC SET
                                    MODIFIER = @MODIFIER,
                                    MODI_DATE = @MODI_DATE,
                                    MODI_TIME = @MODI_TIME,
                                    MODI_AP = @MODI_AP,
                                    MODI_PRID = @MODI_PRID,
                                    FLAG = FLAG + 1,
                                    MC007  = MC007 + @MC007,
                                    MC013  = @MC013 
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
                                MC007 = ConversionQty + ConversionFreeSpareQty,
                                MC013 = ReceiveDateErp,
                                MC001 = MtlItemNo,
                                MC002 = InventoryNo
                            });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //品號有批號設定觸發
                            if (LotManagement != "N")
                            {
                                #region //INVMF 品號批號資料單身 資料刪除
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
                                    MF004 = RoErpPrefix,
                                    MF005 = RoErpNo,
                                    MF006 = RoSequence
                                });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region //INVLF 儲位批號異動明細資料檔 資料刪除
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE INVLF
                                        WHERE LF001 = @LF001
                                          AND LF002 = @LF002
                                          AND LF003 = @LF003
                                          AND LF004 = @LF004
                                          AND LF005 = @LF005
                                          AND LF006 = @LF006
                                          AND LF007 = @LF007";
                                dynamicParameters.AddDynamicParams(
                                new
                                {
                                    LF001 = RoErpPrefix,
                                    LF002 = RoErpNo,
                                    LF003 = RoSequence,
                                    LF004 = MtlItemNo,
                                    LF005 = InventoryNo,
                                    LF006 = Location != "" ? Location : "##########",
                                    LF007 = LotNumber
                                });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region //INVMM 品號庫別儲位批號檔 庫存數量更新
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE INVMM SET
                                        MODIFIER = @MODIFIER,
                                        MODI_DATE = @MODI_DATE,
                                        MODI_TIME = @MODI_TIME,
                                        MODI_AP = @MODI_AP,
                                        MODI_PRID = @MODI_PRID,
                                        FLAG = FLAG + 1,
                                        MM005  = MM005 + @MM005,
                                        MM009  = @MM009
                                        WHERE MM001 = @MM001
                                        AND MM002 = @MM002
                                        AND MM003 = @MM003";
                                dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MODIFIER = UserNo,
                                    MODI_DATE = dateNow,
                                    MODI_TIME = timeNow,
                                    MODI_AP = BaseHelper.ClientComputer(),
                                    MODI_PRID = "BM",
                                    MM005 = ConversionQty+ ConversionFreeSpareQty,
                                    MM009 = ReceiveDateErp,
                                    MM001 = MtlItemNo,
                                    MM002 = InventoryNo,
                                    MM003 = Location != "" ? Location : "##########"
                                });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }
                            #endregion

                            #endregion

                        }
                        #endregion

                        #region //銷貨單單據確認
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE COPTG SET 
                                MODIFIER = @MODIFIER,
                                MODI_DATE = @MODI_DATE,
                                MODI_TIME = @MODI_TIME,
                                MODI_AP = @MODI_AP,
                                MODI_PRID = @MODI_PRID,
                                FLAG = FLAG + 1,
                                TG023 = 'N'
                                WHERE TG001 = @RoErpPrefix
                                AND TG002 = @RoErpNo";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            MODIFIER = UserNo,
                            MODI_DATE = dateNow,
                            MODI_TIME = timeNow,
                            MODI_AP = BaseHelper.ClientComputer(),
                            MODI_PRID = "BM",
                            RoErpPrefix,
                            RoErpNo
                        });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE COPTH SET 
                                MODIFIER = @MODIFIER,
                                MODI_DATE = @MODI_DATE,
                                MODI_TIME = @MODI_TIME,
                                MODI_AP = @MODI_AP,
                                MODI_PRID = @MODI_PRID,
                                FLAG = FLAG + 1,
                                TH020 = 'N'
                                WHERE TH001 = @RoErpPrefix
                                AND TH002 = @RoErpNo";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            MODIFIER = UserNo,
                            MODI_DATE = dateNow,
                            MODI_TIME = timeNow,
                            MODI_AP = BaseHelper.ClientComputer(),
                            MODI_PRID = "BM",
                            RoErpPrefix,
                            RoErpNo
                        });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                    }
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //若成功則將相關資料寫回MES
                        #region //更新 - 銷貨單單頭
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.ReceiveOrder SET
                                ConfirmStatus = @ConfirmStatus,
                                ConfirmUserId = @ConfirmUserId,
                                ConfirmTime = @ConfirmTime,
                                TransferStatusMES = @TransferStatusMES,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RoId = @RoId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmStatus = "N",
                                ConfirmUserId = LastModifiedBy,
                                ConfirmTime = LastModifiedDate,
                                TransferStatusMES = "R",
                                LastModifiedDate,
                                LastModifiedBy,
                                RoId
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //更新 - 銷貨單單單身
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.RoDetail SET
                                ConfirmationCode = 'N',
                                TransferStatusMES = @TransferStatusMES,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RoId = @RoId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TransferStatusMES = "N",
                                LastModifiedDate,
                                LastModifiedBy,
                                RoId
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region /撈取銷貨單單身資料=>更新訂單,暫出單
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.[Type],a.FreeSpareQty,a.Quantity,a.AvailableQty,a.SoDetailId,a.TsnDetailId
                                FROM SCM.RoDetail a
                                WHERE a.RoId = @RoId";
                        dynamicParameters.Add("RoId", RoId);

                        var resultRoDetail = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in resultRoDetail)
                        {
                            #region //更新 - 訂單 已交數量,贈品已交數量,備品已交量
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.SoDetail SET
                                    SiQty = SiQty - @SiQty,
                                    FreebieSiQty = FreebieSiQty - @FreebieSiQty,
                                    SpareSiQty = SpareSiQty - @SpareSiQty,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE SoDetailId = @SoDetailId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    SiQty = item.Quantity,
                                    FreebieSiQty = item.Type == "1" ? item.FreeSpareQty : 0,
                                    SpareSiQty = item.Type == "2" ? item.FreeSpareQty : 0,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    SoDetailId = item.SoDetailId
                                });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //更新 - 暫出單 轉進銷量,轉銷贈/備品量
                            if (item.TsnDetailId > 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE SCM.TsnDetail SET
                                        SaleQty = SaleQty - @SaleQty,
                                        SaleFSQty = SaleFSQty - @SaleFSQty,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE TsnDetailId = @TsnDetailId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        SaleQty = item.Quantity,
                                        SaleFSQty = item.FreeSpareQty,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        TsnDetailId = item.TsnDetailId
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
                        #endregion
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

        #region //UpdateReceiveOrderVoid -- 銷貨單作廢 -- Shintokuro 2024-03-04
        public string UpdateReceiveOrderVoid(int RoId)
        {
            try
            {
                if (RoId < 0) throw new SystemException("銷貨單不可為空,請重新確認!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    string dateNow = CreateDate.ToString("yyyyMMdd"), timeNow = CreateDate.ToString("HH:mm:ss"), errorMsg = "";
                    string MESCompanyNo = "", departmentNo = "", userNo = "", userName = "", userGroup = "";
                    string RoErpPrefix = "", RoErpNo = "", TransferStatusMES = "", ConfirmStatus = "";
                    int OrderCompanyIdBase = -1;

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
                            MESCompanyNo = item.ErpNo;
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
                        dynamicParameters.Add("FunctionCode", "ReceiveOrderManagment");
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

                        #region //確認銷貨單是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.TransferStatusMES, a.ConfirmStatus, a.CompanyId, a.RoErpPrefix, a.RoErpNo
                                FROM SCM.ReceiveOrder a
                                WHERE a.CompanyId = @CompanyId
                                AND a.RoId = @RoId
                                ";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("RoId", RoId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("銷貨單單頭資料找不到，請重新確認");
                        foreach (var item in result)
                        {
                            TransferStatusMES = item.TransferStatusMES;
                            ConfirmStatus = item.ConfirmStatus;
                            OrderCompanyIdBase = item.CompanyId;
                            RoErpPrefix = item.RoErpPrefix;
                            RoErpNo = item.RoErpNo;
                        }
                        if (TransferStatusMES == "N") throw new SystemException("銷貨單處於未拋轉狀態,無法使用作廢功能");
                        if (ConfirmStatus == "Y") throw new SystemException("銷貨單處於確認狀態,如欲修改銷貨單資料請使用反確認功能");
                        if (ConfirmStatus == "V") throw new SystemException("銷貨單處於作廢狀態不可以異動");
                        if (OrderCompanyIdBase != CurrentCompany) throw new SystemException("單據的公司別與目前公司別不同，請嘗試重新登入!!");
                        #endregion

                        #region //更新單頭
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.ReceiveOrder SET
                                ConfirmStatus = @ConfirmStatus,
                                TransferStatusMES = @TransferStatusMES,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RoId = @RoId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmStatus = "V",
                                TransferStatusMES = "Y",
                                LastModifiedDate,
                                LastModifiedBy,
                                RoId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //更新單身
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.RoDetail SET
                                ConfirmationCode = @ConfirmationCode,
                                TransferStatusMES = @TransferStatusMES,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RoId = @RoId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmationCode = "V",
                                TransferStatusMES = "Y",
                                LastModifiedDate,
                                LastModifiedBy,
                                RoId
                            });
                        insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //ERP公司代碼取得
                        string ERPCompanyNo = "";
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 ML001  
                                         FROM CMSML";
                        var resultCompanyNo = sqlConnection.Query(sql, dynamicParameters);
                        foreach (var item in resultCompanyNo)
                        {
                            ERPCompanyNo = item.ML001;
                        }
                        if (MESCompanyNo != ERPCompanyNo) throw new SystemException("ERP的公司代碼與MES的公司代碼不同，請嘗試重新登入!!");
                        #endregion

                        #region //判斷ERP使用者是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 MF005
                                        FROM ADMMF
                                        WHERE COMPANY = @CompanyNo
                                        AND MF001 = @userNo";
                        dynamicParameters.Add("CompanyNo", ERPCompanyNo);
                        dynamicParameters.Add("userNo", userNo);
                        var resultUserExist = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUserExist.Count() <= 0) throw new SystemException("【ERP使用者】不存在!");
                        #endregion

                        #region //審核ERP權限-建立
                        var USR_GROUP = BaseHelper.CheckErpAuthority(userNo, sqlConnection, "COPI08", "CREATE");
                        #endregion

                        #region //判斷銷貨單是否存在+可不可以作廢
                        string erpConfirmStatus = "";
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TG023
                                FROM COPTG
                                WHERE TG001 =@RoErpPrefix
                                AND TG002 = @RoErpNo";
                        dynamicParameters.Add("RoErpPrefix", RoErpPrefix);
                        dynamicParameters.Add("RoErpNo", RoErpNo);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("銷貨單不存在,請重新確認!!!");

                        foreach (var item in result)
                        {
                            erpConfirmStatus = item.TG023;
                        }
                        if (erpConfirmStatus == "Y") throw new SystemException("ERP銷貨單處於確認狀態不可以異動!!!");
                        if (erpConfirmStatus == "V") throw new SystemException("ERP銷貨單處於作廢狀態不可以異動!!!");
                        #endregion

                        #region //更新單頭
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE COPTG SET
                                TG023 = 'V',
                                MODIFIER = @MODIFIER,
                                MODI_DATE = @MODI_DATE,
                                MODI_TIME = @MODI_TIME,
                                MODI_AP = @MODI_AP,
                                MODI_PRID = @MODI_PRID
                                WHERE TG001 =@RoErpPrefix
                                AND TG002 = @RoErpNo";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                MODIFIER = userNo,
                                MODI_DATE = dateNow,
                                MODI_TIME = timeNow,
                                MODI_AP = BaseHelper.ClientComputer(),
                                MODI_PRID = "BM",
                                RoErpPrefix,
                                RoErpNo
                            });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //更新單身
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE COPTH SET
                                TH020 = 'V',
                                MODIFIER = @MODIFIER,
                                MODI_DATE = @MODI_DATE,
                                MODI_TIME = @MODI_TIME,
                                MODI_AP = @MODI_AP,
                                MODI_PRID = @MODI_PRID
                                WHERE TH001 =@RoErpPrefix
                                AND TH002 = @RoErpNo";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                MODIFIER = userNo,
                                MODI_DATE = dateNow,
                                MODI_TIME = timeNow,
                                MODI_AP = BaseHelper.ClientComputer(),
                                MODI_PRID = "BM",
                                RoErpPrefix,
                                RoErpNo
                            });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                    }

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "單據作廢成功!!!"
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

        #region //UpdateReceiveOrderSynchronize -- 銷貨單資料手動同步 -- Shintokuro 2024.02.20
        public string UpdateReceiveOrderSynchronize(string RoErpFullNo, string SyncStartDate, string SyncEndDate
            , string NormalSync, string TranSync, string CompanyNo
            )
        {
            try
            {
                List<ReceiveOrder> receiveOrders = new List<ReceiveOrder>();
                List<RoDetail> roDetails = new List<RoDetail>();

                int mainAffected = 0, detailAffected = 0, mainDelAffected = 0, detailDelAffected = 0; ;
                string companyNoBase = "";
                //if (CompanyNo.Length <= 0) throw new SystemException("【公司】不能為空!");
                int CompanyId = -1;
                string ErpConnectionStrings = "";

                using (TransactionScope transactionScope = new TransactionScope(TransactionScopeOption.Required, new TimeSpan(0, 180, 0)))
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
                            CompanyId = Convert.ToInt32(item.CompanyId);
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        }
                        #endregion
                    }

                    #region //正常同步
                    if (NormalSync == "Y")
                    {
                        using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                        { 
                            #region //判斷ERP銷貨單資料是否存在
                            if (RoErpFullNo.Length > 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM COPTG
                                        WHERE (LTRIM(RTRIM(TG001)) + '-' + LTRIM(RTRIM(TG002))) LIKE '%' + @RoErpFullNo + '%'";
                                dynamicParameters.Add("RoErpFullNo", RoErpFullNo);

                                var resultTsn = sqlConnection.Query(sql, dynamicParameters);
                                if (resultTsn.Count() <= 0) throw new SystemException("ERP銷貨單資料錯誤");
                            }
                            #endregion

                            #region //撈取ERP銷貨單頭資料 (欄位尚未編寫)
                            sql = @"SELECT LTRIM(RTRIM(TG001)) RoErpPrefix       , LTRIM(RTRIM(TG002)) RoErpNo
                                    , LTRIM(RTRIM(TG004)) CustomerNo             , LTRIM(RTRIM(TG005)) DepartmentNo
                                    , LTRIM(RTRIM(TG006)) SalesmenNo             , LTRIM(RTRIM(TG007)) CustomerFullName
                                    , LTRIM(RTRIM(TG008)) CustomerAddressFirst   , LTRIM(RTRIM(TG009)) CustomerAddressSecond
                                    , LTRIM(RTRIM(TG010)) Factory                , LTRIM(RTRIM(TG011)) Currency
                                    , LTRIM(RTRIM(TG012)) ExchangeRate           , LTRIM(RTRIM(TG013)) OriginalAmount
                                    , LTRIM(RTRIM(TG014)) InvNumEnd              , LTRIM(RTRIM(TG015)) UiNo
                                    , LTRIM(RTRIM(TG016)) InvoiceType            , LTRIM(RTRIM(TG017)) TaxType
                                    , LTRIM(RTRIM(TG018)) InvoiceAddressFirst    , LTRIM(RTRIM(TG019)) InvoiceAddressSecond
                                    , LTRIM(RTRIM(TG020)) Remark                 , LTRIM(RTRIM(TG022)) PrintCount
                                    , LTRIM(RTRIM(TG023)) ConfirmStatus          , LTRIM(RTRIM(TG024)) UpdateCode
                                    , LTRIM(RTRIM(TG025)) OriginalTaxAmount      , LTRIM(RTRIM(TG026)) CollectionSalesmenNo
                                    , LTRIM(RTRIM(TG027)) Remark1                , LTRIM(RTRIM(TG028)) Remark2
                                    , LTRIM(RTRIM(TG029)) Remark3                , LTRIM(RTRIM(TG030)) InvoicesVoid
                                    , LTRIM(RTRIM(TG031)) CustomsClearance       , LTRIM(RTRIM(TG032)) RowCnt
                                    , LTRIM(RTRIM(TG033)) TotalQuantity          , LTRIM(RTRIM(TG034)) CashSales
                                    , LTRIM(RTRIM(TG035)) StaffNo                , LTRIM(RTRIM(TG036)) RevenueJournalCode
                                    , LTRIM(RTRIM(TG037)) CostJournalCode        , LTRIM(RTRIM(TG038)) ApplyYYMM
                                    , LTRIM(RTRIM(TG039)) LCNO                   , LTRIM(RTRIM(TG040)) INVOICENO
                                    , LTRIM(RTRIM(TG041)) InvoicesPrintCount     , LTRIM(RTRIM(TG043)) ConfirmUserNo
                                    , LTRIM(RTRIM(TG044)) TaxRate                , LTRIM(RTRIM(TG045)) PretaxAmount
                                    , LTRIM(RTRIM(TG046)) TaxAmount              , LTRIM(RTRIM(TG047)) PaymentTerm
                                    , LTRIM(RTRIM(TG048)) SoErpPrefix            , LTRIM(RTRIM(TG049)) SoErpNo             , LTRIM(RTRIM(TG050)) AdvOrderPrefix
                                    , LTRIM(RTRIM(TG051)) AdvOrderNo             , LTRIM(RTRIM(TG052)) OffsetAmount
                                    , LTRIM(RTRIM(TG053)) OffsetTaxAmount        , LTRIM(RTRIM(TG054)) TotalPackages
                                    , LTRIM(RTRIM(TG055)) SignatureStatus        , LTRIM(RTRIM(TG056)) ChangeInvCode
                                    , LTRIM(RTRIM(TG057)) NewRoPrefix            , LTRIM(RTRIM(TG058)) NewRoNo
                                    , LTRIM(RTRIM(TG059)) TransferStatusERP      , LTRIM(RTRIM(TG060)) ProcessCode
                                    , LTRIM(RTRIM(TG061)) AttachInvWithShip      , LTRIM(RTRIM(TG062)) BondCode
                                    , LTRIM(RTRIM(TG063)) TransmissionCount      , LTRIM(RTRIM(TG064)) Invoicer
                                    , LTRIM(RTRIM(TG065)) InvCode                , LTRIM(RTRIM(TG066)) ContactPerson
                                    , LTRIM(RTRIM(TG067)) Courier                , LTRIM(RTRIM(TG068)) SiteCommCalcMethod
                                    , LTRIM(RTRIM(TG069)) SiteCommRate           , LTRIM(RTRIM(TG070)) CommCalcInclTax
                                    , LTRIM(RTRIM(TG071)) TotalCommAmount        , LTRIM(RTRIM(TG072)) TransportMethod
                                    , LTRIM(RTRIM(TG073)) DispatchOrderPrefix    , LTRIM(RTRIM(TG074)) DispatchOrderNo
                                    , LTRIM(RTRIM(TG075)) DeclarationNumber      , LTRIM(RTRIM(TG076)) FullNameOfDelivCustomer
                                    , LTRIM(RTRIM(TG077)) RoPriceType            , LTRIM(RTRIM(TG078)) TelephoneNumber
                                    , LTRIM(RTRIM(TG079)) FaxNumber              , LTRIM(RTRIM(TG080)) ShipNoticePrefix
                                    , LTRIM(RTRIM(TG081)) ShipNoticeNo           , LTRIM(RTRIM(TG082)) TradingTerms
                                    , LTRIM(RTRIM(TG083)) CustomerEgFullName     , LTRIM(RTRIM(TG086)) InvNumGenMethod
                                    , LTRIM(RTRIM(TG087)) DocSourceCode          , LTRIM(RTRIM(TG089)) NoCredLimitControl
                                    , LTRIM(RTRIM(TG090)) InstallmentSettlement  , LTRIM(RTRIM(TG091)) InstallmentCount
                                    , LTRIM(RTRIM(TG092)) AutoApportionByInstallment , LTRIM(RTRIM(TG093)) StartYearMonth
                                    , LTRIM(RTRIM(TG094)) TaxCode                , LTRIM(RTRIM(TG095)) CustomsManual
                                    , LTRIM(RTRIM(TG096)) RemarkCode             , LTRIM(RTRIM(TG097)) MultipleInvoices
                                    , LTRIM(RTRIM(TG098)) InvNumStart            , LTRIM(RTRIM(TG099)) NumberOfInvoices
                                    , LTRIM(RTRIM(TG100)) MultiTaxRate           , LTRIM(RTRIM(TG101)) TaxCurrencyRate
                                    , LTRIM(RTRIM(TG104)) VoidApprovalDocNum     , LTRIM(RTRIM(TG105)) VoidReason
                                    , LTRIM(RTRIM(TG106)) [Source]               , LTRIM(RTRIM(TG107)) IncomeDraftID
                                    , LTRIM(RTRIM(TG108)) IncomeDraftSeq         , LTRIM(RTRIM(TG109)) IncomeVoucherType
                                    , LTRIM(RTRIM(TG110)) IncomeVoucherNumber    , LTRIM(RTRIM(TG111)) [Status]
                                    , LTRIM(RTRIM(TG112)) ZeroTaxForBuyer        , LTRIM(RTRIM(TG116)) GenLedgerAcctType
                                    , LTRIM(RTRIM(TG117)) InvoiceTime            , LTRIM(RTRIM(TG118)) InvCode2
                                    , LTRIM(RTRIM(TG119)) InvSymbol              , LTRIM(RTRIM(TG120)) DeliveryCountry
                                    , LTRIM(RTRIM(TG124)) VehicleIDshow          , LTRIM(RTRIM(TG125)) VehicleTypeNumber
                                    , LTRIM(RTRIM(TG126)) VehicleIDhide          , LTRIM(RTRIM(TG127)) InvDonationRecipient
                                    , LTRIM(RTRIM(TG128)) InvRandomCode          , LTRIM(RTRIM(TG129)) ReservedField
                                    , LTRIM(RTRIM(TG130)) CreditCard4No          , LTRIM(RTRIM(TG131)) ContactEmail
                                    , LTRIM(RTRIM(TG132)) ExpectedReceiptDate    , LTRIM(RTRIM(TG133)) OrigInvNumber
                                    , CASE WHEN LEN(LTRIM(RTRIM(TG003))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TG003)) as date), 'yyyy-MM-dd') ELSE NULL END ReceiveDate
                                    , CASE WHEN LEN(LTRIM(RTRIM(TG021))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TG021)) as date), 'yyyy-MM-dd') ELSE NULL END InvoiceDate
                                    , CASE WHEN LEN(LTRIM(RTRIM(TG042))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TG042)) as date), 'yyyy-MM-dd') ELSE NULL END DocDate
                                    , CASE WHEN LEN(LTRIM(RTRIM(TG102))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TG102)) as date), 'yyyy-MM-dd') ELSE NULL END VoidDate
                                    , CASE WHEN LEN(LTRIM(RTRIM(CREATE_DATE))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_DATE)) as date), 'yyyy-MM-dd') ELSE NULL END TransferDate
                                    , CASE WHEN LEN(LTRIM(RTRIM(CREATE_TIME))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_TIME)) as datetime), 'HH:mm:ss') ELSE NULL END TransferTime
                                    FROM COPTG
                                    WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "RoErpFullNo", @" AND (LTRIM(RTRIM(TG001)) + '-' + LTRIM(RTRIM(TG002))) LIKE '%' + @RoErpFullNo + '%'", RoErpFullNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SyncStartDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME >= @SyncStartDate OR MODI_DATE + ' ' + MODI_TIME >= @SyncStartDate)", SyncStartDate.Length > 0 ? Convert.ToDateTime(SyncStartDate).ToString("yyyyMMdd 00:00:00") : "");
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SyncEndDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME <= @SyncEndDate OR MODI_DATE + ' ' + MODI_TIME <= @SyncEndDate)", SyncEndDate.Length > 0 ? Convert.ToDateTime(SyncEndDate).ToString("yyyyMMdd 23:59:59") : ""); sql += @" ORDER BY TransferDate, TransferTime";

                            receiveOrders = sqlConnection.Query<ReceiveOrder>(sql, dynamicParameters).ToList();
                            #endregion

                            #region //撈取ERP銷貨單身資料 (欄位尚未編寫)
                            sql = @"SELECT  LTRIM(RTRIM(TH001)) RoErpPrefix      , LTRIM(RTRIM(TH002)) RoErpNo
                                    , LTRIM(RTRIM(TH003)) RoSequence             , LTRIM(RTRIM(TH004)) MtlItemNo
                                    , LTRIM(RTRIM(TH007)) InventoryNo            , LTRIM(RTRIM(TH008)) Quantity
                                    , LTRIM(RTRIM(TH009)) UomNo                  , LTRIM(RTRIM(TH012)) UnitPrice
                                    , LTRIM(RTRIM(TH013)) Amount                 , LTRIM(RTRIM(TH014)) SoErpPrefix
                                    , LTRIM(RTRIM(TH015)) SoErpNo                , LTRIM(RTRIM(TH016)) SoSequence
                                    , LTRIM(RTRIM(TH017)) LotNumber              , LTRIM(RTRIM(TH018)) Remarks
                                    , LTRIM(RTRIM(TH019)) CustomerMtlItemNo      , LTRIM(RTRIM(TH020)) ConfirmationCode
                                    , LTRIM(RTRIM(TH021)) UpdateCode             , LTRIM(RTRIM(TH024)) FreeSpareQty
                                    , LTRIM(RTRIM(TH025)) DiscountRate           , LTRIM(RTRIM(TH026)) CheckOutCode
                                    , LTRIM(RTRIM(TH027)) CheckOutPrefix         , LTRIM(RTRIM(TH028)) CheckOutNo
                                    , LTRIM(RTRIM(TH029)) CheckOutSequence       , LTRIM(RTRIM(TH030)) ProjectCode
                                    , LTRIM(RTRIM(TH031)) [Type]                 , LTRIM(RTRIM(TH032)) TsnErpPrefix
                                    , LTRIM(RTRIM(TH033)) TsnErpNo               , LTRIM(RTRIM(TH034)) TsnSequence
                                    , LTRIM(RTRIM(TH035)) OrigPreTaxAmt          , LTRIM(RTRIM(TH036)) OrigTaxAmt
                                    , LTRIM(RTRIM(TH037)) PreTaxAmt              , LTRIM(RTRIM(TH038)) TaxAmt
                                    , LTRIM(RTRIM(TH039)) PackageQty             , LTRIM(RTRIM(TH040)) PackageFreeSpareQty
                                    , LTRIM(RTRIM(TH041)) PackageUomNo           , LTRIM(RTRIM(TH042)) BondedCode
                                    , LTRIM(RTRIM(TH043)) SrQty                  , LTRIM(RTRIM(TH044)) SrPackageQty
                                    , LTRIM(RTRIM(TH045)) SrOrigPreTaxAmt        , LTRIM(RTRIM(TH046)) SrOrigTaxAmt
                                    , LTRIM(RTRIM(TH047)) SrPreTaxAmt            , LTRIM(RTRIM(TH048)) SrTaxAmt
                                    , LTRIM(RTRIM(TH049)) CommissionRate         , LTRIM(RTRIM(TH050)) CommissionAmount
                                    , LTRIM(RTRIM(TH051)) OriginalCustomer       , LTRIM(RTRIM(TH052)) SrFreeSpareQty
                                    , LTRIM(RTRIM(TH053)) SrPackageFreeSpareQty  , LTRIM(RTRIM(TH054)) NotPayTemp
                                    , LTRIM(RTRIM(TH057)) ProductSerialNumberQty , LTRIM(RTRIM(TH058)) ForecastCode
                                    , LTRIM(RTRIM(TH059)) ForecastSequence       , LTRIM(RTRIM(TH060)) [Location]
                                    , LTRIM(RTRIM(TH061)) AvailableQty           , LTRIM(RTRIM(TH062)) AvailableUomNo
                                    , LTRIM(RTRIM(TH068)) MultiBatch             , LTRIM(RTRIM(TH069)) FreeSpareRate
                                    , LTRIM(RTRIM(TH070)) FinalCustomerCode      , LTRIM(RTRIM(TH071)) ReferenceQty
                                    , LTRIM(RTRIM(TH072)) ReferencePackageQty    , LTRIM(RTRIM(TH073)) TaxRate
                                    , LTRIM(RTRIM(TH074)) CRMSource              , LTRIM(RTRIM(TH075)) CRMPrefix
                                    , LTRIM(RTRIM(TH076)) CRMNo                  , LTRIM(RTRIM(TH077)) CRMSequence
                                    , LTRIM(RTRIM(TH078)) CRMContractCode        , LTRIM(RTRIM(TH079)) CRMAllowDeduction
                                    , LTRIM(RTRIM(TH080)) CRMDeductionQty        , LTRIM(RTRIM(TH081)) CRMDeductionUnit
                                    , LTRIM(RTRIM(TH082)) DebitAccount           , LTRIM(RTRIM(TH083)) CreditAccount
                                    , LTRIM(RTRIM(TH084)) TaxAmountAccount       , LTRIM(RTRIM(TH089)) BusinessItemNumber
                                    , LTRIM(RTRIM(TH090)) TaxCode                , LTRIM(RTRIM(TH091)) DiscountAmount
                                    , LTRIM(RTRIM(TH094)) K2NO                   , LTRIM(RTRIM(TH500)) MarkingBINRecord
                                    , LTRIM(RTRIM(TH501)) MarkingManagement      , LTRIM(RTRIM(TH502)) BillingUnitInPackage
                                    , LTRIM(RTRIM(TH503)) DATECODE
                                    , CASE WHEN LEN(LTRIM(RTRIM(CREATE_DATE))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_DATE)) as date), 'yyyy-MM-dd') ELSE NULL END TransferDate
                                    , CASE WHEN LEN(LTRIM(RTRIM(CREATE_TIME))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_TIME)) as datetime), 'HH:mm:ss') ELSE NULL END TransferTime
                                    FROM COPTH
                                    WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "RoErpFullNo", @" AND (LTRIM(RTRIM(TH001)) + '-' + LTRIM(RTRIM(TH002))) LIKE '%' + @RoErpFullNo + '%'", RoErpFullNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SyncStartDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME >= @SyncStartDate OR MODI_DATE + ' ' + MODI_TIME >= @SyncStartDate)", SyncStartDate.Length > 0 ? Convert.ToDateTime(SyncStartDate).ToString("yyyyMMdd 00:00:00") : "");
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SyncEndDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME <= @SyncEndDate OR MODI_DATE + ' ' + MODI_TIME <= @SyncEndDate)", SyncEndDate.Length > 0 ? Convert.ToDateTime(SyncEndDate).ToString("yyyyMMdd 23:59:59") : ""); sql += @" ORDER BY TransferDate, TransferTime";

                            roDetails = sqlConnection.Query<RoDetail>(sql, dynamicParameters).ToList();
                            #endregion
                        }
                        using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                        {
                            #region //撈取客戶ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT CustomerId, CustomerNo 
                                    FROM SCM.Customer
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<Customer> resultCustomers = sqlConnection.Query<Customer>(sql, dynamicParameters).ToList();

                            receiveOrders = receiveOrders.Join(resultCustomers, x => x.CustomerNo, y => y.CustomerNo, (x, y) => { x.CustomerId = y.CustomerId; return x; }).ToList();
                            #endregion

                            #region //撈取部門ID
                            sql = @"SELECT DepartmentId, DepartmentNo
                                    FROM BAS.Department
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            List<Department> resultDepartments = sqlConnection.Query<Department>(sql, dynamicParameters).ToList();
                            receiveOrders = receiveOrders.GroupJoin(resultDepartments, x => x.DepartmentNo, y => y.DepartmentNo, (x, y) => { x.DepartmentId = y.FirstOrDefault()?.DepartmentId; return x; }).ToList();
                            #endregion

                            #region //撈取業務人員ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.UserId, a.UserNo 
                                    FROM BAS.[User] a";
                            List<User> resultUser = sqlConnection.Query<User>(sql, dynamicParameters).ToList();
                            receiveOrders = receiveOrders.GroupJoin(resultUser, x => x.SalesmenNo, y => y.UserNo, (x, y) => { x.SalesmenId = y.FirstOrDefault()?.UserId; return x; }).ToList();
                            receiveOrders = receiveOrders.GroupJoin(resultUser, x => x.ConfirmUserNo, y => y.UserNo, (x, y) => { x.ConfirmUserId = y.FirstOrDefault()?.UserId; return x; }).ToList();
                            receiveOrders = receiveOrders.GroupJoin(resultUser, x => x.CollectionSalesmenNo, y => y.UserNo, (x, y) => { x.CollectionSalesmenId = y.FirstOrDefault()?.UserId; return x; }).ToList();
                            #endregion

                            #region //撈取單位ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT UomId, UomNo
                                    FROM PDM.UnitOfMeasure
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            List<Uom> resultSoPriceUomrNos = sqlConnection.Query<Uom>(sql, dynamicParameters).ToList(); 
                            roDetails = roDetails.Join(resultSoPriceUomrNos, x => x.AvailableUomNo.ToUpper(), y => y.UomNo, (x, y) => { x.AvailableUomId = y.UomId; return x; }).ToList();
                            roDetails = roDetails.Join(resultSoPriceUomrNos, x => x.UomNo.ToUpper(), y => y.UomNo, (x, y) => { x.UomId = y.UomId; return x; }).ToList();
                            //roDetails = roDetails.GroupJoin(resultSoPriceUomrNos, x => x.AvailableUomNo.ToUpper(), y => y.UomNo, (x, y) => { x.UomId = y.FirstOrDefault()?.UomId; return x; }).ToList();
                            roDetails = roDetails.GroupJoin(resultSoPriceUomrNos, x => x.PackageUomNo.ToUpper(), y => y.UomNo, (x, y) => { x.PackageUomId = y.FirstOrDefault()?.UomId; return x; }).ToList();
                            #endregion

                            #region //撈取品號ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT MtlItemId, MtlItemNo
                                    FROM PDM.MtlItem
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            List<MtlItem> resultMtlitems = sqlConnection.Query<MtlItem>(sql, dynamicParameters).ToList();
                            roDetails = roDetails.Join(resultMtlitems, x => x.MtlItemNo, y => y.MtlItemNo, (x, y) => { x.MtlItemId = y.MtlItemId; return x; }).Select(x => x).ToList();
                            #endregion

                            #region //撈取庫別ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT InventoryId, InventoryNo
                                    FROM SCM.Inventory
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            List<Inventory> resultInventories = sqlConnection.Query<Inventory>(sql, dynamicParameters).ToList();
                            roDetails = roDetails.Join(resultInventories, x => x.InventoryNo, y => y.InventoryNo, (x, y) => { x.InventoryId = y.InventoryId; return x; }).ToList();
                            #endregion

                            #region //判斷SCM.SoDetail是否有資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.SoDetailId, a.SoId,b.SoErpPrefix,b.SoErpNo, a.SoSequence
                                    FROM SCM.SoDetail a
                                    INNER JOIN SCM.SaleOrder b ON a.SoId = b.SoId
                                    WHERE b.CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<SoDetail> resultSoDetails = sqlConnection.Query<SoDetail>(sql, dynamicParameters).ToList();

                            roDetails = roDetails.GroupJoin(resultSoDetails, x => new { x.SoErpPrefix, x.SoErpNo, x.SoSequence }, y => new { y.SoErpPrefix, y.SoErpNo, y.SoSequence }, (x, y) => { x.SoDetailId = y.FirstOrDefault()?.SoDetailId; return x; }).ToList();
                            #endregion

                            #region //判斷SCM.ReceiveOrder是否有資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT RoId, RoErpPrefix, RoErpNo
                                    FROM SCM.ReceiveOrder
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            List<ReceiveOrder> resultReceiveOrder = sqlConnection.Query<ReceiveOrder>(sql, dynamicParameters).ToList();
                            receiveOrders = receiveOrders.GroupJoin(resultReceiveOrder, x => new { x.RoErpPrefix, x.RoErpNo }, y => new { y.RoErpPrefix, y.RoErpNo }, (x, y) => { x.RoId = y.FirstOrDefault()?.RoId; return x; }).Select(x => x).ToList();
                            #endregion

                            #region //判斷SCM.TsnDetail是否有資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.TsnDetailId, b.TsnErpPrefix, b.TsnErpNo, a.TsnSequence
                                    FROM SCM.TsnDetail a
                                    INNER JOIN SCM.TempShippingNote b on a.TsnId = b.TsnId
                                    WHERE b.CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            List<TsnDetail> resultTsnDetail = sqlConnection.Query<TsnDetail>(sql, dynamicParameters).ToList();
                            roDetails = roDetails.GroupJoin(resultTsnDetail, x => new { x.TsnErpPrefix, x.TsnErpNo, x.TsnSequence }, y => new { y.TsnErpPrefix, y.TsnErpNo, y.TsnSequence }, (x, y) => { x.TsnDetailId = y.FirstOrDefault()?.TsnDetailId; return x; }).Select(x => x).ToList();
                            #endregion

                            #region //銷貨單單頭(新增/修改)
                            List<ReceiveOrder> addReceiveOrders = receiveOrders.Where(x => x.RoId == null).ToList();
                            List<ReceiveOrder> updateReceiveOrders = receiveOrders.Where(x => x.RoId != null).ToList();

                            #region //新增
                            dynamicParameters = new DynamicParameters();
                            if (addReceiveOrders.Count > 0)
                            {
                                addReceiveOrders
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.CompanyId = CurrentCompany;
                                        x.TransferStatusMES = "Y";
                                        x.CreateDate = CreateDate;
                                        x.LastModifiedDate = LastModifiedDate;
                                        x.CreateBy = CreateBy;
                                        x.LastModifiedBy = LastModifiedBy;
                                    });

                                #region //新增單頭資料表
                                sql = @"INSERT INTO SCM.ReceiveOrder (CompanyId, RoErpPrefix, RoErpNo, ReceiveDate, CustomerId, DepartmentId, SalesmenId, 
                                        CustomerFullName, CustomerAddressFirst, CustomerAddressSecond, Factory, Currency, ExchangeRate, 
                                        OriginalAmount, InvNumEnd, UiNo, InvoiceType, TaxType, InvoiceAddressFirst, InvoiceAddressSecond, Remark, 
                                        InvoiceDate, PrintCount, ConfirmStatus, UpdateCode, OriginalTaxAmount, CollectionSalesmenId, Remark1, Remark2, Remark3, InvoicesVoid, 
                                        CustomsClearance, RowCnt, TotalQuantity, CashSales, StaffId, RevenueJournalCode, CostJournalCode, ApplyYYMM, 
                                        LCNO, INVOICENO, InvoicesPrintCount, DocDate, ConfirmUserId, TaxRate, PretaxAmount, TaxAmount, PaymentTerm, SoDetailId, 
                                        AdvOrderPrefix, AdvOrderNo, OffsetAmount, OffsetTaxAmount, TotalPackages, SignatureStatus, ChangeInvCode, 
                                        NewRoPrefix, NewRoNo, TransferStatusERP, ProcessCode, AttachInvWithShip, BondCode, TransmissionCount, 
                                        Invoicer, InvCode, ContactPerson, Courier, SiteCommCalcMethod, SiteCommRate, CommCalcInclTax, 
                                        TotalCommAmount, TransportMethod, DispatchOrderPrefix, DispatchOrderNo, DeclarationNumber, 
                                        FullNameOfDelivCustomer, RoPriceType, TelephoneNumber, FaxNumber, ShipNoticePrefix, ShipNoticeNo, 
                                        TradingTerms, CustomerEgFullName, InvNumGenMethod, DocSourceCode, NoCredLimitControl, 
                                        InstallmentSettlement, InstallmentCount, AutoApportionByInstallment, StartYearMonth, TaxCode, CustomsManual, 
                                        RemarkCode, MultipleInvoices, InvNumStart, NumberOfInvoices, MultiTaxRate, TaxCurrencyRate, 
                                        VoidDate, VoidApprovalDocNum, VoidReason, Source, IncomeDraftID, IncomeDraftSeq, IncomeVoucherType, 
                                        IncomeVoucherNumber, Status, ZeroTaxForBuyer, GenLedgerAcctType, InvoiceTime, InvCode2, InvSymbol, 
                                        DeliveryCountry, VehicleIDshow, VehicleTypeNumber, VehicleIDhide, InvDonationRecipient, InvRandomCode, 
                                        ReservedField, CreditCard4No, ContactEmail, ExpectedReceiptDate, OrigInvNumber, TransferStatusMES, 
                                        TransferTime, ConfirmTime, SourceType, PriceSourceTypeMain, SourceOrderId, SourceFull,
                                        CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.RoId
                                        VALUES ( @CompanyId, @RoErpPrefix, @RoErpNo, @ReceiveDate, @CustomerId, @DepartmentId, @SalesmenId, 
                                        @CustomerFullName, @CustomerAddressFirst, @CustomerAddressSecond, @Factory, @Currency, @ExchangeRate, 
                                        @OriginalAmount, @InvNumEnd, @UiNo, @InvoiceType, @TaxType, @InvoiceAddressFirst, @InvoiceAddressSecond, @Remark, 
                                        @InvoiceDate, @PrintCount, @ConfirmStatus, @UpdateCode, @OriginalTaxAmount, @CollectionSalesmenId, @Remark1, @Remark2, @Remark3, @InvoicesVoid, 
                                        @CustomsClearance, @RowCnt, @TotalQuantity, @CashSales, @StaffId, @RevenueJournalCode, @CostJournalCode, @ApplyYYMM, 
                                        @LCNO, @INVOICENO, @InvoicesPrintCount, @DocDate, @ConfirmUserId, @TaxRate, @PretaxAmount, @TaxAmount, @PaymentTerm, @SoDetailId, 
                                        @AdvOrderPrefix, @AdvOrderNo, @OffsetAmount, @OffsetTaxAmount, @TotalPackages, @SignatureStatus, @ChangeInvCode, 
                                        @NewRoPrefix, @NewRoNo, @TransferStatusERP, @ProcessCode, @AttachInvWithShip, @BondCode, @TransmissionCount, 
                                        @Invoicer, @InvCode, @ContactPerson, @Courier, @SiteCommCalcMethod, @SiteCommRate, @CommCalcInclTax, 
                                        @TotalCommAmount, @TransportMethod, @DispatchOrderPrefix, @DispatchOrderNo, @DeclarationNumber, 
                                        @FullNameOfDelivCustomer, @RoPriceType, @TelephoneNumber, @FaxNumber, @ShipNoticePrefix, @ShipNoticeNo, 
                                        @TradingTerms, @CustomerEgFullName, @InvNumGenMethod, @DocSourceCode, @NoCredLimitControl, 
                                        @InstallmentSettlement, @InstallmentCount, @AutoApportionByInstallment, @StartYearMonth, @TaxCode, @CustomsManual, 
                                        @RemarkCode, @MultipleInvoices, @InvNumStart, @NumberOfInvoices, @MultiTaxRate, @TaxCurrencyRate, 
                                        @VoidDate, @VoidApprovalDocNum, @VoidReason, @Source, @IncomeDraftID, @IncomeDraftSeq, @IncomeVoucherType, 
                                        @IncomeVoucherNumber, @Status, @ZeroTaxForBuyer, @GenLedgerAcctType, @InvoiceTime, @InvCode2, @InvSymbol, 
                                        @DeliveryCountry, @VehicleIDshow, @VehicleTypeNumber, @VehicleIDhide, @InvDonationRecipient, @InvRandomCode, 
                                        @ReservedField, @CreditCard4No, @ContactEmail, @ExpectedReceiptDate, @OrigInvNumber, @TransferStatusMES, 
                                        @TransferTime, @ConfirmTime, @SourceType, @PriceSourceTypeMain, @SourceOrderId, @SourceFull, 
                                        @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                int addMain = sqlConnection.Execute(sql, addReceiveOrders);
                                mainAffected += addMain;
                                #endregion
                                
                            }
                            #endregion

                            #region //修改
                            dynamicParameters = new DynamicParameters();
                            if (updateReceiveOrders.Count > 0)
                            {
                                updateReceiveOrders
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.LastModifiedDate = LastModifiedDate;
                                        x.LastModifiedBy = LastModifiedBy;
                                    });

                                #region //更新單頭資料表
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE SCM.ReceiveOrder SET
                                ReceiveDate = @ReceiveDate,
                                CustomerId = @CustomerId,
                                DepartmentId = @DepartmentId,
                                SalesmenId = @SalesmenId,
                                CustomerFullName = @CustomerFullName,
                                CustomerAddressFirst = @CustomerAddressFirst,
                                CustomerAddressSecond = @CustomerAddressSecond,
                                Factory = @Factory,
                                Currency = @Currency,
                                ExchangeRate = @ExchangeRate,
                                OriginalAmount = @OriginalAmount,
                                InvNumEnd = @InvNumEnd,
                                UiNo = @UiNo,
                                InvoiceType = @InvoiceType,
                                TaxType = @TaxType,
                                InvoiceAddressFirst = @InvoiceAddressFirst,
                                InvoiceAddressSecond = @InvoiceAddressSecond,
                                Remark = @Remark,
                                InvoiceDate = @InvoiceDate,
                                OriginalTaxAmount = @OriginalTaxAmount,
                                CollectionSalesmenId = @CollectionSalesmenId,
                                Remark1 = @Remark1,
                                Remark2 = @Remark2,
                                Remark3 = @Remark3,
                                InvoicesVoid = @InvoicesVoid,
                                CustomsClearance = @CustomsClearance,
                                RowCnt = @RowCnt,
                                TotalQuantity = @TotalQuantity,
                                CashSales = @CashSales,
                                StaffId = @StaffId,
                                RevenueJournalCode = @RevenueJournalCode,
                                CostJournalCode = @CostJournalCode,
                                ApplyYYMM = @ApplyYYMM,
                                LCNO = @LCNO,
                                INVOICENO = @INVOICENO,
                                DocDate = @DocDate,
                                TaxRate = @TaxRate,
                                PretaxAmount = @PretaxAmount,
                                TaxAmount = @TaxAmount,
                                PaymentTerm = @PaymentTerm,
                                OffsetAmount = @OffsetAmount,
                                OffsetTaxAmount = @OffsetTaxAmount,
                                ChangeInvCode = @ChangeInvCode,
                                ProcessCode = @ProcessCode,
                                AttachInvWithShip = @AttachInvWithShip,
                                BondCode = @BondCode,
                                ContactPerson = @ContactPerson,
                                Courier = @Courier,
                                SiteCommCalcMethod = @SiteCommCalcMethod,
                                SiteCommRate = @SiteCommRate,
                                TotalCommAmount = @TotalCommAmount,
                                TransportMethod = @TransportMethod,
                                DeclarationNumber = @DeclarationNumber,
                                FullNameOfDelivCustomer = @FullNameOfDelivCustomer,
                                TelephoneNumber = @TelephoneNumber,
                                FaxNumber = @FaxNumber,
                                TradingTerms = @TradingTerms,
                                CustomerEgFullName = @CustomerEgFullName,
                                InvNumGenMethod = @InvNumGenMethod,
                                NoCredLimitControl = @NoCredLimitControl,
                                TaxCode = @TaxCode,
                                MultipleInvoices = @MultipleInvoices,
                                InvNumStart = @InvNumStart,
                                MultiTaxRate = @MultiTaxRate,
                                TaxCurrencyRate = @TaxCurrencyRate,
                                InvoiceTime = @InvoiceTime,
                                VehicleIDshow = @VehicleIDshow,
                                VehicleTypeNumber = @VehicleTypeNumber,
                                VehicleIDhide = @VehicleIDhide,
                                InvDonationRecipient = @InvDonationRecipient,
                                InvRandomCode = @InvRandomCode,
                                CreditCard4No = @CreditCard4No,
                                ContactEmail = @ContactEmail,
                                SourceType = @SourceType,
                                PriceSourceTypeMain = @PriceSourceTypeMain,
                                SourceOrderId = @SourceOrderId,
                                SourceFull = @SourceFull,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RoId = @RoId";
                                int updateMain = sqlConnection.Execute(sql, updateReceiveOrders);
                                mainAffected += updateMain;
                                #endregion

                                
                            }
                            #endregion
                            #endregion

                            #region //銷貨單單身(新增/修改)
                            #region //撈取銷貨單單頭ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT RoId, RoErpPrefix, RoErpNo
                                    FROM SCM.ReceiveOrder
                                    WHERE CompanyId = @CompanyId
                                    ORDER BY RoId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            resultReceiveOrder = sqlConnection.Query<ReceiveOrder>(sql, dynamicParameters).ToList();

                            roDetails = roDetails.Join(resultReceiveOrder, x => new { x.RoErpPrefix, x.RoErpNo }, y => new { y.RoErpPrefix, y.RoErpNo }, (x, y) => { x.RoId = y.RoId; return x; }).Select(x => x).ToList();
                            #endregion

                            #region //判斷SCM.RoDetail是否有資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.RoDetailId, a.RoId, a.RoSequence
                                    FROM SCM.RoDetail a
                                    INNER JOIN SCM.ReceiveOrder b ON a.RoId = b.RoId
                                    WHERE b.CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<RoDetail> resultRoDetails = sqlConnection.Query<RoDetail>(sql, dynamicParameters).ToList();

                            roDetails = roDetails.GroupJoin(resultRoDetails, x => new { x.RoId, x.RoSequence }, y => new { y.RoId, y.RoSequence }, (x, y) => { x.RoDetailId = y.FirstOrDefault()?.RoDetailId; return x; }).ToList();
                            #endregion

                            List<RoDetail> addRoDetails = roDetails.Where(x => x.RoDetailId == null && x.SoDetailId != null).ToList();
                            List<RoDetail> updateRoDetails = roDetails.Where(x => x.RoDetailId != null).ToList();

                            #region //新增
                            dynamicParameters = new DynamicParameters();
                            if (addRoDetails.Count > 0)
                            {
                                addRoDetails
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.TransferStatusMES = "Y";
                                        x.CreateDate = CreateDate;
                                        x.LastModifiedDate = LastModifiedDate;
                                        x.CreateBy = CreateBy;
                                        x.LastModifiedBy = LastModifiedBy;
                                    });
                                sql = @"INSERT INTO SCM.RoDetail (RoId, RoSequence, MtlItemId, InventoryId, Quantity, UomId, UnitPrice, Amount, 
                                        SoDetailId, LotNumber, Remarks, CustomerMtlItemNo, ConfirmationCode, UpdateCode, FreeSpareQty, DiscountRate, 
                                        CheckOutCode, CheckOutPrefix, CheckOutNo, CheckOutSequence, ProjectCode, Type, TsnDetailId, OrigPreTaxAmt, 
                                        OrigTaxAmt, PreTaxAmt, TaxAmt, PackageQty, PackageFreeSpareQty, PackageUomId, BondedCode, SrQty, 
                                        SrPackageQty, SrOrigPreTaxAmt, SrOrigTaxAmt, SrPreTaxAmt, SrTaxAmt, CommissionRate, CommissionAmount, 
                                        OriginalCustomer, SrFreeSpareQty, SrPackageFreeSpareQty, NotPayTemp, ProductSerialNumberQty, ForecastCode, 
                                        ForecastSequence, Location, AvailableQty, AvailableUomId, MultiBatch, FreeSpareRate, FinalCustomerCode, 
                                        ReferenceQty, ReferencePackageQty, TaxRate, CRMSource, CRMPrefix, CRMNo, CRMSequence, CRMContractCode, 
                                        CRMAllowDeduction, CRMDeductionQty, CRMDeductionUnit, DebitAccount, CreditAccount, TaxAmountAccount, 
                                        BusinessItemNumber, TaxCode, DiscountAmount, K2NO, MarkingBINRecord, MarkingManagement, 
                                        BillingUnitInPackage, DATECODE, TransferStatusMES, CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        VALUES ( @RoId, @RoSequence, @MtlItemId, @InventoryId, @Quantity, @UomId, @UnitPrice, @Amount, 
                                        @SoDetailId, @LotNumber, @Remarks, @CustomerMtlItemNo, @ConfirmationCode, @UpdateCode, @FreeSpareQty, @DiscountRate, 
                                        @CheckOutCode, @CheckOutPrefix, @CheckOutNo, @CheckOutSequence, @ProjectCode, @Type, @TsnDetailId, @OrigPreTaxAmt, 
                                        @OrigTaxAmt, @PreTaxAmt, @TaxAmt, @PackageQty, @PackageFreeSpareQty, @PackageUomId, @BondedCode, @SrQty, 
                                        @SrPackageQty, @SrOrigPreTaxAmt, @SrOrigTaxAmt, @SrPreTaxAmt, @SrTaxAmt, @CommissionRate, @CommissionAmount, 
                                        @OriginalCustomer, @SrFreeSpareQty, @SrPackageFreeSpareQty, @NotPayTemp, @ProductSerialNumberQty, @ForecastCode, 
                                        @ForecastSequence, @Location, @AvailableQty, @AvailableUomId, @MultiBatch, @FreeSpareRate, @FinalCustomerCode, 
                                        @ReferenceQty, @ReferencePackageQty, @TaxRate, @CRMSource, @CRMPrefix, @CRMNo, @CRMSequence, @CRMContractCode, 
                                        @CRMAllowDeduction, @CRMDeductionQty, @CRMDeductionUnit, @DebitAccount, @CreditAccount, @TaxAmountAccount, 
                                        @BusinessItemNumber, @TaxCode, @DiscountAmount, @K2NO, @MarkingBINRecord, @MarkingManagement, 
                                        @BillingUnitInPackage, @DATECODE, @TransferStatusMES, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy
                                        )";
                                int addDetail = sqlConnection.Execute(sql, addRoDetails);
                                detailAffected += addDetail;
                            }
                            #endregion

                            #region //修改
                            dynamicParameters = new DynamicParameters();
                            if (updateRoDetails.Count > 0)
                            {
                                updateRoDetails
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.TransferStatusMES = "Y";
                                        x.LastModifiedDate = LastModifiedDate;
                                        x.LastModifiedBy = LastModifiedBy;
                                    });

                                sql = @"UPDATE SCM.RoDetail SET
                                        MtlItemId = @MtlItemId, 
                                        InventoryId = @InventoryId, 
                                        Quantity = @Quantity, 
                                        UomId = @UomId, 
                                        UnitPrice = @UnitPrice, 
                                        Amount = @Amount, 
                                        SoDetailId = @SoDetailId, 
                                        LotNumber = @LotNumber, 
                                        Remarks = @Remarks, 
                                        CustomerMtlItemNo = @CustomerMtlItemNo, 
                                        ConfirmationCode = @ConfirmationCode, 
                                        UpdateCode = @UpdateCode, 
                                        FreeSpareQty = @FreeSpareQty, 
                                        DiscountRate = @DiscountRate, 
                                        CheckOutCode = @CheckOutCode, 
                                        CheckOutPrefix = @CheckOutPrefix, 
                                        CheckOutNo = @CheckOutNo, 
                                        CheckOutSequence = @CheckOutSequence, 
                                        ProjectCode = @ProjectCode, 
                                        Type = @Type, 
                                        TsnDetailId = @TsnDetailId, 
                                        OrigPreTaxAmt = @OrigPreTaxAmt, 
                                        OrigTaxAmt = @OrigTaxAmt, 
                                        PreTaxAmt = @PreTaxAmt, 
                                        TaxAmt = @TaxAmt, 
                                        PackageQty = @PackageQty, 
                                        PackageFreeSpareQty = @PackageFreeSpareQty, 
                                        PackageUomId = @PackageUomId, 
                                        BondedCode = @BondedCode, 
                                        SrQty = @SrQty, 
                                        SrPackageQty = @SrPackageQty, 
                                        SrOrigPreTaxAmt = @SrOrigPreTaxAmt, 
                                        SrOrigTaxAmt = @SrOrigTaxAmt, 
                                        SrPreTaxAmt = @SrPreTaxAmt, 
                                        SrTaxAmt = @SrTaxAmt, 
                                        CommissionRate = @CommissionRate, 
                                        CommissionAmount = @CommissionAmount, 
                                        OriginalCustomer = @OriginalCustomer, 
                                        SrFreeSpareQty = @SrFreeSpareQty, 
                                        SrPackageFreeSpareQty = @SrPackageFreeSpareQty, 
                                        NotPayTemp = @NotPayTemp, 
                                        ProductSerialNumberQty = @ProductSerialNumberQty, 
                                        ForecastCode = @ForecastCode, 
                                        ForecastSequence = @ForecastSequence, 
                                        Location = @Location, 
                                        AvailableQty = @AvailableQty, 
                                        AvailableUomId = @AvailableUomId, 
                                        MultiBatch = @MultiBatch, 
                                        FreeSpareRate = @FreeSpareRate, 
                                        FinalCustomerCode = @FinalCustomerCode, 
                                        ReferenceQty = @ReferenceQty, 
                                        ReferencePackageQty = @ReferencePackageQty, 
                                        TaxRate = @TaxRate, 
                                        CRMSource = @CRMSource, 
                                        CRMPrefix = @CRMPrefix, 
                                        CRMNo = @CRMNo, 
                                        CRMSequence = @CRMSequence, 
                                        CRMContractCode = @CRMContractCode, 
                                        CRMAllowDeduction = @CRMAllowDeduction, 
                                        CRMDeductionQty = @CRMDeductionQty, 
                                        CRMDeductionUnit = @CRMDeductionUnit, 
                                        DebitAccount = @DebitAccount, 
                                        CreditAccount = @CreditAccount, 
                                        TaxAmountAccount = @TaxAmountAccount, 
                                        BusinessItemNumber = @BusinessItemNumber, 
                                        TaxCode = @TaxCode, 
                                        DiscountAmount = @DiscountAmount, 
                                        K2NO = @K2NO, 
                                        MarkingBINRecord = @MarkingBINRecord, 
                                        MarkingManagement = @MarkingManagement, 
                                        BillingUnitInPackage = @BillingUnitInPackage, 
                                        DATECODE = @DATECODE, 
                                        TransferStatusMES = @TransferStatusMES, 
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE RoDetailId = @RoDetailId";
                                int updateDetail = sqlConnection.Execute(sql, updateRoDetails);
                                detailAffected += updateDetail;
                            }
                            #endregion
                            #endregion
                        }
                    }
                    #endregion

                    if (TranSync == "Y")
                    {
                        using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //ERP訂單單頭資料
                            sql = @"SELECT TG001 RoErpPrefix, TG002 RoErpNo
                                    FROM COPTG
                                    WHERE 1=1
                                    ORDER BY TG001, TG002";
                            var resultErpRo = erpConnection.Query<ReceiveOrder>(sql, dynamicParameters);
                            #endregion

                            using (SqlConnection bmConnection = new SqlConnection(MainConnectionStrings))
                            {
                                #region //BM訂單單頭資料
                                sql = @"SELECT RoErpPrefix, RoErpNo
                                        FROM SCM.ReceiveOrder
                                        WHERE CompanyId = @CompanyId
                                        ORDER BY RoErpPrefix, RoErpNo";
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                var resultBmIt = bmConnection.Query<ReceiveOrder>(sql, dynamicParameters);
                                #endregion

                                #region //篩選ERP不存在但BM還存在的訂單單頭
                                var dictionaryErpRo = resultErpRo.ToDictionary(x => x.RoErpPrefix + "-" + x.RoErpNo, x => x.RoErpPrefix + "-" + x.RoErpNo);
                                var dictionaryBmRo = resultBmIt.ToDictionary(x => x.RoErpPrefix + "-" + x.RoErpNo, x => x.RoErpPrefix + "-" + x.RoErpNo);
                                var changeRo = dictionaryBmRo.Where(x => !dictionaryErpRo.ContainsKey(x.Key)).ToList();
                                var changeRoList = changeRo.Select(x => x.Value).ToList();
                                #endregion

                                #region //刪除異動訂單單頭
                                if (changeRoList.Count > 0)
                                {
                                    #region //單身刪除
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE SCM.RoDetail
                                            WHERE RoId IN (
                                                SELECT RoId
                                                FROM SCM.ReceiveOrder
                                                WHERE RoErpPrefix + '-' + RoErpNo IN @RoErpFullNo
                                            )";
                                    dynamicParameters.Add("RoErpFullNo", changeRoList.Select(x => x).ToArray());
                                    mainDelAffected += bmConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //單頭刪除
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE SCM.ReceiveOrder
                                            WHERE RoErpPrefix + '-' + RoErpNo IN @RoErpFullNo";
                                    dynamicParameters.Add("RoErpFullNo", changeRoList.Select(x => x).ToArray());
                                    detailDelAffected += bmConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                                #endregion


                            }

                            #region //ERP銷貨單身資料
                            sql = @"SELECT TH001 RoErpPrefix, TH002 RoErpNo, TH003 RoSequence
                                    FROM COPTH
                                    WHERE 1=1
                                    ORDER BY TH001, TH002, TH003";
                            var resultErpRoDetail = erpConnection.Query<RoDetail>(sql, dynamicParameters);
                            #endregion

                            using (SqlConnection bmConnection = new SqlConnection(MainConnectionStrings))
                            {
                                #region //BM銷貨單單身資料
                                sql = @"SELECT b.RoErpPrefix, b.RoErpNo, a.RoSequence
                                        FROM SCM.RoDetail a
                                        INNER JOIN SCM.ReceiveOrder b ON a.RoId = b.RoId
                                        WHERE b.CompanyId = @CompanyId
                                        ORDER BY b.RoErpPrefix, b.RoErpNo, a.RoSequence";
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                var resultBmRoDetail = bmConnection.Query<RoDetail>(sql, dynamicParameters);
                                #endregion

                                #region //篩選ERP不存在但BM還存在的銷貨單單身
                                var dictionaryErpRoDetail = resultErpRoDetail.ToDictionary(x => x.RoErpPrefix + "-" + x.RoErpNo + "-" + x.RoSequence, x => x.RoErpPrefix + "-" + x.RoErpNo + "-" + x.RoSequence);
                                var dictionaryBmRoDetail = resultBmRoDetail.ToDictionary(x => x.RoErpPrefix + "-" + x.RoErpNo + "-" + x.RoSequence, x => x.RoErpPrefix + "-" + x.RoErpNo + "-" + x.RoSequence);
                                var changeRoDetail = dictionaryBmRoDetail.Where(x => !dictionaryErpRoDetail.ContainsKey(x.Key)).ToList();
                                var changeRoDetailList = changeRoDetail.Select(x => x.Value).ToList();
                                #endregion

                                #region //刪除訂單單身
                                if (changeRoDetailList.Count > 0)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE a
                                            FROM SCM.RoDetail a
                                            INNER JOIN SCM.ReceiveOrder b ON a.RoId = b.RoId
                                            WHERE b.RoErpPrefix + '-' + b.RoErpNo + '-' + RIGHT('0000' + a.RoSequence, 4) IN @RoErpFullNo";
                                    dynamicParameters.Add("RoErpFullNo", changeRoDetailList.Select(x => x).ToArray());
                                    detailDelAffected += bmConnection.Execute(sql, dynamicParameters, null, 300);
                                }
                                #endregion
                            }
                        }
                    }

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

        #region //UpdateReturnReceiveOrderSynchronize -- 銷退單資料手動同步 -- Shintokuro 2024.02.20
        public string UpdateReturnReceiveOrderSynchronize(string RtErpFullNo, string SyncStartDate, string SyncEndDate
            , string NormalSync, string TranSync, string CompanyNo
            )
        {
            try
            {
                List<ReturnReceiveOrder> returnReceiveOrder = new List<ReturnReceiveOrder>();
                List<RtDetail> rtDetails = new List<RtDetail>();

                int mainAffected = 0, detailAffected = 0, mainDelAffected = 0, detailDelAffected = 0; ;
                string companyNo = "";
                string companyNoBase = "";
                //if (CompanyNo.Length <= 0) throw new SystemException("【公司】不能為空!");
                int CompanyId = -1;
                string ErpConnectionStrings = "";

                using (TransactionScope transactionScope = new TransactionScope(TransactionScopeOption.Required, new TimeSpan(0, 180, 0)))
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
                            CompanyId = Convert.ToInt32(item.CompanyId);
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
                            #region //判斷ERP銷退單資料是否存在
                            if (RtErpFullNo.Length > 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM COPTI
                                        WHERE (LTRIM(RTRIM(TI001)) + '-' + LTRIM(RTRIM(TI002))) LIKE '%' + @RtErpFullNo + '%'";
                                dynamicParameters.Add("RtErpFullNo", RtErpFullNo);

                                var resultTsn = sqlConnection.Query(sql, dynamicParameters);
                                if (resultTsn.Count() <= 0) throw new SystemException("ERP銷退單資料錯誤");
                            }
                            #endregion

                            #region //撈取ERP銷退單頭資料 
                            sql = @"SELECT 
                                    LTRIM(RTRIM(TI001)) RtErpPrefix,            LTRIM(RTRIM(TI002)) RtErpNo,
                                                                                LTRIM(RTRIM(TI004)) CustomerNo,
                                    LTRIM(RTRIM(TI005)) DepartmentNo,           LTRIM(RTRIM(TI006)) SalesmenNo,
                                    LTRIM(RTRIM(TI007)) Factory,                LTRIM(RTRIM(TI008)) Currency,
                                    LTRIM(RTRIM(TI009)) ExchangeRate,           LTRIM(RTRIM(TI010)) OriginalAmount,
                                    LTRIM(RTRIM(TI011)) OriginalTaxAmount,      LTRIM(RTRIM(TI012)) InvoiceType,
                                    LTRIM(RTRIM(TI013)) TaxType,                LTRIM(RTRIM(TI014)) InvNumEnd,
                                    LTRIM(RTRIM(TI015)) UiNo,                   LTRIM(RTRIM(TI016)) PrintCount,
                                                                                LTRIM(RTRIM(TI018)) UpdateCode,
                                    LTRIM(RTRIM(TI019)) ConfirmStatus,          LTRIM(RTRIM(TI020)) Remark,
                                    LTRIM(RTRIM(TI021)) CustomerFullName,       LTRIM(RTRIM(TI022)) CollectionSalesmenNo,
                                    LTRIM(RTRIM(TI023)) Remark1,                LTRIM(RTRIM(TI024)) Remark2,
                                    LTRIM(RTRIM(TI025)) Remark3,                LTRIM(RTRIM(TI026)) DeductionDistinction,
                                    LTRIM(RTRIM(TI027)) CustomsClearance,       LTRIM(RTRIM(TI028)) RowCnt,
                                    LTRIM(RTRIM(TI029)) TotalQuantity,          LTRIM(RTRIM(TI030)) StaffNo,
                                    LTRIM(RTRIM(TI031)) RevenueJournalCode,     LTRIM(RTRIM(TI032)) CostJournalCode,
                                    LTRIM(RTRIM(TI033)) ApplyYYMM,              
                                    LTRIM(RTRIM(TI035)) ConfirmUserNo,          LTRIM(RTRIM(TI036)) TaxRate,
                                    LTRIM(RTRIM(TI037)) PretaxAmount,           LTRIM(RTRIM(TI038)) TaxAmount,
                                    LTRIM(RTRIM(TI039)) PaymentTerm,            LTRIM(RTRIM(TI040)) TotalPackages,
                                    LTRIM(RTRIM(TI041)) SignatureStatus,        LTRIM(RTRIM(TI042)) ProcessCode,
                                    LTRIM(RTRIM(TI043)) TransferStatusERP,      LTRIM(RTRIM(TI044)) BondCode,
                                    LTRIM(RTRIM(TI045)) TransmissionCount,      LTRIM(RTRIM(TI046)) CustomerAddressFirst,
                                    LTRIM(RTRIM(TI047)) CustomerAddressSecond,  LTRIM(RTRIM(TI049)) ContactPerson,
                                    LTRIM(RTRIM(TI050)) Courier,                LTRIM(RTRIM(TI051)) SiteCommCalcMethod,
                                    LTRIM(RTRIM(TI052)) SiteCommRate,           LTRIM(RTRIM(TI053)) CommCalcInclTax,
                                    LTRIM(RTRIM(TI054)) TotalCommAmount,        LTRIM(RTRIM(TI057)) TradingTerms,
                                    LTRIM(RTRIM(TI058)) CustomerEgFullName,     LTRIM(RTRIM(TI061)) DebitNote,
                                    LTRIM(RTRIM(TI064)) TaxCode,                LTRIM(RTRIM(TI066)) RemarkCode,
                                    LTRIM(RTRIM(TI067)) MultipleInvoices,       LTRIM(RTRIM(TI068)) InvNumStart,
                                    LTRIM(RTRIM(TI069)) NumberOfInvoices,       LTRIM(RTRIM(TI070)) MultiTaxRate,
                                    LTRIM(RTRIM(TI071)) TaxCurrencyRate,        
                                    LTRIM(RTRIM(TI075)) VoidTime,               LTRIM(RTRIM(TI076)) VoidRemark,
                                    LTRIM(RTRIM(TI077)) IncomeDraftID,          LTRIM(RTRIM(TI078)) IncomeDraftSeq,
                                    LTRIM(RTRIM(TI079)) IncomeVoucherType,      LTRIM(RTRIM(TI080)) IncomeVoucherNumber,
                                    LTRIM(RTRIM(TI081)) [Status],               LTRIM(RTRIM(TI086)) GenLedgerAcctType,
                                    LTRIM(RTRIM(TI087)) InvCode2,               LTRIM(RTRIM(TI088)) InvSymbol,
                                    LTRIM(RTRIM(TI091)) ContactEmail,           LTRIM(RTRIM(TI092)) OrigInvNumber		
                                    , CASE WHEN LEN(LTRIM(RTRIM(TI003))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TI003)) as date), 'yyyy-MM-dd') ELSE NULL END ReturnDate
                                    , CASE WHEN LEN(LTRIM(RTRIM(TI017))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TI017)) as date), 'yyyy-MM-dd') ELSE NULL END InvoiceDate
                                    , CASE WHEN LEN(LTRIM(RTRIM(TI034))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TI034)) as date), 'yyyy-MM-dd') ELSE NULL END DocDate
                                    , CASE WHEN LEN(LTRIM(RTRIM(TI074))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TI074)) as date), 'yyyy-MM-dd') ELSE NULL END VoidDate
                                    , CASE WHEN LEN(LTRIM(RTRIM(CREATE_DATE))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_DATE)) as date), 'yyyy-MM-dd') ELSE NULL END TransferDate
                                    , CASE WHEN LEN(LTRIM(RTRIM(CREATE_TIME))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_TIME)) as datetime), 'HH:mm:ss') ELSE NULL END TransferTime
                                    FROM COPTI
                                    WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "RtErpFullNo", @" AND (LTRIM(RTRIM(TG001)) + '-' + LTRIM(RTRIM(TG002))) LIKE '%' + @RtErpFullNo + '%'", RtErpFullNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SyncStartDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME >= @SyncStartDate OR MODI_DATE + ' ' + MODI_TIME >= @SyncStartDate)", SyncStartDate.Length > 0 ? Convert.ToDateTime(SyncStartDate).ToString("yyyyMMdd 00:00:00") : "");
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SyncEndDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME <= @SyncEndDate OR MODI_DATE + ' ' + MODI_TIME <= @SyncEndDate)", SyncEndDate.Length > 0 ? Convert.ToDateTime(SyncEndDate).ToString("yyyyMMdd 23:59:59") : ""); sql += @" ORDER BY TransferDate, TransferTime";

                            returnReceiveOrder = sqlConnection.Query<ReturnReceiveOrder>(sql, dynamicParameters).ToList();
                            #endregion

                            #region //撈取ERP銷退單身資料 
                            sql = @"SELECT  
                                     LTRIM(RTRIM(TJ001)) RtErpPrefix                 ,LTRIM(RTRIM(TJ002)) RtErpNo
                                    ,LTRIM(RTRIM(TJ003)) RtSequence                 ,LTRIM(RTRIM(TJ004)) MtlItemNo
                                    ,LTRIM(RTRIM(TJ007)) Quantity                    ,LTRIM(RTRIM(TJ008)) UomNo
                                    ,LTRIM(RTRIM(TJ011)) UnitPrice                   ,LTRIM(RTRIM(TJ012)) Amount
                                    ,LTRIM(RTRIM(TJ013)) InventoryNo                 ,LTRIM(RTRIM(TJ014)) LotNumber
                                    ,LTRIM(RTRIM(TJ015)) RoErpPrefix                 ,LTRIM(RTRIM(TJ016)) RoErpNo
                                    ,LTRIM(RTRIM(TJ017)) RoSequence
                                    ,LTRIM(RTRIM(TJ018)) SoErpPrefix                 ,LTRIM(RTRIM(TJ019)) SoErpNo
                                    ,LTRIM(RTRIM(TJ020)) SoSequence
                                    ,LTRIM(RTRIM(TJ021)) ConfirmStatus               ,LTRIM(RTRIM(TJ022)) UpdateCode
                                    ,LTRIM(RTRIM(TJ023)) Remarks                     ,LTRIM(RTRIM(TJ024)) CheckOutCode
                                    ,LTRIM(RTRIM(TJ025)) CheckOutPrefix              ,LTRIM(RTRIM(TJ026)) CheckOutNo
                                    ,LTRIM(RTRIM(TJ027)) CheckOutSeq                 ,LTRIM(RTRIM(TJ028)) ProjectCode
                                    ,LTRIM(RTRIM(TJ029)) CustomerMtlItemNo           ,LTRIM(RTRIM(TJ030)) Type
                                    ,LTRIM(RTRIM(TJ031)) OrigPreTaxAmt               ,LTRIM(RTRIM(TJ032)) OrigTaxAmt
                                    ,LTRIM(RTRIM(TJ033)) PreTaxAmt                   ,LTRIM(RTRIM(TJ034)) TaxAmt
                                    ,LTRIM(RTRIM(TJ035)) PackageQty                  ,LTRIM(RTRIM(TJ036)) PackageUomNo
                                    ,LTRIM(RTRIM(TJ037)) BondedCode                  ,LTRIM(RTRIM(TJ038)) CommissionRate
                                    ,LTRIM(RTRIM(TJ039)) CommissionAmount            ,LTRIM(RTRIM(TJ040)) OriginalCustomer
                                    ,LTRIM(RTRIM(TJ041)) FreeSpareType               ,LTRIM(RTRIM(TJ042)) FreeSpareQty
                                    ,LTRIM(RTRIM(TJ043)) PackageFreeSpareQty         ,LTRIM(RTRIM(TJ048)) NotPayTemp
                                    ,LTRIM(RTRIM(TJ049)) Location                    ,LTRIM(RTRIM(TJ050)) AvailableQty
                                    ,LTRIM(RTRIM(TJ051)) AvailableUomNo              ,LTRIM(RTRIM(TJ058)) TaxRate
                                    ,LTRIM(RTRIM(TJ070)) TaxCode                     ,LTRIM(RTRIM(TJ071)) DiscountAmount
                                    , CASE WHEN LEN(LTRIM(RTRIM(CREATE_DATE))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_DATE)) as date), 'yyyy-MM-dd') ELSE NULL END TransferDate
                                    , CASE WHEN LEN(LTRIM(RTRIM(CREATE_TIME))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_TIME)) as datetime), 'HH:mm:ss') ELSE NULL END TransferTime
                                    FROM COPTJ
                                    WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "RtErpFullNo", @" AND (LTRIM(RTRIM(TJ001)) + '-' + LTRIM(RTRIM(TJ002))) LIKE '%' + @RtErpFullNo + '%'", RtErpFullNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SyncStartDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME >= @SyncStartDate OR MODI_DATE + ' ' + MODI_TIME >= @SyncStartDate)", SyncStartDate.Length > 0 ? Convert.ToDateTime(SyncStartDate).ToString("yyyyMMdd 00:00:00") : "");
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SyncEndDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME <= @SyncEndDate OR MODI_DATE + ' ' + MODI_TIME <= @SyncEndDate)", SyncEndDate.Length > 0 ? Convert.ToDateTime(SyncEndDate).ToString("yyyyMMdd 23:59:59") : ""); sql += @" ORDER BY TransferDate, TransferTime";

                            rtDetails = sqlConnection.Query<RtDetail>(sql, dynamicParameters).ToList();
                            #endregion
                        }
                        using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                        {
                            #region //撈取客戶ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT CustomerId, CustomerNo 
                                    FROM SCM.Customer
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<Customer> resultCustomers = sqlConnection.Query<Customer>(sql, dynamicParameters).ToList();

                            returnReceiveOrder = returnReceiveOrder.Join(resultCustomers, x => x.CustomerNo, y => y.CustomerNo, (x, y) => { x.CustomerId = y.CustomerId; return x; }).ToList();
                            #endregion

                            #region //撈取部門ID
                            sql = @"SELECT DepartmentId, DepartmentNo
                                    FROM BAS.Department
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            List<Department> resultDepartments = sqlConnection.Query<Department>(sql, dynamicParameters).ToList();
                            returnReceiveOrder = returnReceiveOrder.GroupJoin(resultDepartments, x => x.DepartmentNo, y => y.DepartmentNo, (x, y) => { x.DepartmentId = y.FirstOrDefault()?.DepartmentId; return x; }).ToList();
                            #endregion

                            #region //撈取業務人員ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.UserId, a.UserNo 
                                    FROM BAS.[User] a";
                            List<User> resultUser = sqlConnection.Query<User>(sql, dynamicParameters).ToList();
                            returnReceiveOrder = returnReceiveOrder.GroupJoin(resultUser, x => x.SalesmenNo, y => y.UserNo, (x, y) => { x.SalesmenId = y.FirstOrDefault()?.UserId; return x; }).ToList();
                            returnReceiveOrder = returnReceiveOrder.GroupJoin(resultUser, x => x.ConfirmUserNo, y => y.UserNo, (x, y) => { x.ConfirmUserId = y.FirstOrDefault()?.UserId; return x; }).ToList();
                            returnReceiveOrder = returnReceiveOrder.GroupJoin(resultUser, x => x.CollectionSalesmenNo, y => y.UserNo, (x, y) => { x.CollectionSalesmenId = y.FirstOrDefault()?.UserId; return x; }).ToList();
                            returnReceiveOrder = returnReceiveOrder.GroupJoin(resultUser, x => x.StaffNo, y => y.UserNo, (x, y) => { x.StaffId = y.FirstOrDefault()?.UserId; return x; }).ToList();
                            #endregion

                            #region //撈取單位ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT UomId, UomNo
                                    FROM PDM.UnitOfMeasure
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            List<Uom> resultSoPriceUomrNos = sqlConnection.Query<Uom>(sql, dynamicParameters).ToList();
                            rtDetails = rtDetails.Join(resultSoPriceUomrNos, x => x.AvailableUomNo.ToUpper(), y => y.UomNo, (x, y) => { x.AvailableUomId = y.UomId; return x; }).ToList();
                            rtDetails = rtDetails.Join(resultSoPriceUomrNos, x => x.UomNo.ToUpper(), y => y.UomNo, (x, y) => { x.UomId = y.UomId; return x; }).ToList();
                            rtDetails = rtDetails.GroupJoin(resultSoPriceUomrNos, x => x.PackageUomNo.ToUpper(), y => y.UomNo, (x, y) => { x.PackageUomId = y.FirstOrDefault()?.UomId; return x; }).ToList();
                            #endregion

                            #region //撈取品號ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT MtlItemId, MtlItemNo
                                    FROM PDM.MtlItem
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            List<MtlItem> resultMtlitems = sqlConnection.Query<MtlItem>(sql, dynamicParameters).ToList();
                            rtDetails = rtDetails.Join(resultMtlitems, x => x.MtlItemNo, y => y.MtlItemNo, (x, y) => { x.MtlItemId = y.MtlItemId; return x; }).Select(x => x).ToList();
                            #endregion

                            #region //撈取庫別ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT InventoryId, InventoryNo
                                    FROM SCM.Inventory
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            List<Inventory> resultInventories = sqlConnection.Query<Inventory>(sql, dynamicParameters).ToList();
                            rtDetails = rtDetails.Join(resultInventories, x => x.InventoryNo, y => y.InventoryNo, (x, y) => { x.InventoryId = y.InventoryId; return x; }).ToList();
                            #endregion

                            #region //判斷訂單是否有資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.SoDetailId, a.SoId,b.SoErpPrefix,b.SoErpNo, a.SoSequence
                                    FROM SCM.SoDetail a
                                    INNER JOIN SCM.SaleOrder b ON a.SoId = b.SoId
                                    WHERE b.CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<SoDetail> resultSoDetails = sqlConnection.Query<SoDetail>(sql, dynamicParameters).ToList();

                            rtDetails = rtDetails.GroupJoin(resultSoDetails, x => new { x.SoErpPrefix, x.SoErpNo, x.SoSequence }, y => new { y.SoErpPrefix, y.SoErpNo, y.SoSequence }, (x, y) => { x.SoDetailId = y.FirstOrDefault()?.SoDetailId; return x; }).ToList();
                            #endregion

                            #region //判斷銷貨單是否有資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.RoDetailId,b.RoId, b.RoErpPrefix, b.RoErpNo, a.RoSequence
                                    FROM SCM.RoDetail a 
                                    INNER JOIN SCM.ReceiveOrder b ON a.RoId = b.RoId
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            List<RoDetail> resultRoDetail = sqlConnection.Query<RoDetail>(sql, dynamicParameters).ToList();
                            rtDetails = rtDetails.GroupJoin(resultRoDetail, x => new { x.RoErpPrefix, x.RoErpNo, x.RoSequence }, y => new { y.RoErpPrefix, y.RoErpNo, y.RoSequence }, (x, y) => { x.RoDetailId = y.FirstOrDefault()?.RoDetailId; return x; }).Select(x => x).ToList();
                            #endregion

                            #region //判斷銷退單是否有資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT RtId, RtErpPrefix, RtErpNo
                                    FROM SCM.ReturnReceiveOrder
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            List<ReturnReceiveOrder> resultReturnReceiveOrder = sqlConnection.Query<ReturnReceiveOrder>(sql, dynamicParameters).ToList();
                            returnReceiveOrder = returnReceiveOrder.GroupJoin(resultReturnReceiveOrder, x => new { x.RtErpPrefix, x.RtErpNo }, y => new { y.RtErpPrefix, y.RtErpNo }, (x, y) => { x.RtId = y.FirstOrDefault()?.RtId; return x; }).Select(x => x).ToList();
                            #endregion

                            #region //銷貨單單頭(新增/修改)
                            List<ReturnReceiveOrder> addReturnReceiveOrders = returnReceiveOrder.Where(x => x.RtId == null).ToList();
                            List<ReturnReceiveOrder> updateReturnReceiveOrders = returnReceiveOrder.Where(x => x.RtId != null).ToList();

                            #region //新增
                            dynamicParameters = new DynamicParameters();
                            if (addReturnReceiveOrders.Count > 0)
                            {
                                addReturnReceiveOrders
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.CompanyId = CurrentCompany;
                                        x.TransferStatusMES = "Y";
                                        x.CreateDate = CreateDate;
                                        x.LastModifiedDate = LastModifiedDate;
                                        x.CreateBy = CreateBy;
                                        x.LastModifiedBy = LastModifiedBy;
                                    });

                                #region //新增單頭資料表
                                sql = @"INSERT INTO SCM.ReturnReceiveOrder (CompanyId, RtErpPrefix, RtErpNo, ReturnDate, CustomerId, DepartmentId, 
                                        SalesmenId, Factory, Currency, ExchangeRate, OriginalAmount, OriginalTaxAmount, InvoiceType, TaxType, 
                                        InvNumEnd, UiNo, PrintCount, InvoiceDate, UpdateCode, ConfirmStatus, Remark, CustomerFullName, 
                                        CollectionSalesmenId, Remark1, Remark2, Remark3, DeductionDistinction, CustomsClearance, RowCnt, 
                                        TotalQuantity, StaffId, RevenueJournalCode, CostJournalCode, ApplyYYMM, DocDate, ConfirmUserId, TaxRate, 
                                        PretaxAmount, TaxAmount, PaymentTerm, TotalPackages, SignatureStatus, ProcessCode, TransferStatusERP, 
                                        BondCode, TransmissionCount, CustomerAddressFirst, CustomerAddressSecond, ContactPerson, Courier, 
                                        SiteCommCalcMethod, SiteCommRate, CommCalcInclTax, TotalCommAmount, TradingTerms, CustomerEgFullName, 
                                        DebitNote, TaxCode, RemarkCode, MultipleInvoices, InvNumStart, NumberOfInvoices, MultiTaxRate, 
                                        TaxCurrencyRate, VoidDate, VoidTime, VoidRemark, IncomeDraftID, IncomeDraftSeq, IncomeVoucherType, 
                                        IncomeVoucherNumber, Status, GenLedgerAcctType, InvCode2, InvSymbol, ContactEmail, OrigInvNumber, 
                                        TransferStatusMES, CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        VALUES ( @CompanyId, @RtErpPrefix, @RtErpNo, @ReturnDate, @CustomerId, @DepartmentId, 
                                        @SalesmenId, @Factory, @Currency, @ExchangeRate, @OriginalAmount, @OriginalTaxAmount, @InvoiceType, @TaxType, 
                                        @InvNumEnd, @UiNo, @PrintCount, @InvoiceDate, @UpdateCode, @ConfirmStatus, @Remark, @CustomerFullName, 
                                        @CollectionSalesmenId, @Remark1, @Remark2, @Remark3, @DeductionDistinction, @CustomsClearance, @RowCnt, 
                                        @TotalQuantity, @StaffId, @RevenueJournalCode, @CostJournalCode, @ApplyYYMM, @DocDate, @ConfirmUserId, @TaxRate, 
                                        @PretaxAmount, @TaxAmount, @PaymentTerm, @TotalPackages, @SignatureStatus, @ProcessCode, @TransferStatusERP, 
                                        @BondCode, @TransmissionCount, @CustomerAddressFirst, @CustomerAddressSecond, @ContactPerson, @Courier, 
                                        @SiteCommCalcMethod, @SiteCommRate, @CommCalcInclTax, @TotalCommAmount, @TradingTerms, @CustomerEgFullName, 
                                        @DebitNote, @TaxCode, @RemarkCode, @MultipleInvoices, @InvNumStart, @NumberOfInvoices, @MultiTaxRate, 
                                        @TaxCurrencyRate, @VoidDate, @VoidTime, @VoidRemark, @IncomeDraftID, @IncomeDraftSeq, @IncomeVoucherType, 
                                        @IncomeVoucherNumber, @Status, @GenLedgerAcctType, @InvCode2, @InvSymbol, @ContactEmail, @OrigInvNumber, 
                                        @TransferStatusMES, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                int addMain = sqlConnection.Execute(sql, addReturnReceiveOrders);
                                mainAffected += addMain;
                                #endregion

                            }
                            #endregion

                            #region //修改
                            dynamicParameters = new DynamicParameters();
                            if (updateReturnReceiveOrders.Count > 0)
                            {
                                updateReturnReceiveOrders
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.TransferStatusMES = "Y";
                                        x.LastModifiedDate = LastModifiedDate;
                                        x.LastModifiedBy = LastModifiedBy;
                                    });

                                #region //更新單頭資料表
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE SCM.ReturnReceiveOrder SET
                                        ReturnDate			= @ReturnDate,			
                                        CustomerId			= @CustomerId,			
                                        DepartmentId		= @DepartmentId,			
                                        SalesmenId			= @SalesmenId,			
                                        Factory				= @Factory,				
                                        Currency			= @Currency,				
                                        ExchangeRate		= @ExchangeRate,			
                                        OriginalAmount		= @OriginalAmount,		
                                        OriginalTaxAmount	= @OriginalTaxAmount,	
                                        InvoiceType			= @InvoiceType,			
                                        TaxType				= @TaxType,				
                                        InvNumEnd			= @InvNumEnd,			
                                        UiNo				= @UiNo,					
                                        PrintCount			= @PrintCount,			
                                        InvoiceDate			= @InvoiceDate,			
                                        UpdateCode			= @UpdateCode,			
                                        ConfirmStatus		= @ConfirmStatus,		
                                        Remark				= @Remark,				
                                        CustomerFullName	= @CustomerFullName,		
                                        CollectionSalesmenId= @CollectionSalesmenId,	
                                        Remark1				= @Remark1,				
                                        Remark2				= @Remark2,				
                                        Remark3				= @Remark3,				
                                        DeductionDistinction= @DeductionDistinction,	
                                        CustomsClearance	= @CustomsClearance,		
                                        RowCnt				= @RowCnt,				
                                        TotalQuantity		= @TotalQuantity,		
                                        StaffId				= @StaffId,				
                                        RevenueJournalCode	= @RevenueJournalCode,	
                                        CostJournalCode		= @CostJournalCode,		
                                        ApplyYYMM			= @ApplyYYMM,			
                                        DocDate				= @DocDate,				
                                        ConfirmUserId		= @ConfirmUserId,		
                                        TaxRate				= @TaxRate,				
                                        PretaxAmount		= @PretaxAmount,			
                                        TaxAmount			= @TaxAmount,			
                                        PaymentTerm			= @PaymentTerm,			
                                        TotalPackages		= @TotalPackages,		
                                        SignatureStatus		= @SignatureStatus,	
                                        ProcessCode			= @ProcessCode,		
                                        TransferStatusERP	= @TransferStatusERP,	
                                        BondCode			= @BondCode,				
                                        TransmissionCount	= @TransmissionCount,	
                                        CustomerAddressFirst= @CustomerAddressFirst,	
                                        CustomerAddressSecond= @CustomerAddressSecond,
                                        ContactPerson		= @ContactPerson,		
                                        Courier				= @Courier,				
                                        SiteCommCalcMethod	= @SiteCommCalcMethod,	
                                        SiteCommRate		= @SiteCommRate,		
                                        CommCalcInclTax		= @CommCalcInclTax,		
                                        TotalCommAmount		= @TotalCommAmount,		
                                        TradingTerms		= @TradingTerms,			
                                        CustomerEgFullName	= @CustomerEgFullName,	
                                        DebitNote			= @DebitNote,			
                                        TaxCode				= @TaxCode,				
                                        RemarkCode			= @RemarkCode,			
                                        MultipleInvoices	= @MultipleInvoices,		
                                        InvNumStart			= @InvNumStart,			
                                        NumberOfInvoices	= @NumberOfInvoices,		
                                        MultiTaxRate		= @MultiTaxRate,			
                                        TaxCurrencyRate		= @TaxCurrencyRate,		
                                        VoidDate			= @VoidDate,				
                                        VoidTime			= @VoidTime,				
                                        VoidRemark			= @VoidRemark,			
                                        IncomeDraftID		= @IncomeDraftID,		
                                        IncomeDraftSeq		= @IncomeDraftSeq,		
                                        IncomeVoucherType	= @IncomeVoucherType,	
                                        IncomeVoucherNumber	= @IncomeVoucherNumber,
                                        Status			    = @Status,				
                                        GenLedgerAcctType	= @GenLedgerAcctType,	
                                        InvCode2			= @InvCode2,			
                                        InvSymbol			= @InvSymbol,			
                                        ContactEmail		= @ContactEmail,			
                                        OrigInvNumber		= @OrigInvNumber,		
                                        TransferStatusMES	= @TransferStatusMES,	
                                        LastModifiedDate	= @LastModifiedDate,		
                                        LastModifiedBy		= @LastModifiedBy	
                                        WHERE RtId = @RtId";
                                int updateMain = sqlConnection.Execute(sql, updateReturnReceiveOrders);
                                mainAffected += updateMain;
                                #endregion
                            }
                            #endregion
                            #endregion

                            #region //銷貨單單身(新增/修改)
                            #region //撈取銷退單單頭ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT RtId, RtErpPrefix, RtErpNo
                                    FROM SCM.ReturnReceiveOrder
                                    WHERE CompanyId = @CompanyId
                                    ORDER BY RtId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            resultReturnReceiveOrder = sqlConnection.Query<ReturnReceiveOrder>(sql, dynamicParameters).ToList();

                            rtDetails = rtDetails.Join(resultReturnReceiveOrder, x => new { x.RtErpPrefix, x.RtErpNo }, y => new { y.RtErpPrefix, y.RtErpNo }, (x, y) => { x.RtId = y.RtId; return x; }).Select(x => x).ToList();
                            #endregion

                            #region //判斷SCM.RtDetail是否有資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.RtDetailId, a.RtId, a.RtSequence
                                    FROM SCM.RtDetail a
                                    INNER JOIN SCM.ReturnReceiveOrder b ON a.RtId = b.RtId
                                    WHERE b.CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<RtDetail> resultRtDetails = sqlConnection.Query<RtDetail>(sql, dynamicParameters).ToList();

                            rtDetails = rtDetails.GroupJoin(resultRtDetails, x => new { x.RtId, x.RtSequence }, y => new { y.RtId, y.RtSequence }, (x, y) => { x.RtDetailId = y.FirstOrDefault()?.RtDetailId; return x; }).ToList();
                            #endregion

                            List<RtDetail> addRtDetails = rtDetails.Where(x => x.RtDetailId == null).ToList();
                            List<RtDetail> updateRtDetails = rtDetails.Where(x => x.RtDetailId != null).ToList();

                            #region //新增
                            dynamicParameters = new DynamicParameters();
                            if (addRtDetails.Count > 0)
                            {
                                addRtDetails
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.TransferStatusMES = "Y";
                                        x.CreateDate = CreateDate;
                                        x.LastModifiedDate = LastModifiedDate;
                                        x.CreateBy = CreateBy;
                                        x.LastModifiedBy = LastModifiedBy;
                                    });
                                sql = @"INSERT INTO SCM.RtDetail (RtId, RtSequence, MtlItemId, Quantity, UomId, UnitPrice, Amount, InventoryId, LotNumber, 
                                        RoDetailId, SoDetailId, ConfirmStatus, UpdateCode, Remarks, CheckOutCode, CheckOutPrefix, CheckOutNo, 
                                        CheckOutSeq, ProjectCode, CustomerMtlItemNo, Type, OrigPreTaxAmt, OrigTaxAmt, PreTaxAmt, TaxAmt, PackageQty, 
                                        PackageUomId, BondedCode, CommissionRate, CommissionAmount, OriginalCustomer, FreeSpareType, 
                                        FreeSpareQty, PackageFreeSpareQty, NotPayTemp, Location, AvailableQty, AvailableUomId, TaxRate, TaxCode, 
                                        DiscountAmount, TransferStatusMES, CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        VALUES ( @RtId, @RtSequence, @MtlItemId, @Quantity, @UomId, @UnitPrice, @Amount, @InventoryId, @LotNumber, 
                                        @RoDetailId, @SoDetailId, @ConfirmStatus, @UpdateCode, @Remarks, @CheckOutCode, @CheckOutPrefix, @CheckOutNo, 
                                        @CheckOutSeq, @ProjectCode, @CustomerMtlItemNo, @Type, @OrigPreTaxAmt, @OrigTaxAmt, @PreTaxAmt, @TaxAmt, @PackageQty, 
                                        @PackageUomId, @BondedCode, @CommissionRate, @CommissionAmount, @OriginalCustomer, @FreeSpareType, 
                                        @FreeSpareQty, @PackageFreeSpareQty, @NotPayTemp, @Location, @AvailableQty, @AvailableUomId, @TaxRate, @TaxCode, 
                                        @DiscountAmount, @TransferStatusMES, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy
                                        )";
                                int addDetail = sqlConnection.Execute(sql, addRtDetails);
                                detailAffected += addDetail;
                            }
                            #endregion

                            #region //修改
                            dynamicParameters = new DynamicParameters();
                            if (updateRtDetails.Count > 0)
                            {
                                updateRtDetails
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.TransferStatusMES = "Y";
                                        x.LastModifiedDate = LastModifiedDate;
                                        x.LastModifiedBy = LastModifiedBy;
                                    });

                                sql = @"UPDATE SCM.RtDetail SET
                                        RtSequence = @RtSequence, 
                                        MtlItemId = @MtlItemId, 
                                        Quantity = @Quantity, 
                                        UomId = @UomId, 
                                        UnitPrice = @UnitPrice, 
                                        Amount = @Amount, 
                                        InventoryId = @InventoryId, 
                                        LotNumber = @LotNumber, 
                                        RoDetailId = @RoDetailId, 
                                        SoDetailId = @SoDetailId, 
                                        ConfirmStatus = @ConfirmStatus, 
                                        UpdateCode = @UpdateCode, 
                                        Remarks = @Remarks, 
                                        CheckOutCode = @CheckOutCode, 
                                        CheckOutPrefix = @CheckOutPrefix, 
                                        CheckOutNo = @CheckOutNo, 
                                        CheckOutSeq = @CheckOutSeq, 
                                        ProjectCode = @ProjectCode, 
                                        CustomerMtlItemNo = @CustomerMtlItemNo, 
                                        Type = @Type, 
                                        OrigPreTaxAmt = @OrigPreTaxAmt, 
                                        OrigTaxAmt = @OrigTaxAmt, 
                                        PreTaxAmt = @PreTaxAmt, 
                                        TaxAmt = @TaxAmt, 
                                        PackageQty = @PackageQty, 
                                        PackageUomId = @PackageUomId, 
                                        BondedCode = @BondedCode, 
                                        CommissionRate = @CommissionRate, 
                                        CommissionAmount = @CommissionAmount, 
                                        OriginalCustomer = @OriginalCustomer, 
                                        FreeSpareType = @FreeSpareType, 
                                        FreeSpareQty = @FreeSpareQty, 
                                        PackageFreeSpareQty = @PackageFreeSpareQty, 
                                        NotPayTemp = @NotPayTemp, 
                                        Location = @Location, 
                                        AvailableQty = @AvailableQty, 
                                        AvailableUomId = @AvailableUomId, 
                                        TaxRate = @TaxRate, 
                                        TaxCode = @TaxCode, 
                                        DiscountAmount = @DiscountAmount, 
                                        TransferStatusMES = @TransferStatusMES, 
                                        LastModifiedDate = @LastModifiedDate, 
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE RoDetailId = @RoDetailId";
                                int updateDetail = sqlConnection.Execute(sql, updateRtDetails);
                                detailAffected += updateDetail;
                            }
                            #endregion
                            #endregion
                        }
                    }
                    #endregion

                    if (TranSync == "Y")
                    {
                        using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //ERP銷退單單頭資料
                            sql = @"SELECT TI001 RtErpPrefix, TI002 RtErpNo
                                    FROM COPTI
                                    WHERE 1=1
                                    ORDER BY TI001, TI002";
                            var resultErpRt = erpConnection.Query<ReturnReceiveOrder>(sql, dynamicParameters);
                            #endregion

                            using (SqlConnection bmConnection = new SqlConnection(MainConnectionStrings))
                            {
                                #region //BM銷退單單頭資料
                                sql = @"SELECT RtErpPrefix, RtErpNo
                                        FROM SCM.ReturnReceiveOrder
                                        WHERE CompanyId = @CompanyId
                                        ORDER BY RtErpPrefix, RtErpNo";
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                var resultBmRt = bmConnection.Query<ReturnReceiveOrder>(sql, dynamicParameters);
                                #endregion

                                #region //篩選ERP不存在但BM還存在的訂單單頭
                                var dictionaryErpRt = resultErpRt.ToDictionary(x => x.RtErpPrefix + "-" + x.RtErpNo, x => x.RtErpPrefix + "-" + x.RtErpNo);
                                var dictionaryBmRt = resultBmRt.ToDictionary(x => x.RtErpPrefix + "-" + x.RtErpNo, x => x.RtErpPrefix + "-" + x.RtErpNo);
                                var changeRo = dictionaryBmRt.Where(x => !dictionaryErpRt.ContainsKey(x.Key)).ToList();
                                var changeRoList = changeRo.Select(x => x.Value).ToList();
                                #endregion

                                #region //刪除異動訂單單頭
                                if (changeRoList.Count > 0)
                                {
                                    #region //單身刪除
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE SCM.RtDetail
                                            WHERE RoId IN (
                                                SELECT RtId
                                                FROM SCM.ReturnReceiveOrder
                                                WHERE RtErpPrefix + '-' + RtErpNo IN @RtErpFullNo
                                            )";
                                    dynamicParameters.Add("RtErpFullNo", changeRoList.Select(x => x).ToArray());
                                    mainDelAffected += bmConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //單頭刪除
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE SCM.ReturnReceiveOrder
                                            WHERE RtErpPrefix + '-' + RtErpNo IN @RtErpFullNo";
                                    dynamicParameters.Add("RtErpFullNo", changeRoList.Select(x => x).ToArray());
                                    detailDelAffected += bmConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                                #endregion


                            }

                            #region //ERP銷退單身資料
                            sql = @"SELECT TJ001 RtErpPrefix, TJ002 RtErpNo, TJ003 RtSequence
                                    FROM COPTJ
                                    WHERE 1=1
                                    ORDER BY TJ001, TJ002, TJ003";
                            var resultErpRtDetail = erpConnection.Query<RtDetail>(sql, dynamicParameters);
                            #endregion

                            using (SqlConnection bmConnection = new SqlConnection(MainConnectionStrings))
                            {
                                #region //BM銷退單單身資料
                                sql = @"SELECT b.RtErpPrefix, b.RtErpNo, a.RtSequence
                                        FROM SCM.RtDetail a
                                        INNER JOIN SCM.ReturnReceiveOrder b ON a.RtId = b.RtId
                                        WHERE b.CompanyId = @CompanyId
                                        ORDER BY b.RtErpPrefix, b.RtErpNo, a.RtSequence";
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                var resultBmRtDetail = bmConnection.Query<RtDetail>(sql, dynamicParameters);
                                #endregion

                                #region //篩選ERP不存在但BM還存在的銷退單單身
                                var dictionaryErpRtDetail = resultErpRtDetail.ToDictionary(x => x.RtErpPrefix + "-" + x.RtErpNo + "-" + x.RtSequence, x => x.RtErpPrefix + "-" + x.RtErpNo + "-" + x.RtSequence);
                                var dictionaryBmRtDetail = resultBmRtDetail.ToDictionary(x => x.RtErpPrefix + "-" + x.RtErpNo + "-" + x.RtSequence, x => x.RtErpPrefix + "-" + x.RtErpNo + "-" + x.RtSequence);
                                var changeRoDetail = dictionaryBmRtDetail.Where(x => !dictionaryErpRtDetail.ContainsKey(x.Key)).ToList();
                                var changeRoDetailList = changeRoDetail.Select(x => x.Value).ToList();
                                #endregion

                                #region //刪除訂單單身
                                if (changeRoDetailList.Count > 0)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE a
                                            FROM SCM.RtDetail a
                                            INNER JOIN SCM.ReturnReceiveOrder b ON a.RtId = b.RtId
                                            WHERE b.RtErpPrefix + '-' + b.RtErpNo + '-' + RIGHT('0000' + a.Sequence, 4) IN @RtErpFullNo";
                                    dynamicParameters.Add("RtErpFullNo", changeRoDetailList.Select(x => x).ToArray());
                                    detailDelAffected += bmConnection.Execute(sql, dynamicParameters, null, 300);
                                }
                                #endregion
                            }
                        }
                    }

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

        #region //UpdateReturnReceiveOrder -- 銷退單更新 -- Shintokuro 2024.05.10
        public string UpdateReturnReceiveOrder(int RtId, int ViewCompanyId
            , string DocDate
            , string ReturnDate, string RtErpPrefix, string RtErpNo, int CustomerId, string ProcessCode
            , string RevenueJournalCode, string CostJournalCode
            , double OriginalAmount
            , double OriginalTaxAmount, double PretaxAmount, double TaxAmount, double TotalQuantity
            , double TaxCurrencyRate
            , int DepartmentId
            , int SalesmenId, int StaffId, int CollectionSalesmenId, string Factory, string PaymentTerm, string TradingTerms
            , string Currency, double ExchangeRate, string BondCode, int RowCnt, string Remark
            , string TaxCode
            , string InvoiceType, string TaxType, double TaxRate, string UiNo, string InvoiceDate, string InvNumStart, string InvNumEnd
            , string ApplyYYMM, string CustomsClearance
            , string CustomerFullName, string CustomerEgFullName, string MultipleInvoices, string MultiTaxRate, string DebitNote
            , string ContactPerson
            , string ContactEmail, string CustomerAddressFirst, string CustomerAddressSecond, string Remark1, string Remark2, string Remark3
            , string RtDetailData
            )
        {
            try
            {
                if (RtErpPrefix.Length <= 0) throw new SystemException("【銷退單別】不能為空!");
                if (ReturnDate.Length <= 0) throw new SystemException("【銷退日期】不能為空!");
                if (CustomerId <= 0) throw new SystemException("【客戶】不能為空!");
                if (Factory.Length <= 0) throw new SystemException("【廠別】不能為空!");
                if (Currency.Length <= 0) throw new SystemException("【幣別】不能為空!");
                if (ExchangeRate < 0) throw new SystemException("【匯率】不能為負!");
                if (ApplyYYMM.Length < 0) throw new SystemException("【申報年月】不能為空!");
                if (DocDate.Length < 0) throw new SystemException("【單據日期】不能為空!");
                if (TaxRate < 0) throw new SystemException("【營業稅率】不能為負!");
                ApplyYYMM = ApplyYYMM.Replace("-", "");

                List<Dictionary<string, string>> RtDetailJsonList = JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(RtDetailData);
                int? nullData = null;
                int rowsAffected = 0;
                string SoErpPrefix = "";
                string SoErpNo = "";
                string SoSequence = "";
                string TransferStatusMES = "";
                string RoErpPrefix = "";
                string RoErpNo = "";
                string userNo = "";
                string MESCompanyNo = "";
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷公司資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 CompanyId, CompanyName, ErpDb, ErpNo
                                FROM BAS.Company
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【公司】資料錯誤!");

                        foreach (var item in result)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            MESCompanyNo = item.ErpNo;
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
                        dynamicParameters.Add("FunctionCode", "ReturnReceiveOrderManagment");
                        dynamicParameters.Add("DetailCode", "update");

                        var resultUser = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUser.Count() <= 0) throw new SystemException("使用者資料錯誤!");

                        foreach (var item in resultUser)
                        {
                            userNo = item.UserNo;
                        }
                        #endregion

                        #region //撈取銷退單單頭資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CompanyId,　a.TransferStatusMES,  a.ConfirmStatus, a.RtErpPrefix, a.RtErpNo
                                FROM SCM.ReturnReceiveOrder a
                                WHERE a.RtId = @RtId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("RtId", RtId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("銷退單資料找不到，請重新確認");
                        foreach (var item in result)
                        {
                            if (item.TransferStatusMES == "Y") throw new SystemException("銷退單已經拋轉ERP，不可以修改");
                            if (item.ConfirmStatus == "Y") throw new SystemException("銷退單已經核單，不可以修改");
                            if (item.CompanyId != ViewCompanyId) throw new SystemException("頁面的公司別與單據公司別不同，請嘗試重新登入!!");
                            TransferStatusMES = item.TransferStatusMES;
                            RtErpPrefix = item.RtErpPrefix;
                            RtErpNo = item.RtErpNo;
                        }
                        #endregion

                        #region //判斷客戶資料是否存在
                        sql = @"SELECT TOP 1 1
                                FROM SCM.Customer
                                WHERE CompanyId = @CompanyId
                                AND CustomerId = @CustomerId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("CustomerId", CustomerId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("客戶資料找不到，請重新確認!!");
                        #endregion

                        #region //判斷部門資料是否存在
                        if (DepartmentId > 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.Department
                                    WHERE CompanyId = @CompanyId
                                    AND DepartmentId = @DepartmentId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("DepartmentId", DepartmentId);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("部門資料找不到，請重新確認!!");
                        }
                        #endregion

                        #region //判斷業務資料是否存在
                        if (SalesmenId > 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.[User]
                                    WHERE UserId = @UserId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("UserId", SalesmenId);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("業務資料找不到，請重新確認!!");
                        }
                        #endregion

                        #region //判斷收款業務資料是否存在
                        if (CollectionSalesmenId > 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.[User]
                                    WHERE UserId = @UserId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("UserId", CollectionSalesmenId);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("收款業務資料找不到，請重新確認!!");
                        }
                        #endregion

                        #region //判斷員工資料是否存在
                        if (StaffId > 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.[User]
                                    WHERE UserId = @UserId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("UserId", StaffId);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("員工資料找不到，請重新確認!!");
                        }
                        #endregion

                        #region //更新單頭資料表
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.ReturnReceiveOrder SET
                                DocDate = @DocDate,
                                ReturnDate = @ReturnDate,
                                CustomerId = @CustomerId,
                                ProcessCode = @ProcessCode,
                                OriginalAmount = @OriginalAmount,
                                OriginalTaxAmount = @OriginalTaxAmount,
                                PretaxAmount = @PretaxAmount,
                                TaxAmount = @TaxAmount,
                                TotalQuantity = @TotalQuantity,
                                DepartmentId = @DepartmentId,
                                SalesmenId = @SalesmenId,
                                StaffId = @StaffId,
                                CollectionSalesmenId = @CollectionSalesmenId,
                                Factory = @Factory,
                                PaymentTerm = @PaymentTerm,
                                TradingTerms = @TradingTerms,
                                Currency = @Currency,
                                ExchangeRate = @ExchangeRate,
                                BondCode = @BondCode,
                                RowCnt = @RowCnt,
                                Remark = @Remark,
                                TaxCode = @TaxCode,
                                InvoiceType = @InvoiceType,
                                TaxType = @TaxType,
                                TaxRate = @TaxRate,
                                UiNo = @UiNo,
                                InvoiceDate = @InvoiceDate,
                                InvNumStart = @InvNumStart,
                                InvNumEnd = @InvNumEnd,
                                ApplyYYMM = @ApplyYYMM,
                                CustomsClearance = @CustomsClearance,
                                CustomerFullName = @CustomerFullName,
                                CustomerEgFullName = @CustomerEgFullName,
                                MultipleInvoices = @MultipleInvoices,
                                MultiTaxRate = @MultiTaxRate,
                                DebitNote = @DebitNote,
                                ContactPerson = @ContactPerson,
                                ContactEmail = @ContactEmail,
                                CustomerAddressFirst = @CustomerAddressFirst,
                                CustomerAddressSecond = @CustomerAddressSecond,
                                Remark1 = @Remark1,
                                Remark2 = @Remark2,
                                Remark3 = @Remark3,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RtId = @RtId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                DocDate,
                                ReturnDate,
                                CustomerId,
                                ProcessCode,
                                OriginalAmount,
                                OriginalTaxAmount,
                                PretaxAmount,
                                TaxAmount,
                                TotalQuantity,
                                DepartmentId,
                                SalesmenId,
                                StaffId,
                                CollectionSalesmenId,
                                Factory,
                                PaymentTerm,
                                TradingTerms,
                                Currency,
                                ExchangeRate,
                                BondCode,
                                RowCnt,
                                Remark,
                                TaxCode,
                                InvoiceType,
                                TaxType,
                                TaxRate,
                                UiNo,
                                InvoiceDate,
                                InvNumStart,
                                InvNumEnd,
                                ApplyYYMM,
                                CustomsClearance,
                                CustomerFullName,
                                CustomerEgFullName,
                                MultipleInvoices,
                                MultiTaxRate,
                                DebitNote,
                                ContactPerson,
                                ContactEmail,
                                CustomerAddressFirst,
                                CustomerAddressSecond,
                                Remark1,
                                Remark2,
                                Remark3,
                                LastModifiedDate,
                                LastModifiedBy,
                                RtId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);


                        rowsAffected += insertResult.Count();
                        #endregion

                        #region //清空銷貨單單身
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.RtDetail
                                WHERE RtId = @RtId";
                        dynamicParameters.Add("RtId", RtId);
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        if (RtDetailJsonList.Count() > 0)
                        {
                            foreach (var item in RtDetailJsonList)
                            {
                                string MtlItemNo = item["MtlItemNo"] != null ? item["MtlItemNo"].ToString() : throw new SystemException("【資料維護不完整】品號欄位資料不可以為空,請重新確認~~");
                                string Type = item["Type"] != "" ? item["Type"].ToString().Substring(0, 1) : throw new SystemException("【資料維護不完整】類別欄位資料不可以為空,請重新確認~~");
                                Double Quantity = item["Quantity"] != null ? Convert.ToDouble(item["Quantity"].ToString()) : 0;
                                string FreeSpareType = item["Type"] != null ? item["FreeSpareType"].ToString() : throw new SystemException("【資料維護不完整】贈備品類別欄位資料不可以為空,請重新確認~~");
                                Double FreeSpareQty = item["FreeSpareQty"] != null ? Convert.ToDouble(item["FreeSpareQty"].ToString()) : 0;
                                string UomNo = item["UomNo"] != null ? item["UomNo"].ToString() : throw new SystemException("【資料維護不完整】單位欄位資料不可以為空,請重新確認~~");
                                Double AvailableQty = item["AvailableQty"] != null ? Convert.ToDouble(item["AvailableQty"].ToString()) : 0;
                                string AvailableUomNo = item["AvailableUomNo"] != null ? item["AvailableUomNo"].ToString() : throw new SystemException("【資料維護不完整】計價單位欄位資料不可以為空,請重新確認~~");
                                Double UnitPrice = item["UnitPrice"] != null ? Convert.ToDouble(item["UnitPrice"].ToString()) : 0;
                                Double TaxRateDetail = item["TaxRate"] != null ? Convert.ToDouble(item["TaxRate"].ToString()) : 0;
                                Double RoDiscountRate = item["RoDiscountRate"] != "" ? Convert.ToDouble(item["RoDiscountRate"].ToString()) : 0;
                                Double SoDiscountRate = item["SoDiscountRate"] != "" ? Convert.ToDouble(item["SoDiscountRate"].ToString()) : 0;
                                Double Amount = item["Amount"] != null ? Convert.ToDouble(item["Amount"].ToString()) : 0;
                                Double OrigPreTaxAmt = item["OrigPreTaxAmt"] != null ? Convert.ToDouble(item["OrigPreTaxAmt"].ToString()) : 0;
                                Double OrigTaxAmt = item["OrigTaxAmt"] != null ? Convert.ToDouble(item["OrigTaxAmt"].ToString()) : 0;
                                Double PreTaxAmt = item["PreTaxAmt"] != null ? Convert.ToDouble(item["PreTaxAmt"].ToString()) : 0;
                                Double TaxAmt = item["TaxAmt"] != null ? Convert.ToDouble(item["TaxAmt"].ToString()) : 0;
                                string InventoryNo = item["InventoryNo"] != null ? item["InventoryNo"].ToString() : throw new SystemException("【資料維護不完整】庫別欄位資料不可以為空,請重新確認~~");
                                string Location = item["Location"] != null ? item["Location"].ToString() : "";
                                string SaveLocation = item["SaveLocation"] != null ? item["SaveLocation"].ToString() : "";
                                string LotNumber = item["LotNumber"] != null ? item["LotNumber"].ToString() : "";
                                string RoFullNo = item["RoFullNo"] != null ? item["RoFullNo"].ToString() : "";
                                string SoFullNo = item["SoFullNo"] != null ? item["SoFullNo"].ToString() : "";
                                string CustomerMtlItemNo = item["CustomerMtlItemNo"] != null ? item["CustomerMtlItemNo"].ToString() : "";
                                string ProjectCode = item["ProjectCode"] != null ? item["ProjectCode"].ToString() : "";
                                string Remarks = item["Remarks"] != null ? item["Remarks"].ToString() : "";
                                string BondedCode = item["BondedCode"] != null ? item["BondedCode"].ToString() : "";
                                string ReturnNo = item["ReturnNo"] != null ? item["ReturnNo"].ToString() : "";
                                string ReturnRemarks = item["ReturnRemarks"] != null ? item["ReturnRemarks"].ToString() : "";

                                #region //匯率等於1時判定
                                if (ExchangeRate == 1)
                                {
                                    if (OrigPreTaxAmt != PreTaxAmt) throw new SystemException("單據幣別與本國貨幣相同時,原幣金額與本幣金額須相同!!");
                                    if (OrigTaxAmt != TaxAmt) throw new SystemException("單據幣別與本國貨幣相同時,原幣稅額與本幣稅額須相同!!");
                                }
                                #endregion

                                #region //判斷品號是否存在
                                int MtlItemId = -1;
                                string OverDeliveryManagement = "";
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 MtlItemId,OverDeliveryManagement,LotManagement
                                        FROM PDM.MtlItem
                                        WHERE CompanyId = @CompanyId
                                        AND MtlItemNo = @MtlItemNo";
                                dynamicParameters.Add("MtlItemNo", MtlItemNo);
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("【品號:" + MtlItemNo + "】找不到，請重新輸入!");
                                foreach (var item2 in result)
                                {
                                    MtlItemId = item2.MtlItemId;
                                    OverDeliveryManagement = item2.OverDeliveryManagement;
                                    if (item2.LotManagement != "N" && LotNumber == "") throw new SystemException("【品號:" + MtlItemNo + "】須維護批號，請重新輸入!");
                                }
                                #endregion

                                #region //判斷庫別是否存在
                                int InventoryId = -1;
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 InventoryId
                                        FROM SCM.Inventory
                                        WHERE CompanyId = @CompanyId
                                        AND InventoryNo = @InventoryNo";
                                dynamicParameters.Add("InventoryNo", InventoryNo.Split(':')[0]);
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("【庫別:" + InventoryNo + "】找不到，請重新輸入!");
                                foreach (var item2 in result)
                                {
                                    InventoryId = item2.InventoryId;
                                }
                                #endregion

                                #region //判斷單位是否存在
                                int UomId = -1;
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 UomId
                                        FROM PDM.UnitOfMeasure
                                        WHERE CompanyId = @CompanyId
                                        AND UomNo = @UomNo";
                                dynamicParameters.Add("UomNo", UomNo);
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("【單位:" + UomNo + "】找不到，請重新輸入!");
                                foreach (var item2 in result)
                                {
                                    UomId = item2.UomId;
                                }
                                #endregion

                                #region //判斷計價單位是否存在
                                int AvailableUomId = -1;
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 UomId
                                        FROM PDM.UnitOfMeasure
                                        WHERE CompanyId = @CompanyId
                                        AND UomNo = @UomNo";
                                dynamicParameters.Add("UomNo", AvailableUomNo);
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("【計價單位:" + AvailableUomNo + "】找不到，請重新輸入!");
                                foreach (var item2 in result)
                                {
                                    AvailableUomId = item2.UomId;
                                }
                                #endregion

                                #region //判斷訂單單身是否存在
                                int? SoDetailId = null;
                                string ProductType = "";
                                Double FreebieQty = 0;
                                Double SpareQty = 0;
                                if (SoFullNo.Length > 0)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 a.SoDetailId, a.ProductType, a.FreebieQty, a.SpareQty
                                            ,b.SoErpPrefix, b.SoErpNo, a.SoSequence, a.ClosureStatus
                                            FROM SCM.SoDetail a
                                            INNER JOIN SCM.SaleOrder b on a.SoId = b.SoId
                                            WHERE b.CompanyId = @CompanyId
                                            AND b.SoErpPrefix = @SoErpPrefix
                                            AND b.SoErpNo = @SoErpNo
                                            AND a.SoSequence = @SoSequence";
                                    dynamicParameters.Add("CompanyId", CurrentCompany);
                                    dynamicParameters.Add("SoErpPrefix", SoFullNo.Split('-')[0]);
                                    dynamicParameters.Add("SoErpNo", SoFullNo.Split('-')[1]);
                                    dynamicParameters.Add("SoSequence", SoFullNo.Split('-')[2]);
                                    result = sqlConnection.Query(sql, dynamicParameters);
                                    if (result.Count() <= 0) throw new SystemException("【訂單:" + SoFullNo + "】找不到，請重新輸入!");
                                    foreach (var item2 in result)
                                    {
                                        if (item2.ClosureStatus == "Y") throw new SystemException("【訂單:" + SoFullNo + "】已經結案!!!");
                                        SoErpPrefix = item2.SoErpPrefix;
                                        SoErpNo = item2.SoErpNo;
                                        SoSequence = item2.SoSequence;
                                        SoDetailId = item2.SoDetailId;
                                        ProductType = item2.ProductType;
                                        FreebieQty = item2.FreebieQty;
                                        SpareQty = item2.SpareQty;

                                    }
                                }
                                #endregion

                                #region //判斷銷貨單身是否存在
                                int? RoDetailId = null;
                                string RoProductType = "";
                                Double FreebieOrSpareQty = 0;
                                if (RoFullNo.Length > 0)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 b.RoDetailId,b.Type,b.FreeSpareQty,b.CheckOutCode
                                            FROM SCM.ReceiveOrder a
                                            INNER JOIN SCM.RoDetail b on a.RoId = b.RoId
                                            WHERE CompanyId = @CompanyId
                                            AND a.RoErpPrefix = @RoErpPrefix
                                            AND a.RoErpNo = @RoErpNo
                                            AND b.RoSequence = @RoSequence
                                            AND a.ConfirmStatus = 'Y'";
                                    dynamicParameters.Add("CompanyId", CurrentCompany);
                                    dynamicParameters.Add("RoErpPrefix", RoFullNo.Split('-')[0]);
                                    dynamicParameters.Add("RoErpNo", RoFullNo.Split('-')[1]);
                                    dynamicParameters.Add("RoSequence", RoFullNo.Split('-')[2]);
                                    result = sqlConnection.Query(sql, dynamicParameters);
                                    if (result.Count() <= 0) throw new SystemException("【銷貨單:" + RoFullNo + "】找不到，請重新輸入!");
                                    foreach (var item2 in result)
                                    {
                                        if (item2.CheckOutCode == "Y") throw new SystemException("【銷貨單:" + RoFullNo + "】已經結案!!!");
                                        RoDetailId = item2.RoDetailId;
                                        RoProductType = item2.Type;
                                        FreebieOrSpareQty = item2.FreeSpareQty;
                                    }
                                }
                                #endregion

                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO SCM.RtDetail (RtId, RtSequence, MtlItemId, Quantity, UomId, UnitPrice, Amount, InventoryId, LotNumber, 
                                        RoDetailId, SoDetailId, ConfirmStatus, UpdateCode, Remarks, CheckOutCode, CheckOutPrefix, CheckOutNo, 
                                        CheckOutSeq, ProjectCode, CustomerMtlItemNo, Type, OrigPreTaxAmt, OrigTaxAmt, PreTaxAmt, TaxAmt, PackageQty, 
                                        PackageUomId, BondedCode, CommissionRate, CommissionAmount, OriginalCustomer, FreeSpareType, 
                                        FreeSpareQty, PackageFreeSpareQty, NotPayTemp, Location, AvailableQty, AvailableUomId, TaxRate, TaxCode, 
                                        DiscountAmount, TransferStatusMES, CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        VALUES ( @RtId, @RtSequence, @MtlItemId, @Quantity, @UomId, @UnitPrice, @Amount, @InventoryId, @LotNumber, 
                                        @RoDetailId, @SoDetailId, @ConfirmStatus, @UpdateCode, @Remarks, @CheckOutCode, @CheckOutPrefix, @CheckOutNo, 
                                        @CheckOutSeq, @ProjectCode, @CustomerMtlItemNo, @Type, @OrigPreTaxAmt, @OrigTaxAmt, @PreTaxAmt, @TaxAmt, @PackageQty, 
                                        @PackageUomId, @BondedCode, @CommissionRate, @CommissionAmount, @OriginalCustomer, @FreeSpareType, 
                                        @FreeSpareQty, @PackageFreeSpareQty, @NotPayTemp, @Location, @AvailableQty, @AvailableUomId, @TaxRate, @TaxCode, 
                                        @DiscountAmount, @TransferStatusMES, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy
                                        )";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        RtId,
                                        RtSequence = "",
                                        MtlItemId,
                                        Quantity,
                                        UomId,
                                        UnitPrice,
                                        Amount,
                                        InventoryId,
                                        LotNumber,
                                        RoDetailId,
                                        SoDetailId,
                                        ConfirmStatus = "N",
                                        UpdateCode = "N",
                                        Remarks,
                                        CheckOutCode = "N",
                                        CheckOutPrefix = "",
                                        CheckOutNo = "",
                                        CheckOutSeq = "",
                                        ProjectCode,
                                        CustomerMtlItemNo,
                                        Type,
                                        OrigPreTaxAmt,
                                        OrigTaxAmt,
                                        PreTaxAmt,
                                        TaxAmt,
                                        PackageQty = 0,
                                        PackageUomId = "",
                                        BondedCode,
                                        CommissionRate = 0,
                                        CommissionAmount = 0,
                                        OriginalCustomer = "",
                                        FreeSpareType = RoProductType,
                                        FreeSpareQty,
                                        PackageFreeSpareQty = 0,
                                        NotPayTemp = "N",
                                        Location,
                                        AvailableQty,
                                        AvailableUomId,
                                        TaxRate,
                                        TaxCode,
                                        DiscountAmount = 0,
                                        TransferStatusMES = "N",
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();
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

                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        #region //ERP公司代碼取得
                        string ERPCompanyNo = "";
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 ML001  
                                         FROM CMSML";
                        var resultCompanyNo = sqlConnection.Query(sql, dynamicParameters);
                        foreach (var item in resultCompanyNo)
                        {
                            ERPCompanyNo = item.ML001;
                        }
                        if (MESCompanyNo != ERPCompanyNo) throw new SystemException("ERP的公司代碼與MES的公司代碼不同，請嘗試重新登入!!");
                        #endregion

                        #region //判斷ERP使用者是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 MF005
                                FROM ADMMF
                                WHERE COMPANY = @CompanyNo
                                AND MF001 = @userNo";
                        dynamicParameters.Add("CompanyNo", ERPCompanyNo);
                        dynamicParameters.Add("userNo", userNo);
                        var resultUserExist = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUserExist.Count() <= 0) throw new SystemException("【ERP使用者】不存在!");
                        #endregion

                        #region //審核ERP權限-建立
                        var USR_GROUP = BaseHelper.CheckErpAuthority(userNo, sqlConnection, "COPI08", "CREATE");
                        #endregion

                        #region //判斷開立日期是否超過結帳日
                        string CloseDateBase = "";
                        DateTime date = DateTime.Parse(ReturnDate);
                        string ReceiveDateBase = date.ToString("yyyyMMdd");
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
                                throw new SystemException("【銷貨單】ERP已經結帳,不可以在新增【" + CloseDateBase + "】之前的單據");
                            }
                        }
                        else
                        {
                            throw new SystemException("日期字符串无效，无法比较");
                        }
                        #endregion


                        foreach (var item in RtDetailJsonList)
                        {
                            string MtlItemNo = item["MtlItemNo"];
                            string SoFullDetail = item["SoFullNo"] != null ? item["SoFullNo"].ToString() : "";
                            Double Quantity = item["Quantity"] != null ? Convert.ToDouble(item["Quantity"].ToString()) : 0;
                            Double FreeSpareQty = item["FreeSpareQty"] != null ? Convert.ToDouble(item["FreeSpareQty"].ToString()) : 0;
                            Double AvailableQty = item["AvailableQty"] != null ? Convert.ToDouble(item["AvailableQty"].ToString()) : 0;
                            SoErpPrefix = SoFullDetail.Split('-')[0];
                            SoErpNo = SoFullDetail.Split('-')[1];
                            SoSequence = SoFullDetail.Split('-')[2];

                            #region //撈取訂單品號的數量+贈/備品數量
                            string ProductType = "";
                            Double CanUseQty = 0;
                            Double CanUseFreebieQty = 0;
                            Double CanUseSpareQty = 0;
                            Double CanUseAvailableQty = 0;
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.TD003 SoSequence, a.TD049 ProductType
                                            ,(ISNULL(a.TD008,0) -  ISNULL(a.TD009,0) - ISNULL(x.UnconfirmedQty,0)) CanUseQty 
                                            ,(ISNULL(a.TD024,0) -  ISNULL(a.TD025,0) - ISNULL(x.UnconfirmedFreeSpareQty,0)) CanUseFreebieQty 
                                            ,(ISNULL(a.TD050,0) -  ISNULL(a.TD051,0) - ISNULL(x.UnconfirmedFreeSpareQty,0)) CanUseSpareQty 
                                            ,(ISNULL(a.TD076,0) -  ISNULL(a.TD078,0) - ISNULL(x.UnconfirmedAvailableQty,0)) CanUseAvailableQty
                                            ,ISNULL(x.UnconfirmedQty,0) UnconfirmedQty
                                            ,ISNULL(x.UnconfirmedFreeSpareQty,0) UnconfirmedFreeSpareQty 
                                            ,ISNULL(x.UnconfirmedAvailableQty,0) UnconfirmedAvailableQty
                                            FROM COPTD a
                                            OUTER APPLY(
                                                SELECT SUM(ISNULL(x1.TH008,0)) UnconfirmedQty 
                                                      ,SUM(ISNULL(x1.TH024,0)) UnconfirmedFreeSpareQty 
                                                      ,SUM(ISNULL(x1.TH061,0)) UnconfirmedAvailableQty
                                                FROM COPTH x1
                                                WHERE x1.TH014 = a.TD001
                                                AND x1.TH015 = a.TD002
                                                AND x1.TH016 = a.TD003
                                                AND x1.TH020 ='N'
                                            ) x
                                            WHERE a.TD001 = @TG014 
                                            AND a.TD002 = @TG015  
                                            AND a.TD003 = @TG016
                                            AND a.TD021 = 'Y'
                                            ";
                            dynamicParameters.Add("TG014", SoErpPrefix);
                            dynamicParameters.Add("TG015", SoErpNo);
                            dynamicParameters.Add("TG016", SoSequence);
                            var resultCOPTD = sqlConnection.Query(sql, dynamicParameters);
                            if (resultCOPTD.Count() > 0)
                            {
                                foreach (var item2 in resultCOPTD)
                                {
                                    ProductType = item2.ProductType;
                                    CanUseQty = Convert.ToDouble(item2.CanUseQty);
                                    CanUseFreebieQty = Convert.ToDouble(item2.CanUseFreebieQty);
                                    CanUseSpareQty = Convert.ToDouble(item2.CanUseSpareQty);
                                    CanUseAvailableQty = Convert.ToDouble(item2.CanUseAvailableQty);
                                }
                            }
                            //if (Quantity > CanUseQty) throw new SystemException("【銷貨量超交】:品號『" + MtlItemNo + "』本次最大可銷貨數量為" + CanUseQty + "，請重新輸入!");
                            //if (Quantity > CanUseAvailableQty) throw new SystemException("【銷貨量超交】:品號『" + MtlItemNo + "』本次最大可銷貨計價數量為" + CanUseAvailableQty + "，請重新輸入!");
                            //if (ProductType == "1")
                            //{
                            //    if (FreeSpareQty > CanUseFreebieQty) throw new SystemException("【贈品量超交】:品號『" + MtlItemNo + "』本次最大可銷貨贈品數量為" + CanUseFreebieQty + "】，請重新輸入!");
                            //}
                            //else if (ProductType == "2")
                            //{
                            //    if (FreeSpareQty > CanUseSpareQty) throw new SystemException("【備品量超交】:品號『" + MtlItemNo + "』本次最大可銷貨備品數量為" + CanUseSpareQty + "】，請重新輸入!");
                            //}
                            //else
                            //{
                            //    throw new SystemException("贈備品類別異常,請重新確認");
                            //}
                            #endregion
                        }

                        var result = sqlConnection.Query(sql, dynamicParameters);

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
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

        #region //UpdateRtToERP -- 銷退單拋轉ERP -- Shintokuro 2024-05-13
        public string UpdateRtToERP(int RtId)
        {
            try
            {
                if (RtId < 0) throw new SystemException("銷退單不可為空,請重新確認!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    string dateNow = CreateDate.ToString("yyyyMMdd"), timeNow = CreateDate.ToString("HH:mm:ss"), errorMsg = "";
                    string uomNo = "", mtlItemNo = "", departmentNo = "", MESCompanyNo = "", userNo = "", userName = "", userGroup = "";
                    string RtErpPrefix = "", RtErpNo = "", TransferStatusMES = "", ConfirmStatus = "", SourceType = "", ReturnDate = "";
                    int OrderCompanyIdBase = -1;

                    List<COPTI> ReturnReceiveOrders = new List<COPTI>();
                    List<COPTJ> RtDetails = new List<COPTJ>();


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
                            MESCompanyNo = item.ErpNo;
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
                        dynamicParameters.Add("FunctionCode", "ReturnReceiveOrderManagment");
                        dynamicParameters.Add("DetailCode", "import");

                        var resultUser = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUser.Count() <= 0) throw new SystemException("使用者資料錯誤!");

                        foreach (var item in resultUser)
                        {
                            userNo = item.UserNo;
                            userName = item.UserName;
                            departmentNo = item.DepartmentNo;
                        }
                        #endregion

                        #region //撈取銷退單單頭資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.RtId, a.CompanyId, a.TransferStatusMES,
                                        a.RtErpPrefix TI001,                          a.RtErpNo TI002,  FORMAT(a.ReturnDate, 'yyyyMMdd') TI003,
                                         b.CustomerNo TI004,         ISNULL(c.DepartmentNo,'' ) TI005,   ISNULL( d.UserNo,'' ) TI006,
                                            a.Factory TI007,                         a.Currency TI008,          a.ExchangeRate TI009,
                                     a.OriginalAmount TI010,                a.OriginalTaxAmount TI011,           a.InvoiceType TI012,
                                            a.TaxType TI013,                        a.InvNumEnd TI014,                  a.UiNo TI015,
                                         a.PrintCount TI016,                              
                                                CASE 
                                                    WHEN a.InvoiceDate = '' THEN ''
                                                    ELSE FORMAT(a.InvoiceDate, 'yyyyMMdd')
                                                END AS TI017,
                                         a.UpdateCode TI018,
                                      a.ConfirmStatus TI019,                           a.Remark TI020,      a.CustomerFullName TI021,
                               ISNULL( d1.UserNo,'' ) TI022,                          a.Remark1 TI023,               a.Remark2 TI024,
                                            a.Remark3 TI025,             a.DeductionDistinction TI026,      a.CustomsClearance TI027,
                                             a.RowCnt TI028,                    a.TotalQuantity TI029,  ISNULL( d2.UserNo,'' ) TI030,
                                 a.RevenueJournalCode TI031,                  a.CostJournalCode TI032,             a.ApplyYYMM TI033, 
                        FORMAT(a.DocDate, 'yyyyMMdd') TI034,             ISNULL( d3.UserNo,'' ) TI035,               a.TaxRate TI036, 
                                       a.PretaxAmount TI037,                        a.TaxAmount TI038,           a.PaymentTerm TI039, 
                                      a.TotalPackages TI040,                  a.SignatureStatus TI041,           a.ProcessCode TI042, 
                                  a.TransferStatusERP TI043,                         a.BondCode TI044,     a.TransmissionCount TI045, 
                               a.CustomerAddressFirst TI046,            a.CustomerAddressSecond TI047,                       0 TI048, 
                                      a.ContactPerson TI049,                          a.Courier TI050,    a.SiteCommCalcMethod TI051, 
                                       a.SiteCommRate TI052,                  a.CommCalcInclTax TI053,       a.TotalCommAmount TI054, 
                                                   '' TI055,                                 '' TI056,          a.TradingTerms TI057, 
                                 a.CustomerEgFullName TI058,                                  0 TI059,                       0 TI060, 
                                          a.DebitNote TI061,                                 '' TI062,                      '' TI063, 
                                            a.TaxCode TI064,                                 '' TI065,            a.RemarkCode TI066, 
                                   a.MultipleInvoices TI067,                      a.InvNumStart TI068,      a.NumberOfInvoices TI069, 
                                       a.MultiTaxRate TI070,                  a.TaxCurrencyRate TI071,                       0 TI072, 
                                                    0 TI073,                                 '' TI074,                      '' TI075, 
                                         a.VoidRemark TI076,                    a.IncomeDraftID TI077,        a.IncomeDraftSeq TI078, 
                                  a.IncomeVoucherType TI079,              a.IncomeVoucherNumber TI080,                a.Status TI081, 
                                                   '' TI082,                                 '' TI083,                      '' TI084, 
                                                   '' TI085,                a.GenLedgerAcctType TI086,              a.InvCode2 TI087, 
                                          a.InvSymbol TI088,                                 '' TI089,                      '' TI090, 
                                       a.ContactEmail TI091,                    a.OrigInvNumber TI092
                        FROM SCM.ReturnReceiveOrder a
                        INNER JOIN SCM.Customer b on a.CustomerId = b.CustomerId
                        LEFT JOIN BAS.Department c on a.DepartmentId = c.DepartmentId
                        LEFT JOIN BAS.[User] d on a.SalesmenId = d.UserId
                        LEFT JOIN BAS.[User] d1 on a.CollectionSalesmenId = d1.UserId
                        LEFT JOIN BAS.[User] d2 on a.StaffId = d2.UserId
                        LEFT JOIN BAS.[User] d3 on a.ConfirmUserId = d3.UserId
                        WHERE a.CompanyId = @CompanyId
                        AND a.RtId = @RtId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("RtId", RtId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("銷退單單頭資料找不到，請重新確認");
                        foreach (var item in result)
                        {
                            TransferStatusMES = item.TransferStatusMES;
                            OrderCompanyIdBase = item.CompanyId;
                            RtErpPrefix = item.TI001;
                            RtErpNo = item.TI002;
                            ReturnDate = item.TI003;
                            ConfirmStatus = item.TI019;
                        }
                        if (ConfirmStatus == "Y") throw new SystemException("銷退單處於確認狀態不可以拋轉");
                        if (ConfirmStatus == "V") throw new SystemException("銷退單處於作廢狀態不可以拋轉");
                        if (OrderCompanyIdBase != CurrentCompany) throw new SystemException("單據的公司別與後端公司別不同，請嘗試重新登入!!");
                        ReturnReceiveOrders = sqlConnection.Query<COPTI>(sql, dynamicParameters).ToList();
                        #endregion

                        #region //撈取銷退單單身資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.RtDetailId,
                                        b.RtErpPrefix TJ001,                          b.RtErpNo TJ002,       
                                        FORMAT(ROW_NUMBER() OVER(ORDER BY a.RtDetailId),'0000') TJ003,
                                         c.MtlItemNo  TJ004,                      c.MtlItemName TJ005,           c.MtlItemSpec TJ006,
                                           a.Quantity TJ007,                            e.UomNo TJ008,                       0 TJ009,
                                                   '' TJ010,                        a.UnitPrice TJ011,                a.Amount TJ012,
                                        d.InventoryNo TJ013,                        a.LotNumber TJ014,          g2.RoErpPrefix TJ015,
                                           g2.RoErpNo TJ016,                      g1.RoSequence TJ017,          f2.SoErpPrefix TJ018,
                                           f2.SoErpNo TJ019,                      f1.SoSequence TJ020,         a.ConfirmStatus TJ021,
                                         a.UpdateCode TJ022,                          a.Remarks TJ023,          a.CheckOutCode TJ024,
                                     a.CheckOutPrefix TJ025,                       a.CheckOutNo TJ026,           a.CheckOutSeq TJ027,
                                        a.ProjectCode TJ028,                a.CustomerMtlItemNo TJ029,                  a.Type TJ030,
                                      a.OrigPreTaxAmt TJ031,                       a.OrigTaxAmt TJ032,             a.PreTaxAmt TJ033, 
                                             a.TaxAmt TJ034,                       a.PackageQty TJ035, ISNULL( e2.UomNo,'' )   TJ036, 
                                         a.BondedCode TJ037,                   a.CommissionRate TJ038,      a.CommissionAmount TJ039, 
                                   a.OriginalCustomer TJ040,                    a.FreeSpareType TJ041,          a.FreeSpareQty TJ042, 
                                a.PackageFreeSpareQty TJ043,                                 '' TJ044,                      '' TJ045, 
                                                   '' TJ046,                                  0 TJ047,            a.NotPayTemp TJ048, 
                                           a.Location TJ049,                     a.AvailableQty TJ050,                e3.UomNo TJ051, 
                                                   '' TJ052,                                  0 TJ053,                       0 TJ054, 
                                                   '' TJ055,                                 '' TJ056,                      '' TJ057, 
                                            a.TaxRate TJ058,                                 '' TJ059,                       0 TJ060, 
                                                   '' TJ061,                                 '' TJ062,                      '' TJ063, 
                                                   '' TJ064,                                 '' TJ065,                      '' TJ066, 
                                                   '' TJ067,                                 '' TJ068,                      '' TJ069, 
                                                   '' TJ070,                   a.DiscountAmount TJ071,                      '' TJ500, 
                                                   '' TJ501,                                 '' TJ502,                      '' TJ503
                                FROM SCM.RtDetail a
                                INNER JOIN SCM.ReturnReceiveOrder b on a.RtId = b.RtId
                                INNER JOIN PDM.MtlItem c on a.MtlItemId = c.MtlItemId
                                INNER JOIN SCM.Inventory d on a.InventoryId = d.InventoryId
                                INNER JOIN PDM.UnitOfMeasure e on a.UomId = e.UomId
                                LEFT JOIN PDM.UnitOfMeasure e2 on a.PackageUomId = e2.UomId
                                LEFT JOIN PDM.UnitOfMeasure e3 on a.AvailableUomId = e3.UomId
                                LEFT JOIN SCM.SoDetail f1 on a.SoDetailId = f1.SoDetailId
                                LEFT JOIN SCM.SaleOrder f2 on f1.SoId = f2.SoId
                                LEFT JOIN SCM.RoDetail g1 on a.RoDetailId = g1.RoDetailId
                                LEFT JOIN SCM.ReceiveOrder g2 on g1.RoId = g2.RoId
                                WHERE b.CompanyId = @CompanyId
                                AND b.RtId = @RtId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("RtId", RtId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("銷退單單身資料找不到，請重新確認");
                        foreach (var item in result)
                        {
                            if (item.TJ021 == "Y") throw new SystemException("銷貨單單身處於確認狀態不可以拋轉");
                            if (item.TJ021 == "V") throw new SystemException("銷貨單單身處於作廢狀態不可以拋轉");
                            if (item.TJ024 == "Y") throw new SystemException("銷貨單單身已經結帳狀態不可以拋轉");
                        }
                        RtDetails = sqlConnection.Query<COPTJ>(sql, dynamicParameters).ToList();
                        #endregion
                    }


                    if (ReturnReceiveOrders.Select(x => x.TI019).FirstOrDefault() == "N") //銷退單沒有確認
                    {
                        if (ReturnReceiveOrders.Count() > 0)
                        {
                            using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                            {
                                DynamicParameters dynamicParameters = new DynamicParameters();

                                //單頭資料
                                List<COPTI> addReturnReceiveOrders = ReturnReceiveOrders.Where(x => x.TransferStatusMES == "N").ToList();
                                List<COPTI> updateReturnReceiveOrders = ReturnReceiveOrders.Where(x => x.TransferStatusMES == "R").ToList();
                                //單身清單
                                List<COPTJ> addRtDetails = new List<COPTJ>();
                                List<COPTJ> updateRtDetails = new List<COPTJ>();
                                addRtDetails = RtDetails.Where(x => x.TJ021 == "N").ToList();

                                #region //ERP公司代碼取得
                                string ERPCompanyNo = "";
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 ML001  
                                         FROM CMSML";
                                var resultCompanyNo = sqlConnection.Query(sql, dynamicParameters);
                                foreach (var item in resultCompanyNo)
                                {
                                    ERPCompanyNo = item.ML001;
                                }
                                if (MESCompanyNo != ERPCompanyNo) throw new SystemException("ERP的公司代碼與MES的公司代碼不同，請嘗試重新登入!!");
                                #endregion

                                #region //判斷ERP使用者是否存在
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 MF005
                                        FROM ADMMF
                                        WHERE COMPANY = @CompanyNo
                                        AND MF001 = @userNo";
                                dynamicParameters.Add("CompanyNo", ERPCompanyNo);
                                dynamicParameters.Add("userNo", userNo);
                                var resultUserExist = sqlConnection.Query(sql, dynamicParameters);
                                if (resultUserExist.Count() <= 0) throw new SystemException("【ERP使用者】不存在!");
                                #endregion

                                #region //審核ERP權限-建立
                                var USR_GROUP = BaseHelper.CheckErpAuthority(userNo, sqlConnection, "COPI08", "CREATE");
                                #endregion

                                #region //判斷開立日期是否超過結帳日
                                string CloseDateBase = "";
                                string ReturnDateBase = ReturnDate;
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

                                if (DateTime.TryParseExact(ReturnDateBase, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out docDateBase) &&
                                    DateTime.TryParseExact(CloseDateBase, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out closeDateBase))
                                {
                                    int resultDate = DateTime.Compare(docDateBase, closeDateBase);

                                    if (resultDate <= 0)
                                    {
                                        throw new SystemException("【銷退單】ERP已經結帳,不可以在新增【" + CloseDateBase + "】之前的單據");
                                    }
                                }
                                else
                                {
                                    throw new SystemException("日期字符串无效，无法比较");
                                }
                                #endregion

                                if (ReturnReceiveOrders.Select(x => x.TransferStatusMES).FirstOrDefault() == "N") //沒有拋轉過
                                {
                                    #region //撈取單據編號編碼格式設定
                                    RtErpPrefix = ReturnReceiveOrders.Select(x => x.TI001).FirstOrDefault();
                                    string DocDate = ReturnReceiveOrders.Select(x => x.TI034).FirstOrDefault();

                                    string WoType = "";
                                    string encode = ""; // 編碼格式
                                    int yearLength = 0; // 年碼數
                                    int lineLength = 0; // 流水碼數
                                    DateTime referenceTime = default(DateTime);
                                    string docDate = addReturnReceiveOrders.Select(x => x.TI034).FirstOrDefault();
                                    referenceTime = DateTime.ParseExact(docDate, "yyyyMMdd", CultureInfo.InvariantCulture);
                                    #endregion

                                    #region //撈取ERP單據設定
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.MQ004, a.MQ005, a.MQ006, a.MQ044, a.MQ017
                                            FROM CMSMQ a
                                            WHERE MQ001 = @RtErpPrefix";
                                    dynamicParameters.Add("RtErpPrefix", RtErpPrefix);
                                    var resultReceiptNo = sqlConnection.Query(sql, dynamicParameters);
                                    foreach (var item in resultReceiptNo)
                                    {
                                        encode = item.MQ004; //編碼方式
                                        yearLength = Convert.ToInt32(item.MQ005); //年碼數
                                        lineLength = Convert.ToInt32(item.MQ006); //流水號碼數
                                        WoType = item.MQ017;
                                    }
                                    #endregion

                                    #region //單號自動取號
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT CONVERT(int, RIGHT(ISNULL(MAX(RTRIM(LTRIM(TI002))), '" + new string('0', lineLength) + @"'), " + lineLength + @")) + 1 CurrentNum
                                            FROM COPTI
                                            WHERE TI001 = @RtErpPrefix";
                                    dynamicParameters.Add("RtErpPrefix", RtErpPrefix);
                                    #endregion

                                    #region //編碼格式相關
                                    string dateFormat = "";
                                    switch (encode)
                                    {
                                        case "1": //日編
                                            if (yearLength < 0 || yearLength > 4) throw new SystemException("【ERP單據性質】年碼數不能小於等於0或大於4");
                                            if ((lineLength + yearLength + 4) > 11) throw new SystemException("【ERP單據性質】編碼格式碼數不能大於11碼");
                                            dateFormat = new string('y', yearLength) + "MMdd";
                                            sql += @" AND RTRIM(LTRIM(TI002)) LIKE @ReferenceTime  + '" + new string('_', lineLength) + @"'";
                                            //sql += @" AND TA002 LIKE '%' + @ReferenceTime + '%' + '" + new string('_', lineLength) + @"'";
                                            dynamicParameters.Add("ReferenceTime", referenceTime.ToString(dateFormat));

                                            string tedstNo = referenceTime.ToString(dateFormat);
                                            RtErpNo = referenceTime.ToString(dateFormat);
                                            break;
                                        case "2": //月編
                                            if (yearLength < 0 || yearLength > 4) throw new SystemException("【ERP單據性質】年碼數不能小於等於0或大於4");
                                            if ((lineLength + yearLength + 2) > 11) throw new SystemException("【ERP單據性質】編碼格式碼數不能大於11碼");
                                            dateFormat = new string('y', yearLength) + "MM";
                                            sql += @" AND RTRIM(LTRIM(TI002))  LIKE @ReferenceTime + '" + new string('_', lineLength) + @"'";
                                            dynamicParameters.Add("ReferenceTime", referenceTime.ToString(dateFormat));
                                            RtErpNo = referenceTime.ToString(dateFormat);
                                            break;
                                        case "3": //流水號
                                            if (yearLength == 0) throw new SystemException("【ERP單據性質】年碼數必須等於0");
                                            if (lineLength <= 0 || lineLength > 11) throw new SystemException("【ERP單據性質】流水編碼碼數必須大於0小於等於11");
                                            break;
                                        case "4": //手動編號
                                            break;
                                        default:
                                            throw new SystemException("編碼方式錯誤!");
                                    }

                                    int currentNum = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).CurrentNum;
                                    RtErpNo += string.Format("{0:" + new string('0', lineLength) + "}", currentNum);
                                    #endregion

                                    #region //判斷ERP銷退單單號是否重複
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM COPTI
                                            WHERE TI001 = @TI001
                                            AND TI002 = @TI002";
                                    dynamicParameters.Add("TI001", RtErpPrefix);
                                    dynamicParameters.Add("TI002", RtErpNo);

                                    var result = sqlConnection.Query(sql, dynamicParameters);
                                    if (result.Count() > 0) throw new SystemException("【銷貨單單號】重複，請重新取號!");
                                    string CFIELD01 = RtErpPrefix + "-" + RtErpNo;
                                    #endregion

                                    #region //新增
                                    #region //單頭
                                    if (addReturnReceiveOrders.Count > 0)
                                    {
                                        #region //賦值
                                        addReturnReceiveOrders
                                            .ToList()
                                            .ForEach(x =>
                                            {
                                                x.COMPANY = MESCompanyNo;
                                                x.USR_GROUP = USR_GROUP;
                                                x.FLAG = "1";
                                                x.CREATOR = userNo;
                                                x.CREATE_DATE = dateNow;
                                                x.CREATE_TIME = timeNow;
                                                x.CREATE_AP = BaseHelper.ClientComputer();
                                                x.CREATE_PRID = "BM";
                                                x.MODIFIER = "";
                                                x.MODI_DATE = "";
                                                x.MODI_TIME = "";
                                                x.MODI_AP = "";
                                                x.MODI_PRID = "";
                                                x.TI002 = RtErpNo;
                                                x.UDF01 = "";
                                                x.UDF02 = "";
                                                x.UDF03 = "";
                                                x.UDF04 = "";
                                                x.UDF05 = "";
                                                x.UDF06 = 0.0;
                                                x.UDF07 = 0.0;
                                                x.UDF08 = 0.0;
                                                x.UDF09 = 0.0;
                                                x.UDF10 = 0.0;
                                            });
                                        #endregion

                                        #region //新增銷退單頭 COPTI
                                        sql = @"INSERT INTO COPTI (COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE, FLAG, CREATE_TIME, 
                                                CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID, TI001, TI002, TI003, TI004, TI005, TI006, TI007, 
                                                TI008, TI009, TI010, TI011, TI012, TI013, TI014, TI015, TI016, TI017, TI018, TI019, TI020, TI021, TI022, TI023, TI024, 
                                                TI025, TI026, TI027, TI028, TI029, TI030, TI031, TI032, TI033, TI034, TI035, TI036, TI037, TI038, TI039, TI040, TI041, 
                                                TI042, TI043, TI044, TI045, TI046, TI047, TI048, TI049, TI050, TI051, TI052, TI053, TI054, TI055, TI056, TI057, TI058, 
                                                TI059, TI060, TI061, TI062, TI063, TI064, TI065, TI066, TI067, TI068, TI069, TI070, TI071, TI072, TI073, TI074, TI075, 
                                                TI076, TI077, TI078, TI079, TI080, TI081, TI082, TI083, TI084, TI085, TI086, UDF01, UDF02, UDF03, UDF04, UDF05, 
                                                UDF06, UDF07, UDF08, UDF09, UDF10, TI087, TI088, TI089, TI090, TI091, TI092)
                                                VALUES (@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE, @FLAG, @CREATE_TIME, 
                                                @CREATE_AP, @CREATE_PRID, @MODI_TIME, @MODI_AP, @MODI_PRID, @TI001, @TI002, @TI003, @TI004, @TI005, @TI006, @TI007, 
                                                @TI008, @TI009, @TI010, @TI011, @TI012, @TI013, @TI014, @TI015, @TI016, @TI017, @TI018, @TI019, @TI020, @TI021, @TI022, @TI023, @TI024, 
                                                @TI025, @TI026, @TI027, @TI028, @TI029, @TI030, @TI031, @TI032, @TI033, @TI034, @TI035, @TI036, @TI037, @TI038, @TI039, @TI040, @TI041, 
                                                @TI042, @TI043, @TI044, @TI045, @TI046, @TI047, @TI048, @TI049, @TI050, @TI051, @TI052, @TI053, @TI054, @TI055, @TI056, @TI057, @TI058, 
                                                @TI059, @TI060, @TI061, @TI062, @TI063, @TI064, @TI065, @TI066, @TI067, @TI068, @TI069, @TI070, @TI071, @TI072, @TI073, @TI074, @TI075, 
                                                @TI076, @TI077, @TI078, @TI079, @TI080, @TI081, @TI082, @TI083, @TI084, @TI085, @TI086, @UDF01, @UDF02, @UDF03, @UDF04, @UDF05, 
                                                @UDF06, @UDF07, @UDF08, @UDF09, @UDF10, @TI087, @TI088, @TI089, @TI090, @TI091, @TI092)";
                                        rowsAffected += sqlConnection.Execute(sql, addReturnReceiveOrders);
                                        #endregion
                                    }
                                    #endregion

                                    #region //單身                                    
                                    if (addRtDetails.Count > 0)
                                    {
                                        #region //銷退單單身資料驗證                                    
                                        string AllOverMistake = "";
                                        foreach (var item in addRtDetails)
                                        {
                                            string OverMtlItemNo = "";
                                            string OverMistake = "";
                                            string TJ004 = item.TJ004; //品號
                                            string TJ013 = item.TJ013; //庫別
                                            string TJ015 = item.TJ015; //訂單單別 單號 序號
                                            string TJ016 = item.TJ016;
                                            string TJ017 = item.TJ017;
                                            string RoFull = item.TJ015 + '-' + item.TJ016 + '-' + item.TJ017;
                                            string TJ018 = item.TJ018; //銷貨單單別 單號 序號
                                            string TJ019 = item.TJ019;
                                            string TJ020 = item.TJ020;
                                            string SoFull = item.TJ018 + '-' + item.TJ019 + '-' + item.TJ020;
                                            string TJ014 = item.TJ014; //批號
                                            string TJ041 = item.TJ041; //數量類型
                                            Double TJ007 = item.TJ007; //數量
                                            Double TJ042 = item.TJ042; //贈/備品數量

                                            Double TH008 = 0; //銷貨單數量
                                            Double TH024 = 0; //贈/備品數量

                                            double CanUseQty = 0;
                                            double CanUseFreeSpareQty = 0;

                                            #region //品號批號管理檢查
                                            string MB022 = "";
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"SELECT TOP 1 MB022
                                                    FROM INVMB
                                                    WHERE MB001 = @TJ004";
                                            dynamicParameters.Add("TJ004", TJ004);
                                            result = sqlConnection.Query(sql, dynamicParameters);
                                            if (result.Count() <= 0) throw new SystemException("品號找不到,請重新確認");
                                            foreach (var item1 in result)
                                            {
                                                MB022 = item1.MB022;
                                            }
                                            if (MB022 != "N")
                                            {
                                                #region //判斷品號批號庫存是否足夠
                                                dynamicParameters = new DynamicParameters();
                                                sql = @"SELECT a.ME001, a.ME002,a.ME003,a.ME009,b.num,b.MF007
                                                        FROM INVME a
                                                        OUTER APPLY(
                                                              SELECT SUM(x.MF010) num,x.MF007 FROM INVMF x
                                                              WHERE x.MF001 = a.ME001 AND x.MF002 = a.ME002
                                                              GROUP BY x.MF007
                                                          ) b
                                                        WHERE 1 = 1
                                                        AND b.MF007 = @InventoryNo
                                                        AND a.ME001 = @MtlItemNo
                                                        AND a.ME002 = @LotNumber
                                                        ORDER BY a.ME007 DESC   
                                                        ";
                                                dynamicParameters.Add("InventoryNo", TJ013);
                                                dynamicParameters.Add("MtlItemNo", TJ004);
                                                dynamicParameters.Add("LotNumber", TJ014);
                                                result = sqlConnection.Query(sql, dynamicParameters);
                                                foreach (var item1 in result)
                                                {
                                                    if (Convert.ToDouble(item1.num) < (TH008 + TH024)) throw new SystemException("該品號庫存批號不足,請重新確認");
                                                }
                                                #endregion
                                            }
                                            #endregion

                                            #region //撈取銷貨單數量資料
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"SELECT TOP 1 a.TH026,
                                                    ISNULL(a.TH008,0) TH008,
                                                    ISNULL(a.TH024,0) TH024,
                                                    (ISNULL(a.TH008,0) -  ISNULL(a.TH043,0) - ISNULL(x.UnconfirmedQty,0)) CanUseQty,
                                                    (ISNULL(a.TH024,0) -  ISNULL(a.TH052,0) - ISNULL(x.UnconfirmedFreeSpareQty,0)) CanUseFreeSpareQty,
                                                    ISNULL(x.UnconfirmedQty,0) UnconfirmedQty,
                                                    ISNULL(x.UnconfirmedFreeSpareQty,0) UnconfirmedFreeSpareQty
                                                    FROM COPTH a
                                                    OUTER APPLY(
                                                        SELECT SUM(ISNULL(x1.TJ007,0)) UnconfirmedQty 
                                                                ,SUM(ISNULL(x1.TJ042,0)) UnconfirmedFreeSpareQty 
                                                        FROM COPTJ x1
                                                        WHERE x1.TJ018 = a.TH001
                                                        AND x1.TJ019 = a.TH002
                                                        AND x1.TJ020 = a.TH003
                                                        AND x1.TJ021 ='N'
                                                    ) x
                                                    WHERE a.TH001 = @TH001
                                                    AND a.TH002 = @TH002
                                                    AND a.TH003 = @TH003
                                                    AND a.TH020 = 'Y'";
                                            dynamicParameters.Add("TH001", TJ015);
                                            dynamicParameters.Add("TH002", TJ016);
                                            dynamicParameters.Add("TH003", TJ017);
                                            result = sqlConnection.Query(sql, dynamicParameters);
                                            if (result.Count() <= 0) throw new SystemException("銷貨單找不到,請重新確認");
                                            foreach (var item1 in result)
                                            {
                                                if (item1.TH026 == "Y") throw new SystemException("【銷貨單:" + RoFull + "】已經結案!!!");
                                                TH008 = Convert.ToDouble(item1.TH008);
                                                TH024 = Convert.ToDouble(item1.TH024);
                                                CanUseQty = Convert.ToDouble(item1.CanUseQty);
                                                CanUseFreeSpareQty = Convert.ToDouble(item1.CanUseFreeSpareQty);
                                            }
                                            #endregion

                                            #region //判斷是否有超過可退貨數
                                            OverMtlItemNo = "銷退單單身品號:" + TJ004;
                                            if ((CanUseQty - TH008) < 0)
                                            {
                                                OverMistake += "<br>已經超過可銷貨數量<br>銷貨單數量:" + TH008 + ",目前可銷貨最大數:" + CanUseQty + "<br>欲新增銷貨數量:" + TJ007 + ",請重新確認";
                                            }
                                            if (TJ041 == "1")
                                            {
                                                if ((CanUseFreeSpareQty - TJ042) < 0)
                                                {
                                                    OverMistake += "<br>已經超過可銷貨贈品數量,銷貨單贈品數量:" + TH024 + ",目前可銷貨贈品最大數量:" + CanUseFreeSpareQty + "<br>欲新增銷貨贈品數量:" + TJ042 + ",請重新確認<br>";
                                                }
                                            }
                                            else if (TJ041 == "2")
                                            {
                                                if ((CanUseFreeSpareQty - TJ042) < 0)
                                                {
                                                    OverMistake += "<br>已經超過可銷貨備品數量,銷貨單備品數量:" + TH024 + ",目前可銷貨貨品最大數量:" + CanUseFreeSpareQty + "<br>欲新增銷貨品數量:" + TJ042 + ",請重新確認<br>";
                                                }
                                            }
                                            if (OverMistake.Length > 0)
                                            {
                                                AllOverMistake += OverMtlItemNo + OverMistake;
                                            }
                                            #endregion
                                        }

                                        if (AllOverMistake.Length > 0)
                                        {
                                            throw new SystemException(AllOverMistake);
                                        }
                                        #endregion

                                        #region //銷退單身賦值
                                        addRtDetails
                                            .ToList()
                                            .ForEach(x =>
                                            {
                                                x.COMPANY = MESCompanyNo;
                                                x.USR_GROUP = USR_GROUP;
                                                x.FLAG = "1";
                                                x.CREATOR = userNo;
                                                x.CREATE_DATE = dateNow;
                                                x.CREATE_TIME = timeNow;
                                                x.CREATE_AP = BaseHelper.ClientComputer();
                                                x.CREATE_PRID = "BM";
                                                x.MODIFIER = "";
                                                x.MODI_DATE = "";
                                                x.MODI_TIME = "";
                                                x.MODI_AP = "";
                                                x.MODI_PRID = "";
                                                x.TJ002 = RtErpNo;
                                                x.UDF01 = "";
                                                x.UDF02 = "";
                                                x.UDF03 = "";
                                                x.UDF04 = "";
                                                x.UDF05 = "";
                                                x.UDF06 = 0.0;
                                                x.UDF07 = 0.0;
                                                x.UDF08 = 0.0;
                                                x.UDF09 = 0.0;
                                                x.UDF10 = 0.0;
                                            });
                                        #endregion

                                        #region //新增銷退單身 COPTJ                                   
                                        sql = @"INSERT INTO COPTJ(COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE, FLAG, CREATE_TIME, 
                                                CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID, TJ001, TJ002, TJ003, TJ004, TJ005, TJ006, TJ007, 
                                                TJ008, TJ009, TJ010, TJ011, TJ012, TJ013, TJ014, TJ015, TJ016, TJ017, TJ018, TJ019, TJ020, TJ021, TJ022, TJ023, 
                                                TJ024, TJ025, TJ026, TJ027, TJ028, TJ029, TJ030, TJ031, TJ032, TJ033, TJ034, TJ035, TJ036, TJ037, TJ038, TJ039, 
                                                TJ040, TJ041, TJ042, TJ043, TJ044, TJ045, TJ046, TJ047, TJ048, TJ049, TJ050, TJ051, TJ052, TJ053, TJ054, TJ055, 
                                                TJ056, TJ057, TJ058, TJ059, TJ060, TJ061, TJ500, TJ501, TJ502, TJ503, TJ062, TJ063, TJ064, TJ065, TJ066, TJ067, 
                                                TJ068, TJ069, UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10, TJ070, TJ071)
                                                VALUES(@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE, @FLAG, @CREATE_TIME, 
                                                @CREATE_AP, @CREATE_PRID, @MODI_TIME, @MODI_AP, @MODI_PRID, @TJ001, @TJ002, @TJ003, @TJ004, @TJ005, @TJ006, @TJ007, 
                                                @TJ008, @TJ009, @TJ010, @TJ011, @TJ012, @TJ013, @TJ014, @TJ015, @TJ016, @TJ017, @TJ018, @TJ019, @TJ020, @TJ021, @TJ022, @TJ023, 
                                                @TJ024, @TJ025, @TJ026, @TJ027, @TJ028, @TJ029, @TJ030, @TJ031, @TJ032, @TJ033, @TJ034, @TJ035, @TJ036, @TJ037, @TJ038, @TJ039, 
                                                @TJ040, @TJ041, @TJ042, @TJ043, @TJ044, @TJ045, @TJ046, @TJ047, @TJ048, @TJ049, @TJ050, @TJ051, @TJ052, @TJ053, @TJ054, @TJ055, 
                                                @TJ056, @TJ057, @TJ058, @TJ059, @TJ060, @TJ061, @TJ500, @TJ501, @TJ502, @TJ503, @TJ062, @TJ063, @TJ064, @TJ065, @TJ066, @TJ067, 
                                                @TJ068, @TJ069, @UDF01, @UDF02, @UDF03, @UDF04, @UDF05, @UDF06, @UDF07, @UDF08, @UDF09, @UDF10, @TJ070, @TJ071)";
                                        rowsAffected += sqlConnection.Execute(sql, addRtDetails);
                                        #endregion
                                    }
                                    #endregion
                                    #endregion

                                }
                                else if (ReturnReceiveOrders.Select(x => x.TransferStatusMES).FirstOrDefault() == "R") //有拋轉過
                                {
                                    #region //判斷銷退單是否存在+可不可以拋轉
                                    string erpConfirmStatus = "";
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TI019
                                            FROM COPTI
                                            WHERE TI001 =@RtErpPrefix
                                            AND TI002 = @RtErpNo";
                                    dynamicParameters.Add("RtErpPrefix", RtErpPrefix);
                                    dynamicParameters.Add("RtErpNo", RtErpNo);
                                    var result = sqlConnection.Query(sql, dynamicParameters);
                                    if (result.Count() <= 0) throw new SystemException("銷退單不存在,請重新確認!!!");

                                    foreach (var item in result)
                                    {
                                        erpConfirmStatus = item.TI019;
                                    }
                                    if (erpConfirmStatus == "Y") throw new SystemException("ERP銷退單處於確認狀態不可以拋轉!!!");
                                    if (erpConfirmStatus == "V") throw new SystemException("ERP銷退單處於作廢狀態不可以拋轉!!!");
                                    #endregion

                                    #region //單頭修改
                                    if (updateReturnReceiveOrders.Count > 0)
                                    {
                                        updateReturnReceiveOrders
                                            .ToList()
                                            .ForEach(x =>
                                            {
                                                x.MODIFIER = userNo;
                                                x.MODI_DATE = dateNow;
                                                x.MODI_TIME = timeNow;
                                                x.MODI_AP = BaseHelper.ClientComputer();
                                                x.MODI_PRID = "BM";
                                            });

                                        sql = @"UPDATE COPTI SET
                                                TI003 = @TI003,  
                                                TI004 = @TI004,  
                                                TI005 = @TI005,  
                                                TI006 = @TI006,  
                                                TI007 = @TI007,  
                                                TI008 = @TI008,  
                                                TI009 = @TI009,  
                                                TI010 = @TI010,  
                                                TI011 = @TI011,  
                                                TI012 = @TI012,  
                                                TI013 = @TI013,  
                                                TI014 = @TI014,  
                                                TI015 = @TI015,  
                                                TI016 = @TI016,  
                                                TI017 = @TI017,  
                                                TI018 = @TI018,  
                                                TI019 = @TI019,  
                                                TI020 = @TI020,  
                                                TI021 = @TI021,  
                                                TI022 = @TI022,  
                                                TI023 = @TI023,  
                                                TI024 = @TI024,  
                                                TI025 = @TI025,  
                                                TI026 = @TI026,  
                                                TI027 = @TI027,  
                                                TI028 = @TI028,  
                                                TI029 = @TI029,  
                                                TI030 = @TI030,  
                                                TI031 = @TI031,  
                                                TI032 = @TI032,  
                                                TI033 = @TI033,  
                                                TI034 = @TI034,  
                                                TI035 = @TI035,  
                                                TI036 = @TI036,  
                                                TI037 = @TI037,  
                                                TI038 = @TI038,  
                                                TI039 = @TI039,  
                                                TI040 = @TI040,  
                                                TI041 = @TI041,  
                                                TI042 = @TI042,  
                                                TI044 = @TI044,  
                                                TI045 = @TI045,  
                                                TI046 = @TI046,  
                                                TI047 = @TI047,  
                                                TI048 = @TI048,  
                                                TI049 = @TI049,  
                                                TI050 = @TI050,  
                                                TI051 = @TI051,  
                                                TI052 = @TI052,  
                                                TI053 = @TI053,  
                                                TI054 = @TI054,  
                                                TI055 = @TI055,  
                                                TI056 = @TI056,  
                                                TI057 = @TI057,  
                                                TI058 = @TI058,  
                                                TI059 = @TI059,  
                                                TI060 = @TI060,  
                                                TI061 = @TI061,  
                                                TI062 = @TI062,  
                                                TI063 = @TI063,  
                                                TI064 = @TI064,  
                                                TI065 = @TI065,  
                                                TI066 = @TI066,  
                                                TI067 = @TI067,  
                                                TI068 = @TI068,  
                                                TI069 = @TI069,  
                                                TI070 = @TI070,  
                                                TI071 = @TI071,  
                                                TI072 = @TI072,  
                                                TI073 = @TI073,  
                                                TI074 = @TI074,  
                                                TI075 = @TI075,  
                                                TI076 = @TI076,  
                                                TI077 = @TI077,  
                                                TI078 = @TI078,  
                                                TI079 = @TI079,  
                                                TI080 = @TI080,  
                                                TI081 = @TI081,  
                                                TI082 = @TI082,  
                                                TI083 = @TI083,  
                                                TI084 = @TI084,  
                                                TI085 = @TI085,  
                                                TI086 = @TI086,  
                                                TI087 = @TI087,  
                                                TI088 = @TI088,  
                                                TI089 = @TI089,  
                                                TI090 = @TI090,  
                                                TI091 = @TI091,  
                                                TI092 = @TI092  
                                                WHERE TI001 = @TI001
                                                AND TI002 = @TI002";
                                        rowsAffected += sqlConnection.Execute(sql, updateReturnReceiveOrders);
                                    }
                                    #endregion

                                    #region //單身依MES最新版重新拋轉
                                    #region //刪除單身
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE COPTJ
                                            WHERE TJ001 = @RtErpPrefix
                                            AND TJ002 = @RtErpNo";
                                    dynamicParameters.Add("RtErpPrefix", RtErpPrefix);
                                    dynamicParameters.Add("RtErpNo", RtErpNo);

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //新增單身                                    
                                    if (addRtDetails.Count > 0)
                                    {
                                        #region //銷退單單身資料驗證                                    
                                        string AllOverMistake = "";
                                        foreach (var item in addRtDetails)
                                        {
                                            string OverMtlItemNo = "";
                                            string OverMistake = "";
                                            string TJ004 = item.TJ004; //品號
                                            string TJ013 = item.TJ013; //庫別
                                            string TJ015 = item.TJ015; //銷貨單別 單號 序號
                                            string TJ016 = item.TJ016;
                                            string TJ017 = item.TJ017;
                                            string RoFull = item.TJ015 + '-' + item.TJ016 + '-' + item.TJ017;
                                            string TJ018 = item.TJ018; //訂單單單別 單號 序號
                                            string TJ019 = item.TJ019;
                                            string TJ020 = item.TJ020;
                                            string SoFull = item.TJ018 + '-' + item.TJ019 + '-' + item.TJ020;
                                            string TJ014 = item.TJ014; //批號
                                            string TJ041 = item.TJ041; //數量類型
                                            Double TJ007 = item.TJ007; //數量
                                            Double TJ042 = item.TJ042; //贈/備品數量

                                            Double TH008 = 0; //銷貨單數量
                                            Double TH024 = 0; //贈/備品數量

                                            double CanUseQty = 0;
                                            double CanUseFreeSpareQty = 0;

                                            #region //品號批號管理檢查
                                            string MB022 = "";
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"SELECT TOP 1 MB022
                                                    FROM INVMB
                                                    WHERE MB001 = @TJ004";
                                            dynamicParameters.Add("TJ004", TJ004);
                                            result = sqlConnection.Query(sql, dynamicParameters);
                                            if (result.Count() <= 0) throw new SystemException("品號找不到,請重新確認");
                                            foreach (var item1 in result)
                                            {
                                                MB022 = item1.MB022;
                                            }
                                            if (MB022 != "N")
                                            {
                                                #region //判斷品號批號庫存是否足夠
                                                dynamicParameters = new DynamicParameters();
                                                sql = @"SELECT a.ME001, a.ME002,a.ME003,a.ME009,b.num,b.MF007
                                                        FROM INVME a
                                                        OUTER APPLY(
                                                              SELECT SUM(x.MF010) num,x.MF007 FROM INVMF x
                                                              WHERE x.MF001 = a.ME001 AND x.MF002 = a.ME002
                                                              GROUP BY x.MF007
                                                          ) b
                                                        WHERE 1 = 1
                                                        AND b.MF007 = @InventoryNo
                                                        AND a.ME001 = @MtlItemNo
                                                        AND a.ME002 = @LotNumber
                                                        ORDER BY a.ME007 DESC   
                                                        ";
                                                dynamicParameters.Add("InventoryNo", TJ013);
                                                dynamicParameters.Add("MtlItemNo", TJ004);
                                                dynamicParameters.Add("LotNumber", TJ014);
                                                result = sqlConnection.Query(sql, dynamicParameters);
                                                foreach (var item1 in result)
                                                {
                                                    if (Convert.ToDouble(item1.num) < (TH008 + TH024)) throw new SystemException("該品號庫存批號不足,請重新確認");
                                                }
                                                #endregion
                                            }
                                            #endregion

                                            #region //撈取銷貨單數量資料
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"SELECT TOP 1 a.TH026,
                                                    ISNULL(a.TH008,0) TH008,
                                                    ISNULL(a.TH024,0) TH024,
                                                    (ISNULL(a.TH008,0) -  ISNULL(a.TH043,0) - ISNULL(x.UnconfirmedQty,0)) CanUseQty,
                                                    (ISNULL(a.TH024,0) -  ISNULL(a.TH052,0) - ISNULL(x.UnconfirmedFreeSpareQty,0)) CanUseFreeSpareQty,
                                                    ISNULL(x.UnconfirmedQty,0) UnconfirmedQty,
                                                    ISNULL(x.UnconfirmedFreeSpareQty,0) UnconfirmedFreeSpareQty
                                                    FROM COPTH a
                                                    OUTER APPLY(
                                                        SELECT SUM(ISNULL(x1.TJ007,0)) UnconfirmedQty 
                                                                ,SUM(ISNULL(x1.TJ042,0)) UnconfirmedFreeSpareQty 
                                                        FROM COPTJ x1
                                                        WHERE x1.TJ018 = a.TH001
                                                        AND x1.TJ019 = a.TH002
                                                        AND x1.TJ020 = a.TH003
                                                        AND x1.TJ021 ='N'
                                                    ) x
                                                    WHERE a.TH001 = @TH001
                                                    AND a.TH002 = @TH002
                                                    AND a.TH003 = @TH003
                                                    AND a.TH020 = 'Y'";
                                            dynamicParameters.Add("TH001", TJ015);
                                            dynamicParameters.Add("TH002", TJ016);
                                            dynamicParameters.Add("TH003", TJ017);
                                            result = sqlConnection.Query(sql, dynamicParameters);
                                            if (result.Count() <= 0) throw new SystemException("銷貨單找不到,請重新確認");
                                            foreach (var item1 in result)
                                            {
                                                if (item1.TH026 == "Y") throw new SystemException("【銷貨單:" + RoFull + "】已經結案!!!");
                                                TH008 = Convert.ToDouble(item1.TH008);
                                                TH024 = Convert.ToDouble(item1.TH024);
                                                CanUseQty = Convert.ToDouble(item1.CanUseQty);
                                                CanUseFreeSpareQty = Convert.ToDouble(item1.CanUseFreeSpareQty);
                                            }
                                            #endregion

                                            #region //判斷是否有超過可退貨數
                                            OverMtlItemNo = "銷退單單身品號:" + TJ004;
                                            if ((CanUseQty - TH008) < 0)
                                            {
                                                OverMistake += "<br>已經超過可銷貨數量<br>銷貨單數量:" + TH008 + ",目前可銷貨最大數:" + CanUseQty + "<br>欲新增銷貨數量:" + TJ007 + ",請重新確認";
                                            }
                                            if (TJ041 == "1")
                                            {
                                                if ((CanUseFreeSpareQty - TJ042) < 0)
                                                {
                                                    OverMistake += "<br>已經超過可銷貨贈品數量,銷貨單贈品數量:" + TH024 + ",目前可銷貨贈品最大數量:" + CanUseFreeSpareQty + "<br>欲新增銷貨贈品數量:" + TJ042 + ",請重新確認<br>";
                                                }
                                            }
                                            else if (TJ041 == "2")
                                            {
                                                if ((CanUseFreeSpareQty - TJ042) < 0)
                                                {
                                                    OverMistake += "<br>已經超過可銷貨備品數量,銷貨單備品數量:" + TH024 + ",目前可銷貨貨品最大數量:" + CanUseFreeSpareQty + "<br>欲新增銷貨品數量:" + TJ042 + ",請重新確認<br>";
                                                }
                                            }
                                            if (OverMistake.Length > 0)
                                            {
                                                AllOverMistake += OverMtlItemNo + OverMistake;
                                            }
                                            #endregion
                                        }

                                        if (AllOverMistake.Length > 0)
                                        {
                                            throw new SystemException(AllOverMistake);
                                        }
                                        #endregion

                                        #region //銷退單身賦值
                                        addRtDetails
                                            .ToList()
                                            .ForEach(x =>
                                            {
                                                x.COMPANY = MESCompanyNo;
                                                x.USR_GROUP = USR_GROUP;
                                                x.FLAG = "1";
                                                x.CREATOR = userNo;
                                                x.CREATE_DATE = dateNow;
                                                x.CREATE_TIME = timeNow;
                                                x.CREATE_AP = BaseHelper.ClientComputer();
                                                x.CREATE_PRID = "BM";
                                                x.MODIFIER = "";
                                                x.MODI_DATE = "";
                                                x.MODI_TIME = "";
                                                x.MODI_AP = "";
                                                x.MODI_PRID = "";
                                                x.TJ002 = RtErpNo;
                                                x.UDF01 = "";
                                                x.UDF02 = "";
                                                x.UDF03 = "";
                                                x.UDF04 = "";
                                                x.UDF05 = "";
                                                x.UDF06 = 0.0;
                                                x.UDF07 = 0.0;
                                                x.UDF08 = 0.0;
                                                x.UDF09 = 0.0;
                                                x.UDF10 = 0.0;
                                            });
                                        #endregion

                                        #region //新增銷退單身 COPTJ                                   
                                        sql = @"INSERT INTO COPTJ(COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE, FLAG, CREATE_TIME, 
                                                CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID, TJ001, TJ002, TJ003, TJ004, TJ005, TJ006, TJ007, 
                                                TJ008, TJ009, TJ010, TJ011, TJ012, TJ013, TJ014, TJ015, TJ016, TJ017, TJ018, TJ019, TJ020, TJ021, TJ022, TJ023, 
                                                TJ024, TJ025, TJ026, TJ027, TJ028, TJ029, TJ030, TJ031, TJ032, TJ033, TJ034, TJ035, TJ036, TJ037, TJ038, TJ039, 
                                                TJ040, TJ041, TJ042, TJ043, TJ044, TJ045, TJ046, TJ047, TJ048, TJ049, TJ050, TJ051, TJ052, TJ053, TJ054, TJ055, 
                                                TJ056, TJ057, TJ058, TJ059, TJ060, TJ061, TJ500, TJ501, TJ502, TJ503, TJ062, TJ063, TJ064, TJ065, TJ066, TJ067, 
                                                TJ068, TJ069, UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10, TJ070, TJ071)
                                                VALUES(@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE, @FLAG, @CREATE_TIME, 
                                                @CREATE_AP, @CREATE_PRID, @MODI_TIME, @MODI_AP, @MODI_PRID, @TJ001, @TJ002, @TJ003, @TJ004, @TJ005, @TJ006, @TJ007, 
                                                @TJ008, @TJ009, @TJ010, @TJ011, @TJ012, @TJ013, @TJ014, @TJ015, @TJ016, @TJ017, @TJ018, @TJ019, @TJ020, @TJ021, @TJ022, @TJ023, 
                                                @TJ024, @TJ025, @TJ026, @TJ027, @TJ028, @TJ029, @TJ030, @TJ031, @TJ032, @TJ033, @TJ034, @TJ035, @TJ036, @TJ037, @TJ038, @TJ039, 
                                                @TJ040, @TJ041, @TJ042, @TJ043, @TJ044, @TJ045, @TJ046, @TJ047, @TJ048, @TJ049, @TJ050, @TJ051, @TJ052, @TJ053, @TJ054, @TJ055, 
                                                @TJ056, @TJ057, @TJ058, @TJ059, @TJ060, @TJ061, @TJ500, @TJ501, @TJ502, @TJ503, @TJ062, @TJ063, @TJ064, @TJ065, @TJ066, @TJ067, 
                                                @TJ068, @TJ069, @UDF01, @UDF02, @UDF03, @UDF04, @UDF05, @UDF06, @UDF07, @UDF08, @UDF09, @UDF10, @TJ070, @TJ071)";
                                        rowsAffected += sqlConnection.Execute(sql, addRtDetails);
                                        #endregion
                                    }
                                    #endregion
                                    #endregion
                                }
                            }

                            using (SqlConnection sqlConnections = new SqlConnection(MainConnectionStrings))
                            {
                                DynamicParameters dynamicParameters = new DynamicParameters();
                                if (ReturnReceiveOrders.Select(x => x.TransferStatusMES).FirstOrDefault() == "N") //沒有拋轉過
                                {
                                    #region //單頭
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE SCM.ReturnReceiveOrder SET
                                            RtErpNo = @RtErpNo,
                                            TransferStatusMES = @TransferStatusMES,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE RtId = @RtId";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            RtErpNo = ReturnReceiveOrders.Select(x => x.TI002).FirstOrDefault(),
                                            TransferStatusMES = "Y",
                                            LastModifiedDate,
                                            LastModifiedBy,
                                            RtId
                                        });
                                    var insertResult = sqlConnections.Query(sql, dynamicParameters);

                                    rowsAffected += sqlConnections.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //單身
                                    dynamicParameters = new DynamicParameters();
                                    RtDetails
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        sql = @"UPDATE SCM.RtDetail SET
                                                RtSequence = @RtSequence,
                                                TransferStatusMES = @TransferStatusMES,
                                                LastModifiedDate = @LastModifiedDate,
                                                LastModifiedBy = @LastModifiedBy
                                                WHERE RtDetailId = @RtDetailId";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                RtSequence = x.TJ003,
                                                TransferStatusMES = "Y",
                                                LastModifiedDate,
                                                LastModifiedBy,
                                                RtDetailId = x.RtDetailId
                                            });
                                        insertResult = sqlConnections.Query(sql, dynamicParameters);

                                        rowsAffected += sqlConnections.Execute(sql, dynamicParameters);
                                    });

                                    #endregion
                                }
                                else if (ReturnReceiveOrders.Select(x => x.TransferStatusMES).FirstOrDefault() == "R") //有拋轉過
                                {
                                    #region //單頭
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE SCM.ReturnReceiveOrder SET
                                            TransferStatusMES = @TransferStatusMES,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE RtId = @RtId";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            TransferStatusMES = "Y",
                                            LastModifiedDate,
                                            LastModifiedBy,
                                            RtId
                                        });

                                    rowsAffected += sqlConnections.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //單身
                                    dynamicParameters = new DynamicParameters();
                                    RtDetails
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        sql = @"UPDATE SCM.RtDetail SET
                                                RtSequence = @RtSequence,
                                                TransferStatusMES = @TransferStatusMES,
                                                LastModifiedDate = @LastModifiedDate,
                                                LastModifiedBy = @LastModifiedBy
                                                WHERE RtDetailId = @RtDetailId";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                RtSequence = x.TJ003,
                                                TransferStatusMES = "Y",
                                                LastModifiedDate,
                                                LastModifiedBy,
                                                RtDetailId = x.RtDetailId
                                            });

                                        rowsAffected += sqlConnections.Execute(sql, dynamicParameters);
                                    });

                                    #endregion
                                }

                            }

                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                msg = "拋轉成功"
                            });
                            #endregion

                            transactionScope.Complete();
                        }
                    }
                    else
                    {
                        throw new SystemException("銷貨單已拋轉ERP!!!");
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

        #region //UpdateRtReviseMES -- 銷退單回歸MES重新編輯 -- Shintokuro 2024-05-14
        public string UpdateRtReviseMES(int RtId)
        {
            try
            {
                if (RtId < 0) throw new SystemException("銷退單不可為空,請重新確認!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;
                        string dateNow = CreateDate.ToString("yyyyMMdd"), timeNow = CreateDate.ToString("HH:mm:ss"), errorMsg = "";
                        string MESCompanyNo = "", departmentNo = "", userNo = "", userName = "", userGroup = "";
                        string RtErpPrefix = "", RtErpNo = "", TransferStatusMES = "", ConfirmStatus = "";
                        int OrderCompanyIdBase = -1;


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
                            MESCompanyNo = item.ErpNo;
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
                        dynamicParameters.Add("FunctionCode", "ReturnReceiveOrderManagment");
                        dynamicParameters.Add("DetailCode", "import");

                        var resultUser = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUser.Count() <= 0) throw new SystemException("使用者資料錯誤!");

                        foreach (var item in resultUser)
                        {
                            userNo = item.UserNo;
                            userName = item.UserName;
                            departmentNo = item.DepartmentNo;
                        }
                        #endregion

                        #region //確認銷退單是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.TransferStatusMES, a.ConfirmStatus, a.CompanyId 
                                FROM SCM.ReturnReceiveOrder a
                                WHERE a.CompanyId = @CompanyId
                                AND a.RtId = @RtId
                                ";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("RtId", RtId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("銷退單單頭資料找不到，請重新確認");
                        foreach (var item in result)
                        {
                            TransferStatusMES = item.TransferStatusMES;
                            ConfirmStatus = item.ConfirmStatus;
                            OrderCompanyIdBase = item.CompanyId;
                        }
                        if (TransferStatusMES != "Y") throw new SystemException("銷退單未處於拋轉ERP狀態不可以操作");
                        if (ConfirmStatus == "Y") throw new SystemException("銷退單處於確認狀態不可以操作");
                        if (ConfirmStatus == "V") throw new SystemException("銷退單處於作廢狀態不可以操作");
                        if (OrderCompanyIdBase != CurrentCompany) throw new SystemException("單據的公司別與目前公司別不同，請嘗試重新登入!!");
                        #endregion

                        #region //更新單頭
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.ReturnReceiveOrder SET
                                TransferStatusMES = @TransferStatusMES,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RtId = @RtId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TransferStatusMES = "R",
                                LastModifiedDate,
                                LastModifiedBy,
                                RtId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //更新單身
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.RtDetail SET
                                TransferStatusMES = @TransferStatusMES,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RtId = @RtId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TransferStatusMES = "R",
                                LastModifiedDate,
                                LastModifiedBy,
                                RtId
                            });
                        insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "可以重新編輯單據了!"
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

        #region //UpdateRtVoid -- 銷退單作廢 -- Shintokuro 2024-03-04
        public string UpdateRtVoid(int RtId)
        {
            try
            {
                if (RtId < 0) throw new SystemException("銷貨單不可為空,請重新確認!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    string dateNow = CreateDate.ToString("yyyyMMdd"), timeNow = CreateDate.ToString("HH:mm:ss"), errorMsg = "";
                    string MESCompanyNo = "", departmentNo = "", userNo = "", userName = "", userGroup = "";
                    string RtErpPrefix = "", RtErpNo = "", TransferStatusMES = "", ConfirmStatus = "";
                    int OrderCompanyIdBase = -1;

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
                            MESCompanyNo = item.ErpNo;
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
                        dynamicParameters.Add("FunctionCode", "ReturnReceiveOrderManagment");
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

                        #region //確認銷貨單是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.TransferStatusMES, a.ConfirmStatus, a.CompanyId, a.RtErpPrefix, a.RtErpNo
                                FROM SCM.ReturnReceiveOrder a
                                WHERE a.CompanyId = @CompanyId
                                AND a.RtId = @RtId
                                ";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("RtId", RtId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("銷貨單單頭資料找不到，請重新確認");
                        foreach (var item in result)
                        {
                            TransferStatusMES = item.TransferStatusMES;
                            ConfirmStatus = item.ConfirmStatus;
                            OrderCompanyIdBase = item.CompanyId;
                            RtErpPrefix = item.RtErpPrefix;
                            RtErpNo = item.RtErpNo;
                        }
                        if (TransferStatusMES == "N") throw new SystemException("銷退單處於未拋轉狀態,無法使用作廢功能");
                        if (ConfirmStatus == "Y") throw new SystemException("銷退單處於確認狀態,如欲修改銷貨單資料請使用反確認功能");
                        if (ConfirmStatus == "V") throw new SystemException("銷退單處於作廢狀態不可以異動");
                        if (OrderCompanyIdBase != CurrentCompany) throw new SystemException("單據的公司別與目前公司別不同，請嘗試重新登入!!");
                        #endregion

                        #region //更新單頭
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.ReturnReceiveOrder SET
                                ConfirmStatus = @ConfirmStatus,
                                TransferStatusMES = @TransferStatusMES,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RtId = @RtId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmStatus = "V",
                                TransferStatusMES = "Y",
                                LastModifiedDate,
                                LastModifiedBy,
                                RtId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //更新單身
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.RtDetail SET
                                ConfirmStatus = @ConfirmStatus,
                                TransferStatusMES = @TransferStatusMES,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RtId = @RtId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmStatus = "V",
                                TransferStatusMES = "Y",
                                LastModifiedDate,
                                LastModifiedBy,
                                RtId
                            });
                        insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //ERP公司代碼取得
                        string ERPCompanyNo = "";
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 ML001  
                                         FROM CMSML";
                        var resultCompanyNo = sqlConnection.Query(sql, dynamicParameters);
                        foreach (var item in resultCompanyNo)
                        {
                            ERPCompanyNo = item.ML001;
                        }
                        if (MESCompanyNo != ERPCompanyNo) throw new SystemException("ERP的公司代碼與MES的公司代碼不同，請嘗試重新登入!!");
                        #endregion

                        #region //判斷ERP使用者是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 MF005
                                        FROM ADMMF
                                        WHERE COMPANY = @CompanyNo
                                        AND MF001 = @userNo";
                        dynamicParameters.Add("CompanyNo", ERPCompanyNo);
                        dynamicParameters.Add("userNo", userNo);
                        var resultUserExist = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUserExist.Count() <= 0) throw new SystemException("【ERP使用者】不存在!");
                        #endregion

                        #region //審核ERP權限-建立
                        var USR_GROUP = BaseHelper.CheckErpAuthority(userNo, sqlConnection, "COPI08", "CREATE");
                        #endregion

                        #region //判斷銷退單是否存在+可不可以作廢
                        string erpConfirmStatus = "";
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TI019
                                FROM COPTI
                                WHERE TI001 =@RtErpPrefix
                                AND TI002 = @RtErpNo";
                        dynamicParameters.Add("RtErpPrefix", RtErpPrefix);
                        dynamicParameters.Add("RtErpNo", RtErpNo);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("銷退單不存在,請重新確認!!!");

                        foreach (var item in result)
                        {
                            erpConfirmStatus = item.TH020;
                        }
                        if (erpConfirmStatus == "Y") throw new SystemException("ERP銷退單處於確認狀態不可以異動!!!");
                        if (erpConfirmStatus == "V") throw new SystemException("ERP銷退單處於作廢狀態不可以異動!!!");
                        #endregion

                        #region //更新單頭
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE COPTI SET
                                FLAG = FLAG + 1,
                                MODIFIER = @MODIFIER,
                                MODI_DATE = @MODI_DATE,
                                MODI_TIME = @MODI_TIME,
                                MODI_AP = @MODI_AP,
                                MODI_PRID = @MODI_PRID,
                                TI019 = 'V'
                                WHERE TI001 =@RtErpPrefix
                                AND TI002 = @RtErpNo";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                MODIFIER = userNo,
                                MODI_DATE = dateNow,
                                MODI_TIME = timeNow,
                                MODI_AP = BaseHelper.ClientComputer(),
                                MODI_PRID = "BM",
                                RtErpPrefix,
                                RtErpNo
                            });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //更新單身
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE COPTJ SET
                                FLAG = FLAG + 1,
                                MODIFIER = @MODIFIER,
                                MODI_DATE = @MODI_DATE,
                                MODI_TIME = @MODI_TIME,
                                MODI_AP = @MODI_AP,
                                MODI_PRID = @MODI_PRID,
                                TJ021 = 'V'
                                WHERE TJ001 =@RtErpPrefix
                                AND TJ002 = @RtErpNo";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                MODIFIER = userNo,
                                MODI_DATE = dateNow,
                                MODI_TIME = timeNow,
                                MODI_AP = BaseHelper.ClientComputer(),
                                MODI_PRID = "BM",
                                RtErpPrefix,
                                RtErpNo
                            });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                    }

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "單據作廢成功!!!"
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

        #region //UpdateReturnRtConfirm -- 銷退單確認單據 -- Shintokuro 2024-05-21
        public string UpdateReturnRtConfirm(int RtId)
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
                int CompanyIdBase = -1;
                int SoDetailId = -1;
                int RoDetailId = -1;

                string RtErpPrefix = "";
                string RtErpNo = "";
                string ReturnDate = "";

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        string CustomerNo = "";
                        string Currency = "";
                        decimal TotalAmount = 0;
                        string CompanyNo = "";

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
                        dynamicParameters.Add("FunctionCode", "ReturnReceiveOrderManagment");
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
                        sql = @"SELECT a.RtErpPrefix,a.RtErpNo, a.ConfirmStatus,a.CompanyId
                                ,FORMAT(a.ReturnDate, 'yyyyMMdd') ReturnDate
                                ,b.SoDetailId,b.RoDetailId
                                ,c.CustomerNo, a.Currency, a.PretaxAmount ,d.CompanyNo
                                FROM SCM.ReturnReceiveOrder a
                                INNER JOIN SCM.RtDetail b on a.RtId = b.RtId
                                INNER JOIN SCM.Customer c on a.CustomerId = c.CustomerId
                                INNER JOIN BAS.Company d on a.CompanyId = d.CompanyId
                                WHERE a.RtId = @RtId";
                        dynamicParameters.Add("RtId", RtId);

                        var resultRoMes = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in resultRoMes)
                        {
                            if (item.ConfirmStatus == "Y") throw new SystemException("該銷退單單據已經核單,不能確認");
                            if (item.ConfirmStatus == "V") throw new SystemException("該銷退單單據已經作廢,不能確認");
                            if (item.CompanyId != CurrentCompany) throw new SystemException("該銷退單單據公司別與後端公司別不一致,不能確認");
                            CompanyIdBase = item.CompanyId;
                            RtErpPrefix = item.RtErpPrefix;
                            RtErpNo = item.RtErpNo;
                            ReturnDate = item.ReturnDate;
                            SoDetailId = item.SoDetailId;
                            RoDetailId = item.RoDetailId;
                            CustomerNo = item.CustomerNo;
                            Currency = item.Currency;
                            TotalAmount = Convert.ToDecimal(item.PretaxAmount);
                            CompanyNo = item.CompanyNo;
                        }
                        #endregion

                        #region //信用額度檢核
                        //string domainUrl = "http://192.168.20.136:16668/";
                        string domainUrl = "https://bm.zy-tech.com.tw/";


                        string targetDate = DateTime.Now.AddDays(-2).ToString("yyyyMMdd HH:mm:ss"),
                            targetUrl = string.Format("{0}{1}", domainUrl, "api/BM/CheckCreditLimit");

                        var postData = new List<KeyValuePair<string, string>>
                        {
                            new KeyValuePair<string, string>("CustomerNo", CustomerNo),
                            new KeyValuePair<string, string>("Currency", Currency),
                            new KeyValuePair<string, string>("TotalAmount", TotalAmount.ToString()),
                            new KeyValuePair<string, string>("DocType", "ShippingOrder"),
                            new KeyValuePair<string, string>("Amount", "0"),
                            new KeyValuePair<string, string>("CompanyNo", CompanyNo)
                        };

                        string response = BaseHelper.PostWebRequest(targetUrl, postData);

                        if (response.TryParseJson(out JObject tempJObject))
                        {
                            JObject resultJson = JObject.Parse(response);

                            Console.WriteLine("狀態：" + resultJson["status"].ToString());
                            Console.WriteLine("回傳訊息：" + resultJson["msg"].ToString());

                            if (resultJson["status"].ToString() != "ok")
                            {
                                throw new SystemException(resultJson["msg"].ToString());
                            }
                        }
                        else
                        {
                            logger.Error(response);
                        }
                        #endregion
                    }
                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //審核ERP權限
                        USR_GROUP = BaseHelper.CheckErpAuthority(UserNo, sqlConnection, "COPI09", "CONFIRM");
                        #endregion

                        #region //判斷開立日期是否超過結帳日
                        string CloseDateBase = "";
                        string ReturnDateBase = ReturnDate;
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

                        if (DateTime.TryParseExact(ReturnDateBase, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out docDateBase) &&
                            DateTime.TryParseExact(CloseDateBase, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out closeDateBase))
                        {
                            int resultDate = DateTime.Compare(docDateBase, closeDateBase);

                            if (resultDate <= 0)
                            {
                                throw new SystemException("【銷貨單】ERP已經結帳,不可以在新增【" + CloseDateBase + "】之前的單據");
                            }
                        }
                        else
                        {
                            throw new SystemException("日期字符串无效，无法比较");
                        }
                        #endregion

                        #region //取出銷退單單據資料 + 異動更新相關資料表
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.TI003,a.TI019 ,a.TI021,
                                b.TJ003,b.TJ004,b.TJ030,b.TJ041,b.TJ011,b.TJ033,
                                b.TJ013,b.TJ014,b.TJ049,b.TJ007,b.TJ042,b.TJ050,
                                b.TJ015,b.TJ016,b.TJ017,
                                b.TJ018,b.TJ019,b.TJ020,
                                b.TJ031,b.TJ032,b.TJ033,b.TJ034,
                                c.MB022,c.MB004,
                                ISNULL(md1.RtRate,1) RtRate, 
                                ISNULL(md2.RtAvailableRate,1) RtAvailableRate, 
                                (b.TJ007*ISNULL(md1.RtRate,1)) ConversionQty,
                                (b.TJ042*ISNULL(md1.RtRate,1)) ConversionFreeSpareQty,
                                (b.TJ050*ISNULL(md2.RtAvailableRate,1)) ConversionAvailableQty,
                                ROUND((b.TJ033 / (b.TJ050*ISNULL(md2.RtAvailableRate,1))),CAST(md3.MF005 AS INT)) UnitCost
                                FROM COPTI a
                                INNER JOIN COPTJ b on a.TI001 = b.TJ001 AND a.TI002 = b.TJ002
                                INNER JOIN INVMB c on b.TJ004 = c.MB001
                                OUTER APPLY(
                                    SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) RtRate
                                    FROM INVMD x1
                                    WHERE x1.MD001 = b.TJ004
                                    AND  x1.MD002 = b.TJ008
                                ) md1
                                OUTER APPLY(
                                    SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) RtAvailableRate
                                    FROM INVMD x1
                                    WHERE x1.MD001 = b.TJ004
                                    AND  x1.MD002 = b.TJ051
                                ) md2
                                OUTER APPLY(
                                    SELECT x1.MF003,x1.MF004,x1.MF005,x1.MF006
                                    FROM CMSMF x1
                                    INNER JOIN CMSMA x2 on x1.MF001 = x2.MA003 
                                ) md3
                                WHERE a.TI001 = @RtErpPrefix
                                AND a.TI002 = @RtErpNo";
                        dynamicParameters.Add("RtErpPrefix", RtErpPrefix);
                        dynamicParameters.Add("RtErpNo", RtErpNo);
                        var resultRtErp = sqlConnection.Query(sql, dynamicParameters);
                        if (resultRtErp.Count() <= 0) throw new SystemException("找不到ERP銷退單單據!");
                        foreach (var item in resultRtErp)
                        {
                            if (item.TI019 == "Y") throw new SystemException("該銷貨單單據已經核單,不能確認");
                            if (item.TI019 == "V") throw new SystemException("該銷貨單單據已經作廢,不能確認");

                            #region //欄位參數撈取
                            string ReturnDateErp = item.TI003; //TI003 銷退日
                            string ConfirmStatus = item.TI019; //TI019 單頭確認
                            string CustomerFullName = item.TI021; //TI021 客戶全名
                            string RtSequence = item.TJ003; //TJ003 序號
                            string MtlItemNo = item.TJ004; //TJ004 品號
                            string Type = item.TJ030; //TJ030 類型(1.銷貨/2.折舊)
                            string FreeSpareType = item.TJ041; //TJ041 數量類型(1.贈品/2.備品)
                            string InventoryNo = item.TJ013; //TJ013 庫別
                            string LotNumber = item.TJ014; //TJ014 批號
                            string Location = item.TJ049; //TJ049 儲位
                            double Quantity = Convert.ToDouble(item.TJ007); //TJ007 數量
                            double FreeSpareQty = Convert.ToDouble(item.TJ042); //TJ042 贈/備品數量
                            double AvailableQty = Convert.ToDouble(item.TJ050); //TJ050 計價數量
                            string TJ015 = item.TJ015; //TJ015 銷貨單別　/ TJ016 銷貨單號 / TJ017 銷貨序號
                            string TJ016 = item.TJ016;
                            string TJ017 = item.TJ017;
                            string RoFull = item.TJ015 + '-' + item.TJ016 + '-' + item.TJ017;
                            string TJ018 = item.TJ018; //TJ018 訂單單別 / TJ019 訂單單號 / TJ020 訂單序號
                            string TJ019 = item.TJ019;
                            string TJ020 = item.TJ020;
                            double TJ031 = Convert.ToDouble(item.TJ031); //原幣未稅金額  
                            double TJ032 = Convert.ToDouble(item.TJ032); //原幣稅額    
                            double TJ033 = Convert.ToDouble(item.TJ033); //本幣未稅金額  
                            double TJ034 = Convert.ToDouble(item.TJ034); //本幣稅額
                            string LotManagement = item.MB022; //品號批號管理
                            string InventoryUomNo = item.MB004; //庫存單位

                            string AllOverMistake = "";
                            string OverMtlItemNo = "";
                            string OverMistake = "";

                            double CanUseQty = 0;
                            double CanUseFreeSpareQty = 0;

                            double RtRate = Convert.ToDouble(item.RtRate); //銷退-單位換算率
                            double RtAvailableRate = Convert.ToDouble(item.RtAvailableRate); //銷退-計價單位換算率
                            double ConversionQty = Convert.ToDouble(item.ConversionQty); //銷退-單位換算後數量
                            double ConversionFreeSpareQty = Convert.ToDouble(item.ConversionFreeSpareQty); //銷退-單位換算後贈備品數量
                            double ConversionAvailableQty = Convert.ToDouble(item.ConversionAvailableQty); //銷退-單位換算後計價數量
                            double UnitPrice = Convert.ToDouble(item.TJ011); //原幣單價
                            double PreTaxAmt = Convert.ToDouble(item.TJ033); //本幣未稅金額
                            double UnitCost = Convert.ToDouble(item.UnitCost); //單位成本

                            double SoRate = 0;
                            double SoAvailableRate = 0;
                            double RoRate = 0;
                            double RoAvailableRate = 0;
                            #endregion

                            #region //檢查段
                            #region //數量是否有超收檢查
                            #region //撈取訂單 單位換算率
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 
                                    ISNULL(md1.SoRate,1) SoRate, 
                                    ISNULL(md2.SoAvailableRate,1) SoAvailableRate
                                    FROM COPTD a
                                    OUTER APPLY(
                                        SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) SoRate
                                        FROM INVMD x1
                                        WHERE x1.MD001= a.TD004
                                        AND  x1.MD002 = a.TD010
                                    ) md1
                                    OUTER APPLY(
                                        SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) SoAvailableRate
                                        FROM INVMD x1
                                        WHERE x1.MD001= a.TD004
                                        AND  x1.MD002 = a.TD077
                                    ) md2
                                    WHERE 1=1
                                    AND a.TD001 = @TJ018
                                    AND a.TD002 = @TJ019
                                    AND a.TD003 = @TJ020
                                    AND a.TD021 = 'Y'";
                            dynamicParameters.Add("TJ018", TJ018);
                            dynamicParameters.Add("TJ019", TJ019);
                            dynamicParameters.Add("TJ020", TJ020);
                            var resultCOPTD = sqlConnection.Query(sql, dynamicParameters);
                            if (resultCOPTD.Count() <= 0) throw new SystemException("訂單找不到,請重新確認");
                            foreach (var item1 in resultCOPTD)
                            {
                                SoRate = Convert.ToDouble(item1.SoRate);//數量單位轉換成庫存單位的換算率
                                SoAvailableRate = Convert.ToDouble(item1.SoAvailableRate);//計價單位轉換成庫存單位的換算率
                            }
                            #endregion

                            #region //撈取銷貨單數量資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 
                                    ISNULL(a.TH008,0) TH008,
                                    ISNULL(a.TH024,0) TH024,
                                    (ISNULL(a.TH008,0) * ISNULL(md1.RoRate,1)) - (ISNULL(a.TH043,0) * ISNULL(md1.RoRate,1)) CanUseQty,
                                    (ISNULL(a.TH024,0) * ISNULL(md1.RoRate,1)) - (ISNULL(a.TH052,0) * ISNULL(md1.RoRate,1)) CanUseFreeSpareQty,
                                    ISNULL(md1.RoRate,1) RoRate, 
                                    ISNULL(md2.RoAvailableRate,1) RoAvailableRate, 
                                    (ISNULL(a.TH008,0) * ISNULL(md1.RoRate,1)) RoConversionQty,
                                    (ISNULL(a.TH043,0) * ISNULL(md1.RoRate,1)) RoConversionSrQty,
                                    (ISNULL(a.TH024,0) * ISNULL(md1.RoRate,1)) RoConversionFreeSpareQty,
                                    (ISNULL(a.TH052,0) * ISNULL(md1.RoRate,1)) RoConversionFreeSpareSrQty
                                    FROM COPTH a
                                    OUTER APPLY(
                                        SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) RoRate
                                        FROM INVMD x1
                                        WHERE x1.MD001= a.TH004
                                        AND  x1.MD002 = a.TH009
                                    ) md1
                                    OUTER APPLY(
                                        SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) RoAvailableRate
                                        FROM INVMD x1
                                        WHERE x1.MD001= a.TH004
                                        AND  x1.MD002 = a.TH062
                                    ) md2
                                    WHERE a.TH001 = @TJ015
                                    AND a.TH002 = @TJ016
                                    AND a.TH003 = @TJ017
                                    AND a.TH020 = 'Y'";
                            dynamicParameters.Add("TJ015", TJ015);
                            dynamicParameters.Add("TJ016", TJ016);
                            dynamicParameters.Add("TJ017", TJ017);
                            var resultCOPTH = sqlConnection.Query(sql, dynamicParameters);
                            if (resultCOPTH.Count() <= 0) throw new SystemException("銷貨單找不到,請重新確認");
                            foreach (var item1 in resultCOPTH)
                            {
                                //if (item1.TH026 == "Y") throw new SystemException("【訂單:" + RoFull + "】已經結案!!!");
                                //TH008 = Convert.ToDouble(item1.TH008);
                                //TH024 = Convert.ToDouble(item1.TH024);
                                //TH061 = Convert.ToDouble(item1.TH061);
                                CanUseQty = Convert.ToDouble(item1.CanUseQty); //可使用數量
                                CanUseFreeSpareQty = Convert.ToDouble(item1.CanUseFreeSpareQty); //可使用贈備品數量
                                RoRate = Convert.ToDouble(item1.RoRate);//數量單位轉換成庫存單位的換算率
                                RoAvailableRate = Convert.ToDouble(item1.RoAvailableRate);//計價單位轉換成庫存單位的換算率
                                double RoConversionQty = Convert.ToDouble(item1.RoConversionQty); //轉換成庫存單位的銷貨數量
                                double RoConversionFreeSpareQty = Convert.ToDouble(item1.RoConversionFreeSpareQty);//轉換成庫存單位的銷貨贈備品數量

                                #region //判斷是否有超過可銷退數
                                OverMtlItemNo = "銷退單單身品號:" + MtlItemNo;
                                if ((CanUseQty - ConversionQty) < 0)
                                {
                                    OverMistake += "<br>已經超過可銷退數量<br>銷貨單數量:" + RoConversionQty + InventoryUomNo + ",目前可銷退最大數:" + CanUseQty + InventoryUomNo + "<br>欲新增銷貨數量:" + ConversionQty + InventoryUomNo + ",請重新確認";
                                }
                                if ((CanUseFreeSpareQty - ConversionFreeSpareQty) < 0)
                                {
                                    OverMistake += "<br>已經超過可銷退贈備品數量,銷貨單贈備品數量:" + RoConversionFreeSpareQty + InventoryUomNo + ",目前可銷退贈備品最大數量:" + CanUseFreeSpareQty + InventoryUomNo + "<br>欲新增銷退贈備品數量:" + ConversionFreeSpareQty + InventoryUomNo + ",請重新確認<br>";
                                }

                                if (OverMistake.Length > 0)
                                {
                                    AllOverMistake += OverMtlItemNo + OverMistake;
                                }
                                if (AllOverMistake.Length > 0)
                                {
                                    throw new SystemException(AllOverMistake);
                                }
                                #endregion

                            }
                            #endregion

                            #endregion

                            #region //品號有批號管理觸發
                            if (LotManagement != "N")
                            {
                                #region //INVMF 品號批號資料單身 判斷該品號批號庫存是否足夠
                                //dynamicParameters = new DynamicParameters();
                                //sql = @"SELECT SUM(MF008*MF010)  AS LotQty, SUM(MF008*MF014) AS K_QTY
                                //        FROM INVMF
                                //        WHERE MF001 = @MtlItemNo 
                                //        AND MF002 = @LotNumber 
                                //        AND MF007 = @InventoryNo";
                                //dynamicParameters.Add("MtlItemNo", MtlItemNo);
                                //dynamicParameters.Add("LotNumber", LotNumber);
                                //dynamicParameters.Add("InventoryNo", InventoryNo);

                                //var resultCheckLot = sqlConnection.Query(sql, dynamicParameters);
                                //if (resultCheckLot.Count() <= 0) throw new SystemException("找不到庫存資訊!!!");
                                //foreach (var item1 in resultCheckLot)
                                //{
                                //    if (Convert.ToDouble(item1.LotQty) < ConversionQty + ConversionFreeSpareQty) throw new SystemException("【品號：" + MtlItemNo + "庫別:" + InventoryNo + "批號:" + LotNumber + "】庫存不足請重新確認");
                                //}
                                #endregion

                                #region //INVLF 品號庫別儲位批號檔 判斷改品號批號庫別儲位庫存數量是否足夠
                                //dynamicParameters = new DynamicParameters();
                                //sql = @"SELECT SUM(LF008*LF011)  AS LotQty, SUM(LF008*LF012) AS K_QTY
                                //        FROM INVLF
                                //        WHERE LF004 = @MtlItemNo 
                                //        AND LF005 = @InventoryNo 
                                //        AND LF006 = @Location
                                //        AND LF007 = @LotNumber";
                                //dynamicParameters.Add("MtlItemNo", MtlItemNo);
                                //dynamicParameters.Add("InventoryNo", InventoryNo);
                                //dynamicParameters.Add("Location", Location != "" ? Location : "##########");
                                //dynamicParameters.Add("LotNumber", LotNumber);

                                //resultCheckLot = sqlConnection.Query(sql, dynamicParameters);
                                //if (resultCheckLot.Count() <= 0) throw new SystemException("找不到庫存資訊!!!");
                                //foreach (var item1 in resultCheckLot)
                                //{
                                //    if (Convert.ToDouble(item1.LotQty) < ConversionQty + ConversionFreeSpareQty) throw new SystemException("【品號：" + MtlItemNo + "庫別:" + InventoryNo + "儲位:" + Location != "" ? Location : "##########" + "批號:" + LotNumber + "】庫存不足請重新確認");
                                //}
                                #endregion
                            }
                            #endregion
                            #endregion

                            #region //異動更新段-銷退單確認後
                            #region //COPTD 回寫訂單單身 已交數量
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE COPTD SET
                                    MODIFIER = @MODIFIER,
                                    MODI_DATE = @MODI_DATE,
                                    MODI_TIME = @MODI_TIME,
                                    MODI_AP = @MODI_AP,
                                    MODI_PRID = @MODI_PRID,
                                    FLAG = FLAG + 1,
                                    TD009  = TD009 - @TD009,
                                    TD025  = TD025 - @TD025,
                                    TD051  = TD051 - @TD051,
                                    TD078  = TD078 - @TD078,
                                    TD016  = @TD016 
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
                                TD009 = Math.Round(ConversionQty / SoRate, 3, MidpointRounding.AwayFromZero),
                                TD025 = FreeSpareType == "1" ? Math.Round(ConversionFreeSpareQty / SoRate, 3, MidpointRounding.AwayFromZero) : 0,
                                TD051 = FreeSpareType == "2" ? Math.Round(ConversionFreeSpareQty / SoRate, 3, MidpointRounding.AwayFromZero) : 0,
                                TD078 = Math.Round(ConversionAvailableQty / SoAvailableRate, 3, MidpointRounding.AwayFromZero),
                                TD016 = "N",
                                TD001 = TJ018,
                                TD002 = TJ019,
                                TD003 = TJ020
                            });
                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //COPTH 銷貨單單身 銷退相關欄位
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE COPTH SET
                                    MODIFIER = @MODIFIER,
                                    MODI_DATE = @MODI_DATE,
                                    MODI_TIME = @MODI_TIME,
                                    MODI_AP = @MODI_AP,
                                    MODI_PRID = @MODI_PRID,
                                    FLAG = FLAG + 1,
                                    TH043  = TH043 + @TH043,
                                    TH052  = TH052 + @TH052,
                                    TH045  = TH045 + @TH045,
                                    TH046  = TH046 + @TH046,
                                    TH047  = TH047 + @TH047,
                                    TH048  = TH048 + @TH048
                                    WHERE TH001 = @TH001
                                    AND TH002 = @TH002
                                    AND TH003 = @TH003";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                MODIFIER = UserNo,
                                MODI_DATE = dateNow,
                                MODI_TIME = timeNow,
                                MODI_AP = BaseHelper.ClientComputer(),
                                MODI_PRID = "BM",
                                TH043 = Math.Round(ConversionQty / RoRate, 3, MidpointRounding.AwayFromZero),
                                TH052 = Math.Round(ConversionFreeSpareQty / RoRate, 3, MidpointRounding.AwayFromZero),
                                TH045 = TJ031,
                                TH046 = TJ032,
                                TH047 = TJ033,
                                TH048 = TJ034,
                                TH001 = TJ015,
                                TH002 = TJ016,
                                TH003 = TJ017
                            });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //INVLA 交易紀錄建立
                            #region //INVLB 取得月初總數量,月初總成本, 單據單位成本
                            double LB010 = 0;
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 LB003,LB004,LB010
                                    FROM INVLB 
                                    WHERE LB001 =@MtlItemNo
                                    AND LB010 != 0
                                    ORDER BY LB002 DESC";
                            dynamicParameters.Add("MtlItemNo", MtlItemNo);
                            var resultINVLB = sqlConnection.Query(sql, dynamicParameters);

                            if (resultINVLB.Count() > 0)
                            {
                                foreach (var item1 in resultINVLB)
                                {
                                    LB010 = Convert.ToDouble(item1.LB010);
                                }
                            }
                            #endregion

                            #region //金額小數位數取位
                            #region //取得本國幣別
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 LTRIM(RTRIM(MA003)) MA003 FROM CMSMA";

                            var resultCMSMA = sqlConnection.Query(sql, dynamicParameters);

                            if (resultCMSMA.Count() <= 0) throw new SystemException("共用參數檔案資料錯誤!!");

                            string Currency = "";
                            foreach (var item1 in resultCMSMA)
                            {
                                Currency = item1.MA003;
                            }
                            #endregion

                            #region //小數點後取位
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.MF004,a.MF005,a.MF006
                                    FROM CMSMF a 
                                    WHERE a.MF001 = @Currency";
                            dynamicParameters.Add("Currency", Currency);

                            var resultCMSMF = sqlConnection.Query(sql, dynamicParameters);

                            if (resultCMSMF.Count() <= 0) throw new SystemException("金額小數點取位資料錯誤!!");

                            int AmountRounding = 0; //金額取位
                            foreach (var item1 in resultCMSMF)
                            {
                                AmountRounding = Convert.ToInt32(item1.MF00ˋ);
                            }
                            #endregion
                            #endregion

                            #region //新增INVLA
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
                                    LA004 = ReturnDateErp,
                                    LA005 = 1,
                                    LA006 = RtErpPrefix,
                                    LA007 = RtErpNo,
                                    LA008 = RtSequence,
                                    LA009 = InventoryNo,
                                    LA010 = CustomerFullName,
                                    LA011 = ConversionQty + ConversionFreeSpareQty,
                                    LA012 = 0,
                                    LA013 = 0,
                                    LA014 = "2",
                                    LA015 = "N",
                                    LA016 = LotNumber,
                                    LA017 = 0,
                                    LA018 = 0,
                                    LA019 = 0,
                                    LA020 = 0,
                                    LA021 = 0,
                                    LA022 = Location != "" ? Location : ""
                                });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                            #endregion

                            #region //INVMB 品號基本資料檔 品號庫存更新
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE INVMB SET
                                    MODIFIER = @MODIFIER,
                                    MODI_DATE = @MODI_DATE,
                                    MODI_TIME = @MODI_TIME,
                                    MODI_AP = @MODI_AP,
                                    MODI_PRID = @MODI_PRID,
                                    FLAG = FLAG + 1,
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
                                MB064 = ConversionQty + ConversionFreeSpareQty,
                                MB065 = PreTaxAmt,
                                MB001 = MtlItemNo
                            });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //INVMC 品號庫別檔 品號庫別庫存更新
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE INVMC SET
                                    MODIFIER = @MODIFIER,
                                    MODI_DATE = @MODI_DATE,
                                    MODI_TIME = @MODI_TIME,
                                    MODI_AP = @MODI_AP,
                                    MODI_PRID = @MODI_PRID,
                                    FLAG = FLAG + 1,
                                    MC007  = MC007 + @MC007,
                                    MC013  = @MC013 
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
                                MC007 = ConversionQty + ConversionFreeSpareQty,
                                MC013 = ReturnDateErp,
                                MC001 = MtlItemNo,
                                MC002 = InventoryNo
                            });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //品號有批號管理觸發
                            if (LotManagement != "N")
                            {
                                #region //INVMF 品號批號資料單身 資料建立
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO INVMF (COMPANY, CREATOR, USR_GROUP, CREATE_DATE,
                                       MODIFIER, MODI_DATE, FLAG, UDF01, UDF02,
                                       UDF03, UDF04, UDF05, UDF06, UDF07,
                                       UDF08, UDF09, UDF10,
                                       MF001, MF002, MF003, MF004, MF005, 
                                       MF006, MF007, MF008, MF009, MF010, 
                                       MF014)
                                VALUES(@COMPANY, @CREATOR, @USR_GROUP,
                                       CONVERT(varchar(100),GETDATE(),112),@CREATOR,CONVERT(varchar(100),GETDATE(),112),0,
                                       '','','','','',0,0,0,0,0,
                                       @MF001 , @MF002 , @MF003 , @MF004 , @MF005,
                                       @MF006 , @MF007 , @MF008 , @MF009 , @MF010,
                                       @MF014)";
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
                                        MF003 = ReturnDateErp,
                                        MF004 = RtErpPrefix,
                                        MF005 = RtErpNo,
                                        MF006 = RtSequence,
                                        MF007 = InventoryNo,
                                        MF008 = 1,
                                        MF009 = "2",
                                        MF010 = ConversionQty + ConversionFreeSpareQty,
                                        MF014 = 0
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region //INVLF 儲位批號異動明細資料檔 資料建立
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
                                       @LF011, @LF012 , @LF013)";
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
                                        LF001 = RtErpPrefix,
                                        LF002 = RtErpNo,
                                        LF003 = RtSequence,
                                        LF004 = MtlItemNo,
                                        LF005 = InventoryNo,
                                        LF006 = Location != "" ? Location : "##########",
                                        LF007 = LotNumber,
                                        LF008 = 1,
                                        LF009 = ReturnDateErp,
                                        LF010 = "2",
                                        LF011 = ConversionQty + ConversionFreeSpareQty,
                                        LF012 = 0,
                                        LF013 = CustomerFullName
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region //INVMM 品號庫別儲位批號檔 庫存數量更新
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE INVMM SET
                                        MODIFIER = @MODIFIER,
                                        MODI_DATE = @MODI_DATE,
                                        MODI_TIME = @MODI_TIME,
                                        MODI_AP = @MODI_AP,
                                        MODI_PRID = @MODI_PRID,
                                        FLAG = FLAG + 1,
                                        MM005  = MM005 + @MM005,
                                        MM009  = @MM009
                                        WHERE MM001 = @MM001
                                        AND MM002 = @MM002
                                        AND MM003 = @MM003";
                                dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MODIFIER = UserNo,
                                    MODI_DATE = dateNow,
                                    MODI_TIME = timeNow,
                                    MODI_AP = BaseHelper.ClientComputer(),
                                    MODI_PRID = "BM",
                                    MM005 = ConversionQty + ConversionFreeSpareQty,
                                    MM009 = ReturnDateErp,
                                    MM001 = MtlItemNo,
                                    MM002 = InventoryNo,
                                    MM003 = Location != "" ? Location : "##########"
                                });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }
                            #endregion
                            #endregion

                        }
                        #endregion

                        #region //銷退單單據確認
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE COPTI SET 
                                MODIFIER = @MODIFIER,
                                MODI_DATE = @MODI_DATE,
                                MODI_TIME = @MODI_TIME,
                                MODI_AP = @MODI_AP,
                                MODI_PRID = @MODI_PRID,
                                FLAG = FLAG + 1,
                                TI019 = 'Y',
                                TI035 = @TI035
                                WHERE 1=1
                                AND TI001 = @RtErpPrefix
                                AND TI002 = @RtErpNo";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            MODIFIER = UserNo,
                            MODI_DATE = dateNow,
                            MODI_TIME = timeNow,
                            MODI_AP = BaseHelper.ClientComputer(),
                            MODI_PRID = "BM",
                            TI035 = UserNo,
                            RtErpPrefix,
                            RtErpNo
                        });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE COPTJ SET
                                MODIFIER = @MODIFIER,
                                MODI_DATE = @MODI_DATE,
                                MODI_TIME = @MODI_TIME,
                                MODI_AP = @MODI_AP,
                                MODI_PRID = @MODI_PRID,
                                FLAG = FLAG + 1,
                                TJ021 = 'Y'
                                WHERE 1=1
                                AND TJ001 = @RtErpPrefix
                                AND TJ002 = @RtErpNo";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            MODIFIER = UserNo,
                            MODI_DATE = dateNow,
                            MODI_TIME = timeNow,
                            MODI_AP = BaseHelper.ClientComputer(),
                            MODI_PRID = "BM",
                            RtErpPrefix,
                            RtErpNo
                        });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                    }
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //若成功則將相關資料寫回MES
                        #region //更新 - 銷退單單頭
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.ReturnReceiveOrder SET
                                TransferStatusMES = @TransferStatusMES,
                                ConfirmStatus = @ConfirmStatus,
                                ConfirmUserId = @ConfirmUserId,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RtId = @RtId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TransferStatusMES = "Y",
                                ConfirmStatus = "Y",
                                ConfirmUserId = LastModifiedBy,
                                LastModifiedDate,
                                LastModifiedBy,
                                RtId
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //更新 - 銷退單單身
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.RtDetail SET
                                ConfirmStatus = @ConfirmStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RtId = @RtId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmStatus = "Y",
                                ConfirmUserId = LastModifiedBy,
                                LastModifiedDate,
                                LastModifiedBy,
                                RtId
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region /撈取銷退單單身資料=>更新訂單,暫出單
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.FreeSpareType,a.FreeSpareQty,a.Quantity,a.AvailableQty,a.SoDetailId,a.RoDetailId
                                ,a.OrigPreTaxAmt,a.OrigTaxAmt,a.PreTaxAmt,a.TaxAmt
                                FROM SCM.RtDetail a
                                WHERE a.RtId = @RtId";
                        dynamicParameters.Add("RtId", RtId);

                        var resultRoDetail = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in resultRoDetail)
                        {
                            #region //更新 - 訂單 已交數量,贈品已交數量,備品已交量
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.SoDetail SET
                                    SiQty = SiQty - @SiQty,
                                    FreebieSiQty = FreebieSiQty - @FreebieSiQty,
                                    SpareSiQty = SpareSiQty - @SpareSiQty,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE SoDetailId = @SoDetailId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    SiQty = item.Quantity,
                                    FreebieSiQty = item.FreeSpareType == "1" ? item.FreeSpareQty : 0,
                                    SpareSiQty = item.FreeSpareType == "2" ? item.FreeSpareQty : 0,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    SoDetailId = item.SoDetailId
                                });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //更新 - 銷貨單 銷退量,銷退贈/備品量
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.RoDetail SET
                                        SrQty = SrQty + @SrQty,
                                        SrFreeSpareQty = SrFreeSpareQty + @SrFreeSpareQty,
                                        SrOrigPreTaxAmt = SrOrigPreTaxAmt + @SrOrigPreTaxAmt,
                                        SrOrigTaxAmt = SrOrigTaxAmt + @SrOrigTaxAmt,
                                        SrPreTaxAmt = SrPreTaxAmt + @SrPreTaxAmt,
                                        SrTaxAmt = SrTaxAmt + @SrTaxAmt,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE RoDetailId = @RoDetailId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    SrQty = item.Quantity,
                                    SrFreeSpareQty = item.FreeSpareQty,
                                    SrOrigPreTaxAmt = item.OrigPreTaxAmt,
                                    SrOrigTaxAmt = item.OrigTaxAmt,
                                    SrPreTaxAmt = item.PreTaxAmt,
                                    SrTaxAmt = item.TaxAmt,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    RoDetailId = item.RoDetailId
                                });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

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

        #region //UpdateReturnRtReconfirm -- 銷退單反確認單據 -- Shintokuro 2024-05-21
        public string UpdateReturnRtReconfirm(int RtId)
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
                int CompanyIdBase = -1;
                int SoDetailId = -1;
                int RoDetailId = -1;

                string RtErpPrefix = "";
                string RtErpNo = "";
                string ReturnDate = "";

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
                        dynamicParameters.Add("FunctionCode", "ReturnReceiveOrderManagment");
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
                        sql = @"SELECT a.RtErpPrefix,a.RtErpNo, a.ConfirmStatus,a.CompanyId
                                ,FORMAT(a.ReturnDate, 'yyyyMMdd') ReturnDate
                                ,b.SoDetailId,b.RoDetailId
                                FROM SCM.ReturnReceiveOrder a
                                INNER JOIN SCM.RtDetail b on a.RtId = b.RtId
                                WHERE a.RtId = @RtId";
                        dynamicParameters.Add("RtId", RtId);

                        var resultRoMes = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in resultRoMes)
                        {
                            if (item.ConfirmStatus != "Y") throw new SystemException("該銷退單單據非核單狀態,不能反確認");
                            if (item.ConfirmStatus == "V") throw new SystemException("該銷退單單據已經作廢,不能確認");
                            if (item.CompanyId != CurrentCompany) throw new SystemException("該銷退單單據公司別與後端公司別不一致,不能確認");
                            CompanyIdBase = item.CompanyId;
                            RtErpPrefix = item.RtErpPrefix;
                            RtErpNo = item.RtErpNo;
                            ReturnDate = item.ReturnDate;
                            SoDetailId = item.SoDetailId;
                            RoDetailId = item.RoDetailId;
                        }
                        #endregion
                    }
                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //審核ERP權限
                        USR_GROUP = BaseHelper.CheckErpAuthority(UserNo, sqlConnection, "COPI09", "CONFIRM");
                        #endregion

                        #region //判斷開立日期是否超過結帳日
                        string CloseDateBase = "";
                        string ReturnDateBase = ReturnDate;
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

                        if (DateTime.TryParseExact(ReturnDateBase, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out docDateBase) &&
                            DateTime.TryParseExact(CloseDateBase, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out closeDateBase))
                        {
                            int resultDate = DateTime.Compare(docDateBase, closeDateBase);

                            if (resultDate <= 0)
                            {
                                throw new SystemException("【銷貨單】ERP已經結帳,不可以在異動【" + CloseDateBase + "】之前的單據");
                            }
                        }
                        else
                        {
                            throw new SystemException("日期字符串无效，无法比较");
                        }
                        #endregion

                        #region //取出銷退單單據資料 + 異動更新相關資料表
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.TI003,a.TI019 ,a.TI021,
                                b.TJ003,b.TJ004,b.TJ030,b.TJ041,b.TJ011,b.TJ033,
                                b.TJ013,b.TJ014,b.TJ049,b.TJ007,b.TJ042,b.TJ050,
                                b.TJ015,b.TJ016,b.TJ017,
                                b.TJ018,b.TJ019,b.TJ020,
                                b.TJ031,b.TJ032,b.TJ033,b.TJ034,
                                c.MB022,c.MB004,
                                ISNULL(md1.RtRate,1) RtRate, 
                                ISNULL(md2.RtAvailableRate,1) RtAvailableRate, 
                                (b.TJ007*ISNULL(md1.RtRate,1)) ConversionQty,
                                (b.TJ042*ISNULL(md1.RtRate,1)) ConversionFreeSpareQty,
                                (b.TJ050*ISNULL(md2.RtAvailableRate,1)) ConversionAvailableQty,
                                ROUND((b.TJ033 / (b.TJ050*ISNULL(md2.RtAvailableRate,1))),CAST(md3.MF005 AS INT)) UnitCost
                                FROM COPTI a
                                INNER JOIN COPTJ b on a.TI001 = b.TJ001 AND a.TI002 = b.TJ002
                                INNER JOIN INVMB c on b.TJ004 = c.MB001
                                OUTER APPLY(
                                    SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) RtRate
                                    FROM INVMD x1
                                    WHERE x1.MD001 = b.TJ004
                                    AND  x1.MD002 = b.TJ008
                                ) md1
                                OUTER APPLY(
                                    SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) RtAvailableRate
                                    FROM INVMD x1
                                    WHERE x1.MD001 = b.TJ004
                                    AND  x1.MD002 = b.TJ051
                                ) md2
                                OUTER APPLY(
                                    SELECT x1.MF003,x1.MF004,x1.MF005,x1.MF006
                                    FROM CMSMF x1
                                    INNER JOIN CMSMA x2 on x1.MF001 = x2.MA003 
                                ) md3
                                WHERE a.TI001 = @RtErpPrefix
                                AND a.TI002 = @RtErpNo";
                        dynamicParameters.Add("RtErpPrefix", RtErpPrefix);
                        dynamicParameters.Add("RtErpNo", RtErpNo);
                        var resultRtErp = sqlConnection.Query(sql, dynamicParameters);
                        if (resultRtErp.Count() <= 0) throw new SystemException("找不到ERP銷退單單據!");
                        foreach (var item in resultRtErp)
                        {
                            if (item.TI019 != "Y") throw new SystemException("該銷退單單據非核單狀態,不能反確認");
                            if (item.TI019 == "V") throw new SystemException("該銷退單單據已經作廢,不能反確認");

                            #region //欄位參數撈取
                            string ReturnDateErp = item.TI003; //TI003 銷退日
                            string ConfirmStatus = item.TI019; //TI019 單頭確認
                            string CustomerFullName = item.TI021; //TI021 客戶全名
                            string RtSequence = item.TJ003; //TJ003 序號
                            string MtlItemNo = item.TJ004; //TJ004 品號
                            string Type = item.TJ030; //TJ030 類型(1.銷貨/2.折舊)
                            string FreeSpareType = item.TJ041; //TJ041 數量類型(1.贈品/2.備品)
                            string InventoryNo = item.TJ013; //TJ013 庫別
                            string LotNumber = item.TJ014; //TJ014 批號
                            string Location = item.TJ049; //TJ049 儲位
                            double Quantity = Convert.ToDouble(item.TJ007); //TJ007 數量
                            double FreeSpareQty = Convert.ToDouble(item.TJ042); //TJ042 贈/備品數量
                            double AvailableQty = Convert.ToDouble(item.TJ050); //TJ050 計價數量
                            string TJ015 = item.TJ015; //TJ015 銷貨單別　/ TJ016 銷貨單號 / TJ017 銷貨序號
                            string TJ016 = item.TJ016;
                            string TJ017 = item.TJ017;
                            string RoFull = item.TJ015 + '-' + item.TJ016 + '-' + item.TJ017;
                            string TJ018 = item.TJ018; //TJ018 訂單單別 / TJ019 訂單單號 / TJ020 訂單序號
                            string TJ019 = item.TJ019;
                            string TJ020 = item.TJ020;
                            double TJ031 = Convert.ToDouble(item.TJ031); //原幣未稅金額  
                            double TJ032 = Convert.ToDouble(item.TJ032); //原幣稅額    
                            double TJ033 = Convert.ToDouble(item.TJ033); //本幣未稅金額  
                            double TJ034 = Convert.ToDouble(item.TJ034); //本幣稅額
                            string LotManagement = item.MB022; //品號批號管理
                            string InventoryUomNo = item.MB004; //庫存單位

                            string AllOverMistake = "";
                            string OverMtlItemNo = "";
                            string OverMistake = "";

                            double CanUseQty = 0;
                            double CanUseFreeSpareQty = 0;

                            double RtRate = Convert.ToDouble(item.RtRate); //銷退-單位換算率
                            double RtAvailableRate = Convert.ToDouble(item.RtAvailableRate); //銷退-計價單位換算率
                            double ConversionQty = Convert.ToDouble(item.ConversionQty); //銷退-單位換算後數量
                            double ConversionFreeSpareQty = Convert.ToDouble(item.ConversionFreeSpareQty); //銷退-單位換算後贈備品數量
                            double ConversionAvailableQty = Convert.ToDouble(item.ConversionAvailableQty); //銷退-單位換算後計價數量
                            double UnitPrice = Convert.ToDouble(item.TJ011); //原幣單價
                            double PreTaxAmt = Convert.ToDouble(item.TJ033); //本幣未稅金額
                            double UnitCost = Convert.ToDouble(item.UnitCost); //單位成本

                            double SoRate = 0;
                            double SoAvailableRate = 0;
                            string ClosureStatusCOPTD = "N";
                            double RoRate = 0;
                            double RoAvailableRate = 0;
                            #endregion

                            #region //檢查段
                            #region //數量是否有超收檢查
                            #region //撈取訂單 單位換算率
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 TD049,a.TD016,TD010,TD077,TD008,TD009,TD076,TD078,
                                    ISNULL(a.TD008,0) TD008,
                                    ISNULL(a.TD024,0) TD024,
                                    ISNULL(a.TD050,0) TD050,
                                    (ISNULL(a.TD008,0) * ISNULL(md1.SoRate,1)) - (ISNULL(a.TD009,0) * ISNULL(md1.SoRate,1)) CanUseQty,
                                    (ISNULL(a.TD024,0) * ISNULL(md1.SoRate,1)) - (ISNULL(a.TD025,0) * ISNULL(md1.SoRate,1)) CanUseFreebieQty,
                                    (ISNULL(a.TD050,0) * ISNULL(md1.SoRate,1)) - (ISNULL(a.TD051,0) * ISNULL(md1.SoRate,1)) CanUseSpareQty,
                                    (ISNULL(a.TD076,0) * ISNULL(md2.SoAvailableRate,1)) - (ISNULL(a.TD078,0) * ISNULL(md2.SoAvailableRate,1)) CanUseAvailableQty,
                                    ISNULL(md1.SoRate,1) SoRate, 
                                    ISNULL(md2.SoAvailableRate,1) SoAvailableRate, 
                                    (ISNULL(a.TD008,0) * ISNULL(md1.SoRate,1)) SoConversionQty,
                                    (ISNULL(a.TD009,0) * ISNULL(md1.SoRate,1)) SoConversionSiQty,
                                    (ISNULL(a.TD024,0) * ISNULL(md1.SoRate,1)) SoConversionFreebieQty,
                                    (ISNULL(a.TD025,0) * ISNULL(md1.SoRate,1)) SoConversionFreebieSiQty,
                                    (ISNULL(a.TD050,0) * ISNULL(md1.SoRate,1)) SoConversionSpareQty,
                                    (ISNULL(a.TD051,0) * ISNULL(md1.SoRate,1)) SoConversionSpareSiQty,
                                    (ISNULL(a.TD076,0) * ISNULL(md2.SoAvailableRate,1)) SoConversionSoPriceQty,
                                    (ISNULL(a.TD078,0) * ISNULL(md2.SoAvailableRate,1)) SoConversionSoPriceSiQty
                                    FROM COPTD a
                                    OUTER APPLY(
                                        SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) SoRate
                                        FROM INVMD x1
                                        WHERE x1.MD001= a.TD004
                                        AND  x1.MD002 = a.TD010
                                    ) md1
                                    OUTER APPLY(
                                        SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) SoAvailableRate
                                        FROM INVMD x1
                                        WHERE x1.MD001= a.TD004
                                        AND  x1.MD002 = a.TD077
                                    ) md2
                                    WHERE 1=1
                                    AND a.TD001 = @TJ018
                                    AND a.TD002 = @TJ019
                                    AND a.TD003 = @TJ020
                                    AND a.TD021 = 'Y'";
                            dynamicParameters.Add("TJ018", TJ018);
                            dynamicParameters.Add("TJ019", TJ019);
                            dynamicParameters.Add("TJ020", TJ020);
                            var resultCOPTD = sqlConnection.Query(sql, dynamicParameters);
                            if (resultCOPTD.Count() <= 0) throw new SystemException("訂單找不到,請重新確認");
                            foreach (var item1 in resultCOPTD)
                            {
                                string TD049 = item1.TD049;
                                double TD008 = Convert.ToDouble(item1.TD008);
                                double TD024 = Convert.ToDouble(item1.TD024);
                                double TD050 = Convert.ToDouble(item1.TD050);
                                CanUseQty = Convert.ToDouble(item1.CanUseQty); //可使用訂單數量
                                double CanUseFreebieQty = Convert.ToDouble(item1.CanUseFreebieQty); //可使用贈品數量
                                double CanUseSpareQty = Convert.ToDouble(item1.CanUseSpareQty); //可使用備品數量
                                double CanUseAvailableQty = Convert.ToDouble(item1.CanUseAvailableQty); //可使用計價數量

                                SoRate = Convert.ToDouble(item1.SoRate);//數量單位轉換成庫存單位的換算率
                                SoAvailableRate = Convert.ToDouble(item1.SoAvailableRate);//計價單位轉換成庫存單位的換算率
                                double SoConversionQty = Convert.ToDouble(item1.SoConversionQty); //轉換成庫存單位的訂單數量
                                double SoConversionSiQty = Convert.ToDouble(item1.SoConversionSiQty);//轉換成庫存單位的已交訂單數量
                                double SoConversionFreebieQty = Convert.ToDouble(item1.SoConversionFreebieQty);//轉換成庫存單位的贈品數量
                                double SoConversionFreebieSiQty = Convert.ToDouble(item1.SoConversionFreebieSiQty);//轉換成庫存單位的已交贈品數量
                                double SoConversionSpareQty = Convert.ToDouble(item1.SoConversionSpareQty);//轉換成庫存單位的備品數量
                                double SoConversionSpareSiQty = Convert.ToDouble(item1.SoConversionSpareSiQty);//轉換成庫存單位的已交備品數量
                                double SoConversionSoPriceQty = Convert.ToDouble(item1.SoConversionSoPriceQty);//轉換成庫存單位的計價數量
                                double SoConversionSoPriceSiQty = Convert.ToDouble(item1.SoConversionSoPriceSiQty);//轉換成庫存單位的已交計價數量

                                #region //判斷是否有超過可銷貨數
                                OverMtlItemNo = "銷貨單單身品號:" + MtlItemNo;
                                if ((CanUseQty - ConversionQty) < 0)
                                {
                                    OverMistake += "<br>已經超過可銷貨數量<br>訂單數量:" + SoConversionQty + InventoryUomNo + ",目前可銷貨最大數:" + CanUseQty + InventoryUomNo + "<br>欲新增銷貨數量:" + ConversionQty + InventoryUomNo + ",請重新確認";
                                }
                                if (TD049 == "1")
                                {
                                    if ((CanUseFreebieQty - ConversionFreeSpareQty) < 0)
                                    {
                                        OverMistake += "<br>已經超過可銷貨贈品數量,訂單贈品數量:" + SoConversionFreebieQty + InventoryUomNo + ",目前可銷貨贈品最大數量:" + CanUseFreebieQty + InventoryUomNo + "<br>欲新增銷貨贈品數量:" + ConversionFreeSpareQty + InventoryUomNo + ",請重新確認<br>";
                                    }
                                    if (CanUseQty - Quantity == 0 && CanUseFreebieQty - FreeSpareQty == 0)
                                    {
                                        ClosureStatusCOPTD = "Y";
                                    }
                                }
                                else if (TD049 == "2")
                                {
                                    if ((CanUseSpareQty - ConversionFreeSpareQty) < 0)
                                    {
                                        OverMistake += "<br>已經超過可銷貨備品數量,訂單備品數量:" + SoConversionSpareQty + InventoryUomNo + ",目前可銷貨貨品最大數量:" + CanUseSpareQty + InventoryUomNo + "<br>欲新增銷貨品數量:" + ConversionFreeSpareQty + InventoryUomNo + ",請重新確認<br>";
                                    }
                                    if (CanUseQty - ConversionQty == 0 && CanUseSpareQty - ConversionFreeSpareQty == 0)
                                    {
                                        ClosureStatusCOPTD = "Y";
                                    }
                                }
                                if (OverMistake.Length > 0)
                                {
                                    AllOverMistake += OverMtlItemNo + OverMistake;
                                }

                                if (AllOverMistake.Length > 0)
                                {
                                    throw new SystemException(AllOverMistake);
                                }
                                #endregion
                            }
                            #endregion

                            #region //撈取銷貨單數量資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 
                                    ISNULL(a.TH008,0) TH008,
                                    ISNULL(a.TH024,0) TH024,
                                    (ISNULL(a.TH008,0) * ISNULL(md1.RoRate,1)) - (ISNULL(a.TH043,0) * ISNULL(md1.RoRate,1)) CanUseQty,
                                    (ISNULL(a.TH024,0) * ISNULL(md1.RoRate,1)) - (ISNULL(a.TH052,0) * ISNULL(md1.RoRate,1)) CanUseFreeSpareQty,
                                    ISNULL(md1.RoRate,1) RoRate, 
                                    ISNULL(md2.RoAvailableRate,1) RoAvailableRate, 
                                    (ISNULL(a.TH008,0) * ISNULL(md1.RoRate,1)) RoConversionQty,
                                    (ISNULL(a.TH043,0) * ISNULL(md1.RoRate,1)) RoConversionSrQty,
                                    (ISNULL(a.TH024,0) * ISNULL(md1.RoRate,1)) RoConversionFreeSpareQty,
                                    (ISNULL(a.TH052,0) * ISNULL(md1.RoRate,1)) RoConversionFreeSpareSrQty
                                    FROM COPTH a
                                    OUTER APPLY(
                                        SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) RoRate
                                        FROM INVMD x1
                                        WHERE x1.MD001= a.TH004
                                        AND  x1.MD002 = a.TH009
                                    ) md1
                                    OUTER APPLY(
                                        SELECT ROUND(CAST(ISNULL(MD004/MD003,1) AS FLOAT) ,3) RoAvailableRate
                                        FROM INVMD x1
                                        WHERE x1.MD001= a.TH004
                                        AND  x1.MD002 = a.TH062
                                    ) md2
                                    WHERE a.TH001 = @TJ015
                                    AND a.TH002 = @TJ016
                                    AND a.TH003 = @TJ017
                                    AND a.TH020 = 'Y'";
                            dynamicParameters.Add("TJ015", TJ015);
                            dynamicParameters.Add("TJ016", TJ016);
                            dynamicParameters.Add("TJ017", TJ017);
                            var resultCOPTH = sqlConnection.Query(sql, dynamicParameters);
                            if (resultCOPTH.Count() <= 0) throw new SystemException("銷貨單找不到,請重新確認");
                            foreach (var item1 in resultCOPTH)
                            {
                                //if (item1.TH026 == "Y") throw new SystemException("【訂單:" + RoFull + "】已經結案!!!");
                                //TH008 = Convert.ToDouble(item1.TH008);
                                //TH024 = Convert.ToDouble(item1.TH024);
                                //TH061 = Convert.ToDouble(item1.TH061);
                                CanUseQty = Convert.ToDouble(item1.CanUseQty); //可使用數量
                                CanUseFreeSpareQty = Convert.ToDouble(item1.CanUseFreeSpareQty); //可使用贈備品數量
                                RoRate = Convert.ToDouble(item1.RoRate);//數量單位轉換成庫存單位的換算率
                                RoAvailableRate = Convert.ToDouble(item1.RoAvailableRate);//計價單位轉換成庫存單位的換算率
                                double RoConversionQty = Convert.ToDouble(item1.RoConversionQty); //轉換成庫存單位的銷貨數量
                                double RoConversionFreeSpareQty = Convert.ToDouble(item1.RoConversionFreeSpareQty);//轉換成庫存單位的銷貨贈備品數量
                            }
                            #endregion

                            #endregion

                            #region //INVMC 品號庫別檔 搜尋該品號是否有庫別庫存
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT MC007
                                    FROM INVMC
                                    WHERE MC001 = @MtlItemNo
                                    AND MC002 = @InventoryNo";
                            dynamicParameters.Add("MtlItemNo", MtlItemNo);
                            dynamicParameters.Add("InventoryNo", InventoryNo);
                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【品號：" + MtlItemNo + "庫別:" + InventoryNo + "】找不到庫存資料,請重新確認!!");
                            foreach (var item1 in result)
                            {
                                if (Convert.ToDouble(item1.MC007) < ConversionQty + ConversionFreeSpareQty) throw new SystemException("【品號：" + MtlItemNo + "庫別:" + InventoryNo + "】庫存不足請重新確認");
                            }
                            #endregion


                            #region //品號有批號管理觸發
                            if (LotManagement != "N")
                            {
                                #region //INVMF 品號批號資料單身 判斷該品號批號庫存是否足夠
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT SUM(MF008*MF010)  AS LotQty, SUM(MF008*MF014) AS K_QTY
                                        FROM INVMF
                                        WHERE MF001 = @MtlItemNo 
                                        AND MF002 = @LotNumber 
                                        AND MF007 = @InventoryNo";
                                dynamicParameters.Add("MtlItemNo", MtlItemNo);
                                dynamicParameters.Add("LotNumber", LotNumber);
                                dynamicParameters.Add("InventoryNo", InventoryNo);

                                var resultCheckLot = sqlConnection.Query(sql, dynamicParameters);
                                if (resultCheckLot.Count() <= 0) throw new SystemException("找不到庫存資訊!!!");
                                foreach (var item1 in resultCheckLot)
                                {
                                    if (Convert.ToDouble(item1.LotQty) < ConversionQty + ConversionFreeSpareQty) throw new SystemException("【品號：" + MtlItemNo + "庫別:" + InventoryNo + "批號:" + LotNumber + "】庫存不足請重新確認");
                                }
                                #endregion

                                #region //INVLF 品號庫別儲位批號檔 判斷改品號批號庫別儲位庫存數量是否足夠
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT SUM(LF008*LF011)  AS LotQty, SUM(LF008*LF012) AS K_QTY
                                        FROM INVLF
                                        WHERE LF004 = @MtlItemNo 
                                        AND LF005 = @InventoryNo 
                                        AND LF006 = @Location
                                        AND LF007 = @LotNumber";
                                dynamicParameters.Add("MtlItemNo", MtlItemNo);
                                dynamicParameters.Add("InventoryNo", InventoryNo);
                                dynamicParameters.Add("Location", Location != "" ? Location : "##########");
                                dynamicParameters.Add("LotNumber", LotNumber);

                                resultCheckLot = sqlConnection.Query(sql, dynamicParameters);
                                if (resultCheckLot.Count() <= 0) throw new SystemException("找不到庫存資訊!!!");
                                foreach (var item1 in resultCheckLot)
                                {
                                    if (Convert.ToDouble(item1.LotQty) < ConversionQty + ConversionFreeSpareQty) throw new SystemException("【品號：" + MtlItemNo + "庫別:" + InventoryNo + "儲位:" + Location != "" ? Location : "##########" + "批號:" + LotNumber + "】庫存不足請重新確認");
                                }
                                #endregion
                            }
                            #endregion
                            #endregion

                            #region //異動更新段-銷退單確認後
                            #region //COPTD 回寫訂單單身 已交數量
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE COPTD SET
                                    MODIFIER = @MODIFIER,
                                    MODI_DATE = @MODI_DATE,
                                    MODI_TIME = @MODI_TIME,
                                    MODI_AP = @MODI_AP,
                                    MODI_PRID = @MODI_PRID,
                                    FLAG = FLAG + 1,
                                    TD009  = TD009 + @TD009,
                                    TD025  = TD025 + @TD025,
                                    TD051  = TD051 + @TD051,
                                    TD078  = TD078 + @TD078,
                                    TD016  = @TD016 
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
                                TD009 = Math.Round(ConversionQty / SoRate, 3, MidpointRounding.AwayFromZero),
                                TD025 = FreeSpareType == "1" ? Math.Round(ConversionFreeSpareQty / SoRate, 3, MidpointRounding.AwayFromZero) : 0,
                                TD051 = FreeSpareType == "2" ? Math.Round(ConversionFreeSpareQty / SoRate, 3, MidpointRounding.AwayFromZero) : 0,
                                TD078 = Math.Round(ConversionAvailableQty / SoAvailableRate, 3, MidpointRounding.AwayFromZero),
                                TD016 = ClosureStatusCOPTD,
                                TD001 = TJ018,
                                TD002 = TJ019,
                                TD003 = TJ020
                            });
                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //COPTH 銷貨單單身 銷退相關欄位
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE COPTH SET
                                    MODIFIER = @MODIFIER,
                                    MODI_DATE = @MODI_DATE,
                                    MODI_TIME = @MODI_TIME,
                                    MODI_AP = @MODI_AP,
                                    MODI_PRID = @MODI_PRID,
                                    FLAG = FLAG + 1,
                                    TH043  = TH043 - @TH043,
                                    TH052  = TH052 - @TH052,
                                    TH045  = TH045 - @TH045,
                                    TH046  = TH046 - @TH046,
                                    TH047  = TH047 - @TH047,
                                    TH048  = TH048 - @TH048
                                    WHERE TH001 = @TH001
                                    AND TH002 = @TH002
                                    AND TH003 = @TH003";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                MODIFIER = UserNo,
                                MODI_DATE = dateNow,
                                MODI_TIME = timeNow,
                                MODI_AP = BaseHelper.ClientComputer(),
                                MODI_PRID = "BM",
                                TH043 = Math.Round(ConversionQty / RoRate, 3, MidpointRounding.AwayFromZero),
                                TH052 = Math.Round(ConversionFreeSpareQty / RoRate, 3, MidpointRounding.AwayFromZero),
                                TH045 = TJ031,
                                TH046 = TJ032,
                                TH047 = TJ033,
                                TH048 = TJ034,
                                TH001 = TJ015,
                                TH002 = TJ016,
                                TH003 = TJ017
                            });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //INVLA 交易紀錄刪除
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
                                LA004 = ReturnDate,
                                LA005 = 1,
                                LA006 = RtErpPrefix,
                                LA007 = RtErpNo,
                                LA008 = RtSequence
                            });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //INVMB 品號基本資料檔 品號庫存更新
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE INVMB SET
                                    MODIFIER = @MODIFIER,
                                    MODI_DATE = @MODI_DATE,
                                    MODI_TIME = @MODI_TIME,
                                    MODI_AP = @MODI_AP,
                                    MODI_PRID = @MODI_PRID,
                                    FLAG = FLAG + 1,
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
                                MB064 = ConversionQty + ConversionFreeSpareQty,
                                MB065 = PreTaxAmt,
                                MB001 = MtlItemNo
                            });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //INVMC 品號庫別檔 品號庫別庫存更新
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE INVMC SET
                                    MODIFIER = @MODIFIER,
                                    MODI_DATE = @MODI_DATE,
                                    MODI_TIME = @MODI_TIME,
                                    MODI_AP = @MODI_AP,
                                    MODI_PRID = @MODI_PRID,
                                    FLAG = FLAG + 1,
                                    MC007  = MC007 - @MC007,
                                    MC013  = @MC013 
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
                                MC007 = ConversionQty + ConversionFreeSpareQty,
                                MC013 = ReturnDateErp,
                                MC001 = MtlItemNo,
                                MC002 = InventoryNo
                            });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //品號有批號管理觸發
                            if (LotManagement != "N")
                            {
                                #region //INVMF 品號批號資料單身 資料刪除
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
                                    MF003 = ReturnDate,
                                    MF004 = RtErpPrefix,
                                    MF005 = RtErpNo,
                                    MF006 = RtSequence
                                });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region //INVLF 儲位批號異動明細資料檔 資料刪除
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE INVLF
                                        WHERE LF001 = @LF001
                                          AND LF002 = @LF002
                                          AND LF003 = @LF003
                                          AND LF004 = @LF004
                                          AND LF005 = @LF005
                                          AND LF006 = @LF006
                                          AND LF007 = @LF007";
                                dynamicParameters.AddDynamicParams(
                                new
                                {
                                    LF001 = RtErpPrefix,
                                    LF002 = RtErpNo,
                                    LF003 = RtSequence,
                                    LF004 = MtlItemNo,
                                    LF005 = InventoryNo,
                                    LF006 = Location != "" ? Location : "##########",
                                    LF007 = LotNumber
                                });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region //INVMM 品號庫別儲位批號檔 庫存數量更新
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE INVMM SET
                                        MODIFIER = @MODIFIER,
                                        MODI_DATE = @MODI_DATE,
                                        MODI_TIME = @MODI_TIME,
                                        MODI_AP = @MODI_AP,
                                        MODI_PRID = @MODI_PRID,
                                        FLAG = FLAG + 1,
                                        MM005  = MM005 - @MM005,
                                        MM009  = @MM009
                                        WHERE MM001 = @MM001
                                        AND MM002 = @MM002
                                        AND MM003 = @MM003";
                                dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MODIFIER = UserNo,
                                    MODI_DATE = dateNow,
                                    MODI_TIME = timeNow,
                                    MODI_AP = BaseHelper.ClientComputer(),
                                    MODI_PRID = "BM",
                                    MM005 = ConversionQty + ConversionFreeSpareQty,
                                    MM009 = ReturnDateErp,
                                    MM001 = MtlItemNo,
                                    MM002 = InventoryNo,
                                    MM003 = Location != "" ? Location : "##########"
                                });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }
                            #endregion
                            #endregion

                        }
                        #endregion

                        #region //銷退單單據確認
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE COPTI SET 
                                MODIFIER = @MODIFIER,
                                MODI_DATE = @MODI_DATE,
                                MODI_TIME = @MODI_TIME,
                                MODI_AP = @MODI_AP,
                                MODI_PRID = @MODI_PRID,
                                FLAG = FLAG + 1,
                                TI019 = 'N'
                                WHERE 1=1
                                AND TI001 = @RtErpPrefix
                                AND TI002 = @RtErpNo";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            MODIFIER = UserNo,
                            MODI_DATE = dateNow,
                            MODI_TIME = timeNow,
                            MODI_AP = BaseHelper.ClientComputer(),
                            MODI_PRID = "BM",
                            RtErpPrefix,
                            RtErpNo
                        });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE COPTJ SET
                                MODIFIER = @MODIFIER,
                                MODI_DATE = @MODI_DATE,
                                MODI_TIME = @MODI_TIME,
                                MODI_AP = @MODI_AP,
                                MODI_PRID = @MODI_PRID,
                                FLAG = FLAG + 1,
                                TJ021 = 'N'
                                WHERE 1=1
                                AND TJ001 = @RtErpPrefix
                                AND TJ002 = @RtErpNo";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            MODIFIER = UserNo,
                            MODI_DATE = dateNow,
                            MODI_TIME = timeNow,
                            MODI_AP = BaseHelper.ClientComputer(),
                            MODI_PRID = "BM",
                            RtErpPrefix,
                            RtErpNo
                        });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                    }
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //若成功則將相關資料寫回MES
                        #region //更新 - 銷退單單頭
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.ReturnReceiveOrder SET
                                TransferStatusMES = @TransferStatusMES,
                                ConfirmStatus = @ConfirmStatus,
                                ConfirmUserId = @ConfirmUserId,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RtId = @RtId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TransferStatusMES = "R",
                                ConfirmStatus = "N",
                                ConfirmUserId = LastModifiedBy,
                                LastModifiedDate,
                                LastModifiedBy,
                                RtId
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //更新 - 銷退單單身
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.RtDetail SET
                                ConfirmStatus = @ConfirmStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RtId = @RtId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmStatus = "N",
                                ConfirmUserId = LastModifiedBy,
                                LastModifiedDate,
                                LastModifiedBy,
                                RtId
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region /撈取銷退單單身資料=>更新訂單,暫出單
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.FreeSpareType,a.FreeSpareQty,a.Quantity,a.AvailableQty,a.SoDetailId,a.RoDetailId
                                ,a.OrigPreTaxAmt,a.OrigTaxAmt,a.PreTaxAmt,a.TaxAmt
                                FROM SCM.RtDetail a
                                WHERE a.RtId = @RtId";
                        dynamicParameters.Add("RtId", RtId);

                        var resultRoDetail = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in resultRoDetail)
                        {
                            #region //更新 - 訂單 已交數量,贈品已交數量,備品已交量
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.SoDetail SET
                                    SiQty = SiQty + @SiQty,
                                    FreebieSiQty = FreebieSiQty + @FreebieSiQty,
                                    SpareSiQty = SpareSiQty + @SpareSiQty,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE SoDetailId = @SoDetailId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    SiQty = item.Quantity,
                                    FreebieSiQty = item.FreeSpareType == "1" ? item.FreeSpareQty : 0,
                                    SpareSiQty = item.FreeSpareType == "2" ? item.FreeSpareQty : 0,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    SoDetailId = item.SoDetailId
                                });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //更新 - 銷貨單 銷退量,銷退贈/備品量
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.RoDetail SET
                                        SrQty = SrQty - @SrQty,
                                        SrFreeSpareQty = SrFreeSpareQty - @SrFreeSpareQty,
                                        SrOrigPreTaxAmt = SrOrigPreTaxAmt - @SrOrigPreTaxAmt,
                                        SrOrigTaxAmt = SrOrigTaxAmt - @SrOrigTaxAmt,
                                        SrPreTaxAmt = SrPreTaxAmt - @SrPreTaxAmt,
                                        SrTaxAmt = SrTaxAmt - @SrTaxAmt,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE RoDetailId = @RoDetailId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    SrQty = item.Quantity,
                                    SrFreeSpareQty = item.FreeSpareQty,
                                    SrOrigPreTaxAmt = item.OrigPreTaxAmt,
                                    SrOrigTaxAmt = item.OrigTaxAmt,
                                    SrPreTaxAmt = item.PreTaxAmt,
                                    SrTaxAmt = item.TaxAmt,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    RoDetailId = item.RoDetailId
                                });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

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

        #endregion

        #region //Delete
        #region //DeleteReceiveOrder-- 銷貨單刪除 -- Shintokuro 2024.02.06
        public string DeleteReceiveOrder(int RoId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        int rowsAffected = 0;

                        #region //判斷銷貨單資料是否存在
                        sql = @"SELECT TOP 1 TransferStatusMES
                                FROM SCM.ReceiveOrder
                                WHERE CompanyId = @CompanyId
                                AND RoId = @RoId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("RoId", RoId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("銷貨單資料找不到,請重新確認!");
                        foreach(var item in result)
                        {
                            if(item.TransferStatusMES != "N") throw new SystemException("銷貨單未拋轉ERP狀態下,才能刪除!!!");
                        }
                        #endregion

                        #region //刪除次要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.RoDetail
                                WHERE RoId = @RoId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("RoId", RoId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.ReceiveOrder
                                WHERE CompanyId = @CompanyId
                                AND RoId = @RoId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("RoId", RoId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteRoDetail- 銷貨單單身刪除 -- Shintokuro 2024.02.06
        public string DeleteRoDetail(string RoDetailList)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        int rowsAffected = 0;
                        List<int> RoDetailIdList = RoDetailList.Split(',').Select(int.Parse).ToList();
                        foreach (var RoDetailId in RoDetailIdList)
                        {
                            #region //判斷銷貨單資料是否存在
                            sql = @"SELECT TOP 1 b.TransferStatusMES
                                    FROM SCM.RoDetail a
                                    INNER JOIN SCM.ReceiveOrder b on a.RoId = b.RoId
                                    WHERE b.CompanyId = @CompanyId
                                    AND a.RoDetailId = @RoDetailId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("RoDetailId", RoDetailId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("銷貨單單身資料找不到,請重新確認!");
                            foreach (var item in result)
                            {
                                if (item.TransferStatusMES != "N") throw new SystemException("銷貨單單身需處於未拋轉ERP狀態下,才能刪除!!!");
                            }
                            #endregion

                            #region //刪除主要table
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE SCM.RoDetail
                                    WHERE RoDetailId = @RoDetailId";
                            dynamicParameters.Add("RoDetailId", RoDetailId);

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

        #region //DeleteReceiveOrderAllMES-- MES銷貨單全刪除 -- Shintokuro 2024.02.23
        public string DeleteReceiveOrderAllMES()
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    string MESCompanyNo = "";
                    int rowsAffected = 0;
                    List<string> AllReceiveOrdersList = new List<string>();

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();


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
                            MESCompanyNo = item.ErpNo;
                        }
                        #endregion

                        #region //判斷銷貨單資料是否存在
                        sql = @"SELECT (RoErpPrefix + '-' + RoErpNo) RoErpFull
                                FROM SCM.ReceiveOrder";

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("銷貨單資料找不到,請重新確認!");
                        foreach (var item in result)
                        {
                            AllReceiveOrdersList.Add(item.RoErpFull);
                        }
                        #endregion

                        #region //刪除次要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.RoDetail";
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.ReceiveOrder";
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        string sqlErp = "";
                        foreach(var item in AllReceiveOrdersList)
                        {
                            #region //刪除次要table
                            dynamicParameters = new DynamicParameters();
                            sqlErp += @" DELETE COPTG
                                        WHERE TG001 = "+ item.Split('-')[0] + @"
                                        AND TG002 = " + item.Split('-')[1] + @"
                                        ";
                            #endregion

                            #region //刪除主要table
                            dynamicParameters = new DynamicParameters();
                            sqlErp += @" DELETE COPTH
                                        WHERE TH001 = " + item.Split('-')[0] + @"
                                        AND TH002 = " + item.Split('-')[1] + @"
                                        ";
                            #endregion
                        }

                        rowsAffected += sqlConnection.Execute(sqlErp, dynamicParameters);

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

        #region //DeleteReturnReceiveOrder-- 銷退單刪除 -- Shintokuro 2024.05.10
        public string DeleteReturnReceiveOrder(int RtId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        int rowsAffected = 0;

                        #region //判斷銷貨單資料是否存在
                        sql = @"SELECT TOP 1 TransferStatusMES
                                FROM SCM.ReturnReceiveOrder
                                WHERE CompanyId = @CompanyId
                                AND RtId = @RtId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("RtId", RtId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("銷貨單資料找不到,請重新確認!");
                        foreach (var item in result)
                        {
                            if (item.TransferStatusMES != "N") throw new SystemException("銷貨單未拋轉ERP狀態下,才能刪除!!!");
                        }
                        #endregion

                        #region //刪除次要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.RtDetail
                                WHERE RtId = @RtId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("RtId", RtId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.ReturnReceiveOrder
                                WHERE CompanyId = @CompanyId
                                AND RtId = @RtId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("RtId", RtId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteReturnRtDetail- 銷貨單單身刪除 -- Shintokuro 2024.02.06
        public string DeleteReturnRtDetail(string RtDetailList)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        int rowsAffected = 0;
                        List<int> RtDetailIdList = RtDetailList.Split(',').Select(int.Parse).ToList();
                        foreach (var RtDetailId in RtDetailIdList)
                        {
                            #region //判斷銷貨單資料是否存在
                            sql = @"SELECT TOP 1 b.TransferStatusMES
                                    FROM SCM.RtDetail a
                                    INNER JOIN SCM.ReturnReceiveOrder b on a.RtId = b.RtId
                                    WHERE b.CompanyId = @CompanyId
                                    AND a.RtDetailId = @RtDetailId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("RtDetailId", RtDetailId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("銷貨單單身資料找不到,請重新確認!");
                            foreach (var item in result)
                            {
                                if (item.TransferStatusMES != "N") throw new SystemException("銷貨單單身需處於未拋轉ERP狀態下,才能刪除!!!");
                            }
                            #endregion

                            #region //刪除主要table
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE SCM.RtDetail
                                    WHERE RtDetailId = @RtDetailId";
                            dynamicParameters.Add("RtDetailId", RtDetailId);

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

        #region //API

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

        #endregion



        #region //Model
        public class DoDetail
        {
            public int DoDetailId { get; set; }
            public string Inventory { get; set; }

            public string SoFullNo { get; set; }
            public int Qty { get; set; }
            public int FreeQty { get; set; }
            public int SpareQty { get; set; }


            public DoDetail(int DoDetailId, string Inventory,string SoFullNo, int Qty, int FreeQty, int SpareQty)
            {
                this.DoDetailId = DoDetailId; 
                this.Inventory = Inventory;
                this.SoFullNo = SoFullNo;
                this.Qty = Qty;
                this.FreeQty = FreeQty;
                this.SpareQty = SpareQty;
            }
        }
        #endregion

        #endregion
    }
}
