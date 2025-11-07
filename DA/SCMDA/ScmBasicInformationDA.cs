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
    public class ScmBasicInformationDA
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

        public ScmBasicInformationDA()
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
        #region //GetPackingVolume -- 取得包材體積資料 -- Zoey 2022.06.09
        public string GetPackingVolume(int VolumeId, string VolumeNo, string VolumeName, string VolumeType, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.VolumeId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.VolumeNo, a.VolumeName, a.VolumeSpec, a.Volume, c.TypeNo VolumeType, c.TypeName VolumeTypeName, a.Status
                          , b.UomNo UomNo, b.UomId VolumeUomId";
                    sqlQuery.mainTables =
                        @"FROM SCM.PackingVolume a
                          LEFT JOIN PDM.UnitOfMeasure b ON b.UomId = a.VolumeUomId
                          LEFT JOIN BAS.[Type] c ON c.TypeNo = a.VolumeType AND c.TypeSchema = 'PackingVolume.VolumeType'";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "VolumeId", @" AND a.VolumeId = @VolumeId", VolumeId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "VolumeNo", @" AND a.VolumeNo LIKE '%' + @VolumeNo + '%'", VolumeNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "VolumeName", @" AND a.VolumeName LIKE '%' + @VolumeName + '%'", VolumeName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "VolumeType", @" AND a.VolumeType = @VolumeType", VolumeType);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.VolumeId";
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

        #region //GetPackingWeight -- 取得包材重量資料 -- Zoey 2022.06.10
        public string GetPackingWeight(int WeightId, string WeightNo, string WeightName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.WeightId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.WeightNo, a.WeightName, a.WeightSpec, a.Weight, a.Status
                          , b.UomNo UomNo, b.UomId WeightUomId";
                    sqlQuery.mainTables =
                        @"FROM SCM.PackingWeight a
                          LEFT JOIN PDM.UnitOfMeasure b ON b.UomId = a.WeightUomId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "WeightId", @" AND a.WeightId = @WeightId", WeightId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "WeightNo", @" AND a.WeightNo LIKE '%' + @WeightNo + '%'", WeightNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "WeightName", @" AND a.WeightName LIKE '%' + @WeightName + '%'", WeightName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.WeightId";
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

        #region //GetItemWeight -- 取得物件重量資料 -- Zoey 2022.06.10
        public string GetItemWeight(int ItemDefaultWeightId, int MtlModelId, string Status, int StartParent
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    string declareSql = @"DECLARE @rowsAdded int

                                          DECLARE @mtlModel TABLE
                                          ( 
                                            MtlModelId int,
                                            ParentId int,
                                            MtlModelLevel int,
                                            MtlModelRoute nvarchar(MAX),
                                            MtlModelSort nvarchar(MAX),
                                            processed int DEFAULT(0)
                                          )

                                          INSERT @mtlModel
                                              SELECT MtlModelId, ParentId, 1 MtlModelLevel
                                              , CAST(ParentId AS nvarchar(MAX)) AS MtlModelRoute
                                              , CAST(MtlModelSort AS nvarchar(MAX)) AS MtlModelSort, 0
                                              FROM PDM.MtlModel
                                              WHERE CompanyId = @CompanyId
                                              AND ParentId = @StartParent

                                          SET @rowsAdded=@@rowcount

                                          WHILE @rowsAdded > 0
                                          BEGIN

                                            UPDATE @mtlModel SET processed = 1 WHERE processed = 0

                                            INSERT @mtlModel
                                                SELECT a.MtlModelId, a.ParentId, ( b.MtlModelLevel + 1 ) MtlModelLevel
                                                , CAST(b.MtlModelRoute + ',' + CAST(a.ParentId AS nvarchar(MAX)) AS nvarchar(MAX)) AS MtlModelRoute
                                                , CAST(b.MtlModelSort + CAST(a.MtlModelSort AS nvarchar(MAX)) AS nvarchar(MAX)) AS MtlModelSort, 0
                                                FROM PDM.MtlModel a
                                                INNER JOIN @mtlModel b ON a.ParentId = b.MtlModelId
                                                WHERE a.ParentId <> a.MtlModelId 
                                                AND b.processed = 1

                                            SET @rowsAdded = @@rowcount

                                            UPDATE @mtlModel SET processed = 2 WHERE processed = 1

                                          END;";

                    sqlQuery.mainKey = "a.ItemDefaultWeightId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.MtlModelId, a.DefaultWeight, a.DefaultWeightUomId, a.Status
                          , b.MtlModelRoute   
                          , c.MtlModelName, c.ParentId
                          , d.UomNo";
                    sqlQuery.mainTables =
                        @"FROM SCM.ItemDefaultWeight a
                          INNER JOIN @mtlModel b ON b.MtlModelId = a.MtlModelId
                          INNER JOIN PDM.MtlModel c ON c.MtlModelId = a.MtlModelId
                          INNER JOIN PDM.UnitOfMeasure d ON d.UomId = a.DefaultWeightUomId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    dynamicParameters.Add("StartParent", StartParent);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ItemDefaultWeightId", @" AND a.ItemDefaultWeightId = @ItemDefaultWeightId", ItemDefaultWeightId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlModelId", @" AND a.MtlModelId = @MtlModelId", MtlModelId);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.declarePart = declareSql;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ItemDefaultWeightId";
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

        #region //GetCustomer -- 取得客戶資料 -- Zoey 2022.06.15
        public string GetCustomer(int CustomerId, string CustomerNo, string CustomerName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
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
                          , a.CustomerNo + ' ' + a.CustomerShortName CustomerWithNo
                          , ISNULL(b.UserNo, '') SalesmenNo, ISNULL(b.UserName ,'') SalesmenName";
                    sqlQuery.mainTables =
                        @"FROM SCM.Customer a
                          LEFT JOIN BAS.[User] b ON b.UserId = a.SalesmenId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerId", @" AND a.CustomerId = @CustomerId", CustomerId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerNo", @" AND a.CustomerNo LIKE '%' + @CustomerNo + '%'", CustomerNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerName", @" AND a.CustomerName LIKE '%' + @CustomerName + '%'", CustomerName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
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

        #region //GetForwarder -- 取得貨運承攬商資料 -- Zoey 2022.07.01
        public string GetForwarder(int ForwarderId, string ShipMethod, string ForwarderNo, string ForwarderName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.ForwarderId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ShipMethod, a.ForwarderNo, a.ForwarderName, a.Status";
                    sqlQuery.mainTables =
                        @"FROM SCM.Forwarder a";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ForwarderId", @" AND a.ForwarderId = @ForwarderId", ForwarderId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ShipMethod", @" AND a.ShipMethod = @ShipMethod", ShipMethod);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ForwarderNo", @" AND a.ForwarderNo LIKE '%' + @ForwarderNo + '%'", ForwarderNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ForwarderName", @" AND a.ForwarderName LIKE '%' + @ForwarderName + '%'", ForwarderName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ForwarderId";
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

        #region //GetDeliveryCustomer -- 取得送貨客戶資料 -- Zoey 2022.07.01
        public string GetDeliveryCustomer(int DcId, int CustomerId, string DcName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.DcId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.CustomerId, a.DcName, a.DcEnglishName, a.DcShortName
                          , a.Contact, a.TelNo, a.FaxNo, a.RegisteredAddress, a.DeliveryAddress
                          , a.ShipType, a.ForwarderId, a.Status
                          , b.CustomerName, b.CustomerShortName";
                    sqlQuery.mainTables =
                        @"FROM SCM.DeliveryCustomer a
                          LEFT JOIN SCM.Customer b ON b.CustomerId = a.CustomerId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND b.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DcId", @" AND a.DcId = @DcId", DcId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerId", @" AND a.CustomerId = @CustomerId", CustomerId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DcName", @" AND a.DcName LIKE '%' + @DcName + '%'", DcName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.DcId";
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

        #region //GetSupplier -- 取得供應商資料 -- Zoey 2022.07.04
        public string GetSupplier(int SupplierId, string SupplierNo, string SupplierName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.SupplierId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.SupplierNo, a.SupplierName, a.SupplierShortName, a.SupplierEnglishName
                          , a.Country, a.Region, a.GuiNumber, a.ResponsiblePerson, a.ContactFirst, a.ContactSecond, a.ContactThird
                          , a.TelNoFirst, a.TelNoSecond, a.FaxNo, a.FaxNoAccounting, a.Email, a.AddressFirst, a.AddressSecond
                          , a.ZipCodeFirst, a.ZipCodeSecond, a.BillAddressFirst, a.BillAddressSecond, a.PermitStatus
                          , FORMAT(a.InauguateDate, 'yyyy-MM-dd') InauguateDate, a.AccountMonth, a.AccountDay, a.Version
                          , a.Capital, a.Headcount, a.PoDeliver, a.Currency, a.TradeTerm, a.PaymentType, a.PaymentTerm
                          , a.ReceiptReceive, a.InvoiceCount, a.TaxNo, a.Taxation, a.PermitPartialDelivery, a.TaxAmountCalculateType
                          , a.InvocieAttachedStatus, a.CertificateFormatType, a.DepositRate, a.TradeItem, a.RemitBank, a.RemitAccount
                          , a.AccountPayable, a.AccountOverhead, a.AccountInvoice, a.SupplierLevel, a.DeliveryRating, a.QualityRating
                          , a.RelatedPerson, a.PurchaseUserId, a.SupplierRemark, a.PassStationControl, a.TransferStatus, a.TransferDate, a.Status
                          , a.SupplierNo + ' ' + a.SupplierName SupplierWithNo
                          , ISNULL(b.UserName ,'') PurchaseUserName";
                    sqlQuery.mainTables =
                        @"FROM SCM.Supplier a
                          LEFT JOIN BAS.[User] b ON b.UserId = a.PurchaseUserId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SupplierId", @" AND a.SupplierId = @SupplierId", SupplierId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SupplierNo", @" AND a.SupplierNo LIKE '%' + @SupplierNo + '%'", SupplierNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SupplierName", @" AND a.SupplierName LIKE '%' + @SupplierName + '%'", SupplierName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.SupplierId";
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

        #region //GetInventory -- 取得庫別資料 -- Zoey 2022.07.13
        public string GetInventory(int InventoryId, string InventoryNo, string InventoryName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.InventoryId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.InventoryNo, a.InventoryName, a.InventoryType, a.MrpCalculation
                          , a.ConfirmStatus, a.SaveStatus, a.InventoryDesc, a.TransferStatus
                          , a.Status, a.InventoryNo + ' ' + a.InventoryName InventoryWithNo";
                    sqlQuery.mainTables =
                        @"FROM SCM.Inventory a";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "InventoryId", @" AND a.InventoryId = @InventoryId", InventoryId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "InventoryNo", @" AND a.InventoryNo LIKE '%' + @InventoryNo + '%'", InventoryNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "InventoryName", @" AND a.InventoryName LIKE '%' + @InventoryName + '%'", InventoryName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.InventoryId";
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

        #region //GetPacking -- 取得包材資料 -- Zoey 2022.10.11
        public string GetPacking(int PackingId, string PackingName, string PackingType, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.PackingId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.PackingName, a.PackingType, a.VolumeSpec, a.Volume
                          , a.VolumeUomId, a.WeightSpec, a.Weight, a.WeightUomId, a.Status
                          , a.PackingName + ' (' + a.VolumeSpec + ')' PackingNameWithVolume
                          , b.TypeName PackingTypeName
                          , c.UomName VolumeUomNo
                          , d.UomName WeightUomNo";
                    sqlQuery.mainTables =
                        @"FROM SCM.Packing a
                          LEFT JOIN BAS.[Type] b ON b.TypeNo = a.PackingType AND b.TypeSchema = 'Packing.PackingType'
                          LEFT JOIN PDM.UnitOfMeasure c ON c.UomId = a.VolumeUomId
                          LEFT JOIN PDM.UnitOfMeasure d ON d.UomId = a.WeightUomId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PackingId", @" AND a.PackingId = @PackingId", PackingId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PackingName", @" AND a.PackingName LIKE '%' + @PackingName + '%'", PackingName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PackingType", @" AND a.PackingType = @PackingType", PackingType);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.PackingId";
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

        #region //GetRfqProductClass -- 取得RFQ產品類型 -- Chia Yuan 2023.06.30
        public string GetRfqProductClass(int RfqProClassId, string RfqProductClassName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.RfqProClassId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.RfqProductClassName, a.[Status], a.CreateDate, a.LastModifiedDate, a.CreateBy, a.LastModifiedBy, b.StatusName";
                    sqlQuery.mainTables =
                        @"FROM SCM.RfqProductClass a
                        INNER JOIN BAS.[Status] b on b.StatusNo = a.[Status] and b.StatusSchema = 'Status'";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RfqProClassId", @" AND a.RfqProClassId = @RfqProClassId", RfqProClassId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RfqProductClassName", @" AND a.RfqProductClassName LIKE '%' + @RfqProductClassName + '%'", RfqProductClassName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.RfqProClassId";
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

        #region //GetRfqProductType -- 取得RFQ產品類別 -- Chia Yuan 2023.07.03

        public string GetRfqProductType(int RfqProTypeId, int RfqProClassId, string RfqProductTypeName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.RfqProTypeId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.RfqProductTypeName, a.[Status], a.CreateDate, a.LastModifiedDate, a.CreateBy, a.LastModifiedBy, b.StatusName, 'Y' as CanEdit
                        , c.RfqProductClassName + '-' + a.RfqProductTypeName as RfqProClsTypeName";
                    sqlQuery.mainTables =
                        @"FROM SCM.RfqProductType a
                        INNER JOIN BAS.[Status] b ON b.StatusNo = a.[Status] AND b.StatusSchema = 'Status'
                        INNER JOIN SCM.RfqProductClass c ON c.RfqProClassId = a.RfqProClassId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RfqProTypeId", @" AND a.RfqProTypeId = @RfqProTypeId", RfqProTypeId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RfqProClassId", @" AND a.RfqProClassId = @RfqProClassId", RfqProClassId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RfqProductTypeName", @" AND a.RfqProductTypeName LIKE '%' + @RfqProductTypeName + '%'", RfqProductTypeName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.RfqProTypeId";
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

        #region //GetRfqPackageType -- 取得RFQ包裝種類 -- Chia Yuan 2023.07.03

        public string GetRfqPackageType(int RfqPkTypeId, int RfqProClassId, string PackagingMethod, string Status, string SustSupplyStatus
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.RfqPkTypeId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.PackagingMethod, a.[Status], a.SustSupplyStatus, a.CreateDate, a.LastModifiedDate, a.CreateBy, a.LastModifiedBy
                        , b.StatusName, c.StatusName as SustSupplyStatusName, 'Y' as CanEdit, d.RfqProductClassName";
                    sqlQuery.mainTables =
                        @"FROM SCM.RfqPackageType a
                        INNER JOIN BAS.[Status] b on b.StatusNo = a.[Status] and b.StatusSchema = 'Status'
                        INNER JOIN BAS.[Status] c on c.StatusNo = a.SustSupplyStatus and c.StatusSchema = 'Boolean'
                        INNER JOIN SCM.RfqProductClass d ON d.RfqProClassId = a.RfqProClassId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RfqPkTypeId", @" AND a.RfqPkTypeId = @RfqPkTypeId", RfqPkTypeId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RfqProClassId", @" AND a.RfqProClassId = @RfqProClassId", RfqProClassId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PackagingMethod", @" AND a.PackagingMethod LIKE '%' + @PackagingMethod + '%'", PackagingMethod);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    if (SustSupplyStatus.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SustSupplyStatus", @" AND a.SustSupplyStatus IN @SustSupplyStatus", SustSupplyStatus.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.RfqPkTypeId";
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

        #region //GetProductUse -- 取得RFQ產品用途 -- Chia Yuan 2023.07.11
        public string GetProductUse(int ProductUseId, string ProductUseNo, string ProductUseName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.ProductUseId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ProductUseNo, a.ProductUseName, a.[Status], b.StatusName
                        , a.TypeOne, t1.TypeName AS ProductUseTypeName";
                    sqlQuery.mainTables =
                        @"FROM SCM.ProductUse a
                        INNER JOIN BAS.[Status] b on b.StatusNo = a.[Status] and b.StatusSchema = 'Status'
                        INNER JOIN BAS.[Type] t1 ON t1.TypeNo = a.TypeOne AND t1.TypeSchema = 'ProductUse.Type'";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProductUseId", @" AND a.ProductUseId = @ProductUseId", ProductUseId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProductUseNo", @" AND a.ProductUseNo LIKE '%' + @ProductUseNo + '%'", ProductUseNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProductUseName", @" AND a.ProductUseName LIKE '%' + @ProductUseName + '%'", ProductUseName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ProductUseName, a.ProductUseId";
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

        #region //GetMember -- 取得電商註冊資訊 -- Yi 2023.07.17
        public string GetMember(int MemberId, string MemberName, string OrgShortName, string OrganizaitonType, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.MemberId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.MemberName, a.MemberEmail, a.[Description], a.OrgShortName, a.Address, a.ContactName
                        , a.ContactPhone, a.ContactEmail, a.CertCode, a.Address, a.[Status], c.StatusName
                        , b.OrgId, b.OrganizaitonTypeId, b.OrganizaitonType, d.TypeName OrganizaitonTypeName
                        , b.OrganizationCode, b.OrganizaitonScale, e.TypeName OrganizaitonScaleName";
                    sqlQuery.mainTables =
                        @"FROM EIP.Member a
                        LEFT JOIN EIP.MemberOrganization b ON b.MemberId = a.MemberId
                        LEFT JOIN BAS.[Status] c ON c.StatusNo = a.[Status] AND c.StatusSchema = 'EipMember.Status'
                        LEFT JOIN BAS.[Type] d ON d.TypeNo = b.OrganizaitonType AND d.TypeSchema = 'MemberOrganization.OrganizaitonType'
                        LEFT JOIN BAS.[Type] e ON e.TypeNo = b.OrganizaitonScale AND e.TypeSchema = 'MemberOrganization.OrganizaitonScale'";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MemberId", @" AND a.MemberId = @MemberId", MemberId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MemberName", @" AND a.MemberName LIKE '%' + @MemberName + '%'", MemberName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "OrgShortName", @" AND a.OrgShortName LIKE '%' + @OrgShortName + '%'", OrgShortName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "OrganizaitonType", @" AND b.OrganizaitonType = @OrganizaitonType", OrganizaitonType);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
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

        #region //GetPerformanceGoals -- 取得業務績效目標 -- Shintokuro 2023.11.29
        public string GetPerformanceGoals(int PgId, string PgNo, string PgName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.PgId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.PgNo, a.PgName, a.PgDesc
                          , FORMAT(a.StartDate, 'yyyy-MM-dd') StartDate
                          , FORMAT(a.EndDate, 'yyyy-MM-dd') EndDate
                          , a.Status";
                    sqlQuery.mainTables =
                        @"FROM SCM.PerformanceGoals a";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PgId", @" AND a.PgId = @PgId", PgId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PgNo", @" AND a.PgNo LIKE '%' + @PgNo + '%'", PgNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PgName", @" AND a.PgName LIKE '%' + @PgName + '%'", PgName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.PgId";
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

        #region //GetPgDetail -- 取得業務績效目標單身 -- Shintokuro 2023.11.29
        public string GetPgDetail(int PgId, int PgDetailId, int UserId, string IntentType
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.PgDetailId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.PgId, a.UserId, a.IntentType, a.IntentAmount, a.ConfirmStatus
                          , FORMAT(a.ConfirmDate, 'yyyy-MM-dd') ConfirmDate
                          , c.UserName";
                    sqlQuery.mainTables =
                        @"FROM SCM.PgDetail a
                          INNER JOIN SCM.PerformanceGoals b on a.PgId = b.PgId
                          INNER JOIN BAS.[User] c on a.UserId = c.UserId
                          ";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PgId", @" AND a.PgId = @PgId", PgId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PgDetailId", @" AND a.PgDetailId = @PgDetailId", PgDetailId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserId", @" AND a.UserId = @UserId", UserId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "IntentType", @" AND a.IntentType = @IntentType", IntentType);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.PgDetailId";
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

        #region //GetProductTypeGroup -- 取得產品群組資料 -- Chia Yuan 2024.05.24
        public string GetProductTypeGroup(int ProTypeGroupId, int RfqProClassId, string ProTypeGroupName, string CoatingFlag, string Status
            , string SearchKey
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.ProTypeGroupId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ProTypeGroupName, a.[Status], s.StatusName, a.CoatingFlag, s2.StatusName CoatingFlagName
                        , CASE a.Status WHEN 'A' THEN 'true' ELSE 'false' END Checked
                        , ISNULL(b.UserName, '') UserName, a.TypeOne, t1.TypeName GroupTypeName
                        , FORMAT(a.CreateDate, 'yyyy-MM-dd HH:mm:ss') CreateDate
                        , ISNULL(FORMAT(a.LastModifiedDate, 'yyyy-MM-dd HH:mm:ss'), '') LastModifiedDate
                        , d.RfqProClassId, d.RfqProductClassName
                        , (
	                        SELECT aa.ParameterName, aa.ParameterEName, aa.SortNumber, aa.ControlType, aa.[Required], aa.[Status], ab.StatusName
	                        FROM RFI.TemplateSpecParameter aa
	                        INNER JOIN BAS.[Status] ab ON ab.StatusNo = aa.[Status]
	                        WHERE aa.ProTypeGroupId = a.ProTypeGroupId
	                        ORDER BY aa.SortNumber
	                        FOR JSON PATH, ROOT('data')
                        ) TempProdSpec";
                    sqlQuery.mainTables =
                        @" FROM SCM.ProductTypeGroup a
                        INNER JOIN SCM.RfqProductClass d ON d.RfqProClassId = a.RfqProClassId
                        LEFT JOIN BAS.[User] b ON b.UserId = a.CreateBy
                        LEFT JOIN BAS.[User] c ON c.UserId = a.LastModifiedBy
                        INNER JOIN BAS.[Status] s ON s.StatusNo = a.[Status] AND s.StatusSchema = 'Status'
                        INNER JOIN BAS.[Type] t1 ON t1.TypeNo = a.TypeOne AND t1.TypeSchema = 'ProductTypeGroup.Attribute'
                        INNER JOIN BAS.[Status] s2 ON s2.StatusNo = a.CoatingFlag AND s2.StatusSchema = 'Boolean'";
                    sqlQuery.auxTables = "";
                    string queryCondition = ""; // AND b.CompanyId = @CompanyId
                    //dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProTypeGroupId", @" AND a.ProTypeGroupId = @ProTypeGroupId", ProTypeGroupId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RfqProClassId", @" AND a.RfqProClassId = @RfqProClassId", RfqProClassId);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    if (CoatingFlag.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CoatingFlag", @" AND a.CoatingFlag IN @CoatingFlag", CoatingFlag.Split(','));
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey",
                        @" AND (a.ProTypeGroupName LIKE N'%' + @SearchKey + '%'
                            OR d.RfqProductClassName LIKE N'%' + @SearchKey + '%')", SearchKey);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ProTypeGroupId";
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

        #region //GetProductType -- 取得產品類別資料 -- Shintokuro 2024.05.24
        public string GetProductType(int ProTypeId, int ProTypeGroupId, string ProTypeName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.ProTypeId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ProTypeGroupId, a.ProTypeName, a.Status
                          , b.ProTypeGroupName
                          ";
                    sqlQuery.mainTables =
                        @"FROM SCM.ProductType a
                          INNER JOIN SCM.ProductTypeGroup b on a.ProTypeGroupId = b.ProTypeGroupId
                          INNER JOIN SCM.RfqProductClass c on b.RfqProClassId = c.RfqProClassId
                        ";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND c.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProTypeId", @" AND a.ProTypeId = @ProTypeId", ProTypeId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProTypeGroupId", @" AND a.ProTypeGroupId = @ProTypeGroupId", ProTypeGroupId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProTypeName", @" AND a.ProTypeName LIKE '%' + @ProTypeName + '%'", ProTypeName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ProTypeId";
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

        #region //GetTemplateRfiSignFlow -- 取得市場評估單流程 -- Shintokuro 2024.05.27
        public string GetTemplateRfiSignFlow(int TempRfiSfId, int ProTypeGroupId, int DepartmentId, int FlowUser, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.TempRfiSfId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ProTypeGroupId, a.FlowUser, a.SortNumber
                        , a.[Status], s1.StatusName
                        , ISNULL(u.UserId, -1) FlowUserId, ISNULL(u.UserName, '') FlowUserName, ISNULL(u.UserNo, '') FlowUserNo, ISNULL(u.Gender, '') FlowGender
                        , ISNULL(u.Job, '') FlowJob, ISNULL(u.JobType, '') FlowJobType, ISNULL(u.Email, '') FlowEmail
                        , ISNULL(a.FlowJobName, '') FlowJobName, a.FlowStatus, s2.StatusName FlowStatusName
                        , ISNULL(d.DepartmentNo, '') DepartmentNo, ISNULL(d.DepartmentName, '') DepartmentName
                        , ISNULL(c.CompanyNo, '') CompanyNo, ISNULL(c.CompanyName, '') CompanyName, ISNULL(c.LogoIcon, -1) LogoIcon";
                    sqlQuery.mainTables =
                        @"FROM RFI.TemplateRfiSignFlow a
                        INNER JOIN SCM.ProductTypeGroup a1 on a.ProTypeGroupId = a1.ProTypeGroupId
                        INNER JOIN SCM.RfqProductClass a2 on a1.RfqProClassId = a2.RfqProClassId
                        LEFT JOIN BAS.[User] u on a.FlowUser = u.UserId
                        LEFT JOIN BAS.Department d ON d.DepartmentId = u.DepartmentId
                        LEFT JOIN BAS.Company c ON c.CompanyId = d.CompanyId
                        INNER JOIN BAS.[Status] s1 ON s1.StatusNo = a.[Status] AND s1.StatusSchema = 'Status'
                        LEFT JOIN BAS.[Status] s2 ON s2.StatusNo = a.FlowStatus AND s2.StatusSchema = 'RfiDetail.FlowStatus'";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TempRfiSfId", @" AND a.TempRfiSfId = @TempRfiSfId", TempRfiSfId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProTypeGroupId", @" AND a.ProTypeGroupId = @ProTypeGroupId", ProTypeGroupId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DepartmentId", @" AND a.DepartmentId = @DepartmentId", DepartmentId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "FlowUser", @" AND a.FlowUser = @FlowUser", FlowUser);
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

        #region //GetTemplateDesignSignFlow -- 取得設計申請單流程 -- Shintokuro 2024.05.27
        public string GetTemplateDesignSignFlow(int TempDesignSfId, int ProTypeGroupId, int DepartmentId, int FlowUser, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.TempDesignSfId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ProTypeGroupId, a.FlowUser, a.SortNumber
                        , a.[Status], s1.StatusName
                        , ISNULL(u.UserId, -1) AS FlowUserId, ISNULL(u.UserName, '') AS FlowUserName, ISNULL(u.UserNo, '') AS FlowUserNo, ISNULL(u.Gender, '') AS FlowGender
                        , ISNULL(u.Job, '') AS FlowJob, ISNULL(u.JobType, '') AS FlowJobType, ISNULL(u.Email, '') AS FlowEmail
                        , ISNULL(a.FlowJobName, '') AS FlowJobName, a.FlowStatus, s2.StatusName AS FlowStatusName
                        , ISNULL(d.DepartmentNo, '') AS DepartmentNo, ISNULL(d.DepartmentName, '') AS DepartmentName
                        , ISNULL(c.CompanyNo, '') AS CompanyNo, ISNULL(c.CompanyName, '') AS CompanyName, ISNULL(c.LogoIcon, -1) AS LogoIcon";
                    sqlQuery.mainTables =
                        @"FROM RFI.TemplateDesignSignFlow a
                        INNER JOIN SCM.ProductTypeGroup a1 on a.ProTypeGroupId = a1.ProTypeGroupId
                        INNER JOIN SCM.RfqProductClass a2 on a1.RfqProClassId = a2.RfqProClassId
                        LEFT JOIN BAS.[User] u on a.FlowUser = u.UserId
                        LEFT JOIN BAS.Department d ON d.DepartmentId = u.DepartmentId
                        LEFT JOIN BAS.Company c ON c.CompanyId = d.CompanyId
                        INNER JOIN BAS.[Status] s1 ON s1.StatusNo = a.[Status] AND s1.StatusSchema = 'Status'
                        LEFT JOIN BAS.[Status] s2 ON s2.StatusNo = a.FlowStatus AND s2.StatusSchema = 'RfiDesign.FlowStatus'";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TempDesignSfId", @" AND a.TempDesignSfId = @TempDesignSfId", TempDesignSfId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProTypeGroupId", @" AND a.ProTypeGroupId = @ProTypeGroupId", ProTypeGroupId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DepartmentId", @" AND a.DepartmentId = @DepartmentId", DepartmentId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "FlowUser", @" AND a.FlowUser = @FlowUser", FlowUser);
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

        #region //GetQuotationTag -- 取得報價單屬性標籤 -- Shintokuro 2024.05.24
        public string GetQuotationTag(int QtId, string TagNo, string TagName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.QtId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.TagNo, a.TagName, a.Status
                          ";
                    sqlQuery.mainTables =
                        @"FROM SCM.QuotationTag a
                        ";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QtId", @" AND a.QtId = @QtId", QtId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TagNo", @" AND a.TagNo LIKE '%' + @TagNo + '%'", TagNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TagName", @" AND a.TagName LIKE '%' + @TagName + '%'", TagName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.QtId";
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

        #region//GetSupplierProcessEquipment --參與托外掃碼的供應商已登記的製程與機台清單 -- Andrew 2024.10.16
        public string GetSupplierProcessEquipment(int SmId, int SupplierId, int ProcessId, int MachineId, int ShopId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.SmId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @"  , d.SupplierId, d.SupplierNo, d.SupplierShortName
                            , c.ProcessId, c.ProcessName
                            , b.MachineId, b.MachineName, b.MachineDesc, e.ShopName, a.Status";
                    sqlQuery.mainTables =
                        @" FROM SCM.SupplierMachine a
                            LEFT JOIN MES.Machine  b ON b.MachineId=a.MachineId
                            LEFT JOIN MES.Process c ON a.ProcessId=c.ProcessId
                            LEFT JOIN SCM.Supplier d ON d.SupplierId=a.SupplierId
                            LEFT JOIN MES.WorkShop e ON e.ShopId=b.ShopId";
                    string queryTable = @"";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND c.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SmId", @" AND a.SmId = @SmId", SmId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SupplierId", @" AND d.SupplierId = @SupplierId", SupplierId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProcessId", @" AND c.ProcessId = @ProcessId", ProcessId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MachineId", @" AND b.MachineId = @MachineId", MachineId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ShopId", @" AND b.ShopId = @ShopId", ShopId);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.SmId ASC";
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

        #region //GetProcess -- 取得製程資料 -- 曾永至 2024.10.16
        public string GetProcess(int SupplierId, int ProcessId, string ProcessNo, string ProcessName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sql = @"SELECT DISTINCT b.ProcessId, b.ProcessNo, b.ProcessName, b.ProcessDesc, b.Status, (b.ProcessNo + '-' + b.ProcessName) ProcessWithNo 
                         FROM SCM.SupplierMachine a
                         INNER JOIN MES.Process b ON a.ProcessId = b.ProcessId";

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CompanyId", @" AND b.CompanyId = @CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ProcessId", @" AND b.ProcessId = @ProcessId", ProcessId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ProcessNo", @" AND b.ProcessNo LIKE '%' + @ProcessNo + '%'", ProcessNo);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ProcessName", @" AND b.ProcessName LIKE '%' + @ProcessName + '%'", ProcessName);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SupplierId", @" AND a.SupplierId = @SupplierId", SupplierId);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "Status", @" AND b.Status IN @Status", Status.Split(','));

                    sql += @"
                    ORDER BY b.ProcessId";

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

        #region //GetWorkShop -- 取得車間資料 -- 曾永至 2024.10.17
        public string GetWorkShop(int SupplierId, int ShopId, string ShopNo, string ShopName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sql = $@"SELECT DISTINCT c.ShopId, c.ShopNo, c.ShopName
                         FROM SCM.SupplierMachine a 
                         LEFT JOIN MES.Machine b ON a.MachineId = b.MachineId
                         LEFT JOIN MES.WorkShop c ON b.ShopId = c.ShopId
                         WHERE a.SupplierId = {SupplierId}";

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CompanyId", @" AND c.CompanyId = @CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ShopNo", @" AND c.ShopNo LIKE '%' + @ShopNo + '%'", ShopNo);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ShopName", @" AND c.ShopName LIKE '%' + @ShopName + '%'", ShopName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "Status", @" AND c.Status IN @Status", Status.Split(','));

                    sql += @"
                    ORDER BY c.ShopId";

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

        #region //GetMachine -- 取得機台資料 -- 曾永至 2024.10.17
        public string GetMachine(int SupplierId, int ShopId, int MachineId, string MachineNo, string MachineName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sql = $@"SELECT DISTINCT b.MachineId, b.MachineNo, b.MachineName, (b.MachineNo + '-' + b.MachineName) MachineWithNo
                         FROM SCM.SupplierMachine a 
                         LEFT JOIN MES.Machine b ON a.MachineId = b.MachineId
                         LEFT JOIN MES.WorkShop c ON b.ShopId = c.ShopId
                         WHERE a.SupplierId = {SupplierId}";

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CompanyId", @" AND c.CompanyId = @CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ShopId", @" AND c.ShopId = @ShopId", ShopId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MachineNo", @" AND b.MachineNo LIKE '%' + @MachineNo + '%'", MachineNo);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MachineName", @" AND b.MachineName LIKE '%' + @MachineName + '%'", MachineName);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SupplierId", @" AND a.SupplierId = @SupplierId", SupplierId);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "Status", @" AND b.Status IN @Status", Status.Split(','));

                    sql += @"
                    ORDER BY b.MachineId";

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

        #region //GetProcessItem -- 取得托外製程相關資料 -- Andrew 2024.10.16
        public string GetProcessItem(int ProcessId, string ProcessNo, string ProcessName, string ProcessDesc, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.ProcessId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ProcessNo, a.ProcessName, a.LastModifiedDate";
                    sqlQuery.mainTables =
                        @"FROM MES.Process a";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProcessId", @" AND a.ProcessId = @ProcessId", ProcessId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProcessNo", @" AND a.ProcessNo LIKE '%' + @ProcessNo + '%'", ProcessNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProcessName", @" AND a.ProcessName LIKE '%' + @ProcessName + '%'", ProcessName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProcessDesc", @" AND a.ProcessDesc LIKE '%' + @ProcessDesc + '%'", ProcessDesc);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status = @Status ", Status);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ProcessId DESC";
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

        #region //GetMachineItem -- 取得托外機台相關資料 -- Andrew 2024.10.16
        public string GetMachineItem(int ShopId, string ShopName, int MachineId, string MachineName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.MachineId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.MachineNo, a.MachineName, b.ShopId, b.ShopNo, b.ShopName";
                    sqlQuery.mainTables =
                        @"FROM MES.Machine a 
                        INNER JOIN MES.WorkShop b ON a.ShopId = b.ShopId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND b.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ShopId", @" AND a.ShopId = @ShopId", ShopId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ShopName", @" AND b.ShopName LIKE '%' + @ShopName + '%'", ShopName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MachineId", @" AND a.MachineId = @MachineId", MachineId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MachineName", @" AND a.MachineName LIKE '%' + @MachineName + '%'", MachineName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status = @Status ", Status);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.MachineId DESC";
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

        #region //GetProject -- 取得專案資料 -- Ann 2025-04-23
        public string GetProject(int ProjectId, string ProjectNo, string ProjectName
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.ProjectId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.CompanyId, a.ProjectNo, a.ProjectName, a.Remark
                        , ISNULL(FORMAT(a.EffectiveDate, 'yyyy-MM-dd'), '') EffectiveDate, ISNULL(FORMAT(a.ExpirationDate, 'yyyy-MM-dd'), '') ExpirationDate
                        , a.CloseCode, a.TransferErpStatus, ISNULL(FORMAT(a.TransferErpDate, 'yyyy-MM-dd'), '') TransferErpDate
                        , FORMAT(a.CreateDate, 'yyyy-MM-dd') CreateDate, FORMAT(a.LastModifiedDate, 'yyyy-MM-dd') LastModifiedDate";
                    sqlQuery.mainTables =
                        @"FROM SCM.Project a";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProjectId", @" AND a.ProjectId = @ProjectId", ProjectId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProjectNo", @" AND a.ProjectNo LIKE '%' + @ProjectNo + '%'", ProjectNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProjectName", @" AND a.ProjectName LIKE '%' + @ProjectName + '%'", ProjectName);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ProjectId DESC";
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

        #region //GetProjectDetail -- 取得專案詳細資料 -- Ann 2025-04-23
        public string GetProjectDetail(int ProjectDetailId, int ProjectId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.ProjectDetailId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ProjectId, a.ProjectType, a.Currency, a.ExchangeRate, a.BudgetAmount, a.LocalBudgetAmount
                        , a.Edition, a.BpmTransferStatus, a.[Status], a.Remark
                        , b.TypeName ProjectTypeName
                        , c.StatusName
                        , (
                            SELECT x.FileId
                            FROM SCM.ProjectFile x
                            WHERE x.ProjectDetailId = a.ProjectDetailId
                            FOR JSON PATH, ROOT('data')
                        ) ProjectFile
                        , (
                            SELECT TOP 1 x.LogId
                            FROM SCM.ProjectBudgetChangeLog x 
                            WHERE x.ProjectDetailId = a.ProjectDetailId
                            AND x.BpmStatus = 'N'
                        ) LogId";
                    sqlQuery.mainTables =
                        @"FROM SCM.ProjectDetail a 
                        INNER JOIN BAS.[Type] b ON b.TypeSchema = 'ProjectDetail.ProjectType' AND b.TypeNo = a.ProjectType
                        INNER JOIN BAS.[Status] c ON c.StatusSchema = 'ProjectDetail.Status' AND c.StatusNo = a.[Status]
                        INNER JOIN SCM.Project d ON a.ProjectId = d.ProjectId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND d.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProjectDetailId", @" AND a.ProjectDetailId = @ProjectDetailId", ProjectDetailId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProjectId", @" AND a.ProjectId = @ProjectId", ProjectId);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ProjectDetailId DESC";
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

        #region //GetProjectBudgetChangeLog -- 取得專案預算變更Log紀錄 -- Ann 2025-04-24
        public string GetProjectBudgetChangeLog(int ProjectDetailId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.LogId, a.ProjectDetailId, a.Currency, a.ExchangeRate, a.BudgetAmount, a.LocalBudgetAmount
                            , a.OriCurrency, a.OriExchangeRate, a.OriBudgetAmount, a.OriLocalBudgetAmount, a.OriEdition, a.BpmStatus, a.Remark
                            , FORMAT(a.CreateDate, 'yyyy-MM-dd') CreateDate
                            , b.UserNo, b.UserName
                            FROM SCM.ProjectBudgetChangeLog a 
                            INNER JOIN BAS.[User] b ON a.CreateBy = b.UserId
                            WHERE a.ProjectDetailId = @ProjectDetailId";
                    dynamicParameters.Add("ProjectDetailId", ProjectDetailId);

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
        #region //AddPackingVolume -- 包材體積資料新增 -- Zoey 2022.06.10
        public string AddPackingVolume(int VolumeId, string VolumeNo, string VolumeName, string VolumeType, string VolumeSpec, double Volume, int VolumeUomId)
        {
            try
            {
                if (VolumeNo.Length <= 0) throw new SystemException("【體積代號】不能為空!");
                if (VolumeNo.Length > 50) throw new SystemException("【體積代號】長度錯誤!");
                if (VolumeName.Length <= 0) throw new SystemException("【體積名稱】不能為空!");
                if (VolumeName.Length > 100) throw new SystemException("【體積名稱】長度錯誤!");
                if (VolumeType.Length <= 0) throw new SystemException("【體積類別】不能為空!");
                if (VolumeSpec.Length <= 0) throw new SystemException("【體積規格】不能為空!");
                if (VolumeSpec.Length > 100) throw new SystemException("【體積規格】長度錯誤!");
                if (Volume < 0) throw new SystemException("【撿貨體積】不能為空!");
                if (VolumeUomId <= 0) throw new SystemException("【體積單位】不能為空!");


                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷體積代號是否重複
                        sql = @"SELECT TOP 1 1
                                FROM SCM.PackingVolume
                                WHERE CompanyId = @CompanyId
                                AND VolumeNo = @VolumeNo
                                AND VolumeId != @VolumeId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("VolumeNo", VolumeNo);
                        dynamicParameters.Add("VolumeId", VolumeId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【體積代號】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.PackingVolume (CompanyId, VolumeNo, VolumeName
                                , VolumeSpec, VolumeType, Volume, VolumeUomId, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.VolumeId
                                VALUES (@CompanyId, @VolumeNo, @VolumeName
                                , @VolumeSpec, @VolumeType, @Volume, @VolumeUomId, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                VolumeNo,
                                VolumeName,
                                VolumeSpec,
                                VolumeType,
                                Volume,
                                VolumeUomId = VolumeUomId > 0 ? (int?)VolumeUomId : null,
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

        #region //AddPackingWeight -- 包材重量資料新增 -- Zoey 2022.06.10
        public string AddPackingWeight(int WeightId, string WeightNo, string WeightName, string WeightSpec, double Weight, int WeightUomId)
        {
            try
            {
                if (WeightNo.Length <= 0) throw new SystemException("【重量代號】不能為空!");
                if (WeightNo.Length > 50) throw new SystemException("【重量代號】長度錯誤!");
                if (WeightName.Length <= 0) throw new SystemException("【重量名稱】不能為空!");
                if (WeightName.Length > 100) throw new SystemException("【重量名稱】長度錯誤!");
                if (WeightSpec.Length <= 0) throw new SystemException("【重量規格】不能為空!");
                if (WeightSpec.Length > 100) throw new SystemException("【重量規格】長度錯誤!");
                if (Weight < 0) throw new SystemException("【撿貨重量】不能為空!");
                if (WeightUomId <= 0) throw new SystemException("【重量單位】不能為空!");


                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷重量代號是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.PackingWeight
                                WHERE CompanyId = @CompanyId
                                AND WeightNo = @WeightNo
                                AND WeightId != @WeightId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("WeightNo", WeightNo);
                        dynamicParameters.Add("WeightId", WeightId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【重量代號】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.PackingWeight (CompanyId, WeightNo, WeightName
                                , WeightSpec, Weight, WeightUomId, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.WeightId
                                VALUES (@CompanyId, @WeightNo, @WeightName
                                , @WeightSpec, @Weight, @WeightUomId, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                WeightNo,
                                WeightName,
                                WeightSpec,
                                Weight,
                                WeightUomId = WeightUomId > 0 ? (int?)WeightUomId : null,
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

        #region //AddItemWeight -- 物件重量資料新增 -- Zoey 2022.06.13
        public string AddItemWeight(int ItemDefaultWeightId, int MtlModelId, string DefaultWeight, int DefaultWeightUomId, string Status, int ParentId)
        {
            try
            {
                if (MtlModelId <= 0) throw new SystemException("【物件機型名稱】不能為空!");
                if (DefaultWeight.Length <= 0) throw new SystemException("【物件重量】不能為空!");
                if (DefaultWeight.Length > 100) throw new SystemException("【物件重量】長度錯誤!");
                if (DefaultWeightUomId <= 0) throw new SystemException("【重量單位】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷物件機型名稱是否重複
                        sql = @"SELECT TOP 1 1
                                FROM SCM.ItemDefaultWeight
                                WHERE MtlModelId = @MtlModelId
                                AND ItemDefaultWeightId != @ItemDefaultWeightId";
                        dynamicParameters.Add("MtlModelId", MtlModelId);
                        dynamicParameters.Add("ItemDefaultWeightId", ItemDefaultWeightId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【物件機型名稱】重複，請重新輸入!");
                        #endregion

                        #region //抓取目前最大序號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ISNULL(MAX(MtlModelSort), 0) MaxSort
                                FROM PDM.MtlModel
                                WHERE CompanyId = @CompanyId
                                AND ParentId = @ParentId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ParentId", ParentId);
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.ItemDefaultWeight (MtlModelId, DefaultWeight, DefaultWeightUomId, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ItemDefaultWeightId
                                VALUES (@MtlModelId, @DefaultWeight, @DefaultWeightUomId, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                MtlModelId = MtlModelId > 0 ? (int?)MtlModelId : null,
                                DefaultWeight,
                                DefaultWeightUomId = DefaultWeightUomId > 0 ? (int?)DefaultWeightUomId : null,
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

        #region //AddCustomer -- 客戶資料新增 -- Zoey 2022.06.16
        public string AddCustomer(int CustomerId, string CustomerNo, string CustomerName, string CustomerEnglishName, string CustomerShortName, string RelatedPerson
            , string PermitDate, string Version, string ResponsiblePerson, string Contact, string TelNoFirst, string TelNoSecond, string FaxNo, string Email
            , string GuiNumber, double Capital, double AnnualTurnover, int Headcount, string HomeOffice, string Currency, int DepartmentId, string CustomerKind
            , int SalesmenId, int PaymentSalesmenId, string InauguateDate, string CloseDate, string ZipCodeRegister, string RegisterAddressFirst, string RegisterAddressSecond
            , string ZipCodeInvoice, string InvoiceAddressFirst, string InvoiceAddressSecond, string ZipCodeDelivery, string DeliveryAddressFirst, string DeliveryAddressSecond
            , string ZipCodeDocument, string DocumentAddressFirst, string DocumentAddressSecond, string BillReceipient, string ZipCodeBill, string BillAddressFirst, string BillAddressSecond
            , string InvocieAttachedStatus, double DepositRate, string TaxAmountCalculateType, string SaleRating, string CreditRating, string TradeTerm, string PaymentTerm, string PricingType
            , string ClearanceType, string DocumentDeliver, string ReceiptReceive, string PaymentType, string TaxNo, string InvoiceCount, string Taxation, string Country
            , string Region, string Route, string UploadType, string PaymentBankFirst, string BankAccountFirst, string PaymentBankSecond, string BankAccountSecond
            , string PaymentBankThird, string BankAccountThird, string Account, string AccountInvoice, string AccountDay, string ShipMethod, string ShipType, int ForwarderId
            , string CustomerRemark, double CreditLimit, string CreditLimitControl, string CreditLimitControlCurrency, string SoCreditAuditType, string SiCreditAuditType
            , string DoCreditAuditType, string InTransitCreditAuditType, string TransferStatus, string TransferDate, string Status)
        {
            try
            {
                #region //判斷客戶資料長度
                if (CustomerNo.Length <= 0) throw new SystemException("【客戶代號】不能為空!");
                if (CustomerName.Length <= 0) throw new SystemException("【客戶名稱】不能為空!");
                if (CustomerShortName.Length <= 0) throw new SystemException("【客戶簡稱】不能為空!");
                if (ResponsiblePerson.Length <= 0) throw new SystemException("【負責人】不能為空!");
                if (TelNoFirst.Length <= 0) throw new SystemException("【電話(一)】不能為空!");
                if (GuiNumber.Length <= 0) throw new SystemException("【統一編號】不能為空!");
                if (Capital < 0) throw new SystemException("【資本額】不能為空!");
                if (AnnualTurnover < 0) throw new SystemException("【年營業額】不能為空!");
                if (Headcount < 0) throw new SystemException("【員工人數】不能為空!");
                if (Currency.Length <= 0) throw new SystemException("【交易幣別】不能為空!");
                if (DepartmentId <= 0) throw new SystemException("【部門別】不能為空!");
                if (SalesmenId <= 0) throw new SystemException("【業務人員】不能為空!");
                if (PaymentSalesmenId <= 0) throw new SystemException("【收款業務員】不能為空!");
                if (RegisterAddressFirst.Length <= 0) throw new SystemException("【登記地址(中文)】不能為空!");
                if (InvoiceAddressFirst.Length <= 0) throw new SystemException("【發票地址(中文)】不能為空!");
                if (DeliveryAddressFirst.Length <= 0) throw new SystemException("【送貨地址(中文)】不能為空!");
                if (DocumentAddressFirst.Length <= 0) throw new SystemException("【文件地址(中文)】不能為空!");
                if (BillAddressFirst.Length <= 0) throw new SystemException("【帳單地址(中文)】不能為空!");
                if (DepositRate < 0) throw new SystemException("【訂金比率】不能為空!");
                if (TradeTerm.Length <= 0) throw new SystemException("【交易條件】不能為空!");
                if (PaymentTerm.Length <= 0) throw new SystemException("【付款條件】不能為空!");
                if (PricingType.Length <= 0) throw new SystemException("【計價方式】不能為空!");
                if (ClearanceType.Length <= 0) throw new SystemException("【通關方式】不能為空!");
                if (DocumentDeliver.Length <= 0) throw new SystemException("【採購單發送方式】不能為空!");
                if (ReceiptReceive.Length <= 0) throw new SystemException("【票據寄領】不能為空!");
                if (PaymentType.Length <= 0) throw new SystemException("【收款方式】不能為空!");
                if (TaxNo.Length <= 0) throw new SystemException("【稅別碼】不能為空!");
                if (InvoiceCount.Length <= 0) throw new SystemException("【發票聯數】不能為空!");
                if (Taxation.Length <= 0) throw new SystemException("【課稅別】不能為空!");
                if (ShipMethod.Length <= 0) throw new SystemException("【運輸方式】不能為空!");
                if (ForwarderId <= 0) throw new SystemException("【貨運承攬商】不能為空!");
                if (CreditLimit < 0) throw new SystemException("【信用額度】不能為空!");
                if (CreditLimitControlCurrency.Length <= 0) throw new SystemException("【信用額度控管幣別】不能為空!");
                if (SoCreditAuditType.Length <= 0) throw new SystemException("【訂單信用查核方式】不能為空!");
                if (SiCreditAuditType.Length <= 0) throw new SystemException("【銷貨信用查核方式】不能為空!");
                if (DoCreditAuditType.Length <= 0) throw new SystemException("【出貨通知查核方式】不能為空!");
                if (InTransitCreditAuditType.Length <= 0) throw new SystemException("【暫出單信用查核方式】不能為空!");

                if (CustomerNo.Length > 50) throw new SystemException("【客戶代號】長度錯誤!");
                if (CustomerName.Length > 200) throw new SystemException("【客戶名稱】長度錯誤!");
                if (CustomerEnglishName.Length > 80) throw new SystemException("【英文名稱】長度錯誤!");
                if (CustomerShortName.Length > 100) throw new SystemException("【客戶簡稱】長度錯誤!");
                if (ResponsiblePerson.Length > 30) throw new SystemException("【負責人】長度錯誤!");
                if (Contact.Length > 30) throw new SystemException("【聯絡人】長度錯誤!");
                if (TelNoFirst.Length > 20) throw new SystemException("【電話(一)】長度錯誤!");
                if (TelNoSecond.Length > 20) throw new SystemException("【電話(二)】長度錯誤!");
                if (FaxNo.Length > 20) throw new SystemException("【客戶傳真】長度錯誤!");
                if (Email.Length > 60) throw new SystemException("【電子郵件】長度錯誤!");
                if (GuiNumber.Length > 20) throw new SystemException("【統一編號】長度錯誤!");
                if (HomeOffice.Length > 10) throw new SystemException("【總店號】長度錯誤!");
                if (ZipCodeRegister.Length > 6) throw new SystemException("【郵遞區號(登記)】長度錯誤!");
                if (ZipCodeInvoice.Length > 6) throw new SystemException("【郵遞區號(發票)】長度錯誤!");
                if (ZipCodeDelivery.Length > 6) throw new SystemException("【郵遞區號(送貨)】長度錯誤!");
                if (ZipCodeDocument.Length > 6) throw new SystemException("【郵遞區號(文件)】長度錯誤!");
                if (ZipCodeBill.Length > 6) throw new SystemException("【郵遞區號(帳單)】長度錯誤!");
                if (RegisterAddressFirst.Length > 255) throw new SystemException("【登記地址(中文)】長度錯誤!");
                if (InvoiceAddressFirst.Length > 255) throw new SystemException("【發票地址(中文)】長度錯誤!");
                if (DeliveryAddressFirst.Length > 255) throw new SystemException("【送貨地址(中文)】長度錯誤!");
                if (DocumentAddressFirst.Length > 255) throw new SystemException("【文件地址(中文)】長度錯誤!");
                if (BillAddressFirst.Length > 255) throw new SystemException("【帳單地址(中文)】長度錯誤!");
                if (RegisterAddressSecond.Length > 255) throw new SystemException("【登記地址(英文)】長度錯誤!");
                if (InvoiceAddressSecond.Length > 255) throw new SystemException("【發票地址(英文)】長度錯誤!");
                if (DeliveryAddressSecond.Length > 255) throw new SystemException("【送貨地址(英文)】長度錯誤!");
                if (DocumentAddressSecond.Length > 255) throw new SystemException("【文件地址(英文)】長度錯誤!");
                if (BillAddressSecond.Length > 255) throw new SystemException("【帳單地址(英文)】長度錯誤!");
                if (BankAccountFirst.Length > 20) throw new SystemException("【銀行帳號(一)】長度錯誤!");
                if (BankAccountSecond.Length > 20) throw new SystemException("【銀行帳號(二)】長度錯誤!");
                if (BankAccountThird.Length > 20) throw new SystemException("【銀行帳號(三)】長度錯誤!");
                if (CustomerRemark.Length > 500) throw new SystemException("【備註】長度錯誤!");
                #endregion

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷客戶代號是否重複
                        sql = @"SELECT TOP 1 1
                                FROM SCM.Customer
                                WHERE CustomerNo = @CustomerNo
                                AND CustomerId != @CustomerId";
                        dynamicParameters.Add("CustomerNo", CustomerNo);
                        dynamicParameters.Add("CustomerId", CustomerId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【客戶代號】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.Customer (CompanyId, DepartmentId, CustomerNo, CustomerName, CustomerEnglishName, CustomerShortName
                                , Country, Region, Route, GuiNumber, ResponsiblePerson, Contact, TelNoFirst, TelNoSecond, FaxNo, Email
                                , RegisterAddressFirst, RegisterAddressSecond, InvoiceAddressFirst, InvoiceAddressSecond, DeliveryAddressFirst
                                , DeliveryAddressSecond, DocumentAddressFirst, DocumentAddressSecond, BillAddressFirst, BillAddressSecond
                                , ZipCodeRegister, ZipCodeInvoice, ZipCodeDelivery, ZipCodeDocument, ZipCodeBill, BillReceipient, InauguateDate
                                , AccountDay, CloseDate, PermitDate, Version, Capital, Headcount, HomeOffice, AnnualTurnover, Currency, TradeTerm
                                , PaymentTerm, PricingType, PaymentType, ReceiptReceive, TaxAmountCalculateType, DocumentDeliver
                                , InvoiceCount, TaxNo, Taxation, PaymentBankFirst, PaymentBankSecond, PaymentBankThird, BankAccountFirst
                                , BankAccountSecond, BankAccountThird, Account, AccountInvoice, ClearanceType, ShipMethod, InvocieAttachedStatus
                                , CustomerKind, UploadType, DepositRate, SaleRating, CreditRating, CreditLimit, CreditLimitControl
                                , CreditLimitControlCurrency, SoCreditAuditType, SiCreditAuditType, DoCreditAuditType, InTransitCreditAuditType
                                , RelatedPerson, SalesmenId, PaymentSalesmenId, ShipType, ForwarderId, CustomerRemark, TransferStatus
                                , TransferDate, Status, CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.CustomerId
                                VALUES (@CompanyId, @DepartmentId, @CustomerNo, @CustomerName, @CustomerEnglishName, @CustomerShortName
                                , @Country, @Region, @Route, @GuiNumber, @ResponsiblePerson, @Contact, @TelNoFirst, @TelNoSecond, @FaxNo, @Email
                                , @RegisterAddressFirst, @RegisterAddressSecond, @InvoiceAddressFirst, @InvoiceAddressSecond, @DeliveryAddressFirst
                                , @DeliveryAddressSecond, @DocumentAddressFirst, @DocumentAddressSecond, @BillAddressFirst, @BillAddressSecond
                                , @ZipCodeRegister, @ZipCodeInvoice, @ZipCodeDelivery, @ZipCodeDocument, @ZipCodeBill, @BillReceipient, @InauguateDate
                                , @AccountDay, @CloseDate, @PermitDate, @Version, @Capital, @Headcount, @HomeOffice, @AnnualTurnover, @Currency, @TradeTerm
                                , @PaymentTerm, @PricingType, @PaymentType, @ReceiptReceive, @TaxAmountCalculateType, @DocumentDeliver
                                , @InvoiceCount, @TaxNo, @Taxation, @PaymentBankFirst, @PaymentBankSecond, @PaymentBankThird, @BankAccountFirst
                                , @BankAccountSecond, @BankAccountThird, @Account, @AccountInvoice, @ClearanceType, @ShipMethod, @InvocieAttachedStatus
                                , @CustomerKind, @UploadType, @DepositRate, @SaleRating, @CreditRating, @CreditLimit, @CreditLimitControl
                                , @CreditLimitControlCurrency, @SoCreditAuditType, @SiCreditAuditType, @DoCreditAuditType, @InTransitCreditAuditType
                                , @RelatedPerson, @SalesmenId, @PaymentSalesmenId, @ShipType, @ForwarderId, @CustomerRemark, @TransferStatus
                                , @TransferDate, @Status, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                DepartmentId = DepartmentId > 0 ? (int?)DepartmentId : null,
                                CustomerNo,
                                CustomerName,
                                CustomerEnglishName,
                                CustomerShortName,
                                Country,
                                Region,
                                Route,
                                GuiNumber,
                                ResponsiblePerson,
                                Contact,
                                TelNoFirst,
                                TelNoSecond,
                                FaxNo,
                                Email,
                                RegisterAddressFirst,
                                RegisterAddressSecond,
                                InvoiceAddressFirst,
                                InvoiceAddressSecond,
                                DeliveryAddressFirst,
                                DeliveryAddressSecond,
                                DocumentAddressFirst,
                                DocumentAddressSecond,
                                BillAddressFirst,
                                BillAddressSecond,
                                ZipCodeRegister,
                                ZipCodeInvoice,
                                ZipCodeDelivery,
                                ZipCodeDocument,
                                ZipCodeBill,
                                BillReceipient,
                                InauguateDate = InauguateDate.Length > 0 ? InauguateDate : null,
                                AccountDay,
                                CloseDate = CloseDate.Length > 0 ? CloseDate : null,
                                PermitDate = PermitDate.Length > 0 ? PermitDate : null,
                                Version = "0000",
                                Capital,
                                Headcount,
                                HomeOffice,
                                AnnualTurnover,
                                Currency,
                                TradeTerm,
                                PaymentTerm,
                                PricingType,
                                PaymentType,
                                ReceiptReceive,
                                TaxAmountCalculateType,
                                DocumentDeliver,
                                InvoiceCount,
                                TaxNo,
                                Taxation,
                                PaymentBankFirst,
                                PaymentBankSecond,
                                PaymentBankThird,
                                BankAccountFirst,
                                BankAccountSecond,
                                BankAccountThird,
                                Account,
                                AccountInvoice,
                                ClearanceType,
                                ShipMethod,
                                InvocieAttachedStatus,
                                CustomerKind,
                                UploadType,
                                DepositRate,
                                SaleRating,
                                CreditRating,
                                CreditLimit,
                                CreditLimitControl,
                                CreditLimitControlCurrency,
                                SoCreditAuditType,
                                SiCreditAuditType,
                                DoCreditAuditType,
                                InTransitCreditAuditType,
                                RelatedPerson,
                                SalesmenId = SalesmenId > 0 ? (int?)SalesmenId : null,
                                PaymentSalesmenId = PaymentSalesmenId > 0 ? (int?)PaymentSalesmenId : null,
                                ShipType,
                                ForwarderId = ForwarderId > 0 ? (int?)ForwarderId : null,
                                CustomerRemark,
                                TransferStatus = "Y",
                                TransferDate = TransferDate.Length > 0 ? TransferDate : null,
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

        #region //AddForwarder -- 貨運承攬商資料新增 -- Zoey 2022.07.01
        public string AddForwarder(int ForwarderId, string ShipMethod, string ForwarderNo, string ForwarderName, string Status)
        {
            try
            {
                if (ShipMethod.Length <= 0) throw new SystemException("【運輸方式】不能為空!");
                if (ForwarderNo.Length <= 0) throw new SystemException("【貨運承攬商代號】不能為空!");
                if (ForwarderNo.Length > 50) throw new SystemException("【貨運承攬商代號】長度錯誤!");
                if (ForwarderName.Length <= 0) throw new SystemException("【貨運承攬商名稱】不能為空!");
                if (ForwarderName.Length > 100) throw new SystemException("【貨運承攬商名稱】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷貨運承攬商代號是否重複
                        sql = @"SELECT TOP 1 1
                                FROM SCM.Forwarder
                                WHERE ForwarderNo = @ForwarderNo
                                AND ForwarderId != @ForwarderId";
                        dynamicParameters.Add("ForwarderNo", ForwarderNo);
                        dynamicParameters.Add("ForwarderId", ForwarderId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【貨運承攬商代號】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.Forwarder (CompanyId, ShipMethod, ForwarderNo, ForwarderName, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ForwarderId
                                VALUES (@CompanyId, @ShipMethod, @ForwarderNo, @ForwarderName, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                ShipMethod,
                                ForwarderNo,
                                ForwarderName,
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

        #region //AddDeliveryCustomer -- 送貨客戶資料新增 -- Zoey 2022.07.01
        public string AddDeliveryCustomer(int DcId, int CustomerId, string DcName, string DcEnglishName, string DcShortName, string Contact
            , string TelNo, string FaxNo, string RegisteredAddress, string DeliveryAddress, string ShipType, int ForwarderId, string Status)

        {
            try
            {
                if (CustomerId <= 0) throw new SystemException("【客戶來源】不能為空!");
                if (DcName.Length <= 0) throw new SystemException("【客戶名稱】不能為空!");
                if (DcName.Length > 200) throw new SystemException("【客戶名稱】長度錯誤!");
                if (DcEnglishName.Length > 80) throw new SystemException("【英文名稱】長度錯誤!");
                if (DcShortName.Length <= 0) throw new SystemException("【客戶簡稱】不能為空!");
                if (DcShortName.Length > 100) throw new SystemException("【客戶簡稱】長度錯誤!");
                if (Contact.Length > 30) throw new SystemException("【客戶聯絡人】長度錯誤!");
                if (TelNo.Length <= 0) throw new SystemException("【聯絡電話】不能為空!");
                if (TelNo.Length > 20) throw new SystemException("【聯絡電話】長度錯誤!");
                if (FaxNo.Length > 20) throw new SystemException("【聯絡傳真】長度錯誤!");
                if (RegisteredAddress.Length <= 0) throw new SystemException("【登記地址】不能為空!");
                if (RegisteredAddress.Length > 255) throw new SystemException("【登記地址】長度錯誤!");
                if (DeliveryAddress.Length <= 0) throw new SystemException("【送貨地址】不能為空!");
                if (DeliveryAddress.Length > 255) throw new SystemException("【送貨地址】長度錯誤!");
                if (ShipType.Length > 50) throw new SystemException("【運輸種類】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷送貨客戶名稱是否重複
                        //sql = @"SELECT TOP 1 1
                        //        FROM SCM.DeliveryCustomer
                        //        WHERE DcName = @DcName
                        //        AND DcId != @DcId";
                        //dynamicParameters.Add("DcName", DcName);
                        //dynamicParameters.Add("DcId", DcId);

                        //var result = sqlConnection.Query(sql, dynamicParameters);
                        //if (result.Count() > 0) throw new SystemException("【送貨客戶名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.DeliveryCustomer (CustomerId, DcName, DcEnglishName, DcShortName, Contact
                                , TelNo, FaxNo, RegisteredAddress, DeliveryAddress, ShipType, ForwarderId, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.DcId
                                VALUES (@CustomerId, @DcName, @DcEnglishName, @DcShortName, @Contact
                                , @TelNo, @FaxNo, @RegisteredAddress, @DeliveryAddress, @ShipType, @ForwarderId, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CustomerId = CustomerId > 0 ? (int?)CustomerId : null,
                                DcName,
                                DcEnglishName,
                                DcShortName,
                                Contact,
                                TelNo,
                                FaxNo,
                                RegisteredAddress,
                                DeliveryAddress,
                                ShipType,
                                ForwarderId = ForwarderId > 0 ? (int?)ForwarderId : null,
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

        #region //AddSupplier -- 供應商資料新增 -- Zoey 2022.07.04
        public string AddSupplier(int SupplierId, string SupplierNo, string SupplierName, string SupplierShortName, string SupplierEnglishName, string Country
            , string Region, string GuiNumber, string ResponsiblePerson, string ContactFirst, string ContactSecond, string ContactThird, string TelNoFirst, string TelNoSecond
            , string FaxNo, string FaxNoAccounting, string Email, string AddressFirst, string AddressSecond, string ZipCodeFirst, string ZipCodeSecond
            , string BillAddressFirst, string BillAddressSecond, string PermitStatus, string InauguateDate, string AccountMonth, string AccountDay
            , string Version, double Capital, int Headcount, string PoDeliver, string Currency, string TradeTerm, string PaymentType, string PaymentTerm
            , string ReceiptReceive, string InvoiceCount, string TaxNo, string Taxation, string PermitPartialDelivery, string TaxAmountCalculateType, string InvocieAttachedStatus
            , string CertificateFormatType, double DepositRate, string TradeItem, string RemitBank, string RemitAccount, string AccountPayable, string AccountOverhead
            , string AccountInvoice, string SupplierLevel, string DeliveryRating, string QualityRating, string RelatedPerson, int PurchaseUserId, string SupplierRemark
            , string PassStationControl, string TransferStatus, string TransferDate, string Status)
        {
            try
            {
                #region //判斷供應商資料長度
                if (SupplierNo.Length <= 0) throw new SystemException("【供應商代號】不能為空!");
                if (SupplierName.Length <= 0) throw new SystemException("【供應商名稱】不能為空!");
                if (SupplierEnglishName.Length <= 0) throw new SystemException("【英文名稱】不能為空!");
                if (SupplierShortName.Length <= 0) throw new SystemException("【供應商簡稱】不能為空!");
                if (GuiNumber.Length <= 0) throw new SystemException("【統一編號】不能為空!");
                if (PermitStatus.Length <= 0) throw new SystemException("【核准狀態碼】不能為空!");
                if (ResponsiblePerson.Length <= 0) throw new SystemException("【負責人】不能為空!");
                if (TelNoFirst.Length <= 0) throw new SystemException("【電話(一)】不能為空!");
                if (PurchaseUserId <= 0) throw new SystemException("【採購人員】不能為空!");
                if (Capital < 0) throw new SystemException("【資本額】不能為空!");
                if (Headcount < 0) throw new SystemException("【員工人數】不能為空!");
                if (AddressFirst.Length <= 0) throw new SystemException("【聯絡地址】不能為空!");
                if (CertificateFormatType.Length <= 0) throw new SystemException("【憑證列印格式】不能為空!");
                if (Currency.Length <= 0) throw new SystemException("【交易幣別】不能為空!");
                if (TradeTerm.Length <= 0) throw new SystemException("【交易條件】不能為空!");
                if (TaxAmountCalculateType.Length <= 0) throw new SystemException("【稅額計算方式】不能為空!");
                if (PoDeliver.Length <= 0) throw new SystemException("【採購單發送方式】不能為空!");
                if (DepositRate < 0) throw new SystemException("【訂金比率】不能為空!");
                if (PaymentType.Length <= 0) throw new SystemException("【付款方式】不能為空!");
                if (PaymentTerm.Length <= 0) throw new SystemException("【付款條件】不能為空!");
                if (RemitBank.Length <= 0) throw new SystemException("【匯款銀行】不能為空!");
                if (RemitAccount.Length <= 0) throw new SystemException("【銀行帳號】不能為空!");
                if (AccountDay.Length <= 0) throw new SystemException("【結帳日期(日)】不能為空!");
                if (ReceiptReceive.Length <= 0) throw new SystemException("【票據寄領】不能為空!");
                if (TaxNo.Length <= 0) throw new SystemException("【稅別碼】不能為空!");
                if (InvoiceCount.Length <= 0) throw new SystemException("【發票聯數】不能為空!");
                if (Taxation.Length <= 0) throw new SystemException("【課稅別】不能為空!");
                if (DeliveryRating.Length <= 0) throw new SystemException("【交貨評等】不能為空!");
                if (SupplierNo.Length > 50) throw new SystemException("【供應商代號】長度錯誤!");
                if (SupplierName.Length > 200) throw new SystemException("【供應商名稱】長度錯誤!");
                if (SupplierEnglishName.Length > 100) throw new SystemException("【英文名稱】長度錯誤!");
                if (SupplierShortName.Length > 100) throw new SystemException("【客戶簡稱】長度錯誤!");
                if (ResponsiblePerson.Length > 30) throw new SystemException("【負責人】長度錯誤!");
                if (TelNoFirst.Length > 20) throw new SystemException("【電話(一)】長度錯誤!");
                if (TelNoSecond.Length > 20) throw new SystemException("【電話(二)】長度錯誤!");
                if (ContactFirst.Length > 20) throw new SystemException("【聯絡人(一)】長度錯誤!");
                if (ContactSecond.Length > 20) throw new SystemException("【聯絡人(二)】長度錯誤!");
                if (ContactThird.Length > 20) throw new SystemException("【聯絡人(三)】長度錯誤!");
                if (Email.Length > 60) throw new SystemException("【電子郵件】長度錯誤!");
                if (FaxNo.Length > 20) throw new SystemException("【客戶傳真】長度錯誤!");
                if (FaxNoAccounting.Length > 20) throw new SystemException("【客戶傳真(會計)】長度錯誤!");
                if (ZipCodeFirst.Length > 6) throw new SystemException("【郵遞區號(聯絡)】長度錯誤!");
                if (ZipCodeSecond.Length > 6) throw new SystemException("【郵遞區號(帳單)】長度錯誤!");
                if (AddressFirst.Length > 255) throw new SystemException("【聯絡地址(一)】長度錯誤!");
                if (AddressSecond.Length > 255) throw new SystemException("【聯絡地址(二)】長度錯誤!");
                if (BillAddressFirst.Length > 255) throw new SystemException("【帳單地址(一)】長度錯誤!");
                if (BillAddressSecond.Length > 255) throw new SystemException("【帳單地址(二)】長度錯誤!");
                if (TradeItem.Length > 255) throw new SystemException("【交易項目】長度錯誤!");
                if (RemitAccount.Length > 30) throw new SystemException("【銀行帳號】長度錯誤!");
                if (SupplierRemark.Length > 500) throw new SystemException("【備註】長度錯誤!");
                #endregion

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷供應商代號是否重複
                        sql = @"SELECT TOP 1 1
                                FROM SCM.Supplier
                                WHERE SupplierNo = @SupplierNo
                                AND SupplierId != @SupplierId";
                        dynamicParameters.Add("SupplierNo", SupplierNo);
                        dynamicParameters.Add("SupplierId", SupplierId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【供應商代號】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.Supplier (CompanyId, SupplierNo, SupplierName, SupplierShortName, SupplierEnglishName, Country, Region
                                , GuiNumber, ResponsiblePerson, ContactFirst, ContactSecond, ContactThird, TelNoFirst, TelNoSecond, FaxNo
                                , FaxNoAccounting, Email, AddressFirst, AddressSecond, ZipCodeFirst, ZipCodeSecond, BillAddressFirst
                                , BillAddressSecond, PermitStatus, InauguateDate, AccountMonth, AccountDay, Version, Capital, Headcount, PoDeliver
                                , Currency, TradeTerm, PaymentType, PaymentTerm, ReceiptReceive, InvoiceCount, TaxNo, Taxation
                                , PermitPartialDelivery, TaxAmountCalculateType, InvocieAttachedStatus, CertificateFormatType, DepositRate
                                , TradeItem, RemitBank, RemitAccount, AccountPayable, AccountOverhead, AccountInvoice, SupplierLevel
                                , DeliveryRating, QualityRating, RelatedPerson, PurchaseUserId, SupplierRemark, PassStationControl
                                , TransferStatus, TransferDate, Status, CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.SupplierId
                                VALUES (@CompanyId, @SupplierNo, @SupplierName, @SupplierShortName, @SupplierEnglishName, @Country, @Region
                                , @GuiNumber, @ResponsiblePerson, @ContactFirst, @ContactSecond, @ContactThird, @TelNoFirst, @TelNoSecond, @FaxNo
                                , @FaxNoAccounting, @Email, @AddressFirst, @AddressSecond, @ZipCodeFirst, @ZipCodeSecond, @BillAddressFirst
                                , @BillAddressSecond, @PermitStatus, @InauguateDate, @AccountMonth, @AccountDay, @Version, @Capital, @Headcount, @PoDeliver
                                , @Currency, @TradeTerm, @PaymentType, @PaymentTerm, @ReceiptReceive, @InvoiceCount, @TaxNo, @Taxation
                                , @PermitPartialDelivery, @TaxAmountCalculateType, @InvocieAttachedStatus, @CertificateFormatType, @DepositRate
                                , @TradeItem, @RemitBank, @RemitAccount, @AccountPayable, @AccountOverhead, @AccountInvoice, @SupplierLevel
                                , @DeliveryRating, @QualityRating, @RelatedPerson, @PurchaseUserId, @SupplierRemark, @PassStationControl
                                , @TransferStatus, @TransferDate, @Status, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                SupplierNo,
                                SupplierName,
                                SupplierShortName,
                                SupplierEnglishName,
                                Country,
                                Region,
                                GuiNumber,
                                ResponsiblePerson,
                                ContactFirst,
                                ContactSecond,
                                ContactThird,
                                TelNoFirst,
                                TelNoSecond,
                                FaxNo,
                                FaxNoAccounting,
                                Email,
                                AddressFirst,
                                AddressSecond,
                                ZipCodeFirst,
                                ZipCodeSecond,
                                BillAddressFirst,
                                BillAddressSecond,
                                PermitStatus,
                                InauguateDate = InauguateDate.Length > 0 ? InauguateDate : null,
                                AccountMonth,
                                AccountDay,
                                Version = "0000",
                                Capital,
                                Headcount,
                                PoDeliver,
                                Currency,
                                TradeTerm,
                                PaymentType,
                                PaymentTerm,
                                ReceiptReceive,
                                InvoiceCount,
                                TaxNo,
                                Taxation,
                                PermitPartialDelivery,
                                TaxAmountCalculateType,
                                InvocieAttachedStatus,
                                CertificateFormatType,
                                DepositRate,
                                TradeItem,
                                RemitBank,
                                RemitAccount,
                                AccountPayable,
                                AccountOverhead,
                                AccountInvoice,
                                SupplierLevel,
                                DeliveryRating,
                                QualityRating,
                                RelatedPerson,
                                PurchaseUserId = PurchaseUserId > 0 ? (int?)PurchaseUserId : null,
                                SupplierRemark,
                                PassStationControl,
                                TransferStatus = "Y",
                                TransferDate = TransferDate.Length > 0 ? TransferDate : null,
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

        #region //AddInventory -- 庫別資料新增 -- Zoey 2022.07.18
        public string AddInventory(int InventoryId, string InventoryNo, string InventoryName, string InventoryType
            , string MrpCalculation, string ConfirmStatus, string SaveStatus, string InventoryDesc
            , string TransferStatus, string TransferDate, string Status)
        {
            try
            {
                if (InventoryNo.Length <= 0) throw new SystemException("【庫別代號】不能為空!");
                if (InventoryNo.Length > 50) throw new SystemException("【庫別代號】長度錯誤!");
                if (InventoryName.Length <= 0) throw new SystemException("【庫別名稱】不能為空!");
                if (InventoryName.Length > 100) throw new SystemException("【庫別名稱】長度錯誤!");
                if (InventoryType.Length <= 0) throw new SystemException("【庫別性質】不能為空!");
                if (InventoryDesc.Length > 100) throw new SystemException("【庫別描述】長度錯誤!");
                if (MrpCalculation.Length <= 0) throw new SystemException("【納入MRP計算】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷庫別代號是否重複
                        sql = @"SELECT TOP 1 1
                                FROM SCM.Inventory
                                WHERE InventoryNo = @InventoryNo
                                AND InventoryId != @InventoryId";
                        dynamicParameters.Add("InventoryNo", InventoryNo);
                        dynamicParameters.Add("InventoryId", InventoryId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【庫別代號】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.Inventory (CompanyId, InventoryNo, InventoryName, InventoryType, MrpCalculation
                                , ConfirmStatus, SaveStatus, InventoryDesc, TransferStatus, TransferDate, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.InventoryId
                                VALUES (@CompanyId, @InventoryNo, @InventoryName, @InventoryType, @MrpCalculation
                                , @ConfirmStatus, @SaveStatus, @InventoryDesc, @TransferStatus, @TransferDate, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                InventoryNo,
                                InventoryName,
                                InventoryType,
                                MrpCalculation,
                                ConfirmStatus = "N",
                                SaveStatus = "N",
                                InventoryDesc,
                                TransferDate,
                                TransferStatus = "Y",
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

        #region //AddPacking -- 包材資料新增 -- Zoey 2022.10.11
        public string AddPacking(int PackingId, string PackingName, string PackingType, string VolumeSpec, double Volume
             , int VolumeUomId, string WeightSpec, double Weight, int WeightUomId, string Status)
        {
            try
            {
                if (PackingName.Length <= 0) throw new SystemException("【包材名稱】不能為空!");
                if (PackingName.Length > 50) throw new SystemException("【包材名稱】長度錯誤!");
                if (PackingType.Length <= 0) throw new SystemException("【包材種類】不能為空!");
                if (VolumeSpec.Length <= 0) throw new SystemException("【體積規格】不能為空!");
                if (VolumeSpec.Length > 100) throw new SystemException("【體積規格】長度錯誤!");
                if (Volume <= 0) throw new SystemException("【包材體積】不能為空!");
                if (WeightUomId <= 0) throw new SystemException("【體積單位】不能為空!");
                //if (WeightSpec.Length <= 0) throw new SystemException("【重量規格】不能為空!");
                if (WeightSpec.Length > 100) throw new SystemException("【重量規格】長度錯誤!");
                if (Weight <= 0) throw new SystemException("【包材重量】不能為空!");
                if (WeightUomId <= 0) throw new SystemException("【重量單位】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷包材名稱是否重複
                        sql = @"SELECT TOP 1 1
                                FROM SCM.Packing
                                WHERE PackingName = @PackingName
                                AND PackingId != @PackingId";
                        dynamicParameters.Add("PackingName", PackingName);
                        dynamicParameters.Add("PackingId", PackingId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【包材名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.Packing (CompanyId, PackingName, PackingType, VolumeSpec, Volume
                                , VolumeUomId, WeightSpec, Weight, WeightUomId, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.PackingId
                                VALUES (@CompanyId, @PackingName, @PackingType, @VolumeSpec, @Volume
                                , @VolumeUomId, @WeightSpec, @Weight, @WeightUomId, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                PackingName,
                                PackingType,
                                VolumeSpec,
                                Volume,
                                VolumeUomId,
                                WeightSpec,
                                Weight,
                                WeightUomId,
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

        #region //AddRfqProductClass -- RFQ產品類型新增 -- Chia Yuan 2023.07.03
        public string AddRfqProductClass(int RfqProClassId, string RfqProductClassName)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(RfqProductClassName)) throw new SystemException("【RFQ產品類型】不能為空!");

                RfqProductClassName = RfqProductClassName.Trim();

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷RFQ產品類型資料是否重複
                        sql = @"SELECT TOP 1 1
                                FROM SCM.RfqProductClass
                                WHERE RfqProductClassName = @RfqProductClassName
                                AND RfqProClassId != @RfqProClassId";
                        dynamicParameters.Add("RfqProductClassName", RfqProductClassName);
                        dynamicParameters.Add("RfqProClassId", RfqProClassId);

                        var result = sqlConnection.Query(sql, dynamicParameters).FirstOrDefault();
                        if (result != null) throw new SystemException("【RFQ產品類型】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.RfqProductClass (RfqProductClassName,Status,CreateDate,LastModifiedDate,CreateBy,LastModifiedBy)
                                OUTPUT INSERTED.RfqProClassId
                                VALUES (@RfqProductClassName,@Status,@CreateDate,@LastModifiedDate,@CreateBy,@LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RfqProductClassName,
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
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //AddRfqProductType -- RFQ產品類別新增 --Chia Yuan 2023.07.03

        public string AddRfqProductType(int RfqProTypeId, int RfqProClassId, string RfqProductTypeName)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(RfqProductTypeName)) throw new SystemException("【RFQ產品類別】不能為空!");

                RfqProductTypeName = RfqProductTypeName.Trim();

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷RFQ產品類型資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.RfqProductClass
                                WHERE RfqProClassId = @RfqProClassId";
                        dynamicParameters.Add("RfqProClassId", RfqProClassId);

                        var result = sqlConnection.Query(sql, dynamicParameters).FirstOrDefault();
                        if (result == null) throw new SystemException("RFQ產品類型資料錯誤!");
                        #endregion

                        #region //判斷RFQ產品類型資料是否重複
                        sql = @"SELECT TOP 1 1
                                FROM SCM.RfqProductType
                                WHERE RfqProductTypeName = @RfqProductTypeName
                                AND RfqProTypeId != @RfqProTypeId";
                        dynamicParameters.Add("RfqProductTypeName", RfqProductTypeName);
                        dynamicParameters.Add("RfqProTypeId", RfqProTypeId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters).FirstOrDefault();
                        if (result2 != null) throw new SystemException("【RFQ產品類別】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.RfqProductType (RfqProClassId,RfqProductTypeName,Status,CreateDate,LastModifiedDate,CreateBy,LastModifiedBy)
                                OUTPUT INSERTED.RfqProTypeId
                                VALUES (@RfqProClassId,@RfqProductTypeName,@Status,@CreateDate,@LastModifiedDate,@CreateBy,@LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RfqProClassId,
                                RfqProductTypeName,
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

        #region //AddRfqPackageType -- RFQ包裝種類新增 --Chia Yuan 2023.07.03

        public string AddRfqPackageType(int RfqPkTypeId, int RfqProClassId, string PackagingMethod, string SustSupplyStatus)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(PackagingMethod)) throw new SystemException("【RFQ包裝方式】不能為空!");

                PackagingMethod = PackagingMethod.Trim();

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷RFQ產品類型資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.RfqProductClass
                                WHERE RfqProClassId = @RfqProClassId";
                        dynamicParameters.Add("RfqProClassId", RfqProClassId);

                        var result = sqlConnection.Query(sql, dynamicParameters).FirstOrDefault();
                        if (result == null) throw new SystemException("RFQ產品類型資料錯誤!");
                        #endregion

                        #region //判斷RFQ包裝種類資料是否重複
                        sql = @"SELECT TOP 1 1
                                FROM SCM.RfqPackageType
                                WHERE PackagingMethod = @PackagingMethod
                                AND RfqPkTypeId != @RfqPkTypeId";
                        dynamicParameters.Add("PackagingMethod", PackagingMethod);
                        dynamicParameters.Add("RfqPkTypeId", RfqPkTypeId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters).FirstOrDefault();
                        if (result2 != null) throw new SystemException("【RFQ包裝方式】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.RfqPackageType (RfqProClassId,PackagingMethod,SustSupplyStatus,Status,CreateDate,LastModifiedDate,CreateBy,LastModifiedBy)
                                OUTPUT INSERTED.RfqPkTypeId
                                VALUES (@RfqProClassId,@PackagingMethod,@SustSupplyStatus,@Status,@CreateDate,@LastModifiedDate,@CreateBy,@LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RfqProClassId,
                                PackagingMethod,
                                SustSupplyStatus,
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

        #region //AddProductUse -- RFQ產品用途新增 -- Yi 2023.07.12
        public string AddProductUse(int ProductUseId, string ProductUseNo, string ProductUseName, string TypeOne, string Status)
        {
            try
            {
                if (ProductUseName.Length <= 0) throw new SystemException("【產品用途名稱】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷產品用途名稱是否重複
                        sql = @"SELECT TOP 1 1
                                FROM SCM.ProductUse
                                WHERE ProductUseName = @ProductUseName
                                AND ProductUseId != @ProductUseId";
                        dynamicParameters.Add("ProductUseName", ProductUseName);
                        dynamicParameters.Add("ProductUseId", ProductUseId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【產品用途名稱】重複，請重新輸入!");
                        #endregion

                        #region //代號取號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT CONVERT(int, RIGHT(ISNULL(MAX(ProductUseNo), '000'), 3)) + 1 CurrentNum
                            FROM SCM.ProductUse
                            WHERE ProductUseNo LIKE @ProductUseNo";
                        dynamicParameters.Add("ProductUseNo", string.Format("{0}___", "No"));
                        int currentNum = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).CurrentNum;

                        string productUseNo = string.Format("{0}{1}", "No", string.Format("{0:000}", currentNum));
                        #endregion

                        #region //取得產品類別
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Type]
                                WHERE TypeNo = @TypeNo
                                AND TypeSchema = 'ProductUse.Type'";
                        dynamicParameters.Add("TypeNo", TypeOne);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【產品類別】資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.ProductUse (ProductUseNo, ProductUseName, TypeOne, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ProductUseId
                                VALUES (@ProductUseNo, @ProductUseName, @TypeOne, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ProductUseNo = productUseNo,
                                ProductUseName,
                                TypeOne,
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

        #region //AddPerformanceGoals -- 新增業務績效目標 -- Shintokuro 2023.11.29
        public string AddPerformanceGoals(string PgNo, string PgName, string PgDesc, string StartDate, string EndDate)
        {
            try
            {
                if (PgNo.Length <= 0) throw new SystemException("【目標代號】不能為空!");
                if (PgNo.Length > 50) throw new SystemException("【目標代號】長度錯誤!");
                if (PgName.Length <= 0) throw new SystemException("【目標名稱】不能為空!");
                if (PgName.Length > 100) throw new SystemException("【目標名稱】長度錯誤!");
                if (PgDesc.Length <= 0) throw new SystemException("【目標描述】不能為空!");
                if (PgDesc.Length > 100) throw new SystemException("【目標描述】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷績效目標代號是否重複
                        sql = @"SELECT TOP 1 1
                                FROM SCM.PerformanceGoals
                                WHERE PgNo = @PgNo";
                        dynamicParameters.Add("PgNo", PgNo);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【業務績效目標】代號重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.PerformanceGoals (PgNo, PgName, PgDesc
                                , StartDate, EndDate, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.PgId
                                VALUES (@PgNo, @PgName, @PgDesc
                                , @StartDate, @EndDate, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                PgNo,
                                PgName,
                                PgDesc,
                                StartDate,
                                EndDate,
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
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //AddPgDetail -- 新增業務績效目標單身 -- Shintokuro 2023.11.29
        public string AddPgDetail(int PgId, string IntentType, int UserId, int IntentAmount)
        {
            try
            {
                if (PgId <= 0) throw new SystemException("【業務績效目標單頭Id】不能為空!");
                if (IntentType.Length <= 0) throw new SystemException("【目標類別】不能為空!");
                if (UserId <= 0) throw new SystemException("【業務人員】不能為空!");
                if (IntentAmount <= 0) throw new SystemException("【目標金額】必須大於0!");
                string nullDate = null;
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷業務人員是否存在
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[User]
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", UserId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【業務人員】不存在，請重新確認!");
                        #endregion

                        #region //判斷業務人員是否重複
                        sql = @"SELECT TOP 1 1
                                FROM SCM.PgDetail 
                                WHERE IntentType = @IntentType
                                AND UserId = @UserId
                                AND PgId = @PgId";
                        dynamicParameters.Add("IntentType", IntentType);
                        dynamicParameters.Add("UserId", UserId);
                        dynamicParameters.Add("PgId", PgId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【業務人員】已經存在"+ IntentType + "類別中，請重新確認!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.PgDetail (PgId, UserId, IntentType, IntentAmount
                                , ConfirmStatus, ConfirmDate
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.PgDetailId
                                VALUES (@PgId, @UserId, @IntentType, @IntentAmount
                                , @ConfirmStatus, @ConfirmDate
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                PgId,
                                UserId,
                                IntentType,
                                IntentAmount,
                                ConfirmStatus ="N",
                                ConfirmDate = nullDate,
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

        #region //AddProductTypeGroup -- 新增產品群組資料 -- Shintokuro 2024.05.24
        public string AddProductTypeGroup(int RfqProClassId, string TypeOne, string ProTypeGroupName, string CoatingFlag)
        {
            try
            {
                if (RfqProClassId <= 0) throw new SystemException("【RFQ產品類型】不能為空!");
                if (ProTypeGroupName.Length > 50) throw new SystemException("【產品群組】長度不能超過50字元!");
                if (ProTypeGroupName.Length <= 0) throw new SystemException("【產品群組】不能為空!");
                if (CoatingFlag.Length <= 0) throw new SystemException("【是否鍍膜】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷RFQ產品類型是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.RfqProductClass
                                WHERE 1 = 1
                                AND RfqProClassId = @RfqProClassId";
                        dynamicParameters.Add("RfqProClassId", RfqProClassId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【RFQ產品類型】找不到，請重新輸入!");
                        #endregion

                        #region //判斷產品屬性類別是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Type]
                                WHERE TypeSchema = 'ProductTypeGroup.Attribute'
                                AND TypeNo = @TypeOne";
                        dynamicParameters.Add("TypeOne", TypeOne);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【產品屬性】資料錯誤!");
                        #endregion

                        #region //判斷是否重複
                        sql = @"SELECT TOP 1 1
                                FROM SCM.ProductTypeGroup
                                WHERE 1 = 1
                                AND RfqProClassId = @RfqProClassId
                                AND ProTypeGroupName = @ProTypeGroupName";
                        dynamicParameters.Add("RfqProClassId", RfqProClassId);
                        dynamicParameters.Add("ProTypeGroupName", ProTypeGroupName);

                       result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【資料組合】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.ProductTypeGroup (RfqProClassId, TypeOne, ProTypeGroupName, CoatingFlag, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ProTypeGroupId
                                VALUES (@RfqProClassId, @TypeOne, @ProTypeGroupName, @CoatingFlag, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RfqProClassId,
                                TypeOne,
                                ProTypeGroupName,
                                CoatingFlag,
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

        #region //AddProductType -- 新增產品類別資料 -- Shintokuro 2024.05.24
        public string AddProductType(int ProTypeGroupId, string ProTypeName)
        {
            try
            {
                if (ProTypeGroupId <= 0) throw new SystemException("【產品群組】不能為空!");
                if (ProTypeName.Length > 100) throw new SystemException("【類別名稱】長度不能超過100字元!");
                if (ProTypeName.Length <= 0) throw new SystemException("【類別名稱】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷產品群組是否存在
                        sql = @"SELECT TOP 1 1
                                FROM SCM.ProductTypeGroup
                                WHERE 1 = 1
                                AND ProTypeGroupId = @ProTypeGroupId";
                        dynamicParameters.Add("ProTypeGroupId", ProTypeGroupId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【產品群組】找不到，請重新輸入!");
                        #endregion

                        #region //判斷是否重複
                        sql = @"SELECT TOP 1 1
                                FROM SCM.ProductType
                                WHERE 1 = 1
                                AND ProTypeGroupId = @ProTypeGroupId
                                AND ProTypeName = @ProTypeName";
                        dynamicParameters.Add("ProTypeGroupId", ProTypeGroupId);
                        dynamicParameters.Add("ProTypeName", ProTypeName);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【資料組合】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.ProductType (ProTypeGroupId, ProTypeName, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ProTypeId
                                VALUES (@ProTypeGroupId, @ProTypeName, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ProTypeGroupId,
                                ProTypeName,
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

        #region //AddTemplateRfiSignFlow -- 新增市場評估單流程 -- Shintokuro 2024.05.27
        public string AddTemplateRfiSignFlow(int ProTypeGroupId, int DepartmentId, string UserList)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷撈取最大序號是否存在
                        int MaxSortNumber = 0;
                        sql = @"SELECT TOP 1 SortNumber
                                FROM RFI.TemplateRfiSignFlow
                                WHERE ProTypeGroupId = @ProTypeGroupId
                                AND DepartmentId = @DepartmentId
                                ORDER BY SortNumber DESC";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { ProTypeGroupId, DepartmentId });
                        MaxSortNumber = Convert.ToInt32(result?.SortNumber) ?? 0;
                        #endregion

                        #region //取得部門
                        sql = @"SELECT TOP 1 1 FROM BAS.Department WHERE DepartmentId = @DepartmentId";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { DepartmentId }) ?? throw new SystemException("【部門】資料錯誤!");
                        #endregion

                        if (string.IsNullOrWhiteSpace(UserList))
                        {
                            sql = @"INSERT INTO RFI.TemplateRfiSignFlow (ProTypeGroupId, DepartmentId, FlowStatus, SortNumber, Status
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.TempRfiSfId
                                    VALUES (@ProTypeGroupId, @DepartmentId, @FlowStatus, @SortNumber, @Status
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";

                            rowsAffected += sqlConnection.Execute(sql,
                                new
                                {
                                    ProTypeGroupId,
                                    DepartmentId,
                                    FlowStatus = "1",
                                    SortNumber = MaxSortNumber + 1,
                                    Status = "A",
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                        }
                        else
                        {
                            var UserArr = UserList.Split(',');
                            for (var i = 0; i < UserArr.Length; i++)
                            {
                                int SortNumber = i + 1 + MaxSortNumber;
                                int FlowUser = Convert.ToInt32(UserArr[i]);
                                #region //判斷使用者是否存在
                                sql = @"SELECT TOP 1 1
                                        FROM BAS.[User]
                                        WHERE Status = 'A'
                                        AND UserId = @FlowUser";
                                result = sqlConnection.QueryFirstOrDefault(sql, new { FlowUser }) ?? throw new SystemException("使用者不存在，請重新確認!");
                                #endregion

                                #region //判斷關卡人員是否存在 (停用)
                                //dynamicParameters = new DynamicParameters();
                                //sql = @"SELECT TOP 1 1
                                //        FROM RFI.TemplateRfiSignFlow
                                //        WHERE 1 = 1
                                //        AND FlowUser = @FlowUser";
                                //dynamicParameters.Add("FlowUser", FlowUser);
                                //result = sqlConnection.Query(sql, dynamicParameters);
                                //if (result.Count() > 0) throw new SystemException("關卡人員已存在，請重新確認!");
                                #endregion

                                sql = @"INSERT INTO RFI.TemplateRfiSignFlow (ProTypeGroupId, DepartmentId, FlowUser, FlowStatus, SortNumber, Status
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.TempRfiSfId
                                        VALUES (@ProTypeGroupId, @DepartmentId, @FlowUser, @FlowStatus, @SortNumber, @Status
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";

                                rowsAffected += sqlConnection.Execute(sql,
                                    new
                                    {
                                        ProTypeGroupId,
                                        DepartmentId,
                                        FlowUser,
                                        FlowStatus = "1",
                                        SortNumber,
                                        Status = "A",
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
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

        #region //CopyTemplateRfiSignFlow -- 複製市場評估單流程 -- Chia Yuan 2025.07.03
        public string CopyTemplateRfiSignFlow(int ProTypeGroupId, int BaseDepartmentId, int DepartmentId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //取得群組
                        sql = @"SELECT TOP 1 1
                                FROM SCM.ProductTypeGroup
                                WHERE ProTypeGroupId = @ProTypeGroupId";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { ProTypeGroupId }) ?? throw new SystemException("【產品群組】資料錯誤!");
                        #endregion

                        #region //取得部門
                        sql = @"SELECT TOP 1 1 
                                FROM BAS.Department WHERE DepartmentId = @DepartmentId";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { DepartmentId }) ?? throw new SystemException("【部門】資料錯誤!");
                        result = sqlConnection.QueryFirstOrDefault(sql, new { DepartmentId = BaseDepartmentId }) ?? throw new SystemException("【來源部門】資料錯誤!");
                        #endregion

                        #region //新增流程
                        sql = @"INSERT RFI.TemplateRfiSignFlow
                                SELECT a.ProTypeGroupId, @DepartmentId, a.FlowUser, a.FlowJobName, a.FlowStatus, a.SortNumber
                                , @Status, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy
                                FROM RFI.TemplateRfiSignFlow a
                                WHERE a.ProTypeGroupId = @ProTypeGroupId
                                AND a.DepartmentId = @BaseDepartmentId
                                ORDER BY a.SortNumber";
                        rowsAffected += sqlConnection.Execute(sql, new
                            {
                                ProTypeGroupId,
                                DepartmentId,
                                BaseDepartmentId,
                                Status = "A",
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
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

        #region //AddTemplateDesignSignFlow -- 新增設計申請單流程 -- Shintokuro 2024.05.27
        public string AddTemplateDesignSignFlow(int ProTypeGroupId, int DepartmentId, string UserList)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷撈取最大序號是否存在
                        int MaxSortNumber = 0;
                        sql = @"SELECT TOP 1 SortNumber
                                FROM RFI.TemplateDesignSignFlow
                                WHERE ProTypeGroupId = @ProTypeGroupId
                                AND DepartmentId = @DepartmentId
                                ORDER BY SortNumber DESC";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { ProTypeGroupId, DepartmentId });
                        MaxSortNumber = Convert.ToInt32(result?.SortNumber) ?? 0;
                        #endregion

                        #region //取得部門
                        sql = @"SELECT TOP 1 1 FROM BAS.Department WHERE DepartmentId = @DepartmentId";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { DepartmentId }) ?? throw new SystemException("【部門】資料錯誤!");
                        #endregion

                        if (string.IsNullOrWhiteSpace(UserList))
                        {
                            sql = @"INSERT INTO RFI.TemplateDesignSignFlow (ProTypeGroupId, DepartmentId, FlowStatus, SortNumber, Status
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.TempDesignSfId
                                    VALUES (@ProTypeGroupId, @DepartmentId, @FlowStatus, @SortNumber, @Status
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            rowsAffected += sqlConnection.Execute(sql,
                                new
                                {
                                    ProTypeGroupId,
                                    DepartmentId,
                                    FlowStatus = "1",
                                    SortNumber = MaxSortNumber + 1,
                                    Status = "A",
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                        }
                        else
                        {
                            var UserArr = UserList.Split(',');
                            for (var i = 0; i < UserArr.Length; i++)
                            {
                                int SortNumber = i + 1 + MaxSortNumber;
                                int FlowUser = Convert.ToInt32(UserArr[i]);
                                #region //判斷使用者是否存在
                                sql = @"SELECT TOP 1 1
                                        FROM BAS.[User]
                                        WHERE Status = 'A'
                                        AND UserId = @FlowUser";
                                result = sqlConnection.QueryFirstOrDefault(sql, new { FlowUser }) ?? throw new SystemException("使用者不存在，請重新確認!");
                                #endregion

                                #region //判斷關卡人員是否存在 (停用)
                                //dynamicParameters = new DynamicParameters();
                                //sql = @"SELECT TOP 1 1
                                //        FROM RFI.TemplateDesignSignFlow
                                //        WHERE FlowUser = @FlowUser";
                                //dynamicParameters.Add("FlowUser", FlowUser);
                                //result = sqlConnection.Query(sql, dynamicParameters);
                                //if (result.Count() > 0) throw new SystemException("關卡人員已存在，請重新確認!");
                                #endregion

                                sql = @"INSERT INTO RFI.TemplateDesignSignFlow (ProTypeGroupId, DepartmentId, FlowUser, FlowStatus, SortNumber, Status
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.TempDesignSfId
                                        VALUES (@ProTypeGroupId, @DepartmentId, @FlowUser, @FlowStatus, @SortNumber, @Status
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                rowsAffected += sqlConnection.Execute(sql,
                                    new
                                    {
                                        ProTypeGroupId,
                                        DepartmentId,
                                        FlowUser,
                                        FlowStatus = "1",
                                        SortNumber,
                                        Status = "A",
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
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

        #region //CopyTemplateDesignSignFlow -- 複製新設計申請單流程 -- Chia Yuan 2025.07.03
        public string CopyTemplateDesignSignFlow(int ProTypeGroupId, int BaseDepartmentId, int DepartmentId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //取得群組
                        sql = @"SELECT TOP 1 1
                                FROM SCM.ProductTypeGroup
                                WHERE ProTypeGroupId = @ProTypeGroupId";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { ProTypeGroupId }) ?? throw new SystemException("【產品群組】資料錯誤!");
                        #endregion

                        #region //取得部門
                        sql = @"SELECT TOP 1 1 
                                FROM BAS.Department WHERE DepartmentId = @DepartmentId";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { DepartmentId }) ?? throw new SystemException("【部門】資料錯誤!");
                        result = sqlConnection.QueryFirstOrDefault(sql, new { DepartmentId = BaseDepartmentId }) ?? throw new SystemException("【來源部門】資料錯誤!");
                        #endregion

                        #region //新增流程
                        sql = @"INSERT RFI.TemplateDesignSignFlow
                                SELECT a.ProTypeGroupId, @DepartmentId, a.FlowUser, a.FlowJobName, a.FlowStatus, a.SortNumber
                                , @Status, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy
                                FROM RFI.TemplateDesignSignFlow a
                                WHERE a.ProTypeGroupId = @ProTypeGroupId
                                AND a.DepartmentId = @BaseDepartmentId
                                ORDER BY a.SortNumber";
                        rowsAffected += sqlConnection.Execute(sql, new
                        {
                            ProTypeGroupId,
                            DepartmentId,
                            BaseDepartmentId,
                            Status = "A",
                            CreateDate,
                            LastModifiedDate,
                            CreateBy,
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

        #region //AddQuotationTag -- 新增報價單屬性標籤 -- Shinotokuro 2024.08.23
        public string AddQuotationTag(string TagNo, string TagName)
        {
            try
            {
                if (TagNo.Length <= 0) throw new SystemException("【標籤代號】不能為空!");
                if (TagNo.Length > 20) throw new SystemException("【標籤代號】長度錯誤!");
                if (TagName.Length <= 0) throw new SystemException("【標籤名稱】不能為空!");
                if (TagName.Length > 20) throw new SystemException("【標籤名稱】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷標籤代號是否重複
                        sql = @"SELECT TOP 1 1
                                FROM SCM.QuotationTag
                                WHERE TagNo = @TagNo";
                        dynamicParameters.Add("TagNo", TagNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【標籤代號】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.QuotationTag (CompanyId, TagNo, TagName, Status,
                                CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.QtId
                                VALUES (@CompanyId, @TagNo, @TagName, @Status,
                                @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                TagNo,
                                TagName,
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

        #region //AddSupplierProcessEquipment -- 參與托外掃碼的供應商的製程與機台清單新增 -- Andrew 2024-10-16
        public string AddSupplierProcessEquipment(int SupplierId, int ProcessId, int MachineId, string Status)
        {
            try
            {
                // if (!Regex.IsMatch(Status, "^(A|S)$", RegexOptions.IgnoreCase)) throw new SystemException("【狀態】錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        #region //判斷製程資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Process
                                WHERE ProcessId = @ProcessId 
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("ProcessId", ProcessId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("製程資料錯誤!");
                        #endregion

                        #region //判斷機台和車間是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Machine a
                                INNER JOIN MES.WorkShop b ON a.ShopId = b.ShopId
                                WHERE MachineId = @MachineId
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("MachineId", MachineId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("車間和機台資料錯誤!");
                        #endregion

                        #region //判斷 製程+機台資料 是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.SupplierMachine
                                WHERE SupplierId = @SupplierId
                                AND ProcessId = @ProcessId
                                AND MachineId = @MachineId";
                        dynamicParameters.Add("SupplierId", SupplierId);
                        dynamicParameters.Add("ProcessId", ProcessId);
                        dynamicParameters.Add("MachineId", MachineId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【製程 + 機台】資料重複，請重新輸入!");
                        #endregion

                        #region //INSERT SCM.SupplierMachine
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.SupplierMachine (SupplierId, ProcessId, MachineId, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.SmId
                                VALUES (@SupplierId, @ProcessId, @MachineId, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SupplierId,
                                ProcessId,
                                MachineId,
                                Status,
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

        #region //AddProject -- 新增專案 -- Ann 2025-04-23
        public string AddProject(string ProjectNo, string ProjectName, string EffectiveDate, string ExpirationDate, string Remark)
        {
            try
            {
                int rowsAffected = 0;
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

                        string ErpNo = "";
                        foreach (var item in companyResult)
                        {
                            ErpNo = item.ErpNo;
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        }
                        #endregion

                        #region //確認使用者資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT UserNo
                                FROM BAS.[User] 
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", CurrentUser);

                        var UserResult = sqlConnection.Query(sql, dynamicParameters);

                        if (UserResult.Count() <= 0) throw new SystemException("【使用者】資料錯誤!!");

                        string UserNo = UserResult.FirstOrDefault().UserNo;
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            if (ProjectNo.Length < 0) throw new SystemException("【專案代號】不能為空!");
                            if (ProjectName.Length < 0) throw new SystemException("【專案代號】不能為空!");

                            #region //新增MES段
                            #region //確認是否有相同公司別及專案代號
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM SCM.Project a 
                                    WHERE a.CompanyId = @CompanyId
                                    AND a.ProjectNo = @ProjectNo";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("ProjectNo", ProjectNo);

                            var CheckProjectResult = sqlConnection.Query(sql, dynamicParameters);

                            if (CheckProjectResult.Count() > 0) throw new SystemException("已有重複的專案代號!!");
                            #endregion

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO SCM.Project (CompanyId, ProjectNo, ProjectName, Remark, EffectiveDate, ExpirationDate, CloseCode, TransferErpStatus, TransferErpDate
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.ProjectId
                                    VALUES (@CompanyId, @ProjectNo, @ProjectName, @Remark, @EffectiveDate, @ExpirationDate, @CloseCode, @TransferErpStatus, @TransferErpDate
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    CompanyId = CurrentCompany,
                                    ProjectNo,
                                    ProjectName,
                                    Remark,
                                    EffectiveDate,
                                    ExpirationDate,
                                    CloseCode = "N",
                                    TransferErpStatus = "Y",
                                    TransferErpDate = DateTime.Now,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();

                            int ProjectId = -1;
                            foreach (var item in insertResult)
                            {
                                ProjectId = item.ProjectId;
                            }
                            #endregion

                            #region //新增ERP段
                            #region //確認是否有相同公司別及專案代號
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM CMSNB a 
                                    WHERE a.NB001 = @ProjectNo";
                            dynamicParameters.Add("ProjectNo", ProjectNo);

                            var CheckErpProjectResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (CheckErpProjectResult.Count() > 0) throw new SystemException("ERP已有重複的專案代號!!");
                            #endregion

                            #region //取得USR_GROUP
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(a.MF004)) USR_GROUP
                                    FROM ADMMF a
                                    WHERE MF001 = @MF001";
                            dynamicParameters.Add("MF001", UserNo);

                            var ADMMFResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (ADMMFResult.Count() <= 0) throw new SystemException("查無ERP帳號權限，無法進行拋轉及核單!");

                            string USR_GROUP = "";
                            foreach (var item in ADMMFResult)
                            {
                                USR_GROUP = item.USR_GROUP;
                            }
                            #endregion

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO CMSNB (COMPANY, CREATOR, USR_GROUP, CREATE_DATE, FLAG
                                    , CREATE_TIME, CREATE_AP, CREATE_PRID
                                    , NB001, NB002, NB003, NB004, NB005, NB006)
                                    OUTPUT INSERTED.NB001
                                    VALUES (@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @FLAG
                                    , @CREATE_TIME, @CREATE_AP, @CREATE_PRID
                                    , @NB001, @NB002, @NB003, @NB004, @NB005, @NB006)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    COMPANY = ErpNo,
                                    CREATOR = UserNo,
                                    USR_GROUP,
                                    CREATE_DATE = DateTime.Now.ToString("yyyyMMdd"),
                                    FLAG = 1,
                                    CREATE_TIME = DateTime.Now.ToString("HH:mm:ss"),
                                    CREATE_AP = UserNo + "PC",
                                    CREATE_PRID = "BM",
                                    NB001 = ProjectNo,
                                    NB002 = ProjectName,
                                    NB003 = Remark,
                                    NB004 = EffectiveDate.Length > 0 ? DateTime.Parse(EffectiveDate).ToString("yyyyMMdd") : "",
                                    NB005 = ExpirationDate.Length > 0 ? DateTime.Parse(ExpirationDate).ToString("yyyyMMdd") : "",
                                    NB006 = "N",
                                });

                            var erpInsertResult = sqlConnection2.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();
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

        #region //AddProjectDetail -- 新增專案預算資料 -- Ann 2025-04-23
        public string AddProjectDetail(int ProjectId, string ProjectType, string Currency, double ExchangeRate, double BudgetAmount, double LocalBudgetAmount, string Remark, string ProjectFile)
        {
            try
            {
                int rowsAffected = 0;
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

                        string ErpNo = "";
                        foreach (var item in companyResult)
                        {
                            ErpNo = item.ErpNo;
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        }
                        #endregion

                        #region //確認使用者資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT UserNo
                                FROM BAS.[User] 
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", CurrentUser);

                        var UserResult = sqlConnection.Query(sql, dynamicParameters);

                        if (UserResult.Count() <= 0) throw new SystemException("【使用者】資料錯誤!!");

                        string UserNo = UserResult.FirstOrDefault().UserNo;
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            if (ProjectType.Length <= 0) throw new SystemException("【專案代號】不能為空!");
                            if (BudgetAmount <= 0) throw new SystemException("【預算金額】不能為空!");

                            #region //確認MES專案資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT ProjectNo, EffectiveDate, ExpirationDate, CloseCode, TransferErpStatus
                                    FROM SCM.Project
                                    WHERE ProjectId = @ProjectId";
                            dynamicParameters.Add("ProjectId", ProjectId);

                            var ProjectResult = sqlConnection.Query(sql, dynamicParameters);

                            if (ProjectResult.Count() <= 0) throw new SystemException("【專案代號】不能為空!");

                            string ProjectNo = "";
                            foreach (var item in ProjectResult)
                            {
                                ProjectNo = item.ProjectNo;
                                DateTime? EffectiveDate = item.EffectiveDate;
                                DateTime? ExpirationDate = item.ExpirationDate;

                                if (item.TransferErpStatus != "Y")
                                {
                                    throw new SystemException("此專案尚未拋轉ERP!!");
                                }

                                if (item.CloseCode == "Y")
                                {
                                    throw new SystemException("此專案已結案!!");
                                }

                                if (item.EffectiveDate != new DateTime(1900, 1, 1) && DateTime.Now < item.EffectiveDate)
                                {
                                    throw new SystemException("此專案尚未生效，無法使用!!");
                                }

                                if (item.ExpirationDate != new DateTime(1900, 1, 1) && DateTime.Now > item.ExpirationDate)
                                {
                                    throw new SystemException("此專案已失效，無法使用!!");
                                }
                            }
                            #endregion

                            #region //確認ERP專案資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(NB004)) EffectiveDate, LTRIM(RTRIM(NB005)) ExpirationDate
                                    , NB006 CloseCode
                                    FROM CMSNB
                                    WHERE NB001 = @ProjectNo";
                            dynamicParameters.Add("ProjectNo", ProjectNo);

                            var ErpProjectResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (ErpProjectResult.Count() <= 0) throw new SystemException("ERP【專案代號】不能為空!");

                            foreach (var item in ErpProjectResult)
                            {
                                if (item.CloseCode == "Y")
                                {
                                    throw new SystemException("此專案已結案!!");
                                }

                                if (item.EffectiveDate.Length > 0 && item.EffectiveDate != "19000101")
                                {
                                    string EffectiveDate = item.EffectiveDate;
                                    string effYear = EffectiveDate.Substring(0, 4);
                                    string effMonth = EffectiveDate.Substring(4, 2);
                                    string effDay = EffectiveDate.Substring(6, 2);
                                    DateTime effFullDate = Convert.ToDateTime(effYear + "-" + effMonth + "-" + effDay);
                                    int effresult = DateTime.Compare(DateTime.Now, effFullDate);
                                    if (effresult < 0) throw new SystemException("此專案ERP尚未生效，無法使用!!");
                                }

                                if (item.ExpirationDate.Length > 0 && item.ExpirationDate != "19000101")
                                {
                                    string ExpirationDate = item.ExpirationDate;
                                    string effYear = ExpirationDate.Substring(0, 4);
                                    string effMonth = ExpirationDate.Substring(4, 2);
                                    string effDay = ExpirationDate.Substring(6, 2);
                                    DateTime effFullDate = Convert.ToDateTime(effYear + "-" + effMonth + "-" + effDay);
                                    int effresult = DateTime.Compare(CreateDate, effFullDate);
                                    if (effresult > 0) throw new SystemException("此專案ERP已失效，無法使用!!");
                                }
                            }
                            #endregion

                            #region //確認是否已經有相同專案類型資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM SCM.ProjectDetail a 
                                    WHERE a.ProjectId = @ProjectId 
                                    AND a.ProjectType = @ProjectType";
                            dynamicParameters.Add("ProjectId", ProjectId);
                            dynamicParameters.Add("ProjectType", ProjectType);

                            var ProjectDetailResult = sqlConnection.Query(sql, dynamicParameters);

                            if (ProjectDetailResult.Count() > 0) throw new SystemException("此專案已建立相同專案類型預算!!");
                            #endregion

                            #region //確認幣別資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM CMSMF
                                    WHERE MF001 = @Currency";
                            dynamicParameters.Add("Currency", Currency);

                            var CMSMFResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (CMSMFResult.Count() <= 0) throw new SystemException("【幣別】資料錯誤!!");
                            #endregion

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO SCM.ProjectDetail (ProjectId, ProjectType, Currency, ExchangeRate, BudgetAmount, LocalBudgetAmount, Edition, BpmTransferStatus, Status, Remark
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.ProjectDetailId
                                    VALUES (@ProjectId, @ProjectType, @Currency, @ExchangeRate, @BudgetAmount, @LocalBudgetAmount, @Edition, @BpmTransferStatus, @Status, @Remark
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    ProjectId,
                                    ProjectType,
                                    Currency,
                                    ExchangeRate,
                                    BudgetAmount,
                                    LocalBudgetAmount,
                                    Edition = "0000",
                                    BpmTransferStatus = "N",
                                    Status = "N",
                                    Remark,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();

                            int ProjectDetailId = -1;
                            foreach (var item in insertResult)
                            {
                                ProjectDetailId = item.ProjectDetailId;
                            }

                            #region //附檔
                            if (ProjectFile.Length > 0)
                            {
                                string[] projectFiles = ProjectFile.Split(',');
                                foreach (var file in projectFiles)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO SCM.ProjectFile (ProjectDetailId, FileId
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.ProjectFileId
                                            VALUES (@ProjectDetailId, @FileId
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            ProjectDetailId,
                                            FileId = file,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    var fileInsertResult = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += fileInsertResult.Count();
                                }
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

        #region //AddProjectLog -- 新增專案預算BPM LOG -- Ann 2025-04-25
        public string AddProjectLog(int ProjectDetailId, string BpmNo, string Status, string RootId, string ConfirmUser, string LogType)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (LogType == "C")
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.ProjectDetailId
                                    FROM SCM.ProjectBudgetChangeLog a
                                    WHERE LogId = @LogId";
                            dynamicParameters.Add("LogId", ProjectDetailId);

                            var ProjectBudgetChangeLogResult = sqlConnection.Query(sql, dynamicParameters);

                            if (ProjectBudgetChangeLogResult.Count() <= 0) throw new SystemException("專案預算變更資料錯誤!!");

                            ProjectDetailId = ProjectBudgetChangeLogResult.FirstOrDefault().ProjectDetailId;
                        }

                        #region //確認專案預算資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.LastModifiedDate
                                FROM SCM.ProjectDetail a
                                WHERE ProjectDetailId = @ProjectDetailId";
                        dynamicParameters.Add("ProjectDetailId", ProjectDetailId);

                        var ProjectDetailResult = sqlConnection.Query(sql, dynamicParameters);

                        if (ProjectDetailResult.Count() <= 0) throw new SystemException("專案預算資料錯誤!!");

                        DateTime BpmTransferDate = new DateTime();
                        foreach (var item in ProjectDetailResult)
                        {
                            BpmTransferDate = item.LastModifiedDate;
                        }
                        #endregion

                        #region //INSERT SCM.ProjectLog
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.ProjectLog (ProjectDetailId, LogType, RootId, BpmNo, TransferBpmDate, BpmStatus, ConfirmUser, ErrorMessage
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ProjectLogId
                                VALUES (@ProjectDetailId, @LogType, @RootId, @BpmNo, @TransferBpmDate, @BpmStatus, @ConfirmUser, @ErrorMessage
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                          new
                          {
                              ProjectDetailId,
                              LogType,
                              RootId,
                              BpmNo,
                              TransferBpmDate = BpmTransferDate,
                              BpmStatus = Status,
                              ConfirmUser,
                              ErrorMessage = "",
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

        #region //AddProjectBudgetChangeLog -- 新增專案預算變更資料 -- Ann 2025-04-25
        public string AddProjectBudgetChangeLog(int ProjectDetailId, string Currency, double ExchangeRate, double BudgetAmount, double LocalBudgetAmount, string Remark, string ProjectFile)
        {
            try
            {
                int rowsAffected = 0;
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

                        string ErpNo = "";
                        foreach (var item in companyResult)
                        {
                            ErpNo = item.ErpNo;
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        }
                        #endregion

                        #region //確認使用者資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT UserNo
                                FROM BAS.[User] 
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", CurrentUser);

                        var UserResult = sqlConnection.Query(sql, dynamicParameters);

                        if (UserResult.Count() <= 0) throw new SystemException("【使用者】資料錯誤!!");

                        string UserNo = UserResult.FirstOrDefault().UserNo;
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            if (BudgetAmount <= 0) throw new SystemException("【預算金額】不能為空!");

                            #region //確認原專案預算資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.Currency, a.ExchangeRate, a.BudgetAmount, a.LocalBudgetAmount, a.Edition, a.[Status]
                                    , b.CloseCode
                                    FROM SCM.ProjectDetail a 
                                    INNER JOIN SCM.Project b ON a.ProjectId = b.ProjectId
                                    WHERE a.ProjectDetailId = @ProjectDetailId";
                            dynamicParameters.Add("ProjectDetailId", ProjectDetailId);

                            var ProjectDetailResult = sqlConnection.Query(sql, dynamicParameters);

                            if (ProjectDetailResult.Count() <= 0) throw new SystemException("【專案預算】資料錯誤!!");

                            string OriCurrency = "";
                            double OriExchangeRate = -1;
                            double OriBudgetAmount = -1;
                            double OriLocalBudgetAmount = -1;
                            string OriEdition = "";
                            foreach (var item in ProjectDetailResult)
                            {
                                if (item.Status != "Y" && item.Status != "K") throw new SystemException("目前專案預算狀態非簽核完成，無法變更!!");
                                if (item.CloseCode != "N") throw new SystemException("目前專案已結案，無法變更!!");

                                OriCurrency = item.Currency;
                                OriExchangeRate = item.ExchangeRate;
                                OriBudgetAmount = item.BudgetAmount;
                                OriLocalBudgetAmount = item.LocalBudgetAmount;
                                OriEdition = item.Edition;
                            }
                            #endregion

                            #region //確認幣別資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM CMSMF
                                    WHERE MF001 = @Currency";
                            dynamicParameters.Add("Currency", Currency);

                            var CMSMFResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (CMSMFResult.Count() <= 0) throw new SystemException("【幣別】資料錯誤!!");
                            #endregion

                            #region //Insert SCM.ProjectBudgetChangeLog
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO SCM.ProjectBudgetChangeLog (ProjectDetailId, Currency, ExchangeRate, BudgetAmount, LocalBudgetAmount
                                    , OriCurrency, OriExchangeRate, OriBudgetAmount, OriLocalBudgetAmount, OriEdition, BpmStatus, Remark
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.LogId
                                    VALUES (@ProjectDetailId, @Currency, @ExchangeRate, @BudgetAmount, @LocalBudgetAmount
                                    , @OriCurrency, @OriExchangeRate, @OriBudgetAmount, @OriLocalBudgetAmount, @OriEdition, @BpmStatus, @Remark
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    ProjectDetailId,
                                    Currency,
                                    ExchangeRate,
                                    BudgetAmount,
                                    LocalBudgetAmount,
                                    OriCurrency,
                                    OriExchangeRate,
                                    OriBudgetAmount,
                                    OriLocalBudgetAmount,
                                    OriEdition,
                                    BpmStatus = "N",
                                    Remark,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();

                            int LogId = insertResult.FirstOrDefault().LogId;
                            #endregion

                            #region //附檔
                            if (ProjectFile.Length > 0)
                            {
                                string[] projectFiles = ProjectFile.Split(',');
                                foreach (var file in projectFiles)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO SCM.ProjectChangeFile (LogId, FileId
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.ChangeFileId
                                            VALUES (@LogId, @FileId
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            LogId,
                                            FileId = file,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    var fileInsertResult = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += fileInsertResult.Count();
                                }
                            }
                            #endregion

                            #region //更改原專案預算狀態
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.ProjectDetail SET
                                    Status = @Status,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE ProjectDetailId = @ProjectDetailId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    Status = "C",
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    ProjectDetailId
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
        #endregion

        #region //Update
        #region //UpdateCustomerSynchronize -- 客戶資料同步 -- Zoey 2022.06.13
        public string UpdateCustomerSynchronize(string CompanyNo, string UpdateDate)
        {
            try
            {
                List<Customer> customers = new List<Customer>();

                if (CompanyNo.Length <= 0) throw new SystemException("【公司】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
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
                        #region //撈取ERP客戶資料
                        sql = @"SELECT LTRIM(RTRIM(MA015)) DepartmentNo, LTRIM(RTRIM(MA001)) CustomerNo, LTRIM(RTRIM(MA003)) CustomerName
                                , LTRIM(RTRIM(MA110)) CustomerEnglishName, LTRIM(RTRIM(MA002)) CustomerShortName, LTRIM(RTRIM(MA019)) Country
                                , LTRIM(RTRIM(MA018)) Region, LTRIM(RTRIM(MA077)) Route, LTRIM(RTRIM(MA010)) GuiNumber
                                , LTRIM(RTRIM(MA004)) ResponsiblePerson, LTRIM(RTRIM(MA005)) Contact, LTRIM(RTRIM(MA006)) TelNoFirst
                                , LTRIM(RTRIM(MA007)) TelNoSecond, LTRIM(RTRIM(MA008)) FaxNo, LTRIM(RTRIM(MA009)) Email
                                , LTRIM(RTRIM(MA023)) RegisterAddressFirst, LTRIM(RTRIM(MA024)) RegisterAddressSecond, LTRIM(RTRIM(MA025)) InvoiceAddressFirst
                                , LTRIM(RTRIM(MA026)) InvoiceAddressSecond, LTRIM(RTRIM(MA027)) DeliveryAddressFirst, LTRIM(RTRIM(MA064)) DeliveryAddressSecond
                                , LTRIM(RTRIM(MA062)) DocumentAddressFirst, LTRIM(RTRIM(MA063)) DocumentAddressSecond, LTRIM(RTRIM(MA099)) BillAddressFirst
                                , LTRIM(RTRIM(MA100)) BillAddressSecond,LTRIM(RTRIM(MA040)) ZipCodeRegister, LTRIM(RTRIM(MA079)) ZipCodeInvoice
                                , LTRIM(RTRIM(MA080)) ZipCodeDelivery, LTRIM(RTRIM(MA081)) ZipCodeDocument, LTRIM(RTRIM(MA098)) ZipCodeBill
                                , CASE WHEN LEN(LTRIM(RTRIM(MA020))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(MA020)) as date), 'yyyy-MM-dd') ELSE NULL END InauguateDate
                                , CASE WHEN LEN(LTRIM(RTRIM(MA068))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(MA068)) as date), 'yyyy-MM-dd') ELSE NULL END CloseDate
                                , CASE WHEN LEN(LTRIM(RTRIM(MA111))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(MA111)) as date), 'yyyy-MM-dd') ELSE NULL END PermitDate
                                , LTRIM(RTRIM(MA101)) BillReceipient, LTRIM(RTRIM(MA043)) AccountDay, LTRIM(RTRIM(MA112)) Version
                                , LTRIM(RTRIM(MA011)) Capital, LTRIM(RTRIM(MA013)) Headcount, LTRIM(RTRIM(MA065)) HomeOffice
                                , LTRIM(RTRIM(MA012)) AnnualTurnover, LTRIM(RTRIM(MA014)) Currency, LTRIM(RTRIM(MA109)) TradeTerm
                                , LTRIM(RTRIM(MA083)) PaymentTerm, LTRIM(RTRIM(MA035)) PricingType, LTRIM(RTRIM(MA041)) PaymentType
                                , LTRIM(RTRIM(MA042)) ReceiptReceive, LTRIM(RTRIM(MA087)) TaxAmountCalculateType, LTRIM(RTRIM(MA039)) DocumentDeliver
                                , LTRIM(RTRIM(MA037)) InvoiceCount, LTRIM(RTRIM(MA118)) TaxNo, LTRIM(RTRIM(MA038)) Taxation
                                , LTRIM(RTRIM(MA046)) PaymentBankFirst, LTRIM(RTRIM(MA069)) PaymentBankSecond, LTRIM(RTRIM(MA070)) PaymentBankThird
                                , LTRIM(RTRIM(MA071)) BankAccountFirst, LTRIM(RTRIM(MA072)) BankAccountSecond, LTRIM(RTRIM(MA073)) BankAccountThird
                                , LTRIM(RTRIM(MA047)) Account, LTRIM(RTRIM(MA074)) AccountInvoice, LTRIM(RTRIM(MA104)) ClearanceType
                                , LTRIM(RTRIM(MA048)) ShipMethod, LTRIM(RTRIM(MA086)) InvocieAttachedStatus, LTRIM(RTRIM(MA137)) CustomerKind
                                , LTRIM(RTRIM(MA141)) UploadType, LTRIM(RTRIM(MA095)) DepositRate, LTRIM(RTRIM(MA028)) SaleRating
                                , LTRIM(RTRIM(MA029)) CreditRating, LTRIM(RTRIM(MA033)) CreditLimit, LTRIM(RTRIM(MA032)) CreditLimitControl
                                , LTRIM(RTRIM(MA133)) CreditLimitControlCurrency, LTRIM(RTRIM(MA088)) SoCreditAuditType, LTRIM(RTRIM(MA089)) SiCreditAuditType
                                , LTRIM(RTRIM(MA120)) DoCreditAuditType, LTRIM(RTRIM(MA132)) InTransitCreditAuditType, LTRIM(RTRIM(MA124)) RelatedPerson
                                , LTRIM(RTRIM(MA016)) SalesmenNo, LTRIM(RTRIM(MA085)) PaymentSalesmenNo, LTRIM(RTRIM(MA049)) CustomerRemark
                                , CASE WHEN LEN(LTRIM(RTRIM(CREATE_DATE))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_DATE)) as date), 'yyyy-MM-dd') ELSE NULL END TransferDate
                                , CASE WHEN LEN(LTRIM(RTRIM(CREATE_TIME))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_TIME)) as datetime), 'HH:mm:ss') ELSE NULL END TransferTime
                                FROM COPMA
                                WHERE 1=1";

                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "UpdateDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME > @UpdateDate OR MODI_DATE + ' ' + MODI_TIME > @UpdateDate)", UpdateDate);
                        sql += @" ORDER BY TransferDate, TransferTime";

                        customers = sqlConnection.Query<Customer>(sql, dynamicParameters).ToList();
                        #endregion
                    }

                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //撈取部門ID
                        sql = @"SELECT DepartmentId, DepartmentNo
                                FROM BAS.Department
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<Department> resultDepartments = sqlConnection.Query<Department>(sql, dynamicParameters).ToList();

                        customers = customers.GroupJoin(resultDepartments, x => x.DepartmentNo, y => y.DepartmentNo, (x, y) => { x.DepartmentId = y.FirstOrDefault()?.DepartmentId; return x; }).ToList();
                        #endregion

                        #region //撈取人員ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserId, a.UserNo 
                                FROM BAS.[User] a
                                INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                                WHERE b.CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<User> resultUsers = sqlConnection.Query<User>(sql, dynamicParameters).ToList();

                        customers = customers.GroupJoin(resultUsers, x => x.SalesmenNo, y => y.UserNo, (x, y) => { x.SalesmenId = y.FirstOrDefault()?.UserId; return x; }).ToList();
                        customers = customers.GroupJoin(resultUsers, x => x.PaymentSalesmenNo, y => y.UserNo, (x, y) => { x.PaymentSalesmenId = y.FirstOrDefault()?.UserId; return x; }).ToList();
                        #endregion

                        #region //判斷SCM.Customer是否有資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT CustomerId, CustomerNo
                                FROM SCM.Customer
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<Customer> resultCustomers = sqlConnection.Query<Customer>(sql, dynamicParameters).ToList();

                        customers = customers.GroupJoin(resultCustomers, x => x.CustomerNo, y => y.CustomerNo, (x, y) => { x.CustomerId = y.FirstOrDefault()?.CustomerId; return x; }).ToList();
                        #endregion

                        List<Customer> addCustomers = customers.Where(x => x.CustomerId == null).ToList();
                        List<Customer> updateCustomers = customers.Where(x => x.CustomerId != null).ToList();

                        #region //新增
                        if (addCustomers.Count > 0)
                        {
                            addCustomers
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.CompanyId = CompanyId;
                                    x.ShipType = "";
                                    x.ForwarderId = null;
                                    x.TransferStatus = "Y";
                                    x.Status = "A";
                                    x.CreateDate = CreateDate;
                                    x.LastModifiedDate = LastModifiedDate;
                                    x.CreateBy = CreateBy;
                                    x.LastModifiedBy = LastModifiedBy;
                                });

                            sql = @"INSERT INTO SCM.Customer (CompanyId, DepartmentId, CustomerNo, CustomerName, CustomerEnglishName, CustomerShortName
                                    , Country, Region, Route, GuiNumber, ResponsiblePerson, Contact, TelNoFirst, TelNoSecond, FaxNo, Email, RegisterAddressFirst
                                    , RegisterAddressSecond, InvoiceAddressFirst, InvoiceAddressSecond, DeliveryAddressFirst, DeliveryAddressSecond, DocumentAddressFirst
                                    , DocumentAddressSecond, BillAddressFirst, BillAddressSecond, ZipCodeRegister, ZipCodeInvoice, ZipCodeDelivery, ZipCodeDocument
                                    , ZipCodeBill, BillReceipient, InauguateDate, AccountDay, CloseDate, PermitDate, Version, Capital, Headcount, HomeOffice
                                    , AnnualTurnover, Currency, TradeTerm, PaymentTerm, PricingType, PaymentType, ReceiptReceive, TaxAmountCalculateType, DocumentDeliver
                                    , InvoiceCount, TaxNo, Taxation, PaymentBankFirst, PaymentBankSecond, PaymentBankThird, BankAccountFirst, BankAccountSecond
                                    , BankAccountThird, Account, AccountInvoice, ClearanceType, ShipMethod, InvocieAttachedStatus, CustomerKind, UploadType, DepositRate
                                    , SaleRating, CreditRating, CreditLimit, CreditLimitControl, CreditLimitControlCurrency, SoCreditAuditType, SiCreditAuditType
                                    , DoCreditAuditType, InTransitCreditAuditType, RelatedPerson, SalesmenId, PaymentSalesmenId, ShipType, ForwarderId, CustomerRemark
                                    , TransferStatus, TransferDate, Status
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@CompanyId, @DepartmentId, @CustomerNo, @CustomerName, @CustomerEnglishName, @CustomerShortName
                                    , @Country, @Region, @Route, @GuiNumber, @ResponsiblePerson, @Contact, @TelNoFirst, @TelNoSecond, @FaxNo, @Email, @RegisterAddressFirst
                                    , @RegisterAddressSecond, @InvoiceAddressFirst, @InvoiceAddressSecond, @DeliveryAddressFirst, @DeliveryAddressSecond, @DocumentAddressFirst
                                    , @DocumentAddressSecond, @BillAddressFirst, @BillAddressSecond, @ZipCodeRegister, @ZipCodeInvoice, @ZipCodeDelivery, @ZipCodeDocument
                                    , @ZipCodeBill, @BillReceipient, @InauguateDate, @AccountDay, @CloseDate, @PermitDate, @Version, @Capital, @Headcount, @HomeOffice
                                    , @AnnualTurnover, @Currency, @TradeTerm, @PaymentTerm, @PricingType, @PaymentType, @ReceiptReceive, @TaxAmountCalculateType, @DocumentDeliver
                                    , @InvoiceCount, @TaxNo, @Taxation, @PaymentBankFirst, @PaymentBankSecond, @PaymentBankThird, @BankAccountFirst, @BankAccountSecond
                                    , @BankAccountThird, @Account, @AccountInvoice, @ClearanceType, @ShipMethod, @InvocieAttachedStatus, @CustomerKind, @UploadType, @DepositRate
                                    , @SaleRating , @CreditRating, @CreditLimit, @CreditLimitControl, @CreditLimitControlCurrency, @SoCreditAuditType, @SiCreditAuditType
                                    , @DoCreditAuditType, @InTransitCreditAuditType, @RelatedPerson, @SalesmenId, @PaymentSalesmenId, @ShipType, @ForwarderId, @CustomerRemark
                                    , @TransferStatus, @TransferDate, @Status
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            rowsAffected += sqlConnection.Execute(sql, addCustomers);
                        }
                        #endregion

                        #region //修改
                        if (updateCustomers.Count > 0)
                        {
                            updateCustomers
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.TransferStatus = "Y";
                                    x.LastModifiedDate = LastModifiedDate;
                                    x.LastModifiedBy = LastModifiedBy;
                                });

                            sql = @"UPDATE SCM.Customer SET
                                    DepartmentId = @DepartmentId,
                                    CustomerNo = @CustomerNo,
                                    CustomerName = @CustomerName,
                                    CustomerEnglishName = @CustomerEnglishName,
                                    CustomerShortName = @CustomerShortName,
                                    Country = @Country,
                                    Region = @Region,
                                    Route = @Route,
                                    GuiNumber = @GuiNumber,
                                    ResponsiblePerson = @ResponsiblePerson,
                                    Contact = @Contact,
                                    TelNoFirst = @TelNoFirst,
                                    TelNoSecond = @TelNoSecond,
                                    FaxNo = @FaxNo,
                                    Email = @Email,
                                    RegisterAddressFirst = @RegisterAddressFirst,
                                    RegisterAddressSecond = @RegisterAddressSecond,
                                    InvoiceAddressFirst = @InvoiceAddressFirst,
                                    InvoiceAddressSecond = @InvoiceAddressSecond,
                                    DeliveryAddressFirst = @DeliveryAddressFirst,
                                    DeliveryAddressSecond = @DeliveryAddressSecond,
                                    DocumentAddressFirst = @DocumentAddressFirst,
                                    DocumentAddressSecond = @DocumentAddressSecond,
                                    BillAddressFirst = @BillAddressFirst,
                                    BillAddressSecond = @BillAddressSecond,
                                    ZipCodeRegister = @ZipCodeRegister,
                                    ZipCodeInvoice = @ZipCodeInvoice,
                                    ZipCodeDelivery = @ZipCodeDelivery,
                                    ZipCodeDocument = @ZipCodeDocument,
                                    ZipCodeBill = @ZipCodeBill,
                                    BillReceipient = @BillReceipient,
                                    InauguateDate = @InauguateDate,
                                    AccountDay = @AccountDay,
                                    CloseDate = @CloseDate,
                                    PermitDate = @PermitDate,
                                    Version = @Version,
                                    Capital = @Capital,
                                    Headcount = @Headcount,
                                    HomeOffice = @HomeOffice,
                                    AnnualTurnover = @AnnualTurnover,
                                    Currency = @Currency,
                                    TradeTerm = @TradeTerm,
                                    PaymentTerm = @PaymentTerm,
                                    PricingType = @PricingType,
                                    PaymentType = @PaymentType,
                                    ReceiptReceive = @ReceiptReceive,
                                    TaxAmountCalculateType = @TaxAmountCalculateType,
                                    DocumentDeliver = @DocumentDeliver,
                                    InvoiceCount = @InvoiceCount,
                                    TaxNo = @TaxNo,
                                    Taxation = @Taxation,
                                    PaymentBankFirst = @PaymentBankFirst,
                                    PaymentBankSecond = @PaymentBankSecond,
                                    PaymentBankThird = @PaymentBankThird,
                                    BankAccountFirst = @BankAccountFirst,
                                    BankAccountSecond = @BankAccountSecond,
                                    BankAccountThird = @BankAccountThird,
                                    Account = @Account,
                                    AccountInvoice = @AccountInvoice,
                                    ClearanceType = @ClearanceType,
                                    ShipMethod = @ShipMethod,
                                    InvocieAttachedStatus = @InvocieAttachedStatus,
                                    CustomerKind = @CustomerKind,
                                    UploadType = @UploadType,
                                    SaleRating = @SaleRating,
                                    CreditRating = @CreditRating,
                                    CreditLimit = @CreditLimit,
                                    CreditLimitControl = @CreditLimitControl,
                                    CreditLimitControlCurrency = @CreditLimitControlCurrency,
                                    SoCreditAuditType = @SoCreditAuditType,
                                    SiCreditAuditType = @SiCreditAuditType,
                                    DoCreditAuditType = @DoCreditAuditType,
                                    InTransitCreditAuditType = @InTransitCreditAuditType,
                                    RelatedPerson = @RelatedPerson,
                                    SalesmenId = @SalesmenId,
                                    PaymentSalesmenId = @PaymentSalesmenId,                                        
                                    CustomerRemark = @CustomerRemark,
                                    TransferDate = @TransferDate,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE CustomerId = @CustomerId";
                            rowsAffected += sqlConnection.Execute(sql, updateCustomers);
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

        #region //UpdateInventorySynchronize -- 庫別資料同步 -- Zoey 2022.06.27
        public string UpdateInventorySynchronize(string CompanyNo, string UpdateDate)
        {
            try
            {
                List<Inventory> inventories = new List<Inventory>();

                if (CompanyNo.Length <= 0) throw new SystemException("【公司】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
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
                        #region //撈取ERP庫別資料
                        sql = @"SELECT LTRIM(RTRIM(MC001)) InventoryNo, LTRIM(RTRIM(MC002)) InventoryName, LTRIM(RTRIM(MC004)) InventoryType
                                , LTRIM(RTRIM(MC005)) MrpCalculation, LTRIM(RTRIM(MC006)) ConfirmStatus, LTRIM(RTRIM(MC008)) SaveStatus, LTRIM(RTRIM(MC007)) InventoryDesc                            
                                , CASE WHEN LEN(LTRIM(RTRIM(CREATE_DATE))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_DATE)) as date), 'yyyy-MM-dd') ELSE NULL END TransferDate
                                , CASE WHEN LEN(LTRIM(RTRIM(CREATE_TIME))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_TIME)) as datetime), 'HH:mm:ss') ELSE NULL END TransferTime
                                FROM CMSMC
                                WHERE 1=1";

                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "UpdateDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME > @UpdateDate OR MODI_DATE + ' ' + MODI_TIME > @UpdateDate)", UpdateDate);
                        sql += @" ORDER BY TransferDate, TransferTime";

                        inventories = sqlConnection.Query<Inventory>(sql, dynamicParameters).ToList();
                        #endregion
                    }

                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷SCM.Inventory是否有資料
                        sql = @"SELECT InventoryId, InventoryNo
                                FROM SCM.Inventory
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<Inventory> resultInventories = sqlConnection.Query<Inventory>(sql, dynamicParameters).ToList();

                        inventories = inventories.GroupJoin(resultInventories, x => x.InventoryNo, y => y.InventoryNo, (x, y) => { x.InventoryId = y.FirstOrDefault()?.InventoryId; return x; }).ToList();
                        #endregion

                        List<Inventory> addInventories = inventories.Where(x => x.InventoryId == null).ToList();
                        List<Inventory> updateInventories = inventories.Where(x => x.InventoryId != null).ToList();

                        #region //新增
                        if (addInventories.Count > 0)
                        {
                            addInventories
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.CompanyId = CompanyId;
                                    x.TransferStatus = "Y";
                                    x.Status = "A";
                                    x.CreateDate = CreateDate;
                                    x.LastModifiedDate = LastModifiedDate;
                                    x.CreateBy = CreateBy;
                                    x.LastModifiedBy = LastModifiedBy;
                                });

                            sql = @"INSERT INTO SCM.Inventory (CompanyId, InventoryNo, InventoryName, InventoryType, MrpCalculation
                                    , ConfirmStatus, SaveStatus, InventoryDesc, TransferStatus, TransferDate, Status
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@CompanyId, @InventoryNo, @InventoryName, @InventoryType, @MrpCalculation
                                    , @ConfirmStatus, @SaveStatus, @InventoryDesc, @TransferStatus, @TransferDate, @Status
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            rowsAffected += sqlConnection.Execute(sql, addInventories);
                        }
                        #endregion

                        #region //修改
                        if (updateInventories.Count > 0)
                        {
                            updateInventories
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.TransferStatus = "Y";
                                    x.LastModifiedDate = LastModifiedDate;
                                    x.LastModifiedBy = LastModifiedBy;
                                });

                            sql = @"UPDATE SCM.Inventory SET
                                    InventoryNo = @InventoryNo,
                                    InventoryName = @InventoryName,
                                    InventoryType = @InventoryType,
                                    MrpCalculation = @MrpCalculation,
                                    ConfirmStatus = @ConfirmStatus,
                                    SaveStatus = @SaveStatus,
                                    InventoryDesc = @InventoryDesc,
                                    TransferStatus = @TransferStatus,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE InventoryId = @InventoryId";
                            rowsAffected += sqlConnection.Execute(sql, updateInventories);
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

        #region //UpdateSupplierSynchronize -- 供應商資料同步 -- Zoey 2022.06.29
        public string UpdateSupplierSynchronize(string CompanyNo, string UpdateDate)
        {
            try
            {
                List<Supplier> suppliers = new List<Supplier>();

                if (CompanyNo.Length <= 0) throw new SystemException("【公司】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
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
                        #region //撈取ERP供應商資料
                        sql = @"SELECT LTRIM(RTRIM(MA001)) SupplierNo, LTRIM(RTRIM(MA003)) SupplierName, LTRIM(RTRIM(MA002)) SupplierShortName
                                , LTRIM(RTRIM(MA088)) SupplierEnglishName, LTRIM(RTRIM(MA006)) Country, LTRIM(RTRIM(MA007)) Region
                                , LTRIM(RTRIM(MA005)) GuiNumber, LTRIM(RTRIM(MA012)) ResponsiblePerson, LTRIM(RTRIM(MA013)) ContactFirst
                                , LTRIM(RTRIM(MA048)) ContactSecond, LTRIM(RTRIM(MA049)) ContactThird, LTRIM(RTRIM(MA008)) TelNoFirst
                                , LTRIM(RTRIM(MA009)) TelNoSecond, LTRIM(RTRIM(MA010)) FaxNo, LTRIM(RTRIM(MA059)) FaxNoAccounting
                                , LTRIM(RTRIM(MA011)) Email, LTRIM(RTRIM(MA014)) AddressFirst, LTRIM(RTRIM(MA015)) AddressSecond
                                , LTRIM(RTRIM(MA046)) ZipCodeFirst, LTRIM(RTRIM(MA050)) ZipCodeSecond, LTRIM(RTRIM(MA051)) BillAddressFirst
                                , LTRIM(RTRIM(MA052)) BillAddressSecond, LTRIM(RTRIM(MA016)) PermitStatus
                                , null InauguateDate
                                , LTRIM(RTRIM(MA034)) AccountMonth, LTRIM(RTRIM(MA035)) AccountDay, LTRIM(RTRIM(MA069)) Version, LTRIM(RTRIM(MA018)) Capital
                                , LTRIM(RTRIM(MA019)) Headcount, LTRIM(RTRIM(MA020)) PoDeliver, LTRIM(RTRIM(MA021)) Currency, LTRIM(RTRIM(MA077)) TradeTerm
                                , LTRIM(RTRIM(MA024)) PaymentType, LTRIM(RTRIM(MA055)) PaymentTerm, LTRIM(RTRIM(MA029)) ReceiptReceive, LTRIM(RTRIM(MA030)) InvoiceCount
                                , LTRIM(RTRIM(MA076)) TaxNo, LTRIM(RTRIM(MA044)) Taxation, LTRIM(RTRIM(MA045)) PermitPartialDelivery, LTRIM(RTRIM(MA053)) TaxAmountCalculateType
                                , LTRIM(RTRIM(MA056)) InvocieAttachedStatus, LTRIM(RTRIM(MA057)) CertificateFormatType, LTRIM(RTRIM(MA058)) DepositRate, LTRIM(RTRIM(MA066)) TradeItem
                                , LTRIM(RTRIM(MA027)) RemitBank, LTRIM(RTRIM(MA028)) RemitAccount, LTRIM(RTRIM(MA041)) AccountPayable, LTRIM(RTRIM(MA042)) AccountOverhead
                                , LTRIM(RTRIM(MA043)) AccountInvoice, LTRIM(RTRIM(MA031)) SupplierLevel, LTRIM(RTRIM(MA032)) DeliveryRating, LTRIM(RTRIM(MA033)) QualityRating
                                , LTRIM(RTRIM(MA085)) RelatedPerson, LTRIM(RTRIM(MA047)) PurchaseUserNo, LTRIM(RTRIM(MA040)) SupplierRemark
                                , CASE WHEN LEN(LTRIM(RTRIM(CREATE_DATE))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_DATE)) as date), 'yyyy-MM-dd') ELSE NULL END TransferDate
                                , CASE WHEN LEN(LTRIM(RTRIM(CREATE_TIME))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_TIME)) as datetime), 'HH:mm:ss') ELSE NULL END TransferTime
                                FROM PURMA
                                WHERE 1=1";

                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "UpdateDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME > @UpdateDate OR MODI_DATE + ' ' + MODI_TIME > @UpdateDate)", UpdateDate);
                        sql += @" ORDER BY TransferDate, TransferTime";

                        suppliers = sqlConnection.Query<Supplier>(sql, dynamicParameters).ToList();
                        #endregion
                    }

                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //撈取人員ID
                        sql = @"SELECT a.UserId, a.UserNo 
                                FROM BAS.[User] a
                                INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                                WHERE b.CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<User> resultUsers = sqlConnection.Query<User>(sql, dynamicParameters).ToList();

                        suppliers = suppliers.GroupJoin(resultUsers, x => x.PurchaseUserNo, y => y.UserNo, (x, y) => { x.PurchaseUserId = y.FirstOrDefault()?.UserId; return x; }).ToList();
                        #endregion

                        #region //判斷SCM.Supplier是否有資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT SupplierId, SupplierNo
                                FROM SCM.Supplier
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<Supplier> resultSuppliers = sqlConnection.Query<Supplier>(sql, dynamicParameters).ToList();

                        suppliers = suppliers.GroupJoin(resultSuppliers, x => x.SupplierNo, y => y.SupplierNo, (x, y) => { x.SupplierId = y.FirstOrDefault()?.SupplierId; return x; }).ToList();
                        #endregion

                        List<Supplier> addSuppliers = suppliers.Where(x => x.SupplierId == null).ToList();
                        List<Supplier> updateSuppliers = suppliers.Where(x => x.SupplierId != null).ToList();

                        #region //新增
                        if (addSuppliers.Count > 0)
                        {
                            addSuppliers
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.CompanyId = CompanyId;
                                    x.TransferStatus = "Y";
                                    x.Status = "A";
                                    x.CreateDate = CreateDate;
                                    x.LastModifiedDate = LastModifiedDate;
                                    x.CreateBy = CreateBy;
                                    x.LastModifiedBy = LastModifiedBy;
                                });

                            sql = @"INSERT INTO SCM.Supplier (CompanyId, SupplierNo, SupplierName, SupplierShortName, SupplierEnglishName, Country, Region, GuiNumber
                                    , ResponsiblePerson, ContactFirst, ContactSecond, ContactThird, TelNoFirst, TelNoSecond, FaxNo, FaxNoAccounting, Email, AddressFirst, AddressSecond
                                    , ZipCodeFirst, ZipCodeSecond, BillAddressFirst, BillAddressSecond, PermitStatus, InauguateDate, AccountMonth, AccountDay, Version, Capital, Headcount
                                    , PoDeliver, Currency, TradeTerm, PaymentType, PaymentTerm, ReceiptReceive, InvoiceCount, TaxNo, Taxation, PermitPartialDelivery, TaxAmountCalculateType
                                    , InvocieAttachedStatus, CertificateFormatType, DepositRate, TradeItem, RemitBank, RemitAccount, AccountPayable, AccountOverhead, AccountInvoice
                                    , SupplierLevel, DeliveryRating, QualityRating, RelatedPerson, PurchaseUserId, SupplierRemark, TransferStatus, TransferDate, Status
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@CompanyId, @SupplierNo, @SupplierName, @SupplierShortName, @SupplierEnglishName, @Country, @Region, @GuiNumber
                                    , @ResponsiblePerson, @ContactFirst, @ContactSecond, @ContactThird, @TelNoFirst, @TelNoSecond, @FaxNo, @FaxNoAccounting, @Email, @AddressFirst, @AddressSecond
                                    , @ZipCodeFirst, @ZipCodeSecond, @BillAddressFirst, @BillAddressSecond, @PermitStatus, @InauguateDate, @AccountMonth, @AccountDay, @Version, @Capital, @Headcount
                                    , @PoDeliver, @Currency, @TradeTerm, @PaymentType, @PaymentTerm, @ReceiptReceive, @InvoiceCount, @TaxNo, @Taxation, @PermitPartialDelivery, @TaxAmountCalculateType
                                    , @InvocieAttachedStatus, @CertificateFormatType, @DepositRate, @TradeItem, @RemitBank, @RemitAccount, @AccountPayable, @AccountOverhead, @AccountInvoice
                                    , @SupplierLevel, @DeliveryRating, @QualityRating, @RelatedPerson, @PurchaseUserId, @SupplierRemark, @TransferStatus, @TransferDate, @Status
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            rowsAffected += sqlConnection.Execute(sql, addSuppliers);
                        }
                        #endregion

                        #region //修改
                        if (updateSuppliers.Count > 0)
                        {
                            updateSuppliers
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.TransferStatus = "Y";
                                    x.LastModifiedDate = LastModifiedDate;
                                    x.LastModifiedBy = LastModifiedBy;
                                });

                            sql = @"UPDATE SCM.Supplier SET
                                    SupplierNo = @SupplierNo,
                                    SupplierName = @SupplierName,
                                    SupplierShortName = @SupplierShortName,
                                    SupplierEnglishName = @SupplierEnglishName,
                                    Country = @Country,
                                    Region = @Region,
                                    GuiNumber = @GuiNumber,
                                    ResponsiblePerson = @ResponsiblePerson,
                                    ContactFirst = @ContactFirst,
                                    ContactSecond = @ContactSecond,
                                    ContactThird = @ContactThird,
                                    TelNoFirst = @TelNoFirst,
                                    TelNoSecond = @TelNoSecond,
                                    FaxNo = @FaxNo,
                                    FaxNoAccounting = @FaxNoAccounting,
                                    Email = @Email,
                                    AddressFirst = @AddressFirst,
                                    AddressSecond = @AddressSecond,
                                    ZipCodeFirst = @ZipCodeFirst,
                                    ZipCodeSecond = @ZipCodeSecond,
                                    BillAddressFirst = @BillAddressFirst,
                                    BillAddressSecond = @BillAddressSecond,
                                    PermitStatus = @PermitStatus,
                                    InauguateDate = @InauguateDate,
                                    AccountMonth = @AccountMonth,
                                    AccountDay = @AccountDay,
                                    Version = @Version,
                                    Capital = @Capital,
                                    Headcount = @Headcount,
                                    PoDeliver = @PoDeliver,
                                    Currency = @Currency,
                                    TradeTerm = @TradeTerm,
                                    PaymentType = @PaymentType,
                                    PaymentTerm = @PaymentTerm,
                                    ReceiptReceive = @ReceiptReceive,
                                    InvoiceCount = @InvoiceCount,
                                    TaxNo = @TaxNo,
                                    Taxation = @Taxation,
                                    PermitPartialDelivery = @PermitPartialDelivery,
                                    TaxAmountCalculateType = @TaxAmountCalculateType,
                                    InvocieAttachedStatus = @InvocieAttachedStatus,
                                    CertificateFormatType = @CertificateFormatType,
                                    TradeItem = @TradeItem,
                                    RemitBank = @RemitBank,
                                    RemitAccount = @RemitAccount,
                                    AccountPayable = @AccountPayable,
                                    AccountOverhead = @AccountOverhead,
                                    AccountInvoice = @AccountInvoice,
                                    SupplierLevel = @SupplierLevel,
                                    DeliveryRating = @DeliveryRating,
                                    QualityRating = @QualityRating,
                                    RelatedPerson = @RelatedPerson,
                                    PurchaseUserId = @PurchaseUserId,
                                    SupplierRemark = @SupplierRemark,
                                    TransferDate = @TransferDate,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE SupplierId = @SupplierId";
                            rowsAffected += sqlConnection.Execute(sql, updateSuppliers);
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

        #region//UpdateItemInventorySynchronize -- 品號庫別資料同步 -- Luca 2024.11.06
        public string UpdateItemInventorySynchronize(string CompanyNo, string UpdateDate)
        {
            try
            {
                // 設定 TransactionScope 選項，避免分散式交易的額外開銷
                var options = new TransactionOptions
                {
                    IsolationLevel = IsolationLevel.ReadCommitted,
                    Timeout = TimeSpan.FromMinutes(5) // 設定適當的超時時間
                };

                using (var transactionScope = new TransactionScope(TransactionScopeOption.Required, options))
                {
                    List<ItemInventory> itemInventory = new List<ItemInventory>();
                    if (string.IsNullOrEmpty(CompanyNo)) throw new SystemException("【公司】不能為空!");

                    int CompanyId = -1;
                    string ErpConnectionStrings = "";

                    // 取得公司資訊
                    using (var sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        sqlConnection.Open(); // 明確開啟連線

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 CompanyId, CompanyName, ErpDb
                        FROM BAS.Company
                        WHERE CompanyNo = @CompanyNo";
                        dynamicParameters.Add("CompanyNo", CompanyNo);
                        var result = sqlConnection.Query(sql, dynamicParameters).ToList();
                        if (!result.Any()) throw new SystemException("【公司】資料錯誤!");

                        var company = result.First();
                        CompanyId = Convert.ToInt32(company.CompanyId);
                        ErpConnectionStrings = ConfigurationManager.AppSettings[company.ErpDb];
                    }

                    // 從ERP取得資料
                    using (var sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        sqlConnection.Open();

                        sql = @"SELECT LTRIM(RTRIM(MC001)) MtlItemNo, 
                               LTRIM(RTRIM(MC002)) InventoryNo, 
                               LTRIM(RTRIM(MC007)) InventoryQty
                                FROM INVMC
                                WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "UpdateDate",
                                @" AND (CREATE_DATE + ' ' + CREATE_TIME > @UpdateDate 
                                OR MODI_DATE + ' ' + MODI_TIME > @UpdateDate)", UpdateDate);
                        itemInventory = sqlConnection.Query<ItemInventory>(sql, dynamicParameters).ToList();
                    }

                    // 處理本地資料庫操作
                    using (var sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        sqlConnection.Open();
                        dynamicParameters = new DynamicParameters();
                        dynamicParameters.Add("CompanyId", CompanyId);

                        // 建立查詢字典 - 品號對照
                        sql = @"SELECT MtlItemId, MtlItemNo
                        FROM PDM.MtlItem
                        WHERE CompanyId = @CompanyId";
                        var mtlItemDict = sqlConnection.Query<MtlItem>(sql, dynamicParameters)
                            .ToDictionary(x => x.MtlItemNo, x => (int?)x.MtlItemId);  // 明確指定為 nullable int

                        // 建立查詢字典 - 庫別對照
                        sql = @"SELECT InventoryId, InventoryNo
                        FROM SCM.Inventory
                        WHERE CompanyId = @CompanyId";
                        var inventoryDict = sqlConnection.Query<Inventory>(sql, dynamicParameters)
                        .ToDictionary(x => x.InventoryNo, x => (int?)x.InventoryId);  // 明確指定為 nullable int

                        // 轉換 ERP 資料，設定 MtlItemId 和 InventoryId
                        var validInventory = new List<ItemInventory>();
                        foreach (var item in itemInventory)
                        {
                            int? mtlItemId;
                            int? inventoryId;
                            bool hasMtlItem = mtlItemDict.TryGetValue(item.MtlItemNo, out mtlItemId);
                            bool hasInventory = inventoryDict.TryGetValue(item.InventoryNo, out inventoryId);

                            if (hasMtlItem && hasInventory && mtlItemId.HasValue && inventoryId.HasValue)
                            {
                                item.MtlItemId = mtlItemId.Value;
                                item.InventoryId = inventoryId.Value;
                                validInventory.Add(item);
                            }
                        }

                        // 取得現有庫存資料
                        sql = @"SELECT ItemInventoryId, MtlItemId, InventoryId, InventoryQty
                            FROM SCM.ItemInventory 
                            WHERE CompanyId = @CompanyId";
                        var existingInventoryDict = sqlConnection.Query<ItemInventory>(sql, dynamicParameters)
                            .ToDictionary(x => $"{x.MtlItemId}_{x.InventoryId}", x => x);

                        var itemsToAdd = new List<ItemInventory>();
                        var itemsToUpdate = new List<ItemInventory>();


                        // 分類需要新增和更新的資料
                        foreach (var item in validInventory)
                        {
                            var key = $"{item.MtlItemId}_{item.InventoryId}";

                            if (existingInventoryDict.TryGetValue(key, out var existingItem))
                            {
                                // 更新現有資料，保留 ItemInventoryId
                                var updateItem = new ItemInventory
                                {
                                    ItemInventoryId = existingItem.ItemInventoryId,  // 保留原有的 ID
                                    MtlItemId = item.MtlItemId,
                                    InventoryId = item.InventoryId,
                                    InventoryQty = item.InventoryQty,
                                    LastModifiedDate = LastModifiedDate,
                                    LastModifiedBy = LastModifiedBy
                                };
                                itemsToUpdate.Add(updateItem);
                            }
                            else
                            {
                                // 新增資料
                                var newItem = new ItemInventory
                                {
                                    CompanyId = CompanyId,
                                    MtlItemId = item.MtlItemId,
                                    InventoryId = item.InventoryId,
                                    InventoryQty = item.InventoryQty,
                                    CreateDate = CreateDate,
                                    LastModifiedDate = LastModifiedDate,
                                    CreateBy = CreateBy,
                                    LastModifiedBy = LastModifiedBy
                                };
                                itemsToAdd.Add(newItem);
                            }
                        }

                        int rowsAffected = 0;

                        // 執行新增
                        if (itemsToAdd.Any())
                        {
                            sql = @"INSERT INTO SCM.ItemInventory 
                            (CompanyId, MtlItemId, InventoryId, InventoryQty, 
                             CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                            VALUES 
                            (@CompanyId, @MtlItemId, @InventoryId, @InventoryQty, 
                             @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            rowsAffected += sqlConnection.Execute(sql, itemsToAdd);
                        }

                        // 執行更新
                        if (itemsToUpdate.Any())
                        {
                            sql = @"UPDATE SCM.ItemInventory 
                            SET InventoryQty = @InventoryQty,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                            WHERE ItemInventoryId = @ItemInventoryId";
                            rowsAffected += sqlConnection.Execute(sql, itemsToUpdate);
                        }

                        transactionScope.Complete(); // 提交交易

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

        #region //UpdateVolumeStatus -- 包材體積狀態更新 -- Zoey 2022.06.10
        public string UpdateVolumeStatus(int VolumeId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷體積資訊是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM SCM.PackingVolume
                                WHERE CompanyId = @CompanyId
                                AND VolumeId = @VolumeId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("VolumeId", VolumeId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("體積資料錯誤!");

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
                        sql = @"UPDATE SCM.PackingVolume SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE CompanyId = @CompanyId
                                AND VolumeId = @VolumeId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                CompanyId = CurrentCompany,
                                VolumeId
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

        #region //UpdatePackingWeight -- 包材重量資料更新 -- Zoey 2022.06.10
        public string UpdatePackingWeight(int WeightId, string WeightNo, string WeightName, string WeightSpec, double Weight, int WeightUomId)
        {
            try
            {
                if (WeightNo.Length <= 0) throw new SystemException("【重量代號】不能為空!");
                if (WeightNo.Length > 50) throw new SystemException("【重量代號】長度錯誤!");
                if (WeightName.Length <= 0) throw new SystemException("【重量名稱】不能為空!");
                if (WeightName.Length > 100) throw new SystemException("【重量名稱】長度錯誤!");
                if (WeightSpec.Length <= 0) throw new SystemException("【重量規格】不能為空!");
                if (WeightSpec.Length > 100) throw new SystemException("【重量規格】長度錯誤!");
                if (Weight < 0) throw new SystemException("【撿貨重量】不能為空!");
                if (WeightUomId <= 0) throw new SystemException("【重量單位】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷重量資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.PackingWeight
                                WHERE WeightId = @WeightId";
                        dynamicParameters.Add("WeightId", WeightId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("重量資料錯誤!");
                        #endregion

                        #region //判斷重量代號是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.PackingWeight
                                WHERE CompanyId = @CompanyId
                                AND WeightNo = @WeightNo
                                AND WeightId != @WeightId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("WeightNo", WeightNo);
                        dynamicParameters.Add("WeightId", WeightId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【重量代號】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.PackingWeight SET
                                WeightNo = @WeightNo,
                                WeightName = @WeightName,
                                WeightSpec = @WeightSpec,
                                Weight = @Weight,
                                WeightUomId = @WeightUomId,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE WeightId = @WeightId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                WeightNo,
                                WeightName,
                                WeightSpec,
                                Weight,
                                WeightUomId = WeightUomId > 0 ? (int?)WeightUomId : null,
                                LastModifiedDate,
                                LastModifiedBy,
                                WeightId
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

        #region //UpdateWeightStatus -- 包材重量狀態更新 -- Zoey 2022.06.10
        public string UpdateWeightStatus(int WeightId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷重量資訊是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM SCM.PackingWeight
                                WHERE CompanyId = @CompanyId
                                AND WeightId = @WeightId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("WeightId", WeightId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("重量資料錯誤!");

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
                        sql = @"UPDATE SCM.PackingWeight SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE CompanyId = @CompanyId
                                AND WeightId = @WeightId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                CompanyId = CurrentCompany,
                                WeightId
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

        #region //UpdateItemWeight -- 物件重量資料更新 -- Zoey 2022.06.10
        public string UpdateItemWeight(int ItemDefaultWeightId, int MtlModelId, string DefaultWeight, int DefaultWeightUomId)
        {
            try
            {
                if (MtlModelId <= 0) throw new SystemException("【物件機型名稱】不能為空!");
                if (DefaultWeight.Length <= 0) throw new SystemException("【物件重量】不能為空!");
                if (DefaultWeight.Length > 100) throw new SystemException("【物件重量】長度錯誤!");
                if (DefaultWeightUomId <= 0) throw new SystemException("【重量單位】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷物件重量資訊是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.ItemDefaultWeight
                                WHERE ItemDefaultWeightId = @ItemDefaultWeightId";
                        dynamicParameters.Add("ItemDefaultWeightId", ItemDefaultWeightId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("物件重量資料錯誤!");
                        #endregion

                        #region //判斷物件機型名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.ItemDefaultWeight
                                WHERE MtlModelId = @MtlModelId
                                AND ItemDefaultWeightId != @ItemDefaultWeightId";
                        dynamicParameters.Add("MtlModelId", MtlModelId);
                        dynamicParameters.Add("ItemDefaultWeightId", ItemDefaultWeightId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【物件機型名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.ItemDefaultWeight SET
                                MtlModelId = @MtlModelId,
                                DefaultWeight = @DefaultWeight,
                                DefaultWeightUomId = @DefaultWeightUomId,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ItemDefaultWeightId = @ItemDefaultWeightId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                MtlModelId = MtlModelId > 0 ? (int?)MtlModelId : null,
                                DefaultWeight,
                                DefaultWeightUomId = DefaultWeightUomId > 0 ? (int?)DefaultWeightUomId : null,
                                LastModifiedDate,
                                LastModifiedBy,
                                ItemDefaultWeightId
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

        #region //UpdateItemWeightStatus -- 物件重量狀態更新 -- Zoey 2022.06.13
        public string UpdateItemWeightStatus(int ItemDefaultWeightId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷物件重量資訊是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM SCM.ItemDefaultWeight
                                WHERE ItemDefaultWeightId = @ItemDefaultWeightId";
                        dynamicParameters.Add("ItemDefaultWeightId", ItemDefaultWeightId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("物件重量資料錯誤!");

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
                        sql = @"UPDATE SCM.ItemDefaultWeight SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ItemDefaultWeightId = @ItemDefaultWeightId";
                        dynamicParameters.AddDynamicParams(new
                        {
                            Status = status,
                            LastModifiedDate,
                            LastModifiedBy,
                            ItemDefaultWeightId
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

        #region //UpdateCustomer -- 客戶資料更新 -- Zoey 2022.06.16
        public string UpdateCustomer(int CustomerId, string CustomerNo, string CustomerName, string CustomerEnglishName, string CustomerShortName, string RelatedPerson
            , string PermitDate, string Version, string ResponsiblePerson, string Contact, string TelNoFirst, string TelNoSecond, string FaxNo, string Email
            , string GuiNumber, double Capital, double AnnualTurnover, int Headcount, string HomeOffice, string Currency, int DepartmentId, string CustomerKind
            , int SalesmenId, int PaymentSalesmenId, string InauguateDate, string CloseDate, string ZipCodeRegister, string RegisterAddressFirst, string RegisterAddressSecond
            , string ZipCodeInvoice, string InvoiceAddressFirst, string InvoiceAddressSecond, string ZipCodeDelivery, string DeliveryAddressFirst, string DeliveryAddressSecond
            , string ZipCodeDocument, string DocumentAddressFirst, string DocumentAddressSecond, string BillReceipient, string ZipCodeBill, string BillAddressFirst, string BillAddressSecond
            , string InvocieAttachedStatus, double DepositRate, string TaxAmountCalculateType, string SaleRating, string CreditRating, string TradeTerm, string PaymentTerm, string PricingType
            , string ClearanceType, string DocumentDeliver, string ReceiptReceive, string PaymentType, string TaxNo, string InvoiceCount, string Taxation, string Country
            , string Region, string Route, string UploadType, string PaymentBankFirst, string BankAccountFirst, string PaymentBankSecond, string BankAccountSecond
            , string PaymentBankThird, string BankAccountThird, string Account, string AccountInvoice, string AccountDay, string ShipMethod, string ShipType, int ForwarderId
            , string CustomerRemark, double CreditLimit, string CreditLimitControl, string CreditLimitControlCurrency, string SoCreditAuditType, string SiCreditAuditType
            , string DoCreditAuditType, string InTransitCreditAuditType, string TransferStatus, string Status)
        {
            try
            {
                #region //判斷客戶資料長度
                if (CustomerNo.Length <= 0) throw new SystemException("【客戶代號】不能為空!");
                if (CustomerName.Length <= 0) throw new SystemException("【客戶名稱】不能為空!");
                if (CustomerShortName.Length <= 0) throw new SystemException("【客戶簡稱】不能為空!");
                if (ResponsiblePerson.Length <= 0) throw new SystemException("【負責人】不能為空!");
                if (TelNoFirst.Length <= 0) throw new SystemException("【電話(一)】不能為空!");
                if (GuiNumber.Length <= 0) throw new SystemException("【統一編號】不能為空!");
                if (Capital < 0) throw new SystemException("【資本額】不能為空!");
                if (AnnualTurnover < 0) throw new SystemException("【年營業額】不能為空!");
                if (Headcount < 0) throw new SystemException("【員工人數】不能為空!");
                if (Currency.Length <= 0) throw new SystemException("【交易幣別】不能為空!");
                if (DepartmentId <= 0) throw new SystemException("【部門別】不能為空!");
                if (SalesmenId <= 0) throw new SystemException("【業務人員】不能為空!");
                if (PaymentSalesmenId <= 0) throw new SystemException("【收款業務員】不能為空!");
                if (RegisterAddressFirst.Length <= 0) throw new SystemException("【登記地址(中文)】不能為空!");
                if (InvoiceAddressFirst.Length <= 0) throw new SystemException("【發票地址(中文)】不能為空!");
                if (DeliveryAddressFirst.Length <= 0) throw new SystemException("【送貨地址(中文)】不能為空!");
                if (DocumentAddressFirst.Length <= 0) throw new SystemException("【文件地址(中文)】不能為空!");
                if (BillAddressFirst.Length <= 0) throw new SystemException("【帳單地址(中文)】不能為空!");
                if (DepositRate < 0) throw new SystemException("【訂金比率】不能為空!");
                if (TradeTerm.Length <= 0) throw new SystemException("【交易條件】不能為空!");
                if (PaymentTerm.Length <= 0) throw new SystemException("【付款條件】不能為空!");
                if (PricingType.Length <= 0) throw new SystemException("【計價方式】不能為空!");
                if (ClearanceType.Length <= 0) throw new SystemException("【通關方式】不能為空!");
                if (DocumentDeliver.Length <= 0) throw new SystemException("【採購單發送方式】不能為空!");
                if (ReceiptReceive.Length <= 0) throw new SystemException("【票據寄領】不能為空!");
                if (PaymentType.Length <= 0) throw new SystemException("【收款方式】不能為空!");
                if (TaxNo.Length <= 0) throw new SystemException("【稅別碼】不能為空!");
                if (InvoiceCount.Length <= 0) throw new SystemException("【發票聯數】不能為空!");
                if (Taxation.Length <= 0) throw new SystemException("【課稅別】不能為空!");
                if (ShipMethod.Length <= 0) throw new SystemException("【運輸方式】不能為空!");
                if (ForwarderId <= 0) throw new SystemException("【貨運承攬商】不能為空!");
                if (CreditLimit < 0) throw new SystemException("【信用額度】不能為空!");
                if (CreditLimitControlCurrency.Length <= 0) throw new SystemException("【信用額度控管幣別】不能為空!");
                if (SoCreditAuditType.Length <= 0) throw new SystemException("【訂單信用查核方式】不能為空!");
                if (SiCreditAuditType.Length <= 0) throw new SystemException("【銷貨信用查核方式】不能為空!");
                if (DoCreditAuditType.Length <= 0) throw new SystemException("【出貨通知查核方式】不能為空!");
                if (InTransitCreditAuditType.Length <= 0) throw new SystemException("【暫出單信用查核方式】不能為空!");
                if (CustomerNo.Length > 50) throw new SystemException("【客戶代號】長度錯誤!");
                if (CustomerName.Length > 200) throw new SystemException("【客戶名稱】長度錯誤!");
                if (CustomerEnglishName.Length > 80) throw new SystemException("【英文名稱】長度錯誤!");
                if (CustomerShortName.Length > 100) throw new SystemException("【客戶簡稱】長度錯誤!");
                if (ResponsiblePerson.Length > 30) throw new SystemException("【負責人】長度錯誤!");
                if (Contact.Length > 30) throw new SystemException("【聯絡人】長度錯誤!");
                if (TelNoFirst.Length > 20) throw new SystemException("【電話(一)】長度錯誤!");
                if (TelNoSecond.Length > 20) throw new SystemException("【電話(二)】長度錯誤!");
                if (FaxNo.Length > 20) throw new SystemException("【客戶傳真】長度錯誤!");
                if (Email.Length > 60) throw new SystemException("【電子郵件】長度錯誤!");
                if (GuiNumber.Length > 20) throw new SystemException("【統一編號】長度錯誤!");
                if (HomeOffice.Length > 10) throw new SystemException("【總店號】長度錯誤!");
                if (ZipCodeRegister.Length > 6) throw new SystemException("【郵遞區號(登記)】長度錯誤!");
                if (ZipCodeInvoice.Length > 6) throw new SystemException("【郵遞區號(發票)】長度錯誤!");
                if (ZipCodeDelivery.Length > 6) throw new SystemException("【郵遞區號(送貨)】長度錯誤!");
                if (ZipCodeDocument.Length > 6) throw new SystemException("【郵遞區號(文件)】長度錯誤!");
                if (ZipCodeBill.Length > 6) throw new SystemException("【郵遞區號(帳單)】長度錯誤!");
                if (RegisterAddressFirst.Length > 255) throw new SystemException("【登記地址(中文)】長度錯誤!");
                if (InvoiceAddressFirst.Length > 255) throw new SystemException("【發票地址(中文)】長度錯誤!");
                if (DeliveryAddressFirst.Length > 255) throw new SystemException("【送貨地址(中文)】長度錯誤!");
                if (DocumentAddressFirst.Length > 255) throw new SystemException("【文件地址(中文)】長度錯誤!");
                if (BillAddressFirst.Length > 255) throw new SystemException("【帳單地址(中文)】長度錯誤!");
                if (RegisterAddressSecond.Length > 255) throw new SystemException("【登記地址(英文)】長度錯誤!");
                if (InvoiceAddressSecond.Length > 255) throw new SystemException("【發票地址(英文)】長度錯誤!");
                if (DeliveryAddressSecond.Length > 255) throw new SystemException("【送貨地址(英文)】長度錯誤!");
                if (DocumentAddressSecond.Length > 255) throw new SystemException("【文件地址(英文)】長度錯誤!");
                if (BillAddressSecond.Length > 255) throw new SystemException("【帳單地址(英文)】長度錯誤!");
                if (BankAccountFirst.Length > 20) throw new SystemException("【銀行帳號(一)】長度錯誤!");
                if (BankAccountSecond.Length > 20) throw new SystemException("【銀行帳號(二)】長度錯誤!");
                if (BankAccountThird.Length > 20) throw new SystemException("【銀行帳號(三)】長度錯誤!");
                if (CustomerRemark.Length > 500) throw new SystemException("【備註】長度錯誤!");
                #endregion

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷客戶資訊是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.Customer
                                WHERE CustomerId = @CustomerId";
                        dynamicParameters.Add("CustomerId", CustomerId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("客戶資料錯誤!");
                        #endregion

                        #region //判斷客戶代號是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.Customer
                                WHERE CustomerNo = @CustomerNo
                                AND CustomerId != @CustomerId";
                        dynamicParameters.Add("CustomerNo", CustomerNo);
                        dynamicParameters.Add("CustomerId", CustomerId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【客戶代號】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.Customer SET
                                CustomerName = @CustomerName,
                                CustomerEnglishName = @CustomerEnglishName,
                                CustomerShortName = @CustomerShortName,
                                RelatedPerson = @RelatedPerson,
                                PermitDate = @PermitDate,
                                Version = @Version,
                                ResponsiblePerson = @ResponsiblePerson,
                                Contact = @Contact,
                                TelNoFirst = @TelNoFirst,
                                TelNoSecond = @TelNoSecond,
                                FaxNo = @FaxNo,
                                Email = @Email,
                                GuiNumber = @GuiNumber,
                                Capital = @Capital,
                                AnnualTurnover = @AnnualTurnover,
                                Headcount = @Headcount,
                                HomeOffice = @HomeOffice,
                                Currency = @Currency,
                                DepartmentId = @DepartmentId,
                                CustomerKind = @CustomerKind,
                                SalesmenId = @SalesmenId,
                                PaymentSalesmenId = @PaymentSalesmenId,
                                InauguateDate = @InauguateDate,
                                CloseDate = @CloseDate,
                                ZipCodeRegister = @ZipCodeRegister,
                                RegisterAddressFirst = @RegisterAddressFirst,
                                RegisterAddressSecond = @RegisterAddressSecond,
                                ZipCodeInvoice = @ZipCodeInvoice,
                                InvoiceAddressFirst = @InvoiceAddressFirst,
                                InvoiceAddressSecond = @InvoiceAddressSecond,
                                ZipCodeDelivery = @ZipCodeDelivery,
                                DeliveryAddressFirst = @DeliveryAddressFirst,
                                DeliveryAddressSecond = @DeliveryAddressSecond,
                                ZipCodeDocument = @ZipCodeDocument,
                                DocumentAddressFirst = @DocumentAddressFirst,
                                DocumentAddressSecond = @DocumentAddressSecond,
                                BillReceipient = @BillReceipient,
                                ZipCodeBill = @ZipCodeBill,
                                BillAddressFirst = @BillAddressFirst,
                                BillAddressSecond = @BillAddressSecond,
                                InvocieAttachedStatus = @InvocieAttachedStatus,
                                TaxAmountCalculateType = @TaxAmountCalculateType,
                                SaleRating = @SaleRating,
                                CreditRating = @CreditRating,
                                TradeTerm = @TradeTerm,
                                PaymentTerm = @PaymentTerm,
                                PricingType = @PricingType,
                                ClearanceType = @ClearanceType,
                                DocumentDeliver = @DocumentDeliver,
                                ReceiptReceive = @ReceiptReceive,
                                PaymentType = @PaymentType,
                                TaxNo = @TaxNo,
                                InvoiceCount = @InvoiceCount,
                                Taxation = @Taxation,
                                Country = @Country,
                                Region = @Region,
                                Route = @Route,
                                UploadType = @UploadType,
                                PaymentBankFirst = @PaymentBankFirst,
                                BankAccountFirst = @BankAccountFirst,
                                PaymentBankSecond = @PaymentBankSecond,
                                BankAccountSecond = @BankAccountSecond,
                                PaymentBankThird = @PaymentBankThird,
                                BankAccountThird = @BankAccountThird,
                                Account = @Account,
                                AccountInvoice = @AccountInvoice,
                                AccountDay = @AccountDay,
                                ShipMethod = @ShipMethod,
                                ShipType = @ShipType,
                                ForwarderId = @ForwarderId,
                                CustomerRemark = @CustomerRemark,
                                CreditLimit = @CreditLimit,
                                CreditLimitControl = @CreditLimitControl,
                                CreditLimitControlCurrency = @CreditLimitControlCurrency,
                                SoCreditAuditType = @SoCreditAuditType,
                                SiCreditAuditType = @SiCreditAuditType,
                                DoCreditAuditType = @DoCreditAuditType,
                                InTransitCreditAuditType = @InTransitCreditAuditType,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE CustomerId = @CustomerId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CustomerName,
                                CustomerEnglishName,
                                CustomerShortName,
                                RelatedPerson,
                                PermitDate = PermitDate.Length > 0 ? PermitDate : null,
                                Version,
                                ResponsiblePerson,
                                Contact,
                                TelNoFirst,
                                TelNoSecond,
                                FaxNo,
                                Email,
                                GuiNumber,
                                Capital,
                                AnnualTurnover,
                                Headcount,
                                HomeOffice,
                                Currency,
                                DepartmentId = DepartmentId > 0 ? (int?)DepartmentId : null,
                                CustomerKind,
                                SalesmenId = SalesmenId > 0 ? (int?)SalesmenId : null,
                                PaymentSalesmenId = PaymentSalesmenId > 0 ? (int?)PaymentSalesmenId : null,
                                InauguateDate = InauguateDate.Length > 0 ? InauguateDate : null,
                                CloseDate = CloseDate.Length > 0 ? CloseDate : null,
                                ZipCodeRegister,
                                RegisterAddressFirst,
                                RegisterAddressSecond,
                                ZipCodeInvoice,
                                InvoiceAddressFirst,
                                InvoiceAddressSecond,
                                ZipCodeDelivery,
                                DeliveryAddressFirst,
                                DeliveryAddressSecond,
                                ZipCodeDocument,
                                DocumentAddressFirst,
                                DocumentAddressSecond,
                                BillReceipient,
                                ZipCodeBill,
                                BillAddressFirst,
                                BillAddressSecond,
                                InvocieAttachedStatus,
                                TaxAmountCalculateType,
                                SaleRating,
                                CreditRating,
                                TradeTerm,
                                PaymentTerm,
                                PricingType,
                                ClearanceType,
                                DocumentDeliver,
                                ReceiptReceive,
                                PaymentType,
                                TaxNo,
                                InvoiceCount,
                                Taxation,
                                Country,
                                Region,
                                Route,
                                UploadType,
                                PaymentBankFirst,
                                BankAccountFirst,
                                PaymentBankSecond,
                                BankAccountSecond,
                                PaymentBankThird,
                                BankAccountThird,
                                Account,
                                AccountInvoice,
                                AccountDay,
                                ShipMethod,
                                ShipType,
                                ForwarderId = ForwarderId > 0 ? (int?)ForwarderId : null,
                                CustomerRemark,
                                CreditLimit,
                                CreditLimitControl,
                                CreditLimitControlCurrency,
                                SoCreditAuditType,
                                SiCreditAuditType,
                                DoCreditAuditType,
                                InTransitCreditAuditType,
                                LastModifiedDate,
                                LastModifiedBy,
                                CustomerId
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

        #region //UpdateCustomerStatus -- 客戶狀態更新 -- Zoey 2022.06.16
        public string UpdateCustomerStatus(int CustomerId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷客戶資訊是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM SCM.Customer
                                WHERE CustomerId = @CustomerId";
                        dynamicParameters.Add("CustomerId", CustomerId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("客戶資料錯誤!");

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
                        sql = @"UPDATE SCM.Customer SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE CustomerId = @CustomerId";
                        dynamicParameters.AddDynamicParams(new
                        {
                            Status = status,
                            LastModifiedDate,
                            LastModifiedBy,
                            CustomerId
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

        #region //UpdateForwarder -- 貨運承攬商資料更新 -- Zoey 2022.07.01
        public string UpdateForwarder(int ForwarderId, string ShipMethod, string ForwarderNo, string ForwarderName)
        {
            try
            {
                if (ShipMethod.Length <= 0) throw new SystemException("【運輸方式】不能為空!");
                if (ForwarderNo.Length <= 0) throw new SystemException("【貨運承攬商代號】不能為空!");
                if (ForwarderNo.Length > 50) throw new SystemException("【貨運承攬商代號】長度錯誤!");
                if (ForwarderName.Length <= 0) throw new SystemException("【貨運承攬商名稱】不能為空!");
                if (ForwarderName.Length > 100) throw new SystemException("【貨運承攬商名稱】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷貨運承攬商資訊是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.Forwarder
                                WHERE ForwarderId = @ForwarderId";
                        dynamicParameters.Add("ForwarderId", ForwarderId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("貨運承攬商資料錯誤!");
                        #endregion

                        #region //判斷貨運承攬商代號是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.Forwarder
                                WHERE ForwarderNo = @ForwarderNo
                                AND ForwarderId != @ForwarderId";
                        dynamicParameters.Add("ForwarderNo", ForwarderNo);
                        dynamicParameters.Add("ForwarderId", ForwarderId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【貨運承攬商代號】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.Forwarder SET
                                ShipMethod = @ShipMethod,
                                ForwarderNo = @ForwarderNo,
                                ForwarderName = @ForwarderName,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ForwarderId = @ForwarderId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ShipMethod,
                                ForwarderNo,
                                ForwarderName,
                                LastModifiedDate,
                                LastModifiedBy,
                                ForwarderId
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

        #region //UpdateForwarderStatus -- 貨運承攬商狀態更新 -- Zoey 2022.07.01
        public string UpdateForwarderStatus(int ForwarderId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷貨運承攬商資訊是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM SCM.Forwarder
                                WHERE ForwarderId = @ForwarderId";
                        dynamicParameters.Add("ForwarderId", ForwarderId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("貨運承攬商資料錯誤!");

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
                        sql = @"UPDATE SCM.Forwarder SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ForwarderId = @ForwarderId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                ForwarderId
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

        #region //UpdateDeliveryCustomer -- 送貨客戶資料更新 -- Zoey 2022.07.01
        public string UpdateDeliveryCustomer(int DcId, int CustomerId, string DcName, string DcEnglishName, string DcShortName, string Contact
            , string TelNo, string FaxNo, string RegisteredAddress, string DeliveryAddress, string ShipType, int ForwarderId, string Status)
        {
            try
            {
                if (CustomerId <= 0) throw new SystemException("【客戶來源】不能為空!");
                if (DcName.Length <= 0) throw new SystemException("【客戶名稱】不能為空!");
                if (DcName.Length > 200) throw new SystemException("【客戶名稱】長度錯誤!");
                if (DcEnglishName.Length > 80) throw new SystemException("【英文名稱】長度錯誤!");
                if (DcShortName.Length <= 0) throw new SystemException("【客戶簡稱】不能為空!");
                if (DcShortName.Length > 100) throw new SystemException("【客戶簡稱】長度錯誤!");
                if (Contact.Length > 30) throw new SystemException("【客戶聯絡人】長度錯誤!");
                if (TelNo.Length <= 0) throw new SystemException("【聯絡電話】不能為空!");
                if (TelNo.Length > 20) throw new SystemException("【聯絡電話】長度錯誤!");
                if (FaxNo.Length > 20) throw new SystemException("【聯絡傳真】長度錯誤!");
                if (RegisteredAddress.Length <= 0) throw new SystemException("【登記地址】不能為空!");
                if (RegisteredAddress.Length > 255) throw new SystemException("【登記地址】長度錯誤!");
                if (DeliveryAddress.Length <= 0) throw new SystemException("【送貨地址】不能為空!");
                if (DeliveryAddress.Length > 255) throw new SystemException("【送貨地址】長度錯誤!");
                if (ShipType.Length > 50) throw new SystemException("【運輸種類】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷送貨客戶資訊是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.DeliveryCustomer
                                WHERE DcId = @DcId";
                        dynamicParameters.Add("DcId", DcId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("送貨客戶資料錯誤!");
                        #endregion

                        #region //判斷送貨客戶名稱是否重複
                        //dynamicParameters = new DynamicParameters();
                        //sql = @"SELECT TOP 1 1
                        //        FROM SCM.DeliveryCustomer
                        //        WHERE DcName = @DcName
                        //        AND DcId != @DcId";
                        //dynamicParameters.Add("DcName", DcName);
                        //dynamicParameters.Add("DcId", DcId);

                        //var result2 = sqlConnection.Query(sql, dynamicParameters);
                        //if (result2.Count() > 0) throw new SystemException("【送貨客戶名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.DeliveryCustomer SET
                                CustomerId = @CustomerId,
                                DcName = @DcName,
                                DcEnglishName = @DcEnglishName,
                                DcShortName = @DcShortName,
                                Contact = @Contact,
                                TelNo = @TelNo,
                                FaxNo = @FaxNo,
                                RegisteredAddress = @RegisteredAddress,
                                DeliveryAddress = @DeliveryAddress,
                                ShipType = @ShipType,
                                ForwarderId = @ForwarderId,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE DcId = @DcId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CustomerId = CustomerId > 0 ? (int?)CustomerId : null,
                                DcName,
                                DcEnglishName,
                                DcShortName,
                                Contact,
                                TelNo,
                                FaxNo,
                                RegisteredAddress,
                                DeliveryAddress,
                                ShipType,
                                ForwarderId = ForwarderId > 0 ? (int?)ForwarderId : null,
                                LastModifiedDate,
                                LastModifiedBy,
                                DcId
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

        #region //UpdateDeliveryCustomerStatus -- 送貨客戶狀態更新 -- Zoey 2022.07.01
        public string UpdateDeliveryCustomerStatus(int DcId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷送貨客戶資訊是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM SCM.DeliveryCustomer
                                WHERE DcId = @DcId";
                        dynamicParameters.Add("DcId", DcId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("送貨客戶資料錯誤!");

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
                        sql = @"UPDATE SCM.DeliveryCustomer SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE DcId = @DcId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                DcId
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

        #region //UpdateSupplier -- 供應商資料更新 -- Zoey 2022.07.06
        public string UpdateSupplier(int SupplierId, string SupplierNo, string SupplierName, string SupplierShortName, string SupplierEnglishName, string Country
            , string Region, string GuiNumber, string ResponsiblePerson, string ContactFirst, string ContactSecond, string ContactThird, string TelNoFirst, string TelNoSecond
            , string FaxNo, string FaxNoAccounting, string Email, string AddressFirst, string AddressSecond, string ZipCodeFirst, string ZipCodeSecond
            , string BillAddressFirst, string BillAddressSecond, string PermitStatus, string InauguateDate, string AccountMonth, string AccountDay
            , string Version, double Capital, int Headcount, string PoDeliver, string Currency, string TradeTerm, string PaymentType, string PaymentTerm
            , string ReceiptReceive, string InvoiceCount, string TaxNo, string Taxation, string PermitPartialDelivery, string TaxAmountCalculateType, string InvocieAttachedStatus
            , string CertificateFormatType, double DepositRate, string TradeItem, string RemitBank, string RemitAccount, string AccountPayable, string AccountOverhead
            , string AccountInvoice, string SupplierLevel, string DeliveryRating, string QualityRating, string RelatedPerson, int PurchaseUserId, string SupplierRemark
            , string PassStationControl, string TransferStatus, string Status)
        {
            try
            {
                #region //判斷供應商資料長度
                if (SupplierNo.Length <= 0) throw new SystemException("【供應商代號】不能為空!");
                if (SupplierName.Length <= 0) throw new SystemException("【供應商名稱 】不能為空!");
                if (SupplierEnglishName.Length <= 0) throw new SystemException("【英文名稱】不能為空!");
                if (SupplierShortName.Length <= 0) throw new SystemException("【供應商簡稱】不能為空!");
                if (PermitStatus.Length <= 0) throw new SystemException("【核准狀態碼】不能為空!");
                if (ResponsiblePerson.Length <= 0) throw new SystemException("【負責人】不能為空!");
                if (TelNoFirst.Length <= 0) throw new SystemException("【電話(一)】不能為空!");
                if (PurchaseUserId <= 0) throw new SystemException("【採購人員】不能為空!");
                if (Capital < 0) throw new SystemException("【資本額】不能為空!");
                if (Headcount < 0) throw new SystemException("【員工人數】不能為空!");
                if (AddressFirst.Length <= 0) throw new SystemException("【聯絡地址】不能為空!");
                if (CertificateFormatType.Length <= 0) throw new SystemException("【憑證列印格式】不能為空!");
                if (Currency.Length <= 0) throw new SystemException("【交易幣別】不能為空!");
                if (TradeTerm.Length <= 0) throw new SystemException("【交易條件】不能為空!");
                if (TaxAmountCalculateType.Length <= 0) throw new SystemException("【稅額計算方式】不能為空!");
                if (PoDeliver.Length <= 0) throw new SystemException("【採購單發送方式】不能為空!");
                if (DepositRate <= 0) throw new SystemException("【訂金比率】不能為空!");
                if (PaymentType.Length <= 0) throw new SystemException("【付款方式】不能為空!");
                if (PaymentTerm.Length <= 0) throw new SystemException("【付款條件】不能為空!");
                if (RemitBank.Length <= 0) throw new SystemException("【匯款銀行】不能為空!");
                if (RemitAccount.Length <= 0) throw new SystemException("【銀行帳號】不能為空!");
                if (AccountDay.Length <= 0) throw new SystemException("【結帳日期(日)】不能為空!");
                if (ReceiptReceive.Length <= 0) throw new SystemException("【票據寄領】不能為空!");
                if (TaxNo.Length <= 0) throw new SystemException("【稅別碼】不能為空!");
                if (InvoiceCount.Length <= 0) throw new SystemException("【發票聯數】不能為空!");
                if (Taxation.Length <= 0) throw new SystemException("【課稅別】不能為空!");
                if (DeliveryRating.Length <= 0) throw new SystemException("【交貨評等】不能為空!");
                if (SupplierNo.Length > 50) throw new SystemException("【供應商代號】長度錯誤!");
                if (SupplierName.Length > 200) throw new SystemException("【供應商名稱 】長度錯誤!");
                if (SupplierEnglishName.Length > 100) throw new SystemException("【英文名稱】長度錯誤!");
                if (SupplierShortName.Length > 100) throw new SystemException("【客戶簡稱】長度錯誤!");
                if (ResponsiblePerson.Length > 30) throw new SystemException("【負責人】長度錯誤!");
                if (TelNoFirst.Length > 20) throw new SystemException("【電話(一)】長度錯誤!");
                if (TelNoSecond.Length > 20) throw new SystemException("【電話(二)】長度錯誤!");
                if (ContactFirst.Length > 20) throw new SystemException("【聯絡人(一)】長度錯誤!");
                if (ContactSecond.Length > 20) throw new SystemException("【聯絡人(二)】長度錯誤!");
                if (ContactThird.Length > 20) throw new SystemException("【聯絡人(三)】長度錯誤!");
                if (Email.Length > 60) throw new SystemException("【電子郵件】長度錯誤!");
                if (FaxNo.Length > 20) throw new SystemException("【客戶傳真】長度錯誤!");
                if (FaxNoAccounting.Length > 20) throw new SystemException("【客戶傳真(會計)】長度錯誤!");
                if (ZipCodeFirst.Length > 6) throw new SystemException("【郵遞區號(聯絡)】長度錯誤!");
                if (ZipCodeSecond.Length > 6) throw new SystemException("【郵遞區號(帳單)】長度錯誤!");
                if (AddressFirst.Length > 255) throw new SystemException("【聯絡地址(一)】長度錯誤!");
                if (AddressSecond.Length > 255) throw new SystemException("【聯絡地址(二)】長度錯誤!");
                if (BillAddressFirst.Length > 255) throw new SystemException("【帳單地址(一)】長度錯誤!");
                if (BillAddressSecond.Length > 255) throw new SystemException("【帳單地址(二)】長度錯誤!");
                if (TradeItem.Length > 255) throw new SystemException("【交易項目】長度錯誤!");
                if (RemitAccount.Length > 30) throw new SystemException("【銀行帳號】長度錯誤!");
                if (SupplierRemark.Length > 500) throw new SystemException("【備註】長度錯誤!");
                #endregion

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷供應商資訊是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.Supplier
                                WHERE SupplierId = @SupplierId";
                        dynamicParameters.Add("SupplierId", SupplierId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("供應商資料錯誤!");
                        #endregion

                        #region //判斷供應商代號是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.Supplier
                                WHERE SupplierNo = @SupplierNo
                                AND SupplierId != @SupplierId";
                        dynamicParameters.Add("SupplierNo", SupplierNo);
                        dynamicParameters.Add("SupplierId", SupplierId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【供應商代號】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.Supplier SET
                                SupplierNo = @SupplierNo,
                                SupplierName = @SupplierName,
                                SupplierShortName = @SupplierShortName,
                                SupplierEnglishName = @SupplierEnglishName,
                                Country = @Country,
                                Region = @Region,
                                GuiNumber = @GuiNumber,
                                ResponsiblePerson = @ResponsiblePerson,
                                ContactFirst = @ContactFirst,
                                ContactSecond = @ContactSecond,
                                ContactThird = @ContactThird,
                                TelNoFirst = @TelNoFirst,
                                TelNoSecond = @TelNoSecond,
                                FaxNo = @FaxNo,
                                FaxNoAccounting = @FaxNoAccounting,
                                Email = @Email,
                                AddressFirst = @AddressFirst,
                                AddressSecond = @AddressSecond,
                                ZipCodeFirst = @ZipCodeFirst,
                                ZipCodeSecond = @ZipCodeSecond,
                                BillAddressFirst = @BillAddressFirst,
                                BillAddressSecond = @BillAddressSecond,
                                PermitStatus = @PermitStatus,
                                InauguateDate = @InauguateDate,
                                AccountMonth = @AccountMonth,
                                AccountDay = @AccountDay,
                                Version = @Version,
                                Capital = @Capital,
                                Headcount = @Headcount,
                                PoDeliver = @PoDeliver,
                                Currency = @Currency,
                                TradeTerm = @TradeTerm,
                                PaymentType = @PaymentType,
                                ReceiptReceive = @ReceiptReceive,
                                InvoiceCount = @InvoiceCount,
                                TaxNo = @TaxNo,
                                Taxation = @Taxation,
                                PermitPartialDelivery = @PermitPartialDelivery,
                                TaxAmountCalculateType = @TaxAmountCalculateType,
                                InvocieAttachedStatus = @InvocieAttachedStatus,
                                CertificateFormatType = @CertificateFormatType,
                                DepositRate = @DepositRate,
                                TradeItem = @TradeItem,
                                RemitBank = @RemitBank,
                                RemitAccount = @RemitAccount,
                                AccountPayable = @AccountPayable,
                                AccountOverhead = @AccountOverhead,
                                AccountInvoice = @AccountInvoice,
                                SupplierLevel = @SupplierLevel,
                                DeliveryRating = @DeliveryRating,
                                QualityRating = @QualityRating,
                                RelatedPerson = @RelatedPerson,
                                PurchaseUserId = @PurchaseUserId,
                                SupplierRemark = @SupplierRemark,
                                PassStationControl = @PassStationControl,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SupplierId = @SupplierId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SupplierNo,
                                SupplierName,
                                SupplierShortName,
                                SupplierEnglishName,
                                Country,
                                Region,
                                GuiNumber,
                                ResponsiblePerson,
                                ContactFirst,
                                ContactSecond,
                                ContactThird,
                                TelNoFirst,
                                TelNoSecond,
                                FaxNo,
                                FaxNoAccounting,
                                Email,
                                AddressFirst,
                                AddressSecond,
                                ZipCodeFirst,
                                ZipCodeSecond,
                                BillAddressFirst,
                                BillAddressSecond,
                                PermitStatus,
                                InauguateDate = InauguateDate.Length > 0 ? InauguateDate : null,
                                AccountMonth,
                                AccountDay,
                                Version,
                                Capital,
                                Headcount,
                                PoDeliver,
                                Currency,
                                TradeTerm,
                                PaymentType,
                                PaymentTerm,
                                ReceiptReceive,
                                InvoiceCount,
                                TaxNo,
                                Taxation,
                                PermitPartialDelivery,
                                TaxAmountCalculateType,
                                InvocieAttachedStatus,
                                CertificateFormatType,
                                DepositRate,
                                TradeItem,
                                RemitBank,
                                RemitAccount,
                                AccountPayable,
                                AccountOverhead,
                                AccountInvoice,
                                SupplierLevel,
                                DeliveryRating,
                                QualityRating,
                                RelatedPerson,
                                PurchaseUserId = PurchaseUserId > 0 ? (int?)PurchaseUserId : null,
                                SupplierRemark,
                                PassStationControl,
                                LastModifiedDate,
                                LastModifiedBy,
                                SupplierId
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

        #region //UpdateSupplierStatus -- 供應商狀態更新 -- Zoey 2022.07.06
        public string UpdateSupplierStatus(int SupplierId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷供應商資訊是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM SCM.Supplier
                                WHERE SupplierId = @SupplierId";
                        dynamicParameters.Add("SupplierId", SupplierId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("供應商資料錯誤!");

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
                        sql = @"UPDATE SCM.Supplier SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SupplierId = @SupplierId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                SupplierId
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

        #region //UpdateSupplierPassStationControl -- 供應商是否刷過站狀態更新 -- Shintokuro 2024.03.30
        public string UpdateSupplierPassStationControl(int SupplierId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷供應商資訊是否正確
                        sql = @"SELECT TOP 1 PassStationControl
                                FROM SCM.Supplier
                                WHERE SupplierId = @SupplierId";
                        dynamicParameters.Add("SupplierId", SupplierId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("供應商資料錯誤!");

                        string PassStationControl = "";
                        foreach (var item in result)
                        {
                            PassStationControl = item.PassStationControl;
                        }

                        #region //調整為相反狀態
                        switch (PassStationControl)
                        {
                            case "Y":
                                PassStationControl = "N";
                                break;
                            case "N":
                                PassStationControl = "Y";
                                break;
                        }
                        #endregion
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.Supplier SET
                                PassStationControl = @PassStationControl,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SupplierId = @SupplierId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                PassStationControl = PassStationControl,
                                LastModifiedDate,
                                LastModifiedBy,
                                SupplierId
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

        #region //UpdateInventory -- 庫別資料更新 -- Zoey 2022.07.06
        public string UpdateInventory(int InventoryId, string InventoryNo, string InventoryName, string InventoryType
            , string MrpCalculation, string ConfirmStatus, string SaveStatus, string InventoryDesc
            , string TransferStatus, string TransferDate, string Status)
        {
            try
            {
                if (InventoryNo.Length <= 0) throw new SystemException("【庫別代號】不能為空!");
                if (InventoryNo.Length > 50) throw new SystemException("【庫別代號】長度錯誤!");
                if (InventoryName.Length <= 0) throw new SystemException("【庫別名稱】不能為空!");
                if (InventoryName.Length > 100) throw new SystemException("【庫別名稱】長度錯誤!");
                if (InventoryType.Length <= 0) throw new SystemException("【庫別性質】不能為空!");
                if (InventoryDesc.Length > 100) throw new SystemException("【庫別描述】長度錯誤!");
                if (MrpCalculation.Length <= 0) throw new SystemException("【納入MRP計算】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷庫別資訊是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.Inventory
                                WHERE InventoryId = @InventoryId";
                        dynamicParameters.Add("InventoryId", InventoryId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("庫別資料錯誤!");
                        #endregion

                        #region //判斷庫別代號是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.Inventory
                                WHERE InventoryNo = @InventoryNo
                                AND InventoryId != @InventoryId";
                        dynamicParameters.Add("InventoryNo", InventoryNo);
                        dynamicParameters.Add("InventoryId", InventoryId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【庫別代號】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.Inventory SET
                                InventoryNo = @InventoryNo,
                                InventoryName = @InventoryName,
                                InventoryType = @InventoryType,
                                InventoryDesc = @InventoryDesc,
                                MrpCalculation = @MrpCalculation,
                                ConfirmStatus = @ConfirmStatus,
                                SaveStatus = @SaveStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE InventoryId = @InventoryId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                InventoryNo,
                                InventoryName,
                                InventoryType,
                                InventoryDesc,
                                MrpCalculation,
                                ConfirmStatus,
                                SaveStatus,
                                LastModifiedDate,
                                LastModifiedBy,
                                InventoryId
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

        #region //UpdateInventoryStatus -- 庫別狀態更新 -- Zoey 2022.07.18
        public string UpdateInventoryStatus(int InventoryId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷庫別資訊是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM SCM.Inventory
                                WHERE InventoryId = @InventoryId";
                        dynamicParameters.Add("InventoryId", InventoryId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("庫別資料錯誤!");

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
                        sql = @"UPDATE SCM.Inventory SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE InventoryId = @InventoryId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                InventoryId
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

        #region //UpdatePacking -- 包材資料更新 -- Zoey 2022.10.11
        public string UpdatePacking(int PackingId, string PackingName, string PackingType, string VolumeSpec, double Volume
             , int VolumeUomId, string WeightSpec, double Weight, int WeightUomId, string Status)
        {
            try
            {
                if (PackingName.Length <= 0) throw new SystemException("【包材名稱】不能為空!");
                if (PackingName.Length > 50) throw new SystemException("【包材名稱】長度錯誤!");
                if (PackingType.Length <= 0) throw new SystemException("【包材種類】不能為空!");
                if (VolumeSpec.Length <= 0) throw new SystemException("【體積規格】不能為空!");
                if (VolumeSpec.Length > 100) throw new SystemException("【體積規格】長度錯誤!");
                if (Volume <= 0) throw new SystemException("【包材體積】不能為空!");
                if (WeightUomId <= 0) throw new SystemException("【體積單位】不能為空!");
                //if (WeightSpec.Length <= 0) throw new SystemException("【重量規格】不能為空!");
                if (WeightSpec.Length > 100) throw new SystemException("【重量規格】長度錯誤!");
                if (Weight <= 0) throw new SystemException("【包材重量】不能為空!");
                if (WeightUomId <= 0) throw new SystemException("【重量單位】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷包材資訊是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.Packing
                                WHERE PackingId = @PackingId";
                        dynamicParameters.Add("PackingId", PackingId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("包材資料錯誤!");
                        #endregion

                        #region //判斷包材名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.Packing
                                WHERE PackingName = @PackingName
                                AND PackingId != @PackingId";
                        dynamicParameters.Add("PackingName", PackingName);
                        dynamicParameters.Add("PackingId", PackingId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【包材名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.Packing SET
                                PackingName = @PackingName,
                                PackingType = @PackingType,
                                VolumeSpec = @VolumeSpec,
                                Volume = @Volume,
                                VolumeUomId = @VolumeUomId,
                                WeightSpec = @WeightSpec,
                                Weight = @Weight,
                                WeightUomId = @WeightUomId,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE PackingId = @PackingId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                PackingName,
                                PackingType,
                                VolumeSpec,
                                Volume,
                                VolumeUomId,
                                WeightSpec,
                                Weight,
                                WeightUomId,
                                LastModifiedDate,
                                LastModifiedBy,
                                PackingId
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

        #region //UpdatePackingStatus -- 包材狀態更新 -- Zoey 2022.10.11
        public string UpdatePackingStatus(int PackingId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷包材資訊是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM SCM.Packing
                                WHERE PackingId = @PackingId";
                        dynamicParameters.Add("PackingId", PackingId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("包材資料錯誤!");

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
                        sql = @"UPDATE SCM.Packing SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE PackingId = @PackingId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                PackingId
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

        #region //UpdateExchangeRateSynchronize -- 匯率資料同步 -- Zoey 2023.01.31
        public string UpdateExchangeRateSynchronize(string CompanyNo, string UpdateDate)
        {
            try
            {
                List<ExchangeRate> exchangeRates = new List<ExchangeRate>();

                if (CompanyNo.Length <= 0) throw new SystemException("【公司】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
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
                        #region //撈取ERP匯率資料
                        sql = @"SELECT LTRIM(RTRIM(MG001)) Currency, FORMAT(CAST(LTRIM(RTRIM(MG002)) as date), 'yyyy-MM-dd') EffectiveDate
                                , CAST(LTRIM(RTRIM(MG003)) as float) BankBuyingRate, CAST(LTRIM(RTRIM(MG004)) as float) BankSellingRate
                                , CAST(LTRIM(RTRIM(MG005)) as float) CustomsBuyingRate, CAST(LTRIM(RTRIM(MG006)) as float) CustomsSellingRate
                                , CASE WHEN LEN(LTRIM(RTRIM(CREATE_DATE))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_DATE)) as date), 'yyyy-MM-dd') ELSE NULL END TransferDate
                                , CASE WHEN LEN(LTRIM(RTRIM(CREATE_TIME))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_TIME)) as datetime), 'HH:mm:ss') ELSE NULL END TransferTime
                                FROM CMSMG
                                WHERE 1=1";

                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "UpdateDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME > @UpdateDate OR MODI_DATE + ' ' + MODI_TIME > @UpdateDate)", UpdateDate);
                        sql += @" ORDER BY TransferDate, TransferTime";

                        exchangeRates = sqlConnection.Query<ExchangeRate>(sql, dynamicParameters).ToList();
                        #endregion
                    }

                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷SCM.ExchangeRate是否有資料
                        //sql = @"SELECT Currency, EffectiveDate
                        //        FROM SCM.ExchangeRate
                        //        WHERE CompanyId = @CompanyId";

                        // Note by Mark, 07/26 09:12, 之前取 MES2 的 SCM.ExchangeRate, 欠 ExchangeRateId 欄位
                        sql = @"SELECT Currency, EffectiveDate, ExchangeRateId
                                FROM SCM.ExchangeRate
                                WHERE CompanyId = @CompanyId";

                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<ExchangeRate> resultExchangeRates = sqlConnection.Query<ExchangeRate>(sql, dynamicParameters).ToList();

                        exchangeRates = exchangeRates.GroupJoin(resultExchangeRates, x => new { x.Currency, x.EffectiveDate }, y => new { y.Currency, y.EffectiveDate }, (x, y) => { x.ExchangeRateId = y.FirstOrDefault()?.ExchangeRateId; return x; }).ToList();

                        #endregion

                        List<ExchangeRate> addexchangeRates = exchangeRates.Where(x => x.ExchangeRateId == null).ToList();
                        List<ExchangeRate> updateexchangeRates = exchangeRates.Where(x => x.ExchangeRateId != null).ToList();

                        #region //新增
                        if (addexchangeRates.Count > 0)
                        {
                            addexchangeRates
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.CompanyId = CompanyId;
                                    x.CreateDate = CreateDate;
                                    x.LastModifiedDate = LastModifiedDate;
                                    x.CreateBy = CreateBy;
                                    x.LastModifiedBy = LastModifiedBy;
                                });

                            sql = @"INSERT INTO SCM.ExchangeRate (CompanyId, Currency, EffectiveDate, BankBuyingRate
                                    , BankSellingRate, CustomsBuyingRate, CustomsSellingRate
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@CompanyId, @Currency, @EffectiveDate, @BankBuyingRate
                                    , @BankSellingRate, @CustomsBuyingRate, @CustomsSellingRate
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            rowsAffected += sqlConnection.Execute(sql, addexchangeRates);
                        }
                        #endregion

                        #region //修改
                        if (updateexchangeRates.Count > 0)
                        {
                            updateexchangeRates
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.LastModifiedDate = LastModifiedDate;
                                    x.LastModifiedBy = LastModifiedBy;
                                });

                            sql = @"UPDATE SCM.ExchangeRate SET
                                    Currency = @Currency,
                                    EffectiveDate = @EffectiveDate,
                                    BankBuyingRate = @BankBuyingRate,
                                    BankSellingRate = @BankSellingRate,
                                    CustomsBuyingRate = @CustomsBuyingRate,
                                    CustomsSellingRate = @CustomsSellingRate,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE ExchangeRateId = @ExchangeRateId";
                            rowsAffected += sqlConnection.Execute(sql, updateexchangeRates);
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

        #region //UpdateRfqProductClass -- RFQ產品類型更新 -- Chia Yuan 2023.06.30

        public string UpdateRfqProductClass(int RfqProClassId, string RfqProductClassName)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(RfqProductClassName)) throw new SystemException("【RFQ產品類型】不能為空!");

                RfqProductClassName = RfqProductClassName.Trim();

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷RFQ產品類型資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.RfqProductClass
                                WHERE RfqProClassId = @RfqProClassId";
                        dynamicParameters.Add("RfqProClassId", RfqProClassId);

                        var result = sqlConnection.Query(sql, dynamicParameters).FirstOrDefault();
                        if (result == null) throw new SystemException("RFQ產品類型資料錯誤!");
                        #endregion

                        #region //判斷RFQ產品類型資料是否重複
                        sql = @"SELECT TOP 1 1
                                FROM SCM.RfqProductClass
                                WHERE RfqProductClassName = @RfqProductClassName
                                AND RfqProClassId != @RfqProClassId";
                        dynamicParameters.Add("RfqProductClassName", RfqProductClassName);
                        dynamicParameters.Add("RfqProClassId", RfqProClassId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters).FirstOrDefault();
                        if (result2 != null) throw new SystemException("【RFQ產品類型】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"update a set 
                                a.RfqProductClassName = @RfqProductClassName,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                from SCM.RfqProductClass a
                                where a.RfqProClassId = @RfqProClassId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RfqProductClassName,
                                LastModifiedDate,
                                LastModifiedBy,
                                RfqProClassId
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

        #region //UpdateRfqProductClassStatus RFQ產品類型狀態更新 -- RFQ產品類型更新 -- Chia Yuan 2023.06.30

        public string UpdateRfqProductClassStatus(int RfqProClassId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷RFQ產品類型資料是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM SCM.RfqProductClass
                                WHERE RfqProClassId = @RfqProClassId";
                        dynamicParameters.Add("RfqProClassId", RfqProClassId);

                        var result = sqlConnection.Query(sql, dynamicParameters).FirstOrDefault();
                        if (result == null) throw new SystemException("RFQ產品類型資料錯誤!");
                        #endregion

                        #region //調整為相反狀態
                        string status = result.Status;
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

                        dynamicParameters = new DynamicParameters();
                        sql = @"update a set
                                a.Status = @Status,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                from SCM.RfqProductClass a
                                where a.RfqProClassId = @RfqProClassId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                RfqProClassId
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

        #region //UpdateRfqProductType -- RFQ產品類別更新 -- Chia Yuan 2023.07.03

        public string UpdateRfqProductType(int RfqProTypeId, string RfqProductTypeName)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(RfqProductTypeName)) throw new SystemException("【RFQ產品類別】不能為空!");

                RfqProductTypeName = RfqProductTypeName.Trim();

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷RFQ產品類別資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.RfqProductType
                                WHERE RfqProTypeId = @RfqProTypeId";
                        dynamicParameters.Add("RfqProTypeId", RfqProTypeId);

                        var result = sqlConnection.Query(sql, dynamicParameters).FirstOrDefault();
                        if (result == null) throw new SystemException("RFQ產品類別資料錯誤!");
                        #endregion

                        #region //判斷RFQ產品類別資料是否重複
                        sql = @"SELECT TOP 1 1
                                FROM SCM.RfqProductType
                                WHERE RfqProductTypeName = @RfqProductTypeName
                                AND RfqProTypeId != @RfqProTypeId";
                        dynamicParameters.Add("RfqProductTypeName", RfqProductTypeName);
                        dynamicParameters.Add("RfqProTypeId", RfqProTypeId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters).FirstOrDefault();
                        if (result2 != null) throw new SystemException("【RFQ產品類別】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"update a set 
                                a.RfqProductTypeName = @RfqProductTypeName,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                from SCM.RfqProductType a
                                where a.RfqProTypeId = @RfqProTypeId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RfqProductTypeName,
                                LastModifiedDate,
                                LastModifiedBy,
                                RfqProTypeId
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

        #region //UpdateRfqProductTypeStatus -- RFQ產品類別狀態更新 -- Chia Yuan 2023.07.03
        public string UpdateRfqProductTypeStatus(int RfqProTypeId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷RFQ產品類別資料是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM SCM.RfqProductType
                                WHERE RfqProTypeId = @RfqProTypeId";
                        dynamicParameters.Add("RfqProTypeId", RfqProTypeId);

                        var result = sqlConnection.Query(sql, dynamicParameters).FirstOrDefault();
                        if (result == null) throw new SystemException("RFQ產品類別資料錯誤!");
                        #endregion

                        #region //調整為相反狀態
                        string status = result.Status;
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

                        dynamicParameters = new DynamicParameters();
                        sql = @"update a set
                                a.Status = @Status,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                from SCM.RfqProductType a
                                where a.RfqProTypeId = @RfqProTypeId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                RfqProTypeId
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

        #region //UpdateRfqPackageType -- RFQ包裝種類更新 -- Chia Yuan 2023.07.03

        public string UpdateRfqPackageType(int RfqPkTypeId, string PackagingMethod, string SustSupplyStatus)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(PackagingMethod)) throw new SystemException("【RFQ包裝方式】不能為空!");

                PackagingMethod = PackagingMethod.Trim();

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷RFQ包裝種類資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.RfqPackageType
                                WHERE RfqPkTypeId = @RfqPkTypeId";
                        dynamicParameters.Add("RfqPkTypeId", RfqPkTypeId);

                        var result = sqlConnection.Query(sql, dynamicParameters).FirstOrDefault();
                        if (result == null) throw new SystemException("RFQ包裝種類資料錯誤!");
                        #endregion

                        #region //判斷RFQ包裝種類資料是否重複
                        sql = @"SELECT TOP 1 1
                                FROM SCM.RfqPackageType
                                WHERE PackagingMethod = @PackagingMethod
                                AND RfqPkTypeId != @RfqPkTypeId";
                        dynamicParameters.Add("PackagingMethod", PackagingMethod);
                        dynamicParameters.Add("RfqPkTypeId", RfqPkTypeId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters).FirstOrDefault();
                        if (result2 != null) throw new SystemException("【RFQ包裝方式】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"update a set 
                                a.PackagingMethod = @PackagingMethod,
                                a.SustSupplyStatus = @SustSupplyStatus,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                from SCM.RfqPackageType a
                                where a.RfqPkTypeId = @RfqPkTypeId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                PackagingMethod,
                                SustSupplyStatus,
                                LastModifiedDate,
                                LastModifiedBy,
                                RfqPkTypeId
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

        #region //UpdateSustSupplyStatus -- RFQ包裝種類狀態更新 -- Chia Yuan 2023.07.03
        public string UpdateRfqPackageTypeStatus(int RfqPkTypeId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷RFQ產品類別資料是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM SCM.RfqPackageType
                                WHERE RfqPkTypeId = @RfqPkTypeId";
                        dynamicParameters.Add("RfqPkTypeId", RfqPkTypeId);

                        var result = sqlConnection.Query(sql, dynamicParameters).FirstOrDefault();
                        if (result == null) throw new SystemException("RFQ包裝種類資料錯誤!");
                        #endregion

                        #region //調整為相反狀態
                        string status = result.Status;
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

                        dynamicParameters = new DynamicParameters();
                        sql = @"update a set
                                a.Status = @Status,
                                a.LastModifiedDate = @LastModifiedDate,
                                a.LastModifiedBy = @LastModifiedBy
                                from SCM.RfqPackageType a
                                where a.RfqPkTypeId = @RfqPkTypeId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                RfqPkTypeId
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

        #region //UpdateProductUse -- RFQ產品用途更新 -- Yi 2023.07.12
        public string UpdateProductUse(int ProductUseId, string ProductUseNo, string ProductUseName, string TypeOne, string Status)
        {
            try
            {
                if (ProductUseName.Length <= 0) throw new SystemException("【產品用途名稱】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷產品用途資訊是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.ProductUse
                                WHERE ProductUseId = @ProductUseId";
                        dynamicParameters.Add("ProductUseId", ProductUseId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("產品用途資料錯誤!");
                        #endregion

                        #region //判斷產品用途名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.ProductUse
                                WHERE ProductUseName = @ProductUseName
                                AND ProductUseId != @ProductUseId";
                        dynamicParameters.Add("ProductUseName", ProductUseName);
                        dynamicParameters.Add("ProductUseId", ProductUseId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【產品用途名稱】重複，請重新輸入!");
                        #endregion

                        #region //取得產品類別
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Type]
                                WHERE TypeNo = @TypeNo
                                AND TypeSchema = 'ProductUse.Type'";
                        dynamicParameters.Add("TypeNo", TypeOne);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【產品類別】資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.ProductUse SET
                                ProductUseName = @ProductUseName,
                                TypeOne = @TypeOne,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ProductUseId = @ProductUseId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ProductUseName,
                                TypeOne,
                                LastModifiedDate,
                                LastModifiedBy,
                                ProductUseId
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

        #region //UpdateProductUseStatus -- RFQ產品用途狀態更新 -- Yi 2023.07.12
        public string UpdateProductUseStatus(int ProductUseId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷RFQ產品用途資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM SCM.ProductUse
                                WHERE ProductUseId = @ProductUseId";
                        dynamicParameters.Add("ProductUseId", ProductUseId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【產品用途】資料錯誤!");

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
                        sql = @"UPDATE SCM.ProductUse SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ProductUseId = @ProductUseId";
                        var parametersObject = new
                        {
                            Status = status,
                            LastModifiedDate,
                            LastModifiedBy,
                            ProductUseId
                        };
                        dynamicParameters.AddDynamicParams(parametersObject);

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

        #region //UpdateMember -- 電商註冊資訊更新 -- Yi 2023.07.18
        public string UpdateMember(int MemberId, string OrgShortName, string MemberName, string MemberEmail, string ContactName
            , string ContactPhone, string ContactEmail, string Description, string Address, int OrgId, int OrganizaitonType
            , int OrganizaitonTypeId, string OrganizaitonScale, string OrganizationCode, string Status)
        {
            try
            {
                #region //判斷資料必填
                if (OrgShortName.Length <= 0) throw new SystemException("【公司簡稱】不能為空!");
                if (MemberName.Length <= 0) throw new SystemException("【客戶名稱】不能為空!");
                if (MemberEmail.Length <= 0) throw new SystemException("【客戶電子郵件】不能為空!");
                if (ContactName.Length <= 0) throw new SystemException("【公司聯絡人】不能為空!");
                if (ContactPhone.Length <= 0) throw new SystemException("【聯絡電話】不能為空!");
                #endregion

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷電商註冊資訊是否正確
                        sql = @"SELECT TOP 1 1
                                FROM EIP.Member
                                WHERE MemberId = @MemberId";
                        dynamicParameters.Add("MemberId", MemberId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("電商客戶註冊資料錯誤!");
                        #endregion

                        #region //判斷MemberId對應之OrgId是否存在
                        sql = @"SELECT TOP 1 b.OrgId
                                FROM EIP.Member a
                                INNER JOIN EIP.MemberOrganization b ON b.MemberId = a.MemberId
                                WHERE a.MemberId = @MemberId";
                        dynamicParameters.Add("MemberId", MemberId);

                        var resultOrg = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in resultOrg)
                        {
                            OrgId = item.OrgId;
                        }

                        #endregion

                        #region //EIP.Member資料更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EIP.Member SET
                                OrgShortName = @OrgShortName,
                                MemberName = @MemberName,
                                MemberEmail = @MemberEmail,
                                ContactName = @ContactName,
                                ContactPhone = @ContactPhone,
                                ContactEmail = @ContactEmail,
                                Description = @Description,
                                Address = @Address,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE MemberId = @MemberId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                OrgShortName,
                                MemberName,
                                MemberEmail,
                                ContactName,
                                ContactPhone,
                                ContactEmail,
                                Description,
                                Address,
                                LastModifiedDate,
                                LastModifiedBy,
                                MemberId
                            });

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        if (MemberId >= 0 && OrgId >= 0)
                        {
                            #region //EIP.MemberOrganization資料更新
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE EIP.MemberOrganization SET
                                OrganizaitonTypeId = @OrganizaitonTypeId,
                                OrganizaitonType = @OrganizaitonType,
                                OrganizationCode = @OrganizationCode,
                                OrganizaitonScale = @OrganizaitonScale,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE MemberId = @MemberId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    OrganizaitonTypeId,
                                    OrganizaitonType,
                                    OrganizationCode,
                                    OrganizaitonScale,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    MemberId
                                });

                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }
                        else
                        {
                            #region //EIP.MemberOrganization資料新增
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO EIP.MemberOrganization (MemberId, OrganizaitonTypeId, OrganizaitonType, OrganizationCode, OrganizaitonScale
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                VALUES (@MemberId, @OrganizaitonTypeId, @OrganizaitonType, @OrganizationCode ,@OrganizaitonScale
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MemberId,
                                    OrganizaitonTypeId,
                                    OrganizaitonType,
                                    OrganizationCode,
                                    OrganizaitonScale,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected = insertResult.Count();
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

        #region //UpdateMemberStatus -- 電商註冊狀態更新 -- Yi 2023.07.18
        public string UpdateMemberStatus(int MemberId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷電商註冊資訊是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM EIP.Member
                                WHERE MemberId = @MemberId";
                        dynamicParameters.Add("MemberId", MemberId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【電商註冊】資料錯誤!");

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
                        var parametersObject = new
                        {
                            Status = status,
                            LastModifiedDate,
                            LastModifiedBy,
                            MemberId
                        };
                        dynamicParameters.AddDynamicParams(parametersObject);

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

        #region //UpdatePerformanceGoals -- 業務績效目標更新 -- Shintokuro 2023.11.29
        public string UpdatePerformanceGoals(int PgId, string PgName, string PgDesc, string StartDate, string EndDate)
        {
            try
            {
                if (PgName.Length <= 0) throw new SystemException("【目標名稱】不能為空!");
                if (PgName.Length > 100) throw new SystemException("【目標名稱】長度錯誤!");
                if (PgDesc.Length <= 0) throw new SystemException("【目標描述】不能為空!");
                if (PgDesc.Length > 100) throw new SystemException("【目標描述】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷業務績效目標資訊是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.PerformanceGoals
                                WHERE PgId = @PgId";
                        dynamicParameters.Add("PgId", PgId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("庫別資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.PerformanceGoals SET
                                PgName = @PgName,
                                PgDesc = @PgDesc,
                                StartDate = @StartDate,
                                EndDate = @EndDate,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE PgId = @PgId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                PgName,
                                PgDesc,
                                StartDate,
                                EndDate,
                                LastModifiedDate,
                                LastModifiedBy,
                                PgId
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

        #region //UpdatePerformanceGoalsStatus -- 業務績效目狀態更新 -- Shintokuro 2023.11.29
        public string UpdatePerformanceGoalsStatus(int PgId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷庫別資訊是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM SCM.PerformanceGoals
                                WHERE PgId = @PgId";
                        dynamicParameters.Add("PgId", PgId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("業務績效目標資料錯誤!");

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
                        sql = @"UPDATE SCM.PerformanceGoals SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE PgId = @PgId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                PgId
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

        #region //UpdatePgDetail -- 業務績效目標單身更新 -- Shintokuro 2023.11.29
        public string UpdatePgDetail(int PgDetailId, int PgId, string IntentType, int UserId, int IntentAmount)
        {
            try
            {
                if (PgDetailId <= 0) throw new SystemException("【業務績效目標單身Id】不能為空!");
                if (PgId <= 0) throw new SystemException("【業務績效目標單頭Id】不能為空!");
                if (IntentType.Length <= 0) throw new SystemException("【目標類別】不能為空!");
                if (UserId <= 0) throw new SystemException("【業務人員】不能為空!");
                if (IntentAmount <= 0) throw new SystemException("【目標金額】必須大於0!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷業務人員是否存在
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[User]
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", UserId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【業務人員】不存在，請重新確認!");
                        #endregion

                        #region //判斷業務人員是否重複
                        sql = @"SELECT TOP 1 1
                                FROM SCM.PgDetail 
                                WHERE IntentType = @IntentType
                                AND UserId = @UserId
                                AND PgId = PgId
                                AND PgDetailId != PgDetailId";
                        dynamicParameters.Add("IntentType", IntentType);
                        dynamicParameters.Add("UserId", UserId);
                        dynamicParameters.Add("PgId", PgId);
                        dynamicParameters.Add("PgDetailId", PgDetailId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【業務人員】已經存在" + IntentType + "類別中，請重新確認!");
                        #endregion
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.PgDetail SET
                                PgId = @PgId,
                                UserId = @UserId,
                                IntentType = @IntentType,
                                IntentAmount = @IntentAmount,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE PgDetailId = @PgDetailId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                PgId,
                                UserId,
                                IntentType,
                                IntentAmount,
                                LastModifiedDate,
                                LastModifiedBy,
                                PgDetailId
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

        #region //UpdatePgDetailConfirm -- 業務績效目標單身確認 -- Shintokuro 2023.11.29
        public string UpdatePgDetailConfirm(int PgDetailId)
        {
            try
            {
                if (PgDetailId <= 0) throw new SystemException("【業務績效目標單身Id】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.PgDetail SET
                                ConfirmStatus = @ConfirmStatus,
                                ConfirmDate = @ConfirmDate,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE PgDetailId = @PgDetailId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmStatus = 'Y',
                                ConfirmDate = LastModifiedDate,
                                LastModifiedDate,
                                LastModifiedBy,
                                PgDetailId
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

        #region //UpdatePgDetailReConfirm -- 業務績效目標單身反確認 -- Shintokuro 2023.11.29
        public string UpdatePgDetailReConfirm(int PgDetailId)
        {
            try
            {
                if (PgDetailId <= 0) throw new SystemException("【業務績效目標單身Id】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.PgDetail SET
                                ConfirmStatus = @ConfirmStatus,
                                ConfirmDate = @ConfirmDate,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE PgDetailId = @PgDetailId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmStatus = 'N',
                                ConfirmDate = LastModifiedDate,
                                LastModifiedDate,
                                LastModifiedBy,
                                PgDetailId
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

        #region //UpdateProductTypeGroup -- 更新產品群組資料 -- Shintokuro 2024.05.24
        public string UpdateProductTypeGroup(int ProTypeGroupId, string TypeOne, int RfqProClassId, string ProTypeGroupName, string CoatingFlag)
        {
            try
            {
                if (RfqProClassId <= 0) throw new SystemException("【RFQ產品類型】不能為空!");
                if (ProTypeGroupName.Length > 50) throw new SystemException("【產品群組】長度不可以超過50字元!");
                if (ProTypeGroupName.Length <= 0) throw new SystemException("【產品群組】不能為空!");
                if (CoatingFlag.Length <= 0) throw new SystemException("【是否鍍膜】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷RFQ產品類型是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.RfqProductClass
                                WHERE 1 = 1
                                AND RfqProClassId = @RfqProClassId";
                        dynamicParameters.Add("RfqProClassId", RfqProClassId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【RFQ產品類型】找不到，請重新輸入!");
                        #endregion

                        #region //判斷產品屬性類別是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Type]
                                WHERE TypeSchema = 'ProductTypeGroup.Attribute'
                                AND TypeNo = @TypeOne";
                        dynamicParameters.Add("TypeOne", TypeOne);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【產品屬性】資料錯誤!");
                        #endregion

                        #region //判斷產品群組是否存在
                        sql = @"SELECT TOP 1 1
                                FROM SCM.ProductTypeGroup
                                WHERE 1 = 1
                                AND ProTypeGroupId = @ProTypeGroupId";
                        dynamicParameters.Add("ProTypeGroupId", ProTypeGroupId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【產品群組】找不到，請重新輸入!");
                        #endregion

                        #region //判斷是否重複
                        sql = @"SELECT TOP 1 1
                                FROM SCM.ProductTypeGroup
                                WHERE 1 = 1
                                AND RfqProClassId = @RfqProClassId
                                AND ProTypeGroupName = @ProTypeGroupName
                                AND ProTypeGroupId != @ProTypeGroupId";
                        dynamicParameters.Add("RfqProClassId", RfqProClassId);
                        dynamicParameters.Add("ProTypeGroupName", ProTypeGroupName);
                        dynamicParameters.Add("ProTypeGroupId", ProTypeGroupId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【資料組合】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.ProductTypeGroup SET
                                RfqProClassId = @RfqProClassId,
                                TypeOne = @TypeOne,
                                ProTypeGroupName = @ProTypeGroupName,
                                CoatingFlag = @CoatingFlag,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ProTypeGroupId = @ProTypeGroupId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RfqProClassId,
                                TypeOne,
                                ProTypeGroupName,
                                CoatingFlag,
                                LastModifiedDate,
                                LastModifiedBy,
                                ProTypeGroupId
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

        #region //UpdateProductTypeGroupStatus -- 更新產品群組狀態資料 -- Shintokuro 2024.05.24
        public string UpdateProductTypeGroupStatus(int ProTypeGroupId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷產品群組是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM SCM.ProductTypeGroup
                                WHERE 1 = 1
                                AND ProTypeGroupId = @ProTypeGroupId";
                        dynamicParameters.Add("ProTypeGroupId", ProTypeGroupId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【產品群組】找不到!請重新確認!");

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
                        sql = @"UPDATE SCM.ProductTypeGroup SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE 1 = 1
                                AND ProTypeGroupId = @ProTypeGroupId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                CompanyId = CurrentCompany,
                                ProTypeGroupId
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

        #region //UpdateProductType -- 更新產品類別資料 -- Shintokuro 2024.05.24
        public string UpdateProductType(int ProTypeId, int ProTypeGroupId, string ProTypeName)
        {
            try
            {
                if (ProTypeGroupId <= 0) throw new SystemException("【產品群組】不能為空!");
                if (ProTypeName.Length > 100) throw new SystemException("【類別名稱】長度不能超過100字元!");
                if (ProTypeName.Length <= 0) throw new SystemException("【類別名稱】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷產品類別是否存在
                        sql = @"SELECT TOP 1 1
                                FROM SCM.ProductType
                                WHERE 1 = 1
                                AND ProTypeId = @ProTypeId";
                        dynamicParameters.Add("ProTypeId", ProTypeId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【產品類別】找不到，請重新輸入!");
                        #endregion

                        #region //判斷產品群組是否存在
                        sql = @"SELECT TOP 1 1
                                FROM SCM.ProductTypeGroup
                                WHERE 1 = 1
                                AND ProTypeGroupId = @ProTypeGroupId";
                        dynamicParameters.Add("ProTypeGroupId", ProTypeGroupId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【產品群組】找不到，請重新輸入!");
                        #endregion

                        #region //判斷是否重複
                        sql = @"SELECT TOP 1 1
                                FROM SCM.ProductType
                                WHERE 1 = 1
                                AND ProTypeGroupId = @ProTypeGroupId
                                AND ProTypeName = @ProTypeName
                                AND ProTypeId != @ProTypeId";
                        dynamicParameters.Add("ProTypeGroupId", ProTypeGroupId);
                        dynamicParameters.Add("ProTypeName", ProTypeName);
                        dynamicParameters.Add("ProTypeId", ProTypeId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【資料組合】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.ProductType SET
                                ProTypeGroupId = @ProTypeGroupId,
                                ProTypeName = @ProTypeName,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ProTypeId = @ProTypeId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ProTypeGroupId,
                                ProTypeName,
                                LastModifiedDate,
                                LastModifiedBy,
                                ProTypeId
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

        #region //UpdateTemplateRfiSignFlow -- 更新市場評估單簽核流程 -- Chia Yuan 2024.05.30
        public string UpdateTemplateRfiSignFlow(int TempRfiSfId, string UserList, string FlowJobName, string FlowStatus)
        {
            try
            {
                //if (string.IsNullOrWhiteSpace(UserList)) throw new SystemException("【關卡人員】不能為空!!");
                if (!string.IsNullOrWhiteSpace(FlowJobName))
                {
                    if (FlowJobName.Length > 50) throw new SystemException("【關卡人職稱】長度錯誤!");
                }

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷簽核流程是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 ProTypeGroupId
                                FROM RFI.TemplateRfiSignFlow
                                WHERE TempRfiSfId = @TempRfiSfId";
                        dynamicParameters.Add("TempRfiSfId", TempRfiSfId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("簽核流程資料錯誤!");
                        int proTypeGroupId = -1;
                        foreach (var item in result)
                        {
                            proTypeGroupId = item.ProTypeGroupId;
                        }
                        #endregion

                        if (string.IsNullOrWhiteSpace(UserList))
                        {
                            string flowStatus = "1"; //預設業務
                            if (!string.IsNullOrWhiteSpace(FlowStatus))
                            {
                                #region //取得關卡狀態
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 StatusNo
                                    FROM BAS.[Status]
                                    WHERE StatusSchema = @StatusSchema
                                    AND StatusNo = @StatusNo";
                                dynamicParameters.Add("StatusSchema", "RfiDetail.FlowStatus");
                                dynamicParameters.Add("StatusNo", FlowStatus);
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("關卡狀態資料錯誤!");
                                foreach (var item in result)
                                {
                                    flowStatus = item.StatusNo;
                                }
                                #endregion
                            }

                            #region //資料更新
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE a SET 
                                    a.FlowJobName = @FlowJobName,
                                    a.FlowStatus = @FlowStatus,
                                    a.LastModifiedDate = @LastModifiedDate,
                                    a.LastModifiedBy = @LastModifiedBy
                                    FROM RFI.TemplateRfiSignFlow a
                                    WHERE a.TempRfiSfId = @TempRfiSfId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    TempRfiSfId,
                                    FlowJobName = string.IsNullOrWhiteSpace(FlowJobName) ? null : FlowJobName.Trim(),
                                    FlowStatus = flowStatus,
                                    LastModifiedDate,
                                    LastModifiedBy
                                });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }
                        else
                        {
                            int userId = -1;
                            if (UserList != "-1")
                            {
                                var userIds = UserList.Split(',').Select(s => { int.TryParse(s, out int d); return d; }).ToArray();

                                #region //判斷使用者是否存在
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT DISTINCT UserId, UserNo, UserName
                                    FROM BAS.[User]
                                    WHERE Status = 'A'
                                    AND UserId IN @UserIds";
                                dynamicParameters.Add("UserIds", userIds);
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("使用者不存在，請重新確認!");
                                var nonUsers = result.Where(w => !userIds.ToList().Contains(w.UserId)).ToList();
                                if (nonUsers.Count > 0) throw new SystemException("使用者不存在:<br/>" + string.Join("<br/>", nonUsers.Select(s => string.Format("{0}({1})", s.UserName, s.UserNo))));
                                foreach (var item in result)
                                {
                                    userId = item.UserId;
                                }
                                #endregion

                                #region //判斷關卡人員是否存在 (停用)
                                //dynamicParameters = new DynamicParameters();
                                //sql = @"SELECT TOP 1 1
                                //        FROM RFI.TemplateRfiSignFlow
                                //        WHERE ProTypeGroupId = @ProTypeGroupId
                                //        AND FlowUser = @FlowUser
                                //        AND TempRfiSfId <> @TempRfiSfId";
                                //dynamicParameters.Add("ProTypeGroupId", proTypeGroupId);
                                //dynamicParameters.Add("FlowUser", userId);
                                //dynamicParameters.Add("TempRfiSfId", TempRfiSfId);
                                //result = sqlConnection.Query(sql, dynamicParameters);
                                //if (result.Count() > 0) throw new SystemException("關卡人員已存在，請重新確認!");
                                #endregion
                            }

                            #region //資料更新
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE a SET 
                                    a.FlowUser = @FlowUser,
                                    a.LastModifiedDate = @LastModifiedDate,
                                    a.LastModifiedBy = @LastModifiedBy
                                    FROM RFI.TemplateRfiSignFlow a
                                    WHERE a.TempRfiSfId = @TempRfiSfId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    TempRfiSfId,
                                    FlowUser = userId > 0 ? userId : (int?)null,
                                    LastModifiedDate,
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

        #region //UpdateTemplateDesignSignFlow -- 更新設計申請單簽核流程 -- Chia Yuan 2024.05.30
        public string UpdateTemplateDesignSignFlow(int TempDesignSfId, string UserList, string FlowJobName, string FlowStatus)
        {
            try
            {
                //if (string.IsNullOrWhiteSpace(UserList)) throw new SystemException("【關卡人員】不能為空!!");
                if (!string.IsNullOrWhiteSpace(FlowJobName))
                {
                    if (FlowJobName.Length > 50) throw new SystemException("【關卡人職稱】長度錯誤!");
                }

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷簽核流程是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 ProTypeGroupId
                                FROM RFI.TemplateDesignSignFlow
                                WHERE TempDesignSfId = @TempDesignSfId";
                        dynamicParameters.Add("TempDesignSfId", TempDesignSfId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("簽核流程資料錯誤!");
                        int proTypeGroupId = -1;
                        foreach (var item in result)
                        {
                            proTypeGroupId = item.ProTypeGroupId;
                        }
                        #endregion

                        if (string.IsNullOrWhiteSpace(UserList))
                        {
                            string flowStatus = "1"; //預設業務
                            if (!string.IsNullOrWhiteSpace(FlowStatus))
                            {
                                #region //取得關卡狀態
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 StatusNo
                                    FROM BAS.[Status]
                                    WHERE StatusSchema = @StatusSchema
                                    AND StatusNo = @StatusNo";
                                dynamicParameters.Add("StatusSchema", "RfiDesign.FlowStatus");
                                dynamicParameters.Add("StatusNo", FlowStatus);
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("關卡狀態資料錯誤!");
                                foreach (var item in result)
                                {
                                    flowStatus = item.StatusNo;
                                }
                                #endregion
                            }

                            #region //資料更新
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE a SET 
                                    a.FlowJobName = @FlowJobName,
                                    a.FlowStatus = @FlowStatus,
                                    a.LastModifiedDate = @LastModifiedDate,
                                    a.LastModifiedBy = @LastModifiedBy
                                    FROM RFI.TemplateDesignSignFlow a
                                    WHERE a.TempDesignSfId = @TempDesignSfId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    TempDesignSfId,
                                    FlowJobName = string.IsNullOrWhiteSpace(FlowJobName) ? null : FlowJobName.Trim(),
                                    FlowStatus = flowStatus,
                                    LastModifiedDate,
                                    LastModifiedBy
                                });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }
                        else
                        {
                            int userId = -1;
                            if (UserList != "-1")
                            {
                                var userIds = UserList.Split(',').Select(s => { int.TryParse(s, out int d); return d; }).ToArray();

                                #region //判斷使用者是否存在
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT DISTINCT UserId, UserNo, UserName
                                    FROM BAS.[User]
                                    WHERE Status = 'A'
                                    AND UserId IN @UserIds";
                                dynamicParameters.Add("UserIds", userIds);
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("使用者不存在，請重新確認!");
                                var nonUsers = result.Where(w => !userIds.ToList().Contains(w.UserId)).ToList();
                                if (nonUsers.Count > 0) throw new SystemException("使用者不存在:<br/>" + string.Join("<br/>", nonUsers.Select(s => string.Format("{0}({1})", s.UserName, s.UserNo))));
                                foreach (var item in result)
                                {
                                    userId = item.UserId;
                                }
                                #endregion
                            }

                            #region //判斷關卡人員是否存在 (停用)
                            //dynamicParameters = new DynamicParameters();
                            //sql = @"SELECT TOP 1 1
                            //        FROM RFI.TemplateDesignSignFlow
                            //        WHERE ProTypeGroupId = @ProTypeGroupId
                            //        AND FlowUser = @FlowUser
                            //        AND TempDesignSfId <> @TempDesignSfId";
                            //dynamicParameters.Add("ProTypeGroupId", proTypeGroupId);
                            //dynamicParameters.Add("FlowUser", userId);
                            //dynamicParameters.Add("TempDesignSfId", TempDesignSfId);
                            //result = sqlConnection.Query(sql, dynamicParameters);
                            //if (result.Count() > 0) throw new SystemException("關卡人員已存在，請重新確認!");
                            #endregion

                            #region //資料更新
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE a SET 
                                    a.FlowUser = @FlowUser,
                                    a.LastModifiedDate = @LastModifiedDate,
                                    a.LastModifiedBy = @LastModifiedBy
                                    FROM RFI.TemplateDesignSignFlow a
                                    WHERE a.TempDesignSfId = @TempDesignSfId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    TempDesignSfId,
                                    FlowUser = userId > 0 ? userId : (int?)null,
                                    LastModifiedDate,
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

        #region //UpdateTemplateRfiSignFlowSort -- 更新市場評估單簽核流程排序 -- Shintokuro 2024.05.24
        public string UpdateTemplateRfiSignFlowSort(int ProTypeGroupId, int DepartmentId, string SortList)
        {
            try
            {
                if (SortList.Length <= 0) throw new SystemException("【順序】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //取得群組
                        sql = @"SELECT TOP 1 1
                                FROM SCM.ProductTypeGroup
                                WHERE ProTypeGroupId = @ProTypeGroupId";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { ProTypeGroupId }) ?? throw new SystemException("【產品群組】資料錯誤!");
                        #endregion

                        #region //取得部門
                        sql = @"SELECT TOP 1 1 
                                FROM BAS.Department WHERE DepartmentId = @DepartmentId";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { DepartmentId }) ?? throw new SystemException("【部門】資料錯誤!");
                        #endregion

                        string[] SortArr = SortList.Split(',');

                        foreach (var TempRfiSfId in SortArr)
                        {
                            #region //判斷主資料是否有誤
                            sql = @"SELECT TOP 1 1
                                    FROM RFI.TemplateRfiSignFlow
                                    WHERE TempRfiSfId = @TempRfiSfId";
                            result = sqlConnection.QueryFirstOrDefault(sql, new { TempRfiSfId }) ?? throw new SystemException("流程資料錯誤!");
                            #endregion
                        }

                        #region //先將主資料排序改為-1
                        sql = @"UPDATE RFI.TemplateRfiSignFlow SET
                                SortNumber = SortNumber * -1,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ProTypeGroupId = @ProTypeGroupId
                                AND DepartmentId = @DepartmentId";
                        rowsAffected += sqlConnection.Execute(sql, new { ProTypeGroupId, DepartmentId, LastModifiedDate, LastModifiedBy });
                        #endregion

                        #region //更新順序
                        int sort = 1;
                        foreach (var TempRfiSfId in SortArr)
                        {
                            sql = @"UPDATE RFI.TemplateRfiSignFlow SET
                                    SortNumber = @SortNumber,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE TempRfiSfId = @TempRfiSfId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    SortNumber = sort,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    TempRfiSfId
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            sort++;
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

        #region //UpdateTemplateDesignSignFlowSort -- 更新設計申請單流程簽核流程排序 -- Shintokuro 2024.05.24
        public string UpdateTemplateDesignSignFlowSort(int ProTypeGroupId, int DepartmentId, string SortList)
        {
            try
            {
                if (SortList.Length <= 0) throw new SystemException("【順序】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //取得群組
                        sql = @"SELECT TOP 1 1
                                FROM SCM.ProductTypeGroup
                                WHERE ProTypeGroupId = @ProTypeGroupId";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { ProTypeGroupId }) ?? throw new SystemException("【產品群組】資料錯誤!");
                        #endregion

                        #region //取得部門
                        sql = @"SELECT TOP 1 1 
                                FROM BAS.Department WHERE DepartmentId = @DepartmentId";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { DepartmentId }) ?? throw new SystemException("【部門】資料錯誤!");
                        #endregion

                        string[] SortArr = SortList.Split(',');

                        foreach (var TempDesignSfId in SortArr)
                        {
                            #region //判斷主資料是否錯誤
                            sql = @"SELECT TOP 1 1
                                    FROM RFI.TemplateDesignSignFlow
                                    WHERE TempDesignSfId = @TempDesignSfId";
                            result = sqlConnection.QueryFirstOrDefault(sql, new { TempDesignSfId }) ?? throw new SystemException("流程資料錯誤!");
                            #endregion
                        }

                        #region //先將主資料排序改為-1
                        sql = @"UPDATE RFI.TemplateDesignSignFlow SET
                                SortNumber = SortNumber * -1,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ProTypeGroupId = @ProTypeGroupId
                                AND DepartmentId = @DepartmentId";
                        rowsAffected += sqlConnection.Execute(sql, new { ProTypeGroupId, DepartmentId, LastModifiedDate, LastModifiedBy });
                        #endregion

                        #region //更新順序
                        int sort = 1;
                        foreach (var TempDesignSfId in SortArr)
                        {
                            sql = @"UPDATE RFI.TemplateDesignSignFlow SET
                                    SortNumber = @SortNumber,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE TempDesignSfId = @TempDesignSfId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    SortNumber = sort,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    TempDesignSfId
                                });

                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            sort++;
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

        #region //UpdateProductTypeStatus -- 更新產品類別狀態資料 -- Shintokuro 2024.05.24
        public string UpdateProductTypeStatus(int ProTypeId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷產品類別是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM SCM.ProductType
                                WHERE 1 = 1
                                AND ProTypeId = @ProTypeId";
                        dynamicParameters.Add("ProTypeId", ProTypeId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【產品類別】找不到!請重新確認!");

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
                        sql = @"UPDATE SCM.ProductType SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE 1 = 1
                                AND ProTypeId = @ProTypeId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                CompanyId = CurrentCompany,
                                ProTypeId
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

        #region //UpdateQuotationTag -- 更新報價單屬性標籤 -- Shinotokuro 2024.08.23
        public string UpdateQuotationTag(int QtId, string TagNo, string TagName)
        {
            try
            {
                if (TagNo.Length <= 0) throw new SystemException("【標籤代號】不能為空!");
                if (TagNo.Length > 20) throw new SystemException("【標籤代號】長度錯誤!");
                if (TagName.Length <= 0) throw new SystemException("【標籤名稱】不能為空!");
                if (TagName.Length > 20) throw new SystemException("【標籤名稱】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷標籤資訊是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.QuotationTag
                                WHERE QtId = @QtId";
                        dynamicParameters.Add("QtId", QtId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【標籤】資料找不到,請重新確認!");
                        #endregion

                        #region //判斷標籤代號是否重複
                        sql = @"SELECT TOP 1 1
                                FROM SCM.QuotationTag
                                WHERE TagNo = @TagNo
                                AND QtId != @QtId";
                        dynamicParameters.Add("TagNo", TagNo);
                        dynamicParameters.Add("QtId", QtId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【標籤代號】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.QuotationTag SET
                                TagNo = @TagNo,
                                TagName = @TagName,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE QtId = @QtId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TagNo,
                                TagName,
                                LastModifiedDate,
                                LastModifiedBy,
                                QtId
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

        #region //UpdateQuotationTagStatus -- 更新報價單屬性標籤狀態 -- Shintokuro 2024.08.23
        public string UpdateQuotationTagStatus(int QtId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷標籤是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM SCM.QuotationTag
                                WHERE 1 = 1
                                AND QtId = @QtId";
                        dynamicParameters.Add("QtId", QtId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【標籤】找不到!請重新確認!");

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
                        sql = @"UPDATE SCM.QuotationTag SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE 1 = 1
                                AND QtId = @QtId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                QtId
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

        #region //UpdateSupplierMachineStatus -- 更新供應商機台資訊狀態 -- Andrew 2024.10.17
        public string UpdateSupplierMachineStatus(int SmId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷供應商機台資訊是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM SCM.SupplierMachine
                                WHERE 1 = 1
                                AND SmId = @SmId";
                        dynamicParameters.Add("SmId", SmId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【供應商機台】資料找不到!請重新確認!");

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
                        sql = @"UPDATE SCM.SupplierMachine SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE 1 = 1
                                AND SmId = @SmId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                SmId
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

        #region //UpdateSupplierProcessEquipment -- 更新參與托外掃碼的供應商的製程與機台清單 -- Andrew 2024.10.17
        public string UpdateSupplierProcessEquipment(int SmId, int SupplierId, int ProcessId, int MachineId, string Status)
        {
            try
            {
                // if (!Regex.IsMatch(Status, "^(A|S)$", RegexOptions.IgnoreCase)) throw new SystemException("【必要過站】錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();

                        #region //判斷製程 + 機台資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.SupplierMachine
                                WHERE SmId = @SmId";
                        dynamicParameters.Add("SmId", SmId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("製程 + 機台資料錯誤!");
                        #endregion

                        #region //判斷製程資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Process
                                WHERE ProcessId = @ProcessId 
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("ProcessId", ProcessId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("製程資料錯誤!");
                        #endregion

                        #region //判斷機台和車間是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.Machine a
                                INNER JOIN MES.WorkShop b ON a.ShopId = b.ShopId
                                WHERE MachineId = @MachineId
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("MachineId", MachineId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("車間和機台資料錯誤!");
                        #endregion

                        #region //更新SQL
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.SupplierMachine SET
                                SupplierId = @SupplierId,
                                ProcessId = @ProcessId,
                                MachineId = @MachineId,
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SmId = @SmId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SupplierId,
                                ProcessId,
                                MachineId,
                                Status,
                                LastModifiedDate,
                                LastModifiedBy,
                                SmId
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

        #region //UpdateProject -- 更新專案 -- Ann 2025-04-24
        public string UpdateProject(int ProjectId, string ProjectNo, string ProjectName, string EffectiveDate, string ExpirationDate, string Remark)
        {
            try
            {
                int rowsAffected = 0;
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

                        string ErpNo = "";
                        foreach (var item in companyResult)
                        {
                            ErpNo = item.ErpNo;
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        }
                        #endregion

                        #region //確認使用者資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT UserNo
                                FROM BAS.[User] 
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", CurrentUser);

                        var UserResult = sqlConnection.Query(sql, dynamicParameters);

                        if (UserResult.Count() <= 0) throw new SystemException("【使用者】資料錯誤!!");

                        string UserNo = UserResult.FirstOrDefault().UserNo;
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            if (ProjectNo.Length < 0) throw new SystemException("【專案代號】不能為空!");
                            if (ProjectName.Length < 0) throw new SystemException("【專案代號】不能為空!");

                            #region //確認專案資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.ProjectNo
                                    FROM SCM.Project a 
                                    WHERE a.ProjectId = @ProjectId";
                            dynamicParameters.Add("ProjectId", ProjectId);

                            var ProjectResult = sqlConnection.Query(sql, dynamicParameters);

                            if (ProjectResult.Count() <= 0) throw new SystemException("專案資料錯誤!!");

                            string OriProjectNo = "";
                            foreach (var item in ProjectResult)
                            {
                                OriProjectNo = item.ProjectNo;
                            }
                            #endregion

                            #region //確認專案預算資料狀態
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM SCM.ProjectDetail   
                                    WHERE ProjectId = @ProjectId
                                    AND Status != 'N'";
                            dynamicParameters.Add("ProjectId", ProjectId);

                            var ProjectDetailStatusResult = sqlConnection.Query(sql, dynamicParameters);

                            if (ProjectDetailStatusResult.Count() > 0)
                            {
                                if (OriProjectNo != ProjectNo)
                                {
                                    throw new SystemException("此專案已有預算項目進行簽核，無法更改【專案代號】!!");
                                }
                            }
                            #endregion

                            #region //修改MES段
                            #region //確認是否有相同公司別及專案代號
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM SCM.Project a 
                                    WHERE a.CompanyId = @CompanyId
                                    AND a.ProjectNo = @ProjectNo
                                    AND a.ProjectId != @ProjectId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("ProjectNo", ProjectNo);
                            dynamicParameters.Add("ProjectId", ProjectId);

                            var CheckProjectResult = sqlConnection.Query(sql, dynamicParameters);

                            if (CheckProjectResult.Count() > 0) throw new SystemException("已有重複的專案代號!!");
                            #endregion

                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.Project SET
                                    ProjectNo = @ProjectNo,
                                    ProjectName = @ProjectName,
                                    Remark = @Remark,
                                    EffectiveDate = @EffectiveDate,
                                    ExpirationDate = @ExpirationDate,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE ProjectId = @ProjectId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    ProjectNo,
                                    ProjectName,
                                    Remark,
                                    EffectiveDate,
                                    ExpirationDate,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    ProjectId
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //新增ERP段
                            #region //確認是否有相同公司別及專案代號
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM CMSNB a 
                                    WHERE a.NB001 = @ProjectNo
                                    AND a.NB001 != @OriProjectNo";
                            dynamicParameters.Add("ProjectNo", ProjectNo);
                            dynamicParameters.Add("OriProjectNo", OriProjectNo);

                            var CMSNBResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (CMSNBResult.Count() > 0) throw new SystemException("ERP已有重複的專案代號!!");
                            #endregion

                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE CMSNB SET
                                    MODIFIER = @MODIFIER,
                                    MODI_DATE = @MODI_DATE,
                                    FLAG = FLAG + 1,
                                    MODI_TIME = @MODI_TIME,
                                    MODI_AP = @MODI_AP,
                                    MODI_PRID = @MODI_PRID,
                                    NB001 = @ProjectNo,
                                    NB002 = @ProjectName,
                                    NB003 = @Remark,
                                    NB004 = @EffectiveDate,
                                    NB005 = @ExpirationDate
                                    WHERE NB001 = @OriProjectNo";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MODIFIER = UserNo,
                                    MODI_DATE = DateTime.Now.ToString("yyyyMMdd"),
                                    MODI_TIME = DateTime.Now.ToString("HH:mm:ss"),
                                    MODI_AP = UserNo + "PC",
                                    MODI_PRID = "BM",
                                    ProjectNo,
                                    ProjectName,
                                    Remark,
                                    EffectiveDate = EffectiveDate.Length > 0 ? DateTime.Parse(EffectiveDate).ToString("yyyyMMdd") : "",
                                    ExpirationDate = ExpirationDate.Length > 0 ? DateTime.Parse(ExpirationDate).ToString("yyyyMMdd") : "",
                                    OriProjectNo
                                });

                            rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);
                            #endregion

                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                msg = "(" + rowsAffected + " rows affected)",
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

        #region //UpdateProjectDetail -- 更新專案預算資料 -- Ann 2025-04-24
        public string UpdateProjectDetail(int ProjectDetailId, string ProjectType, string Currency, double ExchangeRate, double BudgetAmount, double LocalBudgetAmount, string Remark, string ProjectFile)
        {
            try
            {
                int rowsAffected = 0;
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

                        string ErpNo = "";
                        foreach (var item in companyResult)
                        {
                            ErpNo = item.ErpNo;
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        }
                        #endregion

                        #region //確認使用者資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT UserNo
                                FROM BAS.[User] 
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", CurrentUser);

                        var UserResult = sqlConnection.Query(sql, dynamicParameters);

                        if (UserResult.Count() <= 0) throw new SystemException("【使用者】資料錯誤!!");

                        string UserNo = UserResult.FirstOrDefault().UserNo;
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            if (ProjectType.Length <= 0) throw new SystemException("【專案代號】不能為空!");
                            if (BudgetAmount <= 0) throw new SystemException("【預算金額】不能為空!");

                            #region //確認專案預算資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.ProjectId, a.[Status]
                                    FROM SCM.ProjectDetail a 
                                    WHERE a.ProjectDetailId = @ProjectDetailId";
                            dynamicParameters.Add("ProjectDetailId", ProjectDetailId);

                            var ProjectDetailResult = sqlConnection.Query(sql, dynamicParameters);

                            if (ProjectDetailResult.Count() <= 0) throw new SystemException("專案預算資料錯誤!!");

                            int ProjectId = -1;
                            foreach (var item in ProjectDetailResult)
                            {
                                ProjectId = item.ProjectId;

                                if (item.Status != "N" && item.Status != "E")
                                {
                                    throw new SystemException("此專案預算已簽核，無法更改!!");
                                }
                            }
                            #endregion

                            #region //確認是否已經有相同專案類型資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM SCM.ProjectDetail a 
                                    WHERE a.ProjectId = @ProjectId 
                                    AND a.ProjectType = @ProjectType
                                    AND a.ProjectDetailId != @ProjectDetailId";
                            dynamicParameters.Add("ProjectId", ProjectId);
                            dynamicParameters.Add("ProjectType", ProjectType);
                            dynamicParameters.Add("ProjectDetailId", ProjectDetailId);

                            var CheckProjectDetailResult = sqlConnection.Query(sql, dynamicParameters);

                            if (CheckProjectDetailResult.Count() > 0) throw new SystemException("此專案已建立相同專案類型預算!!");
                            #endregion

                            #region //確認幣別資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM CMSMF
                                    WHERE MF001 = @Currency";
                            dynamicParameters.Add("Currency", Currency);

                            var CMSMFResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (CMSMFResult.Count() <= 0) throw new SystemException("【幣別】資料錯誤!!");
                            #endregion

                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.ProjectDetail SET
                                    ProjectType = @ProjectType,
                                    Currency = @Currency,
                                    ExchangeRate = @ExchangeRate,
                                    BudgetAmount = @BudgetAmount,
                                    LocalBudgetAmount = @LocalBudgetAmount,
                                    Remark = @Remark,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE ProjectDetailId = @ProjectDetailId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    ProjectType,
                                    Currency,
                                    ExchangeRate,
                                    BudgetAmount,
                                    LocalBudgetAmount,
                                    Remark,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    ProjectDetailId
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                            #region //附檔
                            #region //先將原本的砍掉
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE FROM SCM.ProjectFile
                                    WHERE ProjectDetailId = @ProjectDetailId";
                            dynamicParameters.Add("ProjectDetailId", ProjectDetailId);

                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            if (ProjectFile.Length > 0)
                            {
                                string[] projectFiles = ProjectFile.Split(',');
                                foreach (var file in projectFiles)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO SCM.ProjectFile (ProjectDetailId, FileId
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.ProjectFileId
                                            VALUES (@ProjectDetailId, @FileId
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            ProjectDetailId,
                                            FileId = file,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    var fileInsertResult = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += fileInsertResult.Count();
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

        #region //UpdateProjectStatus -- 更新專案預算狀態 -- Ann 2025-04-25
        public string UpdateProjectStatus(int ProjectDetailId, string Status, string ConfirmUser)
        {
            try
            {
                int rowsAffected = 0;
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認使用者資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT UserId
                                FROM BAS.[User] 
                                WHERE UserNo = @UserNo";
                        dynamicParameters.Add("UserNo", ConfirmUser);

                        var UserResult = sqlConnection.Query(sql, dynamicParameters);

                        if (UserResult.Count() <= 0) throw new SystemException("【使用者】資料錯誤!!");

                        int UserId = UserResult.FirstOrDefault().UserId;
                        #endregion

                        #region //確認專案預算資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.[Status]
                                    FROM SCM.ProjectDetail a 
                                    WHERE a.ProjectDetailId = @ProjectDetailId";
                        dynamicParameters.Add("ProjectDetailId", ProjectDetailId);

                        var ProjectDetailResult = sqlConnection.Query(sql, dynamicParameters);

                        if (ProjectDetailResult.Count() <= 0) throw new SystemException("專案預算資料錯誤!!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.ProjectDetail SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ProjectDetailId = @ProjectDetailId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status,
                                LastModifiedDate,
                                LastModifiedBy = UserId,
                                ProjectDetailId
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
        
        #region //UpdateChangeProjectStatus -- 更新專案預算變更狀態 -- Ann 2025-04-28
        public string UpdateChangeProjectStatus(int LogId, string Status, string ConfirmUser)
        {
            try
            {
                int rowsAffected = 0;
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認使用者資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT UserId
                                FROM BAS.[User] 
                                WHERE UserNo = @UserNo";
                        dynamicParameters.Add("UserNo", ConfirmUser);

                        var UserResult = sqlConnection.Query(sql, dynamicParameters);

                        if (UserResult.Count() <= 0) throw new SystemException("【使用者】資料錯誤!!");

                        int UserId = UserResult.FirstOrDefault().UserId;
                        #endregion

                        #region //確認專案預算變更資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ProjectDetailId, a.Currency, a.ExchangeRate, a.BudgetAmount, a.LocalBudgetAmount, a.BpmStatus
                                , b.Currency OriCurrency, b.ExchangeRate OriExchangeRate, b.BudgetAmount OriBudgetAmount, b.LocalBudgetAmount OriLocalBudgetAmount, b.Edition
                                FROM SCM.ProjectBudgetChangeLog a 
                                INNER JOIN SCM.ProjectDetail b ON a.ProjectDetailId = b.ProjectDetailId
                                WHERE a.LogId = @LogId";
                        dynamicParameters.Add("LogId", LogId);

                        var ProjectBudgetChangeLogResult = sqlConnection.Query(sql, dynamicParameters);

                        if (ProjectBudgetChangeLogResult.Count() <= 0) throw new SystemException("【專案預算變更】資料錯誤!!");

                        int ProjectDetailId = -1;
                        string Currency = "";
                        double ExchangeRate = 0;
                        double BudgetAmount = 0;
                        double LocalBudgetAmount = 0;
                        string BpmStatus = "";
                        string OriCurrency = "";
                        double OriExchangeRate = 0;
                        double OriBudgetAmount = 0;
                        double OriLocalBudgetAmount = 0;
                        string Edition = "";
                        foreach (var item in ProjectBudgetChangeLogResult)
                        {
                            if (item.BpmStatus != "N") throw new SystemException("【專案預算變更】狀態錯誤!!");
                            ProjectDetailId = item.ProjectDetailId;
                            Currency = item.Currency;
                            ExchangeRate = item.ExchangeRate;
                            BudgetAmount = item.BudgetAmount;
                            LocalBudgetAmount = item.LocalBudgetAmount;
                            BpmStatus = item.BpmStatus;
                            OriCurrency = item.OriCurrency;
                            OriExchangeRate = item.OriExchangeRate;
                            OriBudgetAmount = item.OriBudgetAmount;
                            OriLocalBudgetAmount = item.OriLocalBudgetAmount;
                            Edition = item.Edition;
                        }
                        #endregion

                        #region //更新專案預算變更狀態
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.ProjectBudgetChangeLog SET
                                BpmStatus = @BpmStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE LogId = @LogId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                BpmStatus = Status == "Y" ? "Y" : "E",
                                LastModifiedDate,
                                LastModifiedBy = UserId,
                                LogId
                            });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //更新專案預算狀態
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.ProjectDetail SET
                                Currency = @Currency,
                                ExchangeRate = @ExchangeRate,
                                BudgetAmount = @BudgetAmount,
                                LocalBudgetAmount = @LocalBudgetAmount,
                                Edition = @Edition,
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ProjectDetailId = @ProjectDetailId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Currency = Status == "Y" ? Currency : OriCurrency,
                                ExchangeRate = Status == "Y" ? ExchangeRate : OriExchangeRate,
                                BudgetAmount = Status == "Y" ? BudgetAmount : OriBudgetAmount,
                                LocalBudgetAmount = Status == "Y" ? LocalBudgetAmount : OriLocalBudgetAmount,
                                Edition = Status == "Y" ? (Convert.ToInt32(Edition) + 1).ToString("D4") : Edition,
                                Status = Status == "Y" ? "Y" : "K",
                                LastModifiedDate,
                                LastModifiedBy = UserId,
                                ProjectDetailId
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

        #region //CancelChangeProject -- 取消專案預算變更 -- Ann 2025-04-28
        public string CancelChangeProject(int ProjectDetailId)
        {
            try
            {
                int rowsAffected = 0;
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認專案預算變更資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.Status
                                FROM SCM.ProjectDetail a 
                                WHERE a.ProjectDetailId = @ProjectDetailId";
                        dynamicParameters.Add("ProjectDetailId", ProjectDetailId);

                        var ProjectDetailResult = sqlConnection.Query(sql, dynamicParameters);

                        if (ProjectDetailResult.Count() <= 0) throw new SystemException("【專案預算變更】資料錯誤!!");

                        foreach (var item in ProjectDetailResult)
                        {
                            if (item.Status != "K") throw new SystemException("【專案預算】狀態錯誤!!");
                        }
                        #endregion

                        #region //確認是否有專案預算變更紀錄還在簽核中
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.ProjectBudgetChangeLog a 
                                WHERE a.ProjectDetailId = @ProjectDetailId 
                                AND a.BpmStatus = 'N'";
                        dynamicParameters.Add("ProjectDetailId", ProjectDetailId);

                        var ProjectBudgetChangeLogResult = sqlConnection.Query(sql, dynamicParameters);

                        if (ProjectBudgetChangeLogResult.Count() > 0) throw new SystemException("尚有專案預算變更還在簽核!!");
                        #endregion

                        #region //更新專案預算變更狀態
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.ProjectDetail SET
                                Status = 'Y',
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ProjectDetailId = @ProjectDetailId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                LastModifiedDate,
                                LastModifiedBy,
                                ProjectDetailId
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

        #region //CloseProject -- 專案結案 -- Ann 2025-04-28
        public string CloseProject(int ProjectId)
        {
            try
            {
                int rowsAffected = 0;
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

                        string ErpNo = "";
                        foreach (var item in companyResult)
                        {
                            ErpNo = item.ErpNo;
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        }
                        #endregion

                        #region //確認使用者資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT UserNo
                                FROM BAS.[User] 
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", CurrentUser);

                        var UserResult = sqlConnection.Query(sql, dynamicParameters);

                        if (UserResult.Count() <= 0) throw new SystemException("【使用者】資料錯誤!!");

                        string UserNo = UserResult.FirstOrDefault().UserNo;
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //確認專案資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.ProjectNo
                                    FROM SCM.Project a 
                                    WHERE a.ProjectId = @ProjectId";
                            dynamicParameters.Add("ProjectId", ProjectId);

                            var ProjectResult = sqlConnection.Query(sql, dynamicParameters);

                            if (ProjectResult.Count() <= 0) throw new SystemException("【專案】資料錯誤!!");

                            string ProjectNo = ProjectResult.FirstOrDefault().ProjectNo;
                            #endregion

                            #region //確認是否有專案預算變更紀錄還在簽核中
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM SCM.ProjectDetail a 
                                    WHERE a.ProjectId = @ProjectId 
                                    AND a.Status != 'N' AND a.Status != 'Y'";
                            dynamicParameters.Add("ProjectId", ProjectId);

                            var ProjectDetailResult = sqlConnection.Query(sql, dynamicParameters);

                            if (ProjectDetailResult.Count() > 0) throw new SystemException("尚有專案預算變更還在簽核!!");
                            #endregion

                            #region //更新MES專案預算變更狀態
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.Project SET
                                    CloseCode = 'Y',
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE ProjectId = @ProjectId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    ProjectId
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //更新ERP專案預算變更狀態
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE CMSNB SET
                                    MODIFIER = @MODIFIER,
                                    MODI_DATE = @MODI_DATE,
                                    FLAG = FLAG + 1,
                                    MODI_TIME = @MODI_TIME,
                                    MODI_AP = @MODI_AP,
                                    MODI_PRID = @MODI_PRID,
                                    NB006 = 'Y'
                                    WHERE NB001 = @NB001";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MODIFIER = UserNo,
                                    MODI_DATE = DateTime.Now.ToString("yyyyMMdd"),
                                    MODI_TIME = DateTime.Now.ToString("HH:mm:ss"),
                                    MODI_AP = UserNo + "PC",
                                    MODI_PRID = "BM",
                                    NB001 = ProjectNo
                                });

                            rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);
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

        #region //UpdateProjectManualSynchronize -- 手動同步專案資料 -- Ann 2025-05-12
        public string UpdateProjectManualSynchronize()
        {
            try
            {
                int rowsAffected = 0;
                List<Project> projects = new List<Project>();
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

                        string ErpNo = "";
                        foreach (var item in companyResult)
                        {
                            ErpNo = item.ErpNo;
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        }
                        #endregion

                        #region //確認使用者資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT UserNo
                                FROM BAS.[User] 
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", CurrentUser);

                        var UserResult = sqlConnection.Query(sql, dynamicParameters);

                        if (UserResult.Count() <= 0) throw new SystemException("【使用者】資料錯誤!!");

                        string UserNo = UserResult.FirstOrDefault().UserNo;
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //取得ERP專案資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(NB001)) ProjectNo, LTRIM(RTRIM(NB002)) ProjectName, LTRIM(RTRIM(NB003)) Remark
                                    , CASE WHEN LEN(LTRIM(RTRIM(NB004))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(NB004)) as date), 'yyyy-MM-dd') ELSE NULL END EffectiveDate
                                    , CASE WHEN LEN(LTRIM(RTRIM(NB005))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(NB005)) as date), 'yyyy-MM-dd') ELSE NULL END ExpirationDate
                                    , LTRIM(RTRIM(NB006)) CloseCode
                                    FROM CMSNB
                                    WHERE 1=1";
                            projects = sqlConnection2.Query<Project>(sql, dynamicParameters).ToList();
                            #endregion

                            #region //取得目前MES專案資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.ProjectId, a.ProjectNo
                                    FROM SCM.Project a
                                    WHERE a.CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<Project> bmProjects = sqlConnection.Query<Project>(sql, dynamicParameters).ToList();

                            projects = projects.GroupJoin(bmProjects, x => new { x.ProjectNo }, y => new { y.ProjectNo }, (x, y) => { x.ProjectId = y.FirstOrDefault()?.ProjectId; return x; }).Select(x => x).ToList();
                            #endregion

                            #region //新增/修改
                            List<Project> addProjects = projects.Where(x => x.ProjectId == null).ToList();
                            List<Project> updateProjects = projects.Where(x => x.ProjectId != null).ToList();

                            #region //新增
                            dynamicParameters = new DynamicParameters();
                            if (addProjects.Count > 0)
                            {
                                addProjects
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.CompanyId = CurrentCompany;
                                        x.TransferErpStatus = "Y";
                                        x.CreateDate = CreateDate;
                                        x.LastModifiedDate = LastModifiedDate;
                                        x.CreateBy = CreateBy;
                                        x.LastModifiedBy = LastModifiedBy;
                                    });

                                sql = @"INSERT INTO SCM.Project (CompanyId, ProjectNo, ProjectName, Remark
                                        , EffectiveDate, ExpirationDate, CloseCode, TransferErpStatus, TransferErpDate
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        VALUES (@CompanyId, @ProjectNo, @ProjectName, @Remark
                                        , @EffectiveDate, @ExpirationDate, @CloseCode, @TransferErpStatus, @TransferErpDate
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                rowsAffected += sqlConnection.Execute(sql, addProjects);
                            }
                            #endregion

                            #region //修改
                            dynamicParameters = new DynamicParameters();
                            if (updateProjects.Count > 0)
                            {
                                updateProjects
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.LastModifiedDate = LastModifiedDate;
                                        x.LastModifiedBy = LastModifiedBy;
                                    });

                                sql = @"UPDATE SCM.Project SET
                                        ProjectNo = @ProjectNo,
                                        ProjectName = @ProjectName,
                                        Remark = @Remark,
                                        EffectiveDate = @EffectiveDate,
                                        ExpirationDate = @ExpirationDate,
                                        CloseCode = @CloseCode,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE ProjectId = @ProjectId";
                                rowsAffected += sqlConnection.Execute(sql, updateProjects);
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
        #region //DeletePackingVolume -- 包材體積資料刪除 -- Zoey 2022.06.10
        public string DeletePackingVolume(int VolumeId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷體積資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.PackingVolume
                                WHERE VolumeId = @VolumeId";
                        dynamicParameters.Add("VolumeId", VolumeId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("體積資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.PackingVolume
                                WHERE VolumeId = @VolumeId";
                        dynamicParameters.Add("VolumeId", VolumeId);

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

        #region //DeletePackingWeight -- 包材重量資料刪除 -- Zoey 2022.06.10
        public string DeletePackingWeight(int WeightId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷重量資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.PackingWeight
                                WHERE WeightId = @WeightId";
                        dynamicParameters.Add("WeightId", WeightId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("重量資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.PackingWeight
                                WHERE WeightId = @WeightId";
                        dynamicParameters.Add("WeightId", WeightId);

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

        #region //DeleteItemWeight -- 物件重量資料刪除 -- Zoey 2022.06.13
        public string DeleteItemWeight(int ItemDefaultWeightId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷物件重量資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.ItemDefaultWeight
                                WHERE ItemDefaultWeightId = @ItemDefaultWeightId";
                        dynamicParameters.Add("ItemDefaultWeightId", ItemDefaultWeightId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("物件重量資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.ItemDefaultWeight
                                WHERE ItemDefaultWeightId = @ItemDefaultWeightId";
                        dynamicParameters.Add("ItemDefaultWeightId", ItemDefaultWeightId);

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

        #region //DeleteForwarder -- 貨運承攬商資料刪除 -- Zoey 2022.07.01
        public string DeleteForwarder(int ForwarderId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷貨運承攬商資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.Forwarder
                                WHERE ForwarderId = @ForwarderId";
                        dynamicParameters.Add("ForwarderId", ForwarderId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("貨運承攬商資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.Forwarder
                                WHERE ForwarderId = @ForwarderId";
                        dynamicParameters.Add("ForwarderId", ForwarderId);

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

        #region //DeleteDeliveryCustomer -- 送貨客戶資料刪除 -- Zoey 2022.07.01
        public string DeleteDeliveryCustomer(int DcId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷送貨客戶資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.DeliveryCustomer
                                WHERE DcId = @DcId";
                        dynamicParameters.Add("DcId", DcId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("送貨客戶資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.DeliveryCustomer
                                WHERE DcId = @DcId";
                        dynamicParameters.Add("DcId", DcId);

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

        #region //DeletePacking -- 包材資料刪除 -- Zoey 2022.10.11
        public string DeletePacking(int PackingId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷包材資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.Packing
                                WHERE PackingId = @PackingId";
                        dynamicParameters.Add("PackingId", PackingId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("包材資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.Packing
                                WHERE PackingId = @PackingId";
                        dynamicParameters.Add("PackingId", PackingId);

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

        #region //DeleteRfqProductClass -- RFQ產品類型刪除 -- Chia Yuan 2023.06.30
        public string DeleteRfqProductClass(int RfqProClassId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷RFQ產品類型資料是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM SCM.RfqProductClass
                                WHERE RfqProClassId = @RfqProClassId";
                        dynamicParameters.Add("RfqProClassId", RfqProClassId);

                        var result = sqlConnection.Query(sql, dynamicParameters).FirstOrDefault();
                        if (result == null) throw new SystemException("RFQ產品類型資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.RfqProductClass
                                WHERE RfqProClassId = @RfqProClassId";
                        dynamicParameters.Add("RfqProClassId", RfqProClassId);

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

        #region //DeleteRfqProductType -- RFQ產品類別刪除 -- Chia Yuan 2023.06.30
        public string DeleteRfqProductType(int RfqProTypeId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷RFQ產品類別資料是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM SCM.RfqProductType
                                WHERE RfqProTypeId = @RfqProTypeId";
                        dynamicParameters.Add("RfqProTypeId", RfqProTypeId);

                        var result = sqlConnection.Query(sql, dynamicParameters).FirstOrDefault();
                        if (result == null) throw new SystemException("RFQ產品類別資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.RfqProductType
                                WHERE RfqProTypeId = @RfqProTypeId";
                        dynamicParameters.Add("RfqProTypeId", RfqProTypeId);

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

        #region //DeleteRfqPackageType -- RFQ包裝種類刪除 -- Chia Yuan 2023.07.03
        public string DeleteRfqPackageType(int RfqPkTypeId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷RFQ包裝種類資料是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM SCM.RfqPackageType
                                WHERE RfqPkTypeId = @RfqPkTypeId";
                        dynamicParameters.Add("RfqPkTypeId", RfqPkTypeId);

                        var result = sqlConnection.Query(sql, dynamicParameters).FirstOrDefault();
                        if (result == null) throw new SystemException("RFQ包裝種類資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.RfqPackageType
                                WHERE RfqPkTypeId = @RfqPkTypeId";
                        dynamicParameters.Add("RfqPkTypeId", RfqPkTypeId);

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

        #region //DeleteProductUse -- RFQ產品用途資料刪除 -- Yi 2023.07.12
        public string DeleteProductUse(int ProductUseId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷產品用途資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.ProductUse
                                WHERE ProductUseId = @ProductUseId";
                        dynamicParameters.Add("ProductUseId", ProductUseId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("產品用途資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.ProductUse
                                WHERE ProductUseId = @ProductUseId";
                        dynamicParameters.Add("ProductUseId", ProductUseId);

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

        #region //DeletePerformanceGoals -- 業務績效目標刪除 -- Shintokuro 2023.11.29
        public string DeletePerformanceGoals(int PgId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷業務績效目標是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.PerformanceGoals
                                WHERE PgId = @PgId";
                        dynamicParameters.Add("PgId", PgId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("業務績效目標資料不存在,請重新確認!!");
                        #endregion

                        #region //判斷業務績效目標單身是否存在
                        sql = @"SELECT TOP 1 1
                                FROM SCM.PgDetail a
                                INNER JOIN SCM.PerformanceGoals b on a.PgId = b.PgId
                                WHERE a.PgId = @PgId
                                AND a.ConfirmStatus 'Y'";
                        dynamicParameters.Add("PgId", PgId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("業務績效目標已經有單身處於確認狀態,無法執行刪除!!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除次要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.PgDetail
                                WHERE PgId = @PgId";
                        dynamicParameters.Add("PgId", PgId);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.PerformanceGoals
                                WHERE PgId = @PgId";
                        dynamicParameters.Add("PgId", PgId);

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

        #region //DeletePgDetail -- 業務績效目標單身刪除 -- Shintokuro 2023.11.29
        public string DeletePgDetail(int PgDetailId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷業務績效目標單身是否正確
                        sql = @"SELECT TOP 1 ConfirmStatus
                                FROM SCM.PgDetail
                                WHERE PgDetailId = @PgDetailId";
                        dynamicParameters.Add("PgDetailId", PgDetailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("業務績效目標單身資料不存在,請重新確認!");
                        foreach(var item in result)
                        {
                            if(item.ConfirmStatus != "N") throw new SystemException("業務績效目標單身確認狀態必須處於未確認才能執行刪除!");
                        }
                        
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.PgDetail
                                WHERE PgDetailId = @PgDetailId";
                        dynamicParameters.Add("PgDetailId", PgDetailId);

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

        #region //DeleteProductTypeGroup -- 刪除產品群組資料 -- Shintokuro 2024.05.24
        public string DeleteProductTypeGroup(int ProTypeGroupId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷產品群組是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.ProductTypeGroup
                                WHERE ProTypeGroupId = @ProTypeGroupId";
                        dynamicParameters.Add("ProTypeGroupId", ProTypeGroupId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("產品群組資料找不到,請重新確認!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.ProductTypeGroup
                                WHERE ProTypeGroupId = @ProTypeGroupId";
                        dynamicParameters.Add("ProTypeGroupId", ProTypeGroupId);

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

        #region //DeleteProductType -- 刪除產品類別資料 -- Shintokuro 2024.05.24
        public string DeleteProductType(int ProTypeId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷產品類別是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.ProductType
                                WHERE ProTypeId = @ProTypeId";
                        dynamicParameters.Add("ProTypeId", ProTypeId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("產品類別資料找不到,請重新確認!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.ProductType
                                WHERE ProTypeId = @ProTypeId";
                        dynamicParameters.Add("ProTypeId", ProTypeId);

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

        #region //DeleteTemplateRfiSignFlow -- 刪除市場評估單簽核流程 -- Shintokuro 2024.05.27
        public string DeleteTemplateRfiSignFlow(int TempRfiSfId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷主資料是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 ProTypeGroupId
                                FROM RFI.TemplateRfiSignFlow
                                WHERE TempRfiSfId = @TempRfiSfId";
                        dynamicParameters.Add("TempRfiSfId", TempRfiSfId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("關卡資料錯誤,請重新確認!");
                        int proTypeGroupId = -1;
                        foreach (var item in result)
                        {
                            proTypeGroupId = item.ProTypeGroupId;
                        }
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE RFI.TemplateRfiSignFlow
                                WHERE TempRfiSfId = @TempRfiSfId";
                        dynamicParameters.Add("TempRfiSfId", TempRfiSfId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //取得主資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.TempRfiSfId
                                FROM RFI.TemplateRfiSignFlow a
                                WHERE a.ProTypeGroupId = @ProTypeGroupId";
                        dynamicParameters.Add("ProTypeGroupId", proTypeGroupId);

                        var SortArr = sqlConnection.Query(sql, dynamicParameters);
                        #endregion

                        #region //先將主資料排序改為-1
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE RFI.TemplateRfiSignFlow SET
                                SortNumber = SortNumber * -1,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ProTypeGroupId = @ProTypeGroupId";
                        dynamicParameters.Add("LastModifiedDate", LastModifiedDate);
                        dynamicParameters.Add("LastModifiedBy", LastModifiedBy);
                        dynamicParameters.Add("ProTypeGroupId", proTypeGroupId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //更新順序
                        int sort = 1;
                        foreach (var item in SortArr)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE RFI.TemplateRfiSignFlow SET
                                    SortNumber = @SortNumber,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE TempRfiSfId = @TempRfiSfId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    SortNumber = sort,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    item.TempRfiSfId
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            sort++;
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

        #region //DeleteTemplateDesignSignFlow -- 刪除設計申請單流程簽核流程 -- Shintokuro 2024.05.27
        public string DeleteTemplateDesignSignFlow(int TempDesignSfId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷主資料是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 ProTypeGroupId
                                FROM RFI.TemplateDesignSignFlow
                                WHERE TempDesignSfId = @TempDesignSfId";
                        dynamicParameters.Add("TempDesignSfId", TempDesignSfId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("關卡資料錯誤,請重新確認!");
                        int proTypeGroupId = -1;
                        foreach (var item in result)
                        {
                            proTypeGroupId = item.ProTypeGroupId;
                        }
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE RFI.TemplateDesignSignFlow
                                WHERE TempDesignSfId = @TempDesignSfId";
                        dynamicParameters.Add("TempDesignSfId", TempDesignSfId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //取得主資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.TempDesignSfId
                                FROM RFI.TemplateDesignSignFlow a
                                WHERE a.ProTypeGroupId = @ProTypeGroupId";
                        dynamicParameters.Add("ProTypeGroupId", proTypeGroupId);

                        var SortArr = sqlConnection.Query(sql, dynamicParameters);
                        #endregion

                        #region //先將主資料排序改為-1
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE RFI.TemplateDesignSignFlow SET
                                SortNumber = SortNumber * -1,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ProTypeGroupId = @ProTypeGroupId";
                        dynamicParameters.Add("LastModifiedDate", LastModifiedDate);
                        dynamicParameters.Add("LastModifiedBy", LastModifiedBy);
                        dynamicParameters.Add("ProTypeGroupId", proTypeGroupId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //更新順序
                        int sort = 1;
                        foreach (var item in SortArr)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE RFI.TemplateDesignSignFlow SET
                                    SortNumber = @SortNumber,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE TempDesignSfId = @TempDesignSfId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    SortNumber = sort,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    item.TempDesignSfId
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            sort++;
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

        #region //DeleteQuotationTag -- 刪除報價單屬性標籤狀態 -- Shintokuro 2024.08.23
        public string DeleteQuotationTag(int QtId)
        {
            try
            {

                int rowsAffected = 0;

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷標籤資訊是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.QuotationTag
                                WHERE QtId = @QtId";
                        dynamicParameters.Add("QtId", QtId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【標籤】資料找不到,請重新確認!");
                        #endregion

                        #region //判斷標籤資訊是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.QuotationTag
                                WHERE QtId = @QtId";
                        dynamicParameters.Add("QtId", QtId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【標籤】資料找不到,請重新確認!");
                        #endregion


                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.QuotationTag
                                WHERE QtId = @QtId";
                        dynamicParameters.Add("QtId", QtId);

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

        #region //DeleteSupplierProcessEquipment -- 參與托外掃碼的供應商的製程相關資料新增 -- Andrew 2024.10.17
        public string DeleteSupplierProcessEquipment(int SmId)
        {
            try
            {

                int rowsAffected = 0;

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷供應商機台資訊是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.SupplierMachine
                                WHERE SmId = @SmId";
                        dynamicParameters.Add("SmId", SmId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【供應商機台】資料找不到,請重新確認!");
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.SupplierMachine
                                WHERE SmId = @SmId";
                        dynamicParameters.Add("SmId", SmId);

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
        #endregion

        #region //BPM相關
        #region //TransferProjectToBpm -- 拋轉專案預算資料到BPM -- Ann 2025-04-25
        public string TransferProjectToBpm(int ProjectDetailId)
        {
            try
            {
                int rowsAffected = 0;
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

                        string ErpNo = "";
                        foreach (var item in companyResult)
                        {
                            ErpNo = item.ErpNo;
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        }
                        #endregion

                        #region //確認使用者資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT UserNo
                                FROM BAS.[User] 
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", CurrentUser);

                        var UserResult = sqlConnection.Query(sql, dynamicParameters);

                        if (UserResult.Count() <= 0) throw new SystemException("【使用者】資料錯誤!!");

                        string UserNo = UserResult.FirstOrDefault().UserNo;
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //確認專案預算資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.ProjectId, a.[Status]
                                    FROM SCM.ProjectDetail a 
                                    WHERE a.ProjectDetailId = @ProjectDetailId";
                            dynamicParameters.Add("ProjectDetailId", ProjectDetailId);

                            var ProjectDetailResult = sqlConnection.Query(sql, dynamicParameters);

                            if (ProjectDetailResult.Count() <= 0) throw new SystemException("專案預算資料錯誤!!");

                            int ProjectId = -1;
                            foreach (var item in ProjectDetailResult)
                            {
                                ProjectId = item.ProjectId;

                                if (item.Status != "N" && item.Status != "E")
                                {
                                    throw new SystemException("此專案預算已簽核，無法重複拋轉!!");
                                }
                            }
                            #endregion

                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.ProjectDetail SET
                                    BpmTransferStatus = 'Y',
                                    Status = 'P',
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE ProjectDetailId = @ProjectDetailId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    ProjectDetailId
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
