using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SCMDA
{
    #region //公司欄位
    public class Company
    {
        public int? CompanyId { get; set; }
        public string CompanyNo { get; set; }
        public string ErpNo { get; set; }
    }
    #endregion

    #region //部門欄位
    public class Department
    {
        public int? DepartmentId { get; set; }
        public string DepartmentNo { get; set; }
    }
    #endregion

    #region //使用者欄位
    public class User
    {
        public int? UserId { get; set; }
        public string UserNo { get; set; }
        public int? DepartmentId { get; set; }
    }
    #endregion

    #region //單位欄位
    public class Uom
    {
        public int? UomId { get; set; }
        public string UomNo { get; set; }
    }
    #endregion

    #region //匯率欄位
    public class ExchangeRate
    {
        public int? ExchangeRateId { get; set; }
        public int? CompanyId { get; set; }
        public string Currency { get; set; }
        public DateTime? EffectiveDate { get; set; }
        public string BankBuyingRate { get; set; }
        public string BankSellingRate { get; set; }
        public string CustomsBuyingRate { get; set; }
        public string CustomsSellingRate { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
    }
    #endregion

    #region //品號欄位
    public class MtlItem
    {
        public int? MtlItemId { get; set; }
        public string MtlItemNo { get; set; }
        public string MtlItemName { get; set; }
        public string MtlItemSpec { get; set; }
        public int SaleUomId { get; set; }
    }
    #endregion

    #region //客戶欄位
    public class Customer
    {
        public int? CustomerId { get; set; }
        public int? CompanyId { get; set; }
        public int? DepartmentId { get; set; }
        public string DepartmentNo { get; set; }
        public string CustomerNo { get; set; }
        public string CustomerName { get; set; }
        public string CustomerEnglishName { get; set; }
        public string CustomerShortName { get; set; }
        public string Country { get; set; }
        public string Region { get; set; }
        public string Route { get; set; }
        public string GuiNumber { get; set; }
        public string ResponsiblePerson { get; set; }
        public string Contact { get; set; }
        public string TelNoFirst { get; set; }
        public string TelNoSecond { get; set; }
        public string FaxNo { get; set; }
        public string Email { get; set; }
        public string RegisterAddressFirst { get; set; }
        public string RegisterAddressSecond { get; set; }
        public string InvoiceAddressFirst { get; set; }
        public string InvoiceAddressSecond { get; set; }
        public string DeliveryAddressFirst { get; set; }
        public string DeliveryAddressSecond { get; set; }
        public string DocumentAddressFirst { get; set; }
        public string DocumentAddressSecond { get; set; }
        public string BillAddressFirst { get; set; }
        public string BillAddressSecond { get; set; }
        public string ZipCodeRegister { get; set; }
        public string ZipCodeInvoice { get; set; }
        public string ZipCodeDelivery { get; set; }
        public string ZipCodeDocument { get; set; }
        public string ZipCodeBill { get; set; }
        public string BillReceipient { get; set; }
        public DateTime? InauguateDate { get; set; }
        public string AccountDay { get; set; }
        public DateTime? CloseDate { get; set; }
        public DateTime? PermitDate { get; set; }
        public string Version { get; set; }
        public float? Capital { get; set; }
        public int? Headcount { get; set; }
        public string HomeOffice { get; set; }
        public string AnnualTurnover { get; set; }
        public string Currency { get; set; }
        public string TradeTerm { get; set; }
        public string PaymentTerm { get; set; }
        public string PricingType { get; set; }
        public string PaymentType { get; set; }
        public string ReceiptReceive { get; set; }
        public string TaxAmountCalculateType { get; set; }
        public string DocumentDeliver { get; set; }
        public string InvoiceCount { get; set; }
        public string TaxNo { get; set; }
        public string Taxation { get; set; }
        public string PaymentBankFirst { get; set; }
        public string PaymentBankSecond { get; set; }
        public string PaymentBankThird { get; set; }
        public string BankAccountFirst { get; set; }
        public string BankAccountSecond { get; set; }
        public string BankAccountThird { get; set; }
        public string Account { get; set; }
        public string AccountInvoice { get; set; }
        public string ClearanceType { get; set; }
        public string ShipMethod { get; set; }
        public string InvocieAttachedStatus { get; set; }
        public string CustomerKind { get; set; }
        public string UploadType { get; set; }
        public float? DepositRate { get; set; }
        public string SaleRating { get; set; }
        public string CreditRating { get; set; }
        public float? CreditLimit { get; set; }
        public string CreditLimitControl { get; set; }
        public string CreditLimitControlCurrency { get; set; }
        public string SoCreditAuditType { get; set; }
        public string SiCreditAuditType { get; set; }
        public string DoCreditAuditType { get; set; }
        public string InTransitCreditAuditType { get; set; }
        public string RelatedPerson { get; set; }
        public int? SalesmenId { get; set; }
        public string SalesmenNo { get; set; }
        public int? PaymentSalesmenId { get; set; }
        public string PaymentSalesmenNo { get; set; }
        public string ShipType { get; set; }
        public int? ForwarderId { get; set; }
        public string CustomerRemark { get; set; }
        public string TransferStatus { get; set; }
        public DateTime? TransferDate { get; set; }
        public string Status { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
    }
    #endregion

    #region //供應商欄位
    public class Supplier
    {
        public int? SupplierId { get; set; }
        public int? CompanyId { get; set; }
        public string SupplierNo { get; set; }
        public string SupplierName { get; set; }
        public string SupplierShortName { get; set; }
        public string SupplierEnglishName { get; set; }
        public string Country { get; set; }
        public string Region { get; set; }
        public string GuiNumber { get; set; }
        public string ResponsiblePerson { get; set; }
        public string ContactFirst { get; set; }
        public string ContactSecond { get; set; }
        public string ContactThird { get; set; }
        public string TelNoFirst { get; set; }
        public string TelNoSecond { get; set; }
        public string FaxNo { get; set; }
        public string FaxNoAccounting { get; set; }
        public string Email { get; set; }
        public string AddressFirst { get; set; }
        public string AddressSecond { get; set; }
        public string ZipCodeFirst { get; set; }
        public string ZipCodeSecond { get; set; }
        public string BillAddressFirst { get; set; }
        public string BillAddressSecond { get; set; }
        public string PermitStatus { get; set; }
        public DateTime? InauguateDate { get; set; }
        public string AccountMonth { get; set; }
        public string AccountDay { get; set; }
        public string Version { get; set; }
        public float? Capital { get; set; }
        public int? Headcount { get; set; }
        public string PoDeliver { get; set; }
        public string Currency { get; set; }
        public string TradeTerm { get; set; }
        public string PaymentType { get; set; }
        public string PaymentTerm { get; set; }
        public string ReceiptReceive { get; set; }
        public string InvoiceCount { get; set; }
        public string TaxNo { get; set; }
        public string Taxation { get; set; }
        public string PermitPartialDelivery { get; set; }
        public string TaxAmountCalculateType { get; set; }
        public string InvocieAttachedStatus { get; set; }
        public string CertificateFormatType { get; set; }
        public float? DepositRate { get; set; }
        public string TradeItem { get; set; }
        public string RemitBank { get; set; }
        public string RemitAccount { get; set; }
        public string AccountPayable { get; set; }
        public string AccountOverhead { get; set; }
        public string AccountInvoice { get; set; }
        public string SupplierLevel { get; set; }
        public string DeliveryRating { get; set; }
        public string QualityRating { get; set; }
        public string RelatedPerson { get; set; }
        public int? PurchaseUserId { get; set; }
        public string PurchaseUserNo { get; set; }
        public string SupplierRemark { get; set; }
        public string TransferStatus { get; set; }
        public DateTime? TransferDate { get; set; }
        public string Status { get; set; }
        public string ContactUser { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
    }
    #endregion

    #region//品號庫存欄位
    public class ItemInventory
    {
        public int? ItemInventoryId { get; set; }
        public int? CompanyId { get; set; }
        public int? MtlItemId { get; set; }
        public string MtlItemNo { get; set; }
        public int? InventoryId { get; set; }
        public string InventoryNo { get; set; }
        public double? InventoryQty { get; set; }
        public DateTime? TransferDate { get; set; }
        public DateTime? TransferTime { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
    }
    #endregion

    #region //庫別欄位
    public class Inventory
    {
        public int? InventoryId { get; set; }
        public int? CompanyId { get; set; }
        public string InventoryNo { get; set; }
        public string InventoryName { get; set; }
        public string InventoryType { get; set; }
        public string MrpCalculation { get; set; }
        public string ConfirmStatus { get; set; }
        public string SaveStatus { get; set; }
        public string InventoryDesc { get; set; }
        public string TransferStatus { get; set; }
        public DateTime? TransferDate { get; set; }
        public string Status { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
    }
    #endregion

    #region //訂單欄位
    #region //BM訂單單頭欄位
    public class SaleOrder
    {
        public int? SoId { get; set; }
        public int? CompanyId { get; set; }
        public int? DepartmentId { get; set; }
        public string DepartmentNo { get; set; }
        public string SoErpPrefix { get; set; }
        public string SoErpNo { get; set; }
        public string Version { get; set; }
        public DateTime? SoDate { get; set; }
        public DateTime? DocDate { get; set; }
        public string DocDateStr { get; set; } //字串日期資料(picker format)
        public string SoRemark { get; set; }
        public int? CustomerId { get; set; }
        public string CustomerNo { get; set; }
        public string CustomerShortName { get; set; }
        public int? SalesmenId { get; set; }
        public string SalesmenNo { get; set; }
        public string SalesmenName { get; set; }
        public string SalesmenGender { get; set; }
        public string CustomerAddressFirst { get; set; }
        public string CustomerAddressSecond { get; set; }
        public string CustomerPurchaseOrder { get; set; }
        public string DepositPartial { get; set; }
        public double? DepositRate { get; set; }
        public string Currency { get; set; }
        public double? ExchangeRate { get; set; }
        public string TaxNo { get; set; }
        public string Taxation { get; set; }
        public double? BusinessTaxRate { get; set; }
        public string DetailMultiTax { get; set; }
        public double? TotalQty { get; set; }
        public double? Amount { get; set; }
        public double? TaxAmount { get; set; }
        public string ShipMethod { get; set; }
        public string TradeTerm { get; set; }
        public string PaymentTerm { get; set; }
        public string PaymentTermName { get; set; }
        public string PriceTerm { get; set; }
        public string ConfirmStatus { get; set; }
        public int? ConfirmUserId { get; set; }
        public string ConfirmUserNo { get; set; }
        public string ConfirmUserName { get; set; }
        public string ConfirmUserGender { get; set; }
        public string TransferStatus { get; set; }
        public DateTime? TransferDate { get; set; }
        public DateTime? TransferTime { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
        public string SoErpFullNo { get; set; }
        public string SoDetail { get; set; }
        public string DocStatus { get; set; }
        public string UserNo { get; set; }
        public string UserName { get; set; }
        public string SalesmenDepNo { get; set; }
        public string SalesmenDepName { get; set; }
        public string DepartmentName { get; set; }
        public string CustomerFullNo { get; set; }
        public string BpmTransferStatus { get; set; }
        public string BpmTransferStatusName { get; set; }
        public int? TotalCount { get; set; }
        public string TelNoFirst { get; set; }
        public string FaxNo { get; set; }
        public string DeliveryAddressFirst { get; set; }
        public string GuiNumber { get; set; }
    }
    #endregion

    #region //BM訂單單身欄位
    public class SoDetail
    {
        public int? SoDetailId { get; set; }
        public int? SoId { get; set; }
        public string SoSequence { get; set; }
        public string SoErpPrefix { get; set; }
        public string SoErpNo { get; set; }
        public int? MtlItemId { get; set; }
        public string MtlItemNo { get; set; }
        public string SoMtlItemName { get; set; }
        public string SoMtlItemSpec { get; set; }
        public int? InventoryId { get; set; }
        public string InventoryNo { get; set; }
        public int? UomId { get; set; }
        public string UomNo { get; set; }
        public double? SoQty { get; set; }
        public float? SiQty { get; set; }
        public string ProductType { get; set; }
        public float? FreebieQty { get; set; }
        public float? FreebieSiQty { get; set; }
        public float? SpareQty { get; set; }
        public float? SpareSiQty { get; set; }
        public double? UnitPrice { get; set; }
        public double? Amount { get; set; }
        public DateTime? PromiseDate { get; set; }
        public DateTime? PcPromiseDate { get; set; }
        public string Project { get; set; }
        public string CustomerMtlItemNo { get; set; }
        public string CustomerDwgNo { get; set; }
        public string Version { get; set; }
        public string SoDetailRemark { get; set; }
        public float? SoPriceQty { get; set; }
        public int? SoPriceUomId { get; set; }
        public string SoPriceUomrNo { get; set; }
        public string TaxNo { get; set; }
        public double BusinessTaxRate { get; set; }
        public double DiscountRate { get; set; }
        public double DiscountAmount { get; set; }
        public string ConfirmStatus { get; set; }
        public string ClosureStatus { get; set; }
        public string InventoryFullNo { get; set; }
        public string Currency { get; set; }
        public string SoPriceUomNo { get; set; }
        public string TypeThree { get; set; }
        public string TypeFour { get; set; }
        public string QuotationErp { get; set; }
        public DateTime? TransferDate { get; set; }
        public DateTime? TransferTime { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
    }
    #endregion

    #region //ERP訂單單頭欄位
    public class COPTC
    {
        public string TransferStatus { get; set; }
        public string COMPANY { get; set; }
        public string CREATOR { get; set; }
        public string USR_GROUP { get; set; }
        public string CREATE_DATE { get; set; }
        public string MODIFIER { get; set; }
        public string MODI_DATE { get; set; }
        public string FLAG { get; set; }
        public string CREATE_TIME { get; set; }
        public string CREATE_AP { get; set; }
        public string CREATE_PRID { get; set; }
        public string MODI_TIME { get; set; }
        public string MODI_AP { get; set; }
        public string MODI_PRID { get; set; }
        public string TC001 { get; set; } //單別
        public string TC002 { get; set; } //單號
        public string TC003 { get; set; } //訂單日期
        public string TC004 { get; set; } //客戶代號
        public string TC005 { get; set; } //部門代號
        public string TC006 { get; set; } //業務人員
        public string TC007 { get; set; } //出貨廠別
        public string TC008 { get; set; } //交易幣別
        public double? TC009 { get; set; } //匯率(報關買進匯率)
        public string TC010 { get; set; } //送貨地址(一)
        public string TC011 { get; set; } //送貨地址(二)
        public string TC012 { get; set; } //客戶單號
        public string TC013 { get; set; } //價格條件
        public string TC014 { get; set; } //付款條件
        public string TC015 { get; set; } //備註
        public string TC016 { get; set; } //課稅別
        public string TC017 { get; set; } //L/CNO.
        public string TC018 { get; set; } //連絡人
        public string TC019 { get; set; } //運輸方式
        public string TC020 { get; set; } //起始港口
        public string TC021 { get; set; } //目的港口
        public string TC022 { get; set; } //代理商
        public string TC023 { get; set; } //報關行
        public string TC024 { get; set; } //驗貨公司
        public string TC025 { get; set; } //運輸公司
        public double? TC026 { get; set; } //佣金比率
        public string TC027 { get; set; } //確認碼
        public int? TC028 { get; set; } //列印次數
        public double? TC029 { get; set; } //訂單金額
        public double? TC030 { get; set; } //訂單稅額
        public double? TC031 { get; set; } //總數量
        public string TC032 { get; set; } //CONSIGNEE
        public string TC033 { get; set; } //NOTIFY
        public string TC034 { get; set; } //嘜頭代號
        public string TC035 { get; set; } //目的地
        public string TC036 { get; set; } //往來銀行
        public string TC037 { get; set; } //INVOICE備註
        public string TC038 { get; set; } //PACKING-LIST備註
        public string TC039 { get; set; } //單據日期
        public string TC040 { get; set; } //確認者
        public double? TC041 { get; set; } //營業稅率
        public string TC042 { get; set; } //付款條件代號
        public double? TC043 { get; set; } //總毛重(Kg)
        public double? TC044 { get; set; } //總材積(CUFT)
        public double? TC045 { get; set; } //訂金比率(%)
        public double? TC046 { get; set; } //總包裝數量
        public string TC047 { get; set; } //押匯銀行
        public string TC048 { get; set; } //簽核狀態碼
        public string TC049 { get; set; } //流程代號(多角貿易)
        public string TC050 { get; set; } //拋轉狀態
        public string TC051 { get; set; } //下游廠商
        public int? TC052 { get; set; } //傳送次數
        public string TC053 { get; set; } //客戶全名
        public string TC054 { get; set; } //正嘜
        public string TC055 { get; set; } //側嘜
        public string TC056 { get; set; } //材積單位
        public string TC057 { get; set; } //EBC確認碼
        public string TC058 { get; set; } //EBC訂單號碼
        public string TC059 { get; set; } //EBC訂單版次
        public string TC060 { get; set; } //匯至EBC
        public string TC061 { get; set; } //正嘜文管代號
        public string TC062 { get; set; } //側嘜文管代號
        public string TC063 { get; set; } //發票地址(一)
        public string TC064 { get; set; } //發票地址(二)
        public string TC065 { get; set; } //送貨客戶全名
        public string TC066 { get; set; } //TEL_NO
        public string TC067 { get; set; } //FAX_NO
        public string TC068 { get; set; } //交易條件
        public string TC069 { get; set; } //版次
        public string TC070 { get; set; } //訂金分批
        public string TC071 { get; set; } //客戶英文全名
        public double? TC072 { get; set; } //預留欄位
        public double? TC073 { get; set; } //收入遞延天數
        public string TC074 { get; set; } //預留欄位
        public string TC075 { get; set; } //預留欄位
        public string TC076 { get; set; } //預留欄位
        public string TC077 { get; set; } //不控管信用額度
        public string TC078 { get; set; } //稅別碼
        public string TC079 { get; set; } //通路別
        public string TC080 { get; set; } //地區別
        public string TC081 { get; set; } //國家別
        public string TC082 { get; set; } //型態別
        public string TC083 { get; set; } //路線別
        public string TC084 { get; set; } //其他別
        public string TC085 { get; set; } //出口港
        public string TC086 { get; set; } //經過港口
        public string TC087 { get; set; } //目的港口
        public string TC088 { get; set; } //最上游客戶
        public string TC089 { get; set; } //最上游交易幣別
        public string TC090 { get; set; } //最上游稅別碼
        public string TC091 { get; set; } //單身多稅率
        public string TC092 { get; set; } //來源
        public string TC093 { get; set; } //送貨國家
        public string UDF01 { get; set; }
        public string UDF02 { get; set; }
        public string UDF03 { get; set; }
        public string UDF04 { get; set; }
        public string UDF05 { get; set; }
        public double? UDF06 { get; set; }
        public double? UDF07 { get; set; }
        public double? UDF08 { get; set; }
        public double? UDF09 { get; set; }
        public double? UDF10 { get; set; }
        public string CFIELD01 { get; set; } //單別-單號
    }
    #endregion

    #region //ERP訂單單身欄位
    public class COPTD
    {
        public string COMPANY { get; set; }
        public string CREATOR { get; set; }
        public string USR_GROUP { get; set; }
        public string CREATE_DATE { get; set; }
        public string MODIFIER { get; set; }
        public string MODI_DATE { get; set; }
        public string FLAG { get; set; }
        public string CREATE_TIME { get; set; }
        public string CREATE_AP { get; set; }
        public string CREATE_PRID { get; set; }
        public string MODI_TIME { get; set; }
        public string MODI_AP { get; set; }
        public string MODI_PRID { get; set; }
        public string TD001 { get; set; } //單別
        public string TD002 { get; set; } //單號
        public string TD003 { get; set; } //序號
        public string TD004 { get; set; } //品號
        public string TD005 { get; set; } //品名
        public string TD006 { get; set; } //規格
        public string TD007 { get; set; } //庫別
        public double? TD008 { get; set; } //訂單數量
        public double? TD009 { get; set; } //已交數量
        public string TD010 { get; set; } //單位
        public double? TD011 { get; set; } //單價
        public double? TD012 { get; set; } //金額
        public string TD013 { get; set; } //預交日
        public string TD014 { get; set; } //客戶品號
        public string TD015 { get; set; } //預測代號
        public string TD016 { get; set; } //結案碼
        public string TD017 { get; set; } //前置單據-單別
        public string TD018 { get; set; } //前置單據-單號
        public string TD019 { get; set; } //前置單據-序號
        public string TD020 { get; set; } //備註
        public string TD021 { get; set; } //確認碼
        public double? TD022 { get; set; } //庫存數量
        public string TD023 { get; set; } //小單位
        public double? TD024 { get; set; } //贈品量
        public double? TD025 { get; set; } //贈品已交量
        public double? TD026 { get; set; } //折扣率
        public string TD027 { get; set; } //專案代號
        public string TD028 { get; set; } //預測序號
        public string TD029 { get; set; } //包裝方式
        public double? TD030 { get; set; } //毛重
        public double? TD031 { get; set; } //材積
        public double? TD032 { get; set; } //訂單包裝數量
        public double? TD033 { get; set; } //已交包裝數量
        public double? TD034 { get; set; } //贈品包裝量
        public double? TD035 { get; set; } //贈品已交包裝量
        public string TD036 { get; set; } //包裝單位
        public string TD037 { get; set; } //原始客戶
        public string TD038 { get; set; } //請採購廠商
        public string TD039 { get; set; } //圖號
        public string TD040 { get; set; } //預留欄位
        public string TD041 { get; set; } //預留欄位
        public int? TD042 { get; set; } //預留欄位
        public string TD043 { get; set; } //EBC訂單號碼
        public string TD044 { get; set; } //EBC訂單版次
        public string TD045 { get; set; } //來源
        public string TD046 { get; set; } //圖號版次
        public string TD047 { get; set; } //原預交日
        public string TD048 { get; set; } //排定交貨日
        public string TD049 { get; set; } //1.贈品量 2.備品量
        public double? TD050 { get; set; } //備品量
        public double? TD051 { get; set; } //備品已交量
        public double? TD052 { get; set; } //備品包裝量
        public double? TD053 { get; set; } //備品已交包裝量
        public double? TD054 { get; set; } //預留欄位
        public double? TD055 { get; set; } //預留欄位
        public string TD056 { get; set; } //預留欄位
        public string TD057 { get; set; } //預留欄位
        public string TD058 { get; set; } //預留欄位
        public double? TD059 { get; set; } //贈品率
        public string TD060 { get; set; } //預留欄位
        public double? TD061 { get; set; } //RFQ
        public string TD062 { get; set; } //NewCode
        public string TD063 { get; set; } //測試備註一
        public string TD064 { get; set; } //測試備註二
        public string TD065 { get; set; } //最終客戶代號
        public string TD066 { get; set; } //計畫批號
        public string TD067 { get; set; } //優先順序
        public string TD068 { get; set; } //預留欄位
        public string TD069 { get; set; } //鎖定交期
        public double? TD070 { get; set; } //營業稅率
        public string TD071 { get; set; } //CRM單別
        public string TD072 { get; set; } //CRM單號
        public string TD073 { get; set; } //CRM序號
        public string TD074 { get; set; } //CRM合約代號
        public string TD075 { get; set; } //業務品號
        public double? TD076 { get; set; } //計價數量
        public string TD077 { get; set; } //計價單位
        public double? TD078 { get; set; } //已交計價數量
        public string TD079 { get; set; } //稅別碼
        public double? TD080 { get; set; } //折扣金額
        public string TD500 { get; set; } //排程日期
        public double? TD501 { get; set; } //可排量
        public string TD502 { get; set; } //產品系列
        public string TD503 { get; set; } //客戶需求日
        public string TD504 { get; set; } //以包裝單位計價
        public string UDF01 { get; set; }
        public string UDF02 { get; set; }
        public string UDF03 { get; set; }
        public string UDF04 { get; set; }
        public string UDF05 { get; set; }
        public double? UDF06 { get; set; }
        public double? UDF07 { get; set; }
        public double? UDF08 { get; set; }
        public double? UDF09 { get; set; }
        public double? UDF10 { get; set; }
    }
    #endregion

    #region //PDF訂單欄位
    public class SaleOrderPdf
    {
        public int? SoId { get; set; }
        public string DepartmentNo { get; set; }
        public string DepartmentName { get; set; }
        public string ShippingSiteNo { get; set; }
        public string ShippingSiteName { get; set; }
        public string SoErpPrefixNo { get; set; }
        public string SoErpPrefixName { get; set; }
        public string SoErpNo { get; set; }
        public DateTime? DocDate { get; set; }
        public string CustomerNo { get; set; }
        public string CustomerName { get; set; }
        public string SalesmenNo { get; set; }
        public string SalesmenName { get; set; }
        public string CustomerAddress { get; set; }
        public string TelNo { get; set; }
        public string FaxNo { get; set; }
        public string CustomerPurchaseOrder { get; set; }
        public string GuiNumber { get; set; }
        public string Currency { get; set; }
        public double? ExchangeRate { get; set; }
        public string TaxNo { get; set; }
        public string TaxName { get; set; }
        public string BusinessTaxRate { get; set; }
        public double? TotalQty { get; set; }
        public double? SoAmount { get; set; }
        public double? TaxAmount { get; set; }
        public double? TotalAmount { get; set; }
        public string PaymentTermNo { get; set; }
        public string PaymentTermName { get; set; }
        public string SoDetail { get; set; }
        public string SoSequence { get; set; }
        public string MtlItemNo { get; set; }
        public string MtlItemName { get; set; }
        public string MtlItemSpec { get; set; }
        public int? SoQty { get; set; }
        public string UomNo { get; set; }
        public double? UnitPrice { get; set; }
        public double? Amount { get; set; }
        public DateTime? PromiseDate { get; set; }
        public string InventoryNo { get; set; }
        public string SoDetailRemark { get; set; }
        public string SoFullNo { get; set; }
        public string Taxation { get; set; }
        public string TaxationName { get; set; }
    }
    #endregion

    #region //訂金欄位
    public class SoDeposit
    {
        public string Sequence { get; set; }
        public double Ratio { get; set; }
        public double Amount { get; set; }
        public DateTime? PaymentDate { get; set; }
        public string ClosingStatus { get; set; }
    }
    #endregion
    #endregion

    #region //庫存異動欄位
    #region //BM庫存異動單頭欄位
    public class InventoryTransaction
    {
        public int? ItId { get; set; }
        public int? CompanyId { get; set; }
        public string ItErpPrefix { get; set; }
        public string ItErpNo { get; set; }
        public DateTime? ItDate { get; set; }
        public DateTime? DocDate { get; set; }
        public int? DepartmentId { get; set; }
        public string DepartmentNo { get; set; }
        public string DepartmentName { get; set; }
        public string Remark { get; set; }
        public double? TotalQty { get; set; }
        public double? Amount { get; set; }
        public string ConfirmStatus { get; set; }
        public int? ConfirmUserId { get; set; }
        public string ConfirmUserNo { get; set; }
        public string ConfirmUserName { get; set; }
        public string ConfirmUserGender { get; set; }
        public string TransferStatus { get; set; }
        public DateTime? TransferDate { get; set; }
        public DateTime? TransferTime { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
        public string ItErpFullNo { get; set; }
        public string ItDetail { get; set; }
        public string DocStatus { get; set; }
        public int? TotalCount { get; set; }
    }
    #endregion

    #region //BM庫存異動單身欄位
    public class ItDetail
    {
        public int? ItDetailId { get; set; }
        public int? ItId { get; set; }
        public string ItErpPrefix { get; set; }
        public string ItErpNo { get; set; }
        public string ItSequence { get; set; }
        public int? MtlItemId { get; set; }
        public string MtlItemNo { get; set; }
        public string ItMtlItemName { get; set; }
        public string ItMtlItemSpec { get; set; }
        public double? ItQty { get; set; }
        public int? UomId { get; set; }
        public string UomNo { get; set; }
        public double? InvQty { get; set; }
        public double? UnitCost { get; set; }
        public double? Amount { get; set; }
        public int? InventoryId { get; set; }
        public string InventoryNo { get; set; }
        public int? ToInventoryId { get; set; }
        public string ToInventoryNo { get; set; }
        public string ItRemark { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
    }
    #endregion

    #region //ERP庫存異動單頭欄位
    public class INVTA
    {
        public string COMPANY { get; set; }
        public string CREATOR { get; set; }
        public string USR_GROUP { get; set; }
        public string CREATE_DATE { get; set; }
        public string MODIFIER { get; set; }
        public string MODI_DATE { get; set; }
        public string FLAG { get; set; }
        public string CREATE_TIME { get; set; }
        public string CREATE_AP { get; set; }
        public string CREATE_PRID { get; set; }
        public string MODI_TIME { get; set; }
        public string MODI_AP { get; set; }
        public string MODI_PRID { get; set; }
        public string TA001 { get; set; } //單別
        public string TA002 { get; set; } //單號
        public string TA003 { get; set; } //異動日期
        public string TA004 { get; set; } //部門代號
        public string TA005 { get; set; } //備註
        public string TA006 { get; set; } //確認碼
        public int? TA007 { get; set; } //列印次數
        public string TA008 { get; set; } //廠別代號
        public string TA009 { get; set; } //單據性質碼
        public int? TA010 { get; set; } //件數
        public double? TA011 { get; set; } //總數量
        public double? TA012 { get; set; } //總金額
        public string TA013 { get; set; } //產生分錄碼
        public string TA014 { get; set; } //單據日期
        public string TA015 { get; set; } //確認者
        public double? TA016 { get; set; } //總包裝數量
        public string TA017 { get; set; } //簽核狀態碼
        public string TA018 { get; set; } //保稅碼
        public int TA019 { get; set; } //傳送次數
        public string TA020 { get; set; } //運輸方式
        public string TA021 { get; set; } //派車單別
        public string TA022 { get; set; } //派車單號
        public double? TA023 { get; set; } //預留欄位
        public double? TA024 { get; set; } //預留欄位
        public string TA025 { get; set; } //來源
        public string TA026 { get; set; } //預留欄位
        public string TA027 { get; set; } //預留欄位
        public string TA028 { get; set; } //海關手冊
        public string UDF01 { get; set; }
        public string UDF02 { get; set; }
        public string UDF03 { get; set; }
        public string UDF04 { get; set; }
        public string UDF05 { get; set; }
        public double? UDF06 { get; set; }
        public double? UDF07 { get; set; }
        public double? UDF08 { get; set; }
        public double? UDF09 { get; set; }
        public double? UDF10 { get; set; }
        public string CFIELD01 { get; set; } //單別-單號
    }
    #endregion

    #region //ERP庫存異動單身欄位
    public class INVTB
    {
        public string COMPANY { get; set; }
        public string CREATOR { get; set; }
        public string USR_GROUP { get; set; }
        public string CREATE_DATE { get; set; }
        public string MODIFIER { get; set; }
        public string MODI_DATE { get; set; }
        public string FLAG { get; set; }
        public string CREATE_TIME { get; set; }
        public string CREATE_AP { get; set; }
        public string CREATE_PRID { get; set; }
        public string MODI_TIME { get; set; }
        public string MODI_AP { get; set; }
        public string MODI_PRID { get; set; }
        public string TB001 { get; set; } //單別
        public string TB002 { get; set; } //單號
        public string TB003 { get; set; } //序號
        public string TB004 { get; set; } //品號
        public string TB005 { get; set; } //品名
        public string TB006 { get; set; } //規格
        public double? TB007 { get; set; } //數量
        public string TB008 { get; set; } //單位
        public double? TB009 { get; set; } //庫存數量
        public double? TB010 { get; set; } //單位成本
        public double? TB011 { get; set; } //金額
        public string TB012 { get; set; } //轉出庫
        public string TB013 { get; set; } //轉入庫
        public string TB014 { get; set; } //批號
        public string TB015 { get; set; } //有效日期
        public string TB016 { get; set; } //複檢日期
        public string TB017 { get; set; } //備註
        public string TB018 { get; set; } //確認碼
        public string TB019 { get; set; } //異動日期
        public string TB020 { get; set; } //小單位
        public string TB021 { get; set; } //專案代號
        public double? TB022 { get; set; } //包裝數量
        public string TB023 { get; set; } //包裝單位
        public string TB024 { get; set; } //保稅碼
        public double? TB025 { get; set; } //產品序號數量
        public string TB026 { get; set; } //轉出儲位
        public string TB027 { get; set; } //轉入儲位
        public double? TB028 { get; set; } //預留欄位
        public double? TB029 { get; set; } //預留欄位
        public string TB030 { get; set; } //預留欄位
        public string TB031 { get; set; } //預留欄位
        public string TB032 { get; set; } //預留欄位
        public string TB500 { get; set; } //刻號/BIN記錄
        public string TB501 { get; set; } //刻號管理
        public string TB502 { get; set; } //DTAECODE
        public string UDF01 { get; set; }
        public string UDF02 { get; set; }
        public string UDF03 { get; set; }
        public string UDF04 { get; set; }
        public string UDF05 { get; set; }
        public double? UDF06 { get; set; }
        public double? UDF07 { get; set; }
        public double? UDF08 { get; set; }
        public double? UDF09 { get; set; }
        public double? UDF10 { get; set; }
    }
    #endregion

    #region //裝置庫別庫存
    public class DeciveInventory
    {
        public int? InventoryId { get; set; }
        public string InventoryNo { get; set; }
        public string InventoryName { get; set; }
        public string InventoryWithNo { get; set; }
        public double? InventoryQty { get; set; }
    }
    #endregion
    #endregion

    #region //暫出單欄位
    #region //BM暫出單單頭欄位(舊)
    public class TemporaryOrder
    {
        public int? ToId { get; set; }
        public int? DoId { get; set; }
        public string SoErpPrefix { get; set; }
        public string SoErpNo { get; set; }
        public int? SoId { get; set; }
        public int? CompanyId { get; set; }
        public int? DepartmentId { get; set; }
        public string DepartmentNo { get; set; }
        public int? UserId { get; set; }
        public string UserNo { get; set; }
        public string ToErpPrefix { get; set; }
        public string ToErpNo { get; set; }
        public DateTime? ToDate { get; set; }
        public DateTime? TocDate { get; set; }
        public string ToObject { get; set; }
        public int? ObjectId { get; set; }
        public string ObjectNo { get; set; }
        public string ShipMethod { get; set; }
        public string ToAddressFirst { get; set; }
        public string ToAddressSecond { get; set; }
        public double? TotalQty { get; set; }
        public double? Amount { get; set; }
        public double? TaxAmount { get; set; }
        public string ToRemark { get; set; }
        public string WareHouseToRemark { get; set; }
        public string ConfirmStatus { get; set; }
        public int? ConfirmUserId { get; set; }
        public string ConfirmUserNo { get; set; }
        public string TransferStatus { get; set; }
        public DateTime? TransferDate { get; set; }
        public string Status { get; set; }
        public DateTime? TransferTime { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
    }
    #endregion

    #region //BM暫出單單身欄位(舊)
    public class ToDetail
    {
        public int? ToDetailId { get; set; }
        public string SoErpPrefix { get; set; }
        public string SoErpNo { get; set; }
        public string SoSequence { get; set; }
        public int? ToId { get; set; }
        public int? SoDetailId { get; set; }
        public string ToErpPrefix { get; set; }
        public string ToErpNo { get; set; }
        public string ToSequence { get; set; }
        public int? TransInInventoryId { get; set; }
        public string TransInInventoryNo { get; set; }
        public int? ToQty { get; set; }
        public double? UnitPrice { get; set; }
        public double? Amount { get; set; }
        public string ToDetailRemark { get; set; }
        public string ConfirmStatus { get; set; }
        public int? ConfirmUserId { get; set; }
        public string ClosureStatus { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
    }
    #endregion

    #region //BM暫出單欄位
    public class TempShippingNote
    {
        public int? TsnId { get; set; }
        public int? CompanyId { get; set; }
        public string TsnErpPrefix { get; set; }
        public string TsnErpNo { get; set; }
        public DateTime? TsnDate { get; set; }
        public DateTime? DocDate { get; set; }
        public string ToObject { get; set; }
        public int? ObjectCustomer { get; set; }
        public int? ObjectSupplier { get; set; }
        public int? ObjectUser { get; set; }
        public string ObjectOther { get; set; }
        public int? DepartmentId { get; set; }
        public string DepartmentNo { get; set; }
        public int? UserId { get; set; }
        public string UserNo { get; set; }
        public string UserName { get; set; }
        public string UserGender { get; set; }
        public string Remark { get; set; }
        public string OtherRemark { get; set; }
        public double? TotalQty { get; set; }
        public double? Amount { get; set; }
        public double? TaxAmount { get; set; }
        public string ConfirmStatus { get; set; }
        public int? ConfirmUserId { get; set; }
        public string ConfirmUserNo { get; set; }
        public string ConfirmUserName { get; set; }
        public string ConfirmUserGender { get; set; }
        public string TransferStatus { get; set; }
        public DateTime? TransferDate { get; set; }
        public DateTime? TransferTime { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
        public string TsnErpFullNo { get; set; }
        public string ToObjectName { get; set; }
        public string ObjectName { get; set; }
        public string TsnDetail { get; set; }
        public string DocStatus { get; set; }
        public int? TotalCount { get; set; }
    }
    #endregion

    #region //BM暫出單單身欄位
    public class TsnDetail
    {
        public int? TsnDetailId { get; set; }
        public int? TsnId { get; set; }
        public string TsnErpPrefix { get; set; }
        public string TsnErpNo { get; set; }
        public string TsnSequence { get; set; }
        public int? MtlItemId { get; set; }
        public string MtlItemNo { get; set; }
        public string TsnMtlItemName { get; set; }
        public string TsnMtlItemSpec { get; set; }
        public int? TsnOutInventory { get; set; }
        public string TsnOutInventoryNo { get; set; }
        public int? TsnInInventory { get; set; }
        public string TsnInInventoryNo { get; set; }
        public double? TsnQty { get; set; }
        public int? UomId { get; set; }
        public string UomNo { get; set; }
        public string ProductType { get; set; }
        public double? FreebieOrSpareQty { get; set; }
        public double? TsnPriceQty { get; set; }
        public int? TsnPriceUomId { get; set; }
        public string TsnPriceUomNo { get; set; }
        public double BusinessTaxRate { get; set; }
        public DateTime? ExpectedReturnDate { get; set; }
        public double? UnitPrice { get; set; }
        public double? Amount { get; set; }
        public int? SoDetailId { get; set; }
        public string LotNumber { get; set; }
        public string SoErpPrefix { get; set; }
        public string SoErpNo { get; set; }
        public string SoSequence { get; set; }
        public string TsnRemark { get; set; }
        public double? SaleQty { get; set; }
        public double? SaleFSQty { get; set; }
        public double? ReturnQty { get; set; }
        public double? ReturnFSQty { get; set; }
        public string ConfirmStatus { get; set; }
        public string ClosureStatus { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
    }
    #endregion

    #region //ERP暫出單單頭欄位
    public class INVTF
    {
        public string TransferStatus { get; set; }
        public string COMPANY { get; set; }
        public string CREATOR { get; set; }
        public string USR_GROUP { get; set; }
        public string CREATE_DATE { get; set; }
        public string MODIFIER { get; set; }
        public string MODI_DATE { get; set; }
        public string FLAG { get; set; }
        public string CREATE_TIME { get; set; }
        public string CREATE_AP { get; set; }
        public string CREATE_PRID { get; set; }
        public string MODI_TIME { get; set; }
        public string MODI_AP { get; set; }
        public string MODI_PRID { get; set; }
        public string TF001 { get; set; } //異動單別
        public string TF002 { get; set; } //異動單號
        public string TF003 { get; set; } //異動日期
        public string TF004 { get; set; } //異動對象
        public string TF005 { get; set; } //對象代碼
        public string TF006 { get; set; } //對象簡稱
        public string TF007 { get; set; } //部門代號
        public string TF008 { get; set; } //員工代號
        public string TF009 { get; set; } //廠別
        public string TF010 { get; set; } //課稅別
        public string TF011 { get; set; } //幣別
        public double? TF012 { get; set; } //匯率
        public int? TF013 { get; set; } //件數
        public string TF014 { get; set; } //備註
        public string TF015 { get; set; } //對象全名
        public string TF016 { get; set; } //地址一
        public string TF017 { get; set; } //地址二
        public string TF018 { get; set; } //其它備註
        public int? TF019 { get; set; } //列印次數
        public string TF020 { get; set; } //確認碼
        public string TF021 { get; set; } //更新碼
        public int? TF022 { get; set; } //總數量
        public double? TF023 { get; set; } //總金額
        public string TF024 { get; set; } //單據日期
        public string TF025 { get; set; } //確認者
        public double? TF026 { get; set; } //營業稅率
        public double? TF027 { get; set; } //稅額
        public int? TF028 { get; set; } //總包裝數量
        public string TF029 { get; set; } //簽核狀態碼
        public int? TF030 { get; set; } //傳送次數
        public string TF031 { get; set; } //運輸方式
        public string TF032 { get; set; } //派車單別
        public string TF033 { get; set; } //派車單號
        public int? TF034 { get; set; } //預留欄位
        public int? TF035 { get; set; } //預留欄位
        public string TF036 { get; set; } //來源
        public string TF037 { get; set; } //銷貨單單價別
        public string TF038 { get; set; } //預留欄位
        public string TF039 { get; set; } //稅別碼
        public string TF040 { get; set; } //單身多稅率
        public string TF041 { get; set; } //預計轉銷日
        public string TF042 { get; set; } //轉銷日
        public string TF043 { get; set; } //出通單更新碼
        public string TF044 { get; set; } //不控管信用額度
        public string UDF01 { get; set; }
        public string UDF02 { get; set; }
        public string UDF03 { get; set; }
        public string UDF04 { get; set; }
        public string UDF05 { get; set; }
        public double? UDF06 { get; set; }
        public double? UDF07 { get; set; }
        public double? UDF08 { get; set; }
        public double? UDF09 { get; set; }
        public double? UDF10 { get; set; }
        public string CFIELD01 { get; set; }
        public string TF045 { get; set; } //L/C_NO
        public string TF046 { get; set; } //INVOICE_NO
    }
    #endregion

    #region //ERP暫出單單身欄位
    public class INVTG
    {
        public string COMPANY { get; set; }
        public string CREATOR { get; set; }
        public string USR_GROUP { get; set; }
        public string CREATE_DATE { get; set; }
        public string MODIFIER { get; set; }
        public string MODI_DATE { get; set; }
        public string FLAG { get; set; }
        public string CREATE_TIME { get; set; }
        public string CREATE_AP { get; set; }
        public string CREATE_PRID { get; set; }
        public string MODI_TIME { get; set; }
        public string MODI_AP { get; set; }
        public string MODI_PRID { get; set; }
        public string TG001 { get; set; } //異動單別
        public string TG002 { get; set; } //異動單號
        public string TG003 { get; set; } //序號
        public string TG004 { get; set; } //品號
        public string TG005 { get; set; } //品名
        public string TG006 { get; set; } //規格
        public string TG007 { get; set; } //轉出庫別
        public string TG008 { get; set; } //轉入庫別
        public int? TG009 { get; set; } //數量
        public string TG010 { get; set; } //單位
        public string TG011 { get; set; } //小單位
        public double? TG012 { get; set; } //單價
        public double? TG013 { get; set; } //金額
        public string TG014 { get; set; } //來源單別
        public string TG015 { get; set; } //來源單號
        public string TG016 { get; set; } //來源序號
        public string TG017 { get; set; } //批號
        public string TG018 { get; set; } //專案代號
        public string TG019 { get; set; } //備註
        public int? TG020 { get; set; } //轉進銷量
        public int? TG021 { get; set; } //歸還量
        public string TG022 { get; set; } //確認碼
        public string TG023 { get; set; } //更新碼
        public string TG024 { get; set; } //結案碼
        public string TG025 { get; set; } //有效日期
        public string TG026 { get; set; } //複檢日期
        public string TG027 { get; set; } //預計歸還日
        public int? TG028 { get; set; } //包裝數量
        public int? TG029 { get; set; } //轉進銷包裝量
        public int? TG030 { get; set; } //歸還包裝量
        public string TG031 { get; set; } //包裝單位
        public string TG032 { get; set; } //預留欄位
        public int? TG033 { get; set; } //產品序號數量
        public int? TG034 { get; set; } //預留欄位
        public string TG035 { get; set; } //轉出儲位
        public string TG036 { get; set; } //轉入儲位
        public int? TG037 { get; set; } //預留欄位
        public int? TG038 { get; set; } //預留欄位
        public string TG039 { get; set; } //預留欄位
        public string TG040 { get; set; } //最終客戶代號
        public string TG041 { get; set; } //預留欄位
        public double? TG042 { get; set; } //營業稅率
        public string TG043 { get; set; } //類型
        public int? TG044 { get; set; } //贈/備品量
        public int? TG045 { get; set; } //贈/備品包裝量
        public int? TG046 { get; set; } //轉銷贈/備品量
        public int? TG047 { get; set; } //轉銷贈/備品包裝量
        public int? TG048 { get; set; } //歸還贈/備品量
        public int? TG049 { get; set; } //歸還贈/備品包裝量
        public string TG050 { get; set; } //業務品號
        public string TG500 { get; set; } //刻號/BIN記錄
        public string TG501 { get; set; } //刻號管理
        public string TG502 { get; set; } //DATECODE
        public string TG503 { get; set; } //產品系列
        public string UDF01 { get; set; }
        public string UDF02 { get; set; }
        public string UDF03 { get; set; }
        public string UDF04 { get; set; }
        public string UDF05 { get; set; }
        public int? UDF06 { get; set; }
        public int? UDF07 { get; set; }
        public int? UDF08 { get; set; }
        public int? UDF09 { get; set; }
        public int? UDF10 { get; set; }
        public string TG051 { get; set; } //稅別碼
        public int? TG052 { get; set; } //計價數量
        public string TG053 { get; set; } //計價單位
        public int? TG054 { get; set; } //轉進銷計價數量
        public int? TG055 { get; set; } //歸還計價數量
    }
    #endregion
    #endregion

    #region //暫出歸還單欄位
    #region //BM暫出歸還單欄位
    public class TempShippingReturnNote
    {
        public int? TsrnId { get; set; }
        public int? CompanyId { get; set; }
        public string TsrnErpPrefix { get; set; }
        public string TsrnErpNo { get; set; }
        public DateTime? TsrnDate { get; set; }
        public DateTime? DocDate { get; set; }

        public string ToObject { get; set; }
        public int? ObjectCustomer { get; set; }
        public int? ObjectSupplier { get; set; }
        public int? ObjectUser { get; set; }
        public string ObjectNo { get; set; }
        public string ObjectOther { get; set; }
        public int? DepartmentId { get; set; }
        public string DepartmentNo { get; set; }
        public int? UserId { get; set; }
        public string UserNo { get; set; }
        public string Factory { get; set; }
        public string TaxType { get; set; }
        public string Currency { get; set; }
        public string Remark { get; set; }
        public double? ExchangeRate { get; set; }
        public string CustomerAddressFirst { get; set; }
        public string CustomerAddressSecond { get; set; }
        public string OtherRemark { get; set; }
        public string ConfirmStatus { get; set; }
        public double? TotalQty { get; set; }
        public double? Amount { get; set; }
        public double? TaxAmount { get; set; }
        public int? ConfirmUserId { get; set; }
        public string ConfirmUserNo { get; set; }
        public string ConfirmUserName { get; set; }
        public string ConfirmUserGender { get; set; }
        public double? TaxRate { get; set; }
        public string TaxCode { get; set; }
        public string MultiTaxRate { get; set; }
        public string TransferStatus { get; set; }
        public DateTime? TransferDate { get; set; }
        public DateTime? TransferTime { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
        public string TsrnErpFullNo { get; set; }
        public string ToObjectName { get; set; }
        public string ObjectName { get; set; }
        public string TsrnDetail { get; set; }
        public string DocStatus { get; set; }
        public int? TotalCount { get; set; }
    }
    #endregion

    #region //BM暫出歸還單單身欄位
    public class TsrnDetail
    {
        public int? TsrnDetailId { get; set; }
        public int? TsrnId { get; set; }
        public string TsrnErpPrefix { get; set; }
        public string TsrnErpNo { get; set; }
        public string TsrnSequence { get; set; }
        public int? MtlItemId { get; set; }
        public string MtlItemNo { get; set; }
        public string TsrnMtlItemName { get; set; }
        public string TsrnMtlItemSpec { get; set; }
        public int? TsrnOutInventory { get; set; }
        public string TsrnOutInventoryNo { get; set; }
        public int? TsrnInInventory { get; set; }
        public string TsrnInInventoryNo { get; set; }
        public double? TsrnQty { get; set; }
        public int? UomId { get; set; }
        public string UomNo { get; set; }

        public double BusinessTaxRate { get; set; }
        public double? UnitPrice { get; set; }
        public double? Amount { get; set; }
        public int? TsnDetailId { get; set; }
        public string TsnErpPrefix { get; set; }
        public string TsnErpNo { get; set; }
        public string TsnSequence { get; set; }
        public string LotNumber { get; set; }
        public DateTime? AvailableDate { get; set; }
        public DateTime? ReCheckDate { get; set; }
        public string ProjectCode { get; set; }
        public string TsrnRemark { get; set; }
        public string ConfirmStatus { get; set; }
        public string OutLocation { get; set; }
        public string IntLocation { get; set; }
        public double? TaxRate { get; set; }
        public string ProductType { get; set; }
        public double? FreebieOrSpareQty { get; set; }
        public string TaxCode { get; set; }
        public double? TsrnPriceQty { get; set; }
        public int? TsrnPriceUomId { get; set; }
        public string TsrnPriceUomNo { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
    }
    #endregion

    #region //ERP暫出歸還單單頭欄位
    public class INVTH
    {
        public string TransferStatus { get; set; }
        public string COMPANY { get; set; }
        public string CREATOR { get; set; }
        public string USR_GROUP { get; set; }
        public string CREATE_DATE { get; set; }
        public string MODIFIER { get; set; }
        public string MODI_DATE { get; set; }
        public string FLAG { get; set; }
        public string CREATE_TIME { get; set; }
        public string CREATE_AP { get; set; }
        public string CREATE_PRID { get; set; }
        public string MODI_TIME { get; set; }
        public string MODI_AP { get; set; }
        public string MODI_PRID { get; set; }
        public string TH001 { get; set; }
        public string TH002 { get; set; }
        public string TH003 { get; set; }
        public string TH004 { get; set; }
        public string TH005 { get; set; }
        public string TH006 { get; set; }
        public string TH007 { get; set; }
        public string TH008 { get; set; }
        public string TH009 { get; set; }
        public string TH010 { get; set; }
        public string TH011 { get; set; }
        public double? TH012 { get; set; }
        public int? TH013 { get; set; }
        public string TH014 { get; set; }
        public string TH015 { get; set; }
        public string TH016 { get; set; }
        public string TH017 { get; set; }
        public string TH018 { get; set; }
        public int? TH019 { get; set; }
        public string TH020 { get; set; }
        public double? TH021 { get; set; }
        public double? TH022 { get; set; }
        public string TH023 { get; set; }
        public string TH024 { get; set; }
        public double? TH025 { get; set; }
        public double? TH026 { get; set; }
        public double? TH027 { get; set; }
        public string TH028 { get; set; }
        public int? TH029 { get; set; }
        public string TH030 { get; set; }
        public string TH031 { get; set; }
        public string TH032 { get; set; }
        public double? TH033 { get; set; }
        public double? TH034 { get; set; }
        public string TH035 { get; set; }
        public string TH036 { get; set; }
        public string TH037 { get; set; }
        public string TH038 { get; set; }
        public string TH039 { get; set; }
        public string UDF01 { get; set; }
        public string UDF02 { get; set; }
        public string UDF03 { get; set; }
        public string UDF04 { get; set; }
        public string UDF05 { get; set; }
        public double? UDF06 { get; set; }
        public double? UDF07 { get; set; }
        public double? UDF08 { get; set; }
        public double? UDF09 { get; set; }
        public double? UDF10 { get; set; }
        public string CFIELD01 { get; set; } //單別-單號
    }
    #endregion

    #region //ERP暫出歸還單單身欄位
    public class INVTI
    {
        public int TsrnDetailId { get; set; }
        public string COMPANY { get; set; }
        public string CREATOR { get; set; }
        public string USR_GROUP { get; set; }
        public string CREATE_DATE { get; set; }
        public string MODIFIER { get; set; }
        public string MODI_DATE { get; set; }
        public string FLAG { get; set; }
        public string CREATE_TIME { get; set; }
        public string CREATE_AP { get; set; }
        public string CREATE_PRID { get; set; }
        public string MODI_TIME { get; set; }
        public string MODI_AP { get; set; }
        public string MODI_PRID { get; set; }
        public string TI001 { get; set; }
        public string TI002 { get; set; }
        public string TI003 { get; set; }
        public string TI004 { get; set; }
        public string TI005 { get; set; }
        public string TI006 { get; set; }
        public string TI007 { get; set; }
        public string TI008 { get; set; }
        public double? TI009 { get; set; }
        public string TI010 { get; set; }
        public string TI011 { get; set; }
        public double? TI012 { get; set; }
        public double? TI013 { get; set; }
        public string TI014 { get; set; }
        public string TI015 { get; set; }
        public string TI016 { get; set; }
        public string TI017 { get; set; }
        public string TI018 { get; set; }
        public string TI019 { get; set; }
        public string TI020 { get; set; }
        public string TI021 { get; set; }
        public string TI022 { get; set; }
        public double? TI023 { get; set; }
        public string TI024 { get; set; }
        public double? TI025 { get; set; }
        public string TI026 { get; set; }
        public string TI027 { get; set; }
        public double? TI028 { get; set; }
        public double? TI029 { get; set; }
        public string TI030 { get; set; }
        public string TI031 { get; set; }
        public string TI032 { get; set; }
        public double? TI033 { get; set; }
        public string TI034 { get; set; }
        public double? TI035 { get; set; }
        public double? TI036 { get; set; }
        public string TI037 { get; set; }
        public double? TI038 { get; set; }
        public string TI039 { get; set; }
        public string TI500 { get; set; }
        public string TI501 { get; set; }
        public string TI502 { get; set; }
        public string TI503 { get; set; }
        public string UDF01 { get; set; }
        public string UDF02 { get; set; }
        public string UDF03 { get; set; }
        public string UDF04 { get; set; }
        public string UDF05 { get; set; }
        public double? UDF06 { get; set; }
        public double? UDF07 { get; set; }
        public double? UDF08 { get; set; }
        public double? UDF09 { get; set; }
        public double? UDF10 { get; set; }
    }
    #endregion
    #endregion

    #region //出貨排程欄位
    public class DeliverySchedule
    {
        public int SoDetailId { get; set; }
        public int SoId { get; set; }
        public string SoSequence { get; set; }
        public int MtlItemId { get; set; }
        public int InventoryId { get; set; }
        public int UomId { get; set; }
        public double SoQty { get; set; }
        public double SiQty { get; set; }
        public string SoDetailRemark { get; set; }
        public string PcPromiseDate { get; set; }
        public string PcPromiseTime { get; set; }
        public string ProductType { get; set; }
        public double FreebieQty { get; set; }
        public double SpareQty { get; set; }
        public double PickRegularQty { get; set; }
        public double PickFreebieQty { get; set; }
        public double PickSpareQty { get; set; }
        public string SoErpPrefix { get; set; }
        public string SoErpNo { get; set; }
        public string SoErpFullNo { get; set; }
        public int CustomerId { get; set; }
        public string CustomerNo { get; set; }
        public string CustomerName { get; set; }
        public string CustomerShortName { get; set; }
        public string CustomerWithNo { get; set; }
        public string MtlItemNo { get; set; }
        public string MtlItemName { get; set; }
        public string MtlItemSpec { get; set; }
        public double PickQty { get; set; }
        public double TsnQty { get; set; }
        public double SaleQty { get; set; }
        public double erp_tsnQty { get; set; }
        public double erp_un_tsnQty { get; set; }
        public double erp_tsrnQty { get; set; }
        public double erp_un_tsrnQty { get; set; }
        public double erp_tsnFreebieQty { get; set; }
        public double erp_un_tsnFreebieQty { get; set; }
        public double erp_tsrnFreebieQty { get; set; }
        public double erp_un_tsrnFreebieQty { get; set; }
        public double erp_tsnSpareQty { get; set; }
        public double erp_un_tsnSpareQty { get; set; }
        public double erp_tsrnSpareQty { get; set; }
        public double erp_un_tsrnSpareQty { get; set; }
        public double erp_snQty { get; set; }
        public double erp_un_snQty { get; set; }
        public double erp_srnQty { get; set; }
        public double erp_un_srnQty { get; set; }
        public double erp_snFreebieQty { get; set; }
        public double erp_un_snFreebieQty { get; set; }
        public double erp_srnFreebieQty { get; set; }
        public double erp_un_srnFreebieQty { get; set; }
        public double erp_snSpareQty { get; set; }
        public double erp_un_snSpareQty { get; set; }
        public double erp_srnSpareQty { get; set; }
        public double erp_un_srnSpareQty { get; set; }
        public int? TotalCount { get; set; }
    }

    public class ErpAccounting
    {
        public string SoErpFullNo { get; set; }
        public string SoSequence { get; set; }
        public double erp_tsnQty { get; set; }
        public double erp_un_tsnQty { get; set; }
        public double erp_tsrnQty { get; set; }
        public double erp_un_tsrnQty { get; set; }
        public double erp_tsnFreebieQty { get; set; }
        public double erp_un_tsnFreebieQty { get; set; }
        public double erp_tsrnFreebieQty { get; set; }
        public double erp_un_tsrnFreebieQty { get; set; }
        public double erp_tsnSpareQty { get; set; }
        public double erp_un_tsnSpareQty { get; set; }
        public double erp_tsrnSpareQty { get; set; }
        public double erp_un_tsrnSpareQty { get; set; }
        public double erp_snQty { get; set; }
        public double erp_un_snQty { get; set; }
        public double erp_srnQty { get; set; }
        public double erp_un_srnQty { get; set; }
        public double erp_snFreebieQty { get; set; }
        public double erp_un_snFreebieQty { get; set; }
        public double erp_srnFreebieQty { get; set; }
        public double erp_un_srnFreebieQty { get; set; }
        public double erp_snSpareQty { get; set; }
        public double erp_un_snSpareQty { get; set; }
        public double erp_srnSpareQty { get; set; }
        public double erp_un_srnSpareQty { get; set; }
    }
    #endregion

    #region //出貨定版欄位
    public class DeliveryFinalize
    {
        public int? index { get; set; }
        public int? soDetailId { get; set; }
        public string pcPromiseDate { get; set; }
        public string pcPromiseTime { get; set; }
        public string customerNo { get; set; }
        public string customerName { get; set; }
        public string soNo { get; set; }
        public string soSeq { get; set; }
        public string mtlItemName { get; set; }
        public double? soQty { get; set; }
        public double? freebieQty { get; set; }
        public double? spareQty { get; set; }
        public int? deliveryCustomer { get; set; }
        public int? deliveryProcess { get; set; }
        public string orderSituation { get; set; }
        public string deliveryMethod { get; set; }
        public string pcDoDetailRemark { get; set; }
    }
    #endregion

    #region //出貨轉暫出欄位
    public class DeliveryToTsn
    {
        public int? DoDetailId { get; set; }
        public string SoErpPrefix { get; set; }
        public string SoErpNo { get; set; }
        public string SoSequence { get; set; }
        public int? TsnOutInventory { get; set; }
        public string TsnOutInventoryNo { get; set; }
        public int? TsnInInventory { get; set; }
        public string TsnInInventoryNo { get; set; }
        public int PickQty { get; set; }
        public string ProductType { get; set; }
        public int FreebieOrSpareQty { get; set; }
        public string LotManagement { get; set; }
        public int SoDetailId { get; set; }
        public string MtlItemNo { get; set; }
        public string LotNumber { get; set; }
    }
    #endregion

    #region //ERP欄位
    public class Erp
    {
        public string ShippingSiteNo { get; set; }
        public string ShippingSiteName { get; set; }
        public string SoErpPrefixNo { get; set; }
        public string SoErpPrefixName { get; set; }
        public string SoErpNo { get; set; }
        public string SoSequence { get; set; }
        public string PaymentTermNo { get; set; }
        public string PaymentTermName { get; set; }
        public string TaxNo { get; set; }
        public string TaxName { get; set; }
        public int? TempQty { get; set; }
    }
    #endregion

    #region //退(換)貨欄位
    public class Rma
    {
        public int? RmaId { get; set; }
        public int? CompanyId { get; set; }
        public string RmaNo { get; set; }
        public int? CustomerId { get; set; }
        public string RmaDate { get; set; }
        public string RmaType { get; set; }
        public string RmaRemark { get; set; }
        public string DocType { get; set; }
        public string ConfirmStatus { get; set; }
        public int? ConfirmUserId { get; set; }
        public string ErpPrefix { get; set; }
        public string ErpNo { get; set; }
        public string TransferStatus { get; set; }
        public string TransferDate { get; set; }
        public string TransferMode { get; set; }
        public string CustomerNo { get; set; }
        public string CustomerShortName { get; set; }
        public string RmaTypeName { get; set; }
        public string ConfirmUserNo { get; set; }
        public string ConfirmUserName { get; set; }
        public string ConfirmUserGender { get; set; }
        public string ConfirmName { get; set; }
        public string DocTypeName { get; set; }
        public string DocStatus { get; set; }
        public int? TotalCount { get; set; }
    }
    #endregion

    #region //ERP單據狀態
    public class ErpDocStatus
    {
        public string ErpPrefix { get; set; }
        public string ErpNo { get; set; }
        public string DocStatus { get; set; }
    }
    #endregion

    #region //製令綁定欄位
    public class WipLink
    {
        public string WoErpPrefix { get; set; }
        public string WoErpNo { get; set; }
        public int PlanQty { get; set; }
        public string SoErpPrefix { get; set; }
        public string SoErpNo { get; set; }
        public string SoSequence { get; set; }
        public string LinkStatus { get; set; }
        public string BindSoStatus { get; set; }
    }
    #endregion

    #region //每日出貨未綁定製令明細欄位
    public class DailyDeliveryWipUnLinkDetail
    {
        public string DoDate { get; set; }
        public string DcShortName { get; set; }
        public string SoErpFullNo { get; set; }
        public string SoQty { get; set; }
        public string MtlItemNo { get; set; }
        public string MtlItemName { get; set; }
        public string MtlItemSpec { get; set; }
        public string PcUserNo { get; set; }
        public string PcUserName { get; set; }
        public int PickQty { get; set; }
        public string DepartmentName { get; set; }
        public bool? WipLink { get; set; } = false;
    }
    #endregion

    #region //庫存整合
    public class InventoryMtlItem
    {
        public int MtlItemId { get; set; }
        public string MtlItemNo { get; set; }
        public string MtlItemName { get; set; }
        public string MtlItemSpec { get; set; }
        public List<InventoryData> InventoryDatas { get; set; }
        public double ScheduledPurchase { get; set; }
        public double ScheduledIn { get; set; }
        public double ScheduledProduce { get; set; }
        public double ShippedQuantity { get; set; }
        public double AvailableInventory { get; set; }
        public int? TotalCount { get; set; }
    }

    public class InventoryData
    {
        public string MtlItemNo { get; set; }
        public string InventoryNo { get; set; }
        public string InventoryName { get; set; }
        public double InventoryQty { get; set; }
    }

    public class InventoryAvailable
    {
        public string MtlItemNo { get; set; }
        public double ScheduledPurchase { get; set; }
        public double ScheduledIn { get; set; }
        public double ScheduledProduce { get; set; }
        public double ShippedQuantity { get; set; }
        public double AvailableInventory { get; set; }
    }
    #endregion

    #region //請購單
    public class PrInfo
    {
        public int CompanyId { get; set; }
        public string CompanyNo { get; set; }               //前端傳入資料
        public int PrId { get; set; }
        public string PrNo { get; set; }
        public string PrErpPrefix { get; set; }            //前端傳入資料
        public string PrErpNo { get; set; }
        public string UserNo { get; set; }                 //前端傳入資料
        public string DepartmentNo { get; set; }           //前端傳入資料
        public string DocDate { get; set; }                //前端傳入資料
        public string PrDate { get; set; }                 //前端傳入資料
        public string Priority { get; set; }               //前端傳入資料
        public double TotalQty { get; set; }
        public double Amount { get; set; }
        public string PrRemark { get; set; }
        public string PrFile { get; set; }
        public List<PrInfoDetail> Details { get; set; }
    }

    public class PrInfoDetail
    {
        public int PrDetailId { get; set; }
        public int PrId { get; set; }
        public string PrSequence { get; set; }
        public int MtlItemId { get; set; }
        public string MtlItemNo { get; set; }               //前端傳入資料
        public string PrMtlItemName { get; set; }           //前端傳入資料
        public string PrMtlItemSpec { get; set; }           //前端傳入資料
        public int InventoryId { get; set; }
        public string InventoryNo { get; set; }             //前端傳入資料
        public int PrUomId { get; set; }
        public double PrQty { get; set; }                   //前端傳入資料
        public string DemandDate { get; set; }              //前端傳入資料
        public int SupplierId { get; set; }
        public string SupplierNo { get; set; }              //前端傳入資料
        public string PrCurrency { get; set; }              //前端傳入資料
        public double PrExchangeRate { get; set; }
        public double PrUnitPrice { get; set; }             //前端傳入資料
        public double PrPrice { get; set; }
        public double PrPriceTw { get; set; }
        public string UrgentMtl { get; set; }               //前端傳入資料
        public string ProductionPlan { get; set; }          //前端傳入資料
        public int SoDetailId { get; set; }
        public string LockStaus { get; set; }
        public string PoStaus { get; set; }
        public string PartialPurchaseStaus { get; set; }
        public string InquiryStatus { get; set; }
        public string TaxNo { get; set; }
        public string Taxation { get; set; }
        public double? BusinessTaxRate { get; set; }
        public string DetailMultiTax { get; set; }
        public string TradeTerm { get; set; }
        public double PrPriceQty { get; set; }
        public int PrPriceUomId { get; set; }
        public double PoQty { get; set; }
        public double PoUomId { get; set; }
        public string PoCurrency { get; set; }
        public double PoUnitPrice { get; set; }
        public double PoPrice { get; set; }
        public string MtlInventory { get; set; }
        public double MtlInventoryQty { get; set; }
        public string ConfirmStatus { get; set; }
        public string ClosureStatus { get; set; }
        public string PrDetailRemark { get; set; }
        public string PrDetailFile { get; set; }
        public DateTime CreateDate { get; set; }
        public DateTime LastModifiedDate { get; set; }
        public int CreateBy { get; set; }
        public int LastModifiedBy { get; set; }
    }

    public class PrDetail
    {
        /// <summary>
        /// TD001 請購單別
        /// </summary>
        public string PrErpPrefix { get; set; }
        /// <summary>
        /// TD002 請購單號
        /// </summary>
        public string PrErpNo { get; set; }
        /// <summary>
        /// TD003 請購序號
        /// </summary>
        public string PrSeq { get; set; }
        /// <summary>
        /// 完整單號
        /// </summary>
        public string PrErpFullNo { get; set; }
        /// <summary>
        /// 完整單號+序號
        /// </summary>
        public string PrErpPrefixNo { get; set; }
        /// <summary>
        /// TB004 品號
        /// </summary>
        public string MtlItemNo { get; set; }
        /// <summary>
        /// TB005 品名
        /// </summary>
        public string PrMtlItemName { get; set; }
        /// <summary>
        /// TB006 規格
        /// </summary>
        public string PrMtlItemSpec { get; set; }
        /// <summary>
        /// TB007 請購單位
        /// </summary>
        public string PrUnit { get; set; }
        /// <summary>
        /// TB008 庫別
        /// </summary>
        public string Inventory { get; set; }
        /// <summary>
        /// 庫別名
        /// </summary>
        public string InventoryName { get; set; }
        /// <summary>
        /// TB009 請購數量
        /// </summary>
        public string PrQty { get; set; }
        /// <summary>
        /// TB011 需求日期
        /// </summary>
        public string DemandDate { get; set; }
        /// <summary>
        /// TB012 備註
        /// </summary>
        public string PrRemark { get; set; }
        /// <summary>
        /// TB013 採購人員
        /// </summary>
        public string PoUser { get; set; }
        /// <summary>
        /// 採購人員
        /// </summary>
        public string PoUserNmae { get; set; }
        /// <summary>
        /// TB014 採購數量
        /// </summary>
        public string PoQty { get; set; }
        /// <summary>
        /// TB015 採購單位
        /// </summary>
        public string PoUnit { get; set; }
        /// <summary>
        /// TB066 計價單位
        /// </summary>
        public string PoPriceUnit { get; set; }
        /// <summary>
        /// TB016 採購幣別
        /// </summary>
        public string PoCurrency { get; set; }
        /// <summary>
        /// 本幣別
        /// </summary>
        public string PoCurrencyLocal { get; set; }
        /// <summary>
        /// TB017 採購單價
        /// </summary>
        public string PoUnitPrice { get; set; }
        /// <summary>
        /// TB018 採購金額
        /// </summary>
        public double PoPrice { get; set; }
        /// <summary>
        /// TB019 交貨日
        /// </summary>
        public string DeliveryDate { get; set; }
        /// <summary>
        /// TB020 鎖定碼
        public string LockStaus { get; set; }
        /// <summary>
        /// 採購碼
        /// </summary>
        public string PoStatus { get; set; }
        /// <summary>
        /// 採購完整單號
        /// </summary>
        public string PoErpPrefixNo { get; set; }
        /// <summary>
        /// TB024 採購註記
        /// </summary>
        public string PoRemark { get; set; }
        /// <summary>
        /// TC014 確認碼
        /// </summary>
        public string ConfirmStatus { get; set; }
        /// <summary>
        /// 課稅別
        /// </summary>
        public string Taxation { get; set; }
        /// <summary>
        /// TD025 急料 >> BAS.[Type] N[DEF]；Y
        /// </summary>
        public string UrgentMtl { get; set; }
        /// <summary>
        /// TB034 請購包裝數量
        /// </summary>
        public string PrPackageQty { get; set; }
        /// <summary>
        /// TB035 採購包裝數量
        /// </summary>
        public string PoPackageQty { get; set; }
        /// <summary>
        /// TB039 結案碼
        /// </summary>
        public string ClosureStatus { get; set; }
        /// <summary>
        /// TB040 詢價碼
        /// </summary>
        public string InquiryStatus { get; set; }
        /// <summary>
        /// TB042 類型
        /// </summary>
        public string PoType { get; set; }
        /// <summary>
        /// TB044 請購匯率
        /// </summary>
        public string PrExchangeRate { get; set; }
        /// <summary>
        /// TB045 本幣金額
        /// </summary>
        public string PrPriceLocal { get; set; }
        /// <summary>
        /// TB046 分批採購
        /// </summary>
        public string PoPartial { get; set; }
        /// <summary>
        /// TB047 預算編號
        /// </summary>
        public string BudgetNo { get; set; }
        /// <summary>
        /// TB048 預算科目
        /// </summary>
        public string BudgetSubject { get; set; }
        /// <summary>
        /// TB049 請購單價
        /// </summary>
        public string PrUnitPrice { get; set; }
        /// <summary>
        /// TB050 請購幣別
        /// </summary>
        public string PrCurrency { get; set; }
        /// <summary>
        /// 本幣別
        /// </summary>
        public string PrCurrencyLocal { get; set; }
        /// <summary>
        /// 請購金額
        /// </summary>
        public string PrPrice { get; set; }
        /// <summary>
        /// TB057 稅別碼
        /// </summary>
        public string TaxCode { get; set; }
        /// <summary>
        /// TB058 交易條件
        /// </summary>
        public string TradeTerm { get; set; }
        /// <summary>
        /// TB065 計價數量
        /// </summary>
        public string PoPriceQty { get; set; }
        /// <summary>
        /// TB063 營業稅率
        /// </summary>
        public string BusinessTaxRate { get; set; }
        /// <summary>
        /// TB064 單身多稅率
        /// </summary>
        public string DetailMultiTax { get; set; }
        /// <summary>
        /// TB065 請購計價數量
        /// </summary>
        public string PrPriceQty { get; set; }
        /// <summary>
        /// TB067 折扣率
        /// </summary>
        public string DiscountRate { get; set; }
        /// <summary>
        /// TB068 折扣金額
        /// </summary>
        public string DiscountAmount { get; set; }
    }
    #endregion

    #region //進貨單
    #region //進貨單單頭(BM)
    public class GoodsReceipt
    {
        public int? GrId { get; set; } //進貨單頭ID
        public string GrNo { get; set; } //MES進貨單號
        public int? CompanyId { get; set; } //公司別
        public string GrErpPrefix { get; set; } //ERP進貨單單別
        public string GrErpNo { get; set; } //ERP進貨單單號
        public DateTime? ReceiptDate { get; set; } //進貨日期
        public string ErpReceiptDate { get; set; } //進貨日期
        public int? SupplierId { get; set; } //MES供應商ID
        public string SupplierNo { get; set; } //供應商
        public string SupplierSo { get; set; } //供應單號
        public string CurrencyCode { get; set; } //幣別
        public double? Exchange { get; set; } //匯率
        public string InvoiceType { get; set; } //發票聯數
        public string TaxType { get; set; } //課稅別
        public string InvoiceNo { get; set; } //發票號碼
        public int? PrintCnt { get; set; } //列印次數
        public string ConfirmStatus { get; set; } //確認碼(Y/N/V)
        public DateTime? DocDate { get; set; } //單據日期
        public string ErpDocDate { get; set; } //單據日期
        public string RenewFlag { get; set; } //生產記錄更新碼(N) 不需維護, ERP用
        public string Remark { get; set; } //備註
        public double? TotalAmount { get; set; } //總進貨金額
        public double? DeductAmount { get; set; } //扣款金額
        public double? OrigTax { get; set; } //原幣稅額
        public double? ReceiptAmount { get; set; } //進貨費用
        public string SupplierName { get; set; } //廠商全名
        public string UiNo { get; set; } //統一編號
        public string DeductType { get; set; } //抵扣區分
        public string ObaccoAndLiquorFlag { get; set; } //菸酒註記
        public string RowCnt { get; set; } //件數
        public double? Quantity { get; set; } //數量合計(單身數量合計)
        public DateTime? InvoiceDate { get; set; } //發票日期
        public string ErpInvoiceDate { get; set; } //發票日期
        public double? OrigPreTaxAmount { get; set; } //原幣未稅貨款金額
        public string ApplyYYMM { get; set; } //申報年月
        public double? TaxRate { get; set; } //營業稅率
        public double? PreTaxAmount { get; set; } //本幣未稅貨款金額
        public double? TaxAmount { get; set; } //本幣稅額
        public string PaymentTerm { get; set; } //付款條件代號
        public int? PoId { get; set; } //MES採購單ID
        public string PoErpPrefix { get; set; } //ERP採購單別
        public string PoErpNo { get; set; } //ERP採購單號
        public string PrepaidErpPrefix { get; set; } //ERP預付待抵單別
        public string PrepaidErpNo { get; set; } //ERP預付待抵單號
        public double? OffsetAmount { get; set; } //沖抵金額
        public double? TaxOffset { get; set; } //沖抵稅額
        public int? PackageQuantity { get; set; } //包裝單身數量合計
        public double? PremiumAmount { get; set; } //本幣沖自籌額
        public string ApproveStatus { get; set; } //簽核狀態碼
        public string InvoiceFlag { get; set; } //隨貨附發票
        public string ReserveTaxCode { get; set; } //保稅碼
        public int? SendCount { get; set; } //傳送次數
        public double? OrigPremiumAmount { get; set; } //原幣沖自籌額
        public string EbcErpPreNo { get; set; } //EBC出貨通知單號
        public string EbcEdition { get; set; } //EBC出貨通知版次
        public string EbcFlag { get; set; } //進貨單是否已匯出給EBC
        public string FromDocType { get; set; } //來源類別
        public string NoticeFlag { get; set; } //通知碼
        public string TaxCode { get; set; } //稅別碼
        public string TradeTerm { get; set; } //交易條件
        public string DetailMultiTax { get; set; } //單身多稅率
        public double? TaxExchange { get; set; } //稅幣匯率
        public string ContactUser { get; set; } //供應商聯絡人
        public string TransferStatus { get; set; } //拋轉ERP狀態
        public DateTime? TransferDate { get; set; } //拋轉ERP時間
        public DateTime? CreateDate { get; set; } //新增日期
        public DateTime? LastModifiedDate { get; set; } //修改日期
        public int? CreateBy { get; set; } //新增人員
        public string CreateUserNo { get; set; } //新增人員工號
        public int? LastModifiedBy { get; set; } //修改人員
        public string LastModifiedUserNo { get; set; } //修改人員工號
    }
    #endregion

    #region //進貨單單身(BM)
    public class GrDetail
    {
        public int? GrDetailId { get; set; } //進貨單身ID
        public int? GrId { get; set; } //進貨單頭ID
        public string GrErpPrefix { get; set; } //ERP進貨單單別
        public string GrErpNo { get; set; } //ERP進貨單單頭單號
        public string GrSeq { get; set; } //序號
        public int? MtlItemId { get; set; } //MES品號ID
        public string MtlItemNo { get; set; } //品號
        public string GrMtlItemName { get; set; } //品名
        public string GrMtlItemSpec { get; set; } //規格
        public double? ReceiptQty { get; set; } //進貨數量
        public int? UomId { get; set; } //MES進貨單位ID
        public string UomNo { get; set; } //進貨單位
        public int? InventoryId { get; set; } //MES進貨庫別ID
        public string InventoryNo { get; set; } //進貨庫別
        public string LotNumber { get; set; } //批號
        public string PoErpPrefix { get; set; } //ERP採購單別
        public string PoErpNo { get; set; } //ERP採購單號
        public string PoSeq { get; set; } //ERP採購序號
        public DateTime? AcceptanceDate { get; set; } //驗收日期
        public string ErpAcceptanceDate { get; set; } //驗收日期
        public int? AcceptQty { get; set; } //驗收數量
        public int? AvailableQty { get; set; } //驗收數量
        public int? ReturnQty { get; set; } //驗收數量
        public double? OrigUnitPrice { get; set; } //原幣進貨單價
        public double? OrigAmount { get; set; } //原幣進貨金額
        public double? OrigDiscountAmt { get; set; } //原幣扣款金額
        public string TrErpPrefix { get; set; } //ERP暫入單單別
        public string TrErpNo { get; set; } //ERP暫入單單號
        public string TrSeq { get; set; } //ERP暫入單序號
        public double? ReceiptExpense { get; set; } //進貨費用
        public string DiscountDescription { get; set; } //扣款說明
        public string PaymentHold { get; set; } //暫不付款
        public string Overdue { get; set; } //逾期碼
        public string QcStatus { get; set; } //檢驗狀態
        public string ReturnCode { get; set; } //驗退碼
        public string ConfirmStatus { get; set; } //確認碼
        public string CloseStatus { get; set; } //結帳碼
        public string ReNewStatus { get; set; } //更新碼
        public string Remark { get; set; } //備註
        public double? InventoryQty { get; set; } //庫存數量
        public string SmallUnit { get; set; } //小單位
        public DateTime? AvailableDate { get; set; } //有效日期
        public string ErpAvailableDate { get; set; } //有效日期
        public DateTime? ReCheckDate { get; set; } //複檢日期
        public string ErpReCheckDate { get; set; } //複檢日期
        public int? ConfirmUser { get; set; } //確認者
        public string ConfirmUserNo { get; set; } //MES確認者ID
        public string ApvErpPrefix { get; set; } //ERP應付憑單單別
        public string ApvErpNo { get; set; } //ERP應付憑單單號
        public string ApvSeq { get; set; } //ERP應付憑單序號
        public string ProjectCode { get; set; } //專案代碼
        public string ExpenseEntry { get; set; } //產生分錄碼-費用
        public string PremiumAmountFlag { get; set; } //沖自籌額碼
        public double? OrigPreTaxAmt { get; set; } //原幣未稅金額
        public double? OrigTaxAmt { get; set; } //原幣稅金額
        public double? PreTaxAmt { get; set; } //本幣未稅金額
        public double? TaxAmt { get; set; } //本幣稅金額
        public int? ReceiptPackageQty { get; set; } //進貨包裝數量
        public int? AcceptancePackageQty { get; set; } //驗收包裝數量
        public int? ReturnPackageQty { get; set; } //驗退包裝數量
        public double? PremiumAmount { get; set; } //本幣沖自籌額
        public string PackageUnit { get; set; } //包裝單位
        public string ReserveTaxCode { get; set; } //保稅碼
        public double? OrigPremiumAmount { get; set; } //原幣沖自籌額
        public int? AvailableUomId { get; set; } //MES計價單位ID
        public string AvailableUomNo { get; set; } //計價單位
        public string OrigCustomer { get; set; } //原始客戶
        public string ApproveStatus { get; set; } //簽核狀態碼
        public string EbcErpPreNo { get; set; } //EBC出貨通知單號
        public string EbcEdition { get; set; } //EBC出貨通知版次
        public int? ProductSeqAmount { get; set; } //產品序號數量
        public string MtlItemType { get; set; } //品號類型
        public string Loaction { get; set; } //儲位
        public double? TaxRate { get; set; } //營業稅率
        public string MultipleFlag { get; set; } //多儲批
        public string GrErpStatus { get; set; } //狀態
        public string TaxCode { get; set; } //稅別碼
        public double? DiscountRate { get; set; } //折扣率
        public double? DiscountPrice { get; set; } //折扣金額
        public string QcType { get; set; } //送測狀態
        public string TransferStatus { get; set; } //拋轉ERP狀態
        public DateTime? TransferDate { get; set; } //拋轉ERP時間
        public DateTime? CreateDate { get; set; } //新增日期
        public DateTime? LastModifiedDate { get; set; } //修改日期
        public int? CreateBy { get; set; } //新增人員
        public string CreateUserNo { get; set; } //新增人員工號
        public int? LastModifiedBy { get; set; } //修改人員
        public string LastModifiedUserNo { get; set; } //修改人員工號
    }
    #endregion

    #region //進貨檢相關
    public class SpreadsheetModel
    {
        public List<Sheets> sheets { get; set; }
        public Formats formats { get; set; }
    }

    public class Sheets
    {
        public string name { get; set; }
        public List<Data> data { get; set; }
        public Cols cols { get; set; }
        public Rows rows { get; set; }
    }

    public class Data
    {
        public string cell { get; set; }
        public string css { get; set; }
        public string format { get; set; }
        public string value { get; set; }
    }

    public class Cols
    {
        public double? width { get; set; }
    }

    public class Rows
    {
        public double? height { get; set; }
    }

    public class Formats
    {
        public string name { get; set; }
        public string id { get; set; }
        public string mask { get; set; }
        public string example { get; set; }
    }
    #endregion
    #endregion

    #region //採購單
    #region //進貨單單頭(BM)
    public class PurchaseOrder
    {
        public int? PoId { get; set; } //採購單頭ID
        public int? CompanyId { get; set; } //公司別
        public string PoErpPrefix { get; set; } //ERP採購單單別
        public string PoErpNo { get; set; } //ERP採購單單號
        public DateTime? PoDate { get; set; } //採購日期
        public int? SupplierId { get; set; } //MES供應商ID
        public string SupplierNo { get; set; } //供應商
        public string CurrencyCode { get; set; } //幣別
        public double? Exchange { get; set; } //匯率
        public string PaymentTerm { get; set; } //付款條件
        public string Remark { get; set; } //備註
        public int? PoUserId { get; set; } //採購人員ID
        public string PoUserNo { get; set; } //採購人員
        public string ConfirmStatus { get; set; } //確認碼
        public double? Taxation { get; set; } //課稅別
        public double? PoPrice { get; set; } //採購金額
        public double? TaxAmount { get; set; } //稅額
        public string FirstAddress { get; set; } //第一送貨地址
        public string SecondAddress { get; set; } //第二送貨地址
        public int? Quantity { get; set; } //數量合計
        public DateTime? DocDate { get; set; } //單據日期
        public int? ConfirmUserId { get; set; } //確認者ID
        public string ConfirmUserNo { get; set; } //確認者工號
        public double? TaxRate { get; set; } //營業稅率
        public string PaymentTermNo { get; set; } //付款條件代號
        public string ApproveStatus { get; set; } //簽核狀態碼
        public string Edition { get; set; } //版次
        public string DepositPartial { get; set; } //訂金分批
        public string TaxNo { get; set; } //稅別碼
        public string TradeTerm { get; set; } //交易條件
        public string DetailMultiTax { get; set; } //單身多稅率
        public string ContactUser { get; set; } //聯絡人
        public string TransferStatus { get; set; } //拋轉ERP狀態
        public DateTime? TransferDate { get; set; } //拋轉ERP時間
        public DateTime? CreateDate { get; set; } //新增日期
        public DateTime? LastModifiedDate { get; set; } //修改日期
        public int? CreateBy { get; set; } //新增人員
        public int? LastModifiedBy { get; set; } //修改人員
    }
    #endregion

    #region //進貨單單身(BM)
    public class PoDetail
    {
        public int? PoDetailId { get; set; } //採購單身ID
        public int? PoId { get; set; } //採購單頭ID
        public string PoErpPrefix { get; set; } //採購單別
        public string PoErpNo { get; set; } //採購單身
        public string PoSeq { get; set; } //序號
        public int? MtlItemId { get; set; } //品號ID
        public string MtlItemNo { get; set; } //品號
        public string PoMtlItemName { get; set; } //採購品名
        public string PoMtlItemSpec { get; set; } //採購規格
        public int? InventoryId { get; set; } //交貨庫別ID
        public string InventoryNo { get; set; } //交貨庫別
        public double? Quantity { get; set; } //採購數量
        public int? UomId { get; set; } //單位ID
        public string UomNo { get; set; } //單位
        public double? PoUnitPrice { get; set; } //採購單價
        public double? PoPrice { get; set; } //採購金額
        public DateTime? PromiseDate { get; set; } //預計交貨日
        public string ReferencePrefix { get; set; } //參考單別
        public string Remark { get; set; } //備註
        public double? SiQty { get; set; } //已交數量
        public string ClosureStatus { get; set; } //結案碼
        public string ConfirmStatus { get; set; } //確認碼
        public double? InventoryQty { get; set; } //庫存數量
        public string SmallUnit { get; set; } //小單位
        public string ReferenceNo { get; set; } //參考單號
        public string Project { get; set; } //專案代號
        public string ReferenceSeq { get; set; } //參考序號
        public string UrgentMtl { get; set; } //急料
        public string PrErpPrefix { get; set; } //請購單別
        public string PrErpNo { get; set; } //請購單號
        public string PrSequence { get; set; } //請購序號
        public string FromDocType { get; set; } //來源
        public string MtlItemType { get; set; } //品號類型
        public string PartialSeq { get; set; } //分批序號
        public DateTime? OriPromiseDate { get; set; } //原預交日
        public DateTime? DeliveryDate { get; set; } //交期確認日
        public double? PoPriceQty { get; set; } //庫存數量
        public int? PoPriceUomId { get; set; } //計價單位ID
        public string PoPriceUomNo { get; set; } //計價單位
        public double? SiPriceQty { get; set; } //已交計價數量
        public double? DiscountRate { get; set; } //折扣率
        public double? DiscountAmount { get; set; } //折扣金額
        public string TaxNo { get; set; } //稅別碼
        public DateTime? CreateDate { get; set; } //新增日期
        public DateTime? LastModifiedDate { get; set; } //修改日期
        public int? CreateBy { get; set; } //新增人員
        public int? LastModifiedBy { get; set; } //修改人員
        public string CurrencyCode { get; set; } //幣別
        public double? Exchange { get; set; } //匯率
        public string PaymentTermNo { get; set; } //付款條件
        public string Taxation { get; set; } //課稅別
        public double? TaxRate { get; set; } //營業稅率
    }
    #endregion

    #region //採購單單頭(ERP)
    public class PURTC
    {
        public string TC001  { get; set; }
        public string TC002  { get; set; }
        public string TC003  { get; set; }
        public string TC004  { get; set; }
        /// <summary>
        /// 交易幣別
        /// </summary>
        public string TC005 { get; set; }
        /// <summary>
        /// 匯率
        /// </summary>
        public double TC006 { get; set; }
        /// <summary>
        /// TC007 價格條件
        /// </summary>
        public string TC007  { get; set; }
        /// <summary>
        /// TC008 付款條件
        /// </summary>
        public string TC008  { get; set; }
        /// <summary>
        /// TC009 備註
        /// </summary>
        public string TC009  { get; set; }
        /// <summary>
        /// TC010 廠別
        /// </summary>
        public string TC010  { get; set; }
        /// <summary>
        /// TC011 採購人員
        /// </summary>
        public string TC011  { get; set; }
        /// <summary>
        /// TC012 列印格式
        /// </summary>
        public string TC012  { get; set; }
        /// <summary>
        /// TC013 列印次數
        /// </summary>
        public double TC013  { get; set; }
        /// <summary>
        /// TC014 確認碼
        /// </summary>
        public string TC014  { get; set; }
        public string TC015  { get; set; }
        public string TC016  { get; set; }
        public string TC017  { get; set; }
        /// <summary>
        /// 課稅別 1.應稅內含、2.應稅外加、3.零稅率、4.免稅、9.不計稅
        /// </summary>
        public string TC018  { get; set; }
        /// <summary>
        /// 採購金額
        /// </summary>
        public decimal TC019  { get; set; }
        /// <summary>
        /// 稅額
        /// </summary>
        public decimal TC020  { get; set; }
        public string TC021  { get; set; }
        public string TC022  { get; set; }
        /// <summary>
        /// 數量合計
        /// </summary>
        public double TC023  { get; set; }
        public string TC024  { get; set; }
        public string TC025  { get; set; }
        /// <summary>
        /// 營業稅率
        /// </summary>
        public double TC026  { get; set; }
        /// <summary>
        /// 付款條件代號
        /// </summary>
        public string TC027  { get; set; }
        /// <summary>
        /// 訂金比率
        /// </summary>
        public double TC028  { get; set; }
        /// <summary>
        /// 包裝數量合計
        /// </summary>
        public double TC029  { get; set; }
        public string TC030  { get; set; }
        /// <summary>
        /// 傳送次數
        /// </summary>
        public int TC031  { get; set; }
        public string TC032  { get; set; }
        public string TC033  { get; set; }
        public string TC034  { get; set; }
        public string TC035  { get; set; }
        public string TC036  { get; set; }
        public string TC037  { get; set; }
        public string TC038  { get; set; }
        public string TC039  { get; set; }
        public string TC040  { get; set; }
        public string TC041  { get; set; }
        public double TC042  { get; set; }
        public double TC043  { get; set; }
        public string TC044  { get; set; }
        public string TC045  { get; set; }
        public string TC046  { get; set; }
        public string TC047  { get; set; }
        public string TC048  { get; set; }
        public string TC049  { get; set; }
        public string TC050  { get; set; }
        public string TC051  { get; set; }
        public string TC052  { get; set; }
        public string TC500  { get; set; }
        public string TC501  { get; set; }
        public string TC502  { get; set; }
        public string TC503  { get; set; }
        public string TC550  { get; set; }
        public string COMPANY { get; set; }
        public string CREATOR { get; set; }
        public string USR_GROUP { get; set; }
        public string CREATE_DATE { get; set; }
        public string MODIFIER { get; set; }
        public string MODI_DATE { get; set; }
        public string FLAG { get; set; } = "1";
        public string CREATE_TIME { get; set; }
        public string CREATE_AP { get; set; }
        public string CREATE_PRID { get; set; } = "SRM";
        public string MODI_TIME { get; set; }
        public string MODI_AP { get; set; }
        public string MODI_PRID { get; set; } = "SRM";
        public string UDF01 { get; set; } = "";
        public string UDF02 { get; set; } = "";
        public string UDF03 { get; set; } = "";
        public string UDF04 { get; set; } = "";
        public string UDF05 { get; set; } = "";
        public double? UDF06 { get; set; } = 0;
        public double? UDF07 { get; set; } = 0;
        public double? UDF08 { get; set; } = 0;
        public double? UDF09 { get; set; } = 0;
        public double? UDF10 { get; set; } = 0;
    }
    #endregion

    #region //採購單單身(ERP)
    public class PURTD
    {
        /// <summary>
        /// 供應商代號
        /// </summary>
        public string TC004 { get; set; }
        /// <summary>
        /// 交易幣別
        /// </summary>
        public string TC005 { get; set; }
        /// <summary>
        /// 匯率
        /// </summary>
        public double TC006 { get; set; }
        /// <summary>
        /// 採購人員
        /// </summary>
        public string TC011 { get; set; }
        /// <summary>
        /// 稅別碼
        /// </summary>
        public string TC047 { get; set; }
        /// <summary>
        /// TC048 交易條件
        /// </summary>
        public string TC048 { get; set; }

        public string TD001  { get; set; } 
        public string TD002  { get; set; } 
        public string TD003  { get; set; } 
        public string TD004  { get; set; } 
        public string TD005  { get; set; } 
        public string TD006  { get; set; } 
        public string TD007  { get; set; }
        /// <summary>
        /// 採購數量
        /// </summary>
        public double TD008  { get; set; } 
        public string TD009  { get; set; } 
        /// <summary>
        /// 採購單價
        /// </summary>
        public decimal TD010  { get; set; } 
        /// <summary>
        /// 採購金額
        /// </summary>
        public decimal TD011  { get; set; } 
        public string TD012  { get; set; } 
        public string TD013  { get; set; }
        /// <summary>
        /// 備註
        /// </summary>
        public string TD014  { get; set; }
        /// <summary>
        /// TD015 已交數量
        /// </summary>
        public double TD015  { get; set; }
        /// <summary>
        /// TD016 結案碼
        /// </summary>
        public string TD016  { get; set; } 
        public string TD017  { get; set; } 
        public string TD018  { get; set; }
        /// <summary>
        /// TD019 庫存數量
        /// </summary>
        public double TD019  { get; set; } 
        public string TD020  { get; set; } 
        public string TD021  { get; set; } 
        public string TD022  { get; set; } 
        public string TD023  { get; set; } 
        public string TD024  { get; set; } 
        public string TD025  { get; set; } 
        public string TD026  { get; set; } 
        public string TD027  { get; set; } 
        public string TD028  { get; set; } 
        public string TD029  { get; set; }
        /// <summary>
        /// TD030 採購包裝數量
        /// </summary>
        public double TD030  { get; set; }
        /// <summary>
        /// TD031 已交包裝數量
        /// </summary>
        public double TD031  { get; set; } 
        public string TD032  { get; set; } 
        public string TD033  { get; set; } 
        public string TD034  { get; set; } 
        public string TD035  { get; set; } 
        public string TD036  { get; set; } 
        public string TD037  { get; set; } 
        public string TD038  { get; set; } 
        public string TD039  { get; set; } 
        public string TD040  { get; set; } 
        public string TD041  { get; set; } 
        public string TD042  { get; set; } 
        public string TD043  { get; set; } 
        public string TD044  { get; set; } 
        public string TD045  { get; set; } 
        public string TD046  { get; set; } 
        public string TD047  { get; set; } 
        public double TD048  { get; set; } 
        public double TD049  { get; set; } 
        public string TD050  { get; set; } 
        public string TD051  { get; set; } 
        public string TD052  { get; set; } 
        public string TD053  { get; set; } 
        public string TD054  { get; set; } 
        public string TD055  { get; set; } 
        public string TD056  { get; set; }
        /// <summary>
        /// TD057 營業稅率
        /// </summary>
        public double TD057  { get; set; }
        /// <summary>
        /// TD058 計價數量
        /// </summary>
        public double TD058  { get; set; }
        /// <summary>
        /// TD059 計價單位
        /// </summary>
        public string TD059  { get; set; } 
        public double TD060  { get; set; }
        /// <summary>
        /// TD061 稅別碼
        /// </summary>
        public string TD061  { get; set; }
        /// <summary>
        /// TD062 折扣率
        /// </summary>
        public double TD062  { get; set; }
        /// <summary>
        /// TD063 折扣金額
        /// </summary>
        public double TD063  { get; set; } 
        public string TD500  { get; set; } 
        public string TD501  { get; set; } 
        public string TD502  { get; set; } 
        public string TD503  { get; set; } 
        public string TD550  { get; set; } 
        public double TD551  { get; set; } 
        public string TD552  { get; set; } 
        public string TD553  { get; set; } 
        public string TD554  { get; set; } 
        public string TD555  { get; set; }
        public string COMPANY { get; set; }
        public string CREATOR { get; set; }
        public string USR_GROUP { get; set; }
        public string CREATE_DATE { get; set; }
        public string MODIFIER { get; set; }
        public string MODI_DATE { get; set; }
        public string FLAG { get; set; } = "1";
        public string CREATE_TIME { get; set; }
        public string CREATE_AP { get; set; }
        public string CREATE_PRID { get; set; } = "SRM";
        public string MODI_TIME { get; set; }
        public string MODI_AP { get; set; }
        public string MODI_PRID { get; set; } = "SRM";
        public string UDF01 { get; set; } = "";
        public string UDF02 { get; set; } = "";
        public string UDF03 { get; set; } = "";
        public string UDF04 { get; set; } = "";
        public string UDF05 { get; set; } = "";
        public double? UDF06 { get; set; } = 0;
        public double? UDF07 { get; set; } = 0;
        public double? UDF08 { get; set; } = 0;
        public double? UDF09 { get; set; } = 0;
        public double? UDF10 { get; set; } = 0;
    }
    #endregion


    #endregion

    #region //批號
    #region //批號資料單頭(BM)
    public class LotNumber
    {
        public int? LotNumberId { get; set; } //批號單頭ID
        public int? MtlItemId { get; set; } //品號ID
        public string MtlItemNo { get; set; } //品號
        public string LotNumberNo { get; set; } //批號
        public DateTime? FirstRecriptDate { get; set; } //最早入庫日
        public string FromErpPrefix { get; set; } //參考單別
        public string FromErpNo { get; set; } //參考單號
        public string CloseStatus { get; set; } //結案碼
        public string Remark { get; set; } //備註
        public string TransferStatus { get; set; } //拋轉ERP狀態
        public DateTime? TransferDate { get; set; } //拋轉ERP時間
        public DateTime? CreateDate { get; set; } //新增日期
        public DateTime? LastModifiedDate { get; set; } //修改日期
        public int? CreateBy { get; set; } //新增人員
        public string CreateUserNo { get; set; } //新增人員工號
        public int? LastModifiedBy { get; set; } //修改人員
        public string LastModifiedUserNo { get; set; } //修改人員工號
    }
    #endregion

    #region //批號資料單身(BM)
    public class LnDetail
    {
        public int? LnDetailId { get; set; } //批號單身ID
        public int? LotNumberId { get; set; } //批號單頭ID
        public int? MtlItemId { get; set; } //品號ID
        public string MtlItemNo { get; set; } //品號
        public string LotNumberNo { get; set; } //批號
        public DateTime? TransactionDate { get; set; } //異動日期
        public string FromErpPrefix { get; set; } //異動單別
        public string FromErpNo { get; set; } //異動單號
        public string FromSeq { get; set; } //異動序號
        public int? InventoryId { get; set; } //異動庫別ID
        public string InventoryNo { get; set; } //異動庫別
        public int? TransactionType { get; set; } //異動類別
        public string DocType { get; set; } //異動別
        public double? Quantity { get; set; } //數量
        public string Remark { get; set; } //備註
        public DateTime? CreateDate { get; set; } //新增日期
        public DateTime? LastModifiedDate { get; set; } //修改日期
        public int? CreateBy { get; set; } //新增人員
        public string CreateUserNo { get; set; } //新增人員工號
        public int? LastModifiedBy { get; set; } //修改人員
        public string LastModifiedUserNo { get; set; } //修改人員工號
    }
    #endregion
    #endregion

    #region //訂單拋轉BPM Model
    #region //訂單BPM單頭
    public class SaleOrderModel
    {
        public string Title { get; set; }
        public string Filler { get; set; }
        public string FillerDepartment { get; set; }
        public string FillerNo { get; set; }
        public string FillCalendar { get; set; }
        public string SoErpFullNo { get; set; }
        public string SoErpFullCa { get; set; }
        public string SoErpNo { get; set; }
        public string createDate { get; set; }
        public string DocDate { get; set; }
        public string SalesmenNo { get; set; }
        public string SalesmenName { get; set; }
        public string SalesName { get; set; }
        public string SalesNo { get; set; }
        public string SalesmenDepNo { get; set; }
        public string SalesmenDepName { get; set; }
        public string SalesDep { get; set; }
        public string CustomerFullNo { get; set; }
        public string CustomerPurchaseOrder { get; set; }
        public string Site { get; set; }
        public string SoRemark { get; set; }
        public double Amount { get; set; }
        public double TaxAmount { get; set; }
        public string DbTable { get; set; }
        public int? SoId { get; set; }
        public string CompanyNo { get; set; }
        public List<SoDetailModel> SoDetailModelList { get; set; }
        public int CounterSign { get; set; }
        public string Currency { get; set; }
        public string ExchangeRate { get; set; }
        public string BusinessTaxRate { get; set; }
        public string Taxation { get; set; }
        public string TotalDetail { get; set; }
        public double TotalAmount { get; set; }
        public string Customer { get; set; }
        public string CustomerOrder { get; set; }
        public string PhoneNo { get; set; }
        public string Fax { get; set; }
        public string DeliveryAdd { get; set; }
        public string TaxRate { get; set; }
        public string TaxID { get; set; }
        public string Payment { get; set; }
        public double SoTotalAmount { get; set; }
    }
    #endregion

    #region //訂單BPM單身
    public class SoDetailModel
    {
        public string SoSequence { get; set; } //序號
        public string MtlItemNo { get; set; } //品號
        public string SoMtlItemName { get; set; } //品名
        public string SoMtlItemSpec { get; set; } //規格
        public string InventoryFullNo { get; set; } //庫別
        public string UomNo { get; set; } //單位
        public double? SoQty { get; set; } //訂單數量
        public double? SpareQty { get; set; } //備品量
        public double? FreebieQty { get; set; } //贈品量
        public double? UnitPrice { get; set; } //訂單單價
        public double? Amount { get; set; } //訂單金額
        public double? SoPriceQty { get; set; } //計價數量
        public string PromiseDate { get; set; } //預交日
        public string SoDetailRemark { get; set; } //單身備註
        public List<string> CounterSignDep { get; set; } //加簽部門
    }
    #endregion
    #endregion

    #region //銷貨單

    #region //MES - 銷貨單單頭 SCM.ReceiveOrder
    public class ReceiveOrder
    {
        public int? RoId { get; set; }
        public int? CompanyId { get; set; }
        public string RoErpPrefix { get; set; }
        public string RoErpNo { get; set; }
        public DateTime? ReceiveDate { get; set; }
        public int? CustomerId { get; set; }
        public string CustomerNo { get; set; }
        public int? DepartmentId { get; set; }
        public string DepartmentNo { get; set; }
        public int? SalesmenId { get; set; }
        public string SalesmenNo { get; set; }
        public string CustomerFullName { get; set; }
        public string CustomerAddressFirst { get; set; }
        public string CustomerAddressSecond { get; set; }
        public string Factory { get; set; }
        public string Currency { get; set; }
        public double ExchangeRate { get; set; }
        public double OriginalAmount { get; set; }
        public string InvNumEnd { get; set; }
        public string UiNo { get; set; }
        public string InvoiceType { get; set; }
        public string TaxType { get; set; }
        public string InvoiceAddressFirst { get; set; }
        public string InvoiceAddressSecond { get; set; }
        public string Remark { get; set; }
        public DateTime? InvoiceDate { get; set; }
        public int PrintCount { get; set; }
        public string ConfirmStatus { get; set; }
        public string UpdateCode { get; set; }
        public double OriginalTaxAmount { get; set; }
        public int? CollectionSalesmenId { get; set; }
        public string CollectionSalesmenNo { get; set; }
        public string Remark1 { get; set; }
        public string Remark2 { get; set; }
        public string Remark3 { get; set; }
        public string InvoicesVoid { get; set; }
        public string CustomsClearance { get; set; }
        public int RowCnt { get; set; }
        public double TotalQuantity { get; set; }
        public string CashSales { get; set; }
        public int StaffId { get; set; }
        public string StaffNo { get; set; }
        public string RevenueJournalCode { get; set; }
        public string CostJournalCode { get; set; }
        public string ApplyYYMM { get; set; }
        public string LCNO { get; set; }
        public string INVOICENO { get; set; }
        public int InvoicesPrintCount { get; set; }
        public DateTime? DocDate { get; set; }
        public int? ConfirmUserId { get; set; }
        public string ConfirmUserNo { get; set; }
        public double TaxRate { get; set; }
        public double PretaxAmount { get; set; }
        public double TaxAmount { get; set; }
        public string PaymentTerm { get; set; }
        public int? SoDetailId { get; set; }
        public string SoErpPrefix { get; set; }
        public string SoErpNo { get; set; }
        public string SoSequence { get; set; }
        public string AdvOrderPrefix { get; set; }
        public string AdvOrderNo { get; set; }
        public double OffsetAmount { get; set; }
        public double OffsetTaxAmount { get; set; }
        public double TotalPackages { get; set; }
        public string SignatureStatus { get; set; }
        public string ChangeInvCode { get; set; }
        public string NewRoPrefix { get; set; }
        public string NewRoNo { get; set; }
        public string TransferStatusERP { get; set; }
        public string ProcessCode { get; set; }
        public string AttachInvWithShip { get; set; }
        public string BondCode { get; set; }
        public int TransmissionCount { get; set; }
        public string Invoicer { get; set; }
        public string InvCode { get; set; }
        public string ContactPerson { get; set; }
        public string Courier { get; set; }
        public string SiteCommCalcMethod { get; set; }
        public double SiteCommRate { get; set; }
        public string CommCalcInclTax { get; set; }
        public double TotalCommAmount { get; set; }
        public string TransportMethod { get; set; }
        public string DispatchOrderPrefix { get; set; }
        public string DispatchOrderNo { get; set; }
        public string DeclarationNumber { get; set; }
        public string FullNameOfDelivCustomer { get; set; }
        public string RoPriceType { get; set; }
        public string TelephoneNumber { get; set; }
        public string FaxNumber { get; set; }
        public string ShipNoticePrefix { get; set; }
        public string ShipNoticeNo { get; set; }
        public string TradingTerms { get; set; }
        public string CustomerEgFullName { get; set; }
        public string InvNumGenMethod { get; set; }
        public string DocSourceCode { get; set; }
        public string NoCredLimitControl { get; set; }
        public string InstallmentSettlement { get; set; }
        public int InstallmentCount { get; set; }
        public string AutoApportionByInstallment { get; set; }
        public string StartYearMonth { get; set; }
        public string TaxCode { get; set; }
        public string CustomsManual { get; set; }
        public string RemarkCode { get; set; }
        public string MultipleInvoices { get; set; }
        public string InvNumStart { get; set; }
        public int NumberOfInvoices { get; set; }
        public string MultiTaxRate { get; set; }
        public double TaxCurrencyRate { get; set; }
        public DateTime? VoidDate { get; set; }
        public string VoidApprovalDocNum { get; set; }
        public string VoidReason { get; set; }
        public string Source { get; set; }
        public string IncomeDraftID { get; set; }
        public string IncomeDraftSeq { get; set; }
        public string IncomeVoucherType { get; set; }
        public string IncomeVoucherNumber { get; set; }
        public string Status { get; set; }
        public string ZeroTaxForBuyer { get; set; }
        public string GenLedgerAcctType { get; set; }
        public string InvoiceTime { get; set; }
        public string InvCode2 { get; set; }
        public string InvSymbol { get; set; }
        public string DeliveryCountry { get; set; }
        public string VehicleIDshow { get; set; }
        public string VehicleTypeNumber { get; set; }
        public string VehicleIDhide { get; set; }
        public string InvDonationRecipient { get; set; }
        public string InvRandomCode { get; set; }
        public string ReservedField { get; set; }
        public string CreditCard4No { get; set; }
        public string ContactEmail { get; set; }
        public string ExpectedReceiptDate { get; set; }
        public string OrigInvNumber { get; set; }
        public string TransferStatusMES { get; set; }
        public DateTime? TransferTime { get; set; }
        public DateTime? ConfirmTime { get; set; }
        public string SourceType { get; set; }
        public string PriceSourceTypeMain { get; set; }
        public int SourceOrderId { get; set; }
        public string SourceFull { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
    }
    #endregion

    #region //MES - 銷貨單單身 SCM.RoDetail
    public class RoDetail
    {
        public int? RoDetailId { get; set; }
        public int? RoId { get; set; }
        public string RoErpPrefix { get; set; }
        public string RoErpNo { get; set; }
        public string RoSequence { get; set; }
        public int? MtlItemId { get; set; }
        public string MtlItemNo { get; set; }
        public int? InventoryId { get; set; }
        public string InventoryNo { get; set; }
        public double Quantity { get; set; }
        public int? UomId { get; set; }
        public string UomNo { get; set; }
        public double UnitPrice { get; set; }
        public double Amount { get; set; }
        public int? SoDetailId { get; set; }
        public string SoErpPrefix { get; set; }
        public string SoErpNo { get; set; }
        public string SoSequence { get; set; }
        public string LotNumber { get; set; }
        public string Remarks { get; set; }
        public string CustomerMtlItemNo { get; set; }
        public string ConfirmationCode { get; set; }
        public string UpdateCode { get; set; }
        public double FreeSpareQty { get; set; }
        public double DiscountRate { get; set; }
        public string CheckOutCode { get; set; }
        public string CheckOutPrefix { get; set; }
        public string CheckOutNo { get; set; }
        public string CheckOutSequence { get; set; }
        public string ProjectCode { get; set; }
        public string Type { get; set; }
        public int? TsnDetailId { get; set; }
        public string TsnErpPrefix { get; set; }
        public string TsnErpNo { get; set; }
        public string TsnSequence { get; set; }
        public double OrigPreTaxAmt { get; set; }
        public double OrigTaxAmt { get; set; }
        public double PreTaxAmt { get; set; }
        public double TaxAmt { get; set; }
        public double PackageQty { get; set; }
        public double PackageFreeSpareQty { get; set; }
        public int? PackageUomId { get; set; }
        public string PackageUomNo { get; set; }
        public string BondedCode { get; set; }
        public double SrQty { get; set; }
        public double SrPackageQty { get; set; }
        public double SrOrigPreTaxAmt { get; set; }
        public double SrOrigTaxAmt { get; set; }
        public double SrPreTaxAmt { get; set; }
        public double SrTaxAmt { get; set; }
        public double CommissionRate { get; set; }
        public double CommissionAmount { get; set; }
        public string OriginalCustomer { get; set; }
        public double SrFreeSpareQty { get; set; }
        public double SrPackageFreeSpareQty { get; set; }
        public string NotPayTemp { get; set; }
        public int ProductSerialNumberQty { get; set; }
        public string ForecastCode { get; set; }
        public string ForecastSequence { get; set; }
        public string Location { get; set; }
        public double AvailableQty { get; set; }
        public int? AvailableUomId { get; set; }
        public string AvailableUomNo { get; set; }
        public string MultiBatch { get; set; }
        public double FreeSpareRate { get; set; }
        public string FinalCustomerCode { get; set; }
        public double ReferenceQty { get; set; }
        public double ReferencePackageQty { get; set; }
        public double TaxRate { get; set; }
        public string CRMSource { get; set; }
        public string CRMPrefix { get; set; }
        public string CRMNo { get; set; }
        public string CRMSequence { get; set; }
        public string CRMContractCode { get; set; }
        public string CRMAllowDeduction { get; set; }
        public double CRMDeductionQty { get; set; }
        public string CRMDeductionUnit { get; set; }
        public string DebitAccount { get; set; }
        public string CreditAccount { get; set; }
        public string TaxAmountAccount { get; set; }
        public string BusinessItemNumber { get; set; }
        public string TaxCode { get; set; }
        public double DiscountAmount { get; set; }
        public string K2NO { get; set; }
        public string MarkingBINRecord { get; set; }
        public string MarkingManagement { get; set; }
        public string BillingUnitInPackage { get; set; }
        public string DATECODE { get; set; }
        public string TransferStatusMES { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
    }
    #endregion

    #region //ERP - 銷貨單單頭 COPTG
    public class COPTG
    {
        public string TransferStatusMES { get; set; }
        public string COMPANY { get; set; }
        public string CREATOR { get; set; }
        public string USR_GROUP { get; set; }
        public string CREATE_DATE { get; set; }
        public string MODIFIER { get; set; }
        public string MODI_DATE { get; set; }
        public string FLAG { get; set; }
        public string CREATE_TIME { get; set; }
        public string CREATE_AP { get; set; }
        public string CREATE_PRID { get; set; }
        public string MODI_TIME { get; set; }
        public string MODI_AP { get; set; }
        public string MODI_PRID { get; set; }
        public string TG001 { get; set; }
        public string TG002 { get; set; }
        public string TG003 { get; set; }
        public string TG004 { get; set; }
        public string TG005 { get; set; }
        public string TG006 { get; set; }
        public string TG007 { get; set; }
        public string TG008 { get; set; }
        public string TG009 { get; set; }
        public string TG010 { get; set; }
        public string TG011 { get; set; }
        public double TG012 { get; set; }
        public double TG013 { get; set; }
        public string TG014 { get; set; }
        public string TG015 { get; set; }
        public string TG016 { get; set; }
        public string TG017 { get; set; }
        public string TG018 { get; set; }
        public string TG019 { get; set; }
        public string TG020 { get; set; }
        public string TG021 { get; set; }
        public int TG022 { get; set; }
        public string TG023 { get; set; }
        public string TG024 { get; set; }
        public double TG025 { get; set; }
        public string TG026 { get; set; }
        public string TG027 { get; set; }
        public string TG028 { get; set; }
        public string TG029 { get; set; }
        public string TG030 { get; set; }
        public string TG031 { get; set; }
        public int TG032 { get; set; }
        public double TG033 { get; set; }
        public string TG034 { get; set; }
        public string TG035 { get; set; }
        public string TG036 { get; set; }
        public string TG037 { get; set; }
        public string TG038 { get; set; }
        public string TG039 { get; set; }
        public string TG040 { get; set; }
        public int TG041 { get; set; }
        public string TG042 { get; set; }
        public string TG043 { get; set; }
        public double TG044 { get; set; }
        public double TG045 { get; set; }
        public double TG046 { get; set; }
        public string TG047 { get; set; }
        public string TG048 { get; set; }
        public string TG049 { get; set; }
        public string TG050 { get; set; }
        public string TG051 { get; set; }
        public double TG052 { get; set; }
        public double TG053 { get; set; }
        public double TG054 { get; set; }
        public string TG055 { get; set; }
        public string TG056 { get; set; }
        public string TG057 { get; set; }
        public string TG058 { get; set; }
        public string TG059 { get; set; }
        public string TG060 { get; set; }
        public string TG061 { get; set; }
        public string TG062 { get; set; }
        public int TG063 { get; set; }
        public string TG064 { get; set; }
        public string TG065 { get; set; }
        public string TG066 { get; set; }
        public string TG067 { get; set; }
        public string TG068 { get; set; }
        public double TG069 { get; set; }
        public string TG070 { get; set; }
        public double TG071 { get; set; }
        public string TG072 { get; set; }
        public string TG073 { get; set; }
        public string TG074 { get; set; }
        public string TG075 { get; set; }
        public string TG076 { get; set; }
        public string TG077 { get; set; }
        public string TG078 { get; set; }
        public string TG079 { get; set; }
        public string TG080 { get; set; }
        public string TG081 { get; set; }
        public string TG082 { get; set; }
        public string TG083 { get; set; }
        public double TG084 { get; set; }
        public double TG085 { get; set; }
        public string TG086 { get; set; }
        public string TG087 { get; set; }
        public string TG088 { get; set; }
        public string TG089 { get; set; }
        public string TG090 { get; set; }
        public int TG091 { get; set; }
        public string TG092 { get; set; }
        public string TG093 { get; set; }
        public string TG094 { get; set; }
        public string TG095 { get; set; }
        public string TG096 { get; set; }
        public string TG097 { get; set; }
        public string TG098 { get; set; }
        public int TG099 { get; set; }
        public string TG100 { get; set; }
        public double TG101 { get; set; }
        public string TG102 { get; set; }
        public string TG103 { get; set; }
        public string TG104 { get; set; }
        public string TG105 { get; set; }
        public string TG106 { get; set; }
        public string TG107 { get; set; }
        public string TG108 { get; set; }
        public string TG109 { get; set; }
        public string TG110 { get; set; }
        public string TG111 { get; set; }
        public string TG112 { get; set; }
        public string TG113 { get; set; }
        public string TG114 { get; set; }
        public string TG115 { get; set; }
        public string TG116 { get; set; }
        public string TG117 { get; set; }
        public string TG118 { get; set; }
        public string TG119 { get; set; }
        public string TG120 { get; set; }
        public string TG121 { get; set; }
        public string TG122 { get; set; }
        public string TG123 { get; set; }
        public string TG124 { get; set; }
        public string TG125 { get; set; }
        public string TG126 { get; set; }
        public string TG127 { get; set; }
        public string TG128 { get; set; }
        public string TG129 { get; set; }
        public string TG130 { get; set; }
        public string TG131 { get; set; }
        public string TG132 { get; set; }
        public string TG133 { get; set; }
        public string UDF01 { get; set; }
        public string UDF02 { get; set; }
        public string UDF03 { get; set; }
        public string UDF04 { get; set; }
        public string UDF05 { get; set; }
        public double? UDF06 { get; set; }
        public double? UDF07 { get; set; }
        public double? UDF08 { get; set; }
        public double? UDF09 { get; set; }
        public double? UDF10 { get; set; }
        public string CFIELD01 { get; set; }
    }
    #endregion

    #region //ERP - 銷貨單單身 COPTH
    public class COPTH
    {
        public string TransferStatusMES { get; set; }
        public int RoDetailId { get; set; }
        public string COMPANY { get; set; }
        public string CREATOR { get; set; }
        public string USR_GROUP { get; set; }
        public string CREATE_DATE { get; set; }
        public string MODIFIER { get; set; }
        public string MODI_DATE { get; set; }
        public string FLAG { get; set; }
        public string CREATE_TIME { get; set; }
        public string CREATE_AP { get; set; }
        public string CREATE_PRID { get; set; }
        public string MODI_TIME { get; set; }
        public string MODI_AP { get; set; }
        public string MODI_PRID { get; set; }
        public string TH001 { get; set; }
        public string TH002 { get; set; }
        public string TH003 { get; set; }
        public string TH004 { get; set; }
        public string TH005 { get; set; }
        public string TH006 { get; set; }
        public string TH007 { get; set; }
        public double TH008 { get; set; }
        public string TH009 { get; set; }
        public double TH010 { get; set; }
        public string TH011 { get; set; }
        public double TH012 { get; set; }
        public double TH013 { get; set; }
        public string TH014 { get; set; }
        public string TH015 { get; set; }
        public string TH016 { get; set; }
        public string TH017 { get; set; }
        public string TH018 { get; set; }
        public string TH019 { get; set; }
        public string TH020 { get; set; }
        public string TH021 { get; set; }
        public string TH022 { get; set; }
        public string TH023 { get; set; }
        public double TH024 { get; set; }
        public double TH025 { get; set; }
        public string TH026 { get; set; }
        public string TH027 { get; set; }
        public string TH028 { get; set; }
        public string TH029 { get; set; }
        public string TH030 { get; set; }
        public string TH031 { get; set; }
        public string TH032 { get; set; }
        public string TH033 { get; set; }
        public string TH034 { get; set; }
        public double TH035 { get; set; }
        public double TH036 { get; set; }
        public double TH037 { get; set; }
        public double TH038 { get; set; }
        public double TH039 { get; set; }
        public double TH040 { get; set; }
        public string TH041 { get; set; }
        public string TH042 { get; set; }
        public double TH043 { get; set; }
        public double TH044 { get; set; }
        public double TH045 { get; set; }
        public double TH046 { get; set; }
        public double TH047 { get; set; }
        public double TH048 { get; set; }
        public double TH049 { get; set; }
        public double TH050 { get; set; }
        public string TH051 { get; set; }
        public double TH052 { get; set; }
        public double TH053 { get; set; }
        public string TH054 { get; set; }
        public string TH055 { get; set; }
        public string TH056 { get; set; }
        public double TH057 { get; set; }
        public string TH058 { get; set; }
        public string TH059 { get; set; }
        public string TH060 { get; set; }
        public double TH061 { get; set; }
        public string TH062 { get; set; }
        public double TH063 { get; set; }
        public double TH064 { get; set; }
        public string TH065 { get; set; }
        public string TH066 { get; set; }
        public string TH067 { get; set; }
        public string TH068 { get; set; }
        public double TH069 { get; set; }
        public string TH070 { get; set; }
        public double TH071 { get; set; }
        public double TH072 { get; set; }
        public double TH073 { get; set; }
        public string TH074 { get; set; }
        public string TH075 { get; set; }
        public string TH076 { get; set; }
        public string TH077 { get; set; }
        public string TH078 { get; set; }
        public string TH079 { get; set; }
        public double TH080 { get; set; }
        public string TH081 { get; set; }
        public string TH082 { get; set; }
        public string TH083 { get; set; }
        public string TH084 { get; set; }
        public string TH085 { get; set; }
        public string TH086 { get; set; }
        public string TH087 { get; set; }
        public string TH088 { get; set; }
        public string TH089 { get; set; }
        public string TH090 { get; set; }
        public double TH091 { get; set; }
        public string TH092 { get; set; }
        public string TH093 { get; set; }
        public string TH094 { get; set; }
        public string TH500 { get; set; }
        public string TH501 { get; set; }
        public string TH502 { get; set; }
        public string TH503 { get; set; }
        public string UDF01 { get; set; }
        public string UDF02 { get; set; }
        public string UDF03 { get; set; }
        public string UDF04 { get; set; }
        public string UDF05 { get; set; }
        public double? UDF06 { get; set; }
        public double? UDF07 { get; set; }
        public double? UDF08 { get; set; }
        public double? UDF09 { get; set; }
        public double? UDF10 { get; set; }
    }
    #endregion

    #endregion

    #region //銷退單

    #region //MES - 銷退單單頭 SCM.ReturnReceiveOrder
    public class ReturnReceiveOrder
    {
        public int? RtId { get; set; }
        public int? CompanyId { get; set; }
        public string RtErpPrefix { get; set; }    //TI001  單別
        public string RtErpNo { get; set; }    //TI002  單號
        public DateTime? ReturnDate { get; set; }  //TI003  銷退日
        public int? CustomerId { get; set; }
        public string CustomerNo { get; set; } //TI004  客戶
        public int? DepartmentId { get; set; }
        public string DepartmentNo { get; set; }   //TI005  部門	
        public int? SalesmenId { get; set; }
        public string SalesmenNo { get; set; } //TI006  業務員
        public string Factory { get; set; }    //TI007  廠別
        public string Currency { get; set; }   //TI008  幣別	
        public double ExchangeRate { get; set; }   //TI009  匯率	
        public double OriginalAmount { get; set; } //TI010  原幣銷退金額
        public double OriginalTaxAmount { get; set; }  //TI011  原幣銷退稅額	
        public string InvoiceType { get; set; }    //TI012  發票聯數
        public string TaxType { get; set; }    //TI013  課稅別
        public string InvNumEnd { get; set; }  //TI014  發票號碼(訖)	
        public string UiNo { get; set; }   //TI015  統一編號	
        public int? PrintCount { get; set; }   //TI016  列印次數
        public DateTime? InvoiceDate { get; set; } //TI017  發票日期
        public string UpdateCode { get; set; } //TI018  更新碼
        public string ConfirmStatus { get; set; }  //TI019  確認碼	
        public string Remark { get; set; } //TI020  備註
        public string CustomerFullName { get; set; }   //TI021  客戶全名	
        public int? CollectionSalesmenId { get; set; }
        public string CollectionSalesmenNo { get; set; }   //TI022  收款業務員	
        public string Remark1 { get; set; }    //TI023  備註一
        public string Remark2 { get; set; }    //TI024  備註二
        public string Remark3 { get; set; }    //TI025  備註三
        public string DeductionDistinction { get; set; }   //TI026  扣抵區分	
        public string CustomsClearance { get; set; }   //TI027  通關方式	
        public int? RowCnt { get; set; }   //TI028  件數
        public double TotalQuantity { get; set; }  //TI029  總數量	
        public int? StaffId { get; set; }
        public string StaffNo { get; set; }    //TI030  員工代號
        public string RevenueJournalCode { get; set; } //TI031  產生分錄碼(收入)
        public string CostJournalCode { get; set; }    //TI032  產生分錄碼(成本)
        public string ApplyYYMM { get; set; }  //TI033  申報年月	
        public DateTime? DocDate { get; set; } //TI034  單據日期
        public int? ConfirmUserId { get; set; }
        public string ConfirmUserNo { get; set; }  //TI035  確認者	
        public double TaxRate { get; set; }    //TI036  營業稅率
        public double PretaxAmount { get; set; }   //TI037  本幣銷退金額	
        public double TaxAmount { get; set; }  //TI038  本幣銷退稅額	
        public string PaymentTerm { get; set; }    //TI039  付款條件代號
        public double TotalPackages { get; set; }  //TI040  總包裝數量	
        public string SignatureStatus { get; set; }    //TI041  簽核狀態碼
        public string ProcessCode { get; set; }    //TI042  流程代號
        public string TransferStatusERP { get; set; }  //TI043  拋轉狀態	
        public string BondCode { get; set; }   //TI044  保稅碼	
        public int? TransmissionCount { get; set; }    //TI045  傳送次數	
        public string CustomerAddressFirst { get; set; }   //TI046  取貨地址一	
        public string CustomerAddressSecond { get; set; }    //TI047  取貨地址二		
        public string ContactPerson { get; set; }  //TI049  連絡人	
        public string Courier { get; set; }    //TI050  代送商代號
        public string SiteCommCalcMethod { get; set; } //TI051  營站佣金計算
        public double SiteCommRate { get; set; }   //TI052  營站佣金比率	
        public string CommCalcInclTax { get; set; }    //TI053  佣金計算含稅
        public double TotalCommAmount { get; set; }    //TI054  佣金總金額
        public string TradingTerms { get; set; }   //TI057  交易條件	
        public string CustomerEgFullName { get; set; } //TI058  客戶英文全名
        public string DebitNote { get; set; }  //TI061  折讓證明單	
        public string TaxCode { get; set; }    //TI064  稅別碼
        public string RemarkCode { get; set; } //TI066  註記號
        public string MultipleInvoices { get; set; }   //TI067  多發票	
        public string InvNumStart { get; set; }    //TI068  發票號碼(起)
        public int? NumberOfInvoices { get; set; } //TI069  發票張數	
        public string MultiTaxRate { get; set; }   //TI070  單身多稅率	
        public double TaxCurrencyRate { get; set; }    //TI071  稅幣匯率
        public DateTime? VoidDate { get; set; }    //TI074  作廢日期	
        public string VoidTime { get; set; }    //TI075  作廢時間	
        public string VoidRemark { get; set; } //TI076  作廢原因
        public string IncomeDraftID { get; set; }  //TI077  收入底稿編號	
        public string IncomeDraftSeq { get; set; } //TI078  收入底稿序號
        public string IncomeVoucherType { get; set; }  //TI079  收入傳票單別	
        public string IncomeVoucherNumber { get; set; }    //TI080  收入傳票單號
        public string Status { get; set; }   //TI081  狀態	
        public string GenLedgerAcctType { get; set; }  //TI086  總帳帳別	
        public string InvCode2 { get; set; }   //TI087  發票代號	
        public string InvSymbol { get; set; }  //TI088  發票記號	
        public string ContactEmail { get; set; }   //TI091  連絡人EMAIL	
        public string OrigInvNumber { get; set; }  //TI092  原發票號碼  	
        public string TransferStatusMES { get; set; }  //MES拋轉ERP狀態	
        public DateTime? TransferTime { get; set; }
        public DateTime? ConfirmTime { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
    }
    #endregion

    #region //MES - 銷退單單身 SCM.RtDetail
    public class RtDetail
    {
        public int? RtDetailId { get; set; }           //
        public int? RtId { get; set; }           //銷貨單單頭Id
        public string RtErpPrefix { get; set; }           //TJ001 單別
        public string RtErpNo { get; set; }           //TJ002 編號
        public string RtSequence { get; set; }           //TJ003 序號
        public int? MtlItemId { get; set; }           //TJ004 品號    
        public string MtlItemNo { get; set; }           //TJ004 品號    
        public double Quantity { get; set; }           //TJ007 數量    
        public int? UomId { get; set; }           //TJ008 單位    
        public string UomNo { get; set; }           //TJ008 單位    
        public double UnitPrice { get; set; }           //TJ011 單價    
        public double Amount { get; set; }           //TJ012 金額    
        public int? InventoryId { get; set; }           //TJ013 退貨庫別    
        public string InventoryNo { get; set; }           //TJ013 退貨庫別    
        public string LotNumber { get; set; }           //TJ014 批號    
        public int? RoDetailId { get; set; }           //      銷貨單單身Id 
        public string RoErpPrefix { get; set; }           //TJ015 銷貨單別 
        public string RoErpNo { get; set; }           //TJ016 銷貨單號 
        public string RoSequence { get; set; }           //TJ017 銷貨序號 
        public int? SoDetailId { get; set; }           //      訂單單身Id 
        public string SoErpPrefix { get; set; }           //TJ019 訂單單別 
        public string SoErpNo { get; set; }           //TJ015 訂單單號 
        public string SoSequence { get; set; }           //TJ020 訂單序號 
        public string ConfirmStatus { get; set; }           //TJ021 確認碼    
        public string UpdateCode { get; set; }           //TJ022 更新碼    
        public string Remarks { get; set; }           //TJ023 備註    
        public string CheckOutCode { get; set; }           //TJ024 結帳碼    
        public string CheckOutPrefix { get; set; }           //TJ025 結帳單別    
        public string CheckOutNo { get; set; }           //TJ026 結帳單號    
        public string CheckOutSeq { get; set; }           //TJ027 結帳序號    
        public string ProjectCode { get; set; }           //TJ028 專案代號    
        public string CustomerMtlItemNo { get; set; }           //TJ029 客戶品號    
        public string Type { get; set; }           //TJ030 類型    
        public double OrigPreTaxAmt { get; set; }           //TJ031 原幣未稅金額    
        public double OrigTaxAmt { get; set; }           //TJ032 原幣稅額    
        public double PreTaxAmt { get; set; }           //TJ033 本幣未稅金額    
        public double TaxAmt { get; set; }           //TJ034 本幣稅額    
        public double PackageQty { get; set; }           //TJ035 包裝數量    
        public int? PackageUomId { get; set; }           //TJ036 包裝單位    
        public string PackageUomNo { get; set; }           //TJ036 包裝單位    
        public string BondedCode { get; set; }           //TJ037 保稅碼    
        public double CommissionRate { get; set; }           //TJ038 佣金比率    
        public double CommissionAmount { get; set; }           //TJ039 佣金金額    
        public string OriginalCustomer { get; set; }           //TJ040 原始客戶    
        public string FreeSpareType { get; set; }           //TJ041 數量類型    
        public double FreeSpareQty { get; set; }           //TJ042 贈/備品量    
        public double PackageFreeSpareQty { get; set; }           //TJ043 贈/備品包裝量    
        public string NotPayTemp { get; set; }           //TJ048 暫不付款    
        public string Location { get; set; }			//TJ049 儲位    
        public double AvailableQty { get; set; }			//TJ050 計價數量    
        public int? AvailableUomId { get; set; }           //TJ051 計價單位    
        public string AvailableUomNo { get; set; }           //TJ051 計價單位    
        public double TaxRate { get; set; }           //TJ058 營業稅率    
        public string TaxCode { get; set; }           //TJ070 稅別碼    
        public double DiscountAmount { get; set; }           //TJ071 折扣金額    
        public string TransferStatusMES { get; set; }           //MES拋轉ERP狀態 (N/Y/R/V) N:未拋轉,Y:已拋轉,R:重新編輯,V:作廢      
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
    }
    #endregion

    #region //ERP - 銷退單單頭 COPTI
    public class COPTI
    {
        public string TransferStatusMES { get; set; }
        public string COMPANY { get; set; }
        public string CREATOR { get; set; }
        public string USR_GROUP { get; set; }
        public string CREATE_DATE { get; set; }
        public string MODIFIER { get; set; }
        public string MODI_DATE { get; set; }
        public string FLAG { get; set; }
        public string CREATE_TIME { get; set; }
        public string CREATE_AP { get; set; }
        public string CREATE_PRID { get; set; }
        public string MODI_TIME { get; set; }
        public string MODI_AP { get; set; }
        public string MODI_PRID { get; set; }
        public string TI001 { get; set; }       //單別
        public string TI002 { get; set; }       //單號
        public string TI003 { get; set; }       //銷退日
        public string TI004 { get; set; }       //客戶
        public string TI005 { get; set; }       //部門
        public string TI006 { get; set; }       //業務員
        public string TI007 { get; set; }       //廠別
        public string TI008 { get; set; }       //幣別
        public double TI009 { get; set; }       //匯率
        public double TI010 { get; set; }       //原幣銷退金額
        public double TI011 { get; set; }       //原幣銷退稅額
        public string TI012 { get; set; }       //發票聯數
        public string TI013 { get; set; }       //課稅別
        public string TI014 { get; set; }       //發票號碼(訖)
        public string TI015 { get; set; }       //統一編號
        public int TI016 { get; set; }       //列印次數
        public string TI017 { get; set; }       //發票日期
        public string TI018 { get; set; }       //更新碼
        public string TI019 { get; set; }       //確認碼
        public string TI020 { get; set; }       //備註
        public string TI021 { get; set; }       //客戶全名
        public string TI022 { get; set; }       //收款業務員
        public string TI023 { get; set; }       //備註一
        public string TI024 { get; set; }       //備註二
        public string TI025 { get; set; }       //備註三
        public string TI026 { get; set; }       //扣抵區分
        public string TI027 { get; set; }       //通關方式
        public int TI028 { get; set; }       //件數
        public double TI029 { get; set; }       //總數量
        public string TI030 { get; set; }       //員工代號
        public string TI031 { get; set; }       //產生分錄碼(收入)
        public string TI032 { get; set; }       //產生分錄碼(成本)
        public string TI033 { get; set; }       //申報年月
        public string TI034 { get; set; }       //單據日期
        public string TI035 { get; set; }       //確認者
        public double TI036 { get; set; }       //營業稅率
        public double TI037 { get; set; }       //本幣銷退金額
        public double TI038 { get; set; }       //本幣銷退稅額
        public string TI039 { get; set; }       //付款條件代號
        public double TI040 { get; set; }       //總包裝數量
        public string TI041 { get; set; }       //簽核狀態碼
        public string TI042 { get; set; }       //流程代號
        public string TI043 { get; set; }       //拋轉狀態
        public string TI044 { get; set; }       //保稅碼
        public int TI045 { get; set; }       //傳送次數
        public string TI046 { get; set; }       //取貨地址一
        public string TI047 { get; set; }       //取貨地址二
        public int TI048 { get; set; }       //折讓列印次數
        public string TI049 { get; set; }       //連絡人
        public string TI050 { get; set; }       //代送商代號
        public string TI051 { get; set; }       //營站佣金計算方式
        public double TI052 { get; set; }       //營站佣金比率
        public string TI053 { get; set; }       //佣金計算含稅否
        public double TI054 { get; set; }       //佣金總金額
        public string TI055 { get; set; }       //EBstring銷退單號
        public string TI056 { get; set; }       //EBstring銷退版次
        public string TI057 { get; set; }       //交易條件
        public string TI058 { get; set; }       //客戶英文全名
        public double TI059 { get; set; }       //預留欄位
        public double TI060 { get; set; }       //預留欄位
        public string TI061 { get; set; }       //折讓證明單
        public string TI062 { get; set; }       //來源
        public string TI063 { get; set; }       //預留欄位
        public string TI064 { get; set; }       //稅別碼
        public string TI065 { get; set; }       //海關手冊
        public string TI066 { get; set; }       //註記號
        public string TI067 { get; set; }       //多發票
        public string TI068 { get; set; }       //發票號碼(起)
        public int TI069 { get; set; }       //發票張數
        public string TI070 { get; set; }       //單身多稅率
        public double TI071 { get; set; }       //稅幣匯率
        public int TI072 { get; set; }       //發票註記碼長度
        public int TI073 { get; set; }       //發票流水碼長度
        public string TI074 { get; set; }       //作廢日期
        public string TI075 { get; set; }       //作廢時間
        public string TI076 { get; set; }       //作廢原因
        public string TI077 { get; set; }       //收入底稿編號
        public string TI078 { get; set; }       //收入底稿序號
        public string TI079 { get; set; }       //收入傳票單別
        public string TI080 { get; set; }       //收入傳票單號
        public string TI081 { get; set; }       //狀態
        public string TI082 { get; set; }       //預留欄位
        public string TI083 { get; set; }       //預留欄位
        public string TI084 { get; set; }       //預留欄位
        public string TI085 { get; set; }       //預留欄位
        public string TI086 { get; set; }       //總帳帳別
        public string TI087 { get; set; }       //發票代號
        public string TI088 { get; set; }       //發票記號
        public string TI089 { get; set; }       //賣方開立折讓單
        public string TI090 { get; set; }       //折讓單簽回日
        public string TI091 { get; set; }       //連絡人EMAIL
        public string TI092 { get; set; }       //原發票號碼
        public string UDF01 { get; set; }
        public string UDF02 { get; set; }
        public string UDF03 { get; set; }
        public string UDF04 { get; set; }
        public string UDF05 { get; set; }
        public double? UDF06 { get; set; }
        public double? UDF07 { get; set; }
        public double? UDF08 { get; set; }
        public double? UDF09 { get; set; }
        public double? UDF10 { get; set; }
        public string CFIELD01 { get; set; }
    }
    #endregion

    #region //ERP - 銷退單單身 COPTJ
    public class COPTJ
    {
        public string TransferStatusMES { get; set; }
        public int RtDetailId { get; set; }
        public string COMPANY { get; set; }
        public string CREATOR { get; set; }
        public string USR_GROUP { get; set; }
        public string CREATE_DATE { get; set; }
        public string MODIFIER { get; set; }
        public string MODI_DATE { get; set; }
        public string FLAG { get; set; }
        public string CREATE_TIME { get; set; }
        public string CREATE_AP { get; set; }
        public string CREATE_PRID { get; set; }
        public string MODI_TIME { get; set; }
        public string MODI_AP { get; set; }
        public string MODI_PRID { get; set; }
        public string TJ001 { get; set; }      //單別
        public string TJ002 { get; set; }      //單號
        public string TJ003 { get; set; }      //序號
        public string TJ004 { get; set; }      //品號
        public string TJ005 { get; set; }      //品名
        public string TJ006 { get; set; }      //規格
        public double TJ007 { get; set; }      //數量
        public string TJ008 { get; set; }      //單位
        public double TJ009 { get; set; }      //庫存數量
        public string TJ010 { get; set; }      //小單位
        public double TJ011 { get; set; }      //單價
        public double TJ012 { get; set; }      //金額
        public string TJ013 { get; set; }      //退貨庫別
        public string TJ014 { get; set; }      //批號
        public string TJ015 { get; set; }      //銷貨單別
        public string TJ016 { get; set; }      //銷貨單號
        public string TJ017 { get; set; }      //銷貨序號
        public string TJ018 { get; set; }      //訂單單別
        public string TJ019 { get; set; }      //訂單單號
        public string TJ020 { get; set; }      //訂單序號
        public string TJ021 { get; set; }      //確認碼
        public string TJ022 { get; set; }      //更新碼
        public string TJ023 { get; set; }      //備註
        public string TJ024 { get; set; }      //結帳碼
        public string TJ025 { get; set; }      //結帳單別
        public string TJ026 { get; set; }      //結帳單號
        public string TJ027 { get; set; }      //結帳序號
        public string TJ028 { get; set; }      //專案代號
        public string TJ029 { get; set; }      //客戶品號
        public string TJ030 { get; set; }      //類型
        public double TJ031 { get; set; }      //原幣未稅金額
        public double TJ032 { get; set; }      //原幣稅額
        public double TJ033 { get; set; }      //本幣未稅金額
        public double TJ034 { get; set; }      //本幣稅額
        public double TJ035 { get; set; }      //包裝數量
        public string TJ036 { get; set; }      //包裝單位
        public string TJ037 { get; set; }      //保稅碼
        public double TJ038 { get; set; }      //佣金比率
        public double TJ039 { get; set; }      //佣金金額
        public string TJ040 { get; set; }      //原始客戶
        public string TJ041 { get; set; }      //數量類型
        public double TJ042 { get; set; }      //贈/備品量
        public double TJ043 { get; set; }      //贈/備品包裝量
        public string TJ044 { get; set; }      //EBstring銷退單號
        public string TJ045 { get; set; }      //EBstring銷退版次
        public string TJ046 { get; set; }      //EBstring銷退序號
        public double TJ047 { get; set; }      //產品序號數量
        public string TJ048 { get; set; }      //暫不付款
        public string TJ049 { get; set; }      //儲位
        public double TJ050 { get; set; }      //計價數量
        public string TJ051 { get; set; }      //計價單位
        public string TJ052 { get; set; }      //銷退原因代號
        public double TJ053 { get; set; }      //預留欄位
        public double TJ054 { get; set; }      //預留欄位
        public string TJ055 { get; set; }      //預留欄位
        public string TJ056 { get; set; }      //預留欄位
        public string TJ057 { get; set; }      //預留欄位
        public double TJ058 { get; set; }      //營業稅率
        public string TJ059 { get; set; }      //stringRM可扣抵
        public double TJ060 { get; set; }      //stringRM扣抵量
        public string TJ061 { get; set; }      //stringRM扣抵單位
        public string TJ062 { get; set; }      //借方科目
        public string TJ063 { get; set; }      //貸方科目
        public string TJ064 { get; set; }      //稅額科目
        public string TJ065 { get; set; }      //預留欄位
        public string TJ066 { get; set; }      //預留欄位
        public string TJ067 { get; set; }      //預留欄位
        public string TJ068 { get; set; }      //預留欄位
        public string TJ069 { get; set; }      //業務品號
        public string TJ070 { get; set; }      //稅別碼
        public double TJ071 { get; set; }      //折扣金額
        public string TJ500 { get; set; }      //刻號/BIdouble記錄
        public string TJ501 { get; set; }      //刻號管理
        public string TJ502 { get; set; }      //以包裝單位計價
        public string TJ503 { get; set; }      //DATEstringODE
        public string UDF01 { get; set; }
        public string UDF02 { get; set; }
        public string UDF03 { get; set; }
        public string UDF04 { get; set; }
        public string UDF05 { get; set; }
        public double? UDF06 { get; set; }
        public double? UDF07 { get; set; }
        public double? UDF08 { get; set; }
        public double? UDF09 { get; set; }
        public double? UDF10 { get; set; }
    }
    #endregion

    #endregion

    #region  //批號
    public class LotBarcode
    {
        public string BarcodeNo { get; set; }
        public int BarcodeCount { get; set; }
    }
    #endregion

    #region //批量請購單
    public class BatchPurchaseRequisition
    {
        public int? PrDetailId { get; set; }
        public double? PrQty { get; set; }
        public double? PrUnitPrice { get; set; }
        public double? PrPrice { get; set; }
        public string PrErpPrefix { get; set; }
        public string PrErpNo { get; set; }
        public string DocDate { get; set; }
        public string DemandDate { get; set; }
        public string DepartmentFullNo { get; set; }
        public int? MtlItemId { get; set; }
        public string MtlItemNo { get; set; }
        public string MtlItemName { get; set; }
        public string MtlItemSpec { get; set; }
        public int? InventoryId { get; set; }
        public string InventoryNo { get; set; }
        public string InventoryName { get; set; }
        public int? UomId { get; set; }
        public string UomNo { get; set; }
        public string UomName { get; set; }
        public string SoErpFullNo { get; set; }
        public string SupplierFullNo { get; set; }
        public double? InventoryQty { get; set; }
        public double? PoQty { get; set; }
        public string PrCurrency { get; set; }
    }
    #endregion

    #region  //報價單
    public class QuotationHead
    {
        public string ColNo { get; set; }
        public string ColName { get; set; }
        public string Charge { get; set; }
        public string ColRemark { get; set; }
        public string UseFlag { get; set; }
        public List<QuotationDetail> DetailData { get; set; }
    }
    public class QuotationDetail
    {
        public string Sort { get; set; }
        public string DetailNo { get; set; }
        public string DetailName { get; set; }
        public string ColumnSetting { get; set; }
        public string Formula { get; set; }
        public string DataFrom { get; set; }
        public string ElementRemark { get; set; }
        public string Flag { get; set; }
    }

    public class QuotationItem
    {
        public int QiId { get; set; }
        public int RfqDetailId { get; set; }
        public string HtmlColumns { get; set; }
        public string HtmlColName { get; set; }
        public string Charge { get; set; }
        public int? ConfirmUserId { get; set; }
        public string ConfirmStatus { get; set; }
        public string ColRemark { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int CreateBy { get; set; }
        public int LastModifiedBy { get; set; }
        public List<QuotationItemElement> DetailData { get; set; }
    }

    public class QuotationItemElement
    {
        public int QieId { get; set; }
        public int QiId { get; set; }
        public int Sort { get; set; }
        public string ItemElementNo { get; set; }
        public string ItemElementName { get; set; }
        public string QuoteValue { get; set; }
        public string ColumnSetting { get; set; }
        public string Formula { get; set; }
        public string DataFrom { get; set; }
        public string ElementRemark { get; set; }
        public string Flag { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int CreateBy { get; set; }
        public int LastModifiedBy { get; set; }
    }
    #endregion

    #region// ERP-品號庫別檔(INVMC)
    public class INVMC
    {
        public string MC001 { get; set; }       //品號
        public string MC002 { get; set; }       //庫別
        public string MC003 { get; set; }       //儲存位置
        public string MC004 { get; set; }       //安全存量
        public string MC005 { get; set; }       //補貨點
        public string MC006 { get; set; }       //經濟批量
        public string MC007 { get; set; }       //庫存數量
        public string MC008 { get; set; }       //庫存金額
        public string MC009 { get; set; }      //標準存貨量
        public string MC010 { get; set; }      //標準週轉率
        public string MC011 { get; set; }      //上次盤點日
        public string MC012 { get; set; }      //最近入庫日
        public string MC013 { get; set; }      //最近出庫日
        public string MC014 { get; set; }      //庫存包裝數量
        public string MC015 { get; set; }      //主要儲位
    }
    #endregion

    #region //包裝揀貨用
    public class DataModel
    {
        public int maxQuantity { get; set; }
        public int bestAchievedQuantity { get; set; }
        public List<Barcode> solution { get; set; }
    }

    public class Solution
    {
        public List<Barcode> barcodes { get; set; }
    }

    public class Barcode
    {
        public int barcodeid { get; set; }
        public int quantity { get; set; }
        public string barcodeno { get; set; }
        public string category { get; set; }
    }
    #endregion

    #region //2024.11.22.新增庫存卡控
    public class TempHoldRecord
    {
        public int TempHoldId { get; set; }
        public string MtlItemNo { get; set; }
        public decimal HoldQty { get; set; }
        public int SoDetailId { get; set; }
        public int DoId { get; set; }
        public string BarCodeNo { get; set; }
        public string InputStatus { get; set; }
        public DateTime CreateDate { get; set; }
        public bool IsReleased { get; set; }
    }

    public class DoHoldItem
    {
        public string MtlItemNo { get; set; }
        public string InputStatus { get; set; }
        public decimal TotalQty { get; set; }
    }
    #endregion

    public class InventoryAgingReport
    {
        public string LA001 { get; set; }
        public string NewMtlItemNo { get; set; }
        public string NewMtlItemName { get; set; }
        public string NewMtlItemSpec { get; set; }
        public string MB002 { get; set; }
        public string MB003 { get; set; }
        public string MB004 { get; set; }
        public string LA009 { get; set; }
        public string MC002 { get; set; }
        public decimal INV_QTY_30D { get; set; }
        public decimal INV_AMT_30D { get; set; }
        public decimal PKG_QTY_30D { get; set; }
        public decimal INV_QTY_30_90D { get; set; }
        public decimal INV_AMT_30_90D { get; set; }
        public decimal PKG_QTY_30_90D { get; set; }
        public decimal INV_QTY_90_180D { get; set; }
        public decimal INV_AMT_90_180D { get; set; }
        public decimal PKG_QTY_90_180D { get; set; }
        public decimal INV_QTY_180_270D { get; set; }
        public decimal INV_AMT_180_270D { get; set; }
        public decimal PKG_QTY_180_270D { get; set; }
        public decimal INV_QTY_270_365D { get; set; }
        public decimal INV_AMT_270_365D { get; set; }
        public decimal PKG_QTY_270_365D { get; set; }
        public decimal INV_QTY_365D_PLUS { get; set; }
        public decimal INV_AMT_365D_PLUS { get; set; }
        public decimal PKG_QTY_365D_PLUS { get; set; }

    }

    #region //SRM專用
    /// <summary>
    /// 供應商<>公司別 對應表
    /// </summary>
    public class SupplierCompany
    {
        public string CompanyNo { get; set; } = "";
        public string NewCompanyNo { get; set; } = "";
        public string CompanyNos { get; set; } = "";
        public string SupplierNo { get; set; } = "";
        public string SupplierNos { get; set; } = "";
    }

    public class PurchaseOrderMaster: EditModel
    {
        public string PoErpPrefix { get; set; }
        public string PoErpNo { get; set; }
        public string PoDate { get; set; }
        public string ShipMethod { get; set; }
        public string ShipMethodName { get; set; }
        public string PoUser { get; set; }
        public string Currency { get; set; }
        public string PoExchangeRate { get; set; }
        public string TaxCode { get; set; }
        public string Taxation { get; set; }
        /// <summary>
        /// TC026 營業稅率
        /// </summary>
        //public double TaxRate { get; set; }
        public decimal TotalAmount { get; set; }
        public decimal TaxAmount { get; set; }
        public string PaymentTermCode { get; set; }
        public string PaymentTerm { get; set; }
        public string FirstAddress { get; set; }
        public string ContactUser { get; set; }
        public string Remark { get; set; }
        public List<NewPoDetail> PoDetailItems { get; set; }
        public PurchaseOrderMaster() 
        {
            PoDetailItems = new List<NewPoDetail>();
        }
    }

    public class NewPoDetail: EditModel
    {
        /// <summary>
        /// TD001 + TD002 + TD003 採購單完整明細單號
        /// </summary>
        public string PoErpPrefixNo { get; set; }
        /// <summary>
        /// TD001 採購單別
        /// </summary>
        public string PoErpPrefix { get; set; }
        /// <summary>
        /// TD002 採購單號
        /// </summary>
        public string PoErpNo { get; set; }
        /// <summary>
        /// TD003 採購單序號
        /// </summary>
        public string PoSeq { get; set; }
        /// <summary>
        /// TD007 交貨庫別
        /// </summary>
        public string Inventory { get; set; }
        /// <summary>
        /// TD008 採購數量
        /// </summary>
        public double PoQty { get; set; }
        /// <summary>
        /// TD009 單位(Unit of Measure) >> 參照 ERP品號換算單位檔(INVMD)
        /// </summary>
        public string PoUnit { get; set; }
        /// <summary>
        /// TD010 採購單價
        /// </summary>
        public decimal PoUnitPrice { get; set; }
        /// <summary>
        /// TD011 採購金額
        /// </summary>
        public decimal PoPrice { get; set; }
        /// <summary>
        /// TD012 預計交貨日
        /// </summary>
        public string PromiseDate { get; set; }
        /// <summary>
        /// TD014 備註
        /// </summary>
        public string Remark { get; set; }
        /// <summary>
        /// TD015 已交數量
        /// </summary>
        public double SiQty { get; set; }
        /// <summary>
        /// TD025 急料 >> BAS.[Type] N[DEF]；Y
        /// </summary>
        public string UrgentMtl { get; set; }
        /// <summary>
        /// TD058 計價數量
        /// </summary>
        public double PoPriceQty { get; set; }
        /// <summary>
        /// TD059 計價單位
        /// </summary>
        public string PoPriceUnit { get; set; }
        /// <summary>
        /// TD060 已交計價數量
        /// </summary>
        public double PoPriceSiQty { get; set; }

        public string Msg { get; set; }
    }

    public class NewPURTD : EditModelERP
    {
        /// <summary>
        /// TD001 採購單別
        /// </summary>
        public string TD001 { get; set; }
        /// <summary>
        /// TD002 採購單號
        /// </summary>
        public string TD002 { get; set; }
        /// <summary>
        /// TD003 採購單序號
        /// </summary>
        public string TD003 { get; set; }
        /// <summary>
        /// TD007 交貨庫別
        /// </summary>
        public string TD007 { get; set; }
        /// <summary>
        /// TD008 採購數量
        /// </summary>
        public double TD008 { get; set; }
        /// <summary>
        /// TD009 單位(Unit of Measure) >> 參照 ERP品號換算單位檔(INVMD)
        /// </summary>
        public string TD009 { get; set; }
        /// <summary>
        /// TD010 採購單價
        /// </summary>
        public decimal TD010 { get; set; }
        /// <summary>
        /// TD011 採購金額
        /// </summary>
        public decimal TD011 { get; set; }
        /// <summary>
        /// TD012 預計交貨日
        /// </summary>
        public string TD012 { get; set; }
        /// <summary>
        /// TD014 備註
        /// </summary>
        public string TD014 { get; set; }
        /// <summary>
        /// TD015 已交數量
        /// </summary>
        public double TD015 { get; set; }
        /// <summary>
        /// TD025 急料 >> BAS.[Type] N[DEF]；Y
        /// </summary>
        public string TD025 { get; set; }
        /// <summary>
        /// TD058 計價數量
        /// </summary>
        public double TD058 { get; set; }
        /// <summary>
        /// TD059 計價單位
        /// </summary>
        public string TD059 { get; set; }
        /// <summary>
        /// TD060 已交計價數量
        /// </summary>
        public double TD060 { get; set; }

        public string Msg { get; set; }
    }

    public class EditModel
    {
        public string COMPANY { get; set; }
        public string CREATOR { get; set; }
        public string USR_GROUP { get; set; }
        public string CREATE_DATE { get; set; }
        public string MODIFIER { get; set; }
        public string MODI_DATE { get; set; }
        public string FLAG { get; set; } = "1";
        public string CREATE_TIME { get; set; }
        public string CREATE_AP { get; set; }
        public string CREATE_PRID { get; set; } = "BMS";
        public string MODI_TIME { get; set; }
        public string MODI_AP { get; set; }
        public string MODI_PRID { get; set; } = "BMS";
        /// <summary>
        /// 公司代號
        /// </summary>
        public string CompanyNo { get; set; }
        /// <summary>
        /// 供應商代號
        /// </summary>
        public string SupplierNo { get; set; }
        /// <summary>
        /// 當前使用者
        /// </summary>
        public string CurrentUser { get; set; }
        /// <summary>
        /// 回傳總筆數
        /// </summary>
        public int TotalCount { get; set; }
        /// <summary>
        /// 使用權杖 (包含使用者資訊/Api權限資訊等) 暫時以工號做為認證
        /// </summary>
        public string SecretKey { get; set; }
        /// <summary>
        /// api server 位置
        /// </summary>
        public string apiUrl { get; set; }
        /// <summary>
        /// 建立時間
        /// </summary>
        public string CreateDate { get; set; }
        /// <summary>
        /// 最後異動時間
        /// </summary>
        public string LastModifiedDate { get; set; }
        /// <summary>
        /// 建立者
        /// </summary>
        public string CreateBy { get; set; }
        /// <summary>
        /// 最後異動者
        /// </summary>
        public string LastModifiedBy { get; set; }
        /// <summary>
        /// 集團供應端對應表
        /// </summary>
        public List<SupplierCompany> SupplierCompany { get; set; }

        public EditModel()
        {
            SupplierCompany = new List<SupplierCompany>();
        }
    }

    public class EditModelERP
    {
        public string COMPANY { get; set; }
        public string CREATOR { get; set; }
        public string USR_GROUP { get; set; }
        public string CREATE_DATE { get; set; }
        public string MODIFIER { get; set; }
        public string MODI_DATE { get; set; }
        public string FLAG { get; set; } = "1";
        public string CREATE_TIME { get; set; }
        public string CREATE_AP { get; set; }
        public string CREATE_PRID { get; set; } = "BMS";
        public string MODI_TIME { get; set; }
        public string MODI_AP { get; set; }
        public string MODI_PRID { get; set; } = "BMS";
        public string UDF01 { get; set; } = "";
        public string UDF02 { get; set; } = "";
        public string UDF03 { get; set; } = "";
        public string UDF04 { get; set; } = "";
        public string UDF05 { get; set; } = "";
        public double? UDF06 { get; set; } = 0;
        public double? UDF07 { get; set; } = 0;
        public double? UDF08 { get; set; } = 0;
        public double? UDF09 { get; set; } = 0;
        public double? UDF10 { get; set; } = 0;
    }
    #endregion

    #region //專案資料
    public class Project
    {
        public int? ProjectId { get; set; }
        public int? CompanyId { get; set; }
        public string ProjectNo { get; set; }
        public string ProjectName { get; set; }
        public string Remark { get; set; }
        public DateTime? EffectiveDate { get; set; }
        public DateTime? ExpirationDate { get; set; }
        public string CloseCode { get; set; }
        public string TransferErpStatus { get; set; }
        public DateTime? TransferErpDate { get; set; }
        public int? CreateBy { get; set; }
        public DateTime? CreateDate { get; set; }
        public int? LastModifiedBy { get; set; }
        public DateTime? LastModifiedDate { get; set; }
    }

    public class ProjectDetail
    {
        public int? ProjectDetailId { get; set; }

        public int? ProjectId { get; set; }

        /// <summary>
        /// 專案類型 1:物料預算 2:人力成本預算 3:製造費用預算
        /// </summary>
        public string ProjectType { get; set; }

        /// <summary>
        /// 費用幣別
        /// </summary>
        public string Currency { get; set; }

        /// <summary>
        /// 匯率
        /// </summary>
        public double? ExchangeRate { get; set; }

        /// <summary>
        /// 預算費用
        /// </summary>
        public double? BudgetAmount { get; set; }

        /// <summary>
        /// 預算費用(本幣)
        /// </summary>
        public double? LocalBudgetAmount { get; set; }

        /// <summary>
        /// 版本(第一版預設為0000)
        /// </summary>
        public string Edition { get; set; }

        /// <summary>
        /// 拋轉BPM狀態 Y/N
        /// </summary>
        public string BpmTransferStatus { get; set; }

        /// <summary>
        /// 專案預算狀態
        /// N:新單據 P:BPM簽核中 E:BPM駁回修改 Y:已核准 
        /// C:預算變更BPM簽核中 K:預算變更駁回修改中
        /// </summary>
        public string Status { get; set; } = "N";

        /// <summary>
        /// 備註
        /// </summary>
        public string Remark { get; set; }

        public DateTime? CreateDate { get; set; }

        public DateTime? LastModifiedDate { get; set; }

        public int? CreateBy { get; set; }

        public int? LastModifiedBy { get; set; }
    }
    #endregion

    #region //供應商出貨單
    public class DeliveryOrder
    {
        public int DoId { get; set; }
        public int CompanyId { get; set; }
        public int SupplierGroupId { get; set; }
        public string DoNo { get; set; }
        public string PoErpPrefix { get; set; }
        public string PoErpNo { get; set; }
        public string PoErpFullNo { get; set; }
        public DateTime? DoDate { get; set; }
        public DateTime DocDate { get; set; }
        public string DoAddressFirst { get; set; }
        public string DoAddressSecond { get; set; }
        public string InvoiceType { get; set; }
        public string InvoiceNo { get; set; }
        public DateTime? InvoiceDate { get; set; }
        public string PaymentType { get; set; }
        public string PaymentTerm { get; set; }
        public string Currency { get; set; }
        public double ShippingFree { get; set; }
        public double Amount { get; set; }
        public string Remark { get; set; }
        public string SupplierNo { get; set; }
    }

    public class DoDetail
    {
        /// <summary>
        /// 出貨單身id
        /// </summary>
        public int DoDetailId { get; set; }
        /// <summary>
        /// 出貨單id
        /// </summary>
        public int DoId { get; set; }
        /// <summary>
        /// 出貨序號
        /// </summary>
        public string DoSequence { get; set; }
        /// <summary>
        /// 出貨計劃id
        /// </summary>
        public int DpId { get; set; }
        /// <summary>
        /// 批號出貨id
        /// </summary>
        public int BatchId { get; set; }
        /// <summary>
        /// 出貨數量
        /// </summary>
        public double Quantity { get; set; }
        /// <summary>
        /// 贈品量(供)
        /// </summary>
        public double? FreebieQty { get; set; }
        /// <summary>
        /// 備品量(供)
        /// </summary>
        public double? SpareQty { get; set; }
        /// <summary>
        /// 單價
        /// </summary>
        public double UnitPrice { get; set; }
        /// <summary>
        /// 稅額
        /// </summary>
        public double SubTax { get; set; }
        /// <summary>
        /// 金額
        /// </summary>
        public double Price { get; set; }
        /// <summary>
        /// 單位
        /// </summary>
        public string Unit { get; set; }
        /// <summary>
        /// 預交日
        /// </summary>
        public DateTime? PromiseDate { get; set; }
        /// <summary>
        /// 備註
        /// </summary>
        public string Remark { get; set; }
        /// <summary>
        /// 確認者名稱
        /// </summary>
        public string ConfirmUserName { get; set; }
        /// <summary>
        /// 採購單別
        /// </summary>
        public string PoErpPrefix { get; set; }
        /// <summary>
        /// 採購單號
        /// </summary>
        public string PoErpNo { get; set; }
        /// <summary>
        /// 採購單身序號
        /// </summary>
        public string PoSeq { get; set; }
        /// <summary>
        /// 採購完整明細單號
        /// </summary>
        public string PoErpPrefixNo { get; set; }
        /// <summary>
        /// 採購品號
        /// </summary>
        public string DoItemNo { get; set; }
        /// <summary>
        /// 採購品名
        /// </summary>
        public string DoItemName { get; set; }
        /// <summary>
        /// 採購規格
        /// </summary>
        public string DoItemSpec { get; set; }
        /// <summary>
        /// 確認碼
        /// </summary>
        public string ConfirmStatus { get; set; }
        /// <summary>
        /// 確認碼名稱
        /// </summary>
        public string ConfirmStatusName { get; set; }
        /// <summary>
        /// 確認者
        /// </summary>
        public string ConfirmUser { get; set; }
        /// <summary>
        /// 供應商名稱
        /// </summary>
        public string SupplierName { get; set; }
        /// <summary>
        /// 客戶名稱
        /// </summary>
        public string CompanyName { get; set; }
    }
    #endregion

    #region //訂單單身-出貨定版用
    public class SoDetailForDelivery
    {
        public int SoDetailId { get; set; }
        public string SoErpPrefix { get; set; }
        public string SoErpNo { get; set; }
        public string SoSequence { get; set; }
        public string PcPromiseDate { get; set; }
        public string PcPromiseTime { get; set; }
        public int CustomerId { get; set; }
        public string CustomerNo { get; set; }
        public string CustomerShortName { get; set; }
        public string SoErpFullNo { get; set; }
        public string MtlItemName { get; set; }
        public double SoQty { get; set; }
        public double SiQty { get; set; }
        public double FreebieQty { get; set; }
        public double SpareQty { get; set; }
        public string SoDetailRemark { get; set; }
        public string ConfirmStatus { get; set; }
        public string ClosureStatus { get; set; }
        public double TempShippingQty { get; set; }
        public double ReturnTempShippingQty { get; set; }
        public double ShippingQty { get; set; }
        public double ReturnShippingQty { get; set; }
        public double TotalDoQty { get; set; }
        public string DoStatus { get; set; }
    }

    public class SoDetailForErpDelivery
    {
        public string OrderStatus { get; set; }
        public string SoErpPrefix { get; set; }
        public string SoErpNo { get; set; }
        public string SoErpFullNo { get; set; }
        public string OrderDate { get; set; }
        public string UserName { get; set; }
        public string OrderDateYearMonth { get; set; }
        public string CustomerShortName { get; set; }
        public string Currency { get; set; }
        public double Exchange { get; set; }
        public string SoSeq { get; set; }
        public string MtlItemNo { get; set; }
        public string MtlItemName { get; set; }
        public string MtlItemSpec { get; set; }
        public string PcPromiseDate { get; set; }
        public double SoQty { get; set; }
        public string UnitPrice { get; set; }
        public double Amount { get; set; }
        public double SiQty { get; set; }
        public double NotSiAmount { get; set; }
        public double InventoryQty { get; set; }
        public double ItemCost { get; set; }
        public double ProcessItemCost { get; set; }
        public double TotalDoQty { get; set; }
        public double TempShippingQty { get; set; }
    }
    #endregion
}