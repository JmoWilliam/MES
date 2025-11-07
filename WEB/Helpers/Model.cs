using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Helpers
{
    public class FileConfig
    {
        public string key { get; set; }
        public string caption { get; set; }
        public string type { get; set; }
        public string filetype { get; set; }
        public float size { get; set; }
        public bool previewAsData { get; set; }
        public string url { get; set; }
        public string downloadUrl { get; set; }
    }

    public class FileModel
    {
        public int? FileId { get; set; }
        public string FileName { get; set; }
        public byte[] FileContent { get; set; }
        public string FileExtension { get; set; }
        public int FileSize { get; set; }
    }

    public class MailConfig
    {
        public string Host { get; set; }
        public int Port { get; set; }
        public int SendMode { get; set; }
        public string From { get; set; }
        public string Subject { get; set; }
        public string Account { get; set; }
        public string Password { get; set; }
        public string MailTo { get; set; }
        public string MailCc { get; set; }
        public string MailBcc { get; set; }
        public string HtmlBody { get; set; }
        public string TextBody { get; set; }
        public List<MailFile> FileInfo { get; set; }
        public string QcFileFlag { get; set; }
    }

    public class MailFile
    {
        public string FileName { get; set; }
        public string FileExtension { get; set; }
        public byte[] FileContent { get; set; }
    }

    public class SqlQuery
    {
        public string mainKey { get; set; }
        public string auxKey { get; set; }
        public string columns { get; set; }
        public string mainTables { get; set; }
        public string auxTables { get; set; }
        public string conditions { get; set; }
        public string declarePart { get; set; } = "";
        public string groupBy { get; set; } = "";
        public string orderBy { get; set; } = "";
        public int pageIndex { get; set; } = -1;
        public int pageSize { get; set; } = 0;
        public bool distinct { get; set; } = false;
        public string aliasKey { get; set; } = "";
    }

    public class Notification
    {
        public string UserNo { get; set; }
        public string TriggerFunction { get; set; }
        public string LogTitle { get; set; }
        public string LogContent { get; set; }
        public List<NotificationMode> NotificationModes { get; set; }
    }

    public enum NotificationMode
    {
       [Description("Mail")] Mail,
       [Description("Push")] Push,
       [Description("WorkWeixin")] WorkWeixin
    }

    public class PushNotification
    {
        public string title { get; set; }
        public string body { get; set; }
        public string icon { get; set; } = "/Content/images/icons/icon-96x96.png";
        public string lang { get; set; } = "zh-TW";
        public int[] vibrate { get; set; } = {100, 50, 200};
        public string badge { get; set; } = "/Content/images/icons/icon-96x96.png";
        public string tag { get; set; } = "confirm-notification";
        public bool renotify { get; set; } = false;
        public List<PushNotificationActions> actions { get; set; }
        public List<PushNotificationActionUrls> urls { get; set; }
    }

    public class PushNotificationActions
    {
        public string action { get; set; }
        public string title { get; set; }
    }

    public class PushNotificationActionUrls
    {
        public string action { get; set; }
        public string url { get; set; }
    }

    public class PushNotificationUser
    {
        public int PushSubscriptionId { get; set; }
        public string ApiEndpoint { get; set; }
        public string Keysp256dh { get; set; }
        public string Keysauth { get; set; }
    }

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
}
