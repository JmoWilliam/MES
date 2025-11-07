using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MESDA
{

    #region //AiProcess by Mark, 09/01
    public class AiProcessWrapper
    {
        [JsonProperty("processInfo")]
        public List<AiProcess> ProcessInfo { get; set; }
    }
    public class AiProcess
    {
        [JsonProperty("TempNumber")]
        public int TempNumber { get; set; }

        [JsonProperty("ProcessNo")]
        public string ProcessNo { get; set; }

        [JsonProperty("ProcessAlias")]
        public string ProcessAlias { get; set; }

        [JsonProperty("SortNumber")]
        public int SortNumber { get; set; } //it was string

        [JsonProperty("ProcessCheckStatus")]
        public string ProcessCheckStatus { get; set; }

        [JsonProperty("ProcessCheckType")]
        public string ProcessCheckType { get; set; }

        [JsonProperty("PackageFlag")]
        public string PackageFlag { get; set; }

        [JsonProperty("DisplayStatus")]
        public bool DisplayStatus { get; set; }

        [JsonProperty("NecessityStatus")]
        public bool NecessityStatus { get; set; }

        [JsonProperty("ProcessId")]
        public int ProcessId { get; set; }// NOTE by Mark, 09/01, not to get it from FrontEnd, but to get it by ProcessNo later on!
    }
    #endregion
    #region //品號欄位
    public class MtlItem
    {
        public int? MtlItemId { get; set; }
        public string MtlItemNo { get; set; }
        public int? WoId { get; set; }
        public string WoErpPrefix { get; set; }
        public string WoErpNo { get; set; }
        public int? SubstituteMtlItemId { get; set; }
        public string SubstituteMtlItemNo { get; set; }
    }
    #endregion

    public class WipOrder
    {
        public int? WoId { get; set; }
        public int? CompanyId { get; set; }
        public int? UserId { get; set; }
        public string UserNo { get; set; }
        public string WoErpPrefix { get; set; }
        public string WoErpNo { get; set; }
        public string Version { get; set; }
        public DateTime? DocDate { get; set; }
        public int? MtlItemId { get; set; }
        public string MtlItemNo { get; set; }
        public int? InventoryId { get; set; }
        public string InventoryNo { get; set; }
        public int? UomId { get; set; }
        public string UomNo { get; set; }
        public int? PlanQty { get; set; }
        public int? RequisitionSetQty { get; set; }
        public int? StockInQty { get; set; }
        public int? ScrapQty { get; set; }
        public DateTime? ExpectedStart { get; set; }
        public DateTime? ExpectedEnd { get; set; }
        public DateTime? ActualStart { get; set; }
        public DateTime? ActualEnd { get; set; }
        public DateTime? BomDate { get; set; }
        public string BomVersion { get; set; }
        public string UrgentMtl { get; set; }
        public string Property { get; set; }
        public string PlanLotNo { get; set; }
        public int? SoDetailId { get; set; }
        public string SoErpPrefix { get; set; }
        public string SoErpNo { get; set; }
        public string SoSequence { get; set; }
        public string Project { get; set; }
        public string ProductionLine { get; set; }
        public string Processor { get; set; }
        public string CurrencyCode { get; set; }
        public double Exchange { get; set; }
        public string TaxCode { get; set; }
        public string TaxType { get; set; }
        public double TaxRate { get; set; }
        public string PricingTerm { get; set; }
        public string PaymentTerm { get; set; }
        public string WoRemark { get; set; }
        public string ConfirmStatus { get; set; }
        public int? ConfirmUserId { get; set; }
        public string ConfirmUserNo { get; set; }
        public DateTime? ConfirmDate { get; set; }
        public string WoStatus { get; set; }
        public string TransferStatus { get; set; }
        public DateTime? TransferDate { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
        public string TransactionType { get; set; }
        public string ErpDb { get; set; }
        public string ErpNo { get; set; }
        public int? TransactionUserId { get; set; }
        public DateTime? TransactionDate { get; set; }

    }

    public class WoDetail
    {
        public int? WoDetailId { get; set; }
        public int? WoId { get; set; }
        public int? MtlItemId { get; set; }
        public string MtlItemNo { get; set; }
        public string WoErpPrefix { get; set; }
        public string WoErpNo { get; set; }
        public int? InventoryId { get; set; }
        public string InventoryNo { get; set; }
        public int? UomId { get; set; }
        public string UomNo { get; set; }
        public double? CompositionQuantity { get; set; }
        public double? Base { get; set; }
        public double? LossRate { get; set; }
        public string BomSequence { get; set; }
        public double? DemandRequisitionQty { get; set; }
        public double? RequisitionQty { get; set; }
        public string ProcessNo { get; set; }
        public string Substitute { get; set; }
        public string SubstituteStatus { get; set; }
        public int? SubstituteMtlItemId { get; set; }
        public string SubstituteMtlItemNo { get; set; }
        public string SubstituteProcessNo { get; set; }
        public string MaterialProperties { get; set; }
        public DateTime? ExpectedRequisitionDate { get; set; }
        public DateTime? ActualRequisitionDate { get; set; }
        public string WipDetailRemark { get; set; }
        public string ConfirmStatus { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
    }

    #region //單位欄位
    public class UnitOfMeasure
    {
        public int? UomId { get; set; }
        public string UomNo { get; set; }
        public string ProduceUomId { get; set; }
        public string ProduceUomNo { get; set; }
    }
    #endregion

    #region //庫別欄位
    public class Inventory
    {
        public int? InventoryId { get; set; }
        public string InventoryNo { get; set; }
    }
    #endregion

    #region //供應商欄位
    public class Supplier
    {
        public int? SupplierId { get; set; }
        public string SupplierNo { get; set; }
    }
    #endregion

    #region //使用者欄位
    public class User
    {
        public int? UserId { get; set; }
        public string UserNo { get; set; }
    }
    #endregion

    #region //訂單欄位
    public class SoDetail
    {
        public int? SoDetailId { get; set; }
        public string SoErpPrefix { get; set; }
        public string SoErpNo { get; set; }
        public string SoSequence { get; set; }
    }
    #endregion

    #region //ERP工單單頭欄位
    public class MOCTA
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
        public string TA001 { get; set; }//單別
        public string TA002 { get; set; }//單號
        public string TA003 { get; set; }//開單日期
        public string TA004 { get; set; }//BOM日期
        public string TA005 { get; set; }//BOM版次
        public string TA006 { get; set; }//產品品號
        public string TA007 { get; set; }//單位
        public string TA008 { get; set; }//預留欄位
        public string TA009 { get; set; }//預計開工
        public string TA010 { get; set; }//預計完工
        public string TA011 { get; set; }//狀態碼
        public string TA012 { get; set; }//實際開工
        public string TA013 { get; set; }//確認碼
        public string TA014 { get; set; }//實際完工
        public double? TA015 { get; set; }//預計產量
        public double? TA016 { get; set; }//已領套數
        public double? TA017 { get; set; }//已生產量
        public double? TA018 { get; set; }//報廢數量
        public string TA019 { get; set; }//廠別代號
        public string TA020 { get; set; }//入庫庫別
        public string TA021 { get; set; }//生產線別
        public double? TA022 { get; set; }//加工單價
        public string TA023 { get; set; }//加工單位
        public string TA024 { get; set; }//母製令別
        public string TA025 { get; set; }//母製令編號
        public string TA026 { get; set; }//訂單單別
        public string TA027 { get; set; }//訂單單號
        public string TA028 { get; set; }//訂單序號
        public string TA029 { get; set; }//備註
        public string TA030 { get; set; }//性質
        public int? TA031 { get; set; }//列印次數
        public string TA032 { get; set; }//加工廠商
        public string TA033 { get; set; }//計劃批號
        public string TA034 { get; set; }//產品品名
        public string TA035 { get; set; }//產品規格
        public string TA036 { get; set; }//預計開工
        public string TA037 { get; set; }//預計完工
        public string TA038 { get; set; }//實際開工
        public string TA039 { get; set; }//實際完工
        public string TA040 { get; set; }//確認日
        public string TA041 { get; set; }//確認者
        public string TA042 { get; set; }//幣別
        public double? TA043 { get; set; }//匯率
        public string TA044 { get; set; }//急料
        public double? TA045 { get; set; }//預計產包裝量
        public double? TA046 { get; set; }//已生產包裝量
        public double? TA047 { get; set; }//報廢包裝數量
        public string TA048 { get; set; }//包裝單位
        public string TA049 { get; set; }//簽核狀態碼
        public int? TA050 { get; set; }//傳送次數
        public string TA051 { get; set; }//客戶品號
        public string TA052 { get; set; }//APS規劃製令號碼
        public int? TA053 { get; set; }//產品序號數量
        public string TA054 { get; set; }//類型
        public string TA055 { get; set; }//生管/採購人員
        public string TA056 { get; set; }//課稅別
        public double? TA057 { get; set; }//營業稅率
        public string TA058 { get; set; }//價格條件
        public string TA059 { get; set; }//付款條件代號
        public string TA060 { get; set; }//送貨地址(一)
        public string TA061 { get; set; }//送貨地址(二)
        public string TA062 { get; set; }//版次
        public string TA063 { get; set; }//預計批號
        public double? TA064 { get; set; }//預留欄位
        public double? TA065 { get; set; }//預留欄位
        public string TA066 { get; set; }//MES拋轉紀錄碼
        public string TA067 { get; set; }//預留欄位
        public string TA068 { get; set; }//預留欄位
        public double? TA069 { get; set; }//原幣加工金額
        public string TA070 { get; set; }//稅別碼
        public string TA071 { get; set; }//Marking代號
        public string TA072 { get; set; }//計價方式
        public double? TA073 { get; set; }//不良品單價
        public double? TA074 { get; set; }//廢品單價
        public string TA075 { get; set; }//製造廠商
        public double? TA076 { get; set; }//計價單價
        public double? TA077 { get; set; }//驗收單價
        public string TA078 { get; set; }//進貨/託外進貨單別
        public string TA079 { get; set; }//進貨/託外進貨單號
        public string TA080 { get; set; }//進貨/託外進貨序號
        public string TA081 { get; set; }//採購限制日期
        public string TA082 { get; set; }//預留欄位
        public string TA083 { get; set; }//預留欄位
        public string TA084 { get; set; }//預留欄位
        public string TA085 { get; set; }//預留欄位
        public string TA086 { get; set; }//預留欄位
        public string TA087 { get; set; }//預留欄位
        public string TA088 { get; set; }//預留欄位
        public string TA089 { get; set; }//預留欄位
        public string TA090 { get; set; }//預留欄位
        public string TA091 { get; set; }//預留欄位
        public string TA092 { get; set; }//製程否
        public string TA093 { get; set; }//預留欄位
        public int? TA094 { get; set; }//預留欄位
        public string TA095 { get; set; }//優先順序
        public string TA096 { get; set; }//預留欄位
        public string TA097 { get; set; }//單身多稅率
        public string TA500 { get; set; }//不良品庫別
        public string TA501 { get; set; }//MARKING
        public string TA502 { get; set; }//BONDING
        public double? TA503 { get; set; }//加工片數
        public double? TA504 { get; set; }//已交片數
        public string TA505 { get; set; }//加工項目代號
        public string TA506 { get; set; }//晶片厚度
        public string TA507 { get; set; }//晶片尺寸
        public string TA508 { get; set; }//領料單單別
        public string TA509 { get; set; }//領料單單號
        public string TA510 { get; set; }//備註說明A
        public string TA511 { get; set; }//備註說明B
        public string TA512 { get; set; }//備註說明C
        public string TA513 { get; set; }//備註說明D
        public double? TA514 { get; set; }//需領用總數
        public double? TA515 { get; set; }//需領用包裝總數
        public double? TA516 { get; set; }//GROSS_DIE總數
        public string TA520 { get; set; }//特性1
        public string TA521 { get; set; }//特性2
        public string TA522 { get; set; }//特性3
        public string TA523 { get; set; }//特性4
        public string TA524 { get; set; }//特性5
        public string TA525 { get; set; }//特性6
        public string TA526 { get; set; }//特性7
        public string TA527 { get; set; }//特性8
        public string TA528 { get; set; }//特性9
        public string TA530 { get; set; }//下站加工項目
        public string TA531 { get; set; }//下站加工廠商
        public string TA532 { get; set; }//下站加工聯絡人
        public string TA533 { get; set; }//下站加工電話
        public string TA534 { get; set; }//送貨地址一
        public string TA535 { get; set; }//送貨地址二
        public string TA550 { get; set; }//性質
        public string TA551 { get; set; }//專案代號
        public double? TA552 { get; set; }//材料單價比率
        public string TA553 { get; set; }//預留欄位
        public string TA554 { get; set; }//自動扣料已轉撥
        public string TA555 { get; set; }//轉撥單別
        public string TA556 { get; set; }//轉撥單號
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
        public string CFIELD01 { get; set; }//單別-單號


    }
    #endregion

    #region //ERP工單單身欄位
    public class MOCTB
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
        public string TB001 { get; set; }//製令單別
        public string TB002 { get; set; }//製令單號
        public string TB003 { get; set; }//材料品號
        public double? TB004 { get; set; }//需領用量
        public double? TB005 { get; set; }//已領用量
        public string TB006 { get; set; }//製程代號
        public string TB007 { get; set; }//單位
        public string TB008 { get; set; }//預留欄位
        public string TB009 { get; set; }//庫別
        public string TB010 { get; set; }//取替代件
        public string TB011 { get; set; }//材料型態
        public string TB012 { get; set; }//材料品名
        public string TB013 { get; set; }//材料規格
        public string TB014 { get; set; }//上階主件品號
        public string TB015 { get; set; }//預計領料
        public string TB016 { get; set; }//實際領料
        public string TB017 { get; set; }//備註
        public string TB018 { get; set; }//確認碼
        public double? TB019 { get; set; }//需領用包裝量
        public double? TB020 { get; set; }//已領用包裝量
        public string TB021 { get; set; }//包裝單位
        public string TB022 { get; set; }//類型
        public string TB023 { get; set; }//被取替代品號
        public double? TB024 { get; set; }//預留欄位
        public string TB025 { get; set; }//被取替代製程
        public double? TB026 { get; set; }//預留欄位
        public string TB027 { get; set; }//來源碼
        public string TB028 { get; set; }//預留欄位
        public double? TB029 { get; set; }//預留欄位
        public double? TB030 { get; set; }//預留欄位
        public double? TB031 { get; set; }//預留欄位
        public string TB032 { get; set; }//預留欄位
        public string TB033 { get; set; }//預留欄位
        public string TB034 { get; set; }//預留欄位
        public string TB035 { get; set; }//發料DATECODE
        public string TB036 { get; set; }//發料儲位
        public string TB037 { get; set; }//加工順序
        public double? TB038 { get; set; }//營業稅率
        public string TB039 { get; set; }//稅別碼
        public string TB500 { get; set; }//DATECODE
        public double? TB501 { get; set; }//GROSS_DIE
        public string TB502 { get; set; }//批號
        public string TB503 { get; set; }//領料單別
        public string TB504 { get; set; }//領料單號
        public string TB505 { get; set; }//領料序號
        public string TB550 { get; set; }//性質
        public string TB551 { get; set; }//入庫批號
        public string TB552 { get; set; }//專案代號
        public string TB553 { get; set; }//刻號/BIN記錄
        public string TB554 { get; set; }//刻號管理
        public string TB555 { get; set; }//預留欄位
        public string TB556 { get; set; }//預留欄位
        public string TB557 { get; set; }//預留欄位
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

    #region //領替料單頭欄位 MaterialRequisition
    public class MaterialRequisition
    {
        public int? MrId { get; set; }
        public int? CompanyId { get; set; }
        public string RequesitionNo { get; set; }
        public string MrErpPrefix { get; set; }
        public string MrErpNo { get; set; }
        public string MrDate { get; set; }
        public string DocDate { get; set; }
        public string DocType { get; set; }
        public string ProductionLine { get; set; }
        public string Remark { get; set; }
        public string JournalStatus { get; set; }
        public string PriorityType { get; set; }
        public string NegativeStatus { get; set; }
        public string SignupStatus { get; set; }
        public string SourceType { get; set; }
        public string TransferStatus { get; set; }
        public string ConfirmStatus { get; set; }
        public int? ConfirmUserId { get; set; }
        public string CreateDate { get; set; }
        public string LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
        public string ConfirmUserNo { get; set; }
        public string CreateUserNo { get; set; }
        public string CreateTime { get; set; }
        public DateTime MrDocDate { get; set; }
        public string ContractManufacturer { get; set; }
    }
    #endregion

    #region //領替料單身設定欄位 MrWipOrder
    public class MrWipOrder
    {
        public int? MrId { get; set; }
        public int? MoId { get; set; }
        public int? WoId { get; set; }
        public string WoErpPrefix { get; set; }
        public string WoErpNo { get; set; }
        public string MrType { get; set; }
        public float Quantity { get; set; }
        public int? InventoryId { get; set; }
        public string SubInventoryCode { get; set; }
        public string RequisitionCode { get; set; }
        public string NegativeStatus { get; set; }
        public string Remark { get; set; }
        public string MaterialCategory { get; set; }
        public string SubinventoryType { get; set; }
        public string LineSeq { get; set; }
        public string CreateDate { get; set; }
        public string LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
        public string InventoryNo { get; set; }
        public string CreateUserNo { get; set; }
        public string CreateTime { get; set; }
    }
    #endregion

    #region //領替料單身欄位 MrDetail
    public class MrDetail
    {
        public int? MrDetailId { get; set; }
        public int? MrId { get; set; }
        public string MrSequence { get; set; }
        public int? MtlItemId { get; set; }
        public decimal? Quantity { get; set; }
        public decimal? ActualQuantity { get; set; }
        public int? UomId { get; set; }
        public string Unit { get; set; }
        public int? InventoryId { get; set; }
        public string SubInventoryCode { get; set; }
        public string ProcessCode { get; set; }
        public string LotNumber { get; set; }
        public int? MoId { get; set; }
        public string WoErpPrefix { get; set; }
        public string WoErpNo { get; set; }
        public string DetailDesc { get; set; }
        public string Remark { get; set; }
        public string MaterialCategory { get; set; }
        public string ConfirmStatus { get; set; }
        public string ProjectCode { get; set; }
        public string BondedStatus { get; set; }
        public string SubstituteStatus { get; set; }
        public string OfficialItemStatus { get; set; }
        public int? SubstitutionId { get; set; }
        public string SubstitutionMtlItemNo { get; set; }
        public string SubstituteProcessCode { get; set; }
        public decimal? SubstituteQty { get; set; }
        public float SubstituteRate { get; set; }
        public string CreateDate { get; set; }
        public string LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
        public string CreateUserNo { get; set; }
        public string MtlItemNo { get; set; }
        public string MtlItemName { get; set; }
        public string MtlItemSpec { get; set; }
        public string CreateTime { get; set; }
        public int? WoSeq { get; set; }
        public string UomNo { get; set; }
        public string UomName { get; set; }
        public string InventoryNo { get; set; }
        public string InventoryName { get; set; }
        public int? BomSubstitutionMtlItemId { get; set; }
        public string BomSubstitutionMtlItemNo { get; set; }
        public string BomSubstitutionMtlItemName { get; set; }
        public string MaterialCategoryNo { get; set; }
        public int? WoDetailId { get; set; }
        public decimal? DemandRequisitionQty { get; set; }
        public decimal? RequisitionQty { get; set; }
        public float? CompositionQuantity { get; set; }
        public string BarcodeCtrl { get; set; }
        public string MainBarcode { get; set; }
        public string ControlType { get; set; }
        public string DocType { get; set; }
        public int? WoId { get; set; }
        public string TotalBarcodeQty { get; set; }
        public string MrErpPrefix { get; set; }
        public string ExcessFlag { get; set; }
        public int? MesQuantity { get; set; }
        public string MoMtlItemNo { get; set; }
        public string MoMtlItemName { get; set; }
        public string MoMtlItemSpec { get; set; }
        public string BillMtlItemNo { get; set; }
        public int? BomMtlItemId { get; set; }
        public string LotManagement { get; set; }
        public string BomBarcodeCtrl { get; set; }
        public string StorageLocation { get; set; }

    }
    #endregion

    #region //MOCTC ERP領退料單單頭欄位
    public class MOCTC
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
        public string TC001 { get; set; }
        public string TC002 { get; set; }
        public string TC003 { get; set; }
        public string TC004 { get; set; }
        public string TC005 { get; set; }
        public string TC006 { get; set; }
        public string TC007 { get; set; }
        public string TC008 { get; set; }
        public string TC009 { get; set; }
        public int? TC010 { get; set; }
        public string TC011 { get; set; }
        public string TC012 { get; set; }
        public string TC013 { get; set; }
        public string TC014 { get; set; }
        public string TC015 { get; set; }
        public string TC016 { get; set; }
        public string TC017 { get; set; }
        public int? TC018 { get; set; }
        public string TC019 { get; set; }
        public string TC020 { get; set; }
        public string TC021 { get; set; }
        public float TC022 { get; set; }
        public float TC023 { get; set; }
        public string TC024 { get; set; }
        public string TC025 { get; set; }
        public string TC026 { get; set; }
        public float? UDF06 { get; set; }
        public float? UDF07 { get; set; }
        public float? UDF08 { get; set; }
        public float? UDF09 { get; set; }
        public float? UDF10 { get; set; }
    }
    #endregion

    #region //MOCTD ERP領退料單單身設定欄位
    public class MOCTD
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
        public string TD001 { get; set; }
        public string TD002 { get; set; }
        public string TD003 { get; set; }
        public string TD004 { get; set; }
        public string TD005 { get; set; }
        public float TD006 { get; set; }
        public string TD007 { get; set; }
        public string TD008 { get; set; }
        public string TD009 { get; set; }
        public string TD010 { get; set; }
        public string TD011 { get; set; }
        public string TD012 { get; set; }
        public string TD013 { get; set; }
        public string TD014 { get; set; }
        public string TD015 { get; set; }
        public string TD016 { get; set; }
        public string TD017 { get; set; }
        public string TD018 { get; set; }
        public string TD019 { get; set; }
        public string TD020 { get; set; }
        public string TD021 { get; set; }
        public string TD022 { get; set; }
        public string TD023 { get; set; }
        public float TD024 { get; set; }
        public float TD025 { get; set; }
        public string TD026 { get; set; }
        public string TD027 { get; set; }
        public string TD028 { get; set; }
        public float? UDF06 { get; set; }
        public float? UDF07 { get; set; }
        public float? UDF08 { get; set; }
        public float? UDF09 { get; set; }
        public float? UDF10 { get; set; }
    }
    #endregion

    #region //MOCTE ERP領退料單單身欄位
    public class MOCTE
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
        public string TE001 { get; set; }
        public string TE002 { get; set; }
        public string TE003 { get; set; }
        public string TE004 { get; set; }
        public int TE005 { get; set; }
        public string TE006 { get; set; }
        public string TE007 { get; set; }
        public string TE008 { get; set; }
        public string TE009 { get; set; }
        public string TE010 { get; set; }
        public string TE011 { get; set; }
        public string TE012 { get; set; }
        public string TE013 { get; set; }
        public string TE014 { get; set; }
        public string TE015 { get; set; }
        public string TE016 { get; set; }
        public string TE017 { get; set; }
        public string TE018 { get; set; }
        public string TE019 { get; set; }
        public string TE020 { get; set; }
        public float? TE021 { get; set; }
        public string TE022 { get; set; }
        public string TE023 { get; set; }
        public string TE024 { get; set; }
        public float? TE025 { get; set; }
        public string TE026 { get; set; }
        public string TE027 { get; set; }
        public string TE028 { get; set; }
        public string TE029 { get; set; }
        public float? TE030 { get; set; }
        public float? TE031 { get; set; }
        public float? TE032 { get; set; }
        public float? TE033 { get; set; }
        public string TE034 { get; set; }
        public string TE035 { get; set; }
        public float? TE036 { get; set; }
        public float? TE037 { get; set; }
        public float? UDF06 { get; set; }
        public float? UDF07 { get; set; }
        public float? UDF08 { get; set; }
        public float? UDF09 { get; set; }
        public float? UDF10 { get; set; }
    }
    #endregion

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
        public string ErpItDate { get; set; }
        public string ErpDocDate { get; set; }
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
        public string StorageLocation { get; set; }
        public string ToStorageLocation { get; set; }
        public string ItRemark { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
    }
    #endregion

    #region //入庫單單頭欄位 MoWarehouseEnry
    public class MoWarehouseEnry
    {
        public int? MweId { get; set; }
        //public int? CompanyId { get; set; }
        public string MweErpPrefix { get; set; }
        public string MweErpNo { get; set; }
        public string DocDate { get; set; }
        public string ReceiptDate { get; set; }
        public string FactoryCode { get; set; }
        public string SupplierId { get; set; }
        public string SupplierNo { get; set; }
        public string SupplierSo { get; set; }
        public string Remark { get; set; }
        public int? Quantity { get; set; }
        public int? PackageQuantity { get; set; }
        public string SupplierPicking { get; set; }
        public string ReserveTaxCode { get; set; }
        public string QcFlag { get; set; }
        public string FlowStatus { get; set; }
        public int? ConfirmStatus { get; set; }
        public string TransferStatus { get; set; }
        public string TransferDate { get; set; }
        public string CreateDate { get; set; }
        public string LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
    }
    #endregion

    #region //入庫單單身欄位 MweDetail
    public class MweDetail
    {
        public int? MweDetailId { get; set; }
        public int? MweId { get; set; }
        public string MweErpPrefix { get; set; }
        public string MweErpNo { get; set; }
        public string MweSeq { get; set; }
        public int? MtlItemId { get; set; }
        public int? MtlItemNo { get; set; }
        public int? ReceiptQty { get; set; }
        public int? UomId { get; set; }
        public int? UomNo { get; set; }
        public int? InventoryId { get; set; }
        public int? InventoryNo { get; set; }
        public string LotNumber { get; set; }
        public string AvailableDate { get; set; }
        public string ReCheckDate { get; set; }
        public int? MoId { get; set; }
        public string WoErpPrefix { get; set; }
        public string WoErpNo { get; set; }
        public string ProcessCode { get; set; }
        public int? ReceiptPackageQty { get; set; }
        public int? AcceptancePackageQty { get; set; }
        public string AcceptanceDate { get; set; }
        public int? AcceptQty { get; set; }
        public int? AvailableQty { get; set; }
        public int? ScriptQty { get; set; }
        public int? ReturnQty { get; set; }
        public string ProjectCode { get; set; }
        public string Overdue { get; set; }
        public string QcStatus { get; set; }
        public string ReturnCode { get; set; }
        public string ConfirmStatus { get; set; }
        public string CloseStatus { get; set; }
        public string ReNewStatus { get; set; }
        public string Remark { get; set; }
        public string ConfirmUserId { get; set; }
        public string ConfirmUserNo { get; set; }
        public string CreateDate { get; set; }
        public string LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
    }
    #endregion

    #region //ERP入庫單單頭欄位 MOCTH
    public class MOCTH
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
        public float? TH008 { get; set; }
        public int? TH009 { get; set; }
        public string TH010 { get; set; }
        public string TH011 { get; set; }
        public string TH012 { get; set; }
        public string TH013 { get; set; }
        public string TH014 { get; set; }
        public string TH015 { get; set; }
        public string TH016 { get; set; }
        public string TH017 { get; set; }
        public float? TH018 { get; set; }
        public float? TH019 { get; set; }
        public float? TH020 { get; set; }
        public float? TH021 { get; set; }
        public float? TH022 { get; set; }
        public string TH023 { get; set; }
        public string TH024 { get; set; }
        public int? TH025 { get; set; }
        public string TH026 { get; set; }
        public float? TH027 { get; set; }
        public string TH028 { get; set; }
        public string TH029 { get; set; }
        public float? TH030 { get; set; }
        public float? TH031 { get; set; }
        public float? TH032 { get; set; }
        public string TH033 { get; set; }
        public float? TH034 { get; set; }
        public string TH035 { get; set; }
        public string TH036 { get; set; }
        public string TH037 { get; set; }
        public int? TH038 { get; set; }
        public float? TH039 { get; set; }
        public float? TH040 { get; set; }
        public string TH041 { get; set; }
        public string TH042 { get; set; }
        public string TH043 { get; set; }
        public string TH044 { get; set; }
        public string TH045 { get; set; }
        public string TH046 { get; set; }
        public float? TH047 { get; set; }
        public int? TH048 { get; set; }
        public int? TH049 { get; set; }
        public string TH050 { get; set; }
        public string TH051 { get; set; }
        public string TH052 { get; set; }
        public string TH053 { get; set; }
        public string TH054 { get; set; }
        public string TH055 { get; set; }
        public string TH056 { get; set; }
        public string TH057 { get; set; }
        public string TH058 { get; set; }
        public string TH059 { get; set; }
        public string TH060 { get; set; }
        public string TH061 { get; set; }
        public string TH062 { get; set; }
        public string TH063 { get; set; }
        public string TH500 { get; set; }
        public string UDF01 { get; set; }
        public string UDF02 { get; set; }
        public string UDF03 { get; set; }
        public string UDF04 { get; set; }
        public string UDF05 { get; set; }
        public int UDF06 { get; set; }
        public int UDF07 { get; set; }
        public int UDF08 { get; set; }
        public int UDF09 { get; set; }
        public int UDF10 { get; set; }
        public string CFIELD01 { get; set; }//單別-單號
        public int CompanyId { get; set; }//單別-單號
        public string ReceiptDate { get; set; }//單據日期

    }
    #endregion

    #region //ERP入庫單單身欄位 MOCTI
    public class MOCTI
    {
        public int MweDetailId { get; set; }
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
        public float? TI007 { get; set; }
        public string TI008 { get; set; }
        public string TI009 { get; set; }
        public string TI010 { get; set; }
        public string TI011 { get; set; }
        public string TI012 { get; set; }
        public string TI013 { get; set; }
        public string TI014 { get; set; }
        public string TI015 { get; set; }
        public float? TI016 { get; set; }
        public float? TI017 { get; set; }
        public string TI018 { get; set; }
        public float? TI019 { get; set; }
        public float? TI020 { get; set; }
        public float? TI021 { get; set; }
        public float? TI022 { get; set; }
        public string TI023 { get; set; }
        public float? TI024 { get; set; }
        public float? TI025 { get; set; }
        public float? TI026 { get; set; }
        public float? TI027 { get; set; }
        public string TI028 { get; set; }
        public string TI029 { get; set; }
        public string TI030 { get; set; }
        public string TI031 { get; set; }
        public string TI032 { get; set; }
        public string TI033 { get; set; }
        public string TI034 { get; set; }
        public string TI035 { get; set; }
        public string TI036 { get; set; }
        public string TI037 { get; set; }
        public string TI038 { get; set; }
        public string TI039 { get; set; }
        public string TI040 { get; set; }
        public string TI041 { get; set; }
        public string TI042 { get; set; }
        public string TI043 { get; set; }
        public float? TI044 { get; set; }
        public float? TI045 { get; set; }
        public float? TI046 { get; set; }
        public float? TI047 { get; set; }
        public string TI048 { get; set; }
        public string TI049 { get; set; }
        public float? TI050 { get; set; }
        public float? TI051 { get; set; }
        public string TI052 { get; set; }
        public string TI053 { get; set; }
        public string TI054 { get; set; }
        public int? TI055 { get; set; }
        public string TI056 { get; set; }
        public string TI057 { get; set; }
        public float? TI058 { get; set; }
        public float? TI059 { get; set; }
        public string TI060 { get; set; }
        public string TI061 { get; set; }
        public string TI062 { get; set; }
        public string TI063 { get; set; }
        public string TI064 { get; set; }
        public float? TI065 { get; set; }
        public float? TI066 { get; set; }
        public float? TI067 { get; set; }
        public string TI068 { get; set; }
        public string TI069 { get; set; }
        public string TI070 { get; set; }
        public float? TI071 { get; set; }
        public string TI072 { get; set; }
        public string TI073 { get; set; }
        public string TI500 { get; set; }
        public string TI501 { get; set; }
        public string TI502 { get; set; }
        public float? TI503 { get; set; }
        public float? TI504 { get; set; }
        public float? TI505 { get; set; }
        public float? TI506 { get; set; }
        public string TI507 { get; set; }
        public string TI550 { get; set; }
        public float? TI551 { get; set; }
        public float? TI552 { get; set; }
        public float? TI553 { get; set; }
        public float? TI554 { get; set; }
        public float? TI555 { get; set; }
        public float? TI556 { get; set; }
        public string TI557 { get; set; }
        public string TI558 { get; set; }
        public string TI559 { get; set; }
        public string TI560 { get; set; }
        public string TI561 { get; set; }
        public string TI562 { get; set; }
        public string TI563 { get; set; }
        public string TI564 { get; set; }
        public string TI565 { get; set; }
        public string TI567 { get; set; }
        public string TI568 { get; set; }
        public string TI569 { get; set; }
        public string TI570 { get; set; }
        public string UDF01 { get; set; }
        public string UDF02 { get; set; }
        public string UDF03 { get; set; }
        public string UDF04 { get; set; }
        public string UDF05 { get; set; }
        public int UDF06 { get; set; }
        public int UDF07 { get; set; }
        public int UDF08 { get; set; }
        public int UDF09 { get; set; }
        public int UDF10 { get; set; }

    }
    #endregion

    #region //MES.OspReceipt MES託外進貨單頭欄位
    public class OspReceipt
    {
        public int? OsrId { get; set; }
        public string OsrErpPrefix { get; set; }
        public string OsrErpNo { get; set; }
        public string ReceiptDate { get; set; }
        public string FactoryCode { get; set; }
        public int? SupplierId { get; set; }
        public string SupplierSo { get; set; }
        public string CurrencyCode { get; set; }
        public float? Exchange { get; set; }
        public int? RowCnt { get; set; }
        public string Remark { get; set; }
        public string UiNo { get; set; }
        public string InvoiceType { get; set; }
        public string InvoiceDate { get; set; }
        public string InvoiceNo { get; set; }
        public string TaxType { get; set; }
        public string DeductType { get; set; }
        public string ReserveFlag { get; set; }
        public double? OrigAmount { get; set; }
        public float? DeductAmount { get; set; }
        public double? OrigTax { get; set; }
        public float? ReceiptAmount { get; set; }
        public int? Quantity { get; set; }
        public string ConfirmStatus { get; set; }
        public string RenewFlag { get; set; }
        public int? PrintCnt { get; set; }
        public string AutoMaterialBilling { get; set; }
        public double? OrigPreTaxAmount { get; set; }
        public string ApplyYYMM { get; set; }
        public string DocumentDate { get; set; }
        public float? TaxRate { get; set; }
        public double? PretaxAmount { get; set; }
        public double? TaxAmount { get; set; }
        public string PaymentTerm { get; set; }
        public int? PackageQuantity { get; set; }
        public string SupplierPicking { get; set; }
        public string ApproveStatus { get; set; }
        public string ReserveTaxCode { get; set; }
        public int? SendCount { get; set; }
        public string NoticeFlag { get; set; }
        public string TaxCode { get; set; }
        public float? TaxExchange { get; set; }
        public string QcFlag { get; set; }
        public string FlowStatus { get; set; }
        public DateTime CreateDate { get; set; }
        public DateTime LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
        public string ConfirmUserNo { get; set; }
        public string ConfirmUserId { get; set; }
        public string SupplierNo { get; set; }
        public string CreateUserNo { get; set; }
        public string CreateTime { get; set; }
        public int? CompanyId { get; set; }
        public string MesDocumentDate { get; set; }
    }
    #endregion

    #region //MES.OspReceiptDetail MES託外進貨單身欄位
    public class OspReceiptDetail
    {
        public int? OsrDetailId { get; set; }
        public int? OsrId { get; set; }
        public int? OspDetailId { get; set; }
        public string OsrErpPrefix { get; set; }
        public string OsrErpNo { get; set; }
        public string OsrSeq { get; set; }
        public int? MtlItemId { get; set; }
        public int? ReceiptQty { get; set; }
        public int? UomId { get; set; }
        public int? InventoryId { get; set; }
        public string LotNumber { get; set; }
        public string AvailableDate { get; set; }
        public string ReCheckDate { get; set; }
        public int? MoId { get; set; }
        public int? MoProcessId { get; set; }
        public string ProcessCode { get; set; }
        public int? ReceiptPackageQty { get; set; }
        public int? AcceptancePackageQty { get; set; }
        public string AcceptanceDate { get; set; }
        public int? AcceptQty { get; set; }
        public int? AvailableQty { get; set; }
        public int? ScriptQty { get; set; }
        public int? ReturnQty { get; set; }
        public int? AvailableUom { get; set; }
        public decimal? OrigUnitPrice { get; set; }
        public decimal? OrigAmount { get; set; }
        public decimal? OrigDiscountAmt { get; set; }
        public decimal? ReceiptExpense { get; set; }
        public string DiscountDescription { get; set; }
        public string ProjectCode { get; set; }
        public string PaymentHold { get; set; }
        public string Overdue { get; set; }
        public string QcStatus { get; set; }
        public string ReturnCode { get; set; }
        public string ConfirmStatus { get; set; }
        public string CloseStatus { get; set; }
        public string ReNewStatus { get; set; }
        public string Remark { get; set; }
        public string CostEntry { get; set; }
        public string ExpenseEntry { get; set; }
        public int? ConfirmUserId { get; set; }
        public string ConfirmUser { get; set; }
        public decimal? OrigPreTaxAmt { get; set; }
        public decimal? OrigTaxAmt { get; set; }
        public decimal? PreTaxAmt { get; set; }
        public decimal? TaxAmt { get; set; }
        public string UrgentMtl { get; set; }
        public string ReserveTaxCode { get; set; }
        public string ApproveStatus { get; set; }
        public string TransferStatus { get; set; }
        public string ProcessStatus { get; set; }
        public DateTime CreateDate { get; set; }
        public DateTime LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
        public string CreateUserNo { get; set; }
        public string CreateTime { get; set; }
        public string MtlItemNo { get; set; }
        public string MtlItemName { get; set; }
        public string MtlItemSpec { get; set; }
        public string UomNo { get; set; }
        public string InventoryNo { get; set; }
        public string WoErpPrefix { get; set; }
        public string WoErpNo { get; set; }
        public string MoStatus { get; set; }
        public string DepartmentName { get; set; }
    }
    #endregion

    #region //MES.ManufactureOrder
    public class ManufactureOrder
    {
        public int? MoId { get; set; }
        public string WoRemark { get; set; }
        public string ExpectedEnd { get; set; }
        public int? WoSeq { get; set; }
        public int? Quantity { get; set; }
        public int? InputQty { get; set; }
        public int? CompleteQty { get; set; }
        public double? ScrapQty { get; set; }
        public string Status { get; set; }
        public string DeliveryProcess { get; set; }
        public string ProjectNo { get; set; }
        public string Remark { get; set; }        
        public int? WoId { get; set; }
        public int? CompanyId { get; set; }
        public string WoErpPrefix { get; set; }
        public string WoErpNo { get; set; }
        public int? PlanQty { get; set; }
        public int? InventoryId { get; set; }
        public double? RequisitionSetQty { get; set; }
        public double? StockInQty { get; set; }
        public int? MtlItemId { get; set; }
        public string MtlItemNo { get; set; }
        public string MtlItemName { get; set; }
        public string Version { get; set; }
        public string MtlItemSpec { get; set; }
        public int InventoryUomId { get; set; }
        public string ModeNo { get; set; }
        public string ModeName { get; set; }
        public int? MoSettingId { get; set; }
        public string LotStatus { get; set; }
        public string StatusNo { get; set; }
        public string StatusName { get; set; }
        public string StatusDesc { get; set; }
        public string MoRoutingId { get; set; }
        public string DeliveryProcessName { get; set; }
        public string LetteringInfo { get; set; }
        public int? QcNoticeId { get; set; }
        public string QcNoticeType { get; set; }
        public string BarcodeNo { get; set; }
        public int? ControlId { get; set; }
        public int? ModeId { get; set; }
        public string ProcessDetail { get; set; }
        public string WoErpFullNo { get; set; }
        public string UserNo { get; set; }
        public string UserName { get; set; }
        public double? ProductionQty { get; set; } //MES製令已生產量
        public double? MoScrapQty { get; set; } //MES製令已報廢數量
        public int? TotalCount { get; set; }
        public string CompanyName { get; set; }
        public string CustomerShortName { get; set; }        
        public string QcTypeStatus { get; set; }
        public string BarcodeCtrl { get; set; }
    }
    #endregion

    #region //Spreadsheet相關Model
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

    #region //ProductQuery
    public class ProductQuery
    {
        public int? MoId { get; set; }
        public int? TotalCount { get; set; }
        public int? WoSeq { get; set; }
        public int? InputQty { get; set; }
        public int? Quantity { get; set; }
        public DateTime? DocDate { get; set; }
        public string WoErpFullNo { get; set; }
        public int StockInQty { get; set; }
        public int PlanQty { get; set; }
        public int? RequisitionSetQty { get; set; }
        public string MtlItemNo { get; set; }
        public string MtlItemName { get; set; }
        public string MtlItemSpec { get; set; }
        public string ProcessDetail { get; set; }
    }
    #endregion

    #region //MonthlySalesDetailReport 資料模型類別
    public class MonthlySalesDetailReport
    {
        public string 公司別 { get; set; }
        public string 單據類別 { get; set; }
        public string 產品分類 { get; set; }
        public string 製程 { get; set; }
        public string BU分類 { get; set; }
        public string 客戶代碼 { get; set; }
        public string 客戶 { get; set; }
        public string 集團交易 { get; set; }
        public string 單別 { get; set; }
        public string 銷貨單號 { get; set; }
        public int 銷貨序號 { get; set; }
        public string 品號 { get; set; }
        public string 品名 { get; set; }
        public string 機種 { get; set; }
        public string 規格 { get; set; }
        public string 銷售日期 { get; set; }
        public string 銷售年月 { get; set; }
        public decimal 銷貨數量 { get; set; }
        public string 單位 { get; set; }
        public string 幣別 { get; set; }
        public decimal 匯率 { get; set; }
        public decimal 單價 { get; set; }
        public decimal 本幣單價 { get; set; }
        public decimal 原幣金額 { get; set; }
        public decimal 本幣金額 { get; set; }
        public decimal 台幣金額 { get; set; }
        public decimal 原幣稅額 { get; set; }
        public decimal 本幣稅額 { get; set; }
        public decimal 單位成本 { get; set; }
        public decimal 銷貨成本 { get; set; }
        public decimal 材料成本 { get; set; }
        public decimal 人工成本 { get; set; }
        public decimal 製費成本 { get; set; }
        public decimal 加工成本 { get; set; }
        public string 結帳碼 { get; set; }
        public decimal 預留1數 { get; set; }
        public decimal 預留2數 { get; set; }
        public string 預留1文 { get; set; }
        public string 預留2文 { get; set; }
        public decimal 毛利 { get; set; }
        public decimal 毛利率 { get; set; }
        public string 業務人員 { get; set; }

        // 用於前端顯示的對應屬性
        public string DocumentType => 單據類別;
        public string SalesDate => 銷售日期;
        public string ProductCategory => 產品分類;
        public string CustomerName => 客戶;
        public string SalesPersonName => 業務人員;
        public string SalesOrderNo => 銷貨單號;
        public string ProductNo => 品號;
        public string ProductName => 品名;
        public decimal SalesQuantity => 銷貨數量;
        public decimal LocalAmount => 本幣金額;
        public decimal TwdAmount => 台幣金額;
        public string CustomerCode => 客戶代碼;
        public string SalesPersonCode => 業務人員;
    }

    public class SalesPersonList
    {
        public string 業務人員 { get; set; }
        public string SalesPersonName { get; set; }

        // 前端兼容性屬性
        public string SalesPersonCode => 業務人員;
    }

    public class CustomerList
    {
        public string 客戶代碼 { get; set; }
        public string 客戶 { get; set; }
        public string CustomerCode { get; set; }
        public string CustomerName { get; set; }
    }

    public class ProductCategoryList
    {
        public string 產品分類 { get; set; }
        public string CategoryCode { get; set; }
        public string CategoryName { get; set; }
    }
    #endregion

    #region //INVMC
    public class INVMC
    {
        public string MtlItemNo { get; set; }
        public string MtlItemName { get; set; }
        public int? InventoryId { get; set; }
        public string InventoryNo { get; set; }
        public string InventoryName { get; set; }
        public string InventorySite { get; set; }
        public double? InventoryQty { get; set; }
        public int? TotalCount { get; set; }
    }
    #endregion

    #region //依製令檢查庫存
    public class CheckInventoryQtyClass
    {
        public string MtlItemNo { get; set; }
        public string MtlItemName { get; set; }
        public int? InventoryId { get; set; }
        public string InventoryNo { get; set; }
        public string InventoryName { get; set; }
        public double? InventoryQty { get; set; }
        public double? MrDetailQty { get; set; }
    }
    #endregion

    #region //MOCTO 製令變更單單頭
    public class MOCTO
    {
        public string TO001 { get; set; } //製令單別
        public string TO002 { get; set; } //製令單號
        public string TO003 { get; set; } //變更版次
        public DateTime? TO004 { get; set; } //變更日期
        public string TO005 { get; set; } //變更日期
        public DateTime? TO006 { get; set; } //新開單日期
        public DateTime? TO007 { get; set; } //新BOM日期
        public string TO008 { get; set; } //新BOM版次
        public string TO009 { get; set; } //新產品品號
        public string TO010 { get; set; } //新單位
        public string TO011 { get; set; } //新欄位預留
    }
    #endregion

    #region //BomSubstitution 取替代料
    public class BomSubstitution
    {
        public int SubstitutionId { get; set; }
        public int BomDetailId { get; set; }
        public string SubstituteStatus { get; set; }
        public string SubstituteStatusNo { get; set; }
        public int MtlItemId { get; set; }
        public string MtlItemNo { get; set; }
        public string MtlItemName { get; set; }
        public string MtlItemSpec { get; set; }
        public double? Quantity { get; set; }
        public int? SortNumber { get; set; }
        public string Precedence { get; set; }
        public DateTime? EffectiveDate { get; set; }
        public DateTime? ExpirationDate { get; set; }
        public string Remark { get; set; }
        public int? InventoryId { get; set; }
        public string InventoryNo { get; set; }
        public string InventoryName { get; set; }
        public int? InventoryQty { get; set; }
        public int? UomId { get; set; }
        public string UomNo { get; set; }
        public string UomName { get; set; }
        public string TotalCount { get; set; }
        public int? BomMtlItemId { get; set; }
        public string BomMtlItemNo { get; set; }
        public string BomMtlItemName { get; set; }
        public string BomMtlItemSpec { get; set; }
    }
    #endregion

    #region //MES.BarcodePrint 批量條碼
    public class BarcodePrint
    {
        public int PrintId { get; set; }
        public int? BarcodeId { get; set; }
        public string BarcodeNo { get; set; }
        public double BarcodeQty { get; set; }
        public int MoSettingId { get; set; }
        public string ParentBarcode { get; set; }
        public int PrintCnt { get; set; }
        public string PrintStatus { get; set; }
        public DateTime CreateDate { get; set; }
        public DateTime LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
    }
    #endregion

    #region //MES.Barcode 條碼
    public class Barcode
    {
        public int BarcodeId { get; set; }
        public string BarcodeNo { get; set; }
        public int CurrentMoProcessId { get; set; }
        public int NextMoProcessId { get; set; }
        public int MoId { get; set; }
        public double BarcodeQty { get; set; }
        public int? ParentBarcode { get; set; }
        public int? BarcodeProcessId { get; set; }
        public string Memo { get; set; }
        public int PrintCount { get; set; }
        public string CurrentProdStatus { get; set; }
        public string BarcodeStatus { get; set; }
        public DateTime CreateDate { get; set; }
        public DateTime LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
    }
    #endregion

    #region //品號製令報廢數報表
    #region //品號
    public class ScrapMain
    {
        public string MtlItemNo { get; set; }//MB001
        public string MtlItemName { get; set; }//MB002
        public string MtlItemspec { get; set; }//MB003
        public List<ScrapMo> ScrapMo { get; set; }

    }
    #endregion

    #region //品號製令
    public class ScrapMo
    {
        public string MtlItemNo { get; set; }//MB001
        public string MtlItemName { get; set; }//MB002
        public string MtlItemspec { get; set; }//MB003
        public string WoErpNo { get; set; }//(b.TA001 + '-' + b.TA002)
        public string MoStatus { get; set; }//TA011
        public decimal ExpectedCount { get; set; }//TA015
        public decimal ReceiveCount { get; set; }//TA016
        public decimal ProducedCount { get; set; }//TA017
        public decimal ScrapCount { get; set; }//TA018
        public List<ScrapMoMtlMerge> ScrapMoMtlMerge { get; set; }
    }
    #endregion

    #region //製令料品合併
    public class ScrapMoMtlMerge
    {
        public List<ScrapMoMtl> ScrapMoMtldata { get; set; }
        public List<ScrapMoReplaceMtl> ScrapMoReplaceMtldata { get; set; }
        public decimal MergeScrapCount { get; set; }//報廢數 -->TA011製令未結案:報廢數=領料數-入庫數/製令已結案:報廢數=預計產量-入庫數-超領數
        public decimal MergeMoScrapCount { get; set; }//製令結案報廢數 --> 紘立
        public string Unit { get; set; }//單位 TB007

    }
    #endregion

    #region //製令原料品
    public class ScrapMoMtl
    {
        public string MtlItemNo { get; set; }//TB003
        public string MtlItemName { get; set; }//TB012
        public string MtlItemspec { get; set; }//TB013
        public string WoErpNo { get; set; }//(a.TB001 + '-' + a.TB002)
        public decimal ExpectedCount { get; set; }//TB004//需領用數
        public decimal ReceiveCount { get; set; }//TB005//已領用數
        public decimal ScrapCount { get; set; }//報廢數 -->TA011製令未結案:報廢數=領料數-入庫數/製令已結案:報廢數=預計產量-入庫數-超領數
        public decimal MoScrapCount { get; set; }//製令結案報廢數 --> 紘立
        public string Unit { get; set; }//單位 TB007

    }
    #endregion

    #region //製令替代料料品
    public class ScrapMoReplaceMtl
    {

        public string MB001 { get; set; }//
        public string MtlItemNo { get; set; }//TB003
        public string MtlItemName { get; set; }//TB012
        public string MtlItemspec { get; set; }//TB013
        public string WoErpNo { get; set; }//(a.TB001 + '-' + a.TB002)
        public decimal ExpectedCount { get; set; }//TB004//需領用數
        public decimal ReceiveCount { get; set; }//TB005//已領用數
        public decimal ScrapCount { get; set; }//報廢數 -->TA011製令未結案:報廢數=領料數-入庫數/製令已結案:報廢數=預計產量-入庫數-超領數
        public decimal MoScrapCount { get; set; }//製令結案報廢數 --> 紘立
        public string Unit { get; set; }//單位 TB007

    }
    #endregion

    #region //製令料品超領報表
    public class ScrapMoExcessMtl
    {
        public string MtlItemNo { get; set; }//
        public string MtlItemName { get; set; }//
        public string MtlItemspec { get; set; }//
        public string WoErpNo { get; set; }//
        public decimal MoExpectedCount { get; set; }//TA015//製令預計產量
        public decimal MoProducedCount { get; set; }//TA017 製令已生產量
        public decimal ExpectedCount { get; set; }//TB004//預計領用量
        public decimal ReceiveCount { get; set; }// TB005//已領用數
        public decimal ExcessCount { get; set; }//TE005超領數
        public decimal ExcessPerson { get; set; }//超領%
        public string Unit { get; set; }//單位 TB007

    }
    #endregion

    #endregion

    #region //業務排名分析
    public class Bar
    {
        public string Company { get; set; }
        public string Sales { get; set; }
        public string SalesNo { get; set; }
        public double Amount { get; set; }
        public double AllAmount { get; set; }
        public double MoonAmount { get; set; }
        public double? IntentAmount { get; set; }
    }

    public class Spline
    {
        public string Money { get; set; }
        public string Sales { get; set; }
        public double Moon01 { get; set; }
        public double Moon02 { get; set; }
        public double Moon03 { get; set; }
        public double Moon04 { get; set; }
        public double Moon05 { get; set; }
        public double Moon06 { get; set; }
        public double Moon07 { get; set; }
        public double Moon08 { get; set; }
        public double Moon09 { get; set; }
        public double Moon10 { get; set; }
        public double Moon11 { get; set; }
        public double Moon12 { get; set; }
    }
    #endregion

    #region //紘立報表
    public class EtergeWoData
    {
        public string MtlItemNo { get; set; }//MB001
        public string MtlItemName { get; set; }//MB002
        public string MtlItemspec { get; set; }//MB003
        public string WoErpNo { get; set; }//(b.TA001 + '-' + b.TA002)
        public string MoStatus { get; set; }//TA011
        public int MoId { get; set; }//TA011


        public List<EtergeWoDataDetail> WoDataDetail { get; set; }
    }

    public class EtergeWoDataDetail
    {
        public string MtlItemNo { get; set; }//MB001
        public string MtlItemName { get; set; }//MB002
        public string MtlItemSpec { get; set; }//MB003
        public string WoErpFullNo { get; set; }//(b.TA001 + '-' + b.TA002)
        public string MoStatus { get; set; }//TA011
        public int MoId { get; set; }//TA011
        public int BarcodeId { get; set; }//TA011
        public int BarcodeProcessId { get; set; }//TA011
        public string BarcodeNo { get; set; }//MB001
        public string ProcessAlias { get; set; }//MB002
        public decimal StationQty { get; set; }//MB003
        public decimal PassQty { get; set; }//(b.TA001 + '-' + b.TA002)
        public decimal NgQty { get; set; }//TA011
        public decimal CycleTime { get; set; }//TA015
        public string StartDate { get; set; }//TA016
        public string FinishDate { get; set; }//TA017
        public string MachineDesc { get; set; }//TA017
        public string StarUser { get; set; }//TA017
        public string FinishUser { get; set; }//TA017
        public string BarcodeStatus { get; set; }//TA017 
        public string ProcessName { get; set; }//MB001CauseNo
        public string CauseNo { get; set; }//MB001CauseNo
        public string LastStopFinish { get; set; }//TA017
        public decimal WaitTime { get; set; }//TA015
        public string isNewFinish { get; set; }//TA017
        public int ROWNUM { get; set; }//TA017
        public string WorkDay { get; set; }//TA017



        public int? TotalCount { get; set; }


    }

    public class ProductManufactureData
    {
        public int MoId { get; set; }
        public string WoErpFullNo { get; set; }
        public string WoSeq { get; set; }

        public int MtlItemId { get; set; }
        public string MtlItemNo { get; set; }
        public string MtlItemName { get; set; }
        public string MtlItemSpec { get; set; }
        public int TotalCount { get; set; }
       
        public List<ProductManufactureDetail> ProcessDetail { get; set; }

    }

    public  class ProductManufactureDetail
    {
        public int MoId { get; set; }
        public int MoProcessId { get; set; }
        public string ProcessId { get; set; }
        public string ProcessAlias { get; set; }

        public int totalStationQty { get; set; }
        public string ProdStatus { get; set; }
        public List<TotalMoProcessProdCount> TotalMoProcessProdCount { get; set; }
        public decimal PassPerson { get; set; }//TA015
    }

    public class MoProcessProdCount
    {
        public int BarcodeProcessId { get; set; }
        public int BarcodeId { get; set; }
        public int StationQty { get; set; }
        public string ProdStatus { get; set; }
    }

    public class TotalMoProcessProdCount
    {
        
        public int StationQty { get; set; }
        public string ProdStatus { get; set; }
    }

    #endregion

    #region //量測相關報表
    #region //ProductionHistory 生產在製
    public class ProductionHistory
    {
        public int? BarcodeProcessId { get; set; }
        public string ProdStatus { get; set; }
        public int? CycleTime { get; set; }
        public int? MoProcessId { get; set; }
        public int? StationQty { get; set; }
        public string CauseNo { get; set; }
        public string CauseName { get; set; }
        public string StartDate { get; set; }
        public string FinishDate { get; set; }
        public int? WoSeq { get; set; }
        public string WoErpFullNo { get; set; }
        public string BarcodeNo { get; set; }
        public int? ProcessId { get; set; }
        public string ProcessAlias { get; set; }
        public string MachineName { get; set; }
        public string MachineDesc { get; set; }
        public string StartUserNo { get; set; }
        public string StartUserName { get; set; }
        public string FinishUserNo { get; set; }
        public string FinishUserName { get; set; }
        public string ItemValue { get; set; }
    }
    #endregion
    #endregion

    #region //點膠機台管理
    public class MachineConsume
    {
        public int?     LotNumberId         { get; set; }
        public int?     MaxDispenseQty      { get; set; }
        public int?     AllowDispenseQty    { get; set; }
        public string OpeningDate         { get; set; }
        public int?     EXPs                { get; set; }
        public string BindDate            { get; set; }
        public string UnBindDate          { get; set; }
        public int?     MachineId           { get; set; }
        public string MachineNo           { get; set; }
        public string   MachineName         { get; set; }
        public string   MachineDesc         { get; set; }
        public string   LotNumberNo         { get; set; }
        public int? BarcodeProcessId    { get; set; }
        public int? MoId                { get; set; }
        public int? BarcodeId           { get; set; }
        public int? MoProcessId         { get; set; }
        public string FinishDate          { get; set; }
        public string BarcodeNo           { get; set; }
        public int? CurrentMoProcessId  { get; set; }
        public int? BarcodeQty          { get; set; }
        public string   WoErpPrefix         { get; set; }
        public string   WoErpNo             { get; set; }
        public string   ProcessAlias        { get; set; }
        public int? MtlItemId           { get; set; }
        public string   MtlItemNo           { get; set; }
        public string   MtlItemName         { get; set; }
    }

    #endregion

    #region //製令串查
    public class MoUnitSearchResult
    {
        public int WoId { get; set; }
        public string WoErpFullNo { get; set; }

        public List<MOCTLRresult> TLResult { get; set; }//托外退貨
        public List<MOCTO2> TOResult { get; set; }//製令變更單

        public List<MoResult> MoResult { get; set; }//製令
        public List<MRResult> MRResult { get; set; }//領退料
        public List<MWEResult> MWEResult { get; set; }//入庫
        public List<OSPresult> OSPresult { get; set; }//託外生產
        public List<OSRresult> OSRresult { get; set; }//託外入庫

    }

    public class MoResult
    {
        public int MoId { get; set; }
        public string WoErpFullNo { get; set; }
        public string WoSeq { get; set; }
        public int MtlItemId { get; set; }
        public string MtlItemNo { get; set; }
        public string MtlItemName { get; set; }
        public int PlanQty { get; set; }
        public string UomNo { get; set; }
        public string CreateDate { get; set; }

       
    }

    #region //入庫單
    public class MWEResult
    {
        public int? MweId { get; set; }
        public string MweFullErpNo { get; set; }
        public string DocDate { get; set; }
        public string TransferStatus { get; set; }
        public string ConfirmStatus { get; set; }
        public int ReceiptQty { get; set; }
        public string UomNo { get; set; }
        public string MtlItemNo { get; set; }
        public string MtlItemName { get; set; }
        public string MweSeq { get; set; }

    }



    #endregion

    #region //領退料
    public class MRResult
    {
        public int? MrId { get; set; }
        public string MrFullErpNo { get; set; }
        public string TransferStatus { get; set; }
        public string DocDate { get; set; }
        public string ConfirmStatus { get; set; }
        public int Quantity { get; set; }
        public string Unit { get; set; }
        public string MtlItemNo { get; set; }
        public string MtlItemName { get; set; }
        public string MrSequence { get; set; }

    }

    #endregion

    #region //託外生產
    public class OSPresult
    {
        public int? OspId { get; set; }
        public string OspNo { get; set; }
        public string Status { get; set; }
        public string OspQty { get; set; }
        public string MtlItemNo { get; set; }
        public string MtlItemName { get; set; }
        public string UomNo { get; set; }
        public string OspDate { get; set; }


    }



    #endregion

    #region //託外入庫
    public class OSRresult
    {
        public int? OsrId { get; set; }
        public string OsrErpFullNo { get; set; }
        public string TransferStatus { get; set; }
        public string DocDate { get; set; }
        public string ConfirmStatus { get; set; }
        public int Quantity { get; set; }
        public string UomNo { get; set; }
        public string MtlItemNo { get; set; }
        public string MtlItemName { get; set; }
        public string OsrSeq { get; set; }
        public List<ACPTBRresult> ACPTBRresult { get; set; }//應付憑單

    }

    public class ACPTBRresult //應付憑單
    {
        public string TB001 { get; set; }
        public string TB002 { get; set; }
        public string TB003 { get; set; }
        public string TA003 { get; set; }
        public string TA024 { get; set; }
        public List<ACPTDRresult> ACPTDRresult { get; set; }//付款單


    }
    public class ACPTDRresult //付款單
    {
        public string TD001 { get; set; }
        public string TD002 { get; set; }
        public string TD003 { get; set; }
        public string TC003 { get; set; }
        public string TC008 { get; set; }

    }

    public class MOCTLRresult //托外退貨
    {
        public string TL001 { get; set; } //退貨單別
        public string TL002 { get; set; } //退貨單號
        public string TL003 { get; set; } //序號
        public string TL004 { get; set; } //品號
        public string TL005 { get; set; } //品名
        public string TL006 { get; set; } //規格
        public string TL007 { get; set; } //退貨數量
        public string TL008 { get; set; } //單位
        public string TL009 { get; set; } //計價數量
        public string TL010 { get; set; } //計價單位
        public string TL011 { get; set; } //加工單價
        public string TL012 { get; set; } //加工金額
        public string TL013 { get; set; } //庫別
        public string TL014 { get; set; } //批號
        public string TL015 { get; set; } //製令單別
        public string TL016 { get; set; } //製令單號
        public string TK001 { get; set; } //退貨單別
        public string TK002 { get; set; } //退貨單號
        public string TK003 { get; set; } //退貨日期
        public string TK004 { get; set; } //加工廠商
        public string TK006 { get; set; } //幣別
        public string TK007 { get; set; } //匯率
        public string TK008 { get; set; } //件數
        public string TK009 { get; set; } //備註
        public string TK010 { get; set; } //統一編號
        public string TK017 { get; set; } //原幣未稅金額
        public string TK019 { get; set; } //原幣稅額
        public string TK020 { get; set; } //數量合計
        public string TK027 { get; set; } //單據日期
        public string TK028 { get; set; } //確認者

        public string TK021 { get; set; } //

    }

    #endregion

    #region //托外退貨
    //退貨單資料
    public class TKResult
    {
        public string TK001 { get; set; } //退貨單別
        public string TK002 { get; set; } //退貨單號
        public DateTime? TK003 { get; set; } //退貨日期
        public string TK004 { get; set; } //加工廠商
        public string TK006 { get; set; } //幣別
        public string TK007 { get; set; } //匯率
        public string TK008 { get; set; } //件數
        public string TK009 { get; set; } //備註
        public string TK010 { get; set; } //統一編號
        public string TK017 { get; set; } //原幣未稅金額
        public string TK019 { get; set; } //原幣稅額
        public string TK020 { get; set; } //數量合計
        public string TK027 { get; set; } //單據日期
        public string TK028 { get; set; } //確認者
        public List<MOCTL> TLDetail { get; set; }

    }

    #region //單頭 MOCTK
    public class MOCTK
    {
        public string TK001 { get; set; } //退貨單別
        public string TK002 { get; set; } //退貨單號
        public DateTime? TK003 { get; set; } //退貨日期
        public string TK004 { get; set; } //加工廠商
        public string TK006 { get; set; } //幣別
        public string TK007 { get; set; } //匯率
        public string TK008 { get; set; } //件數
        public string TK009 { get; set; } //備註
        public string TK010 { get; set; } //統一編號
        public string TK017 { get; set; } //原幣未稅金額
        public string TK019 { get; set; } //原幣稅額
        public string TK020 { get; set; } //數量合計
        public string TK027 { get; set; } //單據日期
        public string TK028 { get; set; } //確認者

    }
    #endregion

    #region //單身 MOCTL
    public class MOCTL
    {
        public string TL001 { get; set; } //退貨單別
        public string TL002 { get; set; } //退貨單號
        public string TL003 { get; set; } //序號
        public string TL004 { get; set; } //品號
        public string TL005 { get; set; } //品名
        public string TL006 { get; set; } //規格
        public string TL007 { get; set; } //退貨數量
        public string TL008 { get; set; } //單位
        public string TL009 { get; set; } //計價數量
        public string TL010 { get; set; } //計價單位
        public string TL011 { get; set; } //加工單價
        public string TL012 { get; set; } //加工金額
        public string TL013 { get; set; } //庫別
        public string TL014 { get; set; } //批號
        public string TL015 { get; set; } //製令單別
        public string TL016 { get; set; } //製令單號
    }
    #endregion



    #endregion

    #region //MOCTO2 製令變更單單頭
    public class MOCTO2
    {
        public string TO001 { get; set; } //製令單別
        public string TO002 { get; set; } //製令單號
        public string TO003 { get; set; } //變更版次
        public string TO004 { get; set; } //變更日期
        public string TO005 { get; set; } //變更日期
        public string TO006 { get; set; } //新開單日期
        public string TO007 { get; set; } //新BOM日期
        public string TO008 { get; set; } //新BOM版次
        public string TO009 { get; set; } //新產品品號
        public string TO010 { get; set; } //新單位
        public string TO011 { get; set; } //新欄位預留
        public string TO041 { get; set; } //

    }

    #endregion
    #endregion

    #region //托外單明細Jarvix
    public class OutsourcingResult
    {
        public string CompanyName { get; set; }
        public string OspNo { get; set; }
        public string OspDate { get; set; }
        public string OspCreateDate { get; set; }
        public string SupplierNo { get; set; }
        public string SupplierShortName { get; set; }
        public string SupplierName { get; set; }
        public string OspDesc { get; set; }
        public string WoErpFullNo { get; set; }
        public string MtlItemNo { get; set; }
        public string MtlItemName { get; set; }
        public string MtlItemSpec { get; set; }
        public double OspQty { get; set; }
        public double SuppliedQty { get; set; }
        public string ProcessCode { get; set; }
        public string ProcessCodeName { get; set; }
        public string WoCreateDate { get; set; }
        public string WoUserName { get; set; }
        public string OspUserName { get; set; }
        public string ExpectedDate { get; set; }
        public string ReceiptDate { get; set; }
        public string OsrErpFullNo { get; set; }
        public string OspTransferStatus { get; set; }
        public string OspHRemark { get; set; }
        public string OsrSeq { get; set; }
        public double ReceiptQty { get; set; }
        public string OspBRemark { get; set; }
        public string TransferStatus { get; set; }
        public string ProcessAlias { get; set; }


        public double OspTimes { get; set; }//托外計時長
        public string BackStatus { get; set; }//判斷 -未回廠 逾期 準時 提前 
        public string BackStatusRemark { get; set; }//-未回廠 逾期 準時 提前 X日
        public string OsrErpPrefix { get; set; }
        public string OsrErpNo { get; set; }
        public string OpsTransferStatus { get; set; }
        public string Status { get; set; }
        public string OspDetailCreateDate { get; set; }
        public int? TotalCount { get; set; }



    }

    #endregion

    #region //製令加工進度
    public class MoProgressDetailResult
    {
        public int MoId { get; set; }
        public string CompanyName { get; set; }
        public string WoStatus { get; set; }
        public string WoCreater { get; set; }
        public int ERPPlanQty { get; set; }
        public string Mode { get; set; }
        public string WoErpFullNo { get; set; }
        public string MtlItemNo { get; set; }
        public string MtlItemName { get; set; }
        public string MtlItemDesc { get; set; }
        public string ExpectedStart { get; set; }
        public string ExpectedEnd { get; set; }
        public string ActualStart { get; set; }
        public string ActualEnd { get; set; }
        public string MoCreateDay { get; set; }
        public string MoStatus { get; set; }
        public string MoCreater { get; set; }
        public int MoQuantity { get; set; }//領料
        public string RequisitionSetQty { get; set; }
        public string QuantityStatus { get; set; }//領料狀況
        public int MoTotalProcess { get; set; }//總站數
        public int CompleteQty { get; set; }//


        //計算用
        public double StationQty { get; set; }
        public int SortNumber { get; set; }
        public double Quantity { get; set; }
        public double ReceiptQty { get; set; }


        //計算後填入
        public double ProcessProgress { get; set; }//加工進度(%)
        public string ReceiptStatus { get; set; }//入庫狀態
        public double StationQtySum { get; set; }
        public double ReceiptQtySum { get; set; }
        public int? realTotalCount { get; set; }


        public int? TotalCount { get; set; }

        //制令在制查询  Jean 2025-01-15
        public string ProcessAlias { get; set; }//制程
        public double NotStartQty { get; set; }//未开工
        public double StartQty { get; set; }//加工中
        public double WipQty { get; set; }//已完工
        public double TotalPassQty { get; set; }//良品
        public double TotalNgQty { get; set; }//不良
        public double TotalScrapQty { get; set; }//报废
        public string PassRate { get; set; }//良率



    }

    #endregion

    #region //BatchManufactureOrder
    public class BatchManufactureOrder
    {
        public int MtlItemId { get; set; }
        public string MtlItemNo { get; set; }
        public string MtlItemName { get; set; }
        public string MtlItemSpec { get; set; }
        public string ItemAttribute { get; set; }
        public string ItemAttributeName { get; set; }
        public string InventoryNo { get; set; }
        public double? CompositionQuantity { get; set; }
        public double? Base { get; set; }
        public string UomNo { get; set; }
        public int? ParentMtlItemId { get; set; }
        public string ParentMtlItemNo { get; set; }
        public string ParentMtlItemName { get; set; }
        public double? InventoryQty { get; set; }
        public double? PoQty { get; set; }
    }
    #endregion

    #region//IOT MTF參數欄位
    public class MtfmeasureValues
    {
        public string Location_X { get; set; }
        public string Location_Y { get; set; }
        public string Class { get; set; }
        public string CheckFreq { get; set; }
        public string FBL { get; set; }
        public string DOF_Index { get; set; }
        public string DOF_minus { get; set; }
        public string DOF_plus { get; set; }
        public string DOF_T { get; set; }
        public string GDOF_Index { get; set; }
        public string GDOF_minus { get; set; }
        public string GDOF_plus { get; set; }
        public string GDOF_T { get; set; }
        public string EFL { get; set; }
        public string W_1S { get; set; }
        public string W_1T { get; set; }
        public string W_2S { get; set; }
        public string W_2T { get; set; }
        public string W_3S { get; set; }
        public string W_3T { get; set; }
        public string W_4S { get; set; }
        public string W_4T { get; set; }
        public string W_5S { get; set; }
        public string W_5T { get; set; }
        public string W_6S { get; set; }
        public string W_6T { get; set; }
        public string W_7S { get; set; }
        public string W_7T { get; set; }
        public string W_8S { get; set; }
        public string W_8T { get; set; }
        public string W_9S { get; set; }
        public string W_9T { get; set; }
        public string W_10S { get; set; }
        public string W_10T { get; set; }
        public string W_11S { get; set; }
        public string W_11T { get; set; }
        public string W_12S { get; set; }
        public string W_12T { get; set; }
        public string W_13S { get; set; }
        public string W_13T { get; set; }
        public string W_14S { get; set; }
        public string W_14T { get; set; }
        public string W_15S { get; set; }
        public string W_15T { get; set; }
        public string W_16S { get; set; }
        public string W_16T { get; set; }
        public string W_17S { get; set; }
        public string W_17T { get; set; }
        public string W1S_Peak { get; set; }
        public string W1T_Peak { get; set; }
        public string W2S_Peak { get; set; }
        public string W2T_Peak { get; set; }
        public string W3S_Peak { get; set; }
        public string W3T_Peak { get; set; }
        public string W4S_Peak { get; set; }
        public string W4T_Peak { get; set; }
        public string W5S_Peak { get; set; }
        public string W5T_Peak { get; set; }
        public string W6S_Peak { get; set; }
        public string W6T_Peak { get; set; }
        public string W7S_Peak { get; set; }
        public string W7T_Peak { get; set; }
        public string W8S_Peak { get; set; }
        public string W8T_Peak { get; set; }
        public string W9S_Peak { get; set; }
        public string W9T_Peak { get; set; }
        public string W10S_Peak { get; set; }
        public string W10T_Peak { get; set; }
        public string W11S_Peak { get; set; }
        public string W11T_Peak { get; set; }
        public string W12S_Peak { get; set; }
        public string W12T_Peak { get; set; }
        public string W13S_Peak { get; set; }
        public string W13T_Peak { get; set; }
        public string W14S_Peak { get; set; }
        public string W14T_Peak { get; set; }
        public string W15S_Peak { get; set; }
        public string W15T_Peak { get; set; }
        public string W16S_Peak { get; set; }
        public string W16T_Peak { get; set; }
        public string W17S_Peak { get; set; }
        public string W17T_Peak { get; set; }
        public string MTF_DateTime { get; set; }
        public string RecordTime { get; set; }
        public string MeasurementSN { get; set; }

    }
    #endregion

    #region // QMS量測項目
    public class QcItemModel
    {
        public int QcItemId { get; set; }       
        public string QcItemName { get; set; }
    }
    #endregion

    #region // CheckOrderAndProduction
    public class CheckOrderAndProduction
    {
        public string TD001 { get; set; }
        public string TD002 { get; set; }
        public string TD004 { get; set; }
        public string TD005 { get; set; }
        public string MA003 { get; set; }
        public string TD013 { get; set; }
        public decimal TD008 { get; set; }
        public decimal TD009 { get; set; }
        public decimal UndeliveredVolume { get; set; }
        public decimal  TotalInventory { get; set; }
        public decimal TotalInProgress { get; set; }
        public decimal ManufacturingOrder { get; set; }
        public decimal ExcessQuantity { get; set; }
    }
    #endregion

    #region //庫存明細
    public class MWEDetailReport
    {
        public int WoId { get; set; }
        public string WoErpFullNo { get; set; }

        public List<MWEDetailReportDetail> MWEDetailReportDetail { get; set; }//各品庫存明細


    }

    public class MWEDetailReportDetail
    {
        
        public string MtlItemNo { get; set; }
        public string MtlItemName { get; set; }
        public string MtlItemSpec { get; set; }
        public string UmoNo { get; set; }
        public string SmalUmoNo { get; set; }
        public string PackageUmoNo { get; set; }
        public string MWENo { get; set; }
        public string MWEName { get; set; }
        public string TypeOneNo { get; set; }
        public string TypeOneName { get; set; }
        public string TypeTwoNo { get; set; }
        public string TypeTwoName { get; set; }


        public decimal InventoryQuantity { get; set; }//庫存數量
        public decimal InventoryAmount { get; set; }//庫存金額
        public decimal UnitCost { get; set; }//單位成本
        public decimal InventoryPackageQuantity { get; set; }//庫存包裝數量
        public decimal MaterialAmount { get; set; }//材料金額
        public decimal LaborAmount { get; set; }//人工金額
        public decimal ManufacturingCostAmount { get; set; }//製費金額
        public decimal ProcessingAmount { get; set; }//加工金額
    }


    #endregion

    #region //成本趋势分析报表
    public class CostResult
    {

        public string classification { get; set; }
        public decimal Jan { get; set; }
        public decimal Feb { get; set; }
        public decimal Mar { get; set; }
        public decimal Apr { get; set; }
        public decimal May { get; set; }
        public decimal Jun { get; set; }
        public decimal Jul { get; set; }
        public decimal Aug { get; set; }
        public decimal Sept { get; set; }
        public decimal Oct { get; set; }
        public decimal Nov { get; set; }
        public decimal Dec { get; set; }
        public decimal Total { get; set; }
        public decimal proportion { get; set; }
    }


    #endregion

    #region //人工成本分配
    public class LaborCostResult
    {
        public string PurposeType { get; set; }
        public string CategoryName { get; set; }
        public decimal ManualOutput { get; set; }
        public decimal MachineOutput { get; set; }
        public decimal DirectLabor { get; set; }
        public decimal distributedLabour { get; set; }
        public decimal allHandsShared { get; set; }
        public decimal departmentalLabor01 { get; set; }
        public decimal departmentalLabor02 { get; set; }
        public decimal LaborCost { get; set; }

        public decimal EffectiveChargeCost { get; set; }
        public decimal InvalidChargeCost { get; set; }
        public decimal LaborPerUnit { get; set; }
        public decimal EffectiveUnitCost { get; set; }
        public decimal TotalCost { get; set; }


    }



    #endregion

    #region //制费成本分配
    public class ManufacturCostResult
    {
        public string PurposeType { get; set; }
        public string CategoryName { get; set; }
        public decimal EffectiveTime { get; set; }
        public decimal DeadTime { get; set; }
        public decimal MachineCapacity { get; set; }
        public decimal EffectiveCharge { get; set; }
        public decimal InvalidCharge { get; set; }
        public decimal EffectiveChargeDep01 { get; set; }
        public decimal InvalidChargeDep01 { get; set; }
        public decimal EffectiveCoayments { get; set; }
        public decimal InvalidCoayments { get; set; }
        public decimal EffectiveChargeDep02 { get; set; }//
        public decimal InvalidChargeDep02 { get; set; }//
        public decimal EffectivePSE { get; set; }
        public decimal InvalidPSE { get; set; }
        public decimal EffectiveChargeDep03 { get; set; }//
        public decimal InvalidChargeDep03 { get; set; }//
        public decimal TotalEffectiveManufee { get; set; }//
        public decimal TotalInvalidManufee { get; set; }//
        public decimal TotalManufee { get; set; }//
        public decimal InvalidUnitfee { get; set; }//


    }

    #endregion

    #region //解析QcItemJson
    public class QcItemJson
    {
        public int QcItemId { get; set; }
        public string QcItemNo { get; set; }
        public string QcItemName { get; set; }
        public string QcItemDesc { get; set; }
        public string QcGroupName { get; set; }
        public string QcClassName { get; set; }
        public string QcMachineName { get; set; }
        public string QcMachineNumber { get; set; }
        public string DesignValue { get; set; }
        public string UpperTolerance { get; set; }
        public string LowerTolerance { get; set; }
    }
    #endregion
    #region
    public class QcMeasureDataModel
    {
        public int QcRecordId { get; set; }
        public int QmmDetailId { get; set; }        
        public int QcItemId { get; set; }
        public string QcItemDesc { get; set; }
        public string MeasureValue { get; set; }
        public string QcResult { get; set; }
        public string TrayLocation { get; set; }
        public int? CauseId { get; set; }
        public int CreateBy { get; set; }
        public int LastModifiedBy { get; set; }
        // 其他欄位
    }
    #endregion

    #region //在製報表相關
    public class BarcodeWipCount
    {
        public int MoProcessId { get; set; }
        public string ProcessAlias { get; set; }
        public string WoSeq { get; set; }
        public string WoErpFullNo { get; set; }
        public int NotStartQty { get; set; }
        public int StartQty { get; set; }
        public int WipQty { get; set; }
        public string PreNecessityStatus { get; set; }
        public int SortNumber { get; set; }
        public int MoId { get; set; }
    }
    #endregion

    #region //委外預排資料相關
    public class OspListData
    {
        public int OspListId { get; set; }
        public int GroupId { get; set; }
        public int MoProcessId { get; set; }
        public string WoErpFullNo { get; set; }
        public string ProcessNo { get; set; }
        public string ProcessCodeName { get; set; }
        public string ProcessAlias { get; set; }
        public int DepartmentId { get; set; }
        public string DepartmentNo { get; set; }
        public string DepartmentName { get; set; }
        public int? SupplierId { get; set; }
        public string SupplierNo { get; set; } = "";
        public string SupplierName { get; set; } = "";
        public string ProcessCode { get; set; } = "";
        public string OspDate { get; set; }
        public int Quantity { get; set; }
        public string ExpectedDate { get; set; }
        public string Remark { get; set; }
        public string ProcessCheckStatus { get; set; }
        public string ProcessCheckType { get; set; }
        public string Status { get; set; }
        public string MtlItemNo { get; set; }
        public string MtlItemName { get; set; }
        public string MtlItemSpec { get; set; }
        public string WoErpPrefix { get; set; }
        public string WoErpNo { get; set; }
        public string WoSeq { get; set; }
    }
    #endregion

    #region //修改托外生產單詳細資料
    public class OspDetailData
    {
        public int OspDetailId { get; set; }
        public int OspId { get; set; }
        public int SupplierId { get; set; }
        public string ProcessCode { get; set; }
        public DateTime ExpectedDate { get; set; }
        public string UserNo { get; set; }
    }
    #endregion

    #region //批量新增套环资料
    public class LCRFileInfo
    {
        public int? FileId { get; set; }
        public string FileName { get; set; }
        public byte[] FileContent { get; set; }
        public string FileExtension { get; set; }
        public int FileSize { get; set; }
        public string ClientIP { get; set; }
        public string Source { get; set; }
        public string DeleteStatus { get; set; }
    }

    public class LCRExcelFormat
    {
        public string ModelName { get; set; }
        public string RingName { get; set; }
        public string Remarks { get; set; }
        public int HoleCount { get; set; }
        public string RingSpec { get; set; }
        public string RingCode { get; set; }
        public string RingShape { get; set; }
        public string Customer { get; set; }
        public decimal DailyDemand { get; set; }
        public decimal SafetyStock { get; set; }
        public int RowNumber { get; set; }
    }
    //批量新增库存异动资料
    public class LCRTransExcelFormat
    {
        public string ModelName { get; set; }
        public string TransType { get; set; }
        public string TransDate { get; set; }
        public int Quantity { get; set; }
        public int RowNumber { get; set; }
    }

    #endregion

    #region //ERP制令投料统计

    public class ERPWipOrders
    {
        public string TA001 { get; set; }
        public string TA002 { get; set; }
        public string TA006 { get; set; }
        public string TB003 { get; set; }
        public double TB004 { get; set; }
        public double TB005 { get; set; }
    }
    
    public class MatFeedRegs
    {
        public int MoId { get; set; }
        public string WoErpPrefix { get; set; }
        public string WoErpPre { get; set; }
        public string WoErpNo { get; set; }
        public string WoErpMtlItemNo { get; set; }
        public string WoErpMtlItemName { get; set; }
        public string WoErpMtlItemSpec { get; set; }
        public string MaterialId { get; set; }
        public string MtlItemNo { get; set; }
        public string MtlItemName { get; set; }
        public double SUMMatInRegQty { get; set; }
        public double ERPQty01 { get; set; }
        public double ERPQty02 { get; set; }
        public double ERPQty03 { get; set; }
        public double ERPQty04 { get; set; }
        public int? TotalCount { get; set; } 
    }


    #endregion


}
 