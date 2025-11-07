using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDMDA
{
    #region //品號欄位
    #region //BM 品號欄位
    public class MtlItem
    {
        public int? MtlItemId { get; set; }
        public string MtlItemNo { get; set; }
        public int? CompanyId { get; set; }
        public int? ItemTypeId { get; set; }
        public string ItemTypeNo { get; set; }
        public string ItemTypeName { get; set; }
        public string MtlItemName { get; set; }
        public string MtlItemSpec { get; set; }
        public int? InventoryUomId { get; set; }
        public string InventoryUomNo { get; set; }
        public int? WeightUomId { get; set; }
        public string WeightUomNo { get; set; }
        public int? PurchaseUomId { get; set; }
        public string PurchaseUomNo { get; set; }
        public int? SaleUomId { get; set; }
        public string SaleUomNo { get; set; }
        public string TypeOne { get; set; }
        public string TypeOneName { get; set; }
        public string TypeTwo { get; set; }
        public string TypeTwoName { get; set; }
        public string TypeThree { get; set; }
        public string TypeThreeName { get; set; }
        public string TypeFour { get; set; }
        public string TypeFourName { get; set; }
        public int? InventoryId { get; set; }
        public string InventoryName { get; set; }
        public string InventoryNo { get; set; }
        public int? RequisitionInventoryId { get; set; }
        public string RequisitionInventoryNo { get; set; }
        public string InventoryManagement { get; set; }
        public string MtlModify { get; set; }
        public string BondedStore { get; set; }
        public string ItemAttribute { get; set; }
        public string LotManagement { get; set; }
        public string ProductionLine { get; set; }
        public string MeasureType { get; set; }
        public DateTime? EffectiveDate { get; set; }
        public DateTime? ExpirationDate { get; set; }
        public string OverReceiptManagement { get; set; }
        public string OverDeliveryManagement { get; set; }
        public string Version { get; set; }
        public string MtlItemDesc { get; set; }
        public string MtlItemRemark { get; set; }
        public string CustomerMtlItem { get; set; }
        public string TransferStatus { get; set; }
        public DateTime? TransferDate { get; set; }
        public string Status { get; set; }
        public string StatusName { get; set; }
        public string ErpStatus { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public string UserNo { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
        public int? TotalCount { get; set; }
        public string UserName { get; set; }
        public int? MtlItemUomCalculateQty { get; set; }
        public string ReplenishmentPolicy { get; set; }

        public int? BomTransferQty { get; set; }
        public int? CusTransferQty { get; set; }
        public int? SubTransferQty { get; set; }
        public int? UomTransferQty { get; set; }
        public int? EfficientDays { get; set; }
        public int? RetestDays { get; set; }

        public int Effective { get; set; }
        public double TotalInventoryQty { get; set; }
    }
    #endregion

    #region //ERP 品號欄位
    public class INVMB
    {
        public int MtlItemId { get; set; }
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
        public string MB001 { get; set; } //品號
        public string MB002 { get; set; } //品名
        public string MB003 { get; set; } //規格
        public string MB004 { get; set; } //庫存單位
        public string MB005 { get; set; } //[品號分類一]
        public string MB006 { get; set; } //[品號分類二]
        public string MB007 { get; set; } //[品號分類三]
        public string MB008 { get; set; } //[品號分類四]
        public string MB009 { get; set; } //商品描述
        public string MB010 { get; set; } //標準途程品號
        public string MB011 { get; set; } //標準途程代號
        public string MB012 { get; set; } //文管代號
        public string MB013 { get; set; } //條碼編號
        public int? MB014 { get; set; } //單位淨重
        public string MB015 { get; set; } //重量單位
        public string MB016 { get; set; } //外包裝單位
        public string MB017 { get; set; } //主要庫別
        public string MB018 { get; set; } //計劃人員
        public string MB019 { get; set; } //庫存管理
        public string MB020 { get; set; } //保稅品
        public string MB021 { get; set; } //循環盤點碼
        public string MB022 { get; set; } //批號管理
        public int? MB023 { get; set; } //有效天數
        public int? MB024 { get; set; } //複檢天數
        public string MB025 { get; set; } //品號屬性
        public string MB026 { get; set; } //低階碼
        public string MB027 { get; set; } //ABC等級
        public string MB028 { get; set; } //備註
        public string MB029 { get; set; } //產品圖號
        public string MB030 { get; set; } //生效日期
        public string MB031 { get; set; } //失效日期
        public string MB032 { get; set; } //主供應商
        public string MB033 { get; set; } //MPS件
        public string MB034 { get; set; } //補貨政策
        public string MB035 { get; set; } //補貨週期
        public int? MB036 { get; set; } //固定前置天數
        public int? MB037 { get; set; } //變動前置天數
        public int? MB038 { get; set; } //批量
        public int? MB039 { get; set; } //最低補量
        public int? MB040 { get; set; } //補貨倍量
        public int? MB041 { get; set; } //領用倍量
        public string MB042 { get; set; } //領料碼
        public string MB043 { get; set; } //檢驗方式
        public string MB044 { get; set; } //超收管理
        public int? MB045 { get; set; } //超收率%
        public int? MB046 { get; set; } //標準進價
        public int? MB047 { get; set; } //標準售價
        public string MB048 { get; set; } //最近進價幣別-原幣別
        public int? MB049 { get; set; } //最近進價-原幣單價
        public int? MB050 { get; set; } //最近進價-本幣單價
        public int? MB051 { get; set; } //零售價
        public string MB052 { get; set; } //零售價含稅
        public int? MB053 { get; set; } //售價定價一
        public int? MB054 { get; set; } //售價定價二
        public int? MB055 { get; set; } //售價定價三
        public int? MB056 { get; set; } //售價定價四
        public int? MB057 { get; set; } //單位標準材料成本
        public int? MB058 { get; set; } //單位標準人工成本
        public int? MB059 { get; set; } //單位標準製造費用
        public int? MB060 { get; set; } //單位標準加工費用
        public int? MB061 { get; set; } //本階人工
        public int? MB062 { get; set; } //本階製費
        public int? MB063 { get; set; } //本階加工
        public int? MB064 { get; set; } //庫存數量
        public int? MB065 { get; set; } //庫存金額
        public string MB066 { get; set; } //修改品名/規格
        public string MB067 { get; set; } //採購人員
        public string MB068 { get; set; } //生產線別
        public int? MB069 { get; set; } //售價定價五
        public int? MB070 { get; set; } //售價定價六
        public int? MB071 { get; set; } //外包裝材積
        public string MB072 { get; set; } //小單位
        public int? MB073 { get; set; } //外包裝含商品數
        public int? MB074 { get; set; } //外包裝淨重
        public int? MB075 { get; set; } //外包裝毛重
        public int? MB076 { get; set; } //檢驗天數
        public string MB077 { get; set; } //品管類別
        public int? MB078 { get; set; } //MRP生產允許交期提前天數
        public int? MB079 { get; set; } //MRP採購允許交期提前天數
        public string MB080 { get; set; } //貨號
        public string MB081 { get; set; } //SIZE
        public int? MB082 { get; set; } //關稅率
        public string MB083 { get; set; } //進價管制
        public int? MB084 { get; set; } //單價上限率
        public string MB085 { get; set; } //售價管制
        public int? MB086 { get; set; } //單價下限率
        public string MB087 { get; set; } //超交管理
        public int? MB088 { get; set; } //超交率
        public int? MB089 { get; set; } //庫存包裝數量
        public string MB090 { get; set; } //包裝單位
        public string MB091 { get; set; } //定重
        public string MB092 { get; set; } //產品序號管理
        public int? MB093 { get; set; } //長(CM)
        public int? MB094 { get; set; } //寬(CM)
        public int? MB095 { get; set; } //高(CM)
        public int? MB096 { get; set; } //工時底數
        public int? MB097 { get; set; } //業務底價
        public string MB098 { get; set; } //業務底價含稅
        public int? MB099 { get; set; } //貨物稅率
        public string MB100 { get; set; } //標準進價含稅
        public string MB101 { get; set; } //標準售價含稅
        public string MB102 { get; set; } //最近進價-單價含稅(原/本幣)
        public string MB103 { get; set; } //售價定價一含稅
        public string MB104 { get; set; } //售價定價二含稅
        public string MB105 { get; set; } //售價定價三含稅
        public string MB106 { get; set; } //售價定價四含稅
        public string MB107 { get; set; } //售價定價五含稅
        public string MB108 { get; set; } //售價定價六含稅
        public string MB109 { get; set; } //新品號
        public string MB110 { get; set; } //料件承認碼
        public int? MB111 { get; set; } //轉撥倍量
        public string MB112 { get; set; } //MRP保留欄位
        public string MB113 { get; set; } //APS預留欄位
        public string MB114 { get; set; } //APS預留欄位
        public string MB115 { get; set; } //APS預留欄位
        public string MB116 { get; set; } //APS預留欄位
        public string MB117 { get; set; } //APS預留欄位
        public string MB118 { get; set; } //APS預留欄位
        public int? MB119 { get; set; } //APS預留欄位
        public int? MB120 { get; set; } //APS預留欄位
        public string MB121 { get; set; } //控制編碼原則
        public string MB122 { get; set; } //序號前置碼
        public string MB123 { get; set; } //序號流水號碼數
        public string MB124 { get; set; } //序號編碼原則
        public string MB125 { get; set; } //已用生管序號
        public string MB126 { get; set; } //已用商品序號
        public string MB127 { get; set; } //來源
        public string MB128 { get; set; } //預留欄位
        public string MB129 { get; set; } //預留欄位
        public string MB130 { get; set; } //預留欄位
        public string MB131 { get; set; } //電子發票須上傳產品追溯串接碼
        public string MB132 { get; set; } //預留欄位
        public string MB133 { get; set; } //屬性代碼一
        public string MB134 { get; set; } //屬性代碼二
        public string MB135 { get; set; } //屬性代碼三
        public string MB136 { get; set; } //屬性代碼四
        public string MB137 { get; set; } //屬性代碼五
        public string MB138 { get; set; } //屬性代碼六
        public string MB139 { get; set; } //屬性代碼七
        public string MB140 { get; set; } //屬性代碼八
        public string MB141 { get; set; } //屬性代碼九
        public string MB142 { get; set; } //屬性組代碼
        public string MB143 { get; set; } //屬性內容一
        public string MB144 { get; set; } //屬性內容二
        public string MB145 { get; set; } //屬性內容三
        public string MB146 { get; set; } //屬性內容四
        public string MB147 { get; set; } //屬性內容五
        public string MB148 { get; set; } //屬性內容六
        public string MB149 { get; set; } //屬性內容七
        public string MB150 { get; set; } //屬性內容八
        public string MB151 { get; set; } //屬性內容九
        public string MB152 { get; set; } //屬性代碼十
        public string MB153 { get; set; } //屬性內容十
        public string MB154 { get; set; } //圖號版次
        public string MB155 { get; set; } //採購單位
        public string MB156 { get; set; } //銷售單位
        public string MB157 { get; set; } //主要領料庫別
        public string MB158 { get; set; } //新品號核准日期
        public int? MB159 { get; set; } //預留欄位
        public int? MB160 { get; set; } //預留欄位
        public int? MB161 { get; set; } //預留欄位
        public string MB162 { get; set; } //預留欄位
        public string MB163 { get; set; } //稅則
        public string MB164 { get; set; } //產品追溯系統串接碼
        public string MB165 { get; set; } //版次
        public int? MB166 { get; set; } //贈品率
        public int? MB167 { get; set; } //預留欄位
        public string MB168 { get; set; } //預留欄位
        public string MB169 { get; set; } //DATECODE管理
        public string MB170 { get; set; } //Bin管理
        public int? MB171 { get; set; } //預留欄位
        public int? MB172 { get; set; } //預留欄位
        public int? MB173 { get; set; } //預留欄位
        public string MB174 { get; set; } //預留欄位
        public string MB175 { get; set; } //預留欄位
        public string MB176 { get; set; } //預留欄位
        public string MB177 { get; set; } //預留欄位
        public string MB178 { get; set; } //預留欄位
        public string MB179 { get; set; } //預留欄位
        public int? MB180 { get; set; } //APS固定工時
        public int? MB181 { get; set; } //APS變動工時
        public int? MB182 { get; set; } //批次加工量
        public string MB183 { get; set; } //預留欄位
        public int? MB184 { get; set; } //基準數量
        public string MB185 { get; set; } //資源群組
        public string MB186 { get; set; } //資源群組名稱
        public string MB187 { get; set; } //機台代號
        public string MB188 { get; set; } //機台名稱
        public string MB189 { get; set; } //指定資源
        public string MB190 { get; set; } //關鍵料號
        public int? MB191 { get; set; } //營業稅率
        public int? MB192 { get; set; } //保固佔售價比率
        public int? MB193 { get; set; } //保固期數(月數)
        public string MB194 { get; set; } //交易設限碼
        public string MB195 { get; set; } //建立作業修改日期
        public string MB196 { get; set; } //建立作業修改者
        public string MB197 { get; set; } //MSIC_Code
        public string MB198 { get; set; } //預留欄位
        public string MB199 { get; set; } //預留欄位
        public string MB500 { get; set; } //產品系列ID
        public int? MB501 { get; set; } //GROSS_DIE
        public string MB502 { get; set; } //說明1
        public string MB503 { get; set; } //說明2
        public string MB504 { get; set; } //光罩層次
        public string MB505 { get; set; } //進項稅別碼
        public string MB506 { get; set; } //銷項稅別碼
        public int? MB507 { get; set; } //預留欄位
        public int? MB508 { get; set; } //存貨週轉目標次數
        public int? MB509 { get; set; } //預留欄位
        public string MB510 { get; set; } //預留欄位
        public string MB511 { get; set; } //預留欄位
        public string MB512 { get; set; } //預留欄位
        public string MB513 { get; set; } //預留欄位
        public string MB514 { get; set; } //預留欄位
        public string MB515 { get; set; } //開帳批號
        public string MB516 { get; set; } //批號開帳日期
        public string MB517 { get; set; } //內箱標籤格式
        public string MB518 { get; set; } //中箱標籤格式
        public string MB519 { get; set; } //外箱標籤格式
        public string MB520 { get; set; } //內箱選擇列印
        public string MB521 { get; set; } //中箱選擇列印
        public string MB522 { get; set; } //外箱選擇列印
        public int? MB545 { get; set; } //標準產出良率
        public string MB550 { get; set; } //產品屬性
        public string MB551 { get; set; } //Wafer型號
        public string MB552 { get; set; } //性質
        public string MB553 { get; set; } //版本
        public string MB554 { get; set; } //作業代號
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
    }
    #endregion

    #region //ERP 新品號欄位
    public class INVMO
    {
        public int MtlItemId { get; set; }
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
        public string MO001 { get; set; } //新品號
        public string MO002 { get; set; } //品名
        public string MO003 { get; set; } //規格
        public string MO004 { get; set; } //庫存單位
        public string MO005 { get; set; } //[品號分類一]
        public string MO006 { get; set; } //[品號分類二]
        public string MO007 { get; set; } //[品號分類三]
        public string MO008 { get; set; } //[品號分類四]
        public string MO009 { get; set; } //商品描述
        public string MO010 { get; set; } //標準途程品號
        public string MO011 { get; set; } //標準途程代號
        public string MO012 { get; set; } //文管代號
        public string MO013 { get; set; } //條碼編號
        public int? MO014 { get; set; } //單位淨重
        public string MO015 { get; set; } //重量單位
        public string MO016 { get; set; } //外包裝單位
        public string MO017 { get; set; } //主要庫別
        public string MO018 { get; set; } //計劃人員
        public string MO019 { get; set; } //庫存管理
        public string MO020 { get; set; } //保稅品
        public string MO021 { get; set; } //循環盤點碼
        public string MO022 { get; set; } //批號管理
        public int? MO023 { get; set; } //有效天數
        public int? MO024 { get; set; } //複檢天數
        public string MO025 { get; set; } //品號屬性
        public string MO026 { get; set; } //低階碼
        public string MO027 { get; set; } //ABC等級
        public string MO028 { get; set; } //備註
        public string MO029 { get; set; } //產品圖號
        public string MO030 { get; set; } //生效日期
        public string MO031 { get; set; } //失效日期
        public string MO032 { get; set; } //主供應商
        public string MO033 { get; set; } //MPS件
        public string MO034 { get; set; } //補貨政策
        public string MO035 { get; set; } //補貨週期
        public int? MO036 { get; set; } //固定前置天數
        public int? MO037 { get; set; } //變動前置天數
        public int? MO038 { get; set; } //批量
        public int? MO039 { get; set; } //最低補量
        public int? MO040 { get; set; } //補貨倍量
        public int? MO041 { get; set; } //領用倍量
        public string MO042 { get; set; } //領料碼
        public string MO043 { get; set; } //檢驗方式
        public string MO044 { get; set; } //超收管理
        public int? MO045 { get; set; } //超收率%
        public int? MO046 { get; set; } //標準進價
        public int? MO047 { get; set; } //標準售價
        public string MO048 { get; set; } //最近進價幣別-原幣別
        public int? MO049 { get; set; } //最近進價-原幣單價
        public int? MO050 { get; set; } //最近進價-本幣單價
        public int? MO051 { get; set; } //零售價
        public string MO052 { get; set; } //零售價含稅
        public int? MO053 { get; set; } //售價定價一
        public int? MO054 { get; set; } //售價定價二
        public int? MO055 { get; set; } //售價定價三
        public int? MO056 { get; set; } //售價定價四
        public int? MO057 { get; set; } //單位標準材料成本
        public int? MO058 { get; set; } //單位標準人工成本
        public int? MO059 { get; set; } //單位標準製造費用
        public int? MO060 { get; set; } //單位標準加工費用
        public int? MO061 { get; set; } //本階人工
        public int? MO062 { get; set; } //本階製費
        public int? MO063 { get; set; } //本階加工
        public int? MO064 { get; set; } //庫存數量
        public int? MO065 { get; set; } //庫存金額
        public string MO066 { get; set; } //修改品名/規格
        public string MO067 { get; set; } //採購人員
        public string MO068 { get; set; } //生產線別
        public int? MO069 { get; set; } //售價定價五
        public int? MO070 { get; set; } //售價定價六
        public int? MO071 { get; set; } //外包裝材積
        public string MO072 { get; set; } //小單位
        public int? MO073 { get; set; } //外包裝含商品數
        public int? MO074 { get; set; } //外包裝淨重
        public int? MO075 { get; set; } //外包裝毛重
        public int? MO076 { get; set; } //檢驗天數
        public string MO077 { get; set; } //品管類別
        public int? MO078 { get; set; } //MRP生產允許交期提前天數
        public int? MO079 { get; set; } //MRP採購允許交期提前天數
        public string MO080 { get; set; } //貨號
        public string MO081 { get; set; } //SIZE
        public int? MO082 { get; set; } //關稅率
        public string MO083 { get; set; } //進價管制
        public int? MO084 { get; set; } //單價上限率
        public string MO085 { get; set; } //售價管制
        public int? MO086 { get; set; } //單價下限率
        public string MO087 { get; set; } //超交管理
        public int? MO088 { get; set; } //超交率
        public int? MO089 { get; set; } //庫存包裝數量
        public string MO090 { get; set; } //包裝單位
        public string MO091 { get; set; } //定重
        public string MO092 { get; set; } //產品序號管理
        public int? MO093 { get; set; } //長(CM)
        public int? MO094 { get; set; } //寬(CM)
        public int? MO095 { get; set; } //高(CM)
        public int? MO096 { get; set; } //工時底數
        public int? MO097 { get; set; } //業務底價
        public string MO098 { get; set; } //業務底價含稅
        public int? MO099 { get; set; } //貨物稅率
        public string MO100 { get; set; } //標準進價含稅
        public string MO101 { get; set; } //標準售價含稅
        public string MO102 { get; set; } //最近進價-單價含稅(原/本幣)
        public string MO103 { get; set; } //售價定價一含稅
        public string MO104 { get; set; } //售價定價二含稅
        public string MO105 { get; set; } //售價定價三含稅
        public string MO106 { get; set; } //售價定價四含稅
        public string MO107 { get; set; } //售價定價五含稅
        public string MO108 { get; set; } //售價定價六含稅
        public string MO109 { get; set; } //新品號
        public string MO110 { get; set; } //料件承認碼
        public int? MO111 { get; set; } //轉撥倍量
        public string MO112 { get; set; } //MRP保留欄位
        public string MO113 { get; set; } //APS預留欄位
        public string MO114 { get; set; } //APS預留欄位
        public string MO115 { get; set; } //APS預留欄位
        public string MO116 { get; set; } //APS預留欄位
        public string MO117 { get; set; } //APS預留欄位
        public string MO118 { get; set; } //APS預留欄位
        public int? MO119 { get; set; } //APS預留欄位
        public int? MO120 { get; set; } //APS預留欄位
        public string MO121 { get; set; } //控制編碼原則
        public string MO122 { get; set; } //序號前置碼
        public string MO123 { get; set; } //序號流水號碼數
        public string MO124 { get; set; } //序號編碼原則
        public string MO125 { get; set; } //已用生管序號
        public string MO126 { get; set; } //已用商品序號
        public string MO127 { get; set; } //預留欄位
        public string MO128 { get; set; } //預留欄位
        public string MO129 { get; set; } //預留欄位
        public string MO130 { get; set; } //預留欄位
        public string MO131 { get; set; } //預留欄位
        public string MO132 { get; set; } //預留欄位
        public string MO133 { get; set; } //屬性代碼一
        public string MO134 { get; set; } //屬性代碼二
        public string MO135 { get; set; } //屬性代碼三
        public string MO136 { get; set; } //屬性代碼四
        public string MO137 { get; set; } //屬性代碼五
        public string MO138 { get; set; } //屬性代碼六
        public string MO139 { get; set; } //屬性代碼七
        public string MO140 { get; set; } //屬性代碼八
        public string MO141 { get; set; } //屬性代碼九
        public string MO142 { get; set; } //屬性組代碼
        public string MO143 { get; set; } //屬性內容一
        public string MO144 { get; set; } //屬性內容二
        public string MO145 { get; set; } //屬性內容三
        public string MO146 { get; set; } //屬性內容四
        public string MO147 { get; set; } //屬性內容五
        public string MO148 { get; set; } //屬性內容六
        public string MO149 { get; set; } //屬性內容七
        public string MO150 { get; set; } //屬性內容八
        public string MO151 { get; set; } //屬性內容九
        public string MO152 { get; set; } //屬性代碼十
        public string MO153 { get; set; } //屬性內容十
        public string MO154 { get; set; } //圖號版次
        public string MO155 { get; set; } //採購單位
        public string MO156 { get; set; } //銷售單位
        public string MO157 { get; set; } //確認碼
        public string MO158 { get; set; } //確認日期
        public string MO159 { get; set; } //確認者
        public string MO160 { get; set; } //簽核狀態碼
        public string MO161 { get; set; } //品號
        public string MO162 { get; set; } //申請日期
        public string MO163 { get; set; } //主要領料庫別
        public int? MO164 { get; set; } //預留欄位
        public int? MO165 { get; set; } //預留欄位
        public string MO166 { get; set; } //預留欄位
        public string MO167 { get; set; } //稅則
        public string MO168 { get; set; } //預留欄位
        public int? MO169 { get; set; } //贈品率
        public int? MO170 { get; set; } //預留欄位
        public string MO171 { get; set; } //預留欄位
        public int? MO172 { get; set; } //預留欄位
        public int? MO173 { get; set; } //預留欄位
        public int? MO174 { get; set; } //預留欄位
        public string MO175 { get; set; } //預留欄位
        public string MO176 { get; set; } //預留欄位
        public string MO177 { get; set; } //預留欄位
        public string MO178 { get; set; } //預留欄位
        public string MO179 { get; set; } //預留欄位
        public string MO180 { get; set; } //預留欄位
        public int? MO181 { get; set; } //APS固定工時
        public int? MO182 { get; set; } //APS變動工時
        public int? MO183 { get; set; } //批次加工量
        public string MO184 { get; set; } //預留欄位
        public int? MO185 { get; set; } //基準數量
        public string MO186 { get; set; } //資源群組
        public string MO187 { get; set; } //資源群組名稱
        public string MO188 { get; set; } //機台代號
        public string MO189 { get; set; } //機台名稱
        public string MO190 { get; set; } //指定資源
        public string MO191 { get; set; } //關鍵料號
        public int? MO192 { get; set; } //營業稅率
        public string MO193 { get; set; } //Bin管理
        public string MO194 { get; set; } //交易設限碼
        public int? MO195 { get; set; } //保固佔售價比率
        public int? MO196 { get; set; } //保固期數(月數)
        public string MO197 { get; set; } //
        public string MO198 { get; set; } //預留欄位
        public string MO199 { get; set; } //預留欄位
        public string MO500 { get; set; } //產品系列ID
        public int? MO501 { get; set; } //GROSS_DIE
        public string MO502 { get; set; } //說明1
        public string MO503 { get; set; } //說明2
        public string MO504 { get; set; } //光罩層次
        public string MO505 { get; set; } //進項稅別碼
        public string MO506 { get; set; } //銷項稅別碼
        public int? MO507 { get; set; } //
        public int? MO508 { get; set; } //
        public int? MO545 { get; set; } //標準產出良率
        public string MO550 { get; set; } //產品屬性
        public string MO551 { get; set; } //Wafer型號
        public string MO552 { get; set; } //性質
        public string MO553 { get; set; } //版本
        public string MO554 { get; set; } //作業代號
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
    }
    #endregion

    #region //ERP 客戶品號欄位
    public class COPMG
    {
        public int CustomerMtlItemId { get; set; }
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
        public string MG001 { get; set; } //客戶代號
        public string MG002 { get; set; } //品號
        public string MG003 { get; set; } //客戶品號
        public string MG004 { get; set; } //客戶商品描述
        public string MG005 { get; set; } //客戶品名
        public string MG006 { get; set; } //客戶規格
        public int? MG007 { get; set; } //預留欄位
        public int? MG008 { get; set; } //預留欄位
        public string MG009 { get; set; } //預留欄位
        public string MG010 { get; set; } //預留欄位
        public string MG011 { get; set; } //預留欄位
        public int? MG012 { get; set; } //保固佔售價比率
        public int? MG013 { get; set; } //保固期數(月數)
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
    }
    #endregion

    #region //品號批量欄位
    public class MtlItemBatch: MtlItem
    {
        public int row { get; set; }

        public string ItemNo { get; set; }

        /// <summary>
        /// 批次作業(前端)使用 請勿直接修改
        /// </summary>
        public string Msg { get; set; }

        public string BillOfMaterials { get; set; }

        public List<MtlItemSegment> MtlItemSegments { get; set; }

        public MtlItemBatch()
        {
            MtlItemSegments = new List<MtlItemSegment>();
        }
    }
    #endregion

    #region //品號節段資料
    public class MtlItemSegment
    {
        public int MtlItemId { get; set; }
        public int ItemTypeId { get; set; }
        public string SegmentType { get; set; }
        public int SegmentSort { get; set; }
        public string SegmentValue { get; set; }
        public string SaveValue { get; set; }
        public string SuffixCode { get; set; }
    }
    #endregion
    #endregion

    #region //客戶品號欄位
    public class CustomerMtlItem
    {
        public int? CustomerMtlItemId { get; set; }
        public int? MtlItemId { get; set; }
        public string MtlItemNo { get; set; }
        public int? CustomerId { get; set; }
        public string CustomerNo { get; set; }
        public string CustomerMtlItemNo { get; set; }
        public string CustomerMtlItemName { get; set; }
        public string CustomerMtlItemSpec { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
        public bool Exist { get; set; } = false;
    }
    #endregion

    #region //BOM欄位
    #region //BOM樹狀結構

    public class BillOfMaterialsTree : MtlItem
    {
        public int ParentBomId { get; set; }

        public int ParentMtlItemId { get; set; }

        public int BomId { get; set; }

        public string BomSequence { get; set; }

        public string MtlItemRoute { get; set; }

        public string ParentMtlItemNo { get; set; }

        public int lvl { get; set; }

        public int SortNumber { get; set; }

        public bool IsNewItem { get; set; } = false;

        public bool IsClone { get; set; } = false;

        public string[] EditColumns { get; set; } = new string[0];

        public int NewMtlItemId { get; set; } = -1; //寫入時產生，給關聯資料使用。
        public int? NewBomId { get; set; } = null; //寫入時產生，給關聯資料使用。
        //
        public int? NewItemTypeId { get; set; } = null;
        public string NewMtlItemNo { get; set; } = "";

        public string NewMtlItemName { get; set; } = "";

        public string NewMtlItemSpec { get; set; } = "";

        public string NewTypeOne { get; set; } = "";

        public string NewTypeTwo { get; set; } = "";

        public string NewTypeThree { get; set; } = "";

        public string NewTypeFour { get; set; } = "";

        public int? NewInventoryId { get; set; } = null;

        public int ErrorFlag { get; set; } = 0;

        /// <summary>
        /// 品號編碼節段資料
        /// </summary>
        public List<ItemTypeSegment> ItemTypeSegments { get; set; } = new List<ItemTypeSegment>();
        public List<ItemTypeSegment> NewItemTypeSegments { get; set; } = new List<ItemTypeSegment>();

        /// <summary>
        /// BOM批次作業(前端)使用 請勿直接修改
        /// </summary>
        public string Msg { get; set; } = string.Empty;
    }
    #endregion

    #region //BM 品號編碼節段資料
    public class ItemTypeSegment
    {
        //public int ItemTypeId { get; set; }
        public int ItemTypeSegmentId { get; set; }
        public int SortNumber { get; set; }
        public string SegmentType { get; set; }
        public string SegmentValue { get; set; }
        public string SuffixCode { get; set; }
        public string EffectiveDate { get; set; }
        public string InactiveDate { get; set; }
    }
    #endregion

    #region //BM BOM主件欄位
    public class BillOfMaterials
    {
        public int? BomId { get; set; }
        public int? MtlItemId { get; set; }
        public string MtlItemNo { get; set; }
        public string MtlItemName { get; set; }
        public string MtlItemSpec { get; set; }
        public int? UomId { get; set; }
        public string UomNo { get; set; }
        public double? StandardLot { get; set; }
        public string WipPrefix { get; set; }
        public string WipPrefixName { get; set; }
        public string ModiPrefix { get; set; }
        public string ModiNo { get; set; }
        public string ModiSequence { get; set; }
        public string Version { get; set; }
        public string Remark { get; set; }
        public string ConfirmStatus { get; set; }
        public int? ConfirmUserId { get; set; }
        public string ConfirmUserNo { get; set; }
        public DateTime? ConfirmDate { get; set; }
        public string TransferStatus { get; set; }
        public DateTime? TransferDate { get; set; }
        public string Status { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
    }
    #endregion

    #region //BM BOM元件欄位
    public class BomDetail
    {
        public int? BomDetailId { get; set; }
        public int? BomId { get; set; }
        public int? BomMtlItemId { get; set; }
        public string BomMtlItemNo { get; set; }
        public string BomSequence { get; set; }
        public int? MtlItemId { get; set; }
        public string MtlItemNo { get; set; }
        public int? UomId { get; set; }
        public string UomNo { get; set; }
        public double? CompositionQuantity { get; set; }
        public double? Base { get; set; }
        public double? LossRate { get; set; }
        public string StandardCostingType { get; set; }
        public string MaterialProperties { get; set; }
        public string ReleaseItem { get; set; }
        public DateTime? EffectiveDate { get; set; }
        public DateTime? ExpirationDate { get; set; }
        public string Remark { get; set; }
        public string SubstitutionRemark { get; set; }
        public string ConfirmStatus { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
        public string ComponentType { get; set; }

    }
    #endregion

    #region //ERP BOM主件欄位
    public class BOMMC
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
        public string MC001 { get; set; } //主件品號
        public string MC002 { get; set; } //單位
        public string MC003 { get; set; } //小單位
        public int? MC004 { get; set; } //標準批量
        public string MC005 { get; set; } //製令單別
        public string MC006 { get; set; } //變更單別
        public string MC007 { get; set; } //變更單號
        public string MC008 { get; set; } //變更序號
        public string MC009 { get; set; } //版次
        public string MC010 { get; set; } //備註
        public int? MC011 { get; set; } //預留欄位
        public int? MC012 { get; set; } //預留欄位
        public string MC013 { get; set; } //來源
        public string MC014 { get; set; } //預留欄位
        public string MC015 { get; set; } //預留欄位
        public string MC016 { get; set; } //確認碼
        public string MC017 { get; set; } //確認日期
        public string MC018 { get; set; } //確認者
        public string MC019 { get; set; } //簽核狀態碼
        public int? MC020 { get; set; } //列印次數
        public int? MC021 { get; set; } //傳送次數
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
    }
    #endregion

    #region //ERP BOM元件欄位
    public class BOMMD
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
        public string MD001 { get; set; } //主件品號
        public string MD002 { get; set; } //序號
        public string MD003 { get; set; } //元件品號
        public string MD004 { get; set; } //單位
        public string MD005 { get; set; } //小單位
        public double? MD006 { get; set; } //組成用量
        public double? MD007 { get; set; } //底數
        public double? MD008 { get; set; } //損耗率%
        public string MD009 { get; set; } //製程
        public string MD010 { get; set; } //取替代件
        public string MD011 { get; set; } //生效日期
        public string MD012 { get; set; } //失效日期
        public string MD013 { get; set; } //選配預設
        public string MD014 { get; set; } //標準成本計算
        public string MD015 { get; set; } //插件位置
        public string MD016 { get; set; } //備註
        public string MD017 { get; set; } //材料型態
        public int? MD018 { get; set; } //投料時距
        public string MD019 { get; set; } //新插件位置1
        public string MD020 { get; set; } //新插件位置2
        public string MD021 { get; set; } //新插件位置3
        public string MD022 { get; set; } //新插件位置4
        public string MD023 { get; set; } //新插件位置5
        public string MD024 { get; set; } //預留欄位
        public string MD025 { get; set; } //預留欄位
        public string MD026 { get; set; } //預留欄位
        public int? MD027 { get; set; } //預留欄位
        public int? MD028 { get; set; } //預留欄位
        public string MD029 { get; set; } //發放品
        public string MD030 { get; set; } //預留欄位
        public string MD031 { get; set; } //預留欄位
        public string MD032 { get; set; } //確認碼
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
    }
    #endregion
    #endregion

    #region //取替代料欄位
    #region //BM 
    public class BomSubstitution
    {
        public int? SubstitutionId { get; set; }
        public int? BomDetailId { get; set; }
        public int? BomDetailMtlItemId { get; set; }
        public string BomDetailMtlItemNo { get; set; }
        public int? BomId { get; set; }
        public int? BomMtlItemId { get; set; }
        public string BomMtlItemNo { get; set; }
        public string SubstituteStatus { get; set; }
        public int? MtlItemId { get; set; }
        public string MtlItemNo { get; set; }
        public double? Quantity { get; set; }
        public int? SortNumber { get; set; }
        public string Precedence { get; set; }
        public DateTime? EffectiveDate { get; set; }
        public DateTime? ExpirationDate { get; set; }
        public string Remark { get; set; }
        public string TransferStatus { get; set; }
        public DateTime? TransferDate { get; set; }
        public string Status { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
    }
    #endregion

    #region //ERP 取替代料單頭 
    public class BOMMA
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
        public string MA001 { get; set; } //元件
        public string MA002 { get; set; } //主件
        public string MA003 { get; set; } //取替代件
        public string MA004 { get; set; } //備註
        public double? MA005 { get; set; } //預留欄位
        public double? MA006 { get; set; } //預留欄位
        public string MA007 { get; set; } //預留欄位
        public string MA008 { get; set; } //預留欄位
        public string MA009 { get; set; } //預留欄位
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
    }
    #endregion

    #region //ERP 取替代料單身 
    public class BOMMB
    {
        public int? SubstitutionId { get; set; }
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
        public string MB001 { get; set; } //元件
        public string MB002 { get; set; } //主件
        public string MB003 { get; set; } //取替代件
        public string MB004 { get; set; } //品號
        public double? MB005 { get; set; } //數量
        public string MB006 { get; set; } //生效日期
        public string MB007 { get; set; } //失效日期
        public string MB008 { get; set; } //備註
        public string MB009 { get; set; } //取替代件順序
        public string MB010 { get; set; } //優先耗用 DEF:'N
        public double? MB011 { get; set; } //預留欄位
        public double? MB012 { get; set; } //預留欄位
        public string MB013 { get; set; } //預留欄位
        public string MB014 { get; set; } //預留欄位
        public string MB015 { get; set; } //預留欄位
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
    }

    #endregion
    #endregion

    #region //單位欄位
    public class UnitOfMeasure
    {
        public int? UomId { get; set; }
        public int? CompanyId { get; set; }
        public string UomNo { get; set; }
        public string UomName { get; set; }
        public string UomType { get; set; }
        public string UomDesc { get; set; }
        public string TransferStatus { get; set; }
        public DateTime? TransferDate { get; set; }
        public string Status { get; set; }
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
        public string InventoryNo { get; set; }
    }
    #endregion

    #region //使用者欄位
    public class User
    {
        public int? UserId { get; set; }
        public string UserNo { get; set; }
    }
    #endregion

    #region //Erp欄位
    public class Erp
    {
        public string MtlItemTypeNo { get; set; }
        public string MtlItemTypeName { get; set; }
        public string WipPrefix { get; set; }
        public string WipPrefixName { get; set; }
    }
    #endregion

    #region //客戶欄位
    public class Customer
    {
        public int? CustomerId { get; set; }
        public string CustomerNo { get; set; }
    }
    #endregion

    #region //ERP - 品號單位換算 INVMD
    public class INVMD
    {
        public int CompanyId { get; set; }
        public int MtlItemId { get; set; }
        public int MtlItemUomCalculateId { get; set; }
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
        public string MD001 { get; set; } //品號
        public string MD002 { get; set; } //換算單位
        public double? MD003 { get; set; } //換算率分子
        public double? MD004 { get; set; } //換算率分母
        public double? MD005 { get; set; } //預留欄位
        public double? MD006 { get; set; } //預留欄位
        public string MD007 { get; set; } //預留欄位
        public string MD008 { get; set; } //預留欄位
        public string MD009 { get; set; } //預留欄位
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

    #region //品號單位換算
    public class MtlItemUomCalculate
    {
        public int? MtlItemUomCalculateId { get; set; }
        public int? MtlItemId { get; set; }
        public string MtlItemNo { get; set; }
        public int? UomId { get; set; }
        public string UomNo { get; set; }
        public int? ConvertUomId { get; set; }
        public string ConvertUomNo { get; set; }
        public double? SwapDenominator { get; set; }
        public double? SwapNumerator { get; set; }
        public string TransferStatus { get; set; }
        public DateTime? TransferDate { get; set; }
        public DateTime? CreateDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }
        public int? CreateBy { get; set; }
        public int? LastModifiedBy { get; set; }
        public string TransactionType { get; set; }
        public int? ConvertUomIdBefore { get; set; }
        public double? SwapDenominatorBefore { get; set; }
        public double? SwapNumeratorBefore { get; set; }
        public int? TransactionUserId { get; set; }
        public DateTime? TransactionDate { get; set; }
    }
    #endregion

    #region //品號變更
    public class MtlItemChange
    {
        public int MtlItemChangeId { get; set; }
        public int MtlItemId { get; set; }
        public string MtlItemNo { get; set; }
        public string ChangeEdition { get; set; }
        public string OriginalEdition { get; set; }
        public string OriginalMtlItemName { get; set; }
        public string OriginalMtlItemSpec { get; set; }
        public string NewMtlItemName { get; set; }
        public string NewMtlItemSpec { get; set; }
        public List<MtlItemChangeDetail> MtlItemChangeDetails { get; set; }
        public string UpdateUserNo { get; set; }
        public string UpdateUserName { get; set; }
        public string ChangeDate { get; set; }
        public string ConfirmUserNo { get; set; }
        public string ConfirmUserName { get; set; }
        public string ConfirmDate { get; set; }
        public string TransferStatus { get; set; }
        public string ConfirmStatus { get; set; }
        public string TransferDate { get; set; }
        public string Remark { get; set; }
        public int? TotalCount { get; set; }
    }

    public class MtlItemChangeDetail
    {
        public int MtlChangeDetailId { get; set; }
        public int MtlItemChangeId { get; set; }
        public string SortNum { get; set; }
        public string FieldNum { get; set; }
        public string FieldName { get; set; }
        public string NewTextField { get; set; }
        public string OriginalTextField { get; set; }
        public string NewNumericField { get; set; }
        public string OriginalNumericField { get; set; }
        public string NewFieldName { get; set; }
        public string OriginalFieldName { get; set; }
        public string Remark { get; set; }
        public string FieldType { get; set; }
    }

    #region //ERP品號變更單頭
    public class INVTL
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

        public string TL001 { get; set; }
        public string TL002 { get; set; }
        public string TL003 { get; set; }
        public string TL004 { get; set; }
        public string TL005 { get; set; }
        public string TL006 { get; set; }
        public string TL007 { get; set; }
        public string TL008 { get; set; }
        public string TL009 { get; set; }
        public string TL010 { get; set; }
        public string TL011 { get; set; }
        public string TL012 { get; set; }
        public string TL013 { get; set; }
        public string TL014 { get; set; }
        public string TL015 { get; set; }
        public int TL016 { get; set; }
        public int TL017 { get; set; }
        public string TL023 { get; set; }

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

        public int MtlItemId { get; set; }
    }
    #endregion

    #region //ERP品號變更單身
    public class INVTM
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

        public string TM001 { get; set; }
        public string TM002 { get; set; }
        public string TM003 { get; set; }
        public string TM004 { get; set; }
        public string TM005 { get; set; }
        public string TM006 { get; set; }
        public string TM007 { get; set; }
        public float TM008 { get; set; }
        public float TM009 { get; set; }
        public string TM010 { get; set; }
        public string TM011 { get; set; }
        public string TM012 { get; set; }
        public string TM013 { get; set; }
        public string TM014 { get; set; }
        public string TM015 { get; set; }

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
    }
    #endregion

    #endregion

    #region //BOM變更
    public class BomChange
    {
        public int BomChangeId { get; set; }
        public string BomChangeNo { get; set; }
        public string BomChangePrefix { get; set; }
        public string BomSortNum { get; set; }
        public string EmergencyCode { get; set; }
        public string ChangeReason { get; set; }
        public string Remark { get; set; }
        public List<BomChangeDetail> BomChangeDetail { get; set; }
        public string CreatUserNo { get; set; }
        public string CreatUserName { get; set; }
        public string ChangeDate { get; set; }
        public string ConfirmUserNo { get; set; }
        public string ConfirmUserName { get; set; }
        public string ConfirmDate { get; set; }
        public string ConfirmStatus { get; set; }
        public string TransferStatus { get; set; }
        public string TransferDate { get; set; }
        public int? TotalCount { get; set; }
    }

    public class BomChangeDetail
    {
        public int BomChangeDetailId { get; set; }
        public int BomChangeId { get; set; }
        public int BomId { get; set; }
        public string NewMtlItemNo { get; set; }
        public string NewMtlItemName { get; set; }
        public string NewMtlItemSpec { get; set; }
        public string NewMtlItemAttribute { get; set; }
        public string NewMtlItemAttributeName { get; set; }
        public string Unit { get; set; }
        public string SmallUnit { get; set; }
        public string ChangeEdition { get; set; }
        public string StandardLot { get; set; }
        public string MoPrefix { get; set; }
        public string ChangeReason { get; set; }
        public string ConfirmStatus { get; set; }
        public string ConfirmCode { get; set; }
        public string NewRemark { get; set; }
    }

    public class BomChangeDetail2
    {
        public int BomChangeDetailId { get; set; }
        public int BomChangeId { get; set; }
        public int BomId { get; set; }
        public string SortNum { get; set; }
        public int MtlItemId { get; set; }
        public int Unit { get; set; }
        public string SmallUnit { get; set; }
        public string ChangeEdition { get; set; }
        public string StandardLot { get; set; }
        public string MoPrefix { get; set; }
        public string ChangeReason { get; set; }
        public string ConfirmStatus { get; set; }
        public string ConfirmCode { get; set; }
        public string NewRemark { get; set; }
        public string BomChangeSource { get; set; }
        public int OriginalMtlItemId { get; set; }
        public string OriginalUnit { get; set; }
        public string OriginalSmallUnit { get; set; }
        public string OriginalChangeEdition { get; set; }
        public string OriginalStandardLot { get; set; }
        public string OriginalMoPrefix { get; set; }
        public int OriginalBomChangeId { get; set; }
        public string OriginalSortNum { get; set; }
        public string OriginalRemark { get; set; }
        public string Status { get; set; }
        public string CreateDate { get; set; }
        public string LastModifiedDate { get; set; }
        public int CreateBy { get; set; }
        public int LastModifiedBy { get; set; }
    }

    public class BomChangeSubDetail
    {
        public int BomChangeId { get; set; }
        public int BomChangeDetailId { get; set; }
        public int BomChangeSubDetailId { get; set; }
        public int BomDetailId { get; set; }
        public string BomSortNum { get; set; }
        public string NewSubMtlItemNo { get; set; }
        public string NewSubMtlItemName { get; set; }
        public string NewSubMtlItemSpec { get; set; }
        public string NewSubMtlItemAttribute { get; set; }
        public string NewSubMtlItemAttributeName { get; set; }
        public string SubUnit { get; set; }
        public string SubSmallUnit { get; set; }
        public float? Base { get; set; }
        public float? LossRate { get; set; }
        public float? CompositionQuantity { get; set; }
        public int SubFeedingInterval { get; set; }
        public string EffectiveDate { get; set; }
        public string ExpirationDate { get; set; }
        public string ChangeReason { get; set; }
        public string NewRemark { get; set; }
        public string StandardCosting { get; set; }
        public string BomType { get; set; }
        public string ComponentType { get; set; }

    }

    public class BomChangeSubDetail2
    {
        public int BomChangeSubDetailId { get; set; }
        public int BomChangeDetailId { get; set; }
        public int BomChangeId { get; set; }
        public int BomDetailId { get; set; }
        public int MtlItemId { get; set; }
        public string BomSortNum { get; set; }
        public int Unit { get; set; }
        public string SmallUnit { get; set; }
        public float CompositionQuantity { get; set; }
        public float Base { get; set; }
        public float LossRate { get; set; }
        public string SubstituteStatus { get; set; }
        public string EffectiveDate { get; set; }
        public string ExpirationDate { get; set; }
        public string OptionalPreset { get; set; }
        public string StandardCosting { get; set; }//標準成本
        public string BomType { get; set; }//材料型態
        public int FeedingInterval { get; set; }
        public string ChangeReason { get; set; }
        public string NewRemark { get; set; }
        public int OriginalBomDetailId { get; set; }
        public int OriginalMtlItemId { get; set; }
        public string OriginalBomSortNum { get; set; }
        public int OriginalUnit { get; set; }
        public string OriginalSmallUnit { get; set; }
        public float OriginalCompositionQuantity { get; set; }
        public float OriginalBase { get; set; }
        public float OriginalLossRate { get; set; }
        public string OriginalSubstituteStatus { get; set; }
        public string OriginalEffectiveDate { get; set; }
        public string OriginalExpirationDate { get; set; }
        public string OriginalOptionalPreset { get; set; }
        public string OriginalStandardCosting { get; set; }
        public string OriginalBomType { get; set; }
        public int OriginalFeedingInterval { get; set; }
        public string Originalemark { get; set; }
        public string Provide { get; set; }
        public string Status { get; set; }
        public string CreateDate { get; set; }
        public string LastModifiedDate { get; set; }
        public int CreateBy { get; set; }
        public int LastModifiedBy { get; set; }
        public string DeleteStatus { get; set; }
        public string ComponentType { get; set; }
        public string OriginalComponentType { get; set; }

    }

    #region //ERP
    public class BOMTA
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

        public string TA001 { get; set; }
        public string TA002 { get; set; }
        public string TA003 { get; set; }
        public string TA004 { get; set; }
        public string TA005 { get; set; }
        public string TA006 { get; set; }
        public string TA007 { get; set; }
        public int TA008 { get; set; }
        public string TA009 { get; set; }
        public string TA010 { get; set; }
        public string TA011 { get; set; } //defaut N-已簽核
        public int TA012 { get; set; }

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

        //public string CFIELD01 { get; set; }

    }

    public class BOMTB
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

        public string TB001 { get; set; } //4101
        public string TB002 { get; set; }
        public string TB003 { get; set; }
        public string TB004 { get; set; }
        public string TB005 { get; set; }
        public string TB006 { get; set; }
        public string TB007 { get; set; }
        public int TB008 { get; set; }
        public string TB009 { get; set; }
        public string TB010 { get; set; }
        public string TB011 { get; set; }
        public string TB012 { get; set; }
        public string TB013 { get; set; }
        public string TB016 { get; set; }
        public string TB104 { get; set; }
        public string TB105 { get; set; }
        public string TB106 { get; set; }
        public string TB107 { get; set; }
        public int TB108 { get; set; }
        public string TB109 { get; set; }
        public string TB110 { get; set; }
        public string TB111 { get; set; }
        public string TB112 { get; set; }
        public string TB113 { get; set; }

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

        public int BomId { get; set; }
        public int BomChangeDetailId { get; set; }

    }

    public class BOMTC
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
        public double TC008 { get; set; }
        public double TC009 { get; set; }
        public double TC010 { get; set; }
        public string TC011 { get; set; }
        public string TC012 { get; set; }
        public string TC013 { get; set; }
        public string TC014 { get; set; }
        public string TC015 { get; set; }
        public string TC016 { get; set; }
        public string TC017 { get; set; }
        public string TC018 { get; set; }
        public string TC019 { get; set; }
        public double TC020 { get; set; }
        public string TC021 { get; set; }
        public string TC022 { get; set; }
        public string TC023 { get; set; }
        public string TC024 { get; set; }
        public string TC025 { get; set; }
        public string TC029 { get; set; }
        public string TC032 { get; set; }
        public string TC104 { get; set; }
        public string TC105 { get; set; }
        public string TC106 { get; set; }
        public string TC107 { get; set; }
        public double TC108 { get; set; }
        public double TC109 { get; set; }
        public double TC110 { get; set; }
        public string TC111 { get; set; }
        public string TC112 { get; set; }
        public string TC113 { get; set; }
        public string TC114 { get; set; }
        public string TC115 { get; set; }
        public string TC116 { get; set; }
        public string TC117 { get; set; }
        public string TC119 { get; set; }
        public double TC120 { get; set; }
        public string TC121 { get; set; }
        public string TC122 { get; set; }
        public string TC123 { get; set; }
        public string TC124 { get; set; }
        public string TC125 { get; set; }
        public string TC129 { get; set; }
        public string TC132 { get; set; }

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
        public string DeleteStatus { get; set; }

        public int MtlItemId { get; set; }
        public int BomDetailId { get; set; }
        public int Unit { get; set; }
        public int BomChangeDetailId { get; set; }

    } 
    #endregion

    #endregion

    #region //報價項目
    public class QuotationItem
    {
        public int DfmQiId { get; set; }
        public int ParentDfmQiId { get; set; }
        public int DfmId { get; set; }
        public string DfmNo { get; set; }
        public int ModeId { get; set; }
        public string ModeName { get; set; }
        public string MfgFlag { get; set; }
        public string QuoteModel { get; set; }
        public float? QuoteNum { get; set; }
        public int? Level { get; set; }
        public string Route { get; set; }
        public int? Sort { get; set; }
        public string MaterialStatus { get; set; }
        public string DfmQiOSPStatus { get; set; }
        public string DfmQiProcessStatus { get; set; }
        public string DfmProcessStatus { get; set; }
        public int DfmItemCategoryId { get; set; }
        public string DfmItemCategoryName { get; set; }
        public string DfmQuotationName { get; set; }


    }
    #endregion

    #region //庫存整合+新舊品號變更
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
        public string CreateDate { get; set; }
        public int? CreateBy { get; set; }
        public string UserName { get; set; }
        public string UserNo { get; set; }
        public string NewMtlItemNo { get; set; }
        public string NewMtlItemName { get; set; }
        public string NewMtlItemSpec { get; set; }
        public string NewCreateDate { get; set; }
        public int? NewCreateBy { get; set; }
        public string NewUserName { get; set; }
        public string NewUserNo { get; set; }
        public int? TotalCount { get; set; }
        public string Adjustment { get; set; }
        public string ReturnMsg { get; set; }

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

    #region //新舊品號調整用

    #region //新舊品號對照表
    public class MtlItemReference
    {
        public string CREATE_DATE { get; set; }
        public string CREATE_TIME { get; set; }
        public string CREATE_AP { get; set; }
        public DateTime ExpirationDate { get; set; }
        public int MtlItemId { get; set; }
        public string MtlItemNo { get; set; }
        public string NewMtlItemNo { get; set; }
        public string NewMtlItemName { get; set; }
        public string NewMtlItemSpec { get; set; }
        public string JmoMtlItemName { get; set; }
        public string JmoMtlItemSpec { get; set; }
        public string DgMtlItemName { get; set; }
        public string DgMtlItemSpec { get; set; }
        public string EtgMtlItemName { get; set; }
        public string EtgMtlItemSpec { get; set; }
        public string ReturnMsg { get; set; }
        public string Adjustment { get; set; } //庫存調整狀態(S:不動作/N:已拋轉/Y:已確認) 
        public string TransferStatus { get; set; }
        public string ManyToOne { get; set; } //多個舊品號對應新品號(Y/N)
        public string OnlyVersion { get; set; } //新品號資料是否有一個版本(Y/N)
        public string MultipleVersions { get; set; } //是否存在多版本在不同公司(Y/N)
    }
    #endregion

    #region //庫存異動相關
    public class MCDetail
    {
        public string MC002 { get; set; }
        public string MC007 { get; set; }
    }
    public class MCDetails
    {
        public MCDetail[] MCDetail { get; set; }
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
        public string MtlItemNo { get; set; }
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

    public class DfmQuotationTree
    {
        public string value { get; set; }
        public string id { get; set; }
        public bool opened { get; set; }
        public List<DfmQuotationTree> items { get; set; }
        public int level { get; set; }
        public int DfmItemCategoryId { get; set; }
        public int ModeId { get; set; }
        public string MfgFlag { get; set; }
        public string QuoteModel { get; set; }
        public float? QuoteNum { get; set; }
        public string MaterialStatus { get; set; }
        public string DfmQiOSPStatus { get; set; }
        public string DfmQiProcessStatus { get; set; }
        public string DfmProcessStatus { get; set; }
        public string DfmQuotationName { get; set; }

    }


    #region //Excel匯入用
    public class BmFileInfo
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
    public class ErpLineMapping
    {
        public string CategoryCode { get; set; }
        public string CategoryName { get; set; }
        public string ErpLineCode { get; set; }
        public string ErpLineName { get; set; }
        public float MhLineCost { get; set; }
        public float McLineCost { get; set; }
        public float MhLaborCost { get; set; }
        public float McLaborCost { get; set; }
    }

    public class StockInDetail
    {
        public string ReceiptDate { get; set; }
        public string ErpFull { get; set; }
        public int MtlItemId { get; set; }
        public string MtlItemNo { get; set; }
        public string MtlItemName { get; set; }
        public string MtlItemSpec { get; set; }
        public int CellMtlItemId { get; set; }
        public string CycleTime { get; set; }
        public string MoldWeight { get; set; }
        public string Cavity { get; set; }
        public double TheoreticalQty { get; set; }
        public double ProcessingFee { get; set; }
        public int ReceiptQty { get; set; }
        public int AcceptQty { get; set; }
        public double PassRate { get; set; }
        public int ScriptQty { get; set; }
        public double DefectRate { get; set; }
        public string Material { get; set; }
        public string CellMtlItemNo { get; set; }
        public string CellMtlItemName { get; set; }
        public string CellMtlItemSpec { get; set; }
        public double UnitPrice { get; set; }
        public double ManCost { get; set; }
        public double MachineCost { get; set; }
    }
    public class ItemCatDept
    {
        public string TypeTwo { get; set; }
        public string TypeThree { get; set; }
        public string InventoryNo { get; set; }
        public string CatDept { get; set; }
    }

    public class TempOrderDetail
    {
        
        public string TF003 { get; set; }
        public string DocNum { get; set; }
        public string TF024 { get; set; }
        public string TF041 { get; set; }
        public string TF005 { get; set; }
        public string TF006 { get; set; }
        public string TF007 { get; set; }
        public string ME002 { get; set; }
        public string TF008 { get; set; }
        public string MV002 { get; set; }
        public string TF009 { get; set; }
        public string MB002 { get; set; }
        public string TF011 { get; set; }
        public string TaxType { get; set; }
        public string TF012 { get; set; }
        public string TF013 { get; set; }
        public string TF014 { get; set; }
        public string TG004 { get; set; }
        public string TG005 { get; set; }
        public string TG006 { get; set; }
        public string CatDept { get; set; }//责任部门
        public string MB006 { get; set; }//类别二
        public string MB007 { get; set; }//类别三
        public string TG007 { get; set; }//转出库别
        public string TG008 { get; set; }
        public string TG027 { get; set; }
        public string TG035 { get; set; }
        public string TG036 { get; set; }
        public string TempOutQty { get; set; }
        public string TempOutSaleQty { get; set; }
        public string TempOutReturnQty { get; set; }
        public string TempInQty { get; set; }
        string TempInSaleQty { get; set; }
        public string TempInReturnQty { get; set; }
        public string TempOutPackQty { get; set; }
        public string TempOutPackSaleQty { get; set; }
        public string TempOutPackReturnQty { get; set; }
        public string TempInPackQty { get; set; }
        public string TempInPackSaleQty { get; set; }
        public string TempInPackReturnQty { get; set; }
        public string TempOutSpareQty { get; set; }
        public string TempOutSpareSaleQty { get; set; }
        public string TempOutSpareReturnQty { get; set; }
        public string TempInSpareQty { get; set; }
        public string TempInSpareSaleQty { get; set; }
        public string TempInSpareReturnQty { get; set; }
        public string TempOutSparePackQty { get; set; }
        public string TempOutSparePackSaleQty { get; set; }
        public string TempOutSparePackReturnQty { get; set; }
        public string TempInSparePackQty { get; set; }
        public string TempInSparePackSaleQty { get; set; }
        public string TempInSparePackReturnQty { get; set; }
        public string TG010 { get; set; }
        public string TG011 { get; set; }
        public string CloseCode { get; set; }
        public string TG031 { get; set; }
        public string TG012 { get; set; }
        public string TempOutAmount { get; set; }
        public string TempInAmount { get; set; }
        public string TG017 { get; set; }
        public string TG025 { get; set; }
        public string TG026 { get; set; }
        public string SourceDocNum { get; set; }
        public string TG019 { get; set; }
        public string TG018 { get; set; }
        public string NB002 { get; set; }
        public string TempOut { get; set; }
        public string TypeTempIn { get; set; }
        public string TypeTwoName02 { get; set; }
        public string TypeTwoName03 { get; set; }
    }

    public class TempReturnOrderDetail
    {

        public string TH003 { get; set; }
        public string DocNum { get; set; }
        public string TH023 { get; set; }
        public string TH005 { get; set; }
        public string TH006 { get; set; }
        public string TH007 { get; set; }
        public string ME002 { get; set; }
        public string TH008 { get; set; }
        public string MV002 { get; set; }
        public string TH009 { get; set; }
        public string MB002 { get; set; }
        public string TH011 { get; set; }
        public string TH010 { get; set; }
        public string TH012 { get; set; }
        public string TH013 { get; set; }
        public string TI004 { get; set; }
        public string TI005 { get; set; }
        public string TI006 { get; set; }
        public string CatDept { get; set; }
        public string MB006 { get; set; }
        public string TypeTwoName02 { get; set; }
        public string MB007 { get; set; }
        public string TypeTwoName03 { get; set; }
        public string TI007 { get; set; }
        public string TempOut { get; set; }
        public string TI008 { get; set; }
        public string TypeTempIn { get; set; }
        public string TI026 { get; set; }
        public string TI027 { get; set; }
        public string TI009 { get; set; }
        public string TI023 { get; set; }
        public string TI035 { get; set; }
        public string TI036 { get; set; }
        public string TI010 { get; set; }
        public string TI011 { get; set; }
        public string TI024 { get; set; }
        public string TI012 { get; set; }
        public string TI013 { get; set; }
        public string TI017 { get; set; }
        public string TI018 { get; set; }
        public string TI019 { get; set; }
        public string SourceDocNum { get; set; }
        public string TI021 { get; set; }
        public string TI020 { get; set; }
        
    }

    public class ItemCatDepts
    {
        public int ItemCatDeptId { get; set; }
        public string TypeTwo { get; set; }
        public string TypeTwoName { get; set; }
        public string TypeThree { get; set; }
        public string TypeThreeName { get; set; }
        public string InventoryNo { get; set; }
        public string InventoryName { get; set; }
        public string CatDept { get; set; }
        public int? TotalCount { get; set; }
    }

    public class MtlItemCategory
    {
        public string MtlItemTypeNo { get; set; }
        public string MtlItemTypeName { get; set; }
        
    }
      
    public class MtlItemInventory
    {
        public string InventoryNo { get; set; }
        public string InventoryName { get; set; }

    }



    #endregion

}
