/* MRPII Manufacturing Resource Planning, 非原本之Materials Requirements Planning) */
BEGIN /*基本資料*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Forwarder' AND uid = 13) > 0 DROP TABLE MRP.Demand
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Forwarder' AND uid = 13) > 0 DROP TABLE MRP.DemandLine
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Forwarder' AND uid = 13) > 0 DROP TABLE MRP.DemandLineDtl

END /*基本資料*/


--******************************************************************************************************************
BEGIN /*MRP*/

CREATE TABLE MRP.Demand( /*MRP計劃單頭*/
 CompanyId					int					NOT NULL,						--公司
 DemandId					int					identity,
 DemandName					nvarchar(200)		NOT NULL,						--計劃名稱
 DemandDate					datetime			NOT NULL,						--計劃日期
 EndDate					datetime			NOT NULL,						--需求截止日
 DemandDesc					nvarchar(500)		NULL,							--需求備註
 [Status]					nvarchar(2)			NOT NULL	DEFAULT '0',		--狀態碼 0:New, 1:需求已匯入, 2.需求計算完成
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Demand PRIMARY KEY (DemandId)
)

CREATE TABLE MRP.DemandLine( /*MRP計劃單身*/
 DemandId					int					NOT NULL,
 DemandLineId				int					identity,
 DemandPriority				int					NOT NULL	DEFAULT 1,			--以訂單計劃出貨日依順序排, 相同日期由業務決定順序
 SourceType					nvarchar(2)			NOT NULL,						--SO:訂單需求, FO:銷售預測 ESO:電子平台需求
 SourceId					int					NOT NULL,						--來源單號ID SCM.SalesOrder.SoId
 SourceNo					nvarchar(32)		NOT NULL,						--需求單號
 SourceSeq					nvarchar(10)		NOT NULL,						--單身序號
 MtlItemId					int					NOT NULL,						--品號ID
 MtlItemNo					nvarchar(30)		NOT NULL,						--品號
 MakeType					nvarchar(2)			NOT NULL,						--品號/屬性
 Quantity					float				NOT NULL,						--需求數量(淨需求)
 CustomerId					int					NOT NULL,						--客戶Id
 ScheduleDate				datetime			NOT NULL,						--訂單計劃出貨日
 ExpectDeliveryDate			datetime			NULL,							--預計出貨日(系統計算)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DemandLine PRIMARY KEY (DemandLineId),
 CONSTRAINT FK1_DemandLine FOREIGN KEY (DemandId) REFERENCES MRP.Demand(DemandId),
 CONSTRAINT FK2_DemandLine FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId)
)

CREATE TABLE MRP.DemandLineDtl( /*MRP計劃需求明細(已展BOM)*/
 DemandLineId				int					NOT NULL,
 LineDtlId					int					identity,
 BomLevel					int					NOT NULL,						--Bom階層
 MtlItemId					int					NOT NULL,						--需求料號ID
 ParentMtlItemId			int					NOT NULL,						--上一階料號
 MtlItemNo					nvarchar(30)		NOT NULL,						--Component品號
 MakeType					nvarchar(2)			NOT NULL,						--P:採購件, M:自製件
 Quantity					float				NOT NULL,						--需求數量
 CompositionQuantity		float				NOT NULL,						--單位用量
 Base						float				NOT NULL,						--底數
 ScheduleDate				datetime			NOT NULL,						--開始日期(需求展下來之原始時間)
 WorkDays					int					NOT NULL	DEFAULT 0,			--工作天數(採購件為L/T, 製造件為標準製造時間)
 ExpectFinishDate			datetime			NULL,							--預計完成日期, 再先推算到採購L/T, 到料日再加上標準製造時間
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DemandLineDtl PRIMARY KEY (LineDtlId),
 CONSTRAINT FK1_DemandLineDtl FOREIGN KEY (DemandLineId) REFERENCES MRP.DemandLine(DemandLineId),
 CONSTRAINT FK2_DemandLineDtl FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId)
)

CREATE TABLE MRP.DemandSuggestPr( /*MRP建議請購清單*/
 DemandPrId					int					identity,
 VendorId					int					NULL,
 CurrencyCode				nvarchar(30)		NOT NULL,
 MtlItemId					int					NOT NULL,
 SupplierId					int					NOT NULL,
 RequisitionDate			datetime			NOT NULL,
 PrCurrency					nvarchar(4)				NULL,						--TB050 請購幣別 >> 參照 ERP幣別匯率檔單頭(CMSMF)
 PrExchangeRate				float					NULL,						--TB044 請購匯率
 PrUnitPrice				float					NULL,						--TB049 請購單位金額
 PrPrice					float					NULL,						--TB051 請購金額
 PrPriceTw					float					NULL,						--TB045 請購本幣金額 
 PoUomId					int						NULL,						--TB015 採購單位(Unit of Measure) >> 參照 ERP品號換算單位檔(INVMD)
 PoQty						int						NULL,						--TB014 採購數量
 PoCurrency					nvarchar(4)				NULL,						--TB016 採購幣別 >> 參照 ERP幣別匯率檔單頭(CMSMF)
 PoUnitPrice				float					NULL,						--TB017 採購單位金額
 PoPrice					float					NULL,						--TB018 採購金額
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DemandLineDtl PRIMARY KEY (LineDtlId),
 CONSTRAINT FK1_DemandLineDtl FOREIGN KEY (DemandLineId) REFERENCES MRP.DemandLine(DemandLineId),
 CONSTRAINT FK2_DemandLineDtl FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId)
)

CREATE TABLE MRP.MtlItemUseQty( /*MRP計劃需求明細(已展BOM)*/
 DemandLineId				int					NOT NULL,
 MtlItemNo					nvarchar(30)		NOT NULL,
 UsePoQty					float					NULL,						--已扣在途採購量
 UseWipQty					float					NULL,						--已扣在製量
 UseOnhandQty				float					NULL						--已扣庫存量
)

END 