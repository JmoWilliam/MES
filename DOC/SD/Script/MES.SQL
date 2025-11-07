BEGIN /*基本資料*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ProcessMachine' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES'))> 0 DROP TABLE MES.ProcessMachine
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ProcessProductionUnit' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.ProcessProductionUnit
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ProcessParameter' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.ProcessParameter
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DeviceMachine' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.DeviceMachine
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Device' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.Device
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ProdUnit' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.ProdUnit 
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Jig' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.Jig
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MachineJig' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.MachineJig
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MachineAsset' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.MachineAsset
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MachineShop' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.MachineShop
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Machine' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.Machine

IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Process' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.Process
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'TimeInterval' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.TimeInterval
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ProdModeShift' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.ProdModeShift
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ProdMode' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.ProdMode
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'WorkShop' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.WorkShop
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'KnivesSupplier' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.KnivesSupplier
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Warehouse' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.Warehouse
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'WarehouseLocation' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.WarehouseLocation
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ProcessQcItem' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.ProcessQcItem
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'QcItem' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.QcItem
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'LabelPrintMachine' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.LabelPrintMachine /*標籤機資料*/

IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'UserEventItem' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.UserEventItem 
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MachineEventItem' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.MachineEventItem 
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ProcessEventItem' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.ProcessEventItem 

IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'LaserEngravedUser' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.LaserEngravedUser 

IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'TrayTemp' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.TrayTemp

END /*基本資料*/

BEGIN /*途程管理*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'RoutingItemProcess' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.RoutingItemProcess
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'RoutingProcessTemplate' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.RoutingProcessTemplate
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'RoutingItem' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.RoutingItem
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'RoutingProcessItem' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.RoutingProcessItem
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'RoutingProcess' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.RoutingProcess
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Routing' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.Routing
END /*途程管理*/

BEGIN /*生產管理*/


IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ProdDemandDetail' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.ProdDemandDetail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ProdDemand' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.ProdDemand


IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MoRouting' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.MoRouting
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'WomDetail' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.WomDetail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'WoModification' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.WoModification
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MoMtlSettingAttribute' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.MoMtlSettingAttribute
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MoMtlSetting' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.MoMtlSetting
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MoSetting' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.MoSetting
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MoProcess' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.MoProcess
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ManufactureOrder' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.ManufactureOrder
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'WoDetail' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.WoDetail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'WipOrder' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.WipOrder

IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MoProcessTransaction' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.MoProcessTransaction
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'BarcodeQcRecord' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.BarcodeQcRecord
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'BarcodeTransfer' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.BarcodeTransfer

IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'BarcodeComponentAttribute' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.BarcodeComponentAttribute
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'BarcodeProcessAttributeLog' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.BarcodeProcessAttributeLog
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'BpalDetail' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.BpalDetail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'BarcodeProcessAttribute' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.BarcodeProcessAttribute
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'BarcodeProcess' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.BarcodeProcess
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'BarcodePrint' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.BarcodePrint
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'BarcodeComponent' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.BarcodeComponent
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'BarcodeComponentLog' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.BarcodeComponentLog
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'BarcodeLog' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.BarcodeLog
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Barcode' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.Barcode

IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'UserEvent' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.UserEvent
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MachineEvent' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.MachineEvent
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'BarcodeProcessEvent' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.BarcodeProcessEvent


END /*生產管理*/

BEGIN /*託外生產*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'OspBarcode' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.OspBarcode
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'OspDetail' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.OspDetail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'OutsourcingProduction' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.OutsourcingProduction
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'OspReceiptBarcode' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.OspReceiptBarcode
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'OspReceiptDetail' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.OspReceiptDetail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'OspReceipt' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.OspReceipt


END /*託外生產*/

BEGIN /*倉儲管理*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MtlTransaction' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.MtlTransaction
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MrBarcodeRegister' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.MrBarcodeRegister
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MrDetail' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.MrDetail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MrWipOrder' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.MrWipOrder
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MaterialRequisition' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.MaterialRequisition

IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'PiBarcode' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.PiBarcode
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'PiDetail' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.PiDetail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ProductionInventory' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.ProductionInventory
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MrBarcodeTransferLog' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.MrBarcodeTransferLog
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MweBarcodeTransfer' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.MweBarcodeTransfer


END /*倉儲管理*/

BEGIN /*編程管理*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'CncProgram' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.CncProgram
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'CncParameter' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.CncParameter
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'CncProgramFile' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.CncProgramFile
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'CncParameterOption' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.CncParameterOption

IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'CncProgramLog' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.CncProgramLog
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'CncProgramRequest' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.CncProgramRequest
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'CncProgramResponsest' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.CncProgramResponsest
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'CncProgramWork' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.CncProgramWork
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'CncProgramWorkFiles' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.CncProgramWorkFiles
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'CncResponsestFileLog' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.CncResponsestFileLog

END /*編程管理*/

BEGIN /*工具管理*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ToolMachineParameter' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.ToolMachineParameter
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ToolMachine' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.ToolMachine
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ToolClassIdSetting' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.ToolClassIdSetting
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ToolSpecLog' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.ToolSpecLog
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Tool' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.Tool
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ToolModelSpec' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.ToolModelSpec
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ToolModel' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.ToolModel
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ToolCategory' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.ToolCategory
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ToolClass' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.ToolClass
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ToolSpec' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.ToolSpec
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ToolGroup' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.ToolGroup 

END /*工具管理*/

BEGIN /*製程參數/途程量測項目設定*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ProcessParameterQcItem' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.ProcessParameterQcItem
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'RoutingItemQcItem' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.RoutingItemQcItem

END /*製程參數/途程量測項目設定*/

BEGIN /*Tray相關*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'TrayDismantle' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.TrayDismantle
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'TrayMtlItemNg' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.TrayMtlItemNg

END /*製程參數/途程量測項目設定*/

BEGIN /*MTF相關*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MtfProcessTemp' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'MES')) > 0 DROP TABLE MES.MtfProcessTemp
END /*製程參數/途程量測項目設定*/


--******************************************************************************************************************
BEGIN /*基本資料*/
CREATE TABLE MES.TimeInterval( /*時間區間資料*/
 IntervalId					int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 BeginTime					time				NOT NULL,						--開始時間
 EndTime					time				NOT NULL,						--結束時間
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_TimeInterval PRIMARY KEY (IntervalId),
 CONSTRAINT AK1_TimeInterval UNIQUE (CompanyId, BeginTime),
 CONSTRAINT FK1_TimeInterval FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)

CREATE TABLE MES.ProdMode( /*生產模式資料*/
 ModeId						int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 ModeNo						nvarchar(50)		NOT NULL,						--模式代碼
 ModeName					nvarchar(100)		NOT NULL,						--模式名稱
 ModeDesc					nvarchar(100)		NOT NULL,						--模式描述
 BarcodeCtrl				nvarchar(1)			NOT NULL	DEFAULT 'Y',		--過站是否要條碼控制(模造只輸入數量)
 ScrapRegister				nvarchar(1)			NOT NULL	DEFAULT 'N',		--報廢品是否入報廢倉
 OQcCheckType				nvarchar(2)			NOT NULL	DEFAULT 'Y',		--未完成出貨檢是否可以入庫
 [Source]					nvarchar(2)				NULL,						--發料來源 MR:領退料單 BP:批量條碼
 PVTQCFlag					nvarchar(2)				NULL,						--量產前是否需要先進行試產檢
 NgToBarcode				nvarchar(2)				NULL,						--是否要條碼控管
 TrayBarcode				nvarchar(2)				NULL,						--Tray盤上是否有雷刻條碼
 LotStatus					nvarchar(2)				NULL,						--未完成出貨檢是否可以過站
 OutputBarcodeFlag			nvarchar(2)				NULL,						--製程中是否會產出新條碼(數量過站用，非數量過站不可選，預設為N)	
 MrType						nvarchar(2)				NULL,						--是否可以領取上個月的料 Y:可領取 N:不可領取
 RelationshipFlag			nvarchar(2)				NULL,						--是否需要檢核前後製令用料關聯性
 --CavityStatus				nvarchar(2)			NOT NULL,						--是否需要綁定穴號(批量條碼用) >> BAS.[Status] Y:需要綁定穴號；N:不需綁定穴號
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ProdMode PRIMARY KEY (ModeId),
 CONSTRAINT AK1_ProdMode UNIQUE (ModeNo),
 CONSTRAINT FK1_ProdMode FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)

CREATE TABLE MES.ProdModeShift( /*生產模式班次資料*/
 ModeShiftId				int					identity,
 ModeId						int					NOT NULL,						--所屬生產模式
 ShiftId					int					NOT NULL,						--所屬班次
 EffectiveDate				datetime			NOT NULL,						--生效日
 ExpirationDate				datetime			    NULL,						--失效日
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ProdModeShift PRIMARY KEY (ModeShiftId),
 CONSTRAINT FK1_ProdModeShift FOREIGN KEY (ModeId) REFERENCES MES.ProdMode(ModeId),
 CONSTRAINT FK2_ProdModeShift FOREIGN KEY (ShiftId) REFERENCES BAS.[Shift](ShiftId)
)

CREATE TABLE MES.Process( /*製程資料*/
 ProcessId					int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 ProcessNo					nvarchar(50)		NOT NULL,						--製程代碼
 ProcessName				nvarchar(100)		NOT NULL,						--製程名稱
 ProcessDesc				nvarchar(100)		NOT NULL,						--製程描述
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Process PRIMARY KEY (ProcessId),
 CONSTRAINT AK1_Process UNIQUE (CompanyId, ProcessNo),
 CONSTRAINT AK2_Process UNIQUE (CompanyId, ProcessName),
 CONSTRAINT FK1_Process FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
)

CREATE TABLE MES.WorkShop( /*車間資料*/
 ShopId						int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 ShopNo						nvarchar(10)		NOT NULL,						--車間代碼
 ShopName					nvarchar(100)		NOT NULL,						--車間名稱
 ShopDesc					nvarchar(100)		NOT NULL,						--車間描述
 [Location]					nvarchar(100)			NULL,						--位置
 [Floor]					nvarchar(100)			NULL,						--樓層
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_WorkShop PRIMARY KEY (ShopId),
 CONSTRAINT AK1_WorkShop UNIQUE(CompanyId, ShopNo),
 CONSTRAINT AK2_WorkShop UNIQUE(CompanyId, ShopName),
 CONSTRAINT FK1_WorkShop FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)

CREATE TABLE MES.UserEventItem( /*人員事件資料項目*/
 UserEventItemId			int					identity,
 CompanyId					int					NOT NULL,
 UserEventItemNo			nvarchar(20)		NOT NULL,						--人員事件編號
 UserEventItemName			nvarchar(20)		NOT NULL,						--人員事件名稱
  [Status]					nvarchar(2)			NOT NULL	DEFAULT 'A',		--狀態碼 A:啟用;S:停用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_UserEventItem PRIMARY KEY (UserEventItemId),
 CONSTRAINT FK1_UserEventItem FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)

CREATE TABLE MES.Machine( /*機台資料*/
 MachineId					int					identity,
 ShopId						int					NOT NULL,						--所屬車間
 MachineNo					nvarchar(10)		NOT NULL,						--機台代碼
 MachineName				nvarchar(100)		NOT NULL,						--機台名稱
 MachineDesc				nvarchar(100)		NOT NULL,						--機台描述
 CncMachineNo				int					NULL,							--編程機台代碼
 IOTtableName				nvarchar(100)		NULL,							--物連網對應table
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Machine PRIMARY KEY (MachineId),
 CONSTRAINT AK1_Machine UNIQUE (ShopId, MachineNo),
 CONSTRAINT FK1_Machine FOREIGN KEY (ShopId) REFERENCES MES.WorkShop(ShopId)
)

CREATE TABLE MES.MachineEventItem( /*機台事件資料項目*/
 MachineEventItemId			int					identity,
 CompanyId					int					NOT NULL,
 MachineId					int					NOT NULL,						--機台
 MachineEventNo				nvarchar(20)		NOT NULL,						--機台事件編號
 MachineEventName			nvarchar(20)		NOT NULL,						--機台事件名稱
 [Status]					nvarchar(2)			NOT NULL	DEFAULT 'A',		--狀態碼 A:啟用;S:停用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MachineEventItem PRIMARY KEY (MachineEventItemId),
 CONSTRAINT FK1_MachineEventItem FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_MachineEventItem FOREIGN KEY (MachineId) REFERENCES MES.Machine(MachineId)
)

CREATE TABLE MES.MachineAsset( /*機台資產資料*/
 MachineId					int					NOT NULL,						--所屬機台
 AssetNo					nvarchar(30)		NOT NULL,						--資產編號
 MachineType				nvarchar(2)			NOT NULL,						--機台類別 1:加工機台 2:量測機器
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT AK1_MachineAsset UNIQUE (MachineId, AssetNo),
 CONSTRAINT FK1_MachineAsset FOREIGN KEY (MachineId) REFERENCES MES.Machine(MachineId)
)

CREATE TABLE MES.Jig( /*治具資料*/
 JigId						int					identity,
 CompanyId					int					NOT NULL,						--所屬公司		
 JigNo						nvarchar(100)		NOT NULL,						--治具代碼
 JigName					nvarchar(100)		NOT NULL,						--治具名稱
 JigDesc					nvarchar(100)		NOT NULL,						--治具描述
 JigLimit					int					NOT	NULL,						--治具容納上限
 AvailableFlag				nvarchar(1)		    NOT NULL	DEFAULT 'Y',		--可用狀態
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Jig PRIMARY KEY (JigId),
 CONSTRAINT AK1_Jig UNIQUE (CompanyId, JigNo),
 CONSTRAINT FK1_Jig FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)

CREATE TABLE MES.MachineJig( /*機台治具資料*/
 JigId						int					NOT NULL,						--所屬治具
 MachineId					int						NULL,						--所屬機台
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT AK1_MachineJig UNIQUE (JigId, MachineId),
 CONSTRAINT FK1_MachineJig FOREIGN KEY (JigId) REFERENCES MES.Jig(JigId),
 CONSTRAINT FK2_MachineJig FOREIGN KEY (MachineId) REFERENCES MES.Machine(MachineId)
)

CREATE TABLE MES.Device( /*裝置資料*/
 DeviceId					int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 DeviceType					nvarchar(2)			NOT NULL,						--設備類別 > 1：電腦，2：Android平板
 DeviceName					nvarchar(100)			NULL,						--設備名稱
 DeviceDesc					nvarchar(100)			NULL,						--設備描述
 DeviceIdentifierCode		nvarchar(50)			NULL,						--設備識別碼 > IP address
 DeviceAuthority			nvarchar(2)			NOT NULL,						--設備權限 > N：Noramal，A：Admin
 InventoryId				int						NULL,						--綁定倉別
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Device PRIMARY KEY (DeviceId),
 CONSTRAINT FK1_Device FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_Device FOREIGN KEY (InventoryId) REFERENCES SCM.Inventory(InventoryId)
)

CREATE TABLE MES.DeviceMachine( /*裝置機台綁定*/
 DeviceId					int					NOT NULL,						--所屬裝置
 MachineId					int					NOT NULL,						--所屬機台
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT FK1_DeviceMachine FOREIGN KEY (DeviceId) REFERENCES MES.Device(DeviceId),
 CONSTRAINT FK2_DeviceMachine FOREIGN KEY (MachineId) REFERENCES MES.Machine(MachineId)
)

CREATE TABLE MES.DeciveInventory( /*裝置綁定庫別*/
 DiId						int					identity,
 DeviceId					int					NOT NULL,						--所屬裝置
 InventoryId				int					NOT NULL,						--所屬倉別
 SortNumber					int					NOT NULL,						--優先順序排列
 [Status]					nvarchar(2)			NOT NULL	DEFAULT 'A',		--狀態碼 A:啟用 S:停用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DeciveInventory PRIMARY KEY (DiId),
 CONSTRAINT AK1_DeciveInventory UNIQUE (DeviceId, InventoryId),
 CONSTRAINT AK2_DeciveInventory UNIQUE (DeviceId, SortNumber),
 CONSTRAINT FK1_DeciveInventory FOREIGN KEY (DeviceId) REFERENCES MES.Device(DeviceId),
 CONSTRAINT FK2_DeciveInventory FOREIGN KEY (InventoryId) REFERENCES SCM.Inventory(InventoryId)
)

CREATE TABLE MES.ProdUnit( /*生產單元資料*/
 UnitId						int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 UnitNo						nvarchar(10)		NOT NULL,						--單元代碼
 UnitName					nvarchar(100)		NOT NULL,						--單元名稱
 UnitDesc					nvarchar(100)		NOT NULL,						--單元描述
 CheckStatus				nvarchar(2)			NOT NULL,						--是否需檢核 >> BAS.[Status]  Y:是 N:否
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ProdUnit PRIMARY KEY (UnitId),
 CONSTRAINT AK1_ProdUnit UNIQUE (CompanyId, UnitNo),
 CONSTRAINT AK2_ProdUnit UNIQUE (CompanyId, UnitName),
 CONSTRAINT FK1_ProdUnit FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)

CREATE TABLE MES.ProcessParameter( /*製程參數資料 Master*/
 ParameterId				int					identity,
 ProcessId					int					NOT NULL,						--所屬製程
 ModeId						int					NOT NULL,						--所屬生產模式
 DepartmentId				int						NULL,						--所屬部門(未關聯)
 ProcessCheckStatus			nvarchar(2)			NOT NULL,						--是否支援工程檢 
 PreCollectionStatus		nvarchar(2)			NOT NULL,						--是否收集開工前資訊
 PostCollectionStatus		nvarchar(2)			NOT NULL,						--是否收集完工後資訊
 PassingMode				nvarchar(2)			NOT NULL,						--過站模式 > A：自動開工完工，S：手動開工自動完工，M：手動開工完工
 NgToBarcode				nvarchar(2)			NOT NULL,						--NG時需刷取NG條碼
 ProcessCheckType			nvarchar(2)			NOT NULL,						--工程檢頻率 1:全檢 2:抽檢
 PackageFlag				nvarchar(2)			NOT NULL	DEFAULT 'N',		--是否支援包裝功能
 ConsumeFlag				nvarchar(2)			NOT NULL	DEFAULT 'N',		--是否有點膠 
 DelayTime					int					NOT NULL	DEFAULT 0,			--製程允許開工完工間隔時間(單位:秒)
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼	
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ProcessParameter PRIMARY KEY (ParameterId),
 CONSTRAINT AK1_ProcessParameter UNIQUE (ProcessId, ModeId),
 CONSTRAINT FK2_ProcessParameter FOREIGN KEY (ProcessId) REFERENCES MES.Process(ProcessId),
 CONSTRAINT FK3_ProcessParameter FOREIGN KEY (ModeId) REFERENCES MES.ProdMode(ModeId)
)

CREATE TABLE MES.ProcessEventItem( /*製程參數事件 */
 ProcessEventItemId			int					identity,
 ParameterId				int					NOT NULL,						--製程參數
 ProcessEventNo				nvarchar(20)		NOT NULL,						--製程事件編號
 ProcessEventName			nvarchar(20)		NOT NULL,						--製程事件名稱
 ProcessEventType			nvarchar(2)			NOT NULL,						--加工事件類型 SELECT * FROM BAS.[Type] WHERE TypeSchema = 'ProcessEventItem';
 [Status]					nvarchar(2)			NOT NULL	DEFAULT 'A',		--狀態碼 A:啟用;S:停用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ProcessEventItem PRIMARY KEY (ProcessEventItemId),
 CONSTRAINT AK1_ProcessEventItem UNIQUE (ParameterId,ProcessEventNo,ProcessEventName,ProcessEventType),
 CONSTRAINT FK1_ProcessEventItem FOREIGN KEY (ParameterId) REFERENCES MES.ProcessParameter(ParameterId)
)

CREATE TABLE MES.ProcessMachine( /*製程機台資料*/
 ParameterId				int					NOT NULL,						--所屬製程參數
 MachineId					int					NOT NULL,						--所屬機台 
 BatchStatus				nvarchar(2)			NOT NULL,						--是否支援批量生產<綁定治具>
 ToolCount					int					NOT NULL	DEFAULT 1,			--製程機台所綁工具上限數
 KeyenceFlag				nvarchar(2)			NOT NULL	DEFAULT 'N',		--是否支援Keyence批量過站
 KeyenceId					int						NULL,						--Keyence設備ID
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT AK1_ProcessMachine UNIQUE (ParameterId, MachineId),
 CONSTRAINT FK1_ProcessMachine FOREIGN KEY (MachineId) REFERENCES MES.Machine(MachineId),
 CONSTRAINT FK2_ProcessMachine FOREIGN KEY (ParameterId) REFERENCES MES.ProcessParameter(ParameterId),
 CONSTRAINT FK3_ProcessMachine FOREIGN KEY (KeyenceId) REFERENCES MES.Keyence(KeyenceId)
)

CREATE TABLE MES.ProcessProductionUnit( /*製程生產單元*//*設定製程所需單元資料*/
 ParameterId				int					NOT NULL,						--所屬製程參數
 UnitId						int					NOT NULL,						--所屬生產單元
 SortNumber					int					NOT NULL,						--排序號碼
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT AK1_ProcessProductionUnit UNIQUE (ParameterId, UnitId),
 CONSTRAINT FK1_ProcessProductionUnit FOREIGN KEY (ParameterId) REFERENCES MES.ProcessParameter(ParameterId),
 CONSTRAINT FK2_ProcessProductionUnit FOREIGN KEY (UnitId) REFERENCES MES.ProdUnit(UnitId)
)

CREATE TABLE MES.Warehouse( /*庫房基本資料*/
 WarehouseId				int					identity,						--流水碼
 ShopId						int					NOT NULL,						--所屬車間
 WarehouseNo				nvarchar(10)		NOT NULL, 						--庫房代碼
 WarehouseName				nvarchar(100)		NOT NULL,						--庫房名稱
 WarehouseDesc				nvarchar(100)		NOT NULL,						--庫房描述
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼	
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_PK_Warehouse PRIMARY KEY (WarehouseId),
 CONSTRAINT AK1_Warehouse UNIQUE (WarehouseNo),
 CONSTRAINT FK1_Warehouse FOREIGN KEY (ShopId) REFERENCES MES.WorkShop(ShopId)
)

CREATE TABLE MES.WarehouseLocation( /*庫房儲位基本資料*/
 LocationId					int					identity,						--流水碼
 WarehouseId				int					NOT NULL,						--所屬庫房
 LocationNo					nvarchar(10)		NOT NULL, 						--儲位代碼
 LocationName				nvarchar(100)		NOT NULL,						--儲位名稱
 LocationDesc				nvarchar(100)		NOT NULL,						--儲位描述
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼	
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_WarehouseLocation PRIMARY KEY (LocationId),
 CONSTRAINT AK1_WarehouseLocation UNIQUE (LocationNo),
 CONSTRAINT FK1_WarehouseLocation FOREIGN KEY (WarehouseId) REFERENCES MES.Warehouse(WarehouseId)
)

CREATE TABLE MES.Keyence( /*Keyence基本資料*/
 KeyenceId					int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 KeyenceNo					nvarchar(100)		NOT NULL,						--Keyence型號
 KeyenceDesc				nvarchar(100)			NULL,						--Keyence設備備註
 KeyenceIP					nvarchar(50)		NOT	NULL,						--KeyenceIP
 KeyencePort				nvarchar(50)		NOT NULL,						--KeyencePort
 [Status]					nvarchar(2)			NOT NULL	DEFAULT 'A',		--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Keyence PRIMARY KEY (KeyenceId),
 CONSTRAINT AK1_Keyence UNIQUE (CompanyId, KeyenceNo, KeyenceIP, KeyencePort),
 CONSTRAINT FK1_Keyence FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)

CREATE TABLE MES.Tray( /*托盤資料*/
 TrayId						int					identity,
 CompanyId					int					NOT NULL,						--所屬公司		
 BarcodeNo					nvarchar(60)			NULL,						--條碼
 TrayNo						nvarchar(100)		NOT NULL,						--萃盤編號
 TrayName					nvarchar(100)		NOT NULL,						--萃盤名稱
 TrayCapacity				int					NOT NULL,						--萃盤容量
 UseTimes					int					NOT	NULL	DEFAULT '0',		--使用次數
 LaserMarkingStatus			nvarchar(2)			NOT NULL	DEFAULT 'N',		--是否可以雷刻並刷取
 Remark						nvarchar(255)			NULL,						--備註
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Tray PRIMARY KEY (TrayId),
 CONSTRAINT AK1_Tray UNIQUE (CompanyId, TrayNo),
 CONSTRAINT FK1_Tray FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_Tray FOREIGN KEY (BarcodeNo) REFERENCES MES.BarcodePrint(BarcodeNo)
)

CREATE TABLE MES.TrayTemp( /*托盤(暫存)資料*/
 TrayTempId					int					identity,
 CompanyId					int					NOT NULL,						--所屬公司	
 TrayTempNo					nvarchar(100)		NOT NULL,						--萃盤編號
 TrayTempName				nvarchar(100)		NOT NULL,						--萃盤名稱
 TrayTempCapacity			int					NOT NULL,						--萃盤容量 
 Remark						nvarchar(255)			NULL,						--備註
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_TrayTemp PRIMARY KEY (TrayTempId),
 CONSTRAINT AK1_TrayTemp UNIQUE (CompanyId, TrayTempNo),
 CONSTRAINT FK1_TrayTemp FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)

CREATE TABLE MES.TrayBarcodeLog( /*托盤曾經綁定條碼資料*/
 TrayBarcodeLogId			int					identity,
 TrayId						int					NOT NULL,
 BarcodeNo					nvarchar(60)			NULL,						--條碼
 BindDate					datetime			    NULL,						--綁定日期
 ReMoveBindDate			datetime			    NULL,							--解除綁定日期
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_TrayBarcodeLog PRIMARY KEY (TrayBarcodeLogId),
 CONSTRAINT FK1_TrayBarcodeLog FOREIGN KEY (TrayId) REFERENCES MES.Tray(TrayId),
 CONSTRAINT FK2_TrayBarcodeLog FOREIGN KEY (BarcodeNo) REFERENCES MES.BarcodePrint(BarcodeNo)
)

CREATE TABLE MES.Vehicle( /*載盤資料檔*/
 VehicleId					int					identity,
 CompanyId					int					NOT NULL,						--公司別ID
 VehicleNo					nvarchar(255)		NOT	NULL,						--載盤條碼
 Remark						nvarchar(255)			NULL,						--備註
 [Status]					nvarchar(2)			NOT NULL	DEFAULT 'S',		--狀態碼 A:啟用 S:未啟用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Vehicle PRIMARY KEY (VehicleId),
 CONSTRAINT AK1_Vehicle UNIQUE (CompanyId, VehicleNo),
 CONSTRAINT FK1_Vehicle FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)

CREATE TABLE MES.VehicleDetail( /*載盤詳細資料檔*/
 VdId						int					identity,
 VehicleId					int					NOT NULL,						--載盤主檔ID
 MoProcessId				int					NOT	NULL,						--製令製程ID
 MtlItemId					int					NOT	NULL,						--物料品號ID
 BarcodeNo					nvarchar(60)			NULL,						--條碼NO
 LotNumberId				int						NULL,						--批號ID
 ExpirationDate				datetime				NULL,						--失效日期
 [Status]					nvarchar(2)			NOT NULL	DEFAULT 'A',		--狀態碼 A:綁定中 S:已解綁
 Remark						nvarchar(255)			NULL,						--備註
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_VehicleDetail PRIMARY KEY (VdId),
 CONSTRAINT FK1_VehicleDetail FOREIGN KEY (VehicleId) REFERENCES MES.Vehicle(VehicleId),
 CONSTRAINT FK2_VehicleDetail FOREIGN KEY (MoProcessId) REFERENCES MES.MoProcess(MoProcessId),
 CONSTRAINT FK3_VehicleDetail FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId),
 CONSTRAINT FK4_VehicleDetail FOREIGN KEY (BarcodeNo) REFERENCES MES.Barcode(BarcodeNo),
 CONSTRAINT FK5_VehicleDetail FOREIGN KEY (LotNumberId) REFERENCES SCM.LotNumber(LotNumberId),
)

CREATE TABLE MES.LabelData( /*標籤檔資料*/
 LabelDataId				int					identity,
 LabelDataNo				nvarchar(50)		NOT NULL,						--標籤檔代號
 LabelDataName				nvarchar(200)		NOT NULL,						--標籤檔名稱
 LabelDataPath				nvarchar(200)		NOT NULL,						--標籤檔路徑
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_LabelData PRIMARY KEY (LabelDataId),
 CONSTRAINT AK1_LabelData UNIQUE (LabelDataNo)
) /*標籤檔資料*/

CREATE TABLE MES.LabelPrintMachine( /*標籤機資料*/
 LabelPrintId				int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 LabelPrintNo				nvarchar(50)		NOT NULL,						--標籤機代號
 LabelPrintName				nvarchar(200)		NOT NULL,						--標籤機名稱
 LabelPrintIp				nvarchar(20)		NOT NULL,						--標籤機IP位置
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_LabelPrintMachine PRIMARY KEY (LabelPrintId),
 CONSTRAINT AK1_LabelPrintMachine UNIQUE (LabelPrintNo)
) /*標籤機資料*/

CREATE TABLE MES.LabelPrintRecord( /*標籤檔所屬標籤機資料*/
 LabelPrintRecordId			int					identity,
 LabelPrintId				nvarchar(50)		NOT NULL,						--標籤機代號
 LabelDataId				nvarchar(200)		NOT NULL,						--標籤檔代號
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_LabelPrintRecord PRIMARY KEY (LabelPrintId),
 CONSTRAINT FK1_LabelPrintRecord FOREIGN KEY (LabelPrintId) REFERENCES MES.LabelPrintMachine(LabelPrintId),
 CONSTRAINT FK2_LabelPrintRecord FOREIGN KEY (LabelDataId) REFERENCES MES.LabelData(LabelDataId),
) /*標籤機資料*/

CREATE TABLE MES.SkyMarsMachine( /*編程上拋機台*/
 SkyMarsMachineId			int					identity,
 MachineId					int					NOT NULL,						--所屬機台
 SkyMarsServerIp			nvarchar(200)		NOT NULL,						--SkyMars Server IP
 SkyMarsServerIpPort		nvarchar(20)		NOT NULL,						--SkyMars Server IP Port
 UploadType					int					NOT NULL,						--上拋類型  //1:記憶體 ; 2:DataServer
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_SkyMarsMachine PRIMARY KEY (SkyMarsMachineId)
) /*標籤機資料*/

CREATE TABLE MES.LaserEngravedUser( /*雷射刻印人員*/
 LaserEngravedUserId		int					identity,
 UserId						int					NOT NULL,						--使用者ID
 LoginNo					nvarchar(200)		NOT NULL,						--帳號
 PassWord					nvarchar(20)		NOT NULL,						--密碼
 RoleType					nvarchar(2)			NOT NULL,						--角色  //1:作業員 ; 2:工程師
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_LaserEngravedUser PRIMARY KEY (LaserEngravedUserId),
 CONSTRAINT FK1_LaserEngravedUser FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId)
) /*雷射刻印人員*/

END /*基本資料*/

BEGIN /*途程管理*/
CREATE TABLE MES.Routing( /* 途程Master */
 RoutingId					int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 ModeId						int					NOT NULL,						--生產模式
 RoutingType				nvarchar(1)			NOT NULL,						--途程類別(1:加工, 2:返修)
 RoutingName				nvarchar(100)		NOT NULL,						--途程名稱
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 RoutingConfirm				nvarchar(2)			NOT NULL,						--是否確認(Y:確認, N:未確認)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Routing PRIMARY KEY (RoutingId),
 CONSTRAINT AK1_Routing UNIQUE (CompanyId, ModeId, RoutingName),
 CONSTRAINT FK1_Routing FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_Routing FOREIGN KEY (ModeId) REFERENCES MES.ProdMode(ModeId)
)

CREATE TABLE MES.RoutingProcess( /* 途程製程 */
 RoutingProcessId			int					identity,						
 RoutingId					int					NOT NULL,						--途程
 ProcessId					int					NOT NULL,						--製程ID
 SortNumber					int					NOT NULL,						--製程順序
 ProcessAlias				nvarchar(32)		NOT NULL,						--製程別名
 DisplayStatus				nvarchar(1)			NOT NULL,						--是否顯示在流程卡 顯示 不顯示
 NecessityStatus			nvarchar(1)			NOT NULL,						--必要過站(Y/N)
 ProcessCheckStatus			nvarchar(2)			NOT NULL,						--是否支援工程檢
 ProcessCheckType			nvarchar(2)				NULL,						--工程檢類型 1:全檢 2:抽檢
 PackageFlag				nvarchar(2)				NULL	DEFAULT 'N',		--是否啟用包裝條碼
 ConsumeFlag				nvarchar(2)			NOT NULL	DEFAULT 'N',		--是否有點膠 
 [Status]					nvarchar(2)			NOT NULL,						--狀態 A:啟用 S:停用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_RoutingProcess PRIMARY KEY (RoutingProcessId),
 CONSTRAINT AK1_RoutingProcess UNIQUE (RoutingId, SortNumber),
 CONSTRAINT FK1_RoutingProcess FOREIGN KEY (RoutingId) REFERENCES MES.Routing(RoutingId),
 CONSTRAINT FK2_RoutingProcess FOREIGN KEY (ProcessId) REFERENCES MES.Process(ProcessId)
)

CREATE TABLE MES.RoutingProcessItem( /* 途程製程需綁定項目(刻字) */
 RoutingProcessItemId		int					identity,						
 RoutingProcessId			int					NOT NULL,						
 ItemNo						nvarchar(30)		NOT NULL,						--製程需綁定項目
 ChkUnique					nvarchar(1)			NOT NULL,						--檢核重覆值 >> BAS.[Status] Y:需檢核重覆值；N:不需檢核重覆值
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_RoutingProcessItem PRIMARY KEY (RoutingProcessItemId),
 CONSTRAINT AK1_RoutingProcessItem UNIQUE (RoutingProcessId, ItemNo),
 CONSTRAINT FK1_RoutingProcessItem FOREIGN KEY (RoutingProcessId) REFERENCES MES.RoutingProcess(RoutingProcessId)
)

CREATE TABLE MES.RoutingItem( /* 途程品號(列印流程卡用) */
 RoutingItemId				int					identity,
 RoutingId					int					NOT NULL,
 ControlId					int						NULL,						--研發設計圖版控ID(模具要設計圖時使用)
 MtlItemId					int					NOT NULL,						--品號ID
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 RoutingItemConfirm			nvarchar(2)			NOT NULL,						--是否確認(Y:確認, N:未確認)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_RoutingItem PRIMARY KEY (RoutingItemId),
 CONSTRAINT FK1_RoutingItem FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId),
 CONSTRAINT FK2_RoutingItem FOREIGN KEY (ControlId) REFERENCES PDM.RdDesignControl(ControlId)
)

CREATE TABLE MES.RoutingProcessTemplate( /* 製程加工細項樣版(流程卡用) */
 ProcessTemplateId			int					identity,
 ProcessId					int					NOT NULL,						--製程ID
 ProcessTemplateDesc		nvarchar(200)			NULL,						--加工細項說明
 Remark						nvarchar(200)			NULL,						--加工細項備註
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_RoutingProcessTemplate PRIMARY KEY (ProcessTemplateId),
 CONSTRAINT FK1_RoutingProcessTemplate FOREIGN KEY (ProcessId) REFERENCES MES.Process(ProcessId)
)

CREATE TABLE MES.RoutingItemProcess( /* 製程加工細項(流程卡用), 用ProcessNo找RoutingProcessTemplate */
 ItemProcessId				int					identity,
 RoutingItemId				int					NOT NULL,
 RoutingProcessId			int					NOT NULL,
 RoutingItemProcessDesc		nvarchar(200)			NULL,						--加工細項說明
 CycleTime					int					NOT	NULL	DEFAULT 0,			--標準工時
 MoveTime					int					NOT	NULL	DEFAULT 0,			--當站完工至下站移轉時間
 ProcessingTime				int					NOT NULL	DEFAULT 0,			--標準工時上下公差秒數
 WatingTime					int					NOT NULL	DEFAULT 0,			--當站最小等待時間
 Remark						nvarchar(200)			NULL,						--加工細項備註
 DisplayStatus				nvarchar(2)			NOT NULL,						--是否顯示於流程卡
 AttrSetting				nvarchar(200)			NULL,						--設定刻字規則
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員-
 CONSTRAINT PK_RoutingItemProcess PRIMARY KEY (ItemProcessId),
 CONSTRAINT AK1_RoutingItemProcess UNIQUE (RoutingItemId, RoutingProcessId),
 CONSTRAINT FK1_RoutingItemProcess FOREIGN KEY (RoutingItemId) REFERENCES MES.RoutingItem(RoutingItemId),
 CONSTRAINT FK2_RoutingItemProcess FOREIGN KEY (RoutingProcessId) REFERENCES MES.RoutingProcess(RoutingProcessId)
)

CREATE TABLE MES.RoutingPlanning( /* 途程製程機台排程 */
 PlanningId					int					identity,
 StructureId				int					NOT NULL,						--PC工程平台需求展開ID
 RoutingItemId				int					NOT NULL,						--途程品號ID
 RoutingProcessId			int					NOT NULL,						--途程製程ID
 Quantity					float				NOT NULL,						--排程數量
 MachineId					int					NOT NULL,						--機台ID
 StartDate					datetime			NOT NULL,						--排程開始時間
 EndDate					datetime			NOT NULL,						--排程結束時間
 SupplierId					int					NOT NULL,						--供應商ID
 ProcessCode				nvarchar(100)			NULL,						--加工代碼(委外單使用)
 Sort						int						NULL,						--排程時順序
 Remark						nvarchar(255)			NULL,						--備註
 [Status]					nvarchar(2)			NOT NULL	DEFAULT 'N',		--確認碼 Y/N
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_RoutingPlanning PRIMARY KEY (PlanningId),
 CONSTRAINT FK1_RoutingPlanning FOREIGN KEY (RoutingProcessId) REFERENCES MES.RoutingProcess(RoutingProcessId),
 CONSTRAINT FK2_RoutingPlanning FOREIGN KEY (MachineId) REFERENCES MES.Machine(MachineId)
)

CREATE TABLE MES.RoutingMaterial( /* 途程物料上料相關設定 */
 RoutingMaterialId			int					identity,
 RoutingItemId				int					NOT NULL,
 MtlItemId					int					NOT NULL,						--品號ID
 BindType					nvarchar(1)			NOT NULL,						--上料類型
 RoutingProcessId			int						NULL,						--上料製程
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_RoutingMaterial PRIMARY KEY (RoutingMaterialId),
 CONSTRAINT FK1_RoutingMaterial FOREIGN KEY (RoutingItemId) REFERENCES MES.RoutingItem(RoutingItemId),
 CONSTRAINT FK2_RoutingMaterial FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId),
 CONSTRAINT FK3_RoutingMaterial FOREIGN KEY (RoutingProcessId) REFERENCES MES.RoutingProcess(RoutingProcessId),
)

CREATE TABLE MES.MoProcessChange( /* 製令途程變更單 */
 MpcId					    int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 MpcNo					    nvarchar(15)		NOT NULL,						--製令途程變更單單號
 MoId   					int					NOT NULL,						--所屬製令
 MpcUserId					int					NOT NULL,						--變更者
 DocDate					datetime    		NOT NULL,                       --單據日期
 Remark					    nvarchar(255)		    NULL,                       --變更原因/備註
 SortChange 				nvarchar(2)		    NOT NULL 	DEFAULT 'K',		--順序改變
 [Status] 					nvarchar(2)		    NOT NULL,                       --狀態
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MoProcessChange PRIMARY KEY (MpcId),
 CONSTRAINT AK1_MoProcessChange UNIQUE (CompanyId, MpcNo),
 CONSTRAINT FK1_MoProcessChange FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_MoProcessChange FOREIGN KEY (MoId) REFERENCES MES.ManufactureOrder(MoId),
 CONSTRAINT FK3_MoProcessChange FOREIGN KEY (MpcUserId) REFERENCES BAS.[User](UserId)
)

CREATE TABLE MES.MpcRoutingProcess( /* 變更途程製程 */
 MpcRoutingProcessId		int					identity,			
 MpcId   					int					NOT NULL,						--所屬途程變更單
 MoRoutingId				int						NULL,						--MES製令途程ID(可能中途加製程，因此可NULL)
 RoutingProcessId			int						NULL,						--途程製程ID
 MoProcessId				int					NOT NULL,						--製令上製程ID
 ProcessId					int						NULL,						--製程ID
 SortNumber					int					NOT NULL,						--製程順序
 ProcessAlias				nvarchar(32)		NOT NULL,						--製程別名
 RoutingItemProcessDesc		nvarchar(200)			NULL,						--加工細項說明
 Remark						nvarchar(200)			NULL,						--加工細項備註
 DisplayStatus				nvarchar(1)			NOT NULL,						--是否顯示在流程卡 顯示 不顯示
 NecessityStatus			nvarchar(1)			NOT NULL,						--必要過站 >> BAS.[Type] Y:必要過站(DEFAULT)；N:非必要過站 (Y時需在[入庫]時檢核是否已成)
 RandomStatus				nvarchar(1)				NULL,		--必要按順序過站 >> BAS.[Type] Y:按順序(DEFAULT)；N:不按順序 (Y時需檢核前製程是否已完成)
 OspFlag					nvarchar(1)			NOT NULL,						--此製程是否為委外加工站(顯示用)
 ProcessCheckStatus			nvarchar(2)			NOT NULL,						--此製程是否支援工程檢
 ProcessCheckType			nvarchar(2)			NOT NULL,						--工程檢類型 1:全檢 2:抽檢
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MpcRoutingProcess PRIMARY KEY (MpcRoutingProcessId),
 CONSTRAINT AK1_MpcRoutingProcess UNIQUE (MoRoutingId, SortNumber),
 CONSTRAINT FK1_MpcRoutingProcess FOREIGN KEY (MpcId) REFERENCES MES.MoProcessChange(MpcId),
 CONSTRAINT FK2_MpcRoutingProcess FOREIGN KEY (MoRoutingId) REFERENCES MES.MoRouting(MoRoutingId),
 CONSTRAINT FK3_MpcRoutingProcess FOREIGN KEY (RoutingProcessId) REFERENCES MES.RoutingProcess(RoutingProcessId),
 CONSTRAINT FK4_MpcRoutingProcess FOREIGN KEY (MoProcessId) REFERENCES MES.MoProcess(MoProcessId),
 CONSTRAINT FK5_MpcRoutingProcess FOREIGN KEY (ProcessId) REFERENCES MES.Process(ProcessId)
)

CREATE TABLE MES.MpcBarcode( /* 變更途程所綁定條碼 */
 MpcBarcodeId				int					identity,
 MpcId				        int					NOT NULL,						--所屬途程變更單
 BarcodeNo					nvarchar(60)		NOT NULL,						--條碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MpcBarcode PRIMARY KEY (MpcBarcodeId),
 CONSTRAINT AK1_MpcBarcode UNIQUE (MpcId, BarcodeNo),
 CONSTRAINT FK1_MpcBarcode FOREIGN KEY (MpcId) REFERENCES MES.MoProcessChange(MpcId)
)

CREATE TABLE MES.RoutingItemAttribute( /* 途程品號屬性檔(如果RdDesignAttribute存在,直接帶過來,否則由介面直接新增) */
 AttributeId				int					identity,
 RoutingItemId				int					NOT NULL,
 AttributeItemId			int						NULL,						--品號屬性ID
 AttributeValue				nvarchar(200)			NULL,						--屬性值
 UpperLimit					nvarchar(20)			NULL,						--上限值
 LowerLimit					nvarchar(20)			NULL,						--下限值
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_RoutingItemAttribute PRIMARY KEY (AttributeId),
 CONSTRAINT AK1_RoutingItemAttribute UNIQUE (RoutingItemId, AttributeItemId),
 CONSTRAINT FK1_RoutingItemAttribute FOREIGN KEY (RoutingItemId) REFERENCES MES.RoutingItem(RoutingItemId),
 CONSTRAINT FK2_RoutingItemAttribute FOREIGN KEY (AttributeItemId) REFERENCES PDM.RdAttributeItem(AttributeItemId)
)
END /*途程管理*/

BEGIN /*生產管理*/
CREATE TABLE MES.ProdDemand( /*生產需求單頭*/
 DemandId					int					identity,
 DepartmentId				int					NOT NULL,						--委託部門
 DemandNo					nvarchar(50)		NOT NULL,						--委託單號
 UserId						int					NOT NULL,						--委託人員
 DemandDate					datetime			NOT NULL,						--委託日期
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ProdDemand PRIMARY KEY (DemandId),
 CONSTRAINT FK1_ProdDemand FOREIGN KEY (DepartmentId) REFERENCES BAS.Department(DepartmentId),
 CONSTRAINT FK2_ProdDemand FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId)
)

CREATE TABLE MES.ProdDemandDetail( /*生產需求單身*/
 DemandDetailId				int					identity,
 DemandId					int					NOT NULL,						--所屬單頭
 MtlItemId					int					NOT NULL,						--委託品號
 DemandQty					int					NOT NULL,						--委託數量
 ExceptedFinshDate			datetime			NOT NULL,						--預計交期
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ProdDemandDetail PRIMARY KEY (DemandDetailId),
 CONSTRAINT FK1_ProdDemandDetail FOREIGN KEY (DemandId) REFERENCES MES.ProdDemand(DemandId),
 CONSTRAINT FK2_ProdDemandDetail FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId)
)

CREATE TABLE MES.WipOrder( /*ERP製令單頭*/ /*ERP MOCTA*/
 WoId						int					identity,
 CompanyId					int					NOT NULL,						--TA019 所屬公司
 UserId						int					    NULL,						--TA055 生管/採購人員
 WoErpPrefix				nvarchar(4)			NOT NULL,						--TA001 製令單別 >> 參照 ERP各系統單據設定檔(CMSMQ)
 WoErpNo					nvarchar(11)		NOT NULL,						--TA002 製令單號
 [Version]					nvarchar(4)			NOT NULL,						--TA062 版次
 DocDate					datetime			NOT NULL,						--TA003 開單日期
 MtlItemId					int					NOT NULL,						--TA006 產品品號
 InventoryId				int					NOT NULL,						--TA020 入庫庫別
 UomId						int					NOT NULL,						--TA007 單位(Unit of Measure) >> 參照 ERP品號換算單位檔(INVMD)
 PlanQty					int					NOT NULL,						--TA015 預計產量
 RequisitionSetQty			int					NOT NULL,						--TA016 已領套數
 StockInQty					int					NOT NULL,						--TA017 已生產量
 ScrapQty					int					NOT NULL,						--TA018 報廢數量
 ExpectedStart				datetime				NULL,						--TA009 預計開工
 ExpectedEnd				datetime				NULL,						--TA010 預計完工
 ActualStart				datetime				NULL,						--TA012 實際開工
 ActualEnd					datetime				NULL,						--TA014 實際完工
 BomDate					datetime			NOT NULL,						--TA004 BOM日期
 BomVersion					nvarchar(4)			NOT NULL,						--TA005 BOM版次
 UrgentMtl					nvarchar(1)			NOT NULL,						--TA044 急料 >> BAS.[Type] N[DEF]；Y
 Property					nvarchar(1)			NOT NULL,						--TA030 性質 >> BAS.[Type] 1：廠內製令；2：託外製令
 SoDetailId					int						NULL,						--TA026/27/28 參考訂單
 Project					nvarchar(20)			NULL,						--TA551	專案代號 >> 參考 ERP專案代號檔(CMSNB)
 ProductionLine				nvarchar(10)			NULL,						--TA021	線別 >> 參考 ERP專案代號檔(CMSNB)
 Processor					nvarchar(10)			NULL,						--TA032	加工廠商 >> 參考 ERP專案代號檔(CMSNB)
 CurrencyCode				nvarchar(4)				NULL,						--TA042 幣別
 Exchange					float					NULL	DEFAULT 0.0000,		--TA043 匯率
 TaxCode					nvarchar(3)				NULL,						--TA070 稅別碼 (要找ERP資料來源)
 TaxType					nvarchar(1)				NULL,						--TA056 課稅別 1.應稅內含、2.應稅外加、3.零稅率、4.免稅、9.不計稅
 TaxRate					float					NULL,						--TA057 營業稅率 營業稅率(%)[DEF:9.9999]&&89-06-01ADD//950808 S00-9508031 N5.4->N8.5
 PricingTerm				nvarchar(40)			NULL,						--TA058 價格條件
 PaymentTerm				nvarchar(6)				NULL,						--TA059 付款條件代號
 RecipientAddress			nvarchar(255)			NULL,						--TA060 送貨地址
 WoRemark					nvarchar(255)			NULL,						--TA029 單頭備註
 ConfirmStatus				nvarchar(1)			NOT NULL,						--TA013 確認碼 >> BAS.[Status] N：未確認；V：作廢；Y：已確認
 ConfirmUserId				int						NULL,						--TA041 確認者
 ConfirmDate				datetime				NULL,						--TA040 確認日期
 WoStatus					nvarchar(1)			NOT NULL,						--TA011 製令狀態碼 >> BAS.[Status] 1：未生產；2：已發料；3：生產中；4：已完工；5：指定完工
 TransferStatus				nvarchar(2)			NOT NULL,						--拋轉ERP狀態 >> BAS.[Status]
 TransferDate				datetime				NULL,						--拋轉ERP時間
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_WipOrder PRIMARY KEY (WoId),
 --CONSTRAINT AK1_WipOrder UNIQUE (CompanyId, WoErpPrefix, WoErpNo),
 CONSTRAINT FK1_WipOrder FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_WipOrder FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK3_WipOrder FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId),
 CONSTRAINT FK4_WipOrder FOREIGN KEY (InventoryId) REFERENCES SCM.Inventory(InventoryId),
 CONSTRAINT FK5_WipOrder FOREIGN KEY (UomId) REFERENCES PDM.UnitOfMeasure(UomId),
 CONSTRAINT FK6_WipOrder FOREIGN KEY (SoDetailId) REFERENCES SCM.SoDetail(SoDetailId)
)

CREATE TABLE MES.WoDetail( /*ERP製令單身*/ /*ERP MOCTB*/
 WoDetailId					int					identity,
 WoId						int					NOT NULL,						--所屬單頭
 MtlItemId					int					NOT NULL,						--TB003 材料品號
 InventoryId				int					NOT NULL,						--TB009 庫別
 UomId						int					NOT NULL,						--TB007 單位(Unit of Measure) >> 參照 ERP品號換算單位檔(INVMD)
 CompositionQuantity		float					NULL,						--組成用量
 Base						float					NULL,						--底數
 LossRate					float					NULL,						--損耗率
 BomSequence				nvarchar(4)				NULL,						--序號
 DemandRequisitionQty		float				NOT NULL,						--TB004 需領用量
 RequisitionQty				float				NOT NULL,						--TB005 已領用量
 Substitute					nvarchar(1)				NULL,						--TB010 取替代件
 SubstituteStatus			nvarchar(1)			NOT	NULL,						--TB010 取替代碼 >> Y:是;N:否
 SubstituteMtlItemId		int					NOT NULL,						--TB023 被取替代品號
 MaterialProperties			nvarchar(4)			    NULL,						--TB011 材料型態
 ExpectedRequisitionDate	datetime			    NULL,						--TB015 預計領料日期
 ActualRequisitionDate		datetime				NULL,						--TB016 實際領料日期
 WipDetailRemark			nvarchar(255)			NULL,						--TB017 單身備註
 ConfirmStatus				nvarchar(1)			NOT NULL,						--TB018 確認碼 >> BAS.[Status] N：未確認；V：作廢；Y：已確認
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_WoDetail PRIMARY KEY (WoDetailId),
 CONSTRAINT FK1_WoDetail FOREIGN KEY (WoId) REFERENCES MES.WipOrder(WoId),
 CONSTRAINT FK2_WoDetail FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId),
 CONSTRAINT FK3_WoDetail FOREIGN KEY (InventoryId) REFERENCES SCM.Inventory(InventoryId),
 CONSTRAINT FK4_WoDetail FOREIGN KEY (UomId) REFERENCES PDM.UnitOfMeasure(UomId),
 CONSTRAINT FK5_WoDetail FOREIGN KEY (SubstituteMtlItemId) REFERENCES PDM.MtlItem(MtlItemId)
)

CREATE TABLE MES.WipOrderLog( /*ERP製令單頭Loge*/ /*ERP MOCTA*/
 WoLogId					int					identity,
 TransactionType			nvarchar(2)			    NULL,						--交易類型 A:新增 U:更新 T:拋轉 RT:反確認後在拋轉 D:刪除 SA:同步新增 SU:同步更新 SD:同步刪除
 ErpDb						nvarchar(50)			NULL,						--ERP資料庫
 ErpNo						nvarchar(10)			NULL,						--ERP公司編號
 TransactionUserId			datetime				NULL,						--交易者
 TransactionDate			datetime				NULL,						--交易時間
 CompanyId					int						NULL,						--TA019 所屬公司
 UserId						int					    NULL,						--TA055 生管/採購人員
 WoErpPrefix				nvarchar(4)				NULL,						--TA001 製令單別 >> 參照 ERP各系統單據設定檔(CMSMQ)
 WoErpNo					nvarchar(11)			NULL,						--TA002 製令單號
 [Version]					nvarchar(4)				NULL,						--TA062 版次
 DocDate					datetime				NULL,						--TA003 開單日期
 MtlItemId					int						NULL,						--TA006 產品品號
 InventoryId				int						NULL,						--TA020 入庫庫別
 UomId						int						NULL,						--TA007 單位(Unit of Measure) >> 參照 ERP品號換算單位檔(INVMD)
 PlanQty					int						NULL,						--TA015 預計產量
 RequisitionSetQty			int						NULL,						--TA016 已領套數
 StockInQty					int						NULL,						--TA017 已生產量
 ScrapQty					int						NULL,						--TA018 報廢數量
 ExpectedStart				datetime				NULL,						--TA009 預計開工
 ExpectedEnd				datetime				NULL,						--TA010 預計完工
 ActualStart				datetime				NULL,						--TA012 實際開工
 ActualEnd					datetime				NULL,						--TA014 實際完工
 BomDate					datetime				NULL,						--TA004 BOM日期
 BomVersion					nvarchar(4)				NULL,						--TA005 BOM版次
 UrgentMtl					nvarchar(1)				NULL,						--TA044 急料 >> BAS.[Type] N[DEF]；Y
 Property					nvarchar(1)				NULL,						--TA030 性質 >> BAS.[Type] 1：廠內製令；2：託外製令
 SoDetailId					int						NULL,						--TA026/27/28 參考訂單
 Project					nvarchar(20)			NULL,						--TA551	專案代號 >> 參考 ERP專案代號檔(CMSNB)
 WoRemark					nvarchar(255)			NULL,						--TA029 單頭備註
 ConfirmStatus				nvarchar(1)				NULL,						--TA013 確認碼 >> BAS.[Status] N：未確認；V：作廢；Y：已確認
 ConfirmUserId				int						NULL,						--TA041 確認者
 ConfirmDate				datetime				NULL,						--TA040 確認日期
 WoStatus					nvarchar(1)				NULL,						--TA011 製令狀態碼 >> BAS.[Status] 1：未生產；2：已發料；3：生產中；4：已完工；5：指定完工
 TransferStatus				nvarchar(2)				NULL,						--拋轉ERP狀態 >> BAS.[Status]
 TransferDate				datetime				NULL,						--拋轉ERP時間
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_WipOrderLog PRIMARY KEY (WoLogId),
)

CREATE TABLE MES.WoDetailLog( /*ERP製令單身Log*/ /*ERP MOCTB*/
 WoDetailLogId				int					identity,
 TransactionType			nvarchar(2)			    NULL,						--交易類型 A0:工單新增 B0:工單新增(返修) A1:單獨新增 A2:套用BOM A3:套用替代料  SA:同步新增 
 ErpDb						nvarchar(50)			NULL,						--ERP資料庫(主要用來紀錄同步)
 ErpNo						nvarchar(10)			NULL,						--ERP公司編號(主要用來紀錄同步)
 TransactionUserId			int						NULL,						--交易者
 TransactionDate			datetime				NULL,						--交易時間
 CompanyId					int						NULL,						--TA019 所屬公司
 WoId						int						NULL,						--工單單頭ID(當ERP單據刪除時,可以與WipOrderLog搭配看)
 WoDetailId					int						NULL,						--工單單身ID(當出多出來的單身可以快速查詢用)
 WoMtlItemId				int						NULL,						--主件品號(主要用來確認同步時,是哪個主件品號)
 WoErpPrefix				nvarchar(4)				NULL,						--TA001 製令單別(主要用來確認同步時,新增到哪個單據)
 WoErpNo					nvarchar(11)			NULL,						--TA002 製令單號(主要用來確認同步時,新增到哪個單據)
 MtlItemId					int						NULL,						--TB003 材料品號
 InventoryId				int						NULL,						--TB009 庫別
 UomId						int						NULL,						--TB007 單位(Unit of Measure) >> 參照 ERP品號換算單位檔(INVMD)
 CompositionQuantity		float					NULL,						--組成用量
 Base						float					NULL,						--底數
 LossRate					float					NULL,						--損耗率
 BomSequence				nvarchar(4)				NULL,						--序號
 DemandRequisitionQty		float					NULL,						--TB004 需領用量
 RequisitionQty				float					NULL,						--TB005 已領用量
 Substitute					nvarchar(1)				NULL,						--TB010 取替代件
 SubstituteStatus			nvarchar(1)				NULL,						--TB010 取替代碼 >> Y:是;N:否
 SubstituteMtlItemId		int						NULL,						--TB023 被取替代品號
 MaterialProperties			nvarchar(4)			    NULL,						--TB011 材料型態
 ExpectedRequisitionDate	datetime			    NULL,						--TB015 預計領料日期
 ActualRequisitionDate		datetime				NULL,						--TB016 實際領料日期
 WipDetailRemark			nvarchar(255)			NULL,						--TB017 單身備註
 ConfirmStatus				nvarchar(1)				NULL,						--TB018 確認碼 >> BAS.[Status] N：未確認；V：作廢；Y：已確認
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_WoDetailLog PRIMARY KEY (WoDetailLogId),
)

CREATE TABLE MES.WoModification( /*ERP製令變更單頭*/ /*ERP MOCTO*/
 WomId						int					identity,
 WoId						int					NOT NULL,						--所屬製令
 [Version]					nvarchar(11)		NOT NULL,						--TO003  變更版次        
 ModiDate					nvarchar(4)			NOT NULL,						--TO004  變更日期  
 ModiReason					nvarchar(255)			NULL,						--TO005  變更原因   
 UserId                     int					    NULL,						--TO057  生管/採購人     
 PlanQty                    int					NOT NULL,						--TO017  預計產量        
 ExpectedStart              datetime				NULL,						--TO012  預計開工        
 ExpectedEnd                datetime				NULL,						--TO013  預計完工        
 BomDate                    datetime			NOT NULL,						--TO007  BOM日期         
 InventoryId                int					NOT NULL,						--TO022  入庫庫別        
 SupplierId                 int						NULL,						--TO033  加工廠商  
 PaymentTerm                nvarchar(6) 			NULL,						--TO061  付款條件        
 ProduceUnitPrice           float				NOT NULL,						--TO070  加工單價   
 ProduceUomId               int						NULL,						--TO025  加工單位        
 Currency                   nvarchar(4)				NULL,						--TO045  幣別      
 ExchangeRate               float					NULL,						--TO046  匯率
 TaxNo                      nvarchar(3)			NOT NULL,						--TO071  稅別碼 
 Taxation                   nvarchar(1)				NULL,						--TO058  課稅別  
 BusinessTaxRate            float				NOT NULL,						--TO059  營業稅率      
 PlanLotNo                  nvarchar(20)			NULL,						--TO034  計畫批號
 ParentWoId                 int						NULL,						--TO026/T0027  母製令
 SoDetailId                 int						NULL,						--TO028/29/30  訂單單號
 Project                    nvarchar(20)			NULL,						--TO532  專案代號
 WoStatus					nvarchar(1)			NOT NULL,						--TO014  製令狀態碼 >> BAS.[Status] 1：未生產；2：已發料；3：生產中；4：已完工；5：指定完工
 WoRemark                   nvarchar(20)			NULL,						--TO032  備註
 DocDate					datetime			NOT NULL,						--TO106  原單據日期
 OriUserId					int					    NULL,						--TO157  原生管/採購人員
 OriPlanQty					int					NOT NULL,						--TO117  原預計產量
 OriExpectedStart			datetime				NULL,						--TO112  原預計開工
 OriExpectedEnd				datetime				NULL,						--TO113  原預計完工
 OriBomDate					datetime			NOT NULL,						--TO107  原BOM日期
 OriInventoryId				int					NOT NULL,						--TO122  原入庫庫別
 OriSupplierId				int						NULL,						--TO133  原加工廠商
 OriPaymentTerm				nvarchar(6) 			NULL,						--TO161  原付款條件
 OriProduceUnitPrice		float				NOT NULL,						--TO124  原加工單價
 OriProduceUomId			int						NULL,						--TO125  原加工單位
 OriCurrency				nvarchar(4)				NULL,						--TO145  原幣別
 OriExchangeRate			float					NULL,						--TO146  原匯率
 OriTaxNo					nvarchar(3)			NOT NULL,						--TO171  原稅別碼
 OriTaxation				nvarchar(1)				NULL,						--TO158  原課稅別
 OriBusinessTaxRate			float				NOT NULL,						--TO159  原營業稅率
 OriPlanLotNo				nvarchar(20)			NULL,						--TO134  原計畫批號
 OriParentWoId				int						NULL,						--TO126/127     母製令
 OriSoDetailId				int						NULL,						--TO128/129/130 原訂單單號
 OriProject					nvarchar(20)			NULL,						--TO558  原專案代號
 OriWoStatus				nvarchar(1)			NOT NULL,						--TO114  原製令狀態碼 >> BAS.[Status] 1：未生產；2：已發料；3：生產中；4：已完工；5：指定完工
 OriWoRemark				nvarchar(20)			NULL,						--TO131  原備註
 ConfirmStatus				nvarchar(1)			NOT NULL,						--TO041 確認碼 >> BAS.[Status] N：未確認；V：作廢；Y：已確認
 ConfirmUserId				int						NULL,						--TO044 確認者
 ConfirmDate				datetime				NULL,						--TO040 確認日期
 TransferStatus				nvarchar(2)			NOT NULL,						--拋轉ERP狀態 >> BAS.[Status]
 TransferDate				datetime				NULL,						--拋轉ERP時間
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_WoModification PRIMARY KEY (WomId),
 --CONSTRAINT AK1_WoModification UNIQUE (CompanyId, [Version]),
 CONSTRAINT FK1_WoModification FOREIGN KEY (WoId) REFERENCES MES.WipOrder(WoId),
 CONSTRAINT FK2_WoModification FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK3_WoModification FOREIGN KEY (InventoryId) REFERENCES SCM.Inventory(InventoryId),
 CONSTRAINT FK4_WoModification FOREIGN KEY (OriUserId) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK5_WoModification FOREIGN KEY (OriInventoryId) REFERENCES SCM.Inventory(InventoryId),
)

CREATE TABLE MES.WomDetail( /*ERP製令變更單身*/ /*ERP MOCTP*/
 WomDetail					int					identity,
 WoId						int					NOT NULL,						--所屬單身
 WoDetailId					int					NOT NULL,						--所屬單頭
 InventoryId				int					NOT NULL,						--TP010 庫別
 DemandRequisitionQty		float				NOT NULL,						--TP005 需領用量
 Substitute					nvarchar(1)				NULL,						--TP011 取替代件
 SubstituteMtlItemId		int					NOT NULL,						--TP024 被取替代品號
 WipDetailRemark			nvarchar(255)			NULL,						--TP018 單身備註
 ExpectedRequisitionDate	datetime			    NULL,						--TP016 預計領料日期
 ActualRequisitionDate		datetime				NULL,						--TP017 實際領料日期
 OriInventoryId				int					NOT NULL,						--TP110 原庫別
 OriDemandRequisitionQty	float				NOT NULL,						--TP105 原需領用量
 OriSubstitute				nvarchar(1)				NULL,						--TP111 原取替代件
 OriExpectedRequisitionDate	datetime			    NULL,						--TP116 原預計領料日期
 OriActualRequisitionDate	datetime				NULL,						--TP117 原實際領料日期
 OriWipDetailRemark			nvarchar(255)			NULL,						--TP118 原單身備註
 ConfirmStatus				nvarchar(1)			NOT NULL,						--TP019 確認碼 >> BAS.[Status] N：未確認；V：作廢；Y：已確認
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_WomDetail PRIMARY KEY (WomDetail),
 CONSTRAINT FK1_WomDetail FOREIGN KEY (WoId) REFERENCES MES.WipOrder(WoId),
 CONSTRAINT FK2_WomDetail FOREIGN KEY (WoDetailId) REFERENCES MES.WoDetail(WoDetailId),
 CONSTRAINT FK3_WomDetail FOREIGN KEY (InventoryId) REFERENCES SCM.Inventory(InventoryId),
 CONSTRAINT FK4_WomDetail FOREIGN KEY (SubstituteMtlItemId) REFERENCES PDM.MtlItem(MtlItemId),
 CONSTRAINT FK5_WomDetail FOREIGN KEY (OriInventoryId) REFERENCES SCM.Inventory(InventoryId),
 CONSTRAINT FK6_WomDetail FOREIGN KEY (SubstituteMtlItemId) REFERENCES PDM.MtlItem(MtlItemId)
)

CREATE TABLE MES.ManufactureOrder( /*MES製令*/ 
 MoId						int					identity,
 WoId						int					NOT NULL, 
 WoSeq						int					NOT NULL,						--MES製令序號
 Quantity					int					NOT NULL,						--MES製令預計產量
 InputQty					int					NOT NULL,						--MES製令已投入數量
 RequisitionSetQty			int					NOT NULL	DEFAULT 0,			--MES製令已領套數
 CompleteQty				int					NOT NULL,						--完工數量
 ScrapQty					int					NOT NULL,						--報廢數量
 ModeId						int						NULL,						--生產模式ID
 ProjectNo					nvarchar(100)			NULL,						--專案代碼
 DeliveryProcess			nvarchar(10)			NULL,						--出貨工程
 Remark						nvarchar(255)			NULL,						--備註
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼 >> BAS.[Status] N：新製令未開工；Ａ：加工中；Ｆ：完工
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ManufactureOrder PRIMARY KEY (MoId),
 CONSTRAINT AK1_ManufactureOrder UNIQUE (WoId, WoSeq), 
 CONSTRAINT FK2_ManufactureOrder FOREIGN KEY (WoId) REFERENCES MES.WipOrder(WoId),
 CONSTRAINT FK3_ManufactureOrder FOREIGN KEY (ModeId) REFERENCES MES.ProdMode(ModeId),
)

CREATE TABLE MES.MoItemPart( /*製令刻字資料*/
 MoItemPartId   			int					identity,
 MoItemPartPrefix			nvarchar(100)		NOT NULL,						--刻自前綴
 MoItemPartSortNumber       int					NOT NULL, 						--刻字序號
 MoItemPartNo               nvarchar(100)		NOT NULL,						--刻字號 
 MoId						int         		NOT NULL,						--製令ID
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MoItemPart PRIMARY KEY (MoItemPartId),
 CONSTRAINT AK1_MoItemPart UNIQUE (MoId, MoItemPartNo),
 CONSTRAINT FK1_MoItemPart FOREIGN KEY (MoId) REFERENCES MES.ManufactureOrder(MoId)
)

CREATE TABLE MES.MoRouting( /*MES製令途程*/
 MoRoutingId				int					identity,
 MoId						int					NOT NULL,						--所屬MES製令ID
 RoutingItemId				int					NOT NULL,						--途程品號ID
 SortNumber					int					NOT NULL,						--製令途程順序
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MoRouting PRIMARY KEY (MoRoutingId),
 CONSTRAINT AK1_MoRouting UNIQUE (MoId, RoutingItemId),
 CONSTRAINT FK1_MoRouting FOREIGN KEY (MoId) REFERENCES MES.ManufactureOrder(MoId),
 CONSTRAINT FK2_MoRouting FOREIGN KEY (RoutingItemId) REFERENCES MES.RoutingItem(RoutingItemId)
)

CREATE TABLE MES.MoProcess( /*MES製令Proces*/
 MoProcessId				int					identity,
 MoId						int					NOT	NULL,						--MES製令ID
 MoRoutingId				int						NULL,						--MES製令途程ID(可能中途加製程，因此可NULL)
 RoutingProcessId			int						NULL,						--途程製程ID
 ProcessId					int					NOT NULL,						--製程ID
 SortNumber					int					NOT NULL,						--製程順序
 ProcessAlias				nvarchar(32)		NOT NULL,						--製程別名
 RoutingItemProcessDesc		nvarchar(200)			NULL,						--加工細項說明
 Remark						nvarchar(200)			NULL,						--加工細項備註
 DisplayStatus				nvarchar(1)			NOT NULL,						--是否顯示在流程卡 顯示 不顯示
 NecessityStatus			nvarchar(1)			NOT NULL,						--必要過站 >> BAS.[Type] Y:必要過站(DEFAULT)；N:非必要過站 (Y時需在[入庫]時檢核是否已成)
 OspFlag					nvarchar(1)			NOT NULL,						--此製程是否為委外加工站(顯示用)
 ProcessCheckStatus			nvarchar(2)			NOT NULL,						--此製程是否支援工程檢
 ProcessCheckType			nvarchar(2)				NULL,						--工程檢類型 1:全檢 2:抽檢
 ConsumeFlag				nvarchar(2)			NOT NULL	DEFAULT 'N',		--是否有點膠 
 TotalInputQty				int					NOT NULL	DEFAULT 0,			--製程累績投入數
 TotalPassQty				int					NOT NULL	DEFAULT 0,			--製程累績良品數
 TotalNgQty					int					NOT NULL	DEFAULT 0,			--製程累績不良品數
 TotalScrapQty				int					NOT NULL	DEFAULT 0,			--製程累績報廢數
 AlreadyQty					int					NOT NULL	DEFAULT 0,			--製程目前可生產但未生產數
 WipQty						int					NOT NULL	DEFAULT 0,			--製程在製累績數量
 ReturnQty					int					NULL		DEFAULT 0,			--返修到此製程的數量
 ParallelStatus				nvarchar(1)			NOT NULL	DEFAULT 'N',		--製程是否可以跳站
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MoProcess PRIMARY KEY (MoProcessId),
 CONSTRAINT AK1_MoProcess UNIQUE (MoId, SortNumber),
 CONSTRAINT FK1_MoProcess FOREIGN KEY (MoId) REFERENCES MES.ManufactureOrder(MoId),
 CONSTRAINT FK2_MoProcess FOREIGN KEY (MoRoutingId) REFERENCES MES.MoRouting(MoRoutingId),
 CONSTRAINT FK3_MoProcess FOREIGN KEY (ProcessId) REFERENCES MES.Process(ProcessId),
 CONSTRAINT FK4_MoProcess FOREIGN KEY (RoutingProcessId) REFERENCES MES.RoutingProcess(RoutingProcessId)
)

CREATE TABLE MES.MoProcessItem( /* 製令製程屬性設定 */
 MoProcessItemId			int					identity,				
 MoProcessId				int					NOT NULL,						--製令製程ID						
 ItemNo						nvarchar(30)		NOT NULL,						--製程需綁定項目
 ChkUnique					nvarchar(1)			NOT NULL,						--檢核重覆值 >> BAS.[Status] Y:需檢核重覆值；N:不需檢核重覆值
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MoProcessItem PRIMARY KEY (MoProcessItemId),
 CONSTRAINT AK1_MoProcessItem UNIQUE (MoProcessId, ItemNo),
 CONSTRAINT FK1_MoProcessItem FOREIGN KEY (MoProcessId) REFERENCES MES.MoProcess(MoProcessId)
)

CREATE TABLE MES.MoProcessTransaction( /* 當生產模式為模造等計量產品時, 不使用BarcodeProcess & Barcode, 使用此Table 計算產量 */
 TransactionId				int					NOT NULL	identity,
 MoId 						int					NOT NULL	DEFAULT -1,
 LineId 					int					NOT NULL	DEFAULT -1,
 MachineId					int					NOT NULL	DEFAULT -1,
 MoProcessId				int					NOT NULL,
 NextMoProcessId			int					NOT NULL,						--下一站製程
 ProdStatus 				nvarchar(2)			NOT NULL,						--過站狀態 P:良品, F:不良品, S:報廢品
 Quantity					int					NOT NULL,						--數量
 CauseNo					nvarchar(100)		NOT NULL,						--NG原因
 InputDate					datetime			NOT NULL,						--過站時間
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MoProcessTransaction PRIMARY KEY (TransactionId),
 CONSTRAINT FK1_MoProcessTransaction FOREIGN KEY (MachineId) REFERENCES MES.Machine(MachineId),
 CONSTRAINT FK2_MoProcessTransaction FOREIGN KEY (MoId) REFERENCES MES.ManufactureOrder(MoId),
 CONSTRAINT FK3_MoProcessTransaction FOREIGN KEY (MoProcessId) REFERENCES MES.MoProcess(MoProcessId)
)

CREATE TABLE MES.MoProcessQcItem( /* 由途程品號量測項目帶過來 */
 MpqId						int					NOT NULL	identity,
 MoProcessId 				int					NOT NULL,						--製令製程ID
 QcItemId 					int					NOT NULL,						--量測項目ID
 QcItemDesc 				nvarchar(100)			NULL,						--量測項目描述
 DesignValue 				float					NULL,						--標準值
 UpperTolerance 			float					NULL,						--上公差
 LowerTolerance 			float					NULL,						--下公差
 BallMark 					nvarchar(100)			NULL,						--球標
 Unit 						nvarchar(100)			NULL,						--單位
 QmmDetailId	 			int						NULL,						--量測設備
 Remark 					nvarchar(255)			NULL,						--檢測備註
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MoProcessQcItem PRIMARY KEY (MpqId),
 CONSTRAINT FK1_MoProcessQcItem FOREIGN KEY (MoProcessId) REFERENCES MES.MoProcess(MoProcessId),
 CONSTRAINT FK2_MoProcessQcItem FOREIGN KEY (QcItemId) REFERENCES QMS.QcItem(QcItemId),
)

CREATE TABLE MES.MoSetting( /*MES製令設定*/
 MoSettingId				int					identity,
 MoId						int					NOT NULL,
 LotStatus					nvarchar(1)			NOT NULL,						 --批量生產 >> BAS.[Status] Y:批量生產；N:非批量生產(DEFAULT)
 LotQty						int					NOT NULL,						 --批量條碼數量
 BarcodePrefix				nvarchar(50)			NULL,						 --條碼前綴碼
 SequenceLen				int						NULL,					     --流水碼長度
 BarcodePostfix				nvarchar(20)			NULL,						 --條碼後綴碼
 NgToBarcode				nvarchar(2)				NULL	DEFAULT 'Y',		 --當批量生產時, NG數量是否需要刷取NG_BARCODE, Y:需刷取 N:不需刷取 T:直接更改原條碼狀態
 MtlControl					nvarchar(1)			NOT NULL	DEFAULT 'Y',		 --是否做物料控管
 OutputBarcodeFlag			nvarchar(2)			NOT NULL	DEFAULT 'N',		 --製程中是否會產出新條碼(數量過站用，非數量過站不可選，預設為N)	
 [Source]					nvarchar(2)			NOT NULL,						 --發料來源 MR:領退料單 BP:批量條碼
 PVTQCFlag					nvarchar(2)			NOT NULL    DEFAULT 'N',		 --量產前是否需要先進行試產檢
 OQcCheckType				nvarchar(2)			NOT NULL	DEFAULT 'N',	     --未完成出貨檢是否可以過站
 BarcodeCtrl				nvarchar(2)			NOT NULL	DEFAULT 'Y',	     --是否要條碼控管
 TrayBarcode				nvarchar(2)			NOT NULL	DEFAULT 'N',	     --Tray盤上是否有雷刻條碼
 MrType						nvarchar(2)			NOT NULL	DEFAULT 'Y',	     --是否可以領取上個月的料 Y:可領取 N:不可領取
 NgToDisassembly			nvarchar(2)			NOT NULL	DEFAULT 'N',	     --批量NG品是否需拆解(如:鏡頭)
 VehicleFlag				nvarchar(2)			NOT NULL	DEFAULT 'N',	     --是否使用載盤上料
 FullPassQcFlag				nvarchar(2)			NOT NULL	DEFAULT 'Y',	     --入庫前是否需QC檢查
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	 --新增日期
 LastModifiedDate			datetime				NULL,						 --修改日期
 CreateBy					int					NOT NULL,						 --新增人員
 LastModifiedBy				int						NULL,						 --修改人員
 CONSTRAINT PK_MoSetting PRIMARY KEY (MoSettingId),
 CONSTRAINT FK1_MoSetting FOREIGN KEY (MoId) REFERENCES MES.ManufactureOrder(MoId)
)

CREATE TABLE MES.MoMtlSetting( /*製令用料設定檔*/
 MoMtlSettingId				int					identity,
 MoId						int					NOT NULL,
 WoDetailId					int					NOT NULL,						--料號(ERP工單單身)
 UomId						int					NOT NULL,						--MD004 單位 >> 參照 ERP品號換算單位檔(INVMD)
 CompositionQuantity		float				    NULL,						--單位用量
 Base						float					NULL,						--底數
 LossRate					float					NULL,						--損耗率
 BomSequence				nvarchar(4)				NULL,						--序號
 ProcessBinding				nvarchar(1)			NOT NULL,						--物料需不需要在製程綁定, 需要才要有MoProcessId >> >> BAS.[Status] Y:需製程綁定；N:不需製程綁定(DEF)
 MoProcessId				int						NULL,						--製程(MoProcess裡選擇)
 BarcodeCtrl				nvarchar(1)			NOT NULL,						--綁定條碼 >> BAS.[Status] Y:需綁定條碼；N:不需綁定條碼
 MainBarcode				nvarchar(1)			NOT NULL,						--檢核條碼 >> BAS.[Status] Y:製令之主條碼,過站時需檢核；N:過站時不需檢核此條碼
 ControlType				nvarchar(1)			NOT NULL,						--扣帳方式 >> BAS.[Type] 1:開工前手動開立領料單；2:完工入庫自動開立領料單(用投入量*單位用量)
 DecompositionFlag			nvarchar(1)			NOT NULL	DEFAULT 'N',		--料號是否會有1PCS領多條碼情況 Y/N
 BindType					nvarchar(1)			NOT NULL	DEFAULT 'N',		--物料上料類型 B:條碼 L:批號 N:不上料
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MoMtlSetting PRIMARY KEY (MoMtlSettingId),
 CONSTRAINT AK1_MoMtlSetting UNIQUE (MoId, WoDetailId),
 CONSTRAINT FK1_MoMtlSetting FOREIGN KEY (MoId) REFERENCES MES.ManufactureOrder(MoId),
 CONSTRAINT FK2_MoMtlSetting FOREIGN KEY (WoDetailId) REFERENCES MES.WoDetail(WoDetailId),
 CONSTRAINT FK3_MoMtlSetting FOREIGN KEY (UomId) REFERENCES PDM.UnitOfMeasure(UomId),
 CONSTRAINT FK4_MoMtlSetting FOREIGN KEY (MoProcessId) REFERENCES MES.MoProcess(MoProcessId)
)

CREATE TABLE MES.MoMtlSettingAttribute( /*製令用料屬性檔, Ex:製令為模架, 上料為模仁/精定位時需要設定要輸入穴號*/
 MoMtlSettingAttrId			int					identity,
 MoMtlSettingId				int					NOT NULL,
 ItemNo						nvarchar(30)		NOT NULL,						--項目代碼
 ItemName					nvarchar(30)		NOT NULL,						--項目名稱 
 ItemDesc					nvarchar(200)			NULL,						--項目說明
 ChkUnique					nvarchar(1)			NOT NULL,						--檢核重覆值 >> BAS.[Status] Y:需檢核重覆值；N:不需檢核重覆值
 ChkRange					nvarchar(1)				NULL	DEFAULT 'N',		--檢核重複值範圍 >> Y:整張製令皆須檢核 N:只檢核單一物料
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MoMtlSettingAttr PRIMARY KEY (MoMtlSettingAttrId),
 CONSTRAINT AK1_MoMtlSettingAttr UNIQUE (MoMtlSettingId, ItemNo),
 CONSTRAINT FK1_MoMtlSettingAttr FOREIGN KEY (MoMtlSettingId) REFERENCES MES.MoMtlSetting(MoMtlSettingId)
)

CREATE TABLE MES.Barcode( /* 條碼主檔 */
 BarcodeId					int					identity,
 BarcodeNo					nvarchar(60)		NOT NULL,
 CurrentMoProcessId			int					NOT NULL,						--目前途程ID 
 NextMoProcessId			int					NOT NULL,						--下一途程ID
 MoId						int						NULL,						--條碼所屬製令(最後一張)
 BarcodeQty					int					NOT NULL,						--條碼數量
 ParentBarcode				int					    NULL,						--上層條碼(鏡頭條碼用)
 PrevBarcodeId				int					    NULL,						--上一層條碼(其他除鏡頭條碼之外的使用)
 BarcodeProcessId           int						NULL,						--條碼生產過站ID
 Memo						nvarchar(100)			NULL,
 PrintCount					int					NOT NULL	DEFAULT 0,			--列印次數
 CurrentProdStatus			nvarchar(10)		NOT NULL,						--產品目前狀態 >> BAS.[Type] P:良品；F:不良品；S:報廢品 ;RW:返修 ;M:誤判 ;R:特採 ;SR:原站返修 ;AM:重新量測 ;PG:包裝條碼狀態
 BarcodeStatus				nvarchar(10)		NOT NULL,						--條碼狀態 >> BAS.[Type] 1:In Process；0:Finish(加工結束/強制結束)；6:檢貨後取消；7:強制結束；8:製令完工入庫；9:檢貨完成；10:已出貨; 11:客退入庫; 12:包裝條碼
 TrayLocation				nvarchar(100)			NULL,						--TRAY盤位置
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Barcode PRIMARY KEY (BarcodeId),
 CONSTRAINT AK1_Barcode UNIQUE (BarcodeNo),
 CONSTRAINT FK1_Barcode FOREIGN KEY (MoId) REFERENCES MES.ManufactureOrder(MoId)
)

CREATE TABLE MES.BarcodeLog( /* 條碼主檔歷史檔 */
 BarcodeId					int					NOT NULL,
 BarcodeNo					nvarchar(60)		NOT NULL,
 CurrentMoProcessId			int					    NULL,						--目前途程ID 
 NextMoProcessId			int					    NULL,						--下一途程ID
 MoId						int					    NULL,						--條碼所屬製令(最後一張)
 BarcodeQty					int					NOT NULL,						--條碼數量
 ParentBarcode				int					    NULL,						--上一層條碼
 BarcodeProcessId           int					NOT NULL,						--條碼生產過站ID
 Memo						nvarchar(100)			NULL,						--備註
 CurrentProdStatus			nvarchar(10)		NOT NULL,						--產品目前狀態 >> BAS.[Type] P:良品；F:不良品；S:報廢品 ;RW:返修 ;M:誤判 ;R:特採 ;SR:原站返修 ;AM:重新量測
 BarcodeStatus				nvarchar(10)		NOT NULL,						--條碼狀態 >> BAS.[Type] 1:In Process；0:Finish(加工結束/強制結束)；6:檢貨後取消；7:強制結束；8:製令完工入庫；9:檢貨完成；10:已出貨; 11:客退入庫
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 CreateBy					int					NOT NULL,						--新增人員
 CONSTRAINT PK_BarcodeLog PRIMARY KEY (BarcodeId,CreateDate)
)

CREATE TABLE MES.BarcodeComponent( /* 條碼元件 (僅上料管理有設定時使用)*/
 BarcodeComponentId			int					identity,
 AssemblyBarcodeId			int					NOT NULL,						--主產品條碼ID
 ComponentBarcodeId			int						NULL,						--零件條碼ID
 LotNumberId				int						NULL,						--批號單頭ID
 MoProcessId				int						NULL,						--製令製程ID
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_BarcodeComponent PRIMARY KEY (BarcodeComponentId),
 --CONSTRAINT AK1_BarcodeComponent UNIQUE (AssemblyBarcodeId, ComponentBarcodeId),
 --CONSTRAINT AK2_BarcodeComponent UNIQUE (AssemblyBarcodeId, LotNumberId),
 CONSTRAINT FK1_BarcodeComponent FOREIGN KEY (AssemblyBarcodeId) REFERENCES MES.Barcode(BarcodeId),
 CONSTRAINT FK2_BarcodeComponent FOREIGN KEY (ComponentBarcodeId) REFERENCES MES.Barcode(BarcodeId),
 CONSTRAINT FK3_BarcodeComponent FOREIGN KEY (LotNumberId) REFERENCES SCM.LotNumber(LotNumberId)
)

CREATE TABLE MES.BarcodeComponentAttribute( /* 條碼元件屬性(穴別) , 鉗工組模時模仁,精定位等穴號資訊*/
 BarcodeComponentId			int					NOT NULL,
 ItemNo						nvarchar(30)		NOT NULL,						--元件屬性項目
 ItemValue					nvarchar(30)		NOT NULL,						--元件屬性項目值
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT AK1_BarcodeComponentAttribute UNIQUE (BarcodeComponentId, ItemNo),
 CONSTRAINT FK1_BarcodeComponentAttribute FOREIGN KEY (BarcodeComponentId) REFERENCES MES.BarcodeComponent(BarcodeComponentId)
)

CREATE TABLE MES.BarcodeComponentDeleteLog( /* 刪除條碼元件紀錄*/
 LogId						int					identity,
 AssemblyBarcodeId			int					NOT NULL,						--主產品條碼ID
 ComponentBarcodeId			int					NOT NULL,						--零件條碼ID
 LotNumberId				int						NULL,						--批號單頭ID
 MoProcessId				int						NULL,						--在那個製程所綁定
 ItemNo						nvarchar(30)		NOT NULL,						--元件屬性項目
 ItemValue					nvarchar(30)		NOT NULL,						--元件屬性項目值
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_BarcodeComponentDeleteLog PRIMARY KEY (LogId),
 CONSTRAINT FK1_BarcodeComponentDeleteLog FOREIGN KEY (AssemblyBarcodeId) REFERENCES MES.Barcode(BarcodeId),
 CONSTRAINT FK2_BarcodeComponentDeleteLog FOREIGN KEY (ComponentBarcodeId) REFERENCES MES.Barcode(BarcodeId),
 CONSTRAINT FK3_BarcodeComponentDeleteLog FOREIGN KEY (MoProcessId) REFERENCES MES.MoProcess(MoProcessId),
 CONSTRAINT FK4_BarcodeComponent FOREIGN KEY (LotNumberId) REFERENCES SCM.LotNumber(LotNumberId)
)

CREATE TABLE MES.BarcodeComponentLog( /* 條碼元件置換記錄 */
 ComponentLogId				int					identity,
 BarcodeComponentId			int					NOT NULL,
 OldComponentId				int					NOT NULL,						--零件條碼ID
 NewComponentId				int					NOT NULL,						--零件條碼ID
 Remark						nvarchar(250)		NULL,							--置換原因
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_BarcodeComponentLog PRIMARY KEY (ComponentLogId),
 CONSTRAINT FK1_BarcodeComponentLog FOREIGN KEY (OldComponentId) REFERENCES MES.Barcode(BarcodeId),
 CONSTRAINT FK2_BarcodeComponentLog FOREIGN KEY (NewComponentId) REFERENCES MES.Barcode(BarcodeId)
)

CREATE TABLE MES.BarcodePrint( /*批量條碼, 之後就會轉到BARCODE_INFO */
 PrintId					int					identity,
 BarcodeNo					nvarchar(60)		NOT NULL,						--批量條碼代碼
 BarcodeQty 				int					NOT NULL,						--批量條碼數量
 MoSettingId				int					NOT NULL,						--製令設定ID
 ParentBarcode				nvarchar(60)			NULL,
 PrintCnt					int					NOT NULL,						--列印次數
 PrintStatus				nvarchar(2)			NOT NULL,						--批量條碼狀態 >> BAS.[Type] P:良品；N:不良品；S:報廢品 
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_BarcodePrint PRIMARY KEY(PrintId),
 CONSTRAINT AK_BarcodePrint UNIQUE(BarcodeNo),
 CONSTRAINT FK1_BarcodePrint FOREIGN KEY (MoSettingId) REFERENCES MES.MoSetting(MoSettingId),
)

CREATE TABLE MES.BarcodeProcess( /*條碼生產過站資訊 */
 BarcodeProcessId			int					identity,
 BarcodeId					int					NOT NULL,
 MoId						int					NOT NULL,
 MachineId					int						NULL,						--機台ID
 MoProcessId				int					NOT NULL,						--製程ID
 ProdStatus 				nvarchar(1)			NOT NULL	DEFAULT 'P',		--狀態: 產品狀態 P:良品, F:不良品, S:報廢品
 StationQty 				int					NOT NULL	DEFAULT 1  ,		--過站數量
 PassQty					int					NOT NULL,						--良品數量
 NgQty						int					NOT NULL,						--不良品/報廢數量
 NgBarcodeId				int						NULL,						--當條碼為有NG品時, NG品放到那個Barcode
 CauseNo					nvarchar(100)			NULL,						--當批量條碼刷取NG條碼時，存放NG CODE
 StartDate					datetime			NOT NULL,						--開工時間
 FinishDate 				datetime				NULL,						--完工時間
 ModeShiftId				int					NOT NULL,						--班別ID
 CycleTime					int						NULL,						--加工時間
 StartUserId				int					NOT NULL,						--開工人員
 FinishUserId				int 					NULL,						--完工人員
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_BarcodeProcess PRIMARY KEY(BarcodeProcessId),
 CONSTRAINT FK1_BarcodeProcess FOREIGN KEY (MoId) REFERENCES MES.ManufactureOrder(MoId),
 CONSTRAINT FK2_BarcodeProcess FOREIGN KEY (MoProcessId) REFERENCES MES.MoProcess(MoProcessId),
 CONSTRAINT FK3_BarcodeProcess FOREIGN KEY (BarcodeId) REFERENCES MES.Barcode(BarcodeId),
 CONSTRAINT FK4_BarcodeProcess FOREIGN KEY (MachineId) REFERENCES MES.Machine(MachineId)
)

CREATE TABLE MES.UserEvent( /*人員事件資料*/
 UserEventId				int					identity,
 UserEventItemId			int					NOT NULL,						--人員事件
 StartDate					datetime			NOT NULL,						--開始時間
 FinishDate					datetime				NULL,						--結束時間
 OperationTime				int						NULL,						--作業時間 > 單位 秒
 Operator					int					NOT NULL,						--操作員 > BAS.[User].UserId
 [Source]					nvarchar(200)		NOT NULL,						--來源
 [Status]					nvarchar(2)			NOT NULL	DEFAULT 'A',		--狀態碼 A:啟用;S:停用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_UserEvent PRIMARY KEY (UserEventId),
 CONSTRAINT FK1_UserEvent FOREIGN KEY (UserEventItemId) REFERENCES MES.UserEventItem(UserEventItemId)
)

CREATE TABLE MES.MachineEvent( /*機台事件資料*/
 MachineEventId				int					identity, 
 MachineEventItemId			int					NOT NULL,						--抓MachineEventItem
 StartDate					datetime			NOT NULL,						--開始時間
 FinishDate					datetime				NULL,						--結束時間
 OperationTime				int						NULL,						--作業時間 > 單位 秒
 Operator					int					NOT NULL,						--操作員 > BAS.[User].UserId
 [Source]					nvarchar(200)		NOT NULL,						--來源
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼		   >> S:關閉 ; A:啟用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MachineEvent PRIMARY KEY (MachineEventId),
 CONSTRAINT FK1_MachineEvent FOREIGN KEY (MachineEventItemId) REFERENCES MES.MachineEventItem(MachineEventItemId)
)

CREATE TABLE MES.BarcodeProcessEvent( /*--條碼過站事件 */
 BarcodeEventId				int					identity,
 BarcodeProcessId			int					NOT NULL,						--條碼生產過站資訊
 ProcessEventItemId			int					NOT NULL,						--製程參數事件
 StartDate					datetime			NOT NULL,						--開始時間
 FinishDate					datetime				NULL,						--結束時間
 CycleTime					int						NULL	DEFAULT 0,			--作業秒數
 OperatorId					int					NOT NULL,						--作業員USER_ID
 [Source]					nvarchar(200)		NOT NULL,						--來源
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼		   >> S:關閉 ; A:啟用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_BarcodeProcessEvent PRIMARY KEY (BarcodeEventId),
 CONSTRAINT FK1_BarcodeProcessEvent FOREIGN KEY (BarcodeProcessId) REFERENCES MES.BarcodeProcess(BarcodeProcessId),
 CONSTRAINT FK2_BarcodeProcessEvent FOREIGN KEY (ProcessEventItemId) REFERENCES MES.ProcessEventItem(ProcessEventItemId)
)

CREATE TABLE MES.BarcodeProcessAttribute( /* 條碼屬性檔, 條碼上刻字碼等都放在這 */
 ProcessAttrId				int					identity,
 BarcodeProcessId			int						NULL,						--可以為空, 因為可能在非製造過程中置換
 ItemNo 					nvarchar(100)		NOT NULL,						--屬性代碼
 ItemValue					nvarchar(100)		NOT NULL,						--屬性值
 ItemSeq					int						NULL,						--流水號
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_BarcodeProcessAttr PRIMARY KEY (ProcessAttrId),
 CONSTRAINT FK1_BarcodeProcessAttr FOREIGN KEY (BarcodeProcessId) REFERENCES MES.BarcodeProcess(BarcodeProcessId)
)

CREATE TABLE MES.BarcodeProcessAttributeLog( /* 條碼屬性Log檔, 置換歷程記錄 */
 ProcessAttrLogId			int					identity,
 OriBarcodeAttrId			int					NOT	NULL,						--原始BarcodeAttrId
 NewBarcodeAttrId			int					NOT	NULL,						--欲置換BarcodeAttrId
 OriBarcodeId			    int					NOT	NULL,						--原始條碼ID
 NewBarcodeId	    		int					NOT	NULL,						--欲置換條碼ID
 OriMoId		        	int					NOT	NULL,						--原始製令ID
 NewMoId		        	int					NOT	NULL,						--欲置換製令ID
 OriItemNo 					nvarchar(10)		NOT NULL,						--原始屬性條碼
 NewItemNo 					nvarchar(10)		NOT NULL,						--欲置換屬性條碼
 OriItemValue				nvarchar(100)		NOT NULL,						--原始屬性資料
 NewItemValue				nvarchar(100)		NOT NULL,						--欲置換屬性資料
 Remark						nvarchar(250)			NULL,						--置換說明
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_BarcodeProcessAttrLog PRIMARY KEY (ProcessAttrLogId),
 CONSTRAINT FK1_BarcodeProcessAttrLog FOREIGN KEY (NewBarcodeAttrId) REFERENCES MES.BarcodeAttribute(BarcodeAttrId)
)

CREATE TABLE MES.BpalDetail( /* 條碼屬性Log檔單身, 置換歷程記錄 */
 BpalDetailId				int					identity,
 ProcessAttrLogId			int					NOT	NULL,						--原始ProcessAttrLogId
 BarcodeId					int					NOT	NULL,						--條碼Id
 MoId						int					NOT	NULL,						--製令ID
 [Status]					nvarchar(2)			NOT	NULL	DEFAULT 'N',		--量測狀態 N:未完成置換 A:置換已開工 Y:已完成置換
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_BpalDetail PRIMARY KEY (BpalDetailId),
 CONSTRAINT FK1_BpalDetail FOREIGN KEY (ProcessAttrLogId) REFERENCES MES.BarcodeProcessAttributeLog(ProcessAttrLogId),
 CONSTRAINT FK2_BpalDetail FOREIGN KEY (BarcodeId) REFERENCES MES.Barcode(BarcodeId)
)

CREATE TABLE MES.BarcodeAttribute( /* 簡化版條碼屬性，新增屬性時同時新增此TABLE，只記最新 */
 BarcodeAttrId				int					identity,
 MoId						int						NULL,						--製令ID
 BarcodeId					int					NOT	NULL,						--條碼ID
 ItemNo 					nvarchar(100)		NOT NULL,						--屬性代碼
 ItemValue					nvarchar(100)		NOT NULL,						--屬性值
 ItemSeq					int						NULL,						--流水號
 DateCode					datetime				NULL,						--屬性日期
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_BarcodeAttribute PRIMARY KEY (BarcodeAttrId),
 CONSTRAINT AK1_BarcodeAttribute UNIQUE (BarcodeId, MoId, ItemNo, ItemValue),
 CONSTRAINT FK1_BarcodeAttribute FOREIGN KEY (BarcodeId) REFERENCES MES.Barcode(BarcodeId)
)

CREATE TABLE MES.BarcodeTransfer( /* 條碼移轉, 拆併批數量移轉使用 */
 TransactionId				int					identity,
 FromBarcodeProcessId		int					NOT NULL,						--來源條碼歷程ID
 ToBarcodeProcessId			int					NOT NULL,						--目標條碼歷程ID
 TransactionDate			datetime			NOT NULL,						--移轉時間
 FromBarcodeId				int					NOT NULL,						--轉出的BARCODE_ID
 ToBarcodeId				int					NOT NULL,						--轉入的BARCODE_ID
 TransactionQty				int					NOT NULL,						--移轉數量
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_BarcodeTransfer PRIMARY KEY(TransactionId),
 CONSTRAINT FK1_BarcodeTransfer FOREIGN KEY (FromBarcodeProcessId) REFERENCES MES.BarcodeProcess(BarcodeProcessId),
 CONSTRAINT FK2_BarcodeTransfer FOREIGN KEY (ToBarcodeProcessId) REFERENCES MES.BarcodeProcess(BarcodeProcessId)
)

CREATE TABLE MES.BarcodeScrapLog( /* 條碼報廢紀錄檔 */
 BarcodeScrapId				int					NOT NULL,
 BarcodeId					int					NOT NULL,						--報廢條碼
 ScrapReason				nvarchar(100)			NULL,						--報廢原因
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_BarcodeScrapLog PRIMARY KEY (BarcodeScrapId),
 CONSTRAINT FK1_BarcodeScrapLog FOREIGN KEY (BarcodeId) REFERENCES MES.Barcode(BarcodeId)
)

CREATE TABLE MES.QcRecord( /* 量測委託單主表 */
 QcRecordId 				int					identity,
 QcNoticeId					int						NULL,						--Foreign Key QMS.QcNotice
 QcTypeId					int					NOT NULL,						--QcTypeId
 InputType					nvarchar(10)		NOT	NULL	DEFAULT 'Lettering',--Spreadsheet輸入方式 lettering:刻字, cavity:穴號, barcodeNo:條碼
 MoId						int						NULL,						--製令ID
 MoProcessId				int						NULL,						--製令製程ID--用於數量過站
 MtlItemId					int						NULL,						--品號ID(委託量測單用)
 PassQty					int						NULL,						--Pass個數
 NgQty						int						NULL,						--NG個數
 SystemStatus				nvarchar(2)				NULL,						--系統判斷結果 N:未判定, P:良品, R:需補正, F:不良品
 QcStatus					nvarchar(2)				NULL,						--檢驗人員判斷結果 N:未判定, P:良品, R:需補正, F:不良品
 QcUserId					int						NULL,						--檢驗員判斷人員
 CauseId					int						NULL,						--不良代碼, FINAL_RESULT = 'F'時, 一定要有值
 Remark						nvarchar(200)			NULL,						--品保判斷時備註
 DefaultSpreadsheetData		nvarchar(MAX)			NULL,						--此紀錄尚未上傳EXCEL，以此SpreadsheetData為主
 SpreadsheetData			nvarchar(MAX)			NULL,						--spreadsheet data(未處理資料)
 PlanningSpreadsheetDate	nvarchar(MAX)			NULL,						--排程暫存Spreadsheet DATA
 DefaultFileId				int						NULL,						--若此紀錄尚未上傳EXCEL，以此FILE ID為主
 CurrentFileId				int						NULL,						--此紀錄目前上傳之EXCEL FILE ID (未來可能停用)
 CheckQcMeasureData			nvarchar(2)			NOT	NULL	DEFAULT 'N',		--此量測記錄是否已上傳數據 Y:已上傳/N:新單據/P:已上傳但未判定/A:已送測但未確認/C:已收件但未量測/F:已駁回/S:開始量測/V:已作廢
 MeasurementPlanning		nvarchar(2)			NOT	NULL	DEFAULT 'N',		--排程是否已經安排完成 Y:已完成 N:未完成
 SupportAqFlag				nvarchar(2)			NOT NULL	DEFAULT 'Y',		--是否支援品異單 Y:支援 N:不支援
 SupportProcessFlag			nvarchar(2)			NOT NULL	DEFAULT 'N',		--是否支援卡控製程站別 Y:支援 N:不支援
 ResolveFileJson			nvarchar(MAX)		NOT NULL,						--解析量測數據暫存JSON
 DisallowanceReason			nvarchar(255)			NULL,						--量測需求駁回原因
 RequestDate				datetime				NULL,						--量測需求送測時間
 ReceiptDate				datetime				NULL,						--量測需求確認收貨時間
 QcStartDate				datetime				NULL,						--開始量測時間
 QcFinishDate				datetime				NULL,						--量測完成時間
 UrgentFlag					nvarchar(1)			NOT	NULL	DEFAULT 'N',		--是否急件 Y:是 N:否
 [Source]					nvarchar(100)		NOT	NULL	DEFAULT 'QcDataManagement',	--建立單據來源 DEFAULT:QcDataManagement MTF手臂:MTF
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QcRecord PRIMARY KEY(QcRecordId),
 CONSTRAINT FK1_QcRecord FOREIGN KEY (MoId) REFERENCES MES.ManufactureOrder(MoId),
 CONSTRAINT FK2_QcRecord FOREIGN KEY (MoProcessId) REFERENCES MES.MoProcess(MoProcessId),
 CONSTRAINT FK3_QcRecord FOREIGN KEY (QcNoticeId) REFERENCES QMS.QcNotice(QcNoticeId),
 CONSTRAINT FK4_QcRecord FOREIGN KEY (QcUserId) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK5_QcRecord FOREIGN KEY (CauseId) REFERENCES QMS.DefectCause(CauseId),
 CONSTRAINT FK6_QcRecord FOREIGN KEY (QcTypeId) REFERENCES QMS.QcType(QcTypeId),
 CONSTRAINT FK7_QcRecord FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId)
)

CREATE TABLE MES.QcBarcode( /* 量測條碼資料 */
 QcBarcodeId 				int					identity,
 QcRecordId					int					NOT NULL,						--量測記錄ID
 BarcodeProcessId			int						NULL,						--過站ID, OQC時不會有, 值為-1
 BarcodeId					int						NULL	DEFAULT -1,			--檢驗的條碼, 試模時條碼應該為模具條碼
 PassQty					int						NULL,						--Pass個數
 NgQty						int						NULL,						--NG個數
 SystemStatus				nvarchar(2)				NULL,						--系統判斷結果 N:未判定, P:良品, R:需補正, F:不良品
 QcStatus					nvarchar(2)			NOT NULL,						--檢驗人員判斷結果 N:未判定, P:良品, R:需補正, F:不良品
 QcUserId					int					NOT NULL,						--檢驗員判斷人員
 CauseId					int						NULL,						--不良代碼, FINAL_RESULT = 'F'時, 一定要有值
 Remark						nvarchar(200)			NULL,						--品保判斷時備註
 [Status]					nvarchar(10)		NOT	NULL	DEFAULT 'Y',		--量測狀態 Y:已量測完畢 N:未量測完畢，仍可重量(判重新量測情況)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QcBarcode PRIMARY KEY(QcBarcodeId),
 CONSTRAINT FK1_QcBarcode FOREIGN KEY (QcRecordId) REFERENCES MES.QcRecord(QcRecordId),
 CONSTRAINT FK2_QcBarcode FOREIGN KEY (BarcodeId) REFERENCES MES.Barcode(BarcodeId),
 CONSTRAINT FK3_QcBarcode FOREIGN KEY (QcUserId) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK4_QcBarcode FOREIGN KEY (CauseId) REFERENCES QMS.DefectCause(CauseId),
 CONSTRAINT FK5_QcBarcode FOREIGN KEY (BarcodeProcessId) REFERENCES MES.BarcodeProcess(BarcodeProcessId)
)

CREATE TABLE MES.QcMtfBarcode( /* MTF量測條碼資料 */
 QcMtfBarcodeId 			int					identity,
 QcRecordId					int					NOT NULL,						--量測記錄ID
 BarcodeProcessId			int					NOT	NULL,						--過站ID
 BarcodeId					int					NOT	NULL,						--MTF BARCODE
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QcMtfBarcode PRIMARY KEY(QcMtfBarcodeId),
 CONSTRAINT FK1_QcMtfBarcode FOREIGN KEY (QcRecordId) REFERENCES MES.QcRecord(QcRecordId),
 CONSTRAINT FK2_QcMtfBarcode FOREIGN KEY (BarcodeId) REFERENCES MES.Barcode(BarcodeId),
 CONSTRAINT FK3_QcMtfBarcode FOREIGN KEY (BarcodeProcessId) REFERENCES MES.BarcodeProcess(BarcodeProcessId)
)

CREATE TABLE MES.QcReceiptBarcode( /* 加工單位送測條碼 */
 QrbId 						int					identity,
 QcRecordId					int					NOT NULL,						--量測記錄ID
 BarcodeId					int						NULL	DEFAULT -1,			--檢驗的條碼, 試模時條碼應該為模具條碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QcReceiptBarcode PRIMARY KEY(QrbId),
 CONSTRAINT FK1_QcReceiptBarcode FOREIGN KEY (QcRecordId) REFERENCES MES.QcRecord(QcRecordId),
 CONSTRAINT FK2_QcReceiptBarcode FOREIGN KEY (BarcodeId) REFERENCES MES.Barcode(BarcodeId)
)

CREATE TABLE MES.QcMamoChannel( /* 量測推播相關頻道資料 */
 QmcId 						int					identity,
 ModeId 					int					NOT NULL,						--生產模式ID
 ChannelId					int					NOT NULL,						--頻道ID
 PushType					nvarchar(100)		NOT NULL,						--推播類型 Y:量測完成 C:已送測 F:已駁回
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QcMamoChannel PRIMARY KEY(QmcId),
 CONSTRAINT AK1_QcMamoChannel UNIQUE (ModeId, PushType, ChannelId),
 CONSTRAINT FK1_QcMamoChannel FOREIGN KEY (ModeId) REFERENCES MES.ProdMode(ModeId),
 CONSTRAINT FK2_QcMamoChannel FOREIGN KEY (ChannelId) REFERENCES MAMO.Channels(ChannelId)
)

CREATE TABLE MES.QcRecordFile( /* 量測上傳檔案(EXCEL) */
 QcRecordFileId 			int					identity,
 QcRecordId					int					NOT NULL,						--量測記錄ID
 FileType					nvarchar(20)		NOT NULL	DEFAULT 'other',	--檔案類型 resolve:解析檔案/other:其他檔案/file-management:檔案歸檔功能
 FileId						int						NULL,						--檔案ID
 PhysicalPath				nvarchar(MAX)			NULL,						--檔案實體路徑
 InputTyp					nvarchar(20)			NULL,						--檔案類型為歸檔時，需維護歸檔方式
 BarcodeId					int						NULL,						--隸屬條碼ID
 Cavity						nvarchar(100)			NULL,						--隸屬穴號
 LotNumber					nvarchar(20)			NULL,						--隸屬批號
 EffectiveDiameter			nvarchar(2)				NULL,						--R1、R2
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QcRecordFile PRIMARY KEY(QcRecordFileId),
 CONSTRAINT FK1_QcRecordFile FOREIGN KEY (QcRecordId) REFERENCES MES.QcRecord(QcRecordId),
 CONSTRAINT FK2_QcRecordFile FOREIGN KEY (BarcodeId) REFERENCES MES.Barcode(BarcodeId)
)

CREATE TABLE MES.QcGoodsReceipt( /* 進貨檢量測單據 */
 QcGoodsReceiptId 			int					identity,
 QcRecordId					int					NOT	NULL,						--主量測單據ID
 GrDetailId					int					NOT	NULL,						--進貨單身ID
 QcStatus					nvarchar(1)			NOT	NULL	DEFAULT 'A',		--進貨檢單據檢驗狀態 Y:合格 W:特採 N:驗退 A:尚未檢驗
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QcGoodsReceipt PRIMARY KEY(QcGoodsReceiptId),
 CONSTRAINT FK1_QcGoodsReceipt FOREIGN KEY (QcRecordId) REFERENCES MES.QcRecord(QcRecordId),
 CONSTRAINT FK2_QcGoodsReceipt FOREIGN KEY (GrDetailId) REFERENCES SCM.GrDetail(GrDetailId),
)

CREATE TABLE MES.QcGoodsReceiptLog( /* 進料檢驗作業紀錄 */
 LogId 						int					identity,
 QcGoodsReceiptId			int					NOT	NULL,						--進貨檢驗單ID
 ReceiptQty					int					NOT NULL,						--進貨數量
 AcceptQty					int					NOT NULL,						--驗收數量
 ReturnQty					int					NOT NULL,						--驗退數量
 AcceptanceDate				datetime			NOT NULL,						--驗收日期
 QcStatus					nvarchar(1)			NOT NULL,						--檢驗狀態 0.免檢、1.待驗、2.合格、3.不良、4.特採
 QuickStatus				nvarchar(1)			NOT NULL,						--是否為急件 Y/N
 Remark						nvarchar(255)			NULL,						--備註
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QcGoodsReceiptLog PRIMARY KEY(LogId),
 CONSTRAINT AK1_QcGoodsReceiptLog UNIQUE(QcGoodsReceiptId),
 CONSTRAINT FK1_QcGoodsReceiptLog FOREIGN KEY (QcGoodsReceiptId) REFERENCES MES.QcGoodsReceipt(QcGoodsReceiptId),
)

CREATE TABLE MES.QcGoodsReceiptLogFile( /* 進料檢驗作業紀錄-檔案 */
 LogId						int					NOT	NULL,						--進料檢驗作業ID
 FileId						int					NOT NULL,						--檔案ID
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QcGoodsReceiptLogFile PRIMARY KEY(FileId,LogId),
 CONSTRAINT AK1_QcGoodsReceiptLogFile UNIQUE(FileId),
 CONSTRAINT FK1_QcGoodsReceiptLogFile FOREIGN KEY (LogId) REFERENCES MES.QcGoodsReceiptLog(LogId),
 CONSTRAINT FK2_QcGoodsReceiptLogFile FOREIGN KEY (FileId) REFERENCES BAS.[File](FileId),
)

CREATE TABLE MES.QcReceiptTemplate( /* 量測送測樣板 */
 TemplateId 				int					identity,
 CompanyId					int					NOT	NULL,						--公司別ID
 ModeId						int					NOT	NULL,						--生產模式ID
 TemplateName				nvarchar(100)		NOT	NULL,						--樣板名稱
 Status						nvarchar(1)			NOT	NULL	DEFAULT 'A',		--狀態 A:啟用/S:停用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QcReceiptTemplate PRIMARY KEY(TemplateId),
 CONSTRAINT FK1_QcReceiptTemplate FOREIGN KEY (ModeId) REFERENCES MES.ProdMode(ModeId),
 CONSTRAINT FK2_QcReceiptTemplate FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)

CREATE TABLE MES.QcReceiptTemplateDetail( /* 量測送測樣板-詳細資料 */
 DetailId 					int					identity,
 TemplateId					int					NOT	NULL,						--樣板ID
 QcItemId					int					NOT	NULL,						--生產模式ID
 DesignValue				float				    NULL,						--尺寸設計值
 UpperTolerance				float				    NULL,						--尺寸設計值上限公差
 LowerTolerance				float				    NULL,						--尺寸設計值下限公差
 QmmDetailId	 			int						NULL,						--量測設備
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QcReceiptTemplateDetail PRIMARY KEY(DetailId),
 CONSTRAINT AK1_QcReceiptTemplateDetail UNIQUE(TemplateId, QcItemId),
 CONSTRAINT FK1_QcReceiptTemplateDetail FOREIGN KEY (TemplateId) REFERENCES MES.QcReceiptTemplate(TemplateId),
 CONSTRAINT FK2_QcReceiptTemplateDetail FOREIGN KEY (QcItemId) REFERENCES QMS.QcItem(QcItemId),
 CONSTRAINT FK3_QcReceiptTemplateDetail FOREIGN KEY (QmmDetailId) REFERENCES QMS.QmmDetail(QmmDetailId),
)

CREATE TABLE MES.QcDeliveryInspection( /* 供應商出貨檢驗單據 */
 QcDeliveryInspectionId 	int					identity,
 QcDiNo 					nvarchar(11)		NOT NULL,						--出貨檢驗單號
 QcRecordId					int					NOT	NULL,						--主量測單據ID
 PoDetailId					int					NOT	NULL,						--採購單身ID
 QcStatus					nvarchar(1)			NOT	NULL	DEFAULT 'A',		--進貨檢單據檢驗狀態 Y:合格 W:特採 N:驗退 A:尚未檢驗
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateUserName				nvarchar(10)		NOT NULL,						--新增人員(供應商)
 UpdateUserName				nvarchar(10)			NULL,						--修改人員(供應商)
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QcDeliveryInspection PRIMARY KEY(QcDeliveryInspectionId),
 CONSTRAINT AK1_QcDeliveryInspection UNIQUE (QcDiNo),
 CONSTRAINT FK1_QcDeliveryInspection FOREIGN KEY (QcRecordId) REFERENCES MES.QcRecord(QcRecordId),
 CONSTRAINT FK2_QcDeliveryInspection FOREIGN KEY (PoDetailId) REFERENCES SCM.PoDetail(PoDetailId),
)

CREATE TABLE MES.JigBarcode( /* 條碼綁定治具資料 */
 JigBarcodeId 				int					identity,
 JigId						int					NOT NULL,						--治具ID
 BarcodeNo					nvarchar(60)		NOT NULL,						--條碼
 MoProcessId				int					NOT NULL,						--製令製程ID
 WorkingFlag				nvarchar(1)			NOT NULL	DEFAULT 'N',		--條碼完工狀態 N:未開工 A:加工中 F:已完工
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_JigBarcode PRIMARY KEY(JigBarcodeId),
 CONSTRAINT FK1_JigBarcode FOREIGN KEY (JigId) REFERENCES MES.Jig(JigId),
 CONSTRAINT FK2_JigBarcode FOREIGN KEY (MoProcessId) REFERENCES MES.MoProcess(MoProcessId),
)

CREATE TABLE MES.JigBarcodeHistory( /* 條碼綁定治具歷史資料 */
 HistoryId					int 				identity,
 JigId						int					NOT NULL,						--治具ID
 BarcodeNo					nvarchar(60)		NOT NULL,						--條碼
 MoProcessId				int					NOT NULL,						--製令製程ID
 WorkingFlag				nvarchar(1)			NOT NULL	DEFAULT 'N',		--條碼完工狀態
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 CreateBy					int					NOT NULL,						--新增人員
 CONSTRAINT PK_JigBarcodeHistory PRIMARY KEY(HistoryId),
 CONSTRAINT FK1_JigBarcodeHistory FOREIGN KEY (JigId) REFERENCES MES.Jig(JigId),
 CONSTRAINT FK2_JigBarcodeHistory FOREIGN KEY (MoProcessId) REFERENCES MES.MoProcess(MoProcessId),
)

CREATE TABLE MES.AdvancedLog( /* 進階功能LOG */
 LogId						int 				identity,
 BarcodeId					int						NULL,						--條碼ID
 MoProcessId				int					NOT NULL,						--當前製令製程ID
 OriNextMoProcessId			int					NOT NULL,						--原下一站製程ID
 NextMoProcessId			int					NOT NULL,						--更改後下一站製程ID
 OriQty						int					NOT NULL,						--原(條碼)數量
 Qty						int					NOT NULL,						--更改後數量
 OriCurrentProdStatus		nvarchar(2)			NOT NULL,						--原條碼狀態
 CurrentProdStatus			nvarchar(2)			NOT NULL,						--更改後條碼狀態
 CauseNo					nvarchar(100)			NULL,						--異常原因
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_AdvancedLog PRIMARY KEY(LogId),
 CONSTRAINT FK1_AdvancedLog FOREIGN KEY (BarcodeId) REFERENCES MES.Barcode(BarcodeId),
 CONSTRAINT FK2_AdvancedLog FOREIGN KEY (MoProcessId) REFERENCES MES.MoProcess(MoProcessId)
)

CREATE TABLE MES.Package( /* 包裝主表 */
 PackageId 					int					identity,
 CompanyId 					int					NOT NULL,						--公司別ID
 PackageNo 					nvarchar(100)		NOT NULL,						--包裝編號
 ParentPackageId 			int						NULL,						--上層包裝ID
 [Status]					nvarchar(2)			NOT NULL	DEFAULT '1',		--包裝條碼狀態 BAS.[Status] 1:包裝中 2:已入庫 3:已領料
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Package PRIMARY KEY(PackageId),
 CONSTRAINT AK1_Package UNIQUE (PackageNo),
 CONSTRAINT FK1_Package FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_Package FOREIGN KEY (ParentPackageId) REFERENCES MES.Package(ParentPackageId)
)

CREATE TABLE MES.PackageBarcode( /* 包裝條碼資料 */
 PackageBarcodeId 			int					identity,
 PackageId 					int					NOT NULL,						--包裝ID
 BarcodeId					int					NOT NULL,						--包裝內條碼ID
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_PackageBarcode PRIMARY KEY(PackageBarcodeId),
 CONSTRAINT AK1_PackageBarcode UNIQUE (PackageId, BarcodeId),
 CONSTRAINT FK1_PackageBarcode FOREIGN KEY (PackageId) REFERENCES MES.Package(PackageId),
 CONSTRAINT FK2_PackageBarcode FOREIGN KEY (BarcodeId) REFERENCES MES.Barcode(BarcodeId)
)

CREATE TABLE MES.PackageBarcodeReference( /* 包裝條碼詳細資料(綁定條碼) */
 PbrId 						int					identity,
 PackageBarcodeId 			int					NOT NULL,						--包裝條碼ID
 BarcodeId					int					NOT NULL,						--產品條碼ID
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_PackageBarcodeReference PRIMARY KEY(PbrId),
 CONSTRAINT AK1_PackageBarcodeReference UNIQUE (PackageBarcodeId, BarcodeId),
 CONSTRAINT FK1_PackageBarcodeReference FOREIGN KEY (PackageBarcodeId) REFERENCES MES.PackageBarcode(PackageBarcodeId),
 CONSTRAINT FK2_PackageBarcodeReference FOREIGN KEY (BarcodeId) REFERENCES MES.Barcode(BarcodeId)
)

CREATE TABLE MES.SignProduct( /*產品進出簽收*/
 SignProductId					int					identity,
 MoId							int					NOT	NULL,						--製令ID
 SourcePlaceId					int						NULL,						--來源目標(倉儲)
 PurposePlaceId					int					NOT	NULL,						--目的目標(倉儲)
 MoveQty						int					NOT NULL,						--移動數量
 Remark							nvarchar(255)			NULL,						--備註
 SignProductUserId				int					NOT NULL,						--簽收者 BAS.[User] 
 DocDate						datetime			NOT NULL,						--開單日期
 CreateDate						datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate				datetime				NULL,						--修改日期
 CreateBy						int					NOT NULL,						--新增人員
 LastModifiedBy					int						NULL,						--修改人員
 CONSTRAINT PK_SignProduct PRIMARY KEY (SignProductId),
 CONSTRAINT FK1_SignProduct FOREIGN KEY (MoId) REFERENCES MES.ManufactureOrder(MoId),
 CONSTRAINT FK2_SignProduct FOREIGN KEY (SourcePlaceId) REFERENCES MES.WarehouseLocation(LocationId),
 CONSTRAINT FK3_SignProduct FOREIGN KEY (PurposePlaceId) REFERENCES MES.WarehouseLocation(LocationId),
 CONSTRAINT FK4_SignProduct FOREIGN KEY (SignProductUserId) REFERENCES BAS.[User](UserId)
)

CREATE TABLE MES.TerminateOfMoLog( /* 強制結束製令流程LOG */
 LogId						int 				identity,
 MoId						int					NOT	NULL,						--製令ID
 Remark						nvarchar(255)		NOT	NULL,						--備註
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_TerminateOfMoLog PRIMARY KEY(LogId),
 CONSTRAINT FK1_TerminateOfMoLog FOREIGN KEY (MoId) REFERENCES MES.ManufactureOrder(MoId)
)

CREATE TABLE MES.TOmBarcodeLog( /* 強制結束條碼LOG */
 TOmLogId					int 				identity,
 LogId						int					NOT	NULL,						--強制結束製令ID
 BarcodeId					int					NOT	NULL,						--強制結束條碼ID
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_TOmBarcodeLog PRIMARY KEY(TOmLogId),
 CONSTRAINT FK1_TOmBarcodeLog FOREIGN KEY (LogId) REFERENCES MES.TerminateOfMoLog(LogId),
 CONSTRAINT FK2_TOmBarcodeLog FOREIGN KEY (BarcodeId) REFERENCES MES.Barcode(BarcodeId)
)

CREATE TABLE MES.KeyenceBarcodeLog( /* 強制結束條碼LOG */
 LogId						int 				identity,
 KeyenceId					int					NOT	NULL,						--Keyence ID
 BarcodeId					int					NOT	NULL,						--條碼ID
 MoProcessId				int					NOT	NULL,						--製令製程ID
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_KeyenceBarcodeLog PRIMARY KEY(LogId),
 CONSTRAINT FK1_KeyenceBarcodeLog FOREIGN KEY (KeyenceId) REFERENCES MES.Keyence(KeyenceId),
 CONSTRAINT FK2_KeyenceBarcodeLog FOREIGN KEY (BarcodeId) REFERENCES MES.Barcode(BarcodeId),
 CONSTRAINT FK3_KeyenceBarcodeLog FOREIGN KEY (MoProcessId) REFERENCES MES.MoProcess(MoProcessId)
)

CREATE TABLE MES.KeyenceProcessTemp( /*  Keyence模式過站入口Temp */
 KPTId 					int						identity,
 Barcode				nvarchar(200)			NULL,						--條碼
 MoProcessId			int						NULL,						--過站ID
 MachineId				int						NULL,						--機器ID
 UserId					int						NULL,						--使用者ID
 KeyenceId				int						NULL,						--KeyenceID
 QtyBarcodeStatus		int						NULL,						-- 
 QtyLotNgCauseNo		int						NULL,						-- 
 QtyNextMoProcessId		int						NULL,						-- 
 ItemNo					nvarchar(200)			NULL,							
 ItemValue				nvarchar(200)			NULL,							
 ChkUnique				nvarchar(200)			NULL	default 'N',							
 DateCode				nvarchar(200)			NULL,							
 CreateDate				datetime				NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate		datetime				NULL,						--修改日期
 CreateBy				int						NULL,						--新增人員
 LastModifiedBy			int						NULL,						--修改人員
 CONSTRAINT PK_KeyenceProcessTemp PRIMARY KEY(KPTId),
 )

CREATE TABLE MES.MfgBarcode( /*MES 1.0 BarcodeNo清單*/
 MfgBarcodeId				int					identity,
 BarcodeNo					nvarchar(50)		NOT NULL,						--條碼編號
 BarcodeQty					int					NOT NULL,						--條碼數量
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MfgBarcode PRIMARY KEY (MfgBarcodeId)
)

CREATE TABLE MES.DepartmentalSecondment( /* 部門借調清單 */
 DsId						int 				identity,
 UserId						int					NOT	NULL,						--人員ID
 DepartmentId				int					NOT	NULL,						--部門ID
 StartDate					datetime			NOT	NULL,						--借調開始時間
 EndDate					datetime			NOT	NULL,						--借調停止時間
 [Status]					nvarchar(1)			NOT	NULL,						--狀態 A:啟用 S:停止
 Remark						nvarchar(255)			NULL,						--備註
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DepartmentalSecondment PRIMARY KEY(DsId),
 CONSTRAINT FK1_DepartmentalSecondment FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK2_DepartmentalSecondment FOREIGN KEY (DepartmentId) REFERENCES BAS.Department(DepartmentId)
)

END /*生產管理*/

BEGIN /*託外生產*/
CREATE TABLE MES.OutsourcingProduction( /*託外生產單頭*/
 OspId						int					identity,
 CompanyId					int					NOT NULL,						--公司
 DepartmentId				int					NOT NULL,						--所屬部門
 OspNo						nvarchar(50)		NOT NULL,						--託外生產單號
 OspDate					datetime			NOT NULL,						--託外生產日期
 SupplierId					int						NULL,						--供應商
 OspStatus					nvarchar(2)			NOT NULL,						--供料狀態 >> BAS.[Type] N：不供料；Y：供料
 OspDesc					nvarchar(100)			NULL,						--描述
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼 Y:確認 N:未確認 P:BPM簽核中 E:BPM駁回修改
 TransferBpmDate			datetime				NULL,						--拋轉BPM日期
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_OutsourcingProduction PRIMARY KEY (OspId),
 CONSTRAINT AK1_OutsourcingProduction UNIQUE (CompanyId, OspNo),
 CONSTRAINT FK1_OutsourcingProduction FOREIGN KEY (DepartmentId) REFERENCES BAS.Department(DepartmentId),
 CONSTRAINT FK2_OutsourcingProduction FOREIGN KEY (SupplierId) REFERENCES SCM.Supplier(SupplierId),
 CONSTRAINT FK3_OutsourcingProduction FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)

CREATE TABLE MES.OspDetail( /*託外生產單身*/
 OspDetailId				int					identity,
 OspId						int					NOT NULL,						--所屬單頭
 MoId						int					NOT NULL,						--製令
 MoProcessId				int					NOT NULL,						--製令製程
 ProcessCheckStatus			nvarchar(2)			NOT NULL,						--是否支援工程檢
 ProcessCheckType			nvarchar(2)			NOT NULL,						--工程檢頻率 1:全檢 2:抽檢
 OspQty						int					NOT NULL,						--託外生產數量
 SuppliedQty				int					NOT NULL,						--供貨生產數量
 ProcessCode				nvarchar(20)			NULL,						--製程代碼(ERP)
 ProcessCodeName			nvarchar(200)			NULL,						--製程代碼名稱(ERP)
 ExpectedDate				datetime			NOT NULL,						--預計回廠日期
 Remark						nvarchar(255)			NULL,						--備註
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_OspDetail PRIMARY KEY (OspDetailId),
 CONSTRAINT FK1_OspDetail FOREIGN KEY (OspId) REFERENCES MES.OutsourcingProduction(OspId),
 CONSTRAINT FK2_OspDetail FOREIGN KEY (MoId) REFERENCES MES.ManufactureOrder(MoId),
 CONSTRAINT FK3_OspDetail FOREIGN KEY (MoProcessId) REFERENCES MES.MoProcess(MoProcessId),
)

CREATE TABLE MES.OspBarcode( /*託外生產單身物件*/
 OspBarcodeId				int					identity,
 OspDetailId				int					NOT NULL,						--所屬單身
 BarcodeNo					nvarchar(50)		NOT NULL,						--條碼
 BarcodeQty					int					NOT NULL,						--數量
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_OspBarcode PRIMARY KEY (OspBarcodeId),
 CONSTRAINT FK1_OspBarcode FOREIGN KEY (OspDetailId) REFERENCES MES.OspDetail(OspDetailId)
)

CREATE TABLE MES.OspBpmLog( /*託外生產單頭 BPM相關LOG*/
 LogId						int					identity,
 OspId						int					NOT NULL,						--托外生產單ID
 RootId						nvarchar(30)		NOT NULL,						--任務編號
 BpmNo						nvarchar(100)		NOT	NULL,						--BPM單號
 TransferBpmDate			datetime			NOT NULL,						--拋轉BPM日期
 BpmStatus					nvarchar(2)			NOT NULL,						--拋轉狀態
 ConfirmUser				nvarchar(100)		NOT NULL,						--最後核單USER NO
 ErrorMessage				nvarchar(1000)			NULL,						--Error訊息
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_OspBpmLog PRIMARY KEY (LogId),
 CONSTRAINT FK1_OspBpmLog FOREIGN KEY (OspId) REFERENCES MES.OutsourcingProduction(OspId)
)

CREATE TABLE MES.OspReceipt( /*託外進貨單頭 MOCTH*/
 OsrId						int					identity,
 CompanyId					int					NOT NULL,						--公司別
 OsrErpPrefix				nvarchar(4)				NULL,						--TH001 單別
 OsrErpNo					nvarchar(11)			NULL,						--TH002	單號
 ReceiptDate				datetime			NOT NULL,						--TH003 進貨日期
 FactoryCode				nvarchar(6)			NOT NULL	DEFAULT 'ZY01',		--TH004 廠別(需要在Company裡加連結
 SupplierId					int					NOT NULL,						--TH005 供應商/廠別
 SupplierSo					nvarchar(20)			NULL,						--TH006	廠商單號/放空值即可,預留欄位
 CurrencyCode				nvarchar(4)			NOT NULL,						--TH007 幣別
 Exchange					float				NOT NULL	DEFAULT 0.0000,		--TH008 匯率
 RowCnt						int					NOT NULL	DEFAULT 0,			--TH009 件數
 Remark						nvarchar(255)			NULL,						--TH010 備註
 UiNo						nvarchar(20)			NULL,						--TH011	統一編號
 InvoiceType				nvarchar(1)				NULL,						--TH012	發票聯數 1.二聯式、2.三聯式、3.二聯式收銀機發票、4.三聯式收銀機發票、5.電子計算機發票、6.免用統一發票、A.農產品收購憑證、G.海關代徵完稅憑證、N.不可抵扣專用發票、S.可抵扣專用發票、T.運輸發票、W.廢舊物資收購憑證、Z.其他
 InvoiceDate				datetime				NULL,						--TH013	發票日期
 InvoiceNo					nvarchar(50)			NULL,						--TH014	發票號碼
 TaxType					nvarchar(1)				NULL,						--TH015 1.應稅內含、2.應稅外加、3.零稅率、4.免稅、9.不計稅
 DeductType					nvarchar(1)				NULL,						--TH016 1.可扣抵進貨及費用 2.可扣抵固定資產 3.不可扣抵進貨及費用 4.不可扣抵固定資產
 ReserveFlag				nvarchar(1)			NOT NULL	DEFAULT 'N',		--TH017 保留欄位(N), 不需維護, ERP用
 OrigAmount					float				NOT NULL	DEFAULT 0.0000,		--TH018 原幣加工金額
 DeductAmount				float				NOT NULL	DEFAULT 0.0000,		--TH019 扣款金額
 OrigTax					float				NOT NULL	DEFAULT 0.0000,		--TH020 原幣稅額
 ReceiptAmount				float				NOT NULL	DEFAULT 0.0000,		--TH021	進貨費用
 Quantity					int					NOT NULL	DEFAULT 0,			--TH022	數量合計(單身數量合計)
 ConfirmStatus				nvarchar(1)			NOT NULL	DEFAULT 'N',		--TH023 確認碼(Y/N/V)
 ConfirmUserId				int						NULL,						--TH023 確認者
 TransferStatus				nvarchar(1)			NOT	NULL	DEFAULT 'N',		--TH023 拋轉狀態 Y/N
 RenewFlag					nvarchar(1)			NOT NULL	DEFAULT 'N',		--TH024 生產記錄更新碼(N) 不需維護, ERP用
 PrintCnt					int					NOT NULL	DEFAULT 0,			--TH025 列印次數
 AutoMaterialBilling		nvarchar(1)			NOT NULL	DEFAULT 'N',		--TH026 自動扣料更新碼
 OrigPreTaxAmount			float				NOT NULL	DEFAULT 0.0000,		--TH027 原未稅貨款金額
 ApplyYYMM					datetime			NOT NULL,						--TH028 申報年月(YYYYMM) 申報年月[FORMATE:YM]&&880701ADD
 DocumentDate				datetime				NULL,						--TH029	單據日期 單據日期[FORMATE:YMD]&&89-05-25ADD
 TaxRate					float					NULL,						--TH030 營業稅率 營業稅率(%)[DEF:9.9999]&&89-06-01ADD//950808 S00-9508031 N5.4->N8.5
 PretaxAmount				float				NOT NULL	DEFAULT 0.0000,		--TH031 本幣未稅金額
 TaxAmount					float				NOT NULL	DEFAULT 0.0000,		--TH032 本幣稅額
 PaymentTerm				nvarchar(6)			NOT NULL,						--TH033 付款條件代號&&89-09-06ADD
 PackageQuantity			int					NOT NULL	DEFAULT 0,			--TH034 包裝單身數量合計
 SupplierPicking			nvarchar(1)			NOT NULL	DEFAULT 'N',		--TH035 廠供料領料更新碼
 ApproveStatus				nvarchar(1)			NOT NULL	DEFAULT '1',		--TH036 簽核狀態碼 0.待處理、1.簽核中、2.退件、3.已核准、4.取消確認中、5.作廢中、6.取消作廢中、N.不執行電子簽核[DEF:"N"] 
 ReserveTaxCode				nvarchar(1)			NOT NULL,						--TH037 0.依品號預設 Y.預設為保稅品 N.預設為非保稅品[DEF:"0"] //920702 MODI 原:Y/N[DEF:"N"]//900802_3.1_S00-9008001_ADD_BY_1325
 SendCount					int					NOT NULL	DEFAULT 0,			--TH038 傳送次數, 永遠傳0
 NoticeFlag					nvarchar(1)			NOT NULL	DEFAULT 'N',		--TH041 通知碼, 永遠傳N
 TaxCode					nvarchar(3)			NOT NULL,						--TH044 稅別碼 (要找ERP資料來源)
 TaxExchange				float				NOT NULL	DEFAULT 0,			--TH047 稅幣匯率, 永遠帶1
 QcFlag						nvarchar(1)			NOT NULL	DEFAULT 'N',		--TH061	N.不處理、1.待送檢、2.送檢中、3.重新待送檢、4.重新送檢中、5.送檢完成、Y.確認通知完成[DEF:"N"] 
 FlowStatus					nvarchar(1)			NOT NULL	DEFAULT '1',		--MES託外收貨單狀態 (1.New(收貨人), 2.收貨確認(收貨人), 3.QC確認(品保), 4.驗收確認(採購))
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_OspReceipt PRIMARY KEY (OsrId),
 CONSTRAINT FK1_OspReceipt FOREIGN KEY (SupplierId) REFERENCES SCM.Supplier(SupplierId),
 CONSTRAINT FK2_OspReceipt FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)

CREATE TABLE MES.OspReceiptDetail( /*託外進貨單身 MOCTI (待君葦確認託外加工需要欄位)*/ 
 OsrDetailId				int					identity,
 OsrId						int					NOT NULL,
 OspDetailId				int					NOT NULL,						--OspDetail.OspDetailId
 OsrErpPrefix				nvarchar(4)			NOT NULL,					    --TI001 單別
 OsrErpNo					nvarchar(11)		NOT NULL,						--TI002	單號
 OsrSeq						nvarchar(4)			NOT NULL,						--TI003 序號 ex:0001,0002,0003
 MtlItemId					int					NOT NULL,						--TI004,TI005,TI006(品號/品名/規格) Join PDM.MtlItem帶出
 ReceiptQty					int					NOT NULL,						--TI007 進貨數量
 UomId						int					NOT NULL,						--TI008, Join UnitOfMeasure
 InventoryId				int					NOT NULL,						--TI009 進貨庫別 Join Inventory
 LotNumber					nvarchar(20)			NULL,						--TI010	批號
 AvailableDate				datetime				NULL,						--TI011 有效日期, 塞ReceiptDate
 ReCheckDate				datetime				NULL,						--TI012 複檢日期, 塞ReceiptDate
 MoId						int					NOT NULL,						--TI013,TI014 製令單別單號 Join WipOrder
 MoProcessId				int					NOT NULL,						--MES託外製程Id
 ProcessCode				nvarchar(4)			NOT NULL,						--TI015 ERP.CMSMW.MW001, MW002, MW006給使用者自行選擇
 ReceiptPackageQty			int					NOT NULL	DEFAULT 0,			--TI016 進貨包裝數量, 沒使用, Default填0 (不需顯示)
 AcceptancePackageQty		int					NOT NULL	DEFAULT 0,			--TI017	驗收包裝數量 (不需顯示)
 AcceptanceDate				datetime			NOT NULL,						--TI018 驗收日期
 AcceptQty					int					NOT NULL	DEFAULT 0,			--TI019 驗收數量
 AvailableQty				int					NOT NULL	DEFAULT 0,			--TI020 計價數量
 ScriptQty					int					NOT NULL	DEFAULT 0,			--TI021 報廢數量
 ReturnQty					int					NOT NULL	DEFAULT 0,			--TI022 驗退數量
 AvailableUom				int						NULL,						--TI023 計價單位 Join UnitOfMeasure
 OrigUnitPrice				float				NOT NULL	DEFAULT 0.000000,	--TI024 原幣加工單價
 OrigAmount					float				NOT NULL	DEFAULT 0.000000,	--TI025 原幣加工金額
 OrigDiscountAmt			float				NOT NULL	DEFAULT 0.000000,	--TI026 原幣扣款金額
 ReceiptExpense				float				NOT NULL	DEFAULT 0.000000,	--TI027 進貨費用,預設0
 DiscountDescription		nvarchar(40)			NULL,						--TI028 扣款說明
 ProjectCode				nvarchar(20)			NULL,						--TI032 專案代碼
 PaymentHold				nvarchar(1)			NOT NULL	DEFAULT 'N',		--TI033 暫不付款(不用顯示)
 Overdue					nvarchar(1)			NOT NULL	DEFAULT 'N',		--TI034	逾期碼 (Y/N)
 QcStatus					nvarchar(1)			NOT NULL,						--TI035 檢驗狀態 0.免檢、1.待驗、2.合格、3.不良、4.特採
 ReturnCode					nvarchar(1)			NOT NULL	DEFAULT 'N',		--TI036 驗退碼 (Y/N)(不用顯示)
 ConfirmStatus				nvarchar(1)			NOT NULL	DEFAULT 'N',		--TI037 確認碼 (Y/N/V)
 CloseStatus				nvarchar(1)			NOT NULL	DEFAULT 'N',		--TI038 結帳碼 (Y/N)
 ReNewStatus				nvarchar(1)			NOT NULL	DEFAULT 'N',		--TI039 更新碼 (Y/N)(隱藏欄位,不需顯示)
 Remark						nvarchar(255)			NULL,						--TI040 備註
 CostEntry                  nvarchar(1)			NOT NULL	DEFAULT 'N',		--TI041 產生分錄碼-成本 (Y/N)(不需顯示)
 ExpenseEntry               nvarchar(1)			NOT NULL	DEFAULT 'N',		--TI042 產生分錄碼-費用 (Y/N)(不需顯示)
 ConfirmUserId				int						NULL,						--帶BAS.[User].UserId 找出確認者
 ConfirmUser				nvarchar(10)			NULL,						--TI043 確認者
 OrigPreTaxAmt				float				NOT NULL	DEFAULT 0.000000,	--TI044 原幣未稅金額
 OrigTaxAmt					float				NOT NULL	DEFAULT 0.000000,	--TI045 原幣稅金額
 PreTaxAmt					float				NOT NULL	DEFAULT 0.000000,	--TI046 本幣未稅金額
 TaxAmt						float				NOT NULL	DEFAULT 0.000000,	--TI047 本幣稅金額
 UrgentMtl                  nvarchar(1)			NOT NULL	DEFAULT 'N',        --TI048 急料 Join WipOrder.UrgentMtl
 ReserveTaxCode             nvarchar(1)			NOT NULL,						--TI052 保稅碼 (Y/N Y.預設為保稅品 N.預設為非保稅品) Join OspReceipt.ReserveTaxCode 單頭為0則Join INVMB.MB020 
 ApproveStatus				nvarchar(1)			NOT NULL	DEFAULT 'N',		--TI053 簽核狀態碼 0.待處理、1.簽核中、2.退件、3.已核准、4.取消確認中、5.作廢中、6.取消作廢中、N.不執行電子簽核[DEF:"N"] 
 TransferStatus             nvarchar(1)			NOT NULL	DEFAULT 'N',		--TI054 拋轉碼 (Y/N)
 ProcessStatus				nvarchar(1)			NOT NULL,						--過站狀態 (Y/N) 預設N
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_OspReceiptDetail PRIMARY KEY (OsrDetailId),
 CONSTRAINT AK1_OspReceiptDetail UNIQUE (OsrId, OspDetailId),
 CONSTRAINT AK2_OspReceiptDetail UNIQUE (OsrId, OsrSeq),
 CONSTRAINT FK1_OspReceiptDetail FOREIGN KEY (OsrId) REFERENCES MES.OspReceipt(OsrId),
 CONSTRAINT FK2_OspReceiptDetail FOREIGN KEY (OspDetailId) REFERENCES MES.OspDetail(OspDetailId),
 CONSTRAINT FK3_OspReceiptDetail FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId),
 CONSTRAINT FK4_OspReceiptDetail FOREIGN KEY (UomId) REFERENCES PDM.UnitOfMeasure(UomId),
 CONSTRAINT FK5_OspReceiptDetail FOREIGN KEY (InventoryId) REFERENCES SCM.Inventory(InventoryId),
 CONSTRAINT FK6_OspReceiptDetail FOREIGN KEY (MoId) REFERENCES MES.ManufactureOrder(MoId),
 CONSTRAINT FK7_OspReceiptDetail FOREIGN KEY (MoProcessId) REFERENCES MES.MoProcess(MoProcessId),
)

CREATE TABLE MES.OspReceiptBarcode( /*託外收貨時的Barcode*/ 
 OsrDetailId				int					NOT NULL,
 OspBarcodeId				int					NOT NULL,						--所屬託外生產條碼
 [Status]					nvarchar(2)			NOT NULL,						--收貨條碼狀態 I:已收貨 R:收貨後NG 需退回
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_OspReceiptBarcode PRIMARY KEY (OsrDetailId,OspBarcodeId),
 CONSTRAINT FK1_OspReceiptBarcode FOREIGN KEY (OsrDetailId) REFERENCES MES.OspReceiptDetail(OsrDetailId),
 CONSTRAINT FK2_OspReceiptBarcode FOREIGN KEY (OspBarcodeId) REFERENCES MES.OspBarcode(OspBarcodeId)
)
END /*託外生產*/

BEGIN /*批次託外*/
CREATE TABLE MES.OspList( /*委外項目清單*/
 OspListId					int					identity,
 DepartmentId				int					NOT NULL,						--委託部門
 MoProcessId				int					NOT NULL,						--製令製程ID
 OspDate					datetime			NOT NULL,						--託外生產日期
 SupplierId					int						NULL,						--供應商
 Quantity					int					NOT NULL,						--委外數量
 ProcessCode				nvarchar(20)			NULL,						--製程代碼(ERP)
 ProcessCodeName			nvarchar(200)			NULL,						--製程代碼名稱(ERP)
 ExpectedDate				datetime			NOT NULL,						--預計回廠日期
 ProcessCheckStatus			nvarchar(2)			NOT NULL,						--是否支援工程檢
 ProcessCheckType			nvarchar(2)			NOT NULL,						--工程檢頻率 1:全檢 2:抽檢
 PoConfirmDate				datetime				NULL,						--採購確認日期
 PoRejectionReason			nvarchar(255)			NULL,						--採購駁回原因
 PoUserId					int						NULL,						--採購人員ID
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼 Y:已拋轉委外單 N:新單據未拋轉 P:採購已確認 F:採購已駁回
 Remark						nvarchar(255)			NULL,						--備註
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_OspList PRIMARY KEY (OspListId),
 CONSTRAINT FK1_OspList FOREIGN KEY (DepartmentId) REFERENCES BAS.Department(DepartmentId),
 CONSTRAINT FK2_OspList FOREIGN KEY (MoProcessId) REFERENCES MES.MoProcess(MoProcessId),
)

CREATE TABLE MES.OspListStatusLog( /*委外項目清單 LOG紀錄*/
 LogId						int					identity,
 OspListId					int					NOT NULL,						--委託預排製程ID
 PoConfirmDate				datetime				NULL,						--採購確認日期
 PoRejectionReason			nvarchar(255)			NULL,						--採購駁回原因
 PoUserId					int						NULL,						--採購人員ID
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 CONSTRAINT PK_OspListStatusLog PRIMARY KEY (LogId),
)

CREATE TABLE MES.OspListMamoChannel( /*委外製程MAMO群組*/
 MamoChannelId				int					identity,
 OspMamoNo					nvarchar(10)		NOT NULL,						--MAMO群組編號
 DepartmentId				int					NOT NULL,						--部門ID
 ChannelId					int					NOT NULL,						--頻道ID
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_OspListMamoChannel PRIMARY KEY (MamoChannelId),
 CONSTRAINT AK_OspListMamoChannel UNIQUE (OspMamoNo),
 CONSTRAINT FK1_OspListMamoChannel FOREIGN KEY (DepartmentId) REFERENCES BAS.Department(DepartmentId),
 CONSTRAINT FK2_OspListMamoChannel FOREIGN KEY (ChannelId) REFERENCES MAMO.Channels(ChannelId),
)
END

BEGIN /*倉儲管理*/
CREATE TABLE MES.MaterialRequisition( /* 製令領料及退料單單頭 */ /*ERP MOCTC*/
 MrId						int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 RequesitionNo				nvarchar(20)		NOT NULL,						--MES領料單號
 MrErpPrefix				nvarchar(4)				NULL,						--TC001 ERP領料單別
 MrErpNo					nvarchar(11)			NULL,						--TC002 ERP領料單別
 MrDate						datetime			NOT NULL,						--TC003 領退料日期 >> YYYYMMDD
 DocDate					datetime			NOT NULL,						--TC014 單據日期
 DocType					nvarchar(2)			NOT NULL,						--TC008 單據性質 >> 54.廠內領料、55.託外領料、56.廠內退料、57.託外退料
 ProductionLine				nvarchar(10)			NULL,						--TC005 生產線別 >> 參照 ERP生產線資料單頭檔(CMSMD)
 ContractManufacturer		nvarchar(10)			NULL,						--TC006 加工廠商 >> 參照 ERP廠商基本資料檔(PURMA)，且只有託外單別(55)可以選填
 Remark						nvarchar(50)			NULL,						--TC007 備註 
 JournalStatus				nvarchar(1)			NOT NULL	DEFAULT 'Y',		--TC011 產生分錄碼 >> Y/N&&880624ADD
 PriorityType				nvarchar(1)			NOT NULL	DEFAULT '1',		--需再確認邏輯, MOCTC.TC012 1.依製令單號、2.依材料品號、3.依輸入序號+材料品號、4.材料品號+依輸入序號 //940915  7.1 ADD 3.4.  &&880701ADD
 NegativeStatus				nvarchar(1)			NOT NULL	DEFAULT 'N',		--庫存不足照領 MOCTC.TC013
 SignupStatus				nvarchar(1)			NOT NULL	DEFAULT '0',		--簽核狀態碼 MOCTC.TC016 0.待處理、1.簽核中、2.退件、3.已核准、4.取消確認中、5.作廢中、6.取消作廢中、N.不執行電子簽核[DEF:"N"] //900117
 SourceType					nvarchar(1)			NOT NULL	DEFAULT '0',		--來源別 MOCTC.TC019 0.自行輸入、1.製令單、2.生產入庫自動扣料、3.託外進自動扣料、4.託外退自動扣料、5.移轉單、6.挪料單、7.託外進貨廠供料、8.託外退貨廠供料[DEF:"0"]       //920421 V5.1 ADD //931127 S00-9311018:多7、8
 TransferStatus				nvarchar(1)			NOT NULL	DEFAULT 'N',		--是否已拋轉ERP Y/N
 ConfirmStatus				nvarchar(1)			NOT NULL	DEFAULT 'N',		--確認碼 MOCTC.TC009
 ConfirmUserId				int						NULL,						--確認者 MOCTC.TC015
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MaterialRequisition PRIMARY KEY (MrId),
 CONSTRAINT AK1_MaterialRequisition UNIQUE (CompanyId, MrErpPrefix, MrErpNo),
 CONSTRAINT FK1_MaterialRequisition FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)

CREATE TABLE MES.MrWipOrder( /* 領料單-WIP(匯入ERP時需匯入MOCTC時使用) */ /*ERP MOCTD*/
 MrId						int					NOT NULL,
 MoId						int					NOT NULL,						--MES製令ID
 WoErpPrefix				nvarchar(4)			NOT NULL,						--TA001 製令單別 >> 參照 ERP各系統單據設定檔(CMSMQ)
 WoErpNo					nvarchar(11)		NOT NULL,						--TA002 製令單號
 MrType						nvarchar(1)			NOT NULL,						--TD005 領料方式 >> 1.成套領料、2.補足已領套數、3.補足需領用量
 Quantity					float				NOT NULL,						--TD006 領退料套數
 UomId						int					NOT NULL,						--領退料單位
 InventoryId				int					NOT NULL,						--MES庫別ID
 StorageLocation			nvarchar(100)			NULL,						--儲位
 SubInventoryCode			nvarchar(10)		NOT NULL,						--TD007 庫別代碼
 RequisitionCode			nvarchar(1)			NOT NULL,						--TD008 領料碼 >> 1.逐批領料、2.自動扣料、3.單獨領料
 NegativeStatus				nvarchar(1)			NOT NULL,						--TD013 庫存不足照領 >> 只帶預設值 DEFAULT 'N'
 Remark						nvarchar(255)			NULL,						--TD014 備註
 MaterialCategory			nvarchar(1)			NOT NULL,						--TD015 材料型態 >> 1.直接材料、2.間接材料、5.客戶供料、*.全部&&880701ADD
 SubinventoryType			nvarchar(1)			NOT NULL,						--TD017 庫別型態 >> 1.篩選條件、2.指定庫別
 LineSeq					nvarchar(4)			NOT NULL,						--TD023 輸入序號
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MrWipOrder PRIMARY KEY (MrId, MoId),
 CONSTRAINT FK1_MrWipOrder FOREIGN KEY (MrId) REFERENCES MES.MaterialRequisition(MrId),
 CONSTRAINT FK2_MrWipOrder FOREIGN KEY (MoId) REFERENCES MES.ManufactureOrder(MoId),
 CONSTRAINT FK3_MrWipOrder FOREIGN KEY (InventoryId) REFERENCES SCM.Inventory(InventoryId),
 CONSTRAINT FK4_MrWipOrder FOREIGN KEY (UomId) REFERENCES PDM.UnitOfMeasure(UomId)
)

CREATE TABLE MES.MrDetail( /*領料單單身 */ /*ERP MOCTE*/
 MrDetailId					int					identity,
 MrId						int					NOT NULL,
 MrSequence					nvarchar(4)			NOT NULL,						--TE003, 領料序號
 MtlItemId					int					NOT NULL,						--所屬品號
 Quantity					int					NOT NULL,						--TE005, 預計領退料數量
 ActualQuantity				int					NOT NULL,						--TE005, 實際領退料數量
 UomId						int					NOT NULL,						--MES單位ID
 Unit						nvarchar(6)			NOT NULL,						--TE006, 單位
 InventoryId				int					NOT NULL,						--MES庫別ID
 SubInventoryCode			nvarchar(10)		NOT NULL,						--TE008, 庫別
 ProcessCode				nvarchar(4)			NOT NULL	DEFAULT '****',		--TE009, 製程代碼, 拋ERP用, MES不用顯示
 LotNumber					nvarchar(20)			NULL,						--TE010, 批號(材料庫存批號)
 MoId						int						NULL,						--MES製令ID(CMSMQ MQ019若為N，可不綁定製令)
 WoErpPrefix				nvarchar(4)			NOT NULL,						--TA001 製令單別 >> 參照 ERP各系統單據設定檔(CMSMQ)
 WoErpNo					nvarchar(11)		NOT NULL,						--TA002 製令單號
 DetailDesc					nvarchar(60)			NULL,						--TE013, 領料說明
 Remark						nvarchar(255)			NULL,						--TE014, 備註
 MaterialCategory			nvarchar(1)			NOT NULL	DEFAULT '1',		--TE016, 1.直接材料、2.間接材料、5.客戶供料
 ConfirmStatus				nvarchar(1)			NOT NULL	DEFAULT 'N',		--TE019, 確認碼
 ProjectCode				nvarchar(20)			NULL,						--TE020, 專案代號
 BondedStatus				nvarchar(1)			NOT NULL	DEFAULT 'N',		--TE023, 保稅碼
 SubstituteStatus			nvarchar(1)			NOT NULL	DEFAULT 'N',		--TD024, 領替代料
 OfficialItemStatus			nvarchar(1)			NOT NULL	DEFAULT '2',		--TD026, 1.工程品號 2.正式品號 MES不需顯示
 SubstitutionId				int						NULL,						--MES被取替代料ID
 SubstitutionMtlItemNo		nvarchar(40)			NULL,						--取替代料品號
 SubstituteProcessCode		nvarchar(4)			NOT NULL	DEFAULT '****',		--TE028, 被取替代料製程 MES不顯示
 StorageLocation			nvarchar(120)			NULL,						--TE029, 儲位
 SubstituteQty				float						NULL,						--TE030, 取替代數量
 SubstituteRate				float					NULL,						--TE032, 取替代比率
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MrDetail PRIMARY KEY (MrDetailId),
 CONSTRAINT AK1_MrDetail UNIQUE (MrId, MrSequence),
 CONSTRAINT FK1_MrDetail FOREIGN KEY (MrId) REFERENCES MES.MaterialRequisition(MrId),
 CONSTRAINT FK2_MrDetail FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId),
 CONSTRAINT FK3_MrDetail FOREIGN KEY (UomId) REFERENCES PDM.UnitOfMeasure(UomId),
 CONSTRAINT FK4_MrDetail FOREIGN KEY (InventoryId) REFERENCES SCM.Inventory(InventoryId),
 CONSTRAINT FK5_MrDetail FOREIGN KEY (MoId) REFERENCES MES.ManufactureOrder(MoId),
)

CREATE TABLE MES.MrBarcodeRegister( /*MES領料條碼限制, 當MoMtlSetting = Y時, 條碼過站需check這裡是否有發料 */
 BarcodeRegisterId			int					identity,
 MrDetailId					int					NOT NULL,
 BarcodeNo					nvarchar(32)		NOT NULL,						--條碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MrBarcodeRegister PRIMARY KEY (BarcodeRegisterId),
 CONSTRAINT AK1_MrBarcodeRegister UNIQUE (MrDetailId, BarcodeNo),
 CONSTRAINT FK1_MrBarcodeRegister FOREIGN KEY (MrDetailId) REFERENCES MES.MrDetail(MrDetailId)
)

CREATE TABLE MES.MrBarcodeReRegister( /*MES退料條碼資料 */
 BarcodeReRegisterId		int					identity,
 MrDetailId					int					NOT NULL,
 BarcodeNo					nvarchar(32)		NOT NULL,						--條碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MrBarcodeReRegister PRIMARY KEY (BarcodeReRegisterId),
 CONSTRAINT AK1_MrBarcodeReRegister UNIQUE (MrDetailId, BarcodeNo),
 CONSTRAINT FK1_MrBarcodeReRegister FOREIGN KEY (MrDetailId) REFERENCES MES.MrDetail(MrDetailId)
)

CREATE TABLE MES.MtlTransaction(
 TransactionId				int					identity,
 CompanyId					int					NOT NULL,
 MoId						int					NOT NULL,						--製令
 MoMtlSettingId				int					NOT NULL,						--物料設定ID
 TransactionDate			datetime			NOT NULL,						--物料耗用時間
 BarcodeProcessId			int					NOT NULL,						--過站ID
 MoProcessId				int					NOT NULL,						--製令途程ID
 Quantity					decimal(8,3)		NOT NULL,						--物料耗用數量
 BarcodeNo					nvarchar(60)			NULL,						--BARCODE_NO
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MtlTransaction PRIMARY KEY (TransactionId),
 CONSTRAINT FK1_MtlTransaction FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_MtlTransaction FOREIGN KEY (MoId) REFERENCES MES.ManufactureOrder(MoId)
)

CREATE TABLE MES.MoWarehouseEnry( /*入庫單頭 MOCTH*/
 MweId						int					identity,
 CompanyId					int					NOT NULL,
 MweErpPrefix				nvarchar(4)			NOT NULL,						--TH001 單別
 MweErpNo					nvarchar(11)		NOT NULL,						--TH002	單號
 DocDate				    date				NOT NULL,						--TH029	單據日期 單據日期[FORMATE:YMD]&&89-05-25ADD
 ReceiptDate				date				NOT NULL,						--TH003 進貨日期
 FactoryCode				nvarchar(6)			NOT NULL	DEFAULT 'ZY01',		--TH004 廠別(需要在Company裡加連結
 SupplierId					int					NOT NULL,						--TH005 供應商/廠別
 Remark						nvarchar(255)			NULL,						--TH010 備註
 PackageQuantity			int					NOT NULL	DEFAULT 0,			--TH034 包裝單身數量合計(都是0沒有使用到)
 ReserveTaxCode				nvarchar(1)			NOT NULL,						--TH037 0.依品號預設 Y.預設為保稅品 N.預設為非保稅品[DEF:"0"] //920702 MODI 原:Y/N[DEF:"N"]//900802_3.1_S00-9008001_ADD_BY_1325
 QcFlag						nvarchar(1)			NOT NULL	DEFAULT 'N',		--TH061	N.不處理、1.待送檢、2.送檢中、3.重新待送檢、4.重新送檢中、5.送檢完成、Y.確認通知完成[DEF:"N"] 
 FlowStatus					nvarchar(1)			NOT NULL	DEFAULT '1',		--MES託外收貨單狀態 (1.New(收貨人), 2.收貨確認(收貨人), 3.QC確認(品保), 4.驗收確認(採購))
 TaxCode					nvarchar(3)			NOT NULL,						--TH044 稅別碼 (要找ERP資料來源)
 TaxType					nvarchar(1)				NULL,						--TH015 課稅別 1.應稅內含、2.應稅外加、3.零稅率、4.免稅、9.不計稅
 TaxRate					float					NULL,						--TH030 營業稅率 營業稅率(%)[DEF:9.9999]&&89-06-01ADD//950808 S00-9508031 N5.4->N8.5
 ApplyYYMM					nvarchar(6)			NOT NULL,						--TH028 申報年月(YYYYMM) 申報年月[FORMATE:YM]&&880701ADD
 InvoiceType				nvarchar(1)				NULL,						--TH012	發票聯數 1.二聯式、2.三聯式、3.二聯式收銀機發票、4.三聯式收銀機發票、5.電子計算機發票、6.免用統一發票、A.農產品收購憑證、G.海關代徵完稅憑證、N.不可抵扣專用發票、S.可抵扣專用發票、T.運輸發票、W.廢舊物資收購憑證、Z.其他
 DeductType					nvarchar(1)				NULL,						--TH016 1.可扣抵進貨及費用 2.可扣抵固定資產 3.不可扣抵進貨及費用 4.不可扣抵固定資產
 UiNo						nvarchar(20)			NULL,						--TH011	統一編號
 InvoiceDate				datetime				NULL,						--TH013	發票日期
 InvoiceNo					nvarchar(50)			NULL,						--TH014	發票號碼
 DepartmentNo				nchar(10)				NULL,						--UDF01	費用部門
 CurrencyCode				nvarchar(4)			NOT NULL,						--TH007 幣別
 Exchange					float				NOT NULL	DEFAULT 0.0000,		--TH008 匯率
 RowCnt						int					NOT NULL	DEFAULT 0,			--TH009 件數
 SupplierSo				    nvarchar(20)			NULL,						--TH006	廠商單號/放空值即可,預留欄位
 PaymentTerm				nvarchar(6)			NOT NULL,						--TH033 付款條件代號&&89-09-06A
 AutoMaterialBilling		nvarchar(1)			NOT NULL	DEFAULT 'N',		--TH026 自動扣料更新碼
 SupplierPicking			nvarchar(1)			NOT NULL	DEFAULT 'N',		--TH035 廠供料領料更新碼

 OrigAmount					float				NOT NULL	DEFAULT 0.0000,		--TH018 原幣加工金額
 OrigPreTaxAmount			float				NOT NULL	DEFAULT 0.0000,		--TH027 原未稅貨款金額
 DeductAmount				float				NOT NULL	DEFAULT 0.0000,		--TH019 扣款金額
 OrigTax					float				NOT NULL	DEFAULT 0.0000,		--TH020 原幣稅額
 ReceiptAmount				float				NOT NULL	DEFAULT 0.0000,		--TH021	進貨費用
 PretaxAmount				float				NOT NULL	DEFAULT 0.0000,		--TH031 本幣未稅金額
 TaxAmount					float				NOT NULL	DEFAULT 0.0000,		--TH032 本幣稅額
 Quantity					int					NOT NULL	DEFAULT 0,			--TH022	數量合計(單身進貨數量合計)
 IsAmortize					nvarchar(1)			NOT	NULL    DEFAULT 'N',		--該入庫單是否做攤提


 ConfirmUserId				int						NULL,						--確認者(串單身 TI043)
 ConfirmStatus				nvarchar(1)			NOT NULL	DEFAULT 'N',		--TH023 確認碼(Y/N/V)
 ReConfirmStatus			nvarchar(1)				NULL,						--反確認狀態 1.Null =>都沒有反確認過 2.Y => 反確認中 3. N => 反確認後重新拋單
 TransferStatus				nvarchar(2)			NOT NULL	DEFAULT 'N',		--拋轉ERP狀態 >> BAS.[Status]
 TransferDate				datetime				NULL,						--拋轉ERP時間
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MoWarehouseEnry PRIMARY KEY (MweId),
 CONSTRAINT FK1_MoWarehouseEnry FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_MoWarehouseEnry FOREIGN KEY (SupplierId) REFERENCES SCM.Supplier(SupplierId)
)

CREATE TABLE MES.MweDetail( /*託外進貨單身 MOCTI (待君葦確認託外加工需要欄位)*/ 
 MweDetailId				int					identity,
 MweId						int					NOT NULL,
 MweSeq						nvarchar(4)			NULL,							--TI003 序號 ex:0001,0002,0003
 MoId						int					NOT NULL,						--TI013,TI014 製令單別單號 Join WipOrder
 MtlItemId					int					NOT NULL,						--TI004,TI005,TI006(品號/品名/規格) Join PDM.MtlItem帶出
 UomId						int					NOT NULL,						--TI008 TI023,單位,計價單位 Join UnitOfMeasure
 InventoryId				int					NOT NULL,						--TI009 進貨庫別 Join Inventory
 AvailableDate				date				NULL,							--TI011 有效日期, 塞ReceiptDate
 ReCheckDate				date				NULL,							--TI012 複檢日期, 塞ReceiptDate
 AcceptanceDate				date				NOT NULL,						--TI018 驗收日期
 ReceiptQty					int					NOT NULL,						--TI007 進貨數量
 AcceptQty					int					NOT NULL	DEFAULT 0,			--TI019 驗收數量
 AvailableQty				int					NOT NULL	DEFAULT 0,			--TI020 計價數量
 ScriptQty					int					NOT NULL	DEFAULT 0,			--TI021 報廢數量
 ReturnQty					int					NOT NULL	DEFAULT 0,			--TI022 驗退數量
 OrigUnitPrice				float				NOT NULL	DEFAULT 0.000000,	--TI024 加工單價
 OrigAmount					float				NOT NULL	DEFAULT 0.000000,	--TI025 加工金額
 OrigDiscountAmt			float				NOT NULL	DEFAULT 0.000000,	--TI026 扣款金額
 ReceiptExpense				float				NOT NULL	DEFAULT 0.000000,	--TI027 進貨費用,預設0
 OrigPreTaxAmt				float				NOT NULL	DEFAULT 0.000000,	--TI044 原幣未稅金額
 OrigTaxAmt					float				NOT NULL	DEFAULT 0.000000,	--TI045 原幣稅金額
 PreTaxAmt					float				NOT NULL	DEFAULT 0.000000,	--TI046 本幣未稅額
 TaxAmt						float				NOT NULL	DEFAULT 0.000000,	--TI047 本幣稅額
 DiscountDescription		nvarchar(40)			NULL,						--TI028 扣款說明
 Remark						nvarchar(255)		NULL,							--TI040 備註
 LotNumber					nvarchar(20)		NULL,							--TI010	批號
 ProjectCode				nvarchar(20)		NULL,							--TI032 專案代碼 ERP有下拉
 ProcessCode				nvarchar(4)			NOT NULL,						--TI015 ERP.CMSMW.MW001, MW002, MW006給使用者自行選擇
 QcStatus					nvarchar(1)			NOT NULL	DEFAULT '2',		--TI035 檢驗狀態* (後端自行判斷 0.免檢(不知道規則)1.待驗(都是0),2.合格(沒有報廢),  3.不良(只要有報廢就是),  4.特採)
 Overdue					nvarchar(1)			NOT NULL	DEFAULT 'N',		--TI034	逾期碼 (可以自行勾選)
 ReserveTaxCode             nvarchar(1)			NOT NULL	DEFAULT 'N',		--TI052 保稅碼 (Y/N Y.預設為保稅品 N.預設為非保稅品) Join OspReceipt.ReserveTaxCode 單頭為0則Join INVMB.MB020 (可以自行勾選)
 PaymentHold				nvarchar(1)			NOT NULL	DEFAULT 'N',		--TI033 暫不付款 (可以自行勾選)
 TransferStatus             nvarchar(1)			NOT NULL	DEFAULT 'N',		--TI054 拋轉碼 (Y/N)
  ConfirmUserId				int					NULL,							--TI043 帶BAS.[User].UserId 找出確認者
 ConfirmStatus				nvarchar(1)			NOT NULL	DEFAULT 'N',		--TI037 確認碼
 ReturnCode					nvarchar(1)			NOT NULL	DEFAULT 'N',		--TI036 驗退碼(不用顯示)(無法自行勾選)
 CloseStatus				nvarchar(1)			NOT NULL	DEFAULT 'N',		--TI038 結帳碼 (系統後續更新)
 ReNewStatus				nvarchar(1)			NOT NULL	DEFAULT 'N',		--TI039 更新碼
 CostEntry                  nvarchar(1)			NOT NULL	DEFAULT 'N',		--TI041 產生分錄碼-成本 (Y/N)(無法自行勾選)
 ExpenseEntry               nvarchar(1)			NOT NULL	DEFAULT 'N',		--TI042 產生分錄碼-費用 (Y/N)(無法自行勾選)
 ApproveStatus				nvarchar(1)			NOT NULL	DEFAULT 'N',		--TI053 簽核狀態碼 0.待處理、1.簽核中、2.退件、3.已核准、4.取消確認中、5.作廢中、6.取消作廢中、N.不執行電子簽核[DEF:"N"] 
  ReceiptPackageQty			int					NOT NULL	DEFAULT 0,			--TI016 進貨包裝數量, 沒使用, Default填0
 AcceptancePackageQty		int					NOT NULL	DEFAULT 0,			--TI017	驗收包裝數量
  VoucherPayablePrefix		nvarchar(4)				NULL,						--TI029 應付憑單別* (系統後續更新)
 VoucherPayableNo			nvarchar(11)			NULL,						--TI030 應付憑單號* (系統後續更新)
 VoucherPayableSeq			nvarchar(4)				NULL,						--TI031 應付憑單序號* (系統後續更新)
 LabelStatus				nvarchar(1)			NOT NULL	DEFAULT 'S',		--標籤印製狀態
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MweDetail PRIMARY KEY (MweDetailId),
 CONSTRAINT FK1_MweDetail FOREIGN KEY (MweId) REFERENCES MES.MoWarehouseEnry(MweId),
 CONSTRAINT FK2_MweDetail FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId),
 CONSTRAINT FK3_MweDetail FOREIGN KEY (UomId) REFERENCES PDM.UnitOfMeasure(UomId),
 CONSTRAINT FK4_MweDetail FOREIGN KEY (InventoryId) REFERENCES SCM.Inventory(InventoryId),
 CONSTRAINT FK5_MweDetail FOREIGN KEY (MoId) REFERENCES MES.ManufactureOrder(MoId),
 CONSTRAINT FK6_MweDetail FOREIGN KEY (ConfirmUserId) REFERENCES BAS.[User](UserId),
)

CREATE TABLE MES.MweBarcode( /*入庫單身物件*/
 MweBarcodeId				int					identity,
 MweDetailId				int					NOT NULL,						--所屬單身
 BarcodeId					int					NOT NULL,						--條碼
 MweBarcodeQty				int						NULL,						--條碼數量
 AmortizeTtype				nvarchar(2)				NULL,					    --是否攤提
 QaResult					nvarchar(1)				NULL	DEFAULT 'N',		--是否為檢驗轉報廢 Y/N
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MweBarcode PRIMARY KEY (MweBarcodeId),
 CONSTRAINT FK1_MweBarcode FOREIGN KEY (MweDetailId) REFERENCES MES.MweDetail(MweDetailId),
 CONSTRAINT FK2_MweBarcode FOREIGN KEY (BarcodeId) REFERENCES MES.Barcode(BarcodeId)
)

CREATE TABLE MES.MrBarcodeTransferLog( /*紀錄條碼發料被另外一張製令綁定時的LOG */
 LogId						int					identity,
 MrDetailId					int					NOT NULL,						--領料單單身ID
 MoId						int					NOT NULL,						--此條碼原先製令ID
 BarcodeNo					nvarchar(32)		NOT NULL,						--條碼
 CurrentMoProcessId			int					NOT NULL,						--原先目前製程
 NextMoProcessId			int					NOT NULL,						--原先下一站製程
 CurrentProdStatus			nvarchar(10)		NOT NULL,						--條碼原先品質狀態
 BarcodeStatus				nvarchar(10)		NOT NULL,						--條碼原先狀態
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MrBarcodeTransferLog PRIMARY KEY (LogId),
 CONSTRAINT FK1_MrBarcodeTransferLog FOREIGN KEY (MrDetailId) REFERENCES MES.MrDetail(MrDetailId),
 CONSTRAINT FK2_MrBarcodeTransferLog FOREIGN KEY (MoId) REFERENCES MES.ManufactureOrder(MoId)
)

CREATE TABLE MES.BarcodeMergeLog( /*入庫後拆併盤LOG */
 LogId						int					identity,
 FromBarcodeId				int					NOT NULL,						--來源條碼ID
 ToBarcodeId				int					NOT NULL,						--目標條碼ID
 FromMoId					int					NOT NULL,						--來源條碼製令ID
 ToMoId						int					NOT NULL,						--目標條碼製令ID
 MergeQty					int					NOT NULL,						--拆併數量
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_BarcodeMergeLog PRIMARY KEY (LogId),
 CONSTRAINT FK1_BarcodeMergeLog FOREIGN KEY (FromBarcodeId) REFERENCES MES.Barcode(BarcodeId),
 CONSTRAINT FK2_BarcodeMergeLog FOREIGN KEY (ToBarcodeId) REFERENCES MES.Barcode(BarcodeId),
 CONSTRAINT FK3_BarcodeMergeLog FOREIGN KEY (FromMoId) REFERENCES MES.ManufactureOrder(MoId),
 CONSTRAINT FK4_BarcodeMergeLog FOREIGN KEY (ToMoId) REFERENCES MES.ManufactureOrder(MoId)
)

CREATE TABLE MES.MweBarcodeTransfer( /* 入庫條碼移轉, 拆併盤移轉使用 */
 MweTransactionId			int					identity,
 MweDetailId				int					NOT NULL,						--入庫單單身ID
 TransactionDate			datetime			NOT NULL,						--移轉時間
 FromBarcodeId				int					NOT NULL,						--轉出的BARCODE_ID
 ToBarcodeId				int					NOT NULL,						--轉入的BARCODE_ID
 TransactionQty				int					NOT NULL,						--移轉數量
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MweBarcodeTransfer PRIMARY KEY(MweTransactionId),
 CONSTRAINT AK1_MweBarcodeTransfer UNIQUE (MweDetailId),
 CONSTRAINT AK2_MweBarcodeTransfer FOREIGN KEY (MweDetailId) REFERENCES MES.MweDetail(MweDetailId)
)

END /*倉儲管理*/

BEGIN /*編程管理*/
CREATE TABLE MES.CncProgram( /*編程程式基本資料*/
 CncProgId					int					identity,						--流水碼
 CncProgNo					int					NULL		DEFAULT -1,			--製程編號
 ProcessId					int					NOT NULL	DEFAULT -1,			--所屬製程
 CncProgName				nvarchar(50)		NOT NULL,						--CNC編程名稱
 CncProgApi					nvarchar(100)		NOT NULL,						--CNC編程API名稱 
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼		   >> S:關閉 ; A:啟用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_CncProgram  PRIMARY KEY (CncProgId),
 CONSTRAINT FK1_CncProgram FOREIGN KEY (ProcessId) REFERENCES MES.Process(ProcessId)
)

CREATE TABLE MES.CncParameter( /*編程程式參數基本資料*/
 CncParamId					int					identity,						--流水碼
 CncProgId					int					NOT NULL,						--編程程式ID
 CncParamNo					nvarchar(300)		NOT NULL, 						--參數編號
 CncParamName				nvarchar(500)		NOT NULL,						--參數名稱
 CncParamType				nvarchar(10)		NOT NULL,						--參數輸入方式 >> BAS.[Type] File：上傳檔案；Combobox：下拉；Text：輸入文字；Radio：單選按鈕；Checkbox：多選按鈕 
 CncParamUom				nvarchar(10)		NULL		DEFAULT -1,			--參數單位
 CncParamValue				nvarchar(10)		NOT NULL,						--參數預設值
 CncParamLevel				nvarchar(2)			NOT NULL,						--參數等級	>> M:主程式參數 ; S:副程式參數
 CncParamStatus				nvarchar(2)			NOT NULL,						--參數性質	>> P:編程參數 ; M:機台參數 ; T:工具參數(ex:大刀；小刀...) ; F:檔案參數
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼		>> S:關閉 ; A:啟用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_CncParameter  PRIMARY KEY (CncParamId),
 CONSTRAINT AK1_CncParameter UNIQUE (CncParamNo,CncProgId),
 CONSTRAINT FK1_CncParameter FOREIGN KEY (CncProgId) REFERENCES MES.CncProgram(CncProgId)
)

CREATE TABLE MES.CncParameterOption( /*編程程式參數選項資料*/
 CncParameterOptionId		int					identity,
 CncParamId					int					NOT NULL,						--編程程式參數ID
 ShowOption					nvarchar(50)		NOT NULL,						--參數選項Key
 OptionValue				nvarchar(10)		NOT NULL,						--參數選項Value
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼		   >> S:關閉 ; A:啟用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_CncParameterOption  PRIMARY KEY (CncParameterOptionId),
 CONSTRAINT AK1_CncParameterOption UNIQUE (CncParamId,ShowOption),
 CONSTRAINT FK1_CncParameterOption FOREIGN KEY (CncParamId) REFERENCES MES.CncParameter(CncParamId)
)

CREATE TABLE MES.CncProgramFile( /*編程程式檔案基本資料*/
 CncProgramFileId			int					identity,						--流水碼
 CncProgId					int					NOT NULL,						--編程程式ID
 FileDesc					nvarchar(100)		NOT NULL,						--檔案說明
 FileId						int					NOT NULL,						--檔案資料ID	>> BAS.[File]
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼		   >> S:關閉 ; A:啟用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_CncProgramFile PRIMARY KEY (CncProgramFileId),
 CONSTRAINT FK1_CncProgramFile FOREIGN KEY (CncProgId) REFERENCES MES.CncProgram(CncProgId),
 CONSTRAINT FK2_CncProgramFile FOREIGN KEY (FileId) REFERENCES BAS.[File](FileId)
)

CREATE TABLE MES.CncProgramLog( /*編程執行紀錄*/
 CncProgLogId				int					identity,						--流水碼
 CncProgId					int					NOT NULL,						--編程程式ID
 MoProcessId				int					NOT NULL,						--製令製程ID(MoProcess裡選擇) 
 MachineId					int					NOT NULL,						--機台ID
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼		   >> S:關閉 ; A:啟用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_CncProgramLog PRIMARY KEY (CncProgLogId),
 CONSTRAINT FK1_CncProgramLog FOREIGN KEY (CncProgId) REFERENCES MES.CncProgram(CncProgId),
 CONSTRAINT FK2_CncProgramLog FOREIGN KEY (MachineId) REFERENCES MES.Machine(MachineId)
)

CREATE TABLE MES.CncProgramRequest( /*編程需求紀錄*/
 CncProgRequestId			int					identity,						--流水碼
 CncProgLogId				int					NOT NULL,						--編程執行ID
 CncParamId					int					NOT NULL,						--編程參數ID
 RequestValues				nvarchar(300)		NOT NULL	DEFAULT 'N/A',		--編程參數值，如果參數輸入方式是檔案上傳，此欄位帶N/A
 FileId						int					NOT NULL	DEFAULT -1,			--上傳參數檔案資料ID	>> BAS.[File]
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼		   >> S:關閉 ; A:啟用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_CncProgramRequest PRIMARY KEY (CncProgRequestId), 
 CONSTRAINT FK1_CncProgramRequest FOREIGN KEY (CncProgLogId) REFERENCES MES.CncProgramLog(CncProgLogId),
 CONSTRAINT FK2_CncProgramRequest FOREIGN KEY (CncParamId) REFERENCES MES.CncParameter(CncParamId)
)

CREATE TABLE MES.CncProgramResponsest( /*編程回傳紀錄*/
 CncProgResponsestId		int					identity,						--流水碼
 CncProgLogId				int					NOT NULL,						--編程執行ID
 CncParamNo					nvarchar(500)		NOT NULL,						--回傳參數NO
 ResponsestValues			nvarchar(500)		NOT NULL	DEFAULT 'N/A',		--編程回傳值，如果回傳輸入方式是檔案上傳，此欄位帶N/A
 FileId						int					NOT NULL	DEFAULT -1,			--回傳檔案資料ID	>> BAS.[File]
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼		   >> S:關閉 ; A:啟用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_CncProgramResponsest PRIMARY KEY (CncProgResponsestId),
 CONSTRAINT FK1_CncProgramResponsest FOREIGN KEY (CncProgLogId) REFERENCES MES.CncProgramLog(CncProgLogId)
)

CREATE TABLE MES.CncResponsestFileLog( /*編程回傳紀錄Log*/
 CncResponsestFileLogId		int					identity,						--流水碼
 CncProgResponsestId		int					NOT NULL,						--編程回傳紀錄ID
 BeforeFileId				int					NOT NULL,						--原本File ID
 NewFileId					int					NOT NULL,						--現在File ID 
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼		   >> S:關閉 ; A:啟用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_CncResponsestFileLog PRIMARY KEY (CncResponsestFileLogId),
 CONSTRAINT FK1_CncResponsestFileLog FOREIGN KEY (CncProgResponsestId) REFERENCES MES.CncProgramResponsest(CncProgResponsestId),
 CONSTRAINT FK2_CncResponsestFileLog FOREIGN KEY (BeforeFileId) REFERENCES BAS.[File](FileId),
 CONSTRAINT FK3_CncResponsestFileLog FOREIGN KEY (NewFileId) REFERENCES BAS.[File](FileId)
)

CREATE TABLE MES.CncProgramWork( /*編程報工紀錄*/
 CncProgramWorkId			int					identity,						--流水碼
 CncProgLogId				int					NOT NULL,						--編程執行ID
 MoProcessId				int					NOT NULL,						--製令製程ID(MoProcess裡選擇)
 MachineId					int					NOT NULL,						--機台ID 
 StartWorkDate				datetime			NOT NULL	DEFAULT GETDATE(),	--開始編程日期
 EndWorkDate				datetime			NOT NULL,						--結束編程日期
 WorkCT						int					NOT NULL,						--編程花費時間(秒數)
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼		   >> S:關閉 ; A:啟用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_CncProgramWork PRIMARY KEY (CncProgramWorkId), 
 CONSTRAINT FK1_CncProgramWork FOREIGN KEY (CncProgLogId) REFERENCES MES.CncProgramLog(CncProgLogId),
 CONSTRAINT FK2_CncProgramWork FOREIGN KEY (MoProcessId) REFERENCES MES.MoProcess(MoProcessId),
 CONSTRAINT FK3_CncProgramWork FOREIGN KEY (MachineId) REFERENCES MES.Machine(MachineId)
)

CREATE TABLE MES.CncProgramWorkFiles( /*編程報工檔案上傳紀錄*/
 CncProgramWorkFileId		int					identity,						--流水碼
 CncProgramWorkId			int					NOT NULL,						--編程報工紀錄ID
 FileId						int					NOT NULL,						--檔案資料ID	>> BAS.[File]
 NcName						nvarchar(50)		NULL,							--編程名字 ex:O0004,O0021...etc
 FileType					nvarchar(2)			NOT NULL,						--檔案類型   SELECT * FROM BAS.[Type] WHERE TypeSchema = 'CncWorkFileType'  >> 1：編程檔案；2：其它程式
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼		   >> S:關閉 ; A:啟用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_CncProgramWorkFiles PRIMARY KEY (CncProgramWorkFileId),
 CONSTRAINT FK1_CncProgramWorkFiles FOREIGN KEY (CncProgramWorkId) REFERENCES MES.CncProgramWork(CncProgramWorkId),
 CONSTRAINT FK2_CncProgramWorkFiles FOREIGN KEY (FileId) REFERENCES BAS.[File](FileId)
)
END /*編程管理*/

BEGIN /*銑床編程管理*/
CREATE TABLE MES.MillPreEditLog( /*銑床預編執行紀錄*/
 MillPreLogId				int					identity,						--流水碼
 MoProcessId				int					NOT NULL,						--製令製程ID
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼		   >> S:關閉 ; A:啟用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MillPreEditLog PRIMARY KEY (MillPreLogId),
 CONSTRAINT FK1_MillPreEditLog FOREIGN KEY (MoProcessId) REFERENCES MES.MoProcess(MoProcessId)
)

CREATE TABLE MES.MillPreEditRequest( /*銑床預編需求紀錄*/
 MillPreRequestId			int					identity,						--流水碼
 MillPreLogId				int					NOT NULL,						--預編執行ID 
 DwgPrtData					nvarchar(500)		NOT NULL,						--設計圖.prt檔 
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼		   >> S:關閉 ; A:啟用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MillPreEditRequest PRIMARY KEY (MillPreRequestId), 
 CONSTRAINT FK1_MillPreEditRequest FOREIGN KEY (MillPreLogId) REFERENCES MES.MillPreEditLog(MillPreLogId)
)

CREATE TABLE MES.MillPreEditResponsest( /*銑床預編回傳紀錄*/
 MillPreResponsestId		int					identity,						--流水碼
 MillPreLogId				int					NOT NULL,						--預編執行ID 
 MillPrePrtData				nvarchar(500)		NOT NULL,						--預編檔案.prt檔 
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼		   >> S:關閉 ; A:啟用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MillPreEditResponsest PRIMARY KEY (MillPreResponsestId),
 CONSTRAINT FK1_MillPreEditResponsest FOREIGN KEY (MillPreLogId) REFERENCES MES.MillPreEditLog(MillPreLogId)
)

/*編程*/
CREATE TABLE MES.MillProgramLog( /*銑床編程執行紀錄*/
 MillProgLogId				int					identity,						--流水碼
 MillPreLogId				int					NOT NULL,						--銑床預編執行紀
 MachineId					int					NOT NULL,						--機台ID 
 JigId						int					NOT NULL,						--治具ID 
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼		   >> S:關閉 ; A:啟用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MillProgramLog PRIMARY KEY (MillProgLogId),
 CONSTRAINT FK1_MillProgramLog FOREIGN KEY (MachineId) REFERENCES MES.Machine(MachineId),
 CONSTRAINT FK2_MillProgramLog FOREIGN KEY (MillPreLogId) REFERENCES MES.MillPreEditLog(MillPreLogId),
 CONSTRAINT FK3_MillProgramLog FOREIGN KEY (JigId) REFERENCES MES.Jig(JigId)
)

CREATE TABLE MES.MillProgramRequest( /*銑床編程需求紀錄*/
 MillProgRequestId			int					identity,						--流水碼
 MillProgLogId				int					NOT NULL,						--編程執行ID 
 MillPrePrtData				nvarchar(500)		NOT NULL,						--加工預編.prt
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼		   >> S:關閉 ; A:啟用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MillProgramRequest PRIMARY KEY (MillProgRequestId), 
 CONSTRAINT FK1_MillProgramRequest FOREIGN KEY (MillProgLogId) REFERENCES MES.MillProgramLog(MillProgLogId)
)


CREATE TABLE MES.MillProgramResponsest( /*銑床編程回傳紀錄*/
 MillProgResponsestId		int					identity,						--流水碼
 MillProgLogId				int					NOT NULL,						--編程執行ID 
 MillPrtData				nvarchar(500)		NOT NULL,						--加工.nc
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼		   >> S:關閉 ; A:啟用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MillProgramResponsest PRIMARY KEY (MillProgResponsestId),
 CONSTRAINT FK1_MillProgramResponsest FOREIGN KEY (MillProgLogId) REFERENCES MES.MillProgramLog(MillProgLogId)
)

CREATE TABLE MES.MillProgramTool( /*銑床編程工具使用明細*/
 MillPreEditToolId			int					identity,						--流水碼
 MillProgLogId				int					NOT NULL,						--銑床編程執行ID 
 ToolBlockSeq				nvarchar(2)			NOT NULL,						--刀座編號
 ToolBlockName				nvarchar(20)		NOT NULL,						--刀座名稱
 MillToolD					nvarchar(20)		NOT NULL,						--刀具【直徑D】
 MillToolRone				nvarchar(20)		NOT NULL,						--刀具【下半徑R1】
 MillToolB					nvarchar(20)		NOT NULL,						--刀具【錐角B】
 MillToolFL					nvarchar(20)		NOT NULL,						--刀具【刀刃長度FL】
 MillToolL					nvarchar(20)		NOT NULL,						--刀具【總長度L】
 MillToolHD					nvarchar(20)		NOT NULL,						--刀具【刀柄直徑d】
 MillToolR					nvarchar(20)		NOT NULL,						--刀具【首角r】
 MillToolRPM				nvarchar(20)		NOT NULL,						--刀具【轉速RPM】
 MillToolFPR				nvarchar(20)		NOT NULL,						--刀具【進給值FPR】
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼		   >> S:關閉 ; A:啟用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MillProgramTool PRIMARY KEY (MillPreEditToolId),
 CONSTRAINT FK1_MillProgramTool FOREIGN KEY (MillProgLogId) REFERENCES MES.MillProgramLog(MillProgLogId)
)

CREATE TABLE MES.MillMachineTool( /*銑床機台工具紀錄*/
 MillMachineToolId			int					identity,						--流水碼
 ToolId						int					NOT NULL,						--工具ID 
 MachineId					int					NOT NULL,						--機台ID
 ToolBlockSeq				nvarchar(2)			NOT NULL,						--刀座編號
 ToolBlockName				nvarchar(20)		NOT NULL,						--刀座名稱
 MillToolD					nvarchar(20)		NOT NULL,						--刀具【直徑D】
 MillToolRone				nvarchar(20)		NOT NULL,						--刀具【下半徑R1】
 MillToolB					nvarchar(20)		NOT NULL,						--刀具【錐角B】
 MillToolFL					nvarchar(20)		NOT NULL,						--刀具【刀刃長度FL】
 MillToolL					nvarchar(20)		NOT NULL,						--刀具【總長度L】
 MillToolHD					nvarchar(20)		NOT NULL,						--刀具【刀柄直徑d】
 MillToolR					nvarchar(20)		NOT NULL,						--刀具【首角r】
 MillToolRPM				nvarchar(20)		NOT NULL,						--刀具【轉速RPM】
 MillToolFPR				nvarchar(20)		NOT NULL,						--刀具【進給值FPR】
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼		   >> S:關閉 ; A:啟用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MillMachineTool PRIMARY KEY (MillMachineToolId),
 CONSTRAINT FK1_MillMachineTool FOREIGN KEY (ToolId) REFERENCES MES.Tool(ToolId),
 CONSTRAINT FK2_MillMachineTool FOREIGN KEY (MachineId) REFERENCES MES.Machine(MachineId)
)

CREATE TABLE MES.MachineToolLocator( /*機台工具儲位維護*/
 MachineToolLocatorId		int					identity,						--流水碼
 MachineId					int					NOT NULL,						--機台ID 
 ToolLocatorId				int					NOT NULL,						--工具儲位ID 
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼		   >> S:關閉 ; A:啟用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MachineToolLocator PRIMARY KEY (MachineToolLocatorId),
 CONSTRAINT FK1_MachineToolLocator FOREIGN KEY (MachineId) REFERENCES MES.Machine(MachineId),
 CONSTRAINT FK2_MachineToolLocator FOREIGN KEY (ToolLocatorId) REFERENCES MES.ToolLocator(ToolLocatorId)
)

CREATE TABLE MES.MillProgramWork( /*銑床編程報工紀錄*/
 MillProgramWorkId			int					identity,						--流水碼
 MillProgLogId				int					NOT NULL,						--編程執行ID
 MoProcessId				int					NOT NULL,						--製令製程ID(MoProcess裡選擇)
 MachineId					int					NOT NULL,						--機台ID 
 StartWorkDate				datetime			NOT NULL	DEFAULT GETDATE(),	--開始編程日期
 EndWorkDate				datetime			NOT NULL,						--結束編程日期
 WorkCT						int					NOT NULL,						--編程花費時間(秒數)
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼		   >> S:關閉 ; A:啟用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MillProgramWork PRIMARY KEY (MillProgramWorkId), 
 CONSTRAINT FK1_MillProgramWork FOREIGN KEY (MillProgLogId) REFERENCES MES.MillProgramLog(MillProgLogId),
 CONSTRAINT FK2_MillProgramWork FOREIGN KEY (MoProcessId) REFERENCES MES.MoProcess(MoProcessId),
 CONSTRAINT FK3_MillProgramWork FOREIGN KEY (MachineId) REFERENCES MES.Machine(MachineId)
)

END /*銑床編程管理*/

BEGIN /*工具管理*/
CREATE TABLE MES.ToolGroup( /*工具群組基本資料*/
 ToolGroupId				int					identity,						--流水碼
 CompanyId					int					NOT NULL,						--所屬公司
 ToolGroupNo				nvarchar(50)		NOT NULL,						--工具群組編號
 ToolGroupName				nvarchar(50)		NOT NULL,						--工具群組名稱
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼		   >> S:關閉 ; A:啟用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ToolGroup PRIMARY KEY (ToolGroupId),
 CONSTRAINT AK1_ToolGroup UNIQUE (CompanyId,ToolGroupNo),
 CONSTRAINT FK1_ToolGroup FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)

CREATE TABLE MES.ToolSpec( /*工具規格基本資料*/
 ToolSpecId					int					identity,						--流水碼
 CompanyId					int					NOT NULL,						--所屬公司
 ToolClassId				int					NOT NULL,						--工具類別基本資料 ID
 ToolSpecNo					nvarchar(50)		NOT NULL,						--工具規格編號
 ToolSpecName				nvarchar(50)			NULL,						--工具規格名稱 
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼		   >> S:關閉 ; A:啟用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ToolSpec PRIMARY KEY (ToolSpecId),
 CONSTRAINT AK1_ToolSpec  UNIQUE (CompanyId,ToolSpecNo),
 CONSTRAINT FK1_ToolSpec  FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_ToolSpec  FOREIGN KEY (ToolClassId) REFERENCES MES.ToolClass(ToolClassId) 
)

CREATE TABLE MES.ToolClass( /*工具類別基本資料*/
 ToolClassId				int					identity,						--流水碼
 ToolGroupId				int					NOT NULL,						--工具群組ID
 ToolClassNo				nvarchar(50)		NOT NULL,						--工具類別編號
 ToolClassName				nvarchar(50)		NOT NULL,						--工具類別名稱
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼		   >> S:關閉 ; A:啟用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ToolClass PRIMARY KEY (ToolClassId),
 CONSTRAINT FK1_ToolClass FOREIGN KEY (ToolGroupId) REFERENCES MES.ToolGroup(ToolGroupId)
)

CREATE TABLE MES.ToolCategory( /*工具種類基本資料*/
 ToolCategoryId				int					identity,						--流水碼
 ToolClassId				int					NOT NULL,						--工具類別
 ToolCategoryNo				nvarchar(50)		NOT NULL,						--工具編號
 ToolCategoryName			nvarchar(50)		NOT NULL,						--工具名稱
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼		   >> S:關閉 ; A:啟用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ToolCategory PRIMARY KEY (ToolCategoryId),
 CONSTRAINT FK1_ToolCategory FOREIGN KEY (ToolClassId) REFERENCES MES.ToolClass(ToolClassId)
)

CREATE TABLE MES.ToolModel( /*工具型號*/
 ToolModelId				int					identity,						--流水碼
 ToolModelNo				nvarchar(50)		NOT NULL,						--工具型號
 ToolModelErpNo				nvarchar(50)			NULL,						--ERP型號
 ToolModelName				nvarchar(50)			NULL,						--工具品名
 ToolCategoryId				int					NOT NULL,						--工具種類ID
 SupplierId					int					NOT NULL	DEFAULT -1,			--供應商		>>SCM.Supplier
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼		>> S:關閉 ; A:啟用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ToolModel PRIMARY KEY (ToolModelId),
 CONSTRAINT FK1_ToolModel FOREIGN KEY (ToolCategoryId) REFERENCES MES.ToolCategory(ToolCategoryId)
)

CREATE TABLE MES.ToolModelSpec( /*工具型號規格*/
 ToolModelSpecId			int					identity,						--流水碼
 ToolModelId				int					NOT NULL,						--工具型號
 ToolSpecId					int					NOT NULL,						--工具規格ID
 ToolSpecValue				float				NOT NULL,						--規格數值
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼		>> S:關閉 ; A:啟用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ToolModelSpec PRIMARY KEY (ToolModelSpecId),
 CONSTRAINT FK1_ToolModelSpec FOREIGN KEY (ToolModelId) REFERENCES MES.ToolModel(ToolModelId),
 CONSTRAINT FK2_ToolModelSpec FOREIGN KEY (ToolSpecId) REFERENCES MES.ToolSpec(ToolSpecId)
)

CREATE TABLE MES.Tool( /*工具明細資訊*/
 ToolId						int					identity,						--流水碼
 ToolModelId				int					NOT NULL,						--工具型號ID
 RevToolCategoryId			int						NULL,						--變更後種類
 ToolNo						nvarchar(50)			NULL,						--工具編號
 UsageCount					int					NOT NULL    DEFAULT 0,			--使用次數
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'S',		--狀態碼		   >> S:關閉(未印出標籤) ; A:啟用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Tool PRIMARY KEY (ToolId),
 CONSTRAINT FK1_Tool FOREIGN KEY (ToolModelId) REFERENCES MES.ToolModel(ToolModelId)
)

CREATE TABLE MES.ToolSpecLog( /*工具對應的規格資訊*/
 ToolSpecLogId				int					identity,						--流水碼
 ToolSpecId					int					NOT NULL,						--工具規格編號
 ToolSpecValue				float				NULL,							--工具規格數值
 ToolId						int					NOT NULL,						--工具ID 
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼		   >> S:關閉 ; A:啟用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ToolSpecLog PRIMARY KEY (ToolSpecLogId),
 CONSTRAINT FK1_ToolSpecLog FOREIGN KEY (ToolSpecId) REFERENCES MES.ToolSpec(ToolSpecId),
 CONSTRAINT FK2_ToolSpecLog FOREIGN KEY (ToolId) REFERENCES MES.Tool(ToolId)
)

CREATE TABLE MES.ToolInventory(/*刀具倉庫*/
 ToolInventoryId			int					identity,						--流水碼
 ToolInventoryNo			nvarchar(50)		NOT NULL,						--工具倉庫編號
 ToolInventoryName			nvarchar(50)		NOT NULL,						--工具倉庫名稱
 ToolInventoryDesc			nvarchar(50)		NOT NULL,						--工具倉庫描述
 ShopId						int					NOT NULL,						--車間
 [Status]					NVARCHAR(2)			NOT NULL	DEFAULT 'A',		--A:啟用;S:停用 
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ToolInventory PRIMARY KEY(ToolInventoryId),
 CONSTRAINT AK1_ToolInventory UNIQUE(ToolInventoryNo),
 CONSTRAINT FK1_ToolInventory FOREIGN KEY (ShopId) REFERENCES MES.WorkShop(ShopId),
)

CREATE TABLE MES.ToolLocator(/*刀具儲位*/
 ToolLocatorId				int					identity,						--流水碼
 ToolLocatorNo				nvarchar(50)		NOT NULL,						--工具儲位編號
 ToolLocatorName			nvarchar(50)		NOT NULL,						--工具儲位名稱
 ToolLocatorDesc			nvarchar(50)		NOT NULL,						--工具儲位描述
 ToolInventoryId			int					NOT NULL,						--工具倉庫
 LimitStatus				nvarchar(2)			NULL		DEFAULT 'N',		--Y:上限數為1;N:無上限
 [Status]					nvarchar(2)			NOT NULL	DEFAULT 'A',		--A:啟用;S:停用 
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ToolLocator PRIMARY KEY(ToolLocatorId),
 CONSTRAINT AK1_ToolLocator UNIQUE(ToolLocatorNo),
 CONSTRAINT FK1_ToolLocator FOREIGN KEY (ToolInventoryId) REFERENCES MES.ToolInventory(ToolInventoryId),
)

CREATE TABLE MES.ToolTransactions(/*刀具異動交易紀錄*/
 ToolTransactionsId			int					identity,						--流水碼
 ToolId						int					NOT NULL,						--工具
 TransactionType			nvarchar(4)			NOT NULL,						--交易類型 In:入庫 Out:出庫
 TransactionDate			datetime			NOT NULL,						--交易日 
 ToolLocatorId				int					NOT NULL,						--工具儲位 
 TraderId					int					NOT NULL,						--交易者 BAS.[User] 
 TransactionReason			nvarchar(100)			NULL,						--工具交易原因
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ToolTransactions PRIMARY KEY(ToolTransactionsId),
 CONSTRAINT FK1_ToolTransactions FOREIGN KEY (ToolId) REFERENCES MES.Tool(ToolId),
 CONSTRAINT FK2_ToolTransactions FOREIGN KEY (ToolLocatorId) REFERENCES MES.ToolLocator(ToolLocatorId),
 CONSTRAINT FK3_ToolTransactions FOREIGN KEY (TraderId) REFERENCES BAS.[User](UserId)
)

CREATE TABLE MES.ToolClassIdSetting( /*不同類別工具設定參數種類基本資料*/
 ToolClassIdSettingId		int					identity,						--流水碼 
 ToolClassId				int					NOT NULL,						--工具Class ID
 ParameterName				nvarchar(50)		NOT NULL,						--編程參數名
 ParameterKey				nvarchar(50)		NOT NULL,						--編程參數Key 
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼		   >> S:關閉 ; A:啟用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ToolClassIdSetting PRIMARY KEY (ToolClassIdSettingId),
 CONSTRAINT AK1_ToolClassIdSetting UNIQUE (ToolClassId,ParameterKey,ParameterName),
 CONSTRAINT FK1_ToolClassIdSetting FOREIGN KEY (ToolClassId) REFERENCES MES.ToolClass(ToolClassId)
)

CREATE TABLE MES.ToolMachine( /*機台對應工具綁定資訊*/
 ToolMachineId				int					identity,						--流水碼
 MachineId					int					NOT NULL,						--機台ID
 ToolId						int					NOT NULL,						--工具ID
 MachineLocation			int					NOT NULL    DEFAULT 1,			--座台編號(ex:刀座...)
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼		   >> S:關閉 ; A:啟用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ToolMachine PRIMARY KEY (ToolMachineId),
 CONSTRAINT AK1_ToolMachine UNIQUE (MachineId,ToolId,MachineLocation),
 CONSTRAINT FK1_ToolMachine FOREIGN KEY (MachineId) REFERENCES MES.Machine(MachineId),
 CONSTRAINT FK2_ToolMachine FOREIGN KEY (ToolId) REFERENCES MES.Tool(ToolId)
)

CREATE TABLE MES.ToolMachineParameter( /*工具在機台設定參數資料*/
 ToolMachineParameterId		int					identity,
 ToolMachineId				int					NOT NULL,						--機台對應工具設定參數ID
 ToolClassIdSettingId		int					NOT NULL,						--不同類別工具設定參數種類ID 
 ParameterValue				nvarchar(10)		NOT NULL,						--編程參數值 
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼		   >> S:關閉 ; A:啟用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ToolMachineParameter PRIMARY KEY (ToolMachineParameterId),
 CONSTRAINT FK1_ToolMachineParameter FOREIGN KEY (ToolMachineId) REFERENCES MES.ToolMachine(ToolMachineId),
 CONSTRAINT FK2_ToolMachineParameter FOREIGN KEY (ToolClassIdSettingId) REFERENCES MES.ToolClassIdSetting(ToolClassIdSettingId)
)
END /*工具管理*/

BEGIN /*自動手臂使用*/
CREATE TABLE MES.RoboticArmProcessExt( /* 自動手臂過站暫存資料 */
 RapExtId	 				int					identity,
 JigId						int					NOT NULL,						--治具ID
 BarcodeNo					nvarchar(60)		NOT NULL,						--條碼
 MoProcessId				int					NOT NULL,						--製令製程ID
 WorkingFlag				nvarchar(1)			NOT NULL	DEFAULT 'N',		--條碼完工狀態 N:未開工 A:加工中 F:已完工
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_JigBarcode PRIMARY KEY(JigBarcodeId),
 CONSTRAINT FK1_JigBarcode FOREIGN KEY (JigId) REFERENCES MES.Jig(JigId),
 CONSTRAINT FK2_JigBarcode FOREIGN KEY (MoProcessId) REFERENCES MES.MoProcess(MoProcessId),
)/*自動手臂過站暫存資料*/

END/*自動手臂使用*/

BEGIN /*製程參數/途程量測項目設定*/
CREATE TABLE MES.ProcessParameterQcItem( /* 製程參數量測項目 */
 ParameterQcItemId			int					identity,
 ParameterId				int					NOT NULL,						--製程參數ID
 QcItemId					int					NOT NULL,						--量測項目Id
 QcItemDesc					nvarchar(100)			NULL,						--量測項目描述
 DesignValue				float				    NULL,						--尺寸設計值
 UpperTolerance				float				    NULL,						--尺寸設計值上限公差
 LowerTolerance				float				    NULL,						--尺寸設計值下限公差
 BallMark 					nvarchar(100)			NULL,						--球標
 Unit 						nvarchar(100)			NULL,						--單位
 QmmDetailId	 			int						NULL,						--量測設備
 Remark						nvarchar(255)			NULL,						--檢測備註
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ProcessParameterQcItem PRIMARY KEY(ParameterQcItemId),
 CONSTRAINT FK1_ProcessParameterQcItem FOREIGN KEY (ParameterId) REFERENCES MES.ProcessParameter(ParameterId),
 CONSTRAINT FK2_ProcessParameterQcItem FOREIGN KEY (QcItemId) REFERENCES QMS.QcItem(QcItemId)

 )/*製程參數量測項目*/

CREATE TABLE MES.RoutingItemQcItem( /* 途程量測項目 */
 RoutingItemQcItemId		int					identity,
 ItemProcessId				int					NOT NULL,						--製程ID
 RoutingItemId				int					NOT NULL,						--途程ID
 QcItemId					int					NOT NULL,						--量測項目Id
 QcItemDesc					nvarchar(100)			NULL,						--量測項目描述
 DesignValue				float				    NULL,						--尺寸設計值
 UpperTolerance				float				    NULL,						--尺寸設計值上限公差
 LowerTolerance				float				    NULL,						--尺寸設計值下限公差
 BallMark 					nvarchar(100)			NULL,						--球標
 Unit 						nvarchar(100)			NULL,						--單位
 QmmDetailId	 			int						NULL,						--量測設備
 Remark						nvarchar(255)			NULL,						--檢測備註
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_RoutingItemQcItem PRIMARY KEY(RoutingItemQcItemId),
 CONSTRAINT FK1_RoutingItemQcItem FOREIGN KEY (RoutingItemId) REFERENCES MES.RoutingItem(RoutingItemId),
 CONSTRAINT FK2_RoutingItemQcItem FOREIGN KEY (QcItemId) REFERENCES QMS.QcItem(QcItemId),
 CONSTRAINT FK3_RoutingItemQcItem FOREIGN KEY (ItemProcessId) REFERENCES MES.RoutingItemProcess(ItemProcessId)
)/*途程量測項目*/

END/*製程參數/途程量測項目設定*/

BEGIN /*Tray物料管理*/
CREATE TABLE MES.TrayBarcodeDismantle( /* Tray Barcode 拆解紀錄表 */
 TbdId						int					identity,
 BarcodeId					int					NOT NULL,						--條碼ID
 Quantity					float				NOT	NULL	DEFAULT 0,			--已拆數量
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_TrayBarcodeDismantle PRIMARY KEY(TbdId),
 CONSTRAINT FK1_TrayBarcodeDismantle FOREIGN KEY (BarcodeId) REFERENCES MES.Barcode(BarcodeId)
)/*Tray Barcode 拆解紀錄表*/

CREATE TABLE MES.TrayMtlItemNg( /* Tray盤物料NG數 領料單核單時修改*/
 TrayMtlItemNgId			int					identity,
 MoId						int					NOT NULL,						--製令ID
 BomDetailId				int					NOT NULL,						--物料品號ID
 NgQty						float				NOT	NULL	DEFAULT 0,			--總NG數
 AlreadyQty					int					NOT	NULL	DEFAULT 0,			--領料單已綁定數量 
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_TrayMtlItemNg PRIMARY KEY(TrayMtlItemNgId),
 CONSTRAINT AK_TrayMtlItemNg UNIQUE(MoId, BomDetailId),
 CONSTRAINT FK1_TrayMtlItemNg FOREIGN KEY (MoId) REFERENCES MES.ManufactureOrder(MoId),
 CONSTRAINT FK2_TrayMtlItemNg FOREIGN KEY (BomDetailId) REFERENCES PDM.BomDetail(BomDetailId)
)/*Tray盤物料NG數*/

CREATE TABLE MES.TrayMtlItemNgLog( /* Tray盤物料NG數 Log */
 LogId						int					identity,
 TrayMtlItemNgId			int					NOT NULL,						--Tray盤物料NG數ID
 MrDetailId					int					NOT NULL,						--領料單身ID
 [Action]					nvarchar			NOT NULL,						--執行動作 Y:核單 N:反確認
 Quantity					int					NOT NULL,						--數量
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_TrayMtlItemNgLog PRIMARY KEY(LogId)
)/*Tray盤物料NG數 Log*/

CREATE TABLE MES.MachineConsumeLog( /* 機台耗材使用紀錄 */
 MachineConsumeLogId		int					identity,
 MachineId  				int					NOT NULL,						--機台ID
 GrBarcodeId 				int					NOT NULL,						--物料條碼ID
 LotNumberId 				int					NOT NULL,						--批號條碼ID
 MaxDispenseQty 			int						NULL,						--可用次數
 AllowDispenseQty			int						NULL,						--已用次數
 OpeningDate				datetime				NULL,						--開封日期
 EXPs						int						NULL,						--使用期限/天數
 BindDate 					datetime				NULL,						--綁定機台日期
 UnBindDate					datetime				NULL,						--耗材更換日期
 CreateDate					datetime			NOT NULL,						--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MachineConsumeLog PRIMARY KEY(MachineConsumeLogId),
 CONSTRAINT FK1_MachineConsumeLog FOREIGN KEY (MachineId) REFERENCES MES.Machine(MachineId)
)/*機台耗材使用紀錄*/

CREATE TABLE MES.BarcodeConsumeLog( /* 條碼耗材使用紀錄 */
 BarcodeConsumeLogId		int					identity,
 BarcodeProcessId			int					NOT NULL,						--條碼過站ID
 MachineConsumeLogId		int					NOT NULL,				--批號物料條碼ID
 AllowDispenseQty			int						NULL,						--耗材已用次數,BarcodeQty條碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_BarcodeConsumeLog PRIMARY KEY(BarcodeConsumeLogId),
 CONSTRAINT FK1_BarcodeConsumeLog FOREIGN KEY (BarcodeProcessId) REFERENCES MES.BarcodeProcess(BarcodeProcessId),
 CONSTRAINT FK2_BarcodeConsumeLog FOREIGN KEY (MachineConsumeLogId) REFERENCES MES.MachineConsumeLog(MachineConsumeLogId)

)/*條碼耗材使用紀錄*/

END/*Tray物料管理*/

BEGIN /*MTF暫存資料管理*/
CREATE TABLE MES.MtfProcessTemp( /* MTF量測參數暫存 */
 MtfProcessTempId			int					identity,
 BarcodeProcessId			int					NOT NULL,						--條碼過站ID
 MachineId					int					NOT NULL,						--機台ID
 StartDate					datetime			NULL,							--MTF開工時間
 FinishDate					datetime			NULL,							--MTF完工時間
 RunStatus					nvarchar(1)			NOT NULL 	DEFAULT 'N',		--執行狀態 Y:已執行 N:未執行
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MtfProcessTemp PRIMARY KEY(MtfProcessTempId),
 CONSTRAINT FK1_MtfProcessTemp FOREIGN KEY (BarcodeProcessId) REFERENCES MES.BarcodeProcess(BarcodeProcessId),
 CONSTRAINT FK2_MtfProcessTemp FOREIGN KEY (MachineId) REFERENCES  MES.Machine(MachineId)
)/* MTF量測參數暫存 */

END/*MTF暫存資料管理*/

BEGIN /*各機種對應Tray盤數資料管理*/
CREATE TABLE MES.ItemTrayCount( /* 各機種對應Tray盤數資料管理 */
 ItemTrayCountId			int					identity,
 MtlItemId					int					NOT NULL,						--品號ID
 TrayCount					int					NOT NULL,						--Tray盤數量
 BatchCount					int					NOT NULL,						--批次數量
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ItemTrayCount PRIMARY KEY(ItemTrayCountId),
 CONSTRAINT FK1_ItemTrayCount FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId)
)/* 各機種對應Tray盤數資料管理 */

END/*各機種對應Tray盤數資料管理*/

BEGIN /*動態報表生成系統*/

CREATE TABLE MES.ReportDefinition -- 報表基本信息表
(
    ReportId INT IDENTITY(1,1) PRIMARY KEY,
    ReportName NVARCHAR(100) NOT NULL,
    ReportDescription NVARCHAR(500),
    ReportCategory NVARCHAR(50),
    SqlStatement NTEXT NOT NULL,
    OrderByDefault NVARCHAR(100),
    IsActive BIT DEFAULT 1,
    AllowExport BIT DEFAULT 1,
    CreatedBy NVARCHAR(50),
    CreatedDate DATETIME DEFAULT GETDATE(),
    UpdatedBy NVARCHAR(50),
    UpdatedDate DATETIME
)

-- 報表欄位定義表
CREATE TABLE MES.ReportColumn
(
    ColumnId INT IDENTITY(1,1) PRIMARY KEY,
    ReportId INT FOREIGN KEY REFERENCES MES.ReportDefinition(ReportId),
    ColumnName NVARCHAR(100) NOT NULL,
    DisplayName NVARCHAR(100) NOT NULL,
    DataType NVARCHAR(20) NOT NULL,
    FormatString NVARCHAR(50),
    ColumnWidth NVARCHAR(20),
    ColumnAlign NVARCHAR(10) DEFAULT 'left',
    SortOrder INT,
    IsVisible BIT DEFAULT 1
)

-- 報表查詢條件定義表
CREATE TABLE MES.ReportParameter
(
    ParameterId INT IDENTITY(1,1) PRIMARY KEY,
    ReportId INT FOREIGN KEY REFERENCES MES.ReportDefinition(ReportId),
    ParameterName NVARCHAR(100) NOT NULL,
    DisplayName NVARCHAR(100) NOT NULL,
    ControlType NVARCHAR(20) NOT NULL,    -- Text, Select, Date, etc.
    DefaultValue NVARCHAR(100),
    DataSource NVARCHAR(500),    -- SQL for dropdown options
    ParentParameter INT,    -- For cascading parameters
    SortOrder INT,
    IsRequired BIT DEFAULT 0
)

END/*動態報表生成系統*/

-- 工序預計日期主表
CREATE TABLE MES.MoProcessExpectedDate (
    Id INT PRIMARY KEY IDENTITY(1,1),
    MoId VARCHAR(50) NOT NULL,            -- 製令ID
    ProcessId INT NOT NULL,               -- 工序ID
    ExpectedStartDate DATETIME NULL,      -- 預計開工日期
    ExpectedEndDate DATETIME NULL,        -- 預計完工日期
    CreateUser VARCHAR(50) NOT NULL,      -- 建立人員
    CreateTime DATETIME NOT NULL DEFAULT GETDATE(), -- 建立時間
    LastUpdateUser VARCHAR(50) NOT NULL,  -- 最後更新人員
    LastUpdateTime DATETIME NOT NULL DEFAULT GETDATE(), -- 最後更新時間    
    -- 複合唯一索引確保每個工序只有一筆資料
    CONSTRAINT UQ_MoProcessExpectedDate_MoProcess UNIQUE (MoId, ProcessId)
);

-- 變更紀錄表
CREATE TABLE MES.MoProcessExpectedDateLog (
    Id INT PRIMARY KEY IDENTITY(1,1),
    MoId VARCHAR(50) NOT NULL,             -- 製令ID
    ProcessId INT NOT NULL,                -- 工序ID
    ChangeType VARCHAR(20) NOT NULL,       -- 變更類型: 'Start', 'End', 'Both'
    OldStartDate DATETIME NULL,            -- 變更前預計開工日
    NewStartDate DATETIME NULL,            -- 變更後預計開工日
    OldEndDate DATETIME NULL,              -- 變更前預計完工日
    NewEndDate DATETIME NULL,              -- 變更後預計完工日
    ChangeReason NVARCHAR(500) NULL,       -- 變更原因
    CreateUser VARCHAR(50) NOT NULL,       -- 變更人員
    CreateTime DATETIME NOT NULL DEFAULT GETDATE() -- 變更時間
);

-- 索引加速查詢
CREATE INDEX IX_MoProcessExpectedDate_MoId ON MES.MoProcessExpectedDate (MoId);
CREATE INDEX IX_MoProcessExpectedDateLog_MoId ON MES.MoProcessExpectedDateLog (MoId);
CREATE INDEX IX_MoProcessExpectedDateLog_CreateTime ON MES.MoProcessExpectedDateLog (CreateTime);
