BEGIN /*基本資料*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'RoleUser') > 0 DROP TABLE SSO.RoleUser
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DemandRole') > 0 DROP TABLE SSO.DemandRole
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DemandFlow') > 0 DROP TABLE SSO.DemandFlow
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DemandSource') > 0 DROP TABLE SSO.DemandSource
END /*基本資料*/

BEGIN /*需求追蹤平台*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DemandNotificationLogDetail') > 0 DROP TABLE SSO.DemandNotificationLogDetail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DemandNotificationLog') > 0 DROP TABLE SSO.DemandNotificationLog
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DemandFlowLog') > 0 DROP TABLE SSO.DemandFlowLog
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DemandFlowUser') > 0 DROP TABLE SSO.DemandFlowUser
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DemandFlowLink') > 0 DROP TABLE SSO.DemandFlowLink
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DemandFlowSetting') > 0 DROP TABLE SSO.DemandFlowSetting
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DemandSalesShipment') > 0 DROP TABLE SSO.DemandSalesShipment
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DemandTempShippingNote') > 0 DROP TABLE SSO.DemandTempShippingNote
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DemandProductionStockInNote') > 0 DROP TABLE SSO.DemandProductionStockInNote
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DemandMaterialRequisition') > 0 DROP TABLE SSO.DemandMaterialRequisition
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DemandWipOrder') > 0 DROP TABLE SSO.DemandWipOrder
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DemandPurchaseRequisition') > 0 DROP TABLE SSO.DemandPurchaseRequisition
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DemandRouting') > 0 DROP TABLE SSO.DemandRouting
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DemandSalesOrder') > 0 DROP TABLE SSO.DemandSalesOrder
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DemandBom') > 0 DROP TABLE SSO.DemandBom
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DemandMtlItem') > 0 DROP TABLE SSO.DemandMtlItem
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DemandCertificate') > 0 DROP TABLE SSO.DemandCertificate
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Demand') > 0 DROP TABLE SSO.Demand
END /*需求追蹤平台*/

BEGIN /*系統設定*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'SourceFlowUser') > 0 DROP TABLE SSO.SourceFlowUser
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'SourceFlowLink') > 0 DROP TABLE SSO.SourceFlowLink
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'SourceFlowSetting') > 0 DROP TABLE SSO.SourceFlowSetting
END /*系統設定*/

--******************************************************************************************************************
BEGIN /*基本資料*/
CREATE TABLE SSO.DemandSource( /*需求來源*/ --EC：電商(E-Commerce)；EM：Email；SO：一般訂單(Sales Order)；OEM：委託加工(OEM)；SF：銷售預測(Sales Forecast)；SP：策略性生產(Strategic Production)
 SourceId					int					identity,
 CompanyId					int					NOT NULL,						--所屬公司ID
 SourceNo					nvarchar(10)		NOT NULL,						--來源代碼
 SourceName					nvarchar(50)		NOT NULL,						--來源名稱
 SourceDesc					nvarchar(50)		NOT NULL,						--來源描述
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DemandSource PRIMARY KEY (SourceId),
 CONSTRAINT AK1_DemandSource UNIQUE (SourceNo),
 CONSTRAINT FK1_DemandSource FOREIGN KEY (CreateBy) REFERENCES BAS.UserInfo(UserId),
 CONSTRAINT FK2_DemandSource FOREIGN KEY (LastModifiedBy) REFERENCES BAS.UserInfo(UserId),
 CONSTRAINT FK3_DemandSource FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
) /*需求來源*/

CREATE TABLE SSO.DemandFlow( /*流程資料*/
 FlowId						int					identity,
 FlowNo						nvarchar(50)		NOT NULL,						--流程代號
 FlowName					nvarchar(50)		NOT NULL,						--流程名稱
 FlowImage					nvarchar(MAX)			NULL,						--流程圖片
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DemandFlow PRIMARY KEY (FlowId),
 CONSTRAINT AK1_DemandFlow UNIQUE (FlowNo),
 CONSTRAINT AK2_DemandFlow UNIQUE (FlowName),
 CONSTRAINT FK1_DemandFlow FOREIGN KEY (CreateBy) REFERENCES BAS.UserInfo(UserId),
 CONSTRAINT FK2_DemandFlow FOREIGN KEY (LastModifiedBy) REFERENCES BAS.UserInfo(UserId)
) /*流程資料*/

CREATE TABLE SSO.DemandRole( /*角色資料*/
 RoleId						int					identity,
 CompanyId					int					NOT NULL,						--所屬公司ID
 RoleName					nvarchar(50)		NOT NULL,						--角色名稱
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DemandRole PRIMARY KEY (RoleId),
 CONSTRAINT AK1_DemandRole UNIQUE (RoleName),
 CONSTRAINT FK1_DemandRole FOREIGN KEY (CreateBy) REFERENCES BAS.UserInfo(UserId),
 CONSTRAINT FK2_DemandRole FOREIGN KEY (LastModifiedBy) REFERENCES BAS.UserInfo(UserId),
 CONSTRAINT FK3_DemandRole FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
) /*角色資料*/

CREATE TABLE SSO.RoleUser( /*角色使用者*/
 RoleId						int					NOT NULL,						--所屬角色Id
 UserId						int					NOT NULL,						--所屬使用者Id
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 CreateBy					int					NOT NULL,						--新增人員
 CONSTRAINT PK_RoleUser PRIMARY KEY (RoleId, UserId),
 CONSTRAINT FK1_RoleUser FOREIGN KEY (CreateBy) REFERENCES BAS.UserInfo(UserId),
 CONSTRAINT FK2_RoleUser FOREIGN KEY (RoleId) REFERENCES SSO.DemandRole(RoleId),
 CONSTRAINT FK3_RoleUser FOREIGN KEY (UserId) REFERENCES BAS.UserInfo(UserId)
) /*角色使用者*/
END /*基本資料*/

BEGIN /*需求追蹤平台*/
CREATE TABLE SSO.Demand( /*需求資料*/
 DemandId					int					identity,
 SourceId					int					NOT NULL,						--所屬需求來源ID
 DemandNo					nvarchar(10)		NOT NULL,						--需求單號
 DemandDesc					nvarchar(50)		NOT NULL,						--需求描述
 DemandDate					datetime			NOT NULL,						--需求日期
 DemandDeadline				datetime				NULL,						--需求截止日期
 DemandDepartment			int						NULL,						--需求部門
 DemandCustomer				int						NULL,						--需求客戶
 DemandUser					int					NOT NULL,						--需求人員
 ItemTypeId					int					NOT NULL,						--需求產品類別
 StartDate					datetime				NULL,						--開始時間
 EndDate					datetime				NULL,						--結束時間
 DemandStatus				nvarchar(2)			NOT NULL,						--需求狀態 >> BAS.[Status] P：需求規劃(Planning)；I：需求進行中(In progress)；C：需求完成(Completed)				
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Demand PRIMARY KEY (DemandId),
 CONSTRAINT AK1_Demand UNIQUE (DemandNo),
 CONSTRAINT FK1_Demand FOREIGN KEY (CreateBy) REFERENCES BAS.UserInfo(UserId),
 CONSTRAINT FK2_Demand FOREIGN KEY (LastModifiedBy) REFERENCES BAS.UserInfo(UserId),
 CONSTRAINT FK3_Demand FOREIGN KEY (SourceId) REFERENCES SSO.DemandSource(SourceId),
 CONSTRAINT FK4_Demand FOREIGN KEY (DemandDepartment) REFERENCES BAS.Department(DepartmentId),
 CONSTRAINT FK5_Demand FOREIGN KEY (DemandCustomer) REFERENCES ERP.Customer(CustomerId),
 CONSTRAINT FK6_Demand FOREIGN KEY (DemandUser) REFERENCES BAS.UserInfo(UserId),
 CONSTRAINT FK7_Demand FOREIGN KEY (ItemTypeId) REFERENCES PDM.ItemType(ItemTypeId)
) /*需求資料*/

CREATE TABLE SSO.DemandCertificate( /*需求憑證*/
 CertificateId				int					identity,
 DemandId					int					NOT NULL,						--所屬需求ID
 CertificateDesc			nvarchar(50)		NOT NULL,						--憑證說明
 FileId						int					NOT NULL,						--所屬檔案ID
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DemandCertificate PRIMARY KEY (CertificateId),
 CONSTRAINT FK1_DemandCertificate FOREIGN KEY (DemandId) REFERENCES SSO.Demand(DemandId),
 CONSTRAINT FK2_DemandCertificate FOREIGN KEY (FileId) REFERENCES BAS.[File](FileId)
) /*需求憑證*/

CREATE TABLE SSO.DemandMtlItem( /*需求相關品號*/
 DemandId					int					NOT NULL,						--所屬需求ID
 MtlItemNo					nvarchar(40)		NOT NULL,						--品號
 MainItem					nvarchar(2)			NOT NULL,						--是否是主項目 >> BAS.[Status] Y：開立訂單品號；N：相關品號
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 CreateBy					int					NOT NULL,						--新增人員
 CONSTRAINT AK1_DemandMtlItem UNIQUE (DemandId, MtlItemNo),
 CONSTRAINT FK1_DemandMtlItem FOREIGN KEY (DemandId) REFERENCES SSO.Demand(DemandId)
) /*需求相關品號*/

CREATE TABLE SSO.DemandBom( /*需求相關BOM*/
 DemandId					int					NOT NULL,						--所屬需求ID
 MtlItemNo					nvarchar(40)		NOT NULL,						--主件品號
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 CreateBy					int					NOT NULL,						--新增人員
 CONSTRAINT AK1_DemandBom UNIQUE (DemandId, MtlItemNo),
 CONSTRAINT FK1_DemandBom FOREIGN KEY (DemandId) REFERENCES SSO.Demand(DemandId)
) /*需求相關品號*/

CREATE TABLE SSO.DemandSalesOrder( /*需求相關訂單*/
 DemandId					int					NOT NULL,						--所屬需求ID
 SoErpPrefix				nvarchar(4)			NOT NULL,						--訂單單別
 SoErpNo					nvarchar(11)		NOT NULL,						--訂單單號
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 CreateBy					int					NOT NULL,						--新增人員
 CONSTRAINT AK1_DemandSalesOrder UNIQUE (DemandId, SoErpPrefix, SoErpNo),
 CONSTRAINT FK1_DemandSalesOrder FOREIGN KEY (DemandId) REFERENCES SSO.Demand(DemandId)
) /*需求相關訂單*/

CREATE TABLE SSO.DemandRouting( /*需求相關途程*/
 DemandId					int					NOT NULL,						--所屬需求ID
 RoutingId					int					NOT NULL,						--所屬途程ID
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 CreateBy					int					NOT NULL,						--新增人員
 CONSTRAINT AK1_DemandRouting UNIQUE (DemandId, RoutingId),
 CONSTRAINT FK1_DemandRouting FOREIGN KEY (DemandId) REFERENCES SSO.Demand(DemandId),
 CONSTRAINT FK2_DemandRouting FOREIGN KEY (RoutingId) REFERENCES MES.Routing(RoutingId),
) /*需求相關途程*/

CREATE TABLE SSO.DemandPurchaseRequisition( /*需求相關請購單*/
 DemandId					int					NOT NULL,						--所屬需求ID
 PrErpPrefix				nvarchar(4)			NOT NULL,						--請購單別
 PrErpNo					nvarchar(11)		NOT NULL,						--請購單號
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 CreateBy					int					NOT NULL,						--新增人員
 CONSTRAINT AK1_DemandPurchaseRequisition UNIQUE (DemandId, PrErpPrefix, PrErpNo),
 CONSTRAINT FK1_DemandPurchaseRequisition FOREIGN KEY (DemandId) REFERENCES SSO.Demand(DemandId)
) /*需求相關請購單*/

CREATE TABLE SSO.DemandWipOrder( /*需求相關製令*/
 DemandId					int					NOT NULL,						--所屬需求ID
 WoErpPrefix				nvarchar(4)			NOT NULL,						--製令單別
 WoErpNo					nvarchar(11)		NOT NULL,						--製令單號
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 CreateBy					int					NOT NULL,						--新增人員
 CONSTRAINT AK1_DemandWipOrder UNIQUE (DemandId, WoErpPrefix, WoErpNo),
 CONSTRAINT FK1_DemandWipOrder FOREIGN KEY (DemandId) REFERENCES SSO.Demand(DemandId)
) /*需求相關製令*/

CREATE TABLE SSO.DemandMaterialRequisition( /*需求相關領料單*/
 DemandId					int					NOT NULL,						--所屬需求ID
 MrErpPrefix				nvarchar(4)			NOT NULL,						--領料單別
 MrErpNo					nvarchar(11)		NOT NULL,						--領料單號
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 CreateBy					int					NOT NULL,						--新增人員
 CONSTRAINT AK1_DemandMaterialRequisition UNIQUE (DemandId, MrErpPrefix, MrErpNo),
 CONSTRAINT FK1_DemandMaterialRequisition FOREIGN KEY (DemandId) REFERENCES SSO.Demand(DemandId)
) /*需求相關領料單*/

CREATE TABLE SSO.DemandProductionStockInNote( /*需求相關入庫單*/
 DemandId					int					NOT NULL,						--所屬需求ID
 PsinErpPrefix				nvarchar(4)			NOT NULL,						--入庫單別
 PsinErpNo					nvarchar(11)		NOT NULL,						--入庫單號
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 CreateBy					int					NOT NULL,						--新增人員
 CONSTRAINT AK1_DemandProductionStockInNote UNIQUE (DemandId, PsinErpPrefix, PsinErpNo),
 CONSTRAINT FK1_DemandProductionStockInNote FOREIGN KEY (DemandId) REFERENCES SSO.Demand(DemandId)
) /*需求相關入庫單*/

CREATE TABLE SSO.DemandTempShippingNote( /*需求相關暫出單*/
 DemandId					int					NOT NULL,						--所屬需求ID
 TsnErpPrefix				nvarchar(4)			NOT NULL,						--入庫單別
 TsnErpNo					nvarchar(11)		NOT NULL,						--入庫單號
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 CreateBy					int					NOT NULL,						--新增人員
 CONSTRAINT AK1_DemandTempShippingNote UNIQUE (DemandId, TsnErpPrefix, TsnErpNo),
 CONSTRAINT FK1_DemandTempShippingNote FOREIGN KEY (DemandId) REFERENCES SSO.Demand(DemandId)
) /*需求相關暫出單*/

CREATE TABLE SSO.DemandSalesShipment( /*需求相關銷貨單*/
 DemandId					int					NOT NULL,						--所屬需求ID
 SsErpPrefix				nvarchar(4)			NOT NULL,						--銷貨單別
 SsErpNo					nvarchar(11)		NOT NULL,						--銷貨單號
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 CreateBy					int					NOT NULL,						--新增人員
 CONSTRAINT AK1_DemandSalesShipment UNIQUE (DemandId, SsErpPrefix, SsErpNo),
 CONSTRAINT FK1_DemandSalesShipment FOREIGN KEY (DemandId) REFERENCES SSO.Demand(DemandId)
) /*需求相關銷貨單*/

CREATE TABLE SSO.DemandFlowSetting( /*需求流程設定*/
 SettingId					int					identity,
 DemandId					int					NOT NULL,						--所屬需求Id
 FlowId						int					NOT NULL,						--所屬流程Id
 Xaxis						int						NULL,						--X座標
 Yaxis						int						NULL,						--Y座標
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DemandFlowSetting PRIMARY KEY (SettingId),
 CONSTRAINT FK1_DemandFlowSetting FOREIGN KEY (DemandId) REFERENCES SSO.Demand(DemandId),
) /*需求流程設定*/

CREATE TABLE SSO.DemandFlowLink( /*需求流程設定連結*/
 SourceSettingId			int					NOT NULL,						--來源流程設定
 TargetSettingId			int					NOT NULL,						--目標流程設定
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT AK1_DemandFlowLink UNIQUE (SourceSettingId, TargetSettingId),
 CONSTRAINT FK1_DemandFlowLink FOREIGN KEY (SourceSettingId) REFERENCES SSO.DemandFlowSetting(SettingId),
 CONSTRAINT FK2_DemandFlowLink FOREIGN KEY (TargetSettingId) REFERENCES SSO.DemandFlowSetting(SettingId)
) /*需求流程設定連結*/

CREATE TABLE SSO.DemandFlowUser( /*需求流程使用者對應*/
 SettingId					int					NOT NULL,						--所屬設定Id
 RoleId						int						NULL,						--所屬角色Id
 UserId						int						NULL,						--所屬使用者Id
 MailAdviceStatus			nvarchar(2)			NOT NULL,						--信件通知
 PushAdviceStatus			nvarchar(2)			NOT NULL,						--推播通知
 WorkWeixinStatus			nvarchar(2)			NOT NULL,						--企業微信通知
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT AK1_DemandFlowUser UNIQUE (SettingId, RoleId, UserId),
 CONSTRAINT FK1_DemandFlowUser FOREIGN KEY (SettingId) REFERENCES SSO.DemandFlowSetting(SettingId),
 CONSTRAINT FK2_DemandFlowUser FOREIGN KEY (RoleId) REFERENCES SSO.DemandRole(RoleId),
 CONSTRAINT FK3_DemandFlowUser FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId)
) /*需求流程使用者對應*/

CREATE TABLE SSO.DemandFlowLog( /*需求流程狀態歷程*/
 LogId						int					identity,
 SettingId					int					NOT NULL,						--所屬需求Id
 UserId						int						NULL,						--所屬使用者Id
 FlowStatus					nvarchar(2)			NOT NULL,						--流程狀態 >> BAS.[Status] B：待處理(Backlog)；I：進行中(In progress)；C：已完成(Completed)
 StartDate					datetime				NULL,						--開始時間
 EndDate					datetime				NULL,						--結束時間
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 CreateBy					int					NOT NULL,						--新增人員
 CONSTRAINT PK_DemandFlowLog PRIMARY KEY (LogId),
 CONSTRAINT FK1_DemandFlowLog FOREIGN KEY (SettingId) REFERENCES SSO.DemandFlowSetting(SettingId),
 CONSTRAINT FK2_DemandFlowLog FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId)
) /*需求流程狀態歷程*/

CREATE TABLE SSO.DemandNotificationLog( /*需求平台通知紀錄*/
 LogId						int					identity,
 SettingId					int					NOT NULL,						--所屬需求Id
 LogContent					nvarchar(MAX)		NOT NULL,						--通知內容
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 CONSTRAINT PK_DemandNotificationLog PRIMARY KEY (LogId),
 CONSTRAINT FK1_DemandNotificationLog FOREIGN KEY (SettingId) REFERENCES SSO.DemandFlowSetting(SettingId),
) /*需求平台通知紀錄*/

CREATE TABLE SSO.DemandNotificationLogDetail( /*需求平台通知詳細紀錄*/
 LogId						int					NOT NULL,						--所屬需求平台通知紀錄
 UserNo						nvarchar(20)		NOT NULL,						--所屬使用者
 CONSTRAINT FK1_DemandNotificationLogDetail FOREIGN KEY (LogId) REFERENCES SSO.DemandNotificationLog(LogId)
) /*需求平台通知詳細紀錄*/
END /*需求追蹤平台*/

BEGIN /*系統設定*/
CREATE TABLE SSO.SourceFlowSetting( /*需求來源流程設定*/
 SettingId					int					identity,
 SourceId					int					NOT NULL,						--所屬需求來源Id
 FlowId						int					NOT NULL,						--所屬流程Id
 Xaxis						int						NULL,						--X座標
 Yaxis						int						NULL,						--Y座標
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_SourceFlowSetting PRIMARY KEY (SettingId),
 CONSTRAINT FK1_SourceFlowSetting FOREIGN KEY (SourceId) REFERENCES SSO.DemandSource(SourceId),
 CONSTRAINT FK2_SourceFlowSetting FOREIGN KEY (FlowId) REFERENCES SSO.DemandFlow(FlowId)
) /*需求來源流程設定*/

CREATE TABLE SSO.SourceFlowLink( /*需求來源流程設定連結*/
 SourceSettingId			int					NOT NULL,						--來源流程設定
 TargetSettingId			int					NOT NULL,						--目標流程設定
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT AK1_SourceFlowLink UNIQUE (SourceSettingId, TargetSettingId),
 CONSTRAINT FK1_SourceFlowLink FOREIGN KEY (SourceSettingId) REFERENCES SSO.SourceFlowSetting(SettingId),
 CONSTRAINT FK2_SourceFlowLink FOREIGN KEY (TargetSettingId) REFERENCES SSO.SourceFlowSetting(SettingId)
) /*需求來源流程連結*/

CREATE TABLE SSO.SourceFlowUser( /*需求來源流程使用者對應*/
 SettingId					int					NOT NULL,						--所屬設定Id
 RoleId						int						NULL,						--所屬角色Id
 UserId						int						NULL,						--所屬使用者Id
 MailAdviceStatus			nvarchar(2)			NOT NULL,						--信件通知
 PushAdviceStatus			nvarchar(2)			NOT NULL,						--推播通知
 WorkWeixinStatus			nvarchar(2)			NOT NULL,						--企業微信通知
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT AK1_SourceFlowUser UNIQUE (SettingId, RoleId, UserId),
 CONSTRAINT FK1_SourceFlowUser FOREIGN KEY (SettingId) REFERENCES SSO.SourceFlowSetting(SettingId),
 CONSTRAINT FK2_SourceFlowUser FOREIGN KEY (RoleId) REFERENCES SSO.DemandRole(RoleId),
 CONSTRAINT FK3_SourceFlowUser FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId)
) /*需求來源流程使用者對應*/
END /*系統設定*/

BEGIN /*研發平台*/

CREATE TABLE SSO.DemandBom( /*需求品號物料*/
 DemandBomId				int					identity,
 DemandId					int					NOT NULL,						--所屬需求來源Id
 BomLevel					int					NOT NULL	DEFAULT 1,			--物料層級(1:第一料皆，例如:主件、2:第二料階，例如:元件)
 ItemTypeId					int					NOT NULL,						--品號類型
 MtlItemNo					nvarchar(40)		NOT NULL,						--INVMB.MB001 品號
 MtlItemName				nvarchar(120)		NOT NULL,						--INVMB.MB002 品名
 MtlItemSpec				nvarchar(120)		NOT NULL,						--INVMB.MB003 規格
 MtlItemDesc				nvarchar(255)			NULL,						--INVMBMB009 商品描述
 MtlItemRemark				nvarchar(255)			NULL,						--INVMB.MB028 品號備註
 UomId						int					NOT NULL,						--單位 >> 參照 ERP品號換算單位檔(INVMD) >> 品號、BOM皆參照
 BomSequence				nvarchar(4)			NOT NULL,						--BOMMD.MD002 元件序號
 CompositionQuantity		float				NOT NULL,						--BOMMD.MD006 單位用量
 Base						float				NOT NULL,						--BOMMD.MD007 底數
 LossRate					float				NOT NULL,						--BOMMD.MD008 損耗率
 StandardCostingType		nvarchar(4)			NOT NULL,						--BOMMD.MD014 標準成本計算
 MaterialProperties			nvarchar(4)			    NULL,						--BOMMD.MD017 材料型態
 TransferUserId				int						NULL,						--拋轉者
 TransferStatus				nvarchar(2)			NOT NULL	DEFAULT 'N',		--拋轉ERP狀態 >> BAS.[Status]
 TransferDate				datetime				NULL,						--拋轉ERP時間
 [Status]					nvarchar(2)			NOT NULL	DEFAULT 'A',		--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DemandBom PRIMARY KEY (DemandBomId),
 CONSTRAINT FK1_DemandBom FOREIGN KEY (DemandId) REFERENCES SSO.Demand(DemandId),
 CONSTRAINT FK2_DemandBom FOREIGN KEY (UomId) REFERENCES PDM.UnitOfMeasure(UomId)
)/*需求品號物料*/

CREATE TABLE SSO.DemandMtlItem(/*需求品號*/
 DemandMtlItemId			int					identity, 
 ItemTypeId					int					NOT NULL,						--品號類型
 DemandId					int					NOT NULL,						--需求來源ID
 PdmMtlItemId				int						NULL,						--PDM.MtlItem.MtlItemId
 MesRoutingProcessItemId	int						NULL,						--MES.RoutingProcessItem.RoutingProcessItemId
 MtlItemNo					nvarchar(40)		NOT NULL,						--MB001 品號
 MtlItemName				nvarchar(120)		NOT NULL,						--MB002 品名
 MtlItemSpec				nvarchar(120)		NOT NULL,						--MB003 規格
 InventoryUomId				int						NULL,						--MB004 庫存單位 >> 參照 ERP品號換算單位檔(INVMD)
 WeightUomId				int						NULL,						--MB015 重量單位 >> 參照 ERP品號換算單位檔(INVMD)
 PurchaseUomId				int						NULL,						--MB155 採購單位 >> 參照 ERP品號換算單位檔(INVMD)
 SaleUomId					int						NULL,						--MB156 銷售單位 >> 參照 ERP品號換算單位檔(INVMD)
 TypeOne					nvarchar(6)			NOT NULL,						--MB005 品號分類一
 TypeTwo					nvarchar(6)			NOT NULL,						--MB006 品號分類二
 TypeThree					nvarchar(6)			NOT NULL,						--MB007 品號分類三
 TypeFour					nvarchar(6)			NOT NULL,						--MB008 品號分類四
 InventoryId				int						NULL,						--MB017 主要庫別
 RequisitionInventoryId		int						NULL,						--MB157 主要領料庫別
 InventoryManagement		nvarchar(2)			NOT NULL,						--MB019 庫存管理 >> BAS.[Status]
 MtlModify					nvarchar(1)			NOT NULL,						--MB066 修改品名/規格 >> BAS.[Status]
 BondedStore				nvarchar(1)			NOT NULL,						--MB020 保稅品 >> BAS.[Status]
 ItemAttribute				nvarchar(1)			NOT NULL,						--MB025 品號屬性 >> BAS.[Type] P：採購件；M：自製件；S：託外加工件；Y：虛設品號；F：Feature件；O:Option件
 ProductionLine				nvarchar(10)			NULL,						--TC005 生產線別 >> 參照 ERP生產線資料單頭檔(CMSMD)
 LotManagement				nvarchar(1)			NOT NULL	DEFAULT 'N',		--MB022 批號管理 >> BAS.[Type] N:不需要；Y:需要不檢查庫存量；W:僅需警告；T:需要且檢查庫存量
 EfficientDays				int						NULL,						--MB023 有效天數
 RetestDays					int						NULL,						--MB024 複檢天數
 MeasureType				nvarchar(1)			NOT NULL	DEFAULT '2',		--MB043 檢驗方式 >> BAS.[Type] 0：免檢；1：抽檢(減量)；2：抽檢(正常)；3：抽檢(加嚴)；4：全檢
 EffectiveDate				datetime				NULL,						--MB030 生效日期
 ExpirationDate				datetime				NULL,						--MB031 失效日期
 OverReceiptManagement		nvarchar(1)			NOT NULL,						--MB044 超收管理 >> BAS.[Status]
 OverDeliveryManagement		nvarchar(1)			NOT NULL	DEFAULT 'N',		--MB087 超交管理 >> BAS.[Status]
 [Version]					nvarchar(4)			NOT NULL,						--MB165 版次
 MtlItemDesc				nvarchar(255)			NULL,						--MB009 商品描述
 MtlItemRemark				nvarchar(255)			NULL,						--MB028 品號備註
 TransferStatus				nvarchar(2)			NOT NULL,						--拋轉ERP狀態 >> BAS.[Status]
 TransferDate				datetime				NULL,						--拋轉ERP時間
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DemandMtlItem PRIMARY KEY (DemandMtlItemId),
 CONSTRAINT AK1_DemandMtlItem UNIQUE (DemandId, MtlItemNo),
 CONSTRAINT FK1_DemandMtlItem FOREIGN KEY (InventoryUomId) REFERENCES PDM.UnitOfMeasure(UomId),
 CONSTRAINT FK2_DemandMtlItem FOREIGN KEY (WeightUomId) REFERENCES PDM.UnitOfMeasure(UomId),
 CONSTRAINT FK3_DemandMtlItem FOREIGN KEY (PurchaseUomId) REFERENCES PDM.UnitOfMeasure(UomId),
 CONSTRAINT FK4_DemandMtlItem FOREIGN KEY (SaleUomId) REFERENCES PDM.UnitOfMeasure(UomId),
 CONSTRAINT FK5_DemandMtlItem FOREIGN KEY (InventoryId) REFERENCES SCM.Inventory(InventoryId),
 CONSTRAINT FK6_DemandMtlItem FOREIGN KEY (RequisitionInventoryId) REFERENCES SCM.Inventory(InventoryId),
 CONSTRAINT FK7_DemandMtlItem FOREIGN KEY (DemandId) REFERENCES SSO.Demand(DemandId)
)/*需求品號*/

CREATE TABLE SSO.DemandBillOfMaterials( /*需求BOM用量資料單頭*/ /*ERP BOMMC*/
 DemandBomId				int					identity,
 DemandMtlItemId			int					NOT NULL,						--所屬品號
 UomId						int					NOT NULL,						--MC002 單位 >> 參照 ERP品號換算單位檔(INVMD)
 StandardLot				float				NOT NULL,						--MC004 標準批量
 WipPrefix					nvarchar(4)				NULL,						--MC005 製令單別
 ModiPrefix					nvarchar(4)				NULL,						--MC006 變更單別
 ModiNo						nvarchar(11)			NULL,						--MC007 變更單號
 ModiSequence				nvarchar(4)				NULL,						--MC008 變更序號
 [Version]					nvarchar(4)			NOT NULL,						--MC009 版次
 Remark						nvarchar(255)		NOT NULL,						--MC010 備註
 ConfirmStatus				nvarchar(1)			NOT NULL,						--MC016 確認碼 >> BAS.[Status]
 ConfirmUserId				int						NULL,						--MC018 確認者
 ConfirmDate				datetime				NULL,						--MC017 確認日期
 TransferStatus				nvarchar(2)			NOT NULL,						--拋轉ERP狀態 >> BAS.[Status]
 TransferDate				datetime				NULL,						--拋轉ERP時間
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DemandBillOfMaterials PRIMARY KEY (DemandBomId),
 CONSTRAINT FK1_DemandBillOfMaterials FOREIGN KEY (DemandMtlItemId) REFERENCES SSO.DemandMtlItem(DemandMtlItemId),
 CONSTRAINT FK2_DemandBillOfMaterials FOREIGN KEY (UomId) REFERENCES PDM.UnitOfMeasure(UomId),
 CONSTRAINT FK3_DemandBillOfMaterials FOREIGN KEY (ConfirmUserId) REFERENCES BAS.[User](UserId)
)/*BOM用量資料單頭*/

CREATE TABLE SSO.DemandBomDetail( /*BOM用量資料單身*/ /*ERP BOMMD*/
 DemandBomDetailId			int					identity,
 DemandBomId				int					NOT NULL,						--MD001 所屬單頭
 BomSequence				nvarchar(4)			NOT NULL,						--MD002 序號
 MtlItemId					int					NOT NULL,						--MD003 所屬品號 (只能選PDM.MtlItem 的 MtlItemId)
 UomId						int					NOT NULL,						--MD004 單位 >> 參照 ERP品號換算單位檔(INVMD)
 CompositionQuantity		float				NOT NULL,						--MD006 單位用量
 Base						float				NOT NULL,						--MD007 底數
 LossRate					float				NOT NULL,						--MD008 損耗率
 StandardCostingType		nvarchar(4)			NOT NULL,						--MD014 標準成本計算
 MaterialProperties			nvarchar(4)			    NULL,						--MD017 材料型態
 ReleaseItem				nvarchar(1)			    NULL,						--MD029 發放品 >> BAS.[Status] (待鼎新回覆)
 EffectiveDate				datetime				NULL,						--MD011 生效日期
 ExpirationDate				datetime				NULL,						--MD012 失效日期
 Remark						nvarchar(255)			NULL,						--MD016 備註
 SubstitutionRemark			nvarchar(255)			NULL,						--BOMMA MA004 取替代料單頭備註
 ConfirmStatus				nvarchar(1)			NOT NULL,						--MD032 確認碼 >> BAS.[Status] N：未確認；Y：已確認
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DemandBomDetail PRIMARY KEY (DemandBomDetailId), 
 CONSTRAINT FK1_DemandBomDetail FOREIGN KEY (DemandBomId) REFERENCES SSO.DemandBillOfMaterials(DemandBomId),
 CONSTRAINT FK2_DemandBomDetail FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId),
 CONSTRAINT FK3_DemandBomDetail FOREIGN KEY (UomId) REFERENCES PDM.UnitOfMeasure(UomId)
)/*BOM用量資料單身*/

CREATE TABLE SSO.DemandCustomerMtlItem( /*需求客戶品號*/
 CustomerMtlItemId			int					identity,
 DemandMtlItemId			int					NOT NULL,						--所屬品號
 CustomerId					int					NOT NULL,						--所屬客戶
 CustomerMtlItemNo			nvarchar(50)			NULL,						--客戶品號(部番)
 CustomerMtlItemName		nvarchar(200)			NULL,						--客戶品名
 CustomerMtlItemSpec		nvarchar(200)			NULL,						--客戶品號規格
 TransferStatus				nvarchar(2)			NOT NULL	DEFAULT 'N',		--拋轉ERP狀態 >> BAS.[Status]
 TransferDate				datetime				NULL,						--拋轉ERP時間
 [Status]					nvarchar(2)			NOT NULL,	DEFAULT 'A',		--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DemandCustomerMtlItem PRIMARY KEY (CustomerMtlItemId),
 CONSTRAINT AK1_DemandCustomerMtlItem UNIQUE (DemandMtlItemId, CustomerId, CustomerMtlItemNo),
 CONSTRAINT FK1_DemandCustomerMtlItem FOREIGN KEY (DemandMtlItemId) REFERENCES SSO.DemandMtlItem(DemandMtlItemId),
 CONSTRAINT FK2_DemandCustomerMtlItem FOREIGN KEY (CustomerId) REFERENCES SCM.Customer(CustomerId)
)/*需求客戶品號*/

CREATE TABLE SSO.DemandMtlUomCalculat( /*需求品號單位換算資料*/
 MtlUomCalculateId			int					identity,
 DemandMtlItemId			int					NOT NULL,						--INVMD.MD001 所屬品號 (只能選PDM.MtlItem 的 MtlItemId)
 ConvertUomId				int					NOT NULL,						--INVMD.MD002 換算單位 
 SwapNumerator				float				NOT NULL,						--INVMD.MD003 換算率分子(換算) 
 SwapDenominator			float				NOT NULL,						--INVMD.MD004 換算率分母(庫存) 
 TransferStatus				nvarchar(2)			NOT NULL    DEFAULT 'N',		--拋轉ERP狀態 >> BAS.[Status]
 TransferDate				datetime				NULL,						--拋轉ERP時間
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DemandMtlUomCalculat PRIMARY KEY (MtlUomCalculateId),
 CONSTRAINT FK1_DemandMtlUomCalculat FOREIGN KEY (DemandMtlItemId) REFERENCES SSO.DemandMtlItem(DemandMtlItemId),
 CONSTRAINT FK2_DemandMtlUomCalculat FOREIGN KEY (ConvertUomId) REFERENCES PDM.UnitOfMeasure(UomId)
)/*需求品號單位換算資料*/

CREATE TABLE SSO.DemandBomSubstitution( /*需求物料取替代料件*/
 SubstitutionId				int					identity,
 DemandBomDetailId			int					NOT NULL,						--所屬BOM用量資料單身
 SubstituteStatus			nvarchar(1)			NOT NULL,						--MB003 取替代料 >> BAS.[Type] 2.取代料件；3.替代料件
 MtlItemId					int					NOT NULL,						--MB004 所屬品號
 Quantity					float				NOT NULL,						--MB005 數量
 SortNumber					int					NOT NULL,						--MB009 取替代件順序
 Precedence					nvarchar(1)			NOT NULL,						--MB010 優先耗用 >> BAS.[Status] N:非優先(DEFAULT)；Y:優先
 EffectiveDate				datetime				NULL,						--MB006 生效日期
 ExpirationDate				datetime				NULL,						--MB007 失效日期
 Remark						nvarchar(255)			NULL,						--MB008 備註
 TransferStatus				nvarchar(2)			NOT NULL,						--拋轉ERP狀態 >> BAS.[Status]
 TransferDate				datetime				NULL,						--拋轉ERP時間
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_BomSubstitution PRIMARY KEY (SubstitutionId),
 CONSTRAINT FK1_BomSubstitution FOREIGN KEY (DemandBomDetailId) REFERENCES SSO.DemandBomDetail(DemandBomDetailId),
 CONSTRAINT FK2_BomSubstitution FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId)
)/*需求物料取替代料件*/

END/*研發平台*/