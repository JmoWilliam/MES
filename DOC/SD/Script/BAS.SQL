BEGIN /*基本資料*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'FunctionDetail') > 0 DROP TABLE BAS.FunctionDetail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Function') > 0 DROP TABLE BAS.[Function]
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Module') > 0 DROP TABLE BAS.Module
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'System') > 0 DROP TABLE BAS.[System]
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'StatusCompany') > 0 DROP TABLE BAS.StatusCompany
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'TypeCompany') > 0 DROP TABLE BAS.TypeCompany
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Status') > 0 DROP TABLE BAS.[Status]
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Type') > 0 DROP TABLE BAS.[Type]
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DepartmentShift') > 0 DROP TABLE BAS.DepartmentShift
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Shift') > 0 DROP TABLE BAS.[Shift]
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'User') > 0 DROP TABLE BAS.[User]
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Department') > 0 DROP TABLE BAS.Department
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Company') > 0 DROP TABLE BAS.Company
END /*基本資料*/

BEGIN /*系統介接*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'InterfacingDataOption') > 0 DROP TABLE BAS.InterfacingDataOption
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Operation') > 0 DROP TABLE BAS.Operation
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'InterfacingConnection') > 0 DROP TABLE BAS.InterfacingConnection
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Interfacing') > 0 DROP TABLE BAS.Interfacing
END /*系統介接*/

BEGIN /*系統設定*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'UserSupplier') > 0 DROP TABLE BAS.UserSupplier
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'UserCustomer') > 0 DROP TABLE BAS.UserCustomer
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'UserEmployee') > 0 DROP TABLE BAS.UserEmployee
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'UserInfo') > 0 DROP TABLE BAS.UserInfo
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'NotificationSetting') > 0 DROP TABLE BAS.NotificationSetting
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'NotificationLogDetail') > 0 DROP TABLE BAS.NotificationLogDetail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'NotificationLog') > 0 DROP TABLE BAS.NotificationLog
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'PushSubscription') > 0 DROP TABLE BAS.PushSubscription
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'UserFunctionDetail') > 0 DROP TABLE BAS.UserFunctionDetail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'UserRole') > 0 DROP TABLE BAS.UserRole
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'RoleFunctionDetail') > 0 DROP TABLE BAS.RoleFunctionDetail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Role') > 0 DROP TABLE BAS.[Role]
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ApiKeyFunction') > 0 DROP TABLE BAS.ApiKeyFunction
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ApiKey') > 0 DROP TABLE BAS.ApiKey
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Log') > 0 DROP TABLE BAS.[Log]
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Calendar') > 0 DROP TABLE BAS.Calendar
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Mail') > 0 DROP TABLE BAS.Mail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MailUser') > 0 DROP TABLE BAS.MailUser
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MailContact') > 0 DROP TABLE BAS.MailContact
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MailServer') > 0 DROP TABLE BAS.MailServer
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'File') > 0 DROP TABLE BAS.[File]
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'UserLoginLog') > 0 DROP TABLE BAS.UserLoginLog
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'UserLoginKey') > 0 DROP TABLE BAS.UserLoginKey
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'SubSystemUser') > 0 DROP TABLE BAS.SubSystemUser
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'SubSystem') > 0 DROP TABLE BAS.SubSystem
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'PasswordSetting') > 0 DROP TABLE BAS.PasswordSetting
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Parameter') > 0 DROP TABLE BAS.Parameter
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DownloadLog' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'BAS')) > 0 DROP TABLE BAS.DownloadLog

END /*系統設定*/
--******************************************************************************************************************
BEGIN /*基本資料*/
CREATE TABLE BAS.Company( /*公司資料*/
 CompanyId					int					identity,
 CompanyNo					nvarchar(10)		NOT NULL,						--公司編號
 CompanyName				nvarchar(100)		NOT NULL,						--公司名稱
 Telephone					nvarchar(30)			NULL,						--電話
 Fax						nvarchar(30)			NULL,						--傳真
 [Address]					nvarchar(200)			NULL,						--中文地址
 AddressEn					nvarchar(200)			NULL,						--英文地址
 LogoIcon					int						NULL,						--Logo Icon >> BAS.[File]
 MainDb						nvarchar(50)			NULL,						--主資料庫
 HrmDb						nvarchar(50)			NULL,						--HRM資料庫
 ErpDb						nvarchar(50)			NULL,						--ERP資料庫
 ErpNo						nvarchar(10)			NULL,						--ERP公司編號
 SysDb						nvarchar(50)			NULL,						--SysDB
 SystemStatus				nvarchar(2)			NOT NULL,						--系統狀態
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼	
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Company PRIMARY KEY (CompanyId),
 CONSTRAINT AK1_Company UNIQUE (CompanyNo),
 CONSTRAINT AK2_Company UNIQUE (CompanyName),
 CONSTRAINT FK1_Company FOREIGN KEY (LogoIcon) REFERENCES BAS.[File](FileId) 
)/*公司資料*/

CREATE TABLE BAS.Department( /*部門資料*/
 DepartmentId				int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 DepartmentNo				nvarchar(10)		NOT NULL,						--部門編號
 DepartmentName				nvarchar(100)		NOT NULL,						--部門名稱 
 SystemStatus				nvarchar(2)			NOT NULL,						--系統狀態
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼	
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員 
 CONSTRAINT PK_Department PRIMARY KEY (DepartmentId),
 CONSTRAINT AK_Department UNIQUE (DepartmentNo, CompanyId),
 CONSTRAINT FK_Department FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)/*部門資料*/

CREATE TABLE BAS.[User]( /*使用者資料*/
 UserId						int					identity, 
 DepartmentId				int					NOT NULL,						--所屬部門
 UserNo						nvarchar(20)		NOT NULL,						--使用者編號
 UserName					nvarchar(30)		NOT NULL,						--使用者名稱
 Gender						nvarchar(2)			NOT NULL,						--性別
 Email						nvarchar(100)			NULL,						--電子郵件
 [Password]					nvarchar(128)		NOT NULL,						--密碼
 Job						nvarchar(50)			NULL,						--職務
 JobType					nvarchar(50)			NULL,						--職務分類
 UserStatus					nvarchar(2)			NOT NULL,						--使用者狀態
 SystemStatus				nvarchar(2)			NOT NULL,						--系統狀態
 PasswordStatus				nvarchar(2)			NOT NULL,						--更改密碼狀態
 PasswordExpire				datetime				NULL,						--密碼期限
 PasswordMistake			int					NOT NULL,						--密碼錯誤次數
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員 
 CONSTRAINT PK_User PRIMARY KEY (UserId),
 CONSTRAINT AK_User UNIQUE (UserNo),
 CONSTRAINT FK_User FOREIGN KEY (DepartmentId) REFERENCES BAS.Department(DepartmentId) 
)/*使用者資料*/

CREATE TABLE BAS.UserToken( /*使用者密鑰(更改密碼用)*/
 TokenId					int					identity, 
 UserId						int					NOT NULL,						--使用者ID
 Token						nvarchar(32)		NOT NULL,						--密鑰
 ValidDate					datetime			NOT NULL,						--最近一次更新日期
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員 
 CONSTRAINT PK_UserToken PRIMARY KEY (TokenId),
 CONSTRAINT AK1_UserToken UNIQUE (UserId),
 CONSTRAINT FK1_UserToken FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId) 
)/*使用者密鑰(更改密碼用)*/

CREATE TABLE BAS.[Shift]( /*班次資料*/
 ShiftId					int					identity, 
 CompanyId					int					NOT NULL,						--所屬公司
 ShiftName					nvarchar(50)		NOT NULL,						--班次名稱
 WorkBeginTime				nvarchar(5)			NOT NULL,						--工作開始時間
 WorkEndTime				nvarchar(5)			NOT NULL,						--工作結束時間
 WorkHours					float				NOT NULL,						--工作時數
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員 
 CONSTRAINT PK_Shift PRIMARY KEY (ShiftId),
 CONSTRAINT AK_Shift UNIQUE (ShiftName,CompanyId),
 CONSTRAINT FK_Shift FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)/*班次資料*/

CREATE TABLE BAS.[Status]( /*狀態資料*/
 StatusId					int					identity,
 StatusSchema				nvarchar(100)		NOT NULL,						--狀態綱要
 StatusNo					nvarchar(10)		NOT NULL,						--狀態代碼
 StatusName					nvarchar(100)		NOT NULL,						--狀態名稱
 StatusDesc					nvarchar(100)		NOT NULL,						--狀態描述
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員 
 CONSTRAINT PK_Status PRIMARY KEY (StatusId),
 CONSTRAINT AK_Status UNIQUE (StatusSchema, StatusNo)
)/*狀態資料*/

CREATE TABLE BAS.[Type]( /*類別資料*/
 TypeId						int					identity,
 TypeSchema					nvarchar(100)		NOT NULL,						--類別綱要
 TypeNo						nvarchar(10)		NOT NULL,						--類別代碼 
 TypeName					nvarchar(100)		NOT NULL,						--類別名稱
 TypeDesc					nvarchar(100)		NOT NULL,						--類別描述
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員 
 CONSTRAINT PK_Type PRIMARY KEY (TypeId),
 CONSTRAINT AK_Type UNIQUE (TypeSchema, TypeNo)
)/*類別資料*/

CREATE TABLE BAS.StatusCompany( /*狀態公司對應資料*/
 StatusId					int					NOT NULL,						--所屬狀態
 CompanyId					int					NOT NULL,						--所屬公司
 StatusName					nvarchar(100)		NOT NULL,						--狀態名稱
 StatusDesc					nvarchar(100)		NOT NULL,						--狀態描述
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員 
 CONSTRAINT AK1_StatusCompany UNIQUE (StatusId, CompanyId),
 CONSTRAINT FK1_StatusCompany FOREIGN KEY (StatusId) REFERENCES BAS.[Status](StatusId),
 CONSTRAINT FK2_StatusCompany FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)/*狀態公司對應資料*/

CREATE TABLE BAS.TypeCompany( /*類別公司對應資料*/
 TypeId						int					NOT NULL,						--所屬狀態
 CompanyId					int					NOT NULL,						--所屬公司
 TypeName					nvarchar(100)		NOT NULL,						--類別名稱
 TypeDesc					nvarchar(100)		NOT NULL,						--類別描述
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員 
 CONSTRAINT AK1_TypeCompany UNIQUE (TypeId, CompanyId),
 CONSTRAINT FK1_TypeCompany FOREIGN KEY (TypeId) REFERENCES BAS.[Type](TypeId),
 CONSTRAINT FK2_TypeCompany FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)/*類別公司對應資料*/

CREATE TABLE BAS.DepartmentShift( /*部門班次資料*/
 DepartmentId				int					NOT NULL,						--所屬部門
 ShiftId					int					NOT NULL,						--所屬班次
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員 
 CONSTRAINT PK_DepartmentShift PRIMARY KEY (DepartmentId, ShiftId),
 CONSTRAINT FK1_DepartmentShift FOREIGN KEY (DepartmentId) REFERENCES BAS.[Department](DepartmentId),
 CONSTRAINT FK2_DepartmentShift FOREIGN KEY (ShiftId) REFERENCES BAS.[Shift](ShiftId)
)/*部門班次資料*/

CREATE TABLE BAS.DepartmentRate( /*部門費用率資料*/
 DepartmentRateId			int					identity,
 DepartmentId				int					NOT NULL,						--部門
 Author						int					NOT NULL,						--編寫者
 ResourceRate				float				NOT NULL,						--人工費用率
 OverheadRate				float				NOT NULL,						--製造費用率
 EnableDate					datetime				NULL,						--啟用日期
 DisabledDate				datetime				NULL,						--關閉日期
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'S',		--狀態碼		   >> S:關閉 ; A:啟用	
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員 
 CONSTRAINT PK_DepartmentRate PRIMARY KEY (DepartmentRateId),
 CONSTRAINT FK1_DepartmentRate FOREIGN KEY (DepartmentId) REFERENCES BAS.Department(DepartmentId),
 CONSTRAINT FK2_DepartmentRate FOREIGN KEY (Author) REFERENCES BAS.[User](UserId)
)/*部門費用率資料*/

CREATE TABLE BAS.[System]( /*系統別資料*/
 SystemId					int					identity,
 SystemCode					nvarchar(10)		NOT NULL,						--系統代碼
 SystemName					nvarchar(100)		NOT NULL,						--系統名稱
 ThemeIcon					nvarchar(100)			NULL,						--主題Icon 
 SortNumber					int					NOT NULL,						--排序號碼
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_System PRIMARY KEY (SystemId),
 CONSTRAINT AK1_System UNIQUE (SystemCode),
 CONSTRAINT AK2_System UNIQUE (SystemName),
 CONSTRAINT AK3_System UNIQUE (SortNumber)
)/*系統別資料*/

CREATE TABLE BAS.Module( /*模組別資料*/
 ModuleId					int					identity,
 SystemId					int					NOT NULL,						--所屬系統
 ModuleCode					nvarchar(50)		NOT NULL,						--模組代碼
 ModuleName					nvarchar(100)		NOT NULL,						--模組名稱
 ThemeIcon					nvarchar(100)			NULL,						--主題Icon
 SortNumber					int					NOT NULL,						--排序號碼
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Module PRIMARY KEY (ModuleId),
 CONSTRAINT AK1_Module UNIQUE (ModuleCode),
 CONSTRAINT AK2_Module UNIQUE (SystemId, ModuleName),
 CONSTRAINT AK3_Module UNIQUE (SystemId, SortNumber),
 CONSTRAINT FK_Module FOREIGN KEY (SystemId) REFERENCES BAS.[System](SystemId)
)/*模組別資料*/

CREATE TABLE BAS.[Function]( /*功能別資料*/
 FunctionId					int					identity,
 ModuleId					int					NOT NULL,						--所屬模組
 FunctionCode				nvarchar(50)		NOT NULL,						--功能代碼
 FunctionName				nvarchar(100)		NOT NULL,						--功能名稱
 SortNumber					int					NOT NULL,						--排序號碼
 UrlTarget					nvarchar(50)		NOT NULL,						--開啟連結
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Function PRIMARY KEY (FunctionId),
 CONSTRAINT AK1_Function UNIQUE (ModuleId, FunctionCode),
 CONSTRAINT AK2_Function UNIQUE (ModuleId, FunctionName),
 CONSTRAINT AK3_Function UNIQUE (ModuleId, SortNumber),
 CONSTRAINT FK_Function FOREIGN KEY (ModuleId) REFERENCES BAS.Module(ModuleId)
)/*功能別資料*/

CREATE TABLE BAS.FunctionDetail( /*詳細功能別資料*/
 DetailId					int					identity,
 FunctionId					int					NOT NULL,						--所屬功能
 DetailCode					nvarchar(50)		NOT NULL,						--功能詳細代碼
 DetailName					nvarchar(100)		NOT NULL,						--功能詳細名稱
 SortNumber					int					NOT NULL,						--排序號碼
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_FunctionDetail PRIMARY KEY (DetailId),
 CONSTRAINT AK1_FunctionDetail UNIQUE (FunctionId, DetailCode),
 CONSTRAINT AK2_FunctionDetail UNIQUE (FunctionId, DetailName),
 CONSTRAINT FK_FunctionDetail FOREIGN KEY (FunctionId) REFERENCES BAS.[Function](FunctionId)
)/*詳細功能別資料*/

CREATE TABLE BAS.UploadWhitelist( /*上傳檔案路徑白名單*/
 ListId						int					identity,
 ListNo						nvarchar(50)		NOT NULL,						--白名單類別
 DepartmentId				int					NOT NULL,						--使用部門ID
 FolderPath					nvarchar(max)		NOT NULL,						--路徑
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_UploadWhitelist PRIMARY KEY (ListId),
 CONSTRAINT FK1_UploadWhitelist FOREIGN KEY (DepartmentId) REFERENCES BAS.Department(DepartmentId),
)/*上傳檔案路徑白名單*/

CREATE TABLE BAS.CompanyUserMapping( /*不同公司與使用者之間的對應關係表(紘立vs中揚)*/
 CUmapId					int						identity,
 BeforeUserId				int						NULL,						--原公司使用者ID
 BeforeDepartmentId			int						NULL,						--原公司部門ID
 NewUserId					int						NULL,						--新公司使用者ID
 NewDepartmentId			int						NULL,						--新公司部門ID
 CreateDate					datetime				NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_CompanyUserMapping PRIMARY KEY (CUmapId)
)/*不同公司與使用者之間的對應關係表(紘立vs中揚)*/


CREATE TABLE BAS.LineCategoryMaster( /*線別歸類主檔表*/
 LcmId						int						identity,
 CompanyId					int					NOT NULL,						--所屬公司
 CategoryCode 				int						NULL,						--線別歸類代碼
 CategoryName	 			int						NULL,						--線別歸類名稱
 CategoryDescription		int						NULL,						--歸類說明
 PurposeType				nvarchar(2)			NOT NULL,						--類型（如：成型事业处/模具事业处） BAS.Type
 DisplayOrder				int					NOT NULL,						--排序
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime				NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_LineCategoryMaster PRIMARY KEY (LcmId),
 CONSTRAINT FK_LineCategoryMaster FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)/*線別歸類主檔表*/

CREATE TABLE BAS.DepartmentLineMapping( /*部門線別對應表*/
 DlmId						int						identity,
 DepartmentId				int						NULL,						--部門ID
 LcmId						int						NULL,						--線別歸類主檔表ID
 Remarks					nvarchar(100)			NOT NULL,					--備註
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime				NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DepartmentLineMapping PRIMARY KEY (DlmId),
 CONSTRAINT FK1_DepartmentLineMapping FOREIGN KEY (DepartmentId) REFERENCES BAS.Department(DepartmentId),
 CONSTRAINT FK2_DepartmentLineMapping FOREIGN KEY (LcmId) REFERENCES BAS.LineCategoryMaster(LcmId)
)/*  部門線別對應表*/

CREATE TABLE BAS.ErpLineMapping ( /*ERP 生產線別對應表*/
 ElmId						int						identity,  
 ErpLineCode				nvarchar(100)			NOT NULL,					--ERP系統中的生產線別代碼
 ErpLineName	 			nvarchar(100)			NOT NULL,					--ERP系統中的生產線別名稱
 ErpFactoryCode				nvarchar(100)			NOT NULL,					--ERP工廠代碼
 LcmId						int						NULL,						--線別歸類主檔表ID
 Remarks					nvarchar(100)			NOT NULL,					--備註
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime				NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ErpLineMapping PRIMARY KEY (DlmId),
 CONSTRAINT FK_ErpLineMappin FOREIGN KEY (LcmId) REFERENCES BAS.LineCategoryMaster(LcmId)
)/*ERP 生產線別對應表*/

END /*基本資料*/

BEGIN /*系統介接*/
CREATE TABLE BAS.Interfacing( /*系統介接設定*/
 InterfacingId				int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 InterfacingType			nvarchar(10)		NOT NULL,						--介接類型 >> BAS.[Type][TypeSchema = Interfacing.InterfacingType] ERP_Main：ERP主檔；ERP_System：ERP系統檔；HRM_Main：HRM主檔；AD_Main：AD主機
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Interfacing PRIMARY KEY (InterfacingId),
 CONSTRAINT AK1_Interfacing UNIQUE (CompanyId, InterfacingType),
 CONSTRAINT FK1_Interfacing FOREIGN KEY (CreateBy) REFERENCES BAS.UserInfo(UserId),
 CONSTRAINT FK2_Interfacing FOREIGN KEY (LastModifiedBy) REFERENCES BAS.UserInfo(UserId),
 CONSTRAINT FK3_Interfacing FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
) /*系統介接設定*/

CREATE TABLE BAS.InterfacingConnection( /*介接連線*/
 ConnectionId				int					identity,
 InterfacingId				int					NOT NULL,						--所屬系統介接設定
 ConnectionServer			nvarchar(50)		NOT NULL,						--連線主機
 ConnectionAccount			nvarchar(50)		NOT NULL,						--連線帳號
 ConnectionPassword			nvarchar(50)		NOT NULL,						--連線密碼
 ConnectionDatabase			nvarchar(50)			NULL,						--連線資料庫
 [Status]					nvarchar(10)		NOT NULL,						--狀態碼 >> BAS.[Status][StatusSchema = Status] A：啟用(Active)；S：停用(Stop)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_InterfacingConnection PRIMARY KEY (ConnectionId),
 CONSTRAINT FK1_InterfacingConnection FOREIGN KEY (CreateBy) REFERENCES BAS.UserInfo(UserId),
 CONSTRAINT FK2_InterfacingConnection FOREIGN KEY (LastModifiedBy) REFERENCES BAS.UserInfo(UserId),
 CONSTRAINT FK3_InterfacingConnection FOREIGN KEY (InterfacingId) REFERENCES BAS.Interfacing(InterfacingId),
) /*介接連線*/

CREATE TABLE BAS.Operation( /*作業模式*/
 OperationId				int					identity,
 InterfacingId				int					NOT NULL,						--所屬系統介接設定
 OperationNo				nvarchar(50)		NOT NULL,						--作業模式代碼
 OperationName				nvarchar(50)		NOT NULL,						--作業模式名稱
 [Status]					nvarchar(10)		NOT NULL,						--狀態碼 >> BAS.[Status][StatusSchema = Status] A：啟用(Active)；S：停用(Stop)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Operation PRIMARY KEY (OperationId),
 CONSTRAINT AK1_Operation UNIQUE (InterfacingId, OperationNo),
 CONSTRAINT FK1_Operation FOREIGN KEY (CreateBy) REFERENCES BAS.UserInfo(UserId),
 CONSTRAINT FK2_Operation FOREIGN KEY (LastModifiedBy) REFERENCES BAS.UserInfo(UserId),
 CONSTRAINT FK3_Operation FOREIGN KEY (InterfacingId) REFERENCES BAS.Interfacing(InterfacingId),
) /*作業模式*/

CREATE TABLE BAS.OperationSetting( /*作業模式設定*/
 OperationId				int					NOT NULL,						--所屬作業模式
 AutoImport					nvarchar(10)		NOT NULL,						--自動拋轉 >> BAS.[Status][StatusSchema = Boolean] N：否(No)；Y：是(Yes)
 AutoSynchronize			nvarchar(10)		NOT NULL,						--自動同步 >> BAS.[Status][StatusSchema = Boolean] N：否(No)；Y：是(Yes)
 ImportTimeInterval			int						NULL,						--拋轉時間間隔(秒)
 SynchronizeTimeInterval	int						NULL,						--同步時間間隔(秒)
 ImportExecuteMode			nvarchar(10)			NULL,						--拋轉執行方式 >> BAS.[Type][TypeSchema = OperationSetting.ExecuteMode] S：單筆(Single)；L：批次(Lot)
 SynchronizeExecuteMode		nvarchar(10)			NULL,						--同步執行方式 >> BAS.[Type][TypeSchema = OperationSetting.ExecuteMode] S：單筆(Single)；L：批次(Lot)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT AK1_OperationSetting UNIQUE (OperationId),
 CONSTRAINT FK1_OperationSetting FOREIGN KEY (CreateBy) REFERENCES BAS.UserInfo(UserId),
 CONSTRAINT FK2_OperationSetting FOREIGN KEY (LastModifiedBy) REFERENCES BAS.UserInfo(UserId),
 CONSTRAINT FK3_OperationSetting FOREIGN KEY (OperationId) REFERENCES BAS.Operation(OperationId),
) /*作業模式設定*/

CREATE TABLE BAS.InterfacingDataOption( /*介接系統資料選項*/
 OptionId					int					identity,
 InterfacingId				int					NOT NULL,						--所屬系統介接設定
 OptionGroup				nvarchar(100)		NOT NULL,						--資料選項群組
 OptionNo					nvarchar(10)		NOT NULL,						--資料選項代碼
 OptionName					nvarchar(100)		NOT NULL,						--資料選項名稱
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_InterfacingDataOption PRIMARY KEY (OptionId),
 CONSTRAINT AK1_InterfacingDataOption UNIQUE (InterfacingId, OptionGroup, OptionNo),
 CONSTRAINT FK1_InterfacingDataOption FOREIGN KEY (CreateBy) REFERENCES BAS.UserInfo(UserId),
 CONSTRAINT FK2_InterfacingDataOption FOREIGN KEY (LastModifiedBy) REFERENCES BAS.UserInfo(UserId),
 CONSTRAINT FK3_InterfacingDataOption FOREIGN KEY (InterfacingId) REFERENCES BAS.Interfacing(InterfacingId),
) /*介接系統資料選項*/
END /*系統介接*/

BEGIN /*系統設定*/
CREATE TABLE BAS.UserInfo( /*使用者資料*/
 UserId						int					identity,
 UserType					nvarchar(10)		NOT NULL,						--使用者類型 >> BAS.[Type] E：內部員工(Employee)；C：客戶(Customer)；S：供應商(Supplier)
 Account					nvarchar(320)		NOT NULL,						--帳號
 [Password]					nvarchar(128)		NOT NULL,						--密碼
 InnerCode					nvarchar(20)			NULL,						--內碼 >> UserType = E：員工工號
 DisplayName				nvarchar(50)		NOT NULL,						--顯示名稱 >> UserType = E：員工姓名
 PasswordStatus				nvarchar(10)		NOT NULL,						--更改密碼狀態 >> BAS.[Status][StatusSchema = Boolean] Y：是(Yes)；N：否(No)
 PasswordExpire				datetime				NULL,						--密碼期限
 PasswordMistake			int					NOT NULL	DEFAULT 0,			--密碼錯誤次數
 [Status]					nvarchar(10)		NOT NULL,						--狀態碼 >> BAS.[Status][StatusSchema = Status] A：啟用(Active)；S：停用(Stop)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int						NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_UserInfo PRIMARY KEY (UserId),
 CONSTRAINT AK1_UserInfo UNIQUE (Account),
 CONSTRAINT FK1_UserInfo FOREIGN KEY (CreateBy) REFERENCES BAS.UserInfo(UserId),
 CONSTRAINT FK2_UserInfo FOREIGN KEY (LastModifiedBy) REFERENCES BAS.UserInfo(UserId),
) /*使用者資料*/

CREATE TABLE BAS.UserEmployee( /*使用者內部員工設定檔*/
 UserId						int					NOT NULL,						--所屬使用者
 DepartmentId				int					NOT NULL,						--所屬部門
 Gender						nvarchar(10)		NOT NULL,						--性別 >> BAS.[Type] F：女(Female)；M：男(Male)；N：無(None)
 Email						nvarchar(100)			NULL,						--電子郵件
 Job						nvarchar(50)			NULL,						--職務
 JobType					nvarchar(50)			NULL,						--職務分類
 EmployeeStatus				nvarchar(10)		NOT NULL,						--內部員工狀態 >> BAS.[Status] F：正式員工(Full time)；L：留職停薪員工(Leave of absence)；P：約聘或其他；R：退休員工(Retire)；S：離職員工(Separation)；T：試用員工
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_UserEmployee PRIMARY KEY (UserId, DepartmentId),
 CONSTRAINT AK1_UserEmployee UNIQUE (UserId),
 CONSTRAINT FK1_UserEmployee FOREIGN KEY (CreateBy) REFERENCES BAS.UserInfo(UserId),
 CONSTRAINT FK2_UserEmployee FOREIGN KEY (LastModifiedBy) REFERENCES BAS.UserInfo(UserId),
 CONSTRAINT FK3_UserEmployee FOREIGN KEY (UserId) REFERENCES BAS.UserInfo(UserId),
 CONSTRAINT FK4_UserEmployee FOREIGN KEY (DepartmentId) REFERENCES BAS.Department(DepartmentId),
) /*使用者內部員工設定檔*/

CREATE TABLE BAS.UserCustomer( /*使用者客戶設定檔*/
 UserId						int					NOT NULL,						--所屬使用者
 Contact					nvarchar(30)			NULL,						--聯絡人
 ContactPhone				nvarchar(20)			NULL,						--聯絡電話
 ContactAddress				nvarchar(255)			NULL,						--聯絡地址
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT AK1_UserCustomer UNIQUE (UserId),
 CONSTRAINT FK1_UserCustomer FOREIGN KEY (CreateBy) REFERENCES BAS.UserInfo(UserId),
 CONSTRAINT FK2_UserCustomer FOREIGN KEY (LastModifiedBy) REFERENCES BAS.UserInfo(UserId),
 CONSTRAINT FK3_UserCustomer FOREIGN KEY (UserId) REFERENCES BAS.UserInfo(UserId),
) /*使用者客戶設定檔*/

CREATE TABLE BAS.UserSupplier( /*使用者供應商設定檔*/
 UserId						int					NOT NULL,						--所屬使用者
 Contact					nvarchar(30)			NULL,						--聯絡人
 ContactPhone				nvarchar(20)			NULL,						--聯絡電話
 ContactAddress				nvarchar(255)			NULL,						--聯絡地址
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT AK1_UserSupplier UNIQUE (UserId),
 CONSTRAINT FK1_UserSupplier FOREIGN KEY (CreateBy) REFERENCES BAS.UserInfo(UserId),
 CONSTRAINT FK2_UserSupplier FOREIGN KEY (LastModifiedBy) REFERENCES BAS.UserInfo(UserId),
 CONSTRAINT FK3_UserSupplier FOREIGN KEY (UserId) REFERENCES BAS.UserInfo(UserId),
) /*使用者供應商設定檔*/

CREATE TABLE BAS.ApiKey( /*Api金鑰*/
 KeyId						int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 KeyText					nvarchar(64)		NOT NULL,						--金鑰
 Purpose					nvarchar(50)		NOT NUll,						--用途
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼	
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員 
 CONSTRAINT PK_ApiKey PRIMARY KEY (KeyId),
 CONSTRAINT AK_ApiKey UNIQUE (KeyText),
 CONSTRAINT FK_ApiKey FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)/*Api金鑰*/

CREATE TABLE BAS.ApiKeyFunction( /*Api適用功能*/
 FunctionId					int					identity,
 KeyId						int					NOT NULL,						--所屬Api金鑰
 FunctionCode				nvarchar(100)		NOT NULL,						--功能代碼
 Purpose					nvarchar(50)		NOT NUll,						--用途
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員 
 CONSTRAINT PK_ApiKeyFunction PRIMARY KEY (FunctionId),
 CONSTRAINT AK_ApiKeyFunction UNIQUE (KeyId, FunctionCode),
 CONSTRAINT FK_ApiKeyFunction FOREIGN KEY (KeyId) REFERENCES BAS.ApiKey(KeyId) 
)/*Api適用功能*/

CREATE TABLE BAS.[Role]( /*角色資料*/
 RoleId						int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 RoleName					nvarchar(50)		NOT NULL,						--角色名稱
 AdminStatus				nvarchar(2)			NOT NULL,						--管理者狀態
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Role PRIMARY KEY (RoleId),
 CONSTRAINT AK_Role UNIQUE (CompanyId, RoleName),
 CONSTRAINT FK_Role FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)/*角色資料*/

CREATE TABLE BAS.RoleFunctionDetail( /*角色詳細功能資料*/
 RoleId						int					NOT NULL,						--所屬角色
 DetailId					int					NOT NULL,						--所屬詳細功能
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_RoleFunctionDetail PRIMARY KEY (RoleId, DetailId),
 CONSTRAINT FK1_RoleFunctionDetail FOREIGN KEY (RoleId) REFERENCES BAS.[Role](RoleId),
 CONSTRAINT FK2_RoleFunctionDetail FOREIGN KEY (DetailId) REFERENCES BAS.FunctionDetail(DetailId)
)/*角色詳細功能資料*/

CREATE TABLE BAS.UserRole( /*使用者角色資料*/
 UserId						int					NOT NULL,						--所屬使用者
 RoleId						int					NOT NULL,						--所屬角色
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT AK1_UserRole UNIQUE (UserId, RoleId),
 CONSTRAINT FK1_UserRole FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK2_UserRole FOREIGN KEY (RoleId) REFERENCES BAS.[Role](RoleId)
)/*使用者角色資料*/

CREATE TABLE BAS.UserFunctionDetail( /*使用者詳細功能資料(停用)*/
 UserId						int					NOT NULL,						--所屬使用者
 DetailId					int					NOT NULL,						--所屬詳細功能
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_UserFunctionDetail PRIMARY KEY (UserId, DetailId),
 CONSTRAINT FK1_UserFunctionDetail FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK2_UserFunctionDetail FOREIGN KEY (DetailId) REFERENCES BAS.FunctionDetail(DetailId)
)/*使用者詳細功能資料(停用)*/

CREATE TABLE BAS.[Log]( /*系統日誌資料*/
 LogId						int					identity,
 [Level]					nvarchar(10)		NOT NULL,						--日誌等級
 [Message]					nvarchar(MAX)		NOT NULL,						--日誌內容
 Host						nvarchar(50)		NOT NULL,						--主機名稱
 UserNo						nvarchar(20)		NOT NULL,						--使用者編號
 TriggerUrl					nvarchar(MAX)		NOT NULL,						--觸發網址
 ServerIP					nvarchar(50)		NOT NULL,						--主機IP
 ClientIP					nvarchar(50)		NOT NULL,						--使用者IP
 [Source]					nvarchar(200)		NOT NULL,						--來源
 Detail						nvarchar(MAX)		NOT NULL,						--詳細說明
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 CreateBy					int					NOT NULL,						--新增人員
 CONSTRAINT PK_Log PRIMARY KEY (LogId)
)/*系統日誌資料*/

CREATE TABLE BAS.MailServer( /*郵件伺服器資料*/
 ServerId					int					identity,						
 CompanyId					int					NOT NULL,						--所屬公司
 Host						nvarchar(50)		NOT NULL,						--主機
 [Port]						int					NOT NULL,						--埠號
 SendMode					nvarchar(10)		NOT NULL,						--寄送模式
 Account					nvarchar(100)		NOT NULL,						--帳號
 [Password] 				nvarchar(100)		NOT NULL,						--密碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MailServer PRIMARY KEY (ServerId),
 CONSTRAINT FK_MailServer FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)/*郵件伺服器資料*/

CREATE TABLE BAS.MailContact( /*郵件聯絡人資料*/
 ContactId					int					identity,						
 CompanyId					int					NOT NULL,						--所屬公司
 ContactName				nvarchar(50)		NOT NULL,						--聯絡人姓名
 Email						nvarchar(100)		NOT NULL,						--聯絡人信箱
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MailContact PRIMARY KEY (ContactId),
 CONSTRAINT FK_MailContact FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)/*郵件聯絡人資料*/

CREATE TABLE BAS.Mail( /*郵件資料*/
 MailId						int					identity,						
 ServerId					int					NOT NULL,						--所屬伺服器
 MailName					nvarchar(50)		NOT NULL,						--郵件名稱
 MailFrom					nvarchar(100)		NOT NULL,						--寄件人
 MailTo						nvarchar(MAX)		NOT NULL,						--收件人
 MailCc						nvarchar(MAX)		NOT NULL,						--副本
 MailBcc 					nvarchar(MAX)		NOT NULL,						--秘密副本
 MailSubject				nvarchar(100)		NOT NULL,						--郵件主旨
 MailContent				nvarchar(MAX)		NOT NULL,						--郵件內容
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Mail PRIMARY KEY (MailId),
 CONSTRAINT FK_Mail FOREIGN KEY (ServerId) REFERENCES BAS.MailServer(ServerId)
)/*郵件資料*/

CREATE TABLE BAS.MailUser( /*郵件人員*/
 MailId						int					NOT NULL,						--所屬郵件
 ContactId					int						NULL,						--所屬聯絡人
 UserId						int						NULL,						--所屬使用者
 MailUserType				nvarchar(2)			NOT NULL,						--郵件人員類別 >> BAS.[Type] F：寄件人(From) T：收件人(To)；C：副本(Cc)；B：秘密副本(Bcc)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 CreateBy					int					NOT NULL,						--新增人員
 CONSTRAINT FK1_MailUser FOREIGN KEY (MailId) REFERENCES BAS.Mail(MailId),
 CONSTRAINT FK2_MailUser FOREIGN KEY (ContactId) REFERENCES BAS.MailContact(ContactId),
 CONSTRAINT FK3_MailUser FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId),
)/*郵件人員*/

CREATE TABLE BAS.MailSendSetting ( /*郵件寄送設定*/
 SettingId					int					identity,
 MailId						int					NOT NULL,						--所屬郵件
 SettingSchema				nvarchar(100)		NOT NULL,						--寄送綱要
 SettingNo					nvarchar(10)			NULL,						--寄送代碼 
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MailSendSetting PRIMARY KEY (SettingId),
 CONSTRAINT FK1_MailSendSetting FOREIGN KEY (MailId) REFERENCES BAS.Mail(MailId)
)

CREATE TABLE BAS.Calendar( /*行事曆資料*/
 CalendarId					int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 CalendarDate				datetime			NOT NULL,						--日期
 CalendarDesc				nvarchar(100)		    NULL,						--描述
 DateType					nvarchar(2)			NOT NULL,						--日期類型 > W(working day)：工作日，R(rest day)：休息日，M(mandatory day off)：例假日，H(holiday)：節日
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Calendar PRIMARY KEY(CalendarId),
 CONSTRAINT FK1_Calendar FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT AK1_Calendar UNIQUE(CompanyId, CalendarDate)
)/*行事曆資料*/

CREATE TABLE BAS.[File]( /*檔案資料*/
 FileId						int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 [FileName]					nvarchar(MAX)		NOT NULL,						--檔案名稱
 FileContent				varbinary(MAX)		NOT NULL,						--檔案內容
 FileExtension				nvarchar(10)		NOT NULL,						--副檔名
 FileSize					int					NOT NULL,						--檔案大小 bytes
 ClientIP					nvarchar(50)		NOT NULL,						--使用者IP
 [Source]					nvarchar(200)		NOT NULL,						--來源
 DeleteStatus				nvarchar(2)			NOT NULL,						--刪除狀態
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_File PRIMARY KEY(FileId),
 CONSTRAINT FK1_File FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)/*檔案資料*/

CREATE TABLE BAS.UserLoginLog( /*使用者登入紀錄*/
 LogId						int					identity,
 UserId						int					NOT NULL,						--所屬使用者
 LoginDate					datetime			NOT NULL,						--登入時間
 LoginIP					nvarchar(50)		NOT NULL,						--登入IP
 LoginSite					nvarchar(50)		NOT NULL,						--登入站台
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_UserLoginLog PRIMARY KEY(LogId)
)/*使用者登入紀錄*/

CREATE TABLE BAS.UserLoginKey( /*使用者登入金鑰*/
 KeyId						int					identity,
 UserId						int					NOT NULL,						--所屬使用者
 KeyText					nvarchar(256)		NOT NULL,						--金鑰
 ExpirationDate				datetime			NOT NULL,						--過期時間
 LoginIP					nvarchar(50)		NOT NULL,						--登入IP
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_UserLoginKey PRIMARY KEY(KeyId),
 CONSTRAINT AK1_UserLoginKey UNIQUE(KeyText),
 CONSTRAINT FK1_UserLoginKey FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId)
)/*使用者登入金鑰*/

CREATE TABLE BAS.SubSystem( /*子系統資料*/
 SubSystemId				int					identity,
 SubSystemCode				nvarchar(50)		NOT NULL,						--子系統代碼
 SubSystemName				nvarchar(100)		NOT NULL,						--子系統名稱
 KeyText					nvarchar(64)		NOT NULL,						--金鑰
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_SubSystem PRIMARY KEY(SubSystemId),
 CONSTRAINT AK1_SubSystem UNIQUE(SubSystemCode),
 CONSTRAINT AK2_SubSystem UNIQUE (KeyText)
) /*子系統資料*/

CREATE TABLE BAS.SubSystemUser( /*子系統使用者資料*/
 SubSystemUserId			int					identity,
 SubSystemId				int					NOT NULL,						--所屬子系統
 UserId						int					NOT NULL,						--所屬使用者
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_SubSystemUser PRIMARY KEY(SubSystemUserId),
 CONSTRAINT AK1_SubSystemUser UNIQUE(SubSystemId, UserId),
 CONSTRAINT FK1_SubSystemUser FOREIGN KEY (SubSystemId) REFERENCES BAS.SubSystem(SubSystemId),
 CONSTRAINT FK2_SubSystemUser FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId)
) /*子系統使用者資料*/

CREATE TABLE BAS.PasswordSetting( /*密碼參數設定資料*/
 PasswordExpiration			int					NOT NULL,						--密碼過期時間(分鐘)
 PasswordFormat				nvarchar(MAX)		NOT NULL,						--密碼格式
 PasswordWrongCount			int					NOT NULL,						--密碼錯誤次數
 LastModifiedDate			datetime				NULL,						--修改日期
 LastModifiedBy				int						NULL,						--修改人員
) /*密碼參數設定資料*/

CREATE TABLE BAS.PushSubscription( /*推播訂閱資料*/
 PushSubscriptionId			int					identity,
 UserId						int					NOT NULL,						--所屬使用者
 ApiEndpoint				nvarchar(MAX)		NOT NULL,						--API端點
 Keysp256dh					nvarchar(MAX)		NOT NULL,						--金鑰
 Keysauth					nvarchar(MAX)		NOT NULL,						--授權
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 CreateBy					int					NOT NULL,						--新增人員
 CONSTRAINT PK_PushSubscription PRIMARY KEY(PushSubscriptionId),
 CONSTRAINT FK1_PushSubscription FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId)
) /*推播訂閱資料*/

CREATE TABLE BAS.NotificationLog( /*系統通知紀錄*/
 LogId						int					identity,
 UserNo						nvarchar(20)		NOT NULL,						--所屬使用者
 TriggerFunction			nvarchar(MAX)		NOT NULL,						--觸發功能
 LogTitle					nvarchar(50)		NOT NULL,						--通知標題
 LogContent					nvarchar(MAX)		NOT NULL,						--通知內容
 ReadStatus					nvarchar(2)			NOT NULL,						--讀取狀態 >> BAS.[Status] Y；N(Default)
 ReadDate					datetime				NULL,						--讀取時間
 ReadIP						nvarchar(50)			NULL,						--讀取IP
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 CONSTRAINT PK_NotificationLog PRIMARY KEY (LogId)
) /*系統通知紀錄*/

CREATE TABLE BAS.NotificationLogDetail( /*系統通知詳細紀錄*/
 LogId						int					NOT NULL,						--所屬系統通知紀錄
 LogType					nvarchar(10)		NOT NULL,						--通知類型 >> BAS.[Type] Mail：信件通知；Push：推播通知；WorkWeixin：企業微信通知
 CONSTRAINT FK1_NotificationLogDetail FOREIGN KEY (LogId) REFERENCES BAS.NotificationLog(LogId)
) /*系統通知詳細紀錄*/

CREATE TABLE BAS.NotificationSetting( /*系統通知個人設定*/
 UserId						int					NOT NULL,
 MailAdviceStatus			nvarchar(2)			NOT NULL,						--信件通知
 PushAdviceStatus			nvarchar(2)			NOT NULL,						--推播通知
 WorkWeixinStatus			nvarchar(2)			NOT NULL,						--企業微信通知
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT AK1_NotificationSetting UNIQUE (UserId),
 CONSTRAINT FK1_NotificationSetting FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId)
) /*需求平台個人設定*/

CREATE TABLE BAS.DownloadLog( /*下載紀錄*/
 DownloadLogId				int					identity,
 FileId						int					NOT NULL,						--檔案ID 
 ClientComputer				nvarchar(50)		NOT NULL,						--使用者電腦名稱
 ClientIP					nvarchar(50)		NOT NULL,						--使用者IP
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 CreateBy					int					NOT NULL,						--新增人員
 CONSTRAINT PK_DownloadLog PRIMARY KEY (DownloadLogId),
 CONSTRAINT FK1_DownloadLog FOREIGN KEY (FileId) REFERENCES BAS.[File](FileId)
)/*下載紀錄*/

CREATE TABLE BAS.UserOperateLog( /* 使用者操作紀錄 */
 LogId						int					identity,				
 DocId						int						NULL,						--單據ID					
 ErpFullNo					nvarchar(150)			NULL,						--完整ERP單據單號					
 FullFunctionName			nvarchar(150)		NOT NULL,						--完整Controller/Function Name						
 IpAddress					nvarchar(150)		NOT NULL,						--使用者IP位置
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_UserOperateLog PRIMARY KEY (LogId)
)

CREATE TABLE BAS.AuthorityUserCompany( /*使用者公司權限*/
 UserId						int					NOT NULL,						--所屬使用者
 CompanyId					int					NOT NULL,						--所屬公司
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 CreateBy					int					NOT NULL,						--新增人員
 CONSTRAINT PK_UserCompanyAuth PRIMARY KEY(UserId, CompanyId),
 CONSTRAINT FK1_UserCompanyAuth FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK2_UserCompanyAuth FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
) /*使用者公司權限*/

CREATE TABLE BAS.AllowUserIp( /* 允許外網登入系統的使用者 */
 AuId						int					identity,				
 UserId						int						NULL,						--所屬使用者					
 IpAddress					nvarchar(150)		NOT NULL,						--使用者IP位置
 Device						nvarchar(150)			NULL,						--使用者裝置
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_AllowUserIp PRIMARY KEY (AuId),
 CONSTRAINT FK1_AllowUserIp FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId)
)

CREATE TABLE BAS.DataConnect( /*連線資訊*/
 DcId						int					identity,
 CompanyId					int					NOT NULL,						--公司別
 CompanyNo					nvarchar(10)		NOT NULL,						--公司代號
 SystemCode					varchar(10)			NOT NULL,						--系統代號
 DataSource					nvarchar(15)		NOT NULL,						--站台Ip
 InitialCatalog				nvarchar(30)		NOT NULL,						--資料庫
 UserID						nvarchar(30)		NOT NULL,						--連線ID
 Password					nvarchar(30)		NOT NULL,						--帳密ID
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員 
 CONSTRAINT PK_DataConnect PRIMARY KEY (DcId),
)


END /*系統設定*/