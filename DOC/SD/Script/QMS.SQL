BEGIN /*基本資料*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DefectGroup') > 0 DROP TABLE QMS.DefectGroup
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DefectClass') > 0 DROP TABLE QMS.DefectClass
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DefectCause' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'QMS')) > 0 DROP TABLE QMS.DefectCause
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'RepairGroup') > 0 DROP TABLE QMS.RepairGroup
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'RepairClass') > 0 DROP TABLE QMS.RepairClass
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'RepairCause') > 0 DROP TABLE QMS.RepairCause
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'QmmDetail' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'QMS')) > 0 DROP TABLE QMS.QmmDetail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'QcMachineMode' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'QMS')) > 0 DROP TABLE QMS.QcMachineMode

END

BEGIN /*品異單*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Abnormalquality'  AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'QMS')) > 0 DROP TABLE QMS.Abnormalquality
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'AqBarcode'  AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'QMS')) > 0 DROP TABLE QMS.AqBarcode
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'AqQcItem'  AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'QMS')) > 0 DROP TABLE QMS.AqQcItem
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'AqFile'  AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'QMS')) > 0 DROP TABLE QMS.AqFile
END

BEGIN /*量測項目測定-NEW*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'QcNotice'  AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'QMS')) > 0 DROP TABLE QMS.QcNotice
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'QcNoticeItemSpec'  AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'QMS')) > 0 DROP TABLE QMS.QcNoticeItemSpec
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'QcNoticeTemplate'  AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'QMS')) > 0 DROP TABLE QMS.QcNoticeTemplate
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'QcNoticeItemSpecTemplate'  AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'QMS')) > 0 DROP TABLE QMS.QcNoticeItemSpecTemplate
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'QcNoticeFile'  AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'QMS')) > 0 DROP TABLE QMS.QcNoticeFile

END

BEGIN /*量測項目測定-NEW*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'QcMeasureData'  AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'QMS')) > 0 DROP TABLE QMS.QcMeasureData
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'TempQcMeasureData'  AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'QMS')) > 0 DROP TABLE QMS.TempQcMeasureData
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'FilePathSetting'  AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'QMS')) > 0 DROP TABLE QMS.FilePathSetting
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'QiDetail'  AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'QMS')) > 0 DROP TABLE QMS.QiDetail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'QcMeasureDataModify'  AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'QMS')) > 0 DROP TABLE QMS.QcMeasureDataModify

END

--*******************************************************************************************************************
BEGIN /*基本資料*/
CREATE TABLE QMS.DefectGroup( /*不良群組資料*/
 GroupId					int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 GroupNo					nvarchar(100)		NOT NULL,						--群組代碼
 GroupName					nvarchar(100)		NOT NULL,						--群組名稱
 GroupDesc					nvarchar(100)		NOT NULL,						--群組描述
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼	
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DefectGroup PRIMARY KEY (GroupId),
 CONSTRAINT AK1_DefectGroup UNIQUE (CompanyId, GroupNo),
 CONSTRAINT FK1_DefectGroup FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)/*不良群組資料*/

CREATE TABLE QMS.DefectClass( /*不良類別資料*/
 ClassId					int					identity,
 GroupId					int					NOT NULL,						--所屬不良群組
 ClassNo					nvarchar(100)		NOT NULL,						--類別代碼
 ClassName					nvarchar(100)		NOT NULL,						--類別名稱
 ClassDesc					nvarchar(100)		NOT NULL,						--類別描述
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼	
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DefectClass PRIMARY KEY (ClassId),
 CONSTRAINT AK1_DefectClass UNIQUE (GroupId, ClassNo),
 CONSTRAINT FK1_DefectClass FOREIGN KEY (GroupId) REFERENCES QMS.DefectGroup(GroupId),
)/*不良類別資料*/

CREATE TABLE QMS.DefectCause( /*不良原因資料*/
 CauseId					int					identity, 
 ClassId					int					NOT NULL,						--所屬不良類別
 CauseNo					nvarchar(100)		NOT NULL,						--原因代碼
 CauseName					nvarchar(100)		NOT NULL,						--原因名稱
 CauseDesc					nvarchar(100)		NOT NULL,						--原因描述
 ResponsibleDepartment		int					NOT NULL	DEFAULT -1,         --權責單位
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼	
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DefectCause PRIMARY KEY (CauseId),
 CONSTRAINT AK1_DefectCause UNIQUE (ClassId, CauseNo),
 CONSTRAINT FK1_DefectCause FOREIGN KEY (ClassId) REFERENCES QMS.DefectClass(ClassId),
)/*不良原因資料*/

CREATE TABLE QMS.RepairGroup( /*維修群組資料*/
 GroupId					int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 GroupNo					nvarchar(100)		NOT NULL,						--群組代碼
 GroupName					nvarchar(100)		NOT NULL,						--群組名稱
 GroupDesc					nvarchar(100)		NOT NULL,						--群組描述
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_RepairGroup PRIMARY KEY (GroupId),
 CONSTRAINT AK1_RepairGroup UNIQUE (CompanyId, GroupNo),
 CONSTRAINT FK1_RepairGroup FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)/*維修群組資料*/

CREATE TABLE QMS.RepairClass( /*維修類別資料*/
 ClassId					int					identity,
 GroupId					int					NOT NULL,						--所屬維修群組
 ClassNo					nvarchar(100)		NOT NULL,						--類別代碼
 ClassName					nvarchar(100)		NOT NULL,						--類別名稱
 ClassDesc					nvarchar(100)		NOT NULL,						--類別描述
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼	
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_RepairClass PRIMARY KEY (ClassId),
 CONSTRAINT AK1_RepairClass UNIQUE (GroupId, ClassNo),
 CONSTRAINT FK1_RepairClass FOREIGN KEY (GroupId) REFERENCES QMS.RepairGroup(GroupId),
)/*維修類別資料*/

CREATE TABLE QMS.RepairCause( /*維修原因資料*/
 CauseId					int					identity, 
 ClassId					int					NOT NULL,						--所屬維修類別
 CauseNo					nvarchar(100)		NOT NULL,						--原因代碼
 CauseName					nvarchar(100)		NOT NULL,						--原因名稱
 CauseDesc					nvarchar(100)		NOT NULL,						--原因描述
 ResponsibleDepartment		int					NOT NULL	DEFAULT -1,         --權責單位
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼	
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_RepairCause PRIMARY KEY (CauseId),
 CONSTRAINT AK1_RepairCause UNIQUE (ClassId, CauseNo),
 CONSTRAINT FK1_RepairCause FOREIGN KEY (ClassId) REFERENCES QMS.RepairClass(ClassId),
)/*維修原因資料*/

CREATE TABLE QMS.QcMachineMode( /*量測機型資料*/
 QcMachineModeId			int					identity, 
 CompanyId					int					NOT NULL,						--所屬公司
 QcMachineModeNumber		nvarchar(10)		NOT NULL,						--量測機型代碼編號
 QcMachineModeNo			nvarchar(50)		NOT NULL,						--量測機型代碼
 QcMachineModeName			nvarchar(50)		NOT	NULL,						--量測機型名稱
 QcMachineModeEnName		nvarchar(100)			NULL,						--量測機型名稱(英文)
 QcMachineModeDesc			nvarchar(50)			NULL,						--量測機型描述
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QcMachineMode PRIMARY KEY (QcMachineModeId),
 CONSTRAINT AK1_QcMachineMode UNIQUE (CompanyId, QcMachineModeName),
 CONSTRAINT FK1_QcMachineMode FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)/*量測機型資料*/

CREATE TABLE QMS.QmmDetail( /*量測機型 - 機台單身*/
 QmmDetailId				int					identity, 
 QcMachineModeId			int					NOT NULL,						--量測機型
 MachineNumber				nvarchar(3)			NOT NULL,						--3碼機台資訊(資料解析用)
 MachineId					int					NOT NULL,						--機台
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼	
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QmmDetail PRIMARY KEY (QmmDetailId),
 CONSTRAINT AK1_QmmDetail UNIQUE (QcMachineModeId, MachineId),
 CONSTRAINT FK1_QmmDetail FOREIGN KEY (QcMachineModeId) REFERENCES QMS.QcMachineMode(QcMachineModeId),
 CONSTRAINT FK2_QmmDetail FOREIGN KEY (MachineId) REFERENCES MES.Machine(MachineId)
)/*量測機型單身 - 機台資料*/

CREATE TABLE QMS.QcGroup( /*量測群組資料*/
 QcGroupId					int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 QcGroupNo					nvarchar(50)		NOT NULL,						--量測群組代碼
 QcGroupName				nvarchar(100)		NOT NULL,						--量測群組名稱
 QcGroupDesc				nvarchar(100)		NOT NULL,						--量測群組描述
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼	
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QcGroup PRIMARY KEY (QcGroupId),
 CONSTRAINT AK1_QcGroup UNIQUE (CompanyId, QcGroupNo),
 CONSTRAINT FK1_QcGroup FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)/*不良群組資料*/

CREATE TABLE QMS.QcClass( /*量測類別資料*/
 QcClassId					int					identity,
 QcGroupId					int					NOT NULL,						--所屬量測群組
 QcClassNo					nvarchar(50)		NOT NULL,						--量測類別代碼
 QcClassName				nvarchar(100)		NOT NULL,						--量測類別名稱
 QcClassDesc				nvarchar(100)		NOT NULL,						--量測類別描述
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼	
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QcClass PRIMARY KEY (QcClassId),
 CONSTRAINT AK1_QcClass UNIQUE (QcGroupId, QcClassNo),
 CONSTRAINT FK1_QcClass FOREIGN KEY (QcGroupId) REFERENCES QMS.QcGroup(QcGroupId),
)/*不良類別資料*/

CREATE TABLE QMS.QcItem( /*量測項目*/
 QcItemId					int					identity, 
 QcClassId					int					NOT NULL,						--所屬量測群組
 QcItemNo					nvarchar(50)		NOT NULL,						--量測項目代碼
 QcItemName					nvarchar(100)		NOT	NULL,						--量測項目名稱
 QcItemDesc					nvarchar(100)		NOT NULL,						--量測項目描述
 QcItemType					nvarchar(2)			NOT NULL,						--量測項目輸入方式 1:一般輸入 2:checkbox
 QcType						nvarchar(10)			NULL,						--檢驗類別 1:IPQC(工程檢) , 2:OQC(出貨檢)
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼	
 Remark						nvarchar(255)			NULL,						--備註
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QcItem PRIMARY KEY (QcItemId),
 CONSTRAINT AK1_QcItem UNIQUE (QcClassId, QcItemNo),
 CONSTRAINT FK1_QcItem FOREIGN KEY (QcClassId) REFERENCES QMS.QcClass(QcClassId)
)/*量測項目*/

CREATE TABLE QMS.QcItemCoding ( /*項目編碼規則管理*/
 QicId						int					identity, 
 CompanyId					int					NOT NULL,						--所屬公司
 QicNo						nvarchar(2)			NOT NULL,						--項目編碼代碼
 QicName					nvarchar(100)		NOT	NULL,						--項目編碼名稱
 QicDesc					nvarchar(100)			NULL,						--項目編碼描述
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼	
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QcItemCoding PRIMARY KEY (QicId),
 CONSTRAINT AK1_QcItemCoding UNIQUE (CompanyId, QicNo),
 CONSTRAINT FK1_QcItemCoding FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)/*項目編碼規則管理*/

CREATE TABLE QMS.QiDetail( /*量測項目綁定機型*/
 QiDetailId					int					identity, 						
 QcItemId					int					NOT NULL,						--所屬量測項目
 QcMachineModeId			int					NOT NULL,						--所屬機型
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QiDetail PRIMARY KEY (QiDetailId),
 CONSTRAINT FK1_QiDetail FOREIGN KEY (QcItemId) REFERENCES QMS.QcItem(QcItemId),
 CONSTRAINT FK2_QiDetail FOREIGN KEY (QcMachineModeId) REFERENCES QMS.QcMachineMode(QcMachineModeId)
)

CREATE TABLE QMS.QaType( /*量測產品別*/
 QaTypeId					int					identity, 
 CompanyId					int					NOT NULL,						--所屬公司
 QaTypeNo					nvarchar(10)		NOT NULL,						--產品別代號
 QaTypeName					nvarchar(50)		NOT	NULL,						--產品別名稱
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼	
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QaType PRIMARY KEY (QaTypeId),
 CONSTRAINT AK1_QaType UNIQUE (CompanyId, QaTypeNo),
 CONSTRAINT AK2_QaType UNIQUE (CompanyId, QaTypeName),
 CONSTRAINT FK1_QaType FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)/*量測產品別*/

CREATE TABLE QMS.UnitOfQcMeasure( /*量測單位基本資料*/
 QcUomId					int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 QcUomNo					nvarchar(50)		NOT NULL,						--單位
 QcUomName					nvarchar(100)		NOT NULL,						--單位名稱
 QcUomDesc					nvarchar(100)			NULL,						--單位描述
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_UnitOfQcMeasure PRIMARY KEY (QcUomId),
 CONSTRAINT FK1_UnitOfQcMeasure FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)

CREATE TABLE QMS.QcItemPrinciple( /*量測項目編碼原則*/
 PrincipleId				int					identity,				
 QcClassId					int					NOT NULL,						--產品類別
 QmmDetailId				int					NOT NULL,						--量測機台ID
 PrincipleNo				nvarchar(4)			NOT NULL,						--編碼NO				
 PrincipleDesc				nvarchar(150)		NOT NULL,						--編碼名稱描述				
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QcItemPrinciple PRIMARY KEY (PrincipleId),
 CONSTRAINT AK1_QcItemPrinciple UNIQUE (QcClassId, QmmDetailId, PrincipleNo),
 CONSTRAINT FK1_QcItemPrinciple FOREIGN KEY (QcClassId) REFERENCES QMS.QcClass(QcClassId),
 CONSTRAINT FK2_QcItemPrinciple FOREIGN KEY (QmmDetailId) REFERENCES QMS.QmmDetail(QmmDetailId),
)/*量測項目編碼原則*/

CREATE TABLE QMS.PrincipleDetail( /*量測項目編碼原則 附加欄位*/
 PdId						int					identity,						
 PrincipleId				int					NOT NULL,						--項目編碼原則ID							
 PrincipleDesc				nvarchar(150)		NOT NULL,						--編碼名稱描述				
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_PrincipleDetail PRIMARY KEY (PdId),
 CONSTRAINT AK1_PrincipleDetail UNIQUE (PrincipleId, PrincipleDesc),
 CONSTRAINT FK1_PrincipleDetail FOREIGN KEY (PrincipleId) REFERENCES QMS.QcItemPrinciple(PrincipleId),
)/*量測項目編碼原則 附加欄位*/

CREATE TABLE QMS.QcpRelationship( /*量測項目編碼原則 別名資料*/
 QcprId						int					identity,				
 PrincipleId				int					NOT NULL,						--量測項目編碼原則ID
 QmmDetailId				int					NOT NULL,						--量測機器ID
 RelationshipItem			nvarchar(150)		NOT NULL,						--關聯項目				
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QcpRelationship PRIMARY KEY (QcprId),
 CONSTRAINT AK1_QcpRelationship UNIQUE (PrincipleId, QmmDetailId, RelationshipItem),
 CONSTRAINT FK1_QcpRelationship FOREIGN KEY (PrincipleId) REFERENCES QMS.QcItemPrinciple(PrincipleId),
 CONSTRAINT FK2_QcpRelationship FOREIGN KEY (QmmDetailId) REFERENCES QMS.QmmDetail(QmmDetailId)
)/*量測項目編碼原則 關聯項目資料*/

CREATE TABLE QMS.QcFileParsing( /*解析量測檔案 參考用表*/
 QfpId						int					identity,				
 QmmDetailId				int					NOT NULL,						--量測機台ID
 FunctionName				nvarchar(150)		NOT NULL,						--解析API名稱	
 Encoding					nvarchar(150)		NOT NULL,						--編碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QcFileParsing PRIMARY KEY (QfpId),
 CONSTRAINT AK1_QcFileParsing UNIQUE (QmmDetailId, FunctionName),
 CONSTRAINT FK1_QcFileParsing FOREIGN KEY (QmmDetailId) REFERENCES QMS.QmmDetail(QmmDetailId)
)/*解析量測檔案 參考用表*/

CREATE TABLE QMS.QcWhitelist( /*量測檔案上傳路徑白名單*/
 ListId						int					identity,
 DepartmentId				int					NOT NULL,						--使用部門ID
 FolderPath					nvarchar(max)		NOT NULL,						--路徑
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QcWhitelist PRIMARY KEY (ListId),
 CONSTRAINT FK1_QcWhitelist FOREIGN KEY (DepartmentId) REFERENCES BAS.Department(DepartmentId),
)/*量測檔案上傳路徑白名單*/
END

BEGIN /*量測數據*/
CREATE TABLE QMS.QcMeasureData( /*量測資料紀錄*/
 QmdId						int					identity,				
 QcRecordId					int					NOT NULL,						--檢驗紀錄
 QcNoticeItemSpecId			int						NULL,						--製程量測項目參數
 QcItemId					int						NULL,						--量測項目ID
 QcItemDesc					nvarchar(500)			NULL,						--量測項目備註
 DesignValue				nvarchar(100)			NULL,						--尺寸設計值
 UpperTolerance				nvarchar(100)			NULL,						--尺寸設計值上限公差
 LowerTolerance				nvarchar(100)			NULL,						--尺寸設計值下限公差
 ZAxis						float					NULL,						--Z軸(通常僅有三次元會有)
 BarcodeId					int						NULL,						--量測條碼ID
 QmmDetailId				int					NOT NULL,						--量測機台
 MeasureValue				nvarchar(500)		NOT NULL,						--量測值
 QcResult					nvarchar(10)		NOT NULL,						--量測結果 P/F
 CauseId					int						NULL,						--異常原因
 CauseDesc					nvarchar(100)			NULL,						--異常原因描述
 Cavity						nvarchar(10)			NULL,						--穴號 (若為鏡片，一穴量多次情況)
 LotNumber					nvarchar(100)			NULL,						--批號
 MakeCount					int						NULL,						--同項目量測次數
 CellHeader					nvarchar(10)			NULL,						--spreadsheet欄位單頭
 [Row]						nvarchar(100)			NULL,						--spreadsheet欄數
 Remark						nvarchar(255)			NULL,						--備註
 QcUserId					int						NULL,						--量測人員
 QcCycleTime				int						NULL,						--量測工時
 BallMark					nvarchar(10)            NULL,                       --球標
 Unit						nvarchar(20)            NULL,                       --單位
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QcMeasureData PRIMARY KEY (QmdId),
 CONSTRAINT FK1_QcMeasureData FOREIGN KEY (QcRecordId) REFERENCES MES.QcRecord(QcRecordId),
 CONSTRAINT FK2_QcMeasureData FOREIGN KEY (QcNoticeItemSpecId) REFERENCES QMS.QcNoticeItemSpec(QcNoticeItemSpecId),
 CONSTRAINT FK3_QcMeasureData FOREIGN KEY (QmmDetailId) REFERENCES QMS.QmmDetail(QmmDetailId),
 CONSTRAINT FK4_QcMeasureData FOREIGN KEY (BarcodeId) REFERENCES MES.Barcode(BarcodeId),
 CONSTRAINT FK5_QcMeasureData FOREIGN KEY (CauseId) REFERENCES QMS.DefectCause(CauseId),
 CONSTRAINT FK6_QcMeasureData FOREIGN KEY (QcItemId) REFERENCES QMS.QcItem(QcItemId),
 CONSTRAINT FK7_QcMeasureData FOREIGN KEY (QcUserId) REFERENCES BAS.[User](UserId)
)/*量測資料紀錄*/

CREATE TABLE QMS.QcMeasureDataTemp( /*量測資料紀錄 Spreadsheet暫存檔*/
 TempId						int					identity,				
 QcRecordId					int					NOT NULL,						--檢驗紀錄
 Cell						nvarchar(10)		NOT NULL,						--欄位
 Css						nvarchar(250)		NOT NULL,						--樣式
 [Format]					nvarchar(45)		NOT NULL,						--資料型態
 [Value]					nvarchar(150)		NOT NULL,						--欄位值
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QcMeasureDataTemp PRIMARY KEY (TempId),
 CONSTRAINT FK1_QcMeasureDataTemp FOREIGN KEY (QcRecordId) REFERENCES MES.QcRecord(QcRecordId),
)/*量測資料紀錄 Spreadsheet暫存檔*/

CREATE TABLE QMS.QcFile( /*量測上傳資料*/
 QcFileId					int					identity,
 FileId						int					NOT NULL,						--檔案ID
 QcNoticeId					int					NOT NULL,						--量測通知單ID
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QcFile PRIMARY KEY (QcFileId),
 CONSTRAINT FK1_QcFile FOREIGN KEY (FileId) REFERENCES BAS.[File](FileId),
 CONSTRAINT FK2_QcFile FOREIGN KEY (QcNoticeId) REFERENCES QMS.QcNotice(QcNoticeId)
)

CREATE TABLE QMS.QcMeasureDataModify( /*量測資料修改後紀錄*/
 QmdmId						int					identity,				
 QmdId						int					NOT NULL,						--量測資料紀錄
 QcRecordId					int					NOT NULL,						--檢驗紀錄
 MoId						int					NOT NULL,						--製令ID
 QcItemId					int					NOT NULL,						--量測項目ID
 QcItemDesc					nvarchar(500)			NULL,						--量測項目備註
 DesignValue				nvarchar(100)			NULL,						--尺寸設計值
 UpperTolerance				nvarchar(100)			NULL,						--尺寸設計值上限公差
 LowerTolerance				nvarchar(100)			NULL,						--尺寸設計值下限公差
 ZAxis						nvarchar(100)			NULL,						--Z軸
 ModifyValue				nvarchar(100)		NOT NULL,						--修改後的量測值
 BarcodeId					int						NULL,						--量測條碼ID
 LetteringSeq				int						NULL,						--刻字流水號
 Cavity						nvarchar(10)			NULL,						--穴號 (若為鏡片，一穴量多次情況)
 LotNumber					nvarchar(100)			NULL,						--批號
 QmmDetailId				int					NOT NULL,						--量測機台
 Remark						nvarchar(255)			NULL,						--備註
 Surveyor					nvarchar(30)			NULL,						--檢測人員姓名
 Confirmer					nvarchar(30)	        NULL,                       --確認人員姓名
 BallMark					nvarchar(10)            NULL,                       --球標
 Unit						nvarchar(20)            NULL,                       --單位
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QcMeasureDataModify PRIMARY KEY (QmdmId),
 CONSTRAINT FK1_QcMeasureDataModify FOREIGN KEY (QcRecordId) REFERENCES MES.QcRecord(QcRecordId),
 CONSTRAINT FK2_QcMeasureDataModify FOREIGN KEY (QmdId) REFERENCES QMS.QcMeasureData(QmdId),
 CONSTRAINT FK3_QcMeasureDataModify FOREIGN KEY (QmmDetailId) REFERENCES QMS.QmmDetail(QmmDetailId),
 CONSTRAINT FK4_QcMeasureDataModify FOREIGN KEY (BarcodeId) REFERENCES MES.Barcode(BarcodeId),
 CONSTRAINT FK5_QcMeasureDataModify FOREIGN KEY (MoId) REFERENCES MES.ManufactureOrder(MoId),
 CONSTRAINT FK6_QcMeasureDataModify FOREIGN KEY (QcItemId) REFERENCES QMS.QcItem(QcItemId)
)/*量測資料修改後紀錄*/

END

BEGIN /*量測機台排程相關*/ 
CREATE TABLE QMS.QcRecordPlanning( /*量測單據排程管理*/
 QrpId						int					identity,				
 QcRecordId					int					NOT NULL,						--量測單號ID
 QcMachineModeId			int					NOT NULL,						--量測機型ID
 TotalQcItemCount			int					NOT NULL	DEFAULT 0,			--總量測項目
 EstimatedMeasurementTime	float				NOT NULL	DEFAULT 0,			--預估量測時間(分鐘)
 ConfirmStatus				nvarchar(1)			NOT NULL	DEFAULT 'N',		--確認狀態 Y:已確認 N:未確認
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QcRecordPlanning PRIMARY KEY (QrpId),
 CONSTRAINT AK_QcRecordPlanning UNIQUE (QcRecordId, QcMachineModeId),
 CONSTRAINT FK1_QcRecordPlanning FOREIGN KEY (QcRecordId) REFERENCES MES.QcRecord(QcRecordId),
 CONSTRAINT FK2_QcRecordPlanning FOREIGN KEY (QcMachineModeId) REFERENCES QMS.QcMachineMode(QcMachineModeId),
)/*量測單據排程管理*/

CREATE TABLE QMS.QcMachineModePlanning( /*量測機台排程*/
 QmmpId						int					identity,				
 QrpId						int					NOT NULL,						--量測單據排程ID
 QmmDetailId				int					NOT NULL,						--量測機台ID
 EstimatedStartDate			datetime				NULL,						--預計開始時間
 EstimatedEndDate			datetime				NULL,						--預計結束時間
 StartDate					datetime				NULL,						--實際開始時間
 EndDate					datetime				NULL,						--實際結束時間
 Sort						int						NULL,						--排序
 [Status]					nvarchar(1)			NOT NULL	DEFAULT 'N',		--排程狀態 N:未開始 A:開工中 F:已完工 P:暫停
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QcMachineModePlanning PRIMARY KEY (QmmpId),
 CONSTRAINT AK1_QcMachineModePlanning UNIQUE (QrpId, QmmDetailId),
 CONSTRAINT FK1_QcMachineModePlanning FOREIGN KEY (QrpId) REFERENCES QMS.QcRecordPlanning(QrpId),
 CONSTRAINT FK2_QcMachineModePlanning FOREIGN KEY (QmmDetailId) REFERENCES QMS.QmmDetail(QmmDetailId),
)/*量測機台排程*/

CREATE TABLE QMS.QcPlanningProcessLog( /*量測排程時間紀錄表*/
 LogId						int					identity,				
 QmmpId						int					NOT NULL,						--量測機台排程ID
 StartDate					datetime				NULL,						--此次量測開始時間
 EndDate					datetime				NULL,						--此次量測結束時間
 [Status]					nvarchar(1)			NOT NULL	DEFAULT 'F',		--觸發狀態 A:開工 P:暫停 F:量測完成
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QcPlanningProcessLog PRIMARY KEY (LogId),
 CONSTRAINT FK1_QcPlanningProcessLog FOREIGN KEY (QmmpId) REFERENCES QMS.QcMachineModePlanning(QmmpId),
)/*量測排程時間紀錄表*/

CREATE TABLE QMS.QcMeasureInTheoryWorkTime( /*量測理論工時*/
 QmwtId                   int                    identity,
 QmmDetailId              int                    NOT NULL,                          --量測機台ID
 ProductType              nvarchar(20)            NOT NULL,                          --品號產品別
 QicId                    int                    NOT NULL,                          --量測編碼ID
 MeasureSize              nvarchar(20)           NULL,								--量測尺吋
 WorkTime                 decimal				 NOT NULL,                          --量測工時(秒)
 CreateDate               datetime				 NOT NULL    DEFAULT GETDATE(),     --新增日期
 LastModifiedDate         datetime               NULL,								--修改日期
 CreateBy                 int                    NOT NULL,                          --新增人員
 LastModifiedBy           int                    NULL,                              --修改人員
 CONSTRAINT PK_QcMeasureInTheoryWorkTime PRIMARY KEY (QmwtId),
 CONSTRAINT FK1_QcMeasureInTheoryWorkTime FOREIGN KEY (QmmDetailId) REFERENCES QMS.QmmDetail(QmmDetailId),
 CONSTRAINT FK2_QcMeasureInTheoryWorkTime FOREIGN KEY (QicId) REFERENCES QMS.QcItemCoding(QicId)
)/*量測理論工時*/
END /*量測機台排程相關*/ 

BEGIN /*品異單*/
CREATE TABLE QMS.Abnormalquality(/*品異單單頭*/
    AbnormalqualityId					int					identity,
	CompanyId							int					NOT NULL,						--所屬公司
	MoId								int						NULL,						--製令
	MoProcessId							int						NULL,						--MES製令Proces
	GrDetailId							int						NULL,						--銷貨單
    AbnormalqualityNo					nvarchar(15)			NULL,						--品異單單號    
	AbnormalqualityStatus				nvarchar(2)			NOT NULL,						--品異單狀態	(1:結單、2:未結單)(需單身狀態全都是 5(會簽)或是0(作廢) 才能算結單)
	QcType								nvarchar(10)		NOT NULL,						--檢驗類別 1:IPQC(工程檢) , 2:OQC(出貨檢), 3 IQC(進貨檢), 4 NON(過站進階開立)
	DocDate								datetime			NOT NULL,	 					--開單日期
	CreateDate							datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
	LastModifiedDate					datetime				NULL,						--修改日期
	CreateBy							int					NOT NULL,						--新增人員
	LastModifiedBy						int						NULL,						--修改人員
	CONSTRAINT PK_Abnormalquality PRIMARY KEY (AbnormalqualityId),
	CONSTRAINT AK1_Abnormalquality UNIQUE (CompanyId, AbnormalqualityNo),
	CONSTRAINT FK1_Abnormalquality FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
	CONSTRAINT FK2_Abnormalquality FOREIGN KEY (MoId) REFERENCES MES.ManufactureOrder(MoId),
	CONSTRAINT FK3_Abnormalquality FOREIGN KEY (MoProcessId) REFERENCES MES.MoProcess(MoProcessId),
	CONSTRAINT FK4_Abnormalquality FOREIGN KEY (GrDetailId) REFERENCES SCM.GrDetail(GrDetailId)
)

CREATE TABLE QMS.AqBarcode/*品異單異常條碼*/
(
	AqBarcodeId							int					identity,
    AbnormalqualityId					int					NOT NULL,						--品異單ID
    BarcodeId							int						NULL,						--條碼ID
	SupplierId							int						NULL,						--供應商(進貨撿用)
	QcRecordId							int						NULL,						--MES.QcRecord
	QcBarcodeId							int						NULL,						--工程檢/出貨檢記錄 ID -MES.BarcodeQcRecord
	DefectCauseId						int						NULL,						--不良代碼    (後續可以更改) - QMS.Defect CauseId
    DefectCauseDesc						nvarchar(100)			NULL,						--不良原因	 (預設帶不良原因的描述,後續可以更改) - QMS.Defect CauseDesc
	ConformUserId						int						NULL,						--合致對象	 (不一定會有所以NULL)
    ResponsibleDeptId					int					NOT NULL,						--責任單位	  
    ResponsibleUserId					int					NOT NULL,						--責任者		 
	SubResponsibleDeptId				int						NULL,						--副責任單位   (不一定會有所以NULL) 
    SubResponsibleUserId				int						NULL,						--副責任者	 (不一定會有所以NULL)
    ProgrammerUserId					int						NULL,						--編程者		 (不一定會有所以NULL)
    RepairCauseId						int						NULL,						--對策代碼	 (因後續更新資料才有資料所以NULL) - QMS.RepairCause CauseId
    RepairCauseDesc						nvarchar(100)			NULL,						--對策敘述	 (因後續更新資料才有資料所以NULL) - QMS.RepairCause CauseDesc
	RepairCauseUserId					int						NULL,						--編寫對策者　　(因後續更新資料才有資料所以NULL)
    ResponsibleSupervisorId				int						NULL,						--責任主管	 (因後續更新資料才有資料所以NULL)
	JudgeStatus							nvarchar(20)			NULL,						--判定結果    (S:Scrap[報廢]、RW:Repair[返修]、M:Misjudgment[誤判]、R:Release[特採]、SR:SituRepair[原站返修]、AM:AgainMeasure[重新量測]、NK:[不良品\流程更換])
    JudgeDesc							nvarchar(300)			NULL,						--判定原因	 (因後續更新資料才有資料所以NULL)
    JudgeUserId							int						NULL,						--判定者		 (因後續更新資料才有資料所以NULL)
    JudgeReturnMoProcessId				int						NULL,						--判定返修站別 (因後後續更新資料才有資料所以NULL)
    JudgeReturnNextMoProcessId			int						NULL,						--判定返修後的下一站 (因後後續更新資料才有資料所以NULL)
	JudgeConfirm						nvarchar(2)			NOT NULL,						--判定確認    V:舊資料 N:未判定 Y:已判定
	JudgeDate							datetime				NULL,						--判定日
	ChangeJudgeFlag						nvarchar(1)				NULL,						--是否改判    N:無 Y:是
	CountersignUserId					int						NULL,						--會簽者		 (因後續更新資料才有資料所以NULL)
	AqBarcodeStatus						nvarchar(2)			NOT NULL,						--異常條碼狀態 (1:異常回報完成、2:異常對策完成、3:異常對策確認、4:異常判定完成、5:品異單會簽完成、6:強制會簽、7:異常判定確認完成、0:作廢) (5是正常跑完所有流程,6可能急單前面都沒資料事後要補,7預加不用做頁面)
	ProcessStatus						nvarchar(2)			NOT NULL	DEFAULT 'N',		--當ProcessStatus = 'Y'時, 代表判定返修結果已執行完成, 當ProcessStatus = 'I'時, 代表判定返修進行中, 'V'時,代表上一個判定沒有製型作廢,以新開品異單為主,'N' 無
	ReturnQty							int						NULL,						--驗退數輛(IQC)
	CreateDate							datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
	LastModifiedDate					datetime				NULL,						--修改日期
	CreateBy							int					NOT NULL,						--新增人員
	LastModifiedBy						int						NULL,						--修改人員
	CONSTRAINT PK_AqBarcode PRIMARY KEY (AqBarcodeId),
	CONSTRAINT AK1_AqBarcode UNIQUE (AbnormalqualityId, BarcodeId),
	CONSTRAINT FK1_AqBarcode FOREIGN KEY (AbnormalqualityId) REFERENCES  QMS.Abnormalquality(AbnormalqualityId),
	CONSTRAINT FK2_AqBarcode FOREIGN KEY (BarcodeId) REFERENCES  MES.Barcode(BarcodeId),
	CONSTRAINT FK3_AqBarcode FOREIGN KEY (QcBarcodeId) REFERENCES  MES.QcBarcode(QcBarcodeId),
	CONSTRAINT FK4_AqBarcode FOREIGN KEY (DefectCauseId) REFERENCES  QMS.DefectCause(CauseId),
	CONSTRAINT FK5_AqBarcode FOREIGN KEY (ConformUserId) REFERENCES BAS.[User](UserId),
	CONSTRAINT FK6_AqBarcode FOREIGN KEY (ResponsibleDeptId) REFERENCES BAS.Department(DepartmentId),
	CONSTRAINT FK7_AqBarcode FOREIGN KEY (ResponsibleUserId) REFERENCES BAS.[User](UserId),
	CONSTRAINT FK8_AqBarcode FOREIGN KEY (SubResponsibleDeptId) REFERENCES BAS.Department(DepartmentId),
	CONSTRAINT FK9_AqBarcode FOREIGN KEY (SubResponsibleUserId) REFERENCES BAS.[User](UserId),
	CONSTRAINT FK10_AqBarcode FOREIGN KEY (ProgrammerUserId) REFERENCES BAS.[User](UserId),
	CONSTRAINT FK11_AqBarcode FOREIGN KEY (RepairCauseId) REFERENCES QMS.RepairCause(CauseId),
	CONSTRAINT FK12_AqBarcode FOREIGN KEY (RepairCauseUserId) REFERENCES BAS.[User](UserId),
	CONSTRAINT FK13_AqBarcode FOREIGN KEY (ResponsibleSupervisorId) REFERENCES BAS.[User](UserId),
	CONSTRAINT FK14_AqBarcode FOREIGN KEY (JudgeUserId) REFERENCES BAS.[User](UserId),
	CONSTRAINT FK15_AqBarcode FOREIGN KEY (JudgeReturnMoProcessId) REFERENCES MES.MoProcess(MoProcessId),
	CONSTRAINT FK16_AqBarcode FOREIGN KEY (JudgeReturnNextMoProcessId) REFERENCES MES.MoProcess(MoProcessId),
	CONSTRAINT FK17_AqBarcode FOREIGN KEY (CountersignUserId) REFERENCES BAS.[User](UserId),
	CONSTRAINT FK18_AqBarcode FOREIGN KEY (QcRecordId) REFERENCES MES.QcRecord(QcRecordId)
)

CREATE TABLE QMS.AqQcItem/*品異條碼異常原因對策表*/
(
	AqQcItemId							int					identity,
    AqBarcodeId							int					NOT NULL,						--品異單異常條碼ID
    QcItemId							int					NOT NULL,						--檢測項目ID
    DefectCauseId						int					NOT NULL,						--不良代碼    (後續可以更改) - QMS.Defect CauseId
    DefectCauseDesc						nvarchar(100)			NULL,						--不良原因	 (預設帶不良原因的描述,後續可以更改) - QMS.Defect CauseDesc
    RepairCauseId						int						NULL,						--對策代碼	 (因後續更新資料才有資料所以NULL) - QMS.RepairCause CauseId
    RepairCauseDesc						nvarchar(100)			NULL,						--對策敘述	 (因後續更新資料才有資料所以NULL) - QMS.RepairCause CauseDesc
	CreateDate							datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
	LastModifiedDate					datetime				NULL,						--修改日期
	CreateBy							int					NOT NULL,						--新增人員
	LastModifiedBy						int						NULL,						--修改人員 
	CONSTRAINT PK_AqQcItem PRIMARY KEY (AqQcItemId),
	CONSTRAINT AK1_AqQcItem UNIQUE(AqBarcodeId, QcItemId),
	CONSTRAINT FK1_AqQcItem FOREIGN KEY (AqBarcodeId) REFERENCES   QMS.AqBarcode(AqBarcodeId),
	CONSTRAINT FK2_AqQcItem FOREIGN KEY (QcItemId) REFERENCES   QMS.QcItem(QcItemId),
	CONSTRAINT FK3_AqQcItem FOREIGN KEY (DefectCauseId) REFERENCES   QMS.DefectCause(CauseId),
	CONSTRAINT FK4_AqQcItem FOREIGN KEY (RepairCauseId) REFERENCES   QMS.RepairCause(CauseId),
)

CREATE TABLE QMS.AqFile/*品異條碼原因佐證檔案*/
(
	AqFileId							int					identity,
    AqBarcodeId							int					NOT NULL,						--品異單異常條碼ID
	AqFileStatus						nvarchar(2)			NULL,							--上傳檔案階段(1:回報)
    FileId								int					NULL,							--檔案ID
	CreateDate							datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
	LastModifiedDate					datetime			NULL,							--修改日期
	CreateBy							int					NOT NULL,						--新增人員
	LastModifiedBy						int					NULL,							--修改人員 
	CONSTRAINT PK_AqFile PRIMARY KEY (AqFileId),
	CONSTRAINT FK1_AqFile FOREIGN KEY (AqBarcodeId) REFERENCES   QMS.AqBarcode(AqBarcodeId),
	CONSTRAINT FK2_AqFile FOREIGN KEY (FileId) REFERENCES   BAS.[File](FileId)
)

CREATE TABLE QMS.AqPhrase/*品異常用片語*/
(
	AqpId								int					identity,
	CompanyId							int					NOT NULL,						--所屬公司
	--UserFrom							nvarchar(2)			NOT NULL,						--使用來源
    PhraseText							nvarchar(300)		NOT NULL,						--片語
	CreateDate							datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
	LastModifiedDate					datetime			NULL,							--修改日期
	CreateBy							int					NOT NULL,						--新增人員
	LastModifiedBy						int					NULL,							--修改人員 
	CONSTRAINT PK_AqPhrase PRIMARY KEY (AqpId),
	CONSTRAINT FK1_AqPhrase FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
)


END

BEGIN /*量測項目測定*/

CREATE TABLE QMS.QcNotice( /*製程量測需求單單頭*/
 QcNoticeId					int					identity,				
 QcNoticeNo					nvarchar(50)		NOT NULL,						--製程參數資料
 QcNoticeType				nvarchar(10)		NOT NULL,						--單據類型--(SELECT * FROM BAS.[Type] WHERE TypeSchema = 'QMS') 1:IPQC(工程檢) , 2:OQC(出貨檢), 3:IQC(進料檢)
 QcNoticeQty				int					NOT NULL,						--送測數量
 DepartmentId               int					NOT NULL,						--送測部門ID
 ControlId					int					NOT NULL    DEFAULT -1,			--圖面管理 ID --SELECT ControlId FROM PDM.RdDesignControl
 RoutingProcessId			int					NOT NULL    DEFAULT -1,			--製程加工 ID --SELECT ItemProcessId FROM MES.RoutingItemProcess
 MoId						int					NOT NULL    DEFAULT -1,			--MES製令ID
 Remark						nvarchar(255)			NULL,						--備註
 ResetFlag					nvarchar(2)			NOT NULL    DEFAULT 'Y',		--過站後是否要將通知單狀態改為已結案(C)	 >> Y:過站後改為C N:過站後維持Y(可重複使用)
 BarcodeProcessFlag			nvarchar(2)			NOT NULL	DEFAULT 'Y',		--是否要條碼控管
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'N',		--狀態碼		   >> N:未完成, Y:已完成,A:量測中,C:結案(下一站完工)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QcNotice PRIMARY KEY (QcNoticeId),
 CONSTRAINT AK1_QcNotice UNIQUE (QcNoticeNo),
 CONSTRAINT FK1_QcNotice FOREIGN KEY (DepartmentId) REFERENCES BAS.Department(DepartmentId)
)/*製程量測需求單單頭*/

CREATE TABLE QMS.QcNoticeItemSpec( /*製程量測需求單-量測項目清單*/
 QcNoticeItemSpecId			int					identity,				
 QcNoticeId					int					NOT NULL,						--製程量測需求單單頭ID
 QcItemId					int					NOT NULL,						--量測項目(可串接機型)
 DesignValue				float				NOT NULL,						--尺寸設計值
 UpperTolerance				float				NOT NULL,						--尺寸設計值上限公差
 LowerTolerance				float				NOT NULL,						--尺寸設計值下限公差
 MakeCount					int					NOT NULL    DEFAULT 1,			--項目量測次數
 ProductQcCount				int					NOT NULL    DEFAULT 1,			--產品量測次數
 QcUomId					int					NOT NULL,						--量測單位
 Depth						int						NULL,						--Z值 
 Remark						nvarchar(255)			NULL,						--備註
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼		   >> S:關閉 ; A:啟用	
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QcNoticeItemSpec PRIMARY KEY (QcNoticeItemSpecId),
 CONSTRAINT FK1_QcNoticeItemSpec FOREIGN KEY (QcNoticeId) REFERENCES QMS.QcNotice(QcNoticeId),
 CONSTRAINT FK2_QcNoticeItemSpec FOREIGN KEY (QcItemId) REFERENCES QMS.QcItem(QcItemId),
 CONSTRAINT FK3_QcNoticeItemSpec FOREIGN KEY (QcUomId) REFERENCES QMS.UnitOfQcMeasure(QcUomId)
)/*製程量測需求單-量測項目清單*/

CREATE TABLE QMS.QcNoticeFile( /*製程量測需求單-上傳檔案資料*/
 QcNoticeFileId				 int					identity,				
 QcNoticeId					 int					NOT NULL,						--製程量測需求單單頭ID
 FileId						 int					NOT NULL,						--FileId
 [Status]					 nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼		   >> S:關閉 ; A:啟用	
 CreateDate					 datetime				NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			 datetime				NULL,							--修改日期
 CreateBy					 int					NOT NULL,						--新增人員
 LastModifiedBy				 int					NULL,							--修改人員
 CONSTRAINT PK_QcNoticeFile  PRIMARY KEY (QcNoticeFileId),
 CONSTRAINT FK1_QcNoticeFile FOREIGN KEY (QcNoticeId) REFERENCES QMS.QcNotice(QcNoticeId)
)/*製程量測需求單-上傳檔案資料*/

CREATE TABLE QMS.QcNoticeTemplate( /*製程量測需求單單頭-樣板*/
 QcNoticeTemplateId			int					identity,
 QcNoticeTemplateNo			nvarchar(50)		NOT NULL,						--樣板編號
 QcNoticeTemplateName		nvarchar(50)		NOT NULL,						--樣板名稱
 QcNoticeTemplateDesc		nvarchar(50)		NOT NULL,						--樣板描述
 QcNoticeType				nvarchar(10)		NOT NULL,						--量測項目--(SELECT * FROM BAS.[Type] WHERE TypeSchema = 'QMS';)  
 ParameterId				int					NOT NULL    DEFAULT -1,			--MES製程參數ID
 Remark						nvarchar(255)			NULL,						--備註
 CompanyId					int					NOT NULL,						--公司別
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼		   >> S:關閉 ; A:啟用	
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QcNoticeTemplate PRIMARY KEY (QcNoticeTemplateId),
 CONSTRAINT FK1_QcNoticeTemplate FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)/*製程量測需求單單頭-樣板*/

CREATE TABLE QMS.QcNoticeItemSpecTemplate( /*製程量測需求單-量測項目清單-樣板*/
 QcNoticeItemSpecTemplateId			int					identity,				
 QcNoticeTemplateId					int					NOT NULL,						--製程量測需求單單頭-樣板ID
 QcItemId							int					NOT NULL,						--量測項目(可串接機型)
 DesignValue						float				NOT NULL,						--尺寸設計值
 UpperTolerance						float				NOT NULL,						--尺寸設計值上限公差
 LowerTolerance						float				NOT NULL,						--尺寸設計值下限公差
 MakeCount							int					NOT NULL    DEFAULT 1,			--項目量測次數
 ProductQcCount						int					NOT NULL    DEFAULT 1,			--產品量測次數
 QcUomId							int					NOT NULL,						--量測單位
 Depth								int						NULL,						--Z值
 Remark								nvarchar(255)			NULL,						--備註
 [Status]							nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼		   >> S:關閉 ; A:啟用	
 CreateDate							datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate					datetime				NULL,						--修改日期
 CreateBy							int					NOT NULL,						--新增人員
 LastModifiedBy						int						NULL,						--修改人員
 CONSTRAINT PK_QcNoticeItemSpecTemplate PRIMARY KEY (QcNoticeItemSpecTemplateId),
 CONSTRAINT FK2_QcNoticeItemSpecTemplate FOREIGN KEY (QcItemId) REFERENCES QMS.QcItem(QcItemId),
 CONSTRAINT FK3_QcNoticeItemSpecTemplate FOREIGN KEY (QcUomId) REFERENCES QMS.UnitOfQcMeasure(QcUomId)
)/*製程量測需求單-量測項目清單-樣板*/

CREATE TABLE QMS.FilePathSetting( /*量測檔案路徑設定資料*/
 FilePathSettingId			int					identity,				
 QcMachineModeId			int					NOT NULL,						--量測機型ID
 ItemNo						nvarchar(50)		NOT NULL,						--項目名稱
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼		   >> S:關閉 ; A:啟用	
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_FilePathSetting PRIMARY KEY (FilePathSettingId),
 CONSTRAINT AK1_FilePathSetting UNIQUE (FilePathSettingId, QcMachineModeId, ItemNo),
 CONSTRAINT FK1_FilePathSetting FOREIGN KEY (QcMachineModeId) REFERENCES QMS.QcMachineMode(QcMachineModeId),
)/*量測檔案路徑設定資料*/

CREATE TABLE QMS.QcType( /*量測類型 基本資料設定檔*/
 QcTypeId					int					identity,				
 ModeId						int					NOT NULL,						--生產模式
 QcTypeNo					nvarchar(50)		NOT NULL,						--量測類型No
 QcTypeName					nvarchar(50)		NOT NULL,						--量測類型名稱
 SupportAqFlag				nvarchar(2)			NOT NULL    DEFAULT 'Y',		--是否支援建立品異單
 SupportProcessFlag			nvarchar(2)			NOT NULL    DEFAULT 'N',		--是否支援卡控製程站別
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼 A:啟用 S:停用	
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QcType PRIMARY KEY (QcTypeId),
 CONSTRAINT AK1_QcType UNIQUE (ModeId, QcTypeNo),
 CONSTRAINT FK1_QcType FOREIGN KEY (ModeId) REFERENCES MES.ProdMode(ModeId),
)/*量測類型 基本資料設定檔*/

END 

BEGIN /*客訴系統*/

CREATE TABLE QMS.MenuCategory( /*客訴選單類別*/
 McId						int					identity,	
 CompanyId					int					NOT NULL,						--公司別 
 McType						nvarchar(5)			NOT NULL,						--選單類型 --BAS.[Type]維護=> D2:問題分類、D3:處理分類、D4:原因分類、D5:長期對策分類、D6:、效果驗證分類 、D7:文件分類、D8:追蹤模式
 McNo						nvarchar(20)		NOT NULL,						--客訴選單代碼
 McName						nvarchar(100)		NOT NULL,						--客訴選單名稱
 McDesc						nvarchar(100)			NULL,						--客訴選單描述
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼>> S:關閉 ; A:啟用	
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MenuCategory PRIMARY KEY (McId),
 CONSTRAINT FK1_MenuCategory FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)/*量測檔案路徑設定資料*/

CREATE TABLE QMS.CustomerComplaint( /*客訴單據資訊*/ 
 CcId						int					identity,
 CompanyId					int					NOT NULL,						--公司別 
 CcNo						nvarchar(50)		NOT NULL,						--客訴編號	>> 客訴代碼【C】_西元年末碼YY+2碼流水號01，ex:C230701	
 MtlItemId					int					NOT NULL,						--品號ID		>> ??(待確認-產品序號欄位)
 ProductionSource			nvarchar(100)			NULL,						--生產/來料單位
 FilingDate					datetime				NULL,						--生產日期
 Quantity					int						NULL,						--數量
 DocDate					datetime				NULL,						--客訴日期(單據日期)
 ReplyDate					datetime				NULL,						--回覆日
 ClosureDate				datetime				NULL,						--結案日
 CustomerId					int					NOT	NULL,						--ERP客戶代碼
 UserId						int						NULL,						--聯絡人(保留欄位)
 UserName					nvarchar(20)		NOT	NULL,						--聯絡人名稱
 UserEMail					nvarchar(100)		NOT	NULL,						--信箱
 CaseUserId					int						NULL,						--承認者
 ReviewUserId				int						NULL,						--審核者
 WriteUserId				int						NULL,						--做成者
 IssueDescription			nvarchar(MAX)		NOT NULL,						--問題說明   >> 限 250字
 D1ShowSetting				nvarchar(2)			NOT NULL    DEFAULT 'A', 		--D1流程卡顯示狀態碼	 >> S:關閉 ; A:啟用	
 D2ShowSetting				nvarchar(2)			NOT NULL    DEFAULT 'A',		--D2流程卡顯示狀態碼	 >> S:關閉 ; A:啟用	
 D3ShowSetting				nvarchar(2)			NOT NULL    DEFAULT 'A',		--D3流程卡顯示狀態碼	 >> S:關閉 ; A:啟用	
 D4ShowSetting				nvarchar(2)			NOT NULL    DEFAULT 'A',		--D4流程卡顯示狀態碼	 >> S:關閉 ; A:啟用	
 D5ShowSetting				nvarchar(2)			NOT NULL    DEFAULT 'A',		--D5流程卡顯示狀態碼	 >> S:關閉 ; A:啟用	
 D6ShowSetting				nvarchar(2)			NOT NULL    DEFAULT 'A',		--D6流程卡顯示狀態碼	 >> S:關閉 ; A:啟用	
 D7ShowSetting				nvarchar(2)			NOT NULL    DEFAULT 'A',		--D7流程卡顯示狀態碼	 >> S:關閉 ; A:啟用	
 D8ShowSetting				nvarchar(2)			NOT NULL    DEFAULT 'A',		--D8流程卡顯示狀態碼	 >> S:關閉 ; A:啟用	
 ChangeCustomerCcNo			nvarchar(50)			NULL,						--新客訴編號(原單據要更換客戶走單據已作廢創建新單據流程)
 ConfirmStatus				nvarchar(1)			NOT NULL,						--確認碼 >> BAS.[Status] N：未確認；V：作廢；Y：已確認
 ConfirmDate				datetime				NULL,						--確認日期
 CurrentStatus				nvarchar(2)			NOT NULL    DEFAULT 'S',		--當下狀態	BAS.[Type]維護>>D0(完成建單)=>D1資料維護中=>D2=>D3=>D4=>D5=>D6=>D7=>D8
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員 
 CONSTRAINT PK_CustomerComplaint PRIMARY KEY (CcId),
 CONSTRAINT FK1_CustomerComplaint FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_CustomerComplaint FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId),
 CONSTRAINT FK3_CustomerComplaint FOREIGN KEY (CustomerId) REFERENCES SCM.Customer(CustomerId)
)/*客訴單據資訊*/

CREATE TABLE QMS.CcFile( /*客訴單據佐證檔案*/
 CcFileId							int					identity,
 CcId								int					NOT NULL,						--客訴單據ID
 FileType							nvarchar(5)			NOT NULL,						--選單類型 --廠內,客戶
 FileId								int					NULL,							--檔案ID
 CreateDate							datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate					datetime			NULL,							--修改日期
 CreateBy							int					NOT NULL,						--新增人員
 LastModifiedBy						int					NULL,							--修改人員 
 CONSTRAINT PK_CcFile PRIMARY KEY (CcFileId),
 CONSTRAINT FK1_CcFile FOREIGN KEY (CcId) REFERENCES   QMS.CustomerComplaint(CcId),
 CONSTRAINT FK2_CcFile FOREIGN KEY (FileId) REFERENCES   BAS.[File](FileId)
)

CREATE TABLE QMS.CcTeamMembers( /*客訴單據-小組成員D1*/
 CcTeamMembersId					int					identity,
 CcId								int					NOT NULL,						--客訴單據ID
 UserId								int					NOT NULL,						--成員ID
 ConfirmStatus						nvarchar(1)				NULL,						--確認碼 >> BAS.[Status] N：未確認；V：作廢；Y：已確認
 CreateDate							datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate					datetime				NULL,						--修改日期
 CreateBy							int					NOT NULL,						--新增人員
 LastModifiedBy						int						NULL,						--修改人員 
 CONSTRAINT PK_CcTeamMembers PRIMARY KEY (CcTeamMembersId),
 CONSTRAINT FK1_CcTeamMembers FOREIGN KEY (CcId) REFERENCES   QMS.CustomerComplaint(CcId),
 CONSTRAINT FK2_CcTeamMembers FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId)
)


CREATE TABLE QMS.CcStageData( /*客訴單據-各階段資料*/
 CcStageDataId						int					identity,
 CcId								int					NOT NULL,						--客訴單據ID
 Stage								nvarchar(5)			NOT NULL,						--階段 -- D2:問題陳述、D3:矯正處置、D4:根本原因、D5:長期對策、D6:效果驗證、D7:、再發防止措施、D8:追蹤與結案
 McId								int						NULL,						--選單分類 --	 階段D2:問題分類 、階段D3:處理分類 、階段D4:原因分類 、階段D7:文件分類、階段D8:追蹤模式 、其他為NULL
 DescriptionTextarea				nvarchar(MAX)		NOT NULL,						--說明
 DeptId								int					NOT NULL,						--擔當單位	  
 UserId								int					NOT NULL,						--擔當人員
 ConfirmStatus						nvarchar(1)				NULL,						--確認碼 >> BAS.[Status] N：未確認；V：作廢；Y：已確認
 ConfirmDate						datetime				NULL,						--確認日期
 ExpectedDate						datetime				NULL,						--預計完成日期
 FinishDate							datetime				NULL,						--完成日期
 CreateDate							datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate					datetime				NULL,						--修改日期
 CreateBy							int					NOT NULL,						--新增人員
 LastModifiedBy						int						NULL,						--修改人員 
 CONSTRAINT PK_CcStageData PRIMARY KEY (CcStageDataId),
 CONSTRAINT AK1_PK_CcStageData UNIQUE (CcId, Stage, McId),
 CONSTRAINT FK1_CcStageData FOREIGN KEY (CcId) REFERENCES   QMS.CustomerComplaint (CcId),
 CONSTRAINT FK2_CcStageData FOREIGN KEY (DeptId) REFERENCES BAS.Department(DepartmentId),
 CONSTRAINT FK3_CcStageData FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId)
)

CREATE TABLE QMS.CcStageDataFile( /*客訴單據-個階段資料佐證檔案*/
 CcStageDataFileId					int					identity,
 CcStageDataId						int					NOT NULL,						--個階段資料ID
 FileId								int					NULL,							--檔案ID
 CreateDate							datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate					datetime			NULL,							--修改日期
 CreateBy							int					NOT NULL,						--新增人員
 LastModifiedBy						int					NULL,							--修改人員 
 CONSTRAINT PK_CcProblemDesFile PRIMARY KEY (CcStageDataFileId),
 CONSTRAINT FK1_CcProblemDesFile FOREIGN KEY (CcStageDataId) REFERENCES   QMS.CcStageData(CcStageDataId),
 CONSTRAINT FK2_CcProblemDesFile FOREIGN KEY (FileId) REFERENCES   BAS.[File](FileId)
)

CREATE TABLE QMS.CCTrackingData( /*客訴單據-D8欄位表格資料*/
 CCTrackingDataId					int					identity,
 CcStageDataId						int					NOT NULL,						--個階段資料ID
 Batch								int					NULL,							--批數(1-3)
 ReportNo						    nvarchar(255)		NOT NULL,						--報告編號
 TrialResult						nvarchar(255)		NOT NULL,						--檢測結果
 TrackingDate						datetime			NOT NULL	DEFAULT GETDATE(),	--追蹤日期
 CreateDate							datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate					datetime			NULL,							--修改日期
 CreateBy							int					NOT NULL,						--新增人員
 LastModifiedBy						int					NULL,							--修改人員 
 CONSTRAINT PK_CCTrackingData PRIMARY KEY (CCTrackingDataId),
 CONSTRAINT FK1_CCTrackingData FOREIGN KEY (CcStageDataId) REFERENCES   QMS.CcStageData(CcStageDataId)
 )

END/*客訴系統*/


CREATE TABLE MES.WoModification( /*ERP製令變更單頭-MOCTO*/
 WomId								int					identity,
 WoId								int					NOT NULL,						--製令ID  TO001/TO002
 ModifyVersion						nvarchar(4)			NOT NULL,						--變更版次 TO003 
 ModifyDate							datetime			NOT NULL,						--變更日期 TO004
 ModifyCause						nvarchar(255)		NOT NULL,						--變更原因 TO005
 TrackingDate						datetime			NOT NULL	DEFAULT GETDATE(),	--追蹤日期
 CreateDate							datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate					datetime				NULL,						--修改日期
 CreateBy							int					NOT NULL,						--新增人員
 LastModifiedBy						int						NULL,						--修改人員 
 CONSTRAINT PK_WoModification PRIMARY KEY (WomId),
 CONSTRAINT FK1_WoModification FOREIGN KEY (WoId) REFERENCES   MES.WipOrder(WoId)
 )
 CREATE TABLE MES.WomDetail( /*ERP製令變更單頭-MOCTP*/
 WomDetailId						int					identity,
 WomId								int					NOT NULL,						--製令變更單ID

 CreateDate							datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate					datetime				NULL,						--修改日期
 CreateBy							int					NOT NULL,						--新增人員
 LastModifiedBy						int						NULL,						--修改人員 
 CONSTRAINT PK_WomDetail PRIMARY KEY (WomDetailId),
 CONSTRAINT FK1_WomDetail FOREIGN KEY (WomId) REFERENCES   MES.WoModification(WomId)
 )




