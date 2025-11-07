BEGIN /*基本資料*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MassProductionReviewDocSetting') > 0 DROP TABLE PDM.MassProductionReviewDocSetting /*20240806*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MtlPrincipleBuildData') > 0 DROP TABLE PDM.MtlPrincipleBuildData
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MtlPrincipleCondition') > 0 DROP TABLE PDM.MtlPrincipleCondition
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MtlPrincipleBuild') > 0 DROP TABLE PDM.MtlPrincipleBuild
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MtlPrinciple') > 0 DROP TABLE PDM.MtlPrinciple
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MtlElement') > 0 DROP TABLE PDM.MtlElement
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MtlModel') > 0 DROP TABLE PDM.MtlModel
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'UomCalculate') > 0 DROP TABLE PDM.UomCalculate
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'UnitOfMeasure') > 0 DROP TABLE PDM.UnitOfMeasure
END /*基本資料*/

BEGIN /*圖面管理*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'RdDesignAttribute') > 0 DROP TABLE PDM.RdDesignAttribute
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'RdAttributeItem') > 0 DROP TABLE PDM.RdAttributeItem
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'RdDesignControlFile') > 0 DROP TABLE PDM.RdDesignControlFile
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'RdDesignControl') > 0 DROP TABLE PDM.RdDesignControl
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'RdDesign') > 0 DROP TABLE PDM.RdDesign
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'CustomerCadControlFile') > 0 DROP TABLE PDM.CustomerCadControlFile
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'CustomerCadControl') > 0 DROP TABLE PDM.CustomerCadControl
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'CustomerCad') > 0 DROP TABLE PDM.CustomerCad
END/*圖面管理*/

BEGIN /*產品資料管理*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MassProductionReviewDoc') > 0 DROP TABLE PDM.MassProductionReviewDoc /*20240806*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'BomSubstitution') > 0 DROP TABLE PDM.BomSubstitution
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'BomDetail') > 0 DROP TABLE PDM.BomDetail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'BillOfMaterials') > 0 DROP TABLE PDM.BillOfMaterials
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'CustomerMtlItem') > 0 DROP TABLE PDM.CustomerMtlItem
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MtlItem') > 0 DROP TABLE PDM.MtlItem
END /*產品資料管理*/

BEGIN /*品號變更管理*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MtlItemChange') > 0 DROP TABLE PDM.MtlItemChange  --INVTL
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MtlItemChangeDetail') > 0 DROP TABLE PDM.MtlItemChangeDetail  --INVTM
END /*品號變更管理*/

BEGIN /*Bom變更管理*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'BomChange') > 0 DROP TABLE PDM.BomChange  --BOMTA
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'BomChangeDetail') > 0 DROP TABLE PDM.BomChangeDetail  --BOMTB
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'BomChangeSubDetail') > 0 DROP TABLE PDM.BomChangeSubDetail  --BOMTC
END /*Bom變更管理*/

BEGIN /*允收標準*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MtlQcItem') > 0 DROP TABLE PDM.MtlQcItem  --MtlQcItem
END /*允收標準*/

--******************************************************************************************************************
BEGIN /*基本資料*/
CREATE TABLE PDM.MtlElement( /*品號元素*/
 ElementId					int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 ElementNo					nvarchar(50)		NOT NULL,						--元素代碼
 ElementName				nvarchar(100)		NOT NULL,						--元素名稱
 SortNumber					int					NOT NULL,						--排序號碼
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MtlElement PRIMARY KEY (ElementId),
 CONSTRAINT AK1_MtlElement UNIQUE (ElementNo),
 CONSTRAINT AK2_MtlElement UNIQUE (ElementName),
 CONSTRAINT FK1_MtlElement FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)

CREATE TABLE PDM.MtlPrinciple( /*品號原則*/
 PrincipleId				int					identity,
 CompanyId					int					NOT NULL,						--所屬公司 
 PrincipleDesc				nvarchar(100)		NOT NULL,						--品號原則描述
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MtlPrinciple PRIMARY KEY (PrincipleId),
 CONSTRAINT FK1_MtlPrinciple FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)

CREATE TABLE PDM.MtlPrincipleBuild( /*品號原則組成要件*/
 BuildId					int					identity,
 PrincipleId				int					NOT NULL,						--所屬原則
 ElementId					int					NOT NULL,						--所屬元素
 Digits						int					NOT NULL,						--位數大小
 SortNumber					int					NOT NULL,						--排序號碼
 Surfix						nvarchar(50)			NULL,						--後綴文字
 AutoCreate					nvarchar(2)			NOT NULL	DEFAULT 'N',		--值由系統自動產生 Y/N
 EditAuthority				nvarchar(2)			NOT NULL	DEFAULT 'Y',		--值可由USER自行修改 Y/N
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MtlPrincipleBuild PRIMARY KEY (BuildId),
 CONSTRAINT FK1_MtllPrincipleBuild FOREIGN KEY (PrincipleId) REFERENCES PDM.MtlPrinciple(PrincipleId),
 CONSTRAINT FK2_MtlPrincipleBuild FOREIGN KEY (ElementId) REFERENCES PDM.MtlElement(ElementId)
)

CREATE TABLE PDM.MtlPrincipleCondition( /*品號原則條件*/
 ConditionId				int					identity,
 BuildId					int					NOT NULL,						--所屬組成要件
 [Value]					nvarchar(50)		NOT NULL,						--條件值
 ToPrinciple				int					NOT NULL,						--轉換原則
 ConditionDesc				nvarchar(100)		NOT NULL,						--品號原則條件描述
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MtlPrincipleCondition PRIMARY KEY (ConditionId),
 CONSTRAINT AK1_MtlPrincipleCondition UNIQUE (BuildId, [Value], ToPrinciple),
 CONSTRAINT FK1_MtlPrincipleCondition FOREIGN KEY (BuildId) REFERENCES PDM.MtlPrincipleBuild(BuildId),
 CONSTRAINT FK2_MtlPrincipleCondition FOREIGN KEY (ToPrinciple) REFERENCES PDM.MtlPrinciple(PrincipleId)
)

CREATE TABLE PDM.MtlPrincipleBuildData( /*品號原則組成要件資料*/
 DataId						int					identity,
 BuildId					int					NOT NULL,						--所屬組成要件
 DataNo						nvarchar(50)		NOT NULL,						--資料代碼
 DataName					nvarchar(100)		NOT NULL,						--資料名稱
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MtlPrincipleBuildData PRIMARY KEY (DataId),
 CONSTRAINT FK1_MtlPrincipleBuildData FOREIGN KEY (BuildId) REFERENCES PDM.MtlPrincipleBuild(BuildId)
)

CREATE TABLE PDM.MtlModel( /*品號機型*/
 MtlModelId					int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 ParentId					int					NOT NULL,						--父層機型
 MtlModelSort				int					NOT NULL,						--機型排序
 MtlModelNo					nvarchar(50)		NOT NULL,						--品號機型代碼
 MtlModelName				nvarchar(100)		NOT NULL,						--品號機型名稱
 MtlModelDesc				nvarchar(100)		NOT NULL,						--品號機型描述
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MtlModel PRIMARY KEY (MtlModelId),
 CONSTRAINT AK1_MtlModel UNIQUE (MtlModelNo),
 CONSTRAINT AK2_MtlModel UNIQUE (MtlModelName),
 CONSTRAINT AK3_MtlModel UNIQUE (ParentId, MtlModelSort),
 CONSTRAINT FK1_MtlModel FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)

CREATE TABLE PDM.UnitOfMeasure( /*單位基本資料*/ /*ERP INVMX*/
 UomId						int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 UomNo						nvarchar(50)		NOT NULL,						--MX001 單位
 UomName					nvarchar(100)		NOT NULL,						--MX002 單位名稱
 UomType					nvarchar(2)			NOT NULL,						--單位類別(1：庫存(同ERP)；2：體積；3：長度)
 UomDesc					nvarchar(100)			NULL,						--MX004 單位描述
 TransferStatus				nvarchar(2)			NOT NULL,						--拋轉ERP狀態 >> BAS.[Status]
 TransferDate				datetime				NULL,						--拋轉ERP時間
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_UnitOfMeasure PRIMARY KEY (UomId),
 CONSTRAINT AK1_UnitOfMeasure UNIQUE (CompanyId, UomNo),
 CONSTRAINT FK1_UnitOfMeasure FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)

CREATE TABLE PDM.UomCalculate( /*單位換算資料*/
 UomCalculateId				int					identity,
 UomId  					int					NOT NULL,						--所屬單位
 ConvertUomId				int					NOT NULL,						--換算單位
 [Value]					float				NOT NULL,						--值
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_UomCalculate PRIMARY KEY (UomCalculateId),
 CONSTRAINT FK1_UomCalculate FOREIGN KEY (UomId) REFERENCES PDM.UnitOfMeasure(UomId),
 CONSTRAINT FK2_UomCalculate FOREIGN KEY (ConvertUomId) REFERENCES PDM.UnitOfMeasure(UomId)
)

CREATE TABLE PDM.MassProductionReviewDocSetting( /*量產審查文件設定*/ /*20240806*/
 DocId						int					identity,
 DocName					nvarchar(50)		NOT NULL,						--文件名稱
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MassProductionReviewDocSetting PRIMARY KEY (DocId),
) /*量產審查文件設定*/

CREATE TABLE PDM.MtlItemDocSetting( /*品號審查文件設定*/ /*20241113*/
 MtDocId					int					identity,
 DocName					nvarchar(50)		NOT NULL,						--文件名稱
 PurchaseMandatory			nvarchar(2)			NOT NULL,						--請購是否要檢核(Y/N)
 SaleOrderMandatory			nvarchar(2)			NOT NULL,						--訂單是否要檢核(Y/N)
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MtlItemDocSetting PRIMARY KEY (MtDocId),

) /*品號審查文件設定*/


CREATE TABLE PDM.MoldingConditions( /*成型產品生產條件*/ 
 McId						int					identity,
 MtlItemId					int					NOT NULL,						--主見品號
 CellMtlItemId				int						NULL,						--元件品號(材料品號)
 Material					nvarchar(100)		NOT	NULL,						--材料
 CycleTime					float				NOT	NULL,						--週期
 MoldWeight					float				NOT	NULL,						--模重
 Cavity 					int					NOT	NULL,						--穴數
 ProcessingFee  			float				NOT	NULL,						--加工費(固定)(1800/1.13)
 UnitPrice  				float					NULL,						--材料採購單價(抓ERP最新)
 TheoreticalQty  			int					NOT	NULL,						--理論生產數(公式 24*60*60/CycleTime*Cavity)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MoldingConditions PRIMARY KEY (McId),
 CONSTRAINT FK1_MoldingConditions FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId)
) /*成型產品生產條件*/

CREATE TABLE PDM.ItemCatDept( /*品号类别责任部门*/
 ItemCatDeptId				int					identity,
 TypeTwo				    nvarchar(6)		    NOT NULL,						--品號分類二
 TypeThree				    nvarchar(6)			NOT NULL,						--品號分類三
 InventoryNo				nvarchar(6)			NOT NULL,						--主要庫別
 CatDept				    nvarchar(100)		NOT NULL,						--责任部门
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ItemCatDept PRIMARY KEY (ItemCatDeptId),
 CONSTRAINT AK1_ItemCatDept UNIQUE (TypeTwo, TypeThree, InventoryNo, CatDept)
)/*品号类别责任部门*/
END /*基本資料*/

BEGIN /*圖面管理*/
CREATE TABLE PDM.CustomerCad( /*客戶設計圖*/
 CadId						int					identity, 
 CompanyId					int					NOT NULL,						--所屬公司
 CustomerMtlItemNo			nvarchar(50)		NOT	NULL,						--客戶品號(部番)
 CustomerDwgNo				nvarchar(50)			NULL,						--客戶圖號
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_CustomerCad PRIMARY KEY (CadId),
 CONSTRAINT AK1_CustomerCad UNIQUE (CompanyId, CustomerMtlItemNo, CustomerDwgNo),
 CONSTRAINT FK1_CustomerCad FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)

CREATE TABLE PDM.CustomerCadControl( /*客戶設計圖版本控制*/
 ControlId					int					identity,
 CadId						int					NOT NULL,						--所屬客戶設計圖
 CadFile					int						NULL,						--所屬檔案
 CadFileAbsolutePath		nvarchar(MAX)			NULL,						--所屬檔案絕對路徑
 Edition					nvarchar(32)		NOT NULL,						--圖面版次
 [Version]					nvarchar(32)		NOT NULL,						--系統版次
 ReleasedStatus				nvarchar(2)			NOT NULL,						--是否為發行版本 Y:發行版本 N:非發行版本
 Cause						nvarchar(100)			NULL,						--改版原因
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_CustomerCadControl PRIMARY KEY (ControlId),
 CONSTRAINT FK1_CustomerCadControl FOREIGN KEY (CadId) REFERENCES PDM.CustomerCad(CadId),
 CONSTRAINT FK2_CustomerCadControl FOREIGN KEY (CadFile) REFERENCES BAS.[File](FileId)
)

CREATE TABLE PDM.CustomerCadControlFile( /*客戶設計圖系統 其他類檔案(非設計圖檔)*/ 
 ControlId					int					NOT NULL,						--所屬客戶設計圖版次
 FileId						int					NOT NULL,						--所屬檔案
 CadOtherFileAbsolutePath	nvarchar(MAX)			NULL,						--所屬檔案絕對路徑
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_CustomerCadControlFile PRIMARY KEY (FileId, ControlId),
 CONSTRAINT FK1_CustomerCadControlFile FOREIGN KEY (ControlId) REFERENCES PDM.CustomerCadControl(ControlId),
 CONSTRAINT FK2_CustomerCadControlFile FOREIGN KEY (FileId) REFERENCES BAS.[File](FileId)
)/*客戶設計圖系統 其他類檔案(非設計圖檔)*/

CREATE TABLE PDM.RdDesign( /*研發設計圖*/
 DesignId					int					identity,  
 CompanyId					int					NOT NULL,						--所屬公司
 CustomerMtlItemNo			nvarchar(50)		NOT	NULL,						--客戶品號(部番)
 CustomerCadControlId		int						NULL,						--綁定客戶設計圖版本
 MtlItemId					int						NULL,						--綁定品號
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_RdDesign PRIMARY KEY (DesignId),
 CONSTRAINT AK1_RdDesign UNIQUE (CompanyId, CustomerMtlItemNo),
 CONSTRAINT FK1_RdDesign FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_RdDesign FOREIGN KEY (CustomerCadControlId) REFERENCES PDM.CustomerCadControl(ControlId),
 CONSTRAINT FK3_RdDesign FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId)
)/*研發設計圖*/

CREATE TABLE PDM.RdDesignControl( /*研發設計圖版本控制*/
 ControlId					int					identity,
 DesignId					int					NOT NULL,						--所屬研發設計圖檔
 Cad3DFile					int						NULL,						--研發3D CAD檔FileId
 Cad2DFile					int						NULL,						--研發2D CAD檔FileId
 Pdf2DFile					int						NULL,						--研發2D PDF檔FileId
 JmoFile					int						NULL,						--JMO檔路徑
 Cad3DFileAbsolutePath		nvarchar(MAX)			NULL,						--研發3D CAD檔 絕對路徑
 Cad2DFileAbsolutePath		nvarchar(MAX)			NULL,						--研發2D CAD檔 絕對路徑
 Pdf2DFileAbsolutePath		nvarchar(MAX)			NULL,						--研發2D PDF檔 絕對路徑
 JmoFileAbsolutePath		nvarchar(MAX)			NULL,						--JMO檔 絕對路徑
 Edition					nvarchar(32)		NOT NULL,						--圖面版次
 [Version]					nvarchar(32)		NOT NULL,						--系統版次
 Cause						nvarchar(100)			NULL,						--改版原因
 DesignDate					datetime			NOT NULL,						--出圖日期
 ReleasedStatus				nvarchar(2)			NOT NULL,						--是否為發行版本 Y:發行版本 N:非發行版本
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_RdDesignControl PRIMARY KEY (ControlId),
 CONSTRAINT FK1_RdDesignControl FOREIGN KEY (DesignId) REFERENCES PDM.RdDesign(DesignId),
 CONSTRAINT FK2_RdDesignControl FOREIGN KEY (Cad3DFile) REFERENCES BAS.[File](FileId),
 CONSTRAINT FK3_RdDesignControl FOREIGN KEY (Cad2DFile) REFERENCES BAS.[File](FileId),
 CONSTRAINT FK4_RdDesignControl FOREIGN KEY (Pdf2DFile) REFERENCES BAS.[File](FileId),
 CONSTRAINT FK5_RdDesignControl FOREIGN KEY (JmoFile) REFERENCES BAS.[File](FileId)
)/*研發設計圖版本控制*/

CREATE TABLE PDM.RdDesignControlFile( /*研發設計圖系統 其他類檔案(非設計圖檔)*/
 ControlId					int					NOT NULL,						--所屬研發設計圖版次
 FileId						int					NOT NULL,						--所屬檔案
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_RdDesignControlFile PRIMARY KEY (FileId, ControlId),
 CONSTRAINT FK1_RdDesignControlFile FOREIGN KEY (ControlId) REFERENCES PDM.RdDesignControl(ControlId),
 CONSTRAINT FK2_RdDesignControlFile FOREIGN KEY (FileId) REFERENCES BAS.[File](FileId)
)/*研發設計圖系統 其他類檔案(非設計圖檔)*/

CREATE TABLE PDM.RdAttributeItem( /*設計屬性項目*/ --模仁外徑,APS係數,澆口寬...等
 AttributeItemId			int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 ModeId						int					NOT NULL,						--所屬生產模式
 ItemNo						nvarchar(32)		NOT NULL,						--項目代碼
 ItemDesc					nvarchar(200)		NOT NULL,						--項目說明
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_RdAttributeItem PRIMARY KEY (AttributeItemId),
 CONSTRAINT FK1_RdAttributeItem FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_RdAttributeItem FOREIGN KEY (ModeId) REFERENCES MES.ProdMode(ModeId)
)

CREATE TABLE PDM.RdDesignAttribute( /*研發設計基本資料(如果由圖面就可得到資料放此,否則直接放RoutingItemAttribute)*/
 AttributeId				int					identity,
 ControlId					int					NOT NULL,						--所屬研發設計圖
 AttributeItemId			int					NOT NULL,						--所屬設計屬性
 AttributeValue				nvarchar(200)			NULL,						--屬性值
 UpperLimit					nvarchar(20)			NULL,						--上限值
 LowerLimit					nvarchar(20)			NULL,						--下限值
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_RdDesignAttribute PRIMARY KEY (AttributeId),
 CONSTRAINT FK1_RdDesignAttribute FOREIGN KEY (ControlId) REFERENCES PDM.RdDesignControl(ControlId),
 CONSTRAINT FK2_RdDesignAttribute FOREIGN KEY (AttributeItemId) REFERENCES PDM.RdAttributeItem(AttributeItemId)
)

CREATE TABLE PDM.RdWhitelist( /*RD白名單路徑資料*/
 ListId						int					identity,
 DepartmentId				int					NOT NULL,						--使用部門ID
 FolderPath					nvarchar(max)		NOT NULL,						--路徑
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_RdWhilelist PRIMARY KEY (ListId),
 CONSTRAINT AK1_RdWhilelist UNIQUE (DepartmentId, VirtualPathName),
 CONSTRAINT FK1_RdWhilelist FOREIGN KEY (DepartmentId) REFERENCES BAS.Department(DepartmentId),
)/*RD白名單路徑資料*/
END /*圖面管理*/

BEGIN /*產品資料管理*/
CREATE TABLE PDM.MtlItem( /*品號基本資料*/ /*ERP INVMB*/
 MtlItemId					int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 ItemTypeId					int					NOT NULL,						--品號類型
 MtlItemNo					nvarchar(40)		NOT NULL,						--MB001 品號
 MtlItemName				nvarchar(120)		NOT NULL,						--MB002 品名
 MtlItemSpec				nvarchar(120)		NOT NULL,						--MB003 規格
 InventoryUomId				int						NULL,						--MB004 庫存單位 >> 參照 ERP品號換算單位檔(INVMD)
 UowId						int						NULL,						--MB015 重量單位 >> 參照 ERP品號換算單位檔(INVMD)
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
 ReplenishmentPolicy		nvarchar(1)			NOT NULL	DEFAULT 'R',		--MB034 補貨政策
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
 CONSTRAINT PK_MtlItem PRIMARY KEY (MtlItemId),
 CONSTRAINT AK1_MtlItem UNIQUE (CompanyId, MtlItemNo),
 CONSTRAINT FK1_MtlItem FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_MtlItem FOREIGN KEY (InventoryUomId) REFERENCES PDM.UnitOfMeasure(UomId),
 CONSTRAINT FK3_MtlItem FOREIGN KEY (WeightUomId) REFERENCES PDM.UnitOfMeasure(UomId),
 CONSTRAINT FK4_MtlItem FOREIGN KEY (PurchaseUomId) REFERENCES PDM.UnitOfMeasure(UomId),
 CONSTRAINT FK5_MtlItem FOREIGN KEY (SaleUomId) REFERENCES PDM.UnitOfMeasure(UomId),
 CONSTRAINT FK6_MtlItem FOREIGN KEY (InventoryId) REFERENCES SCM.Inventory(InventoryId),
 CONSTRAINT FK7_MtlItem FOREIGN KEY (RequisitionInventoryId) REFERENCES SCM.Inventory(InventoryId)
)/*品號基本資料*/

CREATE TABLE PDM.MtlItemSetting( /*品號設定*/
 MtlItemId					int					NOT NULL,						--所屬品號
 FixedLeadTime				int					NOT NULL,						--固定前置天數
 ChangeLeadTime				int					NOT NULL,						--變動前置天數
 MinPoQty					int					NOT NULL,						--最低補量/最低採購數量
 MultiplePoQty				int					NOT NULL,						--補貨倍量/採購數量倍數
 MultipleReqQty				int					NOT NULL,						--領用倍量/領用倍數
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT FK1_MtlItemSetting FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId)
)/*品號設定*/

CREATE TABLE PDM.CustomerMtlItem( /*客戶品號*/
 CustomerMtlItemId			int					identity,
 MtlItemId					int					NOT NULL,						--所屬品號
 CustomerId					int					NOT NULL,						--所屬客戶
 PoErpFullNo				nvarchar(20)			NULL,						--採購單來源單號
 PoSeq						nvarchar(6)				NULL,						--採購身序號
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
 CONSTRAINT PK_CustomerMtlItem PRIMARY KEY (CustomerMtlItemId),
 CONSTRAINT AK1_CustomerMtlItem UNIQUE (MtlItemId, CustomerId, CustomerMtlItemNo),
 CONSTRAINT FK1_CustomerMtlItem FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId),
 CONSTRAINT FK2_CustomerMtlItem FOREIGN KEY (CustomerId) REFERENCES SCM.Customer(CustomerId)
)/*客戶品號*/

CREATE TABLE PDM.BillOfMaterials( /*BOM用量資料單頭*/ /*ERP BOMMC*/
 BomId						int					identity,
 MtlItemId					int					NOT NULL,						--所屬品號
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
 CONSTRAINT PK_BillOfMaterials PRIMARY KEY (BomId),
 --CONSTRAINT AK1_BillOfMaterials UNIQUE (MtlItemId),
 CONSTRAINT FK1_BillOfMaterials FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId),
 CONSTRAINT FK2_BillOfMaterials FOREIGN KEY (UomId) REFERENCES PDM.UnitOfMeasure(UomId),
 CONSTRAINT FK3_BillOfMaterials FOREIGN KEY (ConfirmUserId) REFERENCES BAS.[User](UserId)
)/*BOM用量資料單頭*/

CREATE TABLE PDM.BomDetail( /*BOM用量資料單身*/ /*ERP BOMMD*/
 BomDetailId				int					identity,
 BomId						int					NOT NULL,						--MD001 所屬單頭
 BomSequence				nvarchar(4)			NOT NULL,						--MD002 序號
 MtlItemId					int					NOT NULL,						--MD003 所屬品號
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
 ComponentType				nvarchar(2)			NOT NULL	DEFAULT 'N',		--上料設定
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_BomDetail PRIMARY KEY (BomDetailId),
 --CONSTRAINT AK1_BomDetail UNIQUE (BomId, BomSequence),
 CONSTRAINT FK1_BomDetail FOREIGN KEY (BomId) REFERENCES PDM.BillOfMaterials(BomId),
 CONSTRAINT FK2_BomDetail FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId),
 CONSTRAINT FK3_BomDetail FOREIGN KEY (UomId) REFERENCES PDM.UnitOfMeasure(UomId)
)/*BOM用量資料單身*/

CREATE TABLE PDM.BomSubstitution( /*取替代料件*/ /*ERP BOMMB*/
 SubstitutionId				int					identity,
 BomDetailId				int					NOT NULL,						--所屬BOM用量資料單身
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
 --CONSTRAINT AK1_BomSubstitution UNIQUE (BomDetailId, SubstituteStatus, MtlItemId),
 --CONSTRAINT AK2_BomSubstitution UNIQUE (BomDetailId, SubstituteStatus, SortNumber),
 CONSTRAINT FK1_BomSubstitution FOREIGN KEY (BomDetailId) REFERENCES PDM.BomDetail(BomDetailId),
 CONSTRAINT FK2_BomSubstitution FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId)
)/*取替代料件*/

CREATE TABLE PDM.ItemType( /*品號類別*/
 ItemTypeId					int					identity,
 ParentItemTypeId			int					NULL,							--上層ItemTypeId
 CompanyId					int					NOT NULL,						--所屬公司
 ItemTypeNo					nvarchar(30)		NOT	NULL,						--品號類別代碼
 ItemTypeName				nvarchar(30)		NOT	NULL,						--品號類別名稱
 ItemTypeDesc				nvarchar(200)			NULL,						--品號類別敘述
 EffectiveDate				datetime			NOT NULL,						--生效日(預設1900-01-01)
 InactiveDate				datetime			NULL,							--失效日
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ItemType PRIMARY KEY (ItemTypeId),
 CONSTRAINT AK1_ItemType UNIQUE (CompanyId, ItemTypeNo),
 CONSTRAINT FK1_ItemType FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_ItemType FOREIGN KEY (ParentItemTypeId) REFERENCES PDM.ItemType(ItemTypeId)
)

CREATE TABLE PDM.ItemTypeSegment( /*品號類別結構(編碼節斷)*/
 ItemTypeSegmentId			int					identity,
 ItemTypeId					int					NOT NULL,						--PDM.ItemTypeId	
 SortNumber					int					NOT NULL,						--排序值
 SegmentType				nvarchar(2)			NOT NULL,						--編碼內容類別(1.Fix[固定值] = F,2.Date[日期類,YYYYMMDD or YYYYMM or YYMM] = D,3.Sequence[流水號] = S,4.SegmentValue[編碼節段] = SV, 5.Customize[自訂義] = C)
 SegmentValue				nvarchar(20)		NULL,							--StructureType = 1: 固定值
                                                                                --StructureType = 2: YYYYMMDD / YYYYMM / YYMM
																				--StructureType = 3: 流水碼長度
																				--StructureType = 4: ItemSegmentId (那一個ItemSegment.ItemSegmentId)
																				--StructureType = 5: 自訂義
 SuffixCode					nvarchar(30)		NULL,							--後綴碼(後面要不要加"-"),可以為空
 EffectiveDate				datetime			NOT NULL,						--生效日(預設1900-01-01)
 InactiveDate				datetime			NULL,							--失效日
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ItemTypeSegment PRIMARY KEY (ItemTypeSegmentId),
 CONSTRAINT FK1_ItemTypeSegment FOREIGN KEY (ItemTypeId) REFERENCES PDM.ItemType(ItemTypeId)
)

CREATE TABLE PDM.ItemTypeDefault( /*品號類別預設資料*/ 
 ItemTypeDefaultId			int					identity,
 ItemTypeId					int					NOT NULL,						--PDM.ItemTypeId
 TypeOne					nvarchar(6)			NOT NULL,						--MB005 品號分類一
 TypeTwo					nvarchar(6)			NOT NULL,						--MB006 品號分類二
 TypeThree					nvarchar(6)			NOT NULL,						--MB007 品號分類三
 TypeFour					nvarchar(6)			NOT NULL,						--MB008 品號分類四
 InventoryId				int						NULL,						--MB017 主要庫別
 RequisitionInventoryId		int						NULL,						--MB157 主要領料庫別
 InventoryUomId				int						NULL,						--MB004 庫存單位 >> 參照 ERP品號換算單位檔(INVMD)
 ItemAttribute				nvarchar(1)				NULL,						--MB025 品號屬性 >> BAS.[Type] P：採購件；M：自製件；S：託外加工件；Y：虛設品號；F：Feature件；O:Option件
 MeasureType				nvarchar(1)				NULL	DEFAULT '2',		--MB043 檢驗方式 >> BAS.[Type] 0：免檢；1：抽檢(減量)；2：抽檢(正常)；3：抽檢(加嚴)；4：全檢
 WeightUomId				int						NULL,						--MB015 重量單位 >> 參照 ERP品號換算單位檔(INVMD)
 PurchaseUomId				int						NULL,						--MB155 採購單位 >> 參照 ERP品號換算單位檔(INVMD)
 SaleUomId					int						NULL,						--MB156 銷售單位 >> 參照 ERP品號換算單位檔(INVMD)
 LotManagement				nvarchar(1)				NULL	DEFAULT 'N',		--MB022 批號管理 >> BAS.[Type] N:不需要；Y:需要不檢查庫存量；W:僅需警告；T:需要且檢查庫存量
 InventoryManagement		nvarchar(2)				NULL,						--MB019 庫存管理 >> BAS.[Status]
 MtlModify					nvarchar(1)				NULL,						--MB066 修改品名/規格 >> BAS.[Status]
 BondedStore				nvarchar(1)				NULL,						--MB020 保稅品 >> BAS.[Status]
 OverReceiptManagement		nvarchar(1)				NULL,						--MB044 超收管理 >> BAS.[Status]
 OverDeliveryManagement		nvarchar(1)				NULL	DEFAULT 'N',		--MB087 超交管理 >> BAS.[Status]
 EffectiveDate				datetime				NULL,						--MB030 生效日期
 ExpirationDate				datetime				NULL,						--MB031 失效日期
 MtlItemDesc				nvarchar(255)			NULL,						--MB009 商品描述
 MtlItemRemark				nvarchar(255)			NULL,						--MB028 品號備註
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ItemTypeDefault PRIMARY KEY (ItemTypeDefaultId),
 CONSTRAINT FK1_ItemTypeDefault FOREIGN KEY (ItemTypeId) REFERENCES PDM.ItemType(ItemTypeId),
 CONSTRAINT FK2_ItemTypeDefault FOREIGN KEY (InventoryId) REFERENCES SCM.Inventory(InventoryId),
 CONSTRAINT FK3_ItemTypeDefault FOREIGN KEY (RequisitionInventoryId) REFERENCES SCM.Inventory(InventoryId)
)

CREATE TABLE PDM.ItemSegment( /*Segment(編碼節段主檔)*/
 ItemSegmentId				int					identity,
 CompanyId					int					NOT NULL,
 SegmentNo					nvarchar(20)		NOT NULL,						--編碼節段代碼
 SegmentName				nvarchar(20)		NOT	NULL,						--編碼節段名稱
 SegmentDesc				nvarchar(200)			NULL,						--編碼節段說明
 EffectiveDate				datetime			NOT NULL,						--生效日(預設1900-01-01)
 InactiveDate				datetime			NULL,							--失效日
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ItemSegment PRIMARY KEY (ItemSegmentId),
 CONSTRAINT AK1_ItemSegment UNIQUE (CompanyId, SegmentNo),
 CONSTRAINT FK1_ItemSegment FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)

CREATE TABLE PDM.ItemSegmentValue( /*SegmentValue*/
 SegmentValueId				int					identity,
 ItemSegmentId				int					NOT NULL,
 SegmentValueNo				nvarchar(20)		NOT NULL,						--代碼
 SegmentValue				nvarchar(200)		NOT NULL,						--代碼值
 SegmentValueDesc			nvarchar(200)			NULL,						--代碼說明
 EffectiveDate				datetime			NOT NULL,						--生效日(預設1900-01-01)
 InactiveDate				datetime			NULL,							--失效日
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ItemSegmentValue PRIMARY KEY (SegmentValueId),
 CONSTRAINT AK1_ItemSegmentValue UNIQUE (ItemSegmentId, SegmentValueNo),
 CONSTRAINT FK1_ItemSegmentValue FOREIGN KEY (ItemSegmentId) REFERENCES PDM.ItemSegment(ItemSegmentId)
)

CREATE TABLE PDM.MtlItemSegment( /*品號類別編碼設定資料*/ 
 MtlItemSegmentId			int					identity,
 MtlItemId					int					NOT NULL,						--品號Id
 ItemTypeId					int					NOT NULL,						--品號類別Id
 SegmentType				nvarchar(2)			NOT	NULL,						--編碼內容類別
 SegmentSort				int					NOT	NULL,						--排序值
 SegmentValue				nvarchar(20)		NOT	NULL,						--截段碼
 SaveValue					nvarchar(100)			NULL,						--儲存值
 SuffixCode					nvarchar(30)		NULL,							--後綴碼(後面要不要加"-"),可以為空
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MtlItemSegment PRIMARY KEY (MtlItemSegmentId),
 --CONSTRAINT AK1_ItemSegment UNIQUE (MtlItemId,SegmentSort),
 CONSTRAINT FK1_MtlItemSegment FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId),
 CONSTRAINT FK2_MtlItemSegment FOREIGN KEY (ItemTypeId) REFERENCES PDM.ItemType(ItemTypeId),
)

CREATE TABLE PDM.MtlItemUomCalculate( /*品號單位換算資料　INVMD*/
 MtlItemUomCalculateId		int					identity,
 MtlItemId					int					NOT NULL,						--所屬品號 MD001
 UomId  					int					NOT NULL,						--庫存單位 
 ConvertUomId				int					NOT NULL,						--換算單位 MD002
 SwapNumerator				float				NOT NULL,						--換算率分子(換算) MD003
 SwapDenominator			float				NOT NULL,						--換算率分母(庫存) MD004
 TransferStatus				nvarchar(2)			NOT NULL    DEFAULT 'N',		--拋轉ERP狀態 >> BAS.[Status]
 TransferDate				datetime				NULL,						--拋轉ERP時間
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MtlItemUomCalculate PRIMARY KEY (MtlItemUomCalculateId),
 CONSTRAINT FK1_MtlItemUomCalculate FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId),
 CONSTRAINT FK2_MtlItemUomCalculate FOREIGN KEY (UomId) REFERENCES PDM.UnitOfMeasure(UomId),
 CONSTRAINT FK3_MtlItemUomCalculate FOREIGN KEY (ConvertUomId) REFERENCES PDM.UnitOfMeasure(UomId)
)

CREATE TABLE PDM.MtlItemUomCalculateLog( /*品號單位換算LOG資料*/
 MtlItemUomCalculateLogId	int					identity,
 TransactionType			nvarchar(2)				NULL,						--交易類型 A:新增 U:更新 T:拋轉 D:刪除 SA:同步新增 SU:同步更新 SD:同步刪除
 MtlItemUomCalculateId		int						NULL,						--品號單位換算資料ID
 MtlItemId					int						NULL,						--所屬品號 
 UomId  					int						NULL,						--庫存單位 
 ConvertUomIdBefore			int						NULL,						--換算單位變更前  
 SwapDenominatorBefore		float					NULL,						--換算率分母(庫存)變更前  
 SwapNumeratorBefore		float					NULL,						--換算率分子(換算)變更前  
 ConvertUomId				int						NULL,						--換算單位變更後 
 SwapDenominator			float					NULL,						--換算率分母(庫存)變更後  
 SwapNumerator				float					NULL,						--換算率分子(換算)變更後  
 TransactionUserId			int						NULL,						--交易者
 TransactionDate			datetime				NULL,						--交易時間
 TransferStatus				nvarchar(2)				NULL    DEFAULT 'N',		--拋轉ERP狀態 >> BAS.[Status]
 TransferDate				datetime				NULL,						--拋轉ERP時間
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MtlItemUomCalculateLog PRIMARY KEY (MtlItemUomCalculateLogId),
)

CREATE TABLE PDM.MassProductionReviewDoc( /*量產審查文件*/ /*20240806*/
 MtlItemId					int					NOT NULL,						--所屬品號
 DocId						int					NOT NULL,						--所屬文件
 FileId						int					NOT NULL,						--所屬文件
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT AK1_MassProductionReviewDoc UNIQUE (MtlItemId, DocId),
 CONSTRAINT FK1_MassProductionReviewDoc FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId),
 CONSTRAINT FK2_MassProductionReviewDoc FOREIGN KEY (DocId) REFERENCES PDM.MassProductionReviewDocSetting(DocId),
 CONSTRAINT FK3_MassProductionReviewDoc FOREIGN KEY (FileId) REFERENCES BAS.[File](FileId),
) /*量產審查文件*/

CREATE TABLE PDM.MtlItemFile( /*量產審查文件*/ /*20240806*/
 MtlItemId					int					NOT NULL,						--所屬品號
 DocId						int					NOT NULL,						--所屬文件
 FileId						int					NOT NULL,						--所屬文件
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT AK1_MassProductionReviewDoc UNIQUE (MtlItemId, DocId),
 CONSTRAINT FK1_MassProductionReviewDoc FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId),
 CONSTRAINT FK2_MassProductionReviewDoc FOREIGN KEY (DocId) REFERENCES PDM.MassProductionReviewDocSetting(DocId),
 CONSTRAINT FK3_MassProductionReviewDoc FOREIGN KEY (FileId) REFERENCES BAS.[File](FileId),
) /*量產審查文件*/

CREATE TABLE PDM.MtlItemFile( /*品號審查文件檔*/ /*20241113*/
 MtlItemNo					nvarchar(40)		NOT NULL,						--所屬品號
 MtDocId					int					NOT NULL,						--所屬品號審查文件
 FileId						int					NOT NULL,						--所屬檔案
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT AK1_MtlItemFile UNIQUE (MtlItemNo, MtDocId),
 CONSTRAINT FK1_MtlItemFile FOREIGN KEY (MtDocId) REFERENCES PDM.MtlItemDocSetting(MtDocId),
 CONSTRAINT FK2_MtlItemFile FOREIGN KEY (FileId) REFERENCES BAS.[File](FileId),
) /*品號相關文件檔*/

CREATE TABLE PDM.MtlItemNoIntegration( /*新舊品號整合*/ /*20241108*/
 MtlItemId					int					NOT NULL,						--所屬品號
 NewMtlItemNo				nvarchar(40)		NOT NULL,						--MB001 品號
 NewMtlItemName				nvarchar(120)			NULL,						--MB002 品名
 NewMtlItemSpec				nvarchar(120)			NULL,						--MB003 規格

 ItemTransfer				bit					NOT NULL	DEFAULT 0,			--品號是否轉換
 InventoryTransfer			bit					NOT NULL	DEFAULT 0,			--庫存是否轉換
 TransferMsg				nvarchar(MAX)			NULL,						--轉換訊息

 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT FK1_MtlItemNoIntegration FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId),
) /*新舊品號整合*/

CREATE TABLE PDM.NewMtlItemDetail( /*新品號基本資料*/ /*20241206*/
 NewMtDetailId				int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 MtlItemNo					nvarchar(40)		NOT NULL,						--MB001 品號
 InventoryUomId				int						NULL,						--MB004 庫存單位 >> 參照 ERP品號換算單位檔(INVMD)
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
 ReplenishmentPolicy		nvarchar(1)			NOT NULL	DEFAULT 'R',		--MB034 補貨政策
 MeasureType				nvarchar(1)			NOT NULL	DEFAULT '2',		--MB043 檢驗方式 >> BAS.[Type] 0：免檢；1：抽檢(減量)；2：抽檢(正常)；3：抽檢(加嚴)；4：全檢
 EffectiveDate				datetime				NULL,						--MB030 生效日期
 ExpirationDate				datetime				NULL,						--MB031 失效日期
 OverReceiptManagement		nvarchar(1)			NOT NULL,						--MB044 超收管理 >> BAS.[Status]
 OverDeliveryManagement		nvarchar(1)			NOT NULL	DEFAULT 'N',		--MB087 超交管理 >> BAS.[Status]
 [Version]					nvarchar(4)			NOT NULL,						--MB165 版次
 MtlItemDesc				nvarchar(255)			NULL,						--MB009 商品描述
 MtlItemRemark				nvarchar(255)			NULL,						--MB028 品號備註
 TransferStatus				nvarchar(2)				NULL	DEFAULT 'N',		--新品號是否拋轉ERP
 TransferDate				datetime				NULL,						--品號拋轉ERP時間
 BOMVersion					nvarchar(255)			NULL,						--BOM版本
 BOMTransferStatus			nvarchar(2)				NULL	DEFAULT 'N',		--BOM表是否拋轉ERP
 BOMTransferDate			datetime				NULL,						--BOM表拋轉ERP時間
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼 
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_NewMtlItemDetail PRIMARY KEY (NewMtDetailId),
 CONSTRAINT AK1_NewMtlItemDetail UNIQUE (CompanyId, MtlItemNo),
 CONSTRAINT FK1_NewMtlItemDetail FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_NewMtlItemDetail FOREIGN KEY (InventoryUomId) REFERENCES PDM.UnitOfMeasure(UomId),
 CONSTRAINT FK3_NewMtlItemDetail FOREIGN KEY (PurchaseUomId) REFERENCES PDM.UnitOfMeasure(UomId),
 CONSTRAINT FK4_NewMtlItemDetail FOREIGN KEY (SaleUomId) REFERENCES PDM.UnitOfMeasure(UomId),
 CONSTRAINT FK5_NewMtlItemDetail FOREIGN KEY (InventoryId) REFERENCES SCM.Inventory(InventoryId),
 CONSTRAINT FK6_NewMtlItemDetail FOREIGN KEY (RequisitionInventoryId) REFERENCES SCM.Inventory(InventoryId)
) /*新舊品號整合*/
END

BEGIN /*品號變更管理*/
CREATE TABLE PDM.MtlItemChange( /*品號變更單單頭資料*/ /*ERP INVTL*/
 MtlItemChangeId			int					identity, 
 CompanyId					int					NOT NULL,						--所屬公司 
 MtlItemId					int					NOT NULL,						--TL001 品號
 OriginalMtlItemName		nvarchar(120)			NULL,						--TL002 原品名 
 OriginalMtlItemSpec		nvarchar(120)			NULL,						--TL003 原規格
 ChangeEdition				nvarchar(4)			NOT NULL,						--TL004 變更版次
 ChangeDate					datetime				NULL,						--TL005 變更日期
 DelMtlItemStatus			nvarchar(1)				NULL	DEFAULT 'N',		--TL006 刪除品號
 NewMtlItemName				nvarchar(120)			NULL,						--TL007 新品名
 NewMtlItemSpec				nvarchar(120)			NULL,						--TL008 新規格
 UpdateBy					int						NULL,						--TL009 修改者
 Remark						nvarchar(255)		    NULL,						--TL010 備註
 FieldType					nvarchar(1)				NULL	DEFAULT '1',		--TL011 修改欄位類別 
 ConfirmBy					int						NULL,						--TL012 確認者
 ConfirmDate				datetime				NULL,						--TL013 確認日期
 ConfirmStatus				nvarchar(2)				NULL	DEFAULT 'N',		--TL014 確認碼
 SignOffStatus				nvarchar(1)				NULL	DEFAULT 'N',		--TL015 簽核狀態
 PrintCount					nvarchar(1)				NULL,						--TL016 列印次數
 SendCount					nvarchar(1)				NULL,						--TL017 傳送次數
 OriginalEdition			nvarchar(4)				NULL,						--TL023 原版次
 TransferStatus				nvarchar(2)			NOT NULL,						--拋轉ERP狀態 >> BAS.[Status]
 TransferDate				datetime				NULL,						--拋轉ERP時間
 [Status]					nvarchar(2)			NOT NULL	DEFAULT 'A',		--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MtlItemChange PRIMARY KEY (MtlItemChangeId),
 CONSTRAINT AK1_MtlItemChange UNIQUE (CompanyId,MtlItemId,ChangeEdition),
 CONSTRAINT FK1_MtlItemChange FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_MtlItemChange FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId)
)/*品號變更單單頭資料*/

CREATE TABLE PDM.MtlItemChangeDetail( /*品號變更單單身資料*/ /*ERP INVTM */
 MtlChangeDetailId			int					identity, 
 MtlItemChangeId			int					NOT NULL,
 MtlItemId					int					NOT NULL,						--TM001 品號 
 ChangeEdition				nvarchar(4)			NOT NULL,						--TM002 變更版次
 SortNum					nvarchar(4)			NOT NULL,						--TM003 序號
 FieldNumBM					nvarchar(100)			NULL,						--BM 欄位
 FieldTypeBM				nvarchar(1)			NULL,							--BM 欄位型態
 FieldNum					nvarchar(15)			NULL,						--TM004 欄位編號
 FieldName					nvarchar(40)			NULL,						--TM005 欄位名稱
 NewTextField				nvarchar(255)			NULL,						--TM006 新文字型欄位值
 OriginalTextField			nvarchar(255)			NULL,						--TM007 原文字型欄位值
 NewNumericField			float					NULL,						--TM008 新數值欄位值
 OriginalNumericField		float					NULL,						--TM009 原數值欄位值
 FieldType					nvarchar(1)			    NULL,						--TM010 欄位型態 
 RetentionField				nvarchar(20)			NULL,						--TM011 保留欄位
 Remark						nvarchar(255)			NULL,						--TM012 備註
 ConfirmStatus				nvarchar(1)			    NULL	DEFAULT 'N',		--TM013 確認碼
 NewFieldName				nvarchar(255)		    NULL,						--TM014 新值名稱
 OriginalFieldName			nvarchar(255)		    NULL,						--TM015 原值名稱
 [Status]					nvarchar(2)			NOT NULL	DEFAULT 'A',		--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MtlItemChangeDetail PRIMARY KEY (MtlChangeDetailId),
 --CONSTRAINT AK1_MtlItemChangeDetail UNIQUE (MtlItemId,ChangeEdition,SortNum),
 CONSTRAINT FK1_MtlItemChangeDetail FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId),
 CONSTRAINT FK2_MtlItemChangeDetail FOREIGN KEY (MtlItemChangeId) REFERENCES PDM.MtlItemChange(MtlItemChangeId)
)/*品號變更單單身資料*/

CREATE TABLE PDM.MtlItemUomCalculateChange( /*單位換算變更資料*/
 UomChangeId				int					identity,
 CompanyId					int					NOT NULL,						--所屬公司 
 MtlItemChangeId			int					NOT NULL,
 MtlItemId					int					NOT NULL,						--所屬品號 TO001
 UomChangeEdit  			nvarchar(10)		NOT NULL,						--版次 TO002
 ConvertUomId				int					NOT NULL,						--原換算單位 TO003
 SwapNumerator				float				NOT NULL,						--原換算率分子(換算) TO004
 SwapDenominator			float				NOT NULL,						--原換算率分母(庫存) TO005
 OriginConvertUomId			int					NOT NULL,						--新換算單位 TO006
 OriginSwapNumerator		float				NOT NULL,						--新換算率分子(換算) TO007
 OriginSwapDenominator		float				NOT NULL,						--新換算率分母(庫存) TO008
 TransferStatus				nvarchar(2)			NOT NULL    DEFAULT 'N',		--拋轉ERP狀態 >> BAS.[Status]
 TransferDate				datetime				NULL,						--拋轉ERP時間
 ConfirmStatus				nvarchar(2)				NULL	DEFAULT 'N',		--確認 N/Y
 ConfirmDate				datetime				NULL,						--確認日期
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MtlItemUomCalculateChange PRIMARY KEY (UomChangeId),
 CONSTRAINT FK1_MtlItemUomCalculateChange FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId),
 CONSTRAINT FK2_MtlItemUomCalculateChange FOREIGN KEY (ConvertUomId) REFERENCES PDM.UnitOfMeasure(UomId),
 CONSTRAINT FK3_MtlItemUomCalculateChange FOREIGN KEY (OriginConvertUomId) REFERENCES PDM.UnitOfMeasure(UomId),
 CONSTRAINT FK4_MtlItemUomCalculateChange FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK5_MtlItemUomCalculateChange FOREIGN KEY (MtlItemChangeId) REFERENCES PDM.MtlItemChange(MtlItemChangeId)
)/*單位換算變更資料*/

END /*品號變更管理*/

BEGIN /*Bom變更管理*/
CREATE TABLE PDM.BomChange( /*Bom變更單單頭資料*/ /*ERP BOMTA*/
 BomChangeId				int					identity, 
 CompanyId					int					NOT NULL,						--所屬公司 
 BomChangePrefix			nvarchar(120)		NOT NULL,						--TA001 變更單別
 BomChangeNo				nvarchar(120)			NULL,						--TA002 變更單號
 ChangeDate					datetime				NULL,						--TA003 變更日期
 EmergencyCode				nvarchar(2)				NULL,						--TA004 緊急碼
 ChangeReason				nvarchar(255)			NULL,						--TA005 變更原因
 Remark						nvarchar(255)			NULL,						--TA006 備註
 ConfirmStatus				nvarchar(2)				NULL	DEFAULT 'N',		--TA007 確認碼
 PrintCount					int						NULL,						--TA008 列印次數
 ConfirmDate				datetime						NULL,						--TA009 單據日期
 ConfirmBy					int						NULL,						--TA010 確認者
 SignOffStatus				nvarchar(1)				NULL	DEFAULT 'N',		--TA011 簽核狀態碼
 SendCount					int						NULL,						--TA012 傳送次數
 TransferStatus				nvarchar(2)			NOT NULL,						--拋轉ERP狀態 >> BAS.[Status]
 TransferDate				datetime				NULL,						--拋轉ERP時間
 [Status]					nvarchar(2)			NOT NULL	DEFAULT 'A',		--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_BomChange PRIMARY KEY (BomChangeId),
 CONSTRAINT AK1_BomChange UNIQUE (CompanyId,BomChangeNo),
 CONSTRAINT FK1_BomChange FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)/*Bom變更單單頭資料*/

CREATE TABLE PDM.BomChangeDetail( /*Bom變更單單身資料*/ /*ERP BOMTB */
 BomChangeDetailId			int					identity, 
 BomChangeId				int					NOT NULL,
 BomId						int					NOT NULL,						--主件ID
 SortNum					nvarchar(4)				NULL,						--TB003 變更序號 
 MtlItemId					int					NOT NULL,						--TB004 主件品號
 Unit						int						NULL,						--TB005 單位 ERP端為nvarchar(6)	
 SmallUnit					nvarchar(6)				NULL,						--TB006 小單位
 ChangeEdition				nvarchar(4)				NULL,						--TB007 版次
 StandardLot				nvarchar(20)			NULL,						--TB008 標準批量
 MoPrefix					nvarchar(4)				NULL,						--TB009 製令單別
 ChangeReason				nvarchar(255)			NULL,						--TB010 變更原因
 ConfirmStatus				nvarchar(30)			NULL,						--TB011 確認狀態
 ConfirmCode				nvarchar(1)				NULL,						--TB012 確認碼 Y/N
 NewRemark					nvarchar(255)			NULL,						--TB013 新備註
 BomChangeSource			nvarchar(1)				NULL,						--TB016 來源-1.PLM
 OriginalMtlItemId			int						NULL,						--TB104 原主件品號
 OriginalUnit				nvarchar(6)				NULL	DEFAULT 'N',		--TB105 原單位
 OriginalSmallUnit			nvarchar(6)				NULL,						--TB106 原小單位
 OriginalChangeEdition		nvarchar(8)				NULL,						--TB107 原版次
 OriginalStandardLot		nvarchar(20)			NULL,						--TB108 原標準批量
 OriginalMoPrefix			nvarchar(4)				NULL,						--TB109 原製令單別
 --OriginalBomChangePrefix	nvarchar(4)				NULL,						--TB110 原變更單別
 OriginalBomChangeId		int						NULL,						--TB111 原變更單號
 OriginalSortNum			nvarchar(4)				NULL,						--TB112 原變更單序號
 OriginalRemark				nvarchar(255)		    NULL,						--TB113 原備註
 [Status]					nvarchar(2)			NOT NULL	DEFAULT 'A',		--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_BomChangeDetail PRIMARY KEY (BomChangeDetailId),
 CONSTRAINT FK1_BomChangeDetail FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId),
 CONSTRAINT FK2_BomChangeDetail FOREIGN KEY (OriginalMtlItemId) REFERENCES PDM.MtlItem(MtlItemId),
 CONSTRAINT FK3_BomChangeDetail FOREIGN KEY (BomChangeId) REFERENCES PDM.BomChange(BomChangeId),
 CONSTRAINT FK4_BomChangeDetail FOREIGN KEY (OriginalBomChangeId) REFERENCES PDM.BomChange(BomChangeId)
)/*Bom變更單單身資料*/


CREATE TABLE PDM.BomChangeSubDetail( /*Bom變更單子單身資料*/ /*ERP BOMTC */
 BomChangeSubDetailId			int					identity, 
 BomChangeDetailId				int					NOT NULL, 
 BomChangeId					int					NOT NULL,
 BomDetailId					int					NOT NULL,						--TC005 元件ID
 MtlItemId						int					NOT NULL,						--元件品號ID
 BomSortNum						nvarchar(4)			NOT	NULL,						--TC004 BOM序號 
 Unit							int					NOT	NULL,						--TC006 單位 ERP端為nvarchar(6)	
 SmallUnit						nvarchar(6)				NULL,						--TC007 小單位
 CompositionQuantity			float				NOT NULL,						--TC008 單位用量
 Base							float				NOT NULL,						--TC009 底數
 LossRate						float				NOT NULL,						--TC010 損耗率
 SubstituteStatus				nvarchar(1)			NOT NULL,						--TC012 取替代料 >> BAS.[Type] 2.取代料件；3.替代料件
 EffectiveDate					datetime				NULL,						--TC013 生效日期
 ExpirationDate					datetime			 	NULL,						--TC014 失效日期
 OptionalPreset					nvarchar(1)			NOT	NULL	DEFAULT 'N',		--TC015 選配預設 Y/N
 StandardCosting				nvarchar(1)				NULL	DEFAULT 'Y',		--TC016 標準成本計算 Y/N
 BomType						nvarchar(1)			NOT	NULL,						--TC019 材料型態 1.2.3
 ComponentType					nvarchar(1)			NOT	NULL	DEFAULT 'N',		--UDF01 材料型態 1.2.3
 OriginalComponentType			nvarchar(1)			NOT	NULL	DEFAULT 'N',		--UDF02 材料型態 1.2.3

 FeedingInterval				int					NOT	NULL,						--TC020 投料時距
 ChangeReason					nvarchar(255)			NULL,						--TC018 變更原因
 NewRemark						nvarchar(255)			NULL,						--TC029 新備註
 OriginalBomDetailId			int						NULL,						--原元件ID
 OriginalMtlItemId				int						NULL,						--TC105 原元件品號ID
 OriginalBomSortNum				nvarchar(4)				NULL,						--TC104 原BOM序號 
 OriginalUnit					int						NULL,						--TC106 原單位 ERP端為nvarchar(6)	
 OriginalSmallUnit				nvarchar(6)				NULL,						--TC107 原小單位
 OriginalCompositionQuantity	float					NULL,						--TC108 原單位用量
 OriginalBase					float					NULL,						--TC109 原底數
 OriginalLossRate				float					NULL,						--TC110 原損耗率
 OriginalSubstituteStatus		nvarchar(1)			    NULL,						--TC112 原取替代料 >> BAS.[Type] 2.取代料件；3.替代料件
 OriginalEffectiveDate			datetime			    NULL,						--TC113 原生效日期
 OriginalExpirationDate			datetime			    NULL,						--TC114 原失效日期
 OriginalOptionalPreset			nvarchar(1)			    NULL     DEFAULT 'N',		--TC115 原選配預設 Y/N
 OriginalStandardCosting		nvarchar(1)			    NULL     DEFAULT 'Y',		--TC116 原標準成本計算 Y/N
 OriginalBomType				nvarchar(1)			    NULL,						--TC119 原材料型態 1.2.3
 OriginalFeedingInterval		int					    NULL,						--TC120 原投料時距
 Originalemark					nvarchar(255)			NULL,						--TC129 原備註
 Provide						nvarchar(1)			    NULL    DEFAULT 'N',		--發放品 改原本BOM的資料-Y,增新BOM-N
 DeleteStatus					nvarchar(1)				NULL,						--是否為刪除元件
 [Status]						nvarchar(2)			NOT NULL	DEFAULT 'A',		--狀態碼
 CreateDate						datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate				datetime				NULL,						--修改日期
 CreateBy						int					NOT NULL,						--新增人員
 LastModifiedBy					int						NULL,						--修改人員
 CONSTRAINT PK_BomChangeSubDetail PRIMARY KEY (BomChangeSubDetailId),
 --CONSTRAINT FK1_BomChangeSubDetail FOREIGN KEY (BomDetailId) REFERENCES PDM.BomDetail(BomDetailId),
 CONSTRAINT FK2_BomChangeSubDetail FOREIGN KEY (BomChangeDetailId) REFERENCES PDM.BomChangeDetail(BomChangeDetailId),
 CONSTRAINT FK3_BomChangeSubDetail FOREIGN KEY (BomChangeId) REFERENCES PDM.BomChange(BomChangeId),
 CONSTRAINT FK4_BomChangeSubDetail FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId)

)/*Bom變更單子單身資料*/

CREATE TABLE PDM.MtlQcItem( /*允收標準*/
 MtlQcItemId				int					identity, 
 MtlItemId					int					NOT NULL,						--品號Id
 QcItemId					int					NOT NULL,						--量測項目Id
 DesignValue				decimal(18, 6)			NULL,						--設計值
 UpperTolerance				decimal(18, 6)			NULL,						--設計值上限公差
 LowerTolerance				decimal(18, 6)			NULL,						--設計值下限公差
 QcItemDesc					nvarchar(100)			NULL,						--檢測備註
 SortNumber					int					NOT NULL,						--排序
 Graphics					nvarchar(20)			NULL,						--圖形
 BallMark					nvarchar(10)			NULL,						--球標
 Unit						nvarchar(20)			NULL,						--單位
 Regulation					nvarchar(255)			NULL,						--相關規範
 Notice						nvarchar(255)			NULL,						--注意事項
 QmmDetailId				int						NULL,						--量測機台編號
 [Status]					nvarchar(2)			NOT NULL	DEFAULT 'A',		--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MtlQcItem PRIMARY KEY (MtlQcItemId),
 CONSTRAINT FK1_MtlQcItem FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId),
 CONSTRAINT AK1_MtlQcItem UNIQUE (MtlItemId,QcItemId),
 CONSTRAINT AK2_MtlQcItem UNIQUE (MtlItemId,SortNumber)
)/*允收標準*/

END /*Bom變更管理*/
