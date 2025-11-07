BEGIN /*基本資料*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DesignPatent') > 0 DROP TABLE RFI.DesignPatent
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'LensFOV') > 0 DROP TABLE RFI.LensFOVData


IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'RequestForInformation') > 0 DROP TABLE RFI.RequestForInformation
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'RfiDetail') > 0 DROP TABLE RFI.RfiDetail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ProductSpec') > 0 DROP TABLE RFI.ProductSpec
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ProductSpecDetail') > 0 DROP TABLE RFI.ProductSpecDetail

IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DemandDesign') > 0 DROP TABLE RFI.DemandDesign
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DemandDesignSpec') > 0 DROP TABLE RFI.DemandDesignSpec
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DemandDesignSpecDetail') > 0 DROP TABLE RFI.DemandDesignSpecDetail

IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'TemplateProdSpec') > 0 DROP TABLE RFI.TemplateProdSpec
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'TemplateProdSpecDetail') > 0 DROP TABLE RFI.TemplateProdSpecDetail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'TemplateSpecParameter') > 0 DROP TABLE RFI.TemplateSpecParameter
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'TemplateSpecParameterDetail') > 0 DROP TABLE RFI.TemplateSpecParameterDetail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'TemplateRfiSignFlow') > 0 DROP TABLE RFI.TemplateRfiSignFlow
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'TemplateDesignSignFlow') > 0 DROP TABLE RFI.TemplateDesignSignFlow
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'RfiSignFlow') > 0 DROP TABLE RFI.RfiSignFlow
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DesignSignFlow') > 0 DROP TABLE RFI.DesignSignFlow
END /*基本資料*/


--******************************************************************************************************************
BEGIN /*基本資料*/
CREATE TABLE RFI.DesignPatent( /*RFI 研發專利管理*/
 DpId						int					identity,
 CompanyId					int					NOT NULL,						--公司別
 ApplicationNo				nvarchar(100)		NOT NULL,						--申請號
 ApplicationDate			datetime				NULL,						--申請日
 Applicant					nvarchar(100)			NULL,						--申請人
 PublicationNo				nvarchar(100)		　	NULL,						--公開號
 PublicationDate			datetime			　	NULL,						--公開日
 AnnouncementNo				nvarchar(100)		　	NULL,						--公告號
 AnnouncementDate			datetime			　	NULL,						--公告日
 ImplementationExample		nvarchar(100)		NOT NULL,						--實施例
 LensQuantity				int					NOT NULL,						--透鏡數量
 AperturePosition			nvarchar(100)		　	NULL,						--光圈位置
 HalfFieldofView			nvarchar(100)		　	NULL,						--半視場角 (HFOV)	
 ImageHeight				nvarchar(100)		　	NULL,						--像高 (ImgH)
 L1RefractivePower			nvarchar(1)			　	NULL,						--L1屈折力
 L1aSurfaceType				nvarchar(1)			　	NULL,						--L1a面型
 L1bSurfaceType				nvarchar(1)			　	NULL,						--L1b面型	
 L2RefractivePower			nvarchar(1)			　	NULL,						--L2屈折力
 L2aSurfaceType				nvarchar(1)			　	NULL,						--L2a面型
 L2bSurfaceType				nvarchar(1)			　	NULL,						--L2b面型	
 L3RefractivePower			nvarchar(1)			　	NULL,						--L3屈折力
 L3aSurfaceType				nvarchar(1)			　	NULL,						--L3a面型
 L3bSurfaceType				nvarchar(1)			　	NULL,						--L3b面型
 L4RefractivePower			nvarchar(1)			　	NULL,						--L4屈折力
 L4aSurfaceType				nvarchar(1)			　	NULL,						--L4a面型
 L4bSurfaceType				nvarchar(1)			　	NULL,						--L4b面型
 L5RefractivePower			nvarchar(1)			　	NULL,						--L5屈折力
 L5aSurfaceType				nvarchar(1)			　	NULL,						--L5a面型
 L5bSurfaceType				nvarchar(1)			　	NULL,						--L5b面型
 L6RefractivePower			nvarchar(1)			　	NULL,						--L6屈折力
 L6aSurfaceType				nvarchar(1)			　	NULL,						--L6a面型
 L6bSurfaceType				nvarchar(1)			　	NULL,						--L6b面型
 L7RefractivePower			nvarchar(1)			　	NULL,						--L7屈折力
 L7aSurfaceType				nvarchar(1)			　	NULL,						--L7a面型
 L7bSurfaceType				nvarchar(1)			　	NULL,						--L7b面型
 L8RefractivePower			nvarchar(1)			　	NULL,						--L8屈折力
 L8aSurfaceType				nvarchar(1)			　	NULL,						--L8a面型
 L8bSurfaceType				nvarchar(1)			　	NULL,						--L8b面型
 L9RefractivePower			nvarchar(1)			　	NULL,						--L9屈折力
 L9aSurfaceType				nvarchar(1)			　	NULL,						--L9a面型
 L9bSurfaceType				nvarchar(1)			　	NULL,						--L9b面型
 LXRefractivePower			nvarchar(1)			　	NULL,						--L10屈折力
 LXaSurfaceType				nvarchar(1)			　	NULL,						--L10a面型
 LXbSurfaceType				nvarchar(1)			　	NULL,						--L10b面型
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DesignPatent PRIMARY KEY (DpId),
  CONSTRAINT FK1_DesignPatent FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)/*RFI 研發專利管理*/

CREATE TABLE RFI.ModelDesign( /*RFI 機種設計值*/
 MdId						int					identity,
 CompanyId					int					NOT NULL,						--公司別
 Model						nvarchar(100)		NOT NULL,						--Model
 Construction				nvarchar(100)		NOT NULL,						--Construction
 EFL						nvarchar(100)		NOT NULL,						--EFL(mm)
 FNo						nvarchar(100)		NOT NULL,						--F/No
 TTL						nvarchar(100)		NOT NULL,						--TTL
 CRA						nvarchar(100)		NOT NULL,						--CRA
 OpticalDistortionFTan		nvarchar(100)			NULL,						--OpticalDistortion
 OpticalDistortionF			nvarchar(100)			NULL,						--OpticalDistortion
 MaxImageCircle				nvarchar(100)			NULL,						--MaxImageCircle
 Sensor						nvarchar(100)			NULL,						--Sensor
 VFOV						nvarchar(100)			NULL,						--VFOV
 HFOV						nvarchar(100)			NULL,						--HFOV
 DFOV						nvarchar(100)			NULL,						--DFOV
 PartName					nvarchar(100)			NULL,						--PartName
 IR							nvarchar(100)			NULL,						--IR
 MechanicalRetainer			nvarchar(100)			NULL,						--Mechanical-Retainer
 MechanicalFBL				nvarchar(100)			NULL,						--Mechanical-F.B.L.
 MechanicalThread			nvarchar(100)			NULL,						--Mechanical-Thread
 IPRating					nvarchar(100)			NULL,						--IPRating
 Stage						nvarchar(100)			NULL,						--Stage
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ModelDesign PRIMARY KEY (MdId),
 CONSTRAINT FK1_ModelDesign FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)/*RFI 機種設計值*/

CREATE TABLE RFI.LensFOV( /*RFI 鏡頭FOV*/
 LfId						int					identity,
 CompanyId					int					NOT NULL,						--公司別
 Model						nvarchar(100)		NOT NULL,						--機種
 FOV						float				NOT NULL,						--FOV
 Realheight					float				NOT NULL,						--Realheight
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_LensFOVData PRIMARY KEY (LfId),
 CONSTRAINT FK1_LensFOVData FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)/*RFI 鏡頭FOV*/

CREATE TABLE RFI.RequestForInformation( /*RFI 單頭*/
 RfiId						int					identity,
 RfiNo						nvarchar(20)		NOT NULL,						--編號
 ProTypeGroupId				int					NOT NULL,						--產品類別群組
 ProductUseId				int					NOT NULL,						--產品應用
 SalesId					int					NOT NULL,						--業務名稱
 ExistFlag					nvarchar(10)		NOT NULL,						--是否存在
 CustomerName				nvarchar(100)		NOT NULL,						--客戶名稱
 CustomerEName				nvarchar(100)			NULL,						--客戶名稱(英)
 CustomerId					int						NULL,						--客戶ID(業務指派) 未關聯
 OrganizaitonType			int					NOT NULL,						--客戶類型
 AnalysisDate				datetime			NOT	NULL,						--分析日期
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼(0待審核 1審核中 2已取號)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_RequestForInformation PRIMARY KEY (RfiId),
 CONSTRAINT FK1_RequestForInformation FOREIGN KEY (ProductUseId) REFERENCES SCM.[ProductUse](ProductUseId),
 CONSTRAINT FK2_RequestForInformation FOREIGN KEY (SalesId) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK3_RequestForInformation FOREIGN KEY (CustomerId) REFERENCES SCM.[Customer](CustomerId),
 CONSTRAINT FK4_RequestForInformation FOREIGN KEY (ProTypeGroupId) REFERENCES SCM.[ProductTypeGroup](ProTypeGroupId),
 CONSTRAINT AK1_RequestForInformation UNIQUE (RfiNo)
)/*RFI 單頭*/

CREATE TABLE RFI.RfiDetail( /*RFI 單身*/
 RfiDetailId				int					identity,
 RfiId						int					NOT NULL,						--RFI 單頭ID
 RfiSequence				nvarchar(4)			NOT NULL,						--流水號
 CompanyId					int					NOT NULL,						--公司ID
 SalesId					int					NOT NULL,						--業務名稱
 ProdDate					date				NOT NULL,						--量產日期
 ProdLifeCycleStart			date				NOT NULL,						--生命週期(起)
 ProdLifeCycleEnd			date				NOT NULL,						--生命週期(迄)
 TargetUnitPrice			decimal(18,6)		NOT NULL,						--目標單價
 LifeCycleQty				int					NOT NULL,						--週期數量(預估)
 Revenue					decimal(18,6)		NOT NULL,						--營收
 RevenueFC					decimal(18,6)		NOT NULL,						--營收(外幣)
 Currency					nvarchar(5)			NOT NULL,						--幣別
 ExchangeRate				decimal(9,6)		NOT NULL	DEFAULT 1,			--匯率
 AdditionalFile				int					    NULL,						--附加檔案(串File)
 --Sales						int						NULL,						--簽核人-業務
 --SuperSales					int						NULL,						--簽核人-業務主管
 --ManagerPlan				int						NULL,						--簽核人-總經理
 --NotifyUser					int						NULL,						--回饋客戶者
 --SalesDesc					nvarchar(2048)			NULL,						--簽核內容-業務
 --SuperSalesDesc				nvarchar(2048)			NULL,						--簽核內容-業務主管
 --ManagerPlanDesc			nvarchar(2048)			NULL,						--簽核內容-總經理
 --CustomerDesc				nvarchar(2048)			NULL,						--簽核內容-反饋客戶
 --ConfirmSSTime				datetime				NULL,						--通知-業務主管
 --ConfirmMPTime				datetime				NULL,						--通知-總經理時間
 --ConfirmCustTime			datetime				NULL,						--反饋-客戶時間
 --ConfirmSTime				datetime				NULL,						--通知-業務時間
 ConfirmFinalTime			datetime				NULL,						--最後完成時間
 Edition					nvarchar(2)			NOT NULL	DEFAULT '00',		--版次
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼(0新單據, 簽核: 1業務、2業務主管、3許總)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_RfiDetail PRIMARY KEY (RfiDetailId),
 CONSTRAINT FK1_RfiDetail FOREIGN KEY (RfiId) REFERENCES RFI.[RequestForInformation](RfiId),
 CONSTRAINT FK2_RfiDetail FOREIGN KEY (CompanyId) REFERENCES BAS.[Company](CompanyId),
 CONSTRAINT FK3_RfiDetail FOREIGN KEY (SalesId) REFERENCES BAS.[User](UserId),
 --CONSTRAINT FK4_RfqDetail FOREIGN KEY (Sales) REFERENCES BAS.[User](UserId),
 --CONSTRAINT FK5_RfiDetail FOREIGN KEY (SuperSales) REFERENCES BAS.[User](UserId),
 --CONSTRAINT FK6_RfiDetail FOREIGN KEY (ManagerPlan) REFERENCES BAS.[User](UserId),
 --CONSTRAINT FK7_RfiDetail FOREIGN KEY (NotifyUser) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK8_RfqDetail FOREIGN KEY (AdditionalFile) REFERENCES BAS.[File](FileId),
 CONSTRAINT AK1_RfqDetail UNIQUE (RfiId, RfiSequence),
 CONSTRAINT AK2_RfqDetail UNIQUE (RfiId, Edition)
)/*RFI 單身*/

CREATE TABLE RFI.ProductSpec( /*RFI 產品規格*/
 ProdSpecId					int					identity,
 RfiDetailId				int					NOT NULL,						--RFI 單身ID
 SpecName					nvarchar(100)		NOT NULL,						--規格名稱(中)
 SpecEName					nvarchar(100)			NULL,						--規格名稱(英)
 SortNumber					int					NOT NULL,						--排序
 ControlType				nvarchar(10)		NOT NULL,						--控制項類型
 [Required]					bit					NOT NULL,						--是否必填(1是 2否)
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼(A啟用 S停用)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ProductSpec PRIMARY KEY (ProdSpecId),
 CONSTRAINT FK1_ProductSpec FOREIGN KEY (RfiDetailId) REFERENCES RFI.[RfiDetail](RfiDetailId),
 CONSTRAINT AK1_ProductSpec UNIQUE (RfiDetailId, SpecName),
 CONSTRAINT AK2_ProductSpec UNIQUE (RfiDetailId, SortNumber)
)/*RFI 產品規格*/

CREATE TABLE RFI.ProductSpecDetail( /*RFI 產品規格特徵*/
 PsDetailId					int					identity,
 ProdSpecId					int					NOT NULL,						--產品規格ID
 FeatureName				nvarchar(50)		NOT NULL,						--規格特徵名稱
 SortNumber					int					NOT NULL,						--排序
 DataType					nvarchar(10)			NULL,						--選項類別
 [Description]				nvarchar(MAX)			NULL,						--文字描述
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼(A啟用 S停用)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ProductSpecDetail PRIMARY KEY (PsDetailId),
 CONSTRAINT FK1_ProductSpecDetail FOREIGN KEY (ProdSpecId) REFERENCES RFI.[ProductSpec](ProdSpecId),
 CONSTRAINT AK1_ProductSpecDetail UNIQUE (ProdSpecId, FeatureName),
 CONSTRAINT AK2_ProductSpecDetail UNIQUE (ProdSpecId, SortNumber)
)/*RFI 產品規格特徵*/



CREATE TABLE RFI.DemandDesign ( /*Design 新設計申請單*/
 DesignId					int					identity,
 RfiId						int					NOT NULL,
 CompanyId					int					NOT NULL,						--公司ID
 DesignNo					nvarchar(20)		NOT NULL,						--設計代號
 DesignSequence				nvarchar(4)			NOT NULL,						--流水號
 AssemblyName				nvarchar(20)		NOT NULL,						--機種名稱
 SalesId					int					NOT NULL,						--申請人
 SuperSalesId				int					NOT NULL,						--單位主管
 AdditionalFile				int						NULL,						--附加檔案 BAS.File
 DesignFile					int						NULL,						--設計附件檔 BAS.File
 --Sales						int						NULL,						--簽核人-簽核人-業務
 --SuperSales					int						NULL,						--簽核人-業務主管
 --Rd							int						NULL,						--簽核人-RD
 --SuperRd					int						NULL,						--簽核人-RD主管
 --ManagerPlan				int						NULL,						--簽核人-總經理
 --SalesDesc					nvarchar(2048)			NULL,						--業務內容
 --SuperSalesDesc				nvarchar(2048)			NULL,						--業主內容
 --RdDesc						nvarchar(2048)			NULL,						--研主內容
 --SuperRdDesc				nvarchar(2048)			NULL,						--研主內容
 --ManagerPlanDesc			nvarchar(2048)			NULL,						--總經理內容
 --ConfirmSSTime				datetime				NULL,						--通知業主時間
 --ConfirmRDTime				datetime				NULL,						--通知研發時間
 --ConfirmSRDTime				datetime				NULL,						--通知研主時間
 --ConfirmMPTime				datetime				NULL,						--通知總經理時間
 --ConfirmSTime				datetime				NULL,						--通知業務時間
 ConfirmFinalTime			datetime				NULL,						--最後完成時間
 Edition					nvarchar(2)			NOT NULL,						--版次
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼(A啟用 S停用)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DemandDesign PRIMARY KEY (DesignId),
 CONSTRAINT FK1_DemandDesign FOREIGN KEY (RfiId) REFERENCES RFI.[RequestForInformation](RfiId),
 --CONSTRAINT FK2_DemandDesign FOREIGN KEY (Sales) REFERENCES BAS.[User](UserId),
 --CONSTRAINT FK3_DemandDesign FOREIGN KEY (SuperSales) REFERENCES BAS.[User](UserId),
 --CONSTRAINT FK4_DemandDesign FOREIGN KEY (SuperRd) REFERENCES BAS.[User](UserId),
 --CONSTRAINT FK5_DemandDesign FOREIGN KEY (ManagerPlan) REFERENCES BAS.[User](UserId),
 --CONSTRAINT FK6_DemandDesign FOREIGN KEY (Rd) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK7_DemandDesign FOREIGN KEY (SalesId) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK8_DemandDesign FOREIGN KEY (SuperSalesId) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK9_DemandDesign FOREIGN KEY (CompanyId) REFERENCES BAS.[Company](CompanyId),
 CONSTRAINT FK10_DemandDesign FOREIGN KEY (AdditionalFile) REFERENCES BAS.[File](FileId),
 CONSTRAINT FK11_DemandDesign FOREIGN KEY (DesignFile) REFERENCES BAS.[File](FileId),
 CONSTRAINT AK1_DemandDesign UNIQUE (RfiId, Edition),
 CONSTRAINT AK2_DemandDesign UNIQUE (RfiId, DesignSequence),
 CONSTRAINT AK3_DemandDesign UNIQUE (RfiId, DesignSequence, AssemblyName)
)/*Design 新設計申請單*/

CREATE TABLE RFI.DemandDesignSpec ( /*Design 設計規格*/
 DdSpecId					int					identity,
 DesignId					int					NOT NULL,						--新設計申請單ID
 ParameterName				nvarchar(255)		NOT NULL,						--設計規格名稱
 SortNumber					int					NOT NULL,						--排序
 ControlType				nvarchar(10)		NOT NULL,						--控制類別
 [Required]					bit					NOT NULL,						--是否必填?
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼(A啟用 S停用)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DemandDesignSpec PRIMARY KEY (DdSpecId),
 CONSTRAINT FK1_DemandDesignSpec FOREIGN KEY (DesignId) REFERENCES RFI.[DemandDesign](DesignId),
 CONSTRAINT AK1_DemandDesignSpec UNIQUE (DesignId, ParameterName),
 CONSTRAINT AK2_DemandDesignSpec UNIQUE (DesignId, SortNumber)
)/*Design 設計規格*/


CREATE TABLE RFI.DemandDesignSpecDetail ( /*Design 設計規格明細*/
 DdsDetailId				int					identity,
 DdSpecId					int					NOT NULL,						--設計規格ID
 PmtDetailName				nvarchar(255)		NOT NULL,						--設計規格特徵
 SortNumber					int					NOT NULL,						--排序
 DataType					nvarchar(10)			NULL,						--選項類別
 [Required]					bit					NOT NULL,						--是否必填?
 RequireFlag				nvarchar(10)			NULL,						--是否必填?
 DesignInput				nvarchar(512)			NULL,						--設計輸入
 [Description]				nvarchar(MAX)			NULL,						--設計描述
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼(A啟用 S停用)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DemandDesignSpecDetail PRIMARY KEY (DdsDetailId),
 CONSTRAINT FK1_DemandDesignSpecDetail FOREIGN KEY (DdSpecId) REFERENCES RFI.[DemandDesignSpec](DdSpecId),
 CONSTRAINT AK1_DemandDesignSpecDetail UNIQUE (DdSpecId, PmtDetailName),
 CONSTRAINT AK2_DemandDesignSpecDetail UNIQUE (DdSpecId, SortNumber)
)/*Design 設計規格明細*/

CREATE TABLE RFI.TemplateProdSpec( /*RFI 產品規格樣板*/
 TempProdSpecId				int					identity,
 ProTypeGroupId				int					NOT NULL,						--產品群組 2階段
 SpecName					nvarchar(100)		NOT NULL,						--規格名稱
 SpecEName					nvarchar(100)			NULL,						--規格名稱(英)
 SortNumber					int					NOT NULL,						--排序
 ControlType				nvarchar(10)		NOT NULL,						--控制類別
 [Required]					bit					NOT NULL,						--是否必填?
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼(A啟用 S停用)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_TemplateProdSpec PRIMARY KEY (TempProdSpecId),
 CONSTRAINT FK1_TemplateProdSpec FOREIGN KEY (ProTypeGroupId) REFERENCES SCM.[ProductTypeGroup](ProTypeGroupId),
 CONSTRAINT AK1_TemplateProdSpec UNIQUE (ProTypeGroupId, SpecName),
 CONSTRAINT AK2_TemplateProdSpec UNIQUE (ProTypeGroupId, SortNumber),
)/*RFI 產品規格樣板*/

CREATE TABLE RFI.TemplateProdSpecDetail( /*RFI 規格特徵樣板*/
 TempPsDetailId				int					identity,						
 TempProdSpecId				int					NOT NULL,						--產品規格ID
 FeatureName				nvarchar(50)		NOT NULL,						--特徵名稱
 SortNumber					int					NOT NULL,						--排序
 DataType					nvarchar(10)			NULL,						--選項類別
 [Description]				nvarchar(50)			NULL,						--其他描述
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼(A啟用 S停用)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_TemplateProdSpecDetail PRIMARY KEY (TempPsDetailId),
 CONSTRAINT FK1_TemplateProdSpecDetail FOREIGN KEY (TempProdSpecId) REFERENCES RFI.[TemplateProdSpec](TempProdSpecId),
 CONSTRAINT AK1_TemplateProdSpecDetail UNIQUE (TempProdSpecId, FeatureName),
 CONSTRAINT AK2_TemplateProdSpecDetail UNIQUE (TempProdSpecId, SortNumber)
)/*RFI 規格特徵樣板*/

CREATE TABLE RFI.TemplateSpecParameter( /*Design 規格樣板*/
 TempSpId					int					identity,
 ProTypeGroupId				int					NOT NULL,						--產品群組 2階段
 ParameterName				nvarchar(100)		NOT NULL,						--設計規格名稱
 SortNumber					int					NOT NULL,						--排序
 ControlType				nvarchar(10)		NOT NULL,						--控制類別
 [Required]					bit					NOT NULL,						--是否必填
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼(A啟用 S停用)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_TemplateSpecParameter PRIMARY KEY (TempSpId),
 CONSTRAINT FK1_TemplateSpecParameter FOREIGN KEY (ProTypeGroupId) REFERENCES SCM.[ProductTypeGroup](ProTypeGroupId),
 CONSTRAINT AK1_TemplateSpecParameter UNIQUE (ProTypeGroupId, ParameterName),
 CONSTRAINT AK2_TemplateSpecParameter UNIQUE (ProTypeGroupId, SortNumber)
)/*Design 規格樣板*/

CREATE TABLE RFI.TemplateSpecParameterDetail( /*Design 規格特徵樣板*/
 TempSpDetailId				int					identity,
 TempSpId					int					NOT NULL,						--產品規格ID
 PmtDetailName				nvarchar(100)		NOT NULL,						--設計規格特徵名稱
 SortNumber					int					NOT NULL,						--排序
 DataType					nvarchar(10)			NULL,						--選項類別
 [Description]				nvarchar(50)			NULL,						--其他描述
 [Required]					bit					NOT NULL,						--是否必填
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼(A啟用 S停用)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_TemplateSpecParameterDetail PRIMARY KEY (TempSpDetailId),
 CONSTRAINT FK1_TemplateSpecParameterDetail FOREIGN KEY (TempSpId) REFERENCES RFI.[TemplateSpecParameter](TempSpId),
 CONSTRAINT AK1_TemplateSpecParameterDetail UNIQUE (TempSpId, PmtDetailName),
 CONSTRAINT AK2_TemplateSpecParameterDetail UNIQUE (TempSpId, SortNumber)
)/*Design 規格特徵樣板*/

CREATE TABLE RFI.TemplateRfiSignFlow( /*市場評估單簽核流程樣板*/
 TempRfiSfId				int					identity,
 ProTypeGroupId				int					NOT NULL,						--產品群組 2階段
 FlowUser					int						NULL,						--簽核人
 FlowJobName				nvarchar(50)			NULL,						--關卡人職稱
 FlowStatus					nvarchar(2)			NOT NULL,						--關卡狀態 BAS.[Status] WHERE StatusSchema='RfiDetail.Status'
 SortNumber					int					NOT NULL,						--排序
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼(A啟用 S停用)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_TemplateRfiSignFlow PRIMARY KEY (TempRfiSfId),
 CONSTRAINT FK1_TemplateRfiSignFlow FOREIGN KEY (ProTypeGroupId) REFERENCES RFI.[ProductTypeGroup](ProTypeGroupId),
 CONSTRAINT FK2_TemplateRfiSignFlow FOREIGN KEY (FlowUser) REFERENCES BAS.[User](UserId),
 CONSTRAINT AK1_TemplateRfiSignFlow UNIQUE (ProTypeGroupId, SortNumber)
)/*市場評估單簽核流程樣板*/

CREATE TABLE RFI.TemplateDesignSignFlow( /*設計申請單簽核流程樣板*/
 TempDesignSfId				int					identity,
 ProTypeGroupId				int					NOT NULL,						--產品群組 2階段
 FlowUser					int						NULL,						--簽核人
 FlowJobName				nvarchar(50)			NULL,						--關卡人職稱
 FlowStatus					nvarchar(2)			NOT NULL,						--關卡狀態 BAS.[Status] WHERE StatusSchema='RfiDesign.Status'
 SortNumber					int					NOT NULL,						--排序
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼(A啟用 S停用)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_TemplateDesignSignFlow PRIMARY KEY (TempDesignSfId),
 CONSTRAINT FK1_TemplateDesignSignFlow FOREIGN KEY (ProTypeGroupId) REFERENCES RFI.[ProductTypeGroup](ProTypeGroupId),
 CONSTRAINT FK2_TemplateDesignSignFlow FOREIGN KEY (FlowUser) REFERENCES BAS.[User](UserId),
 CONSTRAINT AK1_TemplateDesignSignFlow UNIQUE (ProTypeGroupId, SortNumber)
)/*設計申請單簽核流程樣板*/

CREATE TABLE RFI.RfiSignFlow( /*市場評估單簽核流程*/
 RfiSfId					int					identity,
 RfiDetailId				int					NOT NULL,						--評估單身
 FlowUser					int					NOT NULL,						--關卡人
 AdditionalFile				int						NULL,						--附件上傳
 ArriveDateTime				datetime				NULL,						--到達時間
 SignUser					int						NULL,						--簽核人 
 SignJobName				nvarchar(50)			NULL,						--簽核人職稱
 SignDateTime				datetime				NULL,						--簽核時間
 SignDesc					nvarchar(2048)			NULL,						--簽核內容
 SortNumber					int					NOT NULL,						--排序
 Approved					nvarchar(2)				NULL,						--簽核狀態 BAS.[Status] WHERE StatusSchema='Boolean' N|Y
 UrgentSign					int					NOT NULL,						--催簽次數
 Edition					nvarchar(2)			NOT NULL	DEFAULT '00',		--版次
 [Status]					nvarchar(2)			NOT NULL,						--關卡狀態 BAS.[Status] WHERE StatusSchema='RfiDetail.FlowStatus' 業務|業務主管|研發|研發主管|副總|總經理
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_RfiSignFlow PRIMARY KEY (TempRfiSfId),
 CONSTRAINT FK1_RfiSignFlow FOREIGN KEY (RfiDetailId) REFERENCES RFI.[RfiDetail](RfiDetailId),
 CONSTRAINT FK2_RfiSignFlow FOREIGN KEY (FlowUser) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK3_RfiSignFlow FOREIGN KEY (SignUser) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK4_RfiSignFlow FOREIGN KEY (AdditionalFile) REFERENCES BAS.[File](FileId),
 CONSTRAINT AK1_RfiSignFlow UNIQUE (RfiDetailId, SortNumber)
)/*市場評估單簽核流程*/

CREATE TABLE RFI.DesignSignFlow( /*設計申請單簽核流程*/
 DesignSfId					int					identity,
 DesignId					int					NOT NULL,						--設計申請單
 FlowUser					int					NOT NULL,						--關卡人
 AdditionalFile				int						NULL,						--附件上傳
 ArriveDateTime				datetime				NULL,						--到達時間
 SignUser					int						NULL,						--簽核人 
 SignJobName				nvarchar(50)			NULL,						--簽核人職稱
 SignDateTime				datetime				NULL,						--簽核時間
 SignDesc					nvarchar(2048)			NULL,						--簽核內容
 SortNumber					int					NOT NULL,						--排序
 Approved					nvarchar(2)				NULL,						--簽核狀態 BAS.[Status] WHERE StatusSchema='Boolean' N|Y
 UrgentSign					int					NOT NULL,						--催簽次數
 [Status]					nvarchar(2)			NOT NULL,						--關卡狀態 BAS.[Status] WHERE StatusSchema='RfiDesign.FlowStatus' 業務|業務主管|研發|研發主管|副總|總經理
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DesignSignFlow PRIMARY KEY (DesignSfId),
 CONSTRAINT FK1_DesignSignFlow FOREIGN KEY (DesignId) REFERENCES RFI.[RfiDetail](DesignId),
 CONSTRAINT FK2_DesignSignFlow FOREIGN KEY (FlowUser) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK3_DesignSignFlow FOREIGN KEY (SignUser) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK4_DesignSignFlow FOREIGN KEY (AdditionalFile) REFERENCES BAS.[File](FileId),
 CONSTRAINT AK1_DesignSignFlow UNIQUE (DesignId, SortNumber)
)/*設計申請單簽核流程*/

END /*基本資料*/
