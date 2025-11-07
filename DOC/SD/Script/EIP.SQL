BEGIN /*基本資料*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'RfqProductClass' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'SCM'))> 0 DROP TABLE SCM.RfqProductClass
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'RfqProductType' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'SCM'))> 0 DROP TABLE SCM.RfqProductType
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'RfqPackageType' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'SCM'))> 0 DROP TABLE SCM.RfqProductType
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ProductUse' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'SCM'))> 0 DROP TABLE SCM.ProductUse

IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DfmItem' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'PDM'))> 0 DROP TABLE PDM.DfmItem
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DfmItemCategory' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'PDM'))> 0 DROP TABLE PDM.DfmItemCategory
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DfmMaterial' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'PDM'))> 0 DROP TABLE PDM.DfmMaterial
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DfmCategoryMaterial' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'PDM'))> 0 DROP TABLE PDM.DfmCategoryMaterial
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DfmCategoryTemplate' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'PDM'))> 0 DROP TABLE PDM.DfmCategoryTemplate

IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'NotifyRole' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'EIP'))> 0 DROP TABLE EIP.NotifyRole
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'NotifyUser' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'EIP'))> 0 DROP TABLE EIP.NotifyUser
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DocNotify' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'EIP'))> 0 DROP TABLE EIP.DocNotify
END /*基本資料*/

BEGIN /*電商平台*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Member' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'EIP'))> 0 DROP TABLE EIP.Member
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MemberOrganization' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'EIP'))> 0 DROP TABLE EIP.MemberOrganization
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MemberLoginKey' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'EIP'))> 0 DROP TABLE EIP.MemberLoginKey
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'LanguageExchange' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'EIP'))> 0 DROP TABLE EIP.LanguageExchange
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'RequestForQuotation' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'SCM'))> 0 DROP TABLE SCM.RequestForQuotation
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'RfqDetail' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'SCM'))> 0 DROP TABLE SCM.RfqDetail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'RfqLineSolution' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'SCM'))> 0 DROP TABLE SCM.RfqLineSolution
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'RfqPackage' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'SCM'))> 0 DROP TABLE SCM.RfqPackage
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DesignForManufacturing' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'PDM'))> 0 DROP TABLE PDM.DesignForManufacturing
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DfmDetail' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'PDM'))> 0 DROP TABLE PDM.DfmDetail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DfmDetailLog' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'PDM'))> 0 DROP TABLE PDM.DfmDetailLog
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DfmQuotationItem' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'PDM'))> 0 DROP TABLE PDM.DfmQuotationItem
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DfmQuotationItemLog' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'PDM'))> 0 DROP TABLE PDM.DfmQuotationItemLog
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DfmQiMaterial' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'PDM'))> 0 DROP TABLE PDM.DfmQiMaterial
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DfmQiMaterialLog' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'PDM'))> 0 DROP TABLE PDM.DfmQiMaterialLog
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DfmQiOSP' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'PDM'))> 0 DROP TABLE PDM.DfmQiOSP
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DfmQiOSPLog' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'PDM'))> 0 DROP TABLE PDM.DfmQiOSPLog
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DfmQiProcess' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'PDM'))> 0 DROP TABLE PDM.DfmQiProcess
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DfmQiProcessLog' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'PDM'))> 0 DROP TABLE PDM.DfmQiProcessLog
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DfmQiSolution' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'PDM'))> 0 DROP TABLE PDM.DfmQiSolution
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DfmQiSolutionProcess' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'PDM'))> 0 DROP TABLE PDM.DfmQiSolutionProcess
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DfmTemplate' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'PDM'))> 0 DROP TABLE PDM.DfmTemplate
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DfmTemplateParameter' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'PDM'))> 0 DROP TABLE PDM.DfmTemplateParameter
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'QuotationHtmlElement' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'SCM'))> 0 DROP TABLE SCM.QuotationHtmlElement
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'QuotationHtmlElementDetail' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'SCM'))> 0 DROP TABLE SCM.QuotationHtmlElementDetail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'QuotationItem' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'SCM'))> 0 DROP TABLE SCM.QuotationItem
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'QuotationItemElement' AND uid = (SELECT SCHEMA_ID FROM sys.schemas WHERE name = 'SCM'))> 0 DROP TABLE SCM.QuotationItemElement


END /*電商平台*/

--******************************************************************************************************************
BEGIN /*基本資料*/
CREATE TABLE SCM.RfqProductClass( /*RFQ產品大類別  ex:鏡片 */ 
 RfqProClassId				int					identity,
 CompanyId					int					NOT NULL,						--公司別
 RfqProductClassName		nvarchar(100)		NOT NULL,						--RFQ 產品類別 
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼(S/A:關閉/啟用)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_RfqProductClass PRIMARY KEY (RfqProClassId),
 CONSTRAINT FK1_Company FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
) /*RFQ產品大類別  ex:鏡片 */ 

CREATE TABLE SCM.RfqProductType( /*RFQ產品小類別  G鏡片 P鏡片 */
 RfqProTypeId				int					identity,
 RfqProClassId				int					NOT NULL,
 RfqProductTypeName			nvarchar(100)		NOT NULL,						--RFQ 產品類別 
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼(S/A:關閉/啟用)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_RfqProductType PRIMARY KEY (RfqProTypeId),
 CONSTRAINT FK1_RfqProductType FOREIGN KEY (RfqProClassId) REFERENCES SCM.RfqProductClass(RfqProClassId)
) /*RFQ產品小類別  G鏡片 P鏡片 */

CREATE TABLE SCM.RfqPackageType( /*RFQ包裝種類*/
 RfqPkTypeId				int					identity,
 RfqProClassId				int					NOT NULL,						--Foreign Key
 PackagingMethod			nvarchar(50)		NOT NULL,						--包裝方式
 SustSupplyStatus			nvarchar(2)			NOT NULL,						--客供盛盤(Y/N:是/否)
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼(S/A:關閉/啟用)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_RfqPackagType PRIMARY KEY (RfqPkTypeId),
 CONSTRAINT FK1_RfqPackagType FOREIGN KEY (RfqProClassId) REFERENCES SCM.RfqProductClass(RfqProClassId)
) /*RFQ包裝種類*/

CREATE TABLE SCM.ProductUse( /*RFQ產品用途  車載、醫療 */
 ProductUseId				int					identity,
 ProductUseNo				nvarchar(50)		NOT NULL,						--RFQ 產品用途代號 
 ProductUseName				nvarchar(100)		NOT NULL,						--RFQ 產品用途名稱 
 TypeOne					nvarchar(2)			NOT NULL,						--RFI 產品類別編碼
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼(S/A:關閉/啟用)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ProductUse PRIMARY KEY (ProductUseId)
) /*RFQ產品用途  G鏡片 P鏡片 */

CREATE TABLE PDM.DfmItem( /*DFM 參數*/
 DfmItemId					int					identity,
 CompanyId					int					NOT NULL,						--公司別 
 DfmItemNo					nvarchar(20)		NOT NULL,						--參數編號 
 DfmItemName				nvarchar(20)		NOT NULL,						--參數名稱
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼(S/A:關閉/啟用)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DfmItem PRIMARY KEY (DfmItemId),
 CONSTRAINT FK1_DfmItem FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
) /*DFM 參數*/

CREATE TABLE PDM.DfmItemCategory( /*DFM項目種類*/
 DfmItemCategoryId			int					identity,
 ModeId						int					NOT NULL,						--所屬生產模式
 DfmItemCategoryNo			nvarchar(20)		NOT NULL,						--DFM項目種類代碼
 DfmItemCategoryName		nvarchar(100)		NOT NULL,						--DFM項目種類名稱
 DfmItemCategoryDesc		nvarchar(100)			NULL,						--DFM項目種類描述
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DfmItemCategory PRIMARY KEY (DfmItemCategoryId),
 CONSTRAINT FK1_DfmItemCategory FOREIGN KEY (ModeId) REFERENCES MES.ProdMode(ModeId)
) /*DFM項目種類*/

CREATE TABLE PDM.DfmMaterial( /*DFM物料代碼*/
 DfmMaterialId				int					identity,
 CompanyId					int					NOT NULL,						--公司別 
 DfmMaterialNo				nvarchar(20)		NOT NULL,						--DFM物料代碼
 DfmMaterialName			nvarchar(100)		NOT NULL,						--DFM物料名稱
 DfmMaterialDesc			nvarchar(100)			NULL,						--DFM物料描述
 DfmMaterialMoney			float					NULL,						--DFM物料金額
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DfmMaterial PRIMARY KEY (DfmMaterialId),
 CONSTRAINT FK1_DfmMaterial FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
) /*DFM物料代碼*/

CREATE TABLE PDM.DfmCategoryMaterial( /*DFM項目種類物料*/
 DfmCategoryMaterialId		int					identity,
 DfmMaterialId				int					NOT NULL,
 DfmItemCategoryId			int					NOT NULL,
 DfmMaterialDesc			nvarchar(100)			NULL,						--DFM物料描述
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DfmCategoryMaterial PRIMARY KEY (DfmCategoryMaterialId),
 CONSTRAINT FK1_DfmCategoryMaterial FOREIGN KEY (DfmMaterialId) REFERENCES PDM.DfmMaterial(DfmMaterialId)
) /*DFM項目種類物料*/

CREATE TABLE PDM.DfmCategoryTemplate( /*DFM樣板*/
 DfmCtId					int					identity,
 DfmItemCategoryId			int					NOT NULL,						--項目種類
 QiMaterialTemplateDefault	nvarchar(MAX)			NULL,						--物料樣板標頭預設
 QiMaterialTemplate			nvarchar(MAX)			NULL,						--物料樣板
 QiProcessTemplateDefault	nvarchar(MAX)			NULL,						--製程樣板標頭預設
 QiProcessTemplate			nvarchar(MAX)			NULL,						--製程樣板
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DfmCategoryTemplate PRIMARY KEY (DfmCtId),
 CONSTRAINT FK1_DfmCategoryTemplate FOREIGN KEY (DfmItemCategoryId) REFERENCES PDM.DfmItemCategory(DfmItemCategoryId)
) /*DFM物料代碼*/

CREATE TABLE EIP.NotifyRole( /*通知角色資料*/
 RoleId						int					identity,
 RoleName					nvarchar(50)		NOT NULL,						--角色名稱
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_NotifyRole PRIMARY KEY (RoleId),
 CONSTRAINT AK1_NotifyRole UNIQUE (RoleName)
) /*通知角色資料*/

CREATE TABLE EIP.NotifyUser( /*通知角色使用者*/
 RoleId						int					NOT NULL,						--所屬角色Id
 UserId						int					NOT NULL,						--所屬使用者Id
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 CreateBy					int					NOT NULL,						--新增人員
 CONSTRAINT PK_NotifyUser PRIMARY KEY (RoleId, UserId),
 CONSTRAINT FK1_NotifyUser FOREIGN KEY (RoleId) REFERENCES EIP.NotifyRole(RoleId),
 CONSTRAINT FK2_NotifyUser FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId)
) /*通知角色使用者*/

CREATE TABLE EIP.DocNotify( /*單據通知權限資料*/
 DocNotifyId				int					identity,						
 RoleId						int					NOT NULL,						--所屬角色Id
 DocType					nvarchar(5)			NOT NULL,						--單據類型 -- RFQ、DFM
 SubProdId					int					NOT NULL,						--產品Id -- RFQ:SCM.RfqProductType ; DFM:PDM.DfmItemCategory
 SubProdType				nvarchar(20)			NULL,						--報價項目 -- NULL,Material,Process
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 CreateBy					int					NOT NULL,						--新增人員
 CONSTRAINT PK_DocNotify PRIMARY KEY (DocNotifyId)
) /*單據通知權限資料*/

CREATE TABLE SCM.QuotationTag( /*報價單屬性標籤*/
 QtId						int					identity,
 CompanyId					int					NOT NULL,						--公司別 
 TagNo						nvarchar(20)		NOT NULL,						--標籤編號 
 TagName					nvarchar(20)		NOT NULL,						--標籤名稱
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼(S/A:關閉/啟用)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QuotationTag PRIMARY KEY (QtId),
 CONSTRAINT FK1_QuotationTag FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
) /*DFM 參數*/

END/*基本資料*/

BEGIN /*電商平台*/
--CREATE TABLE EIP.Member( /*註冊資訊-公司內員工 --240410改使用Official.EIP.Member*/
-- MemberId					int					identity,
-- MemberName				nvarchar(30)		NOT NULL,						--使用者名稱 
-- MemberEmail				nvarchar(100)		NOT	NULL,						--電子郵件 
-- [Password]				nvarchar(128)		NOT NULL,						--密碼
-- Description				nvarchar(100)			NULL,						--描述
-- OrgShortName				nvarchar(100)		NOT NULL,						--公司簡稱
-- Address					nvarchar(100)			NULL,						--公司地址
-- ContactName				nvarchar(100)			NULL,						--公司聯絡人
-- ContactPhone				nvarchar(128)		NOT	NULL,						--聯絡電話
-- ContactEmail				nvarchar(200)			NULL,						--公司電子郵件
-- CertCode					nvarchar(100)		NOT	NULL,						--認證碼(用於第一次登入) 
-- PasswordStatus			nvarchar(2)			NOT NULL	DEFAULT 'N',		--更改密碼狀態
-- PasswordMistake			int					NOT NULL	DEFAULT 0, 			--密碼錯誤次數
-- [Status]					nvarchar(2)			NOT NULL,						--狀態碼 --S:停用、A:已審核、T:待審核
-- CreateDate				datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
-- LastModifiedDate			datetime				NULL,						--修改日期
-- CreateBy					int					NOT NULL,						--新增人員
-- LastModifiedBy			int						NULL,						--修改人員
-- CONSTRAINT PK_Member PRIMARY KEY (MemberId)
--) /*註冊資訊-公司內員工*/

CREATE TABLE EIP.MemberOrganization( /*註冊資訊-公司*/
 OrgId						int					identity,
 MemberId					int					NOT NULL,
 OrganizaitonTypeId			int					NOT NULL,						--客戶或供應商ID
 OrganizaitonType			int					NOT NULL,						--1.客戶, 2.供應商
 OrganizationCode			nvarchar(30)			NULL,						--註冊代號(ERP客戶或供應商代碼)
 OrganizaitonScale			nvarchar(30)		NOT NULL,						--規模(串BAS.Type)【非公發】【上市櫃公司】【高風險】
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MemberOrganization PRIMARY KEY (OrgId)
 --CONSTRAINT FK1_MemberOrganization FOREIGN KEY (MemberId) REFERENCES EIP.Member(MemberId)
) /*註冊資訊-公司*/

--CREATE TABLE EIP.MemberLoginKey( /*使用者登入金鑰 --240410改使用Official.EIP.MemberLoginKey*/
-- KeyId						int					identity,
-- MemberId					int					NOT NULL,						--所屬使用者
-- KeyText					nvarchar(256)		NOT NULL,						--金鑰
-- ExpirationDate				datetime			NOT NULL,						--過期時間
-- LoginIP					nvarchar(50)		NOT NULL,						--登入IP
-- CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
-- LastModifiedDate			datetime				NULL,						--修改日期
-- CreateBy					int					NOT NULL,						--新增人員
-- LastModifiedBy				int						NULL,						--修改人員
-- CONSTRAINT PK_MemberLoginKey PRIMARY KEY(KeyId),
-- CONSTRAINT AK1_MemberLoginKey UNIQUE(KeyText),
-- CONSTRAINT FK1_MemberLoginKey FOREIGN KEY (MemberId) REFERENCES EIP.Member(MemberId)
--) /*使用者登入金鑰*/

CREATE TABLE EIP.LanguageExchange( /*語言切換設定 */ 
 LeId						int					identity,
 LanguageNo					nvarchar(100)		NOT NULL,						--語系 參考Status
 ExchangeNo					nvarchar(100)		NOT NULL,						--代表文字
 Content					nvarchar(MAX)		NOT NULL,						--內容
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼(S/A:關閉/啟用)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_LanguageExchange PRIMARY KEY (LeId)
) /*語言切換設定 */ 

CREATE TABLE SCM.RequestForQuotation( /*RFQ 單頭*/
 RfqId						int					identity,
 RfqNo						nvarchar(20)		NOT NULL,						--需求單號 
 AssemblyName				nvarchar(200)		NOT NULL,						--機種名稱
 ProductUseId				int					NOT NULL,						--產品應用(串SCM.ProductUse)
 --MemberType					int					NOT NULL,						--1.EIP , 2.BM
 MemberId					int						NULL,						--EIP.Member.MemberId (按照電商註冊人員ID) --240410改使用Official.EIP.Member
 UserId						int						NULL,						--BM.BAS.User.UserId
 OrganizaitonType			int						NULL,						--1.客戶, 2.供應商
 CustomerId					int						NULL,						--客戶ID(業務指派)
 CustomerName				nvarchar(200)			NULL,						--客戶名稱(客戶輸入)
 SupplierId					int						NULL,						--供給商ID
 SupplierName				nvarchar(200)			NULL,						--供給商名稱
 ContactPerson				nvarchar(30)			NULL,						--聯絡人
 ContactPhone				nvarchar(20)			NULL,						--聯絡電話
 [Status]					nvarchar(2)			NOT NULL,						--0未送出 1詢價中 2已報價
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員(業務)
 LastModifiedBy				int						NULL,						--修改人員(業務)
 CONSTRAINT PK_RequestForQuotation PRIMARY KEY (RfqId),
 CONSTRAINT FK1_RequestForQuotation FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK2_RequestForQuotation FOREIGN KEY (ProductUseId) REFERENCES SCM.ProductUse(ProductUseId),
 CONSTRAINT FK3_RequestForQuotation FOREIGN KEY (CustomerId) REFERENCES SCM.Customer(CustomerId)
) /*RFQ 單頭*/

CREATE TABLE SCM.RfqDetail( /*RFQ 單身*/
 RfqDetailId				int					identity,
 RfqId						int					NOT NULL,						--Foreign Key
 CompanyId					int					NOT NULL,						--公司別 
 RfqProTypeId				int					NOT NULL,						--產品類型(白物鏡片) 
 RfqSequence				nvarchar(4)			NOT NULL,						--流水號
 MtlName					nvarchar(100)		NOT NULL,						--品名/P1~P8
 CustProdDigram				int					NOT NULL,						--客戶產品圖(串File) 有需要再開Table
 PlannedOpeningDate			Date					NULL,						--預計開案日
 PrototypeQty				int					NOT NULL,						--試作需求數量
 UomId						int						NULL,						--單位
 ProtoSchedule				nvarchar(30)			NULL,						--試作時程(48H(2天)、168H(7天)..等)
 MassProductionDemand		nvarchar(1)				NULL,						--量產需求 Y/N
 KickOffType				nvarchar(50)			NULL,						--開案型態(幾P幾黑)
 PlasticName				nvarchar(100)			NULL,						--塑料品名
 OutsideDiameter			nvarchar(100)			NULL,						--部品外徑大小
 ProdLifeCycleStart			Date				NOT NULL,						--生命週期(起)
 ProdLifeCycleEnd			Date				NOT NULL,						--生命週期(迄)
 LifeCycleQty				int					NOT NULL,						--週期數量(預估)
 DemandDate					Date				NOT NULL,						--需求日期
 CoatingFlag				nvarchar(1)			NOT NULL	DEFAULT 'N',		--是否鍍膜()
 Currency					nvarchar(5)			NOT NULL,						--幣別 >> BAS.[Type] NT$：台幣；CNY：人民幣；USD：美元；JPY：日元；HKD：港元；EUR：歐元；GBP：英鎊；CHF：瑞士法郎
 PortOfDelivery				nvarchar(256)			NULL,						--交貨地點
 SalesId					int					    NULL,						--業務人員(由業務主管指派)
 AdditionalFile				int					    NULL,						--附加檔案(串File)
 QuotationFile				int					    NULL,						--報價單檔案(串File)
 [Description]				nvarchar(100)			NULL,						--描述
 ConfirmVPTime				datetime				NULL,						--通知副總時間
 ConfirmSalesTime			datetime				NULL,						--指派人員後，通知Sales時間
 ConfirmRdTime				datetime				NULL,						--單據核准後，通知RD時間
 ConfirmCustTime			datetime				NULL,						--報價單核准後，通知Cust時間
 MonthlyQty					int						NULL,						--月需求量
 QuotationRemark			nvarchar(512)			NULL,						--報價單詳細資訊
 TagList					nvarchar(MAX)			NULL,						--報價單屬性標籤
 Edition					nvarchar(2)			NOT NULL 	DEFAULT '000',		--版次
 BaseCavities				int					NOT NULL,						--模架穴数
 InsertCavities				int					NOT NULL,						--模仁穴数
 CoreThickness				float					NULL,						--芯厚
 CommonMode					nvarchar(1)			NOT NULL,						--是否共模
 DocType					nvarchar(1)			NOT	NULL,						--單據狀態 >> BAS.[Type]:RfqDetail.DocType E：生效(Effective)；F：失效(Failure)
 ExpirationDate				datetime			 	NULL,						--失效日期
 QuotationStatus			nvarchar(2)			NOT NULL	DEFAULT 'N',		--報價單確認狀態 N：未確認；Y：已確認；F：最終確認；B：議價中; S:尚未生成
 CycleTimeAdoption			nvarchar(2)				NULL,						--成形週期採用:S:不運用,MP:人工流程,AI:系統計算,AN:新材料系統計算,NP:新材料人工流程
 AICycleTime				float					NULL,						--AI預測成形週期
 MPCycleTime				float					NULL,						--人工預測成形週期
 RealCycleTime				float					NULL,						--實際成形週期
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼(Create, Assign Sales, Sales Confirm RFQ, RD Confirm DFM, IE Confirm DFM Process, Sales Confirm Quotation)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_RfqDetail PRIMARY KEY (RfqDetailId),
 CONSTRAINT FK1_RfqDetail FOREIGN KEY (RfqId) REFERENCES SCM.RequestForQuotation(RfqId),
 CONSTRAINT FK2_RfqDetail FOREIGN KEY (SalesId) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK3_RfqDetail FOREIGN KEY (RfqProTypeId) REFERENCES SCM.RfqProductType(RfqProTypeId),
 CONSTRAINT FK4_RfqDetail FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK5_RfqDetail FOREIGN KEY (CustProdDigram) REFERENCES BAS.[File](FileId),
 CONSTRAINT FK6_RfqDetail FOREIGN KEY (AdditionalFile) REFERENCES BAS.[File](FileId),
 CONSTRAINT FK7_RfqDetail FOREIGN KEY (QuotationFile) REFERENCES BAS.[File](FileId)
) /*RFQ 單身*/

CREATE TABLE SCM.RfqTagList( /*RFQ 單身*/
 RtlId						int					identity,
 RfqDetailId				int					NOT NULL,						--RFQ單身Id
 QtId						int					NOT NULL,						--報價單屬性標籤
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_RfqTagList PRIMARY KEY (RtlId),
 CONSTRAINT FK1_RfqTagList FOREIGN KEY (RfqDetailId) REFERENCES SCM.RfqDetail(RfqDetailId),
 CONSTRAINT FK2_RfqTagList FOREIGN KEY (QtId) REFERENCES SCM.QuotationTag(QtId)
) /*RFQ 單身*/

CREATE TABLE SCM.RfqLineSolution( /*報價方案數量*/
 RfqLineSolutionId			int					identity,
 RfqDetailId				int					NOT NULL,						--Foreign Key
 SortNumber					int					NOT NULL,						--順序
 SolutionQty				int					NOT NULL,						--方案數量
 PeriodicDemandType			nvarchar(2)			NOT NULL,						--週期需求(M/B:月/批)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_RfqLineSolution PRIMARY KEY (RfqLineSolutionId),
 CONSTRAINT FK1_RfqLineSolution FOREIGN KEY (RfqDetailId) REFERENCES SCM.RfqDetail(RfqDetailId)
) /*報價方案數量*/

CREATE TABLE SCM.RfqPackage( /*RFQ包裝方式*/
 RfqPkId					int					identity,
 RfqDetailId				int					NOT NULL,						--Foreign Key 
 RfqPkTypeId				int					NOT NULL,						--RFQ包裝種類 
 SustSupplyStatus			nvarchar(2)			NOT NULL,						--客供盛盤(Y/N:是/否)
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼(S/A:關閉/啟用)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_RfqPackage PRIMARY KEY (RfqPkId),
 CONSTRAINT FK1_RfqPackage FOREIGN KEY (RfqDetailId) REFERENCES SCM.RfqDetail(RfqDetailId),
 CONSTRAINT FK2_RfqPackage FOREIGN KEY (RfqPkTypeId) REFERENCES SCM.RfqPackageType(RfqPkTypeId)
) /*RFQ包裝方式*/


BEGIN/*報價單管理*/

CREATE TABLE SCM.QuotationHtmlElement( /*報價單樣板項目*/
 QeId						int					identity,				
 CompanyId					int					NOT NULL,						--公司別
 RfqProTypeId				int					NOT NULL,						--RFQ產品類別ID RfqProductType
 CoatingFlag				nvarchar(2)				NULL	DEFAULT 'N',		--是否鍍膜(Y/N)
 QuotationData				nvarchar(MAX)			NULL,						--報價單樣版=>參考 SCM.QuotationHtmlElement,SCM.QuotationHtmlElementDetail
 HtmlColumns				nvarchar(20)		NOT NULL,						--項目欄位
 HtmlColName				nvarchar(100)		NOT NULL,						--項目欄位名稱
 ColRemark					nvarchar(255)			NULL,						--欄位備註
 Charge						nvarchar(100)		NOT NULL,						--維護人員
 [Status]					nvarchar(2)			NOT NULL	DEFAULT 'A',		--啟用狀態
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QuotationHtmlElement PRIMARY KEY (QeId),
 CONSTRAINT FK1_QuotationHtmlElement FOREIGN KEY (RfqProTypeId) REFERENCES SCM.RfqProductType(RfqProTypeId)
)/*報價單網頁元素 參考用表*/

CREATE TABLE SCM.QuotationHtmlElementDetail( /*報價單樣板項目明細*/
 QedId						int					identity,				
 QeId						int					NOT NULL,						--報價單樣板項目ID QuotationHtmlElement
 Sort						int					NOT NULL,						--排序
 DetailNo					nvarchar(2)			NOT NULL,						--項目明細代號
 DetailName					nvarchar(100)		NOT NULL,						--項目明細名稱
 ColumnSetting				nvarchar(2)			NOT NULL,						--欄位格式 L:文字欄位(label),L%:採用%數的文字欄位(公式上會有卡控設定),C:%數欄位,F:公式欄位,S:選擇欄位,
 Formula					nvarchar(MAX)			NULL,						--公式
 DataFrom					nvarchar(MAX)			NULL,						--資料來源
 ElementRemark				nvarchar(255)			NULL,						--欄位備註
 Flag						nvarchar(2)				NULL	DEFAULT '',			--使用設定 代碼 FR:要回寫最終報價單資料表,OP:開機費 PDF輸出開機欄位用
 [Status]					nvarchar(2)			NOT NULL	DEFAULT 'A',		--啟用狀態
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QuotationHtmlElementDetail PRIMARY KEY (QedId),
 CONSTRAINT FK1_QuotationHtmlElementDetail FOREIGN KEY (QeId) REFERENCES SCM.QuotationHtmlElement(QeId)
)/*報價單網頁元素明細 參考用表*/

CREATE TABLE SCM.QuotationItem( /*報價單項目*/
 QiId						int					identity,				
 RfqDetailId				int					NOT NULL,						--RFQ Detail ID
 HtmlColumns				nvarchar(20)		NOT NULL,						--網頁項目欄位
 HtmlColName				nvarchar(100)		NOT NULL,						--網頁項目欄位名稱
 Charge						nvarchar(100)		NOT NULL,						--維護人員
 ConfirmUserId				int						NULL,						--確認者
 ConfirmStatus				nvarchar(1)			NOT NULL	DEFAULT 'N',		--確認狀態 >> BAS.[Status] ConfirmStatus N：未確認；Y：已確認
 ColRemark					nvarchar(255)			NULL,						--欄位備註
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QuotationItem PRIMARY KEY (QiId),
 CONSTRAINT FK1_QuotationItem FOREIGN KEY (RfqDetailId) REFERENCES SCM.RfqDetail(RfqDetailId)
)/*報價單網頁項目 參考用表*/

CREATE TABLE SCM.QuotationItemElement( /*報價單項目明細*/
 QieId						int					identity,				
 QiId						int					NOT NULL,						--報價單項目ID QuotationItemElement
 Sort						int					NOT NULL,						--排序
 ItemElementNo				nvarchar(3)			NOT NULL,						--網頁項目明細代號
 ItemElementName			nvarchar(100)		NOT NULL,						--網頁項目明細名稱
 QuoteValue					nvarchar(100)			NULL,						--網頁元素的輸入值
 ColumnSetting				nvarchar(2)			NOT NULL,						--欄位格式 L:文字欄位(label),F:公式欄位,S:選擇欄位,
 Formula					nvarchar(MAX)			NULL,						--公式
 DataFrom					nvarchar(MAX)			NULL,						--資料來源
 ElementRemark				nvarchar(255)			NULL,						--欄位備註
 Flag						nvarchar(50)			NULL,						--標示 OEE:稼動率;Cavity:穴號;MoldingCycle:成形周期;SC:開機費;FR:議價後回寫
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QuotationItemElement PRIMARY KEY (QieId),
 CONSTRAINT FK1_QuotationItemElement FOREIGN KEY (QiId) REFERENCES SCM.QuotationItem (QiId)
)/*報價單網頁項目元素 參考用表*/

CREATE TABLE SCM.QuotationFinalPrice( /*最終報價單價格紀錄*/
 QfpId						int					identity,				
 RfqDetailId				int					NOT NULL,						--RFQ Detail ID
 FinalPrice					float				NOT NULL,						--價格
 GrossMargin				float				NOT NULL,						--毛利率
 ConfirmStatus				nvarchar(1)			NOT NULL	DEFAULT 'N',		--確認狀態 >> BAS.[Status] ConfirmStatus N：未確認；Y：已確認；H：歷史紀錄
 Remark						nvarchar(255)			NULL,						--欄位備註
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QuotationFinalPrice PRIMARY KEY (QfpId),
 CONSTRAINT FK1_QuotationFinalPrice FOREIGN KEY (RfqDetailId) REFERENCES SCM.RfqDetail(RfqDetailId)
)/*最終報價單價格紀錄 參考用表*/

END/*報價單管理*/



CREATE TABLE PDM.DesignForManufacturing( /*DFM單頭*/
 DfmId						int					identity,
 DfmNo						nvarchar(20)		NOT NULL,						--DFM單號
 RfqDetailId				int					NOT NULL,						--RFQ單身
 DfmDate					datetime			NOT NULL,						--Dfm日期
 DifficultyLevel			nvarchar(2)				NULL,						--難易等級 
 RdUserId					int						NULL,						--研發人員
 [Version]					nvarchar(10)		NOT NULL 	DEFAULT '000',		--版本號
 DefaultSpreadsheetData		nvarchar(MAX)			NULL,						--預設單身Spreadsheet資料
 SpreadsheetData			nvarchar(MAX)			NULL,						--單身Spreadsheet資料
 DfmProcessStatus			nvarchar(2)			NOT NULL,						--1.DFM評估項目中 2.項目成本資料維護中 3.項目成本計算完畢 9.DFM異動變更
 ProcessStatus2Date			datetime				NULL,						--DFM評估項目維護完成時間
 ProcessStatus3Date			datetime				NULL,						--DFM成本計算完成時間
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼(Y/N:已確認/未確認)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DesignForManufacturing PRIMARY KEY (DfmId),
 CONSTRAINT FK1_DesignForManufacturing FOREIGN KEY (RfqDetailId) REFERENCES SCM.RfqDetail(RfqDetailId)
) /*DFM單頭*/

CREATE TABLE PDM.DfmDetail( /*DFM 單身*/
 DfmDetailId				int					identity,
 DfmId						int					NOT NULL,						--DFM單頭Id
 DfmItemId					int					NOT NULL,						--項目名稱 (DfmItem)
 [Standard]					nvarchar(200)			NULL,						--DFM項目標準
 FinalSpec					nvarchar(200)			NULL,						--最後判定規格
 CustomerSpec				nvarchar(200)			NULL,						--客戶規格
 SuggestSpec				nvarchar(200)			NULL,						--RD建議值規格
 RdRemark					nvarchar(200)			NULL,						--研發評論
 CustomerFeedback			nvarchar(200)			NULL,						--客戶回應
 [Version]					nvarchar(10)		NOT NULL 	DEFAULT '000',		--版本號
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DfmDetail PRIMARY KEY (DfmDetailId),
 CONSTRAINT FK1_DfmDetail FOREIGN KEY (DfmId) REFERENCES PDM.DesignForManufacturing(DfmId),
 CONSTRAINT FK2_DfmDetail FOREIGN KEY (DfmItemId) REFERENCES PDM.DfmItem(DfmItemId)
) /*DFM 單身*/

CREATE TABLE PDM.DfmDetailLog( /*DFM 單身Log*/
 DfmDetailLogId				int					identity,
 DfmId						int					NOT NULL,						--Dfm單頭
 DfmItemId					int					NOT NULL,						--項目名稱 (DfmItem)
 Standard					nvarchar(200)			NULL,						--DFM項目標準
 FinalSpec					nvarchar(200)			NULL,						--最後判定規格
 CustomerSpec				nvarchar(200)			NULL,						--客戶規格
 SuggestSpec				nvarchar(200)			NULL,						--RD建議值規格
 RdRemark					nvarchar(200)			NULL,						--研發評論
 CustomerFeedback			nvarchar(200)			NULL,						--客戶回應
 [Version]					nvarchar(10)		NOT NULL 	DEFAULT '000',		--版本號
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DfmDetailLog PRIMARY KEY (DfmDetailLogId),
 CONSTRAINT FK1_DfmDetailLog FOREIGN KEY (DfmId) REFERENCES PDM.DesignForManufacturing(DfmId),
 CONSTRAINT FK2_DfmDetailLog FOREIGN KEY (DfmItemId) REFERENCES PDM.DfmItem(DfmItemId)
) /*DFM 單身*/

CREATE TABLE PDM.DfmQuotationItem( /*DFM 所需報價項目*/
 DfmQiId					int					identity,
 DfmId						int					NOT NULL,						--DFMId
 DfmQuotationName			nvarchar(20)		NOT NULL,						--報價項目名稱
 DfmItemCategoryId			int					NOT NULL,						--項目種類
 QuotationLevel				int					NOT NULL	DEFAULT 1,			--階級(層)
 ParentDfmQiId				int					NOT NULL	DEFAULT -1,			--上階報價項目ID
 MfgFlag					nvarchar(1)			NOT NULL,						--是否製造件(Y/N) 製造件, 標準件
 StandardCost				float				NOT NULL	DEFAULT 0,			--標準金額
 TotalManHours				float				NOT NULL    DEFAULT 0,			--總人力工時 單PCS金額
 TotalMachineHours			float				NOT NULL	DEFAULT 0,			--總人力工時 單PCS金額
 TotalOspAmount				float				NOT NULL	DEFAULT 0,			--總預估外包費用 單PCS金額
 TotalMtlAmount				float				NOT NULL	DEFAULT	0,			--總材料費用(模架/模仁粗胚) 單PCS金額
 MaterialStatus				nvarchar(2)			NOT NULL,						--物料維護狀態碼(Y/A/S:已確認/需考慮/不考慮)
 DfmQiOSPStatus				nvarchar(2)			NOT NULL,						--托外維護狀態碼(Y/A/S:已確認/需考慮/不考慮)
 DfmQiProcessStatus			nvarchar(2)			NOT NULL,						--製程維護狀態碼(Y/A/S:已確認/需考慮/不考慮)
 DefaultMaterialSpreadsheetData		nvarchar(MAX)			NULL,				--預設MaterialSpreadsheet資料
 MaterialSpreadsheetData			nvarchar(MAX)			NULL,				--MaterialSpreadsheet資料
 DefaultQiProcessSpreadsheetData	nvarchar(MAX)			NULL,				--預設QiProcessSpreadsheet資料
 QiProcessSpreadsheetData			nvarchar(MAX)			NULL,				--QiProcessSpreadsheet資料
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼(A/S:啟用/停用)
 [Version]					nvarchar(10)		NOT NULL 	DEFAULT '000',		--版本號
 MaterialConfirmDate		datetime				NULL,						--物料確認完成時間
 QiProcessConfirmDate		datetime				NULL,						--製程確認完成時間
 DfmQiOSPConfirmDate		datetime				NULL,						--托外確認完成時間
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL, 						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DfmQuotationItem PRIMARY KEY (DfmQiId),
 CONSTRAINT FK1_DfmQuotationItem FOREIGN KEY (DfmId) REFERENCES PDM.DesignForManufacturing(DfmId),
 CONSTRAINT FK2_DfmQuotationItem FOREIGN KEY (DfmItemCategoryId) REFERENCES PDM.DfmItemCategory(DfmItemCategoryId)
) /*DFM 所需報價項目*/

CREATE TABLE PDM.DfmQuotationItemLog( /*DFM 所需報價項目Log*/
 DfmQiLogId					int					identity,
 DfmQiId					int					NOT NULL,						--所需報價項目Id
 DfmId						int					NOT NULL,						--DFM單頭Id
 DfmItemCategoryId			int					NOT NULL,						--生產模式ID (只顯示 DfmFlag=Y)
 MfgFlag					nvarchar(1)			NOT NULL,						--是否製造件(Y/N) 製造件, 標準件
 ParentDfmQiId				int					NOT NULL	DEFAULT -1,			--上階報價項目ID
 StandardCost				float				NOT NULL	DEFAULT 0,			--標準金額
 TotalManHours				float				NOT NULL	DEFAULT 0,			--總人力工時 單PCS金額
 TotalMachineHours			float				NOT NULL	DEFAULT 0,			--總人力工時 單PCS金額
 TotalOspAmount				float				NOT NULL	DEFAULT 0,			--總預估外包費用 單PCS金額
 TotalMtlAmount				float				NOT NULL	DEFAULT	0,			--總材料費用(模架/模仁粗胚) 單PCS金額
 MaterialStatus				nvarchar(2)			NOT NULL,						--狀態碼(Y/N:已確認/未確認)
 DfmQiOSPStatus				nvarchar(2)			NOT NULL,						--狀態碼(Y/N:已確認/未確認)
 DfmQiProcessCostStatus		nvarchar(2)			NOT NULL,						--狀態碼(Y/N:已確認/未確認)
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼(Y/N:已確認/未確認)
 [Version]					nvarchar(10)		NOT NULL 	DEFAULT '000',		--版本號
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL, 						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DfmQuotationItemLog PRIMARY KEY (DfmQiLogId),
 CONSTRAINT FK1_DfmQuotationItemLog FOREIGN KEY (DfmQiId) REFERENCES PDM.DfmQuotationItem(DfmQiId),
 CONSTRAINT FK2_DfmQuotationItemLog FOREIGN KEY (DfmId) REFERENCES PDM.DesignForManufacturing(DfmId),
 CONSTRAINT FK3_DfmQuotationItemLog FOREIGN KEY (DfmItemCategoryId) REFERENCES PDM.DfmItemCategory(DfmItemCategoryId)
) /*DFM 所需報價項目*/

CREATE TABLE PDM.DfmQiMaterial( /*DFM 使用物料*/
 DfmQiMaterialId			int					identity,
 DfmQiId					int					NOT NULL,						--所需報價項目Id
 DfmMaterialId				int					NOT NULL,					--物料ID
 MaterialNo					nvarchar(20)		NOT NULL,						--物料代碼
 SolutionQtyFlag			int					NOT NULL	DEFAULT 1,			--報價方式(1.考慮方案數量, 2.不考慮方案數量)
 MaterialName				nvarchar(20)		NOT NULL,						--物料名稱 ex:1.模架, 2.粗胚 (不需要製造的標準件也在這選), 委外加工也算在這
 Quantity					float				NOT NULL,						--數量
 UnitPrice					float				NOT NULL,						--單價
 Amount						float				NOT NULL,						--材料金額
 RefSolutionQty				nvarchar(1)			NOT NULL	DEFAULT 'N',		--連動報價方案數量
 [Version]					nvarchar(10)		NOT NULL 	DEFAULT '000',		--版本號
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DfmQiMaterial PRIMARY KEY (DfmQiMaterialId),
 CONSTRAINT FK1_DfmQiMaterial FOREIGN KEY (DfmQiId) REFERENCES PDM.DfmQuotationItem(DfmQiId),
 CONSTRAINT FK2_DfmQiMaterial FOREIGN KEY (DfmMaterialId) REFERENCES PDM.DfmMaterial(DfmMaterialId)
) /*DFM 使用物料*/

CREATE TABLE PDM.DfmQiMaterialLog( /*DFM 使用物料Log*/
 DfmQiMaterialLogId			int					identity,
 DfmQiMaterialId			int					NOT NULL,						--使用物料Id
 DfmQiId					int					NOT NULL,						--所需報價項目Id
 MaterialNo					nvarchar(20)		NOT NULL,						--物料代碼
 MaterialName				nvarchar(20)		NOT NULL,						--物料名稱 ex:1.模架, 2.粗胚 (不需要製造的標準件也在這選), 委外加工也算在這
 Quantity					float				NOT NULL,						--數量
 UnitPrice					float				NOT NULL,						--單價
 Amount						float				NOT NULL,						--材料金額
 RefSolutionQty				nvarchar(1)			NOT NULL	DEFAULT 'N',		--連動報價方案數量
 [Version]					nvarchar(10)		NOT NULL 	DEFAULT '000',		--版本號
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DfmQiMaterialLog PRIMARY KEY (DfmQiMaterialLogId),
 CONSTRAINT FK1_DfmQiMaterialLog FOREIGN KEY (DfmQiId) REFERENCES PDM.DfmQuotationItem(DfmQiId)
) /*DFM 使用物料Log*/

CREATE TABLE PDM.DfmQiOSP( /*DFM 使用委外加工*/
 DfmQiOSPId					int					identity,
 DfmQiId					int					NOT NULL,						--所需報價項目Id
 OspName					nvarchar(32)		NOT NULL,						--委外名稱
 Amount						float				NOT NULL,						--委外金額
 [Version]					nvarchar(10)		NOT NULL 	DEFAULT '000',		--版本號
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DfmQiOSP PRIMARY KEY (DfmQiOSPId),
 CONSTRAINT FK1_DfmQiOSP FOREIGN KEY (DfmQiId) REFERENCES PDM.DfmQuotationItem(DfmQiId)
) /*DFM 使用委外加工*/

CREATE TABLE PDM.DfmQiOSPLog( /*DFM 使用委外加工Log*/
 DfmQiOSPLogId				int					identity,
 DfmQiOSPId					int					NOT NULL,						--使用委外加工Id
 DfmQiId					int					NOT NULL,						--所需報價項目Id
 OspName					nvarchar(32)		NOT NULL,						--委外名稱
 Amount						float				NOT NULL,						--委外金額
 [Version]					nvarchar(10)		NOT NULL 	DEFAULT '000',		--版本號
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DfmQiOSPLog PRIMARY KEY (DfmQiOSPLogId),
 CONSTRAINT FK1_DfmQiOSPLog FOREIGN KEY (DfmQiId) REFERENCES PDM.DfmQuotationItem(DfmQiId)
) /*DFM 使用委外加工Log*/

CREATE TABLE PDM.DfmQiProcess( /*DFM 製程成本*/
 DfmQiProcessId				int					identity,
 DfmQiId					int					NOT NULL,						--所需報價項目Id
 SortNumber					int						NULL,						--製程順序
 ParameterId				int					NOT NULL,						--製程ID
 AllocationType				int					NOT NULL,						--工時分配類別(1.單PCS工時, 2.By製程批量, 3.By報價數量)
 AllocationQty				int					NOT NULL	DEFAULT 0,			--分配基數數量
 ManHours					float				NOT NULL,						--工時
 ManHoursUnit				nvarchar(5)			NOT NULL,						--工時單位(s,m,hr,day,month,year)
 MachineHours				float				NOT NULL,						--機時
 MachineHoursUnit			nvarchar(5)			NOT NULL,						--機時單位(s,m,hr,day,month,year)
 YieldRate					int					NOT NULL	DEFAULT 100,		--良率Default 100(最高100)
 [Version]					nvarchar(10)		NOT NULL 	DEFAULT '000',		--版本號
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DfmQiProcess PRIMARY KEY (DfmQiProcessId),
 CONSTRAINT FK1_DfmQiProcess FOREIGN KEY (ParameterId) REFERENCES MES.ProcessParameter(ParameterId)
) /*DFM 製程*/

CREATE TABLE PDM.DfmQiProcessLog( /*DFM 製程成本Log*/
 DfmQiProcessCostLogId		int					identity,
 DfmQiProcessId				int					NOT NULL,						--DFM製程ID
 DfmQiId					int					NOT NULL,						--所需報價項目Id
 SortNumber					int					NOT NULL,						--製程順序
 ParameterId				int					NOT NULL,						--製程ID
 AllocationType				int						NULL,						--工時分配類別(1.單PCS工時, 2.By製程批量, 3.By製造批量)
 AllocationQty				int						NULL	DEFAULT 0,			--分配基數數量
 ManHours					float				NOT NULL,						--工時
 MachineHours				float				NOT NULL,						--機時
 YieldRate					int					NOT NULL	DEFAULT 100,		--良率Default 100(最高100)
 [Version]					nvarchar(10)		NOT NULL 	DEFAULT '000',		--版本號
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DfmQiProcessCostLog PRIMARY KEY (DfmQiProcessCostLogId),
 CONSTRAINT FK1_DfmQiProcessCostLog FOREIGN KEY (DfmQiProcessId) REFERENCES PDM.DfmQiProcess(DfmQiProcessId),
 CONSTRAINT FK2_DfmQiProcessCostLog FOREIGN KEY (ParameterId) REFERENCES MES.ProcessParameter(ParameterId)
) /*DFM 製程Log*/

CREATE TABLE PDM.DfmQiSolution( /*DFM 製程成本By方案數量計算*/
 DfmQiSolutionId			int					NOT NULL	identity,
 DfmQiId					int					NOT NULL,
 RfqLineSolutionId			int					NOT NULL,
 MaterialAmt				float				NOT NULL	DEFAULT 0,			--物料費用
 ResourceAmt				float				NOT NULL	DEFAULT 0,			--人工費用
 OverheadAmt				float				NOT NULL	DEFAULT 0,			--製造費用
 StandardCost				float			    NOT NULL    DEFAULT 0,			--標準成本(費用), 由下層Summary回
 DiscountAmount				float				NOT NULL	DEFAULT 0,			--折扣金額
 GrossProfitMargin			float				NOT NULL	DEFAULT 0,			--毛利率, 最高100
 AfterProfitMargin			float				NOT NULL	DEFAULT 0,			--毛利率, 最高100
 QuotationAmount			float				NOT NULL	DEFAULT 0,			--報價金額
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DfmQiSolution PRIMARY KEY (DfmQiSolutionId),
 CONSTRAINT FK1_DfmQiSolution FOREIGN KEY (DfmQiId) REFERENCES PDM.DfmQuotationItem(DfmQiId),
 CONSTRAINT FK2_DfmQiSolution FOREIGN KEY (RfqLineSolutionId) REFERENCES SCM.RfqLineSolution(RfqLineSolutionId)
) /*DFM 製程*/

CREATE TABLE PDM.DfmQiSolutionProcess( /*DFM 製程成本By方案數量計算*/
 DfmQiProcessId				int					NOT NULL,
 DfmQiSolutionId			int					NOT NULL,
 ResourceAmt				float				NOT NULL	DEFAULT 0,			--人工費用
 OverheadAmt				float				NOT NULL	DEFAULT 0,			--製造費用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT FK1_DfmQiSolutionProcess FOREIGN KEY (DfmQiProcessId) REFERENCES PDM.DfmQiProcess(DfmQiProcessId),
 CONSTRAINT FK2_DfmQiSolutionProcess FOREIGN KEY (DfmQiSolutionId) REFERENCES PDM.DfmQiSolution(DfmQiSolutionId)
) /*DFM 製程*/

CREATE TABLE PDM.DfmTemplate( /*DFM 表單樣板--依照RFQ*/
 DfmTemplateId				int					identity,						--流水碼
 CompanyId					int					NOT NULL,						--公司別
 DfmTempNo			        nvarchar(100)		NOT NULL,                       --樣板編號
 DfmTempName			    nvarchar(100)		NOT NULL,                       --樣板名稱 
 RfqProTypeId				int					NOT NULL,						--RFQ產品小類別
 ProductUseId				int					NOT NULL,						--產品用途
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼		>> S:關閉 ; A:啟用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DfmTemplate  PRIMARY KEY (DfmTemplateId),
 CONSTRAINT AK1_DfmTemplate UNIQUE (DfmTempNo,RfqProTypeId,ProductUseId),
 CONSTRAINT FK1_DfmTemplate FOREIGN KEY (RfqProTypeId) REFERENCES SCM.RfqProductType(RfqProTypeId),
 CONSTRAINT FK2_DfmTemplate FOREIGN KEY (ProductUseId) REFERENCES SCM.ProductUse(ProductUseId),
 CONSTRAINT FK3_DfmTemplate FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)/*DFM 表單樣板--依照生產模式*/

CREATE TABLE PDM.DfmTemplateParameter( /*DFM 表單樣板參數--依照RFQ*/
 DfmTempParameterId			int					identity,						--流水碼
 DfmTemplateId				int					NOT NULL,
 DfmItemId					int					NOT NULL, 
 DfmTempParamType			nvarchar(10)		NOT NULL,						--參數輸入方式 >> BAS.[Type] File：上傳檔案；Combobox：下拉；Text：輸入文字；Radio：單選按鈕；Checkbox：多選按鈕 
 DfmTempParamValue			nvarchar(10)		NOT NULL,						--參數預設值
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼		>> S:關閉 ; A:啟用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DfmTemplateParameter  PRIMARY KEY (DfmTempParameterId), 
 CONSTRAINT FK1_DfmTemplateParameter FOREIGN KEY (DfmTemplateId) REFERENCES PDM.DfmTemplate(DfmTemplateId),
 CONSTRAINT FK2_DfmTemplateParameter FOREIGN KEY (DfmItemId) REFERENCES PDM.DfmItem(DfmItemId)
)/*DFM 表單樣板--依照生產模式*/

END/*電商平台*/