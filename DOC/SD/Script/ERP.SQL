BEGIN /*基本資料*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DocumentProfile') > 0 DROP TABLE ERP.DocumentProfile
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'LimitedUser') > 0 DROP TABLE ERP.LimitedUser
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'CurrencyExchangeRate') > 0 DROP TABLE ERP.CurrencyExchangeRate
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'CerDetail') > 0 DROP TABLE ERP.CerDetail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'PositionsCategory') > 0 DROP TABLE ERP.PositionsCategory
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'PositionsPersonnel') > 0 DROP TABLE ERP.PositionsPersonnel
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'PaymentTermData') > 0 DROP TABLE ERP.PaymentTermData
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'TaxCodeProfile') > 0 DROP TABLE ERP.TaxCodeProfile
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'TradeTermData') > 0 DROP TABLE ERP.TradeTermData
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ShipMethodData') > 0 DROP TABLE ERP.ShipMethodData
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'TradingObjectCategory') > 0 DROP TABLE ERP.TradingObjectCategory
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'FinancialInstitutionData') > 0 DROP TABLE ERP.FinancialInstitutionData
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'InvoiceCountData') > 0 DROP TABLE ERP.InvoiceCountData
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'CommonParameterProfile') > 0 DROP TABLE ERP.CommonParameterProfile
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Inventory') > 0 DROP TABLE ERP.Inventory
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'InvDetail') > 0 DROP TABLE ERP.InvDetail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'FactoryData') > 0 DROP TABLE ERP.FactoryData
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Customer') > 0 DROP TABLE ERP.Customer
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Supplier') > 0 DROP TABLE ERP.Supplier
END /*基本資料*/

BEGIN /*訂單資料*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Quotation') > 0 DROP TABLE ERP.Quotation
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'QuotDetail') > 0 DROP TABLE ERP.QuotDetail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'QuotDetail') > 0 DROP TABLE ERP.QuotDetailFile
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'SalesForecast') > 0 DROP TABLE ERP.SalesForecast
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'SfDetail') > 0 DROP TABLE ERP.SfDetail
END /*訂單資料*/

BEGIN /*品號*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'CustomerMtlItemData') > 0 DROP TABLE ERP.CustomerMtlItemData
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'CustomerMtlPrice') > 0 DROP TABLE ERP.CustomerMtlPrice
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'CmpDetail') > 0 DROP TABLE ERP.CmpDetail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MtlItemCategoryData') > 0 DROP TABLE ERP.MtlItemCategoryData
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MtlItem') > 0 DROP TABLE ERP.MtlItem --補FK及相關table
END /*品號*/

--******************************************************************************************************************
BEGIN /*基本資料*/
CREATE TABLE ERP.DocumentProfile( /*各系統單據設定檔*/ /*ERP CMSMQ*/
 DpId						int					identity,
 CompanyId					int					NOT NULL,						--所屬公司ID
 DpErpPrefix				nvarchar(4)			NOT NULL,						--MQ001 單別
 DpErpName					nvarchar(40)		NOT NULL,						--MQ002 單據名稱
 DpType						nvarchar(2)			NOT NULL,						--MQ003	單據性質 >> 1?:庫存；2?:訂單；3?:採購；4?:BOM；5?:製令；6?:應收；7?:應付；8?:票據；9?:會計；A?:進口；B?:出口；C?:固定資產；D?:製程；E?:維修；F?:模具；G?：保稅；H?:軍福品；I?:零用金；J?:人力資源；K?:海關合同；L?:IDL；M?:會計多帳本；N?:設備保養
 EncodingScheme				nvarchar(1)			NOT NULL,						--MQ004	編碼方式 >> 1:日編；2:月編；3:流水號；4:手動編號
 YearCodeNumber				nvarchar(1)			NOT NULL,						--MQ005	年碼數
 SeqCodeNumber				nvarchar(2)			NOT NULL,						--MQ006	流水號碼數
 CustomerName				nvarchar(80)		NOT NULL,						--MQ034 單據全名
 AutoCommit					nvarchar(1)			NOT NULL,						--MQ015 自動確認 >> BAS.[Status] N:未勾選；Y:已勾選
 AutomaticPrinting			nvarchar(1)			NOT NULL,						--MQ016 自動列印 >> BAS.[Status] N:未勾選；Y:已勾選
 UpdatePricing				nvarchar(1)			NOT NULL,						--MQ018 更新核價 >> BAS.[Status] N:未勾選；Y:已勾選
 MtlInputType				nvarchar(1)			NOT NULL,						--MQ024	品號輸入方式 >> 1:採品號輸入；2:採條碼輸入
 AnnotateNumber				nvarchar(6)			NOT NULL,						--MQ025	註記代號 >> 參照 ERP註記/簽核資料建立作業(CMSMS)
 PrintModifyAnnotate		nvarchar(1)			NOT NULL,						--MQ026 列印時修改註記 >> BAS.[Status] N:未勾選；Y:已勾選
 SignOffNumber				nvarchar(6)			NOT NULL,						--MQ027	簽核代號 >> 參照 ERP註記/簽核資料建立作業(CMSMS)
 PrintModifySignOff			nvarchar(1)			NOT NULL,						--MQ028 列印時修改簽核 >> BAS.[Status] N:未勾選；Y:已勾選
 LimitedUser				nvarchar(1)			NOT NULL,						--MQ029 單別限定輸入使用者 >> 參照 ERP.LimitedUser BAS.[Status] N:未勾選；Y:已勾選
 PrintAnnotate				nvarchar(1)			NOT NULL,						--MQ030 每頁列印註記 >> BAS.[Status] N(DEF)；Y
 PrintSignOff				nvarchar(1)			NOT NULL,						--MQ031 每頁列印簽核 >> BAS.[Status] N(DEF)；Y
 CtrlMtlExpiryDate			nvarchar(1)			NOT NULL,						--MQ040 使用未生效己失效品號 >> N；Y(DEF)
 CertificateFormat			nvarchar(15)		NOT NULL,						--MQ032	憑證格式
 ChooseCertificateFormat	nvarchar(1)			NOT NULL,						--MQ033 列印時選擇憑證格式 >> BAS.[Status] N(DEF)；Y
 PrintTotal					nvarchar(1)			NOT NULL,						--MQ035 每頁列印合計 >> BAS.[Status] N(DEF)；Y
 ExchangeRateSource			nvarchar(1)			NOT NULL,						--MQ044	匯率來源 >> I:銀行買進匯率(DEF:"I")；O:銀行賣出匯率；E:報關買進匯率；W:報關賣出匯率
 ScreenRetain				nvarchar(1)			NOT NULL,						--MQ069 新增狀態畫面保留 >> BAS.[Status] N(DEF)；Y
 ModifyPaymentTerm			nvarchar(1)			NOT NULL,						--MQ078 修改付款條件 >> N；Y(DEF)
 Reamrk						nvarchar(255)			NULL,						--MQ022 備註
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DocumentProfile PRIMARY KEY (DpId),
 CONSTRAINT AK1_DocumentProfile UNIQUE (CompanyId, DpErpPrefix),
 CONSTRAINT FK1_DocumentProfile FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_DocumentProfile FOREIGN KEY (CreateBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK3_DocumentProfile FOREIGN KEY (LastModifiedBy) REFERENCES BAS.[User](UserId)
) /*各系統單據設定檔*/

CREATE TABLE ERP.LimitedUser( /*單別限定輸入使用者*/ /*ERP CMSMU*/
 LuId						int					identity,
 DpId						int					NOT NULL,						--MU001 所屬單別
 LuType						nvarchar(2)			NOT NULL,						--MU002	單據性質(依各系統單據性質設定之規格設定/傳票種類) >> 1?:庫存；2?:訂單；3?:採購；4?:BOM；5?:製令；6?:應收；7?:應付；8?:票據；9?:會計；A?:進口；B?:出口；C?:固定資產；D?:製程；E?:維修；F?:模具；G?：保稅；H?:軍福品；I?:零用金；J?:人力資源；K?:海關合同；L?:IDL；M?:會計多帳本；N?:設備保養
 UserId						int					NOT NULL,						--MU003 使用者代號
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_LimitedUser PRIMARY KEY (LuId),
 CONSTRAINT FK1_LimitedUser FOREIGN KEY (CreateBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK2_LimitedUser FOREIGN KEY (LastModifiedBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK3_LimitedUser FOREIGN KEY (DpId) REFERENCES ERP.DocumentProfile(DpId),
 CONSTRAINT FK4_LimitedUser FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId)
) /*單別限定輸入使用者*/

CREATE TABLE ERP.CurrencyExchangeRate( /*幣別匯率檔單頭*/ /*ERP CMSMF*/
 CerId						int					identity,
 CompanyId					int					NOT NULL,						--所屬公司ID
 Currency					nvarchar(4) 		NOT NULL,						--MF001 幣別
 CurrencyName				nvarchar(40)		NOT NULL,						--MF002	幣別名稱
 UnitPriceFetchBit			nvarchar(1)			NOT NULL,						--MF003 單價取位 >> 0.整數；1.小數一位；2.小數二位；3.小數三位；4.小數四位；5.小數五位；6.小數六位
 AmountFetchBit				nvarchar(1)			NOT NULL,						--MF004 金額取位 >> 0.整數；1.小數一位；2.小數二位
 UnitCostFetchBit			nvarchar(1)			NOT NULL,						--MF005 單位成本取位 >> 0.整數；1.小數一位；2.小數二位；3.小數三位；4.小數四位；5.小數五位；6.小數六位
 CostAmountFetchBit			nvarchar(1)			NOT NULL,						--MF006 成本金額取位 >> 0.整數；1.小數一位；2.小數二位
 Reamrk						nvarchar(255)			NULL,						--MF007 備註
 EnPriceCapitalCurrency		nvarchar(60)			NULL,						--MF012 英文金額大寫輸出幣別
 FixedExchangeRate			float				NOT NULL,						--MF013	固定匯率
 PriceCapitalFormat			nvarchar(1)			NOT NULL,						--MF014 金額大寫格式 >> 1.越盾；2.美元
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_CurrencyExchangeRate PRIMARY KEY (CerId),
 CONSTRAINT AK1_CurrencyExchangeRate UNIQUE (CompanyId, Currency),
 CONSTRAINT FK1_CurrencyExchangeRate FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_CurrencyExchangeRate FOREIGN KEY (CreateBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK3_CurrencyExchangeRate FOREIGN KEY (LastModifiedBy) REFERENCES BAS.[User](UserId)
) /*幣別匯率檔單頭*/

CREATE TABLE ERP.CerDetail( /*幣別匯率檔單身*/ /*ERP CMSMG*/
 CerDetailId				int					identity,
 CerId						int					NOT NULL,						--所屬單頭
 EffectiveDate				datetime			NOT	NULL,						--MG002 生效日期 [FORMATE:YMD]
 BankBuyingExchangeRate		float				NOT NULL,						--MG003	銀行買進匯率
 BankSellingExchangeRate	float				NOT NULL,						--MG004	銀行賣出匯率
 CclBuyingExchangeRate		float				NOT NULL,						--MG005	報關買進匯率
 CclSellingExchangeRate		float				NOT NULL,						--MG006	報關賣出匯率
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_CerDetail PRIMARY KEY (CerDetailId),
 CONSTRAINT FK1_CerDetail FOREIGN KEY (CreateBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK2_CerDetail FOREIGN KEY (LastModifiedBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK3_CerDetailId FOREIGN KEY (CerId) REFERENCES ERP.CurrencyExchangeRate(CerId)
) /*幣別匯率檔單身*/

CREATE TABLE ERP.PositionsCategory( /*職務類別檔*/ /*ERP CMSMJ*/
 PcId						int					identity,
 CompanyId					int					NOT NULL,						--所屬公司ID
 PcNo						nvarchar(6)			NOT NULL,						--MJ001	職務代號
 PcType						nvarchar(1)			NOT NULL,						--MJ002 職務分類 >> 1.物管、2.生管、3.業務、4.採購、5.會計、6.出納、7.倉管、8.研發、9.製造、A.品管、B.管理、C.工程、D.生技、E.船務、F.廠務、G.貿易、H.總務、I.人事、J.保稅、K.稽核、L.企劃、M.文管、N.產品、O.行政、P.外點(專櫃前抬)、Z.其它
 PcName						nvarchar(40)		NOT NULL,						--MJ003	職務名稱
 Reamrk						nvarchar(255)			NULL,						--MJ004 備註
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_PositionsCategory PRIMARY KEY (PcId),
 CONSTRAINT AK1_PositionsCategory UNIQUE (CompanyId, PcNo),
 CONSTRAINT FK1_PositionsCategory FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_PositionsCategory FOREIGN KEY (CreateBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK3_PositionsCategory FOREIGN KEY (LastModifiedBy) REFERENCES BAS.[User](UserId)
) /*職務類別檔*/

CREATE TABLE ERP.PositionsPersonnel( /*職務人員檔*/ /*ERP CMSMK*/
 PpId						int					identity,
 PcId						int					NOT NULL,						--類別職務ID
 UserId						int					NOT NULL,						--MK002	人員代號
 Reamrk						nvarchar(255)			NULL,						--MK005 備註
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_PositionsPersonnel PRIMARY KEY (PpId),
 CONSTRAINT FK1_PositionsPersonnel FOREIGN KEY (CreateBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK2_PositionsPersonnel FOREIGN KEY (LastModifiedBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK3_PositionsPersonnel FOREIGN KEY (PcId) REFERENCES ERP.PositionsCategory(PcId),
 CONSTRAINT FK4_PositionsPersonne2 FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId)
) /*職務人員檔*/

CREATE TABLE ERP.PaymentTermData( /*付款條件檔*/ /*ERP CMSNA*/
 PtdId						int					identity,
 CompanyId					int					NOT NULL,						--所屬公司ID
 PtdNo						nvarchar(6)			NOT NULL,						--NA002	付款條件代號
 PtdName					nvarchar(40)		NOT NULL,						--NA003	名稱
 ReceiptPaymentMode			nvarchar(1)			NOT NULL,						--NA004 收(付)款方式別 >> 1.加日數；2.加月數
 AfterClosingDate			nvarchar(3)			NOT NULL,						--NA005 結帳後__天
 AfterClosingMonth			nvarchar(1)				NULL,						--NA006 結帳後__月
 AfterClosingDays			nvarchar(2)				NULL,						--NA007 結帳後__日
 ReceiptPaymentDateType		nvarchar(1)			NOT NULL,						--NA018 收款日起算日 >> 1.結帳日 2.次月初 (DEF:1)
 CapitalRealizeMode			nvarchar(1)			NOT NULL,						--NA008 資金實現方式別 >> 1.加日數；2.加月數
 AfterPaymentDate			nvarchar(3)			NOT NULL,						--NA009 付款後__天
 AfterPaymentMonth			nvarchar(1)				NULL,						--NA010 付款後__月
 AfterPaymentDays			nvarchar(2)				NULL,						--NA011 付款後__日
 RealizeDateType			nvarchar(1)			NOT NULL,						--NA019 實現日起算日 >> 1.付款日 2.次月初 (DEF:1)
 GetDiscountRate			nvarchar(1)			NOT NULL,						--NA012 取得折扣 >> BAS.[Status] N:未勾選；Y:已勾選
 GetDiscountRateType		nvarchar(1)			NOT NULL,						--NA013 取得折扣方式 >> 1.付款提前 2.票期縮短
 BeforehandPaymentDays		nvarchar(3)			NOT NULL,						--NA014 提早付款天數
 BillTermBeforehandDays		nvarchar(3)			NOT NULL,						--NA015 票期提前天數
 DiscountRate				float				NOT NULL,						--NA016	折扣
 Reamrk						nvarchar(255)			NULL,						--NA017 備註
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_PaymentTermData PRIMARY KEY (PtdId),
 CONSTRAINT AK1_PaymentTermData UNIQUE (CompanyId, PtdNo),
 CONSTRAINT FK1_PaymentTermData FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_PaymentTermData FOREIGN KEY (CreateBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK3_PaymentTermData FOREIGN KEY (LastModifiedBy) REFERENCES BAS.[User](UserId)
) /*付款條件檔*/

CREATE TABLE ERP.TaxCodeProfile( /*稅別碼設定檔*/ /*ERP CMSNN*/
 TcpId						int					identity,
 CompanyId					int					NOT NULL,						--所屬公司ID
 TcpNo						nvarchar(3)			NOT NULL,						--NN001	稅別碼代號
 TcpName					nvarchar(30)		NOT NULL,						--NN002	稅別名稱
 PurchaseSalesItem			nvarchar(1)			NOT NULL,						--NN003 進/銷項 >> 1.進項；2.銷項
 InvoiceCount				int					NOT NULL,						--NN005 發票聯數 >> 參照 ERP.InvoiceCountData
 Taxation					nvarchar(1)			NOT NULL,						--NN006 課稅別 >> 1.應稅內含、2.應稅外加、3.零稅率、4.免稅、9.不計稅
 BusinessTaxRate			float				NOT NULL,						--NN004 營業稅率
 Reamrk						nvarchar(255)			NULL,						--NN009 備註
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_TaxCodeProfile PRIMARY KEY (TcpId),
 CONSTRAINT AK1_TaxCodeProfile UNIQUE (CompanyId, TcpNo),
 CONSTRAINT FK1_TaxCodeProfile FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_TaxCodeProfile FOREIGN KEY (CreateBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK3_TaxCodeProfile FOREIGN KEY (LastModifiedBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK4_TaxCodeProfile FOREIGN KEY (InvoiceCount) REFERENCES ERP.InvoiceCountData(IcdId)
) /*稅別碼設定檔*/

CREATE TABLE ERP.TradeTermData( /*交易條件代號檔*/ /*ERP CMSNK*/
 TtdId						int					identity,
 CompanyId					int					NOT NULL,						--所屬公司ID
 TtdNo						nvarchar(1)			NOT NULL,						--NK001	交易條件代號
 TtdName					nvarchar(40)		NOT NULL,						--NK002	交易條件名稱
 Reamrk						nvarchar(255)			NULL,						--NK004 備註
 MultiplyFormula			nvarchar(60)		NOT	NULL,						--NK010 計算公式乘法
 DivisionFormula			nvarchar(60)		NOT	NULL,						--NK011 計算公式除法
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_TradeTermData PRIMARY KEY (TtdId),
 CONSTRAINT AK1_TradeTermData UNIQUE (CompanyId, TtdNo),
 CONSTRAINT FK1_TradeTermData FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_TradeTermData FOREIGN KEY (CreateBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK3_TradeTermData FOREIGN KEY (LastModifiedBy) REFERENCES BAS.[User](UserId)
) /*交易條件代號檔*/

CREATE TABLE ERP.ShipMethodData( /*運輸方式代號檔*/ /*ERP CMSNJ*/
 SmdId						int					identity,
 CompanyId					int					NOT NULL,						--所屬公司ID
 SmdNo						nvarchar(1)			NOT NULL,						--NJ001	運輸方式代號
 SmdCnName					nvarchar(40)		NOT NULL,						--NJ002	運輸方式中文名稱
 SmdEnName					nvarchar(80)		NOT NULL,						--NJ003	運輸方式英文名稱
 Reamrk						nvarchar(255)			NULL,						--NJ004 備註
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ShipMethodData PRIMARY KEY (SmdId),
 CONSTRAINT AK1_ShipMethodData UNIQUE (CompanyId, SmdNo),
 CONSTRAINT FK1_ShipMethodData FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_ShipMethodData FOREIGN KEY (CreateBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK3_ShipMethodData FOREIGN KEY (LastModifiedBy) REFERENCES BAS.[User](UserId)
) /*運輸方式代號檔*/

CREATE TABLE ERP.TradingObjectCategory( /*交易對象分類方式檔*/ /*ERP CMSMR*/
 TocId						int					identity,
 CompanyId					int					NOT NULL,						--所屬公司ID
 TocType					nvarchar(1)			NOT NULL,						--MR001	分類方式 >> 1:通路、2:型態、3:地區、4:國家、5:路線、6:其他、7:抽成、8:活動、9:廠商分類
 TocNo						nvarchar(6)			NOT NULL,						--MR002	代號
 TocShortName				nvarchar(30)		NOT NULL,						--MR003	簡稱
 TocName					nvarchar(80)		NOT NULL,						--MR004	全名
 Reamrk						nvarchar(255)			NULL,						--MR005 備註
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_TradingObjectCategory PRIMARY KEY (TocId),
 CONSTRAINT AK1_TradingObjectCategory UNIQUE (CompanyId, TocNo),
 CONSTRAINT FK1_TradingObjectCategory FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_TradingObjectCategory FOREIGN KEY (CreateBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK3_TradingObjectCategory FOREIGN KEY (LastModifiedBy) REFERENCES BAS.[User](UserId)
) /*交易對象分類方式檔*/

CREATE TABLE ERP.FinancialInstitutionData( /*金融機構資料檔*/ /*ERP CMSMO*/
 FidId						int					identity,
 CompanyId					int					NOT NULL,						--所屬公司ID
 FidNo						nvarchar(20)		NOT NULL,						--MO001	金融機構代號
 FidType					nvarchar(1)			NOT NULL,						--MO002	金融機構種類 >> 1:本國銀行、2:外國銀行、3:信託投資、4:票券金融、5:信用合作社、6:產物保險、7:漁會信用、8:農會信用、9:郵局
 FidHeadOffice				nvarchar(50)		    NULL,						--MO003	金融機構總行
 FidBranchOffice			nvarchar(30)			NULL,						--MR005	金融機構分行
 FidName					nvarchar(80)		NOT NULL,						--MO006	金融機構名稱
 PostingDate				float				NOT NULL,						--MO007 入帳日數 (DEF:0)
 FidAddress					nvarchar(255)			NULL,						--MO008 地址
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_FinancialInstitutionData PRIMARY KEY (FidId),
 CONSTRAINT AK1_FinancialInstitutionData UNIQUE (CompanyId, FidNo),
 CONSTRAINT FK1_FinancialInstitutionData FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_FinancialInstitutionData FOREIGN KEY (CreateBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK3_FinancialInstitutionData FOREIGN KEY (LastModifiedBy) REFERENCES BAS.[User](UserId)
) /*金融機構資料檔*/

CREATE TABLE ERP.InvoiceCountData( /*發票聯數檔*/ /*ERP CMSNM*/
 IcdId						int					identity,
 CompanyId					int					NOT NULL,						--所屬公司ID
 IcdNo						nvarchar(1)			NOT NULL,						--NM001	發票聯數
 IcdName					nvarchar(30)		NOT NULL,						--NM002	聯數名稱
 IcdNumberLength			nvarchar(2)			NOT	NULL,						--NM003 發票號碼長度
 Reamrk						nvarchar(255)			NULL,						--NM004 備註
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_InvoiceCountData PRIMARY KEY (IcdId),
 CONSTRAINT AK1_InvoiceCountData UNIQUE (CompanyId, IcdNo),
 CONSTRAINT FK1_InvoiceCountData FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_InvoiceCountData FOREIGN KEY (CreateBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK3_InvoiceCountData FOREIGN KEY (LastModifiedBy) REFERENCES BAS.[User](UserId)
) /*發票聯數檔*/

CREATE TABLE ERP.CommonParameterProfile( /*共用參數設定檔*/ /*ERP CMSMA*/
 CppId						int					identity,
 CompanyId					int					NOT NULL,						--所屬公司ID
 DateFormatType				nvarchar(1)			NOT NULL,						--MA001	日期格式 >> 1.西元年 YYYYMMDD、2.西元年 MMDDYYYY、3.西元年 DDMMYYYY、4.民國年 YYMMDD、5.民國年 YYYMMDD
 DateSegmentSymbol			nvarchar(1)			NOT NULL,						--MA002 日期區隔符號 >> 1.-、2./
 DomesticCurrency			int			 		NOT NULL,						--MA003 本國幣別 >> 參照 ERP.CurrencyExchangeRate 
 BusinessTaxRate			float				NOT NULL,						--MA004 營業稅率 >> 由稅別碼決定(NN004)
 QtyExpressionMode			nvarchar(1)			NOT NULL,						--MA024 數量表達方式 >> 1.單一單位；2.雙單位；3.製造雙單位
 AccordingToConfirmDate		nvarchar(1)			NOT NULL,						--MA025 確認日依據 >> 1.系統日、2.單據日 (DEF:'1')
 QuotTax					nvarchar(1)			NOT NULL,						--TA023 稅額 >> 1.整張資料計算、2.單身單筆資料計算
 MtlItemOneToManyBarcode	nvarchar(1)			NOT NULL,						--MA023 一品號對應多條碼 >> BAS.[Status] N；Y
 TenPercentile				nvarchar(1)			NOT NULL,						--MA031 數值欄位三位一撇 >> BAS.[Status] N(DEF)；Y
 ElectronicForm				nvarchar(1)			NOT NULL,						--MA032 執行電子表單流程 >> BAS.[Status] N(DEF)；Y
 MtlItemDataControlMode		nvarchar(1)			NOT NULL,						--MA037 品號資料管制方式 >> 1.品名不可重複；2.品名重複時僅需警告；N.不控管(DEF:'N')
 CustomerDataControlMode	nvarchar(1)			NOT NULL,						--MA038 客戶資料管制方式 >> 1.客戶簡稱不可重複；2.客戶簡稱重複時僅需警告；N.不控管(DEF:'N')
 VendorDataControlMode		nvarchar(1)			NOT NULL,						--MA039 廠商資料管制方式 >> 1.廠商簡稱不可重複；2.廠商簡稱重複時僅需警告；N.不控管(DEF:'N')
 RemarkDefault				nvarchar(1)			NOT NULL,						--MA044 備註預設 >> 1.公用片語；2.個人片語；N.不預設(DEF:'N')
 InvoicePrintMtlItemSeq		nvarchar(1)			NOT NULL,						--MA102 電子發票列印品號來源順序 >> 1.品號->業務品號；2.業務品號->品號(DEF:'1')
 MrpNetChange				nvarchar(1)			NOT NULL,						--MA005 MRP使用NetChange(MRP淨變法) >> BAS.[Status] N；Y
 GoodsSorting1				nvarchar(10)		NOT NULL,						--MA007 商品分類方式一
 GoodsSorting2				nvarchar(10)		NOT NULL,						--MA008 商品分類方式二
 GoodsSorting3				nvarchar(10)		NOT NULL,						--MA009 商品分類方式三
 GoodsSorting4				nvarchar(10)		NOT NULL,						--MA010 商品分類方式四
 ShippingMarkPrint			nvarchar(1)			NOT NULL,						--MA035 嘜頭列印依據 >> 1.資料庫欄位；2.嘜頭文管檔案(DEF:"1")
 StoringLocationManagement	nvarchar(1)			NOT NULL,						--MA040 儲位管理 >> BAS.[Status] N(DEF)；Y
 SpecifyStoringLocation		nvarchar(1)			NOT NULL,						--MA041 出庫單據存檔時需指定儲位 >> BAS.[Status] N；Y(DEF)
 SpecifyLotNumber			nvarchar(1)			NOT NULL,						--MA042 出庫單據存檔時需指定批號 >> BAS.[Status] N；Y(DEF)
 CostPricingMethod			nvarchar(1)			NOT NULL,						--MA014 成本計價方式 >> 1.標準成本制、2.月加權平均成本制、3.移動平均成本制
 InvInForceDate				nvarchar(6)			NOT NULL,						--MA011 庫存現行年月 >> [FORMATE:YM]
 InvFinalizeLedgerDate		nvarchar(6)			NOT NULL,						--MA012 庫存關帳年月 >> [FORMATE:YM]
 AccountFrozenDate			nvarchar(8)			NOT NULL,						--MA013 帳務凍結日期 >> [FORMATE:YMD]
 AccountingPeriod			nvarchar(1)			NOT NULL,						--MA016 會計期制(會計期間) >> 1.12期、2.13期
 BankDepositInForceDate		nvarchar(6)			NOT NULL,						--MA021 銀行存款現行年月 >> [FORMATE:YM]
 ArInForceDate				nvarchar(6)			NOT NULL,						--MA027 應收現行年月 >> [FORMATE:YM]
 ArFinalizeLedgerDate		nvarchar(6)			NOT NULL,						--MA028 應收關帳年月 >> [FORMATE:YM]
 ApInForceDate				nvarchar(6)			NOT NULL,						--MA029 應付現行年月 >> [FORMATE:YM]
 ApFinalizeLedgerDate		nvarchar(6)			NOT NULL,						--MA030 應付關帳年月 >> [FORMATE:YM]
 PettyCashInForceDate		nvarchar(6)			NOT NULL,						--MA033 零用金現行年月 >> [FORMATE:YM]
 ExciseTaxInForceDate		nvarchar(6)			NOT NULL,						--MA047 營業稅現行年月 >> [FORMATE:YM]
 BudgetManagement			nvarchar(1)			NOT NULL,						--MA045 預算管理 >> BAS.[Status] N(DEF)；Y
 BudgetItemSorting			nvarchar(1)			NOT NULL,						--MA046 預算科目分類方式 >> (DEF:'1')
 CalculatePayrollDate		nvarchar(6)			NOT NULL,						--MA022 計薪年月 >> [FORMATE:YM]
 UpdateLowLevelCodes		nvarchar(6)			NOT NULL,						--MA017 更新低階碼
 ProductFunctionCategory	nvarchar(1)			NOT NULL,						--MA026 產品功能別 >> 1.台灣版、2.大陸版 (DEF:'1')
 ProductVersion				nvarchar(15)		NOT NULL,	DEFAULT 'GP4.0',	--MA034 產品版號
 FixedAssetsInForceDate		nvarchar(6)			NOT NULL,						--MA036 固定資產現行年月 >> [FORMATE:YM]
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_CommonParameterProfile PRIMARY KEY (CppId),
 CONSTRAINT FK1_CommonParameterProfile FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_CommonParameterProfile FOREIGN KEY (CreateBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK3_CommonParameterProfile FOREIGN KEY (LastModifiedBy) REFERENCES BAS.[User](UserId)
) /*共用參數設定檔*/

CREATE TABLE ERP.Inventory( /*庫別資料檔*/ /*ERP CMSMC*/
 InventoryId				int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 InventoryNo				nvarchar(10)		NOT NULL,						--MC001 庫別代號
 InventoryName				nvarchar(40)		NOT NULL,						--MC002 庫別名稱
 InventoryType				nvarchar(2)			NOT NULL,						--MC004 庫別性質 >> BAS.[Type] 1.存貨倉、2.非存貨倉
 MrpCalculation				nvarchar(2)			NOT NULL,						--MC005 納入MRP計算 >> BAS.[Status]
 ConfirmStatus				nvarchar(2)			NOT NULL,						--MC006 確認時庫存量不足准許出庫 >> BAS.[Status] Y；N
 SaveStatus					nvarchar(2)			NOT NULL,						--MC008 存檔時庫存量不足准許出庫 >> BAS.[Status] Y；N
 InventoryDesc				nvarchar(100)			NULL,						--MC007 庫別描述
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Inventory PRIMARY KEY (InventoryId),
 CONSTRAINT AK1_Inventory UNIQUE (CompanyId, InventoryNo),
 CONSTRAINT AK2_Inventory UNIQUE (CompanyId, InventoryName),
 CONSTRAINT FK1_Inventory FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_Inventory FOREIGN KEY (CreateBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK3_Inventory FOREIGN KEY (LastModifiedBy) REFERENCES BAS.[User](UserId)
) /*庫別資料檔*/

CREATE TABLE ERP.InvDetail( /*庫別儲位檔*/ /*ERP CMSNL*/
 InvDetailId				int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 StoringLocationNo			nvarchar(10)		NOT NULL,						--NL002 儲位編號
 StoringLocationDesc		nvarchar(255)		NOT NULL,						--NL003 儲位說明
 Remark						nvarchar(255)			NULL,						--NL004 備註
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_InvDetail PRIMARY KEY (InvDetailId),
 CONSTRAINT AK1_InvDetail UNIQUE (CompanyId, StoringLocationNo),
 CONSTRAINT FK1_InvDetail FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_InvDetail FOREIGN KEY (CreateBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK3_InvDetail FOREIGN KEY (LastModifiedBy) REFERENCES BAS.[User](UserId)
) /*庫別儲位檔*/

CREATE TABLE ERP.FactoryData( /*廠別資料檔*/ /*ERP CMSMB*/
 FdId						int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 FdNo						nvarchar(6)			NOT NULL,						--MB001 廠別代號
 FdName						nvarchar(40)		NOT NULL,						--MB002 廠別名稱
 TelNo						nvarchar(20)		NOT NULL,						--MB003 電話
 FaxNo						nvarchar(20)			NULL,						--MB004 傳真
 RegisterAddressFirst		nvarchar(255)		NOT NULL,						--MB005 地址(一)
 RegisterAddressFirst		nvarchar(255)			NULL,						--MB006 地址(二)
 Remark						nvarchar(255)			NULL,						--MB007 備註
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_FactoryData PRIMARY KEY (FdId),
 CONSTRAINT AK1_FactoryData UNIQUE (CompanyId, FdNo),
 CONSTRAINT FK1_FactoryData FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_FactoryData FOREIGN KEY (CreateBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK3_FactoryData FOREIGN KEY (LastModifiedBy) REFERENCES BAS.[User](UserId)
) /*廠別資料檔*/

CREATE TABLE ERP.Customer( /*客戶基本資料檔*/ /*ERP COPMA*/
 CustomerId					int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 DepartmentId				int					    NULL,						--MA015 所屬部門
 CustomerNo					nvarchar(50)		NOT NULL,						--MA001 客戶代號
 CustomerName				nvarchar(200)		NOT NULL,						--MA003 客戶全名
 CustomerEnglishName		nvarchar(80)			NULL,						--MA110 客戶英文全名
 CustomerShortName			nvarchar(100)		NOT NULL,						--MA002 客戶簡稱
 Country					nvarchar(6)				NULL,						--MA019 國家別 >> 參照 ERP.TradingObjectCategory
 Region						nvarchar(6)				NULL,						--MA018 地區別 >> 參照 ERP.TradingObjectCategory
 [Route]					nvarchar(6)				NULL,						--MA077 路線別 >> 參照 ERP.TradingObjectCategory
 FirstTradeDate				datetime				NULL,						--MA021 初次交易日
 NearestTradeDate			datetime				NULL,						--MA022 最近交易日
 GetPriceSeq				nvarchar(10)		NOT NULL,						--MA075 取價順序 >> 1客戶計價,2標準售價,3零售價,4售價定價一,5售價定價二,6售價定價三,7售價定價四,8售價定價五,9售價定價六,A折扣後客戶計價,B折扣後標準售價,C折扣後零售價,D折扣後售定價一,E折扣後售定價二,F折扣後售定價三,G折扣後售定價四,H折扣後售定價五,I折扣後售定價六
 GuiNumber					nvarchar(20)		NOT NULL,						--MA010 統一編號
 ResponsiblePerson			nvarchar(30)		NOT NULL,						--MA004 客戶負責人
 Contact					nvarchar(30)			NULL,						--MA005 客戶聯絡人
 TelNoFirst					nvarchar(20)		NOT NULL,						--MA006 客戶聯絡電話(一)
 TelNoSecond				nvarchar(20)			NULL,						--MA007 客戶聯絡電話(二)
 FaxNo						nvarchar(20)			NULL,						--MA008 客戶聯絡傳真
 Email						nvarchar(60)			NULL,						--MA009 客戶電子郵件
 RegisterAddressFirst		nvarchar(255)		NOT NULL,						--MA023 登記地址(一)
 RegisterAddressSecond		nvarchar(255)			NULL,						--MA024 登記地址(二)
 InvoiceAddressFirst		nvarchar(255)		NOT NULL,						--MA025 發票地址(一)
 InvoiceAddressSecond		nvarchar(255)			NULL,						--MA026 發票地址(二)
 DeliveryAddressFirst		nvarchar(255)		NOT NULL,						--MA027 送貨地址(一)
 DeliveryAddressSecond		nvarchar(255)			NULL,						--MA064 送貨地址(二)
 DocumentAddressFirst		nvarchar(255)			NULL,						--MA062 文件地址(一)
 DocumentAddressSecond		nvarchar(255)			NULL,						--MA063 文件地址(二)
 BillAddressFirst			nvarchar(255)			NULL,						--MA099 帳單地址(一)
 BillAddressSecond			nvarchar(255)			NULL,						--MA100 帳單地址(二)
 ZipCodeRegister			nvarchar(6)				NULL,						--MA040 郵遞區號(登記)
 ZipCodeInvoice				nvarchar(6)				NULL,						--MA079 郵遞區號(發票)
 ZipCodeDelivery			nvarchar(6)				NULL,						--MA080 郵遞區號(送貨)
 ZipCodeDocument			nvarchar(6)				NULL,						--MA081 郵遞區號(文件)
 ZipCodeBill				nvarchar(6)				NULL,						--MA098 郵遞區號(帳單)
 BillReceipient				nvarchar(30)			NULL,						--MA101 帳單收件人
 InauguateDate				datetime				NULL,						--MA020 開業日
 AccountDay					nvarchar(2)				NULL,						--MA043 結帳日
 CloseDate					datetime				NULL,						--MA068 歇業日
 PermitDate					datetime				NULL,						--MA111 新客戶核准日
 [Version]					nvarchar(4)			NOT NULL,						--MA112 版次
 Capital					float				NOT NULL,						--MA011 資本額
 Headcount					int					NOT NULL,						--MA013 員工人數
 HomeOffice					nvarchar(10)			NULL,						--MA065 總店號
 AnnualTurnover				float				NOT NULL,						--MA012 年營業額
 Currency					int					NOT NULL,						--MA014 交易幣別 >> 參照 ERP.CurrencyExchangeRate
 TradeTerm					int					NOT NULL,						--MA109 所屬交易條件 >> 參照 ERP.TradeTermData
 PriceTerm					nvarchar(40)			NULL,						--MA030 價格條件
 PaymentTerm				int					NOT NULL,						--MA031 付款條件 >> 參照 ERP.PaymentTermData
 ValuationType				nvarchar(1)			NOT NULL,						--MA035 計價方式 >> BAS.[Type] 1：利潤率；2：折扣率
 PaymentType				nvarchar(1)			NOT NULL,						--MA041 收款方式 >> BAS.[Type] 1：現金；2：電匯；3：支票；4：其他
 ReceiptReceive				nvarchar(1)			NOT NULL,						--MA042 票據寄領 >> BAS.[Type] 1：郵寄；2：自領；3：其它
 TaxAmountCalculateType		nvarchar(1)			NOT NULL,						--MA087 稅額計算方式 >> BAS.[Type] 1：整張資料計算；2：單身單筆資料計算
 DocumentDeliver			nvarchar(1)			NOT NULL,						--MA039 單據發送方式 >> BAS.[Type] 1：郵寄；2：FAX；3：EDI；4：E-MAIL
 IcdId						int					NOT NULL,						--MA037 發票聯數 >> 參照 ERP.InvoiceCountData
 TaxNo						int					NOT NULL,						--MA118 稅別碼 >> 參照 ERP.TaxCodeProfile
 Taxation					nvarchar(1)			NOT NULL,						--MA038 課稅別 >> 由稅別碼決定(NN006) 1：內含；2：外加；3：零稅率；4：免稅；9：不計稅
 RevenueDelayDays			nvarchar(3)				NULL,						--MA127 收入遞延天數
 PaymentBankFirst			nvarchar(20)			NULL,						--MA046 付款銀行(一) >> 參照 ERP.FinancialInstitutionData
 PaymentBankSecond			nvarchar(20)			NULL,						--MA069 付款銀行(二) >> 參照 ERP.FinancialInstitutionData
 PaymentBankThird			nvarchar(20)			NULL,						--MA070 付款銀行(三) >> 參照 ERP.FinancialInstitutionData
 BankAccountFirst			nvarchar(30)			NULL,						--MA071 銀行帳號(一)
 BankAccountSecond			nvarchar(30)			NULL,						--MA072 銀行帳號(二)
 BankAccountThird			nvarchar(30)	 		NULL,						--MA073 銀行帳號(三)
 Account					nvarchar(20)			NULL,						--MA047 帳款科目 >> 參照 ERP會計科目資料檔(ACTMA)
 AccountInvoice				nvarchar(20)			NULL,						--MA074 票據科目 >> 參照 ERP會計科目資料檔(ACTMA)
 Clearance					nvarchar(1)			NOT NULL,						--MA104 通關方式 >> BAS.[Type] 1：非經海關(DEF)；2：經海關
 ShipMethod					int					NOT NULL,						--MA048 所屬運輸方式 >> 參照 ERP.ShipMethodData
 Destination				nvarchar(40)			NULL,						--MA051 目的地
 OceanShippingPort			nvarchar(40)			NULL,						--MA052 海運港口
 AirfreightAerodrome		nvarchar(40)			NULL,						--MA053 空運機場
 InvocieAttached			nvarchar(1)			NOT NULL,						--MA086 隨貨附發票 >> BAS.[Status] Y/N
 CustomerKind				nvarchar(1)				NULL,						--MA137 客戶分類 >> BAS.[Type] 1：自然人；2：法人
 UploadType					nvarchar(1)				NULL,						--MA141 上傳類型 >> BAS.[Type] 1：B2C存證；2：B2B存證(DEF)；3：B2B交換
 DepositRate				float				NOT NULL,						--MA095 訂金比率
 SaleRating					nvarchar(1)				NULL,						--MA028 銷售評等 >> BAS.[Type] 1：A；2：B；3：C
 CreditRating				nvarchar(1)				NULL,						--MA029 信用評等 >> BAS.[Type] 1：A；2：B；3：C
 CreditLimit				float					NULL,						--MA033 信用額度
 CreditLimitCtrlToCompany	nvarchar(1)			NOT	NULL,						--MA082 信用額度依總公司控管 >> BAS.[Type] N(DEF)；Y
 CreditLimitControl			nvarchar(1)				NULL,						--MA032 信用額度管制 >> BAS.[Type] Y：依公司參數管制；N：信用額度不管制；y：依客戶資料管制
 CreditLimitControlCurrency nvarchar(1)			NOT NULL,						--MA133 信用額度控管幣別 >> BAS.[Type] 1：記帳幣別(DEF)；2：主要交易幣別
 SoCreditAudit				nvarchar(1)			NOT NULL,						--MA088 訂單信用查核方式 >> BAS.[Type] 1.不檢查(DEF) 2.警告 3.拒絕 
 SiCreditAudit				nvarchar(1)			NOT NULL,						--MA089 銷貨信用查核方式 >> BAS.[Type] 1.不檢查(DEF) 2.警告 3.拒絕
 SnCreditAudit				nvarchar(1)			NOT NULL,						--MA120 出貨通知信用查核方式 >> BAS.[Type] 1.不檢查(DEF) 2.警告 3.拒絕
 TsnCreditAudit				nvarchar(1)			NOT NULL,						--MA132 暫出單信用查核方式 >> BAS.[Type] 1.不檢查(DEF) 2.警告 3.拒絕
 OutstandingBillRate		float				NOT NULL,						--MA091 未兌現應收票據比率 (DEF:0)
 AccountReceivableRate		float				NOT NULL,						--MA092 應收帳款比率 (DEF:0)
 UnsettledAccountSiRate		float				NOT NULL,						--MA093 未結帳銷貨金額比率 (DEF:0)
 UnsettledShipmentSoRate	float				NOT NULL,						--MA094 未出貨訂單金額比率 (DEF:0)
 UnsettledReturnTsnRate		float				NOT NULL,						--MA102 未歸還暫出金額比率 (DEF:0)
 RelatedPerson				nvarchar(20)			NULL,						--MA124 關係人代號 >> 參照 ERP子母公司資料檔(FCSMB)
 SalesmenId					int						NULL,						--MA016 業務員
 PaymentSalesmenId			int						NULL,						--MA085 收款業務員
 ShipType					nvarchar(50)			NULL,						--運輸種類 >> BAS.[Type] 1：出貨；2：三角貿易
 ForwarderId				int						NULL,						--貨運承攬商
 CustomerRemark				nvarchar(500)			NULL,						--MA049 客戶備註
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Customer PRIMARY KEY (CustomerId),
 CONSTRAINT AK1_Customer UNIQUE (CompanyId, CustomerNo),
 CONSTRAINT FK1_Customer FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_Customer FOREIGN KEY (DepartmentId) REFERENCES BAS.Department(DepartmentId),
 CONSTRAINT FK3_Customer FOREIGN KEY (CreateBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK4_Customer FOREIGN KEY (LastModifiedBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK5_Customer FOREIGN KEY (SalesmenId) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK6_Customer FOREIGN KEY (PaymentSalesmenId) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK7_Customer FOREIGN KEY (ForwarderId) REFERENCES SCM.Forwarder(ForwarderId),
 CONSTRAINT FK8_Customer FOREIGN KEY (Country) REFERENCES ERP.TradingObjectCategory(TocId),
 CONSTRAINT FK9_Customer FOREIGN KEY (Region) REFERENCES ERP.TradingObjectCategory(TocId),
 CONSTRAINT FK10_Customer FOREIGN KEY ([Route]) REFERENCES ERP.TradingObjectCategory(TocId),
 CONSTRAINT FK11_Customer FOREIGN KEY (PaymentBankFirst) REFERENCES ERP.PaymentBankFirst(FidId),
 CONSTRAINT FK12_Customer FOREIGN KEY (PaymentBankSecond) REFERENCES ERP.PaymentBankFirst(FidId),
 CONSTRAINT FK13_Customer FOREIGN KEY (PaymentBankThird) REFERENCES ERP.PaymentBankFirst(FidId),
 CONSTRAINT FK14_Customer FOREIGN KEY (Currency) REFERENCES ERP.CurrencyExchangeRate(CerId),
 CONSTRAINT FK15_Customer FOREIGN KEY (TradeTerm) REFERENCES ERP.TradeTermData(TtdId),
 CONSTRAINT FK16_Customer FOREIGN KEY (PaymentTerm) REFERENCES ERP.PaymentTermData(PtdId),
 CONSTRAINT FK17_Customer FOREIGN KEY (IcdId) REFERENCES ERP.InvoiceCountData(IcdId),
 CONSTRAINT FK18_Customer FOREIGN KEY (TaxNo) REFERENCES ERP.TaxCodeProfile(TcpId),
 CONSTRAINT FK19_Customer FOREIGN KEY (ShipMethod) REFERENCES ERP.ShipMethodData(SmdId)
)/*客戶基本資料檔*/

CREATE TABLE ERP.Supplier( /*廠商基本資料檔*/ /*ERP PURMA*/
 SupplierId					int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 SupplierNo					nvarchar(50)		NOT NULL,						--MA001 供應商代號
 SupplierName				nvarchar(200)		NOT NULL,						--MA003 供應商全名
 SupplierShortName			nvarchar(100)		NOT NULL,						--MA002 供應商簡稱
 SupplierEnglishName		nvarchar(100)		NOT NULL,						--MA088 供應商英文全名
 Country					nvarchar(6)				NULL,						--MA006 國家別 >> 參照 ERP.TradingObjectCategory
 Region						nvarchar(6)				NULL,						--MA007 地區別 >> 參照 ERP.TradingObjectCategory
 GuiNumber					nvarchar(20)		NOT NULL,						--MA005 統一編號
 ResponsiblePerson			nvarchar(30)		NOT NULL,						--MA012 供應商負責人
 ContactFirst				nvarchar(30)			NULL,						--MA013 供應商聯絡人(一)
 ContactSecond				nvarchar(30)			NULL,						--MA048 供應商聯絡人(二)
 ContactThird				nvarchar(30)			NULL,						--MA049 供應商聯絡人(三)
 TelNoFirst					nvarchar(20)		NOT NULL,						--MA008 供應商聯絡電話(一)
 TelNoSecond				nvarchar(20)			NULL,						--MA009 供應商聯絡電話(二)
 FaxNo						nvarchar(20)			NULL,						--MA010 供應商聯絡傳真
 FaxNoAccounting			nvarchar(20)			NULL,						--MA059 供應商聯絡傳真(會計)
 Email						nvarchar(60)			NULL,						--MA011 供應商電子郵件
 AddressFirst				nvarchar(255)		NOT NULL,						--MA014 供應商聯絡地址(一)
 AddressSecond				nvarchar(255)			NULL,						--MA015 供應商聯絡地址(二)
 ZipCodeFirst				nvarchar(6)				NULL,						--MA046 郵遞區號(一)
 ZipCodeSecond				nvarchar(6)				NULL,						--MA050 郵遞區號(二)
 BillAddressFirst			nvarchar(255)			NULL,						--MA051 帳單地址(一)
 BillAddressSecond			nvarchar(255)			NULL,						--MA052 帳單地址(二)
 PermitStatus				nvarchar(2)			NOT NULL,						--MA016 核准狀態碼 >> BAS.[Status] 1：已核准；2：尚待核准；3：不准交易
 SupplierPermitDate			datetime				NULL,						--MA070 廠商核准日 >> [FORMATE:YMD]
 InauguateDate				datetime				NULL,						--MA017 開業日
 AccountMonth				nvarchar(1)				NULL,						--MA034 結帳月
 AccountDay					nvarchar(2)			NOT NULL,						--MA035 結帳日
 [Version]					nvarchar(4)			NOT NULL,						--MA069 版次
 Capital					float				NOT NULL,						--MA018 資本額
 Headcount					int					NOT NULL,						--MA019 員工人數
 PoDeliver					nvarchar(1)			NOT NULL,						--MA020 採購單發送方式 >> BAS.[Type] 1：郵寄；2：FAX；3：EDI；4：E-MAIL
 FirstTradeDate				datetime				NULL,						--MA022 初次交易日
 NearestTradeDate			datetime				NULL,						--MA023 最近交易日
 Currency					int			 		NOT NULL,						--MA021 交易幣別 >> 參照 ERP.CurrencyExchangeRate
 TradeTerm					int					NOT NULL,						--MA077 交易條件 >> 參照 ERP.TradeTermData
 PaymentType				nvarchar(1)			NOT NULL,						--MA024 付款方式 >> BAS.[Type] 1：現金；2：電匯；3：支票；4：其他
 PaymentTermNo				int			 		NOT NULL,						--MA055 付款條件代碼 >> 參照 ERP.PaymentTermData
 PaymentTermName			nvarchar(16) 		NOT NULL,						--MA025 付款條件
 ReceiptReceive				nvarchar(1)			NOT NULL,						--MA029 票據寄領 >> BAS.[Type] 1：郵寄；2：自領；3：其它
 InvoiceCount				int					NOT NULL,						--MA030 發票聯數 >> 參照 ERP.InvoiceCountData
 TaxNo						int					NOT NULL,						--MA076 稅別碼 >> 參照 ERP.TaxCodeProfile
 Taxation					nvarchar(1)			NOT NULL,						--MA044 課稅別 >> 由稅別碼決定(NN006) 1：內含；2：外加；3：零稅率；4：免稅；9：不計稅
 PermitPartialDelivery		nvarchar(1)			NOT NULL,						--MA045 允許分批交貨 >> BAS.[Status]
 TaxAmountCalculateType		nvarchar(1)			NOT NULL,						--MA053 稅額計算方式 >> BAS.[Type] 1：整張資料計算；2：單身單筆資料計算
 InvocieAttachedStatus		nvarchar(1)			NOT NULL,						--MA056 隨貨附發票 >> BAS.[Status]
 CertificateFormatType		nvarchar(1)			NOT NULL,						--MA057 憑證列印格式 >> BAS.[Type]
 DepositRate				float				NOT NULL,						--MA058 訂金比率
 TradeItem					nvarchar(255)			NULL,						--MA066 交易項目
 RemitBank					nvarchar(20)		NOT NULL,						--MA027 匯款銀行
 RemitAccount				nvarchar(30)		NOT NULL,						--MA028 匯款帳號
 AccountPayable				nvarchar(20)			NULL,						--MA041 應付帳款科目 >> 參照 ERP會計科目資料檔(ACTMA)
 AccountOverhead			nvarchar(20)			NULL,						--MA042 加工費用科目 >> 參照 ERP會計科目資料檔(ACTMA)
 AccountInvoice				nvarchar(20)			NULL,						--MA043 票據科目 >> 參照 ERP會計科目資料檔(ACTMA)
 SupplierLevel				nvarchar(1)				NULL,						--MA031 ABC等級 >> BAS.[Type] 1：A；2：B；3：C
 DeliveryRating				nvarchar(1)				NULL,						--MA032 交貨評等 >> BAS.[Type] 1：A；2：B；3：C；4:D
 QualityRating				nvarchar(1)				NULL,						--MA033 品質評等 >> BAS.[Type] 1：A；2：B；3：C
 RelatedPerson				nvarchar(20)			NULL,						--MA085 關係人代號 >> 參照 ERP子母公司資料檔(FCSMB)
 PurchaseUserId				int						NULL,						--MA047 採購人員
 SupplierRemark				nvarchar(255)			NULL,						--MA040 供應商備註
 TransferStatus				nvarchar(2)			NOT NULL,						--拋轉ERP狀態 >> BAS.[Status]
 TransferDate				datetime				NULL,						--拋轉ERP時間
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Supplier PRIMARY KEY (SupplierId),
 CONSTRAINT AK1_Supplier UNIQUE (CompanyId, SupplierNo),
 --CONSTRAINT AK2_Supplier UNIQUE (SupplierName, PermitStatus),
 CONSTRAINT FK1_Supplier FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_Supplier FOREIGN KEY (PurchaseUserId) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK3_Supplier FOREIGN KEY (PaymentTermNo) REFERENCES ERP.PaymentTermData(PtdId),
 CONSTRAINT FK4_Supplier FOREIGN KEY (Currency) REFERENCES ERP.CurrencyExchangeRate(CerId),
 CONSTRAINT FK5_Supplier FOREIGN KEY (TradeTerm) REFERENCES ERP.TradeTermData(TtdId),
 CONSTRAINT FK6_Supplier FOREIGN KEY (InvoiceCount) REFERENCES ERP.InvoiceCountData(IcdId),
 CONSTRAINT FK7_Supplier FOREIGN KEY (TaxNo) REFERENCES ERP.TaxCodeProfile(TcpId)
)/*廠商基本資料檔*/

END /*基本資料*/

--BEGIN /*系統設定*/
--CREATE TABLE ERP.ErpCompany( /*多公司資料建立作業*/ /*ERP DSCMB*/
 --ErpCompanyId				int					identity,
 --CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 --LastModifiedDate			datetime				NULL,						--修改日期
 --CreateBy					int					NOT NULL,						--新增人員
 --LastModifiedBy				int						NULL,						--修改人員
 --CONSTRAINT PK_ErpCompany PRIMARY KEY (ErpCompanyId)
--)
--END /*系統設定*/

BEGIN /*訂單*/
CREATE TABLE ERP.Quotation( /*報價單頭*/ /*ERP COPTA*/
 QuotId						int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 QuotErpPrefix				nvarchar(4)			NOT NULL,						--TA001 報價單單別 >> 參照 ERP.DocumentProfile
 QuotErpNo					nvarchar(11)		NOT NULL,						--TA002 報價單單號
 QuotDate					datetime			NOT NULL,						--TA003 報價日期
 DocDate					datetime			NOT NULL,						--TA013 單據日期
 [Version]					nvarchar(4)			NOT NULL,						--TA064 版次
 CustomerConfrim			nvarchar(1)			NOT NULL,						--TA016 客戶確認 >> BAS.[Status] N:未勾選；Y:已勾選
 CustomerId					int					NOT NULL,						--TA004 客戶代號
 CustomerName				nvarchar(80)		    NULL,						--TA006 客戶全名
 SalesmenId					int					    NULL,						--TA005 業務人員
 Currency					int			 		NOT NULL,						--TA007 幣別 >> 參照 ERP.CurrencyExchangeRate
 ExchangeRate				float				NOT NULL,						--TA008 匯率 >> 參照 ERP.CerDetail
 DetailMultiTax				nvarchar(1)			NOT NULL,						--TA063 單身多稅率 >> BAS.[Status] N(DEF)；Y
 TaxNo						int					NOT NULL,						--TA049 稅別碼 >> 參照 ERP.TaxCodeProfile
 Taxation					nvarchar(1)			NOT NULL,						--TA022 課稅別 >> 由稅別碼決定(NN006) 1：內含；2：外加；3：零稅率；4：免稅；9：不計稅
 BusinessTaxRate			float				NOT NULL,						--TA024 營業稅率 >> 由稅別碼決定(NN004)
 PriceTerm					nvarchar(40) 			NULL,						--TA010 價格條件
 PaymentTermNo				int			 		NOT NULL,						--TA026 付款條件代碼 >> 參照 ERP.PaymentTermData
 PaymentTermName			nvarchar(16) 		NOT NULL,						--TA011 付款條件
 CertificateFormatType		nvarchar(1)			NOT NULL,						--TA017 憑證列印格式 >> BAS.[Type] 1：中式；2：英式
 QuotAmount					float				NOT NULL,						--TA009 報價金額
 QuotTax					float				NOT NULL,						--TA023 稅額
 TotalQty					float				NOT NULL,						--TA025 總數量
 ValuationType				nvarchar(1)			NOT NULL,						--TA012 計價方式 >> BAS.[Type] 1：利潤率；2：折扣率
 DeliveryDate				float				NOT NULL,						--TA014 交貨日
 ContactPerson				nvarchar(30)		NOT	NULL,						--TA033 連絡人
 Reamrk1					nvarchar(255)			NULL,						--TA020 備註一
 Reamrk2					nvarchar(255)			NULL,						--TA021 備註二
 CuftUnit					nvarchar(1)			NOT NULL,						--TA035 材積單位 1：英制；2：公制 1(DEF)；2
 TradeTerm					int					NOT NULL,						--TA039 所屬交易條件 >> 參照 ERP.TradeTermData
 TradeTerm1					int					NOT NULL,						--TA058 交易條件1 >> 參照 ERP.TradeTermData
 TradeTerm2					int					NOT NULL,						--TA059 交易條件2 >> 參照 ERP.TradeTermData
 ShipMethod					int					NOT NULL,						--TA045 所屬運輸方式 >> 參照 ERP.ShipMethodData
 ConfirmStatus				nvarchar(1)			NOT NULL,						--TA019 確認碼　>> BAS.[Status] N：未確認；V：作廢；Y：已確認
 ConfirmUserId				int						NULL,						--TA015 確認者
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Quotation PRIMARY KEY (QuotId),
 CONSTRAINT AK1_Quotation UNIQUE (CompanyId, QuotErpPrefix, QuotErpNo),
 CONSTRAINT FK1_Quotation FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_Quotation FOREIGN KEY (CreateBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK3_Quotation FOREIGN KEY (LastModifiedBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK4_Quotation FOREIGN KEY (CustomerId) REFERENCES ERP.Customer(CustomerId),
 CONSTRAINT FK5_Quotation FOREIGN KEY (SalesmenId) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK6_Quotation FOREIGN KEY (QuotErpPrefix) REFERENCES ERP.DocumentProfile(DpId),
 CONSTRAINT FK7_Quotation FOREIGN KEY (Currency) REFERENCES ERP.CurrencyExchangeRate(CerId),
 CONSTRAINT FK8_Quotation FOREIGN KEY (TaxNo) REFERENCES ERP.TaxCodeProfile(TcpId),
 CONSTRAINT FK9_Quotation FOREIGN KEY (PaymentTermNo) REFERENCES ERP.PaymentTermData(PtdId),
 CONSTRAINT FK10_Quotation FOREIGN KEY (TradeTerm) REFERENCES ERP.TradeTermData(TtdId),
 CONSTRAINT FK11_Quotation FOREIGN KEY (TradeTerm1) REFERENCES ERP.TradeTermData(TtdId),
 CONSTRAINT FK12_Quotation FOREIGN KEY (TradeTerm2) REFERENCES ERP.TradeTermData(TtdId),
 CONSTRAINT FK13_Quotation FOREIGN KEY (ShipMethod) REFERENCES ERP.ShipMethodData(SmdId)
) /*報價單頭*/

CREATE TABLE ERP.QuotDetail( /*報價單身*/ /*ERP COPTB*/
 QuotDetailId				int					identity,
 QuotId						int					NOT NULL,						--所屬單頭
 QuotSequence				nvarchar(4)			NOT NULL,						--TB003	流水號
 MtlItemId					int					NOT	NULL,						--TB004	品號
 QuotMtlItemName			nvarchar(120)		NOT	NULL,						--TB005	品名
 QuotMtlItemSpec			nvarchar(120)		NOT	NULL,						--TB006	規格
 QuotQty					float				NOT NULL,						--TB007	數量
 FreebieQty					float				NOT NULL,						--TB015	贈品數量
 UomId						int					NOT	NULL,						--TB008	單位(Unit of Measure) >> 參照 ERP品號換算單位檔(INVMD)
 UnitPrice					float				NOT NULL,						--TB009	單價
 Amount						float				NOT NULL,						--TB010	金額
 EffectiveDate				datetime			NOT	NULL,						--TB016 生效日期
 ExpirationDate				datetime				NULL,						--TB017 失效日期
 Reamrk						nvarchar(255)			NULL,						--TB012 備註
 ValuationQty				float				NOT NULL,						--TB044 計價數量
 ComponentValuation			nvarchar(1)			NOT NULL,						--TB013	分量計價 >> BAS.[Status] N：未確認；Y：已確認
 BusinessTaxRate			float				NOT NULL,						--TB042 營業稅率 >> 由稅別碼決定(NN004)
 DiscountRate				float				NOT NULL,						--TB025	折扣率 1(DEF)
 TradeTermUnitPrice1		float				NOT NULL,						--TB037 交易條件單價1
 TradeTermAmount1			float				NOT NULL,						--TB038 交易條件金額1
 TradeTermUnitPrice2		float					NULL,						--TB040 交易條件單價2
 TradeTermAmount2			float					NULL,						--TB041 交易條件金額2
 ComponentUomId				int					NOT NULL,						--TB045	計價單位(Unit of Measure) >> 參照 ERP品號換算單位檔(INVMD)
 ConfirmStatus				nvarchar(1)			NOT NULL,						--TB011	確認碼 >> BAS.[Status] N：未確認；V：作廢；Y：已確認
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QuotDetail PRIMARY KEY (QuotDetailId),
 CONSTRAINT AK1_QuotDetail UNIQUE (QuotId, QuotSequence),
 CONSTRAINT FK1_QuotDetail FOREIGN KEY (CreateBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK2_QuotDetail FOREIGN KEY (LastModifiedBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK3_QuotDetail FOREIGN KEY (QuotId) REFERENCES ERP.Quotation(QuotId),
 CONSTRAINT FK4_QuotDetail FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId)
) /*報價單身*/

CREATE TABLE ERP.QuotDetailFile( /*報價單身明細檔*/ /*ERP COPTK*/
 QdfId						int					identity,
 QuotDetailId				int					NOT NULL,						--所屬單身
 QuotQty					float				NOT NULL,						--TK004	數量
 DiscountRate				float				NOT NULL,						--TK005	折扣率 1(DEF)
 UnitPrice					float				NOT NULL,						--TK006	單價
 Reamrk						nvarchar(255)			NULL,						--TK007 備註
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_QuotDetailFile PRIMARY KEY (QdfId),
 CONSTRAINT FK1_QuotDetailFile FOREIGN KEY (CreateBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK2_QuotDetailFile FOREIGN KEY (LastModifiedBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK3_QuotDetailFile FOREIGN KEY (QuotDetailId) REFERENCES ERP.QuotDetail(QuotDetailId)
) /*報價單身明細檔*/

CREATE TABLE ERP.SalesForecast( /*銷售預測單頭*/ /*ERP COPME*/
 SfId						int					identity,
 CompanyId					int					NOT NULL,						--ME013 廠別(所屬公司)
 SfErpNo					nvarchar(11)		NOT NULL,						--ME001 預測代號
 CustomerId					int					NOT NULL,						--ME002 客戶
 ProductionPlan				nvarchar(1)			NOT NULL,						--ME008 納入生產計畫　>> BAS.[Status] N：不納入；Y：納入
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_SalesForecast PRIMARY KEY (SfId),
 CONSTRAINT AK1_SalesForecast UNIQUE (CompanyId, SfErpNo),
 CONSTRAINT FK1_SalesForecast FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_SalesForecast FOREIGN KEY (CreateBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK3_SalesForecast FOREIGN KEY (LastModifiedBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK4_SalesForecast FOREIGN KEY (CustomerId) REFERENCES ERP.Customer(CustomerId)
) /*銷售預測單頭*/

CREATE TABLE ERP.SfDetail( /*銷售預測單身*/ /*ERP COPMF*/
 SfDetailId					int					identity,
 SfId						int					NOT NULL,						--所屬單頭
 SfSequence					nvarchar(4)			NOT NULL,						--MF002	流水號
 MtlItemId					int						NULL,						--MF003	品號
 SfMtlItemName				nvarchar(120)			NULL,						--MF004	品名
 SfMtlItemSpec				nvarchar(120)			NULL,						--MF005	規格
 InventoryId				int						NULL,						--MF007	預測庫別
 SfQty						float				NOT NULL,						--MF008	預測數量
 OrderQty					float				NOT NULL,						--MF009	已受訂量(已開訂單數量)
 UnitPrice					float				NOT NULL,						--MF012	單價
 Amount						float				NOT NULL,						--MF014	金額
 ClosureStatus				nvarchar(1)			NOT NULL,						--MF020	結案碼 >> BAS.[Status] N：未結案；y：指定結案
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_SfDetail PRIMARY KEY (SfDetailId),
 CONSTRAINT AK1_SfDetail UNIQUE (SfId, SfSequence),
 CONSTRAINT FK1_SfDetail FOREIGN KEY (CreateBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK2_SfDetail FOREIGN KEY (LastModifiedBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK3_SfDetail FOREIGN KEY (SfId) REFERENCES ERP.Quotation(SfId),
 CONSTRAINT FK4_SfDetail FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId),
 CONSTRAINT FK5_SoDetail FOREIGN KEY (InventoryId) REFERENCES ERP.Inventory(InventoryId)
) /*銷售預測單身*/

END /*訂單*/

BEGIN /*品號*/
CREATE TABLE ERP.CustomerMtlItemData( /*客戶品號資料檔*/ /*ERP COPMG*/
 CmidId						int					identity,
 CustomerId					int					NOT NULL,						--所屬客戶
 CmidNo						nvarchar(10)			NULL,						--MG001	客戶代號(備註：中揚未使用，紘立為必填)
 MtlItemId					int					NOT	NULL,						--MG002	品號
 CmidNo						nvarchar(40)		NOT	NULL,						--MG003	客戶品號
 CmidDesc					nvarchar(255)			NULL,						--MG004	客戶商品描述
 CmidName					nvarchar(120)			NULL,						--MG005	客戶品名
 CmidSpec					nvarchar(120)			NULL,						--MG006	客戶規格
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_CustomerMtlItemData PRIMARY KEY (CmidId),
 CONSTRAINT AK1_CustomerMtlItemData UNIQUE (CustomerId, MtlItemId),
 CONSTRAINT FK1_CustomerMtlItemData FOREIGN KEY (CreateBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK2_CustomerMtlItemData FOREIGN KEY (LastModifiedBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK3_CustomerMtlItemData FOREIGN KEY (CustomerId) REFERENCES ERP.Customer(CustomerId), --Table未開
 CONSTRAINT FK4_CustomerMtlItemData FOREIGN KEY (MtlItemId) REFERENCES ERP.MtlItem(MtlItemId) --Table未開
)/*客戶品號資料檔*/

CREATE TABLE ERP.CustomerMtlPrice( /*客戶商品計價單頭資料檔*/ /*ERP COPMB*/
 CmpId						int					identity,
 CustomerId					int					NOT NULL,						--所屬客戶
 MtlItemId					int					NOT NULL,						--所屬客戶品號
 TradeTerm					int					NOT NULL,						--MB019 所屬交易條件 >> 參照 ERP.TradeTermData
 ValuationUnit				nvarchar(6)			NOT	NULL,						--MB003 計價單位 >> 參照 ERP品號換算單位檔(INVMD)
 Currency					nvarchar(4)			NOT NULL,						--MB004 幣別
 ComponentValuation			nvarchar(1)			NOT NULL,						--MB007	分量計價 >> BAS.[Status] N：未確認；Y：已確認
 UnitPrice					float				NOT NULL,						--MB008	單價
 PricingDate				datetime			NOT NULL,						--MB009 核價日期
 NearestSiDate				datetime				NULL,						--MB010 上次銷貨日
 Reamrk						nvarchar(255)			NULL,						--MB012 備註
 InclusiveOfTax				nvarchar(1)			NOT NULL,						--MB013	含稅 >> BAS.[Status] N：未確認；Y：已確認
 FirstTradeDate				datetime				NULL,						--MB014 初次交易日
 EffectiveDate				datetime			NOT	NULL,						--MB017 生效日期
 ExpirationDate				datetime				NULL,						--MB018 失效日期
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_CustomerMtlPrice PRIMARY KEY (CmpId),
 CONSTRAINT FK1_CustomerMtlPrice FOREIGN KEY (CreateBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK2_CCustomerMtlPrice FOREIGN KEY (LastModifiedBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK3_CustomerMtlPrice FOREIGN KEY (CustomerId) REFERENCES ERP.Customer(CustomerId),
 CONSTRAINT FK4_CustomerMtlPrice FOREIGN KEY (MtlItemId) REFERENCES ERP.MtlItem(MtlItemId),
 CONSTRAINT FK5_CustomerMtlPrice FOREIGN KEY (TradeTerm) REFERENCES ERP.TradeTermData(TtdId)
)/*客戶商品計價單頭資料檔*/

CREATE TABLE ERP.CmpDetail( /*客戶商品計價單身資料檔*/ /*ERP COPMC*/
 CmpDetailId				int					identity,
 CmpId						int					NOT NULL,						--所屬客戶商品計價單頭
 ValuationUnit				nvarchar(6)			NOT	NULL,						--MC003 計價單位 >> 參照 ERP品號換算單位檔(INVMD)
 Currency					nvarchar(4)			NOT NULL,						--MC004 幣別
 CmpDetailQty				float				NOT NULL,						--MC005	數量
 UnitPrice					float				NOT NULL,						--MC006	單價
 Reamrk						nvarchar(255)			NULL,						--MC007 備註
 DiscountRate				float				NOT NULL,						--MC008	折扣率
 EffectiveDate				datetime			NOT	NULL,						--MC009 生效日期
 TradeTerm					int					NOT NULL,						--MC011 所屬交易條件 >> 參照 ERP.TradeTermData
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_CmpDetail PRIMARY KEY (CmpDetailId),
 CONSTRAINT FK1_CmpDetail FOREIGN KEY (CreateBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK2_CmpDetail FOREIGN KEY (LastModifiedBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK3_CmpDetail FOREIGN KEY (CmpId) REFERENCES ERP.CustomerMtlPrice(CmpId),
 CONSTRAINT FK4_CmpDetail FOREIGN KEY (TradeTerm) REFERENCES ERP.TradeTermData(TtdId)
)/*客戶商品計價單身資料檔*/

CREATE TABLE ERP.MtlItemCategoryData( /*品號類別資料檔*/ /*ERP INVMA*/
 MicdId						int					identity,
 MicdType					nvarchar(1)			NOT	NULL,						--MA001	分類方式
 MicdNo						nvarchar(6)			NOT	NULL,						--MA002	品號類別代號
 MicdName					nvarchar(40)		NOT	NULL,						--MA003	品號類別名稱
 InventoryAccount			nvarchar(20)			NULL,						--MA004	存貨會計科目
 SalesRevenueAccount		nvarchar(20)			NULL,						--MA005	銷貨收入科目
 SalesReturnAccount			nvarchar(20)			NULL,						--MA006	銷貨退回科目
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MtlItemCategoryData PRIMARY KEY (MicdId),
 CONSTRAINT FK1_MtlItemCategoryData FOREIGN KEY (CreateBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK2_MtlItemCategoryData FOREIGN KEY (LastModifiedBy) REFERENCES BAS.[User](UserId)
)/*品號類別資料檔*/

CREATE TABLE ERP.MtlItem( /*品號基本資料檔*/ /*ERP INVMB*/
 MtlItemId					int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 MtlItemNo					nvarchar(40)		NOT NULL,						--MB001 品號
 MtlItemName				nvarchar(120)		NOT NULL,						--MB002 品名
 MtlItemSpec				nvarchar(120)		NOT NULL,						--MB003 規格
 InventoryUnit				nvarchar(6)			NOT	NULL,						--MB004 庫存單位 >> 參照 ERP品號換算單位檔(INVMD)
 TypeOne					nvarchar(6)			NOT NULL,						--MB005 品號分類一
 TypeTwo					nvarchar(6)			NOT NULL,						--MB006 品號分類二
 TypeThree					nvarchar(6)			NOT NULL,						--MB007 品號分類三
 TypeFour					nvarchar(6)			NOT NULL,						--MB008 品號分類四
 ArticleNo					nvarchar(40)		NOT NULL,						--MB080 貨號(同品號資料)
 [Version]					nvarchar(4)			NOT NULL,						--MB165 版次
 MtlModify					nvarchar(1)			NOT NULL,						--MB066 修改品名/規格 >> BAS.[Status] N：未確認；Y：已確認
 InventoryId				int						NULL,						--MB017 主要庫別
 RequisitionInventoryId		int						NULL,						--MB157 主要領料庫別
 ProductionLine				nvarchar(10)			NULL,						--MB068 生產線別 >> 參照 ERP生產線資料單頭檔(CMSMD)
 PurchaseUomId				int						NULL,						--MB155 採購單位 >> 參照 ERP品號換算單位檔(INVMD)
 SaleUomId					int						NULL,						--MB156 銷售單位 >> 參照 ERP品號換算單位檔(INVMD)
 WeightUomId				int						NULL,						--MB015 重量單位 >> 參照 ERP品號換算單位檔(INVMD)
 BondedStore				nvarchar(1)			NOT NULL,						--MB020 保稅品 >> BAS.[Status] N：未確認；Y：已確認
 InventoryManagement		nvarchar(2)			NOT NULL,						--MB019 庫存管理 >> BAS.[Status] N：未確認；Y：已確認
 LotManagement				nvarchar(1)			NOT NULL	DEFAULT 'N',		--MB022 批號管理 >> BAS.[Type] N:不需要；Y:需要不檢查庫存量；W:僅需警告；T:需要且檢查庫存量
 OverDeliveryManagement		nvarchar(1)			NOT NULL	DEFAULT 'N',		--MB087 超交管理 >> BAS.[Status] N：未確認(DEF)；Y：已確認
 ItemAttribute				nvarchar(1)			NOT NULL,						--MB025 品號屬性 >> BAS.[Type] P：採購件；M：自製件；S：託外加工件；Y：虛設品號；F：Feature件；O:Option件
 EffectiveDate				datetime				NULL,						--MB030 生效日期
 ExpirationDate				datetime				NULL,						--MB031 失效日期
 MtlItemRemark				nvarchar(255)			NULL,						--MB028 品號備註
 MtlItemDesc				nvarchar(255)			NULL,						--MB009 商品描述
 PurchasingOfficer			nvarchar(10)		NOT NULL,						--MB067 採購人員 >> 參照 ERP職務人員檔(CMSMK)
 ReplenishmentPolicy		nvarchar(1)			NOT NULL,						--MB034 補貨政策 >> BAS.[Type] R:依補貨點；M:依MRP需求；L:依LRP需求；N:不需；H:依歷史銷售
 MeasureType				nvarchar(1)			NOT NULL	DEFAULT '2',		--MB043 檢驗方式 >> BAS.[Type] 0：免檢；1：抽檢(減量)；2：抽檢(正常)；3：抽檢(加嚴)；4：全檢
 MaterialRequisitionNo		nvarchar(1)			NOT NULL,						--MB042 領料碼 >> BAS.[Type] 1:逐批領料；2:自動扣料；3:單獨領料
 OverReceiptManagement		nvarchar(1)			NOT NULL,						--MB044 超收管理 >> BAS.[Status] N：未確認；Y：已確認
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
 CONSTRAINT FK2_MtlItem FOREIGN KEY (CreateBy) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK3_MtlItem FOREIGN KEY (LastModifiedBy) REFERENCES BAS.[User](UserId)
)/*品號基本資料檔*/

END /*品號*/
