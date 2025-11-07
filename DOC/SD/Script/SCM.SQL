BEGIN /*基本資料*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ExchangeRate') > 0 DROP TABLE SCM.ExchangeRate
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Forwarder') > 0 DROP TABLE SCM.Forwarder
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'CustomerSalesmen') > 0 DROP TABLE SCM.CustomerSalesmen
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DeliveryCustomer') > 0 DROP TABLE SCM.DeliveryCustomer
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Customer') > 0 DROP TABLE SCM.Customer
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Supplier') > 0 DROP TABLE SCM.Supplier
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Inventory') > 0 DROP TABLE SCM.Inventory 
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Packing') > 0 DROP TABLE SCM.Packing
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ProductTypeGroup') > 0 DROP TABLE SCM.ProductTypeGroup
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ProductType') > 0 DROP TABLE SCM.ProductType
END /*基本資料*/

BEGIN /*訂單*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'SomDetail') > 0 DROP TABLE SCM.SomDetail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'SoModification') > 0 DROP TABLE SCM.SoModification
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'SoCertificate') > 0 DROP TABLE SCM.SoCertificate
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'SoDetail') > 0 DROP TABLE SCM.SoDetail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'SaleOrder') > 0 DROP TABLE SCM.SaleOrder
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'SoDetailTemp') > 0 DROP TABLE SCM.SoDetailTemp
END /*訂單*/

BEGIN /*出貨*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'InventoryTempHold') > 0 DROP TABLE SCM.InventoryTempHold
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'PickingItem') > 0 DROP TABLE SCM.PickingItem
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DoDetail') > 0 DROP TABLE SCM.DoDetail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DeliveryOrder') > 0 DROP TABLE SCM.DeliveryOrder
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DeliveryDateLog') > 0 DROP TABLE SCM.DeliveryDateLog
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'PickingCarton') > 0 DROP TABLE SCM.PickingCarton
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DeliveryCustomerReport') > 0 DROP TABLE SCM.DeliveryCustomerReport
END /*出貨*/

BEGIN /*請購*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'PrmFile') > 0 DROP TABLE SCM.PrmFile
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'PrmLog') > 0 DROP TABLE SCM.PrmLog
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'PrmDetail') > 0 DROP TABLE SCM.PrmDetail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'PrModification') > 0 DROP TABLE SCM.PrModification
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'PrFile') > 0 DROP TABLE SCM.PrFile
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'PrLog') > 0 DROP TABLE SCM.PrLog
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'PrDetail') > 0 DROP TABLE SCM.PrDetail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'PurchaseRequisition') > 0 DROP TABLE SCM.PurchaseRequisition
END /*請購*/

BEGIN /*進貨*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'GrDetail') > 0 DROP TABLE SCM.GrDetail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'GoodsReceipt') > 0 DROP TABLE SCM.GoodsReceipt
END /*進貨*/

BEGIN /*庫存*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ItDetail') > 0 DROP TABLE SCM.ItDetail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'InventoryTransaction') > 0 DROP TABLE SCM.InventoryTransaction
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'TsrnDetail') > 0 DROP TABLE SCM.TsrnDetail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'TempShippingReturnNote') > 0 DROP TABLE SCM.TempShippingReturnNote
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'TsnDetail') > 0 DROP TABLE SCM.TsnDetail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'TempShippingNote') > 0 DROP TABLE SCM.TempShippingNote
END /*庫存*/

BEGIN /*退(換)貨*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'RmaItem') > 0 DROP TABLE SCM.RmaItem
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'RmaDetail') > 0 DROP TABLE SCM.RmaDetail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'RmaTransferLog') > 0 DROP TABLE SCM.RmaTransferLog
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ReturnMerchandiseAuthorization') > 0 DROP TABLE SCM.ReturnMerchandiseAuthorization
END /*退貨*/

BEGIN /*需求單管理*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Demand') > 0 DROP TABLE SCM.Demand
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DemandLine') > 0 DROP TABLE SCM.DemandLine
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DemandProcess') > 0 DROP TABLE SCM.DemandProcess
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DemandProductType') > 0 DROP TABLE SCM.DemandProductType
END /*需求單管理*/

--******************************************************************************************************************
BEGIN /*基本資料*/
CREATE TABLE SCM.ExchangeRate( /*匯率*/ /*ERP CMSMG*/
 ExchangeRateId				int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 Currency					nvarchar(5)			NOT NULL,						--MG001 幣別 >> BAS.[Type] NT$：台幣；CNY：人民幣；USD：美元；JPY：日元；HKD：港元；EUR：歐元；GBP：英鎊；CHF：瑞士法郎
 EffectiveDate				datetime			NOT NULL,						--MG002 生效日期
 BankBuyingRate				nvarchar(10)		NOT NULL,						--MG003 銀行買進匯率
 BankSellingRate			nvarchar(10)		NOT NULL,						--MG004 銀行賣出匯率
 CustomsBuyingRate			nvarchar(10)		NOT NULL,						--MG005 報關買進匯率
 CustomsSellingRate			nvarchar(10)		NOT NULL,						--MG006 報關賣出匯率
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ExchangeRate PRIMARY KEY (ExchangeRateId),
 CONSTRAINT AK1_ExchangeRate UNIQUE (CompanyId, Currency, EffectiveDate)
)/*匯率*/

CREATE TABLE SCM.Customer( /*客戶資訊*/ /*ERP COPMA*/
 CustomerId					int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 DepartmentId				int					    NULL,						--MA015 所屬部門
 CustomerNo					nvarchar(50)		NOT NULL,						--MA001 客戶代號
 CustomerName				nvarchar(200)		NOT NULL,						--MA003 客戶全名
 CustomerEnglishName		nvarchar(80)			NULL,						--MA110 客戶英文全名
 CustomerShortName			nvarchar(100)		NOT NULL,						--MA002 客戶簡稱
 Country					nvarchar(6)				NULL,						--MA019 國家別 >> 參照 ERP交易對象分類方式檔(CMSMR)
 Region						nvarchar(6)				NULL,						--MA018 地區別 >> 參照 ERP交易對象分類方式檔(CMSMR)
 [Route]					nvarchar(6)				NULL,						--MA077 路線別 >> 參照 ERP交易對象分類方式檔(CMSMR)
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
 Currency					nvarchar(4)			NOT NULL,						--MA014 交易幣別 >> 參照 ERP幣別匯率檔單頭(CMSMF)
 TradeTerm					nvarchar(1)			NOT NULL,						--MA109 交易條件 >> 參照 ERP交易條件代號檔(CMSNK)
 PaymentTerm				nvarchar(16)		NOT NULL,						--MA031 付款條件 >> 參照 ERP付款條件檔(CMSNA) >>COPMA MA083
 PricingType				nvarchar(1)			NOT NULL,						--MA035 計價方式 >> BAS.[Type] 1：利潤率；2：折扣率
 PaymentType				nvarchar(1)			NOT NULL,						--MA041 收款方式 >> BAS.[Type] 1：現金；2：電匯；3：支票；4：其他
 ReceiptReceive				nvarchar(1)			NOT NULL,						--MA042 票據寄領 >> BAS.[Type] 1：郵寄；2：自領；3：其它
 TaxAmountCalculateType		nvarchar(1)			NOT NULL,						--MA087 稅額計算方式 >> BAS.[Type] 1：整張資料計算；2：單身單筆資料計算
 DocumentDeliver			nvarchar(1)			NOT NULL,						--MA039 單據發送方式 >> BAS.[Type] 1：郵寄；2：FAX；3：EDI；4：E-MAIL
 InvoiceCount				nvarchar(1)			NOT NULL,						--MA037 發票聯數 >> 參照 ERP發票聯數檔(CMSNM)
 TaxNo						nvarchar(3)			NOT NULL,						--MA118 稅別碼 >> 參照 ERP稅別碼設定檔(CMSNN)
 Taxation					nvarchar(1)			NOT NULL,						--MA038 課稅別 >> 由稅別碼決定(NN006) 1：內含；2：外加；3：零稅率；4：免稅；9：不計稅
 PaymentBankFirst			nvarchar(20)			NULL,						--MA046 付款銀行(一) >> 參照 ERP金融機構資料檔(CMSMO)
 PaymentBankSecond			nvarchar(20)			NULL,						--MA069 付款銀行(二) >> 參照 ERP金融機構資料檔(CMSMO)
 PaymentBankThird			nvarchar(20)			NULL,						--MA070 付款銀行(三) >> 參照 ERP金融機構資料檔(CMSMO)
 BankAccountFirst			nvarchar(30)			NULL,						--MA071 銀行帳號(一)
 BankAccountSecond			nvarchar(30)			NULL,						--MA072 銀行帳號(二)
 BankAccountThird			nvarchar(30)	 		NULL,						--MA073 銀行帳號(三)
 Account					nvarchar(20)			NULL,						--MA047 帳款科目 >> 參照 ERP會計科目資料檔(ACTMA)
 AccountInvoice				nvarchar(20)			NULL,						--MA074 票據科目 >> 參照 ERP會計科目資料檔(ACTMA)
 ClearanceType				nvarchar(1)			NOT NULL,						--MA104 通關方式 >> BAS.[Type] 1：非經海關(DEF)；2：經海關
 ShipMethod					nvarchar(2)			NOT NULL,						--MA048 運輸方式 >> 參照 ERP運輸方式代號檔(CMSNJ)
 InvocieAttachedStatus		nvarchar(1)			NOT NULL,						--MA086 隨貨附發票 >> BAS.[Status]
 CustomerKind				nvarchar(1)				NULL,						--MA137 客戶分類 >> BAS.[Type] 1：自然人；2：法人
 UploadType					nvarchar(1)				NULL,						--MA141 上傳類型 >> BAS.[Type] 1：B2C存證；2：B2B存證；3：B2B交換
 DepositRate				float				NOT NULL,						--MA095 訂金比率
 SaleRating					nvarchar(1)				NULL,						--MA028 銷售評等 >> BAS.[Type] 1：A；2：B；3：C
 CreditRating				nvarchar(1)				NULL,						--MA029 信用評等 >> BAS.[Type] 1：A；2：B；3：C
 CreditLimit				float					NULL,						--MA033 信用額度
 CreditLimitControl			nvarchar(1)				NULL,						--MA032 信用額度管制 >> BAS.[Type] Y：依公司參數管制；N：信用額度不管制；y：依客戶資料管制
 CreditLimitControlCurrency nvarchar(1)			NOT NULL,						--MA133 信用額度控管幣別 >> BAS.[Type] 1：記帳幣別(DEF)；2：主要交易幣別
 SoCreditAuditType			nvarchar(1)			NOT NULL,						--MA088 訂單信用查核方式 >> BAS.[Type] 1.不檢查(DEF) 2.警告 3.拒絕 
 SiCreditAuditType			nvarchar(1)			NOT NULL,						--MA089 銷貨信用查核方式 >> BAS.[Type] 1.不檢查(DEF) 2.警告 3.拒絕
 DoCreditAuditType			nvarchar(1)			NOT NULL,						--MA120 出貨通知信用查核方式 >> BAS.[Type] 1.不檢查(DEF) 2.警告 3.拒絕
 InTransitCreditAuditType	nvarchar(1)			NOT NULL,						--MA132 暫出單信用查核方式 >> BAS.[Type] 1.不檢查(DEF) 2.警告 3.拒絕
 RelatedPerson				nvarchar(20)			NULL,						--MA124 關係人代號 >> 參照 ERP子母公司資料檔(FCSMB)
 SalesmenId					int						NULL,						--MA016 業務員
 PaymentSalesmenId			int						NULL,						--MA085 收款業務員
 ShipType					nvarchar(50)			NULL,						--運輸種類 >> BAS.[Type] 1：出貨；2：三角貿易
 ForwarderId				int						NULL,						--貨運承攬商
 CustomerRemark				nvarchar(500)			NULL,						--MA049 客戶備註
 TransferStatus				nvarchar(2)			NOT NULL,						--拋轉ERP狀態 >> BAS.[Status]
 TransferDate				datetime				NULL,						--拋轉ERP時間
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Customer PRIMARY KEY (CustomerId),
 CONSTRAINT AK1_Customer UNIQUE (CompanyId, CustomerNo),
 --CONSTRAINT AK2_Customer UNIQUE (CustomerName),
 CONSTRAINT FK1_Customer FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_Customer FOREIGN KEY (DepartmentId) REFERENCES BAS.Department(DepartmentId),
 CONSTRAINT FK3_Customer FOREIGN KEY (SalesmenId) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK4_Customer FOREIGN KEY (PaymentSalesmenId) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK5_Customer FOREIGN KEY (ForwarderId) REFERENCES SCM.Forwarder(ForwarderId)
)/*客戶資訊*/

CREATE TABLE SCM.CustomerSalesmen( /*客戶業務*/
 CustomerId					int					NOT NULL,						--客戶
 SalesmenId					int					NOT NULL,						--業務
 SortNumber					int					NOT NULL,						--排序號碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT FK1_CustomerSalesmen FOREIGN KEY (CustomerId) REFERENCES SCM.Customer(CustomerId),
 CONSTRAINT FK2_CustomerSalesmen FOREIGN KEY (SalesmenId) REFERENCES BAS.[User](UserId)
)/*客戶業務*/

CREATE TABLE SCM.DeliveryCustomer( /*送貨客戶(間接客戶)*/
 DcId						int					identity,
 CustomerId					int					NOT	NULL,						--客戶來源(客戶對應種類為間接者，則須維護客戶來源)
 DcName						nvarchar(200)		NOT NULL,						--客戶全名
 DcEnglishName				nvarchar(80)			NULL,						--客戶英文全名
 DcShortName				nvarchar(100)		NOT NULL,						--客戶簡稱
 Contact					nvarchar(30)			NULL,						--客戶聯絡人
 TelNo						nvarchar(20)			NULL,						--客戶聯絡電話
 FaxNo						nvarchar(20)			NULL,						--客戶聯絡傳真
 RegisteredAddress			nvarchar(255)			NULL,						--登記地址
 DeliveryAddress			nvarchar(255)		NOT NULL,						--送貨地址
 ShipType					nvarchar(50)			NULL,						--運輸種類 >> BAS.[Type] 1：出貨；2：三角貿易
 ForwarderId				int						NULL,						--貨運承攬商
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DeliveryCustomer PRIMARY KEY (DcId),
 CONSTRAINT FK1_DeliveryCustomer FOREIGN KEY (CustomerId) REFERENCES SCM.Customer(CustomerId),
 CONSTRAINT FK2_DeliveryCustomer FOREIGN KEY (ForwarderId) REFERENCES SCM.Forwarder(ForwarderId)
)

CREATE TABLE SCM.Forwarder( /*貨運承攬商*/
 ForwarderId				int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 ShipMethod					nvarchar(2)			NOT NULL,						--所屬運輸方式 >> 參照 ERP運輸方式代號檔(CMSNJ)
 ForwarderNo				nvarchar(50)		NOT NULL,						--貨運承攬商代號 >> Fedex；DHL；UPS；順豐；空運
 ForwarderName				nvarchar(100)		NOT NULL,						--貨運承攬商名稱
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Forwarder PRIMARY KEY (ForwarderId),
 CONSTRAINT AK1_Forwarder UNIQUE (ForwarderNo),
 CONSTRAINT AK2_Forwarder UNIQUE (ForwarderName),
 CONSTRAINT FK1_Forwarder FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)/*貨運承攬商*/

CREATE TABLE SCM.Supplier( /*供應商*/ /*ERP PURMA*/
 SupplierId					int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 SupplierNo					nvarchar(50)		NOT NULL,						--MA001 供應商代號
 SupplierName				nvarchar(200)		NOT NULL,						--MA003 供應商全名
 SupplierShortName			nvarchar(100)		NOT NULL,						--MA002 供應商簡稱
 SupplierEnglishName		nvarchar(100)		NOT NULL,						--MA088 供應商英文全名
 Country					nvarchar(6)				NULL,						--MA006 國家別 >> 參照 ERP交易對象分類方式檔(CMSMR)
 Region						nvarchar(6)				NULL,						--MA007 地區別 >> 參照 ERP交易對象分類方式檔(CMSMR)
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
 InauguateDate				datetime				NULL,						--MA017 開業日
 AccountMonth				nvarchar(1)				NULL,						--MA034 結帳月
 AccountDay					nvarchar(2)			NOT NULL,						--MA035 結帳日
 [Version]					nvarchar(4)			NOT NULL,						--MA069 版次
 Capital					float				NOT NULL,						--MA018 資本額
 Headcount					int					NOT NULL,						--MA019 員工人數
 PoDeliver					nvarchar(1)			NOT NULL,						--MA020 採購單發送方式 >> BAS.[Type] 1：郵寄；2：FAX；3：EDI；4：E-MAIL
 Currency					nvarchar(4)			NOT NULL,						--MA021 交易幣別 >> 參照 ERP幣別匯率檔單頭(CMSMF)
 TradeTerm					nvarchar(1)			NOT NULL,						--MA077 交易條件 >> 參照 ERP交易條件代號檔(CMSNK)
 PaymentType				nvarchar(1)			NOT NULL,						--MA024 付款方式 >> BAS.[Type] 1：現金；2：電匯；3：支票；4：其他
 PaymentTerm				nvarchar(16)		NOT NULL,						--MA025 付款條件 >> 參照 ERP付款條件檔(CMSNA) >>PURMA MA055
 ReceiptReceive				nvarchar(1)			NOT NULL,						--MA029 票據寄領 >> BAS.[Type] 1：郵寄；2：自領；3：其它
 InvoiceCount				nvarchar(1)			NOT NULL,						--MA030 發票聯數 >> 參照 ERP發票聯數檔(CMSNM)
 TaxNo						nvarchar(3)			NOT NULL,						--MA076 稅別碼 >> 參照 ERP稅別碼設定檔(CMSNN)
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
 HideSupplier				nvarchar(1)			NOT	NULL	DEFAULT 'N',		--集團內/集團外廠商
 PassStationControl			nvarchar(1)			NOT	NULL	DEFAULT 'N',		--是否支援刷取過站
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
)/*供應商*/

CREATE TABLE SCM.SupplierMachine( /*供應商機台資訊*/
 SmId						int					identity,
 SupplierId					int					NOT NULL,						--客戶ID
 ProcessId					int					NOT NULL,						--製程ID
 MachineId					int					NOT NULL,						--機台ID
 [Status]					nvarchar(2)			NOT NULL	DEFAULT 'A',		--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_SupplierMachine PRIMARY KEY (SmId),
 CONSTRAINT AK1_SupplierMachine UNIQUE (SupplierId, ProcessId, MachineId),
 CONSTRAINT FK1_SupplierMachine FOREIGN KEY (SupplierId) REFERENCES SCM.Supplier(SupplierId),
 CONSTRAINT FK2_SupplierMachine FOREIGN KEY (ProcessId) REFERENCES MES.Process(ProcessId),
 CONSTRAINT FK3_SupplierMachine FOREIGN KEY (MachineId) REFERENCES MES.Machine(MachineId),
)

CREATE TABLE SCM.Inventory( /*庫別資訊*/ /*ERP CMSMC*/
 InventoryId				int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 InventoryNo				nvarchar(50)		NOT NULL,						--MC001 庫別代號
 InventoryName				nvarchar(100)		NOT NULL,						--MC002 庫別名稱
 InventoryType				nvarchar(2)			NOT NULL,						--MC004 庫別性質 >> BAS.[Type] 1.存貨倉、2.非存貨倉
 MrpCalculation				nvarchar(2)			NOT NULL,						--MC005 納入MRP計算 >> BAS.[Status]
 ConfirmStatus				nvarchar(2)			NOT NULL,						--MC006 確認時庫存量不足准許出庫 >> BAS.[Status]
 SaveStatus					nvarchar(2)			NOT NULL,						--MC008 存檔時庫存量不足准許出庫 >> BAS.[Status]
 InventoryDesc				nvarchar(100)		    NULL,						--MC007 庫別描述
 TransferStatus				nvarchar(2)			NOT NULL,						--拋轉ERP狀態 >> BAS.[Status]
 TransferDate				datetime				NULL,						--拋轉ERP時間
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Inventory PRIMARY KEY (InventoryId),
 CONSTRAINT AK1_Inventory UNIQUE (CompanyId, InventoryNo),
 --CONSTRAINT AK2_Inventory UNIQUE (InventoryName),
 CONSTRAINT FK1_Inventory FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)

CREATE TABLE SCM.Packing( /*包材資訊*/
 PackingId					int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 PackingName				nvarchar(50)		NOT NULL,						--包材名稱
 PackingType				nvarchar(2)			NOT NULL,						--包材種類 >> BAS.[Type] 1：棧板；2：紙箱；3：木箱
 VolumeSpec					nvarchar(100)		NOT NULL,						--體積規格
 Volume						float				NOT NULL,						--包材體積
 VolumeUomId				int					NOT NULL,						--體積單位(Unit of Measure)
 WeightSpec					nvarchar(100)		NOT NULL,						--重量規格
 [Weight]					float				NOT NULL,						--包材重量
 WeightUomId				int					NOT NULL,						--重量單位(Unit of Measure)
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Packing PRIMARY KEY (PackingId),
 CONSTRAINT AK1_Packing UNIQUE (PackingName),
 CONSTRAINT FK1_Packing FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_Packing FOREIGN KEY (VolumeUomId) REFERENCES PDM.UnitOfMeasure(UomId),
 CONSTRAINT FK3_Packing FOREIGN KEY (WeightUomId) REFERENCES PDM.UnitOfMeasure(UomId)
)

CREATE TABLE SCM.PerformanceGoals ( /*業務績效目標*/
 PgId								int					identity,
 PgNo								nvarchar(20)		NOT NULL,						--目標代碼
 PgName								nvarchar(100)		NOT NULL,						--目標名稱
 PgDesc								nvarchar(100)			NULL,						--目標描述
 StartDate							datetime			NOT NULL	DEFAULT GETDATE(),	--起始時間
 EndDate							datetime			NOT NULL	DEFAULT GETDATE(),	--結束日期
 [Status]							nvarchar(2)			NOT NULL,						--狀態碼	
 CreateDate							datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate					datetime				NULL,						--修改日期
 CreateBy							int					NOT NULL,						--新增人員
 LastModifiedBy						int						NULL,						--修改人員 
 CONSTRAINT PK_PerformanceGoals PRIMARY KEY (PgId)
 )

 CREATE TABLE SCM.PgDetail ( /*業務績效目標單身*/
 PgDetailId							int					identity,
 PgId								int					NOT	NULL,						--業務績效目標單頭Id
 UserId								int					NOT	NULL,						--業務人員
 IntentType							nvarchar(2)			NOT NULL,						--目標類別(接單:S,出貨:D,銷貨:R)
 IntentAmount						int					NOT NULL,						--目標金額
 ConfirmStatus						nvarchar(1)				NULL,						--確認碼 >> BAS.[Status] N：未確認；V：作廢；Y：已確認
 ConfirmDate						datetime				NULL,						--確認日期
 CreateDate							datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate					datetime				NULL,						--修改日期
 CreateBy							int					NOT NULL,						--新增人員
 LastModifiedBy						int						NULL,						--修改人員 
 CONSTRAINT PK_PgDetail PRIMARY KEY (PgDetailId),
 CONSTRAINT AK1_PgDetail UNIQUE (PgId, UserId, IntentType),
 CONSTRAINT FK1_PgDetail FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId)
 )

 CREATE TABLE SCM.ProductTypeGroup ( /*產品群組 RFI*/
 ProTypeGroupId						int					identity,
 RfqProClassId						int					NOT NULL,
 ProTypeGroupName					nvarchar(50)		NOT NULL,						--2階段
 CoatingFlag						nvarchar(1)			NOT NULL,						--是否鍍膜
 [Status]							nvarchar(2)			NOT NULL,						--狀態
 CreateDate							datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate					datetime				NULL,						--修改日期
 CreateBy							int					NOT NULL,						--新增人員
 LastModifiedBy						int						NULL,						--修改人員 
 CONSTRAINT PK_ProductTypeGroup PRIMARY KEY (ProTypeGroupId),
 CONSTRAINT AK1_ProductTypeGroup UNIQUE (RfqProClassId, ProTypeGroupName),
 CONSTRAINT FK1_ProductTypeGroup FOREIGN KEY (RfqProClassId) REFERENCES SCM.[RfqProductClass](RfqProClassId)
 )

 CREATE TABLE SCM.ProductType ( /*產品群組 RFI*/
 ProTypeId							int					identity,
 ProTypeGroupId						int					NOT NULL,
 ProTypeName						nvarchar(100)		NOT NULL,						--2階段
 [Status]							nvarchar(2)			NOT NULL,						--狀態
 CreateDate							datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate					datetime				NULL,						--修改日期
 CreateBy							int					NOT NULL,						--新增人員
 LastModifiedBy						int						NULL,						--修改人員 
 CONSTRAINT PK_ProductType PRIMARY KEY (ProTypeId),
 CONSTRAINT AK1_ProductType UNIQUE (ProTypeGroupId, ProTypeName),
 CONSTRAINT FK1_ProductType FOREIGN KEY (ProTypeGroupId) REFERENCES SCM.[ProductTypeGroup](ProTypeGroupId)
 )

 CREATE TABLE SCM.Project( /*專案主檔 CMSNB*/
 ProjectId					int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 ProjectNo					nvarchar(20)		NOT NULL,						--NB001 專案代號
 ProjectName				nvarchar(40)		NOT NULL,						--NB002 專案名稱
 Remark						nvarchar(255)			NULL,						--NB003 備註
 EffectiveDate				datetime				NULL,						--NB004 生效日期
 ExpirationDate				datetime				NULL,						--NB005 失效日期
 CloseCode					nvarchar(1)			NOT	NULL,						--NB006 結案碼 Y/N
 TransferErpStatus			nvarchar(1)			NOT	NULL	DEFAULT 'N',		--是否拋轉ERP Y/N
 TransferErpDate			datetime				NULL	DEFAULT 'N',		--拋轉至ERP時間
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Project PRIMARY KEY (ProjectId),
 CONSTRAINT AK1_Project UNIQUE (CompanyId, ProjectNo),
 CONSTRAINT FK1_Project FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
)/*專案主檔 CMSNB*/

 CREATE TABLE SCM.ProjectDetail( /*專案詳細資料*/
 ProjectDetailId			int					identity,
 ProjectId					int					NOT NULL,						--所屬專案ID
 ProjectType				nvarchar(1)			NOT NULL,						--專案類型 1:物料預算 2:人力成本預算 3:製造費用預算
 Currency					nvarchar(4)			NOT	NULL,						--費用幣別
 ExchangeRate				float				NOT	NULL,						--匯率
 BudgetAmount				float				NOT NULL,						--預算費用
 LocalBudgetAmount			float				NOT NULL,						--預算費用(本幣)
 Edition					nvarchar(4)			NOT NULL	DEFAULT '0000',		--版本(第一版預設為0000)，若有變更過則進版 EX:0001
 BpmTransferStatus			nvarchar(1)			NOT NULL	DEFAULT 'N',		--拋轉BPM狀態 Y/N
 [Status]					nvarchar(1)			NOT NULL	DEFAULT 'N',		--目前專案預算狀態 N:新單據 P:BPM簽核中 E:BPM駁回修改 Y:已核准 C:預算變更BPM簽核中 K:預算變更駁回修改中
 Remark						nvarchar(255)			NULL,						--備註
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ProjectDetail PRIMARY KEY (ProjectDetailId),
 CONSTRAINT AK1_ProjectDetail UNIQUE (ProjectId, ProjectType),
 CONSTRAINT FK1_ProjectDetail FOREIGN KEY (ProjectId) REFERENCES SCM.Project(ProjectId),
)/*專案詳細資料*/

 CREATE TABLE SCM.ProjectBudgetChangeLog( /*專案預算變更紀錄*/
 LogId						int					identity,
 ProjectDetailId			int					NOT NULL,						--所屬專案詳細資料ID
 Currency					nvarchar(4)			NOT	NULL,						--費用幣別
 ExchangeRate				float				NOT	NULL,						--匯率
 BudgetAmount				float				NOT NULL,						--目前預算費用
 LocalBudgetAmount			float				NOT NULL,						--目前預算費用(本幣)
 OriCurrency				nvarchar(4)			NOT	NULL,						--前一版本費用幣別
 OriExchangeRate			float				NOT	NULL,						--前一版本匯率
 OriBudgetAmount			float				NOT NULL,						--前一版本預算費用
 OriLocalBudgetAmount		float				NOT NULL,						--前一版本預算費用(本幣)
 OriEdition					nvarchar(4)			NOT NULL,						--前一版本號碼
 BpmStatus					nvarchar(1)			NOT NULL	DEFAULT 'N',		--BPM簽核狀態 Y:已核准 N:簽核中 E:駁回
 Remark						nvarchar(255)			NULL,						--變更備註
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ProjectBudgetChangeLog PRIMARY KEY (LogId),
 CONSTRAINT FK1_ProjectBudgetChangeLog FOREIGN KEY (ProjectDetailId) REFERENCES SCM.ProjectDetail(ProjectDetailId),
)/*專案預算變更紀錄*/

CREATE TABLE SCM.ProjectFile( /*專案預算附檔*/
 ProjectFileId				int					identity,
 ProjectDetailId			int					NOT NULL,						--所屬專案預算ID
 FileId						int						NULL,						--檔案ID
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ProjectFile PRIMARY KEY (ProjectFileId),
 CONSTRAINT FK1_ProjectFile FOREIGN KEY (ProjectDetailId) REFERENCES SCM.ProjectDetail(ProjectDetailId),
 CONSTRAINT FK2_ProjectFile FOREIGN KEY (FileId) REFERENCES BAS.[File](FileId)
)/*專案預算附檔*/

CREATE TABLE SCM.ProjectChangeFile( /*專案預算變更附檔*/
 ChangeFileId				int					identity,
 LogId						int					NOT NULL,						--變更紀錄ID
 FileId						int						NULL,						--檔案ID
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ProjectChangeFile PRIMARY KEY (ChangeFileId),
 CONSTRAINT FK1_ProjectChangeFile FOREIGN KEY (LogId) REFERENCES SCM.ProjectBudgetChangeLog(LogId),
 CONSTRAINT FK2_ProjectChangeFile FOREIGN KEY (FileId) REFERENCES BAS.[File](FileId)
)/*專案預算變更附檔*/

CREATE TABLE SCM.ProjectLog( /*專案預算拋轉BPM LOG*/
 ProjectLogId				int					identity,
 ProjectDetailId			int					NOT NULL,						--所屬專案預算ID
 LogType					nvarchar(1)			NOT NULL	DEFAULT 'P',		--LOG類型 P:專案預算 C:專案預算變更
 RootId						nvarchar(30)		NOT NULL,						--任務編號
 BpmNo						nvarchar(100)		NOT	NULL,						--BPM單號
 TransferBpmDate			datetime			NOT NULL,						--拋轉BPM日期
 BpmStatus					nvarchar(2)			NOT NULL,						--拋轉狀態
 ConfirmUser				nvarchar(100)		NOT NULL,						--最後核單USER NO
 ErrorMessage				nvarchar(1000)		NOT NULL,						--Error訊息
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ProjectLog PRIMARY KEY (ProjectLogId),
 CONSTRAINT FK1_ProjectLog FOREIGN KEY (ProjectDetailId) REFERENCES SCM.ProjectDetail(ProjectDetailId),
)
END /*基本資料*/

BEGIN /*批號相關資料*/
CREATE TABLE SCM.LotNumber( /*批號資料單頭*/ /*ERP INVME*/
 LotNumberId				int					identity,
 MtlItemId					int					NOT NULL,						--ME001 品號
 LotNumberNo				nvarchar(20)		NOT NULL,						--ME002 批號
 FirstRecriptDate			datetime			NOT NULL	DEFAULT GETDATE(),	--ME003 最早入庫日
 FromErpPrefix				nvarchar(4)				NULL,						--ME005 參考單別
 FromErpNo					nvarchar(11)			NULL,						--ME006 參考單號
 CloseStatus				nvarchar(1)			NOT NULL	DEFAULT 'N',		--ME007 結案碼 N:未結案、Y:自動結案
 Remark						nvarchar(255)			NULL,						--ME008 備註
 TransferStatus				nvarchar(2)				NULL	DEFAULT 'N',		--拋轉ERP狀態 >> BAS.[Status]
 TransferDate				datetime				NULL,						--拋轉ERP時間
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_LotNumber PRIMARY KEY (LotNumberId),
 CONSTRAINT AK1_LotNumber UNIQUE (MtlItemId, LotNumberNo),
 CONSTRAINT FK1_LotNumber FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId),
)/*批號資料單頭*/

CREATE TABLE SCM.LnDetail( /*批號資料單身*/ /*ERP INVMF*/
 LnDetailId					int					identity,
 LotNumberId				int					NOT NULL,						--批號單頭ID
 TransactionDate			datetime			NOT NULL,						--MF003 異動日期
 FromErpPrefix				nvarchar(4)			NOT NULL,						--MF004 異動單別
 FromErpNo					nvarchar(11)		NOT NULL,						--MF005 異動單號
 FromSeq					nvarchar(4)			NOT NULL,						--MF006 異動序號
 InventoryId				int					NOT NULL,						--MF007 異動庫別
 TransactionType			int					NOT NULL,						--MF008 異動類別 入:1 出:-1
 DocType					nvarchar(1)			NOT NULL,						--MF009 異動別 1.入庫、2.銷貨、3.領用、4.轉撥、5.調整
 Quantity					float				NOT NULL,						--MF010 數量
 Remark						nvarchar(255)			NULL,						--MF013 備註
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_LnDetail PRIMARY KEY (LnDetailId),
 CONSTRAINT FK1_LnDetail FOREIGN KEY (LotNumberId) REFERENCES SCM.LotNumber(LotNumberId),
 CONSTRAINT FK2_LnDetail FOREIGN KEY (InventoryId) REFERENCES SCM.Inventory(InventoryId),
)/*批號資料單頭*/

CREATE TABLE SCM.LnBarcode( /*批號條碼資料*/
 LnBarcodeId				int					identity,
 LotNumberId				int					NOT NULL,						--批號單頭ID
 BarcodeNo					nvarchar(60)		NOT NULL,						--條碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_LnBarcode PRIMARY KEY (LnBarcodeId),
 CONSTRAINT AK1_LnBarcode UNIQUE (LotNumberId, BarcodeNo),
 CONSTRAINT FK1_LnBarcode FOREIGN KEY (LotNumberId) REFERENCES SCM.LotNumber(LotNumberId),
 CONSTRAINT FK2_LnBarcode FOREIGN KEY (BarcodeNo) REFERENCES MES.Barcode(BarcodeNo),
)/*批號條碼資料*/
END /*批號相關資料*/

BEGIN /*訂單*/
CREATE TABLE SCM.SaleOrder( /*訂單單頭*/ /*ERP COPTC*/
 SoId						int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 DepartmentId				int					    NULL,						--TC005 所屬部門
 SoErpPrefix				nvarchar(4)			NOT NULL,						--TC001 訂單單別 >> 參照 ERP各系統單據設定檔(CMSMQ)
 SoErpNo					nvarchar(11)		NOT NULL,						--TC002 訂單單號
 [Version]					nvarchar(4)			NOT NULL,						--TC069 版次
 SoDate						datetime			NOT NULL,						--TC003 訂單日期
 DocDate					datetime			NOT NULL,						--TC039 單據日期
 SoRemark					nvarchar(255)			NULL,						--TC015 單頭備註
 CustomerId					int					NOT NULL,						--TC004 客戶
 SalesmenId					int					    NULL,						--TC006 業務員
 CustomerAddressFirst		nvarchar(200)		NOT NULL,						--TC010 客戶聯絡地址(一)
 CustomerAddressSecond		nvarchar(200)			NULL,						--TC011 客戶聯絡地址(二)
 CustomerPurchaseOrder		nvarchar(20)			NULL,						--TC012 客戶採購單
 DepositPartial				nvarchar(1)			NOT NULL,						--TC070 訂金分批 >> BAS.[Status] N(DEF)；Y
 DepositRate				float				NOT NULL,						--TC045 訂金比率
 Currency					nvarchar(4) 		NOT NULL,						--TC008 交易幣別 >> 參照 ERP幣別匯率檔單頭(CMSMF)
 ExchangeRate				float				NOT NULL,						--TC009 匯率 >> 參照 ERP幣別匯率檔單身(CMSMG)
 TaxNo						nvarchar(3)			NOT NULL,						--TC078 稅別碼 >> 參照 ERP稅別碼設定檔(CMSNN)
 Taxation					nvarchar(1)			NOT NULL,						--TC016 課稅別 >> 由稅別碼決定(NN006) 1：內含；2：外加；3：零稅率；4：免稅；9：不計稅
 BusinessTaxRate			float				NOT NULL,						--TC041 營業稅率 >> 由稅別碼決定(NN004)
 DetailMultiTax				nvarchar(1)			NOT NULL,						--TC091 單身多稅率 >> BAS.[Status] N(DEF)；Y
 TotalQty					float				NOT NULL,						--TC031 訂單總數量
 Amount						float				NOT NULL,						--TC029 訂單總金額
 TaxAmount					float				NOT NULL,						--TC030 訂單總稅額
 ShipMethod					nvarchar(1)			NOT NULL,						--TC019 運輸方式 >> 參照 ERP運輸方式代號檔(CMSNJ)
 TradeTerm					nvarchar(1)			NOT NULL,						--TC068 交易條件 >> 參照 ERP交易條件代號檔(CMSNK)
 PaymentTerm				nvarchar(6) 		NOT NULL,						--TC014 付款條件 >> 參照 ERP付款條件檔(CMSNA) TC042 付款條件代碼
 PriceTerm					nvarchar(40) 			NULL,						--TC013 價格條件
 ConfirmStatus				nvarchar(1)			NOT NULL,						--TC027 確認碼　>> BAS.[Status] N：未確認；V：作廢；Y：已確認
 ConfirmUserId				int						NULL,						--TC040 確認者
 TransferStatus				nvarchar(2)			NOT NULL,						--拋轉ERP狀態 >> BAS.[Status]
 TransferDate				datetime				NULL,						--拋轉ERP時間
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_SaleOrder PRIMARY KEY (SoId),
 CONSTRAINT AK1_SaleOrder UNIQUE (CompanyId, SoErpPrefix, SoErpNo),
 CONSTRAINT FK1_SaleOrder FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_SaleOrder FOREIGN KEY (DepartmentId) REFERENCES BAS.Department(DepartmentId),
 CONSTRAINT FK3_SaleOrder FOREIGN KEY (CustomerId) REFERENCES SCM.Customer(CustomerId),
 CONSTRAINT FK4_SaleOrder FOREIGN KEY (SalesmenId) REFERENCES BAS.[User](UserId)
)

CREATE TABLE SCM.SoDetail( /*訂單單身*/ /*ERP COPTD*/
 SoDetailId					int					identity,
 SoId						int					NOT NULL,						--所屬單頭
 SoSequence					nvarchar(4)			NOT NULL,						--TD003	流水號
 MtlItemId					int						NULL,						--TD004	品號
 SoMtlItemName				nvarchar(120)			NULL,						--TD005	品名
 SoMtlItemSpec				nvarchar(120)			NULL,						--TD006	規格
 InventoryId				int						NULL,						--TD007	庫別
 UomId						int						NULL,						--TD010	單位(Unit of Measure) >> 參照 ERP品號換算單位檔(INVMD)
 SoQty						float				NOT NULL,						--TD008	訂單數量
 SiQty						float				NOT NULL,						--TD009	已交數量
 ProductType				nvarchar(1)			NOT	NULL	DEFAULT '1',		--TD049 類型 1.贈品量、2.備品量
 FreebieQty					float				NOT NULL,						--TD024	贈品數量
 FreebieSiQty				float				NOT NULL,						--TD025	贈品已交數量
 SpareQty					float					NULL,						--TD050	備品量
 SpareSiQty					float					NULL,						--TD051	備品已交量
 UnitPrice					float				NOT NULL,						--TD011	單價
 Amount						float				NOT NULL,						--TD012	金額
 PromiseDate				datetime			    NULL,						--TD013	預交日
 PcPromiseDate				datetime			    NULL,						--TD048	排定交貨日
 Project					nvarchar(20)			NULL,						--TD027	專案代號 >> 參考 ERP專案代號檔(CMSNB)
 CustomerMtlItemNo			nvarchar(40)			NULL,						--TD014	客戶品號
 SoDetailRemark				nvarchar(255)			NULL,						--TD020	單身備註
 SoPriceQty					float				NOT NULL,						--TD076	計價數量
 SoPriceUomId				int						NULL,						--TD077	計價單位
 TaxNo						nvarchar(3)			NOT NULL,						--TD079	稅別碼 >> 參照 ERP稅別碼設定檔(CMSNN)
 BusinessTaxRate			float				NOT NULL,						--TD070	營業稅率 >> 由稅別碼決定	
 DiscountRate				float				NOT NULL,						--TD026	折扣率
 DiscountAmount				float				NOT NULL,						--TD080	折扣金額
 QuotationErp				nvarchar(21)			NULL,						--TD017/18/19 前置單據-報價單
 ConfirmStatus				nvarchar(1)			NOT NULL,						--TD021	確認碼 >> BAS.[Status] N：未確認；V：作廢；Y：已確認
 ClosureStatus				nvarchar(1)			NOT NULL,						--TD016	結案碼 >> BAS.[Status] N：未結案；y：指定結案；Y：自動結案
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_SoDetail PRIMARY KEY (SoDetailId),
 --CONSTRAINT AK1_SoDetail UNIQUE (SoId, SoSequence),
 CONSTRAINT FK1_SoDetail FOREIGN KEY (SoId) REFERENCES SCM.SaleOrder(SoId),
 CONSTRAINT FK2_SoDetail FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId),
 CONSTRAINT FK3_SoDetail FOREIGN KEY (InventoryId) REFERENCES SCM.Inventory(InventoryId),
 CONSTRAINT FK4_SoDetail FOREIGN KEY (UomId) REFERENCES PDM.UnitOfMeasure(UomId),
 CONSTRAINT FK5_SoDetail FOREIGN KEY (SoPriceUomId) REFERENCES PDM.UnitOfMeasure(UomId)
)

CREATE TABLE SCM.SoDetailTemp( /*採購單轉訂單身暫存*/ /*ERP PURTD*/
 SoDetailTempId				int					identity,
 SoId						int					NOT NULL,						--所屬單頭
 CompanyId					int					NOT NULL,						--公司別(集團內客戶)
 SoDetailId					int						NULL,						--綁定的訂單身ID
 BaseDocType				nvarchar(10)		NOT NULL,						--來源單據類別
 BaseErpPrefix				nvarchar(4)			NOT NULL,						--TD001 來源單別
 BaseErpNo					nvarchar(12)		NOT NULL,						--TD002 來源單號
 BaseErpSNo					nvarchar(6)			NOT NULL,						--TD003 來源單號
 CustomerMtlItemId			int						NULL,						--客戶部番id
 UomNo						nvarchar(50)		NOT NULL,						--TD009 單位
 SoQty						float				NOT NULL,						--TD008 訂單數量
 UnitPrice					float				NOT NULL,						--TD010 單價
 Amount						float				NOT NULL,						--TD011	金額
 BaseMtlItemNo				nvarchar(120)		NOT NULL,						--TD004 來原品號
 BaseMtlItemName			nvarchar(120)		NOT NULL,						--TD005 來源品名
 BaseMtlItemSpec			nvarchar(120)		NOT NULL,						--TD006 來源規格
 SortNumber					int					NOT NULL,						--排序		
 [Status]					nvarchar(1)			NOT NULL,						--狀態 >> BAS.[Status] A：啟用；S：停用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_SoDetailTemp PRIMARY KEY (SoDetailTempId),
 CONSTRAINT FK1_SoDetailTemp FOREIGN KEY (SoId) REFERENCES SCM.SaleOrder(SoId),
 CONSTRAINT FK2_SoDetailTemp FOREIGN KEY (CustomerMtlItemId) REFERENCES PDM.CustomerMtlItem(CustomerMtlItemId),
 CONSTRAINT FK3_SoDetailTemp FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK4_SoDetailTemp FOREIGN KEY (SoDetailId) REFERENCES SCM.SoDetail(SoDetailId),
 CONSTRAINT AK1_SoDetailTemp UNIQUE (SoId, SortNumber),
 CONSTRAINT AK2_SoDetailTemp UNIQUE (SoId, BaseErpPrefix, BaseErpNo, BaseErpSNo)
)

CREATE TABLE SCM.SoModification( /*訂單變更單頭*/ /*ERP COPTE*/
 SomId						int					identity,
 SoId						int					NOT NULL,						--所屬訂單
 [Version]					nvarchar(4)			NOT NULL,						--TE003 變更版次
 DepartmentId				int					    NULL,						--TE008 所屬部門
 DocDate					datetime			NOT NULL,						--TE038 單據日期
 SoRemark					nvarchar(255)			NULL,						--TE050 單頭備註
 SalesmenId					int					    NULL,						--TE009 業務員
 CustomerAddressFirst		nvarchar(200)		NOT NULL,						--TE013 客戶聯絡地址(一)
 CustomerAddressSecond		nvarchar(200)			NULL,						--TE014 客戶聯絡地址(二)
 CustomerPurchaseOrder		nvarchar(20)			NULL,						--TE015 客戶採購單
 DepositPartial				nvarchar(1)			NOT NULL,						--TE079 訂金分批 >> BAS.[Status] N(DEF)；Y
 DepositRate				float				NOT NULL,						--TE042 訂金比率
 Currency					nvarchar(4) 		NOT NULL,						--TE011 交易幣別 >> 參照 ERP幣別匯率檔單頭(CMSMF)
 ExchangeRate				float				NOT NULL,						--TE012 匯率 >> 參照 ERP幣別匯率檔單身(CMSMG)
 TaxNo						nvarchar(3)			NOT NULL,						--TE069 稅別碼 >> 參照 ERP稅別碼設定檔(CMSNN)
 Taxation					nvarchar(1)			NOT NULL,						--TE018 課稅別 >> 由稅別碼決定(NN006) 1：內含；2：外加；3：零稅率；4：免稅；9：不計稅
 BusinessTaxRate			float				NOT NULL,						--TE040 營業稅率 >> 由稅別碼決定(NN004)
 DetailMultiTax				nvarchar(1)			NOT NULL,						--TE080 單身多稅率 >> BAS.[Status] N(DEF)；Y
 ShipMethod					nvarchar(1)			NOT NULL,						--TE021 運輸方式 >> 參照 ERP運輸方式代號檔(CMSNJ)
 TradeTerm					nvarchar(1)			NOT NULL,						--TE061 交易條件 >> 參照 ERP交易條件代號檔(CMSNK)
 PaymentTerm				nvarchar(6) 		NOT NULL,						--TE017 付款條件 >> 參照 ERP付款條件檔(CMSNA) TE041 付款條件代碼
 PriceTerm					nvarchar(40) 			NULL,						--TE016 價格條件
 OriVersion					nvarchar(4)			NOT NULL,						--TE103 原版次
 OriDepartmentId			int					    NULL,						--TE108 原所屬部門
 OriSoRemark				nvarchar(255)			NULL,						--TE144 原單頭備註
 OriSalesmenId				int					    NULL,						--TE109 原業務員
 OriCustomerAddressFirst	nvarchar(200)		NOT NULL,						--TE113 原客戶聯絡地址(一)
 OriCustomerAddressSecond	nvarchar(200)			NULL,						--TE114 原客戶聯絡地址(二)
 OriCustomerPurchaseOrder	nvarchar(20)			NULL,						--TE115 原客戶採購單
 OriDepositPartial			nvarchar(1)			NOT NULL,						--TE179 原訂金分批 >> BAS.[Status] N(DEF)；Y
 OriDepositRate				float				NOT NULL,						--TE138 原訂金比率
 OriCurrency				nvarchar(4) 		NOT NULL,						--TE111 原交易幣別 >> 參照 ERP幣別匯率檔單頭(CMSMF)
 OriExchangeRate			float				NOT NULL,						--TE112 原匯率 >> 參照 ERP幣別匯率檔單身(CMSMG)
 OriTaxNo					nvarchar(3)			NOT NULL,						--TE169 原稅別碼 >> 參照 ERP稅別碼設定檔(CMSNN)
 OriTaxation				nvarchar(1)			NOT NULL,						--TE118 原課稅別 >> 由稅別碼決定(NN006) 1：內含；2：外加；3：零稅率；4：免稅；9：不計稅
 OriBusinessTaxRate			float				NOT NULL,						--TE136 原營業稅率 >> 由稅別碼決定(NN004)
 OriDetailMultiTax			nvarchar(1)			NOT NULL,						--TE180 原單身多稅率 >> BAS.[Status] N(DEF)；Y
 OriShipMethod				nvarchar(1)			NOT NULL,						--TE121 原運輸方式 >> 參照 ERP運輸方式代號檔(CMSNJ)
 OriTradeTerm				nvarchar(1)			NOT NULL,						--TE151 原交易條件 >> 參照 ERP交易條件代號檔(CMSNK)
 OriPaymentTerm				nvarchar(6) 		NOT NULL,						--TE117 原付款條件 >> 參照 ERP付款條件檔(CMSNA) TE137 付款條件代碼
 OriPriceTerm				nvarchar(40) 			NULL,						--TE116 原價格條件
 ClosureStatus				nvarchar(1)			NOT NULL,						--TE005 整張結案　>> BAS.[Status]  N：未結案；y：指定結案；Y：自動結案
 ModiReason					nvarchar(255)			NULL,						--TE006 變更原因
 ConfirmStatus				nvarchar(1)			NOT NULL,						--TE029 確認碼　>> BAS.[Status] N：未確認；V：作廢；Y：已確認
 ConfirmUserId				int						NULL,						--TE039 確認者
 TransferStatus				nvarchar(2)			NOT NULL,						--拋轉ERP狀態 >> BAS.[Status]
 TransferDate				datetime				NULL,						--拋轉ERP時間
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_SoModification PRIMARY KEY (SomId),
 CONSTRAINT AK1_SoModification UNIQUE (SoId, [Version]),
 CONSTRAINT FK1_SoModification FOREIGN KEY (SoId) REFERENCES SCM.SaleOrder(SoId),
 CONSTRAINT FK2_SoModification FOREIGN KEY (DepartmentId) REFERENCES BAS.Department(DepartmentId),
 CONSTRAINT FK3_SoModification FOREIGN KEY (SalesmenId) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK4_SoModification FOREIGN KEY (OriDepartmentId) REFERENCES BAS.Department(DepartmentId),
 CONSTRAINT FK5_SoModification FOREIGN KEY (OriSalesmenId) REFERENCES BAS.[User](UserId)
)

CREATE TABLE SCM.SomDetail( /*訂單變更單身*/ /*ERP COPTF*/
 SomDetailId				int					identity,
 SomId						int					NOT NULL,						--所屬單頭
 SoDetailId					int					NOT NULL,						--所屬訂單單身
 [Version]					nvarchar(4)			NOT NULL,						--TF003 變更版次
 SoSequence					nvarchar(4)			NOT NULL,						--TF004	流水號
 MtlItemId					int					NOT NULL,						--TF005	品號
 SoMtlItemName				nvarchar(120)		NOT NULL,						--TF006	品名
 SoMtlItemSpec				nvarchar(120)		NOT NULL,						--TF007	規格
 InventoryId				int					NOT NULL,						--TF008	庫別
 UomId						int					NOT NULL,						--TF010	單位(Unit of Measure) >> 參照 ERP品號換算單位檔(INVMD)
 SoQty						int					NOT NULL,						--TF009	訂單數量
 FreebieQty					int					NOT NULL,						--TF020	贈品數量
 UnitPrice					float				NOT NULL,						--TF013	單價
 Amount						float				NOT NULL,						--TF014	金額
 PromiseDate				datetime			    NULL,						--TF015	預交日
 Project					nvarchar(20)			NULL,						--TF022	專案代號 >> 參考 ERP專案代號檔(CMSNB)
 CustomerMtlItemNo			nvarchar(40)			NULL,						--TF016	客戶品號
 SoDetailRemark				nvarchar(255)			NULL,						--TF039	單身備註
 SoPriceQty					int					NOT NULL,						--TF074	計價數量
 SoPriceUomId				int					NOT NULL,						--TF075	計價單位
 TaxNo						nvarchar(3)			NOT NULL,						--TF076	稅別碼 >> 參照 ERP稅別碼設定檔(CMSNN)
 BusinessTaxRate			float				NOT NULL,						--TF062	營業稅率 >> 由稅別碼決定
 OriSoSequence				nvarchar(4)			    NULL,						--TF004	原流水號
 OriMtlItemId				int					    NULL,						--TF005	原品號
 OriSoMtlItemName			nvarchar(120)		    NULL,						--TF006	原品名
 OriSoMtlItemSpec			nvarchar(120)		    NULL,						--TF007	原規格
 OriInventoryId				int					    NULL,						--TF008	原庫別
 OriUomId					int					    NULL,						--TF010	原單位(Unit of Measure) >> 參照 ERP品號換算單位檔(INVMD)
 OriSoQty					int					    NULL,						--TF009	原訂單數量
 OriFreebieQty				int					    NULL,						--TF020	原贈品數量
 OriUnitPrice				float				    NULL,						--TF013	原單價
 OriAmount					float				    NULL,						--TF014	原金額
 OriPromiseDate				datetime			    NULL,						--TF015	原預交日
 OriProject					nvarchar(20)			NULL,						--TF022	原專案代號 >> 參考 ERP專案代號檔(CMSNB)
 OriCustomerMtlItemNo		nvarchar(40)			NULL,						--TF016	原客戶品號
 OriSoDetailRemark			nvarchar(255)			NULL,						--TF039	原單身備註
 OriSoPriceQty				int					    NULL,						--TF074	原計價數量
 OriSoPriceUomId			int					    NULL,						--TF075	原計價單位
 OriTaxNo					nvarchar(3)			    NULL,						--TF076	原稅別碼 >> 參照 ERP稅別碼設定檔(CMSNN)
 OriBusinessTaxRate			float				    NULL,						--TF062	原營業稅率 >> 由稅別碼決定
 ModiReason					nvarchar(255)			NULL,						--TF018 變更原因
 ConfirmStatus				nvarchar(1)			NOT NULL,						--TF019	確認碼 >> BAS.[Status] N：未確認；V：作廢；Y：已確認
 ClosureStatus				nvarchar(1)			NOT NULL,						--TF017	結案碼 >> BAS.[Status] N：未結案；y：指定結案；Y：自動結案
 OriClosureStatus			nvarchar(1)			NOT NULL,						--TF117	原結案碼 >> BAS.[Status] N：未結案；y：指定結案；Y：自動結案
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_SomDetail PRIMARY KEY (SoDetailId),
 CONSTRAINT AK1_SomDetail UNIQUE (SomId, SoSequence),
 CONSTRAINT FK1_SomDetail FOREIGN KEY (SomId) REFERENCES SCM.SoModification(SomId),
 CONSTRAINT FK2_SomDetail FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId),
 CONSTRAINT FK3_SomDetail FOREIGN KEY (InventoryId) REFERENCES SCM.Inventory(InventoryId),
 CONSTRAINT FK4_SomDetail FOREIGN KEY (UomId) REFERENCES PDM.UnitOfMeasure(UomId),
 CONSTRAINT FK5_SomDetail FOREIGN KEY (SoPriceUomId) REFERENCES PDM.UnitOfMeasure(UomId)
)

CREATE TABLE SCM.SoBpmInfo( /*訂單拋轉BPM相關資料*/
 SoBpmInfoId				int					identity,
 SoId						int					NOT NULL,						--訂單ID
 BpmNo						nvarchar(100)			NULL,						--BPM單號
 BpmTransferStatus			nvarchar(2)			NOT NULL,						--訂單BPM狀態碼 >> BAS.[Status] P：BPM簽核流程；Y：單據核準；R：拋轉BPM失敗；E：單據修改；F：確認ERP失敗
 BpmTransferUserId			int						NULL,						--拋轉BPM人員
 BpmTransferDate			datetime				NULL,						--拋轉BPM時間
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_SoBpmInfo PRIMARY KEY (SoBpmInfoId),
 CONSTRAINT AK1_SoBpmInfo UNIQUE (SoId),
 CONSTRAINT FK1_SoBpmInfo FOREIGN KEY (SoId) REFERENCES SCM.SaleOrder(SoId),
 CONSTRAINT FK2_SoBpmInfo FOREIGN KEY (BpmTransferUserId) REFERENCES BAS.[User](UserId)
)

CREATE TABLE SCM.SoBpmLog( /*訂單拋轉BPM LOG*/
 SoBpmLogId					int					identity,
 SoId						int					NOT NULL,						--所屬訂單ID
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
 CONSTRAINT PK_SoBpmLog PRIMARY KEY (SoBpmLogId),
 CONSTRAINT FK1_SoBpmLog FOREIGN KEY (SoId) REFERENCES SCM.SaleOrder(SoId)
)

CREATE TABLE SCM.PMDPage ( /* PMD頁面 */ 
 PMDPageId					  INT					identity,
 PageName					  VARCHAR(50)			NOT NULL,						--頁面名
 PageNo						  VARCHAR(50)			NOT NULL,						--頁面(@(itemName))
 CreateDate                   datetime				NOT NULL    DEFAULT GETDATE(),  --新增日期
 LastModifiedDate             datetime              NULL,							--修改日期
 CreateBy                     int                   NOT NULL,                       --新增人員
 LastModifiedBy               int                   NULL,							--修改人員
 CONSTRAINT PK_PMDPage PRIMARY KEY(PMDPageId),
);

CREATE TABLE SCM.PageLock ( /* 使用者頁面編輯鎖定紀錄 */
 PageLockId INT                identity,
 PageName VARCHAR(50)        NOT NULL,                --頁面名(@(itemName))
 IsLocked INT                NOT NULL DEFAULT 0,        --是否鎖定 0:未鎖 1:鎖定
 LockedBy INT                NULL,                    --當前使用者
 LockTime DATETIME            NULL,                    --鎖定時間
 CreateDate                    datetime            NOT NULL    DEFAULT GETDATE(),    --新增日期
 LastModifiedDate            datetime                NULL,                        --修改日期
 CreateBy                    int                    NOT NULL,                        --新增人員
 LastModifiedBy                int                        NULL,                        --修改人員
 CONSTRAINT PK_PageLock PRIMARY KEY(PageLockId),
);

CREATE TABLE SCM.PMDOrder( /* PMD報表(舜宇) */
 PMDOrderId						int                    identity,
 MtlItemId						int                      NULL,                        --品號ID
 SoQty							int                      NULL,                        --訂單數
 SpareQty						int                      NULL,                        --備品數
 OrderDate						nvarchar(MAX)            NULL,                        --接單日
 DrawingAttachmentDate			nvarchar(MAX)            NULL,                        --下圖日
 CurrentProcess					nvarchar(MAX)            NULL,                        --當前進度
 CustomerDueDate				nvarchar(MAX)            NULL,                        --客戶交期
 ConfirmedDueDate				nvarchar(MAX)            NULL,                        --回復交期
 DelayRemark					nvarchar(MAX)            NULL,                        --延誤交期時間
 MaterialNumber                 nvarchar(MAX)            NULL,                        --物料編碼
 Remark                         nvarchar(MAX)            NULL,                        --備註
 TrackingNumber                 nvarchar(MAX)            NULL,                        --物流單號
 ShipFrom						nvarchar(MAX)            NULL,                        --發貨地
 SemiDelivery					nvarchar(MAX)            NULL,                        --半成品發貨(ZY-JC)O/X
 TypeId                         int                      NULL,                        --資料所屬頁面
 PMDItem01						nvarchar(MAX)            NULL,                        --保留欄位
 PMDItem02						nvarchar(MAX)            NULL,                        --保留欄位
 PMDItem03						nvarchar(MAX)            NULL,                        --保留欄位
 PMDItem04						nvarchar(MAX)            NULL,                        --保留欄位
 PlannedShipmentDate			nvarchar(MAX)            NULL,                        --預計出貨日
 ActualShipmentDate             nvarchar(MAX)            NULL,                        --實際出貨日
 PlannedShipmentQty				int                      NULL,                        --預計出貨數量
 ActualShipmentQty				int                      NULL,                        --實際出貨數量
 MoldFrameCtrlNo				nvarchar(MAX)            NULL,                        --模架管理號
 ReturnQty						int                      NULL,                        --退貨數量
 ReturnIssue					nvarchar(MAX)            NULL,                        --退貨問題點
 PrevReturnQty					int                      NULL,                        --前期退貨數量
 PrevShortageQty				int                      NULL,                        --前期欠貨數量
 TotalReplenishmentQty			int                      NULL,                        --共需補貨數量    
 CAV							nvarchar(MAX)            NULL,                        --
 ReturnDate						nvarchar(MAX)            NULL,                        --
 CustomerMtlItemNo				nvarchar(MAX)            NULL,                        --
 MtlItemName					nvarchar(MAX)            NULL,                        --
 SendType						int                      NULL,                        --
 CustomerId						int                      NULL,                        --    
 SoDetailId						int                      NULL,                        --   
 PMDOrderTempId					int                      NULL,                        --    
 CreateDate						datetime				 NOT NULL    DEFAULT GETDATE(),    --新增日期
 LastModifiedDate				datetime                 NULL,                        --修改日期
 CreateBy						int                      NOT NULL,                        --新增人員
 LastModifiedBy					int						 NULL,                        --修改人員
 CONSTRAINT PK_PMDOrder PRIMARY KEY(PMDOrderId)
)

CREATE TABLE SCM.PMDOrderTemp( /* PMD報表(舜宇) */
 PMDOrderTempId						int                    identity,
 MtlItemId						int                      NULL,                        --品號ID
 SoQty							int                      NULL,                        --訂單數
 SpareQty						int                      NULL,                        --備品數
 OrderDate						nvarchar(MAX)            NULL,                        --接單日
 DrawingAttachmentDate			nvarchar(MAX)            NULL,                        --下圖日
 CurrentProcess					nvarchar(MAX)            NULL,                        --當前進度
 CustomerDueDate				nvarchar(MAX)            NULL,                        --客戶交期
 ConfirmedDueDate				nvarchar(MAX)            NULL,                        --回復交期
 DelayRemark					nvarchar(MAX)            NULL,                        --延誤交期時間
 MaterialNumber                 nvarchar(MAX)            NULL,                        --物料編碼
 Remark                         nvarchar(MAX)            NULL,                        --備註
 TrackingNumber                 nvarchar(MAX)            NULL,                        --物流單號
 ShipFrom						nvarchar(MAX)            NULL,                        --發貨地
 SemiDelivery					nvarchar(MAX)            NULL,                        --半成品發貨(ZY-JC)O/X
 TypeId                         int                      NULL,                        --資料所屬頁面
 PMDItem01						nvarchar(MAX)            NULL,                        --保留欄位
 PMDItem02						nvarchar(MAX)            NULL,                        --保留欄位
 PMDItem03						nvarchar(MAX)            NULL,                        --保留欄位
 PMDItem04						nvarchar(MAX)            NULL,                        --保留欄位
 PlannedShipmentDate			nvarchar(MAX)            NULL,                        --預計出貨日
 ActualShipmentDate             nvarchar(MAX)            NULL,                        --實際出貨日
 PlannedShipmentQty				int                      NULL,                        --預計出貨數量
 ActualShipmentQty				int                      NULL,                        --實際出貨數量
 MoldFrameCtrlNo				nvarchar(MAX)            NULL,                        --模架管理號
 ReturnQty						int                      NULL,                        --退貨數量
 ReturnIssue					nvarchar(MAX)            NULL,                        --退貨問題點
 PrevReturnQty					int                      NULL,                        --前期退貨數量
 PrevShortageQty				int                      NULL,                        --前期欠貨數量
 TotalReplenishmentQty			int                      NULL,                        --共需補貨數量    
 CAV							nvarchar(MAX)            NULL,                        --
 ReturnDate						nvarchar(MAX)            NULL,                        --
 CustomerMtlItemNo				nvarchar(MAX)            NULL,                        --
 MtlItemName					nvarchar(MAX)            NULL,                        --
  SendType						int                      NULL,                        --
 CustomerId						int                      NULL,                        --    
 SoDetailId						int                      NULL,                        --    
 CreateDate						datetime				 NOT NULL    DEFAULT GETDATE(),    --新增日期
 LastModifiedDate				datetime                 NULL,                        --修改日期
 CreateBy						int                      NOT NULL,                        --新增人員
 LastModifiedBy					int						 NULL,                        --修改人員
 CONSTRAINT PK_PMDOrderTemp PRIMARY KEY(PMDOrderTempId)
)

CREATE TABLE SCM.CustomerMail ( /* 客戶信箱 */ 
 CustomerMailId				  INT					identity,
 CustomerId					  INT				    NULL,							--客戶ID
 CustomerNo					  VARCHAR(50)		    NULL,							--客戶No
 CustomerName				  VARCHAR(50)		    NULL,							--客戶Name
 SalesName					  VARCHAR(50)			NULL,							--業務名
 SalesMail					  VARCHAR(50)			NULL,							--業務信箱
 SalesMailCc				  VARCHAR(50)			NULL,							--業務信箱CC
 SalesMailBcc				  VARCHAR(50)			NULL,							--業務信箱BCC
 SalesTax					  VARCHAR(50)			NULL,							--業務電話
 ServerId					  INT				    NULL,							--
 MailFrom					  VARCHAR(50)			NULL,							--業務電話
 CreateDate                   datetime				NOT NULL    DEFAULT GETDATE(),  --新增日期
 LastModifiedDate             datetime              NULL,							--修改日期
 CreateBy                     int                   NOT NULL,                       --新增人員
 LastModifiedBy               int                   NULL,							--修改人員
 CONSTRAINT PK_CustomerMail PRIMARY KEY(CustomerMailId),
)

CREATE TABLE SCM.PMDOrderLog ( /* PMD編輯Log */ 
 PMDOrderLogId				  INT					identity,
 UserId						  INT				    NULL,							--編輯者
 EditTime					  datetime		        NULL,							--編輯時間
 EditItem					  INT				    NULL,							--編輯資料ID
 EditTable					  INT					NULL,							--編輯表格名
 CreateDate                   datetime				NOT NULL    DEFAULT GETDATE(),  --新增日期
 LastModifiedDate             datetime              NULL,							--修改日期
 CreateBy                     int                   NOT NULL,                       --新增人員
 LastModifiedBy               int                   NULL,							--修改人員
 CONSTRAINT PK_PMDOrderLog PRIMARY KEY(PMDOrderLogId)
)
END /*訂單*/

BEGIN /*出貨*/
CREATE TABLE SCM.DeliveryOrder( /*出貨單頭*/
 DoId						int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 DepartmentId				int						NULL,						--TF007 所屬部門
 UserId						int					NOT NULL,						--TF008 所屬人員
 DoErpPrefix				nvarchar(4)			NOT NULL,						--TF001 出貨單別 >> 參照 ERP各系統單據設定檔(CMSMQ)
 DoErpNo					nvarchar(11)		NOT NULL,						--TF002 出貨單號
 DoDate						datetime			NOT NULL,						--TF003 出貨日期 >> 自動帶入TF041 預計轉銷日
 DocDate					datetime			NOT NULL,						--TF024 單據日期
 CustomerId					int					NOT NULL,						--TF005 所屬客戶
 DcId						int						NULL,						--送貨客戶
 WayBill					nvarchar(50)			NULL,						--貨運單號(指由承運人簽發的一種與託運的貨物有關的細節文件)
 Traffic					nvarchar(50)			NULL,						--快遞單號(倉庫使用)
 ShipMethod					nvarchar(1)				NULL,						--TF031 運輸方式 >> 參照 ERP運輸方式代號檔(CMSNJ)
 DoAddressFirst				nvarchar(255)		NOT NULL,						--TF016 送貨地址(一)
 DoAddressSecond			nvarchar(255)			NULL,						--TF017 送貨地址(二)
 TotalQty					float					NULL,						--TF022 暫出總數量
 Amount						float					NULL,						--TF023 暫出總金額
 TaxAmount					float					NULL,						--TF027 暫出總稅額
 DoRemark					nvarchar(255)			NULL,						--TF014 單頭備註
 WareHouseDoRemark			nvarchar(255)			NULL,						--TF018 倉庫備註 --(其他備註)
 MeasureMailStatus			nvarchar(2)				NULL,						--量測郵件狀態(出貨確認之郵件) >> BAS.[Status] N:未執行；S：已執行但資料不足；Y:已執行
 ConfirmStatus				nvarchar(1)			NOT NULL,						--TF020 確認碼　>> BAS.[Status] N：未確認；V：作廢；Y：已確認
 ConfirmUserId				int						NULL,						--TF025 確認者
 TransferStatus				nvarchar(2)			NOT NULL,						--拋轉ERP狀態 >> BAS.[Status]
 TransferDate				datetime				NULL,						--拋轉ERP時間
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼 >> BAS.[Status] N：未揀貨(Not picked yet)；P：揀貨中(Picking)；R：可出貨(Ready to ship)；S：已出貨(Shipped)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DeliveryOrder PRIMARY KEY (DoId),
 CONSTRAINT AK1_DeliveryOrder UNIQUE (CompanyId, DoErpPrefix, DoErpNo),
 CONSTRAINT FK1_DeliveryOrder FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_DeliveryOrder FOREIGN KEY (DepartmentId) REFERENCES BAS.Department(DepartmentId),
 CONSTRAINT FK3_DeliveryOrder FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK4_DeliveryOrder FOREIGN KEY (DcId) REFERENCES SCM.DeliveryCustomer(DcId),
 CONSTRAINT FK5_DeliveryOrder FOREIGN KEY (ConfirmUserId) REFERENCES BAS.[User](UserId)
)

CREATE TABLE SCM.DoDetail( /*出貨單身*/
 DoDetailId					int					identity,
 DoId						int					NOT NULL,						--所屬單頭
 SoDetailId					int					NOT NULL,						--TG014/15/16 訂單單身(單別/單號/流水號)
 DoSequence					nvarchar(4)			NOT NULL,						--TG003 流水號
 TransInInventoryId			int						NULL,						--TG008 轉入庫別
 DoQty						int					NOT NULL,						--TG009 出貨數量
 FreebieQty					float					NULL,						--贈品量
 SpareQty					float					NULL,						--備品量
 UnitPrice					float					NULL,						--TG012 單價
 Amount						float					NULL,						--TG013 金額
 DoDetailRemark				nvarchar(255)			NULL,						--TG019 單身備註
 WareHouseDoDetailRemark	nvarchar(255)			NULL,						--倉庫備註
 PcDoDetailRemark			nvarchar(255)			NULL,						--生管備註
 DeliveryProcess			nvarchar(2)				NULL,						--出貨工程 >> BAS.[Type] 目前內容13個
 DeliveryRoutine			nvarchar(2)				NULL,						--出貨工序 >> BAS.[Type] 目前內容10個
 DeliveryDocument			nvarchar(100)			NULL,						--出貨文件名稱
 OrderSituation				nvarchar(50)			NULL,						--訂單狀況 >> BAS.[Type] 目前內容5個
 DeliveryMethod				nvarchar(50)			NULL,						--寄送方式 >> BAS.[Type] 目前內容4個
 ConfirmStatus				nvarchar(1)			NOT NULL,						--TG022 確認碼 >> BAS.[Status] N：未確認；V：作廢；Y：已確認
 ClosureStatus				nvarchar(1)			NOT NULL,						--TG024 結案碼 >> BAS.[Status] N：未結案；y：指定結案；Y：自動結案
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DoDetail PRIMARY KEY (DoDetailId),
 CONSTRAINT AK1_DoDetail UNIQUE (DoId, DoSequence),
 CONSTRAINT FK1_DoDetail FOREIGN KEY (DoId) REFERENCES SCM.DeliveryOrder(DoId),
 CONSTRAINT FK2_DoDetail FOREIGN KEY (SoDetailId) REFERENCES SCM.SoDetail(SoDetailId),
 CONSTRAINT FK3_DoDetail FOREIGN KEY (TransInInventoryId) REFERENCES SCM.Inventory(InventoryId)
)

CREATE TABLE SCM.DeliveryDateLog( /*排定交貨日歷史紀錄*/
 DeliveryDateLogId			int					identity,
 ParentLogId				int						NULL,						--父層歷史紀錄
 SoDetailId					int					NOT NULL,						--所屬訂單單身
 DoDetailId					int						NULL,						--所屬出貨單身
 PcPromiseDate				datetime			NOT NULL,						--排定交貨日
 DoQty						int						NULL,						--出貨數量
 OriDoQty					int						NULL,						--原出貨數量
 DoDate						datetime				NULL,						--出貨日
 OriDoDate					datetime				NULL,						--原出貨日
 DepartmentId				int						NULL,						--責任部門
 SupervisorId				int						NULL,						--責任主管
 CauseType					int						NULL,						--原因類別
 CauseDescription			nvarchar(255)			NULL,						--原因說明
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DeliveryDateLog PRIMARY KEY (DeliveryDateLogId),
 CONSTRAINT FK1_DeliveryDateLog FOREIGN KEY (ParentLogId) REFERENCES SCM.DeliveryDateLog(DeliveryDateLogId),
 CONSTRAINT FK2_DeliveryDateLog FOREIGN KEY (SoDetailId) REFERENCES SCM.SoDetail(SoDetailId),
 CONSTRAINT FK3_DeliveryDateLog FOREIGN KEY (DepartmentId) REFERENCES BAS.Department(DepartmentId),
 CONSTRAINT FK4_DeliveryDateLog FOREIGN KEY (SupervisorId) REFERENCES BAS.[User](UserId)
)

CREATE TABLE SCM.DeliveryOrderLog( /*出貨日期更改歷史紀錄*/
 LogId						int					identity,
 DoId						int					NOT NULL,						--所屬出貨單ID
 DoDate						datetime			NOT NULL,						--更改後的出貨日期
 OriDoDate					datetime			NOT NULL,						--更改前的出貨日期
 DepartmentId				int						NULL,						--責任部門
 SupervisorId				int						NULL,						--責任主管
 CauseType					int						NULL,						--原因類別
 CauseDescription			nvarchar(255)			NULL,						--原因說明
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DeliveryOrderLog PRIMARY KEY (LogId),
 CONSTRAINT FK1_DeliveryOrderLog FOREIGN KEY (DoId) REFERENCES SCM.DeliveryOrder(DoId),
 CONSTRAINT FK2_DeliveryOrderLog FOREIGN KEY (DepartmentId) REFERENCES BAS.Department(DepartmentId),
 CONSTRAINT FK3_DeliveryOrderLog FOREIGN KEY (SupervisorId) REFERENCES BAS.[User](UserId)
)

CREATE TABLE SCM.PickingCarton( /*揀貨箱號*/
 PickingCartonId			int					identity,
 PackingId					int					NOT NULL,						--所屬包材
 CartonName					nvarchar(100)		NOT NULL,						--箱號名稱
 CartonBarcode				nvarchar(50)		NOT NULL,						--箱號條碼
 CartonRemark				nvarchar(100)			NULL,						--箱號備註
 TotalWeight				float				NOT NULL,						--總重
 UomId						int						NULL,						--單位
 PrintStatus				nvarchar(2)			NOT NULL,						--列印狀態碼 >> BAS.[Status]
 PrintCount					int					NOT NULL,						--列印次數
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_PickingCarton PRIMARY KEY (PickingCartonId),
 CONSTRAINT AK1_PickingCarton UNIQUE (CartonBarcode),
 CONSTRAINT FK1_PickingCarton FOREIGN KEY (PackingId) REFERENCES SCM.Packing(PackingId),
 CONSTRAINT FK2_PickingCarton FOREIGN KEY (UomId) REFERENCES PDM.UnitOfMeasure(UomId)
)

CREATE TABLE SCM.PickingItem( /*揀貨物件*/
 PickingItemId				int					identity,
 PickingCartonId			int						NULL,						--所屬箱號
 SoDetailId					int					NOT NULL,						--所屬物件
 DoId						int						NULL,						--所屬出貨單
 MoId						int						NULL,						--所屬製令
 BarcodeId					int						NULL,						--所屬條碼
 LotNumber					nvarchar(100)			NULL,						--批號
 ItemType					nvarchar(1)				NULL	DEFAULT '1',		--物件類型 >> BAS.[Type] 1:正常品；2:贈品；3:備品
 ItemQty					int						NULL,						--物件數量
 ItemStatus					nvarchar(2)			NOT NULL,						--物件狀態 >> BAS.[Status] B：條碼(Barcode)；N：無條碼(No Barcode)；S：替代品(Substitute)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_PickingItem PRIMARY KEY (PickingItemId),
 CONSTRAINT AK1_PickingItem UNIQUE (PickingCartonId, SoDetailId, DoId, BarcodeId, ItemType, ItemStatus),
 CONSTRAINT FK1_PickingItem FOREIGN KEY (PickingCartonId) REFERENCES SCM.PickingCarton(PickingCartonId),
 CONSTRAINT FK2_PickingItem FOREIGN KEY (SoDetailId) REFERENCES SCM.SoDetail(SoDetailId),
 CONSTRAINT FK3_PickingItem FOREIGN KEY (DoId) REFERENCES SCM.DeliveryOrder(DoId),
 CONSTRAINT FK4_PickingItem FOREIGN KEY (MoId) REFERENCES MES.ManufactureOrder(MoId),
 CONSTRAINT FK5_PickingItem FOREIGN KEY (BarcodeId) REFERENCES MES.Barcode(BarcodeId)
)

CREATE TABLE SCM.PickingLot( /*揀貨物件無條碼儲存批號資訊*/
 PickingLotId				int					identity,
 PickingItemId				int						NULL,						--所屬箱號
 LotNumber					nvarchar(100)			NULL,						--批號
 ItemQty					int						NULL,						--物件數量
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_PickingLot PRIMARY KEY (PickingLotId),
 CONSTRAINT FK1_PickingLot FOREIGN KEY (PickingItemId) REFERENCES SCM.PickingItem(PickingItemId),
)

CREATE TABLE SCM.WipLinkLog(
 WipLinkLogId				int					identity,
 SoErpPrefix				nvarchar(4)			NOT NULL,						--訂單單別
 SoErpNo					nvarchar(11)		NOT NULL,						--訂單單號
 SoSequence					nvarchar(4)			NOT NULL,						--流水號
 WoErpPrefix				nvarchar(4)			NOT NULL,						--訂單單別
 WoErpNo					nvarchar(11)		NOT NULL,						--訂單單號
 BindStatus					nvarchar(2)			NOT NULL,						--綁定動作 >> BAS.[Status] B：綁定；U：解除綁定
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 CreateBy					int					NOT NULL,						--新增人員
 CONSTRAINT PK_WipLinkLog PRIMARY KEY (WipLinkLogId)
)

CREATE TABLE SCM.DeliveryCustomerReport( /*客戶出貨報表 參考用表*/
 DcrId						int					identity,				
 CustomerId					int					NOT NULL,						--客戶ID，串 DeliveryCustomer DcId
 ProductType				nvarchar(6)			NOT NULL,						--產品類別，串 MtlItem TypeThree
 FunctionName				nvarchar(150)		NOT NULL,						--出貨報表API名稱
 Encoding					nvarchar(150)		NOT NULL,						--編碼
 EmptyReportName			nvarchar(150)		NOT NULL,						--表格空白檔名
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DeliveryCustomerReport PRIMARY KEY (DcrId),
 CONSTRAINT AK1_DeliveryCustomerReport UNIQUE (CustomerId, ProductType, FunctionName, EmptyReportName),
 CONSTRAINT FK1_DeliveryCustomerReport FOREIGN KEY (CustomerId) REFERENCES SCM.DeliveryCustomer(DcId)
)/*客戶出貨報表 參考用表*/

CREATE TABLE SCM.InventoryTempHold(
 TempHoldId					int					identity,                       
 MtlItemNo					nvarchar(50)		NOT NULL,                       --料號
 HoldQty					decimal(18,2)		NOT NULL,                       --暫扣數量
 DoId						int					NOT NULL,                       --出貨單號
 SoDetailId					int					NOT NULL,
 BarCodeNo					nvarchar(50)			NULL,                       --條碼號
 InputStatus				nvarchar(1)         NOT NULL,                       --輸入狀態 >> BAS.[Status] B：條碼；N：無條碼；S：替代品
 IsReleased					bit                 NOT NULL DEFAULT 0,             --是否已釋放
 ReleaseDate				datetime                NULL,                       --釋放日期
 CreateDate					datetime            NOT NULL    DEFAULT GETDATE(),  --新增日期
 LastModifiedDate			datetime				NULL,                       --修改日期
 CreateBy					int					NOT NULL,                       --新增人員
 LastModifiedBy				int						NULL,                       --修改人員
 CONSTRAINT PK_InventoryTempHold PRIMARY KEY (TempHoldId)
)
END /*出貨*/

BEGIN /*請購*/
CREATE TABLE SCM.PurchaseRequisition( /*請購單頭*/ /*ERP PURTA*/
 PrId						int					identity,
 PrNo						nvarchar(100)		NOT NULL,						--請購單編碼
 CompanyId					int					NOT NULL,						--所屬公司
 DepartmentId				int					NOT NULL,						--TA004 所屬部門
 UserId						int					NOT NULL,						--TA012 所屬請購人員
 BpmNo						nvarchar(100)			NULL,						--BPM單號
 PrErpPrefix				nvarchar(4)			NOT NULL,						--TA001 請購單別 >> 參照 ERP各系統單據設定檔(CMSMQ)
 PrErpNo					nvarchar(11)		NOT NULL,						--TA002 請購單號
 Edition					nvarchar(4)			NOT NULL,						--TA021 版次
 PrDate						datetime			NOT NULL,						--TA003 請購日期
 DocDate					datetime			NOT NULL,						--TA013 單據日期
 PrRemark					nvarchar(255)			NULL,						--TA006 單頭備註
 TotalQty					float					NULL,						--TA011 請購總數量
 Amount						float					NULL,						--TA020 請購總金額
 BudgetDepartmentId			int						NULL,						--TA022 預算部門
 PrStatus					nvarchar(2)			NOT NULL,						--請購狀態碼 >> BAS.[Status] N：新單據；L：拋轉BPM中；P：BPM簽核流程；R：拋轉BPM失敗；F：拋轉ERP失敗；E：單據修改；S：MES作廢；Y：單據核准；V：ERP作廢
 SignupStaus				nvarchar(1)			NOT NULL,						--TA016 簽核狀態碼 >> BAS.[Status] 0.待處理、1.簽核中、2.退件、3.已核准、4.取消確認中、5.作廢中、6.取消作廢中
 LockStaus					nvarchar(1)			NOT NULL,						--TA028 鎖定碼 >> BAS.[Status] N；Y
 ConfirmStatus				nvarchar(1)			NOT NULL,						--TA007 確認碼 >> BAS.[Status] N：未確認；V：作廢；Y：已確認
 ConfirmUserId				int						NULL,						--TA014 確認者
 [Priority]					nvarchar(2)			NOT	NULL	DEFAULT '2',		--優先度 >> BAS.[Type] TypeSchema = 'PR.Priority' 0:最速件 1:速件 2:普通件
 [Source]					nvarchar(200)		NOT NULL,						--來源
 BomType					nvarchar(2)			NOT NULL	DEFAULT 'Y',		--生產用BOM表內 or 其他
 BpmTransferStatus			nvarchar(2)			NOT NULL,						--拋轉BPM狀態 >> BAS.[Status]
 BpmTransferUserId			int						NULL,						--拋轉BPM人員
 BpmTransferDate			datetime				NULL,						--拋轉BPM時間
 TransferStatus				nvarchar(2)			NOT NULL,						--拋轉ERP狀態 >> BAS.[Status]
 TransferDate				datetime				NULL,						--拋轉ERP時間
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_PurchaseRequisition PRIMARY KEY (PrId),
 CONSTRAINT AK1_PurchaseRequisition UNIQUE (CompanyId, PrErpPrefix, PrErpNo),
 CONSTRAINT AK2_PurchaseRequisition UNIQUE (CompanyId, PrNo),
 CONSTRAINT FK1_PurchaseRequisition FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_PurchaseRequisition FOREIGN KEY (DepartmentId) REFERENCES BAS.Department(DepartmentId),
 CONSTRAINT FK3_PurchaseRequisition FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId)
)

CREATE TABLE SCM.PrDetail( /*請購單身*/ /*ERP PURTB*/
 PrDetailId					int					identity,
 PrId						int					NOT NULL,						--所屬單頭
 PrSequence					nvarchar(4)			NOT NULL,						--TB003 流水號
 MtlItemId					int					NOT NULL,						--TB004 品號
 PrMtlItemName				nvarchar(120)		NOT NULL,						--TB005 品名
 PrMtlItemSpec				nvarchar(120)		NOT NULL,						--TB006 規格
 InventoryId				int					NOT NULL,						--TB008 庫別
 PrUomId					int					NOT NULL,						--TB007 請購單位(Unit of Measure) >> 參照 ERP品號換算單位檔(INVMD)
 PrQty						int					NOT NULL,						--TB009 請購數量
 DemandDate					datetime			NOT NULL,						--TB011 需求日期
 SupplierId					int						NULL,						--TB010 供應商
 PrCurrency					nvarchar(4)				NULL,						--TB050 請購幣別 >> 參照 ERP幣別匯率檔單頭(CMSMF)
 PrExchangeRate				float					NULL,						--TB044 請購匯率
 PrUnitPrice				float					NULL,						--TB049 請購單位金額
 PrPrice					float					NULL,						--TB051 請購金額
 PrPriceTw					float					NULL,						--TB045 請購本幣金額
 UrgentMtl					nvarchar(1)			NOT NULL,						--TB032 急料 >> BAS.[Type] N[DEF]；Y
 ProductionPlan				nvarchar(2)			NOT NULL,						--納入生產計畫 >> BAS.[Type] N；Y
 Project					nvarchar(20)			NULL,						--TB043 專案代號 >> 參考 ERP專案代號檔(CMSNB)
 BudgetDepartmentNo			nvarchar(6)				NULL,						--TB047 預算編號
 BudgetDepartmentSubject	nvarchar(20)			NULL,						--TB048 預算科目
 SoDetailId					int						NULL,						--TB029/30/31 參考訂單
 DeliveryDate				datetime				NULL,						--TB019 交貨日期
 PoUserId					int						NULL,						--TB013 採購人員
 PoUomId					int						NULL,						--TB015 採購單位(Unit of Measure) >> 參照 ERP品號換算單位檔(INVMD)
 PoQty						int						NULL,						--TB014 採購數量
 PoCurrency					nvarchar(4)				NULL,						--TB016 採購幣別 >> 參照 ERP幣別匯率檔單頭(CMSMF)
 PoUnitPrice				float					NULL,						--TB017 採購單位金額
 PoPrice					float					NULL,						--TB018 採購金額
 PoErpPrefixNo				nvarchar(21)			NULL,						--TB022 採購單號
 LockStaus					nvarchar(1)			NOT NULL,						--TB020 鎖定碼 >> BAS.[Status] N[DEF]；Y
 PoStaus					nvarchar(1)			NOT NULL,						--TB021 採購碼 >> BAS.[Status] N[DEF]；Y
 PartialPurchaseStaus		nvarchar(1)			NOT NULL,						--TB046 分批採購 >> BAS.[Status] N[DEF]；Y
 InquiryStatus				nvarchar(1)			NOT NULL,						--TB040 詢價碼 >> BAS.[Type] 1：不需詢價[DEF]；2：詢價中；3：審核中；4：已核准
 TaxNo						nvarchar(3)				NULL,						--TB057 稅別碼 >> 參照 ERP稅別碼設定檔(CMSNN)
 Taxation					nvarchar(1)				NULL,						--TB026 課稅別 >> 由稅別碼決定 1：內含；2：外加；3：零稅率；4：免稅；9：不計稅
 BusinessTaxRate			float					NULL,						--TB063 營業稅率 >> 由稅別碼決定
 DetailMultiTax				nvarchar(1)				NULL,						--TB064 單身多稅率 >> BAS.[Status] N(DEF)；Y
 TradeTerm					nvarchar(1)				NULL,						--TB058 交易條件 >> 參照 ERP交易條件代號檔(CMSNK)
 PrPriceQty					int					NOT NULL,						--TB065 計價數量
 PrPriceUomId				int					NOT NULL,						--TB066 計價單位
 DiscountRate				float					NULL,						--TB067 折扣率
 DiscountAmount				float					NULL,						--TB068 折扣金額
 MtlInventory				nvarchar(255)			NULL,						--品號庫存狀況
 MtlInventoryQty			float				NOT NULL,						--品號庫存總數量
 ConfirmStatus				nvarchar(1)			NOT NULL,						--TB025 確認碼 >> BAS.[Status] N：未確認；V：作廢；Y：已確認
 ConfirmUserId				int						NULL,						--確認者
 ClosureStatus				nvarchar(1)			NOT NULL,						--TB039 結案碼 >> BAS.[Status] N：未結案；y：指定結案；Y：自動結案
 PrDetailRemark				nvarchar(255)			NULL,						--TB012 單身備註
 PoRemark					nvarchar(255)			NULL,						--TB024 採購備註
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_PrDetail PRIMARY KEY (PrDetailId),
 CONSTRAINT AK1_PrDetail UNIQUE (PrId, PrSequence),
 CONSTRAINT FK1_PrDetail FOREIGN KEY (PrId) REFERENCES SCM.PurchaseRequisition(PrId),
 CONSTRAINT FK2_PrDetail FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId),
 CONSTRAINT FK3_PrDetail FOREIGN KEY (InventoryId) REFERENCES SCM.Inventory(InventoryId),
 CONSTRAINT FK4_PrDetail FOREIGN KEY (PrUomId) REFERENCES PDM.UnitOfMeasure(UomId),
 CONSTRAINT FK5_PrDetail FOREIGN KEY (PrPriceUomId) REFERENCES PDM.UnitOfMeasure(UomId)
)

CREATE TABLE SCM.PrFile( /*請購附檔*/
 PrFileId					int					identity,
 PrId						int					NOT NULL,						--所屬請購
 PrDetailId					int						NULL,						--所屬請購單單身ID
 FileId						int						NULL,						--檔案ID
 FilePath					nvarchar(max)			NULL,						--檔案路徑(實體路徑)
 ErpDocId					nvarchar(10)			NULL,						--ERP檔案ID
 ErpDocPath					nvarchar(200)			NULL,						--ERP檔案路徑
 ErpDocDate					datetime				NULL,						--ERP檔案建立日期
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_PrFile PRIMARY KEY (PrFileId),
 CONSTRAINT FK1_PrFile FOREIGN KEY (PrId) REFERENCES SCM.PurchaseRequisition(PrId),
 CONSTRAINT FK2_PrFile FOREIGN KEY (FileId) REFERENCES BAS.[File](FileId),
 CONSTRAINT FK3_PrFile FOREIGN KEY (PrDetailId) REFERENCES SCM.PrDetail(PrDetailId)
 /*Note：ERP DocName nvarchar(100)，拋轉ERP需判斷大小*/
)

CREATE TABLE SCM.PrLog( /*請購單拋轉LOG*/
 PrLogId					int					identity,
 PrId						int					NOT NULL,						--所屬請購單ID
 RootId						nvarchar(30)		NOT NULL,						--任務編號
 BpmNo						nvarchar(100)		NOT	NULL,						--BPM單號
 TransferBpmDate			datetime			NOT NULL,						--拋轉BPM日期
 BpmStatus					nvarchar(2)			NOT NULL,						--拋轉狀態
 ConfirmUser				nvarchar(100)		NOT NULL,						--最後核單USER NO
 ErrorMessage				nvarchar(1000)		NOT NULL,						--Error訊息
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_PrLog PRIMARY KEY (PrLogId),
 CONSTRAINT FK1_PrLog FOREIGN KEY (PrId) REFERENCES SCM.PurchaseRequisition(PrId)
)

CREATE TABLE SCM.PrModification( /*請購變更單頭*/ /*ERP PURTU*/
 PrmId						int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 PrId						int					NOT NULL,						--所屬請購單
 DepartmentId				int					NOT NULL,						--TU006 所屬部門
 UserId						int					NOT NULL,						--TU007 所屬人員
 BpmNo						nvarchar(100)		NOT NULL,						--BPM單號
 Edition					nvarchar(4)			NOT NULL,						--TU003 版次
 ModiDate					datetime			NOT NULL,						--TU004 變更日期
 DocDate					datetime			NOT NULL,						--TU005 單據日期
 ModiReason					nvarchar(255)			NULL,						--TU009 變更原因
 PrmRemark					nvarchar(255)			NULL,						--TU018 單頭備註
 TotalQty					float				NOT NULL,						--請購總數量
 Amount						float				NOT NULL,						--TU019 請購總金額
 BudgetDepartmentId			int						NULL,						--TU020 預算部門
 OriDepartmentId			int					NOT NULL,						--TU106 原所屬部門
 OriUserId					int					NOT NULL,						--TU107 原所屬人員
 OriEdition					nvarchar(4)			NOT NULL,						--TU103 原版次
 OriPrDate					datetime			NOT NULL,						--TU017 原請購日期
 OriPrRemark				nvarchar(255)			NULL,						--TU108 原單頭備註
 OriAmount					float				NOT NULL,						--TU109 原請購總金額
 OriBudgetDepartmentId		int						NULL,						--TU120 原預算部門
 PrmStatus					nvarchar(2)			NOT NULL,						--請購變更狀態碼 >> BAS.[Status] N：新單據；L：拋轉BPM中；P：BPM簽核流程；R：拋轉BPM失敗；F：拋轉ERP失敗；E：單據修改；S：MES作廢；Y：單據核准；V：ERP作廢
 SignupStaus				nvarchar(1)			NOT NULL,						--TU014 簽核狀態碼 >> BAS.[Status] 0.待處理、1.簽核中、2.退件、3.已核准、4.取消確認中、5.作廢中、6.取消作廢中
 ConfirmStatus				nvarchar(1)			NOT NULL,						--TU012 確認碼 >> BAS.[Status] N：未確認；V：作廢；Y：已確認
 ConfirmUserId				int						NULL,						--TU013 確認者
 WholeClosureStatus			nvarchar(1)			NOT NULL,						--TU008 結案碼 >> BAS.[Status] N：未結案；Y：自動結案
 BpmTransferStatus			nvarchar(2)			NOT NULL,						--拋轉BPM狀態 >> BAS.[Status]
 BpmTransferUserId			int						NULL,						--拋轉BPM人員
 BpmTransferDate			datetime				NULL,						--拋轉BPM時間
 [Priority]					nvarchar(2)			NOT	NULL	DEFAULT '2',		--優先度 >> BAS.[Type] TypeSchema = 'PR.Priority' 0:最速件 1:速件 2:普通件
 TransferStatus				nvarchar(2)			NOT NULL,						--拋轉ERP狀態 >> BAS.[Status]
 TransferDate				datetime				NULL,						--拋轉ERP時間
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_PrModification PRIMARY KEY (PrmId),
 CONSTRAINT AK1_PrModification UNIQUE (PrId, Edition, PrmStatus),
 CONSTRAINT FK1_PrModification FOREIGN KEY (PrId) REFERENCES SCM.PurchaseRequisition(PrId),
 CONSTRAINT FK2_PrModification FOREIGN KEY (DepartmentId) REFERENCES BAS.Department(DepartmentId),
 CONSTRAINT FK3_PrModification FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK4_PrModification FOREIGN KEY (OriDepartmentId) REFERENCES BAS.Department(DepartmentId),
 CONSTRAINT FK5_PrModification FOREIGN KEY (OriUserId) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK6_PrModification FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)

CREATE TABLE SCM.PrmDetail( /*請購變更單身*/ /*ERP PURTV*/
 PrmDetailId				int					identity,
 PrmId						int					NOT NULL,						--所屬單頭
 PrDetailId					int						NULL,						--所屬請購單身
 PrmSequence				nvarchar(4)			NOT NULL,						--TV004 流水號
 MtlItemId					int					NOT NULL,						--TV005 品號
 PrMtlItemName				nvarchar(120)		NOT NULL,						--TV006 品名
 PrMtlItemSpec				nvarchar(120)		NOT NULL,						--TV007 規格
 InventoryId				int					NOT NULL,						--TV009 庫別
 PrUomId					int					NOT NULL,						--TV008 請購單位(Unit of Measure) >> 參照 ERP品號換算單位檔(INVMD)
 PrQty						int					NOT NULL,						--TV010 請購數量
 DemandDate					datetime			NOT NULL,						--TV012 需求日期
 SupplierId					int						NULL,						--TV011 供應商
 PrCurrency					nvarchar(4)				NULL,						--TV047 請購幣別 >> 參照 ERP幣別匯率檔單頭(CMSMF)
 PrExchangeRate				float					NULL,						--TV042 請購匯率
 PrUnitPrice				float					NULL,						--TV046 請購單位金額
 PrPrice					float					NULL,						--TV048 請購金額
 PrPriceTw					float					NULL,						--TV043 請購本幣金額
 UrgentMtl					nvarchar(1)			NOT NULL,						--TV027 急料 >> BAS.[Type] N[DEF]；Y
 ProductionPlan				nvarchar(2)			NOT NULL,						--納入生產計畫 >> BAS.[Type] N；Y
 Project					nvarchar(20)			NULL,						--TV041 專案代號 >> 參考 ERP專案代號檔(CMSNB)
 BudgetDepartmentNo			nvarchar(6)				NULL,						--TV044 預算編號
 BudgetDepartmentSubject	nvarchar(20)			NULL,						--TV045 預算科目
 SoDetailId					int						NULL,						--TV024/25/26 參考訂單
 DeliveryDate				datetime				NULL,						--TV019 交貨日期
 PoUserId					int						NULL,						--TV013 採購人員
 PoUomId					int						NULL,						--TV015 採購單位(Unit of Measure) >> 參照 ERP品號換算單位檔(INVMD)
 PoQty						int						NULL,						--TV014 採購數量
 PoCurrency					nvarchar(4)				NULL,						--TV016 採購幣別 >> 參照 ERP幣別匯率檔單頭(CMSMF)
 PoUnitPrice				float					NULL,						--TV017 採購單位金額
 PoPrice					float					NULL,						--TV018 採購金額
 LockStaus					nvarchar(1)			NOT NULL,						--TV035 鎖定碼 >> BAS.[Status] N[DEF]；Y
 PoStaus					nvarchar(1)			NOT NULL,						--TV036 採購碼 >> BAS.[Status] N[DEF]；Y
 TaxNo						nvarchar(3)				NULL,						--TV054 稅別碼 >> 參照 ERP稅別碼設定檔(CMSNN)
 Taxation					nvarchar(1)				NULL,						--TV021 課稅別 >> 由稅別碼決定 1：內含；2：外加；3：零稅率；4：免稅；9：不計稅
 BusinessTaxRate			float					NULL,						--TV057 營業稅率 >> 由稅別碼決定
 DetailMultiTax				nvarchar(1)				NULL,						--TV056 單身多稅率 >> BAS.[Status] N(DEF)；Y
 TradeTerm					nvarchar(1)				NULL,						--TV055 交易條件 >> 參照 ERP交易條件代號檔(CMSNK)
 PrPriceQty					int					NOT NULL,						--TV058 計價數量
 PrPriceUomId				int					NOT NULL,						--TV059 計價單位
 DiscountRate				float					NULL,						--TV060 折扣率
 DiscountAmount				float					NULL,						--TV061 折扣金額
 MtlInventory				nvarchar(255)			NULL,						--品號庫存狀況
 MtlInventoryQty			float				NOT NULL,						--品號庫存總數量
 ConfirmStatus				nvarchar(1)			NOT NULL,						--TV037 確認碼 >> BAS.[Status] N：未確認；V：作廢；Y：已確認
 ConfirmUserId				int						NULL,						--確認者
 ClosureStatus				nvarchar(1)			NOT NULL,						--TV034 結案碼 >> BAS.[Status] N：未結案；y：指定結案；Y：自動結案
 PrDetailRemark				nvarchar(255)			NULL,						--TV039 單身備註
 PoRemark					nvarchar(255)			NULL,						--TV033 採購備註
 ModiReason					nvarchar(255)			NULL,						--TV038 變更原因
 OriPrSequence				nvarchar(4)				NULL,						--TV104 原流水號
 OriMtlItemId				int						NULL,						--TV105 原品號
 OriPrMtlItemName			nvarchar(120)			NULL,						--TV106 原品名
 OriPrMtlItemSpec			nvarchar(120)			NULL,						--TV107 原規格
 OriInventoryId				int						NULL,						--TV109 原庫別
 OriPrUomId					int						NULL,						--TV108 原請購單位(Unit of Measure) >> 參照 ERP品號換算單位檔(INVMD)
 OriPrQty					int						NULL,						--TV110 原請購數量
 OriDemandDate				datetime				NULL,						--TV112 原需求日期
 OriSupplierId				int						NULL,						--TV111 原供應商
 OriPrCurrency				nvarchar(4)				NULL,						--TV147 原請購幣別 >> 參照 ERP幣別匯率檔單頭(CMSMF)
 OriPrExchangeRate			float					NULL,						--TV142 原請購匯率
 OriPrUnitPrice				float					NULL,						--TV146 原請購單位金額
 OriPrPrice					float					NULL,						--TV148 原請購金額
 OriPrPriceTw				float					NULL,						--TV143 原請購本幣金額
 OriUrgentMtl				nvarchar(1)			NOT NULL,						--TV127 原急料 >> BAS.[Type] N[DEF]；Y
 OriProductionPlan			nvarchar(2)			NOT NULL,						--原納入生產計畫 >> BAS.[Type] N；Y
 OriProject					nvarchar(20)			NULL,						--TV141 原專案代號 >> 參考 ERP專案代號檔(CMSNB)
 OriBudgetDepartmentNo		nvarchar(6)				NULL,						--TV144 原預算編號
 OriBudgetDepartmentSubject	nvarchar(20)			NULL,						--TV145 原預算科目
 OriSoDetailId				int						NULL,						--TV124/125/126 原參考訂單
 OriDeliveryDate			datetime				NULL,						--TV119 原交貨日期
 OriPoUserId				int						NULL,						--TV113 原採購人員
 OriPoUomId					int						NULL,						--TV115 原採購單位(Unit of Measure) >> 參照 ERP品號換算單位檔(INVMD)
 OriPoQty					int						NULL,						--TV114 原採購數量
 OriPoCurrency				nvarchar(4)				NULL,						--TV116 原採購幣別 >> 參照 ERP幣別匯率檔單頭(CMSMF)
 OriPoUnitPrice				float					NULL,						--TV117 原採購單位金額
 OriPoPrice					float					NULL,						--TV118 原採購金額
 OriLockStaus				nvarchar(1)			NOT NULL,						--TV135 原鎖定碼 >> BAS.[Status] N[DEF]；Y
 OriPoStaus					nvarchar(1)			NOT NULL,						--TV136 原採購碼 >> BAS.[Status] N[DEF]；Y
 OriTaxNo					nvarchar(3)				NULL,						--TV154 原稅別碼 >> 參照 ERP稅別碼設定檔(CMSNN)
 OriTaxation				nvarchar(1)				NULL,						--TV121 原課稅別 >> 由稅別碼決定 1：內含；2：外加；3：零稅率；4：免稅；9：不計稅
 OriBusinessTaxRate			float					NULL,						--TV157 原營業稅率 >> 由稅別碼決定
 OriDetailMultiTax			nvarchar(1)				NULL,						--TV156 原單身多稅率 >> BAS.[Status] N(DEF)；Y
 OriTradeTerm				nvarchar(1)				NULL,						--TV155 原交易條件 >> 參照 ERP交易條件代號檔(CMSNK)
 OriPrPriceQty				int					NOT NULL,						--TV158 計價數量
 OriPrPriceUomId			int						NULL,						--TV159 計價單位
 OriDiscountRate			float					NULL,						--TV160 原折扣率
 OriDiscountAmount			float					NULL,						--TV161 原折扣金額
 OriClosureStatus			nvarchar(1)			NOT NULL,						--TV134 原結案碼 >> BAS.[Status] N：未結案；y：指定結案；Y：自動結案
 OriPrDetailRemark			nvarchar(255)			NULL,						--TV137 原單身備註
 OriPoRemark				nvarchar(255)			NULL,						--TV133 原採購備註
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_PrmDetail PRIMARY KEY (PrmDetailId),
 CONSTRAINT AK1_PrmDetail UNIQUE (PrmId, PrmSequence),
 CONSTRAINT FK1_PrmDetail FOREIGN KEY (PrmId) REFERENCES SCM.PrModification(PrmId),
 CONSTRAINT FK2_PrmDetail FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId),
 CONSTRAINT FK3_PrmDetail FOREIGN KEY (InventoryId) REFERENCES SCM.Inventory(InventoryId),
 CONSTRAINT FK4_PrmDetail FOREIGN KEY (PrUomId) REFERENCES PDM.UnitOfMeasure(UomId),
 CONSTRAINT FK5_PrmDetail FOREIGN KEY (PrPriceUomId) REFERENCES PDM.UnitOfMeasure(UomId)
)

CREATE TABLE SCM.PrmFile( /*請購變更附檔*/
 PrmFileId					int					identity,
 PrmId						int					NOT NULL,						--所屬請購
 FileId						int					NOT NULL,						--所屬檔案
 ErpDocId					nvarchar(10)			NULL,						--ERP檔案ID
 ErpDocPath					nvarchar(200)			NULL,						--ERP檔案路徑
 ErpDocDate					datetime				NULL,						--ERP檔案建立日期
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_PrmFile PRIMARY KEY (PrmFileId),
 CONSTRAINT FK1_PrmFile FOREIGN KEY (PrmId) REFERENCES SCM.PrModification(PrmId),
 CONSTRAINT FK2_PrmFile FOREIGN KEY (FileId) REFERENCES BAS.[File](FileId)
 /*Note：ERP DocName nvarchar(100)，拋轉ERP需判斷大小*/
)

CREATE TABLE SCM.PrmLog( /*請購變更單拋轉LOG*/
 PrmLogId					int					identity,
 PrmId						int					NOT NULL,						--所屬請購單ID
 RootId						nvarchar(30)		NOT NULL,						--任務編號
 BpmNo						nvarchar(100)		NOT	NULL,						--BPM單號
 TransferBpmDate			datetime			NOT NULL,						--拋轉BPM日期
 BpmStatus					nvarchar(2)			NOT NULL,						--拋轉狀態
 ConfirmUser				nvarchar(100)		NOT NULL,						--最後核單USER NO
 ErrorMessage				nvarchar(1000)		NOT NULL,						--Error訊息
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_PrmLog PRIMARY KEY (PrmLogId),
 CONSTRAINT FK1_PrmLog FOREIGN KEY (PrmId) REFERENCES SCM.PrModification(PrmId)
)
END /*請購*/

BEGIN /*進貨*/
CREATE TABLE SCM.GoodsReceipt( /*進貨單頭*/ /*ERP PURTG*/
 GrId						int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 GrErpPrefix				nvarchar(4)			NOT NULL,						--TG001 進貨單單別
 GrErpNo					nvarchar(11)		NOT NULL,						--TG002 進貨單單號
 ReceiptDate				datetime			NOT NULL,						--TG003 進貨日期
 SupplierId					int					NOT NULL,						--MES供應商ID
 SupplierNo					nvarchar(100)		NOT NULL,						--TG005 供應商
 SupplierSo					nvarchar(20)			NULL,						--TG006	廠商單號/放空值即可,預留欄位
 CurrencyCode				nvarchar(4)			NOT NULL,						--TG007 幣別
 Exchange					float				NOT NULL	DEFAULT 0.0000,		--TG008 匯率
 InvoiceType				nvarchar(1)				NULL,						--TG009	發票聯數 1.二聯式、2.三聯式、3.二聯式收銀機發票、4.三聯式收銀機發票、5.電子計算機發票、6.免用統一發票、A.農產品收購憑證、G.海關代徵完稅憑證、N.不可抵扣專用發票、S.可抵扣專用發票、T.運輸發票、W.廢舊物資收購憑證、Z.其他
 TaxType					nvarchar(1)				NULL,						--TG010 課稅別 1.應稅內含、2.應稅外加、3.零稅率、4.免稅、9.不計稅
 InvoiceNo					nvarchar(50)			NULL,						--TG011	發票號碼
 PrintCnt					int					NOT NULL	DEFAULT 0,			--TG012 列印次數
 ConfirmStatus				nvarchar(1)			NOT NULL	DEFAULT 'N',		--TG013 確認碼(Y/N/V)
 DocDate					datetime			NOT NULL,						--TG014 單據日期
 RenewFlag					nvarchar(1)			NOT NULL	DEFAULT 'N',		--TG015 生產記錄更新碼(N) 不需維護, ERP用
 Remark						nvarchar(255)			NULL,						--TG016 備註
 TotalAmount				float				NOT NULL	DEFAULT 0.0000,		--TG017	進貨金額(單身原幣進貨金額TH019合計)
 DeductAmount				float				NOT NULL	DEFAULT 0.0000,		--TG018 扣款金額(單身原幣扣款金額TH020合計)
 OrigTax					float				NOT NULL	DEFAULT 0.0000,		--TG019 原幣稅額(單身原幣稅額TH046合計)
 ReceiptAmount				float				NOT NULL	DEFAULT 0.0000,		--TG020	進貨費用(單身進貨費用TH024合計)
 SupplierName				nvarchar(150)		NOT	NULL,						--TG021 廠商全名
 UiNo						nvarchar(20)			NULL,						--TG022	統一編號
 DeductType					nvarchar(1)				NULL,						--TG023 抵扣區分 1.可扣抵進貨及費用 2.可扣抵固定資產 3.不可扣抵進貨及費用 4.不可扣抵固定資產
 ObaccoAndLiquorFlag		nvarchar(1)			NOT	NULL	DEFAULT 'N',		--TG024 菸酒註記 Y/N
 RowCnt						int					NOT NULL	DEFAULT 0,			--TG025 件數
 Quantity					int					NOT NULL	DEFAULT 0,			--TG026	數量合計(單身計價數量合計)
 InvoiceDate				datetime				NULL,						--TG027	發票日期
 OrigPreTaxAmount			float				NOT NULL	DEFAULT 0.0000,		--TG028 原幣未稅貨款金額(原幣未稅金額TH045合計)
 ApplyYYMM					nvarchar(50)		NOT NULL,						--TG029 申報年月(YYYYMM) 申報年月[FORMATE:YM]&&880701ADD
 TaxRate					float					NULL,						--TG030 營業稅率 營業稅率(%)[DEF:9.9999]&&89-06-01ADD//950808 S00-9508031 N5.4->N8.5
 PreTaxAmount				float				NOT NULL	DEFAULT 0.0000,		--TG031 本幣未稅貨款金額(本幣未稅金額TH047合計)
 TaxAmount					float				NOT NULL	DEFAULT 0.0000,		--TG032 本幣稅額(單身本幣稅額TH048合計)
 PaymentTerm				nvarchar(6)			NOT NULL,						--TG033 付款條件代號&&89-09-06ADD
 PoId						int						NULL,						--MES採購單ID
 PoErpPrefix				nvarchar(4)				NULL,						--TG034 採購單別
 PoErpNo					nvarchar(11)			NULL,						--TG035 採購單號
 PrepaidErpPrefix			nvarchar(4)				NULL,						--TG036 預付待抵單別
 PrepaidErpNo				nvarchar(11)			NULL,						--TG037 預付待抵單號
 OffsetAmount				float				NOT	NULL	DEFAULT 0.0000,		--TG038 沖抵金額
 TaxOffset					float				NOT	NULL	DEFAULT 0.0000,		--TG039 沖抵稅額
 PackageQuantity			int					NOT NULL	DEFAULT 0,			--TG040 包裝單身數量合計
 PremiumAmount				float				NOT	NULL	DEFAULT 0.0000,		--TG041 本幣沖自籌額(單身總計)
 ApproveStatus				nvarchar(1)			NOT NULL	DEFAULT 'N',		--TG042 簽核狀態碼 0.待處理、1.簽核中、2.退件、3.已核准、4.取消確認中、5.作廢中、6.取消作廢中、N.不執行電子簽核[DEF:"N"] 
 InvoiceFlag				nvarchar(1)			NOT NULL	DEFAULT 'N',		--TG043 隨貨附發票 Y/N
 ReserveTaxCode             nvarchar(1)			NOT NULL	DEFAULT '0',		--TG044 保稅碼 0.依品號預設 Y.預設為保稅品 N.預設為非保稅品
 SendCount					int					NOT NULL	DEFAULT 0,			--TH038 傳送次數, 永遠傳0
 OrigPremiumAmount			float				NOT	NULL	DEFAULT 0.0000,		--TG046 原幣沖自籌額(單身總計)
 EbcErpPreNo				nvarchar(25)			NULL,						--TG047 EBC出貨通知單號
 EbcEdition					nvarchar(4)				NULL,						--TG048 EBC出貨通知版次
 EbcFlag					nvarchar(1)			NOT	NULL	DEFAULT 'N',		--TG049 進貨單是否已匯出給EBC Y/N
 FromDocType				nvarchar(1)			NOT	NULL	DEFAULT '1',		--TG050 來源類別 1.採購 2.進口 3.製令
 NoticeFlag					nvarchar(1)			NOT NULL	DEFAULT 'N',		--TG055 通知碼, 永遠傳N
 TaxCode					nvarchar(3)			NOT NULL,						--TG058 稅別碼 (要找ERP資料來源)
 TradeTerm					nvarchar(1)			NOT NULL,						--TG059 交易條件 >> 參照 ERP交易條件代號檔(CMSNK)
 DetailMultiTax				nvarchar(1)			NOT NULL	DEFAULT 'N',		--TG063 單身多稅率 Y/N
 TaxExchange				float				NOT NULL	DEFAULT 1,			--TG064 稅幣匯率, 永遠帶1
 ContactUser				nvarchar(30)			NULL,						--TG067 供應商聯絡人
 TransferStatus				nvarchar(2)			NOT NULL	DEFAULT 'N',		--拋轉ERP狀態 >> BAS.[Status]
 TransferDate				datetime				NULL,						--拋轉ERP時間
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_GoodsReceipt PRIMARY KEY (GrId),
 CONSTRAINT AK1_GoodsReceipt UNIQUE (CompanyId, GrErpPrefix, GrErpNo),
 CONSTRAINT FK1_GoodsReceipt FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_GoodsReceipt FOREIGN KEY (SupplierId) REFERENCES SCM.Supplier(SupplierId),
 CONSTRAINT FK3_GoodsReceipt FOREIGN KEY (PoId) REFERENCES SCM.PurchaseOrder(PoId)
)

CREATE TABLE SCM.GrDetail( /*進貨單身*/ /*ERP PURTH*/
 GrDetailId					int					identity,
 GrId						int					NOT NULL,						--所屬進貨單頭ID
 GrErpPrefix				nvarchar(4)			NOT NULL,						--TH001 進貨單單別
 GrErpNo					nvarchar(11)		NOT NULL,						--TH002 進貨單單號
 GrSeq						nvarchar(4)			NOT NULL,						--TH003 序號 ex:0001,0002,0003
 MtlItemId					int					NOT NULL,						--MES品號ID
 MtlItemNo					nvarchar(100)		NOT NULL,						--TH004 品號
 GrMtlItemName				nvarchar(120)		NOT NULL,						--TH005 品名
 GrMtlItemSpec				nvarchar(120)		NOT NULL,						--TH006 規格
 ReceiptQty					int					NOT NULL,						--TH007 進貨數量
 UomId						int					NOT NULL,						--MES進貨單位
 UomNo						nvarchar(120)		NOT NULL,						--TH008 進貨單位
 InventoryId				int					NOT NULL,						--MES進貨庫別
 InventoryNo				nvarchar(120)		NOT NULL,						--TH009 進貨庫別
 LotNumber					nvarchar(20)			NULL,						--TH010	批號
 PoDetailId					int						NULL,						--採購單身ID
 PoErpPrefix				nvarchar(4)			NOT NULL,						--TH011 採購單別
 PoErpNo					nvarchar(11)		NOT NULL,						--TH012 採購單號
 PoSeq						nvarchar(4)			NOT NULL,						--TH013 採購序號
 AcceptanceDate				datetime			NOT NULL,						--TH014 驗收日期
 AcceptQty					int					NOT NULL	DEFAULT 0,			--TH015 驗收數量
 AvailableQty				int					NOT NULL	DEFAULT 0,			--TH016 計價數量
 ReturnQty					int					NOT NULL	DEFAULT 0,			--TH017 驗退數量
 OrigUnitPrice				float				NOT NULL	DEFAULT 0.000000,	--TH018 原幣進貨單價
 OrigAmount					float				NOT NULL	DEFAULT 0.000000,	--TH019 原幣進貨金額
 OrigDiscountAmt			float				NOT NULL	DEFAULT 0.000000,	--TH020 原幣扣款金額
 TrErpPrefix				nvarchar(4)			NOT NULL,						--TH021 暫入單別
 TrErpNo					nvarchar(11)		NOT NULL,						--TH022 暫入單號
 TrSeq						nvarchar(4)			NOT NULL,						--TH013 暫入序號
 ReceiptExpense				float				NOT NULL	DEFAULT 0.000000,	--TH024 進貨費用
 DiscountDescription		nvarchar(40)			NULL,						--TH025 扣款說明
 PaymentHold				nvarchar(1)			NOT NULL	DEFAULT 'N',		--TH026 暫不付款(不用顯示)
 Overdue					nvarchar(1)			NOT NULL	DEFAULT 'N',		--TH027	逾期碼 (Y/N)
 QcStatus					nvarchar(1)			NOT NULL,						--TH028 檢驗狀態 0.免檢、1.待驗、2.合格、3.不良、4.特採
 ReturnCode					nvarchar(1)			NOT NULL	DEFAULT 'N',		--TH029 驗退碼 (Y/N)(不用顯示)
 ConfirmStatus				nvarchar(1)			NOT NULL	DEFAULT 'N',		--TH030 確認碼 (Y/N/V)
 CloseStatus				nvarchar(1)			NOT NULL	DEFAULT 'N',		--TH031 結帳碼 (Y/N)
 ReNewStatus				nvarchar(1)			NOT NULL	DEFAULT 'N',		--TH032 更新碼 (Y/N)(隱藏欄位,不需顯示)
 Remark						nvarchar(255)			NULL,						--TH033 備註
 InventoryQty				float				NOT	NULL	DEFAULT 0.000000,	--TH034 庫存數量
 SmallUnit					nvarchar(6)				NULL,						--TH035 小單位
 AvailableDate				datetime				NULL,						--TH036 有效日期(可空)
 ReCheckDate				datetime				NULL,						--TH037 複檢日期(可空)
 ConfirmUser				int						NULL,						--TH038 確認者
 ApvErpPrefix				nvarchar(4)				NULL,						--TH039 應付憑單別
 ApvErpNo					nvarchar(11)			NULL,						--TH040 應付憑單號
 ApvSeq						nvarchar(4)				NULL,						--TH041 應付憑序號
 ProjectCode				nvarchar(20)			NULL,						--TH042 專案代碼
 ExpenseEntry               nvarchar(1)			NOT NULL	DEFAULT 'N',		--TH043 產生分錄碼-費用 (Y/N)(不需顯示)
 PremiumAmountFlag			nvarchar(1)			NOT	NULL	DEFAULT 'N',		--TH044 沖自籌額碼
 OrigPreTaxAmt				float				NOT NULL	DEFAULT 0.000000,	--TH045 原幣未稅金額
 OrigTaxAmt					float				NOT NULL	DEFAULT 0.000000,	--TH046 原幣稅金額
 PreTaxAmt					float				NOT NULL	DEFAULT 0.000000,	--TH047 本幣未稅金額
 TaxAmt						float				NOT NULL	DEFAULT 0.000000,	--TH048 本幣稅金額
 ReceiptPackageQty			int					NOT NULL	DEFAULT 0,			--TH049 進貨包裝數量, 沒使用, Default填0 (不需顯示)
 AcceptancePackageQty		int					NOT NULL	DEFAULT 0,			--TH050	驗收包裝數量 (不需顯示)
 ReturnPackageQty			int					NOT NULL	DEFAULT 0,			--TH051	驗退包裝數量 (不需顯示)
 PremiumAmount				float				NOT	NULL	DEFAULT 0.0000,		--TH052 本幣沖自籌額
 PackageUnit				nvarchar(6)				NULL,						--TH053 包裝單位
 ReserveTaxCode             nvarchar(1)			NOT NULL	DEFAULT 'N',		--TH054 保稅碼 Y/N 
 OrigPremiumAmount			float				NOT	NULL	DEFAULT 0.0000,		--TH055 原幣沖自籌額
 AvailableUomId				int					NOT NULL,						--MES計價單位ID => 預設與進貨單位相同，暫不開放修改(採購需求)
 AvailableUomNo				nvarchar(120)		NOT NULL,						--TH056 計價單位 => 預設與進貨單位相同，暫不開放修改(採購需求)
 OrigCustomer				nvarchar(10)			NULL,						--TH057 原始客戶
 ApproveStatus				nvarchar(1)			NOT NULL	DEFAULT 'N',		--TH058 簽核狀態碼 0.待處理、1.簽核中、2.退件、3.已核准、4.取消確認中、5.作廢中、6.取消作廢中、N.不執行電子簽核[DEF:"N"] 
 EbcErpPreNo				nvarchar(25)			NULL,						--TH059 EBC出貨通知單號
 EbcEdition					nvarchar(4)				NULL,						--TH060 EBC出貨通知版次
 ProductSeqAmount			int					NOT NULL	DEFAULT 0,			--TH061 產品序號數量
 MtlItemType				nvarchar(1)			NOT	NULL	DEFAULT '2',		--TH062 品號類型 1.工程品號 2.正式品號[DEF:"2"]
 Loaction					nvarchar(10)			NULL,						--TH063 儲位
 TaxRate					float					NULL,						--TH073 營業稅率(%)[DEF:9.9999]&&89-06-01ADD//950808 S00-9508031 N5.4->N8.5
 MultipleFlag				nvarchar(1)			NOT	NULL	DEFAULT 'N',		--TH074 多儲批(不顯示，預設N)
 GrErpStatus				nvarchar(1)			NOT	NULL	DEFAULT '0',		--TH082 狀態 0.待處理 1.底稿中 2.傳票中 3.已處理[DEF:"0"]
 TaxCode					nvarchar(3)			NOT NULL,						--TH088 稅別碼
 DiscountRate				float				NOT NULL	DEFAULT 0.000000,	--TH089 折扣率(預設值與功用待確認)
 DiscountPrice				float				NOT NULL	DEFAULT 0.000000,	--TH090 折扣金額
 QcType						nvarchar(1)			NOT NULL	DEFAULT 'N',		--TH091 送測狀態 N.不處理、1.待送檢、2.送檢中、3.重新待送檢、4.重新送檢中、5.送檢完成、Y.確認通知完成[DEF:"N"]
 TransferStatus				nvarchar(2)			NOT NULL,						--拋轉ERP狀態 >> BAS.[Status]
 TransferDate				datetime				NULL,						--拋轉ERP時間
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_GrDetail PRIMARY KEY (GrDetailId),
 CONSTRAINT AK1_GrDetail UNIQUE (GrId, GrDetailId),
 CONSTRAINT AK2_GrDetail UNIQUE (GrId, GrSeq),
 CONSTRAINT FK1_GrDetail FOREIGN KEY (GrId) REFERENCES SCM.GoodsReceipt(GrId),
 CONSTRAINT FK2_GrDetail FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId),
 CONSTRAINT FK3_GrDetail FOREIGN KEY (UomId) REFERENCES PDM.UnitOfMeasure(UomId),
 CONSTRAINT FK4_GrDetail FOREIGN KEY (InventoryId) REFERENCES SCM.Inventory(InventoryId),
 CONSTRAINT FK5_GrDetail FOREIGN KEY (PoDetailId) REFERENCES SCM.PoDetail(PoDetailId)
)

CREATE TABLE SCM.GrBarcode( /*進貨單條碼資料*/
 GrBarcodeId				int					identity,
 GrDetailId					int					NOT NULL,						--進貨單單身ID
 BarcodeNo					nvarchar(60)		NOT NULL,						--條碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_GrBarcode PRIMARY KEY (GrBarcodeId),
 CONSTRAINT AK1_GrBarcode UNIQUE (GrDetailId, BarcodeNo),
 CONSTRAINT FK1_GrBarcode FOREIGN KEY (GrDetailId) REFERENCES SCM.GrDetail(GrDetailId),
 CONSTRAINT FK2_GrBarcode FOREIGN KEY (BarcodeNo) REFERENCES MES.Barcode(BarcodeNo),
)/*進貨單條碼資料*/
END

BEGIN /*採購*/

CREATE TABLE SCM.MaterialManagement( /*材料管理 */
 MtId						int					identity,				
 MaterialName				nvarchar(100)		NOT NULL,						--材料品名
 SupplierId					int					NOT NULL,						--廠商
 MtlItemId					int						NULL,						--ERP品號
 Remark						nvarchar(255)			NULL,						--備註
 Status						nvarchar(2)			NOT NULL	DEFAULT 'A',		--狀態(A/S)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MaterialManagement PRIMARY KEY (MtId),
 CONSTRAINT FK1_MaterialManagement FOREIGN KEY (SupplierId) REFERENCES SCM.Supplier(SupplierId),
 CONSTRAINT FK2_MaterialManagement FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId)
)/*材料管理 */

CREATE TABLE SCM.MtAmount( /*材料金額 */
 MaId						int					identity,				
 MtId						int					NOT	NULL,						--材料Id
 Amount						float				NOT NULL,						--單位金額(金額/KG)
 UnitQuantity				float				NOT NULL,						--單位量(KG/包)
 ConfirmDate				datetime				NULL,						--確認日
 ConfirmUserId				int						NULL,						--確認人員
 ConfirmStatus				nvarchar(2)			NOT NULL	DEFAULT 'N',		--確認碼(Y/N/H) H:歷史
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MtAmount PRIMARY KEY (MaId),
 CONSTRAINT FK1_MtAmount FOREIGN KEY (MtId) REFERENCES SCM.MaterialManagement(MtId),
 CONSTRAINT FK2_MtAmount FOREIGN KEY (ConfirmUserId) REFERENCES BAS.[User](UserId)
)/*材料金額 */

CREATE TABLE SCM.PriceConfirmationOrder( /*核價單頭*/ /*ERP PURTL*/
 PcId						int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 PcErpPrefix				nvarchar(4)			NOT NULL,						--TL001 採購單別
 PcErpNo					nvarchar(11)		NOT NULL,						--TL002 採購單身
 PcDate						datetime			NOT NULL,						--TL003 核價日期
 SupplierId					int					NOT NULL,						--TL004 供應商代號
 Currency					nvarchar(4)			NOT NULL,						--TL005 幣別
 ConfirmStatus				nvarchar(1)			NOT NULL,						--TL006 確認碼 Y/N/V
 Remark						nvarchar(255)			NULL,						--TL007 備註
 TaxIncluded				nvarchar(1)				NULL,						--TL008 含稅 Y/N
 DocDate					datetime				NULL,						--TL010 單據日期
 ConfirmUserId				int					NOT	NULL,						--TL011 確認者
 ApproveStatus				nvarchar(1)			NOT NULL	DEFAULT 'N',		--TL012 簽核狀態碼 0.待處理、1.簽核中、2.退件、3.已核准、4.取消確認中、5.作廢中、6.取消作廢中、N.不執行電子簽核[DEF:"N"]
 TradeTerm					nvarchar(1)			NOT NULL,						--TL019 交易條件 >> 參照 ERP交易條件代號檔(CMSNK)
 TransferStatus				nvarchar(2)			NOT NULL,						--拋轉ERP狀態 >> BAS.[Status]
 TransferDate				datetime				NULL,						--拋轉ERP時間
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_PriceConfirmationOrder PRIMARY KEY (PcId),
 CONSTRAINT AK1_PriceConfirmationOrder UNIQUE (CompanyId, PcErpPrefix, PcErpNo),
 CONSTRAINT FK1_PriceConfirmationOrder FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_PriceConfirmationOrder FOREIGN KEY (SupplierId) REFERENCES SCM.Supplier(SupplierId)
)

CREATE TABLE SCM.PcDetail( /*核價單身*/ /*ERP PURTM*/
 PcDetailId					int					identity,
 PcId						int					NOT NULL,						--核價單頭(TM001,TM002)
 PcSeq						nvarchar(4)			NOT NULL,						--TM003 核價單序號
 MtlItemId					int					NOT NULL,						--TM004 品號,品名,規格(TM004,TM005,TM006)
 SupplierMtlItemNo			nvarchar(40)			NULL,						--TM007 廠商品號
 TaxIncluded				nvarchar(1)			NOT	NULL,						--TM008 分量計價 Y/N
 PriceUomId					int					NOT	NULL,						--TM009	計價單位
 UnitPrice					float				NOT	NULL,						--TM010 單價
 ConfirmStatus				nvarchar(1)			NOT NULL,						--TM011 確認碼 Y/N/V
 Remark						nvarchar(255)			NULL,						--TM012 備註
 EffectiveDate				datetime			NOT	NULL,						--TM014 生效日期
 ExpirationDate				datetime				NULL,						--TM015 失效日期
 OrEffectiveDate			datetime				NULL,						--TM017 原生效日
 OrUnitPrice				float					NULL,						--TM018 原單價
 AdjustmentRange			float					NULL,						--TM019 調整幅度
 TransferStatus				nvarchar(2)			NOT NULL,						--拋轉ERP狀態 >> BAS.[Status]
 TransferDate				datetime				NULL,						--拋轉ERP時間
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_PcDetail PRIMARY KEY (PcDetailId),
 CONSTRAINT FK1_PcDetail FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId),
 CONSTRAINT FK2_PcDetail FOREIGN KEY (PriceUomId) REFERENCES PDM.UnitOfMeasure(UomId)
)


CREATE TABLE SCM.PurchaseOrder( /*採購單頭*/ /*ERP PURTC*/
 PoId						int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 PoErpPrefix				nvarchar(4)			NOT NULL,						--TC001 採購單別
 PoErpNo					nvarchar(11)		NOT NULL,						--TC002 採購單身
 PoDate						datetime			NOT NULL,						--TC003 採購日期
 SupplierId					int					NOT NULL,						--TC004 供應商
 CurrencyCode				nvarchar(4)			NOT NULL,						--TC005 幣別
 Exchange					float				NOT NULL	DEFAULT 0.0000,		--TC006 匯率
 PaymentTerm				nvarchar(16)		NOT NULL,						--TC008 付款條件
 Remark						nvarchar(255)			NULL,						--TC009 備註
 PoUserId					int					NOT	NULL,						--TC011 採購人員
 ConfirmStatus				nvarchar(1)			NOT NULL,						--TC014 確認碼 Y/N/V
 Taxation					nvarchar(1)			NOT NULL,						--TC018 課稅別  >> 由稅別碼決定(NN006) 1：內含；2：外加；3：零稅率；4：免稅；9：不計稅
 PoPrice					float				NOT NULL,						--TC019 採購金額
 TaxAmount					float				NOT NULL,						--TC020 稅額
 FirstAddress				nvarchar(255)			NULL,						--TC021 第一送貨地址
 SecondAddress				nvarchar(255)			NULL,						--TC022 第二送貨地址
 Quantity					int					NOT NULL	DEFAULT 0,			--TC023	數量合計(單身數量合計)
 DocDate					datetime			NOT NULL,						--TC024 單據日期
 ConfirmUserId				int						NULL,						--TC025 確認者
 TaxRate					float					NULL,						--TC026 營業稅率
 PaymentTermNo				nvarchar(6)			NOT NULL,						--TC027 付款條件代號
 ApproveStatus				nvarchar(1)			NOT NULL	DEFAULT 'N',		--TC030 簽核狀態碼 0.待處理、1.簽核中、2.退件、3.已核准、4.取消確認中、5.作廢中、6.取消作廢中、N.不執行電子簽核[DEF:"N"]
 Edition					nvarchar(4)			NOT NULL	DEFAULT '0000',		--TC039 版次
 DepositPartial				nvarchar(1)			NOT NULL,						--TC040 訂金分批 >> BAS.[Status] N(DEF)；Y
 TaxNo						nvarchar(3)			NOT NULL,						--TC047	稅別碼 >> 參照 ERP稅別碼設定檔(CMSNN)
 TradeTerm					nvarchar(1)			NOT NULL,						--TC048 交易條件 >> 參照 ERP交易條件代號檔(CMSNK)
 DetailMultiTax				nvarchar(1)			NOT NULL,						--TC051 單身多稅率 >> BAS.[Status] N(DEF)；Y
 ContactUser				nvarchar(30)			NULL,						--TC052 聯絡人
 TransferStatus				nvarchar(2)			NOT NULL,						--拋轉ERP狀態 >> BAS.[Status]
 TransferDate				datetime				NULL,						--拋轉ERP時間
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_PurchaseOrder PRIMARY KEY (PoId),
 CONSTRAINT AK1_PurchaseOrder UNIQUE (CompanyId, PoErpPrefix, PoErpNo),
 CONSTRAINT FK1_PurchaseOrder FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_PurchaseOrder FOREIGN KEY (SupplierId) REFERENCES SCM.Supplier(SupplierId),
 CONSTRAINT FK3_PurchaseOrder FOREIGN KEY (PoUserId) REFERENCES BAS.[User](UserId)
)

CREATE TABLE SCM.PoDetail( /*採購單身*/ /*ERP PURTD*/
 PoDetailId					int					identity,
 PoId						int					NOT NULL,						--採購單頭ID
 PoErpPrefix				nvarchar(4)			NOT NULL,						--TD001 採購單別
 PoErpNo					nvarchar(11)		NOT NULL,						--TD002 採購單身
 PoSeq						nvarchar(4)			NOT NULL,						--TD003 序號
 MtlItemId					int					NOT NULL,						--TD004 品號
 PoMtlItemName				nvarchar(120)		NOT NULL,						--TD005 採購品名
 PoMtlItemSpec				nvarchar(120)		NOT NULL,						--TD006 採購規格
 InventoryId				int					NOT NULL,						--TD007 交貨庫別
 Quantity					float				NOT NULL,						--TD008 採購數量
 UomId						int					NOT	NULL,						--TD009	單位(Unit of Measure) >> 參照 ERP品號換算單位檔(INVMD)
 PoUnitPrice				float				NOT	NULL,						--TD010 採購單價
 PoPrice					float				NOT	NULL,						--TD011 採購金額
 PromiseDate				datetime				NULL,						--TD012 預計交貨日
 ReferencePrefix			nvarchar(4)				NULL,						--TD013 參考單別
 Remark						nvarchar(255)			NULL,						--TD014 備註
 SiQty						float				NOT NULL,						--TD015	已交數量
 ClosureStatus				nvarchar(1)			NOT NULL,						--TD016 結案碼 >> BAS.[Status] N：未結案；y：指定結案；Y：自動結案
 ConfirmStatus				nvarchar(1)			NOT NULL,						--TD018 確認碼 Y/N/V
 InventoryQty				float				NOT	NULL	DEFAULT 0.000000,	--TD019 庫存數量
 SmallUnit					nvarchar(6)				NULL,						--TD020 小單位
 ReferenceNo				nvarchar(11)			NULL,						--TD021 參考單號
 Project					nvarchar(20)			NULL,						--TD022	專案代號 >> 參考 ERP專案代號檔(CMSNB)
 ReferenceSeq				nvarchar(4)				NULL,						--TD023	參考序號 >> 參考 ERP專案代號檔(CMSNB)
 UrgentMtl					nvarchar(1)			NOT NULL,						--TD025 急料 >> BAS.[Type] N[DEF]；Y
 PrErpPrefix				nvarchar(4)			NOT NULL,						--TD026 請購單別
 PrErpNo					nvarchar(11)		NOT NULL,						--TD027 請購單號
 PrSequence					nvarchar(4)			NOT NULL,						--TD028 請購序號
 FromDocType				nvarchar(1)			NOT	NULL	DEFAULT '1',		--TD034 來源 1.請購 2.LRP 3.MRP 4.訂單 5.合約採購 6.採購變更 9.其他[DEF:"9"]
 MtlItemType				nvarchar(1)			NOT	NULL	DEFAULT '2',		--TD038 品號類型 1.工程品號 2.正式品號[DEF:"2"]
 PartialSeq					nvarchar(4)				NULL,						--TD041 分批序號
 OriPromiseDate				datetime			    NULL,						--TD045	原預交日
 DeliveryDate				datetime			    NULL,						--TD046	交期確認日
 PoPriceQty					float				NOT NULL,						--TD058	計價數量
 PoPriceUomId				int						NULL,						--TD059	計價單位
 SiPriceQty					float					NULL,						--TD060	已交計價數量
 DiscountRate				float				NOT NULL	DEFAULT 0.000000,	--TD062	折扣率
 DiscountAmount				float				NOT NULL	DEFAULT 0.000000,	--TD063	折扣金額
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_PoDetail PRIMARY KEY (PoDetailId),
 CONSTRAINT AK1_PoDetail UNIQUE (PoId, PoSeq),
 CONSTRAINT FK1_PoDetail FOREIGN KEY (PoId) REFERENCES SCM.PurchaseOrder(PoId),
 CONSTRAINT FK2_PoDetail FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId),
 CONSTRAINT FK3_PoDetail FOREIGN KEY (InventoryId) REFERENCES SCM.Inventory(InventoryId),
 CONSTRAINT FK4_PoDetail FOREIGN KEY (UomId) REFERENCES PDM.UnitOfMeasure(UomId),
 CONSTRAINT FK5_PoDetail FOREIGN KEY (PoPriceUomId) REFERENCES PDM.UnitOfMeasure(UomId)
)

CREATE TABLE SCM.PoDeliveryDateLog( /*採購單交期更改紀錄*/
 LogId						int					identity,
 PoDetailId					int					NOT NULL,						--採購單身ID
 Edition					nvarchar(4)			NOT NULL,						--異動版次
 OriDeliveryDate			datetime			NOT NULL,						--原始交期
 DeliveryDate				datetime			NOT NULL,						--異動交期
 ChangeDate					datetime			NOT NULL,						--異動日期
 Remark						nvarchar(255)		NOT NULL,						--備註
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_PoDeliveryDateLog PRIMARY KEY (LogId),
 CONSTRAINT AK1_PoDeliveryDateLog UNIQUE (PoDetailId, Edition),
 CONSTRAINT FK1_PoDeliveryDateLog FOREIGN KEY (PoDetailId) REFERENCES SCM.PoDetail(PoDetailId),
) /*採購單交期更改紀錄*/
END

BEGIN /*庫存*/
CREATE TABLE SCM.InventoryTransaction( /*庫存異動*/ /*ERP INVTA*/
 ItId						int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 ItErpPrefix				nvarchar(4)			NOT NULL,						--TA001 單別 >> 參照 ERP各系統單據設定檔(CMSMQ)
 ItErpNo					nvarchar(11)		NOT NULL,						--TA002 單號
 ItDate						datetime			NOT NULL,						--TA003 異動日期
 DocDate					datetime			NOT NULL,						--TA014 單據日期
 DepartmentId				int						NULL,						--TA004 所屬部門
 Remark						nvarchar(255)			NULL,						--TA005 備註
 TotalQty					float					NULL,						--TA011 總數量
 Amount						float					NULL,						--TA012 總金額
 ConfirmStatus				nvarchar(1)			NOT NULL,						--TA006 確認碼　>> BAS.[Status] N：未確認；V：作廢；Y：已確認
 ConfirmUserId				int						NULL,						--TA015 確認者
 TransferStatus				nvarchar(2)			NOT NULL,						--拋轉ERP狀態 >> BAS.[Status]
 TransferDate				datetime				NULL,						--拋轉ERP時間
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_InventoryTransaction PRIMARY KEY (ItId),
 CONSTRAINT AK1_InventoryTransaction UNIQUE (CompanyId, ItErpPrefix, ItErpNo),
 CONSTRAINT FK1_InventoryTransaction FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_InventoryTransaction FOREIGN KEY (DepartmentId) REFERENCES BAS.Department(DepartmentId),
 CONSTRAINT FK3_InventoryTransaction FOREIGN KEY (ConfirmUserId) REFERENCES BAS.[User](UserId)
) /*庫存異動*/

CREATE TABLE SCM.ItDetail( /*庫存異動單身*/ /*ERP INVTB*/
 ItDetailId					int					identity,
 ItId						int					NOT NULL,						--所屬單頭
 ItSequence					nvarchar(4)			NOT NULL,						--TB003	流水號
 MtlItemId					int					NOT NULL,						--TB004 所屬品號
 ItMtlItemName				nvarchar(120)			NULL,						--TB005	品名
 ItMtlItemSpec				nvarchar(120)			NULL,						--TB006	規格
 ItQty						decimal(16,3)		NOT NULL,						--TB007 數量
 UomId						int					NOT NULL,						--TB008 單位
 InvQty						decimal(16,3)		NOT NULL,						--TB009 庫存數量
 UnitCost					decimal(21,6)		NOT NULL,						--TB010	單位成本 『品號資料建立作業』單頭的【庫存數量】/【庫存金額】=單位成本。
 Amount						decimal(21,6)		NOT NULL,						--TB011	金額
 InventoryId				int					NOT NULL,						--TB012 轉出庫
 ToInventoryId				int						NULL,						--TB012 轉入庫
 StorageLocation			nvarchar(120)			NULL,						--TB026 轉出儲位
 ToStorageLocation			nvarchar(120)			NULL,						--TB027 轉入儲位
 ItRemark					nvarchar(255)			NULL,						--TB017 備註
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ItDetail PRIMARY KEY (ItDetailId),
 CONSTRAINT AK1_ItDetail UNIQUE (ItId, ItSequence),
 CONSTRAINT FK1_ItDetail FOREIGN KEY (ItId) REFERENCES SCM.InventoryTransaction(ItId),
 CONSTRAINT FK2_ItDetail FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId),
 CONSTRAINT FK3_ItDetail FOREIGN KEY (InventoryId) REFERENCES SCM.Inventory(InventoryId),
 CONSTRAINT FK4_ItDetail FOREIGN KEY (ToInventoryId) REFERENCES SCM.Inventory(InventoryId),
 CONSTRAINT FK5_ItDetail FOREIGN KEY (UomId) REFERENCES PDM.UnitOfMeasure(UomId)
) /*庫存異動單身*/

CREATE TABLE SCM.TempShippingNote( /*暫出單單頭*/ /*ERP INVTF*/
 TsnId						int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 DoId						int						NULL,						--所屬出貨單
 SoId						int						NULL,						--所屬訂單
 TsnErpPrefix				nvarchar(4)			NOT NULL,						--TF001 單別 >> 參照 ERP各系統單據設定檔(CMSMQ)
 TsnErpNo					nvarchar(11)		NOT NULL,						--TF002 單號
 TsnDate					datetime			NOT NULL,						--TF003 暫出日期(出貨/借用) >> 自動帶入TF041 預計轉銷日
 DocDate					datetime			NOT NULL,						--TF024 單據日期
 ToObject					nvarchar(1)			NOT NULL,						--TF004 異動對象 >> BAS.[Type] 1.客戶,2.廠商,3.人員,9.其它
 ObjectCustomer				int						NULL,						--TF005 異動對象(客戶)
 ObjectSupplier				int						NULL,						--TF005 異動對象(廠商)
 ObjectUser					int						NULL,						--TF005 異動對象(人員)
 ObjectOther				nvarchar(30)			NULL,						--TF005 異動對象(其他)
 DepartmentId				int						NULL,						--TF007 所屬部門
 UserId						int					NOT NULL,						--TF008 所屬人員
 Remark						nvarchar(255)			NULL,						--TF014 單頭備註
 OtherRemark				nvarchar(255)			NULL,						--TF018 倉庫備註 --(其他備註)
 TotalQty					float					NULL,						--TF022 總數量
 Amount						float					NULL,						--TF023 總金額
 TaxAmount					float					NULL,						--TF027 總稅額
 ConfirmStatus				nvarchar(1)			NOT NULL,						--TF020 確認碼　>> BAS.[Status] N：未確認；V：作廢；Y：已確認
 ConfirmUserId				int						NULL,						--TF025 確認者
 TransferStatus				nvarchar(2)			NOT NULL,						--拋轉ERP狀態 >> BAS.[Status]
 TransferDate				datetime				NULL,						--拋轉ERP時間
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_TempShippingNote PRIMARY KEY (TsnId),
 CONSTRAINT AK1_TempShippingNote UNIQUE (CompanyId, TsnErpPrefix, TsnErpNo),
 CONSTRAINT FK1_TempShippingNote FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_TempShippingNote FOREIGN KEY (DoId) REFERENCES SCM.DeliveryOrder(DoId),
 CONSTRAINT FK3_TempShippingNote FOREIGN KEY (SoId) REFERENCES SCM.SaleOrder(SoId),
 CONSTRAINT FK4_TempShippingNote FOREIGN KEY (ObjectCustomer) REFERENCES SCM.Customer(CustomerId),
 CONSTRAINT FK5_TempShippingNote FOREIGN KEY (ObjectSupplier) REFERENCES SCM.Supplier(SupplierId),
 CONSTRAINT FK6_TempShippingNote FOREIGN KEY (ObjectUser) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK7_TempShippingNote FOREIGN KEY (DepartmentId) REFERENCES BAS.Department(DepartmentId),
 CONSTRAINT FK8_TempShippingNote FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK9_TempShippingNote FOREIGN KEY (ConfirmUserId) REFERENCES BAS.[User](UserId)
) /*暫出單單頭*/

CREATE TABLE SCM.TsnDetail( /*暫出單單身*/ /*ERP INVTG*/
 TsnDetailId				int					identity,
 TsnId						int					NOT NULL,						--所屬單頭
 TsnSequence				nvarchar(4)			NOT NULL,						--TG003	流水號
 MtlItemId					int					NOT NULL,						--TG004 所屬品號
 TsnMtlItemName				nvarchar(120)			NULL,						--TG005	品名
 TsnMtlItemSpec				nvarchar(120)			NULL,						--TG006	規格
 TsnOutInventory			int						NULL,						--TG007 轉出庫別
 TsnInInventory				int						NULL,						--TG008 轉入庫別
 TsnQty						decimal(16,3)		NOT NULL,						--TG009 數量
 ProductType				nvarchar(1)				NULL	DEFAULT '1',		--TG043 類型 1.贈品量、2.備品量
 FreebieOrSpareQty			float					NULL,						--TG044	贈/備品量
 UnitPrice					decimal(21,6)			NULL,						--TG012 單價
 Amount						decimal(21,6)			NULL,						--TG013 金額
 SoDetailId					int						NULL,						--TG014/15/16 訂單單身(單別/單號/流水號)
 LotNumber					nvarchar(20)			NULL,						--TG017 批號
 TsnPriceQty				float					NULL,						--TG052	計價數量
 TsnPriceUomId				int						NULL,						--TG053	計價單位
 BusinessTaxRate			float					NULL,						--TG042	營業稅率 >> 由稅別碼決定
 TsnRemark					nvarchar(255)			NULL,						--TG019 單身備註
 ExpectedReturnDate			datetime				NULL,						--TG027 預計歸還日
 SaleQty					decimal(16,3)		NOT NULL,						--TG020 轉進銷量
 SaleFSQty					decimal(16,3)			NULL,						--TG046 轉銷贈/備品量
 ReturnQty					decimal(16,3)		NOT NULL,						--TG021 歸還量
 ReturnFSQty				decimal(16,3)			NULL,						--TG048 歸還贈/備品量
 ConfirmStatus				nvarchar(1)			NOT NULL,						--TG022 確認碼 >> BAS.[Status] N：未確認；V：作廢；Y：已確認
 ClosureStatus				nvarchar(1)			NOT NULL,						--TG024 結案碼 >> BAS.[Status] N：未結案；y：指定結案；Y：自動結案
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_TsnDetail PRIMARY KEY (TsnDetailId),
 CONSTRAINT AK1_TsnDetail UNIQUE (TsnId, TsnSequence),
 CONSTRAINT FK1_TsnDetail FOREIGN KEY (TsnId) REFERENCES SCM.TempShippingNote(TsnId),
 CONSTRAINT FK2_TsnDetail FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId),
 CONSTRAINT FK3_TsnDetail FOREIGN KEY (TsnOutInventory) REFERENCES SCM.Inventory(InventoryId),
 CONSTRAINT FK4_TsnDetail FOREIGN KEY (TsnInInventory) REFERENCES SCM.Inventory(InventoryId),
 CONSTRAINT FK5_TsnDetail FOREIGN KEY (SoDetailId) REFERENCES SCM.SoDetail(SoDetailId)
) /*暫出單單身*/

CREATE TABLE SCM.TempShippingReturnNote( /*暫出歸還單單頭*/ /*ERP INVTH*/
 TsrnId						int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 TsrnErpPrefix				nvarchar(4)			NOT NULL,						--TH001 單別 >> 參照 ERP各系統單據設定檔(CMSMQ)
 TsrnErpNo					nvarchar(11)		NOT NULL,						--TH002 單號
 TsrnDate					datetime			NOT NULL,						--TH003 暫出日期(出貨/借用) >> 自動帶入TF041 預計轉銷日
 ToObject					nvarchar(1)			NOT NULL,						--TH004 異動對象 >> BAS.[Type] 1.客戶,2.廠商,3.人員,9.其它
 ObjectCustomer				int						NULL,						--TH005 異動對象(客戶)
 ObjectSupplier				int						NULL,						--TH005 異動對象(廠商)
 ObjectUser					int						NULL,						--TH005 異動對象(人員)
 ObjectOther				nvarchar(30)			NULL,						--TH005 異動對象(其他)
 DepartmentId				int						NULL,						--TH007 所屬部門
 UserId						int					NOT NULL,						--TH008 所屬人員
 Factory					nvarchar(6)			NOT NULL,						--TH009 廠別
 TaxType					nvarchar(1)			NOT NULL,						--TH010 課稅別   
 Currency					nvarchar(4)			NOT NULL,						--TH011 幣別        
 ExchangeRate				float				NOT NULL,						--TH012 匯率
 RowCnt						int						NULL,						--TH013 件數
 Remark						nvarchar(255)			NULL,						--TH014 單頭備註
 CustomerAddressFirst		nvarchar(255)		NOT NULL,						--TH016 地址一        
 CustomerAddressSecond		nvarchar(255)		NOT NULL,						--TH017 地址二   
 OtherRemark				nvarchar(255)			NULL,						--TH018 其他備註
 ConfirmStatus				nvarchar(1)			NOT NULL,						--TH020 確認碼　>> BAS.[Status] N：未確認；V：作廢；Y：已確認
 TotalQty					float					NULL,						--TH021 總數量
 Amount						float					NULL,						--TH022 總金額
 DocDate					datetime			NOT NULL,						--TH023 單據日期
 ConfirmUserId				int						NULL,						--TH024 確認者
 TaxRate					float				NOT	NULL	DEFAULT 0.0000,		--TH025 營業稅率         
 TaxAmount					float					NULL,						--TH026 總稅額
 TransportMethod			nvarchar(1)				NULL,						--TG030 運輸方式         
 TaxCode					nvarchar(3)			NOT	NULL,						--TH038 稅別碼         
 MultiTaxRate				nvarchar(1)				NULL,						--TH039 單身多稅率         
 TransferStatus				nvarchar(2)			NOT NULL,						--拋轉ERP狀態 >> BAS.[Status]
 TransferDate				datetime				NULL,						--拋轉ERP時間
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_TempShippingReturnNote PRIMARY KEY (TsrnId),
 CONSTRAINT AK1_TempShippingReturnNote UNIQUE (CompanyId, TsrnErpPrefix, TsrnErpNo),
 CONSTRAINT FK1_TempShippingReturnNote FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_TempShippingReturnNote FOREIGN KEY (ObjectCustomer) REFERENCES SCM.Customer(CustomerId),
 CONSTRAINT FK3_TempShippingReturnNote FOREIGN KEY (ObjectSupplier) REFERENCES SCM.Supplier(SupplierId),
 CONSTRAINT FK4_TempShippingReturnNote FOREIGN KEY (ObjectUser) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK5_TempShippingReturnNote FOREIGN KEY (DepartmentId) REFERENCES BAS.Department(DepartmentId),
 CONSTRAINT FK6_TempShippingReturnNote FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK7_TempShippingReturnNote FOREIGN KEY (ConfirmUserId) REFERENCES BAS.[User](UserId)
) /*暫出歸還單單頭*/

CREATE TABLE SCM.TsrnDetail( /*暫出歸還單單身*/ /*ERP INVTI*/
 TsrnDetailId				int					identity,
 TsrnId						int					NOT NULL,						--所屬單頭
 TsrnSequence				nvarchar(4)			NOT NULL,						--TI003	流水號
 MtlItemId					int					NOT NULL,						--TI004 所屬品號 / TI005 品名 / TI006 規格
 TsrnMtlItemName			nvarchar(120)			NULL,						--TI005	品名
 TsrnMtlItemSpec			nvarchar(120)			NULL,						--TI006	規格
 TsrnOutInventory			int						NULL,						--TI007 轉出庫別
 TsrnInInventory			int						NULL,						--TI008 轉入庫別
 TsrnQty					decimal(16,3)		NOT NULL,						--TI009 數量
 UomId						int					NOT	NULL,						--TI010	單位
 UnitPrice					decimal(21,6)			NULL,						--TI012 單價
 Amount						decimal(21,6)			NULL,						--TI013 金額
 TsnDetailId				int					NOT NULL,						--TI014 暫出單別 / TI015 暫出單號 / TI016 暫出單序號 
 LotNumber					nvarchar(20)			NULL,						--TI017	批號
 AvailableDate				datetime				NULL,						--TI018 有效日期
 ReCheckDate				datetime				NULL,						--TI019 複檢日期
 ProjectCode				nvarchar(20)			NULL,						--TI020 專案代碼
 TsrnRemark					nvarchar(255)			NULL,						--TI021 單身備註
 ConfirmStatus				nvarchar(1)			NOT NULL,						--TI022 確認碼 >> BAS.[Status] N：未確認；V：作廢；Y：已確認
 OutLocation				nvarchar(10)			NULL,						--TI026 轉出儲位
 IntLocation				nvarchar(10)			NULL,						--TI027 轉入儲位
 TaxRate					float				NOT	NULL	DEFAULT 0.0000,		--TI033 營業稅率         
 ProductType				nvarchar(1)				NULL	DEFAULT '1',		--TI034 類型 1.贈品量、2.備品量
 FreebieOrSpareQty			float					NULL,						--TI035	贈/備品量
 TaxCode					nvarchar(3)			NOT	NULL,						--TI037 稅別碼         
 TsrnPriceQty				float					NULL,						--TI038	計價數量
 TsrnPriceUomId				int						NULL,						--TI039	計價單位
 BusinessTaxRate			float					NULL,						--TI033	營業稅率 >> 由稅別碼決定
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_TsrnDetail PRIMARY KEY (TsrnDetailId),
 CONSTRAINT AK1_TsrnDetail UNIQUE (TsrnId, TsrnSequence),
 CONSTRAINT FK1_TsrnDetail FOREIGN KEY (TsrnId) REFERENCES SCM.TempShippingReturnNote(TsrnId),
 CONSTRAINT FK2_TsrnDetail FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId),
 CONSTRAINT FK3_TsrnDetail FOREIGN KEY (TsrnOutInventory) REFERENCES SCM.Inventory(InventoryId),
 CONSTRAINT FK4_TsrnDetail FOREIGN KEY (TsrnInInventory) REFERENCES SCM.Inventory(InventoryId),
 CONSTRAINT FK5_TsrnDetail FOREIGN KEY (TsnDetailId) REFERENCES SCM.TsnDetail(TsnDetailId)
) /*暫出歸還單單身*/
END /*庫存*/

BEGIN /*退(換)貨*/
CREATE TABLE SCM.ReturnMerchandiseAuthorization( /*退(換)貨*/
 RmaId						int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 RmaNo						nvarchar(50)		NOT NULL,						--退貨單號
 CustomerId					int						NULL,						--客戶
 RmaDate					datetime			NOT NULL,						--退貨日期
 RmaType					nvarchar(2)			NOT NULL,						--退貨種類 >> BAS.[Type] 1：折讓(不退貨)；2：銷退；3：退換貨
 RmaRemark					nvarchar(200)			NULL,						--備註
 DocType					nvarchar(1)			NOT NULL,						--開立單據類型 >> BAS.[Type] A：暫出歸還單；B：客供入料單；C：0成本入庫單
 ConfirmStatus				nvarchar(1)			NOT NULL,						--確認碼 >> BAS.[Status] N：未確認；Y：已確認
 ConfirmUserId				int						NULL,						--確認者
 TransferStatus				nvarchar(2)			NOT NULL,						--拋轉ERP狀態 >> BAS.[Status]
 TransferDate				datetime				NULL,						--拋轉ERP時間
 TransferMode				nvarchar(2)			NOT NULL,						--拋轉模式 >> BAS.[Type] N：開立新單據；U：更新原單據
 ErpPrefix					nvarchar(4)				NULL,						--ERP單別
 ErpNo						nvarchar(11)			NULL,						--ERP單號
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ReturnMerchandiseAuthorization PRIMARY KEY (RmaId),
 CONSTRAINT AK1_ReturnMerchandiseAuthorization UNIQUE (RmaNo),
 CONSTRAINT FK1_ReturnMerchandiseAuthorization FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_ReturnMerchandiseAuthorization FOREIGN KEY (CustomerId) REFERENCES SCM.Customer(CustomerId)
) /*退(換)貨*/

CREATE TABLE SCM.RmaDetail( /*退(換)貨項目*/
 RmaDetailId				int					identity,
 RmaId						int					NOT NULL,						--所屬退(換)貨
 MtlItemId					int					NOT NULL,						--所屬品號
 ItemName					nvarchar(50)		NOT NULL,						--部品名稱
 RmaQty						int					NOT NULL,						--數量
 ItemType					nvarchar(1)				NULL	DEFAULT '1',		--物件類型
 FreebieOrSpareQty			float					NULL,						--贈/備品量
 RmaDesc					nvarchar(200)			NULL,						--不良項目描述
 TsnDetailId				int						NULL,						--所屬暫出單身
 TargetInventory			int						NULL,						--目標倉別
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_RmaDetail PRIMARY KEY (RmaDetailId),
 CONSTRAINT FK1_RmaDetail FOREIGN KEY (RmaId) REFERENCES SCM.ReturnMerchandiseAuthorization(RmaId),
 CONSTRAINT FK2_RmaDetail FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId),
 CONSTRAINT FK3_RmaDetail FOREIGN KEY (TsnDetailId) REFERENCES SCM.TsnDetail(TsnDetailId),
 CONSTRAINT FK4_RmaDetail FOREIGN KEY (TargetInventory) REFERENCES SCM.Inventory(InventoryId)
) /*退(換)貨項目*/

CREATE TABLE SCM.RmaTransferLog( /*退(換)貨拋轉單據紀錄*/
 RmaTransferLogId			int					identity,
 RmaId						int					NOT NULL,						--所屬退(換)貨
 ErpPrefix					nvarchar(4)				NULL,						--ERP單別
 ErpNo						nvarchar(11)			NULL,						--ERP單號
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼	
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 CreateBy					int					NOT NULL,						--新增人員
 CONSTRAINT PK_RmaTransferLog PRIMARY KEY (RmaTransferLogId),
 CONSTRAINT FK1_RmaTransferLog FOREIGN KEY (RmaId) REFERENCES SCM.ReturnMerchandiseAuthorization(RmaId),
) /*退(換)貨拋轉單據紀錄*/

CREATE TABLE SCM.RmaItem( /*退(換)貨物件*/
 RmaItemId					int					identity,
 RmaDetailId				int					NOT NULL,						--所屬項目
 BarcodeId					int					NOT NULL,						--所屬條碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_RmaItem PRIMARY KEY (RmaItemId),
 CONSTRAINT FK1_RmaItem FOREIGN KEY (RmaDetailId) REFERENCES SCM.RmaDetail(RmaDetailId),
 CONSTRAINT FK2_RmaItem FOREIGN KEY (BarcodeId) REFERENCES MES.Barcode(BarcodeId)
) /*退(換)貨物件*/
END /*退(換)貨*/

BEGIN/*需求單管理*/
CREATE TABLE SCM.Demand( /*需求單 單頭*/
 DemandId					int					identity,						
 CompanyId					int					NOT NULL,						--所屬公司
 DemandNo					nvarchar(50)		NOT NULL,						--需求單編號
 DemandSource				nvarchar(2)			NOT NULL,						--單據來源 SELECT * FROM BAS.Type WHERE TypeSchema = 'DemandSource' 
 DemandDesc					nvarchar(200)		NOT NULL,						--描述 (同於Mail主旨、電商的訂單內容) 
 CustNo						nvarchar(50)		NOT NULL,						--客戶編號	--SELECT * FROM SCM.Customer
 DemandStatus				nvarchar(30)		NOT NULL    DEFAULT '0',		--單據狀態	--SELECT * FROM BAS.[Status] WHERE StatusSchema = 'DemandStatus'
 Remark						nvarchar(30)		NOT	NULL,						--備註
 DemandDate					datetime			NOT NULL,						--提出日
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Demand PRIMARY KEY (DemandId),
 CONSTRAINT AK1_Demand UNIQUE (DemandNo),
 CONSTRAINT FK1_Demand FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)

CREATE TABLE SCM.DemandLine(/*需求單 單身*/
 DemandLineId 				int					identity,                       --ID
 DemandLineNo 				nvarchar(50)		NOT NULL,                       --需求單單身編號
 DemandId    				int					NOT NULL,                       --所屬Demand
 ProdTypeId                 int					NOT NULL,                       --產品結構類別
 TypeNo                     nvarchar(10)		NOT NULL,                       --產品屬性			 --SELECT MA001,LTRIM(RTRIM(MA002)) TypeNo,LTRIM(RTRIM(MA003)) TypeName FROM INVMA WHERE 1=1 AND MA001=4
 WipCategory                nvarchar(10)		NOT NULL,                       --製令單別
 CustMtlName                nvarchar(50)		NOT NULL,                       --客戶部番
 DeliveryDate				datetime			NOT NULL,                       --要望交期
 UnitPrice                  float				NOT NULL,                       --單價
 UomId                      int					NOT NULL,                       --單位				  --SELECT * FROM PDM.UnitOfMeasure
 OrderQty                   float				NOT NULL,                       --訂單數
 PlanQty					int					NOT NULL,						--預計產量
 OrderType	                nvarchar(2)			NOT NULL,                       --首/複製番			  --SELECT * FROM BAS.[Type] WHERE TABLE_NAME='DemandOrderType'
 MtlSpec                    nvarchar(50)		NOT NULL,                       --規格 
 DeliveryProcess		    nvarchar(2)			NULL,                           --出貨工程			  --SELECT * FROM BAS.[Type] WHERE TypeSchema='ManufactureOrder.DeliveryProcess'
 Remark                     nvarchar(200)   	NULL,                           --單據單身備註
 ProcessStatus				nvarchar(2)			NOT NULL    DEFAULT '0',        --單據流程狀態		  --SELECT * FROM BAS.[Status] WHERE StatusSchema = 'DemandProcessStatus'
 DemandStatus				nvarchar(30)		NOT NULL    DEFAULT '0',		--單據狀態			  --SELECT * FROM BAS.[Status] WHERE StatusSchema = 'DemandStatus'
 SoId	    				int					NULL		DEFAULT -1,         --訂單單頭ID			  --SELECT * FROM SCM.SaleOrder
 MtlItemId                  int					NULL		DEFAULT -1,         --品號ID				  --SELECT * FROM PDM.MtlItem 
 BomId		                int					NULL		DEFAULT -1,         --BOM主件ID			  --SELECT * FROM PDM.BillOfMaterials
 RountingItemId             int					NULL		DEFAULT -1,         --途程品號			  --SELECT * FROM MES.RoutingItem
 WipId                      int					NULL		DEFAULT -1,         --ERP工單ID			  --SELECT * FROM MES.WipOrder
 MoId	                    int					NULL		DEFAULT -1,         --MES製令ID			  --SELECT * FROM MES.ManufactureOrder
 MrId	                    int					NULL		DEFAULT -1,         --領料單ID			  --SELECT * FROM MES.MaterialRequisition
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime			NULL,						    --修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int					NULL,						    --修改人員
 CONSTRAINT PK_DemandLine PRIMARY KEY(DemandLineId),
 CONSTRAINT AK1_DemandLine UNIQUE (DemandLineNo),
 CONSTRAINT FK1_DemandLine FOREIGN KEY (DemandId) REFERENCES SCM.Demand(DemandId)
)

CREATE TABLE SCM.DemandProcess(/*需求單執行歷程資料*/
 DemandProcessId		    int					identity,                       --ID
 DemandLineNo   		    nvarchar(50)		NOT NULL,                       --所屬DemandLine
 ProcessStatus              nvarchar(2)			NOT NULL    DEFAULT '0',        --單據流程狀態 --SELECT * FROM BAS.[Status] WHERE StatusSchema = 'DemandProcessStatus' 
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime			NULL,						    --修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int					NULL,						    --修改人員
 CONSTRAINT PK_DemandProcess PRIMARY KEY(DemandProcessId),
 CONSTRAINT FK_DemandProcess FOREIGN KEY(DemandLineNo) REFERENCES SCM.DemandLine(DemandLineNo)
)

CREATE TABLE SCM.DemandProductType( /*需求單 品號命名原則*/
 ProductTypeId				int					identity,						
 CompanyId					int					NOT NULL,						--所屬公司
 ProductTypeName			nvarchar(200)		NOT NULL,						--產品類別名稱
 MadeStatus					nvarchar(10)		NOT NULL	DEFAULT 'M',		--M: 自製, P: 採購 
 RuleMtlitemNo				nvarchar(500)		NOT NULL,						--品號規則
 CustomerItemRequired		nvarchar(1)			NOT NULL	DEFAULT 'N',		--客戶品號必填 Y/N
 WipCategory				nvarchar(10)			NULL,						--製令單別
 CargoNo					nvarchar(40)			NULL,						--貨號
 Edition					nvarchar(4)				NULL,						--版次
 Unit						nvarchar(6)			NOT NULL,						--單位
 TypeOne					nvarchar(6)			NOT NULL,						--會計代碼
 TypeTwo					nvarchar(6)			NOT NULL,						--倉管代碼
 TypeThree					nvarchar(6)			NOT NULL,						--業務代碼
 TypeFour					nvarchar(6)			NOT NULL,						--產品類別代碼
 WarehouseNo				nvarchar(10)		NOT NULL,						--主要庫別代碼
 WarehousePickNo			nvarchar(10)		NOT NULL,						--主要領料庫別代碼
 PurchaseUnit				nvarchar(6)			NOT NULL,						--採購單位
 SalesUnit					nvarchar(6)			NOT NULL,						--銷售單位
 StockManagement			nvarchar(1)			NOT NULL	DEFAULT 'N',		--庫存管理 Y/N
 BondedGoods				nvarchar(1)			NOT NULL	DEFAULT 'N',		--保稅品 Y/N
 BatchNumberManagement		nvarchar(1)			NOT NULL	DEFAULT 'N',		--批號管理 N: 不需要, Y: 需要不檢查庫存量, W: 僅需警告, T: 需要且檢查庫存量
 Attributes					nvarchar(1)			NOT NULL	DEFAULT 'P',		--品號屬性 P: 採購件, M: 自製件, S: 託外加工件, Y: 虛設品號, F: Feature件, O: Option件
 Remark						nvarchar(255)		NOT NULL,						--備註
 TestMethod					nvarchar(1)			NOT NULL	DEFAULT '4',		--檢驗方式 0: 免檢, 1: 抽檢(減量), 2: 抽檢(正常), 3: 抽檢(加嚴), 4: 全檢
 ProductDescription			nvarchar(255)		NOT NULL,						--商品描述
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DemandProductType PRIMARY KEY (ProductTypeId),
 CONSTRAINT FK1_DemandProductType FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)
END/*需求單管理*/

BEGIN/*銷貨單管理*/

CREATE TABLE SCM.ReceiveOrder( /*ERP銷貨單 COPTG*/
 RoId								int					identity,
 CompanyId							int					NOT NULL,						--✓ 所屬公司
 RoErpPrefix						nvarchar(4)			NOT NULL,						--✓　TG001 單別         
 RoErpNo							nvarchar(11)		NOT NULL,						--✓　TG002 單號         
 ReceiveDate						datetime			NOT NULL,						--✓　TG003 銷貨日期         
 CustomerId							int					NOT NULL,						--✓　TG004 客戶代號         
 DepartmentId						int					    NULL,						--✓　TG005 部門         
 SalesmenId							int					    NULL,						--✓　TG006 業務員         
 CustomerFullName					nvarchar(80)			NULL,						--✓　TG007 客戶全名         
 CustomerAddressFirst				nvarchar(255)			NULL,						--✓　TG008 送貨地址一         
 CustomerAddressSecond				nvarchar(255)			NULL,						--✓　TG009 送貨地址二         
 Factory							nvarchar(6)			NOT NULL,						--✓　TG010 出貨廠別         
 Currency							nvarchar(4) 		NOT	NULL,						--✓　TG011 交易幣別 >> 參照 ERP幣別匯率檔單頭(CMSMF)
 ExchangeRate						float				NOT	NULL	DEFAULT 0.0000,		--✓　TG012 匯率 >> 參照 ERP幣別匯率檔單身(CMSMG)　(須大於0)
 OriginalAmount						float				NOT	NULL	DEFAULT 0.0000,		--✓　TG013 原幣銷貨金額 (預設0)
 InvNumEnd							nvarchar(50)			NULL,						--✓　TG014 發票號碼(訖)         
 UiNo								nvarchar(20)			NULL,						--✓　TG015 統一編號
 InvoiceType						nvarchar(1)				NULL,						--✓　TG016 發票聯數 1.二聯式、2.三聯式、3.二聯式收銀機發票、4.三聯式收銀機發票、5.電子計算機發票、6.免用統一發票、A.農產品收購憑證、G.海關代徵完稅憑證、N.不可抵扣專用發票、S.可抵扣專用發票、T.運輸發票、W.廢舊物資收購憑證、Z.其他
 TaxType							nvarchar(1)				NULL,						--✓　TG017 課稅別 1.應稅內含、2.應稅外加、3.零稅率、4.免稅、9.不計稅
 InvoiceAddressFirst				nvarchar(255)			NULL,						--✓　TG018 發票地址(一)
 InvoiceAddressSecond				nvarchar(255)			NULL,						--✓　TG019 發票地址(二)
 Remark								nvarchar(255)			NULL,						--✓　TG020 備註       
 InvoiceDate						datetime				NULL,						--✓　TG021 發票日期         
 PrintCount							int					    NULL,						--✓　TG022 列印次數         
 ConfirmStatus						nvarchar(1)			NOT NULL	DEFAULT 'N',		--✓　TG023 確認碼         
 UpdateCode							nvarchar(1)			NOT NULL	DEFAULT 'N',		--　 TG024 更新碼         
 OriginalTaxAmount					float				NOT	NULL	DEFAULT 0.0000,		--✓　TG025 原幣銷貨稅額  (預設0)        
 CollectionSalesmenId				int					    NULL,						--✓　TG026 收款業務員         
 Remark1							nvarchar(20)			NULL,						--✓　TG027 備註一         
 Remark2							nvarchar(20)			NULL,						--✓　TG028 備註二         
 Remark3							nvarchar(20)			NULL,						--✓　TG029 備註三         
 InvoicesVoid						nvarchar(1)				NULL,						--✓　TG030 發票作廢         
 CustomsClearance					nvarchar(1)				NULL,						--✓　TG031 通關方式         
 RowCnt								int						NULL,						--✓　TG032 件數         
 TotalQuantity						float				NOT	NULL	DEFAULT 0.0000,		--✓　TG033 總數量         
 CashSales							nvarchar(1)				NULL,						--✓　TG034 現銷         
 StaffId							int						NULL,						--✓　TG035 員工代號         
 RevenueJournalCode					nvarchar(1)				NULL,						--✓　TG036 產生分錄碼(收入)         
 CostJournalCode					nvarchar(1)				NULL,						--✓　TG037 產生分錄碼(成本)         
 ApplyYYMM							nvarchar(6)			NOT NULL,						--✓　TG038 申報年月         
 LCNO								nvarchar(20)			NULL,						--✓　TG039 L/C_NO         
 INVOICENO							nvarchar(50)			NULL,						--✓　TG040 INVOICE_NO         
 InvoicesPrintCount					int						NULL,						--✓　TG041 發票列印次數         
 DocDate							datetime			NOT NULL,						--✓　TG042 單據日期         
 ConfirmUserId						int						NULL,						--✓　TG043 確認者         
 TaxRate							float				NOT	NULL	DEFAULT 0.0000,		--✓　TG044 營業稅率         
 PretaxAmount						float				NOT	NULL	DEFAULT 0.0000,		--✓　TG045 本幣銷貨金額         
 TaxAmount							float				NOT	NULL	DEFAULT 0.0000,		--✓　TG046 本幣銷貨稅額         
 PaymentTerm						nvarchar(6)				NULL,						--✓　TG047 付款條件代號         
 SoDetailId							int						NULL,						--✓　TG048/49 訂單單別         
 AdvOrderPrefix						nvarchar(4)				NULL,						--✓　TG050 預收待抵單別         
 AdvOrderNo							nvarchar(11)			NULL,						--✓　TG051 預收待抵單號         
 OffsetAmount						float				NOT	NULL	DEFAULT 0.0000,		--✓　TG052 沖抵金額         
 OffsetTaxAmount					float				NOT	NULL	DEFAULT 0.0000,		--✓　TG053 沖抵稅額         
 TotalPackages						float					NULL,						--　 TG054 總包裝數量         
 SignatureStatus					nvarchar(1)				NULL,						--✓　TG055 簽核狀態碼         
 ChangeInvCode						nvarchar(1)				NULL,						--✓　TG056 更換發票碼         
 NewRoPrefix						nvarchar(4)				NULL,						--✓　TG057 新銷貨單別         
 NewRoNo							nvarchar(11)			NULL,						--✓　TG058 新銷貨單號         
 TransferStatusERP					nvarchar(1)				NULL,						--　 TG059 拋轉狀態(ERP自己的)         
 ProcessCode						nvarchar(2)				NULL,						--✓　TG060 流程代號(若單據性質設定為「28.多角貿易銷貨單」時，此欄位不可空白。)    
 AttachInvWithShip					nvarchar(1)				NULL,						--✓　TG061 隨貨附發票         
 BondCode							nvarchar(1)				NULL,						--✓　TG062 保稅碼         
 TransmissionCount					int						NULL,						--✓　TG063 傳送次數         
 Invoicer							nvarchar(10)			NULL,						--　 TG064 開票人         
 InvCode							nvarchar(8)				NULL,						--　 TG065 發票代號         
 ContactPerson						nvarchar(30)			NULL,						--✓　TG066 連絡人         
 Courier							nvarchar(10)			NULL,						--✓　TG067 代送商代號         
 SiteCommCalcMethod					nvarchar(1)				NULL,						--✓　TG068 營站佣金計算方式         
 SiteCommRate						float				NOT	NULL	DEFAULT 0.0000,		--✓　TG069 營站佣金比率         
 CommCalcInclTax					nvarchar(1)				NULL,						--✓　TG070 佣金計算含稅否         
 TotalCommAmount					float				NOT	NULL	DEFAULT 0.0000,		--✓　TG071 佣金總金額         
 TransportMethod					nvarchar(1)				NULL,						--✓　TG072 運輸方式         
 DispatchOrderPrefix				nvarchar(4)				NULL,						--✓　TG073 派車單別         
 DispatchOrderNo					nvarchar(11)			NULL,						--✓　TG074 派車單號   車號,車次,送貨者 顯示欄位回寫
 DeclarationNumber					nvarchar(20)			NULL,						--✓　TG075 報單號碼         
 FullNameOfDelivCustomer			nvarchar(80)			NULL,						--✓　TG076 送貨客戶全名         
 RoPriceType						nvarchar(30)			NULL,						--　 TG077 銷貨單單價別         
 TelephoneNumber					nvarchar(20)			NULL,						--✓　TG078 TEL_NO         
 FaxNumber							nvarchar(20)			NULL,						--✓　TG079 FAX_NO         
 ShipNoticePrefix					nvarchar(4)				NULL,						--✓　TG080 出貨通知單別         
 ShipNoticeNo						nvarchar(11)			NULL,						--✓　TG081 出貨通知單號         
 TradingTerms						nvarchar(1)				NULL,						--　 TG082 交易條件         
 CustomerEgFullName					nvarchar(80)			NULL,						--✓　TG083 客戶英文全名         
 --ReservedField					float				NOT	NULL	DEFAULT 0.0000,		--　 TG084 預留欄位         
 --ReservedField					float				NOT	NULL	DEFAULT 0.0000,,	--　 TG085 預留欄位         
 InvNumGenMethod					nvarchar(1)				NULL,						--✓　TG086 發票號碼產生方式         
 DocSourceCode						nvarchar(30)			NULL,						--　 TG087 單據來源碼         
 --ReservedField					nvarchar(60)			NULL,						--　 TG088 預留欄位         
 NoCredLimitControl					nvarchar(1)				NULL,						--　 TG089 不控管信用額度         
 InstallmentSettlement				nvarchar(1)				NULL,						--　 TG090 分期結帳         
 InstallmentCount					int						NULL,						--　 TG091 期數(2位數)         
 AutoApportionByInstallment			nvarchar(1)				NULL,						--　 TG092 依期數自動分攤         
 StartYearMonth						nvarchar(6)				NULL,						--　 TG093 起始年月         
 TaxCode							nvarchar(3)			NOT	NULL,						--✓　TG094 稅別碼         
 CustomsManual						nvarchar(15)			NULL,						--X　TG095 海關手冊         
 RemarkCode							nvarchar(20)			NULL,						--X　TG096 註記號         
 MultipleInvoices					nvarchar(1)				NULL,						--✓　TG097 多發票         
 InvNumStart						nvarchar(50)			NULL,						--✓　TG098 發票號碼(起)         
 NumberOfInvoices					int						NULL,						--　 TG099 發票張數 (2位數)        
 MultiTaxRate						nvarchar(1)				NULL,						--✓　TG100 單身多稅率         
 TaxCurrencyRate					float				NOT	NULL	DEFAULT 0.0000,		--　 TG101 稅幣匯率         
 VoidDate							datetime				NULL,						--✓　TG102 作廢日期(8字元) ,TG103 作廢時間 (8字元)        
 VoidApprovalDocNum					nvarchar(60)			NULL,						--　 TG104 專案作廢核准文號         
 VoidReason							nvarchar(200)			NULL,						--　 TG105 作廢原因         
 [Source]							nvarchar(1)				NULL,						--　 TG106 來源         
 IncomeDraftID						nvarchar(11)			NULL,						--　 TG107 收入底稿編號         
 IncomeDraftSeq						nvarchar(4)				NULL,						--　 TG108 收入底稿序號         
 IncomeVoucherType					nvarchar(4)				NULL,						--　 TG109 收入傳票單別         
 IncomeVoucherNumber				nvarchar(11)			NULL,						--　 TG110 收入傳票單號         
 [Status]							nvarchar(1)				NULL,						--　 TG111 狀態         
 ZeroTaxForBuyer					nvarchar(255)			NULL,						--　 TG112 買受人適用零稅率註         
 --ReservedField					nvarchar(255)			NULL,						--　 TG113 預留欄位         
 --ReservedField					nvarchar(255)			NULL,						--　 TG114 預留欄位         
 --ReservedField					nvarchar(255)			NULL,						--　 TG115 預留欄位         
 GenLedgerAcctType					nvarchar(6)				NULL,						--　 TG116 總帳帳別         
 InvoiceTime						nvarchar(8)				NULL,						--　 TG117 發票開立時間         
 InvCode2							nvarchar(20)			NULL,						--　 TG118 發票代號         
 InvSymbol							nvarchar(30)			NULL,						--　 TG119 發票記號         
 DeliveryCountry					nvarchar(255)			NULL,						--　 TG120 送貨國家         
 --ReservedField					nvarchar(1)				NULL,						--　 TG121 預留欄位         
 --ReservedField					nvarchar(1)				NULL,						--　 TG122 預留欄位         
 --ReservedField					nvarchar(10)			NULL,						--　 TG123 預留欄位         
 VehicleIDshow						nvarchar(80)			NULL,						--✓　TG124 載具顯碼ID         
 VehicleTypeNumber					nvarchar(20)			NULL,						--✓　TG125 載具類別號碼         
 VehicleIDhide						nvarchar(80)			NULL,						--✓　TG126 載具隱碼ID         
 InvDonationRecipient				nvarchar(30)			NULL,						--✓　TG127 發票捐贈對象         
 InvRandomCode						nvarchar(30)			NULL,						--✓　TG128 發票防偽隨機碼         
 ReservedField						nvarchar(1)				NULL,						--　 TG129 預留欄位         
 CreditCard4No						nvarchar(4)				NULL,						--✓　TG130 信用卡末四碼         
 ContactEmail						nvarchar(255)			NULL,						--✓　TG131 連絡人EMAIL         
 ExpectedReceiptDate				nvarchar(8)				NULL,						--　 TG132 預計收款日         
 OrigInvNumber						nvarchar(50)			NULL,						--　 TG133 原發票號碼   
 TransferStatusMES					nvarchar(2)			NOT	NULL,						--　MES拋轉ERP狀態 (N/Y/R/V) N:未拋轉,Y:已拋轉,R:重新編輯,V:作廢          
 TransferTime						datetime				NULL,						--　拋轉時間        
 ConfirmTime						datetime				NULL,						--　確認時間        
 SourceType							nvarchar(2)				NULL,						--　銷貨單據來源        
 PriceSourceTypeMain				nvarchar(2)				NULL,						--　銷貨單價來源   
 SourceOrderId						int						NULL,						-- 來源單據Id         
 SourceFull							nvarchar(16)			NULL,						-- 來源單據         
 CreateDate							datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate					datetime				NULL,						--修改日期
 CreateBy							int					NOT NULL,						--新增人員
 LastModifiedBy						int						NULL,						--修改人員 
 CONSTRAINT PK_ReceiveOrder PRIMARY KEY (RoId),
 CONSTRAINT AK1_ReceiveOrder UNIQUE (CompanyId, RoErpPrefix, RoErpNo),
 CONSTRAINT FK1_ReceiveOrder FOREIGN KEY (CompanyId) REFERENCES   BAS.Company(CompanyId),
 CONSTRAINT FK2_ReceiveOrder FOREIGN KEY (CustomerId) REFERENCES SCM.Customer(CustomerId)
 )

CREATE TABLE SCM.RoDetail( /*ERP銷貨單單身 COPTH*/
 RoDetailId							int					identity,
 RoId								int					NOT NULL,						-- 銷貨單單頭Id
 RoSequence                         nvarchar(4)				NULL,						--✓ TH003 序號
 MtlItemId                          int					NOT	NULL,						--✓ TH004 品號 / TH005 品名 / TH006 規格
 InventoryId						int					NOT	NULL,						--✓ TH007 庫別
 Quantity                           float				NOT	NULL　	DEFAULT 0.0000,		--✓ TH008 數量
 UomId								int						NULL,						--✓ TH009 單位
 --InventoryQty                       float					NULL,						--  TH010 庫存數量
 --SmallUomId                         int					NULL,						--  TH011 小單位
 UnitPrice                          float				NOT NULL	DEFAULT 0.0000,		--✓ TH012 單價
 Amount                             float				NOT NULL	DEFAULT 0.0000,		--✓ TH013 金額
 SoDetailId							int					NOT	NULL,						--✓ TH014 訂單單別　/ TH015 訂單單號 / TH016 訂單序號
 LotNumber							nvarchar(20)			NULL,						--✓ TH017	批號
 Remarks                            nvarchar(255)			NULL,						--✓ TH018 備註
 CustomerMtlItemNo					nvarchar(40)			NULL,						--✓ TH019	客戶品號
 ConfirmationCode                   nvarchar(1)			NOT	NULL	DEFAULT 'N',		--✓ TH020 確認碼　(Y/N)
 UpdateCode                         nvarchar(1)			NOT	NULL	DEFAULT 'N',		--  TH021 更新碼
 --ReservedField                      nvarchar(8)				NULL,					--  TH022 保留欄位
 --ReservedField                      nvarchar(10)				NULL,					--  TH023 保留欄位
 FreeSpareQty						float				NOT NULL	DEFAULT 0.0000,		--✓ TH024 贈/備品數量
 DiscountRate                       float				NOT NULL	DEFAULT 0.0000,     --✓ TH025 折扣率
 CheckOutCode                       nvarchar(1)			NOT	NULL	DEFAULT 'N',		--✓ TH026 結帳碼(Y/N)
 CheckOutPrefix						nvarchar(4)				NULL,						--✓ TH027 結帳單別
 CheckOutNo							nvarchar(11)			NULL,						--✓ TH028 結帳單號
 CheckOutSequence					nvarchar(4)				NULL,						--✓ TH029 結帳序號
 ProjectCode                        nvarchar(20)			NULL,						--✓ TH030 專案代號
 [Type]                             nvarchar(1)			NOT	NULL,						--✓ TH031 類型(1.贈品/2.備品)(預設1)
 TsnDetailId						int						NULL,						--✓ TH032 暫出單別 / TH033 暫出單號 / TH034 暫出序號
 OrigPreTaxAmt						float				NOT NULL	DEFAULT 0.0000,     --✓ TH035 原幣未稅金額
 OrigTaxAmt							float				NOT NULL	DEFAULT 0.0000,     --✓ TH036 原幣稅額
 PreTaxAmt							float				NOT NULL	DEFAULT 0.0000,     --✓ TH037 本幣未稅金額
 TaxAmt								float				NOT NULL	DEFAULT 0.0000,     --✓ TH038 本幣稅額
 PackageQty							float				NOT NULL	DEFAULT 0.0000,     --  TH039 包裝數量
 PackageFreeSpareQty				float				NOT NULL	DEFAULT 0.0000,     --✓ TH040 贈/備品包裝量
 PackageUomId                       int						NULL,						--  TH041 包裝單位
 BondedCode                         nvarchar(1)				NULL,						--✓ TH042 保稅碼
 SrQty								float				NOT NULL	DEFAULT 0.0000,     --✓ TH043 銷退數量
 SrPackageQty						float				NOT NULL	DEFAULT 0.0000,     --  TH044 銷退包裝數量
 SrOrigPreTaxAmt					float				NOT NULL	DEFAULT 0.0000,     --✓ TH045 原幣銷退未稅金額
 SrOrigTaxAmt						float				NOT NULL	DEFAULT 0.0000,     --✓ TH046 原幣銷退稅額
 SrPreTaxAmt						float				NOT NULL	DEFAULT 0.0000,     --✓ TH047 本幣銷退未稅金額
 SrTaxAmt							float				NOT NULL	DEFAULT 0.0000,     --✓ TH048 本幣銷退稅額
 CommissionRate                     float				NOT NULL	DEFAULT 0.0000,     --  TH049 佣金比率
 CommissionAmount                   float				NOT NULL	DEFAULT 0.0000,     --  TH050 佣金金額
 OriginalCustomer                   nvarchar(10)			NULL,						--  TH051 原始客戶
 SrFreeSpareQty						float				NOT NULL	DEFAULT 0.0000,     --✓ TH052 銷退贈/備品量
 SrPackageFreeSpareQty				float				NOT NULL	DEFAULT 0.0000,     --  TH053 銷退贈/備品包裝量
 NotPayTemp                         nvarchar(4)			NOT	NULL	DEFAULT 'N',		--  TH054 暫不付款(Y/N)
 --ReservedField                        nvarchar(11)			NULL,					--  TH055 預留欄位
 --ReservedField                        nvarchar(4)			NULL,						--  TH056 預留欄位
 ProductSerialNumberQty				int						NULL,						--  TH057 產品序號數量(10字元)
 ForecastCode                       nvarchar(11)			NULL,						--✓ TH058 預測代號
 ForecastSequence					nvarchar(4)				NULL,						--✓ TH059 預測序號
 [Location]							nvarchar(10)			NULL,						--✓ TH060 儲位
 AvailableQty						float				NOT NULL	DEFAULT 0.0000,     --  TH061 計價數量
 AvailableUomId                     int						NULL,						--✓ TH062 計價單位
 --ReservedField                        float					NULL,					--  TH063 預留欄位
 --ReservedField                        float					NULL,					--  TH064 預留欄位
 --ReservedField                        nvarchar(1)			NULL,						--  TH065 預留欄位
 --ReservedField                        nvarchar(30)			NULL,					--  TH066 預留欄位
 --ReservedField                        nvarchar(60)			NULL,					--  TH067 預留欄位
 MultiBatch                         nvarchar(1)			NOT	NULL	DEFAULT 'N',		--✓ TH068 多儲批
 FreeSpareRate                      float				NOT NULL	DEFAULT 0.0000,     --  TH069 贈備品率
 FinalCustomerCode                  nvarchar(10)			NULL,						--  TH070 最終客戶代號
 ReferenceQty						float				NOT NULL	DEFAULT 0.0000,     --  TH071 參考數量
 ReferencePackageQty				float				NOT NULL	DEFAULT 0.0000,     --  TH072 參考包裝數量
 TaxRate                            float				NOT NULL	DEFAULT 0.0000,     --✓ TH073 營業稅率
 CRMSource                          nvarchar(1)				NULL,						--  TH074 CRM來源
 CRMPrefix                          nvarchar(4)				NULL,						--  TH075 CRM單別
 CRMNo								nvarchar(12)			NULL,						--  TH076 CRM單號
 CRMSequence						nvarchar(5)				NULL,						--  TH077 CRM序號
 CRMContractCode                    nvarchar(20)			NULL,						--  TH078 CRM合約代號
 CRMAllowDeduction                  nvarchar(1)				NULL,						--  TH079 CRM可扣抵
 CRMDeductionQty					float				NOT NULL	DEFAULT 0.0000,     --  TH080 CRM扣抵量
 CRMDeductionUnit                   nvarchar(1)				NULL,						--  TH081 CRM扣抵單位
 DebitAccount                       nvarchar(20)			NULL,						--  TH082 借方科目
 CreditAccount                      nvarchar(20)			NULL,						--  TH083 貸方科目
 TaxAmountAccount                   nvarchar(20)			NULL,						--  TH084 稅額科目
 --ReservedField                        nvarchar(255)			NULL,					--  TH085 預留欄位
 --ReservedField                        nvarchar(255)			NULL,					--  TH086 預留欄位
 --ReservedField                        nvarchar(255)			NULL,					--  TH087 預留欄位
 --ReservedField                        nvarchar(255)			NULL,					--  TH088 預留欄位
 BusinessItemNumber                 nvarchar(40)			NULL,						--✓ TH089 業務品號
 TaxCode                            nvarchar(3)				NULL,						--  TH090 稅別碼
 DiscountAmount                     float				NOT NULL	DEFAULT 0.0000,     --  TH091 折扣金額
 --ReservedField                        nvarchar(1)			NULL,						--  TH092 預留欄位
 --ReservedField                        nvarchar(8)			NULL,						--  TH093 預留欄位
 K2NO                               nvarchar(255)			NULL,						--  TH094 K2_NO
 MarkingBINRecord                   nvarchar(1250)			NULL,						--  TH500 刻號/BIN記錄
 MarkingManagement                  nvarchar(1)				NULL,						--  TH501 刻號管理
 BillingUnitInPackage               nvarchar(1)				NULL,						--  TH502 以包裝單位計價
 DATECODE							nvarchar(20)			NULL,						--  TH503 DATECODE
 TransferStatusMES					nvarchar(1)			NOT	NULL,						--MES拋轉ERP狀態 (N/Y/R/V) N:未拋轉,Y:已拋轉,R:重新編輯,V:作廢      
 CreateDate							datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate					datetime				NULL,						--修改日期
 CreateBy							int					NOT NULL,						--新增人員
 LastModifiedBy						int						NULL,						--修改人員 
 CONSTRAINT PK_RoDetailu PRIMARY KEY (RoDetailId),
 CONSTRAINT FK1_RoDetail FOREIGN KEY (RoId) REFERENCES   SCM.ReceiveOrder(RoId)
 )

CREATE TABLE SCM.ReturnReceiveOrder( /*ERP銷貨退回單 COPTI*/
  RtId								int					identity,
  CompanyId							int					NOT NULL,						--✓ 所屬公司
  RtErpPrefix						nvarchar(4)			NOT NULL,						--TI001  單別
  RtErpNo							nvarchar(11)		NOT NULL,						--TI002  單號
  ReturnDate						datetime			NOT NULL,						--TI003  銷退日
  CustomerId						int					NOT NULL,						--TI004  客戶
  DepartmentId						int					    NULL,						--TI005  部門
  SalesmenId						int					    NULL,						--TI006  業務員
  Factory							nvarchar(6)			NOT NULL,						--TI007  廠別
  Currency							nvarchar(4) 		NOT	NULL,						--TI008  幣別
  ExchangeRate						float				NOT	NULL	DEFAULT 0.0000,		--TI009  匯率
  OriginalAmount					float				NOT	NULL	DEFAULT 0.0000,		--TI010  原幣銷退金額
  OriginalTaxAmount					float				NOT	NULL	DEFAULT 0.0000,		--TI011  原幣銷退稅額
  InvoiceType						nvarchar(1)				NULL,						--TI012  發票聯數
  TaxType							nvarchar(1)				NULL,						--TI013  課稅別
  InvNumEnd							nvarchar(50)			NULL,						--TI014  發票號碼(訖)
  UiNo								nvarchar(20)			NULL,						--TI015  統一編號
  PrintCount						int					    NULL,						--TI016  列印次數
  InvoiceDate						datetime				NULL,						--TI017  發票日期
  UpdateCode						nvarchar(1)			NOT NULL	DEFAULT 'N',		--TI018  更新碼
  ConfirmStatus						nvarchar(1)			NOT NULL	DEFAULT 'N',		--TI019  確認碼
  Remark							nvarchar(20)			NULL,					    --TI020  備註
  CustomerFullName					nvarchar(80)			NULL,						--TI021  客戶全名
  CollectionSalesmenId				int					    NULL,						--TI022  收款業務員
  Remark1							nvarchar(20)			NULL,						--TI023  備註一
  Remark2							nvarchar(20)			NULL,						--TI024  備註二
  Remark3							nvarchar(20)			NULL,						--TI025  備註三
  DeductionDistinction				nvarchar(1)				NULL,						--TI026  扣抵區分
  CustomsClearance					nvarchar(1)				NULL,						--TI027  通關方式
  RowCnt							int						NULL,						--TI028  件數
  TotalQuantity						float				NOT	NULL	DEFAULT 0.0000,		--TI029  總數量
  StaffId							int						NULL,						--TI030  員工代號
  RevenueJournalCode				nvarchar(1)				NULL,						--TI031  產生分錄碼(收入)
  CostJournalCode					nvarchar(1)				NULL,						--TI032  產生分錄碼(成本)
  ApplyYYMM							nvarchar(6)			NOT NULL,						--TI033  申報年月
  DocDate							datetime			NOT NULL,						--TI034  單據日期
  ConfirmUserId						int						NULL,						--TI035  確認者
  TaxRate							float				NOT	NULL	DEFAULT 0.0000,		--TI036  營業稅率
  PretaxAmount						float				NOT	NULL	DEFAULT 0.0000,		--TI037  本幣銷退金額
  TaxAmount							float				NOT	NULL	DEFAULT 0.0000,		--TI038  本幣銷退稅額
  PaymentTerm						nvarchar(6)				NULL,						--TI039  付款條件代號
  TotalPackages						float					NULL,						--TI040  總包裝數量
  SignatureStatus					nvarchar(1)				NULL,						--TI041  簽核狀態碼
  ProcessCode						nvarchar(2)				NULL,						--TI042  流程代號
  TransferStatusERP					nvarchar(1)				NULL,						--TI043  拋轉狀態
  BondCode							nvarchar(1)				NULL,						--TI044  保稅碼
  TransmissionCount					int						NULL,						--TI045  傳送次數
  CustomerAddressFirst				nvarchar(255)			NULL,						--TI046  取貨地址一
  CustomerAddressSecond				nvarchar(255)			NULL,						--TI047  取貨地址二
  --TI048  折讓列印次數
  ContactPerson						nvarchar(30)			NULL,						--TI049  連絡人
  Courier							nvarchar(10)			NULL,						--TI050  代送商代號
  SiteCommCalcMethod				nvarchar(1)				NULL,						--TI051  營站佣金計算
  SiteCommRate						float				NOT	NULL	DEFAULT 0.0000,		--TI052  營站佣金比率
  CommCalcInclTax					nvarchar(1)				NULL,						--TI053  佣金計算含稅
  TotalCommAmount					float				NOT	NULL	DEFAULT 0.0000,		--TI054  佣金總金額
  --TI055  EBC銷退單號
  --TI056  EBC銷退版次
  TradingTerms						nvarchar(1)				NULL,						--TI057  交易條件
  CustomerEgFullName				nvarchar(80)			NULL,						--TI058  客戶英文全名
  --TI059  預留欄位
  --TI060  預留欄位
  DebitNote							nvarchar(1)			NOT	NULL,						--TI061  折讓證明單
  --TI062  來源
  --TI063  預留欄位
  TaxCode							nvarchar(3)			NOT	NULL,						--TI064  稅別碼
  --TI065  海關手冊
  RemarkCode						nvarchar(20)			NULL,						--TI066  註記號
  MultipleInvoices					nvarchar(1)				NULL,						--TI067  多發票
  InvNumStart						nvarchar(50)			NULL,						--TI068  發票號碼(起)
  NumberOfInvoices					int						NULL,						--TI069  發票張數
  MultiTaxRate						nvarchar(1)				NULL,						--TI070  單身多稅率
  TaxCurrencyRate					float				NOT	NULL	DEFAULT 0.0000,		--TI071  稅幣匯率
  --TI072  發票註記碼長
  --TI073  發票流水碼長
  VoidDate							datetime				NULL,						--TI074  作廢日期
  VoidTime							datetime				NULL,						--TI075  作廢時間
  VoidRemark						nvarchar(20)			NULL,						--TI076  作廢原因
  IncomeDraftID						nvarchar(11)			NULL,						--TI077  收入底稿編號
  IncomeDraftSeq					nvarchar(4)				NULL,						--TI078  收入底稿序號
  IncomeVoucherType					nvarchar(4)				NULL,						--TI079  收入傳票單別
  IncomeVoucherNumber				nvarchar(11)			NULL,						--TI080  收入傳票單號
  [Status]							nvarchar(1)				NULL,						--TI081  狀態
  --TI082  預留欄位
  --TI083  預留欄位
  --TI084  預留欄位
  --TI085  預留欄位
  GenLedgerAcctType					nvarchar(6)				NULL,						--TI086  總帳帳別
  InvCode2							nvarchar(20)			NULL,						--TI087  發票代號
  InvSymbol							nvarchar(30)			NULL,						--TI088  發票記號
  --TI089  賣方開立折讓
  --TI090  折讓單簽回日
  ContactEmail						nvarchar(255)			NULL,						--TI091  連絡人EMAIL
  OrigInvNumber						nvarchar(50)			NULL,						--TI092  原發票號碼    
  TransferStatusMES					nvarchar(1)			NOT	NULL,						--MES拋轉ERP狀態 (N/Y/R/V) N:未拋轉,Y:已拋轉,R:重新編輯,V:作廢      
  CreateDate						datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
  LastModifiedDate					datetime				NULL,						--修改日期
  CreateBy							int					NOT NULL,						--新增人員
  LastModifiedBy					int						NULL,						--修改人員 
 CONSTRAINT PK_ReturnReceiveOrder PRIMARY KEY (RtId),
 CONSTRAINT AK1_ReturnReceiveOrder UNIQUE (CompanyId, RtErpPrefix, RtErpNo),
 CONSTRAINT FK1_ReturnReceiveOrder FOREIGN KEY (CompanyId) REFERENCES   BAS.Company(CompanyId),
 CONSTRAINT FK2_ReturnReceiveOrder FOREIGN KEY (CustomerId) REFERENCES SCM.Customer(CustomerId)
 )

CREATE TABLE SCM.RtDetail( /*ERP銷貨退回單單身 COPTJ*/
 RtDetailId							int					identity,
 RtId								int					NOT NULL,						--銷貨單單頭Id
 RtSequence		                    nvarchar(4)				NULL,						--TJ003 序號
 MtlItemId                          int					NOT	NULL,						--TJ004 品號 TJ005 品名 TJ006 規格    
 Quantity                           float				NOT	NULL　	DEFAULT 0.0000,		--TJ007 數量    
 UomId								int						NULL,						--TJ008 單位    
 --TJ009 庫存數量    
 --TJ010 小單位    
 UnitPrice                          float				NOT NULL	DEFAULT 0.0000,		--TJ011 單價    
 Amount                             float				NOT NULL	DEFAULT 0.0000,		--TJ012 金額    
 InventoryId						int					NOT	NULL,						--TJ013 退貨庫別    
 LotNumber							nvarchar(20)			NULL,						--TJ014 批號    
 RoDetailId							int					NOT	NULL,						--TJ015 銷貨單別 TJ016 銷貨單號 TJ017 銷貨序號 
 SoDetailId							int					NOT	NULL,						--TJ018 訂單單別 TJ019 訂單單號 TJ020 訂單序號 
 ConfirmStatus						nvarchar(1)			NOT NULL	DEFAULT 'N',		--TJ021 確認碼    
 UpdateCode							nvarchar(1)			NOT NULL	DEFAULT 'N',		--TJ022 更新碼    
 Remarks                            nvarchar(255)			NULL,						--TJ023 備註    
 CheckOutCode                       nvarchar(1)			NOT	NULL	DEFAULT 'N',		--TJ024 結帳碼    
 CheckOutPrefix						nvarchar(4)				NULL,						--TJ025 結帳單別    
 CheckOutNo							nvarchar(11)			NULL,						--TJ026 結帳單號    
 CheckOutSeq						nvarchar(4)				NULL,						--TJ027 結帳序號    
 ProjectCode                        nvarchar(20)			NULL,						--TJ028 專案代號    
 CustomerMtlItemNo					nvarchar(40)			NULL,						--TJ029 客戶品號    
 [Type]                             nvarchar(1)			NOT	NULL,						--TJ030 類型    
 OrigPreTaxAmt						float				NOT NULL	DEFAULT 0.0000,		--TJ031 原幣未稅金額    
 OrigTaxAmt							float				NOT NULL	DEFAULT 0.0000,		--TJ032 原幣稅額    
 PreTaxAmt							float				NOT NULL	DEFAULT 0.0000,		--TJ033 本幣未稅金額    
 TaxAmt								float				NOT NULL	DEFAULT 0.0000,		--TJ034 本幣稅額    
 PackageQty							float				NOT NULL	DEFAULT 0.0000,		--TJ035 包裝數量    
 PackageUomId                       int						NULL,						--TJ036 包裝單位    
 BondedCode                         nvarchar(1)				NULL,						--TJ037 保稅碼    
 CommissionRate                     float				NOT NULL	DEFAULT 0.0000,		--TJ038 佣金比率    
 CommissionAmount                   float				NOT NULL	DEFAULT 0.0000,		--TJ039 佣金金額    
 OriginalCustomer                   nvarchar(10)			NULL,						--TJ040 原始客戶    
 FreeSpareType                      nvarchar(1)			NOT	NULL,						--TJ041 數量類型    
 FreeSpareQty						float				NOT NULL	DEFAULT 0.0000,		--TJ042 贈/備品量    
 PackageFreeSpareQty                float					NULL,						--TJ043 贈/備品包裝量    
 --TJ044 EBC銷退單號    
 --TJ045 EBC銷退版次    
 --TJ046 EBC銷退序號    
 --TJ047 產品序號數量    
 NotPayTemp                         nvarchar(4)			NOT	NULL	DEFAULT 'N',		--TJ048 暫不付款    
 [Location]							nvarchar(10)			NULL,						--TJ049 儲位    
 AvailableQty						float				NOT NULL	DEFAULT 0.0000,		--TJ050 計價數量    
 AvailableUomId                     int						NULL,						--TJ051 計價單位    
 --TJ052 銷退原因代號    
 --TJ053 預留欄位    
 --TJ054 預留欄位    
 --TJ055 預留欄位    
 --TJ056 預留欄位    
 --TJ057 預留欄位    
 TaxRate                            float				NOT NULL	DEFAULT 0.0000,		--TJ058 營業稅率    
 --TJ059 CRM可扣抵    
 --TJ060 CRM扣抵量    
 --TJ061 CRM扣抵單位    
 --TJ062 借方科目    
 --TJ063 貸方科目    
 --TJ064 稅額科目    
 --TJ065 預留欄位    
 --TJ066 預留欄位    
 --TJ067 預留欄位    
 --TJ068 預留欄位    
 --TJ069 業務品號    
 TaxCode                            nvarchar(3)				NULL,						--TJ070 稅別碼    
 DiscountAmount                     float				NOT NULL	DEFAULT 0.0000,		--TJ071 折扣金額    
 --TJ500 刻號/BIN記錄    
 --TJ501 刻號管理    
 --TJ502 以包裝單位計價    
 --TJ503 DATECODE    
 TransferStatusMES					nvarchar(1)			NOT	NULL,						--MES拋轉ERP狀態 (N/Y/R/V) N:未拋轉,Y:已拋轉,R:重新編輯,V:作廢      
 CreateDate							datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate					datetime				NULL,						--修改日期
 CreateBy							int					NOT NULL,						--新增人員
 LastModifiedBy						int						NULL,						--修改人員 
 CONSTRAINT PK_RtDetail PRIMARY KEY (RtDetailId),
 CONSTRAINT FK1_RtDetail FOREIGN KEY (RtId) REFERENCES SCM.ReturnReceiveOrder(RtId),
 CONSTRAINT FK2_RtDetail FOREIGN KEY (SoDetailId) REFERENCES SCM.SoDetail(SoDetailId),
 CONSTRAINT FK3_RtDetail FOREIGN KEY (RoDetailId) REFERENCES SCM.RoDetail(RoDetailId)
 )

END/*銷貨單管理*/

BEGIN/*品號庫別檔*/
CREATE TABLE SCM.ItemInventory( /*ERP品號庫別檔 INVMC */
  ItemInventoryId				int						identity,
  CompanyId						int						NOT NULL,						--所屬公司
  MtlItemId						int						NOT NULL,						--品號
  InventoryId					int						NOT NULL,						--庫別
  InventoryQty					float					NOT NULL	DEFAULT 0,			--庫存數量
  CreateDate					datetime				NOT NULL	DEFAULT GETDATE(),	--新增日期
  LastModifiedDate				datetime				NULL,							--修改日期
  CreateBy						int						NOT NULL,						--新增人員
  LastModifiedBy				int						NULL,							--修改人員 
 CONSTRAINT PK_ItemInventory PRIMARY KEY (ItemInventoryId),
 CONSTRAINT AK1_ItemInventory UNIQUE (CompanyId, MtlItemId , InventoryId),
 CONSTRAINT FK1_ItemInventory FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId),
 CONSTRAINT FK2_ItemInventory FOREIGN KEY (MtlItemId) REFERENCES PDM.MtlItem(MtlItemId),
 CONSTRAINT FK3_ItemInventory FOREIGN KEY (InventoryId) REFERENCES SCM.Inventory(InventoryId)
 )
END/*品號庫別檔*/

