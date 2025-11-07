BEGIN /*中揚*/
--(TypeSchema, CompanyId, TypeNo, TypeName, TypeDescription, CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)

BEGIN /*性別*/
DELETE FROM BAS.[Type] WHERE TypeSchema = 'UserGender';
INSERT INTO BAS.[Type] VALUES('UserGender', 2, 'F', '女', '女', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('UserGender', 2, 'M', '男', '男', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('UserGender', 2, 'N', '不顯示', '不顯示', GETDATE(), GETDATE(), 0, 0);
END

BEGIN /*編程管理*/
DELETE FROM BAS.[Type] WHERE TypeSchema = 'ViewCncParameterSetting';
INSERT INTO BAS.[Type] VALUES('ViewCncParameterSetting', 'File', '上傳檔案', '上傳檔案', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('ViewCncParameterSetting', 'Combobox', '下拉', '下拉', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('ViewCncParameterSetting', 'Text', '輸入文字', '輸入文字', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('ViewCncParameterSetting', 'Radio', '單選按鈕', '單選按鈕', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('ViewCncParameterSetting', 'Checkbox', '多選按鈕', '多選按鈕', GETDATE(), GETDATE(), 0, 0);

DELETE FROM BAS.[Type] WHERE TypeSchema = 'ViewCncProgramQcItemSpec';
INSERT INTO BAS.[Type] VALUES('ViewCncProgramQcItemSpec', '1', '全檢', '全檢', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('ViewCncProgramQcItemSpec', '2', '抽檢', '抽檢', GETDATE(), GETDATE(), 0, 0);

DELETE FROM BAS.[Type] WHERE TypeSchema = 'CncWorkType';
INSERT INTO BAS.[Type] VALUES('CncWorkType', '1', '編程報工', '編程報工', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('CncWorkType', '2', '直接報工', '直接報工', GETDATE(), GETDATE(), 0, 0);

DELETE FROM BAS.[Type] WHERE TypeSchema = 'CncWorkFileType';
INSERT INTO BAS.[Type] VALUES('CncWorkFileType', '1', '編程檔案', '編程檔案', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('CncWorkFileType', '2', '其它程式', '其它程式', GETDATE(), GETDATE(), 0, 0);

END

BEGIN /*量測資料種類*/
DELETE FROM BAS.[Type] WHERE TypeSchema = 'QMS';
INSERT INTO BAS.[Type] VALUES('QMS','1','OQC','出貨檢', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('QMS','2','IQC','進貨檢', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('QMS','3','IPQC','工程檢',GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('QMS','4','PVTQC','試樣檢',GETDATE(), GETDATE(), 0, 0);
END

BEGIN /*需求單來源*/
DELETE FROM BAS.[Type] WHERE TypeSchema = 'DemandSource';
INSERT INTO BAS.[Type] VALUES('DemandSource','1','ECS','電子商務', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('DemandSource','2','MAIL','信件爬蟲', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('DemandSource','3','MANUAL','手動建單',GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('DemandSource','4','PMD','PMD',GETDATE(), GETDATE(), 0, 0);
END

BEGIN /*需求單來源*/
DELETE FROM BAS.[Type] WHERE TypeSchema = 'DemandTypeNo';
INSERT INTO BAS.[Type] VALUES('DemandTypeNo','G','成品','成品', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('DemandTypeNo','M','材料','材料', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('DemandTypeNo','S','半成品','半成品', GETDATE(), GETDATE(), 0, 0);
END

BEGIN /*需求單來源*/
DELETE FROM BAS.[Type] WHERE TypeSchema = 'DemandOrderType';
INSERT INTO BAS.[Type] VALUES('DemandOrderType','COPY','複製','複製', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('DemandOrderType','NEW','首番','首番', GETDATE(), GETDATE(), 0, 0);
END

BEGIN /*外膳廠商發票聯數來源*/
DELETE FROM BAS.[Type] WHERE TypeSchema = 'FoodVendor.InvoiceCount';
INSERT INTO BAS.[Type] VALUES('FoodVendor.InvoiceCount','1','二聯式','二聯式', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('FoodVendor.InvoiceCount','2','三聯式','三聯式', GETDATE(), GETDATE(), 0, 0);
END

BEGIN /*外膳餐廳開放時段來源*/
DELETE FROM BAS.[Type] WHERE TypeSchema = 'FoodVendor.OpeningHours';
INSERT INTO BAS.[Type] VALUES('FoodVendor.OpeningHours','1','早餐','早餐', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('FoodVendor.OpeningHours','2','中餐','中餐', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('FoodVendor.OpeningHours','3','下午茶','下午茶', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('FoodVendor.OpeningHours','4','晚餐','晚餐', GETDATE(), GETDATE(), 0, 0);
END

BEGIN /*編程上拋機台*/
DELETE FROM BAS.[Type] WHERE TypeSchema = 'SkyMarsMachine';
INSERT INTO BAS.[Type] VALUES('SkyMarsMachine','1','記憶體','記憶體', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('SkyMarsMachine','2','DataServer','DataServer', GETDATE(), GETDATE(), 0, 0);
END

BEGIN /*組立標籤樣式*/
DELETE FROM BAS.[Type] WHERE TypeSchema = 'Label.LensLabelType';
INSERT INTO BAS.[Type] VALUES('Label.LensLabelType','1','大張標籤','大張標籤', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('Label.LensLabelType','2','小張標籤','小張標籤', GETDATE(), GETDATE(), 0, 0);
END

BEGIN /*組立標籤樣式*/
DELETE FROM BAS.[Type] WHERE TypeSchema = 'Label.PrintUser';
INSERT INTO BAS.[Type] VALUES('Label.PrintUser','1','倉庫','倉庫', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('Label.PrintUser','2','量測室','量測室', GETDATE(), GETDATE(), 0, 0);
END

BEGIN /*編程*/
DELETE FROM BAS.[Type] WHERE TypeSchema = 'MES.SkyMarsMachine';
INSERT INTO BAS.[Type] VALUES('MES.SkyMarsMachine','1','記憶體','記憶體', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('MES.SkyMarsMachine','2','DataServer','DataServer', GETDATE(), GETDATE(), 0, 0);
END

BEGIN /*品號變更單 - 修改欄位類別*/
DELETE FROM BAS.[Type] WHERE TypeSchema = 'MtlItemChange.FieldType';
INSERT INTO BAS.[Type] VALUES('MtlItemChange.FieldType','1','全部','全部', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('MtlItemChange.FieldType','0','通用','通用', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('MtlItemChange.FieldType','2','基本資料','基本資料', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('MtlItemChange.FieldType','7','倉管','倉管', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('MtlItemChange.FieldType','8','採購','採購', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('MtlItemChange.FieldType','9','業務','業務', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('MtlItemChange.FieldType','A','會計','會計', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('MtlItemChange.FieldType','B','SHIPPING','SHIPPING', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('MtlItemChange.FieldType','C','生管','生管', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('MtlItemChange.FieldType','D','品管','品管', GETDATE(), GETDATE(), 0, 0);
END

BEGIN /*加工事件類型*/
DELETE FROM BAS.[Type] WHERE TypeSchema = 'ProcessEventItem';
INSERT INTO BAS.[Type] VALUES('ProcessEventItem',0, '加工前', '加工前', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('ProcessEventItem',1, '加工後', '加工後', GETDATE(), GETDATE(), 0, 0);
END

BEGIN /*加工事件類型*/
DELETE FROM BAS.[Type] WHERE TypeSchema = 'EventType';
INSERT INTO BAS.[Type] VALUES('EventType',0, '加工事件', '加工事件', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('EventType',1, '機台事件', '機台事件', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('EventType',2, '人員事件', '人員事件', GETDATE(), GETDATE(), 0, 0);
END

BEGIN /*材料型態*/
DELETE FROM BAS.[Type] WHERE TypeSchema = 'MaterialProperties';
INSERT INTO BAS.[Type] VALUES('MaterialProperties',1, '直接材料', '直接材料', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('MaterialProperties',2, '間接材料', '間接材料', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('MaterialProperties',3, '廠商供料', '廠商供料', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('MaterialProperties',4, '不發料', '不發料', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('MaterialProperties',5, '客戶供料', '客戶供料', GETDATE(), GETDATE(), 0, 0);
END

BEGIN /*補貨政策*/
DELETE FROM BAS.[Type] WHERE TypeSchema = 'ReplenishmentPolicy';
INSERT INTO BAS.[Type] VALUES('ReplenishmentPolicy','R', '依補貨點', '依補貨點', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('ReplenishmentPolicy','M', '依MRP需求', '依MRP需求', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('ReplenishmentPolicy','L', '依LRP需求', '依LRP需求', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('ReplenishmentPolicy','N', '不需', '不需', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Type] VALUES('ReplenishmentPolicy','H', '依歷史銷售', '依歷史銷售', GETDATE(), GETDATE(), 0, 0);
END

END