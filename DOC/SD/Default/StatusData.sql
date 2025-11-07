BEGIN /*中揚*/
--(StatusSchema, CompanyId, StatusNo, StatusName, StatusDescription, CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)

BEGIN /*通用狀態*/
DELETE FROM BAS.[Status] WHERE StatusSchema = 'Status';
INSERT INTO BAS.[Status] VALUES('Status', 2, 'A', '啟用', '啟用', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Status] VALUES('Status', 2, 'S', '停用', '停用', GETDATE(), GETDATE(), 0, 0);
END

BEGIN /*系統狀態*/
DELETE FROM BAS.[Status] WHERE StatusSchema = 'SystemStatus';
INSERT INTO BAS.[Status] VALUES('SystemStatus', 2, 'M', '手動', '手動', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Status] VALUES('SystemStatus', 2, 'A', '自動', '自動', GETDATE(), GETDATE(), 0, 0);
END

BEGIN /*使用者狀態*/
DELETE FROM BAS.[Status] WHERE StatusSchema = 'User.UserStatus';
INSERT INTO BAS.[Status] VALUES('User.UserStatus', 2, 'T', '試用員工', '試用員工', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Status] VALUES('User.UserStatus', 2, 'F', '正式員工', '正式員工', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Status] VALUES('User.UserStatus', 2, 'S', '離職員工', '離職員工', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Status] VALUES('User.UserStatus', 2, 'R', '退休員工', '退休員工', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Status] VALUES('User.UserStatus', 2, 'L', '留職停薪員工', '留職停薪員工', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Status] VALUES('User.UserStatus', 2, 'P', '約聘或其他', '約聘或其他', GETDATE(), GETDATE(), 0, 0);
END

BEGIN /*需求單狀態*/
DELETE FROM BAS.[Status] WHERE StatusSchema = 'DemandStatus';
INSERT INTO BAS.[Status] VALUES('DemandStatus','0', '待處理', '待處理', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Status] VALUES('DemandStatus','1', '處理中', '處理中', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Status] VALUES('DemandStatus','2', '已處理', '已處理', GETDATE(), GETDATE(), 0, 0);
END

BEGIN /*需求單 簽核歷程狀態*/
DELETE FROM BAS.[Status] WHERE StatusSchema = 'DemandProcessStatus';
INSERT INTO BAS.[Status] VALUES('DemandProcessStatus','0', '新單據', '新單據', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Status] VALUES('DemandProcessStatus','1', '訂單建立完成', '訂單建立完成', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Status] VALUES('DemandProcessStatus','2', '訂單核准完成', '訂單核准完成', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Status] VALUES('DemandProcessStatus','3', '品號建立完成', '品號建立完成', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Status] VALUES('DemandProcessStatus','4', '品號核准完成', '品號核准完成', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Status] VALUES('DemandProcessStatus','5', '主件建立完成', '主件建立完成', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Status] VALUES('DemandProcessStatus','6', '主件核准完成', '主件核准完成', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Status] VALUES('DemandProcessStatus','7', '元件建立完成', '元件建立完成', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Status] VALUES('DemandProcessStatus','8', '元件核准完成', '元件核准完成', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Status] VALUES('DemandProcessStatus','9', '客戶圖面建立完成', '客戶圖面建立完成', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Status] VALUES('DemandProcessStatus','10', '研發圖面建立完成', '研發圖面建立完成', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Status] VALUES('DemandProcessStatus','11', '途程品號建立完成', '途程品號建立完成', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Status] VALUES('DemandProcessStatus','12', 'ERP工單建立完成', 'ERP工單建立完成', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Status] VALUES('DemandProcessStatus','13', 'ERP工單核准完成', 'ERP工單核准完成', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Status] VALUES('DemandProcessStatus','14', 'MES製令建立完成', 'MES製令建立完成', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Status] VALUES('DemandProcessStatus','15', '領料單建立完成', '領料單建立完成', GETDATE(), GETDATE(), 0, 0);
END

BEGIN /*品號變更單 - ERP拋單狀態*/
DELETE FROM BAS.[Status] WHERE StatusSchema = 'MtlItemChange.TransferStatus';
INSERT INTO BAS.[Status] VALUES('MtlItemChange.TransferStatus','0', '未拋轉', '未拋轉', GETDATE(), GETDATE(), 0, 0);
INSERT INTO BAS.[Status] VALUES('MtlItemChange.TransferStatus','1', '已拋轉', '已拋轉', GETDATE(), GETDATE(), 0, 0);
END


END

/**/