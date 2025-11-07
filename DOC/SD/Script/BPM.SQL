BEGIN /*文件簽核*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'SystemToken') > 0 DROP TABLE BPM.SystemToken
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'BpmUser') > 0 DROP TABLE BPM.BpmUser
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DeptDocument') > 0 DROP TABLE BPM.DeptDocument
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Folder') > 0 DROP TABLE BPM.Folder
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DemandMessage') > 0 DROP TABLE BPM.DemandMessage
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DMS') > 0 DROP TABLE BPM.DMS
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ATTACH') > 0 DROP TABLE BPM.[ATTACH]
END /*文件簽核*/

--******************************************************************************************************************
BEGIN /*基本資料*/
CREATE TABLE BPM.SystemToken( /*Bpm Token*/
 IpAddress					nvarchar(100)		NOT NULL,						--IP位置		
 CompanyId					int					NOT NULL,						--所屬公司
 Token						nvarchar(MAX)		NOT NULL,						--TOKEN
 VerifyDate					datetime			NOT NULL,						--驗證時間
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_SystemToken PRIMARY KEY (IpAddress, CompanyId),
 CONSTRAINT FK1_SystemToken FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)

CREATE TABLE BPM.BpmUser( /*Bpm使用者資訊*/
 UserId						int					NOT NULL,						--使用者ID	
 BpmUserId					nvarchar(500)		NOT NULL,						--Bpm使用者ID
 BpmUserNo					nvarchar(500)		NOT NULL,						--Bpm使用者NO
 BpmUserName				nvarchar(100)		NOT NULL,						--Bpm使用者Name
 BpmRoleId					nvarchar(500)		NOT NULL,						--Bpm使用者Role
 BpmRoleName				nvarchar(100)		NOT NULL,						--Bpm使用者Role Name
 BpmDepId					nvarchar(100)		NOT NULL,						--Bpm使用者部門ID
 BpmDepNo					nvarchar(100)		NOT NULL,						--Bpm使用者部門NO
 BpmDepName					nvarchar(100)		NOT NULL,						--Bpm使用者部門NO
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_BpmUser PRIMARY KEY (UserId, BpmUserId),
 CONSTRAINT FK1_BpmUser FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId)
)


END /*基本資料*/

BEGIN /*文件簽核*/

CREATE TABLE BPM.Folder( /*卷宗資料*/
 FolderId					int					identity,						
 CompanyId					int					NOT NULL,						--所屬公司
 FolderNo					nvarchar(100)		NOT NULL,						--文件夾代碼
 FolderType					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Folder PRIMARY KEY (FolderId),
 CONSTRAINT FK1_Folder FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)


CREATE TABLE BPM.DeptDocument( /*簽核文件組成要件*/
 DocId						int					identity,						
 FolderId					int					NOT NULL,						--卷宗夾ID
 DepartmentId				int					NOT NULL,						--所屬部門
 DocName					nvarchar(200)			NULL,						--文件名稱
 SenderId					int					NOT NULL,						--寄送者ID
 SendTime					datetime			NOT NULL,						--寄送時間
 ApproverId					int					NOT NULL,						--批准者ID
 ApproverTime				datetime			NOT NULL,						--批准時間
 Remark						nvarchar(200)		NULL,							--備註
 [Status]					nvarchar(1)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DeptDocument PRIMARY KEY (DocId),
 CONSTRAINT FK1_DeptDocument FOREIGN KEY (DepartmentId) REFERENCES BAS.Department(DepartmentId),
 CONSTRAINT FK2_DeptDocument FOREIGN KEY (FolderId) REFERENCES BPM.Folder(FolderId)
)

CREATE TABLE BPM.DemandMessage( /*推播訊息資訊*/
 DemandMessageId			int					identity,						
 MessageType				int					NOT NULL,						--訊息類型
 DemandMessageContent		nvarchar(MAX)		NOT NULL,						--訊息內容
 DocId						int					NOT NULL,						--DocID
 RoomId						nvarchar(200)		NOT NULL,						--RoomID
 SendUserId					int					NOT NULL,						--寄送者ID
 SendStatus					nvarchar(2)			NOT NULL,						--寄送狀態
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DemandMessage PRIMARY KEY (DemandMessageId),
 CONSTRAINT FK1_DemandMessage FOREIGN KEY (DocId) REFERENCES BPM.DeptDocument(DocId)
)
END /*文件簽核*/

BEGIN /*模擬ERP設定資料表*/
CREATE TABLE BPM.DMS( /**/
 DocID						nvarchar(10)		NOT NULL,						--DOC ID						
 Revision					nvarchar(10)		NOT NULL,						--
 [Owner]					nvarchar(10)			NULL,						--
 [Status]					nvarchar(10)			NULL,						--狀態碼
 AddDate					nvarchar(10)			NULL,						--新增日期
 AddTime					nvarchar(10)			NULL,						--修改時間
 DocName					nvarchar(100)			NULL,						--檔案名稱
 DocType					nvarchar(100)			NULL,						--檔案副檔名
 [Description]				nvarchar(255)			NULL,						--描述
 CONSTRAINT PK_DMS PRIMARY KEY (DocID)
)

CREATE TABLE BPM.[ATTACH]( /**/
 Parent						nvarchar(10)		NOT NULL,						--						
 KeyValues					nvarchar(100)		NOT NULL,						--
 CompanyID					nvarchar(20)		NOT NULL,						--公司別
 UserID						nvarchar(10)		NOT NULL,						--建立者NO
 [Type]						nvarchar(10)			NULL,						--
 SeqNo						nvarchar(10)		NOT	NULL,						--序號
 [FileName]					nvarchar(100)			NULL,						--檔案名稱
 DocID						nvarchar(10)			NULL,						--檔案副檔名
 Revision					nvarchar(10)		NOT NULL,						--
 AddDate					nvarchar(10)			NULL,						--新增日期
 AddTime					nvarchar(10)			NULL,						--修改時間
 KeyFields					nvarchar(100)			NULL,						--修改時間
)
END