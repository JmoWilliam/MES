BEGIN /*基本資料*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'UserSetting') > 0 DROP TABLE IM.UserSetting
END /*基本資料*/

BEGIN /*訊息管理*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MessageLog') > 0 DROP TABLE IM.MessageLog
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Usually') > 0 DROP TABLE IM.Usually
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'GroupUsers') > 0 DROP TABLE IM.GroupUsers
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Groups') > 0 DROP TABLE IM.Groups
END /*訊息管理*/
--******************************************************************************************************************
BEGIN /*基本資料*/
CREATE TABLE IM.UserSetting( /*使用者設定資料*/
 UserId						int					NOT NULL,						--所屬使用者
 Avatar						int						NULL,						--頭像
 Nickname					nvarchar(20)		NOT NULL,						--暱稱
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT AK1_UserSetting UNIQUE (Nickname),
 CONSTRAINT FK1_UserSetting FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK2_UserSetting FOREIGN KEY (Avatar) REFERENCES BAS.[File](FileId)
)/*使用者設定資料*/
END /*基本資料*/

BEGIN /*訊息管理*/
CREATE TABLE IM.Groups( /*群組資料*/
 GroupsId					int					identity,
 GroupsNo					nvarchar(MAX)		NOT NULL,						--群組編號
 GroupsName					nvarchar(10)		NOT NULL,						--群組名稱
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Groups PRIMARY KEY (GroupsId),
 CONSTRAINT AK1_Groups UNIQUE (GroupsNo),
 CONSTRAINT AK2_Groups UNIQUE (GroupsName)
)/*群組資料*/

CREATE TABLE IM.GroupUsers( /*群組使用者資料*/
 GroupsId					int					NOT NULL,						--所屬群組
 UserId						int					NOT NULL,						--所屬使用者
 AdminStatus				nvarchar(2)			NOT NULL,						--管理員狀態 >> BAS.[Status]  Y：是 N：否
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT AK1_GroupUsers UNIQUE (GroupsId, UserId),
 CONSTRAINT FK1_GroupUsers FOREIGN KEY (GroupsId) REFERENCES IM.Groups(GroupsId),
 CONSTRAINT FK2_GroupUsers FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId)
)/*群組使用者資料*/

CREATE TABLE IM.Usually( /*常用資料*/
 UsuallyId					int					identity,
 GroupsId					int						NULL,						--所屬群組
 UserId						int						NULL,						--所屬使用者
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Usually PRIMARY KEY (UsuallyId),
 CONSTRAINT FK1_Usually FOREIGN KEY (GroupsId) REFERENCES IM.Groups(GroupsId),
 CONSTRAINT FK2_Usually FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId)
) /*常用資料*/

CREATE TABLE IM.MessageLog( /*訊息紀錄資料*/
 LogId						int					identity,
 CurrentUser				int					NOT NULL,						--目前使用者
 MessageType				nvarchar(2)			NOT NULL,						--訊息類別 >> BAS.[Type] S：Send；R：Receive
 TargetUser					int						NULL,						--目標使用者
 TargetGroup				int						NULL,						--目標群組
 SendMessage				nvarchar(MAX)		NOT NULL,						--發送訊息
 SendFile					int						NULL,						--發送檔案
 ReplyLogId					int						NULL,						--回覆之訊息
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 CreateBy					int					NOT NULL,						--新增人員
 CONSTRAINT PK_MessageLog PRIMARY KEY (LogId),
 CONSTRAINT FK1_MessageLog FOREIGN KEY (CurrentUser) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK2_MessageLog FOREIGN KEY (TargetUser) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK3_MessageLog FOREIGN KEY (TargetGroup) REFERENCES IM.Groups(GroupsId),
 CONSTRAINT FK4_MessageLog FOREIGN KEY (MessageFile) REFERENCES BAS.[File](FileId),
 CONSTRAINT FK5_MessageLog FOREIGN KEY (ReplyLogId) REFERENCES IM.MessageLog(LogId)
) /*訊息紀錄資料*/

END /*訊息管理*/