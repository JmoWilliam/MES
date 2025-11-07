BEGIN /*LOG相關資料*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MamoLog') > 0 DROP TABLE MAMO.MamoLog
END /*LOG相關資料*/

BEGIN /*訊息相關資料*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'PnFile') > 0 DROP TABLE MAMO.PnFile
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'PushNotification') > 0 DROP TABLE MAMO.PushNotification
END /*訊息相關資料*/

BEGIN /*頻道資料*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ChannelMembers') > 0 DROP TABLE MAMO.ChannelMembers
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Channels') > 0 DROP TABLE MAMO.Channels
END /*頻道資料*/

BEGIN /*團隊資料*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'TeamMembers') > 0 DROP TABLE MAMO.TeamMembers
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Team') > 0 DROP TABLE MAMO.Team
END /*團隊資料*/
--******************************************************************************************************************
BEGIN /*團隊資料*/
CREATE TABLE MAMO.Teams( /*團隊基本資料*/
 TeamId						int					identity,
 CompanyId					int					NOT NULL,						--公司ID
 MamoTeamId					nvarchar(200)		NOT NULL,						--Mamo團隊ID碼
 TeamName					nvarchar(100)		NOT NULL,						--團隊名稱
 TeamNo						nvarchar(100)			NULL,						--Mamo Project No
 [Status]					nvarchar(2)			NOT NULL	DEFAULT 'A',		--狀態碼	A:啟用中 S:停用
 Remark						nvarchar(258)			NULL,						--備註
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Teams PRIMARY KEY (TeamId),
 CONSTRAINT AK1_Teams UNIQUE (MamoTeamId),
 CONSTRAINT FK1_Teams FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)/*團隊基本資料*/

CREATE TABLE MAMO.TeamMembers( /*團隊成員*/
 MemberId					int					identity,
 TeamId						int					NOT NULL,						--團隊ID
 UserId						int					NOT NULL,						--USER ID
 [Status]					nvarchar(2)			NOT NULL	DEFAULT 'A',		--狀態碼	A:啟用中 S:停用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_TeamMembers PRIMARY KEY (MemberId),
 CONSTRAINT AK1_TeamMembers UNIQUE (TeamId, UserId),
 CONSTRAINT FK1_TeamMembers FOREIGN KEY (TeamId) REFERENCES MAMO.Teams(TeamId),
 CONSTRAINT FK2_TeamMembers FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId)
)/*團隊成員*/
END /*團隊資料*/

BEGIN /*頻道資料*/
CREATE TABLE MAMO.Channels( /*頻道基本資料*/
 ChannelId					int					identity,
 TeamId						int					NOT NULL,						--團隊ID
 MamoChannelId				nvarchar(200)		NOT NULL,						--Mamo頻道ID碼
 ChannelName				nvarchar(100)		NOT NULL,						--團隊名稱
 ChannelNo					nvarchar(100)			NULL,						--Mamo Channel No
 [Status]					nvarchar(2)			NOT NULL	DEFAULT 'A',		--狀態碼	A:啟用中 S:停用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Channels PRIMARY KEY (ChannelId),
 CONSTRAINT AK1_Channels UNIQUE (ChannelId, MamoChannelId),
 CONSTRAINT FK1_Channels FOREIGN KEY (TeamId) REFERENCES MAMO.Teams(TeamId)
)/*頻道基本資料*/

CREATE TABLE MAMO.ChannelMembers( /*頻道成員*/
 MemberId					int					identity,
 ChannelId					int					NOT NULL,						--頻道ID
 UserId						int					NOT NULL,						--USER ID
 [Status]					nvarchar(2)			NOT NULL	DEFAULT 'A',		--狀態碼	A:啟用中 S:停用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ChannelMembers PRIMARY KEY (MemberId),
 CONSTRAINT AK1_ChannelMembers UNIQUE (ChannelId, UserId),
 CONSTRAINT FK1_ChannelMembers FOREIGN KEY (ChannelId) REFERENCES MAMO.Channels(ChannelId),
 CONSTRAINT FK2_ChannelMembers FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId)
)/*團隊成員*/
END /*團隊資料*/

BEGIN /*訊息相關資料*/
CREATE TABLE MAMO.PushNotification( /*推播紀錄*/
 PnId						int					identity,
 PushType					nvarchar(100)		NOT NULL,						--推播類型 Personal:個人, Channel:頻道
 SendId						nvarchar(100)		NOT NULL,						--寄送ID(MAMO)
 Content					nvarchar(MAX)		NOT NULL,						--訊息內容
 Tags						nvarchar(MAX)			NULL,						--標記USER ID 格式: ZY01001,ZY01002
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_PushNotification PRIMARY KEY (PnId)
)/*推播紀錄*/

CREATE TABLE MAMO.PnFile( /*推播檔案*/
 PnFileId					int					identity,
 PnId						int					NOT NULL,						--推播ID
 FileId						int					NOT NULL,						--檔案ID
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_PnFile PRIMARY KEY (PnFileId),
 CONSTRAINT FK1_PnFile FOREIGN KEY (PnId) REFERENCES MAMO.PushNotification(PnId),
 CONSTRAINT FK2_PnFile FOREIGN KEY (FileId) REFERENCES BAS.[File](FileId),
)/*推播檔案*/
END /*訊息相關資料*/

BEGIN /*LOG*/
CREATE TABLE MAMO.MamoLog( /*異動LOG*/
 LogId						int					identity,
 MamoId						nvarchar(200)		NOT NULL,						--MamoID碼
 MamoTeamId					nvarchar(200)			NULL,						--Mamo TeamID碼
 ControlType				nvarchar(200)		NOT NULL,						--異動紀錄類型(EX:CreateTeams、AddTeamMembers)
 UserList					nvarchar(MAX)			NULL,						--USER列表
 Remark						nvarchar(256)			NULL,						--備註
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_TeamsLog PRIMARY KEY (LogId)
)/*異動LOG*/
END