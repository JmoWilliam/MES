BEGIN /*基本資料*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'TemplateTaskUserAuthority') > 0 DROP TABLE PMIS.TemplateTaskUserAuthority
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'TemplateTaskUser') > 0 DROP TABLE PMIS.TemplateTaskUser
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'TemplateTaskLink') > 0 DROP TABLE PMIS.TemplateTaskLink
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'TemplateTask') > 0 DROP TABLE PMIS.TemplateTask
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'TemplateCooperation') > 0 DROP TABLE PMIS.TemplateCooperation
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Template') > 0 DROP TABLE PMIS.Template
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Authority') > 0 DROP TABLE PMIS.Authority
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'FileLevel') > 0 DROP TABLE PMIS.FileLevel
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'TrackFilterDefault') > 0 DROP TABLE PMIS.TrackFilterDefault
END /*基本資料*/

BEGIN /*專案管理*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ProjectFile') > 0 DROP TABLE PMIS.ProjectFile
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ProjectUserAuthority') > 0 DROP TABLE PMIS.ProjectUserAuthority
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ProjectUser') > 0 DROP TABLE PMIS.ProjectUser
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Project') > 0 DROP TABLE PMIS.Project
END /*專案管理*/

BEGIN /*任務管理*/
/*專案任務*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ProjectTask') > 0 DROP TABLE PMIS.ProjectTask
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ProjectTaskLink') > 0 DROP TABLE PMIS.ProjectTaskLink
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ProjectTaskFile') > 0 DROP TABLE PMIS.ProjectTaskFile
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ProjectTaskUser') > 0 DROP TABLE PMIS.ProjectTaskUser
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ProjectTaskUserAuthority') > 0 DROP TABLE PMIS.ProjectTaskUserAuthority

/*個人任務*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'PersonalTask') > 0 DROP TABLE PMIS.PersonalTask
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'PersonalTaskFile') > 0 DROP TABLE PMIS.PersonalTaskFile

/*共用資料表*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'TaskReply') > 0 DROP TABLE PMIS.TaskReply
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'TaskUser') > 0 DROP TABLE PMIS.TaskUser
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'TaskUserAuthority') > 0 DROP TABLE PMIS.TaskUserAuthority
END /*任務管理*/

--******************************************************************************************************************

BEGIN /*基本資料*/
CREATE TABLE PMIS.Authority( /*權限項目*/ 
 AuthorityId				int					identity,
 AuthorityType				nvarchar(2)			NOT NULL,						--權限分類 >> BAS.[Type] P：專案；T：任務
 AuthorityCode				nvarchar(50)		NOT NULL,						--權限代碼
 AuthorityName				nvarchar(100)		NOT NULL,						--權限名稱
 SortNumber					int					NOT NULL,						--排序號碼
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Authority PRIMARY KEY (AuthorityId),
 CONSTRAINT AK1_Authority UNIQUE (AuthorityType, AuthorityCode),
 CONSTRAINT AK2_Authority UNIQUE (AuthorityType, SortNumber)
)/*專案權限*/

CREATE TABLE PMIS.FileLevel( /*檔案等級*/
 LevelId					int					identity,
 LevelCode					nvarchar(50)		NOT NULL,						--檔案等級代碼
 LevelName					nvarchar(100)		NOT NULL,						--檔案等級名稱
 SortNumber					int					NOT NULL,						--排序號碼
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_FileLevel PRIMARY KEY (LevelId),
 CONSTRAINT AK1_FileLevel UNIQUE (LevelCode),
 CONSTRAINT AK2_FileLevel UNIQUE (LevelName)
)/*檔案等級*/

CREATE TABLE PMIS.Template( /*樣板資料*/
 TemplateId					int					identity,
 TemplateName				nvarchar(100)		NOT NULL,						--樣板名稱
 TemplateDesc				nvarchar(300)		NOT NULL,						--樣板描述
 WorkTimeStatus				nvarchar(2)			NOT NULL,						--工作天模式 >> BAS.[Status]  Y：是 N：否
 WorkTimeInterval			nvarchar(100)			NULL,						--工作時段
 TemplateOwner				int					NOT NULL,						--樣板擁有人員
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Template PRIMARY KEY (TemplateId),
 CONSTRAINT FK1_Template FOREIGN KEY (TemplateOwner) REFERENCES BAS.[User](UserId)
)/*樣板資料*/

CREATE TABLE PMIS.TemplateCooperation( /*樣板協作人員*/
 TemplateCooperationId		int					identity,
 TemplateId					int					NOT NULL,						--所屬樣板
 UserId						int					NOT NULL,						--人員
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_TemplateCooperation PRIMARY KEY (TemplateCooperationId),
 CONSTRAINT FK1_TemplateCooperation FOREIGN KEY (TemplateId) REFERENCES PMIS.Template(TemplateId),
 CONSTRAINT FK2_TemplateCooperation FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId)
)/*樣板協作人員*/

CREATE TABLE PMIS.TemplateTask( /*樣板任務資料*/
 TemplateTaskId				int					identity,
 TemplateId					int					NOT NULL,						--所屬樣板
 ParentTaskId				int						NULL,						--父層任務
 TaskSort					int					NOT NULL,						--任務排序
 TaskName					nvarchar(50)		NOT NULL,						--任務名稱
 StartPoint					int					NOT NULL,						--起始點(分鐘)
 Duration					int						NULL,						--持續時間(分鐘)
 ReplyFrequency				int						NULL,						--回報頻率(分鐘)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_TemplateTask PRIMARY KEY (TemplateTaskId),
 CONSTRAINT FK1_TemplateTask FOREIGN KEY (TemplateId) REFERENCES PMIS.Template(TemplateId),
 CONSTRAINT FK2_TemplateTask FOREIGN KEY (ParentTaskId) REFERENCES PMIS.TemplateTask(TemplateTaskId)
)/*樣板任務資料*/

CREATE TABLE PMIS.TemplateTaskLink( /*樣板任務連結資料*/
 TemplateTaskLinkId			int					identity,
 SourceTaskId				int					NOT NULL,						--來源任務
 TargetTaskId				int					NOT NULL,						--目標任務
 LinkType					nvarchar(2)			NOT NULL,						--連結類型 >> BAS.[Type]  0：結束到開始；1：開始到開始；2：結束到結束；3：開始到結束
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_TemplateTaskLink PRIMARY KEY (TemplateTaskLinkId),
 CONSTRAINT AK1_TemplateTaskLink UNIQUE (SourceTaskId, TargetTaskId, LinkType),
 CONSTRAINT FK1_TemplateTaskLink FOREIGN KEY (SourceTaskId) REFERENCES PMIS.TemplateTask(TemplateTaskId),
 CONSTRAINT FK2_TemplateTaskLink FOREIGN KEY (TargetTaskId) REFERENCES PMIS.TemplateTask(TemplateTaskId)
)/*樣板任務連結資料*/

CREATE TABLE PMIS.TemplateTaskUser( /*樣板任務成員*/
 TemplateTaskUserId			int					identity,
 TemplateTaskId				int					NOT NULL,						--所屬樣板任務
 UserId						int					NOT NULL,						--所屬使用者
 LevelId					int					NOT NULL,						--所屬檔案等級
 AgentSort					int					NOT NULL,						--代理順序
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_TemplateTaskUser PRIMARY KEY (TemplateTaskUserId),
 CONSTRAINT AK1_TemplateTaskUser UNIQUE (TemplateTaskId, UserId),
 CONSTRAINT AK2_TemplateTaskUser UNIQUE (TemplateTaskId, AgentSort),
 CONSTRAINT FK1_TemplateTaskUser FOREIGN KEY (TemplateTaskId) REFERENCES PMIS.TemplateTask(TemplateTaskId),
 CONSTRAINT FK2_TemplateTaskUser FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK3_TemplateTaskUser FOREIGN KEY (LevelId) REFERENCES PMIS.FileLevel(LevelId)
)/*樣板任務成員*/

CREATE TABLE PMIS.TemplateTaskUserAuthority( /*樣板任務成員權限*/
 TemplateTaskUserId			int					NOT NULL,						--所屬樣板任務成員
 AuthorityId				int					NOT NULL,						--所屬權限
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_TemplateTaskUserAuthority PRIMARY KEY (TemplateTaskUserId, AuthorityId),
 CONSTRAINT FK1_TemplateTaskUserAuthority FOREIGN KEY (TemplateTaskUserId) REFERENCES PMIS.TemplateTaskUser(TemplateTaskUserId),
 CONSTRAINT FK2_TemplateTaskUserAuthority FOREIGN KEY (AuthorityId) REFERENCES PMIS.Authority(AuthorityId)
)/*樣板任務成員權限*/


CREATE TABLE PMIS.TrackFilterDefault( /*專案任務追蹤查詢條件快取*/
 UserId						int					NOT NULL,						--使用者
 FilterName					varchar(255)		NOT NULL,						--快取欄位名稱
 FilterValue				nvarchar(512)		NOT NULL,						--快取欄位值
 LastModifiedBy				int						NULL,						--修改日期
 LastModifiedDate			datetime				NULL,						--修改人員
 CONSTRAINT PK_TrackFilterDefault PRIMARY KEY (UserId, FilterName),
 CONSTRAINT FK1_TrackFilterDefault FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId)
)/*專案任務追蹤查詢條件快取*/


END /*基本資料*/

BEGIN /*專案管理*/
CREATE TABLE PMIS.Project( /*專案資料*/
 ProjectId					int					identity,
 DepartmentId				int						NULL,						--所屬部門
 TeamId						int						NULL,						--MAMO團隊ID
 CompanyType				nvarchar(2)			NOT NULL,						--公司別 >> BAS.[Type] A：中揚；B：紘立
 ProjectType				nvarchar(2)			NOT NULL,						--專案分類 >> BAS.[Type] 01：製造；02：銷售；03：管理；04：研發；11：試作；12：量產；99：其他
 ProjectNo					nvarchar(10)		NOT NULL,						--專案代碼
 ProjectName				nvarchar(100)		NOT NULL,						--專案名稱
 ProjectDesc				nvarchar(300)		NOT NULL,						--專案描述
 ProjectAttribute			nvarchar(2)			NOT NULL,						--專案屬性 >> BAS.[Type] P：專案；O：OKR 
 WorkTimeStatus				nvarchar(2)			NOT NULL,						--工作天模式 >> BAS.[Status]  Y：是 N：否
 WorkTimeInterval			nvarchar(100)			NULL,						--工作時段
 CustomerRemark				nvarchar(500)			NULL,						--客戶備註
 ProjectRemark				nvarchar(500)			NULL,						--專案備註
 ProjectStatus				nvarchar(2)			NOT NULL,						--專案狀態 >> BAS.[Status] P：專案規劃(Planning)；I：專案進行中(In progress)；C：專案完成(Completed)
 PlannedStart				datetime				NULL,						--計畫開始時間
 PlannedEnd					datetime				NULL,						--計畫結束時間
 PlannedDuration			int						NULL,						--計畫持續時間(分鐘)
 EstimateStart				datetime				NULL,						--預估開始時間
 EstimateEnd				datetime				NULL,						--預估結束時間
 EstimateDuration			int						NULL,						--預估持續時間(分鐘)
 ActualStart				datetime				NULL,						--實際開始時間
 ActualEnd					datetime				NULL,						--實際結束時間
 ActualDuration				int						NULL,						--實際持續時間(分鐘)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Project PRIMARY KEY (ProjectId),
 CONSTRAINT AK1_Project UNIQUE (ProjectNo),
 CONSTRAINT FK1_Project FOREIGN KEY (DepartmentId) REFERENCES BAS.Department(DepartmentId),
 CONSTRAINT FK2_Project FOREIGN KEY (TeamId) REFERENCES MAMO.Teams(TeamId)
)/*專案資料*/

CREATE TABLE PMIS.ProjectUser( /*專案成員*/
 ProjectUserId				int					identity,
 ProjectId					int					NOT NULL,						--所屬專案
 UserId						int					NOT NULL,						--所屬使用者
 LevelId					int					NOT NULL,						--所屬檔案等級
 AgentSort					int					NOT NULL,						--代理順序
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ProjectUser PRIMARY KEY (ProjectUserId),
 CONSTRAINT AK1_ProjectUser UNIQUE (ProjectId, UserId),
 CONSTRAINT AK2_ProjectUser UNIQUE (ProjectId, AgentSort),
 CONSTRAINT FK1_ProjectUser FOREIGN KEY (ProjectId) REFERENCES PMIS.Project(ProjectId),
 CONSTRAINT FK2_ProjectUser FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK3_ProjectUser FOREIGN KEY (LevelId) REFERENCES PMIS.FileLevel(LevelId)
)/*專案成員*/

CREATE TABLE PMIS.ProjectUserAuthority( /*專案成員權限*/
 ProjectUserId				int					NOT NULL,						--所屬專案成員
 AuthorityId				int					NOT NULL,						--所屬權限
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ProjectUserAuthority PRIMARY KEY (ProjectUserId, AuthorityId),
 CONSTRAINT FK1_ProjectUserAuthority FOREIGN KEY (ProjectUserId) REFERENCES PMIS.ProjectUser(ProjectUserId),
 CONSTRAINT FK2_ProjectUserAuthority FOREIGN KEY (AuthorityId) REFERENCES PMIS.Authority(AuthorityId)
)/*專案成員權限*/

CREATE TABLE PMIS.ProjectFile( /*專案檔案*/
 ProjectFileId				int					identity,
 ProjectId					int					NOT NULL,						--所屬專案
 FileId						int					NOT NULL,						--所屬檔案
 LevelId					int					NOT NULL,						--所屬檔案等級
 FileType					nvarchar(2)			    NULL,						--檔案類型 >> BAS.[Type] 01：會議記錄；02：合約；99：其他
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ProjectFile PRIMARY KEY (ProjectFileId),
 CONSTRAINT FK1_ProjectFile FOREIGN KEY (ProjectId) REFERENCES PMIS.Project(ProjectId),
 CONSTRAINT FK2_ProjectFile FOREIGN KEY (FileId) REFERENCES BAS.[File](FileId),
 CONSTRAINT FK3_ProjectFile FOREIGN KEY (LevelId) REFERENCES PMIS.FileLevel(LevelId)
)/*專案檔案*/

CREATE TABLE PMIS.ProjectTask( /*專案任務資料*/
 ProjectTaskId				int					identity,
 ProjectId					int					NOT NULL,						--所屬專案
 ParentTaskId				int						NULL,						--父層任務
 ChannelId					int						NULL,						--MAMO頻道ID
 TaskSort					int					NOT NULL,						--任務排序
 TaskName					nvarchar(50)		NOT NULL,						--任務名稱
 TaskDesc					nvarchar(300)			NULL,						--任務描述
 PlannedStart				datetime				NULL,						--計畫開始時間
 PlannedEnd					datetime				NULL,						--計畫結束時間
 PlannedDuration			int						NULL,						--計畫持續時間(分鐘)
 EstimateStart				datetime				NULL,						--預估開始時間
 EstimateEnd				datetime				NULL,						--預估結束時間
 EstimateDuration			int						NULL,						--預估持續時間(分鐘)
 ActualStart				datetime				NULL,						--實際開始時間
 ActualEnd					datetime				NULL,						--實際結束時間
 ActualDuration				int						NULL,						--實際持續時間(分鐘)
 ReplyFrequency				int						NULL,						--回報頻率(分鐘)
 RequiredFile				bit					NOT NULL,	DEFAULT 0,			--是否需要回報檔案
 TaskStatus					nvarchar(2)			NOT NULL,						--任務狀態 >> BAS.[Status] B：待處理(Backlog)；I：進行中(In progress)；T：待檢驗(Testing)；C：已完成(Completed)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ProjectTask PRIMARY KEY (ProjectTaskId),
 CONSTRAINT FK1_ProjectTask FOREIGN KEY (ProjectId) REFERENCES PMIS.Project(ProjectId),
 CONSTRAINT FK2_ProjectTask FOREIGN KEY (ParentTaskId) REFERENCES PMIS.ProjectTask(ProjectTaskId)
)/*專案任務資料*/

CREATE TABLE PMIS.ProjectTaskLink( /*專案任務連結資料*/
 ProjectTaskLinkId			int					identity,
 SourceTaskId				int					NOT NULL,						--來源任務
 TargetTaskId				int					NOT NULL,						--目標任務
 LinkType					nvarchar(2)			NOT NULL,						--連結類型 >> BAS.[Type]  0：結束到開始；1：開始到開始；2：結束到結束；3：開始到結束
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ProjectTaskLink PRIMARY KEY (ProjectTaskLinkId),
 CONSTRAINT AK1_ProjectTaskLink UNIQUE (SourceTaskId, TargetTaskId, LinkType),
 CONSTRAINT FK1_ProjectTaskLink FOREIGN KEY (SourceTaskId) REFERENCES PMIS.ProjectTask(ProjectTaskId),
 CONSTRAINT FK2_ProjectTaskLink FOREIGN KEY (TargetTaskId) REFERENCES PMIS.ProjectTask(ProjectTaskId)
)/*樣板任務連結資料*/

CREATE TABLE PMIS.ProjectTaskFile( /*任務回報檔案*/
 ProjectTaskFileId			int					identity,
 ProjectTaskId				int					NOT NULL,						--所屬專案任務
 FileId						int					NOT NULL,						--所屬檔案
 LevelId					int					NOT NULL,						--所屬檔案等級
 FileType					nvarchar(2)			    NULL,						--檔案類型 >> BAS.[Type] 01：會議記錄；02：合約；99：其他
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ProjectTaskFile PRIMARY KEY (ProjectTaskFileId),
 CONSTRAINT FK1_ProjectTaskFile FOREIGN KEY (ProjectTaskId) REFERENCES PMIS.ProjectTask(ProjectTaskId),
 CONSTRAINT FK2_ProjectTaskFile FOREIGN KEY (FileId) REFERENCES BAS.[File](FileId),
 CONSTRAINT FK3_ProjectTaskFile FOREIGN KEY (LevelId) REFERENCES PMIS.FileLevel(LevelId)
)/*任務回報檔案*/

CREATE TABLE PMIS.ProjectTaskUser( /*專案任務成員*/
 ProjectTaskUserId			int					identity,
 ProjectTaskId				int					NOT NULL,						--所屬專案任務
 UserId						int					NOT NULL,						--所屬使用者
 LevelId					int					NOT NULL,						--所屬檔案等級
 AgentSort					int					NOT NULL,						--代理順序
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ProjectTaskUser PRIMARY KEY (ProjectTaskUserId),
 CONSTRAINT AK1_ProjectTaskUser UNIQUE (ProjectTaskId, UserId),
 CONSTRAINT AK2_ProjectTaskUser UNIQUE (ProjectTaskId, AgentSort),
 CONSTRAINT FK1_ProjectTaskUser FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK2_ProjectTaskUser FOREIGN KEY (LevelId) REFERENCES PMIS.FileLevel(LevelId)
)/*專案任務成員*/

CREATE TABLE PMIS.ProjectTaskUserAuthority( /*專案任務成員權限*/
 ProjectTaskUserId			int					NOT NULL,						--所屬專案任務成員
 AuthorityId				int					NOT NULL,						--所屬權限
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ProjectTaskUserAuthority PRIMARY KEY (ProjectTaskUserId, AuthorityId),
 CONSTRAINT FK1_ProjectTaskUserAuthority FOREIGN KEY (ProjectTaskUserId) REFERENCES PMIS.ProjectTaskUser(ProjectTaskUserId),
 CONSTRAINT FK2_ProjectTaskUserAuthority FOREIGN KEY (AuthorityId) REFERENCES PMIS.Authority(AuthorityId)
)/*專案任務成員權限*/
END /*專案管理*/

BEGIN /*任務管理*/
CREATE TABLE PMIS.PersonalTask( /*個人任務資料*/
 PersonalTaskId				int					identity,
 UserId						int					NOT NULL,						--所屬使用者
 TaskName					nvarchar(50)		NOT NULL,						--任務名稱
 TaskDesc					nvarchar(300)		NOT NULL,						--任務說明
 PlannedStart				datetime			NOT NULL,						--計畫開始時間
 PlannedEnd					datetime			NOT NULL,						--計畫完成時間
 PlannedDuration			int					NOT NULL,						--計畫持續時間(分鐘)
 ActualStart				datetime			NOT NULL,						--實際開始時間
 ActualEnd					datetime			NOT NULL,						--實際完成時間
 ActualDuration				int					NOT NULL,						--實際持續時間(分鐘)
 ReplyFrequency				int					NOT NULL,						--回報頻率(分鐘)
 TaskStatus					nvarchar(2)			NOT NULL,						--任務狀態 >> BAS.[Status] B：待處理(Backlog)；I：進行中(In progress)；T：待檢驗(Testing)；C：已完成(Completed)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_PersonalTask PRIMARY KEY (PersonalTaskId),
 CONSTRAINT FK1_PersonalTask FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId)
)/*個人任務資料*/

CREATE TABLE PMIS.PersonalTaskFile( /*任務回報檔案*/
 PersonalTaskFileId			int					identity,
 PersonalTaskId				int					NOT NULL,						--所屬專案任務
 FileId						int					NOT NULL,						--所屬檔案
 LevelId					int					NOT NULL,						--所屬檔案等級
 FileType					nvarchar(2)			    NULL,						--檔案類型 >> BAS.[Type] 01：會議記錄；02：合約；99：其他
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_PersonalTaskFile PRIMARY KEY (PersonalTaskFileId),
 CONSTRAINT FK1_PersonalTaskFile FOREIGN KEY (PersonalTaskId) REFERENCES PMIS.PersonalTask(PersonalTaskId),
 CONSTRAINT FK2_PersonalTaskFile FOREIGN KEY (FileId) REFERENCES BAS.[File](FileId),
 CONSTRAINT FK3_PersonalTaskFile FOREIGN KEY (LevelId) REFERENCES PMIS.FileLevel(LevelId)
)/*任務回報檔案*/


CREATE TABLE PMIS.TaskUser( /*任務成員-New*/
 TaskUserId					int					identity,
 TaskId						int					NOT NULL,						--個人任務
 TaskType					nvarchar(10)		NOT NULL,						--任務類型 Project|Personal
 UserId						int					NOT NULL,						--任務成員
 LevelId					int					NOT NULL,						--所屬檔案等級
 AgentSort					int					NOT NULL,						--代理順序
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_TaskUser PRIMARY KEY (TaskUserId),
 CONSTRAINT AK1_TaskUser UNIQUE (TaskId, TaskType, UserId),
 CONSTRAINT AK2_TaskUser UNIQUE (TaskId, TaskType, AgentSort),
 CONSTRAINT FK1_TaskUser FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK2_TaskUser FOREIGN KEY (LevelId) REFERENCES PMIS.FileLevel(LevelId)
)/*任務成員-New*/

CREATE TABLE PMIS.TaskUserAuthority( /*任務權限-New*/
 TaskUserId					int					NOT NULL,
 AuthorityId				int					NOT NULL,
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員	
 CONSTRAINT PK_TaskUserAuthority PRIMARY KEY (TaskUserId, AuthorityId),
 CONSTRAINT FK1_TaskUserAuthority FOREIGN KEY (AuthorityId) REFERENCES PMIS.Authority(AuthorityId),
 CONSTRAINT FK2_TaskUserAuthority FOREIGN KEY (TaskUserId) REFERENCES PMIS.TaskUser(TaskUserId)
)/*任務權限-New*/

CREATE TABLE PMIS.TaskReply( /*任務回報內容*/
 ReplyId					int					identity,
 TaskId						int					NOT NULL,						--所屬任務
 ReplyType					nvarchar(10)		NOT NULL,						--回報類型 >> BAS.[Type] Project：專案；Personal：個人
 ReplyContent				nvarchar(MAX)		NOT NULL,						--回報內容
 ReplyUser					int					NOT NULL,						--回報人員
 ReplyDate					datetime			NOT NULL,						--回報日期
 ReplyStatus				nvarchar(2)			NOT NULL,						--回報狀態 >> BAS.[Status] S：開始(Start)；C：完成(Complete)；O：按計劃進行(On schedule)；N：尚未開始(Not started yet)；D：延宕(Delay)
 ReplyFile					int						NULL,						--回報檔案
 DeferredDuration			int						NULL,						--延遲時間(分鐘)
 EventDate					datetime				NULL,						--事件時間 (回報時間)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_TaskReply PRIMARY KEY (ReplyId),
 CONSTRAINT FK1_TaskReply FOREIGN KEY (ReplyUser) REFERENCES BAS.[User](UserId)
)/*任務回報內容*/
END /*任務管理*/


BEGIN /*任務管家*/
CREATE TABLE PMIS.TaskType( /*任務類別*/
 TaskTypeId					int					identity,
 TaskTypeName				nvarchar(100)		NOT NULL,						--任務類別名稱
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_TaskType PRIMARY KEY (TaskTypeId)
)/*任務派放*/

CREATE TABLE PMIS.TaskDistribution( /*任務派放*/
 TdId						int					identity,
 TaskTypeId					int					NOT NULL,						--任務類別ID
 TaskName					nvarchar(MAX)		NOT NULL,						--主旨
 ProductionModel			nvarchar(150)			NULL,						--機種名稱
 CustomerId					int						NULL,						--客戶ID
 DepartmentId				int						NULL,						--部門ID
 TaskContent				nvarchar(MAX)			NULL,						--任務內容
 StartDate					datetime			NOT NULL,						--開始日期
 ExpectedDate				datetime				NULL,						--預計結束日期
 AttendUser					nvarchar(MAX)		NOT NULL,						--參與者
 ReturnFrequency			int						NULL,						--回報頻率(天)
 RepeatType					nvarchar(20)			NULL,						--重複類型(NONE:不重複, DAILY:每天, WEEKLY:每週, INTERVAL:間隔天數)
 RepeatValue				int						NULL,						--重複值(若RepeatType=WEEKLY則存星期幾1-7, RepeatType=INTERVAL則存間隔天數)
 RepeatEndType				nvarchar(20)			NULL,						--結束類型(NEVER:永不結束, COUNT:重複次數, UNTIL:直到某日)
 RepeatEndValue				nvarchar(50)			NULL,						--結束值(若RepeatEndType=COUNT則存次數, RepeatEndType=UNTIL則存結束日期)
 FileList					nvarchar(MAX)			NULL,						--附檔檔案
 ImgList					nvarchar(MAX)			NULL,						--附檔圖片
 [Status]					nvarchar(2)			NOT NULL,						--任務狀態 >> N：尚未開始; S：開始；D：即將到期(預計結束日前兩天)；C：完成；V：作廢
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_TaskDistribution PRIMARY KEY (TdId),
 CONSTRAINT FK1_TaskDistribution FOREIGN KEY (TaskTypeId) REFERENCES PMIS.TaskType(TaskTypeId),
 CONSTRAINT FK2_TaskDistribution FOREIGN KEY (CustomerId) REFERENCES SCM.Customer(CustomerId),
 CONSTRAINT FK3_TaskDistribution FOREIGN KEY (DepartmentId) REFERENCES BAS.Department(DepartmentId),
)/*任務派放*/

CREATE TABLE PMIS.MissionReport( /*任務回報*/
 MrId						int					identity,
 TdId						int					NOT NULL,						--所屬任務
 Sort						int					NOT NULL,						--任務排序
 ReportUserId				int					NOT NULL,						--回報者
 [GetDate]					datetime			NOT	NULL,						--接收時間
 StartDate					datetime				NULL,						--開始時間
 EndDate					datetime				NULL,						--結案時間
 NextReportDate				datetime				NULL,						--下次回報時間
 [Status]					nvarchar(2)			NOT NULL,						--任務狀態 >> N:未開始;S:進行中;T:追蹤中;E:結案
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MissionReport PRIMARY KEY (MrId),
 CONSTRAINT FK1_MissionReport FOREIGN KEY (TdId) REFERENCES PMIS.TaskDistribution(TdId),
 CONSTRAINT FK2_MissionReport FOREIGN KEY (ReportUserId) REFERENCES BAS.[User](UserId)
)/*任務派放*/

CREATE TABLE PMIS.MrDetail( /*任務回報*/
 MrDetailId					int					identity,
 MrId						int					NOT NULL,						--所屬任務回報者
 TdId						int					NOT NULL,						--所屬任務
 ReportContent				nvarchar(MAX)		NOT NULL,						--回報內容
 ReportDate					datetime				NULL,						--回報日期(確認日)
 FileList					nvarchar(MAX)			NULL,						--回報檔案
 ImgList					nvarchar(MAX)			NULL,						--回報圖片
 [Status]					nvarchar(2)			NOT NULL,						--回報狀態 >> Y:已確認/N:未確認
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MrDetail PRIMARY KEY (MrDetailId),
 CONSTRAINT FK1_MrDetail FOREIGN KEY (MrId) REFERENCES PMIS.MissionReport(MrId),
 CONSTRAINT FK2_MrDetail FOREIGN KEY (TdId) REFERENCES PMIS.TaskDistribution(TdId)
)/*任務派放*/

END /*任務管家*/
