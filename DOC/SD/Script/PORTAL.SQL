BEGIN /*基本資料*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'TemplateTaskUserAuthority') > 0 DROP TABLE PMIS.TemplateTaskUserAuthority
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'TemplateTaskUser') > 0 DROP TABLE PMIS.TemplateTaskUser
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'TemplateTask') > 0 DROP TABLE PMIS.TemplateTask
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Template') > 0 DROP TABLE PMIS.Template
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Authority') > 0 DROP TABLE PMIS.Authority
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'FileLevel') > 0 DROP TABLE PMIS.FileLevel
END /*基本資料*/

BEGIN /*專案管理*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ProjectTaskUserAuthority') > 0 DROP TABLE PMIS.ProjectTaskUserAuthority
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ProjectTaskUser') > 0 DROP TABLE PMIS.ProjectTaskUser
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ProjectTask') > 0 DROP TABLE PMIS.ProjectTask
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ProjectFile') > 0 DROP TABLE PMIS.ProjectFile
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ProjectUserAuthority') > 0 DROP TABLE PMIS.ProjectUserAuthority
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ProjectUser') > 0 DROP TABLE PMIS.ProjectUser
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Project') > 0 DROP TABLE PMIS.Project
END /*專案管理*/

BEGIN /*任務管理*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'TaskReplyFile') > 0 DROP TABLE PMIS.TaskReplyFile
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'TaskReply') > 0 DROP TABLE PMIS.TaskReply
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'PersonalTask') > 0 DROP TABLE PMIS.PersonalTask
END /*任務管理*/
--******************************************************************************************************************
BEGIN /*基本資料*/
CREATE TABLE PMIS.Authority( /*權限項目*/ 
 AuthorityId				int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
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
 CONSTRAINT AK1_Authority UNIQUE (AuthorityCode),
 CONSTRAINT AK2_Authority UNIQUE (AuthorityName),
 CONSTRAINT FK1_Authority FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)/*專案權限*/

CREATE TABLE PMIS.FileLevel( /*檔案等級*/
 LevelId					int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
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
 CONSTRAINT AK2_FileLevel UNIQUE (LevelName),
 CONSTRAINT FK1_FileLevel FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)/*檔案等級*/

CREATE TABLE PMIS.Template( /*樣板資料*/
 TemplateId					int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 TemplateName				nvarchar(100)		NOT NULL,						--樣板名稱
 TemplateDesc				nvarchar(300)		NOT NULL,						--樣板描述
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Template PRIMARY KEY (TemplateId),
 CONSTRAINT FK1_Template FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
)/*樣板資料*/

CREATE TABLE PMIS.TemplateTask( /*樣板任務資料*/
 TemplateTaskId				int					identity,
 TemplateId					int					NOT NULL,						--所屬樣板
 ParentTaskId				int					NOT NULL,						--父層任務
 TaskSort					int					NOT NULL,						--任務排序
 TaskName					nvarchar(50)		NOT NULL,						--任務名稱
 Duration					int					NOT NULL,						--持續時間(分鐘)
 ReplyFrequency				int					NOT NULL,						--回報頻率(分鐘)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_TemplateTask PRIMARY KEY (TemplateTaskId),
 CONSTRAINT AK1_TemplateTask UNIQUE (ParentTaskId, TaskSort),
 CONSTRAINT FK1_TemplateTask FOREIGN KEY (TemplateId) REFERENCES PMIS.Template(TemplateId)
)/*樣板任務資料*/

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
END /*基本資料*/

BEGIN /*專案管理*/
CREATE TABLE PMIS.Project( /*專案資料*/
 ProjectId					int					identity,
 CompanyId					int					NOT NULL,						--所屬公司
 CustomerId					int						NULL,						--所屬客戶
 DepartmentId				int						NULL,						--所屬部門
 ProjectType				nvarchar(2)			NOT NULL,						--專案分類 >> BAS.[Type] 01：製造；02：銷售；03：管理；04：研發；11：試作；12：量產；99：其他
 ProjectNo					nvarchar(10)		NOT NULL,						--專案代碼
 ProjectName				nvarchar(100)		NOT NULL,						--專案名稱
 ProjectDesc				nvarchar(100)		NOT NULL,						--專案描述
 ProjectAttribute			nvarchar(2)			NOT NULL,						--專案屬性 >> BAS.[Type] P：專案；O：OKR
 ProjectStatus				nvarchar(2)			NOT NULL,						--專案狀態 >> BAS.[Status] P：專案規劃(Planning)；I：專案進行中(In progress)；C：專案完成(Completed)
 WorkTime					nvarchar(2)			NOT NULL,						--是否計算工作天 >> BAS.[Status]  Y:是 N:否
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Project PRIMARY KEY (ProjectId),
 CONSTRAINT AK1_Project UNIQUE (ProjectNo),
 CONSTRAINT FK1_Project FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
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
 FileId						int					NOT NULL,						--所屬檔案
 LevelId					int					NOT NULL,						--所屬檔案等級
 FileType					nvarchar(2)			NOT NULL,						--檔案類型 >> BAS.[Type] 01：會議記錄；02：合約
 FileDesc					nvarchar(100)		NOT NULL,						--檔案描述
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ProjectFile PRIMARY KEY (ProjectFileId),
 CONSTRAINT FK1_ProjectFile FOREIGN KEY (FileId) REFERENCES BAS.[File](FileId),
 CONSTRAINT FK2_ProjectFile FOREIGN KEY (LevelId) REFERENCES PMIS.FileLevel(LevelId)
)/*專案檔案*/

CREATE TABLE PMIS.ProjectTask( /*專案任務資料*/
 ProjectTaskId				int					identity,
 ProjectId					int					NOT NULL,						--所屬專案
 ParentTaskId				int					NOT NULL,						--父層任務
 TaskSort					int					NOT NULL,						--任務排序
 TaskName					nvarchar(50)		NOT NULL,						--任務名稱
 PlanningStart				datetime			NOT NULL,						--計畫開始時間
 PlanningEnd				datetime			NOT NULL,						--計畫結束時間
 StartDate					datetime			NOT NULL,						--開始時間
 EndDate					datetime			NOT NULL,						--結束時間
 Duration					int					NOT NULL,						--持續時間(分鐘)
 ReplyFrequency				int					NOT NULL,						--回報頻率(分鐘)
 TaskStatus					nvarchar(2)			NOT NULL,						--任務狀態 >> BAS.[Status] B：待處理(Backlog)；I：進行中(In progress)；T：待檢驗(Testing)；C：已完成(Completed)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ProjectTask PRIMARY KEY (ProjectTaskId),
 CONSTRAINT AK1_ProjectTask UNIQUE (ParentTaskId, TaskSort),
 CONSTRAINT FK1_ProjectTask FOREIGN KEY (ProjectId) REFERENCES PMIS.Project(ProjectId)
)/*專案任務資料*/

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
 CONSTRAINT FK1_ProjectTaskUser FOREIGN KEY (ProjectTaskId) REFERENCES PMIS.ProjectTask(ProjectTaskId),
 CONSTRAINT FK2_ProjectTaskUser FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK3_ProjectTaskUser FOREIGN KEY (LevelId) REFERENCES PMIS.FileLevel(LevelId)
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
 TaskDesc					nvarchar(200)		NOT NULL,						--任務說明
 PlanningStart				datetime			NOT NULL,						--計畫開始時間
 PlanningEnd				datetime			NOT NULL,						--計畫結束時間
 StartDate					datetime			NOT NULL,						--開始時間
 EndDate					datetime			NOT NULL,						--結束時間
 Duration					int					NOT NULL,						--持續時間(分鐘)
 ReplyFrequency				int					NOT NULL,						--回報頻率(分鐘)
 TaskStatus					nvarchar(2)			NOT NULL,						--任務狀態 >> BAS.[Status] B：待處理(Backlog)；I：進行中(In progress)；T：待檢驗(Testing)；C：已完成(Completed)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_PersonalTask PRIMARY KEY (PersonalTaskId),
 CONSTRAINT FK1_PersonalTask FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId)
)/*個人任務資料*/

CREATE TABLE PMIS.TaskReply( /*任務回報內容*/
 ReplyId					int					identity,
 TaskId						int					NOT NULL,						--所屬任務
 ReplyType					nvarchar(2)			NOT NULL,						--回報類型 >> BAS.[Type] Project：專案；Personal：個人
 ReplyContent				nvarchar(MAX)		NOT NULL,						--回報內容
 ReplyUser					int					NOT NULL,						--回報人員
 ReplyDate					datetime			NOT NULL,						--回報日期
 AheadOfSchedule			nvarchar(2)			NOT NULL,						--是否提前完成 >> BAS.[Status] Y：是；N：否
 Deferred					nvarchar(2)			NOT NULL,						--是否延期 >> BAS.[Status] Y：是；N：否
 DeferredDuration			int						NULL,						--延遲時間(分鐘)
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_TaskReply PRIMARY KEY (ReplyId),
 CONSTRAINT FK1_TaskReply FOREIGN KEY (ReplyUser) REFERENCES BAS.[User](UserId)
)/*任務回報內容*/

CREATE TABLE PMIS.TaskReplyFile( /*任務回報檔案*/
 ReplyId					int					NOT NULL,						--所屬任務回報
 FileId						int					NOT NULL,						--所屬檔案
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_TaskReplyFile PRIMARY KEY (ReplyId, FileId),
 CONSTRAINT FK1_TaskReplyFile FOREIGN KEY (ReplyId) REFERENCES PMIS.TaskReply(ReplyId),
 CONSTRAINT FK2_TaskReplyFile FOREIGN KEY (FileId) REFERENCES BAS.[File](FileId)
)/*任務回報檔案*/
END /*任務管理*/