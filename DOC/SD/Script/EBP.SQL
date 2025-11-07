BEGIN /*基本資料*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'BoardType') > 0 DROP TABLE EBP.BoardType
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Annual') > 0 DROP TABLE EBP.Annual
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ClubJobType') > 0 DROP TABLE EBP.ClubJob
END /*基本資料*/

BEGIN /*公佈欄管理*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'BoardFile') > 0 DROP TABLE EBP.BoardFile
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Board') > 0 DROP TABLE EBP.Board
END /*公佈欄管理*/

BEGIN /*積分管理*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Points') > 0 DROP TABLE EBP.Points
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'RedeemPrize') > 0 DROP TABLE EBP.RedeemPrize
END /*積分管理*/

BEGIN /*活動管理*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ParticipantPoint') > 0 DROP TABLE EBP.ParticipantPoint
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Participant') > 0 DROP TABLE EBP.Participant
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ActivityPoint') > 0 DROP TABLE EBP.ActivityPoint
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'AwardsUser') > 0 DROP TABLE EBP.AwardsUser
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Awards') > 0 DROP TABLE EBP.Awards
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Activity') > 0 DROP TABLE EBP.Activity
END /*活動管理*/

BEGIN /*社團管理*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ClubMember') > 0 DROP TABLE EBP.ClubMember
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Club') > 0 DROP TABLE EBP.Club
END /*社團管理*/

BEGIN /*遊戲管理*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'BingoLotteryRecord') > 0 DROP TABLE EBP.BingoLotteryRecord
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'BingoRound') > 0 DROP TABLE EBP.BingoRound
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'BingoMap') > 0 DROP TABLE EBP.BingoMap
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Bingo') > 0 DROP TABLE EBP.Bingo
END /*遊戲管理*/

BEGIN /*表單管理*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ClubBudgetParticipant') > 0 DROP TABLE EBP.ClubBudgetParticipant
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ClubBudgetDetail') > 0 DROP TABLE EBP.ClubBudgetDetail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'ClubBudgetSubsidy') > 0 DROP TABLE EBP.ClubBudgetSubsidy
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DinnerParticipant') > 0 DROP TABLE EBP.DinnerParticipant
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DinnerSubsidy') > 0 DROP TABLE EBP.DinnerSubsidy
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'SubsidyCertificate') > 0 DROP TABLE EBP.SubsidyCertificate
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'SubsidyRange') > 0 DROP TABLE EBP.SubsidyRange
END /*表單管理*/

BEGIN /*訂餐系統*/
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'Restaurant') > 0 DROP TABLE EBP.Restaurant
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'RestaurantFile') > 0 DROP TABLE EBP.RestaurantFile
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'RestaurantMeal') > 0 DROP TABLE EBP.RestaurantMeal
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MealAttribute') > 0 DROP TABLE EBP.MealCategory
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'MaDetail') > 0 DROP TABLE EBP.McDetail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'DailyMeal') > 0 DROP TABLE EBP.DailyMeal
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'OrderMealSteward') > 0 DROP TABLE EBP.OrderMealSteward
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'UmoAdditional') > 0 DROP TABLE EBP.UmoAdditional
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'UmoDetail') > 0 DROP TABLE EBP.UmoDetail
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'UserMealOrder') > 0 DROP TABLE EBP.UserMealOrder
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'UserMealOrderLog') > 0 DROP TABLE EBP.UserMealOrderLog
IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'OrderSetting') > 0 DROP TABLE EBP.OrderSetting

END /*訂餐系統*/
--******************************************************************************************************************
BEGIN /*基本資料*/
CREATE TABLE EBP.BoardType( /*公告類別*/
 BoardTypeId				int					identity,
 TypeName					nvarchar(100)		NOT NULL,						--類別名稱
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime			NULL,							--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int					NULL,							--修改人員
 CONSTRAINT PK_BoardType PRIMARY KEY (BoardTypeId)
)/*公告類別*/

CREATE TABLE EBP.Annual(/*年份*/
 AnnualId					int					identity,
 Annual						int					NOT NULL,						--年份
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Annual PRIMARY KEY (AnnualId)
)/*年份*/

CREATE TABLE EBP.ClubJob(/*社團職位*/
 ClubJobId					int					identity,
 JobName					nvarchar(50)		NOT NULL,						--職位名稱
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ClubJob PRIMARY KEY (ClubJobId)
)/*社團職位*/
END /*基本資料*/

BEGIN/*公佈欄管理*/
CREATE TABLE EBP.Board( /*公佈欄資料*/
 BoardId					int					identity,
 Title						nvarchar(100)		NOT NULL,						--公告標題
 Content					nvarchar(MAX)		NOT NULL,						--公告內容
 AnnounceDate				datetime			NOT NULL,						--公告日期
 BoardTypeId				int					NOT NULL,						--所屬類別
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Board PRIMARY KEY (BoardId), 
 CONSTRAINT FK1_Board FOREIGN KEY (BoardTypeId) REFERENCES EBP.BoardType(BoardTypeId),
)/*公佈欄資料*/

CREATE TABLE EBP.BoardFile( /*公佈欄檔案*/
 BoardFileId				int					identity,
 BoardId					int					NOT NULL,						--所屬公告
 FileId						int					NOT NULL,						--所屬檔案
 FileType					nvarchar(5)			NOT NULL,						--檔案類別 >> BAS.[Type] File：文件；Image：圖片；Form：表單；Other：其他
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_BoardFile PRIMARY KEY (BoardFileId), 
 CONSTRAINT FK1_BoardFile FOREIGN KEY (BoardId) REFERENCES EBP.Board(BoardId),
 CONSTRAINT FK2_BoardFile FOREIGN KEY (FileId) REFERENCES BAS.[File](FileId)
)/*公佈欄檔案*/
END/*公佈欄管理*/

BEGIN /*積分管理*/
CREATE TABLE EBP.RedeemPrize(/*兌換獎勵*/
 RedeemPrizeId				int					identity,
 PrizeName					nvarchar(100)		NOT NULL,						--獎勵名稱
 RedeemPoint				int					NOT NULL,						--兌換點數
 AnnualId					int					NOT NULL,						--所屬年份
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_RedeemPrize PRIMARY KEY (RedeemPrizeId),
 CONSTRAINT FK1_RedeemPrize FOREIGN KEY (AnnualId) REFERENCES EBP.Annual(AnnualId)
)/*點數兌換*/

CREATE TABLE EBP.Point(/*人員點數*/
 PointId					int					identity,
 UserId						int					NOT NULL,						--所屬人員
 AnnualId					int					NOT NULL,						--所屬年份
 PointQty					int					NOT NULL,						--點數數量
 PointType					nvarchar(10)		NOT NULL,						--點數類別 >> BAS.[Type] Activity：活動；Club：社團；Redeem：兌換獎勵；Other：其他
 ActivityId					int						NULL,						--所屬活動
 ClubId						int						NULL,						--所屬社團
 RedeemPrizeId				int						NULL,						--所屬兌換獎勵
 Remark						nvarchar(500)			NULL,						--備註
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Point PRIMARY KEY (PointId),
 CONSTRAINT FK1_Point FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK2_Point FOREIGN KEY (AnnualId) REFERENCES EBP.Annual(AnnualId),
 CONSTRAINT FK3_Point FOREIGN KEY (ActivityId) REFERENCES EBP.Activity(ActivityId),
 CONSTRAINT FK4_Point FOREIGN KEY (ClubId) REFERENCES EBP.Club(ClubId),
 CONSTRAINT FK5_Point FOREIGN KEY (RedeemPrizeId) REFERENCES EBP.RedeemPrize(RedeemPrizeId)
)/*人員點數*/
END /*積分管理*/

BEGIN /*活動管理*/
CREATE TABLE EBP.Activity( /*活動*/
 ActivityId					int					identity,
 ActivityName				nvarchar(100)		NOT NULL,						--活動名稱
 ActivityStartDate			datetime			NOT NULL,						--活動開始日期
 ActivityEndDate			datetime			NOT NULL,						--活動結束日期
 ActivityContent			nvarchar(255)		NOT NULL,						--活動內容
 AnnualId					int					NOT NULL,						--所屬年份
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Activity PRIMARY KEY (ActivityId),
 CONSTRAINT FK1_Activity FOREIGN KEY (AnnualId) REFERENCES EBP.Annual(AnnualId)
)/*活動*/

CREATE TABLE EBP.Awards( /*活動獎項*/
 AwardsId					int					identity,
 ActivityId					int					NOT NULL,						--所屬活動
 AwardsName					nvarchar(100)		NOT NULL,						--獎項名稱
 AwardsPrice				int						NULL,						--獎項金額
 AwardsQuantity				int						NULL,						--獎項數量
 AwardsUnit					nvarchar(50)			NULL,						--獎項數量單位
 AwardsRemark				nvarchar(500)			NULL,						--獎項備註
 SortNumber					int					NOT NULL,						--排序號碼
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Awards PRIMARY KEY (AwardsId),
 CONSTRAINT AK1_Awards UNIQUE (ActivityId, AwardsName),
 CONSTRAINT FK1_Awards FOREIGN KEY (ActivityId) REFERENCES EBP.Activity(ActivityId)
)/*活動獎項*/

CREATE TABLE EBP.AwardsUser( /*獎項人員*/
 AwardsId					int					NOT NULL,						--所屬活動
 UserId						int					NOT NULL,						--所屬人員
 GroupName					nvarchar(50)			NULL,						--群組名稱
 Remark						nvarchar(500)			NULL,						--備註
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT FK1_AwardsUser FOREIGN KEY (AwardsId) REFERENCES EBP.Awards(AwardsId),
 CONSTRAINT FK2_AwardsUser FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId)
)/*獎項人員*/

CREATE TABLE EBP.ActivityPoint( /*活動點數*/
 ActivityPointId			int					identity,
 ActivityId					int					NOT NULL,						--所屬活動
 PointType					nvarchar(5)				NULL,						--點數類型 >> BAS.[Type] Basic：基本；Bonus：獎勵；Other：其他
 PointQty					int					NOT NULL,						--點數數量
 Remark						nvarchar(500)			NULL,						--備註
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ActivityPoint PRIMARY KEY (ActivityPointId),
 CONSTRAINT FK1_ActivityPoint FOREIGN KEY (ActivityId) REFERENCES EBP.Activity(ActivityId),
)/*活動點數*/

CREATE TABLE EBP.Participant(/*活動參與人員*/
 ParticipantId				int					identity,
 UserId						int					NOT NULL,						--所屬人員
 ActivityId					int					NOT NULL,						--所屬活動
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼 >> Y：已發放點數；N：未發放點數
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Participant PRIMARY KEY (ParticipantId),
 CONSTRAINT FK1_Participant FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK2_Participant FOREIGN KEY (ActivityId) REFERENCES EBP.Activity(ActivityId)
)/*活動參與人員*/

CREATE TABLE EBP.ParticipantPoint(/*活動人員點數*/
 ParticipantId				int					NOT NULL,						--所屬參與人員
 ActivityPointId			int					NOT NULL,						--所屬活動點數
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 CreateBy					int					NOT NULL,						--新增人員
 CONSTRAINT AK1_ParticipantPoint UNIQUE (ParticipantId, ActivityPointId),
 CONSTRAINT FK1_ParticipantPoint FOREIGN KEY (ParticipantId) REFERENCES EBP.Participant(ParticipantId),
 CONSTRAINT FK2_ParticipantPoint FOREIGN KEY (ActivityPointId) REFERENCES EBP.ActivityPoint(ActivityPointId)
)/*活動人員點數*/
END /*活動管理*/

BEGIN /*社團管理*/
CREATE TABLE EBP.Club( /*社團*/
 ClubId						int					identity,
 ClubName					nvarchar(100)		NOT NULL,						--社團名稱
 ClubApplicant				int					NOT NULL,						--申請人
 EstablishedDate			datetime			NOT NULL,						--創立時間
 ClubProperty				nvarchar(MAX)			NULL,						--社團性質 >> BAS.[Type]  1：自治類；2：學術類；3：才藝類；4：文學類；5：音樂類；6：服務類；7：體育類；99：其他
 ClubGoal					nvarchar(MAX)			NULL,						--社團宗旨
 Appropriation				nvarchar(MAX)			NULL,						--經費說明
 ActiveRegion				nvarchar(MAX)			NULL,						--活動地點
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Club PRIMARY KEY (ClubId),
 CONSTRAINT FK1_Club FOREIGN KEY (ClubApplicant) REFERENCES BAS.[User](UserId)
)/*社團*/

CREATE TABLE EBP.ClubMember(/*社團成員*/
 MemberId					int					identity,
 ClubId						int					NOT NULL,						--所屬社團
 UserId						int					NOT NULL,						--所屬人員
 ClubJobId					int						NULL,						--所屬職位
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ClubMember PRIMARY KEY (MemberId),
 CONSTRAINT FK1_ClubMember FOREIGN KEY (ClubId) REFERENCES EBP.Club(ClubId),
 CONSTRAINT FK2_ClubMember FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK3_ClubMember FOREIGN KEY (ClubJobId) REFERENCES EBP.ClubJob(ClubJobId)
)/*社團成員*/
END /*社團管理*/

BEGIN /*遊戲管理*/
CREATE TABLE EBP.Bingo( /*賓果*/
 BingoId					int					identity,
 BingoNo					nvarchar(100)		NOT NULL,						--賓果代號
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 RoundId					int						NULL,						--所屬賓果開局
 WinStatus					nvarchar(2)			NOT NULL,						--中獎狀態
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Bingo PRIMARY KEY (BingoId),
 CONSTRAINT AK1_Bingo UNIQUE (BingoNo)
)/*賓果*/

CREATE TABLE EBP.BingoMap( /*賓果分佈圖*/
 MapId						int					identity,
 BingoId					int					NOT NULL,						--所屬活動
 MapContent					nvarchar(100)		NOT NULL,						--內容
 MapIndex					int					NOT NULL,						--位置
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_BingoMap PRIMARY KEY (MapId),
 CONSTRAINT FK1_BingoMap FOREIGN KEY (BingoId) REFERENCES EBP.Bingo(BingoId),
)/*賓果分佈圖*/

CREATE TABLE EBP.BingoRound( /*賓果開局*/
 RoundId					int					identity,
 RoundName					nvarchar(100)		NOT NULL,						--開局名稱
 RoundDate					datetime			NOT NULL,						--開局時間
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_BingoRound PRIMARY KEY (RoundId)
)/*賓果開局*/

CREATE TABLE EBP.BingoLotteryRecord( /*賓果開獎記錄*/
 RecordId					int					identity,
 RoundId					int					NOT NULL,						--所屬賓果開局
 RecordContent				nvarchar(100)		NOT NULL,						--開獎內容
 RecordSeq					int					NOT NULL,						--開獎順序
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_BingoLotteryRecord PRIMARY KEY (RecordId)
)/*賓果開獎記錄*/
END /*遊戲管理*/

BEGIN /*表單管理*/
CREATE TABLE EBP.SubsidyRange( /*補助區間*/
 SubsidyRangeId				int					identity,
 AnnualId					int					NOT NULL,						--所屬年份
 RangeType					nvarchar(10)		NOT NULL,						--區間類型 >> BAS.[Type] Dinner：聚餐；Club：社團
 RangeName					nvarchar(100)		NOT NULL,						--區間名稱
 StartDate					datetime			NOT NULL,						--開始日期
 EndDate					datetime			NOT NULL,						--結束日期
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_SubsidyRange PRIMARY KEY (SubsidyRangeId),
 CONSTRAINT FK1_SubsidyRange FOREIGN KEY (AnnualId) REFERENCES EBP.Annual(AnnualId)
)/*補助區間*/

CREATE TABLE EBP.SubsidyCertificate( /*補助憑證*/
 CertificateId				int					identity,
 CertificateType			nvarchar(10)		NOT NULL,						--憑證類型 >> BAS.[Type] Dinner：聚餐；Club：社團
 FileId						int					NOT NULL,						--所屬檔案
 DinnerSubsidyId			int						NULL,						--所屬聚餐補助
 ClubBudgetSubsidyId		int						NULL,						--所屬社團經費補助
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_SubsidyCertificate PRIMARY KEY (CertificateId), 
 CONSTRAINT FK1_SubsidyCertificate FOREIGN KEY (FileId) REFERENCES BAS.[File](FileId),
 CONSTRAINT FK2_SubsidyCertificate FOREIGN KEY (DinnerSubsidyId) REFERENCES EBP.DinnerSubsidy(DinnerSubsidyId),
 CONSTRAINT FK3_SubsidyCertificate FOREIGN KEY (ClubBudgetSubsidyId) REFERENCES EBP.ClubBudgetSubsidy(ClubBudgetSubsidyId)
)/*補助憑證*/

CREATE TABLE EBP.DinnerSubsidy( /*聚餐補助*/
 DinnerSubsidyId			int					identity,
 SubsidyRangeId				int					NOT NULL,						--所屬補助區間
 Amount						int					NOT NULL,						--補助金額
 ApplyDate					datetime			NOT NULL,						--申請時間
 ApplyId					int					NOT NULL,						--申請人
 PayeeId					int					NOT NULL,						--領款人
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼 >> Y：已補助；N：未補助
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DinnerSubsidy PRIMARY KEY (DinnerSubsidyId),
 CONSTRAINT FK1_DinnerSubsidy FOREIGN KEY (SubsidyRangeId) REFERENCES EBP.SubsidyRange(SubsidyRangeId),
 CONSTRAINT FK2_DinnerSubsidy FOREIGN KEY (ApplyId) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK3_DinnerSubsidy FOREIGN KEY (PayeeId) REFERENCES BAS.[User](UserId)
)/*聚餐補助*/

CREATE TABLE EBP.DinnerParticipant( /*聚餐補助人員*/
 ParticipantId				int					identity,
 DinnerSubsidyId			int					NOT NULL,						--所屬聚餐補助
 UserId						int					NOT NULL,						--所屬人員
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_DinnerParticipant PRIMARY KEY (ParticipantId),
 CONSTRAINT FK1_DinnerParticipant FOREIGN KEY (DinnerSubsidyId) REFERENCES EBP.DinnerSubsidy(DinnerSubsidyId),
 CONSTRAINT FK2_DinnerParticipant FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId)
)/*聚餐補助人員*/

CREATE TABLE EBP.ClubBudgetSubsidy( /*社團經費補助*/
 ClubBudgetSubsidyId		int					identity,
 ClubId						int					NOT NULL,						--所屬社團
 SubsidyRangeId				int					NOT NULL,						--所屬補助區間
 Amount						int					NOT NULL,						--補助金額
 ApplyDate					datetime			NOT NULL,						--申請時間
 ApplyId					int					NOT NULL,						--申請人
 PayeeId					int					NOT NULL,						--領款人
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼 >> Y：已補助；N：未補助
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ClubBudgetSubsidy PRIMARY KEY (ClubBudgetSubsidyId),
 CONSTRAINT FK1_ClubBudgetSubsidy FOREIGN KEY (ClubId) REFERENCES EBP.Club(ClubId),
 CONSTRAINT FK2_ClubBudgetSubsidy FOREIGN KEY (SubsidyRangeId) REFERENCES EBP.SubsidyRange(SubsidyRangeId),
 CONSTRAINT FK3_ClubBudgetSubsidy FOREIGN KEY (ApplyId) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK4_ClubBudgetSubsidy FOREIGN KEY (PayeeId) REFERENCES BAS.[User](UserId)
)/*社團經費補助*/

CREATE TABLE EBP.ClubBudgetDetail( /*社團經費明細*/
 ClubBudgetDetailId			int					identity,
 ClubBudgetSubsidyId		int					NOT NULL,						--所屬社團經費補助
 OccurrenceDate				datetime			NOT NULL,						--發生日期
 DetailDesc					nvarchar(MAX)		NOT NULL,						--內容說明
 Player						int					NOT NULL,						--參與人數
 Amount						int					NOT NULL,						--金額
 ActiveRegion				nvarchar(MAX)			NULL,						--活動地點
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ClubBudgetDetail PRIMARY KEY (ClubBudgetDetailId),
 CONSTRAINT FK1_ClubBudgetDetail FOREIGN KEY (ClubBudgetSubsidyId) REFERENCES EBP.ClubBudgetSubsidy(ClubBudgetSubsidyId),
)

CREATE TABLE EBP.ClubBudgetParticipant( /*社團經費補助人員*/
 ParticipantId				int					identity,
 ClubBudgetSubsidyId		int					NOT NULL,						--所屬聚餐補助
 UserId						int					NOT NULL,						--所屬人員
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_ClubBudgetParticipant PRIMARY KEY (ParticipantId),
 CONSTRAINT FK1_ClubBudgetParticipant FOREIGN KEY (ClubBudgetSubsidyId) REFERENCES EBP.ClubBudgetSubsidy(ClubBudgetSubsidyId),
 CONSTRAINT FK2_ClubBudgetParticipant FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId)
)/*社團經費補助人員*/
END /*表單管理*/

BEGIN /*訂餐系統*/
CREATE TABLE EBP.Restaurant( /*餐廳基本資料*/
 RestaurantId				int					identity,						
 CompanyId					int					NOT NULL,						--所屬公司ID
 RestaurantNo				nvarchar(50)		NOT NULL,						--餐廳代號
 RestaurantName				nvarchar(200)		NOT NULL,						--餐廳名稱
 GuiNumber					nvarchar(20)		NOT NULL,						--統一編號
 ResponsiblePerson			nvarchar(30)		NOT NULL,						--廠商負責人
 Contact					nvarchar(30)		NOT	NULL,						--廠商聯絡人
 TelNo						nvarchar(20)		NOT NULL,						--廠商聯絡電話
 FaxNo						nvarchar(20)			NULL,						--廠商聯絡傳真
 Email						nvarchar(60)			NULL,						--廠商電子郵件
 [Address]					nvarchar(255)		NOT NULL,						--廠商聯絡地址
 AccountDay					nvarchar(2)			NOT NULL,						--結帳日
 PaymentType				nvarchar(1)			NOT NULL,						--付款方式 >> BAS.[Type] 1：現金；2：電匯；3：支票；4：其他
 InvoiceCount				nvarchar(10)		NOT NULL,						--發票聯數--(SELECT * FROM BAS.[Type] WHERE TypeSchema = 'Restaurant.InvoiceCount') 1:二聯式, 2:三聯式
 RemitBank					nvarchar(20)		NOT NULL,						--匯款銀行
 RemitAccount				nvarchar(30)		NOT NULL,						--匯款帳號
 RestaurantRemark			nvarchar(500)			NULL,						--廠商備註
 StartTime					nvarchar(5)			NOT NULL	DEFAULT GETDATE(),	--開始日期
 EndTime					nvarchar(5)			NOT NULL	DEFAULT GETDATE(),	--結束日期
 [Status]					nvarchar(2)			NOT NULL    DEFAULT 'A',		--狀態碼
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_Restaurant PRIMARY KEY (RestaurantId),
 CONSTRAINT FK1_Restaurant FOREIGN KEY (CompanyId) REFERENCES BAS.Company(CompanyId)
) /*餐廳基本資料*/

CREATE TABLE EBP.RestaurantFile( /*餐廳檔案*/
 FileId						int					identity,
 RestaurantId				int					NOT NULL,						--所屬餐廳
 FileDoc					int					NOT NULL,						--餐廳檔案
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_RestaurantFile PRIMARY KEY (FileId),
 CONSTRAINT FK1_RestaurantFile FOREIGN KEY (FileDoc) REFERENCES BAS.[File](FileId),
 CONSTRAINT FK2_RestaurantFile FOREIGN KEY (RestaurantId) REFERENCES EBP.Restaurant(RestaurantId)
) /*餐廳檔案*/

CREATE TABLE EBP.RestaurantMeal( /*餐廳餐點*/
 MealId						int					identity,
 RestaurantId				int					NOT NULL,						--所屬餐廳
 MealName					nvarchar(30)		NOT NULL,						--餐點名稱
 MealImage					int						NULL,						--餐點圖片
 MealPrice					float				NOT NULL,						--餐點費用
 [Status]					nvarchar(2)			NOT NULL,						--狀態碼	
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_RestaurantMeal PRIMARY KEY (MealId),
 CONSTRAINT FK1_RestaurantMeal FOREIGN KEY (RestaurantId) REFERENCES EBP.Restaurant(RestaurantId),
 CONSTRAINT FK2_RestaurantMeal FOREIGN KEY (MealImage) REFERENCES BAS.[File](FileId)
) /*餐廳餐點*/

CREATE TABLE EBP.MealCategory( /*餐點可供選擇之類別*/
 MealCgId					int					identity,
 MealId						int					NOT NULL,						--使用者點餐紀錄ID
 RestaurantId				int					NOT NULL,						--所屬餐廳
 CategoryName				nvarchar(30)		NOT NULL,						--類別名稱
 CategoryType				nvarchar(1)			NOT NULL,						--類別 >> N:Radio, Y:Checkbox
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_MealCategory PRIMARY KEY (MealCgId),
 CONSTRAINT FK1_MealCategory FOREIGN KEY (MealId) REFERENCES EBP.RestaurantMeal(MealId)
) /*餐點可供選擇之類別*/

CREATE TABLE EBP.McDetail( /*餐點可供選擇之細節*/
 McDetailId					int					identity,
 MealCgId					int					NOT NULL,						--使用者點餐紀錄單頭ID
 McDetailName				nvarchar(30)		NOT NULL,						--細節名稱
 McDetailPrice				float				NOT NULL,						--細節費用
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_McDetail PRIMARY KEY (McDetailId),
 CONSTRAINT FK1_McDetail FOREIGN KEY (MealCgId) REFERENCES EBP.MealCategory(MealCgId)
) /*餐點可供選擇之細節*/

CREATE TABLE EBP.DailyMeal( /*每日餐點*/
 MealId						int					NOT NULL,						--所屬餐點ID
 CalendarId					int					NOT NULL,						--行事曆ID
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 CreateBy					int					NOT NULL,						--新增人員
 CONSTRAINT PK_DailyMeal PRIMARY KEY (MealId, CalendarId),
 CONSTRAINT FK1_DailyMeal FOREIGN KEY (MealId) REFERENCES EBP.RestaurantMeal(MealId),
 CONSTRAINT FK2_DailyMeal FOREIGN KEY (CalendarId) REFERENCES BAS.Calendar(CalendarId)
) /*每日餐點*/

CREATE TABLE EBP.OrderMealSteward( /*餐點事務人員資料*/
 StewardId					int					identity,
 UserId						int					NOT NULL,						--事務ID
 StaffUserId				int					NOT NULL,						--員工ID
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_OrderMealSteward PRIMARY KEY (StewardId),
 CONSTRAINT FK1_OrderMealSteward FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId),
 CONSTRAINT FK2_OrderMealSteward FOREIGN KEY (StaffUserId) REFERENCES BAS.[User](UserId)
) /*餐點事務人員資料*/

CREATE TABLE EBP.UserMealOrder( /*使用者餐點訂單*/
 UmoId						int					identity,
 CompanyId					int					NOT NULL,						--所屬公司ID
 UserId						int					NOT NULL,						--所屬使用者ID
 UmoDate					datetime			NOT NULL,						--訂餐日期
 UmoNo						nvarchar(20)		NOT NULL,						--訂單編號
 UmoDiscount				float				NULL,							--折扣價格
 UmoAmount 					float				NOT NULL,						--總金額
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_UserMealOrder PRIMARY KEY (UmoId),
 CONSTRAINT AK1_UserMealOrder UNIQUE (UmoNo),
 CONSTRAINT FK1_UserMealOrder FOREIGN KEY (CompanyId) REFERENCES BAS.[Company](CompanyId),
 CONSTRAINT FK2_UserMealOrder FOREIGN KEY (UserId) REFERENCES BAS.[User](UserId)
) /*使用者餐點訂單*/

CREATE TABLE EBP.UmoDetail( /*餐點明細*/
 UmoDetailId				int					identity,
 UmoId						int					NOT NULL,
 MealId						int					NOT NULL,						--餐點ID
 MealName					nvarchar(30)		NOT NULL,						--餐點名稱
 UmoDetailRemark			nvarchar(100)		NULL,		    				--備註
 UmoDetailQty				int					NOT NULL,						--份數
 UmoDetailPrice				float				NOT NULL,						--價格
 Pickup						nvarchar(1)			NOT NULL,						--取餐區 >>BAS.Type UmoDetail.Pickup 1:本廠2.5樓; 2:本廠4樓; 3:2廠
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_UmoDetail PRIMARY KEY (UmoDetailId),
 CONSTRAINT FK1_UmoDetail FOREIGN KEY (UmoId) REFERENCES EBP.UserMealOrder(UmoId),
 CONSTRAINT FK2_UmoDetail FOREIGN KEY (MealId) REFERENCES EBP.RestaurantMeal(MealId),
) /*餐點明細*/

CREATE TABLE EBP.UmoAdditional( /*餐點加購明細*/
 UmoDetailId				int					NOT NULL,
 McDetailId					int					NOT NULL,						--餐點細節ID
 McDetailName				nvarchar(30)		NOT NULL,						--細節名稱
 UmoAdditionalPrice			float				NOT NULL,						--加購價格
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_UmoAdditional PRIMARY KEY (UmoDetailId, McDetailId),
 CONSTRAINT FK1_UmoAdditional FOREIGN KEY (UmoDetailId) REFERENCES EBP.UmoDetail(UmoDetailId),
 CONSTRAINT FK2_UmoAdditional FOREIGN KEY (McDetailId) REFERENCES EBP.McDetail(McDetailId),
) /*餐點加購明細*/

CREATE TABLE EBP.UserMealOrderLog( /*使用者點餐Log*/
 UmoLogId					int					identity,
 UmoId						int					NOT NULL,						--訂單ID
 UserNo						nvarchar(20)		NOT NULL,						--使用者編號
 UmoNo						nvarchar(20)		NOT NULL,						--訂單編號
 MotionDate					datetime			NOT NULL,						--動作日期
 MotionRemark				nvarchar(50)		NOT NULL,						--動作備註
 CreateDate					datetime			NOT NULL	DEFAULT GETDATE(),	--新增日期
 LastModifiedDate			datetime				NULL,						--修改日期
 CreateBy					int					NOT NULL,						--新增人員
 LastModifiedBy				int						NULL,						--修改人員
 CONSTRAINT PK_UserMealOrderLog PRIMARY KEY (UmoLogId)
) /*使用者點餐Log*/

CREATE TABLE EBP.OrderSetting( /*點餐設定檔*/
 StartTime					datetime			NULL,	--開始日期
 EndTime					datetime			NULL,	--結束日期
 OrderType					nvarchar(1)			NULL	--設定類型 >>BAS.Type OrderSetting.OrderType 1:點餐時間; 2:優惠時段
) /*點餐設定檔*/
END /*訂餐系統*/