IF(SELECT COUNT(NAME) FROM sysobjects WHERE TYPE = 'U' AND NAME = 'RoleUser') > 0 DROP TABLE WX.AccessToken

--******************************************************************************************************************
CREATE TABLE WX.AccessToken( /*存取令牌*/
 Plaintext					nvarchar(512)		NOT NULL,						--內容
 Expires					datetime			NOT NULL,						--有效時間
) /*存取令牌*/