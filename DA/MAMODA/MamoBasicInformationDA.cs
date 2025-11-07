
using Dapper;
using Helpers;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Transactions;
using System.Web;

namespace MAMODA
{
    public class MamoBasicInformationDA
    {
        public string MainConnectionStrings = "";
        public string ErpConnectionStrings = "";
        public string ErpSysConnectionStrings = "";
        public string HrmConnectionStrings = "";

        public int CurrentCompany = -1;
        public int CurrentUser = -1;
        public int CreateBy = -1;
        public int LastModifiedBy = -1;
        public DateTime CreateDate = default(DateTime);
        public DateTime LastModifiedDate = default(DateTime);

        public string sql = "";
        public JObject jsonResponse = new JObject();
        public Logger logger = LogManager.GetCurrentClassLogger();
        public DynamicParameters dynamicParameters = new DynamicParameters();
        public SqlQuery sqlQuery = new SqlQuery();
        public MamoHelper mamoHelper = new MamoHelper();

        public MamoBasicInformationDA()
        {
            MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];

            GetUserInfo();
            CreateDate = DateTime.Now;
            LastModifiedDate = DateTime.Now;

            dynamicParameters = new DynamicParameters();
            sqlQuery = new SqlQuery();
        }

        #region //GetUserInfo 取得使用者資訊
        private void GetUserInfo()
        {
            try
            {
                CurrentCompany = Convert.ToInt32(HttpContext.Current.Session["UserCompany"]);
                CurrentUser = Convert.ToInt32(HttpContext.Current.Session["UserId"]);
                CreateBy = Convert.ToInt32(HttpContext.Current.Session["UserId"]);
                LastModifiedBy = Convert.ToInt32(HttpContext.Current.Session["UserId"]);

                if (HttpContext.Current.Session["CompanySwitch"] != null)
                {
                    if (HttpContext.Current.Session["CompanySwitch"].ToString() == "manual")
                    {
                        CurrentCompany = Convert.ToInt32(HttpContext.Current.Session["CompanyId"]);
                    }
                }
            }
            catch (Exception e)
            {
                logger.Error(e.Message);
            }
        }
        #endregion

        #region //Get

        #region //GetMamoTeams              -- 取得團隊資料         -- Tim 2024.10.18
        public string GetMamoTeams(
            string  TeamId,
            string  MamoTeamId,
            string  TeamName,
            string  Remark,
            string  Status,
            string  OrderBy,
            int     PageIndex,
            int     PageSize
        )
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //Request
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.TeamId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.CompanyId,
	                        a.MamoTeamId, a.TeamName, a.TeamNo, a.[Status], a.Remark";

                    sqlQuery.mainTables =
                        @"FROM MAMO.Teams a";
                    sqlQuery.auxTables = "";

                    string queryCondition = " AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TeamId", @" AND a.TeamId = @TeamId", TeamId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MamoTeamId", @" AND a.MamoTeamId LIKE '%' + @MamoTeamId + '%'", MamoTeamId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TeamName", @" AND a.TeamName LIKE '%' + @TeamName + '%'", TeamName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Remark", @" AND a.Remark LIKE '%' + @Remark + '%'", Remark);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.TeamId DESC";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);
                    #endregion //Request

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = result
                    });

                    #endregion //Response
                }
            }
            catch(Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion //Response
            }

            return jsonResponse.ToString();
        }

        #endregion //GetMamoTeams           取得團隊資料

        #region //GetMamoTeamMembers        -- 取得團隊成員資料     -- Tim 2024.10.18
        public string GetMamoTeamMembers(
            int MemberId,
            int TeamId,
            int CompanyId,
            string Departments,
            string UserNo,
            string UserName,
            string OrderBy,
            int PageIndex,
            int PageSize
        )
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.MemberId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.TeamId, a.UserId, a.[Status]
                        , b.UserNo, b.UserName, b.Gender
                        , c.DepartmentId, c.DepartmentNo, c.DepartmentName
                        , d.CompanyId, d.CompanyNo, d.CompanyName, ISNULL(d.LogoIcon, -1) LogoIcon";
                    sqlQuery.mainTables =
                        @"FROM MAMO.TeamMembers a
                        INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                        INNER JOIN BAS.Department c ON b.DepartmentId = c.DepartmentId
                        INNER JOIN BAS.Company d ON c.CompanyId = d.CompanyId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MemberId", @" AND a.MemberId = @MemberId", MemberId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TeamId", @" AND a.TeamId = @TeamId", TeamId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CompanyId", @" AND d.CompanyId = @CompanyId", CompanyId);
                    if (Departments.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Departments", @" AND c.DepartmentId IN @Departments", Departments.Split(','));
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserNo", @" AND b.UserNo LIKE '%' + @UserNo + '%'", UserNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserName", @" AND b.UserName LIKE '%' + @UserName + '%'", UserName);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.TeamId, d.CompanyId, b.UserNo";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = result
                    });
                    #endregion
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }

        #endregion //GetMamoTeamMembers        取得團隊成員資料

        #region //GetUserDetail             -- 取得使用者資訊       -- Tim 2024.10.18
        public string GetUserDetail(
            int  UserId
        )
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //Request
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.UserId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.UserId, a.UserNo, a.UserName, a.Email,
	                        a.UserStatus, a.SystemStatus, a.PasswordStatus, a.Status";

                    sqlQuery.mainTables =
                        @"FROM BAS.[User] a";
                    sqlQuery.auxTables = "";

                    string queryCondition = " AND a.UserId = @UserId";
                    dynamicParameters.Add("UserId", UserId);

                    sqlQuery.conditions = queryCondition;

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);
                    #endregion //Request

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = result
                    });

                    #endregion //Response
                }
            }
            catch(Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion //Response
            }

            return jsonResponse.ToString();
        }

        #endregion //GetUserDetail

        #region //GetCompanyDetail          -- 取得使用者公司資訊   -- Tim 2024.10.18
        public string GetCompanyDetail(
            int  CompanyId
        )
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //Request
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.CompanyId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.CompanyId, a.CompanyNo, a.CompanyName, a.CompanyName, a.LogoIcon";

                    sqlQuery.mainTables =
                        @"FROM BAS.Company a";
                    sqlQuery.auxTables = "";

                    string queryCondition = " AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CompanyId);

                    sqlQuery.conditions = queryCondition;

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);
                    #endregion //Request

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = result
                    });

                    #endregion //Response
                }
            }
            catch(Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion //Response
            }

            return jsonResponse.ToString();
        }

        #endregion //GetCompanyDetail

        #region //GetMamoChannelMembers     -- 取得頻道成員資料     -- Tim 2024.10.18
        #endregion //GetMamoChannelMembers  取得頻道成員資料

        #region //GetMamoChannels           -- 取得頻道資料         -- Tim 2024.10.18
        public string GetMamoChannels(
            string  ChannelId,
            string  MamoChannelId,
            string  TeamId,
            string  ChannelName,
            string  ChannelNo,
            string  Status,
            string  OrderBy,
            int     PageIndex,
            int     PageSize
        )
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //Request
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.ChannelId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.MamoChannelId,
                        	b.TeamId, b.TeamName,
                        	a.ChannelName, a.ChannelNo, a.Status";

                    sqlQuery.mainTables =
                        @"FROM MAMO.Channels a
                        INNER JOIN MAMO.Teams b ON a.TeamId = b.TeamId";
                    sqlQuery.auxTables = "";

                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ChannelId", @" AND a.ChannelId = @ChannelId", ChannelId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MamoChannelId", @" AND a.MamoChannelId LIKE '%' + @MamoChannelId + '%'", MamoChannelId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TeamId", @" AND a.TeamId = @TeamId", TeamId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ChannelName", @" AND a.ChannelName LIKE '%' + @ChannelName + '%'", ChannelName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ChannelNo", @" AND a.ChannelNo LIKE '%' + @ChannelNo + '%'", ChannelNo);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "b.TeamId DESC, a.ChannelId";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);
                    #endregion //Request

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = result
                    });

                    #endregion //Response
                }
            }
            catch(Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion //Response
            }

            return jsonResponse.ToString();
        }

        #endregion //GetMamoChannels           取得團隊資料



        #endregion //Get

        #region //Add

        #region //AddMamoTeams              -- 建立團隊             -- Tim 2024.10.18
        public string AddMamoTeams(
            string CompanyNo,
            int UserId,
            string TeamName,
            string Remark
        )
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    var rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int teamId = -1;
                        #region //判斷專案資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT
                                    *
                                FROM MAMO.Teams a
                                WHERE a.TeamName = @TeamName";

                        dynamicParameters.Add("TeamName", TeamName);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【MAMO】團隊名稱重覆!");

                        #endregion

                        if (teamId <= 0)
                        {
                            #region //MAMO團隊新增
                            var mamoResult = mamoHelper.CreateTeams(CompanyNo, CurrentUser, TeamName, Remark);
                            JObject mamoResultJson = JObject.Parse(mamoResult);

                            if (mamoResultJson["status"].ToString() == "success")
                            {
                                foreach (var data in mamoResultJson["result"])
                                {
                                    teamId = Convert.ToInt32(data["TeamId"].ToString());
                                    rowsAffected += 1;
                                }
                            }

                            #endregion
                        }

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion
                    }

                    transactionScope.Complete();
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion //AddMamoTeams -- 建立MAMO團隊

        #region //AddMamoTeamMembers        -- 新增團隊成員         -- Tim 2024.10.18
        public string AddMamoTeamMembers(
            string CompanyNo,
            int UserId,
            int TeamId,
            string Users
        )
        {
            try
            {
                if (Users.Length <= 0) throw new SystemException("【團隊成員】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷團隊資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MAMO.Teams
                                WHERE TeamId = @TeamId";
                        dynamicParameters.Add("TeamId", TeamId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("團隊資料錯誤!");
                        #endregion

                        #region //判斷團隊成員資料是否正確
                        string[] usersList = Users.Split(',');

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT COUNT(1) TotalUsers
                                FROM BAS.[User]
                                WHERE UserId IN @UserId";
                        dynamicParameters.Add("UserId", usersList);

                        int totalUsers = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).TotalUsers;
                        if (totalUsers != usersList.Length) throw new SystemException("團隊成員資料錯誤!");

                        #endregion

                        #region //UserListToApi
                        sql = @"SELECT
                                    *
                                FROM BAS.[User]
                                WHERE UserId IN (" + Users + ")";
                        // dynamicParameters.Add("UserId", usersList);

                        var resultUsers = sqlConnection.Query(sql, dynamicParameters);

                        List<string> UserNo = new List<string>();

                        foreach (var item in resultUsers)
                        {
                            UserNo.Add(item.UserNo);
                        }
                        #endregion //UserListToApi



                        int rowsAffected = 0;

                        var mamoResult = mamoHelper.AddTeamMembers(CompanyNo, UserId, TeamId, UserNo);
                        JObject mamoResultJson = JObject.Parse(mamoResult);

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion
                    }

                    transactionScope.Complete();
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //AddMamoChannels           -- 建立頻道             -- Tim 2024.10.21
        public string AddMamoChannels(
            string CompanyNo,
            int UserId,
            string ChannelName,
            int TeamId,
            string Remark
        )
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    var rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int teamId = -1;
                        #region //判斷專案資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT
                                    *
                                FROM MAMO.Channels
                                WHERE TeamId = @TeamId
                                AND ChannelName = @ChannelName";
                        dynamicParameters.Add("TeamId", TeamId);
                        dynamicParameters.Add("ChannelName", ChannelName);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【MAMO】頻道名稱重覆!");

                        #endregion

                        if (teamId <= 0)
                        {
                            #region //MAMO頻道新增
                            var mamoResult = mamoHelper.CreateChannels(CompanyNo, CurrentUser, TeamId, ChannelName, Remark);
                            JObject mamoResultJson = JObject.Parse(mamoResult);

                            if (mamoResultJson["status"].ToString() == "success")
                            {
                                foreach (var data in mamoResultJson["result"])
                                {
                                    teamId = Convert.ToInt32(data["TeamId"].ToString());
                                    rowsAffected += 1;
                                }
                            }

                            #endregion
                        }

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion
                    }

                    transactionScope.Complete();
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion //AddMamoChannels -- 建立MAMO頻道

        #endregion //Add

        #region //Update

        #region //UpdateMamoTeamsStatus     -- 團隊狀態更新         -- Tim 2024.10.18
        public string UpdateMamoTeamsStatus(
            string CompanyNo,
            int UserId,
            int TeamId
        )
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷團隊資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1
                                    Status
                                FROM MAMO.Teams
                                WHERE TeamId = @TeamId";
                        dynamicParameters.Add("TeamId", TeamId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("團隊資料錯誤!");

                        string status = "";
                        foreach (var item in result)
                        {
                            status = item.Status;
                        }

                        #region //調整為相反狀態
                        switch (status)
                        {
                            case "A":
                                status = "S";
                                break;
                            case "S":
                                status = "A";
                                break;
                        }
                        #endregion
                        #endregion


                        if (String.Equals(status, "A"))
                        {
                            mamoHelper.RestoreTeams(CompanyNo, CurrentUser, TeamId);
                        }
                        else if (String.Equals(status, "S"))
                        {
                            mamoHelper.DeleteTeams(CompanyNo, CurrentUser, TeamId);
                        }

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MAMO.Teams
                                SET
                                    Status = @Status,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                WHERE TeamId = @TeamId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                TeamId
                            });

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion
                    }

                    transactionScope.Complete();

                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #endregion //Update

        #region //Delete

        #region //DeleteMamoTeamMembers     -- 移除團隊成員         -- Tim 2024.10.18
        public string DeleteMamoTeamMembers(
            string CompanyNo,
            int UserId,
            int MemberId
        )
        {
            try
            {
                int TeamId = -1;
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷子系統使用者資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MAMO.TeamMembers
                                WHERE MemberId = @MemberId";
                        dynamicParameters.Add("MemberId", MemberId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("子系統使用者資料錯誤!");
                        #endregion

                        #region //UserListToApi
                        sql = @"SELECT
                                    a.MemberId, a.TeamId, a.UserId,
                                    b.UserNo, b.UserName,
                                    a.Status
                                FROM MAMO.TeamMembers a
                                INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                                WHERE a.MemberId IN (" + MemberId + ")";
                        // dynamicParameters.Add("UserId", usersList);

                        var resultUsers = sqlConnection.Query(sql, dynamicParameters);

                        List<string> UserNo = new List<string>();

                        foreach (var item in resultUsers)
                        {
                            TeamId = item.TeamId;
                            UserNo.Add(item.UserNo);
                        }
                        #endregion //UserListToApi



                        int rowsAffected = 0;

                        var mamoResult = mamoHelper.DeleteTeamMembers(CompanyNo, UserId, TeamId, UserNo);
                        JObject mamoResultJson = JObject.Parse(mamoResult);

                        // int rowsAffected = 0;
                        #region //刪除主要table
                        // dynamicParameters = new DynamicParameters();
                        // sql = @"DELETE MAMO.TeamMembers
                        //         WHERE MemberId = @MemberId";
                        // dynamicParameters.Add("MemberId", MemberId);

                        // rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion
                    }

                    transactionScope.Complete();
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #endregion //Delete

    }
}
