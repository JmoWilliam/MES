using Dapper;
using Helpers;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Transactions;
using System.Web;

namespace EBPDA
{
    public class GamingDA
    {
        public string MainConnectionStrings = "";

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

        public GamingDA()
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
        #region //GetBingo -- 取得賓果資料 -- Ben Ma 2023.01.04
        public string GetBingo(int BingoId, string BingoNo, string Status, string WinStatus
           , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.BingoId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.BingoNo, a.[Status], a.WinStatus
                        , ISNULL(b.RoundName, '') RoundName, ISNULL(FORMAT(b.RoundDate, 'yyyy-MM-dd'), '') RoundDate";
                    sqlQuery.mainTables =
                        @"FROM EBP.Bingo a
                        LEFT JOIN EBP.BingoRound b ON a.RoundId = b.RoundId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "BingoId", @" AND a.BingoId = @BingoId", BingoId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "BingoNo", @" AND a.BingoNo LIKE '%' + @BingoNo + '%'", BingoNo);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    if (WinStatus.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "WinStatus", @" AND a.WinStatus IN @Status", WinStatus.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.BingoId";
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
        #endregion

        #region //GetBingoMap -- 取得賓果分佈圖資料 -- Ben Ma 2023.01.05
        public string GetBingoMap(int BingoId, string BingoNo)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.MapId, a.MapContent, a.MapIndex
                            FROM EBP.BingoMap a
                            INNER JOIN EBP.Bingo b ON a.BingoId = b.BingoId
                            WHERE 1=1";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "BingoId", @" AND b.BingoId = @BingoId", BingoId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "BingoNo", @" AND b.BingoNo = @BingoNo", BingoNo);
                    sql += @" ORDER BY a.MapIndex";

                    var result = sqlConnection.Query(sql, dynamicParameters);

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
        #endregion

        #region //GetBingoRound -- 取得賓果開局資料 -- Ben Ma 2023.01.06
        public string GetBingoRound(int RoundId, string RoundName, string StartDate, string EndDate, string Status
           , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.RoundId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.RoundName, FORMAT(a.RoundDate, 'yyyy-MM-dd') RoundDate, a.[Status]";
                    sqlQuery.mainTables =
                        @"FROM EBP.BingoRound a";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RoundId", @" AND a.RoundId = @RoundId", RoundId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RoundName", @" AND a.RoundName LIKE '%' + @RoundName + '%'", RoundName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.RoundDate >= @StartDate ", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.RoundDate <= @EndDate ", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.[Status], a.RoundDate DESC, a.RoundId";
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
        #endregion

        #region //GetBingoLotteryRecord -- 取得賓果開獎記錄資料 -- Ben Ma 2023.01.06
        public string GetBingoLotteryRecord(int RoundId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.RecordId, a.RecordContent, a.RecordSeq, FORMAT(a.CreateDate, 'yyyy-MM-dd HH:mm:ss') CreateDate
                            , b.UserNo, b.UserName, b.Gender
                            , c.DepartmentNo, c.DepartmentName
                            , d.LogoIcon
                            FROM EBP.BingoLotteryRecord a
                            INNER JOIN BAS.[User] b ON a.CreateBy = b.UserId
                            INNER JOIN BAS.Department c ON b.DepartmentId = c.DepartmentId
                            INNER JOIN BAS.Company d ON c.CompanyId = d.CompanyId
                            WHERE 1=1
                            AND RoundId = @RoundId
                            ORDER BY a.RecordSeq";
                    dynamicParameters.Add("RoundId", RoundId);

                    var result = sqlConnection.Query(sql, dynamicParameters);

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
        #endregion
        #endregion

        #region //Add
        #region //AddBingo -- 賓果資料新增 -- Ben Ma 2023.01.04
        public string AddBingo(string BingoMap)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;

                        #region //賓果新增
                        #region //判斷賓果代號是否重複
                        string bingoNo = "";
                        bool isExist = true;
                        while (isExist)
                        {
                            bingoNo = BaseHelper.RandomCode(10);

                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM EBP.Bingo
                                    WHERE BingoNo = @BingoNo";
                            dynamicParameters.Add("BingoNo", bingoNo);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            isExist = result.Count() > 0;
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EBP.Bingo (BingoNo, Status, WinStatus
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.BingoId
                                VALUES (@BingoNo, @Status, @WinStatus
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                BingoNo = bingoNo,
                                Status = "A",
                                WinStatus = "N",
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var bingoResult = sqlConnection.Query(sql, dynamicParameters);
                        rowsAffected += bingoResult.Count();

                        int BingoId = 0;
                        foreach (var item in bingoResult)
                        {
                            BingoId = Convert.ToInt32(item.BingoId);
                        }
                        #endregion

                        #region //賓果分佈圖新增
                        string[] mapContent = BingoMap.Split(';');

                        for (int i = 0; i < 12; i++)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO EBP.BingoMap (BingoId, MapContent, MapIndex
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.MapId
                                    VALUES (@BingoId, @MapContent, @MapIndex
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    BingoId,
                                    MapContent = mapContent[i],
                                    MapIndex = i + 1,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            var result = sqlConnection.Query(sql, dynamicParameters);
                            rowsAffected += result.Count();
                        }

                        for (int i = 12; i < 24; i++)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO EBP.BingoMap (BingoId, MapContent, MapIndex
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.MapId
                                    VALUES (@BingoId, @MapContent, @MapIndex
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    BingoId,
                                    MapContent = mapContent[i],
                                    MapIndex = i + 2,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            var result = sqlConnection.Query(sql, dynamicParameters);
                            rowsAffected += result.Count();
                        }
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

        #region //AddBingoRound -- 賓果開局資料新增 -- Ben Ma 2023.01.06
        public string AddBingoRound(string RoundName, string RoundDate)
        {
            try
            {
                if (RoundName.Length <= 0) throw new SystemException("【開局名稱】不能為空!");
                if (RoundName.Length > 100) throw new SystemException("【開局名稱】長度錯誤!");
                if (RoundDate.Length <= 0) throw new SystemException("【開局時間】不能為空!");
                if (!DateTime.TryParse(RoundDate, out DateTime tempDate)) throw new SystemException("【開局時間】格式錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EBP.BingoRound (RoundName, RoundDate, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.RoundId
                                VALUES (@RoundName, @RoundDate, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RoundName,
                                RoundDate,
                                Status = "S", //停用
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = insertResult.Count();

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

        #region //AddBingoLotteryRecord -- 賓果開獎紀錄新增 -- Ben Ma 2023.01.06
        public string AddBingoLotteryRecord(int RoundId, string RecordContent)
        {
            try
            {
                if (RecordContent.Length <= 0) throw new SystemException("【開獎內容】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷賓果開局資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM EBP.BingoRound a
                                WHERE a.RoundId = @RoundId";
                        dynamicParameters.Add("RoundId", RoundId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【賓果開局】資料錯誤!");
                        #endregion

                        #region //搜尋目前賓果開獎紀錄順序
                        sql = @"SELECT MAX(a.RecordSeq) CurrentSeq
                                FROM EBP.BingoLotteryRecord a
                                WHERE a.RoundId = @RoundId";
                        dynamicParameters.Add("RoundId", RoundId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);

                        int currentSeq = 0;
                        foreach (var item in result2)
                        {
                            currentSeq = Convert.ToInt32(item.CurrentSeq);
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EBP.BingoLotteryRecord (RoundId, RecordContent, RecordSeq
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.RecordId
                                VALUES (@RoundId, @RecordContent, @RecordSeq
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RoundId,
                                RecordContent,
                                RecordSeq = currentSeq + 1,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = insertResult.Count();

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
        #endregion

        #region //Update
        #region //UpdateBingoStatus -- 賓果狀態更新 -- Ben Ma 2023.01.07
        public string UpdateBingoStatus(int BingoId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷賓果資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM EBP.Bingo
                                WHERE BingoId = @BingoId";
                        dynamicParameters.Add("BingoId", BingoId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("賓果資料錯誤!");

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

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EBP.Bingo SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE BingoId = @BingoId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                BingoId
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

        #region //UpdateBingoWinStatus -- 賓果中獎狀態更新 -- Ben Ma 2023.01.07
        public string UpdateBingoWinStatus(int RoundId, string BingoNo)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷賓果開局資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.BingoRound
                                WHERE RoundId = @RoundId
                                AND Status = @Status";
                        dynamicParameters.Add("RoundId", RoundId);
                        dynamicParameters.Add("Status", "A");

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("賓果開局資料錯誤!");
                        #endregion

                        #region //判斷賓果資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.Bingo
                                WHERE BingoNo = @BingoNo";
                        dynamicParameters.Add("BingoNo", BingoNo);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("賓果資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EBP.Bingo SET
                                RoundId = @RoundId,
                                WinStatus = @WinStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE BingoNo = @BingoNo";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RoundId,
                                WinStatus = "Y",
                                LastModifiedDate,
                                LastModifiedBy,
                                BingoNo
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

        #region //UpdateBingoRound -- 賓果開局資料更新 -- Ben Ma 2023.01.06
        public string UpdateBingoRound(int RoundId, string RoundName, string RoundDate)
        {
            try
            {
                if (RoundName.Length <= 0) throw new SystemException("【開局名稱】不能為空!");
                if (RoundName.Length > 100) throw new SystemException("【開局名稱】長度錯誤!");
                if (RoundDate.Length <= 0) throw new SystemException("【開局時間】不能為空!");
                if (!DateTime.TryParse(RoundDate, out DateTime tempDate)) throw new SystemException("【開局時間】格式錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷賓果開局資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.BingoRound
                                WHERE RoundId = @RoundId";
                        dynamicParameters.Add("RoundId", RoundId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("賓果開局資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EBP.BingoRound SET
                                RoundName = @RoundName,
                                RoundDate = @RoundDate,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RoundId = @RoundId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RoundName,
                                RoundDate,
                                LastModifiedDate,
                                LastModifiedBy,
                                RoundId
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

        #region //UpdateBingoRoundStatus -- 賓果開局狀態更新 -- Ben Ma 2023.01.07
        public string UpdateBingoRoundStatus(int RoundId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷賓果開局資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM EBP.BingoRound
                                WHERE RoundId = @RoundId";
                        dynamicParameters.Add("RoundId", RoundId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("賓果開局資料錯誤!");

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

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EBP.BingoRound SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RoundId = @RoundId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                RoundId
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

        #region //UpdateBingoRoundReset -- 賓果開局重置 -- Ben Ma 2023.01.07
        public string UpdateBingoRoundReset(int RoundId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷賓果開局資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.BingoRound
                                WHERE RoundId = @RoundId";
                        dynamicParameters.Add("RoundId", RoundId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("賓果開局資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //賓果開獎記錄初始化
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.BingoLotteryRecord
                                WHERE RoundId = @RoundId";
                        dynamicParameters.Add("RoundId", RoundId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //賓果中獎初始化
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EBP.Bingo SET
                                RoundId = null,
                                WinStatus = @WinStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RoundId = @RoundId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                WinStatus = "N",
                                LastModifiedDate,
                                LastModifiedBy,
                                RoundId
                            });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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
        #endregion

        #region //Delete
        #region //DeleteBingo -- 賓果資料刪除 -- Ben Ma 2023.01.07
        public string DeleteBingo(int BingoId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷賓果資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 ISNULL(RoundId, -1) RoundId
                                FROM EBP.Bingo
                                WHERE BingoId = @BingoId";
                        dynamicParameters.Add("BingoId", BingoId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("賓果資料錯誤!");

                        #region //判斷是否有中獎資料
                        foreach (var item in result)
                        {
                            if (Convert.ToInt32(item.RoundId) > 0) throw new SystemException("該賓果已有中獎紀錄，無法刪除!");
                        }
                        #endregion
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除關聯table
                        #region //刪除賓果分佈圖
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.BingoMap
                                WHERE BingoId = @BingoId";
                        dynamicParameters.Add("BingoId", BingoId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.Bingo
                                WHERE BingoId = @BingoId";
                        dynamicParameters.Add("BingoId", BingoId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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
        #endregion
    }
}
