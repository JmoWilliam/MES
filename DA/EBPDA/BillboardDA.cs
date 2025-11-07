using Dapper;
using Helpers;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Transactions;
using System.Web;

namespace EBPDA
{
    public class BillboardDA
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

        public BillboardDA()
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
        #region //GetBoard -- 取得公佈欄資料 -- Zoey 2023.02.06
        public string GetBoard(int BoardId, int BoardTypeId, string Title, string Status, string StartDate, string EndDate
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.BoardId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.Title, a.Content, FORMAT(a.AnnounceDate, 'yyyy-MM-dd') AnnounceDate, a.BoardTypeId, a.Status
                          , b.TypeName";
                    sqlQuery.mainTables =
                        @"FROM EBP.Board a
                          INNER JOIN EBP.BoardType b ON b.BoardTypeId = a.BoardTypeId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "BoardId", @" AND a.BoardId = @BoardId", BoardId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "BoardTypeId", @" AND a.BoardTypeId = @BoardTypeId", BoardTypeId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Title", @" AND a.Title LIKE '%' + @Title + '%'", Title);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.AnnounceDate >= @StartDate ", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.AnnounceDate <= @EndDate ", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.BoardId";
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

        #region //GetBoardFile -- 取得公佈欄檔案 -- Zoey 2023.02.06
        public string GetBoardFile(int BoardId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.BoardFileId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.BoardId, a.FileId, a.FileType
                          , FORMAT(a.CreateDate, 'yyyy-MM-dd') CreateDate
                          , b.FileName, b.FileContent, b.FileExtension, b.FileSize";
                    sqlQuery.mainTables =
                        @"FROM EBP.BoardFile a
                          INNER JOIN BAS.[File] b ON b.FileId = a.FileId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "BoardId", @" AND a.BoardId = @BoardId", BoardId);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.BoardFileId";
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
        #endregion

        #region //Add
        #region //AddBoard -- 公佈欄資料新增 -- Zoey 2023.02.06
        public string AddBoard(int BoardTypeId, string Title, string AnnounceDate, string Content)
        {
            try
            {
                if (BoardTypeId <= 0) throw new SystemException("【公告類別】不能為空!");
                if (Title.Length <= 0) throw new SystemException("【標題】不能為空!");
                if (AnnounceDate.Length <= 0) throw new SystemException("【公告日期】不能為空!");
                if (Content.Length <= 0) throw new SystemException("【內容】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EBP.Board (Title, Content, AnnounceDate, BoardTypeId, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.BoardId
                                VALUES (@Title, @Content, @AnnounceDate, @BoardTypeId, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Title,
                                Content,
                                AnnounceDate,
                                BoardTypeId,
                                Status = "A",
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

        #region //AddBoardFile -- 公佈欄檔案新增 -- Zoey 2023.02.07
        public string AddBoardFile(int BoardId, int FileId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷公佈欄資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.Board
                                WHERE BoardId = @BoardId";
                        dynamicParameters.Add("BoardId", BoardId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("公佈欄資料錯誤!");
                        #endregion

                        #region //判斷檔案資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[File]
                                WHERE FileId = @FileId";
                        dynamicParameters.Add("FileId", FileId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("檔案資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EBP.BoardFile (BoardId, FileId, FileType, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.BoardFileId
                                VALUES (@BoardId, @FileId, @FileType, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                BoardId,
                                FileId,
                                FileType = "Other",
                                Status = "A",
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
        #region //UpdateBoard -- 公佈欄資料更新-- Zoey 2023.02.06
        public string UpdateBoard(int BoardId, int BoardTypeId, string Title, string AnnounceDate, string Content)
        {
            try
            {
                if (BoardTypeId <= 0) throw new SystemException("【公告類別】不能為空!");
                if (Title.Length <= 0) throw new SystemException("【標題】不能為空!");
                if (AnnounceDate.Length <= 0) throw new SystemException("【公告日期】不能為空!");
                if (Content.Length <= 0) throw new SystemException("【內容】不能為空!");


                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷公佈欄資訊是否正確
                        sql = @"SELECT TOP 1 1
                                FROM EBP.Board
                                WHERE BoardId = @BoardId";
                        dynamicParameters.Add("BoardId", BoardId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("公佈欄資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EBP.Board SET
                                Title = @Title,
                                Content = @Content,
                                AnnounceDate = @AnnounceDate,
                                BoardTypeId = @BoardTypeId,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE BoardId = @BoardId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Title,
                                Content,
                                AnnounceDate,
                                BoardTypeId,
                                LastModifiedDate,
                                LastModifiedBy,
                                BoardId
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

        #region //UpdateBoardStatus -- 公佈欄狀態更新 -- Zoey 2023.02.06
        public string UpdateBoardStatus(int BoardId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷公佈欄資訊是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM EBP.Board
                                WHERE BoardId = @BoardId";
                        dynamicParameters.Add("BoardId", BoardId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("公佈欄資料錯誤!");

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
                        sql = @"UPDATE EBP.Board SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE BoardId = @BoardId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                BoardId
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

        #region //UpdateBoardFileType -- 公佈欄檔案類別更新 -- Zoey 2023.02.07
        public string UpdateBoardFileType(int BoardFileId, string FileType)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷公佈欄檔案是否正確
                        sql = @"SELECT TOP 1 1
                                FROM EBP.BoardFile
                                WHERE BoardFileId = @BoardFileId";
                        dynamicParameters.Add("BoardFileId", BoardFileId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("公佈欄檔案錯誤!");
                        #endregion

                        #region //判斷檔案類型資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Type]
                                WHERE TypeNo = @TypeNo
                                AND TypeSchema = @TypeSchema";
                        dynamicParameters.Add("TypeNo", FileType);
                        dynamicParameters.Add("TypeSchema", "BoardFile.FileType");

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("檔案類型資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EBP.BoardFile SET
                                FileType = @FileType,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE BoardFileId = @BoardFileId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                FileType,
                                LastModifiedDate,
                                LastModifiedBy,
                                BoardFileId
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
        #endregion

        #region //Delete
        #region //DeleteBoard -- 公佈欄資料刪除 -- Zoey 2023.02.06
        public string DeleteBoard(int BoardId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷公佈欄資訊是否正確
                        sql = @"SELECT TOP 1 1
                                FROM EBP.Board
                                WHERE BoardId = @BoardId";
                        dynamicParameters.Add("BoardId", BoardId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("公佈欄資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除關聯table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.BoardFile
                                WHERE BoardId = @BoardId";
                        dynamicParameters.Add("BoardId", BoardId);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.Board
                                WHERE BoardId = @BoardId";
                        dynamicParameters.Add("BoardId", BoardId);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteBoardFile -- 公佈欄檔案刪除 -- Zoey 2023.02.07
        public string DeleteBoardFile(int BoardFileId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷公佈欄檔案是否正確
                        sql = @"SELECT TOP 1 1
                                FROM EBP.BoardFile
                                WHERE BoardFileId = @BoardFileId";
                        dynamicParameters.Add("BoardFileId", BoardFileId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("公佈欄檔案錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.BoardFile
                                WHERE BoardFileId = @BoardFileId";
                        dynamicParameters.Add("BoardFileId", BoardFileId);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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
