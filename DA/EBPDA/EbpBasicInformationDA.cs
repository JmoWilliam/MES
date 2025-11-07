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
    public class EbpBasicInformationDA
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

        public EbpBasicInformationDA()
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
        #region //GetBoardType -- 取得公告類別資料 -- Zoey 2023.02.02
        public string GetBoardType(int BoardTypeId, string TypeName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.BoardTypeId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.TypeName, a.Status";
                    sqlQuery.mainTables =
                        @"FROM EBP.BoardType a";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "BoardTypeId", @" AND a.BoardTypeId = @BoardTypeId", BoardTypeId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TypeName", @" AND a.TypeName LIKE '%' + @TypeName + '%'", TypeName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.BoardTypeId";
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

        #region //GetAnnual -- 取得年份資料 -- Zoey 2023.02.02
        public string GetAnnual(int AnnualId, int Annual, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.AnnualId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.Annual, a.Status";
                    sqlQuery.mainTables =
                        @"FROM EBP.Annual a";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "AnnualId", @" AND a.AnnualId = @AnnualId", AnnualId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Annual", @" AND a.Annual = @Annual", Annual);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.AnnualId";
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

        #region //GetClubJob -- 取得社團職位資料 -- Zoey 2023.02.06
        public string GetClubJob(int ClubJobId, string JobName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.ClubJobId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.JobName, a.Status";
                    sqlQuery.mainTables =
                        @"FROM EBP.ClubJob a";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ClubJobId", @" AND a.ClubJobId = @ClubJobId", ClubJobId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "JobName", @" AND a.JobName LIKE '%' + @JobName + '%'", JobName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ClubJobId";
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

        #region //GetRestaurant -- 取得餐廳資料 -- Daiyi 2022.12.21
        public string GetRestaurant(int RestaurantId, string RestaurantNo, string RestaurantName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.RestaurantId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.RestaurantNo, a.RestaurantName, a.GuiNumber, a.ResponsiblePerson, a.Contact, a.TelNo, a.FaxNo, a.Email
                          , a.Address, a.AccountDay, a.PaymentType, a.InvoiceCount, a.RemitBank, a.RemitAccount
                          , a.RestaurantRemark, a.StartTime, a.EndTime, a.Status";
                    sqlQuery.mainTables =
                        @"FROM EBP.Restaurant a";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RestaurantId", @" AND a.RestaurantId = @RestaurantId", RestaurantId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RestaurantNo", @" AND a.RestaurantNo LIKE '%' + @RestaurantNo + '%'", RestaurantNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RestaurantName", @" AND a.RestaurantName LIKE '%' + @RestaurantName + '%'", RestaurantName);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.RestaurantId DESC";
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

        #region//GetRestaurantFile -- 取得餐廳上傳檔案資料 -- 2022.12.23
        public string GetRestaurantFile(int RestaurantId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.FileId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",a.RestaurantId, b.FileId, b.[FileName], FORMAT(a.CreateDate, 'yyyy-MM-dd HH:mm') AS CreateDate";
                    sqlQuery.mainTables =
                        @" FROM EBP.RestaurantFile a
                        INNER JOIN BAS.[File] b ON a.FileDoc = b.FileId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "AND a.RestaurantId = @RestaurantId";
                    dynamicParameters.Add("RestaurantId", RestaurantId);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.distinct = false;
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

        #region //GetStewardInfo -- 取得事務人員資料 -- Yi 2023.05.11
        public string GetStewardInfo(int StewardId, int UserId, string UserName
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.UserId";
                    sqlQuery.distinct = true;
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.UserNo, a.UserName";
                    sqlQuery.mainTables =
                        @"FROM BAS.[User] a
                        INNER JOIN BAS.Department b ON b.DepartmentId = a.DepartmentId
                        INNER JOIN BAS.Company c ON c.CompanyId = b.CompanyId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND c.CompanyId = @CompanyId
                                            AND a.UserId IN (
                                                SELECT DISTINCT x.UserId
                                                FROM EBP.OrderMealSteward x
                                            )";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    dynamicParameters.Add("UserId", UserId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserId", @" AND a.UserId = @UserId", UserId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserName", @" AND b.UserName LIKE '%' + @UserName + '%'", UserName);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.UserId";
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

        #region //GetStaffInfo -- 取得每日餐點資料 -- Yi 2023.05.12
        public string GetStaffInfo(int UserId, int StaffUserId, string UserName)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.UserId, a.StaffUserId, b.UserName
                            FROM EBP.OrderMealSteward a
                            INNER JOIN BAS.[User] b ON b.UserId = a.StaffUserId
                            WHERE a.UserId IN (
                                SELECT DISTINCT x.UserId
                                FROM BAS.[User] x
                            )";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "UserId", @" AND a.UserId = @UserId", UserId);
                    sql += @"
                            ORDER BY a.StaffUserId";

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
        #region //AddBoardType -- 公告類別資料新增 -- Zoey 2023.02.02
        public string AddBoardType(string TypeName)
        {
            try
            {
                if (TypeName.Length <= 0) throw new SystemException("【類別名稱】不能為空!");
                if (TypeName.Length > 100) throw new SystemException("【類別名稱】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷類別名稱是否重複
                        sql = @"SELECT TOP 1 1
                                FROM EBP.BoardType
                                WHERE TypeName = @TypeName";
                        dynamicParameters.Add("TypeName", TypeName);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【類別名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EBP.BoardType (TypeName, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.BoardTypeId
                                VALUES (@TypeName, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TypeName,
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

        #region //AddAnnual -- 年份資料新增 -- Zoey 2023.02.02
        public string AddAnnual(int Annual)
        {
            try
            {
                if (Annual <= 0) throw new SystemException("【年份】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷年份是否重複
                        sql = @"SELECT TOP 1 1
                                FROM EBP.Annual
                                WHERE Annual = @Annual";
                        dynamicParameters.Add("Annual", Annual);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【年份】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EBP.Annual (Annual, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.AnnualId
                                VALUES (@Annual, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Annual,
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

        #region //AddClubJob -- 社團職位資料新增 -- Zoey 2023.02.06
        public string AddClubJob(string JobName)
        {
            try
            {
                if (JobName.Length <= 0) throw new SystemException("【職位名稱】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷職位是否重複
                        sql = @"SELECT TOP 1 1
                                FROM EBP.ClubJob
                                WHERE JobName = @JobName";
                        dynamicParameters.Add("JobName", JobName);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【年份】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EBP.ClubJob (JobName, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ClubJobId
                                VALUES (@JobName, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                JobName,
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

        #region //AddRestaurant -- 餐廳資料新增 -- Daiyi 2022.12.21
        public string AddRestaurant(string RestaurantNo, string RestaurantName
            , string GuiNumber, string ResponsiblePerson, string Contact
            , string TelNo, string FaxNo, string Email, string Address
            , string AccountDay, string PaymentType, string InvoiceCount
            , string RemitBank, string RemitAccount, string RestaurantRemark
            , string StartTime, string EndTime)
        {
            try
            {
                #region //判斷餐廳資料長度
                if (RestaurantNo.Length <= 0) throw new SystemException("【餐廳代號】不能為空!");
                if (RestaurantName.Length <= 0) throw new SystemException("【餐廳名稱】不能為空!");
                //if (GuiNumber.Length <= 0) throw new SystemException("【統一編號】不能為空!");
                //if (ResponsiblePerson.Length <= 0) throw new SystemException("【負責人】不能為空!");
                if (Contact.Length <= 0) throw new SystemException("【聯絡人】不能為空!");
                if (TelNo.Length <= 0) throw new SystemException("【電話】不能為空!");
                //if (FaxNo.Length <= 0) throw new SystemException("【傳真】不能為空!");
                //if (Email.Length <= 0) throw new SystemException("【郵件】不能為空!");
                if (Address.Length <= 0) throw new SystemException("【地址】不能為空!");
                if (PaymentType.Length <= 0) throw new SystemException("【付款方式】不能為空!");
                //if (InvoiceCount.Length <= 0) throw new SystemException("【發票聯數】不能為空!");
                if (StartTime.Length <= 0) throw new SystemException("【開始時間】不能為空!");
                if (EndTime.Length <= 0) throw new SystemException("【結束時間】不能為空!");

                if (RestaurantNo.Length > 50) throw new SystemException("【餐廳代號】長度錯誤!");
                if (RestaurantName.Length > 200) throw new SystemException("【餐廳名稱】長度錯誤!");
                //if (ResponsiblePerson.Length > 30) throw new SystemException("【負責人】長度錯誤!");
                if (Contact.Length > 30) throw new SystemException("【聯絡人】長度錯誤!");
                if (TelNo.Length > 20) throw new SystemException("【電話】長度錯誤!");
                //if (FaxNo.Length > 20) throw new SystemException("【傳真】長度錯誤!");
                //if (Email.Length > 60) throw new SystemException("【郵件】長度錯誤!");
                //if (GuiNumber.Length > 20) throw new SystemException("【統一編號】長度錯誤!");
                if (Address.Length > 255) throw new SystemException("【地址】長度錯誤!");
                //if (RemitAccount.Length > 30) throw new SystemException("【銀行帳號】長度錯誤!");
                //if (RestaurantRemark.Length > 500) throw new SystemException("【備註】長度錯誤!");
                #endregion

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷餐廳代號是否重複
                        sql = @"SELECT TOP 1 1
                                FROM EBP.Restaurant
                                WHERE RestaurantNo = @RestaurantNo";
                        dynamicParameters.Add("RestaurantNo", RestaurantNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() < 0) throw new SystemException("【餐廳代號】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EBP.Restaurant (CompanyId, RestaurantNo, RestaurantName
                                , GuiNumber, ResponsiblePerson, Contact
                                , TelNo, FaxNo, Email, Address
                                , AccountDay, PaymentType, InvoiceCount
                                , RemitBank, RemitAccount, RestaurantRemark
                                , StartTime, EndTime
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.RestaurantId
                                VALUES (@CompanyId, @RestaurantNo, @RestaurantName
                                , @GuiNumber, @ResponsiblePerson, @Contact
                                , @TelNo, @FaxNo, @Email, @Address
                                , @AccountDay, @PaymentType, @InvoiceCount
                                , @RemitBank, @RemitAccount, @RestaurantRemark
                                , @StartTime, @EndTime
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                RestaurantNo,
                                RestaurantName,
                                GuiNumber,
                                ResponsiblePerson,
                                Contact,
                                TelNo,
                                FaxNo,
                                Email,
                                Address,
                                AccountDay,
                                PaymentType,
                                InvoiceCount,
                                RemitBank,
                                RemitAccount,
                                RestaurantRemark,
                                StartTime,
                                EndTime,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int RestaurantId = -1;
                        foreach (var item in insertResult)
                        {
                            RestaurantId = item.RestaurantId;
                        }
                        int rowsAffected = insertResult.Count();

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            data = insertResult
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

        #region //AddRestaurantFile 餐廳檔案新增 -- Daiyi 2023.02.03
        public string AddRestaurantFile(int RestaurantId, string FileList, int FileDoc)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷餐廳ID是否正確
                        sql = @"SELECT TOP 1 1
                                FROM EBP.Restaurant
                                WHERE RestaurantId = @RestaurantId";
                        dynamicParameters.Add("RestaurantId", RestaurantId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() < 0) throw new SystemException("【餐廳ID】錯誤!");
                        #endregion

                        int rowsAffected = result.Count();

                        if (FileList != "")
                        {
                            string[] FileIdItem = FileList.Split(',');
                            foreach (var id in FileIdItem)
                            {
                                #region //判斷【File ID】是否存在
                                int FileId = int.Parse(id);
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.FileId 
                                FROM BAS.[File] a
                                WHERE a.FileId = @FileId";
                                dynamicParameters.Add("FileId", FileId);
                                var resultF = sqlConnection.Query(sql, dynamicParameters);
                                if (resultF.Count() <= 0) throw new SystemException("該【File ID】不存在");
                                foreach (var item in resultF)
                                {
                                    FileId = item.FileId;
                                }
                                #endregion

                                FileDoc = FileId;
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO EBP.RestaurantFile (RestaurantId, FileDoc
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.RestaurantId
                                VALUES (@RestaurantId, @FileDoc
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        RestaurantId,
                                        FileDoc,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                var insertResult = sqlConnection.Query(sql, dynamicParameters);
                            }
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
        #endregion

        #region //AddRestaurantMeal 餐廳餐點新增 -- Daiyi 2023.02.03
        public string AddRestaurantMeal(int RestaurantId, string MealList, int MealImage, double MealPrice)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷餐廳ID是否正確
                        sql = @"SELECT TOP 1 1
                                FROM EBP.Restaurant
                                WHERE RestaurantId = @RestaurantId";
                        dynamicParameters.Add("RestaurantId", RestaurantId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() < 0) throw new SystemException("【餐廳ID】錯誤!");
                        #endregion

                        int rowsAffected = result.Count();

                        if (MealList != "")
                        {
                            string[] MenuIdItem = MealList.Split(',');
                            foreach (var id in MenuIdItem)
                            {
                                int FileId = int.Parse(id);
                                #region 
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM EBP.RestaurantMeal a
                                        WHERE a.MealImage = @FileId";
                                dynamicParameters.Add("FileId", FileId);
                                var Imgresult = sqlConnection.Query(sql, dynamicParameters);
                                if (Imgresult.Count() > 0)
                                {
                                    continue;
                                }
                                #endregion

                                #region //判斷【File ID】是否存在
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.FileId 
                                FROM BAS.[File] a
                                WHERE a.FileId = @FileId";
                                dynamicParameters.Add("FileId", FileId);
                                var resultM = sqlConnection.Query(sql, dynamicParameters);
                                if (resultM.Count() <= 0) throw new SystemException("該【File ID】不存在");
                                #endregion

                                MealImage = FileId;
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO EBP.RestaurantMeal (RestaurantId, MealImage, MealPrice
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.RestaurantId
                                VALUES (@RestaurantId, @MealImage, @MealPrice
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        RestaurantId,
                                        MealImage,
                                        MealPrice,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                var insertResult = sqlConnection.Query(sql, dynamicParameters);
                            }
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
        #endregion

        #region //AddOrderMealSteward -- 新增事務員相關資料 -- Yi 2023.05.11
        public string AddOrderMealSteward(int UserId, string Staffs)
        {
            try
            {
                if (Staffs.Length <= 0) throw new SystemException("【所屬同仁資料】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;

                        #region //判斷事務資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[User]
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", UserId);

                        var totalStewards = sqlConnection.Query(sql, dynamicParameters);
                        if (totalStewards.Count() <= 0) throw new SystemException("事務員資料錯誤!");
                        #endregion

                        #region //判斷事務維護資料是否正確
                        string[] staffsList = Staffs.Split(',');

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT COUNT(1) TotalStaffs
                                FROM BAS.[User]
                                WHERE UserId IN @UserId";
                        dynamicParameters.Add("UserId", staffsList);

                        int totalStaffs = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).TotalStaffs;
                        if (totalStaffs != staffsList.Length) throw new SystemException("所屬同仁資料錯誤!");
                        #endregion

                        foreach (var staff in staffsList)
                        {
                            #region //判斷該人員是否已維護
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 b.UserNo, b.UserName
                                    FROM EBP.OrderMealSteward a
                                    INNER JOIN BAS.[User] b ON b.UserId = a.UserId
                                    WHERE a.UserId = @UserId
                                    AND a.StaffUserId = @StaffUserId";
                            dynamicParameters.Add("UserId", UserId);
                            dynamicParameters.Add("StaffUserId", Convert.ToInt32(staff));

                            var resultStaffs = sqlConnection.Query(sql, dynamicParameters);
                            if (resultStaffs.Count() <= 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO EBP.OrderMealSteward (UserId, StaffUserId
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.StewardId
                                    VALUES (@UserId, @StaffUserId
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        UserId,
                                        StaffUserId = Convert.ToInt32(staff),
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();
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
        #endregion

        #region //AddCopyRestaurant-- 複製餐廳資料 -- Yi 2023-05-25
        public string AddCopyRestaurant(int CompanyId, int RestaurantId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷餐廳資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.RestaurantNo, a.RestaurantName, a.GuiNumber, a.ResponsiblePerson
                                , a.Contact, a.TelNo, a.FaxNo, a.Email, a.[Address], a.AccountDay
                                , a.PaymentType, a.InvoiceCount, a.RemitBank, a.RemitAccount, a.RestaurantRemark, a.[Status]
                                FROM EBP.Restaurant a
                                WHERE a.RestaurantId = @RestaurantId";
                        dynamicParameters.Add("RestaurantId", RestaurantId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("餐廳資料錯誤!");

                        string RestaurantNo = "";
                        string RestaurantName = "";
                        string GuiNumber = "";
                        string ResponsiblePerson = "";
                        string Contact = "";
                        string TelNo = "";
                        string FaxNo = "";
                        string Email = "";
                        string Address = "";
                        string AccountDay = "";
                        string PaymentType = "";
                        string InvoiceCount = "";
                        string RemitBank = "";
                        string RemitAccount = "";
                        string RestaurantRemark = "";
                        string Status = "";
                        foreach (var item in result)
                        {
                            RestaurantNo = item.RestaurantNo;
                            RestaurantName = "Copy From: " + item.RestaurantName;
                            GuiNumber = item.GuiNumber;
                            ResponsiblePerson = item.ResponsiblePerson;
                            Contact = item.Contact;
                            TelNo = item.TelNo;
                            FaxNo = item.FaxNo;
                            Email = item.Email;
                            Address = item.Address;
                            AccountDay = item.AccountDay;
                            PaymentType = item.PaymentType;
                            InvoiceCount = item.InvoiceCount;
                            RemitBank = item.RemitBank;
                            RemitAccount = item.RemitAccount;
                            RestaurantRemark = item.RestaurantRemark;
                            Status = item.Status;
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EBP.Restaurant (CompanyId, RestaurantNo, RestaurantName, GuiNumber, ResponsiblePerson
                                , Contact, TelNo, FaxNo, Email, Address, AccountDay
                                , PaymentType, InvoiceCount, RemitBank, RemitAccount, RestaurantRemark, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.RestaurantId
                                VALUES (@CompanyId, @RestaurantNo, @RestaurantName, @GuiNumber, @ResponsiblePerson
                                , @Contact, @TelNo, @FaxNo, @Email, @Address, @AccountDay
                                , @PaymentType, @InvoiceCount, @RemitBank, @RemitAccount, @RestaurantRemark, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId,
                                RestaurantNo,
                                RestaurantName,
                                GuiNumber,
                                ResponsiblePerson,
                                Contact,
                                TelNo,
                                FaxNo,
                                Email,
                                Address,
                                AccountDay,
                                PaymentType,
                                InvoiceCount,
                                RemitBank,
                                RemitAccount,
                                RestaurantRemark,
                                Status,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });

                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = insertResult.Count();

                        int newRestaurantId = 0;
                        foreach (var item in insertResult)
                        {
                            newRestaurantId = item.RestaurantId;
                        }

                        #region //複製餐廳檔案
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.RestaurantId, a.FileDoc
                                FROM EBP.RestaurantFile a
                                WHERE RestaurantId = @RestaurantId";
                        dynamicParameters.Add("RestaurantId", RestaurantId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0)
                        {
                            foreach (var item in result2)
                            {
                                int FileDoc = item.FileDoc;

                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO EBP.RestaurantFile (RestaurantId, FileDoc
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.FileId
                                        VALUES (@RestaurantId, @FileDoc
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        RestaurantId = newRestaurantId,
                                        FileDoc,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });

                                insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();
                            }
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
        #endregion

        #region //Update
        #region //UpdateBoardType -- 公告類別資料更新 -- Zoey 2023.02.02
        public string UpdateBoardType(int BoardTypeId, string TypeName)
        {
            try
            {
                if (TypeName.Length <= 0) throw new SystemException("【類別名稱】不能為空!");
                if (TypeName.Length > 100) throw new SystemException("【類別名稱】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷公告類別資訊是否正確
                        sql = @"SELECT TOP 1 1
                                FROM EBP.BoardType
                                WHERE BoardTypeId = @BoardTypeId";
                        dynamicParameters.Add("BoardTypeId", BoardTypeId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("公告類別資料錯誤!");
                        #endregion

                        #region //判斷類別名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.BoardType
                                WHERE TypeName = @TypeName
                                AND BoardTypeId != @BoardTypeId";
                        dynamicParameters.Add("TypeName", TypeName);
                        dynamicParameters.Add("BoardTypeId", BoardTypeId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【類別名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EBP.BoardType SET
                                TypeName = @TypeName,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE BoardTypeId = @BoardTypeId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TypeName,
                                LastModifiedDate,
                                LastModifiedBy,
                                BoardTypeId
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

        #region //UpdateBoardTypeStatus -- 公告類別狀態更新 -- Zoey 2023.02.02
        public string UpdateBoardTypeStatus(int BoardTypeId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷公告類別資訊是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM EBP.BoardType
                                WHERE BoardTypeId = @BoardTypeId";
                        dynamicParameters.Add("BoardTypeId", BoardTypeId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("公告類別資料錯誤!");

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
                        sql = @"UPDATE EBP.BoardType SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE BoardTypeId = @BoardTypeId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                BoardTypeId
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

        #region //UpdateAnnual -- 年份資料更新 -- Zoey 2023.02.02
        public string UpdateAnnual(int AnnualId, int Annual)
        {
            try
            {
                if (Annual <= 0) throw new SystemException("【年份】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷年份資訊是否正確
                        sql = @"SELECT TOP 1 1
                                FROM EBP.Annual
                                WHERE AnnualId = @AnnualId";
                        dynamicParameters.Add("AnnualId", AnnualId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("年份資料錯誤!");
                        #endregion

                        #region //判斷年份是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.Annual
                                WHERE Annual = @Annual
                                AND AnnualId != @AnnualId";
                        dynamicParameters.Add("Annual", Annual);
                        dynamicParameters.Add("AnnualId", AnnualId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【年份】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EBP.Annual SET
                                Annual = @Annual,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE AnnualId = @AnnualId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Annual,
                                LastModifiedDate,
                                LastModifiedBy,
                                AnnualId
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

        #region //UpdateAnnualStatus -- 年份狀態更新 -- Zoey 2023.02.02
        public string UpdateAnnualStatus(int AnnualId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷公告類別資訊是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM EBP.Annual
                                WHERE AnnualId = @AnnualId";
                        dynamicParameters.Add("AnnualId", AnnualId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("公告類別資料錯誤!");

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
                        sql = @"UPDATE EBP.Annual SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE AnnualId = @AnnualId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                AnnualId
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

        #region //UpdateClubJob -- 社團職位資料更新-- Zoey 2023.02.06
        public string UpdateClubJob(int ClubJobId, string JobName)
        {
            try
            {
                if (JobName.Length <= 0) throw new SystemException("【職位名稱】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷社團職位資訊是否正確
                        sql = @"SELECT TOP 1 1
                                FROM EBP.ClubJob
                                WHERE ClubJobId = @ClubJobId";
                        dynamicParameters.Add("ClubJobId", ClubJobId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("社團職位資料錯誤!");
                        #endregion

                        #region //判斷年份是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.ClubJob
                                WHERE JobName = @JobName
                                AND ClubJobId != @ClubJobId";
                        dynamicParameters.Add("JobName", JobName);
                        dynamicParameters.Add("ClubJobId", ClubJobId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【職位名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EBP.ClubJob SET
                                JobName = @JobName,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ClubJobId = @ClubJobId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                JobName,
                                LastModifiedDate,
                                LastModifiedBy,
                                ClubJobId
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

        #region //UpdateClubJobStatus -- 社團職位狀態更新 -- Zoey 2023.02.06
        public string UpdateClubJobStatus(int ClubJobId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷社團職位資訊是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM EBP.ClubJob
                                WHERE ClubJobId = @ClubJobId";
                        dynamicParameters.Add("ClubJobId", ClubJobId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("社團職位資料錯誤!");

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
                        sql = @"UPDATE EBP.ClubJob SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ClubJobId = @ClubJobId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                ClubJobId
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

        #region //UpdateRestaurant -- 餐廳資料更新 -- Daiyi 2022.12.21
        public string UpdateRestaurant(int RestaurantId, string RestaurantNo, string RestaurantName
            , string GuiNumber, string ResponsiblePerson, string Contact
            , string TelNo, string FaxNo, string Email, string Address
            , string AccountDay, string PaymentType, string InvoiceCount
            , string RemitBank, string RemitAccount, string RestaurantRemark
            , string StartTime, string EndTime, string Status)
        {
            try
            {
                #region //判斷餐廳資料長度
                if (RestaurantNo.Length <= 0) throw new SystemException("【餐廳代號】不能為空!");
                if (RestaurantName.Length <= 0) throw new SystemException("【餐廳名稱】不能為空!");
                //if (GuiNumber.Length <= 0) throw new SystemException("【統一編號】不能為空!");
                //if (ResponsiblePerson.Length <= 0) throw new SystemException("【負責人】不能為空!");
                if (Contact.Length <= 0) throw new SystemException("【聯絡人】不能為空!");
                if (TelNo.Length <= 0) throw new SystemException("【電話】不能為空!");
                //if (FaxNo.Length <= 0) throw new SystemException("【傳真】不能為空!");
                //if (Email.Length <= 0) throw new SystemException("【郵件】不能為空!");
                if (Address.Length <= 0) throw new SystemException("【地址】不能為空!");
                if (PaymentType.Length <= 0) throw new SystemException("【付款方式】不能為空!");
                //if (InvoiceCount.Length <= 0) throw new SystemException("【發票聯數】不能為空!");
                if (StartTime.Length <= 0) throw new SystemException("【開始時間】不能為空!");
                if (EndTime.Length <= 0) throw new SystemException("【結束時間】不能為空!");

                if (RestaurantNo.Length > 50) throw new SystemException("【餐廳代號】長度錯誤!");
                if (RestaurantName.Length > 200) throw new SystemException("【餐廳名稱】長度錯誤!");
                //if (ResponsiblePerson.Length > 30) throw new SystemException("【負責人】長度錯誤!");
                if (Contact.Length > 30) throw new SystemException("【聯絡人】長度錯誤!");
                if (TelNo.Length > 20) throw new SystemException("【電話】長度錯誤!");
                //if (FaxNo.Length > 20) throw new SystemException("【傳真】長度錯誤!");
                //if (Email.Length > 60) throw new SystemException("【郵件】長度錯誤!");
                //if (GuiNumber.Length > 20) throw new SystemException("【統一編號】長度錯誤!");
                if (Address.Length > 255) throw new SystemException("【地址】長度錯誤!");
                //if (RemitAccount.Length > 30) throw new SystemException("【銀行帳號】長度錯誤!");
                //if (RestaurantRemark.Length > 500) throw new SystemException("【備註】長度錯誤!");
                #endregion

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷餐廳資訊是否正確
                        sql = @"SELECT TOP 1 1
                                FROM EBP.Restaurant
                                WHERE RestaurantId = @RestaurantId";
                        dynamicParameters.Add("RestaurantId", RestaurantId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("餐廳資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EBP.Restaurant SET
                                RestaurantNo = @RestaurantNo,
                                RestaurantName = @RestaurantName,
                                GuiNumber = @GuiNumber,
                                ResponsiblePerson = @ResponsiblePerson,
                                Contact = @Contact,
                                TelNo = @TelNo,
                                FaxNo = @FaxNo,
                                Email = @Email,
                                Address = @Address,
                                AccountDay = @AccountDay,
                                PaymentType = @PaymentType,
                                InvoiceCount = @InvoiceCount,
                                RemitBank = @RemitBank,
                                RemitAccount = @RemitAccount,
                                RestaurantRemark = @RestaurantRemark,
                                StartTime = @StartTime,
                                EndTime = @EndTime,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RestaurantId = @RestaurantId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RestaurantNo,
                                RestaurantName,
                                GuiNumber,
                                ResponsiblePerson,
                                Contact,
                                TelNo,
                                FaxNo,
                                Email,
                                Address,
                                AccountDay,
                                PaymentType,
                                InvoiceCount,
                                RemitBank,
                                RemitAccount,
                                RestaurantRemark,
                                StartTime,
                                EndTime,
                                LastModifiedDate,
                                LastModifiedBy,
                                RestaurantId
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

        #region //UpdateRestaurantFile -- 餐廳檔案更新 -- Daiyi 2023.02.02
        public string UpdateRestaurantFile(int RestaurantId, string FileList, int FileDoc)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷餐廳資訊是否正確
                        sql = @"SELECT TOP 1 1
                                FROM EBP.Restaurant
                                WHERE RestaurantId = @RestaurantId";
                        dynamicParameters.Add("RestaurantId", RestaurantId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("餐廳資料錯誤!");
                        #endregion

                        int rowsAffected = result.Count();

                        if (FileList != "")
                        {
                            string[] FileIdItem = FileList.Split(',');
                            foreach (var id in FileIdItem)
                            {
                                #region //判斷【File ID】是否存在
                                int FileId = int.Parse(id);
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.FileId 
                                FROM BAS.[File] a
                                WHERE a.FileId = @FileId";
                                dynamicParameters.Add("FileId", FileId);
                                var resultF = sqlConnection.Query(sql, dynamicParameters);
                                if (resultF.Count() <= 0) throw new SystemException("該【File ID】不存在");
                                FileDoc = -1;
                                foreach (var item in resultF)
                                {
                                    FileId = item.FileId;
                                }
                                #endregion

                                FileDoc = FileId;
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO EBP.RestaurantFile (RestaurantId, FileDoc
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.RestaurantId
                                VALUES (@RestaurantId, @FileDoc
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        RestaurantId,
                                        FileDoc,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                var insertResult = sqlConnection.Query(sql, dynamicParameters);
                            }
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
        #endregion

        #region //UpdateRestaurantMeal -- 餐廳餐點更新 -- Daiyi 2023.02.03
        public string UpdateRestaurantMeal(int RestaurantId, string MealList, int MealImage, double MealPrice)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷餐廳資訊是否正確
                        sql = @"SELECT TOP 1 1
                                FROM EBP.Restaurant
                                WHERE RestaurantId = @RestaurantId";
                        dynamicParameters.Add("RestaurantId", RestaurantId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("餐廳資料錯誤!");
                        #endregion

                        int rowsAffected = result.Count();

                        if (MealList != "")
                        {
                            string[] MenuIdItem = MealList.Split(',');
                            foreach (var id in MenuIdItem)
                            {
                                #region //判斷【File ID】是否存在
                                int FileId = int.Parse(id);
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.FileId 
                                FROM BAS.[File] a
                                WHERE a.FileId = @FileId";
                                dynamicParameters.Add("FileId", FileId);
                                var resultM = sqlConnection.Query(sql, dynamicParameters);
                                if (resultM.Count() <= 0) throw new SystemException("該【File ID】不存在");
                                #endregion

                                MealImage = FileId;
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO EBP.RestaurantMeal (RestaurantId, MealImage, MealPrice
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.RestaurantId
                                VALUES (@RestaurantId, @MealImage, @MealPrice
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        RestaurantId,
                                        MealImage,
                                        MealPrice,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                var insertResult = sqlConnection.Query(sql, dynamicParameters);
                            }
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
        #endregion

        #region //UpdateRestaurantStatus -- 餐廳狀態更新 -- Daiyi 2022.12.21
        public string UpdateRestaurantStatus(int RestaurantId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷餐廳資訊是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM EBP.Restaurant
                                WHERE RestaurantId = @RestaurantId";
                        dynamicParameters.Add("RestaurantId", RestaurantId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("餐廳資料錯誤!");

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
                        sql = @"UPDATE EBP.Restaurant SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RestaurantId = @RestaurantId";
                        dynamicParameters.AddDynamicParams(new
                        {
                            Status = status,
                            LastModifiedDate,
                            LastModifiedBy,
                            RestaurantId
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
        #region //DeleteBoardType -- 公告類別資料刪除 -- Zoey 2023.02.02
        public string DeleteBoardType(int BoardTypeId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷公告類別資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM EBP.BoardType
                                WHERE BoardTypeId = @BoardTypeId";
                        dynamicParameters.Add("BoardTypeId", BoardTypeId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("公告類別資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.BoardType
                                WHERE BoardTypeId = @BoardTypeId";
                        dynamicParameters.Add("BoardTypeId", BoardTypeId);

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

        #region //DeleteAnnual -- 年份資料刪除 -- Zoey 2023.02.02
        public string DeleteAnnual(int AnnualId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷年份資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM EBP.Annual
                                WHERE AnnualId = @AnnualId";
                        dynamicParameters.Add("AnnualId", AnnualId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("年份資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.Annual
                                WHERE AnnualId = @AnnualId";
                        dynamicParameters.Add("AnnualId", AnnualId);

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

        #region //DeleteClubJob -- 社團職位資料刪除 -- Zoey 2023.02.06
        public string DeleteClubJob(int ClubJobId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷社團職位資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM EBP.ClubJob
                                WHERE ClubJobId = @ClubJobId";
                        dynamicParameters.Add("ClubJobId", ClubJobId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("社團職位資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.ClubJob
                                WHERE ClubJobId = @ClubJobId";
                        dynamicParameters.Add("ClubJobId", ClubJobId);

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

        #region//DeleteRestaurantFile --餐廳刪除上傳檔案資料 -- Daiyi 2022.12.23
        public string DeleteRestaurantFile(int FileDoc)
        {
            try
            {
                if (FileDoc <= 0) throw new SystemException("【餐廳資料附件】查無資料!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷資訊是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.FileId, a.FileDoc, b.FileId
                                FROM EBP.RestaurantFile a
                                LEFT JOIN BAS.[File] b ON b.FileId = a.FileDoc
                                WHERE a.FileDoc = @FileDoc";
                        dynamicParameters.Add("FileDoc", FileDoc);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() != 1) throw new SystemException("【餐廳資料附件】錯誤!");
                        int FileId = -1;
                        foreach (var item in result)
                        {
                            FileDoc = item.FileDoc;
                        }
                        #endregion

                        #region//EBP.RestaurantFile
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.RestaurantFile 
                                WHERE FileDoc = @FileDoc";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                FileDoc
                            });
                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        if (rowsAffected != 1) throw new SystemException("【EBP.RestaurantFile】刪除失敗!");
                        #endregion

                        #region//修改 BAS.File狀態
                        FileId = FileDoc;
                        rowsAffected = 0;
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE BAS.[File] SET
                                DeleteStatus = @DeleteStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE FileId = @FileId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                DeleteStatus = "N",
                                LastModifiedDate,
                                LastModifiedBy,
                                FileId
                            });
                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        if (rowsAffected != 1) throw new SystemException("【BAS.[File]】DeleteStatus修改失敗!");
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

        #region//DeleteRestaurantMeal --餐廳刪除上傳菜單資料 -- Daiyi 2022.12.23
        public string DeleteRestaurantMeal(int MealId)
        {
            try
            {
                if (MealId <= 0) throw new SystemException("【餐廳菜單附件】查無資料!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷資訊是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT MealImage
                                FROM EBP.RestaurantMeal
                                WHERE MealId = @MealId";
                        dynamicParameters.Add("MealId", MealId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() != 1) throw new SystemException("【餐廳菜單附件】錯誤!");
                        int MealImage = -1;
                        foreach (var item in result)
                        {
                            MealImage = item.MealImage;
                        }
                        #endregion

                        #region//EBP.RestaurantMeal
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.RestaurantMeal
                                WHERE MealId = @MealId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                MealId
                            });
                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        if (rowsAffected != 1) throw new SystemException("【EBP.RestaurantMeal】刪除失敗!");
                        #endregion

                        #region//修改 BAS.File狀態
                        int FileId = MealImage;
                        rowsAffected = 0;
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE BAS.[File] SET
                                DeleteStatus = @DeleteStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE FileId = @FileId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                DeleteStatus = "Y",
                                LastModifiedDate,
                                LastModifiedBy,
                                FileId
                            });
                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        if (rowsAffected != 1) throw new SystemException("【BAS.[File]】DeleteStatus修改失敗!");
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

        #region //DeleteStaffUser -- 刪除該事務下所屬同仁資料 -- Yi 2023.05.12
        public string DeleteStaffUser(int UserId, int StaffUserId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷該事務下所屬同仁資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EBP.OrderMealSteward
                                WHERE UserId = @UserId
                                AND StaffUserId = @StaffUserId";
                        dynamicParameters.Add("UserId", UserId);
                        dynamicParameters.Add("StaffUserId", StaffUserId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("所屬同仁資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EBP.OrderMealSteward
                                WHERE UserId = @UserId
                                AND StaffUserId = @StaffUserId";
                        dynamicParameters.Add("UserId", UserId);
                        dynamicParameters.Add("StaffUserId", StaffUserId);

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
