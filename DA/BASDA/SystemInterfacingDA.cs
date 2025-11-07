using Dapper;
using Helpers;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Transactions;
using System.Web;

namespace BASDA
{
    public class SystemInterfacingDA
    {
        public static string MainConnectionStrings = "";

        public static int CurrentCompany = -1;
        public static int CurrentUser = -1;
        public static int CreateBy = -1;
        public static int LastModifiedBy = -1;
        public static DateTime CreateDate = default(DateTime);
        public static DateTime LastModifiedDate = default(DateTime);

        public static string sql = "";
        public static JObject jsonResponse = new JObject();
        public static Logger logger = LogManager.GetCurrentClassLogger();
        public static DynamicParameters dynamicParameters = new DynamicParameters();
        public static SqlQuery sqlQuery = new SqlQuery();

        public SystemInterfacingDA()
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
        #region //GetInterfacing -- 取得系統介接設定資料 -- Ben Ma 2023.09.01
        public string GetInterfacing(int InterfacingId, int CompanyId, string InterfacingType
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.InterfacingId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.CompanyId, a.InterfacingType
                        , b.LogoIcon, b.CompanyName
                        , c.TypeName InterfacingTypeName
                        , ISNULL(LAG(a.InterfacingId) OVER (ORDER BY a.InterfacingId), -1) PreviousId, ISNULL(LEAD(a.InterfacingId) OVER (ORDER BY a.InterfacingId), -1) NextId
                        , (
                            SELECT aa.ConnectionId, aa.ConnectionServer, aa.ConnectionAccount, aa.ConnectionPassword, aa.ConnectionDatabase, aa.[Status]
                            FROM BAS.InterfacingConnection aa
                            WHERE aa.InterfacingId = a.InterfacingId
                            ORDER BY aa.ConnectionId
                            FOR JSON PATH, ROOT('data')
                        ) InterfacingConnection
                        , b.CompanyName + ' ' + c.TypeName InterfacingName";
                    sqlQuery.mainTables =
                        @"FROM BAS.Interfacing a
                        INNER JOIN BAS.Company b ON a.CompanyId = b.CompanyId
                        INNER JOIN BAS.[Type] c ON a.InterfacingType = c.TypeNo AND c.TypeSchema = 'Interfacing.InterfacingType'";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "InterfacingId", @" AND a.InterfacingId = @InterfacingId", InterfacingId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CompanyId", @" AND a.CompanyId = @CompanyId", CompanyId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "InterfacingType", @" AND a.InterfacingType = @InterfacingType", InterfacingType);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.InterfacingId";
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

        #region //GetOperation -- 取得作業模式資料 -- Ben Ma 2023.09.26
        public string GetOperation(int OperationId, int CompanyId, int InterfacingId, string OperationNo
            , string OperationName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.OperationId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.InterfacingId, a.OperationNo, a.OperationName, a.[Status]
                        , b.AutoImport, b.ImportTimeInterval, b.ImportExecuteMode
                        , b.AutoSynchronize, b.SynchronizeTimeInterval, b.SynchronizeExecuteMode
                        , c.InterfacingType
                        , d.CompanyName, d.LogoIcon
                        , e.TypeName InterfacingTypeName
                        , f.StatusName AutoImportName, g.StatusName AutoSynchronizeName
                        , h.TypeName ImportExecuteModeName, i.TypeName SynchronizeExecuteModeName
                        , ISNULL(LAG(a.OperationId) OVER (ORDER BY a.OperationNo), -1) PreviousId, ISNULL(LEAD(a.OperationId) OVER (ORDER BY a.OperationNo), -1) NextId
                        , a.OperationNo + ' ' + a.OperationName OperationWithNo";
                    sqlQuery.mainTables =
                        @"FROM BAS.Operation a
                        INNER JOIN BAS.OperationSetting b ON a.OperationId = b.OperationId
                        INNER JOIN BAS.Interfacing c ON a.InterfacingId = c.InterfacingId
                        INNER JOIN BAS.Company d ON c.CompanyId = d.CompanyId
                        INNER JOIN BAS.[Type] e ON c.InterfacingType = e.TypeNo AND e.TypeSchema = 'Interfacing.InterfacingType'
                        INNER JOIN BAS.[Status] f ON b.AutoImport = f.StatusNo AND f.StatusSchema = 'Boolean'
                        INNER JOIN BAS.[Status] g ON b.AutoSynchronize = g.StatusNo AND g.StatusSchema = 'Boolean'
                        INNER JOIN BAS.[Type] h ON b.ImportExecuteMode = h.TypeNo AND h.TypeSchema = 'OperationSetting.ExecuteMode'
                        INNER JOIN BAS.[Type] i ON b.SynchronizeExecuteMode = i.TypeNo AND i.TypeSchema = 'OperationSetting.ExecuteMode'";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "OperationId", @" AND a.OperationId = @OperationId", OperationId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CompanyId", @" AND c.CompanyId = @CompanyId", CompanyId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "InterfacingId", @" AND a.InterfacingId = @InterfacingId", InterfacingId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "OperationNo", @" AND a.OperationNo '%' + @DemandNo + '%'", OperationNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "OperationName", @" AND a.OperationName '%' + @OperationName + '%'", OperationName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.[Status] IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.OperationNo";
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
        #region //AddInterfacing -- 系統介接設定資料新增 -- Ben Ma 2023.09.06
        public string AddInterfacing(string InterfacingJson)
        {
            try
            {
                if (!InterfacingJson.TryParseJson(out JObject tempJObject)) throw new SystemException("資料格式錯誤");
                
                Interfacing interfacings = JsonConvert.DeserializeObject<Interfacing>(InterfacingJson);

                if (interfacings.CompanyId <= 0) throw new SystemException("【所屬公司】不能為空!");
                if (interfacings.InterfacingType.Length <= 0) throw new SystemException("【介接類型】不能為空!");
                if (interfacings.Details.Count <= 0) throw new SystemException("【介接連線】不能為空!");
                if (interfacings.Details.Any(x => x.ConnectionServer.Length <= 0)) throw new SystemException("【連線主機】不能為空!");
                if (interfacings.Details.Any(x => x.ConnectionServer.Length > 50)) throw new SystemException("【連線主機】長度錯誤!");
                if (interfacings.Details.Any(x => x.ConnectionAccount.Length <= 0)) throw new SystemException("【連線帳號】不能為空!");
                if (interfacings.Details.Any(x => x.ConnectionAccount.Length > 50)) throw new SystemException("【連線帳號】長度錯誤!");
                if (interfacings.Details.Any(x => x.ConnectionPassword.Length <= 0)) throw new SystemException("【連線密碼】不能為空!");
                if (interfacings.Details.Any(x => x.ConnectionPassword.Length > 50)) throw new SystemException("【連線密碼】長度錯誤!");
                if (interfacings.Details.Any(x => x.ConnectionDatabase.Length > 50)) throw new SystemException("【連線資料庫】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷公司資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Company
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", interfacings.CompanyId);

                        var resultCompany = sqlConnection.Query(sql, dynamicParameters);
                        if (resultCompany.Count() <= 0) throw new SystemException("所屬公司資料錯誤!");
                        #endregion

                        #region //判斷介接類型資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Type]
                                WHERE TypeSchema = @TypeSchema
                                AND TypeNo = @TypeNo";
                        dynamicParameters.Add("TypeSchema", "Interfacing.InterfacingType");
                        dynamicParameters.Add("TypeNo", interfacings.InterfacingType);

                        var resultInterfacingType = sqlConnection.Query(sql, dynamicParameters);
                        if (resultInterfacingType.Count() <= 0) throw new SystemException("介接類型資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //單頭新增
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO BAS.Interfacing (CompanyId, InterfacingType
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.InterfacingId
                                VALUES (@CompanyId, @InterfacingType
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                interfacings.CompanyId,
                                interfacings.InterfacingType,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += insertResult.Count();

                        int InterfacingId = -1;
                        foreach (var item in insertResult)
                        {
                            InterfacingId = Convert.ToInt32(item.InterfacingId);
                        }
                        #endregion

                        #region //單身新增
                        interfacings.Details
                            .ToList()
                            .ForEach(x =>
                            {
                                x.InterfacingId = InterfacingId;
                                x.Status = "S";
                                x.CreateDate = CreateDate;
                                x.LastModifiedDate = LastModifiedDate;
                                x.CreateBy = CreateBy;
                                x.LastModifiedBy = LastModifiedBy;
                            });

                        sql = @"INSERT INTO BAS.InterfacingConnection (InterfacingId, ConnectionServer
                                , ConnectionAccount, ConnectionPassword, ConnectionDatabase, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                VALUES (@InterfacingId, @ConnectionServer
                                , @ConnectionAccount, @ConnectionPassword, @ConnectionDatabase, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        rowsAffected += sqlConnection.Execute(sql, interfacings.Details);
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

        #region //AddOperation -- 作業模式資料新增 -- Ben Ma 2023.09.26
        public string AddOperation(string OperationJson)
        {
            try
            {
                if (!OperationJson.TryParseJson(out JObject tempJObject)) throw new SystemException("資料格式錯誤");

                Operation operations = JsonConvert.DeserializeObject<Operation>(OperationJson);

                if (operations.InterfacingId <= 0) throw new SystemException("【系統介接設定】不能為空!");
                if (operations.OperationNo.Length <= 0) throw new SystemException("【作業模式代碼】不能為空!");
                if (operations.OperationName.Length <= 0) throw new SystemException("【作業模式名稱】不能為空!");
                if (operations.Import.Auto.Length <= 0) throw new SystemException("拋轉區塊【自動】不能為空!");
                if (operations.Import.TimeInterval < 0) throw new SystemException("拋轉區塊【時間間隔】不能為空!");
                if (operations.Import.ExecuteMode.Length <= 0) throw new SystemException("拋轉區塊【執行方式】不能為空!");
                if (operations.Synchronize.Auto.Length <= 0) throw new SystemException("同步區塊【自動】不能為空!");
                if (operations.Synchronize.TimeInterval < 0) throw new SystemException("同步區塊【時間間隔】不能為空!");
                if (operations.Synchronize.ExecuteMode.Length <= 0) throw new SystemException("同步區塊【執行方式】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷系統介接設定資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Interfacing
                                WHERE InterfacingId = @InterfacingId";
                        dynamicParameters.Add("InterfacingId", operations.InterfacingId);

                        var resultInterfacing = sqlConnection.Query(sql, dynamicParameters);
                        if (resultInterfacing.Count() <= 0) throw new SystemException("所屬系統介接設定資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //作業模式新增
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO BAS.Operation (InterfacingId, OperationNo, OperationName, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.OperationId
                                VALUES (@InterfacingId, @OperationNo, @OperationName, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                operations.InterfacingId,
                                operations.OperationNo,
                                operations.OperationName,
                                Status = "A",
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += insertResult.Count();

                        int OperationId = -1;
                        foreach (var item in insertResult)
                        {
                            OperationId = Convert.ToInt32(item.OperationId);
                        }
                        #endregion

                        #region //作業模式設定新增
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO BAS.OperationSetting (OperationId, AutoImport, AutoSynchronize
                                , ImportTimeInterval, SynchronizeTimeInterval, ImportExecuteMode, SynchronizeExecuteMode
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                VALUES (@OperationId, @AutoImport, @AutoSynchronize
                                , @ImportTimeInterval, @SynchronizeTimeInterval, @ImportExecuteMode, @SynchronizeExecuteMode
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                OperationId,
                                AutoImport = operations.Import.Auto,
                                AutoSynchronize = operations.Synchronize.Auto,
                                ImportTimeInterval = operations.Import.TimeInterval,
                                SynchronizeTimeInterval = operations.Synchronize.TimeInterval,
                                ImportExecuteMode = operations.Import.ExecuteMode,
                                SynchronizeExecuteMode = operations.Synchronize.ExecuteMode,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
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

        #region //Update
        #region //UpdateInterfacing -- 系統介接設定資料更新 -- Ben Ma 2023.09.07
        public string UpdateInterfacing(int InterfacingId, string InterfacingJson)
        {
            try
            {
                if (!InterfacingJson.TryParseJson(out JObject tempJObject)) throw new SystemException("資料格式錯誤");

                Interfacing interfacings = JsonConvert.DeserializeObject<Interfacing>(InterfacingJson);

                if (interfacings.CompanyId <= 0) throw new SystemException("【所屬公司】不能為空!");
                if (interfacings.InterfacingType.Length <= 0) throw new SystemException("【介接類型】不能為空!");
                if (interfacings.Details.Count <= 0) throw new SystemException("【介接連線】不能為空!");
                if (interfacings.Details.Any(x => x.ConnectionServer.Length <= 0)) throw new SystemException("【連線主機】不能為空!");
                if (interfacings.Details.Any(x => x.ConnectionServer.Length > 50)) throw new SystemException("【連線主機】長度錯誤!");
                if (interfacings.Details.Any(x => x.ConnectionAccount.Length <= 0)) throw new SystemException("【連線帳號】不能為空!");
                if (interfacings.Details.Any(x => x.ConnectionAccount.Length > 50)) throw new SystemException("【連線帳號】長度錯誤!");
                if (interfacings.Details.Any(x => x.ConnectionPassword.Length <= 0)) throw new SystemException("【連線密碼】不能為空!");
                if (interfacings.Details.Any(x => x.ConnectionPassword.Length > 50)) throw new SystemException("【連線密碼】長度錯誤!");
                if (interfacings.Details.Any(x => x.ConnectionDatabase.Length > 50)) throw new SystemException("【連線資料庫】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷公司資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Company
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", interfacings.CompanyId);

                        var resultCompany = sqlConnection.Query(sql, dynamicParameters);
                        if (resultCompany.Count() <= 0) throw new SystemException("所屬公司資料錯誤!");
                        #endregion

                        #region //判斷介接類型資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Type]
                                WHERE TypeSchema = @TypeSchema
                                AND TypeNo = @TypeNo";
                        dynamicParameters.Add("TypeSchema", "Interfacing.InterfacingType");
                        dynamicParameters.Add("TypeNo", interfacings.InterfacingType);

                        var resultInterfacingType = sqlConnection.Query(sql, dynamicParameters);
                        if (resultInterfacingType.Count() <= 0) throw new SystemException("介接類型資料錯誤!");
                        #endregion

                        #region //判斷系統介接設定資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Interfacing
                                WHERE InterfacingId = @InterfacingId";
                        dynamicParameters.Add("InterfacingId", InterfacingId);

                        var resultInterfacing = sqlConnection.Query(sql, dynamicParameters);
                        if (resultInterfacing.Count() <= 0) throw new SystemException("【系統介接設定】資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //單頭更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE BAS.Interfacing SET
                                CompanyId = @CompanyId,
                                InterfacingType = @InterfacingType,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE InterfacingId = @InterfacingId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                interfacings.CompanyId,
                                interfacings.InterfacingType,
                                LastModifiedDate,
                                LastModifiedBy,
                                InterfacingId
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //單身更新
                        foreach (var item in interfacings.Details)
                        {
                            if (item.ConnectionId > 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE BAS.InterfacingConnection SET
                                        ConnectionServer = @ConnectionServer,
                                        ConnectionAccount = @ConnectionAccount,
                                        ConnectionPassword = @ConnectionPassword,
                                        ConnectionDatabase = @ConnectionDatabase,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE ConnectionId = @ConnectionId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        item.ConnectionServer,
                                        item.ConnectionAccount,
                                        item.ConnectionPassword,
                                        item.ConnectionDatabase,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        item.ConnectionId
                                    });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            }
                            else
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO BAS.InterfacingConnection (InterfacingId, ConnectionServer
                                        , ConnectionAccount, ConnectionPassword, ConnectionDatabase, Status
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        VALUES (@InterfacingId, @ConnectionServer
                                        , @ConnectionAccount, @ConnectionPassword, @ConnectionDatabase, @Status
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        InterfacingId,
                                        item.ConnectionServer,
                                        item.ConnectionAccount,
                                        item.ConnectionPassword,
                                        item.ConnectionDatabase,
                                        Status = "S",
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //UpdateInterfacingConnectionStatus -- 介接連線狀態更新 -- Ben Ma 2023.09.07
        public string UpdateInterfacingConnectionStatus(int ConnectionId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷需求來源資料是否正確
                        sql = @"SELECT TOP 1 InterfacingId, Status
                                FROM BAS.InterfacingConnection
                                WHERE ConnectionId = @ConnectionId";
                        dynamicParameters.Add("ConnectionId", ConnectionId);

                        var resultInterfacingConnection = sqlConnection.Query(sql, dynamicParameters);
                        if (resultInterfacingConnection.Count() <= 0) throw new SystemException("介接連線資料錯誤!");

                        int interfacingId = -1;
                        string status = "";
                        foreach (var item in resultInterfacingConnection)
                        {
                            interfacingId = Convert.ToInt32(item.InterfacingId);
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

                        int rowsAffected = 0;
                        if (status == "A")
                        {
                            #region //將其他連線停用，只保留一筆
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE BAS.InterfacingConnection SET
                                    Status = @Status,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE InterfacingId = @InterfacingId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    Status = "S",
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    InterfacingId = interfacingId
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE BAS.InterfacingConnection SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ConnectionId = @ConnectionId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                ConnectionId
                            });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

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

        #region //UpdateOperation -- 作業模式資料新增 -- Ben Ma 2023.09.26
        public string UpdateOperation(int OperationId, string OperationJson)
        {
            try
            {
                if (!OperationJson.TryParseJson(out JObject tempJObject)) throw new SystemException("資料格式錯誤");

                Operation operations = JsonConvert.DeserializeObject<Operation>(OperationJson);

                if (operations.InterfacingId <= 0) throw new SystemException("【系統介接設定】不能為空!");
                if (operations.OperationNo.Length <= 0) throw new SystemException("【作業模式代碼】不能為空!");
                if (operations.OperationName.Length <= 0) throw new SystemException("【作業模式名稱】不能為空!");
                if (operations.Import.Auto.Length <= 0) throw new SystemException("拋轉區塊【自動】不能為空!");
                if (operations.Import.TimeInterval < 0) throw new SystemException("拋轉區塊【時間間隔】不能為空!");
                if (operations.Import.ExecuteMode.Length <= 0) throw new SystemException("拋轉區塊【執行方式】不能為空!");
                if (operations.Synchronize.Auto.Length <= 0) throw new SystemException("同步區塊【自動】不能為空!");
                if (operations.Synchronize.TimeInterval < 0) throw new SystemException("同步區塊【時間間隔】不能為空!");
                if (operations.Synchronize.ExecuteMode.Length <= 0) throw new SystemException("同步區塊【執行方式】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷作業模式資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Operation
                                WHERE OperationId = @OperationId";
                        dynamicParameters.Add("OperationId", OperationId);

                        var resultOperation = sqlConnection.Query(sql, dynamicParameters);
                        if (resultOperation.Count() <= 0) throw new SystemException("【作業模式】資料錯誤!");
                        #endregion

                        #region //判斷系統介接設定資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Interfacing
                                WHERE InterfacingId = @InterfacingId";
                        dynamicParameters.Add("InterfacingId", operations.InterfacingId);

                        var resultInterfacing = sqlConnection.Query(sql, dynamicParameters);
                        if (resultInterfacing.Count() <= 0) throw new SystemException("所屬系統介接設定資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //作業模式資料更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE BAS.Operation SET
                                InterfacingId = @InterfacingId,
                                OperationNo = @OperationNo,
                                OperationName = @OperationName,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE OperationId = @OperationId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                operations.InterfacingId,
                                operations.OperationNo,
                                operations.OperationName,
                                LastModifiedDate,
                                LastModifiedBy,
                                OperationId
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //作業模式設定刪除
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE BAS.OperationSetting
                                WHERE OperationId = @OperationId";
                        dynamicParameters.Add("OperationId", OperationId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //作業模式設定新增
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO BAS.OperationSetting (OperationId, AutoImport, AutoSynchronize
                                , ImportTimeInterval, SynchronizeTimeInterval, ImportExecuteMode, SynchronizeExecuteMode
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                VALUES (@OperationId, @AutoImport, @AutoSynchronize
                                , @ImportTimeInterval, @SynchronizeTimeInterval, @ImportExecuteMode, @SynchronizeExecuteMode
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                OperationId,
                                AutoImport = operations.Import.Auto,
                                AutoSynchronize = operations.Synchronize.Auto,
                                ImportTimeInterval = operations.Import.TimeInterval,
                                SynchronizeTimeInterval = operations.Synchronize.TimeInterval,
                                ImportExecuteMode = operations.Import.ExecuteMode,
                                SynchronizeExecuteMode = operations.Synchronize.ExecuteMode,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //UpdateOperationStatus -- 作業模式狀態更新 -- Ben Ma 2023.09.26
        public string UpdateOperationStatus(int OperationId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷作業模式資料是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM BAS.Operation
                                WHERE OperationId = @OperationId";
                        dynamicParameters.Add("OperationId", OperationId);

                        var resultOperation = sqlConnection.Query(sql, dynamicParameters);
                        if (resultOperation.Count() <= 0) throw new SystemException("介接連線資料錯誤!");

                        string status = "";
                        foreach (var item in resultOperation)
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
                        sql = @"UPDATE BAS.Operation SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE OperationId = @OperationId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                OperationId
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
        #region //DeleteInterfacing -- 系統介接設定資料刪除 -- Ben Ma 2023.09.06
        public string DeleteInterfacing(int InterfacingId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷系統介接設定資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Interfacing
                                WHERE InterfacingId = @InterfacingId";
                        dynamicParameters.Add("InterfacingId", InterfacingId);

                        var resultInterfacing = sqlConnection.Query(sql, dynamicParameters);
                        if (resultInterfacing.Count() <= 0) throw new SystemException("系統介接設定資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除關聯table
                        #region //介接連線刪除
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE BAS.InterfacingConnection
                                WHERE InterfacingId = @InterfacingId";
                        dynamicParameters.Add("InterfacingId", InterfacingId);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE BAS.Interfacing
                                WHERE InterfacingId = @InterfacingId";
                        dynamicParameters.Add("InterfacingId", InterfacingId);

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

        #region //DeleteInterfacingConnection -- 介接連線資料刪除 -- Ben Ma 2023.09.06
        public string DeleteInterfacingConnection(int ConnectionId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷介接連線資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.InterfacingConnection
                                WHERE ConnectionId = @ConnectionId";
                        dynamicParameters.Add("ConnectionId", ConnectionId);

                        var resultInterfacingConnection = sqlConnection.Query(sql, dynamicParameters);
                        if (resultInterfacingConnection.Count() <= 0) throw new SystemException("介接連線資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE BAS.InterfacingConnection
                                WHERE ConnectionId = @ConnectionId";
                        dynamicParameters.Add("ConnectionId", ConnectionId);

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
