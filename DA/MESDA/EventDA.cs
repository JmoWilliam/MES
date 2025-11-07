using Dapper;
using Helpers;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Transactions;
using System.Web;


namespace MESDA
{
    public class EventDA
    {
        public string MainConnectionStrings = "";
        public string ErpConnectionStrings = "";
        public string HrmConnectionStrings = "";
        public string OldMesConnectionStrings = "";

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

        public EventDA()
        {
            MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];
            ErpConnectionStrings = ConfigurationManager.AppSettings["ErpDb"];
            HrmConnectionStrings = ConfigurationManager.AppSettings["HrmDb"];
            OldMesConnectionStrings = ConfigurationManager.AppSettings["OldDb"];

            GetUserInfo();
            CreateDate = DateTime.Now;
            LastModifiedDate = DateTime.Now;
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

        #region//Get

        #region //GetUserEvent -- 取得人員事件 -- Ellie 2023.10.30
        public string GetUserEvent(int UserEventId,int UserEventItemId
            , string StartDateBegin, string StartDateEnd, int Operator
            , string OrderBy, int PageIndex, int PageSize)
        {
            try 
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.UserEventId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" , b.UserEventItemId, b.UserEventItemName
                           , FORMAT(a.StartDate, 'yyyy-MM-dd HH:mm') StartDate
                           , FORMAT(a.StartDate, 'yyyy-MM-dd') StartYmd
                           , FORMAT(a.StartDate, 'HH:mm:ss') StartTime
                           , FORMAT(a.FinishDate, 'yyyy-MM-dd HH:mm') FinishDate
                           , FORMAT(a.FinishDate, 'yyyy-MM-dd') FinishYmd
                           , FORMAT(a.FinishDate, 'HH:mm:ss') FinishTime
                           , CONVERT(VARCHAR(8), DATEADD(ss, a.OperationTime, '00:00'), 108) OperationTime
                           , a.Operator, ca.UserName OperationUser
                           , FORMAT(a.LastModifiedDate, 'yyyy-MM-dd') LastModifiedDate
                           , cb.UserName LastModifiedBy
                           , a.[Source] Source, cd.CompanyId
                           , a.Status
                          ";
                    sqlQuery.mainTables =
                        @"FROM MES.UserEvent a
                        INNER JOIN MES.UserEventItem b ON a.UserEventItemId = b.UserEventItemId
                        INNER JOIN BAS.[User] ca ON ca.UserId = a.Operator
                        INNER JOIN BAS.[User] cb ON cb.UserId = a.LastModifiedBy
                        INNER JOIN BAS.Department cc ON cc.DepartmentId = ca.DepartmentId
                        INNER JOIN BAS.Company cd ON cd.CompanyId = cc.CompanyId
                    ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND cd.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserEventId", @" AND a.UserEventId = @UserEventId", UserEventId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserEventItemId", @" AND a.UserEventItemId = @UserEventItemId", UserEventItemId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDateBegin", @" AND a.StartDate >= @StartDateBegin", StartDateBegin);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDateEnd", @" AND a.StartDate <= @StartDateEnd", StartDateEnd);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Operator", @" AND a.Operator = @Operator", Operator);


                    sqlQuery.conditions = queryCondition;

                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.UserEventId DESC";
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

        #region //GetMachineEvent -- 取得機台事件 -- Ellie 2023.11.06
        public string GetMachineEvent(int MachineEventId, int MachineEventItemId, int ShopId, int MachineId
            , string StartDateBegin, string StartDateEnd, int Operator
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.MachineEventId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @"  
                           , b.MachineEventItemId
                           , b.MachineEventName 
                           , FORMAT(a.StartDate, 'yyyy-MM-dd HH:mm:ss') StartDate
                           , FORMAT(a.StartDate, 'yyyy-MM-dd') StartYmd
                           , FORMAT(a.StartDate, 'HH:mm:ss') StartTime
                           , FORMAT(a.FinishDate, 'yyyy-MM-dd HH:mm:ss') FinishDate
                           , FORMAT(a.FinishDate, 'yyyy-MM-dd') FinishYmd
                           , FORMAT(a.FinishDate, 'HH:mm:ss') FinishTime
                           , CONVERT(VARCHAR(8), DATEADD(ss, a.OperationTime, '00:00'), 108) OperationTime
                           , e.ShopId
                           , e.ShopName
                           , b.MachineId 
                           , d.MachineName
                           , d.MachineDesc
                           , a.Operator, ca.UserName OperationUser
                           , FORMAT(a.LastModifiedDate, 'yyyy-MM-dd') LastModifiedDate
                           , ISNULL(a.LastModifiedBy, '') LastModifiedBy
                           , cb.UserName LastModifiedName
                           , a.[Source] Source
                           , cd.CompanyId
                          ";
                    sqlQuery.mainTables =
                        @"FROM MES.MachineEvent a
                        INNER JOIN MES.MachineEventItem b ON a.MachineEventItemId = b.MachineEventItemId
                        INNER JOIN MES.Machine d ON b.MachineId = d.MachineId
                        INNER JOIN MES.WorkShop e ON e.ShopId = d.ShopId
                        INNER JOIN BAS.[User] ca ON ca.UserId = a.Operator
                        LEFT JOIN BAS.[User] cb ON cb.UserId = a.LastModifiedBy
                        INNER JOIN BAS.Department cc ON cc.DepartmentId = ca.DepartmentId
                        INNER JOIN BAS.Company cd ON cd.CompanyId = cc.CompanyId
                    ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND cd.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MachineEventId", @" AND a.MachineEventId = @MachineEventId", MachineEventId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MachineEventItemId", @" AND a.MachineEventItemId = @MachineEventItemId", MachineEventItemId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ShopId", @" AND e.ShopId = @ShopId", ShopId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MachineId", @" AND b.MachineId = @MachineId", MachineId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Operator", @" AND a.Operator = @Operator", Operator);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDateBegin", @" AND a.StartDate >= @StartDateBegin", StartDateBegin);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDateEnd", @" AND a.StartDate <= @StartDateEnd", StartDateEnd);



                    sqlQuery.conditions = queryCondition;

                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.MachineEventId DESC";
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

        #region //GetProcessEvent -- 取得製程參數事件 -- Ellie 2023.11.13
        public string GetProcessEvent(int BarcodeEventId, int ProcessEventItemId, int ModeId, int ParameterId, int ProcessId, int TypeNo
            , string StartDateBegin, string StartDateEnd, int OperatorId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.BarcodeEventId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" , i.WoErpPrefix, i.WoErpNo, h.WoSeq, g.BarcodeNo, b.ProcessEventName, d.ModeName, e.ProcessName, ce.TypeName, ca.UserName OperationUser
                           , FORMAT(a.StartDate, 'yyyy-MM-dd HH:mm:ss') StartDate
                           , FORMAT(a.StartDate, 'yyyy-MM-dd') StartYmd, FORMAT(a.StartDate, 'HH:mm:ss') StartTime
                           , FORMAT(a.FinishDate, 'yyyy-MM-dd HH:mm:ss') FinishDate, FORMAT(a.FinishDate, 'yyyy-MM-dd') FinishYmd, FORMAT(a.FinishDate, 'HH:mm:ss') FinishTime
                           , CONVERT(VARCHAR(8), DATEADD(ss, a.CycleTime, '00:00'), 108) OperationTime
                           , FORMAT(a.LastModifiedDate, 'yyyy-MM-dd') LastModifiedDate, ISNULL(a.LastModifiedBy, '') LastModifiedBy, cb.UserName LastModifiedName, a.[Source] Source
                           , cd.CompanyId, f.BarcodeProcessId, b.ProcessEventItemId,g.BarcodeId, d.ModeId, c.ParameterId, e.ProcessId , b.ProcessEventType TypeNo, a.OperatorId, g.MoId
                          ";
                    sqlQuery.mainTables =
                        @"FROM MES.BarcodeProcessEvent a
                           INNER JOIN MES.ProcessEventItem b ON b.ProcessEventItemId = a.ProcessEventItemId
                           INNER JOIN MES.ProcessParameter c ON c.ParameterId = b.ParameterId
                           INNER JOIN MES.ProdMode d ON d.ModeId = c.ModeId 
                           INNER JOIN MES.Process e ON e.ProcessId = c.ProcessId 
                           INNER JOIN BAS.[User] ca ON ca.UserId = a.OperatorId
                           INNER JOIN BAS.[User] cb ON cb.UserId = a.CreateBy
                           INNER JOIN BAS.Type ce on ce.TypeNo = b.ProcessEventType and ce.TypeSchema='ProcessEventItem'
                           INNER JOIN BAS.Department cc ON cc.DepartmentId = ca.DepartmentId
                           INNER JOIN BAS.Company cd ON cd.CompanyId = cc.CompanyId
						   INNER JOIN MES.BarcodeProcess f ON a.BarcodeProcessId = f.BarcodeProcessId
						   INNER JOIN MES.Barcode g ON f.BarcodeId = g.BarcodeId
						   INNER JOIN MES.ManufactureOrder h ON g.MoId = h.MoId
						   INNER JOIN MES.WipOrder i ON h.WoId = i.WoId
                    ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND cd.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "BarcodeEventId", @" AND a.BarcodeEventId = @BarcodeEventId", BarcodeEventId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProcessEventItemId", @" AND b.ProcessEventItemId = @ProcessEventItemId", ProcessEventItemId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ModeId", @" AND d.ModeId = @ModeId", ModeId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ParameterId", @" AND c.ParameterId = @ParameterId", ParameterId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProcessId", @" AND e.ProcessId = @ProcessId", ProcessId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TypeNo", @" AND b.ProcessEventType = @TypeNo", TypeNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "OperatorId", @" AND a.OperatorId = @OperatorId", OperatorId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDateBegin", @" AND a.StartDate >= @StartDateBegin", StartDateBegin);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDateEnd", @" AND a.StartDate <= @StartDateEnd", StartDateEnd);

                    sqlQuery.conditions = queryCondition;

                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.BarcodeEventId DESC";
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

        #region//Add

        #region //AddUserEvent -- 人員事件新增 -- Ellie 2023.10.30
        public string AddUserEvent(int UserEventId, int UserEventItemId, string StartYmd, string StartTime
            , string FinishYmd, string FinishTime, int Operator, string Status)
        {
            try
            {
                DateTime NewStartDate = Convert.ToDateTime(StartYmd + " " + StartTime);
                DateTime NewFinishDate = Convert.ToDateTime(FinishYmd + " " + FinishTime);

                TimeSpan ts = NewFinishDate - NewStartDate;
                string time = ts.TotalSeconds.ToString();

                if (StartYmd.Length <= 0) throw new SystemException("【開始日期】不能為空!");
                if (StartTime.Length <= 0) throw new SystemException("【開始時間】不能為空!");
                if (FinishYmd.Length <= 0) throw new SystemException("【結束日期】不能為空!");
                if (FinishTime.Length <= 0) throw new SystemException("【結束時間】不能為空!");
                if (NewFinishDate < NewStartDate) throw new SystemException("【結束日期】要大於【開始日期】");
                if (Operator <= 0) throw new SystemException("【操作員】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.UserEvent (UserEventItemId, StartDate, FinishDate
                                , OperationTime, Operator, Source, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.UserEventId
                                VALUES (@UserEventItemId , @StartDate, @FinishDate
                                , @OperationTime, @Operator, @Source, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                UserEventItemId,
                                StartDate = NewStartDate,
                                FinishDate = NewFinishDate,
                                OperationTime = time,
                                Operator,
                                Source = "/UserEventManagement/AddUserEvent",
                                Status,
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
                            msg = "(" + rowsAffected + " rows affected)",
                            data = insertResult
                        });
                        #endregion

                        transactionScope.Complete();
                    }
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

        #region //AddMachineEvent -- 機台事件新增 -- Ellie 2023.11.07
        public string AddMachineEvent(int MachineEventId, int MachineEventItemId
            , int Operator, string StartYmd, string StartTime
            , string FinishYmd, string FinishTime)
        {
            try
            {
                DateTime NewStartDate = Convert.ToDateTime(StartYmd + " " + StartTime);
                DateTime NewFinishDate = Convert.ToDateTime(FinishYmd + " " + FinishTime);

                TimeSpan ts = NewFinishDate - NewStartDate;
                string time = ts.TotalSeconds.ToString();

                if (Operator <= 0) throw new SystemException("【操作員】不能為空!");
                if (StartYmd.Length <= 0) throw new SystemException("【開始日期】不能為空!");
                if (StartTime.Length <= 0) throw new SystemException("【開始時間】不能為空!");
                if (FinishYmd.Length <= 0) throw new SystemException("【結束日期】不能為空!");
                if (FinishTime.Length <= 0) throw new SystemException("【結束時間】不能為空!");
                if (NewFinishDate <= NewStartDate) throw new SystemException("【結束日期】要大於【開始日期】");


                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.MachineEvent (MachineEventItemId
                                , Operator
                                , StartDate, FinishDate
                                , OperationTime, Source
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.MachineEventId
                                VALUES (@MachineEventItemId 
                                , @Operator
                                , @StartDate, @FinishDate
                                , @OperationTime, @Source
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                MachineEventItemId,
                                Operator,
                                StartDate = NewStartDate,
                                FinishDate = NewFinishDate,
                                OperationTime = time,
                                Source = "/MachineEventManagement/AddMachineEvent",
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
                            msg = "(" + rowsAffected + " rows affected)",
                            data = insertResult
                        });
                        #endregion

                        transactionScope.Complete();
                    }
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

        #region //AddProcessEvent -- 製程事件管理新增 -- Ellie 2023.11.14
        public string AddProcessEvent(int BarcodeEventId, int ProcessEventItemId
            , int OperatorId, string StartYmd, string StartTime
            , string FinishYmd, string FinishTime)
        {
            try
            {
                DateTime NewStartDate = Convert.ToDateTime(StartYmd + " " + StartTime);
                DateTime NewFinishDate = Convert.ToDateTime(FinishYmd + " " + FinishTime);

                TimeSpan ts = NewFinishDate - NewStartDate;
                string time = ts.TotalSeconds.ToString();

                if (OperatorId <= 0) throw new SystemException("【操作員】不能為空!");
                if (StartYmd.Length <= 0) throw new SystemException("【開始日期】不能為空!");
                if (StartTime.Length <= 0) throw new SystemException("【開始時間】不能為空!");
                if (FinishYmd.Length <= 0) throw new SystemException("【結束日期】不能為空!");
                if (FinishTime.Length <= 0) throw new SystemException("【結束時間】不能為空!");
                if (NewFinishDate <= NewStartDate) throw new SystemException("【結束日期】要大於【開始日期】");


                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.BarcodeProcessEvent (ProcessEventItemId
                                , OperatorId
                                , StartDate, FinishDate
                                , CycleTime, Source
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.BarcodeEventId
                                VALUES (@ProcessEventItemId 
                                , @OperatorId
                                , @StartDate, @FinishDate
                                , @CycleTime, @Source
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ProcessEventItemId,
                                OperatorId,
                                StartDate = NewStartDate,
                                FinishDate = NewFinishDate,
                                CycleTime = time,
                                Source = "/ProcessEventManagement/AddProcessEvent",
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
                            msg = "(" + rowsAffected + " rows affected)",
                            data = insertResult
                        });
                        #endregion

                        transactionScope.Complete();
                    }
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

        #region//Update

        #region //UpdateUserEvent -- 人員事件更新 -- Ellie 2023.11.02
        public string UpdateUserEvent(int UserEventId, int UserEventItemId, string StartYmd, string StartTime
            , string FinishYmd, string FinishTime , int Operator)
        {
            try
            {
                DateTime NewStartDate = Convert.ToDateTime(StartYmd + " " + StartTime);
                DateTime NewFinishDate = Convert.ToDateTime(FinishYmd + " " + FinishTime);

               // TimeSpan ts = NewFinishDate - NewStartDate;
             //   string time = ts.TotalSeconds.ToString();

                if (StartYmd.Length <= 0) throw new SystemException("【開始日期】不能為空!");
                if (StartTime.Length <= 0) throw new SystemException("【開始時間】不能為空!");
                if (FinishYmd.Length <= 0) throw new SystemException("【結束日期】不能為空!");
                if (FinishTime.Length <= 0) throw new SystemException("【結束時間】不能為空!");

                if (NewFinishDate <= NewStartDate) throw new SystemException("【結束日期】要大於【開始日期】");
                if (Operator <= 0) throw new SystemException("操作員【名稱】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷人員事件是否正確
                        sql = @"SELECT TOP 1 1
                                FROM MES.UserEvent
                                WHERE UserEventId = @UserEventId";
                        dynamicParameters.Add("UserEventId", UserEventId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("人員事件資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.UserEvent SET
                                UserEventItemId = @UserEventItemId,
                                StartDate = @StartDate,
                                FinishDate = @FinishDate,
                                OperationTime = DATEDIFF(ss,StartDate,@FinishDate),
                                Operator = @Operator,
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE UserEventId = @UserEventId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                UserEventItemId,
                                StartDate = NewStartDate,
                                FinishDate = NewFinishDate,
                             //   OperationTime = time,
                                Operator,
                                Status = "A", //啟用
                                LastModifiedDate,
                                LastModifiedBy,
                                UserEventId
                            });
                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion

                        transactionScope.Complete();
                    }
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

        #region //UpdateMachineEvent -- 機台事件更新 -- Ellie 2023.11.07
        public string UpdateMachineEvent(int MachineEventId, int MachineEventItemId
            , int ShopId , int MachineId
            , int Operator, string StartYmd, string StartTime
            , string FinishYmd, string FinishTime)
        {
            try
            {
                DateTime NewStartDate = Convert.ToDateTime(StartYmd + " " + StartTime);
                DateTime NewFinishDate = Convert.ToDateTime(FinishYmd + " " + FinishTime);

              // TimeSpan ts = NewFinishDate - NewStartDate;
           //    string time = ts.TotalSeconds.ToString();


                if (Operator <= 0) throw new SystemException("操作員【名稱】不能為空!");
                if (StartYmd.Length <= 0) throw new SystemException("【開始日期】不能為空!");
                if (StartTime.Length <= 0) throw new SystemException("【開始時間】不能為空!");
                if (FinishYmd.Length <= 0) throw new SystemException("【結束日期】不能為空!");
                if (FinishTime.Length <= 0) throw new SystemException("【結束時間】不能為空!");

                if (NewFinishDate <= NewStartDate) throw new SystemException("【結束日期】要大於【開始日期】");


                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷機台事件是否正確
                        sql = @"SELECT TOP 1 1
                                FROM MES.MachineEvent
                                WHERE MachineEventId = @MachineEventId";
                        dynamicParameters.Add("MachineEventId", MachineEventId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("人員事件資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.MachineEvent SET
                                MachineEventItemId = @MachineEventItemId,
                                Operator = @Operator,
                                StartDate = @StartDate,
                                FinishDate = @FinishDate,
                                OperationTime = DATEDIFF(ss,StartDate,@FinishDate),
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE MachineEventId = @MachineEventId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                MachineEventItemId,
                                Operator,
                                StartDate = NewStartDate,
                                FinishDate = NewFinishDate,
                               // OperationTime = time,
                                LastModifiedDate,
                                LastModifiedBy,
                                MachineEventId
                            });
                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion

                        transactionScope.Complete();
                    }
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

        #region //UpdateProcessEvent -- 製程事件更新 -- Ellie 2023.11.14
        public string UpdateProcessEvent(int BarcodeEventId, int ProcessEventItemId, int OperatorId, string StartYmd, string StartTime
            , string FinishYmd, string FinishTime)
        {
            try
            {
                DateTime NewStartDate = Convert.ToDateTime(StartYmd + " " + StartTime);
                DateTime NewFinishDate = Convert.ToDateTime(FinishYmd + " " + FinishTime);

               // TimeSpan ts = NewFinishDate - NewStartDate;
               // string time = ts.TotalSeconds.ToString();


                if (OperatorId <= 0) throw new SystemException("操作員【名稱】不能為空!");
                if (StartYmd.Length <= 0) throw new SystemException("【開始日期】不能為空!");
                if (StartTime.Length <= 0) throw new SystemException("【開始時間】不能為空!");
                if (FinishYmd.Length <= 0) throw new SystemException("【結束日期】不能為空!");
                if (FinishTime.Length <= 0) throw new SystemException("【結束時間】不能為空!");
                if (NewFinishDate < NewStartDate) throw new SystemException("【結束日期】要大於【開始日期】");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷機台事件是否正確
                        sql = @"SELECT TOP 1 1
                                FROM MES.BarcodeProcessEvent
                                WHERE BarcodeEventId = @BarcodeEventId";
                        dynamicParameters.Add("BarcodeEventId", BarcodeEventId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("製程事件資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.BarcodeProcessEvent SET
                                ProcessEventItemId = @ProcessEventItemId,
                                OperatorId = @OperatorId,
                                StartDate = @StartDate,
                                FinishDate = @FinishDate,
                                CycleTime = DATEDIFF(ss,StartDate,@FinishDate),
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE BarcodeEventId = @BarcodeEventId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ProcessEventItemId,
                                OperatorId,
                                StartDate = NewStartDate,
                                FinishDate = NewFinishDate,
                               // OperationTime = time,
                                LastModifiedDate,
                                LastModifiedBy,
                                BarcodeEventId
                            });
                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion

                        transactionScope.Complete();
                    }
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

        #region//Delete

        #region //DeleteUserEvent -- 人員事件刪除 -- Ellie 2023.10.30
        public string DeleteUserEvent(int UserEventId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷人員事件資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.UserEvent
                                WHERE UserEventId = @UserEventId";
                        dynamicParameters.Add("UserEventId", UserEventId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("人員事件資料錯誤!");
                        #endregion


                        int rowsAffected = 0;

                        #region //刪除群組table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.UserEvent
                                WHERE UserEventId = @UserEventId";
                        dynamicParameters.Add("UserEventId", UserEventId);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion

                        transactionScope.Complete();
                    }
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

        #region //DeleteMachineEvent -- 機台事件刪除 -- Ellie 2023.11.07
        public string DeleteMachineEvent(int MachineEventId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷機台事件資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.MachineEvent
                                WHERE MachineEventId = @MachineEventId";
                        dynamicParameters.Add("MachineEventId", MachineEventId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("人員事件資料錯誤!");
                        #endregion


                        int rowsAffected = 0;

                        #region //刪除群組table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.MachineEvent
                                WHERE MachineEventId = @MachineEventId";
                        dynamicParameters.Add("MachineEventId", MachineEventId);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion

                        transactionScope.Complete();
                    }
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

        #region //DeleteProcessEvent -- 製程事件刪除 -- Ellie 2023.11.14
        public string DeleteProcessEvent(int BarcodeEventId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷機台事件資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.BarcodeProcessEvent
                                WHERE BarcodeEventId = @BarcodeEventId";
                        dynamicParameters.Add("BarcodeEventId", BarcodeEventId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("製程事件資料錯誤!");
                        #endregion


                        int rowsAffected = 0;

                        #region //刪除群組table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.BarcodeProcessEvent
                                WHERE BarcodeEventId = @BarcodeEventId";
                        dynamicParameters.Add("BarcodeEventId", BarcodeEventId);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion

                        transactionScope.Complete();
                    }
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