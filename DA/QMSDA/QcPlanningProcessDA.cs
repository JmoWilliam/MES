using Dapper;
using Helpers;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Text.RegularExpressions;
using System.Transactions;
using System.Web;

namespace QMSDA
{
    public class QcPlanningProcessDA
    {
        public string MainConnectionStrings = "";
        public string ErpConnectionStrings = "";
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

        public QcPlanningProcessDA()
        {
            MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];
            ErpConnectionStrings = ConfigurationManager.AppSettings["ErpDb"];

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
        #region //GetLoginUserInfo -- 取得使用者資料 -- Ann 2024-03-15
        public string GetLoginUserInfo()
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT UserNo, UserName
                            FROM BAS.[User]
                            WHERE UserId = @UserId";
                    dynamicParameters.Add("UserId", CurrentUser);

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

        #region //GetQcMachineModePlanning -- 取得量測機台排程資料 -- Ann 2024-08-27
        public string GetQcMachineModePlanning(int QmmpId, int QrpId, int QmmDetailId, int QcMachineModeId, string StartDate, string Status, string MtlItemName, int QcRecordId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.QmmpId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.QrpId, a.QmmDetailId
                        , FORMAT(a.EstimatedStartDate, 'MM-dd HH:mm:ss') EstimatedStartDate
                        , FORMAT(a.EstimatedEndDate, 'MM-dd HH:mm:ss') EstimatedEndDate
                        , FORMAT(a.StartDate, 'MM-dd HH:mm:ss') StartDate
                        , FORMAT(a.EndDate, 'MM-dd HH:mm:ss') EndDate
                        , a.Sort, a.Status
                        , d.QcRecordId
                        , h.MtlItemId, h.MtlItemNo, h.MtlItemName
                        , i.MachineNo, i.MachineDesc
                        , j.StatusName
                        , DATEDIFF(HOUR, ISNULL(a.StartDate, GETDATE()), a.EstimatedStartDate) DateDifferenceInDays";
                    sqlQuery.mainTables =
                        @"FROM QMS.QcMachineModePlanning a 
                        INNER JOIN QMS.QmmDetail b ON a.QmmDetailId = b.QmmDetailId
                        INNER JOIN QMS.QcMachineMode c ON b.QcMachineModeId = c.QcMachineModeId
                        INNER JOIN QMS.QcRecordPlanning d ON a.QrpId = d.QrpId
                        INNER JOIN MES.QcRecord e ON d.QcRecordId = e.QcRecordId
                        INNER JOIN MES.ManufactureOrder f ON e.MoId = f.MoId
                        INNER JOIN MES.WipOrder g ON f.WoId = g.WoId
                        INNER JOIN PDM.MtlItem h ON g.MtlItemId = h.MtlItemId
                        INNER JOIN MES.Machine i ON b.MachineId = i.MachineId
                        INNER JOIN BAS.[Status] j ON j.StatusSchema = 'QcMachineModePlanning.Status' AND j.StatusNo = a.[Status]";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND c.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QmmpId", @" AND a.QmmpId = @QmmpId", QmmpId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QrpId", @" AND a.QrpId = @QrpId", QrpId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QmmDetailId", @" AND a.QmmDetailId = @QmmDetailId", QmmDetailId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcMachineModeId", @" AND b.QcMachineModeId = @QcMachineModeId", QcMachineModeId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND FORMAT(a.EstimatedStartDate, 'yyyy-MM-dd') = @StartDate", StartDate);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemName", @" AND h.MtlItemName LIKE '%' + @MtlItemName + '%'", MtlItemName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcRecordId", @" AND d.QcRecordId = @QcRecordId", QcRecordId);

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.QmmDetailId, a.Sort";
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

        #endregion

        #region //Update
        #region //UpdateQcPlanningProcess -- 更新量測排程過站狀態 -- Ann 2024-09-18
        public string UpdateQcPlanningProcess(int QmmpId, string PlanningStatus)
        {
            try
            {
                int rowsAffected = -1;
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認量測機台排程資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.Status, a.StartDate, a.EndDate
                                FROM QMS.QcMachineModePlanning a 
                                WHERE a.QmmpId = @QmmpId";
                        dynamicParameters.Add("QmmpId", QmmpId);

                        var QcMachineModePlanningResult = sqlConnection.Query(sql, dynamicParameters);

                        if (QcMachineModePlanningResult.Count() <= 0) throw new SystemException("量測機台排程資料錯誤!!");

                        DateTime? StartDate = new DateTime();
                        DateTime? EndDate = new DateTime();

                        string currentStatus = "";
                        foreach (var item in QcMachineModePlanningResult)
                        {
                            currentStatus = item.Status;
                            StartDate = item.StartDate;
                            EndDate = item.EndDate;
                        }
                        #endregion

                        switch (PlanningStatus)
                        {
                            #region //開工
                            case "A":
                                if (currentStatus != "N" && currentStatus != "F" && currentStatus != "P") throw new SystemException("排程狀態錯誤!!");
                                UpdateQcMachineModePlanning(DateTime.Now, (DateTime?)null, PlanningStatus, QmmpId);
                                InsertQcPlanningProcessLog(DateTime.Now, (DateTime?)null, "A", QmmpId);
                                break;
                            #endregion
                            #region //完工
                            case "F":
                                if (currentStatus != "A" && currentStatus != "P") throw new SystemException("排程狀態錯誤!!");
                                UpdateQcMachineModePlanning(StartDate, DateTime.Now, PlanningStatus, QmmpId);
                                UpdateQcPlanningProcessLog(DateTime.Now, PlanningStatus, QmmpId);
                                break;
                            #endregion
                            #region //暫停
                            case "P":
                                if (currentStatus != "A") throw new SystemException("排程狀態錯誤!!");
                                UpdateQcMachineModePlanning(StartDate, (DateTime?)null, PlanningStatus, QmmpId);
                                UpdateQcPlanningProcessLog(DateTime.Now, PlanningStatus, QmmpId);
                                break;
                            #endregion
                            #region //取消
                            case "C":
                                if (currentStatus != "A") throw new SystemException("排程狀態錯誤!!");
                                UpdateQcMachineModePlanning((DateTime?)null, (DateTime?)null, "N", QmmpId);
                                UpdateQcPlanningProcessLog((DateTime?)null, PlanningStatus, QmmpId);
                                break;
                            #endregion
                        }

                        #region //Update QMS.QcMachineModePlanning
                        void UpdateQcMachineModePlanning(DateTime? startDate, DateTime? endDate, string status, int qmmpId)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE QMS.QcMachineModePlanning SET
                                    StartDate = @StartDate,
                                    EndDate = @EndDate,
                                    Status = @Status,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE QmmpId = @QmmpId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    StartDate = startDate,
                                    EndDate = endDate,
                                    Status = status,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    QmmpId = qmmpId
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        }
                        #endregion

                        #region //Insert QMS.QcPlanningProcessLog
                        void InsertQcPlanningProcessLog(DateTime? startDate, DateTime? endDate, string status, int qmmpId)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO QMS.QcPlanningProcessLog (QmmpId, StartDate, EndDate, Status
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.LogId
                                    VALUES (@QmmpId, @StartDate, @EndDate, @Status
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    QmmpId = qmmpId,
                                    StartDate = startDate,
                                    EndDate = endDate,
                                    Status = status,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();
                        }
                        #endregion

                        #region //Update QMS.QcPlanningProcessLog
                        void UpdateQcPlanningProcessLog(DateTime? endDate, string status, int qmmpId)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE a SET 
                                    EndDate = @EndDate,
                                    Status = @Status,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    FROM QMS.QcPlanningProcessLog a 
                                    WHERE a.LogId = (
                                        SELECT TOP 1 a.LogId 
                                        FROM QMS.QcPlanningProcessLog a 
                                        WHERE a.QmmpId = @QmmpId
                                        AND a.EndDate IS NULL
                                        ORDER BY a.CreateDate DESC
                                    )";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    EndDate = endDate,
                                    Status = status,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    QmmpId = qmmpId
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //Delete

        #endregion
    }
}
