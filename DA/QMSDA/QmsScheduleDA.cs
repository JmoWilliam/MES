using System;
using System.Configuration;
using System.Data.SqlClient;
using Dapper;
using NLog;
using System.Collections.Generic;
using System.Linq;

namespace QMSDA
{
    public class QmsScheduleDA
    {
        public string MainConnectionStrings = "";
        private readonly Logger logger = LogManager.GetCurrentClassLogger();

        public QmsScheduleDA()
        {
            MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];
        }

        public IEnumerable<dynamic> Query(string sql, object parameters = null)
        {
            using (var conn = new SqlConnection(MainConnectionStrings))
            {
                return conn.Query(sql, parameters);
            }
        }

        public T QueryFirstOrDefault<T>(string sql, object parameters = null)
        {
            using (var conn = new SqlConnection(MainConnectionStrings))
            {
                return conn.QueryFirstOrDefault<T>(sql, parameters);
            }
        }

        public dynamic QueryFirstOrDefault(string sql, object parameters = null)
        {
            using (var conn = new SqlConnection(MainConnectionStrings))
            {
                return conn.QueryFirstOrDefault(sql, parameters);
            }
        }

        public int Execute(string sql, object parameters = null)
        {
            using (var conn = new SqlConnection(MainConnectionStrings))
            {
                return conn.Execute(sql, parameters);
            }
        }

        public T ExecuteScalar<T>(string sql, object parameters = null)
        {
            using (var conn = new SqlConnection(MainConnectionStrings))
            {
                return conn.ExecuteScalar<T>(sql, parameters);
            }
        }

        public void RunInTransaction(Action<SqlConnection, SqlTransaction> action)
        {
            using (var conn = new SqlConnection(MainConnectionStrings))
            {
                conn.Open();
                using (var tx = conn.BeginTransaction())
                {
                    try
                    {
                        action(conn, tx);
                        tx.Commit();
                    }
                    catch (Exception ex)
                    {
                        try { tx.Rollback(); } catch { }
                        logger.Error(ex);
                        throw;
                    }
                }
            }
        }

        #region //基础数据查询
        
        /// <summary>
        /// 验证员工号是否存在
        /// </summary>
        public dynamic ValidateEmployee(string employeeNo)
        {
            var sql = @"
                SELECT UserId, UserNo, UserName 
                FROM BAS.[User] 
                WHERE UserNo = @EmployeeNo";
            return QueryFirstOrDefault(sql, new { EmployeeNo = employeeNo });
        }

        /// <summary>
        /// 获取机型列表
        /// </summary>
        public IEnumerable<dynamic> GetMachineModes()
        {
            var sql = @"
                SELECT QcMachineModeId, QcMachineModeName 
                FROM QMS.QcMachineMode 
                ORDER BY QcMachineModeId";
            return Query(sql);
        }

        /// <summary>
        /// 根据用户账号取得所属部门与公司资讯
        /// </summary>
        public dynamic GetUserOrganizationByUserNo(string userNo)
        {
            if (string.IsNullOrWhiteSpace(userNo))
            {
                return null;
            }

            var sql = @"
                SELECT TOP 1
                    u.UserId,
                    u.UserNo,
                    u.UserName,
                    d.DepartmentId,
                    d.DepartmentName,
                    c.CompanyId,
                    c.CompanyName
                FROM BAS.[User] u
                LEFT JOIN BAS.Department d ON u.DepartmentId = d.DepartmentId
                LEFT JOIN BAS.Company c ON d.CompanyId = c.CompanyId
                WHERE u.UserNo = @UserNo";

            return QueryFirstOrDefault(sql, new { UserNo = userNo });
        }

        /// <summary>
        /// 获取指定机型的机台列表
        /// </summary>
        public IEnumerable<dynamic> GetMachinesByMode(int qcMachineModeId)
        {
            var sql = @"
                SELECT a.MachineId, b.MachineDesc, a.QcMachineModeId
                FROM QMS.QmmDetail a
                INNER JOIN MES.Machine b ON a.MachineId = b.MachineId
                WHERE 1=1";

            if (qcMachineModeId > 0)
            {
                sql += " AND a.QcMachineModeId = @QcMachineModeId";
            }

            sql += " ORDER BY b.MachineDesc";

            return Query(sql, new { QcMachineModeId = qcMachineModeId });
        }

        /// <summary>
        /// 获取机台详情
        /// </summary>
        public dynamic GetMachineInfo(int machineId)
        {
            var sql = @"
                SELECT a.MachineId, b.MachineDesc, a.QcMachineModeId, c.QcMachineModeName
                FROM QMS.QmmDetail a
                INNER JOIN MES.Machine b ON a.MachineId = b.MachineId
                INNER JOIN QMS.QcMachineMode c ON a.QcMachineModeId = c.QcMachineModeId
                WHERE a.MachineId = @MachineId";
            return QueryFirstOrDefault(sql, new { MachineId = machineId });
        }

        /// <summary>
        /// 获取所有机台类型
        /// </summary>
        public IEnumerable<dynamic> GetAllMachineTypes()
        {
            var sql = @"
                SELECT QcMachineModeId, QcMachineModeNo, QcMachineModeName, QcMachineModeDesc
                FROM QMS.QcMachineMode
                ORDER BY QcMachineModeId";
            return Query(sql);
        }

        #endregion

        #region //量测单据查询

        /// <summary>
        /// 根据量测单号查询量测单据
        /// </summary>
        public dynamic SearchQcRecordById(string qcRecordId, int? companyId)
        {
            var sql = @"
                SELECT 
                    a.QcRecordId,
                    k.QcTypeName, 
                    g.UserName AS CreateByName,
                    d.MtlItemId, 
                    d.MtlItemNo, 
                    d.MtlItemName, 
                    d.MtlItemSpec,
                    i.StatusName AS CheckQcMeasureDataName,
                    c.WoErpPrefix + '-' + c.WoErpNo AS WipNo,
                    a.CreateDate,
                    a.CheckQcMeasureData,
                    a.DisallowanceReason,
                    a.RequestDate,
                    a.ReceiptDate,
                    a.QcStartDate,
                    a.QcFinishDate
                FROM MES.QcRecord a
                INNER JOIN MES.ManufactureOrder b ON a.MoId = b.MoId
                INNER JOIN MES.WipOrder c ON b.WoId = c.WoId
                INNER JOIN PDM.MtlItem d ON c.MtlItemId = d.MtlItemId
                LEFT JOIN MES.MoProcess e ON a.MoProcessId = e.MoProcessId
                LEFT JOIN BAS.[Status] f ON a.QcStatus = f.StatusNo AND f.StatusSchema = 'QcRecord.QcStatus'
                INNER JOIN BAS.[User] g ON a.CreateBy = g.UserId
                LEFT JOIN QMS.QcNotice h ON a.QcNoticeId = h.QcNoticeId
                INNER JOIN QMS.QcType k ON a.QcTypeId = k.QcTypeId
                INNER JOIN BAS.Status i ON a.CheckQcMeasureData = i.StatusNo AND i.StatusSchema = 'QcRecord.CheckQcMeasureData'
                WHERE c.CompanyId = @CompanyId 
                  AND i.StatusName = N'新單據'
                  AND a.QcRecordId = @QcRecordId
                ORDER BY a.CreateDate DESC";
            int resolvedCompanyId = companyId.HasValue && companyId.Value > 0 ? companyId.Value : 4;
            return QueryFirstOrDefault(sql, new { QcRecordId = qcRecordId, CompanyId = resolvedCompanyId });
        }

        /// <summary>
        /// 获取量测单据列表
        /// </summary>
        public IEnumerable<dynamic> GetQcRecords(int qcRecordId, string status, int pageIndex, int pageSize)
        {
            var sql = @"
                SELECT 
                    a.QcRecordId,
                    k.QcTypeName, 
                    g.UserName AS CreateByName,
                    d.MtlItemId, 
                    d.MtlItemNo, 
                    d.MtlItemName, 
                    d.MtlItemSpec,
                    i.StatusName AS CheckQcMeasureDataName,
                    c.WoErpPrefix + '-' + c.WoErpNo AS WipNo,
                    a.CreateDate,
                    a.CheckQcMeasureData,
                    a.DisallowanceReason
                FROM MES.QcRecord a
                INNER JOIN MES.ManufactureOrder b ON a.MoId = b.MoId
                INNER JOIN MES.WipOrder c ON b.WoId = c.WoId
                INNER JOIN PDM.MtlItem d ON c.MtlItemId = d.MtlItemId
                LEFT JOIN MES.MoProcess e ON a.MoProcessId = e.MoProcessId
                LEFT JOIN BAS.[Status] f ON a.QcStatus = f.StatusNo AND f.StatusSchema = 'QcRecord.QcStatus'
                INNER JOIN BAS.[User] g ON a.CreateBy = g.UserId
                LEFT JOIN QMS.QcNotice h ON a.QcNoticeId = h.QcNoticeId
                INNER JOIN QMS.QcType k ON a.QcTypeId = k.QcTypeId
                INNER JOIN BAS.Status i ON a.CheckQcMeasureData = i.StatusNo AND i.StatusSchema = 'QcRecord.CheckQcMeasureData'
                WHERE c.CompanyId = 4";

            if (qcRecordId > 0)
            {
                sql += " AND a.QcRecordId = @QcRecordId";
            }

            if (!string.IsNullOrEmpty(status))
            {
                sql += " AND a.CheckQcMeasureData = @Status";
            }

            sql += " ORDER BY a.CreateDate DESC";

            // 分页
            if (pageIndex > 0 && pageSize > 0)
            {
                int offset = (pageIndex - 1) * pageSize;
                sql += string.Format(" OFFSET {0} ROWS FETCH NEXT {1} ROWS ONLY", offset, pageSize);
            }

            return Query(sql, new { QcRecordId = qcRecordId, Status = status });
        }

        /// <summary>
        /// 获取单个量测单详情
        /// </summary>
        public dynamic GetQcRecordDetail(int qcRecordId)
        {
            var sql = @"
                SELECT 
                    a.QcRecordId,
                    k.QcTypeName, 
                    g.UserName AS CreateByName,
                    d.MtlItemId, 
                    d.MtlItemNo, 
                    d.MtlItemName, 
                    d.MtlItemSpec,
                    i.StatusName AS CheckQcMeasureDataName,
                    c.WoErpPrefix + '-' + c.WoErpNo AS WipNo,
                    a.CreateDate,
                    a.CheckQcMeasureData,
                    a.DisallowanceReason
                FROM MES.QcRecord a
                INNER JOIN MES.ManufactureOrder b ON a.MoId = b.MoId
                INNER JOIN MES.WipOrder c ON b.WoId = c.WoId
                INNER JOIN PDM.MtlItem d ON c.MtlItemId = d.MtlItemId
                INNER JOIN BAS.[User] g ON a.CreateBy = g.UserId
                INNER JOIN QMS.QcType k ON a.QcTypeId = k.QcTypeId
                INNER JOIN BAS.Status i ON a.CheckQcMeasureData = i.StatusNo AND i.StatusSchema = 'QcRecord.CheckQcMeasureData'
                WHERE a.QcRecordId = @QcRecordId AND c.CompanyId = 4";
            return QueryFirstOrDefault(sql, new { QcRecordId = qcRecordId });
        }

        #endregion

        #region //机台排程管理

        /// <summary>
        /// 获取指定机台的排程列表
        /// </summary>
        public IEnumerable<dynamic> GetScheduleList(int machineId, string status, int pageIndex, int pageSize)
        {
            var sql = @"
                SELECT 
                    a.QmmId,
                    a.MachineId,
                    m.MachineDesc,
                    a.QcRecordId,
                    a.CreateDate AS ReceiveTime,
                    a.StartTime,
                    a.EndTime,
                    a.Exector,
                    u.UserName AS ExecutorName,
                    a.Status,
                    a.TimeToDelivery,
                    a.TimeToProcess,
                    a.TimeToFinish,
                    qr.CreateDate AS QcCreateDate,
                    creator.UserName AS CreateByName,
                    CASE 
                        WHEN a.Status = '1' THEN N'排程中'
                        WHEN a.Status = '2' THEN N'量测中'
                        WHEN a.Status = '3' THEN N'已完成'
                        ELSE N'未知'
                    END AS StatusName
                FROM QMS.QcMachineMeasurement a
                INNER JOIN MES.Machine m ON a.MachineId = m.MachineId
                INNER JOIN MES.QcRecord qr ON a.QcRecordId = qr.QcRecordId
                INNER JOIN BAS.[User] creator ON qr.CreateBy = creator.UserId
                LEFT JOIN BAS.[User] u ON a.Exector = u.UserId
                WHERE 1=1";

            if (machineId > 0)
            {
                sql += " AND a.MachineId = @MachineId";
            }

            if (!string.IsNullOrEmpty(status))
            {
                sql += " AND a.Status = @Status";
            }

            sql += " ORDER BY CASE WHEN a.EndTime IS NULL THEN 0 ELSE 1 END, a.ReceiveTime ASC";

            // 分页
            if (pageIndex > 0 && pageSize > 0)
            {
                int offset = (pageIndex - 1) * pageSize;
                sql += string.Format(" OFFSET {0} ROWS FETCH NEXT {1} ROWS ONLY", offset, pageSize);
            }

            return Query(sql, new { MachineId = machineId, Status = status });
        }

        /// <summary>
        /// 驳回送检单
        /// </summary>
        public int RejectQcRecord(int qcRecordId, string reason, int userId)
        {
            var sql = @"
                UPDATE MES.QcRecord 
                SET CheckQcMeasureData = 'F',
                    DisallowanceReason = @Reason,
                    LastModifiedDate = GETDATE(),
                    LastModifiedBy = @UserId
                WHERE QcRecordId = @QcRecordId";
            return Execute(sql, new { QcRecordId = qcRecordId, Reason = reason, UserId = userId });
        }

        #endregion

        #region //机台状态

        /// <summary>
        /// 获取机台状态统计
        /// </summary>
        public IEnumerable<dynamic> GetMachineStatus(int qcMachineModeId)
        {
            var sql = @"
                SELECT 
                    a.MachineId,
                    b.MachineDesc,
                    a.QcMachineModeId,
                    c.QcMachineModeName,
                    ISNULL(d.QueueCount, 0) AS QueueCount,
                    ISNULL(d.MeasuringCount, 0) AS MeasuringCount,
                    CASE 
                        WHEN ISNULL(d.MeasuringCount, 0) > 0 THEN 'detecting'
                        ELSE 'standby'
                    END AS StatusCode,
                    CASE 
                        WHEN ISNULL(d.MeasuringCount, 0) > 0 THEN '检测中'
                        ELSE '待机中'
                    END AS StatusText
                FROM QMS.QmmDetail a
                INNER JOIN MES.Machine b ON a.MachineId = b.MachineId
                INNER JOIN QMS.QcMachineMode c ON a.QcMachineModeId = c.QcMachineModeId
                LEFT JOIN (
                    SELECT 
                        MachineId,
                        SUM(CASE WHEN Status = '1' THEN 1 ELSE 0 END) AS QueueCount,
                        SUM(CASE WHEN Status = '2' THEN 1 ELSE 0 END) AS MeasuringCount
                    FROM QMS.QcMachineMeasurement
                    WHERE Status IN ('1', '2')
                    GROUP BY MachineId
                ) d ON a.MachineId = d.MachineId
                WHERE 1=1";

            if (qcMachineModeId > 0)
            {
                sql += " AND a.QcMachineModeId = @QcMachineModeId";
            }

            sql += " ORDER BY a.QcMachineModeId, b.MachineDesc";

            return Query(sql, new { QcMachineModeId = qcMachineModeId });
        }

        /// <summary>
        /// 根据机台类型获取机台状态
        /// </summary>
        public IEnumerable<dynamic> GetMachineStatusByType(int qcMachineModeId)
        {
            var sql = @"
                SELECT 
                    qd.MachineId,
                    m.MachineDesc,
                    qd.QcMachineModeId,
                    qmm.QcMachineModeName,
                    qmm.QcMachineModeDesc,
                    ISNULL(d.QueueCount, 0) AS QueueCount,
                    ISNULL(d.MeasuringCount, 0) AS MeasuringCount,
                    ISNULL(d.CompletedCount, 0) AS CompletedCount
                FROM QMS.QmmDetail qd
                INNER JOIN MES.Machine m ON qd.MachineId = m.MachineId
                INNER JOIN QMS.QcMachineMode qmm ON qd.QcMachineModeId = qmm.QcMachineModeId
                LEFT JOIN (
                    SELECT 
                        MachineId,
                        SUM(CASE WHEN StartTime IS NULL THEN 1 ELSE 0 END) AS QueueCount,
                        SUM(CASE WHEN StartTime IS NOT NULL AND EndTime IS NULL THEN 1 ELSE 0 END) AS MeasuringCount,
                        SUM(CASE WHEN EndTime IS NOT NULL THEN 1 ELSE 0 END) AS CompletedCount
                    FROM QMS.QcMachineMeasurement
                    GROUP BY MachineId
                ) d ON qd.MachineId = d.MachineId
                WHERE 1=1";

            if (qcMachineModeId > 0)
            {
                sql += " AND qd.QcMachineModeId = @QcMachineModeId";
            }

            sql += " ORDER BY qd.QcMachineModeId, m.MachineDesc";

            return Query(sql, new { QcMachineModeId = qcMachineModeId });
        }

        #endregion

        #region //事务操作

        /// <summary>
        /// 收件 - 将送检单添加到机台排程
        /// </summary>
        public void AddToSchedule(int qcRecordId, int machineId, int userId)
        {
            RunInTransaction((conn, transaction) =>
            {
                // 1. 检查 QcRecord 是否存在
                var checkSql = @"
                    SELECT QcRecordId, CheckQcMeasureData, CreateDate 
                    FROM MES.QcRecord 
                    WHERE QcRecordId = @QcRecordId";
                
                var qcRecord = conn.QueryFirstOrDefault(checkSql, new { QcRecordId = qcRecordId }, transaction);
                
                if (qcRecord == null)
                {
                    throw new Exception("送检单不存在");
                }

                // 2. 检查是否已经在排程中（只检查排程中和量测中状态，允许已驳回的重新加入）
                var existsSql = @"
                    SELECT COUNT(*) 
                    FROM QMS.QcMachineMeasurement 
                    WHERE QcRecordId = @QcRecordId AND Status IN ('1', '2')";
                
                var exists = conn.ExecuteScalar<int>(existsSql, new { QcRecordId = qcRecordId }, transaction);
                
                if (exists > 0)
                {
                    throw new Exception("该送检单已在排程中");
                }
                
                // 3. 如果存在已完成的记录，先删除（允许重新加入）
                var deleteCompletedSql = @"
                    DELETE FROM QMS.QcMachineMeasurement 
                    WHERE QcRecordId = @QcRecordId AND Status = '3'";
                
                conn.Execute(deleteCompletedSql, new { QcRecordId = qcRecordId }, transaction);

                // 4. 计算 TimeToDelivery (从创建到收件的秒数)
                DateTime createDate = qcRecord.CreateDate;
                DateTime receiveTime = DateTime.Now;
                int timeToDelivery = (int)(receiveTime - createDate).TotalSeconds;

                // 5. 新增到 QMS.QcMachineMeasurement
                var insertSql = @"
                    INSERT INTO QMS.QcMachineMeasurement 
                    (MachineId, QcRecordId, StartTime, EndTime, Exector, Status, 
                     TimeToDelivery, TimeToProcess, TimeToFinish,
                     CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                    VALUES 
                    (@MachineId, @QcRecordId, NULL, NULL, NULL, '1',
                     @TimeToDelivery, NULL, NULL,
                     GETDATE(), GETDATE(), @UserId, @UserId)";
                
                conn.Execute(insertSql, new 
                { 
                    MachineId = machineId,
                    QcRecordId = qcRecordId,
                    TimeToDelivery = timeToDelivery,
                    UserId = userId
                }, transaction);

                // 6. 更新 MES.QcRecord.CheckQcMeasureData = 'C' (已收件未量测)
                var updateSql = @"
                    UPDATE MES.QcRecord 
                    SET CheckQcMeasureData = 'C',
                        LastModifiedDate = GETDATE(),
                        LastModifiedBy = @UserId
                    WHERE QcRecordId = @QcRecordId";
                conn.Execute(updateSql, new { QcRecordId = qcRecordId, UserId = userId }, transaction);
            });
        }

        /// <summary>
        /// 开始量测
        /// </summary>
        public void StartMeasurement(int qmmId, int executorId)
        {
            RunInTransaction((conn, transaction) =>
            {
                // 1. 获取记录状态
                var getSql = @"
                    SELECT Status, CreateDate
                    FROM QMS.QcMachineMeasurement 
                    WHERE QmmId = @QmmId";
                
                var record = conn.QueryFirstOrDefault(getSql, new { QmmId = qmmId }, transaction);
                
                if (record == null)
                {
                    throw new Exception("排程记录不存在");
                }

                if (record.Status != "1")
                {
                    throw new Exception("该记录不是排程中状态");
                }

                DateTime startTime = DateTime.Now;
                int timeToProcess = (int)(startTime - ((DateTime)record.CreateDate)).TotalSeconds;

                // 2. 更新 QMS.QcMachineMeasurement
                var updateSql = @"
                    UPDATE QMS.QcMachineMeasurement 
                    SET StartTime = @StartTime,
                        Exector = @ExecutorId,
                        Status = '2',
                        TimeToProcess = @TimeToProcess,
                        LastModifiedDate = GETDATE(),
                        LastModifiedBy = @ExecutorId
                    WHERE QmmId = @QmmId";
                
                conn.Execute(updateSql, new 
                { 
                    QmmId = qmmId,
                    StartTime = startTime,
                    ExecutorId = executorId,
                    TimeToProcess = timeToProcess
                }, transaction);

                // 3. 更新 MES.QcRecord.CheckQcMeasureData = 'S' (开始量测)
                var updateQcSql = @"
                    UPDATE MES.QcRecord 
                    SET CheckQcMeasureData = 'S',
                        LastModifiedDate = GETDATE(),
                        LastModifiedBy = @ExecutorId
                    WHERE QcRecordId = (
                        SELECT QcRecordId FROM QMS.QcMachineMeasurement WHERE QmmId = @QmmId
                    )";
                conn.Execute(updateQcSql, new { QmmId = qmmId, ExecutorId = executorId }, transaction);
            });
        }

        /// <summary>
        /// 结束量测
        /// </summary>
        public void EndMeasurement(int qmmId, int userId)
        {
            RunInTransaction((conn, transaction) =>
            {
                // 1. 获取 StartTime 以计算 TimeToFinish
                var getSql = @"
                    SELECT StartTime, Status 
                    FROM QMS.QcMachineMeasurement 
                    WHERE QmmId = @QmmId";
                
                var record = conn.QueryFirstOrDefault(getSql, new { QmmId = qmmId }, transaction);
                
                if (record == null)
                {
                    throw new Exception("排程记录不存在");
                }

                if (record.Status != "2")
                {
                    throw new Exception("该记录不是量测中状态");
                }

                if (record.StartTime == null)
                {
                    throw new Exception("未找到开始时间");
                }

                DateTime endTime = DateTime.Now;
                int timeToFinish = (int)(endTime - ((DateTime)record.StartTime)).TotalSeconds;

                // 2. 更新 QMS.QcMachineMeasurement
                var updateSql = @"
                    UPDATE QMS.QcMachineMeasurement 
                    SET EndTime = @EndTime,
                        Status = '3',
                        TimeToFinish = @TimeToFinish,
                        LastModifiedDate = GETDATE(),
                        LastModifiedBy = @UserId
                    WHERE QmmId = @QmmId";
                
                conn.Execute(updateSql, new 
                { 
                    QmmId = qmmId,
                    EndTime = endTime,
                    TimeToFinish = timeToFinish,
                    UserId = userId
                }, transaction);

                // 3. 更新 MES.QcRecord.CheckQcMeasureData = 'E' (量测完成)
                var updateQcSql = @"
                    UPDATE MES.QcRecord 
                    SET CheckQcMeasureData = 'E',
                        LastModifiedDate = GETDATE(),
                        LastModifiedBy = @UserId
                    WHERE QcRecordId = (
                        SELECT QcRecordId FROM QMS.QcMachineMeasurement WHERE QmmId = @QmmId
                    )";
                conn.Execute(updateQcSql, new { QmmId = qmmId, UserId = userId }, transaction);
            });
        }

        /// <summary>
        /// 更换机台
        /// </summary>
        public void ChangeMachine(int qmmId, int toMachineId, int userId)
        {
            RunInTransaction((conn, tx) =>
            {
                // 1. 直接更新 QMS.QcMachineMeasurement.MachineId
                var updateSql = @"
                    UPDATE QMS.QcMachineMeasurement 
                    SET MachineId = @ToMachineId,
                        LastModifiedDate = GETDATE(),
                        LastModifiedBy = @UserId
                    WHERE QmmId = @QmmId";
                int rowsAffected = conn.Execute(updateSql, new 
                { 
                    QmmId = qmmId,
                    ToMachineId = toMachineId,
                    UserId = userId
                }, tx);

                if (rowsAffected == 0)
                {
                    throw new Exception("排程记录不存在或更新失败");
                }

                // 2. 记录到 QMS.QcMachineMeasurementLog
                var insertLogSql = @"
                    INSERT INTO QMS.QcMachineMeasurementLog 
                    (MachineId, QcRecordId, StartTime, EndTime, Exector,
                     CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                    SELECT MachineId, QcRecordId, StartTime, EndTime, Exector,
                           GETDATE(), GETDATE(), @UserId, @UserId
                    FROM QMS.QcMachineMeasurement 
                    WHERE QmmId = @QmmId";
                conn.Execute(insertLogSql, new 
                { 
                    QmmId = qmmId,
                    UserId = userId
                }, tx);
            });
        }

        /// <summary>
        /// 根据送检单号查询机台作业列表
        /// </summary>
        public dynamic QueryOrderByNumber(string orderNumber, int page, int pageSize)
        {
            dynamic result = null;
            
            RunInTransaction((conn, tx) =>
            {
                // 1. 首先查找该送检单号是否在排程表中
                var checkSql = @"
                    SELECT TOP 1 a.MachineId, m.MachineDesc, a.QcRecordId
                    FROM QMS.QcMachineMeasurement a
                    INNER JOIN MES.Machine m ON a.MachineId = m.MachineId
                    WHERE a.QcRecordId = @OrderNumber";

                var orderInfo = conn.QueryFirstOrDefault(checkSql, new { OrderNumber = orderNumber }, tx);

                if (orderInfo == null)
                {
                    throw new Exception("未找到该送检单号");
                }

                // 2. 获取该机台的所有送检单（分页）
                var offset = (page - 1) * pageSize;
                var dataSql = @"
                    SELECT 
                        a.QmmId,
                        a.MachineId,
                        m.MachineDesc,
                        a.QcRecordId,
                        a.CreateDate AS ReceiveTime,
                        a.StartTime,
                        a.EndTime,
                        a.Exector,
                        u.UserName AS ExecutorName,
                        a.Status,
                        a.TimeToDelivery,
                        a.TimeToProcess,
                        a.TimeToFinish,
                        qr.CreateDate AS QcCreateDate,
                        creator.UserName AS CreateByName,
                        CASE 
                            WHEN a.Status = '1' THEN N'排程中'
                            WHEN a.Status = '2' THEN N'量测中'
                            WHEN a.Status = '3' THEN N'已完成'
                            ELSE N'未知'
                        END AS StatusName
                    FROM QMS.QcMachineMeasurement a
                    INNER JOIN MES.Machine m ON a.MachineId = m.MachineId
                    INNER JOIN MES.QcRecord qr ON a.QcRecordId = qr.QcRecordId
                    INNER JOIN BAS.[User] creator ON qr.CreateBy = creator.UserId
                    LEFT JOIN BAS.[User] u ON a.Exector = u.UserId
                    WHERE a.MachineId = @MachineId
                      AND (a.EndTime IS NULL OR a.EndTime >= DATEADD(HOUR, -36, GETDATE()))
                    ORDER BY 
                        CASE 
                            WHEN a.EndTime IS NULL THEN 0
                            ELSE 1
                        END,
                        a.CreateDate ASC
                    OFFSET @Offset ROWS FETCH NEXT @PageSize ROWS ONLY";

                var orders = conn.Query(dataSql, new 
                { 
                    MachineId = orderInfo.MachineId,
                    Offset = offset,
                    PageSize = pageSize
                }, tx).ToList();

                // 3. 获取当前单号的状态和创建时间
                var currentOrderSql = @"
                    SELECT a.EndTime, a.CreateDate, a.Status
                    FROM QMS.QcMachineMeasurement a
                    WHERE a.MachineId = @MachineId AND a.QcRecordId = @OrderNumber";

                var currentOrderInfo = conn.QueryFirstOrDefault(currentOrderSql, new { 
                    MachineId = orderInfo.MachineId, 
                    OrderNumber = orderNumber 
                }, tx);

                // 4. 获取排程中送检单的总数（不包含已完成的）
                var scheduledCountSql = @"
                    SELECT COUNT(*)
                    FROM QMS.QcMachineMeasurement a
                    WHERE a.MachineId = @MachineId AND a.EndTime IS NULL";

                var scheduledCount = conn.QuerySingle<int>(scheduledCountSql, new { MachineId = orderInfo.MachineId }, tx);

                // 5. 获取该机台所有送检单的总数（包含已完成的，用于分页）
                var totalCountSql = @"
                    SELECT COUNT(*)
                    FROM QMS.QcMachineMeasurement a
                    WHERE a.MachineId = @MachineId
                      AND (a.EndTime IS NULL OR a.EndTime >= DATEADD(HOUR, -36, GETDATE()))";

                var totalCount = conn.QuerySingle<int>(totalCountSql, new { MachineId = orderInfo.MachineId }, tx);

                // 6. 计算当前单号在排程中的排位（按送检时间逆序，后加入的排后面）
                int currentRank = 0;
                if (currentOrderInfo != null && currentOrderInfo.EndTime == null)
                {
                    var rankSql = @"
                        SELECT COUNT(*) + 1 as Rank
                        FROM QMS.QcMachineMeasurement a
                        WHERE a.MachineId = @MachineId 
                        AND a.EndTime IS NULL 
                        AND a.CreateDate < @CurrentOrderCreateDate";

                    currentRank = conn.QuerySingle<int>(rankSql, new { 
                        MachineId = orderInfo.MachineId, 
                        CurrentOrderCreateDate = currentOrderInfo.CreateDate 
                    }, tx);
                }

                result = new
                {
                    orderNumber = orderNumber,
                    machineId = orderInfo.MachineId,
                    machineDesc = orderInfo.MachineDesc,
                    currentRank = currentRank,
                    scheduledCount = scheduledCount,
                    isCompleted = currentOrderInfo?.EndTime != null,
                    orders = orders,
                    pagination = new
                    {
                        currentPage = page,
                        pageSize = pageSize,
                        totalCount = totalCount,
                        totalPages = (int)Math.Ceiling((double)totalCount / pageSize)
                    }
                };
            });

            return result;
        }

        #endregion
    }
}


