using Dapper;
using Helpers;
using Newtonsoft.Json.Linq;
using NLog;
using SCMDA;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Transactions;

namespace BPMDA
{
    public class DeptDocumentDA
    {
        public string MainConnectionStrings = "";
        public string ErpConnectionStrings = "";
        public string HrmConnectionStrings = "";
        public string HrmEtergeConnectionStrings = "";
        public string MESEtergeConnectionStrings = "";

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
        public LinePushNotifiHelper linePushNotifiHelper = new LinePushNotifiHelper();

        public DeptDocumentDA()
        {
            MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];
            ErpConnectionStrings = ConfigurationManager.AppSettings["ErpDb"];
            MESEtergeConnectionStrings = "Data Source=192.168.20.137;Initial Catalog=MES;User Id=sa;Password=JmoMes2019;integrated security=false;persist security info=True;MultipleActiveResultSets=True;";
            //HrmConnectionStrings = ConfigurationManager.AppSettings["HrmDb"];
            //HrmEtergeConnectionStrings = ConfigurationManager.AppSettings["HrmDb_Eterge"];

            //CurrentCompany = BaseHelper.CurrentCompany();
            //CurrentUser = BaseHelper.CurrentUser();
            // CreateBy = BaseHelper.CurrentUser();
            // LastModifiedBy = BaseHelper.CurrentUser();
            CreateDate = DateTime.Now;
            LastModifiedDate = DateTime.Now;

            dynamicParameters = new DynamicParameters();
            sqlQuery = new SqlQuery();
        }

        #region //Get
        #endregion

        #region //Add
        #region //AddDocumentSynchronize -- 文件簽核資料建立 -- Daiyi 2022.12.14
        public string AddDocumentSynchronize(string Company, string FolderNo, string SenderUserNo
            , string DocName, int SenderId, string SendTime, string ApproverUserNo, int ApproverId, string ApproverTime
            , string Remark, string Status)
        {
            try
            {
                if (SenderUserNo.Length <= 0) throw new SystemException("【寄件者】不能為空!");
                if (ApproverUserNo.Length <= 0) throw new SystemException("【收件者】不能為空!");
                if (Company.Length <= 0) throw new SystemException("【公司】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int CompanyId = -1;
                    string CompanyName = "";
                    int SenderUserId = -1;
                    string SenderUserName = "";
                    int DepartmentId = -1;
                    int ApproverUserId = -1;
                    string ApproverUserName = "";
                    int FolderId = -1;

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷公司資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 CompanyId, CompanyName
                                FROM BAS.Company
                                WHERE CompanyNo = @Company";
                        dynamicParameters.Add("Company", Company);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【公司】資料錯誤!");

                        foreach (var item in result)
                        {
                            CompanyId = Convert.ToInt32(item.CompanyId);
                            CompanyName = item.CompanyName;
                        }
                        #endregion

                        #region //判斷寄件人員資料是否正確及所屬部門
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserId, a.UserName, b.DepartmentId, b.DepartmentNo
                                FROM BAS.[User] a
                                LEFT JOIN BAS.Department b ON b.DepartmentId = a.DepartmentId
                                WHERE UserNo = @SenderUserNo";
                        dynamicParameters.Add("SenderUserNo", SenderUserNo);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【寄件人員】資料錯誤!");

                        foreach (var item in result)
                        {
                            SenderUserId = Convert.ToInt32(item.UserId);
                            SenderUserName = item.UserName;
                            DepartmentId = item.DepartmentId;
                        }
                        #endregion

                        #region //判斷收件人員資料是否正確及所屬部門
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserId, a.UserName, b.DepartmentId, b.DepartmentNo
                                FROM BAS.[User] a
                                LEFT JOIN BAS.Department b ON b.DepartmentId = a.DepartmentId
                                WHERE UserNo = @ApproverUserNo";
                        dynamicParameters.Add("ApproverUserNo", ApproverUserNo);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【收件人員】資料錯誤!");

                        foreach (var item in result)
                        {
                            ApproverUserId = Convert.ToInt32(item.UserId);
                            ApproverUserName = item.UserName;
                        }
                        #endregion

                        #region //卷宗夾No轉ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.FolderId, a.FolderNo
                                FROM BPM.Folder a
                                LEFT JOIN BPM.DeptDocument b ON b.FolderId = a.FolderId
                                WHERE FolderNo = @FolderNo";
                        dynamicParameters.Add("FolderNo", FolderNo);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【卷宗】資料錯誤!");

                        foreach (var item in result)
                        {
                            FolderId = Convert.ToInt32(item.FolderId);
                        }
                        #endregion

                        #region //insert app資料至DB
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO BPM.DeptDocument (FolderId, DepartmentId, DocName, SenderId, SendTime, 
                                ApproverId, ApproverTime, Remark,
                                [Status], CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.DocId
                                VALUES (@FolderId, @DepartmentId, @DocName, @SenderId, @SendTime, 
                                @ApproverId, @ApproverTime, @Remark,
                                @Status, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                FolderId,
                                DepartmentId,
                                DocName,
                                SenderId = SenderUserId,
                                SendTime,
                                ApproverId = ApproverUserId,
                                ApproverTime,
                                Remark,
                                Status = '1',
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);
                        #endregion

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
        #endregion

        #region //Update
        #region //UpdateDocumentSynchronize -- 文件簽核資料更新 -- Daiyi 2022.12.14
        public string UpdateDocumentSynchronize(string Company, string UserNo
            , string FolderNo, string DocName, int SenderId, string SendTime
            , string ApproverUserNo, int ApproverId, string ApproverTime, string Remark, string Status)

        {
            try
            {
                if (ApproverUserNo.Length <= 0) throw new SystemException("【寄件者】不能為空!");

                if (Company.Length <= 0) throw new SystemException("【公司】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int CompanyId = -1;
                    string CompanyName = "";
                    int ApproverUserId = -1;
                    string ApproverUserName = "";
                    int DocId = -1;

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷公司資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 CompanyId, CompanyName
                                FROM BAS.Company
                                WHERE CompanyNo = @Company";
                        dynamicParameters.Add("Company", Company);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【公司】資料錯誤!");

                        foreach (var item in result)
                        {
                            CompanyId = Convert.ToInt32(item.CompanyId);
                            CompanyName = item.CompanyName;
                        }
                        #endregion

                        #region //判斷收件人員資料是否正確及所屬部門
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserId, a.UserName, b.DepartmentId, b.DepartmentNo
                                FROM BAS.[User] a
                                LEFT JOIN BAS.Department b ON b.DepartmentId = a.DepartmentId
                                WHERE a.UserNo = @ApproverUserNo";
                        dynamicParameters.Add("ApproverUserNo", ApproverUserNo);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【收件人員】資料錯誤!");

                        foreach (var item in result)
                        {
                            ApproverUserId = Convert.ToInt32(item.UserId);
                            ApproverUserName = item.UserName;
                        }
                        #endregion

                        int rowsAffected = -1;
                        if (Status == "1")
                        {
                            #region //判斷BPM.DeptDocument是否有資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.DocId, a.DocName, a.ApproverId, a.FolderId, b.FolderNo
                                , c.StatusNo StatusNo, c.StatusName StatusName
                                FROM BPM.DeptDocument a
                                LEFT JOIN BPM.Folder b ON b.FolderId = a.FolderId
                                LEFT JOIN BAS.[Status] c ON c.StatusNo = a.[Status] AND c.StatusSchema = 'DeptDocument.Status'
                                WHERE b.FolderNo = @FolderNo
                                AND a.Status = '1'";
                            dynamicParameters.Add("FolderNo", FolderNo);
                            var result1 = sqlConnection.Query(sql, dynamicParameters);
                            if (result1.Count() <= 0) throw new SystemException("該【DocId】不存在");

                            foreach (var item in result1)
                            {
                                DocId = Convert.ToInt32(item.DocId);
                            }
                            #endregion

                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE BPM.DeptDocument SET
                                ApproverId = @ApproverId,
                                ApproverTime = @ApproverTime,
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE DocId = @DocId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    ApproverId = ApproverUserId,
                                    ApproverTime,
                                    Status = '2',
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    DocId
                                });
                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            if (rowsAffected != 1) throw new SystemException("該筆【文件簽核資料】修改失敗");
                        }
                        else if (Status == "2")
                        {
                            #region //判斷BPM.DeptDocument是否有資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.DocId, a.DocName, a.ApproverId, a.FolderId, b.FolderNo
                                , c.StatusNo StatusNo, c.StatusName StatusName
                                FROM BPM.DeptDocument a
                                LEFT JOIN BPM.Folder b ON b.FolderId = a.FolderId
                                LEFT JOIN BAS.[Status] c ON c.StatusNo = a.[Status] AND c.StatusSchema = 'DeptDocument.Status'
                                WHERE b.FolderNo = @FolderNo
                                AND a.Status = '2'";
                            dynamicParameters.Add("FolderNo", FolderNo);
                            var result1 = sqlConnection.Query(sql, dynamicParameters);
                            if (result1.Count() <= 0) throw new SystemException("該【DocId】不存在");

                            foreach (var item in result1)
                            {
                                DocId = Convert.ToInt32(item.DocId);
                            }
                            #endregion

                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE BPM.DeptDocument SET
                                ApproverId = @ApproverId,
                                ApproverTime = @ApproverTime,
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE DocId = @DocId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    ApproverId = ApproverUserId,
                                    ApproverTime,
                                    Status = '3',
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    DocId
                                });
                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            if (rowsAffected != 1) throw new SystemException("該筆【文件簽核資料】修改失敗");
                        }

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            data = DocId
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

        #region //Inquiry 查詢
        #region //InquiryDocumentSynchronize -- 文件簽核資料查詢 -- Daiyi 2022.12.16
        public string InquiryDocumentSynchronize(string Company, string FolderNo, string SenderUserNo
            , string DocName, int SenderId, string SendTime, string ApproverUserNo, int ApproverId, string ApproverTime
            , string Remark, string Status)

        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int CompanyId = -1;
                    string CompanyName = "";
                    int SenderUserId = -1;
                    string SenderUserName = "";
                    int ApproverUserId = -1;
                    string ApproverUserName = "";

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷公司資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 CompanyId, CompanyName
                                FROM BAS.Company
                                WHERE CompanyNo = @Company";
                        dynamicParameters.Add("Company", Company);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【公司】資料錯誤!");

                        foreach (var item in result)
                        {
                            CompanyId = Convert.ToInt32(item.CompanyId);
                            CompanyName = item.CompanyName;
                        }
                        #endregion

                        #region //判斷寄件人員資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserId, a.UserName, b.DepartmentId, b.DepartmentNo, c.ApproverId
                                FROM BAS.[User] a
                                LEFT JOIN BAS.Department b ON b.DepartmentId = a.DepartmentId
                                LEFT JOIN BPM.DeptDocument c ON c.DepartmentId = b.DepartmentId
                                LEFT JOIN BAS.[User] d ON d.UserId = c.ApproverId 
                                WHERE a.UserNo = @SenderUserNo";
                        dynamicParameters.Add("SenderUserNo", SenderUserNo);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        //if (result.Count() <= 0) throw new SystemException("【寄件人員】資料錯誤!");

                        foreach (var item in result)
                        {
                            SenderUserId = Convert.ToInt32(item.UserId);
                            SenderUserName = item.UserName;
                            ApproverUserId = Convert.ToInt32(item.ApproverId);
                        }
                        #endregion

                        #region //查詢BPM.DeptDocument資料
                        dynamicParameters = new DynamicParameters();
                        sqlQuery.mainKey = "a.DocId";
                        sqlQuery.auxKey = "";
                        sqlQuery.columns =
                        @", a.DocName, a.FolderId, b.FolderNo, a.SenderId, c.UserName, a.SendTime,
                                a.ApproverId, d.UserName, a.ApproverTime, a.Remark
                                , e.StatusNo StatusNo, e.StatusName StatusName";
                        sqlQuery.mainTables =
                        @"FROM BPM.DeptDocument a
                                LEFT JOIN BPM.Folder b ON b.FolderId = a.FolderId
                                LEFT JOIN BAS.[User] c ON c.UserId = a.SenderId
                                LEFT JOIN BAS.[User] d ON d.UserId = a.ApproverId
                                LEFT JOIN BAS.[Status] e ON e.StatusNo = a.[Status] AND e.StatusSchema = 'DeptDocument.Status'";
                        sqlQuery.auxTables = "";
                        string queryCondition = @"
                                AND a.Status != '3'
                                AND a.SenderId= @SenderUserId
                                ";
                        dynamicParameters.Add("SenderUserName", SenderUserName);
                        dynamicParameters.Add("SenderUserId", SenderUserId);
                        sqlQuery.conditions = queryCondition;
                        sqlQuery.distinct = false;
                        result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            data = result
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

        #region //DemandMessage 推播
        #region //DemandMessage -- 文件簽核資料推播(尚未完成編寫) -- Daiyi 2022.12.19 (參考電商推播API)
        public string DemandMessage(string Company, string FolderNo
            , string txType, int MessageType, string DemandMessageContent, string SendUserNo, string SendStatus)
        {
            try
            {
                if (DemandMessageContent.Length <= 0) throw new SystemException("【訊息內容】不能為空!");
                if (SendUserNo.Length <= 0) throw new SystemException("【員工工號】不能為空!");
                if (SendStatus.Length <= 0) throw new SystemException("【訊息狀態】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    string RoomId = "";
                    int FolderId = -1;
                    int SendUserId = -1;
                    string SendUserName = "";
                    int DepartmentId = -1;

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷寄件人員資料是否正確及所屬部門
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserId, a.UserName, b.DepartmentId, b.DepartmentNo
                                FROM BAS.[User] a
                                LEFT JOIN BAS.Department b ON b.DepartmentId = a.DepartmentId
                                WHERE UserNo = @SendUserNo";
                        dynamicParameters.Add("SendUserNo", SendUserNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【寄件人員】資料錯誤!");

                        foreach (var item in result)
                        {
                            SendUserId = Convert.ToInt32(item.UserId);
                            SendUserName = item.UserName;
                            DepartmentId = item.DepartmentId;
                        }
                        #endregion

                        #region //查詢SendUserId 是否有建立過Room id，若無則新建
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT DemandMessageContent, MessageType, DocId, RoomId, SendUserId, CreateDate AS SendDate
                                FROM BPM.DemandMessage
                                WHERE SendUserId = @SendUserId
                                AND SendStatus != ''";
                        dynamicParameters.Add("SendUserId", SendUserId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (RoomId != "0")
                        {
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MessageType,
                                    DemandMessageContent,
                                    RoomId,
                                    SendUserId,
                                    SendStatus,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                        }
                        else
                        {
                            #region //新建UUid
                            Guid UUId = Guid.NewGuid();
                            RoomId = UUId.ToString();
                            #endregion
                        }

                        foreach (var item in result)
                        {
                            SendUserId = Convert.ToInt32(item.SendUserId);
                            SendStatus = item.SendStatus;
                        }
                        #endregion

                        #region //卷宗夾No轉ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.FolderId, a.FolderNo, b.DocId
                                FROM BPM.Folder a
                                LEFT JOIN BPM.DeptDocument b ON b.FolderId = a.FolderId
                                WHERE FolderNo = @FolderNo";
                        dynamicParameters.Add("FolderNo", FolderNo);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【卷宗】資料錯誤!");

                        foreach (var item in result)
                        {
                            FolderId = Convert.ToInt32(item.FolderId);
                        }
                        #endregion


                        if (SendStatus == "N")
                        {
                            #region //insert資料至Table
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO BPM.DemandMessage (DemandMessageId, MessageType, DemandMessageContent, RoomId, 
                                SendUserId, SendStatus, CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.DocId
                                VALUES (@DemandMessageId, @MessageType, @DemandMessageContent, @RoomId, 
                                @SendUserId, @SendStatus, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MessageType,
                                    DemandMessageContent,
                                    RoomId,
                                    SendUserId,
                                    SendStatus,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);
                            #endregion

                            int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            if (rowsAffected != 1) throw new SystemException("該筆【推播訊息】新增失敗");

                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                msg = "(" + rowsAffected + " rows affected)"
                            });
                            #endregion
                        }
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

        #region //IOS製令在製查詢
        public string GetMESProcessForiOS(int Company, int ModeId, string WoErpNo, string MtlItemNo, string StartDate, string EndDate)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1500
a.MoId,
(b.WoErpPrefix + '-' + b.WoErpNo) WipNo,
b.StockInQty, 
b.PlanQty,
b.StockInQty,
c.MtlItemNo,
c.MtlItemName,
c.MtlItemSpec,
(
SELECT 
ISNULL(d3.ProcessAlias,'未投入') CurrentProcess, 
ISNULL(d4.ProcessAlias,'已結束') NextProcess,
SUM (ISNULL(d2.BarcodeQty,0)) BarcodeQty
FROM MES.WipOrder d
INNER JOIN MES.ManufactureOrder d1 ON d.WoId = d1.WoId
LEFT JOIN MES.Barcode d2 ON d1.MoId = d2.MoId
LEFT JOIN MES.MoProcess d3 ON d2.CurrentMoProcessId = d3.MoProcessId
LEFT JOIN MES.MoProcess d4 ON d2.NextMoProcessId = d4.MoProcessId
WHERE d1.MoId = a.MoId
GROUP BY 
ISNULL(d3.ProcessAlias,'未投入'),
ISNULL(d4.ProcessAlias,'已結束')
FOR JSON PATH, ROOT('data')
) StationDetail
FROM MES.ManufactureOrder a
INNER JOIN MES.WipOrder b ON b.WoId = a.WoId
INNER JOIN PDM.MtlItem c ON c.MtlItemId = b.MtlItemId
";
                    string queryCondition = @"AND b.CompanyId = @Company
                                              ";
                    dynamicParameters.Add("CompanyId", Company);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ModeId", @" AND a.ModeId = @ModeId", ModeId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "WoErpNo", @" AND (b.WoErpPrefix + '-' + b.WoErpNo) LIKE '%' + @WoErpNo + '%'", WoErpNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemNo", @" AND (c.MtlItemNo LIKE '%' + @MtlItemNo + '%' OR c.MtlItemName LIKE '%' + @MtlItemNo + '%')", MtlItemNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND b.DocDate >= @StartDate ", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND b.DocDate <= @EndDate ", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    sql += @" ORDER BY  a.MoId DESC";
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

        #region //IOS取得管理職名單
        public string GetDepartmentLeader()
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"select UserId,DepartmentId,(UserNo + ' ' + UserName) [User],Email,Job
from BAS.[User]
where JobType = '管理制'
and Status = 'A'
";

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

        #region //取得APP功能
        public string GetAPPFunction(int UserId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.SubSystemId,a.SubSystemCode,SubSystemName,a.KeyText,b.SubSystemUserId
FROM BAS.SubSystem a
INNER JOIN BAS.SubSystemUser b ON b.UserId = @UserId AND b.SubSystemId = a.SubSystemId AND b.[Status] = 'A'
ORDER BY a.SubSystemCode ASC
";
                    dynamicParameters.Add("UserId", UserId);
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

        #region //IOS出貨計畫查詢
        public string GetDeliveriOS(string Company, int CompanyId, string DeliveryStatus, string SoIds, string SoErpFullNo, string CustomerMtlItemNo, int CustomerId
            , string MtlItemNo, int SalesmenId, string StartDate, string EndDate
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {

                List<string> SoErp = new List<string>();

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //判斷公司資料是否正確
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 CompanyId, CompanyName
                                FROM BAS.Company
                                WHERE CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CompanyId);

                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0) throw new SystemException("【公司】資料錯誤!");

                    foreach (var item in result)
                    {
                        ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                    }
                    #endregion
                }

                if (DeliveryStatus == "tracked")
                {
                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        List<Erp> erps = new List<Erp>();

                        #region //撈取暫出數量
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.TD001 SoErpPrefixNo, a.TD002 SoErpNo, a.TD003 SoSequence
                                FROM COPTD a
                                WHERE a.TD016 = 'N'
                                AND a.TD021 = 'Y'
                                AND a.TD008 - a.TD009 > 0";
                        erps = sqlConnection.Query<Erp>(sql, dynamicParameters).ToList();

                        SoErp = erps.Select(x => x.SoErpPrefixNo + '-' + x.SoErpNo + '-' + x.SoSequence).ToList();
                        #endregion
                    }
                }

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT TOP 1500
a.SoDetailId, 
a.SoId, 
a.SoSequence, 
a.MtlItemId, 
a.InventoryId, 
a.UomId, 
a.SoQty, 
a.SiQty,
a.SoDetailRemark
, FORMAT(a.PcPromiseDate, 'yyyy-MM-dd') PcPromiseDate
, FORMAT(a.PcPromiseDate, 'HH:mm:ss') PcPromiseTime
, b.SoErpPrefix,
b.SoErpNo, 
b.SoErpPrefix + '-' + b.SoErpNo SoErpFullNo, 
b.CustomerId,
c.CustomerNo, 
c.CustomerName, 
c.CustomerShortName, 
c.CustomerNo + ' ' + c.CustomerShortName CustomerWithNo
, ISNULL(d.MtlItemNo, '') MtlItemNo
, ISNULL(d.MtlItemName, '') MtlItemName
, ISNULL(d.MtlItemSpec, '') MtlItemSpec
, ISNULL(e.ItemQty, 0) PickQty
,(
	SELECT c.*
	FROM SCM.SoDetail a1
	INNER JOIN SCM.SaleOrder b1 ON b1.SoId = a1.SoId
	OUTER APPLY(
			SELECT (x1.WoErpPrefix +'-'+ x1.WoErpNo) WipOrder,
			y1.MtlItemNo,
			y1.MtlItemSpec,
			z1.CustomerMtlItemNo
			,ISNULL(d3.ProcessAlias,'未投入') CurrentProcess, 
			SUM (ISNULL(d2.BarcodeQty,0)) BarcodeQty
			FROM MES.WipOrder x1
			INNER JOIN PDM.MtlItem y1 ON y1.MtlItemId = x1.MtlItemId
			INNER JOIN PDM.CustomerMtlItem z1 ON z1.MtlItemId = x1.MtlItemId AND z1.CustomerId = b.CustomerId
			LEFT JOIN MES.ManufactureOrder d1 ON x1.WoId = d1.WoId
			LEFT JOIN MES.Barcode d2 ON d1.MoId = d2.MoId
			LEFT JOIN MES.MoProcess d3 ON d2.CurrentMoProcessId = d3.MoProcessId
			LEFT JOIN MES.MoProcess d4 ON d2.NextMoProcessId = d4.MoProcessId
			WHERE x1.MtlItemId = a.MtlItemId
			AND x1.WoStatus NOT IN ('4', '5')
			AND x1.ConfirmStatus = 'Y'
			GROUP BY ISNULL(d3.ProcessAlias,'未投入'),x1.WoErpPrefix,x1.WoErpNo,y1.MtlItemNo,y1.MtlItemSpec,z1.CustomerMtlItemNo
			
				) c
WHERE c.WipOrder IS NOT NULL
AND a1.SoDetailId = a.SoDetailId
order by c.WipOrder ASC
FOR JSON PATH, ROOT('data')
) MESOrderDetail 
FROM SCM.SoDetail a
INNER JOIN SCM.SaleOrder b ON b.SoId = a.SoId
INNER JOIN SCM.Customer c ON c.CustomerId = b.CustomerId
LEFT JOIN PDM.MtlItem d ON a.MtlItemId = d.MtlItemId
AND b.CompanyId = 1
AND a.ConfirmStatus = 'Y'
AND a.ClosureStatus = 'N'
OUTER APPLY(
		SELECT x.SoDetailId, SUM(x.ItemQty) ItemQty
		FROM SCM.PickingItem x
		WHERE x.SoDetailId = a.SoDetailId
		GROUP BY x.SoDetailId 
			) e
order by a.SoDetailId DESC
";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND b.CompanyId = @CompanyId 
                                              AND a.ConfirmStatus = 'Y'";
                    switch (DeliveryStatus)
                    {
                        case "tracked":
                            queryCondition += @" AND a.ClosureStatus = 'N'";
                            BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SoErp", @" AND (b.SoErpPrefix + '-' + b.SoErpNo + '-' + a.SoSequence) IN @SoErp", SoErp.ToArray());
                            break;
                        case "triTrade":
                            queryCondition += @"AND d.MtlItemNo LIKE '5%'";
                            break;
                        default:
                            break;
                    }
                    if (SoIds.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SoIds", @" AND a.SoId IN @SoIds", SoIds.Split(','));
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerId", @" AND b.CustomerId = @CustomerId", CustomerId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemNo", @" AND (d.MtlItemNo LIKE '%' + @MtlItemNo + '%' OR d.MtlItemName LIKE '%' + @MtlItemNo + '%' OR a.CustomerMtlItemNo LIKE '%' + @MtlItemNo + '%')", MtlItemNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SalesmenId", @" AND b.SalesmenId = @SalesmenId", SalesmenId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.PcPromiseDate >= @StartDate ", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.PcPromiseDate <= @EndDate ", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.SoId DESC, a.SoSequence";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;
                    var result = sqlConnection.Query(sql, dynamicParameters);
                    //var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);

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

        #region //取得客戶資料
        public string GetCustomeriOS(int CustomerId, string CustomerNo, string CustomerName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.CustomerId, a.CustomerNo, a.CustomerName, a.CustomerEnglishName, a.CustomerShortName
                          , a.RelatedPerson, FORMAT(a.PermitDate, 'yyyy-MM-dd') PermitDate, a.Version
                          , a.ResponsiblePerson, a.Contact, a.TelNoFirst, a.TelNoSecond, a.FaxNo, a.Email
                          , a.GuiNumber, a.Capital, a.AnnualTurnover, a.Headcount, a.HomeOffice, a.Currency
                          , a.DepartmentId, a.CustomerKind, a.SalesmenId, a.PaymentSalesmenId
                          , FORMAT(a.InauguateDate, 'yyyy-MM-dd') InauguateDate, FORMAT(a.CloseDate, 'yyyy-MM-dd') CloseDate
                          , a.ZipCodeRegister, a.RegisterAddressFirst, a.RegisterAddressSecond
                          , a.ZipCodeInvoice, a.InvoiceAddressFirst, a.InvoiceAddressSecond
                          , a.ZipCodeDelivery, a.DeliveryAddressFirst, a.DeliveryAddressSecond
                          , a.ZipCodeDocument, a.DocumentAddressFirst, a.DocumentAddressSecond
                          , a.BillReceipient, a.ZipCodeBill, a.BillAddressFirst, a.BillAddressSecond
                          , a.InvocieAttachedStatus, a.DepositRate, a.TaxAmountCalculateType, a.SaleRating, a.CreditRating
                          , a.TradeTerm, a.PaymentTerm, a.PricingType, a.ClearanceType, a.DocumentDeliver
                          , a.ReceiptReceive, a.PaymentType, a.TaxNo, a.InvoiceCount, a.Taxation
                          , a.Country, a.Region, a.Route, a.UploadType, a.PaymentBankFirst, a.BankAccountFirst
                          , a.PaymentBankSecond, a.BankAccountSecond, a.PaymentBankThird, a.BankAccountThird
                          , a.Account, a.AccountInvoice, a.AccountDay, a.ShipMethod, a.ShipType, a.ForwarderId
                          , a.CustomerRemark, a.CreditLimit, a.CreditLimitControl, a.CreditLimitControlCurrency
                          , a.SoCreditAuditType, a.SiCreditAuditType, a.DoCreditAuditType, a.InTransitCreditAuditType
                          , a.TransferStatus, a.TransferDate, a.Status
                          , a.CustomerNo + ' ' + a.CustomerShortName CustomerWithNo
                          , ISNULL(b.UserNo, '') SalesmenNo, ISNULL(b.UserName ,'') SalesmenName 
						  FROM SCM.Customer a
                          LEFT JOIN BAS.[User] b ON b.UserId = a.SalesmenId
";

                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerId", @" AND a.CustomerId = @CustomerId", CustomerId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerNo", @" AND a.CustomerNo LIKE '%' + @CustomerNo + '%'", CustomerNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerName", @" AND a.CustomerName LIKE '%' + @CustomerName + '%'", CustomerName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.CustomerNo";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    var result = sqlConnection.Query(sql, dynamicParameters);
                    //var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);

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

        #region //使用者列表
        public string GetUserList()
        {
            try
            {
                List<UserList> userList = new List<UserList>();
                List<MESUserList> mesUserList = new List<MESUserList>();
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.UserId, a.UserNo, a.UserName,a.Gender,a.Email,b.DepartmentName,a.Job,a.JobType, c.CompanyId, c.CompanyName, a.Status
FROM BAS.[User] a
INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
INNER JOIN BAS.Company c ON b.CompanyId = c.CompanyId
WHERE a.Status = 'A' and (a.UserNo LIKE 'DC%' OR a.UserNo LIKE 'ZY%' /*OR a.UserNo LIKE 'E%'*/ )
ORDER BY a.UserNo ASC
";


                    //var result = sqlConnection.Query(sql, dynamicParameters);
                    userList = sqlConnection.Query<UserList>(sql, dynamicParameters).ToList();



                }

                using (SqlConnection sqlConnection = new SqlConnection(MESEtergeConnectionStrings))
                {

                    var sql2 = "";
                    sql2 = @"SELECT USER_ID,BPM_USERNO,BPM_USERNAME FROM BAS.BPM_USER";
                    mesUserList = sqlConnection.Query<MESUserList>(sql2, dynamicParameters).ToList();

                    userList = userList.GroupJoin(mesUserList, x => x.UserNo, y => y.BPM_USERNO, (x, y) => { x.UserId = y.FirstOrDefault()?.USER_ID ?? x.UserId; return x; }).ToList();

                }

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    data = userList
                });
                #endregion
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

        #region //取得公司代碼
        public string GetUserCompany(int UserId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.UserId, a.UserName, c.CompanyId, c.CompanyName ,c.CompanyNo
                            FROM BAS.[User] a
                            INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                            INNER JOIN BAS.Company c ON b.CompanyId = c.CompanyId
                            WHERE a.Status = 'A'
                            AND b.Status = 'A'
                            AND c.Status = 'A'
                            AND a.UserId = @UserId
";
                    dynamicParameters.Add("UserId", UserId);
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

        #region //取得Token
        public string GetUserToken(int UserId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT * FROM BAS.UserLoginKey where UserId = @UserId
";
                    dynamicParameters.Add("UserId", UserId);
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

        #region //DeleteUserLoginKey -- 使用者登入金鑰刪除
        public string DeleteUserLoginKey(string UserNo/*, string KeyText*/)
        {
            try
            {
                if (UserNo.Length <= 0) throw new SystemException("【使用者編號】錯誤!");
                //if (KeyText.Length <= 0) throw new SystemException("【金鑰】錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷使用者登入金鑰是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.UserLoginKey a
                                INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                                WHERE b.UserNo = @UserNo
                                /*AND a.KeyText = @KeyText*/";
                        dynamicParameters.Add("UserNo", UserNo);
                        //dynamicParameters.Add("KeyText", KeyText);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("使用者登入金鑰資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM BAS.UserLoginKey a
                                INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                                WHERE b.UserNo = @UserNo
                                /*AND a.KeyText = @KeyText*/";
                        dynamicParameters.Add("UserNo", UserNo);
                        //dynamicParameters.Add("KeyText", KeyText);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            data = rowsAffected
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


        #region //Line Notify GPAI 20230122

        #region GetUserNotifyAuthorize //取得用戶Code Get https://notify-bot.line.me/oauth/authorize
        public string GetUserNotifyAuthorize(/*string clientId, string redirecturi*/)
        {
            try
            {
                //if (MoId <= 0) throw new SystemException("【製令】不能為空!");
                //if (MoProcessId <= 0) throw new SystemException("【製程】不能為空!");
                string clientId = "";
                string redirecturi = "";
                string resultURL = "";

                

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //組URL
                    string postUrl = "https://notify-bot.line.me/oauth/authorize?response_type=code&scope=notify&state=Gpai&client_id=";
                    clientId = linePushNotifiHelper.ClientID;
                    redirecturi = linePushNotifiHelper.callbackURL;

                    //var linePushNotifiHelperResult = linePushNotifiHelper.LineLogin();
                    //JObject linePushNotifiHelperResultJson = JObject.Parse(linePushNotifiHelperResult);

                    //if (linePushNotifiHelperResultJson["status"].ToString() == "success")
                    //{
                    //    foreach (var item in linePushNotifiHelperResultJson["result"])
                    //    {
                    //        clientId = item["client_id"].ToString();
                    //        redirecturi = item["redirect_uri"].ToString();

                    //    }

                    //    //resultURL = postUrl + clientId + "&redirect_uri=" + redirecturi;
                    //}

                    resultURL = postUrl + clientId + "&redirect_uri=" + redirecturi;

                    //JObject postDataJson = new JObject();
                    //postDataJson = JObject.FromObject(new
                    //{
                    //    response_type = "code",
                    //    client_id = clientId,
                    //    redirect_uri = redirecturi,
                    //    scope = "notify",
                    //    state = "Gpai",
                    //    response_mode = "form_post"
                    //});
                    //string postData = postDataJson.ToString();
                    //UTF8Encoding uTF8 = new UTF8Encoding();
                    //var result = BaseHelper.PostWebRequest(postUrl, postData, uTF8);
                    //var resultJson = JToken.Parse(result);
                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = resultURL
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

        #region //GetLineNotifyToken 取得用戶Token https://notify-bot.line.me/oauth/token
        public string GetLineNotifyToken(string usercode)
        {
            try
            {
                //if (MoId <= 0) throw new SystemException("【製令】不能為空!");
                //if (MoProcessId <= 0) throw new SystemException("【製程】不能為空!");
                string clientId = "";
                string callbackURL = "";

                string uuurr = "";

                if (usercode == "") throw new SystemException("【usercode】不能為空!");

                clientId = linePushNotifiHelper.ClientID;
                callbackURL = linePushNotifiHelper.callbackURL;

                string SecretKey = linePushNotifiHelper.SecretKey; 

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //取得數據
                    var linePushNotifiHelperResult = linePushNotifiHelper.GetLineNotifyToken(usercode);
                    JObject linePushNotifiHelperResultJson = JObject.Parse(linePushNotifiHelperResult);

                    if (linePushNotifiHelperResultJson["status"].ToString() == "200")
                    {
                        uuurr = linePushNotifiHelperResultJson["access_token"].ToString();

                        //resultURL = postUrl + clientId + "&redirect_uri=" + redirecturi;
                    }
                    else {
                        uuurr = linePushNotifiHelperResultJson.ToString();
                    }
                    //string postUrl = "https://notify-bot.line.me/oauth/token" + uuurr;
                    //string postUrl = "https://notify-bot.line.me/oauth/token";
                    //JObject postDataJson = new JObject();
                    //postDataJson = JObject.FromObject(new
                    //{
                    //    grant_type = "authorization_code",
                    //    code = usercode,
                    //    redirect_uri = callbackURL,
                    //    client_id = clientId,
                    //    client_secret = SecretKey
                    //});
                    //string postData = postDataJson.ToString();
                    //UnicodeEncoding uTF8 = new UnicodeEncoding();
                    //var result = BaseHelper.PostWebRequest(postUrl, postData, uTF8);
                    //var resultJson = JToken.Parse(result);
                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = uuurr
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

        #region //SendUserNotifyMessage 發送訊息給用戶 https://notify-api.line.me/api/notify token
        public string SendUserNotifyMessage(string message)
        {
            try
            {
                //if (MoId <= 0) throw new SystemException("【製令】不能為空!");
                //if (MoProcessId <= 0) throw new SystemException("【製程】不能為空!");
                //if (MeasurementNo == "") throw new SystemException("【機台】不能為空!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //取得數據
                    string postUrl = "https://notify-api.line.me/api/notify";
                    JObject postDataJson = new JObject();
                    postDataJson = JObject.FromObject(new
                    {
                        message = "TESTTTTTTTTTT",
                        
                    });
                    string postData = postDataJson.ToString();
                    UTF8Encoding uTF8 = new UTF8Encoding();
                    var result = BaseHelper.PostWebRequest(postUrl, postData, uTF8);
                    var resultJson = JToken.Parse(result);
                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = resultJson
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

        #region //CancelUserNotifyAuthorize 取消用戶訂閱通知 https://notify-api.line.me/api/revoke token
        public string CancelUserNotifyAuthorize(int UserId)
        {
            try
            {
                //if (MoId <= 0) throw new SystemException("【製令】不能為空!");
                //if (MoProcessId <= 0) throw new SystemException("【製程】不能為空!");
                //if (MeasurementNo == "") throw new SystemException("【機台】不能為空!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //取得數據
                    string postUrl = "https://notify-api.line.me/api/revoke";
                    JObject postDataJson = new JObject();
                    postDataJson = JObject.FromObject(new
                    {
                        
                    });
                    string postData = postDataJson.ToString();
                    UTF8Encoding uTF8 = new UTF8Encoding();
                    var result = BaseHelper.PostWebRequest(postUrl, postData, uTF8);
                    var resultJson = JToken.Parse(result);
                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = resultJson
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

    }
}