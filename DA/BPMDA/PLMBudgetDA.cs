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
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Transactions;

namespace BPMDA
{
    public class PLMBudgetDA
    {
        public string MainConnectionStrings = "";
        public string ErpConnectionStrings = "";
        public string BpmDbConnectionStrings = "";
        public string BpmServerPath = "";
        public string BpmAccount = "";
        public string BpmPassword = "";
        public DateTime CreateDate = default(DateTime);
        public DateTime LastModifiedDate = default(DateTime);

        public string sql = "";
        public JObject jsonResponse = new JObject();
        public Logger logger = LogManager.GetCurrentClassLogger();
        public SqlQuery sqlQuery = new SqlQuery();
        public DynamicParameters dynamicParameters = new DynamicParameters();
        public BpmHelper bpmHelper = new BpmHelper();

        public PLMBudgetDA() {

            MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];
            ErpConnectionStrings = ConfigurationManager.AppSettings["ErpSysDb"];
            BpmDbConnectionStrings = ConfigurationManager.AppSettings["BpmDb"];
            BpmServerPath = ConfigurationManager.AppSettings["BpmServerPath"];
            BpmAccount = ConfigurationManager.AppSettings["BpmAccount"];
            BpmPassword = ConfigurationManager.AppSettings["BpmPassword"];

            CreateDate = DateTime.Now;
            LastModifiedDate = DateTime.Now;

        }

        #region//Update
        #region//UpdatePLMBudgetTransferBpm  --拋轉專案預算
        public string UpdatePLMBudgetTransferBpm(string CompanyNo, string UserNo,string PLMBudget, int ProjectType) {
            try {
                string token = "";
                int CompanyId = -1,UserId=-1, rowsAffected = 0;
                using (TransactionScope transactionScope = new TransactionScope()) {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //取得USER資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserNo,a.UserId
                                FROM BAS.[User] a
                                WHERE a.UserNo = @UserNo";
                        dynamicParameters.Add("UserNo", UserNo);

                        var UserResult = sqlConnection.Query(sql, dynamicParameters);


                        foreach (var item in UserResult)
                        {
                            UserNo = item.UserNo;
                            UserId = item.UserId;
                        }
                        #endregion

                        #region //確認公司別DB
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CompanyId,a.CompanyNo, a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        dynamicParameters.Add("CompanyNo", CompanyNo);

                        var companyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (companyResult.Count() <= 0) throw new SystemException("公司別錯誤!");

                        foreach (var item in companyResult)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            CompanyId = item.CompanyId;
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //取得BPM TOKEN
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.IpAddress, a.CompanyId, a.Token, FORMAT(a.VerifyDate, 'yyyy-MM-dd HH:mm:ss') VerifyDate
                                    FROM BPM.SystemToken a
                                    WHERE IpAddress = @IpAddress
                                    AND CompanyId = @CompanyId";
                            dynamicParameters.Add("IpAddress", BpmServerPath);
                            dynamicParameters.Add("CompanyId", CompanyId);

                            var systemTokenResult = sqlConnection.Query(sql, dynamicParameters);
                            if (systemTokenResult.Count() <= 0) throw new SystemException("查無此憑證!");

                            foreach (var item in systemTokenResult)
                            {
                                DateTime verifyDate = Convert.ToDateTime(item.VerifyDate);
                                DateTime nowDate = DateTime.Now;
                                var CheckMin = (nowDate - verifyDate).TotalMinutes;
                                if (CheckMin >= 30)
                                {
                                    #region //取得新BPM TOKEN
                                    string tokenResponse = BpmHelper.GetBpmToken(BpmServerPath, BpmAccount, BpmPassword);
                                    var tokenJson = JObject.Parse(tokenResponse);
                                    foreach (var item2 in tokenJson)
                                    {
                                        if (item2.Key == "status")
                                        {
                                            if (item2.Value.ToString() != "success") throw new SystemException("取得token失敗!");
                                        }
                                        else if (item2.Key == "data")
                                        {
                                            token = item2.Value.ToString();
                                        }
                                    }
                                    #endregion

                                    #region //將新的TOKEN更新回MES
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE BPM.SystemToken SET
                                            Token = @Token,
                                            VerifyDate = @VerifyDate,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE IpAddress = @IpAddress
                                            AND CompanyId = @CompanyId";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            Token = token,
                                            VerifyDate = nowDate,
                                            LastModifiedDate,
                                            LastModifiedBy = UserId,
                                            IpAddress = BpmServerPath,
                                            CompanyId
                                        });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                                else
                                {
                                    token = item.Token;
                                }
                            }
                            #endregion

                            #region //取得BpmUser資料
                            string BpmUserId = "";
                            string BpmRoleId = "";
                            string BpmUserNo = "";
                            string UserName = "";
                            string BpmDepNo = "";
                            string BpmDepName = "";
                            using (SqlConnection sqlConnection3 = new SqlConnection(BpmDbConnectionStrings))
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"WITH BasicUserInfo(MemID, LoginID, UserName, MainRoleID, RolID, RolName, ParentRol) AS(
                                        SELECT a.MemID, a.LoginID, a.UserName, a.MainRoleID
                                        , b.RolID, b.Name, b.DepID AS ParentRol
                                        FROM Mem_GenInf a
                                        LEFT JOIN Rol_GenInf b ON b.RolID = a.MainRoleID
                                        WHERE a.LoginID = @LoginID
                                        UNION ALL
                                        SELECT a.MemID, a.LoginID, a.UserName, a.MainRoleID
                                        , b.RolID, b.Name, b.DepID AS ParentRol
                                        FROM BasicUserInfo a, Rol_GenInf b
                                        WHERE a.ParentRol = b.RolID
                                        )
                                        SELECT a.MemID, a.LoginID, a.UserName, a.MainRoleID
                                        , b.Name AS RolName
                                        , c.DepID, c.ID AS DepNo, c.Name AS DepName
                                        FROM BasicUserInfo a
                                        LEFT JOIN Rol_GenInf AS parentRol_GenInf ON a.RolID = parentRol_GenInf.RolID
                                        LEFT JOIN Dep_GenInf c ON parentRol_GenInf.DepID = c.DepID
                                        , Rol_GenInf b
                                        WHERE a.MainRoleID = b.RolID
                                        AND c.DepID IS NOT NULL
                                        UNION
                                        SELECT a.MemID, a.LoginID, a.UserName, a.MainRoleID
                                        , b.Name AS RolName
                                        , c.ComID AS DepID, c.ID AS DepNo, c.Name AS DepName
                                        FROM Mem_GenInf a
                                        LEFT JOIN Rol_GenInf b ON a.MainRoleID = b.RolID
                                        LEFT JOIN Company c ON b.DepID = c.ComID
                                        WHERE c.ComID IS NOT NULL
                                        AND a.LoginID = @LoginID
                                        ORDER BY a.LoginID";
                                dynamicParameters.Add("LoginID", UserNo);

                                var MemGenInfResult = sqlConnection3.Query(sql, dynamicParameters);

                                if (MemGenInfResult.Count() <= 0) throw new SystemException("取得BPM使用者資訊時發生錯誤!!");

                                foreach (var item in MemGenInfResult)
                                {
                                    BpmUserId = item.MemID;
                                    BpmRoleId = item.MainRoleID;
                                    BpmUserNo = item.LoginID;
                                    UserName = item.UserName;
                                    BpmDepNo = item.DepNo;
                                    BpmDepName = item.DepName;
                                }
                            }
                            #endregion

                            #region //依公司別取得ProId
                            string proId = "";
                            if (ProjectType==1) {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.TypeName ProId
                                    FROM BAS.[Type] a
                                    WHERE a.TypeSchema = 'BPM.PrProId'
                                    AND a.TypeNo = @CompanyNo";
                                dynamicParameters.Add("CompanyNo", CompanyNo);

                                var ProIdResult = sqlConnection.Query(sql, dynamicParameters);

                                if (ProIdResult.Count() <= 0) throw new SystemException("此公司別【" + CompanyNo + "】尚未建立拋轉起單碼，請聯繫資訊人員!!");

                                
                                foreach (var item in ProIdResult)
                                {
                                    proId = item.ProId;
                                }
                                proId = "PRO23221747028234464";
                            }
                            else {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.TypeName ProId
                                    FROM BAS.[Type] a
                                    WHERE a.TypeSchema = 'BPM.PrProId'
                                    AND a.TypeNo = @CompanyNo";
                                dynamicParameters.Add("CompanyNo", CompanyNo);

                                var ProIdResult = sqlConnection.Query(sql, dynamicParameters);

                                if (ProIdResult.Count() <= 0) throw new SystemException("此公司別【" + CompanyNo + "】尚未建立拋轉起單碼，請聯繫資訊人員!!");


                                foreach (var item in ProIdResult)
                                {
                                    proId = item.ProId;
                                }
                                proId = "PRO23551747964031518";
                            }

                            #endregion

                            string memId = BpmUserId;
                            string rolId = BpmRoleId;
                            string startMethod = "NoOpFirst";
                            JObject plmBudgetObject = null;
                            JObject artInsAppData = null;

                            // 驗證 PLMBudget 是否為有效的 JSON
                            if (string.IsNullOrEmpty(PLMBudget))
                            {
                                throw new ArgumentException("PLMBudget 不能為空");
                            }

                            try
                            {
                                plmBudgetObject = JObject.Parse(PLMBudget);
                            }
                            catch (JsonException)
                            {
                                throw new SystemException("PLMBudget 不是有效的 JSON 格式");
                            }

                            if (plmBudgetObject["ApprovalUrl"] == null)
                            {
                                throw new SystemException("PLMBudget 缺少必要的 ApprovalUrl 欄位");
                            }

                            // 建立 artInsAppData（物件格式）
                            artInsAppData = JObject.FromObject(new
                            {
                                PLMBudget 
                            });                            
                            
                            string sData = BpmHelper.PostFormToBpm(token, proId, memId, rolId, startMethod, artInsAppData, BpmServerPath);


                            if (sData != "true")
                            {
                                // 解析JSON數據
                                var responseJson = JObject.Parse(PLMBudget);
                                // 取出回傳資訊                            
                                string approvalUrl = responseJson["ApprovalUrl"]?.Value<string>() ?? "";
                                string budgetId = responseJson["BudgetId"]?.Value<string>() ?? "";
                                string changeId = responseJson["ChangeId"] is JArray arr ? string.Join(",", arr) : responseJson["ChangeId"]?.Value<string>() ?? "";
                                string approvedBy = UserNo;
                                string approvalStatus = "F"; //F 審核失敗
                                string approvalComments = "審核失敗";
                                // 呼叫審核API
                                string approvalResult = CallApprovalApiSync(approvalUrl, changeId, budgetId, approvedBy, approvalStatus, approvalComments);
                            }
                            else
                            {
                                #region //Response
                                jsonResponse = JObject.FromObject(new
                                {
                                    status = "success",
                                    msg = "(" + rowsAffected + " rows affected)"
                                });
                                #endregion
                            }

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

        // 新增的同步方法：使用HttpClient呼叫審核API（同步版本）
        private string CallApprovalApiSync(string approvalUrl,string changeId, string budgetId, string approvedBy, string approvalStatus, string approvalComments)
        {
            try
            {
                if (budgetId!="") {
                    #region //budgetId
                    using (var httpClient = new HttpClient())
                    {
                        // 設定請求標頭
                        httpClient.DefaultRequestHeaders.Add("Accept", "application/json");

                        // 準備POST資料
                        var postData = new
                        {
                            BudgetId = budgetId,
                            ApprovedBy = approvedBy,
                            ApprovalStatus = approvalStatus,
                            ApprovalComments = approvalComments
                        };

                        // 序列化為JSON
                        string jsonContent = JsonConvert.SerializeObject(postData);
                        var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                        // 發送POST請求（同步方式）
                        HttpResponseMessage response = httpClient.PostAsync(approvalUrl, content).Result;

                        // 檢查回應狀態
                        if (response.IsSuccessStatusCode)
                        {
                            string responseContent = response.Content.ReadAsStringAsync().Result;
                            logger.Info($"審核API呼叫成功: BudgetId={budgetId}, Status={approvalStatus}, Response={responseContent}");
                            return $"審核API呼叫成功: {responseContent}";
                        }
                        else
                        {
                            string errorContent = response.Content.ReadAsStringAsync().Result;
                            string errorMsg = $"審核API呼叫失敗: StatusCode={response.StatusCode}, Error={errorContent}";
                            logger.Error(errorMsg);
                            return errorMsg;
                        }
                    }
                    #endregion
                }
                else {
                    #region //changeId
                    using (var httpClient = new HttpClient())
                    {
                        // 設定請求標頭
                        httpClient.DefaultRequestHeaders.Add("Accept", "application/json");

                        // 準備POST資料
                        var postData = new
                        {
                            ChangeId = changeId,
                            ApprovedBy = approvedBy,
                            ApprovalStatus = approvalStatus,
                            ApprovalComments = approvalComments
                        };

                        // 序列化為JSON
                        string jsonContent = JsonConvert.SerializeObject(postData);
                        var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                        // 發送POST請求（同步方式）
                        HttpResponseMessage response = httpClient.PostAsync(approvalUrl, content).Result;

                        // 檢查回應狀態
                        if (response.IsSuccessStatusCode)
                        {
                            string responseContent = response.Content.ReadAsStringAsync().Result;
                            logger.Info($"審核API呼叫成功: BudgetId={budgetId}, Status={approvalStatus}, Response={responseContent}");
                            return $"審核API呼叫成功: {responseContent}";
                        }
                        else
                        {
                            string errorContent = response.Content.ReadAsStringAsync().Result;
                            string errorMsg = $"審核API呼叫失敗: StatusCode={response.StatusCode}, Error={errorContent}";
                            logger.Error(errorMsg);
                            return errorMsg;
                        }
                    }
                    #endregion
                }
            }
            catch (HttpRequestException httpEx)
            {
                string errorMsg = $"HTTP請求錯誤: {httpEx.Message}";
                logger.Error(errorMsg);
                return errorMsg;
            }
            catch (AggregateException aggEx)
            {
                // 處理同步調用異步方法可能產生的AggregateException
                Exception innerEx = aggEx.GetBaseException();
                string errorMsg = $"請求處理錯誤: {innerEx.Message}";
                logger.Error(errorMsg);
                return errorMsg;
            }
            catch (Exception ex)
            {
                string errorMsg = $"審核API呼叫發生未預期錯誤: {ex.Message}";
                logger.Error(errorMsg);
                return errorMsg;
            }
        }        
        #endregion


        #endregion
    }
}
