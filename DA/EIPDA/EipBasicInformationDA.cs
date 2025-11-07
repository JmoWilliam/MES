using System;
using Dapper;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Text.RegularExpressions;
using System.Transactions;
using Helpers;
using NLog;
using System.Web;
using System.IO;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;

namespace EIPDA
{
    public class EipBasicInformationDA
    {
        public string MainConnectionStrings = string.Empty;
        public string OfficialConnectionStrings = string.Empty;

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

        public EipBasicInformationDA()
        {
            MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];
            OfficialConnectionStrings = ConfigurationManager.AppSettings["OfficialDb"];
            GetUserInfo();
            CreateDate = DateTime.Now;
            LastModifiedDate = DateTime.Now;

            dynamicParameters = new DynamicParameters();
            sqlQuery = new SqlQuery();
        }

        #region //GetUserInfo -- 取得使用者資訊 -- Chia Yuan 2023.7.7
        private void GetUserInfo()
        {
            try
            {
                CurrentUser = Convert.ToInt32(HttpContext.Current.Session["UserId"]);
                CreateBy = Convert.ToInt32(HttpContext.Current.Session["UserId"]);
                LastModifiedBy = Convert.ToInt32(HttpContext.Current.Session["UserId"]);
            }
            catch (Exception e)
            {
                logger.Error(e.Message);
            }
        }
        #endregion

        #region //GetServerInfo -- 取得伺服器資訊 -- Chia Yuan 2023.07.27
        public static string GetServerInfo()
        {
            string LocalIp = string.Empty;
            string Domain = System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties().DomainName;
            string Host = System.Net.Dns.GetHostName();
            if (!System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable())
                return null;

            if (!string.IsNullOrEmpty(HttpContext.Current.Request.ServerVariables["HTTP_X_FORWARDED_FOR"]))
            {
                return HttpContext.Current.Request.ServerVariables["HTTP_X_FORWARDED_FOR"];
            }
            else
            {
                return HttpContext.Current.Request.ServerVariables["REMOTE_ADDR"];
            }

            System.Net.IPHostEntry host = System.Net.Dns.GetHostEntry(System.Net.Dns.GetHostName());

            foreach (System.Net.IPAddress ip in host.AddressList)
            {
                if (ip.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                {
                    LocalIp = ip.ToString();
                    break;
                }

            }
            //return string.Format("[Domain-{0} : Host-{1} : IP-{2}]", Domain, Host, LocalIp);
            return LocalIp;
        }
        #endregion

        #region //Get
        #region //GetCustomer -- 取得客戶資料 -- Zoey 2022.06.15
        public string GetCustomer(int CustomerId, string CustomerNo, string CustomerName, string Status
            , string Account, string KeyText
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //取得使用者公司別
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 b.CompanyId 
                            FROM BAS.[User] a join BAS.Department b on b.DepartmentId = a.DepartmentId 
                            WHERE a.UserNo = @Account";
                    dynamicParameters.Add("Account", Account);
                    var userResult = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                    //if (userResult == null) throw new SystemException("【使用者】資料錯誤!");
                    int CompanyId = userResult == null ? -1 : userResult.CompanyId;
                    #endregion

                    sqlQuery.mainKey = "a.CustomerId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.CustomerNo, a.CustomerName, a.CustomerEnglishName, a.CustomerShortName
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
                          , ISNULL(b.UserNo, '') SalesmenNo, ISNULL(b.UserName ,'') SalesmenName";
                    sqlQuery.mainTables =
                        @"FROM SCM.Customer a
                          LEFT JOIN BAS.[User] b ON b.UserId = a.SalesmenId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CompanyId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerId", @" AND a.CustomerId = @CustomerId", CustomerId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerNo", @" AND a.CustomerNo LIKE '%' + @CustomerNo + '%'", CustomerNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerName", @" AND a.CustomerName LIKE '%' + @CustomerName + '%'", CustomerName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.CustomerNo";
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

        #region //GetLogin -- 取得登入者資訊 -- Chia Yuan 2023.7.7
        public string GetLogin(string Account, string Password)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(Account)) throw new SystemException("【信箱】不能為空!");
                if (string.IsNullOrWhiteSpace(Password)) throw new SystemException("【密碼】不能為空!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //密碼參數設定
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.PasswordExpiration, a.PasswordFormat, a.PasswordWrongCount
                            FROM BAS.PasswordSetting a";
                    var resultPasswordSetting = sqlConnection.Query(sql, dynamicParameters);
                    if (!resultPasswordSetting.Any()) throw new SystemException("【密碼參數設定】資料錯誤!");

                    int PasswordExpiration = -1;
                    string PasswordFormat = "";
                    int PasswordWrongCount = -1;
                    foreach (var item in resultPasswordSetting)
                    {
                        PasswordExpiration = Convert.ToInt32(item.PasswordExpiration);
                        PasswordFormat = item.PasswordFormat;
                        PasswordWrongCount = Convert.ToInt32(item.PasswordWrongCount);
                    }
                    #endregion

                    //MemberType 1=EIP User, 2=BM User
                    sql = @"SELECT a.MemberId as UserId,a.MemberName as UserName,a.MemberEmail as Email,a.Address,a.[Password],a.PasswordStatus,a.PasswordMistake,1 as MemberType,a.[Status] 
                            FROM EIP.[Member] a 
                            WHERE a.Status = @Status and a.MemberEmail = @Account";
                    dynamicParameters.Add("Status", "A");
                    dynamicParameters.Add("Account", Account);
                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (!result.Any()) throw new SystemException("查無此帳號資料：" + Account);

                    foreach (var item in result)
                    {
                        //if (Convert.ToInt32(item.PasswordMistake) >= PasswordWrongCount) throw new SystemException(string.Format("密碼錯誤已達{0}次，帳號鎖定!", PasswordWrongCount));

                        if (item.PasswordStatus != "Y")
                        {
                            if (!Regex.IsMatch(Password, PasswordFormat, RegexOptions.IgnoreCase)) throw new SystemException("密碼格式錯誤!");
                        }

                        if (item.Password != BaseHelper.Sha256Encrypt(Password))
                        {
                            UpdatePasswordMistake(Convert.ToInt32(item.MemberId));

                            throw new SystemException("登入密碼錯誤：" + Account);
                        }
                        else
                        {
                            UpdatePasswordMistakeReset(Convert.ToInt32(item.MemberId));
                        }
                    }

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

        #region //GetLoginByKey -- 取得登入者資訊(cookie) -- Chia Yuan 2023.08.01
        public string GetLoginByKey(string Account, string KeyText)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(Account)) throw new SystemException("【使用者】資料錯誤!");
                if (string.IsNullOrWhiteSpace(KeyText)) throw new SystemException("【金鑰】錯誤!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.KeyId, a.KeyText, b.MemberId, b.MemberEmail, b.MemberName
                            FROM EIP.MemberLoginKey a
                            INNER JOIN EIP.[Member] b on b.MemberId = a.MemberId
                            WHERE a.LoginIP = @LoginIP
                            AND a.KeyText = @KeyText
                            AND a.ExpirationDate > @ExpirationDate
                            AND b.[Status] = @Status
                            AND b.MemberEmail = @Account";
                    dynamicParameters.Add("Status", "A");
                    dynamicParameters.Add("Account", Account);
                    dynamicParameters.Add("LoginIP", BaseHelper.ClientIP());
                    dynamicParameters.Add("KeyText", KeyText);
                    dynamicParameters.Add("ExpirationDate", DateTime.Now);
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

        #region //GetLoginIPCheck -- 取得登入者最近IP -- Chia Yuan 2023.7.7
        public string GetLoginIPCheck(string KeyText)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(KeyText)) throw new SystemException("【使用者】資料錯誤!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT TOP 1 b.MemberEmail, a.LoginIP
                            FROM EIP.MemberLoginKey a
                            INNER JOIN EIP.[Member] b ON a.MemberId = b.MemberId
                            WHERE b.[Status] = @Status AND a.KeyText = @KeyText
                            ORDER BY a.CreateDate DESC";
                    dynamicParameters.Add("Status", "A");
                    dynamicParameters.Add("KeyText", KeyText);
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

        #region //GetMember -- 取得使用者資料 -- Chia Yuan -- 2023.07.27
        public string GetMember(int MemberId, string MemberName, string Status
            , int MemberType, string Account, string KeyText
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    string currentDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                    sqlQuery.mainKey = "a.MemberId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",a.MemberName,a.MemberEmail,a.OrgShortName,a.[Address],a.ContactName,a.ContactPhone,a.[Status]";
                    sqlQuery.mainTables =
                        @"FROM EIP.[Member] a 
                        INNER JOIN EIP.MemberLoginKey b on b.MemberId = a.MemberId
                            and b.KeyText = @KeyText and b.ExpirationDate >= @Now";
                    dynamicParameters.Add("KeyText", KeyText);
                    dynamicParameters.Add("Now", currentDate);
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MemberId", @" AND a.MemberId = @MemberId", MemberId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Account", @" AND a.MemberEmail = @Account", Account);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MemberName", @" AND a.MemberName LIKE '%' + @MemberName + '%'", MemberName);
                    //BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MemberEmail", @" AND a.MemberEmail LIKE '%' + @MemberEmail + '%'", MemberEmail);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.MemberId";
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

        #region //GetRequestForQuotation -- 取得詢價清單列表(客戶端 RFQ)資訊 -- Chia Yuan 2023.07.11
        public string GetRequestForQuotation(int RfqId, string RfqNo, string MemberName, string AssemblyName, int ProductUseId, int SalesId, int RfqProTypeId, string Status
            , int MemberType, string Account, string KeyText
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //取得客戶資料
                    int MemberId = -1, UserId = -1, CustomerId = -1;
                    if (MemberType == 1)
                    {
                        #region //判斷使用者登入金鑰是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 b.MemberId,ISNULL(c.OrganizaitonTypeId,-1) AS CustomerId
                                FROM EIP.MemberLoginKey a
                                INNER JOIN EIP.[Member] b ON a.MemberId = b.MemberId
                                LEFT JOIN EIP.MemberOrganization c on c.MemberId = b.MemberId and c.OrganizaitonType = @OrganizaitonType
                                WHERE b.MemberEmail = @Account
                                AND a.KeyText = @KeyText";
                        dynamicParameters.Add("Account", Account);
                        dynamicParameters.Add("KeyText", KeyText);
                        dynamicParameters.Add("OrganizaitonType", "1");
                        var userResult = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                        if (userResult == null) throw new SystemException("【使用者】資料錯誤!");
                        MemberId = userResult.MemberId;
                        CustomerId = userResult.CustomerId;
                        #endregion
                    }
                    if (MemberType == 2)
                    {
                        #region //判斷使用者登入金鑰是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 b.UserId
                                FROM BAS.UserLoginKey a
                                INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                                WHERE b.UserNo = @Account
                                AND a.KeyText = @KeyText";
                        dynamicParameters.Add("Account", Account);
                        dynamicParameters.Add("KeyText", KeyText);
                        var userResult = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                        if (userResult == null) throw new SystemException("【使用者】資料錯誤!");
                        UserId = userResult.UserId;
                        #endregion
                    }
                    #endregion

                    sqlQuery.mainKey = "a.RfqId";
                    sqlQuery.auxKey = "";
                    sqlQuery.distinct = true;
                    sqlQuery.columns =
                        @", a.RfqNo, a.AssemblyName, a.ProductUseId, a.MemberId, a.UserId, FORMAT(a.CreateDate, 'yyyy-MM-dd') AS CreateDate, a.LastModifiedDate, a.[Status]
                            , b.ProductUseName, d.StatusName, e.TypeName
                            , g.*, ISNULL(a.CustomerId,f.OrganizaitonTypeId) as CustomerId,case when ISNULL(a.CustomerName,'') = '' then h.CustomerName else a.CustomerName end as CustomerName
                            , (
		                        SELECT aa.RfqId, aa.RfqDetailId, aa.RfqSequence, ab.RfqProClassId, aj.RfqProductClassName, aa.RfqProTypeId, ab.RfqProductTypeName, aa.MtlName
		                        , aa.CustProdDigram, aa.PlannedOpeningDate
		                        , aa.PrototypeQty, ac.TypeName AS ProtoScheduleName
		                        , aa.MassProductionDemand, ah.StatusName AS MassProductionDemandName
		                        , aa.KickOffType, aa.PlasticName, aa.OutsideDiameter
		                        , aa.ProdLifeCycleStart, aa.ProdLifeCycleEnd, aa.LifeCycleQty, aa.DemandDate
		                        , aa.CoatingFlag, ai.StatusName AS CoatingFlagName
		                        , ISNULL(aa.SalesId, -1) SalesId, ISNULL(ad.UserName, '') AS SalesName
		                        , ISNULL((ag.CompanyName + '-' + ad.UserName), '') AS SalesInfo
		                        , aa.AdditionalFile, aa.QuotationFile, aa.[Status], ae.StatusName AS RfqDetailStatusName
		                        , ISNULL(FORMAT(aa.ConfirmSalesTime, 'yyyy-MM-dd HH:mm:ss'), '') AS ConfirmSalesTime
		                        , ISNULL(FORMAT(aa.ConfirmRdTime, 'yyyy-MM-dd HH:mm:ss'), '') AS ConfirmRdTime, aa.[Description]
	                            , (
		                            SELECT aaa.RfqPkId, aaa.RfqPkTypeId, aaa.[Status], aac.StatusName, aab.PackagingMethod FROM SCM.RfqPackage aaa
		                            INNER JOIN SCM.RfqPackageType aab on aab.RfqPkTypeId = aaa.RfqPkTypeId
		                            INNER JOIN BAS.[Status] aac on aac.StatusNo = aaa.[Status] and aac.StatusSchema = 'Status'
		                            WHERE aaa.RfqDetailId = aa.RfqDetailId
		                            FOR JSON PATH, ROOT('data')
	                            ) AS RfqPackage
		                        , (
			                        SELECT aaa.RfqLineSolutionId, aaa.PeriodicDemandType, aaa.SolutionQty FROM SCM.RfqLineSolution aaa
			                        WHERE aaa.RfqDetailId = aa.RfqDetailId
			                        ORDER BY aaa.SortNumber
			                        FOR JSON PATH, ROOT('data')
		                        ) AS RfqLineSolution
                                FROM SCM.RfqDetail aa
		                        INNER JOIN SCM.RfqProductType ab ON ab.RfqProTypeId = aa.RfqProTypeId
		                        INNER JOIN SCM.RfqProductClass aj ON aj.RfqProClassId = ab.RfqProClassId
		                        LEFT JOIN BAS.[Type] ac ON ac.TypeNo = aa.ProtoSchedule AND ac.TypeSchema = 'RfqDetail.ProtoSchedule'
		                        LEFT JOIN BAS.[User] ad ON ad.UserId = aa.SalesId
		                        INNER JOIN BAS.[Status] ae ON ae.StatusNo = aa.[Status] AND ae.StatusSchema = 'RfqDetail.Status'
		                        LEFT JOIN BAS.Department af ON af.DepartmentId = ad.DepartmentId
		                        LEFT JOIN BAS.Company ag ON ag.CompanyId = af.CompanyId
		                        LEFT JOIN BAS.[Status] ah ON ah.StatusNo = aa.MassProductionDemand AND ah.StatusSchema = 'Boolean'
		                        INNER JOIN BAS.[Status] ai ON ai.StatusNo = aa.CoatingFlag AND ai.StatusSchema = 'Boolean'
                                WHERE aa.RfqId = a.RfqId
                                ORDER BY aa.RfqDetailId
                                FOR JSON PATH, ROOT('data')
                            ) AS RfqDetail";
                    sqlQuery.mainTables =
                        @"FROM SCM.RequestForQuotation a
                        INNER JOIN SCM.ProductUse b ON b.ProductUseId = a.ProductUseId
                        INNER JOIN BAS.[Status] d ON d.StatusNo = a.[Status] and d.StatusSchema = 'RequestForQuotation.Status'
                        INNER JOIN BAS.[Type] e ON e.TypeNo = a.MemberType and e.TypeSchema = 'RequestForQuotation.MemberType'
                        LEFT JOIN SCM.Customer c ON c.CustomerId = a.CustomerId
                        LEFT JOIN EIP.MemberOrganization f on f.MemberId = a.MemberId and f.OrganizaitonType = 1
                        LEFT JOIN SCM.Customer h on h.CustomerId = ISNULL(a.CustomerId,f.OrganizaitonTypeId)
                        LEFT JOIN (
	                        SELECT distinct a.MemberId as UserId,a.MemberEmail as UserNo, a.MemberName as UseeName, 1 as MemberType FROM EIP.[Member] a 
	                        UNION SELECT a.UserId, a.UserNo, a.UserName, 2 as MemberType FROM BAS.[User] a
                        ) g on ISNULL(a.MemberId,a.UserId) = g.UserId and g.MemberType = case when a.MemberId is null then 2 else 1 end";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    //string queryCondition = " and a.MemberType = @MemberType";
                    //dynamicParameters.Add("MemberType", MemberType);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RfqId", @" AND a.RfqId = @RfqId", RfqId);
                    if (CustomerId > 0)
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerId", @" AND ISNULL(a.CustomerId,f.OrganizaitonTypeId) = @CustomerId", CustomerId);
                    else
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MemberId", @" AND a.MemberId = @MemberId", MemberId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserId", @" AND a.UserId = @UserId", UserId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RfqNo", @" AND a.RfqNo = @RfqNo", RfqNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "AssemblyName", @" AND a.AssemblyName LIKE '%' + @AssemblyName + '%'", AssemblyName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProductUseId", @" AND a.ProductUseId = @ProductUseId", ProductUseId);
                    //BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SalesId", @" AND b.SalesId = @SalesId", SalesId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status = @Status", Status);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.RfqNo DESC";
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

        #region //GetRfqDetail -- 取得RFQ單身資訊 -- Chia Yuan 2023.07.21
        public string GetRfqDetail(int RfqId, int RfqDetailId
            , int MemberType, string Account, string KeyText
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //取得客戶資料
                    int MemberId = -1, UserId = -1, CustomerId = -1;
                    if (MemberType == 1)
                    {
                        #region //判斷使用者登入金鑰是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 b.MemberId,ISNULL(c.OrganizaitonTypeId,-1) AS CustomerId
                                FROM EIP.MemberLoginKey a
                                INNER JOIN EIP.[Member] b ON a.MemberId = b.MemberId
                                LEFT JOIN EIP.MemberOrganization c on c.MemberId = b.MemberId and c.OrganizaitonType = @OrganizaitonType
                                WHERE b.MemberEmail = @Account
                                AND a.KeyText = @KeyText";
                        dynamicParameters.Add("Account", Account);
                        dynamicParameters.Add("KeyText", KeyText);
                        dynamicParameters.Add("OrganizaitonType", "1");
                        var userResult = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                        if (userResult == null) throw new SystemException("【使用者】資料錯誤!");
                        MemberId = userResult.MemberId;
                        CustomerId = userResult.CustomerId;
                        #endregion
                    }
                    if (MemberType == 2)
                    {
                        #region //判斷使用者登入金鑰是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 b.UserId
                                FROM BAS.UserLoginKey a
                                INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                                WHERE b.UserNo = @Account
                                AND a.KeyText = @KeyText";
                        dynamicParameters.Add("Account", Account);
                        dynamicParameters.Add("KeyText", KeyText);
                        var userResult = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                        if (userResult == null) throw new SystemException("【使用者】資料錯誤!");
                        UserId = userResult.UserId;
                        #endregion
                    }
                    #endregion

                    sqlQuery.mainKey = "a.RfqDetailId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.RfqId, a.CompanyId, a.RfqSequence
                        , a.RfqProTypeId, ISNULL(e.RfqProductTypeName, '') AS RfqProductTypeName
                        , a.MtlName, ISNULL(n.FileId, -1) AS CustProdDigram, n.FileName, n.FileExtension, ISNULL(FORMAT(a.PlannedOpeningDate, 'yyyy-MM-dd'), '') AS PlannedOpeningDate
                        , a.PrototypeQty, a.ProtoSchedule, ISNULL(f.TypeName, '') AS ProtoScheduleName
                        , a.MassProductionDemand, k.StatusName AS MassProductionDemandName
                        , ISNULL(a.KickOffType, '') AS KickOffType, ISNULL(a.PlasticName, '') AS PlasticName, ISNULL(a.OutsideDiameter, '') AS OutsideDiameter
                        , FORMAT(a.ProdLifeCycleStart, 'yyyy-MM-dd') AS ProdLifeCycleStart, FORMAT(a.ProdLifeCycleEnd, 'yyyy-MM-dd') AS ProdLifeCycleEnd
                        , a.LifeCycleQty, FORMAT(a.DemandDate, 'yyyy-MM-dd') AS DemandDate
                        , a.CoatingFlag, l.StatusName AS CoatingFlagName, a.PortOfDelivery, a.Status
                        , ISNULL(a.SalesId, -1) AS SalesId, g.UserName AS SalesName, ISNULL((j.CompanyName + '-' + g.UserName), '') AS SalesInfo, a.AdditionalFile, a.QuotationFile, c.ProductUseName, h.StatusName AS RfqDetailStatusName
                        , ISNULL(FORMAT(a.ConfirmSalesTime, 'yyyy-MM-dd HH:mm:ss'), '') AS ConfirmSalesTime, ISNULL(FORMAT(a.ConfirmRdTime, 'yyyy-MM-dd HH:mm:ss'), '') AS ConfirmRdTime, a.Description
                        , b.RfqNo, b.AssemblyName, b.ProductUseId, b.MemberId, b.Status AS RequestForQuotationStatus, b.MemberType, b.CustomerId, b.CustomerName
                        , m.RfqProClassId, ISNULL(m.RfqProductClassName, '') AS RfqProductClassName 
                        , (
                            SELECT l.RfqPkTypeId, l.PackagingMethod, k.SustSupplyStatus, m.StatusName SustSupplyStatusName, n.RfqProductClassName
                            FROM SCM.RfqPackage k
                            INNER JOIN SCM.RfqPackageType l ON l.RfqPkTypeId = k.RfqPkTypeId
                            INNER JOIN BAS.[Status] m ON m.StatusNo = k.SustSupplyStatus AND m.StatusSchema = 'Boolean'
                            INNER JOIN SCM.RfqProductClass n ON n.RfqProClassId = l.RfqProClassId
                            WHERE k.RfqDetailId = a.RfqDetailId
                            FOR JSON PATH, ROOT('data')
                        ) AS PackagingMethod
                        , (
	                        SELECT aa.RfqLineSolutionId, aa.PeriodicDemandType, bb.TypeName, aa.SolutionQty FROM SCM.RfqLineSolution aa
                            INNER JOIN BAS.[Type] bb on bb.TypeNo = aa.PeriodicDemandType AND bb.TypeSchema = 'RfqDetail.PeriodicDemandType'
	                        WHERE aa.RfqDetailId = a.RfqDetailId
	                        ORDER BY aa.SortNumber
	                        FOR JSON PATH, ROOT('data')
                        ) AS RfqLineSolution";
                    sqlQuery.mainTables =
                        @"FROM SCM.RfqDetail a
                        INNER JOIN SCM.RequestForQuotation b ON b.RfqId = a.RfqId
                        INNER JOIN SCM.ProductUse c ON c.ProductUseId = b.ProductUseId
                        LEFT JOIN SCM.RfqProductType e ON e.RfqProTypeId = a.RfqProTypeId
                        LEFT JOIN SCM.RfqProductClass m ON m.RfqProClassId = e.RfqProClassId
                        LEFT JOIN BAS.[User] g ON g.UserId = a.SalesId
                        LEFT JOIN BAS.Department i ON i.DepartmentId = g.DepartmentId
                        LEFT JOIN BAS.Company j ON j.CompanyId = i.CompanyId
                        INNER JOIN BAS.[File] n ON n.FileId = a.CustProdDigram AND n.DeleteStatus = 'N'
                        LEFT JOIN BAS.[Type] f ON f.TypeNo = a.ProtoSchedule AND f.TypeSchema = 'RfqDetail.ProtoSchedule'
                        INNER JOIN BAS.[Status] h ON h.StatusNo = a.[Status] AND h.StatusSchema = 'RfqDetail.Status'
                        LEFT JOIN BAS.[Status] k ON k.StatusNo = a.MassProductionDemand AND k.StatusSchema = 'Boolean'
                        INNER JOIN BAS.[Status] l ON l.StatusNo = a.CoatingFlag AND l.StatusSchema = 'Boolean'";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RfqId", @" AND a.RfqId = @RfqId", RfqId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RfqDetailId", @" AND a.RfqDetailId = @RfqDetailId", RfqDetailId);
                    sqlQuery.conditions = queryCondition;

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

        #region //GetRfqProductClassForList -- 取得RFQ產品型別種類(客戶端 清單用) -- Chia Yuan 2023.07.11
        public string GetRfqProductForList(int RfqProClassId, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.RfqProClassId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",a.RfqProductClassName,a.[Status] as RfqProClassStatus
                        , (SELECT b.RfqProTypeId,b.RfqProductTypeName,b.[Status] as RfqProTypeStatus FROM SCM.RfqProductType b 
                            WHERE b.RfqProClassId = a.RfqProClassId and b.[Status] = @Status FOR JSON PATH, ROOT('data')) as RfqProductType";
                    sqlQuery.mainTables =
                        @"FROM SCM.RfqProductClass a";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RfqProClassId", @" AND a.RfqProClassId = @RfqProClassId", RfqProClassId);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.RfqProClassId";
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

        #region //GetSales -- 取得RFQ負責業務(客戶端 Cmb用) -- Chia Yuan 2023.07.11
        public string GetSales(int UserId, string UserName
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.UserId, a.UserName, b.DepartmentName, c.CompanyName
                            FROM BAS.[User] a
                            INNER JOIN BAS.Department b ON b.DepartmentId = a.DepartmentId
                            INNER JOIN BAS.Company c ON c.CompanyId = b.CompanyId
                            WHERE b.DepartmentName LIKE '%業務%'
                            AND a.UserStatus != 'S'
                            AND a.UserName != '來賓'
                            AND c.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserId", @" AND a.UserId = @UserId", UserId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserName", @" AND a.UserName LIKE '%' + @UserName + '%'", UserName);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.UserId";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

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

        #region //GetRfqQtyLevel -- 取得RFQ報價方案(客戶端) -- Chia Yuan 2023.07.12
        public string GetRfqQtyLevel(int RfqQtyLevelId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.RfqQtyLevelId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.SolutionQty, a.RfqQtyLevelName";
                    sqlQuery.mainTables =
                        @"FROM SCM.RfqQtyLevel a";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RfqQtyLevelId", @" AND a.RfqQtyLevelId = @RfqQtyLevelId", RfqQtyLevelId);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.SortNumber";
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

        #region //GetReturnRfqDetailDoc -- 取得報價單資料 --Chia Yuan -- 2023.07.31
        public string GetReturnRfqDetailDoc(int FileId
            , string Account, string KeyText)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    int rowsAffected = 0;
                    #region //判斷使用者登入金鑰是否存在
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 1
                                FROM EIP.MemberLoginKey a
                                INNER JOIN EIP.[Member] b ON a.MemberId = b.MemberId
                                WHERE b.MemberEmail = @Account
                                AND a.KeyText = @KeyText";
                    dynamicParameters.Add("Account", Account);
                    dynamicParameters.Add("KeyText", KeyText);
                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (!result.Any()) throw new SystemException("【使用者】資料錯誤!");
                    #endregion

                    #region //取得資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.FileId, a.[FileName], a.FileContent, a.FileExtension, a.FileSize, a.[Source], a.DeleteStatus
                            FROM BAS.[File] a
                            WHERE a.FileId = @FileId";
                    dynamicParameters.Add("FileId", FileId);
                    result = sqlConnection.Query(sql, dynamicParameters);
                    if (!result.Any()) throw new SystemException("【報價單】資料錯誤!!");
                    #endregion

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

        #region //GetFile -- 取得報價單資料 --Chia Yuan -- 2023.07.31
        public string GetFile(int FileId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //取得報價單
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.FileId, a.[FileName], a.FileContent, a.FileExtension, a.FileSize, a.[Source], a.DeleteStatus
                            FROM BAS.[File] a WHERE a.FileId = @FileId";
                    dynamicParameters.Add("FileId", FileId);
                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (!result.Any()) throw new SystemException("【報價單】資料錯誤!!");
                    #endregion

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

        #region //GetDocRoleInformation -- 取得角色單據資料 -- Xuan -- 2023.08.24
        public string GetDocRoleInformation(string Status, string RoleName, int RoleId, string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.RoleId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",a.RoleName, a.Status";
                    sqlQuery.mainTables =
                        @"FROM EIP.NotifyRole a";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RoleId", @" AND a.RoleId = @RoleId", RoleId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RoleName", @" AND a.RoleName = @RoleName", RoleName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.CreateDate";
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

        #region //GetDocUserInformation -- 取得角色使用者單據資料 -- Xuan -- 2023.08.24
        public string GetDocUserInformation(string UserName, string RoleName, int RoleId, string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.RoleId,c.UserId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",b.RoleName,c.UserName";
                    sqlQuery.mainTables =
                        @"FROM EIP.NotifyUser a
                          INNER JOIN EIP.NotifyRole b ON b.RoleId=a.RoleId
                          INNER JOIN BAS.[User] c ON c.UserId=a.UserId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.RoleId=@RoleId";
                    dynamicParameters.Add("RoleId", RoleId);
                    //BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RoleId", @" AND a.RoleId = @RoleId", RoleId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RoleName", @" AND b.RoleName = @RoleName", RoleName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserName", @" AND c.UserName = @UserName", UserName);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.CreateDate";
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

        #region //GetDocInformation -- 取得單據權限資料 -- Xuan -- 2023.08.29
        public string GetDocInformation(string DocType, string RoleName, int RoleId, string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "b.DocNotifyId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",a.RoleName
                          ,b.DocType
                          ,b.SubProdType
                          ,CASE 
                              WHEN b.DocType = 'RFQ' THEN c.RfqProductTypeName 
                              WHEN b.DocType = 'DFM' THEN d.DfmItemCategoryName 
                              ELSE NULL 
                          END AS SubProductName";
                    sqlQuery.mainTables =
                        @"FROM 
                          EIP.NotifyRole a
                          INNER JOIN EIP.DocNotify b ON b.RoleId = a.RoleId
                          LEFT JOIN SCM.RfqProductType c ON b.DocType = 'RFQ' AND c.RfqProTypeId = b.SubProdId
                          LEFT JOIN PDM.DfmItemCategory d ON b.DocType = 'DFM' AND d.DfmItemCategoryId = b.SubProdId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.RoleId=@RoleId";
                    dynamicParameters.Add("RoleId", RoleId);
                    //BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RoleId", @" AND a.RoleId = @RoleId", RoleId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RoleName", @" AND a.RoleName = @RoleName", RoleName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DocType", @" AND b.DocType = @DocType", DocType);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "b.CreateDate";
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

        #region //GetUser_S -- 取得人員在職資訊 -- Xuan -- 2023.08.29
        public string GetUser_S(string DepartmentName, int DepartmentId, string UserName, int UserId, string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.UserId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",a.UserNo,a.UserName,(a.UserNo+'  '+a.UserName) UserNoName,b.DepartmentId,b.DepartmentName";
                    sqlQuery.mainTables =
                        @"FROM BAS.[User] a
                          INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
 　                       INNER JOIN BAS.Company c ON b.CompanyId = c.CompanyId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND b.CompanyId=2 AND a.UserStatus='F'";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserId", @" AND a.UserId = @UserId", UserId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserName", @" AND a.UserName = @UserName", UserName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DepartmentId", @" AND b.DepartmentId = @DepartmentId", DepartmentId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DepartmentName", @" AND b.DepartmentName = @DepartmentName", DepartmentName);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "c.CompanyId, a.UserNo";
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

        #region //GetDepartment_A -- 取得部門啟用資訊 -- Xuan -- 2023.08.30
        public string GetDepartment_A(string UserName, int UserId, string DepartmentName, int DepartmentId, string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.UserId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",a.UserName,b.DepartmentId,b.DepartmentName";
                    sqlQuery.mainTables =
                        @"FROM BAS.[User] a
                          INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
 　                       INNER JOIN BAS.Company c ON b.CompanyId = c.CompanyId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND b.CompanyId=2 AND a.UserStatus='F'";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserId", @" AND a.UserId = @UserId", UserId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserName", @" AND a.UserName = @UserName", UserName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DepartmentId", @" AND b.DepartmentId = @DepartmentId", DepartmentId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DepartmentName", @" AND b.DepartmentName = @DepartmentName", DepartmentName);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "c.CompanyId, a.UserNo";
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

        #region //GetRfqProductType -- 取得RFQ資訊 -- Xuan -- 2023.08.29
        public string GetRfqProductType(string RfqProductTypeName, int RfqProTypeId, string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.RfqProTypeId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",a.RfqProductTypeName";
                    sqlQuery.mainTables =
                        @"FROM SCM.RfqProductType a";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RfqProTypeId", @" AND a.RfqProTypeId = @RfqProTypeId", RfqProTypeId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RfqProductTypeName", @" AND a.RfqProductTypeName = @RfqProductTypeName", RfqProductTypeName);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.LastModifiedDate DESC";
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

        #region //GetDfmItemCategory -- 取得DFM資訊 -- Xuan -- 2023.08.29
        public string GetDfmItemCategory(string DfmItemCategoryName, int DfmItemCategoryId, string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.DfmItemCategoryId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",a.DfmItemCategoryName";
                    sqlQuery.mainTables =
                        @"FROM PDM.DfmItemCategory a";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DfmItemCategoryId", @" AND a.DfmItemCategoryId = @DfmItemCategoryId", DfmItemCategoryId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DfmItemCategoryName", @" AND a.DfmItemCategoryName = @DfmItemCategoryName", DfmItemCategoryName);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.LastModifiedDate DESC";
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

        #region //GetMemberIcon -- 取得會員圖像 -- Chia Yuan -- 2024.03.13
        public string GetMemberIcon(int FileId, string DeleteStatus, string StartDate, string EndDate
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(OfficialConnectionStrings))
                {
                    sqlQuery.mainKey = "a.FileId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.[FileName], a.FileExtension, a.FileSize, a.Source, a.DeleteStatus
                        , Format(a.CreateDate, 'yyyy-MM-dd HH:mm:ss') AS UploadDate
                        , Format(a.LastModifiedDate, 'yyyy-MM-dd HH:mm:ss') DeleteDate";
                    if (FileId > 0) sqlQuery.columns += @", a.FileContent";
                    sqlQuery.mainTables =
                        @"FROM EIP.MemberFile a
                        INNER JOIN EIP.[Member] b ON b.MemberIcon = a.FileId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "FileId", @" AND a.FileId = @FileId", FileId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DeleteStatus", @" AND a.DeleteStatus = @DeleteStatus", DeleteStatus);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.CreateDate >= @StartDate ", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.CreateDate <= @EndDate ", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.FileId DESC";
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

        #region //AddUser -- 使用者資料新增 -- Chia Yuan 2023.7.7
        public string RegisterUser(string Account, string Password, string MemberName, string OrgShortName, string ContactName, string ContactPhone)
        {
            int atIdx = Account.IndexOf('@');
            if (atIdx < 0) throw new SystemException("【信箱】格式錯誤!");
            if (string.IsNullOrWhiteSpace(Account)) throw new SystemException("【信箱】不能為空!");
            if (Account.Length > 100) throw new SystemException("【信箱】長度錯誤!");
            if (string.IsNullOrWhiteSpace(Password)) throw new SystemException("【密碼】不能為空!");
            if (Password.Length > 128) throw new SystemException("【密碼】長度錯誤!");
            if (string.IsNullOrWhiteSpace(MemberName)) throw new SystemException("【使用者名稱】長度錯誤!");
            if (string.IsNullOrWhiteSpace(OrgShortName)) throw new SystemException("【公司名稱】長度錯誤!");
            //if (string.IsNullOrWhiteSpace(ContactName)) throw new SystemException("【聯絡人】長度錯誤!");
            if (string.IsNullOrWhiteSpace(ContactPhone)) throw new SystemException("【電話】長度錯誤!");

            //bool telcheck = Regex.IsMatch(ContactPhone, @"^[0-9]$");//規則:09開頭，後面接著8

            Account = Account.Trim();
            ContactName = ContactName.Trim();
            ContactPhone = ContactPhone.Trim();

            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //密碼參數設定
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.PasswordExpiration, a.PasswordFormat, a.PasswordWrongCount
                            FROM BAS.PasswordSetting a";

                        var resultPasswordSetting = sqlConnection.Query(sql, dynamicParameters);
                        if (!resultPasswordSetting.Any()) throw new SystemException("【密碼參數設定】資料錯誤!");

                        int PasswordExpiration = -1;
                        string PasswordFormat = "";
                        int PasswordWrongCount = -1;
                        foreach (var item in resultPasswordSetting)
                        {
                            PasswordExpiration = Convert.ToInt32(item.PasswordExpiration);
                            PasswordFormat = item.PasswordFormat;
                            PasswordWrongCount = Convert.ToInt32(item.PasswordWrongCount);
                        }
                        #endregion

                        if (!Regex.IsMatch(Password, PasswordFormat, RegexOptions.IgnoreCase)) throw new SystemException("【密碼】格式錯誤!");

                        #region //判斷使用者編號是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EIP.[Member]
                                WHERE MemberEmail = @Account";
                        dynamicParameters.Add("Account", Account);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Any()) throw new SystemException("【信箱】已被使用，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EIP.[Member] (MemberName,MemberEmail,Password,Address,OrgShortName,ContactName,ContactPhone,ContactEmail,CertCode,Status,CreateDate,LastModifiedDate,CreateBy,LastModifiedBy)
                                OUTPUT INSERTED.MemberId
                                VALUES (@MemberName, @MemberEmail, @Password, @Address, @OrgShortName, @ContactName, @ContactPhone, @ContactEmail, @CertCode
                                    , @Status, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                MemberName, //Account.Substring(0, atIdx > 0 ? atIdx : Account.Length),
                                MemberEmail = Account,
                                Password = BaseHelper.Sha256Encrypt(Password),
                                Address = (string)null, //string.Empty,
                                OrgShortName,
                                ContactName = (string)null,
                                ContactPhone,
                                ContactEmail = Account, // string.Empty,
                                CertCode = string.Empty,
                                Status = "A", //啟用
                                CreateDate,
                                LastModifiedDate,
                                CreateBy = -1,
                                LastModifiedBy = -1
                            });

                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += insertResult.Count();

                        //if (rowsAffected > 0) {
                        //    int memberId = insertResult.FirstOrDefault().MemberId;
                        //    dynamicParameters = new DynamicParameters();
                        //    sql = @"insert into EIP.[MemberOrganization] 
                        //        OUTPUT INSERTED.OrgId
                        //        VALUES (@MemberId, @OrganizaitonType, @OrganizationCode, @OrganizaitonScale, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        //    dynamicParameters.AddDynamicParams(
                        //        new
                        //        {
                        //            MemberId = memberId,
                        //            OrganizaitonType = 1, //預設客戶身分
                        //            OrganizationCode = string.Empty,
                        //            OrganizaitonScale = string.Empty,
                        //            CreateDate,
                        //            LastModifiedDate,
                        //            CreateBy = memberId,
                        //            LastModifiedBy = memberId
                        //        });
                        //    result = sqlConnection.Query(sql, dynamicParameters);
                        //    if (!result.Any()) throw new SystemException("【註冊資訊-公司】錯誤!");
                        //}

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

        #region //AddRequestForQuotation --RFQ單頭新增 --Chia Yuan 2023.7.7
        public string AddRequestForQuotation(string AssemblyName, int ProductUseId, string CustomerName, int CustomerId, string Status
            , int MemberType, string Account, string KeyText)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(AssemblyName)) throw new SystemException("【機種名稱】不能為空!");

                AssemblyName = AssemblyName.Trim();

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //取得客戶資料
                        int MemberId = -1, UserId = -1;
                        if (MemberType == 1)
                        {
                            #region //判斷使用者登入金鑰是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 b.MemberId
                                FROM EIP.MemberLoginKey a
                                INNER JOIN EIP.[Member] b ON a.MemberId = b.MemberId
                                WHERE b.MemberEmail = @Account
                                AND a.KeyText = @KeyText";
                            dynamicParameters.Add("Account", Account);
                            dynamicParameters.Add("KeyText", KeyText);
                            var userResult = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                            if (userResult == null) throw new SystemException("【使用者】資料錯誤!");
                            MemberId = userResult.MemberId;
                            #endregion
                        }
                        if (MemberType == 2)
                        {
                            #region //判斷使用者登入金鑰是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 b.UserId
                                FROM BAS.UserLoginKey a
                                INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                                WHERE b.UserNo = @Account
                                AND a.KeyText = @KeyText";
                            dynamicParameters.Add("Account", Account);
                            dynamicParameters.Add("KeyText", KeyText);
                            var userResult = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                            if (userResult == null) throw new SystemException("【使用者】資料錯誤!");
                            UserId = userResult.UserId;
                            #endregion

                            #region //判斷客戶資料是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.CustomerId
                                    FROM SCM.Customer a 
                                    WHERE a.CustomerId = @CustomerId";
                            dynamicParameters.Add("CustomerId", CustomerId);
                            userResult = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                            if (userResult == null) throw new SystemException("【客戶】資料錯誤!");
                            #endregion
                        }
                        #endregion

                        #region //判斷RFQ產品應用資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.ProductUse
                                WHERE ProductUseId = @ProductUseId";
                        dynamicParameters.Add("ProductUseId", ProductUseId);
                        var result = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                        if (result == null) throw new SystemException("【產品應用】資料錯誤!");
                        #endregion

                        #region //需求單號取號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT CONVERT(int, RIGHT(ISNULL(MAX(RfqNo), '000'), 3)) + 1 CurrentNum
                            FROM SCM.RequestForQuotation
                            WHERE RfqNo LIKE @RfqNo";
                        dynamicParameters.Add("RfqNo", string.Format("{0}{1}___", "RFQ", DateTime.Now.ToString("yyyyMMdd")));
                        int currentNum = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).CurrentNum;
                        string RfqNo = string.Format("{0}{1}{2}", "RFQ", DateTime.Now.ToString("yyyyMMdd"), string.Format("{0:000}", currentNum));
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.RequestForQuotation (RfqNo, AssemblyName, ProductUseId, MemberType, MemberId, UserId, CustomerId, CustomerName, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.RfqId
                                VALUES (@RfqNo, @AssemblyName, @ProductUseId, @MemberType, @MemberId, @UserId, @CustomerId, @CustomerName, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RfqNo,
                                AssemblyName,
                                ProductUseId,
                                MemberType,
                                MemberId = MemberId > 0 ? MemberId : (int?)null,
                                UserId = UserId > 0 ? UserId : (int?)null,
                                CustomerId = CustomerId > 0 ? CustomerId : (int?)null,
                                CustomerName = string.IsNullOrWhiteSpace(CustomerName) ? null : CustomerName,
                                Status = "0",
                                CreateDate,
                                LastModifiedDate,
                                CreateBy = MemberId > 0 ? MemberId : UserId,
                                LastModifiedBy = MemberId > 0 ? MemberId : UserId
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

        #region //AddRequestForQuotationFromJson --RFQ單頭一次新增 --Chia Yuan 2023.7.7
        public string AddRequestForQuotationFromJson(string RfqEstimate, int MemberType, string Account, string KeyText)
        {
            try
            {
                //if (!RfqEstimate.TryParseJson(out JObject tempJObject)) throw new SystemException("購物車資料格式錯誤");

                JObject jObject = new JObject(); //新建 操作對象
                List<RfqMainVM> arrayData = JsonConvert.DeserializeObject<List<RfqMainVM>>(RfqEstimate);

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷使用者登入金鑰是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 b.MemberId
                                FROM EIP.MemberLoginKey a
                                INNER JOIN EIP.[Member] b ON a.MemberId = b.MemberId
                                WHERE b.MemberEmail = @Account
                                AND a.KeyText = @KeyText";
                        dynamicParameters.Add("Account", Account);
                        dynamicParameters.Add("KeyText", KeyText);

                        var result = sqlConnection.Query(sql, dynamicParameters).FirstOrDefault();
                        if (result == null) throw new SystemException("【使用者】資料錯誤!");
                        int MemberId = result.MemberId;
                        #endregion

                        int rowsAffected = 0;
                        arrayData.ForEach(item =>
                        {
                            if (string.IsNullOrWhiteSpace(item.AssemblyName)) throw new SystemException("【機種名稱】資料錯誤!");
                            if (item.ProductUseId <= 0) throw new SystemException("【產品應用】資料錯誤!");

                            #region //判斷RFQ產品應用資料是否正確
                            sql = @"SELECT TOP 1 ProductUseId
                                FROM SCM.ProductUse
                                WHERE ProductUseId = @ProductUseId";
                            dynamicParameters.Add("ProductUseId", item.ProductUseId);
                            var result1 = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                            if (result1 == null) throw new SystemException("【產品應用】資料錯誤!");
                            #endregion

                            #region //需求單號取號
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT CONVERT(int, RIGHT(ISNULL(MAX(RfqNo), '000'), 3)) + 1 CurrentNum
                            FROM SCM.RequestForQuotation
                            WHERE RfqNo LIKE @RfqNo";
                            dynamicParameters.Add("RfqNo", string.Format("{0}{1}___", "RFQ", DateTime.Now.ToString("yyyyMMdd")));
                            int currentNum = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).CurrentNum;
                            string RfqNo = string.Format("{0}{1}{2}", "RFQ", DateTime.Now.ToString("yyyyMMdd"), string.Format("{0:000}", currentNum));
                            #endregion

                            #region //新增報價單頭
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO SCM.RequestForQuotation (RfqNo, AssemblyName, ProductUseId, MemberId, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.RfqId
                                VALUES (@RfqNo, @AssemblyName, @ProductUseId, @MemberId, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    RfqNo,
                                    item.AssemblyName,
                                    item.ProductUseId,
                                    MemberId,
                                    CustomerId = MemberType == 1 ? MemberId : item.CustomerId,
                                    Status = "0", //預設啟用
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy = MemberId,
                                    LastModifiedBy = MemberId
                                });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);
                            rowsAffected += insertResult.Count();
                            int RfqId = -1;
                            foreach (var innerItem in insertResult)
                            {
                                RfqId = Convert.ToInt32(innerItem.RfqId);
                            }
                            #endregion

                            item.RfqDetail.ForEach(detail =>
                            {
                                #region //取得流水號
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT CONVERT(int, RIGHT(ISNULL(MAX(RfqSequence), '000'), 3)) + 1 as CurrentRfqSequence
                                    FROM SCM.RfqDetail
                                    WHERE RfqId = @RfqId";
                                dynamicParameters.AddDynamicParams(new { RfqId });
                                int currentRfqSequence = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).CurrentRfqSequence;
                                string RfqSequence = string.Format("{0:0000}", currentRfqSequence);
                                #endregion

                                #region //新增報價單身
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO SCM.RfqDetail (RfqId,CompanyId,RfqProTypeId,RfqSequence,MtlName
                                    ,CustProdDigram,PlannedOpeningDate,PrototypeQty,ProtoSchedule,MassProductionDemand,KickOffType
                                    ,PlasticName,OutsideDiameter,ProdLifeCycleStart,ProdLifeCycleEnd
                                    ,LifeCycleQty,DemandDate,CoatingFlag,SalesId,AdditionalFile,QuotationFile
                                    ,Description,ConfirmSalesTime,ConfirmRdTime,Status
                                    ,CreateDate,LastModifiedDate,CreateBy,LastModifiedBy)
                                    OUTPUT INSERTED.RfqDetailId
                                    VALUES (@RfqId,@CompanyId,@RfqProTypeId,@RfqSequence,@MtlName
                                    ,@CustProdDigram,@PlannedOpeningDate,@PrototypeQty,@ProtoSchedule,@MassProductionDemand,@KickOffType
                                    ,@PlasticName,@OutsideDiameter,@ProdLifeCycleStart,@ProdLifeCycleEnd
                                    ,@LifeCycleQty,@DemandDate,@CoatingFlag,@SalesId,@AdditionalFile,@QuotationFile
                                    ,@Description,@ConfirmSalesTime,@ConfirmRdTime,@Status
                                    ,@CreateDate,@LastModifiedDate,@CreateBy,@LastModifiedBy)";

                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        RfqId,
                                        CompanyId = 2, //前端傳入
                                        detail.RfqProTypeId,
                                        RfqSequence,
                                        detail.MtlName,
                                        CustProdDigram = (int?)null,
                                        PlannedOpeningDate = (DateTime?)null,
                                        detail.PrototypeQty,
                                        ProtoSchedule = (string)null,
                                        MassProductionDemand = "N",
                                        KickOffType = (string)null,
                                        PlasticName = (string)null,
                                        OutsideDiameter = (int?)null,
                                        detail.ProdLifeCycleStart,
                                        detail.ProdLifeCycleEnd,
                                        detail.LifeCycleQty,
                                        detail.DemandDate,
                                        detail.CoatingFlag,
                                        SalesId = (int?)null,
                                        AdditionalFile = (int?)null,
                                        QuotationFile = (int?)null,
                                        detail.Description,
                                        ConfirmSalesTime = (DateTime?)null,
                                        ConfirmRdTime = (DateTime?)null,
                                        Status = "1",
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy = MemberId,
                                        LastModifiedBy = MemberId
                                    });

                                insertResult = sqlConnection.Query(sql, dynamicParameters);
                                rowsAffected += insertResult.Count();
                                int RfqDetailId = -1;
                                foreach (var innerItem in insertResult)
                                {
                                    RfqDetailId = Convert.ToInt32(innerItem.RfqDetailId);
                                }
                                //RfqDetailId = insertResult.FirstOrDefault().RfqDetailId;
                                #endregion

                                #region //判斷RFQ包裝種類資料是否正確
                                sql = @"SELECT TOP 1 RfqPkTypeId,PackagingMethod,SustSupplyStatus
                                FROM SCM.RfqPackageType
                                WHERE RfqPkTypeId = @RfqPkTypeId
                                AND RfqProClassId = @RfqProClassId
                                AND Status = @Status";
                                dynamicParameters.Add("RfqPkTypeId", detail.RfqPkTypeId);
                                dynamicParameters.Add("RfqProClassId", detail.RfqProClassId);
                                dynamicParameters.Add("Status", "A");
                                var result2 = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                                if (result2 == null) throw new SystemException("【包裝種類】資料錯誤!");
                                #endregion

                                #region //新增包裝方式
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO SCM.RfqPackage (RfqDetailId,RfqPkTypeId,SustSupplyStatus,Status
                                    ,CreateDate,LastModifiedDate,CreateBy,LastModifiedBy)
                                    OUTPUT INSERTED.RfqPkId
                                    VALUES (@RfqDetailId,@RfqPkTypeId,@SustSupplyStatus,@Status
                                    ,@CreateDate,@LastModifiedDate,@CreateBy,@LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        RfqDetailId,
                                        result2.RfqPkTypeId,
                                        result2.SustSupplyStatus,
                                        Status = "A",
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy = MemberId,
                                        LastModifiedBy = MemberId
                                    });
                                insertResult = sqlConnection.Query(sql, dynamicParameters);
                                rowsAffected += insertResult.Count();
                                int RfqPkId = -1;
                                foreach (var innerItem in insertResult)
                                {
                                    RfqPkId = Convert.ToInt32(innerItem.RfqPkId);
                                }
                                #endregion

                                int idx = 1;
                                detail.RfqLineSolutionList.ForEach(solution =>
                                {
                                    #region //新增報價方案
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO SCM.RfqLineSolution (RfqDetailId,SortNumber,SolutionQty,PeriodicDemandType
                                    ,CreateDate,LastModifiedDate,CreateBy,LastModifiedBy)
                                    OUTPUT INSERTED.RfqLineSolutionId
                                    VALUES (@RfqDetailId,@SortNumber,@SolutionQty,@PeriodicDemandType
                                    ,@CreateDate,@LastModifiedDate,@CreateBy,@LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            RfqDetailId,
                                            SortNumber = solution.idx,
                                            solution.SolutionQty,
                                            solution.PeriodicDemandType,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy = MemberId,
                                            LastModifiedBy = MemberId
                                        });
                                    insertResult = sqlConnection.Query(sql, dynamicParameters);
                                    rowsAffected += insertResult.Count();
                                    int RfqLineSolutionId = -1;
                                    foreach (var innerItem in insertResult)
                                    {
                                        RfqLineSolutionId = Convert.ToInt32(innerItem.RfqLineSolutionId);
                                    }
                                    #endregion
                                    idx++;
                                });
                            });
                        });

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

        #region //AddRfqDetail -- RFQ單身新增 --Chia Yuan 2023.7.20
        public string AddRfqDetail(int RfqId, int RfqProTypeId, string MtlName, int PrototypeQty
            , string ProdLifeCycleStart, string ProdLifeCycleEnd, int LifeCycleQty, string DemandDate, string CoatingFlag, int AdditionalFile, string Description
            , int RfqPkTypeId, string RfqLineSolution, string Currency, string PortOfDelivery
            , List<FileModel> Files, string ClientIP, string savePath
            , int MemberType, string Account, string KeyText)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(MtlName)) throw new SystemException("【品名】不能為空!");
                if (string.IsNullOrWhiteSpace(ProdLifeCycleStart)) throw new SystemException("【生命週期(起)】不能為空!");
                if (string.IsNullOrWhiteSpace(ProdLifeCycleEnd)) throw new SystemException("【生命週期(迄)】不能為空!");
                if (string.IsNullOrWhiteSpace(DemandDate)) throw new SystemException("【需求日期】不能為空!");
                if (string.IsNullOrWhiteSpace(CoatingFlag)) throw new SystemException("【是否鍍膜】不能為空!");
                if (string.IsNullOrWhiteSpace(PortOfDelivery)) throw new SystemException("【交貨地點】不能為空!");
                if (Description.Length >= 100) throw new SystemException("【描述】格式錯誤!");
                if (PrototypeQty <= 0) throw new SystemException("【試作需求數量】格式錯誤!");
                if (LifeCycleQty <= 0) throw new SystemException("【週期數量】格式錯誤!");
                if (!DateTime.TryParse(ProdLifeCycleStart, out DateTime prodLifeCycleStart)) throw new SystemException("【生命週期(起)】格式錯誤!");
                if (!DateTime.TryParse(ProdLifeCycleEnd, out DateTime prodLifeCycleEnd)) throw new SystemException("【生命週期(迄)】格式錯誤!");
                if (!DateTime.TryParse(DemandDate, out DateTime demandDate)) throw new SystemException("【需求日期】格式錯誤!");
                if (prodLifeCycleStart.CompareTo(prodLifeCycleEnd) > 0) throw new SystemException("【生命週期】區間錯誤!");

                List<RfqLineSolutionVM> RfqLineSolutionData = JsonConvert.DeserializeObject<List<RfqLineSolutionVM>>(RfqLineSolution);
                if (!RfqLineSolutionData.Any()) throw new SystemException("【報價依據】錯誤!");

                CoatingFlag = CoatingFlag.Trim();

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //取得客戶資料
                        int MemberId = -1, UserId = -1;
                        if (MemberType == 1)
                        {
                            #region //判斷使用者登入金鑰是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 b.MemberId
                                FROM EIP.MemberLoginKey a
                                INNER JOIN EIP.[Member] b ON a.MemberId = b.MemberId
                                WHERE b.MemberEmail = @Account
                                AND a.KeyText = @KeyText";
                            dynamicParameters.Add("Account", Account);
                            dynamicParameters.Add("KeyText", KeyText);
                            var userResult = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                            if (userResult == null) throw new SystemException("【使用者】資料錯誤!");
                            MemberId = userResult.MemberId;
                            #endregion
                        }
                        if (MemberType == 2)
                        {
                            #region //判斷使用者登入金鑰是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 b.UserId
                                FROM BAS.UserLoginKey a
                                INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                                WHERE b.UserNo = @Account
                                AND a.KeyText = @KeyText";
                            dynamicParameters.Add("Account", Account);
                            dynamicParameters.Add("KeyText", KeyText);
                            var userResult = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                            if (userResult == null) throw new SystemException("【使用者】資料錯誤!");
                            UserId = userResult.UserId;
                            #endregion
                        }
                        #endregion

                        #region //判斷RFQ單頭資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.RequestForQuotation
                                WHERE RfqId = @RfqId";
                        dynamicParameters.Add("RfqId", RfqId);
                        var result = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                        if (result == null) throw new SystemException("【詢價單】資料錯誤!");
                        #endregion

                        #region //判斷RFQ產品類別資料是否正確
                        sql = @"SELECT TOP 1 RfqProClassId
                                FROM SCM.RfqProductType
                                WHERE RfqProTypeId = @RfqProTypeId";
                        dynamicParameters.Add("RfqProTypeId", RfqProTypeId);
                        result = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                        if (result == null) throw new SystemException("【產品類別】資料錯誤!");
                        int RfqProClassId = result.RfqProClassId;
                        #endregion

                        #region //判斷RFQ是否度模資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Status]
                                WHERE StatusSchema = 'Boolean' and StatusNo = @CoatingFlag";
                        dynamicParameters.Add("CoatingFlag", CoatingFlag);
                        result = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                        if (result == null) throw new SystemException("【是否鍍膜】資料錯誤!");
                        #endregion

                        #region //圖檔新增
                        int FileId = -1;
                        Files.ForEach(file =>
                        {
                            sql = @"INSERT INTO BAS.[File] (CompanyId, FileName, FileContent, FileExtension, FileSize
                                   , ClientIP, Source, DeleteStatus
                                   , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                   OUTPUT INSERTED.FileId
                                   VALUES (@CompanyId, @FileName, @FileContent, @FileExtension, @FileSize
                                   , @ClientIP, @Source, @DeleteStatus
                                   , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    CompanyId = 2,
                                    file.FileName,
                                    file.FileContent,
                                    file.FileExtension,
                                    file.FileSize,
                                    ClientIP,
                                    Source = savePath,
                                    DeleteStatus = "N",
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            var insertResult1 = sqlConnection.Query(sql, dynamicParameters);
                            var rowsAffected1 = insertResult1.Count();
                            foreach (var innerItem in insertResult1)
                            {
                                FileId = Convert.ToInt32(innerItem.FileId);
                            }
                        });
                        #endregion

                        #region //取得流水號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT CONVERT(int, RIGHT(ISNULL(MAX(RfqSequence), '000'), 3)) + 1 as CurrentRfqSequence
                                    FROM SCM.RfqDetail
                                    WHERE RfqId = @RfqId";
                        dynamicParameters.AddDynamicParams(new { RfqId });
                        int currentRfqSequence = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).CurrentRfqSequence;
                        string RfqSequence = string.Format("{0:0000}", currentRfqSequence);
                        #endregion

                        #region //取得預設業務
                        sql = @"SELECT ISNULL(x.SalesId,y.SalesId) as SalesId FROM (
	                                SELECT * FROM (
		                                SELECT distinct a.MemberId,a.UserId,ISNULL(a.CustomerId,c.OrganizaitonTypeId) as CustomerId,b.SalesId
		                                ,row_number() over (partition by a.MemberId order by b.CreateDate desc) as Rnak 
		                                FROM SCM.RequestForQuotation a
		                                join SCM.RfqDetail b on b.RfqId = a.RfqId 
		                                join EIP.MemberOrganization c on c.MemberId = a.MemberId and c.OrganizaitonType = 1
	                                ) a WHERE a.Rnak = 1
                                ) x
                                LEFT JOIN (
	                                SELECT * FROM (
		                                SELECT distinct a.MemberId,b.UserId,ISNULL(b.CustomerId, a.OrganizaitonTypeId) as CustomerId,c.SalesId
		                                ,row_number() over (partition by ISNULL(b.CustomerId, a.OrganizaitonTypeId) order by c.CreateDate desc) as Rnak 
		                                FROM EIP.MemberOrganization a
		                                join SCM.RequestForQuotation b on b.MemberId = a.MemberId
		                                join SCM.RfqDetail c on c.RfqId = b.RfqId
		                                WHERE a.OrganizaitonType = 1 and ISNULL(b.CustomerId, a.OrganizaitonTypeId) > 0
	                                ) a WHERE a.Rnak = 1
                                ) y on y.CustomerId = x.CustomerId
                                WHERE 1=1";


                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MemberId", @" AND x.MemberId = @MemberId", MemberId);
                        //BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "UserId", @" AND x.UserId = @UserId", UserId);
                        result = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                        //if (result == null) throw new SystemException("【產品類別】資料錯誤!");
                        int? SalesId = result?.SalesId;

                        #endregion

                        #region //新增報價單身
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.RfqDetail (RfqId,CompanyId,RfqProTypeId,RfqSequence,MtlName
                                    ,CustProdDigram,PlannedOpeningDate,PrototypeQty,ProtoSchedule,MassProductionDemand,KickOffType
                                    ,PlasticName,OutsideDiameter,ProdLifeCycleStart,ProdLifeCycleEnd
                                    ,LifeCycleQty,DemandDate,CoatingFlag,SalesId,AdditionalFile,QuotationFile
                                    ,Description,ConfirmVPTime,ConfirmSalesTime,ConfirmRdTime,Currency,PortOfDelivery,Status
                                    ,CreateDate,LastModifiedDate,CreateBy,LastModifiedBy)
                                    OUTPUT INSERTED.RfqDetailId
                                    VALUES (@RfqId,@CompanyId,@RfqProTypeId,@RfqSequence,@MtlName
                                    ,@CustProdDigram,@PlannedOpeningDate,@PrototypeQty,@ProtoSchedule,@MassProductionDemand,@KickOffType
                                    ,@PlasticName,@OutsideDiameter,@ProdLifeCycleStart,@ProdLifeCycleEnd
                                    ,@LifeCycleQty,@DemandDate,@CoatingFlag,@SalesId,@AdditionalFile,@QuotationFile
                                    ,@Description,@ConfirmVPTime,@ConfirmSalesTime,@ConfirmRdTime,@Currency,@PortOfDelivery,@Status
                                    ,@CreateDate,@LastModifiedDate,@CreateBy,@LastModifiedBy)";

                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RfqId,
                                CompanyId = 2, //前端傳入
                                RfqProTypeId,
                                RfqSequence,
                                MtlName,
                                CustProdDigram = FileId == -1 ? (int?)null : FileId,
                                PlannedOpeningDate = (DateTime?)null,
                                PrototypeQty,
                                ProtoSchedule = (string)null,
                                MassProductionDemand = "N",
                                KickOffType = (string)null,
                                PlasticName = (string)null,
                                OutsideDiameter = (int?)null,
                                ProdLifeCycleStart,
                                ProdLifeCycleEnd,
                                LifeCycleQty,
                                DemandDate,
                                CoatingFlag,
                                SalesId = MemberType == 2 ? UserId : SalesId,
                                AdditionalFile = AdditionalFile == -1 ? (int?)null : AdditionalFile,
                                QuotationFile = (int?)null,
                                Description,
                                ConfirmVPTime = CreateDate,
                                ConfirmSalesTime = (DateTime?)null,
                                ConfirmRdTime = (DateTime?)null,
                                Currency,
                                PortOfDelivery,
                                Status = "1",
                                CreateDate,
                                LastModifiedDate,
                                CreateBy = MemberId > 0 ? MemberId : UserId,
                                LastModifiedBy = MemberId > 0 ? MemberId : UserId
                            });

                        var insertResult = sqlConnection.Query(sql, dynamicParameters);
                        var rowsAffected = insertResult.Count();
                        int RfqDetailId = -1;
                        foreach (var innerItem in insertResult)
                        {
                            RfqDetailId = Convert.ToInt32(innerItem.RfqDetailId);
                        }
                        #endregion

                        #region //判斷RFQ包裝種類資料是否正確
                        sql = @"SELECT TOP 1 RfqPkTypeId,PackagingMethod,SustSupplyStatus
                                FROM SCM.RfqPackageType
                                WHERE RfqPkTypeId = @RfqPkTypeId
                                AND RfqProClassId = @RfqProClassId
                                AND Status = @Status";
                        dynamicParameters.Add("RfqPkTypeId", RfqPkTypeId);
                        dynamicParameters.Add("RfqProClassId", RfqProClassId);
                        dynamicParameters.Add("Status", "A");
                        result = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                        if (result == null) throw new SystemException("【包裝種類】資料錯誤!");
                        string SustSupplyStatus = result.SustSupplyStatus;
                        #endregion

                        #region //新增包裝方式
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.RfqPackage (RfqDetailId,RfqPkTypeId,SustSupplyStatus,Status
                                    ,CreateDate,LastModifiedDate,CreateBy,LastModifiedBy)
                                    OUTPUT INSERTED.RfqPkId
                                    VALUES (@RfqDetailId,@RfqPkTypeId,@SustSupplyStatus,@Status
                                    ,@CreateDate,@LastModifiedDate,@CreateBy,@LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RfqDetailId,
                                RfqPkTypeId,
                                SustSupplyStatus,
                                Status = "A",
                                CreateDate,
                                LastModifiedDate,
                                CreateBy = MemberId > 0 ? MemberId : UserId,
                                LastModifiedBy = MemberId > 0 ? MemberId : UserId
                            });
                        insertResult = sqlConnection.Query(sql, dynamicParameters);
                        rowsAffected += insertResult.Count();
                        #endregion

                        #region //新增報價方案
                        RfqLineSolutionData.ForEach(solution =>
                        {
                            dynamicParameters = new DynamicParameters();
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO SCM.RfqLineSolution (RfqDetailId,SortNumber,SolutionQty,PeriodicDemandType
                                    ,CreateDate,LastModifiedDate,CreateBy,LastModifiedBy)
                                    OUTPUT INSERTED.RfqLineSolutionId
                                    VALUES (@RfqDetailId,@SortNumber,@SolutionQty,@PeriodicDemandType
                                    ,@CreateDate,@LastModifiedDate,@CreateBy,@LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    RfqDetailId,
                                    SortNumber = solution.idx,
                                    solution.SolutionQty,
                                    solution.PeriodicDemandType,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy = MemberId > 0 ? MemberId : UserId,
                                    LastModifiedBy = MemberId > 0 ? MemberId : UserId
                                });
                            insertResult = sqlConnection.Query(sql, dynamicParameters);
                            rowsAffected += insertResult.Count();
                        });
                        #endregion

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

        #region //ForgotPassword -- 寄送更換密碼通知信 -- Chia Yuan -- 2023.07.27
        public string ForgotPassword(string Account)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    string currentDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    string MemberEmail = "", MemberName = "", OrgShortName = "";
                    int MemberId = -1;

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷使用者資料是否正確
                        sql = @"SELECT TOP 1 MemberId, MemberEmail, MemberName, OrgShortName
                                FROM EIP.[Member]
                                WHERE MemberEmail = @Account";
                        dynamicParameters.Add("Account", Account);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (!result.Any()) throw new SystemException("【使用者】資料錯誤!");
                        foreach (var item in result)
                        {
                            MemberEmail = item.MemberEmail;
                            MemberName = item.MemberName;
                            OrgShortName = item.OrgShortName;
                            MemberId = item.MemberId;
                        }
                        #endregion

                        #region //刪除過期的LoginKey
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EIP.MemberLoginKey
                                WHERE MemberId = @MemberId
                                AND ExpirationDate <= @Now OR LoginIP = @LoginIP";
                        dynamicParameters.Add("MemberId", MemberId);
                        dynamicParameters.Add("LoginIP", BaseHelper.ClientIP());
                        dynamicParameters.Add("Now", currentDate);
                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //新增金鑰
                        string KeyText = BaseHelper.Sha256Encrypt(Account + DateTime.Now.ToString("yyyyMMdd") + BaseHelper.ClientIP() + DateTime.Now.ToString("HHmmss"));
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EIP.MemberLoginKey (MemberId, KeyText, ExpirationDate, LoginIP
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.KeyId
                                VALUES (@MemberId, @KeyText, @ExpirationDate, @LoginIP
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                MemberId,
                                KeyText,
                                ExpirationDate = DateTime.Now.AddMinutes(15), //mail認證期限為15分鐘
                                LoginIP = BaseHelper.ClientIP(),
                                CreateDate,
                                LastModifiedDate,
                                CreateBy = MemberId,
                                LastModifiedBy = MemberId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);
                        rowsAffected = result.Count();
                        #endregion

                        #region //Mail資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MailFrom, a.MailTo, a.MailCc, a.MailBcc, a.MailSubject, a.MailContent
                                , b.Host, b.Port, b.SendMode, b.Account, b.Password
                                FROM BAS.Mail a
                                LEFT JOIN BAS.MailServer b ON b.ServerId = a.ServerId
                                WHERE a.MailId IN (
                                    SELECT z.MailId
                                    FROM BAS.MailSendSetting z
                                    WHERE z.SettingSchema = @SettingSchema
                                    AND z.SettingNo = @SettingNo)";
                        dynamicParameters.Add("SettingSchema", "RfqMemberPwdForgot");
                        dynamicParameters.Add("SettingNo", "Y");
                        var resultMailTemplate = sqlConnection.Query(sql, dynamicParameters);
                        if (!resultMailTemplate.Any()) throw new SystemException("Mail設定錯誤!");
                        #endregion

                        string serverInfo = GetServerInfo();

                        foreach (var item in resultMailTemplate)
                        {
                            string mailSubject = item.MailSubject;
                            string mailContent = HttpUtility.UrlDecode(item.MailContent);
                            //ForgotUrl = string.Format("<a href=\"http://{0}:16669/Header.html?page=resetpwd&Token={1}\">{2}</a>", serverInfo, KeyText, "點我重設密碼");

                            string validationUrl = HttpContext.Current.Request.UrlReferrer.OriginalString.Replace("forgot", "resetpwd&Token");

                            string ForgotUrl = string.Format("<a href=\"{0}={1}\">{2}</a>", validationUrl, KeyText, "點我重設密碼");

                            #region //Mail內容
                            mailContent = mailContent.Replace("[MemberName]", MemberName);
                            mailContent = mailContent.Replace("[OrgShortName]", OrgShortName);
                            mailContent = mailContent.Replace("[ForgotUrl]", ForgotUrl);
                            #endregion

                            #region 發送
                            MailConfig mailConfig = new MailConfig
                            {
                                Host = item.Host,
                                Port = Convert.ToInt32(item.Port),
                                SendMode = Convert.ToInt32(item.SendMode),
                                From = item.MailFrom,
                                Subject = mailSubject,
                                Account = item.Account,
                                Password = item.Password,
                                MailTo = item.MailTo, //測試用  正式為使用者 MemberEmail
                                MailCc = item.MailCc,
                                MailBcc = item.MailBcc,
                                HtmlBody = mailContent,
                                TextBody = "-"
                            };

                            BaseHelper.MailSend(mailConfig);
                            #endregion
                        }

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

        #region //AddDocRoleInformation -- 單據角色資料新增 -- Xuan 2023.8.24
        public string AddDocRoleInformation(string RoleName)
        {

            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷角色名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EIP.NotifyRole
                                WHERE RoleName = @RoleName";
                        dynamicParameters.Add("RoleName", RoleName);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【角色名稱】已存在，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EIP.NotifyRole (RoleName, Status, CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.RoleId
                                VALUES (@RoleName, @Status, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RoleName,
                                Status = "A", //啟用
                                CreateDate,
                                LastModifiedDate,
                                CreateBy = CurrentUser,
                                LastModifiedBy = CurrentUser
                            });

                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += insertResult.Count();


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

        #region //AddDocUserInformation -- 單據使用者資料新增 -- Xuan 2023.8.28
        public string AddDocUserInformation(int RoleId, int UserId, string UserName, string RoleName)
        {

            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (UserId < 0) throw new SystemException("【人員】不能為空!");
                        #region //判斷角色名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EIP.NotifyRole a
								INNER JOIN EIP.NotifyUser b ON b.RoleId=a.RoleId
								INNER JOIN BAS.[User] c ON c.UserId=b.UserId
                                WHERE a.RoleName = @RoleName AND c.UserName=@UserName";
                        dynamicParameters.Add("RoleName", RoleName);
                        dynamicParameters.Add("UserName", UserName);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【該人員已具有該角色名稱】，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EIP.NotifyUser (RoleId, UserId, CreateDate, CreateBy)
                                OUTPUT INSERTED.RoleId
                                VALUES (@RoleId, @UserId, @CreateDate, @CreateBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RoleId,
                                UserId,
                                CreateDate,
                                CreateBy = CurrentUser
                            });

                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += insertResult.Count();


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

        #region //AddDocInformation -- 單據權限資料新增 -- Xuan 2023.8.29
        public string AddDocInformation(string SubProdType, int SubProdId, int RoleId, string DocType, string RoleName)
        {

            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (SubProdId == -1) throw new SystemException("【產品名稱】不能為空!");
                        if (DocType == "DFM")
                        {
                            if (SubProdType == "") throw new SystemException("【產品類別】不能為空!");
                        }
                       
                        #region //判斷單據權限是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EIP.DocNotify a
                                WHERE a.RoleId = @RoleId AND a.DocType=@DocType AND a.SubProdId=@SubProdId AND a.SubProdType=@SubProdType";
                        dynamicParameters.Add("RoleId", RoleId);
                        dynamicParameters.Add("DocType", DocType);
                        dynamicParameters.Add("SubProdId", SubProdId);
                        dynamicParameters.Add("SubProdType", SubProdType);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【該筆資料已存在】，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO EIP.DocNotify (RoleId, DocType, SubProdId, SubProdType,CreateDate, CreateBy)
                                OUTPUT INSERTED.DocNotifyId, INSERTED.RoleId
                                VALUES (@RoleId, @DocType, @SubProdId, @SubProdType, @CreateDate, @CreateBy)";
                        if(DocType == "DFM")
                        {
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                RoleId,
                                DocType,
                                SubProdId,
                                SubProdType,
                                CreateDate,
                                CreateBy = CurrentUser
                            });
                        }
                        else
                        {
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                RoleId,
                                DocType,
                                SubProdId,
                                SubProdType="NULL",
                                CreateDate,
                                CreateBy = CurrentUser
                            });
                        }

                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += insertResult.Count();


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

        #region //UpdateMember -- 使用者資料更新 -- Chia Yuan  -- 2023.07.28
        public string UpdateMember(string MemberName, string OrgShortName, string Address, string ContactName, string ContactPhone
            , string Account, string KeyText)
        {
            if (string.IsNullOrWhiteSpace(MemberName)) throw new SystemException("【使用者名稱】長度錯誤!");
            if (string.IsNullOrWhiteSpace(OrgShortName)) throw new SystemException("【公司名稱】長度錯誤!");
            if (string.IsNullOrWhiteSpace(ContactPhone)) throw new SystemException("【電話】長度錯誤!");

            //bool telcheck = Regex.IsMatch(ContactPhone, @"^[0-9]$");//規則:09開頭，後面接著8

            ContactName = ContactName.Trim();
            ContactPhone = ContactPhone.Trim();

            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷使用者登入金鑰是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 b.MemberId
                                FROM EIP.MemberLoginKey a
                                INNER JOIN EIP.[Member] b ON a.MemberId = b.MemberId
                                WHERE b.MemberEmail = @Account
                                AND a.KeyText = @KeyText";
                        dynamicParameters.Add("Account", Account);
                        dynamicParameters.Add("KeyText", KeyText);
                        var result = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                        if (result == null) throw new SystemException("【使用者】資料錯誤!");
                        int MemberId = result.MemberId;
                        #endregion

                        #region //使用者資料更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EIP.[Member] SET
                                MemberName = @MemberName,
                                OrgShortName = @OrgShortName,
                                Address = @Address,
                                ContactName = @ContactName,
                                ContactPhone = @ContactPhone,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE MemberId = @MemberId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                MemberId,
                                MemberName,
                                OrgShortName,
                                Address,
                                ContactName,
                                ContactPhone,
                                LastModifiedDate = MemberId,
                                LastModifiedBy = MemberId
                            });
                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //UpdatePasswordReset -- 重設密碼 -- Chia Yuan -- 2023.07.27
        public string UpdatePasswordReset(string Password, string NewPassword, string KeyText)
        {
            if (Password != NewPassword) throw new SystemException("【密碼確認】資料錯誤!");

            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0, MemberId = -1;
                    string currentDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷LoginKey是否正確
                        sql = @"SELECT TOP 1 KeyText, MemberId
                                FROM EIP.MemberLoginKey
                                WHERE KeyText = @KeyText 
                                AND ExpirationDate >= @Now
                                AND LoginIP = @LoginIP";
                        dynamicParameters.Add("KeyText", KeyText);
                        dynamicParameters.Add("Now", currentDate);
                        dynamicParameters.Add("LoginIP", BaseHelper.ClientIP());
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() != 1) throw new SystemException("【使用者】狀態異常!");
                        foreach (var item in result)
                        {
                            MemberId = item.MemberId;
                        }
                        #endregion

                        #region //刪除過期的LoginKey
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EIP.MemberLoginKey
                                WHERE KeyText = @KeyText";
                        dynamicParameters.Add("KeyText", KeyText);
                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //密碼參數設定
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.PasswordExpiration, a.PasswordFormat, a.PasswordWrongCount
                            FROM BAS.PasswordSetting a";
                        var resultPasswordSetting = sqlConnection.Query(sql, dynamicParameters);
                        if (!resultPasswordSetting.Any()) throw new SystemException("【密碼參數設定】資料錯誤!");

                        int PasswordExpiration = -1;
                        string PasswordFormat = "";
                        int PasswordWrongCount = -1;
                        foreach (var item in resultPasswordSetting)
                        {
                            PasswordExpiration = Convert.ToInt32(item.PasswordExpiration);
                            PasswordFormat = item.PasswordFormat;
                            PasswordWrongCount = Convert.ToInt32(item.PasswordWrongCount);
                        }
                        if (!Regex.IsMatch(Password, PasswordFormat, RegexOptions.IgnoreCase)) throw new SystemException("【密碼】格式錯誤!");
                        #endregion

                        #region //密碼更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EIP.[Member] SET
                                Password = @Password,
                                PasswordStatus = @PasswordStatus,
                                PasswordMistake = @PasswordMistake,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE MemberId = @MemberId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Password = BaseHelper.Sha256Encrypt(Password),
                                PasswordStatus = "N",
                                PasswordMistake = 0,
                                LastModifiedDate,
                                LastModifiedBy,
                                MemberId
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

        #region //UpdateMemberAccount -- 密碼更新 -- Chia Yuan -- 2023.07.27
        public string UpdateMemberAccount(string Password, string NewPassword, string ConfirmPassword
            , string Account, string KeyText)
        {
            if (NewPassword != ConfirmPassword) throw new SystemException("【密碼確認】資料錯誤!");
            if (NewPassword == Password) throw new SystemException("【新舊密碼】不能相同!");

            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0, MemberId = -1;
                    string enyPassword = BaseHelper.Sha256Encrypt(Password);
                    string enyNewPassword = BaseHelper.Sha256Encrypt(NewPassword);
                    string currentDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷使用者登入金鑰是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 b.MemberId
                                FROM EIP.MemberLoginKey a
                                INNER JOIN EIP.[Member] b ON a.MemberId = b.MemberId
                                WHERE b.MemberEmail = @Account
                                AND a.KeyText = @KeyText
                                AND b.Password = @Password";
                        dynamicParameters.Add("Account", Account);
                        dynamicParameters.Add("KeyText", KeyText);
                        dynamicParameters.Add("Password", enyPassword);
                        var result = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                        if (result == null) throw new SystemException("【使用者】資料錯誤!");
                        MemberId = result.MemberId;
                        #endregion

                        #region //密碼參數設定
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.PasswordExpiration, a.PasswordFormat, a.PasswordWrongCount
                            FROM BAS.PasswordSetting a";
                        var resultPasswordSetting = sqlConnection.Query(sql, dynamicParameters);
                        if (!resultPasswordSetting.Any()) throw new SystemException("【密碼參數設定】資料錯誤!");

                        int PasswordExpiration = -1;
                        string PasswordFormat = "";
                        int PasswordWrongCount = -1;
                        foreach (var item in resultPasswordSetting)
                        {
                            PasswordExpiration = Convert.ToInt32(item.PasswordExpiration);
                            PasswordFormat = item.PasswordFormat;
                            PasswordWrongCount = Convert.ToInt32(item.PasswordWrongCount);
                        }
                        if (!Regex.IsMatch(NewPassword, PasswordFormat, RegexOptions.IgnoreCase)) throw new SystemException("【密碼】格式錯誤!");
                        #endregion

                        #region //刪除過期的LoginKey
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EIP.MemberLoginKey
                                WHERE KeyText = @KeyText";
                        dynamicParameters.Add("KeyText", KeyText);
                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //密碼更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EIP.[Member] SET
                                Password = @Password,
                                PasswordStatus = @PasswordStatus,
                                PasswordMistake = @PasswordMistake,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE MemberId = @MemberId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Password = enyNewPassword,
                                PasswordStatus = "N",
                                PasswordMistake = 0,
                                LastModifiedDate,
                                LastModifiedBy,
                                MemberId
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

        #region //UpdatePasswordMistake -- 密碼錯誤累積 -- Chia Yuan -- 2023.7.7
        public string UpdatePasswordMistake(int MemberId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷使用者資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EIP.[Member]
                                WHERE MemberId = @MemberId";
                        dynamicParameters.Add("MemberId", MemberId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (!result.Any()) throw new SystemException("【使用者】資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EIP.[Member] SET
                                PasswordMistake = PasswordMistake + 1,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE MemberId = @MemberId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                LastModifiedDate,
                                LastModifiedBy,
                                MemberId
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

        #region //UpdatePasswordMistakeReset -- 密碼錯誤次數重置 -- Chia Yuan -- 2023.7.7
        public string UpdatePasswordMistakeReset(int MemberId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷使用者資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EIP.[Member]
                                WHERE MemberId = @MemberId";
                        dynamicParameters.Add("MemberId", MemberId);

                        var result = sqlConnection.Query(sql, dynamicParameters).FirstOrDefault();
                        if (result == null) throw new SystemException("【使用者】資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EIP.[Member] SET
                                PasswordMistake = @PasswordMistake,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE MemberId = @MemberId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                PasswordMistake = 0,
                                LastModifiedDate,
                                LastModifiedBy,
                                MemberId
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

        #region //UpdateUserLoginKey -- 使用者登入金鑰 -- Chia Yuan -- 2023.7.7
        public string UpdateUserLoginKey(int MemberId, string Account)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;

                        HttpCookie Login = HttpContext.Current.Request.Cookies.Get("Login");
                        HttpCookie LoginKey = HttpContext.Current.Request.Cookies.Get("LoginKey");

                        #region //刪除過期的LoginKey
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EIP.MemberLoginKey
                                WHERE MemberId = @MemberId
                                AND ExpirationDate <= @Now";
                        dynamicParameters.Add("MemberId", MemberId);
                        dynamicParameters.Add("Now", DateTime.Now);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        if (Login == null || LoginKey == null)
                        {
                            string keyText = BaseHelper.Sha256Encrypt(Account + DateTime.Now.ToString("yyyyMMdd") + BaseHelper.ClientIP() + DateTime.Now.ToString("HHmmss"));

                            Login = new HttpCookie("Login")
                            {
                                Value = Account,
                                Expires = DateTime.Now.AddDays(1)
                            };
                            HttpContext.Current.Response.Cookies.Add(Login);

                            LoginKey = new HttpCookie("LoginKey")
                            {
                                Value = keyText,
                                Expires = DateTime.Now.AddDays(1)
                            };
                            HttpContext.Current.Response.Cookies.Add(LoginKey);

                            #region //刪除目前IP的LoginKey
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE EIP.MemberLoginKey
                                    WHERE LoginIP = @LoginIP";
                            dynamicParameters.Add("LoginIP", BaseHelper.ClientIP());

                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //新增金鑰
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO EIP.MemberLoginKey (MemberId, KeyText, ExpirationDate, LoginIP
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.KeyId
                                    VALUES (@MemberId, @KeyText, @ExpirationDate, @LoginIP
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MemberId,
                                    KeyText = keyText,
                                    ExpirationDate = DateTime.Now.AddDays(1),
                                    LoginIP = BaseHelper.ClientIP(),
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected = insertResult.Count();
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

        #region //UpdateRequestForQuotation --RFQ單頭更新 --Chia Yuan -- 2023.7.20
        public string UpdateRequestForQuotation(int RfqId, string AssemblyName, int ProductUseId, string CustomerName, int CustomerId, string Status
            , int MemberType, string Account, string KeyText)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(AssemblyName)) throw new SystemException("【機種名稱】不能為空!");

                AssemblyName = AssemblyName.Trim();

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷RFQ產品應用資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.ProductUse
                                WHERE ProductUseId = @ProductUseId";
                        dynamicParameters.Add("ProductUseId", ProductUseId);
                        var result = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                        if (result == null) throw new SystemException("【產品應用】資料錯誤!");
                        #endregion

                        #region //取得客戶資料
                        int MemberId = -1, UserId = -1;
                        if (MemberType == 1)
                        {
                            #region //判斷使用者登入金鑰是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 b.MemberId
                                FROM EIP.MemberLoginKey a
                                INNER JOIN EIP.[Member] b ON a.MemberId = b.MemberId
                                WHERE b.MemberEmail = @Account
                                AND a.KeyText = @KeyText";
                            dynamicParameters.Add("Account", Account);
                            dynamicParameters.Add("KeyText", KeyText);
                            var userResult = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                            if (userResult == null) throw new SystemException("【使用者】資料錯誤!");
                            MemberId = userResult.MemberId;
                            #endregion

                            #region //更新RFQ單頭
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.RequestForQuotation SET
                                AssemblyName = @AssemblyName,
                                ProductUseId = @ProductUseId,
                                CustomerName = @CustomerName,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RfqId = @RfqId AND Status = '0'";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    RfqId,
                                    AssemblyName,
                                    ProductUseId,
                                    CustomerName = string.IsNullOrWhiteSpace(CustomerName) ? null : CustomerName,
                                    LastModifiedDate,
                                    LastModifiedBy = MemberId
                                });
                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }
                        if (MemberType == 2)
                        {
                            #region //判斷使用者登入金鑰是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 b.UserId
                                FROM BAS.UserLoginKey a
                                INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                                WHERE b.UserNo = @Account
                                AND a.KeyText = @KeyText";
                            dynamicParameters.Add("Account", Account);
                            dynamicParameters.Add("KeyText", KeyText);
                            var userResult = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                            if (userResult == null) throw new SystemException("【使用者】資料錯誤!");
                            UserId = userResult.UserId;
                            #endregion

                            #region //判斷客戶資料是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.CustomerId
                                    FROM SCM.Customer a 
                                    WHERE a.CustomerId = @CustomerId";
                            dynamicParameters.Add("CustomerId", CustomerId);
                            userResult = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                            if (userResult == null) throw new SystemException("【客戶】資料錯誤!");
                            #endregion

                            #region //更新RFQ單頭
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.RequestForQuotation SET
                                AssemblyName = @AssemblyName,
                                ProductUseId = @ProductUseId,
                                CustomerId = @CustomerId,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RfqId = @RfqId AND Status = '0'";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    RfqId,
                                    AssemblyName,
                                    ProductUseId,
                                    CustomerId,
                                    LastModifiedDate,
                                    LastModifiedBy = MemberId > 0 ? MemberId : UserId
                                });
                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
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

        #region //UpdateRfqDetail -- RFQ單身更新 -- Chia Yuan -- 2023.7.21
        public string UpdateRfqDetail(int RfqDetailId, int RfqProTypeId, string MtlName, int PrototypeQty
            , string ProdLifeCycleStart, string ProdLifeCycleEnd, int LifeCycleQty, string DemandDate, string CoatingFlag, int AdditionalFile, string Description
            , int RfqPkTypeId, string RfqLineSolution, string Currency, string PortOfDelivery
            , List<FileModel> Files, string ClientIP, string savePath
            , int MemberType, string Account, string KeyText)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(MtlName)) throw new SystemException("【品名】不能為空!");
                if (string.IsNullOrWhiteSpace(ProdLifeCycleStart)) throw new SystemException("【生命週期(起)】不能為空!");
                if (string.IsNullOrWhiteSpace(ProdLifeCycleEnd)) throw new SystemException("【生命週期(迄)】不能為空!");
                if (string.IsNullOrWhiteSpace(DemandDate)) throw new SystemException("【需求日期】不能為空!");
                if (string.IsNullOrWhiteSpace(CoatingFlag)) throw new SystemException("【是否鍍膜】不能為空!");
                if (string.IsNullOrWhiteSpace(PortOfDelivery)) throw new SystemException("【交貨地點】不能為空!");
                if (Description.Length >= 100) throw new SystemException("【描述】格式錯誤!");
                if (PrototypeQty <= 0) throw new SystemException("【試作需求數量】格式錯誤!");
                if (LifeCycleQty <= 0) throw new SystemException("【週期數量】格式錯誤!");
                if (!DateTime.TryParse(ProdLifeCycleStart, out DateTime prodLifeCycleStart)) throw new SystemException("【生命週期(起)】格式錯誤!");
                if (!DateTime.TryParse(ProdLifeCycleEnd, out DateTime prodLifeCycleEnd)) throw new SystemException("【生命週期(迄)】格式錯誤!");
                if (!DateTime.TryParse(DemandDate, out DateTime demandDate)) throw new SystemException("【需求日期】格式錯誤!");
                if (prodLifeCycleStart.CompareTo(prodLifeCycleEnd) > 0) throw new SystemException("【生命週期】區間錯誤!");

                List<RfqLineSolutionVM> RfqLineSolutionData = JsonConvert.DeserializeObject<List<RfqLineSolutionVM>>(RfqLineSolution);
                if (!RfqLineSolutionData.Any()) throw new SystemException("【報價依據】錯誤!");

                CoatingFlag = CoatingFlag.Trim();

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;
                        #region //取得客戶資料
                        int MemberId = -1, UserId = -1;
                        if (MemberType == 1)
                        {
                            #region //判斷使用者登入金鑰是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 b.MemberId
                                FROM EIP.MemberLoginKey a
                                INNER JOIN EIP.[Member] b ON a.MemberId = b.MemberId
                                WHERE b.MemberEmail = @Account
                                AND a.KeyText = @KeyText";
                            dynamicParameters.Add("Account", Account);
                            dynamicParameters.Add("KeyText", KeyText);
                            var userResult = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                            if (userResult == null) throw new SystemException("【使用者】資料錯誤!");
                            MemberId = userResult.MemberId;
                            #endregion
                        }
                        if (MemberType == 2)
                        {
                            #region //判斷使用者登入金鑰是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 b.UserId
                                FROM BAS.UserLoginKey a
                                INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                                WHERE b.UserNo = @Account
                                AND a.KeyText = @KeyText";
                            dynamicParameters.Add("Account", Account);
                            dynamicParameters.Add("KeyText", KeyText);
                            var userResult = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                            if (userResult == null) throw new SystemException("【使用者】資料錯誤!");
                            UserId = userResult.UserId;
                            #endregion
                        }
                        #endregion

                        #region //判斷RFQ單身資料是否正確
                        sql = @"SELECT TOP 1 ISNULL(b.FileId, -1) AS CustProdDigram
                                FROM SCM.RfqDetail a
                                LEFT JOIN BAS.[File] b ON b.FileId = a.CustProdDigram
                                WHERE RfqDetailId = @RfqDetailId";
                        dynamicParameters.Add("RfqDetailId", RfqDetailId);
                        var result = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                        if (result == null) throw new SystemException("【詢價明細】資料錯誤!");
                        int FileId = result.CustProdDigram;
                        #endregion

                        #region //判斷RFQ產品類別資料是否正確
                        sql = @"SELECT TOP 1 RfqProClassId
                                FROM SCM.RfqProductType
                                WHERE RfqProTypeId = @RfqProTypeId";
                        dynamicParameters.Add("RfqProTypeId", RfqProTypeId);
                        result = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                        if (result == null) throw new SystemException("【產品類別】資料錯誤!");
                        int RfqProClassId = result.RfqProClassId;
                        #endregion

                        #region //判斷RFQ是否度模資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Status]
                                WHERE StatusSchema = 'Boolean' and StatusNo = @CoatingFlag";
                        dynamicParameters.Add("CoatingFlag", CoatingFlag);
                        result = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                        if (result == null) throw new SystemException("【是否鍍膜】資料錯誤!");
                        #endregion

                        #region //判斷AdditionalFile資料是否正確(未完成)

                        #endregion

                        #region //判斷RFQ包裝種類資料是否正確
                        sql = @"SELECT TOP 1 RfqPkTypeId,PackagingMethod,SustSupplyStatus
                                FROM SCM.RfqPackageType
                                WHERE RfqPkTypeId = @RfqPkTypeId
                                AND RfqProClassId = @RfqProClassId
                                AND Status = @Status";
                        dynamicParameters.Add("RfqPkTypeId", RfqPkTypeId);
                        dynamicParameters.Add("RfqProClassId", RfqProClassId);
                        dynamicParameters.Add("Status", "A");
                        result = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                        if (result == null) throw new SystemException("【包裝種類】資料錯誤!");
                        string SustSupplyStatus = result.SustSupplyStatus;
                        #endregion

                        #region 判斷包裝方式是否正確(單筆寫入)
                        sql = @"SELECT TOP 1 RfqPkId
                                FROM SCM.RfqPackage
                                WHERE RfqDetailId = @RfqDetailId";
                        dynamicParameters.Add("RfqDetailId", RfqDetailId);
                        result = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                        int RfqPkId = result == null ? -1 : result.RfqPkId;
                        #endregion

                        #region //包裝方式處理
                        if (RfqPkId < 0)
                        {
                            #region //新增包裝方式
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO SCM.RfqPackage (RfqDetailId,RfqPkTypeId,SustSupplyStatus,Status
                                    ,CreateDate,LastModifiedDate,CreateBy,LastModifiedBy)
                                    OUTPUT INSERTED.RfqPkId
                                    VALUES (@RfqDetailId,@RfqPkTypeId,@SustSupplyStatus,@Status
                                    ,@CreateDate,@LastModifiedDate,@CreateBy,@LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    RfqDetailId,
                                    RfqPkTypeId,
                                    SustSupplyStatus,
                                    Status = "A",
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy = MemberId > 0 ? MemberId : UserId,
                                    LastModifiedBy = MemberId > 0 ? MemberId : UserId
                                });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);
                            rowsAffected += insertResult.Count();
                            #endregion
                        }
                        else
                        {
                            #region //更新包裝方式
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.RfqPackage SET
                                RfqPkTypeId = @RfqPkTypeId,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RfqDetailId = @RfqDetailId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    RfqDetailId,
                                    RfqPkTypeId,
                                    Status = "A",
                                    LastModifiedDate,
                                    LastModifiedBy = MemberId > 0 ? MemberId : UserId
                                });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }
                        #endregion

                        #region 圖檔處理
                        if (FileId < 0)
                        {
                            #region //新增圖檔
                            Files.ForEach(file =>
                            {
                                sql = @"INSERT INTO BAS.[File] (CompanyId, FileName, FileContent, FileExtension, FileSize
                                      , ClientIP, Source, DeleteStatus
                                      , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                      OUTPUT INSERTED.FileId
                                      VALUES (@CompanyId, @FileName, @FileContent, @FileExtension, @FileSize
                                      , @ClientIP, @Source, @DeleteStatus
                                      , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        CompanyId = 2,
                                        file.FileName,
                                        file.FileContent,
                                        file.FileExtension,
                                        file.FileSize,
                                        ClientIP,
                                        Source = savePath,
                                        DeleteStatus = "N",
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy = MemberId > 0 ? MemberId : UserId,
                                        LastModifiedBy = MemberId > 0 ? MemberId : UserId
                                    });
                                var insertResult1 = sqlConnection.Query(sql, dynamicParameters);
                                var rowsAffected1 = insertResult1.Count();
                                foreach (var innerItem in insertResult1)
                                {
                                    FileId = Convert.ToInt32(innerItem.FileId);
                                }
                            });
                            #endregion
                        }
                        else
                        {
                            #region //更新圖檔
                            Files.ForEach(file =>
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE BAS.[File] SET
                                CompanyId = @CompanyId,
                                FileName = @FileName,
                                FileContent = @FileContent,
                                FileExtension = @FileExtension,
                                FileSize = @FileSize,
                                ClientIP = @ClientIP,
                                Source = @Source,
                                DeleteStatus = @DeleteStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE FileId = @FileId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        FileId,
                                        CompanyId = 2,
                                        file.FileName,
                                        file.FileContent,
                                        file.FileExtension,
                                        file.FileSize,
                                        ClientIP,
                                        Source = savePath,
                                        DeleteStatus = "N",
                                        LastModifiedDate,
                                        LastModifiedBy = MemberId > 0 ? MemberId : UserId
                                    });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            });
                            #endregion
                        }
                        #endregion

                        #region //刪除圖檔(需改為多圖處理)
                        //if (FileId.HasValue)
                        //{
                        //    dynamicParameters = new DynamicParameters();
                        //    sql = @"DELETE BAS.[File]
                        //        WHERE FileId = @FileId";
                        //    dynamicParameters.Add("FileId", FileId);
                        //    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        //}
                        #endregion

                        #region //新增圖檔(需改為多圖處理)
                        //FileId = -1;
                        //Files.ForEach(file =>
                        //{
                        //    dynamicParameters = new DynamicParameters();
                        //    sql = @"INSERT INTO BAS.[File] (CompanyId, FileName, FileContent, FileExtension, FileSize
                        //        , ClientIP, Source, DeleteStatus
                        //        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                        //        OUTPUT INSERTED.FileId
                        //        VALUES (@CompanyId, @FileName, @FileContent, @FileExtension, @FileSize
                        //        , @ClientIP, @Source, @DeleteStatus
                        //        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        //    dynamicParameters.AddDynamicParams(
                        //        new
                        //        {
                        //            CompanyId = 2,
                        //            file.FileName,
                        //            file.FileContent,
                        //            file.FileExtension,
                        //            file.FileSize,
                        //            ClientIP,
                        //            Source = savePath,
                        //            DeleteStatus = "N",
                        //            CreateDate,
                        //            LastModifiedDate,
                        //            CreateBy = MemberId,
                        //            LastModifiedBy = MemberId
                        //        });
                        //    var insertResult = sqlConnection.Query(sql, dynamicParameters);
                        //    rowsAffected += insertResult.Count();                            
                        //    foreach (var innerItem in insertResult)
                        //    {
                        //        FileId = Convert.ToInt32(innerItem.FileId);
                        //    }
                        //});
                        //if (FileId < 0) throw new SystemException("【圖片】上傳失敗!");
                        #endregion

                        #region//更新RFQ單身明細
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.RfqDetail SET
                                RfqProTypeId = @RfqProTypeId,
                                MtlName = @MtlName,
                                CustProdDigram = @CustProdDigram,
                                PrototypeQty = @PrototypeQty,
                                ProdLifeCycleStart = @ProdLifeCycleStart,
                                ProdLifeCycleEnd = @ProdLifeCycleEnd,
                                LifeCycleQty = @LifeCycleQty,
                                DemandDate = @DemandDate,
                                CoatingFlag = @CoatingFlag,
                                AdditionalFile = @AdditionalFile,
                                Currency = @Currency,
                                PortOfDelivery = @PortOfDelivery,
                                Description = @Description,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RfqDetailId = @RfqDetailId";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            RfqDetailId,
                            RfqProTypeId,
                            MtlName,
                            CustProdDigram = FileId == -1 ? (int?)null : FileId,
                            PrototypeQty,
                            ProdLifeCycleStart,
                            ProdLifeCycleEnd,
                            LifeCycleQty,
                            DemandDate,
                            CoatingFlag,
                            AdditionalFile = AdditionalFile == -1 ? (int?)null : AdditionalFile,
                            Currency,
                            PortOfDelivery,
                            Description,
                            LastModifiedDate,
                            LastModifiedBy = MemberId > 0 ? MemberId : UserId
                        });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除報價方案
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.RfqLineSolution
                                WHERE RfqDetailId = @RfqDetailId";
                        dynamicParameters.Add("RfqDetailId", RfqDetailId);
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //新增報價方案
                        int idx = 1;
                        RfqLineSolutionData.ForEach(solution =>
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO SCM.RfqLineSolution (RfqDetailId,SortNumber,SolutionQty,PeriodicDemandType
                                    ,CreateDate,LastModifiedDate,CreateBy,LastModifiedBy)
                                    OUTPUT INSERTED.RfqLineSolutionId
                                    VALUES (@RfqDetailId,@SortNumber,@SolutionQty,@PeriodicDemandType
                                    ,@CreateDate,@LastModifiedDate,@CreateBy,@LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    RfqDetailId,
                                    SortNumber = idx,
                                    solution.SolutionQty,
                                    solution.PeriodicDemandType,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy = MemberId > 0 ? MemberId : UserId,
                                    LastModifiedBy = MemberId > 0 ? MemberId : UserId
                                });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);
                            rowsAffected += insertResult.Count();
                            idx++;
                        });
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

        #region //UpdateRequestForQuotationStatus -- RFQ單頭狀態更新 -- Chia Yuan -- 2023.07.21
        public string UpdateRequestForQuotationStatus(int RfqId, string Status
            , int MemberType, string Account, string KeyText)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //取得客戶資料
                        int MemberId = -1, UserId = -1;
                        if (MemberType == 1)
                        {
                            #region //判斷使用者登入金鑰是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 b.MemberId
                                FROM EIP.MemberLoginKey a
                                INNER JOIN EIP.[Member] b ON a.MemberId = b.MemberId
                                WHERE b.MemberEmail = @Account
                                AND a.KeyText = @KeyText";
                            dynamicParameters.Add("Account", Account);
                            dynamicParameters.Add("KeyText", KeyText);
                            var userResult = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                            if (userResult == null) throw new SystemException("【使用者】資料錯誤!");
                            MemberId = userResult.MemberId;
                            #endregion
                        }
                        if (MemberType == 2)
                        {
                            #region //判斷使用者登入金鑰是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 b.UserId
                                FROM BAS.UserLoginKey a
                                INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                                WHERE b.UserNo = @Account
                                AND a.KeyText = @KeyText";
                            dynamicParameters.Add("Account", Account);
                            dynamicParameters.Add("KeyText", KeyText);
                            var userResult = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                            if (userResult == null) throw new SystemException("【使用者】資料錯誤!");
                            UserId = userResult.UserId;
                            #endregion
                        }
                        #endregion

                        #region //判斷單身是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT b.RfqDetailId
                                FROM SCM.RequestForQuotation a
                                INNER JOIN SCM.RfqDetail b on b.RfqId = a.RfqId
                                WHERE a.RfqId = @RfqId";
                        dynamicParameters.Add("RfqId", RfqId);
                        var result = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                        if (result == null) throw new SystemException("尚未建立詢價項目!");
                        #endregion

                        #region //取得RFQ單頭、客戶基本資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.RfqId,a.RfqNo,f.ProductUseName,a.AssemblyName
                                ,ISNULL(e.CustomerId,c.CustomerId) as CustomerId
                                ,ISNULL(ISNULL(ISNULL(e.CustomerName,case when a.CustomerName = '' then null else a.CustomerName end),case when b.OrgShortName = '' then null else b.OrgShortName end), c.CustomerName) as CustomerName
                                ,ISNULL(ISNULL(case when b.ContactEmail = '' then null else b.ContactEmail end,b.MemberEmail),c.Email) as MemberEmail
                                ,ISNULL(b.ContactPhone,c.TelNoFirst) as ContactPhone 
                                FROM SCM.RequestForQuotation a
                                join SCM.ProductUse f on f.ProductUseId = a.ProductUseId
                                LEFT JOIN EIP.[Member] b on b.MemberId = a.MemberId
                                LEFT JOIN EIP.MemberOrganization d on d.MemberId = a.MemberId and d.OrganizaitonType = 1
                                LEFT JOIN SCM.Customer e on e.CustomerId = d.OrganizaitonTypeId
                                LEFT JOIN SCM.Customer c on c.CustomerId = a.CustomerId
                                WHERE a.RfqId = @RfqId";
                        dynamicParameters.Add("RfqId", RfqId);
                        var rfqResult = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                        if (rfqResult == null) throw new SystemException("無法取得詢價單!");
                        #endregion

                        #region //取得RFQ單身
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.RfqDetailId,b.RfqProductTypeName,c.RfqProductClassName,a.MtlName,a.LifeCycleQty,a.PrototypeQty,a.DemandDate,a.SalesId,ISNULL(d.Email, '') as Email
                                FROM SCM.RfqDetail a
                                join SCM.RfqProductType b on b.RfqProTypeId = a.RfqProTypeId
                                join SCM.RfqProductClass c on c.RfqProClassId = b.RfqProClassId
                                LEFT JOIN BAS.[User] d on d.UserId = a.SalesId
                                WHERE a.RfqId = @RfqId";
                        dynamicParameters.Add("RfqId", RfqId);
                        var rfqDetailResult = sqlConnection.Query(sql, dynamicParameters);
                        if (!rfqDetailResult.Any()) throw new SystemException("無法取得詢價項目!");

                        string[] salesMail = rfqDetailResult.Where(w => !string.IsNullOrWhiteSpace((string)w.Email)).Select(s => (string)s.Email).Distinct().ToArray();
                        #endregion

                        #region //更新RFQ單頭
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.RequestForQuotation SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RfqId = @RfqId AND Status = '0'";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RfqId,
                                Status = "1",
                                LastModifiedDate,
                                LastModifiedBy = MemberId > 0 ? MemberId : UserId
                            });
                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //Mail資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MailFrom, a.MailTo, a.MailCc, a.MailBcc, a.MailSubject, a.MailContent
                                , b.Host, b.Port, b.SendMode, b.Account, b.Password
                                FROM BAS.Mail a
                                LEFT JOIN BAS.MailServer b ON b.ServerId = a.ServerId
                                WHERE a.MailId IN (
                                    SELECT z.MailId
                                    FROM BAS.MailSendSetting z
                                    WHERE z.SettingSchema = @SettingSchema
                                    AND z.SettingNo = @SettingNo)";
                        dynamicParameters.Add("SettingSchema", "NewRequestForQuotation");
                        dynamicParameters.Add("SettingNo", "Y");
                        var resultMailTemplate = sqlConnection.Query(sql, dynamicParameters);
                        if (!resultMailTemplate.Any()) throw new SystemException("Mail設定錯誤!");
                        #endregion

                        string serverInfo = GetServerInfo();

                        foreach (var item in resultMailTemplate)
                        {
                            string mailSubject = item.MailSubject;
                            string mailContent = HttpUtility.UrlDecode(item.MailContent);

                            string rfqLink = HttpContext.Current.Request.Url.AbsoluteUri.Replace("api/EIP/UpdateRequestForQuotationStatus", "RequestForQuotation/RfqManagment");
                            rfqLink = string.Format("<a href=\"{0}\">{1}</a>", rfqLink, string.Format("查看客戶詢價單 單號：{0}", rfqResult.RfqNo));

                            #region //Mail內容

                            string contentDetail = "";

                            int i = 1;
                            foreach (var detail in rfqDetailResult)
                            {
                                contentDetail += string.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td></tr>",
                                    i.ToString().PadLeft(2, '0'), detail.RfqProductTypeName, detail.MtlName, detail.LifeCycleQty.ToString("N0"), detail.DemandDate.ToString("yyyy-MM-dd"));
                                i++;
                            }

                            contentDetail = @"<table border>
	                                <thead>
		                                <tr>
			                                <th>清單序號</th><th>產品類別</th><th>品名</th><th>數量</th><th>需求日期</th>
		                                </tr>
	                                </thead>
	                                <tbody>" + contentDetail + "</tbody></table>";

                            mailContent = mailContent.Replace("[CustomerName]", rfqResult.CustomerName);
                            mailContent = mailContent.Replace("[MemberEmail]", rfqResult.MemberEmail);
                            mailContent = mailContent.Replace("[ContactPhone]", rfqResult.ContactPhone);
                            mailContent = mailContent.Replace("[RfqNo]", rfqResult.RfqNo);
                            mailContent = mailContent.Replace("[AssemblyName]", rfqResult.AssemblyName);
                            mailContent = mailContent.Replace("[ProductUseName]", rfqResult.ProductUseName);
                            mailContent = mailContent.Replace("[RFQContent]", contentDetail);
                            mailContent = mailContent.Replace("[RFQLink]", rfqLink);

                            #endregion

                            #region 發送

                            if (rfqDetailResult.Any(a => a.SalesId == null))
                            {
                                MailConfig mailConfig = new MailConfig
                                {
                                    Host = item.Host,
                                    Port = Convert.ToInt32(item.Port),
                                    SendMode = Convert.ToInt32(item.SendMode),
                                    From = item.MailFrom,
                                    Subject = mailSubject,
                                    Account = item.Account,
                                    Password = item.Password,
                                    MailTo = item.MailTo, //測試用  正式為使用者 MemberEmail
                                    MailCc = item.MailCc,
                                    MailBcc = item.MailBcc,
                                    HtmlBody = mailContent,
                                    TextBody = "-"
                                };
                                BaseHelper.MailSend(mailConfig);
                            }
                            foreach (string mailTo in salesMail)
                            {
                                MailConfig mailConfig = new MailConfig
                                {
                                    Host = item.Host,
                                    Port = Convert.ToInt32(item.Port),
                                    SendMode = Convert.ToInt32(item.SendMode),
                                    From = item.MailFrom,
                                    Subject = mailSubject,
                                    Account = item.Account,
                                    Password = item.Password,
                                    MailTo = mailTo, //測試用  正式為使用者 MemberEmail
                                    MailCc = item.MailCc,
                                    MailBcc = item.MailBcc,
                                    HtmlBody = mailContent,
                                    TextBody = "-"
                                };
                                BaseHelper.MailSend(mailConfig);
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

        #region //UpdateDocRoleInformation -- 單據角色資料更新 -- Xuan  -- 2023.08.24
        public string UpdateDocRoleInformation(int RoleId, string RoleName)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷角色ID是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.RoleId
                                FROM EIP.NotifyRole a
                                WHERE a.RoleId = @RoleId";
                        dynamicParameters.Add("RoleId", RoleId);
                        var result = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                        if (result == null) throw new SystemException("【角色ID】不存在!");
                        #endregion

                        #region //使用者資料更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE EIP.NotifyRole SET
                                RoleName=@RoleName
                                ,LastModifiedDate=@LastModifiedDate
                                ,LastModifiedBy=@LastModifiedBy
                                WHERE RoleId=@RoleId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RoleName,
                                RoleId,
                                LastModifiedDate,
                                LastModifiedBy
                            });
                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //UpdateDocRoleStatus -- 單據角色狀態更新 -- Xuan 2023.08.24
        public string UpdateDocRoleStatus(int RoleId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷餐點資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM EIP.NotifyRole
                                WHERE RoleId = @RoleId";
                        dynamicParameters.Add("RoleId", RoleId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【單據】資料錯誤!");

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
                        sql = @"UPDATE EIP.NotifyRole SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RoleId = @RoleId";
                        var parametersObject = new
                        {
                            Status = status,
                            LastModifiedDate,
                            LastModifiedBy,
                            RoleId
                        };
                        dynamicParameters.AddDynamicParams(parametersObject);

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

        #region //DeleteUserLoginKeyFromEIP -- 使用者登入金鑰刪除 -- Chia Yuan -- 2023.7.10
        public string DeleteUserLoginKeyFromEIP(string Account, string KeyText)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(Account)) throw new SystemException("【信箱】錯誤!");
                if (string.IsNullOrWhiteSpace(KeyText)) throw new SystemException("【金鑰】錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;
                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM EIP.MemberLoginKey a
                                INNER JOIN EIP.[Member] b ON a.MemberId = b.MemberId
                                WHERE b.MemberEmail = @Account
                                AND a.KeyText = @KeyText";
                        dynamicParameters.Add("Account", Account);
                        dynamicParameters.Add("KeyText", KeyText);

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

        #region //DeleteUserLoginKeyFromBM -- 使用者登入金鑰刪除 -- Chia Yuan -- 2023.7.10
        public string DeleteUserLoginKeyFromBM(string Account, string KeyText)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(Account)) throw new SystemException("【帳號】錯誤!");
                if (string.IsNullOrWhiteSpace(KeyText)) throw new SystemException("【金鑰】錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;
                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM BAS.UserLoginKey a
                                INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                                WHERE b.UserNo = @Account
                                AND a.KeyText = @KeyText";
                        dynamicParameters.Add("Account", Account);
                        dynamicParameters.Add("KeyText", KeyText);

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

        #region DeleteRequestForQuotation -- RFQ單頭刪除 -- Chia Yuan -- 2023.7.22
        public string DeleteRequestForQuotation(int RfqId
            , int MemberType, string Account, string KeyText)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;
                        #region //取得客戶資料
                        int MemberId = -1, UserId = -1;
                        if (MemberType == 1)
                        {
                            #region //判斷使用者登入金鑰是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 b.MemberId
                                FROM EIP.MemberLoginKey a
                                INNER JOIN EIP.[Member] b ON a.MemberId = b.MemberId
                                WHERE b.MemberEmail = @Account
                                AND a.KeyText = @KeyText";
                            dynamicParameters.Add("Account", Account);
                            dynamicParameters.Add("KeyText", KeyText);
                            var userResult = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                            if (userResult == null) throw new SystemException("【使用者】資料錯誤!");
                            MemberId = userResult.MemberId;
                            #endregion
                        }
                        if (MemberType == 2)
                        {
                            #region //判斷使用者登入金鑰是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 b.UserId
                                FROM BAS.UserLoginKey a
                                INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                                WHERE b.UserNo = @Account
                                AND a.KeyText = @KeyText";
                            dynamicParameters.Add("Account", Account);
                            dynamicParameters.Add("KeyText", KeyText);
                            var userResult = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                            if (userResult == null) throw new SystemException("【使用者】資料錯誤!");
                            UserId = userResult.UserId;
                            #endregion
                        }
                        #endregion

                        #region //判斷RFQ單頭資料是否正確
                        sql = @"SELECT DISTINCT a.RfqId, b.RfqDetailId, b.CustProdDigram, c.RfqLineSolutionId, d.RfqPkId
                                FROM SCM.RequestForQuotation a
                                LEFT JOIN SCM.RfqDetail b ON b.RfqId = a.RfqId
                                LEFT JOIN SCM.RfqLineSolution c ON c.RfqDetailId = b.RfqDetailId
                                LEFT JOIN SCM.RfqPackage d ON d.RfqDetailId = c.RfqDetailId
                                WHERE a.RfqId = @RfqId";
                        dynamicParameters.Add("RfqId", RfqId);
                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (!result2.Any()) throw new SystemException("【詢價單】資料錯誤!");
                        #endregion

                        foreach (var item in result2)
                        {
                            int? RfqDetailId = item.RfqDetailId;
                            if (RfqDetailId != null && RfqDetailId > 0)
                            {
                                #region 刪除報價方案
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE a FROM SCM.RfqLineSolution a WHERE a.RfqDetailId = @RfqDetailId";
                                dynamicParameters.Add("RfqDetailId", RfqDetailId);
                                rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region 刪除包裝方式
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE a FROM SCM.RfqPackage a WHERE a.RfqDetailId = @RfqDetailId";
                                dynamicParameters.Add("RfqDetailId", RfqDetailId);
                                rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region 刪除RFQ單身
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE a FROM SCM.RfqDetail a WHERE a.RfqDetailId = @RfqDetailId";
                                dynamicParameters.Add("RfqDetailId", RfqDetailId);
                                rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region 刪除圖檔
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE a FROM BAS.[File] a WHERE a.FileId = @FileId";
                                dynamicParameters.Add("FileId", item.CustProdDigram);
                                rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }
                        }

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM SCM.RequestForQuotation a
                                WHERE a.RfqId = @RfqId";
                        dynamicParameters.Add("RfqId", RfqId);
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

        #region //DeleteRfqDetail -- RFQ單身刪除 -- Chia Yuan -- 2023.7.21
        public string DeleteRfqDetail(int RfqDetailId
            , int MemberType, string Account, string KeyText)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;
                        #region //取得客戶資料
                        int MemberId = -1, UserId = -1;
                        if (MemberType == 1)
                        {
                            #region //判斷使用者登入金鑰是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 b.MemberId
                                FROM EIP.MemberLoginKey a
                                INNER JOIN EIP.[Member] b ON a.MemberId = b.MemberId
                                WHERE b.MemberEmail = @Account
                                AND a.KeyText = @KeyText";
                            dynamicParameters.Add("Account", Account);
                            dynamicParameters.Add("KeyText", KeyText);
                            var userResult = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                            if (userResult == null) throw new SystemException("【使用者】資料錯誤!");
                            MemberId = userResult.MemberId;
                            #endregion
                        }
                        if (MemberType == 2)
                        {
                            #region //判斷使用者登入金鑰是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 b.UserId
                                FROM BAS.UserLoginKey a
                                INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                                WHERE b.UserNo = @Account
                                AND a.KeyText = @KeyText";
                            dynamicParameters.Add("Account", Account);
                            dynamicParameters.Add("KeyText", KeyText);
                            var userResult = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                            if (userResult == null) throw new SystemException("【使用者】資料錯誤!");
                            UserId = userResult.UserId;
                            #endregion
                        }
                        #endregion

                        #region //判斷RFQ單身資料是否正確
                        sql = @"SELECT DISTINCT a.RfqDetailId,a.CustProdDigram,b.RfqLineSolutionId,c.RfqPkId
                                FROM SCM.RfqDetail a
                                JOIN SCM.RfqLineSolution b on b.RfqDetailId = a.RfqDetailId
                                JOIN SCM.RfqPackage c on c.RfqDetailId = a.RfqDetailId
                                JOIN SCM.RequestForQuotation d ON d.RfqId = a.RfqId
                                WHERE a.RfqDetailId = @RfqDetailId";
                        dynamicParameters.Add("RfqDetailId", RfqDetailId);
                        var result = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters);
                        if (result == null) throw new SystemException("【詢價明細】資料錯誤!");
                        int CustProdDigram = result.CustProdDigram;
                        #endregion

                        #region 刪除報價方案
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a FROM SCM.RfqLineSolution a WHERE a.RfqDetailId = @RfqDetailId";
                        dynamicParameters.Add("RfqDetailId", RfqDetailId);
                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region 刪除包裝方式
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a FROM SCM.RfqPackage a WHERE a.RfqDetailId = @RfqDetailId";
                        dynamicParameters.Add("RfqDetailId", RfqDetailId);
                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM SCM.RfqDetail a
                                JOIN SCM.RequestForQuotation b on b.RfqId = a.RfqId
                                WHERE a.RfqDetailId = @RfqDetailId";
                        dynamicParameters.Add("RfqDetailId", RfqDetailId);
                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                        #region 刪除圖檔
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a FROM BAS.[File] a WHERE a.FileId = @FileId";
                        dynamicParameters.Add("FileId", CustProdDigram);
                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion
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

        #region //DeleteDocRoleInformation -- 單據角色資料刪除 -- Xuan -- 2023.8.24
        public string DeleteDocRoleInformation(int RoleId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷角色ID是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EIP.NotifyRole
                                WHERE RoleId = @RoleId";
                        dynamicParameters.Add("RoleId", RoleId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() < 0) throw new SystemException("【角色ID不存在】!");
                        #endregion

                        #region //判斷是否繫結[EIP].[NotifyUser]
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EIP.NotifyUser
                                WHERE RoleId = @RoleId";
                        dynamicParameters.Add("RoleId", RoleId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【該角色已被使用者使用】!");
                        #endregion

                        #region //判斷是否繫結[EIP].[DocNotify]
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EIP.DocNotify
                                WHERE RoleId = @RoleId";
                        dynamicParameters.Add("RoleId", RoleId);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() > 0) throw new SystemException("【該角色已被產品類別使用】!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EIP.NotifyRole
								WHERE RoleId = @RoleId";
                        dynamicParameters.Add("RoleId", RoleId);

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

        #region //DeleteDocUserInformation -- 單據人員資料刪除 -- Xuan -- 2023.8.28
        public string DeleteDocUserInformation(int UserId, int RoleId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷該筆資料是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EIP.NotifyUser a
								WHERE a.RoleId=@RoleId AND a.UserId=@UserId";
                        dynamicParameters.Add("RoleId", RoleId);
                        dynamicParameters.Add("UserId", UserId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() < 0) throw new SystemException("【資料不存在】!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EIP.NotifyUser
								WHERE RoleId = @RoleId AND UserId=@UserId";
                        dynamicParameters.Add("RoleId", RoleId);
                        dynamicParameters.Add("UserId", UserId);

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

        #region //DeleteDocInformation -- 單據權限資料刪除 -- Xuan -- 2023.8.29
        public string DeleteDocInformation(int DocNotifyId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷該筆資料是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM EIP.DocNotify a
                                WHERE a.DocNotifyId=@DocNotifyId";
                        dynamicParameters.Add("DocNotifyId", DocNotifyId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() < 0) throw new SystemException("【資料不存在】!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE EIP.DocNotify
                                WHERE DocNotifyId=@DocNotifyId";
                        dynamicParameters.Add("DocNotifyId", DocNotifyId);

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
