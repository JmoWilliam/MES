using Dapper;
using Helpers;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Transactions;
using System.Web;

namespace BASDA
{
    public class BasicInformationDA
    {
        public string MainConnectionStrings = "";
        public string ErpConnectionStrings = "";
        public string ErpSysConnectionStrings = "";
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
        public MamoHelper mamoHelper = new MamoHelper();

        public BasicInformationDA()
        {
            MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];
            ErpSysConnectionStrings = ConfigurationManager.AppSettings["ErpSysDb"];
            HrmConnectionStrings = ConfigurationManager.AppSettings["HrmDb"];

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
        #region //GetLogin -- 取得登入者資訊 -- Ben Ma 2022.04.01
        public string GetLogin(string Account, string Password)
        {
            try
            {
                if (Account.Length <= 0) throw new SystemException("【帳號】不能為空!");
                if (Password.Length <= 0) throw new SystemException("【密碼】不能為空!");

                IEnumerable<dynamic> resultUser;
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    bool adVerify = false, classicVerify = false;

                    #region //判斷是否要用AD登入驗證
                    string companyNo = "";

                    switch (true)
                    {
                        case bool _ when Regex.IsMatch(Account, @"^(DC|DD|ZY)[0-9]{5}$"): //中揚
                            adVerify = true;
                            companyNo = "JMO";
                            break;
                        case bool _ when Regex.IsMatch(Account, @"^E[0-9]{4}$"): //紘立
                            adVerify = true;
                            companyNo = "Eterge";
                            break;
                        default:
                            classicVerify = true;
                            break;
                    }
                    #endregion

                    adVerify = false;
                    classicVerify = true;

                    if (adVerify)
                    {
                        #region //AD主機資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT c.ConnectionServer, c.ConnectionAccount, c.ConnectionPassword
                                FROM BAS.Interfacing a
                                INNER JOIN BAS.Company b ON a.CompanyId = b.CompanyId
                                OUTER APPLY (
                                    SELECT TOP 1 ca.ConnectionServer, ca.ConnectionAccount, ca.ConnectionPassword
                                    FROM BAS.InterfacingConnection ca
                                    WHERE ca.InterfacingId = a.InterfacingId
                                    AND ca.[Status] = @Status
                                ) c
                                WHERE a.InterfacingType = @InterfacingType
                                AND b.CompanyNo = @CompanyNo";
                        dynamicParameters.Add("Status", "A");
                        dynamicParameters.Add("InterfacingType", "AD_Main");
                        dynamicParameters.Add("CompanyNo", companyNo);

                        var resultInterfacing = sqlConnection.Query(sql, dynamicParameters);
                        if (resultInterfacing.Count() <= 0) throw new SystemException("AD主機資料錯誤!");

                        string connectionServer = "", connectionAccount = "", connectionPassword = "";
                        foreach (var item in resultInterfacing)
                        {
                            connectionServer = item.ConnectionServer;
                            connectionAccount = item.connectionAccount;
                            connectionPassword = item.connectionPassword;
                        }
                        #endregion

                        #region //檢驗AD帳號狀態
                        using (PrincipalContext principalContext = new PrincipalContext(ContextType.Domain, connectionServer, connectionAccount, connectionPassword))
                        {
                            using (UserPrincipal userPrincipal = UserPrincipal.FindByIdentity(principalContext, IdentityType.SamAccountName, Account))
                            {
                                if (userPrincipal == null)
                                {
                                    classicVerify = true;
                                }
                                else
                                {
                                    if (!(bool)userPrincipal.Enabled) throw new Exception("AD帳號已關閉!");
                                    if (userPrincipal.IsAccountLockedOut()) throw new Exception("AD帳號已鎖定!");
                                }
                            }
                        }
                        #endregion

                        if (!classicVerify)
                        {
                            #region //檢驗AD帳號密碼
                            using (DirectoryEntry directoryEntry = new DirectoryEntry(string.Format("LDAP://{0}", connectionServer), Account, Password))
                            {
                                DirectorySearcher directorySearcher = new DirectorySearcher(directoryEntry)
                                {
                                    Filter = string.Format("SAMAccountName={0}", Account)
                                };

                                SearchResult searchResult = directorySearcher.FindOne();

                                if (searchResult != null)
                                {
                                    #region //取得使用者資料
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.UserId, 'N' PasswordStatus, DATEADD(day, 1, GETDATE()) PasswordExpire
                                            FROM BAS.[User] a
                                            WHERE a.Status = @Status
                                            AND a.UserNo = @Account";
                                    dynamicParameters.Add("Status", "A");
                                    dynamicParameters.Add("Account", Account);

                                    resultUser = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultUser.Count() <= 0) throw new SystemException("使用者資料錯誤");
                                    #endregion

                                    #region //Response
                                    jsonResponse = JObject.FromObject(new
                                    {
                                        status = "success",
                                        data = resultUser
                                    });
                                    #endregion
                                }
                            }
                            #endregion
                        }
                    }
                    
                    if (classicVerify)
                    {
                        #region //密碼參數設定
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.PasswordExpiration, a.PasswordFormat, a.PasswordWrongCount
                                FROM BAS.PasswordSetting a";

                        var resultPasswordSetting = sqlConnection.Query(sql, dynamicParameters);
                        if (resultPasswordSetting.Count() <= 0) throw new SystemException("【密碼參數設定】資料錯誤!");

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

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserId, a.UserName, a.Email, a.Password, a.PasswordStatus, a.PasswordExpire, a.PasswordMistake, a.Status
                                , c.CompanyId, c.CompanyName
                                FROM BAS.[User] a
                                INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                                INNER JOIN BAS.Company c ON b.CompanyId = c.CompanyId
                                WHERE a.Status = @Status
                                AND b.Status = @Status
                                AND c.Status = @Status
                                AND a.UserNo = @Account";
                        dynamicParameters.Add("Status", "A");
                        dynamicParameters.Add("Account", Account);

                        resultUser = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUser.Count() <= 0) throw new SystemException("查無此帳號資料：" + Account);

                        foreach (var item in resultUser)
                        {
                            if (Convert.ToInt32(item.PasswordMistake) >= PasswordWrongCount) throw new SystemException(string.Format("密碼錯誤已達{0}次，帳號鎖定!", PasswordWrongCount));

                            if (item.PasswordStatus != "Y")
                            {
                                if (!Regex.IsMatch(Password, PasswordFormat, RegexOptions.IgnoreCase)) throw new SystemException("密碼格式錯誤!");
                            }

                            if (item.Password != BaseHelper.Sha256Encrypt(Password))
                            {
                                UpdatePasswordMistake(Convert.ToInt32(item.UserId));

                                throw new SystemException("登入密碼錯誤：" + Account);
                            }
                            else
                            {
                                UpdatePasswordMistakeReset(Convert.ToInt32(item.UserId));
                            }
                        }

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            data = resultUser
                        });
                        #endregion
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

        #region //GetSubSystemLogin -- 取得子系統登入者資訊 -- Ben Ma 2022.11.23
        public string GetSubSystemLogin(string SystemKey, string Account, string SubSystemCode)
        {
            try
            {
                if (SystemKey.Length <= 0) throw new SystemException("【	金鑰】不能為空!");
                if (Account.Length <= 0) throw new SystemException("【帳號】不能為空!");
                if (SubSystemCode.Length <= 0) throw new SystemException("【子系統代碼】不能為空!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.UserId, a.UserName, a.Email
                            , c.CompanyId, c.CompanyName
                            FROM BAS.[User] a
                            INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                            INNER JOIN BAS.Company c ON b.CompanyId = c.CompanyId
                            INNER JOIN BAS.SubSystemUser d ON a.UserId = d.UserId
                            INNER JOIN BAS.SubSystem e ON d.SubSystemId = e.SubSystemId
                            WHERE 1=1
                            AND a.[Status] = @Status
                            AND b.[Status] = @Status
                            AND c.[Status] = @Status
                            AND d.[Status] = @Status
                            AND e.[Status] = @Status
                            AND e.KeyText = @SystemKey
                            AND a.UserNo = @Account
                            AND e.SubSystemCode = @SubSystemCode";
                    dynamicParameters.Add("Status", "A");
                    dynamicParameters.Add("SystemKey", SystemKey);
                    dynamicParameters.Add("Account", Account);
                    dynamicParameters.Add("SubSystemCode", SubSystemCode);

                    var result = sqlConnection.Query(sql, dynamicParameters);

                    if (result.Count() <= 0) throw new SystemException("子系統查無此帳號資料：" + Account);

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

        #region //GetLoginByKey -- 取得登入者資訊(cookie) -- Ben Ma 2023.03.07
        public string GetLoginByKey(string UserNo, string KeyText)
        {
            try
            {
                if (UserNo.Length <= 0) throw new SystemException("【使用者編號】錯誤!");
                if (KeyText.Length <= 0) throw new SystemException("【金鑰】錯誤!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    //sql = @"SELECT a.KeyId, a.KeyText
                    //        , b.UserId, b.UserNo, b.UserName
                    //        , c.DepartmentId, c.DepartmentNo, c.DepartmentName
                    //        , d.CompanyId, d.CompanyNo, d.CompanyName
                    //        FROM BAS.UserLoginKey a
                    //        INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                    //        INNER JOIN BAS.Department c ON b.DepartmentId = c.DepartmentId
                    //        INNER JOIN BAS.Company d ON c.CompanyId = d.CompanyId
                    //        WHERE 1=1
                    //        AND b.[Status] = @Status
                    //        AND c.[Status] = @Status
                    //        AND d.[Status] = @Status
                    //        AND a.LoginIP = @LoginIP
                    //        AND a.KeyText = @KeyText
                    //        AND b.UserNo = @UserNo";
                    sql = @"SELECT a.KeyId, a.KeyText
                            , b.UserId, b.UserNo, b.UserName
                            , c.DepartmentId, c.DepartmentNo, c.DepartmentName
                            , d.CompanyId, d.CompanyNo, d.CompanyName
                            FROM BAS.UserLoginKey a
                            INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                            INNER JOIN BAS.Department c ON b.DepartmentId = c.DepartmentId
                            INNER JOIN BAS.Company d ON c.CompanyId = d.CompanyId
                            WHERE 1=1
                            AND b.[Status] = @Status
                            AND c.[Status] = @Status
                            AND d.[Status] = @Status
                            AND a.KeyText = @KeyText
                            AND b.UserNo = @UserNo";
                    dynamicParameters.Add("Status", "A");
                    //dynamicParameters.Add("LoginIP", BaseHelper.ClientIP());
                    dynamicParameters.Add("KeyText", KeyText);
                    dynamicParameters.Add("UserNo", UserNo);
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

        #region //GetLoginIPCheck -- 取得登入者最近IP -- Ben Ma 2023.06.07
        public string GetLoginIPCheck(string UserNo)
        {
            try
            {
                if (UserNo.Length <= 0) throw new SystemException("【使用者編號】錯誤!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT TOP 1 b.UserNo, a.LoginIP
                            FROM BAS.UserLoginKey a
                            INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                            WHERE 1=1
                            AND b.[Status] = @Status
                            AND b.UserNo = @UserNo
                            ORDER BY a.CreateDate DESC";
                    dynamicParameters.Add("Status", "A");
                    dynamicParameters.Add("UserNo", UserNo);
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

        #region //GetStatus -- 取得狀態資料 -- Ben Ma 2022.04.08
        public string GetStatus(string StatusSchema, string StatusNo, string StatusName
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.StatusId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.StatusSchema, a.StatusNo, a.StatusName, a.StatusDesc
                        , ISNULL(b.StatusName, a.StatusName) ApplyStatusName, ISNULL(b.StatusDesc, a.StatusDesc) ApplyStatusDesc";
                    sqlQuery.mainTables =
                        @"FROM BAS.[Status] a
                        OUTER APPLY (
                            SELECT ba.StatusName, ba.StatusDesc
                            FROM BAS.StatusCompany ba
                            WHERE ba.StatusId = a.StatusId
                            AND ba.CompanyId = @CompanyId
                        ) b";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StatusSchema", @" AND a.StatusSchema = @StatusSchema", StatusSchema);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StatusNo", @" AND a.StatusNo = @StatusNo", StatusNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StatusName", @" AND (a.StatusName LIKE '%' + @StatusName + '%' OR a.StatusDesc LIKE '%' + @StatusName + '%')", StatusName);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.StatusSchema, a.StatusNo";
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

        #region //GetStatusSchema -- 取得狀態綱要 -- Zoey 2022.05.18
        public string GetStatusSchema(string StatusSchema)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT DISTINCT a.StatusSchema
                            FROM BAS.[Status] a";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "StatusSchema", @" AND a.StatusSchema = @StatusSchema", StatusSchema);

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

        #region //GetStatusCompany -- 取得狀態公司對應資料 -- Ben Ma 2022.10.17
        public string GetStatusCompany(int StatusId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.CompanyId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.CompanyNo, a.CompanyName, a.LogoIcon
                        , ISNULL(b.StatusName, '') StatusName, ISNULL(b.StatusDesc, '') StatusDesc";
                    sqlQuery.mainTables =
                        @"FROM BAS.Company a
                        OUTER APPLY (
                            SELECT ba.StatusName, ba.StatusDesc
                            FROM BAS.StatusCompany ba
                            WHERE ba.CompanyId = a.CompanyId
                            AND ba.StatusId = @StatusId
                        ) b";
                    sqlQuery.auxTables = "";
                    dynamicParameters.Add("StatusId", StatusId);
                    string queryCondition = "";
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = "a.CompanyId";

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

        #region //GetType -- 取得類別資料 -- Ben Ma 2022.04.11
        public string GetType(string TypeSchema, string TypeNo, string TypeName
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.TypeId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.TypeSchema, a.TypeNo, a.TypeName, a.TypeDesc
                        , ISNULL(b.TypeName, a.TypeName) ApplyTypeName, ISNULL(b.TypeDesc, a.TypeDesc) ApplyTypeDesc";
                    sqlQuery.mainTables =
                        @"FROM BAS.[Type] a
                        OUTER APPLY (
                            SELECT ba.TypeName, ba.TypeDesc
                            FROM BAS.TypeCompany ba
                            WHERE ba.TypeId = a.TypeId
                            AND ba.CompanyId = @CompanyId
                        ) b";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TypeSchema", @" AND a.TypeSchema = @TypeSchema", TypeSchema);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TypeNo", @" AND a.TypeNo = @TypeNo", TypeNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TypeName", @" AND (a.TypeName LIKE '%' + @TypeName + '%' OR a.TypeDesc LIKE '%' + @TypeName + '%')", TypeName);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.TypeSchema, a.TypeNo";
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

        #region //GetTypeSchema -- 取得類別綱要 -- Zoey 2022.05.18
        public string GetTypeSchema(string TypeSchema
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT DISTINCT a.TypeSchema
                            FROM BAS.[Type] a";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TypeSchema", @" AND a.TypeSchema = @TypeSchema", TypeSchema);

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

        #region //GetTypeCompany -- 取得類別公司對應資料 -- Ben Ma 2022.10.17
        public string GetTypeCompany(int TypeId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.CompanyId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.CompanyNo, a.CompanyName, a.LogoIcon
                        , ISNULL(b.TypeName, '') TypeName, ISNULL(b.TypeDesc, '') TypeDesc";
                    sqlQuery.mainTables =
                        @"FROM BAS.Company a
                        OUTER APPLY (
                            SELECT ba.TypeName, ba.TypeDesc
                            FROM BAS.TypeCompany ba
                            WHERE ba.CompanyId = a.CompanyId
                            AND ba.TypeId = @TypeId
                        ) b";
                    sqlQuery.auxTables = "";
                    dynamicParameters.Add("TypeId", TypeId);
                    string queryCondition = "";
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = "a.CompanyId";

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

        #region //GetCompany -- 取得公司資料 -- Ben Ma 2022.04.11
        public string GetCompany(int CompanyId, string CompanyNo, string CompanyName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.CompanyId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.CompanyNo, a.CompanyName, a.Telephone, a.Fax, a.Address
                        , a.AddressEn, ISNULL(a.LogoIcon, -1) LogoIcon, a.Status
                        , a.CompanyNo + ' ' + a.CompanyName CompanyWithNo";
                    sqlQuery.mainTables =
                        @"FROM BAS.Company a";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CompanyId", @" AND a.CompanyId = @CompanyId", CompanyId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CompanyNo", @" AND a.CompanyNo LIKE '%' + @CompanyNo + '%'", CompanyNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CompanyName", @" AND a.CompanyName LIKE '%' + @CompanyName + '%'", CompanyName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.CompanyId";
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

        #region //GetDepartment -- 取得部門資料 -- Ben Ma 2022.04.11
        public string GetDepartment(string CompanyNo, int DepartmentId, int CompanyId, string DepartmentNo, string DepartmentName, string Status
            , string SearchKey, string SearchKeys
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.DepartmentId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.DepartmentNo, a.DepartmentName, a.Status
                        , b.CompanyId, b.CompanyNo, b.CompanyName, ISNULL(b.LogoIcon, -1) LogoIcon
                        , a.DepartmentNo + ' ' + a.DepartmentName DepartmentWithNo";
                    sqlQuery.mainTables =
                        @"FROM BAS.Department a
                        INNER JOIN BAS.Company b ON a.CompanyId = b.CompanyId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DepartmentId", @" AND a.DepartmentId = @DepartmentId", DepartmentId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CompanyId", @" AND b.CompanyId = @CompanyId", CompanyId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DepartmentNo", @" AND a.DepartmentNo LIKE '%' + @DepartmentNo + '%'", DepartmentNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DepartmentName", @" AND a.DepartmentName LIKE '%' + @DepartmentName + '%'", DepartmentName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CompanyNo", @" AND b.CompanyNo = @CompanyNo", CompanyNo);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    if (SearchKeys.Length > 0) 
                    {
                        int i = 0;
                        queryCondition += " AND (";
                        SearchKeys.Split(',').ToList().ForEach(x =>
                        {
                            if (i > 0) queryCondition += " OR";
                            BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, $"x{i}",
                                $@" a.DepartmentName LIKE N'%' + @x{i} + '%' OR a.DepartmentNo LIKE N'%' + @x{i} + '%'", x);
                            i++;
                        });
                        queryCondition += ")";
                    }
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey",
                        @" AND (a.DepartmentName LIKE N'%' + @SearchKey + '%'
                            OR a.DepartmentNo LIKE N'%' + @SearchKey + '%')", SearchKey);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "b.CompanyId, a.DepartmentNo";
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

        #region //GetDepartmentShift -- 取得部門班次資料 -- Zoey 2022.05.20
        public string GetDepartmentShift(int DepartmentId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.DepartmentId, a.ShiftId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", c.ShiftName, c.WorkBeginTime, c.WorkEndTime";
                    sqlQuery.mainTables =
                        @"FROM BAS.DepartmentShift a
                        LEFT JOIN BAS.Department b ON b.DepartmentId = a.DepartmentId
                        LEFT JOIN BAS.Shift c ON c.ShiftId = a.ShiftId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DepartmentId", @" AND a.DepartmentId = @DepartmentId", DepartmentId);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.DepartmentId, a.ShiftId";
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

        #region //GetDepartmentRate -- 取得部門費率資料 -- Chia Yuan 2023.6.26

        public string GetDepartmentRate(int DepartmentId, int AuthorId, string Status
            , string StartEnableDate, string EndEnableDate
            , string StartDisabledDate, string EndDisabledDate
            , string StartCreateDate, string EndCreateDate
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.DepartmentId, a.DepartmentRateId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
@", a.ResourceRate, a.OverheadRate
, case when a.EnableDate is null then 'Y' else 'N' end as CaanEdit
, ISNULL(CONVERT(varchar(20), a.EnableDate, 23), '') as EnableDate
, ISNULL(CONVERT(varchar(20), a.DisabledDate, 23), '') as DisabledDate
, ISNULL(CONVERT(varchar(20), a.CreateDate, 23), '') as CreateDate
, ISNULL(CONVERT(varchar(20), a.LastModifiedDate, 23), '') as LastModifiedDate
, a.[Status], c.UserNo, c.UserName as Author, d.UserName, b.DepartmentNo, b.DepartmentName, e.StatusName";
                    sqlQuery.mainTables =
@"from BAS.DepartmentRate a
join BAS.Department b on b.DepartmentId = a.DepartmentId
join BAS.[User] c on c.UserId = a.Author
join BAS.[User] d on d.UserId = a.CreateBy
join (
	select a.StatusId, a.StatusSchema, a.StatusNo, a.StatusName, a.StatusDesc
	--, ISNULL(b.StatusName, a.StatusName) ApplyStatusName, ISNULL(b.StatusDesc, a.StatusDesc) ApplyStatusDesc
	FROM BAS.[Status] a
	--OUTER APPLY (
	--	SELECT ba.StatusName, ba.StatusDesc
	--	FROM BAS.StatusCompany ba
	--	WHERE ba.StatusId = a.StatusId AND ba.CompanyId = @CompanyId
	--) b  
    where a.StatusSchema ='Status'
) e on e.StatusNo = a.[Status]";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DepartmentId", @" AND a.DepartmentId = @DepartmentId", DepartmentId);

                    if (AuthorId > 0)
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "AuthorId", @" AND a.Author = @AuthorId", AuthorId);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartEnableDate", @" AND a.EnableDate >= @StartEnableDate", StartEnableDate.Length > 0 ? Convert.ToDateTime(StartEnableDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndEnableDate", @" AND a.EnableDate <= @EndEnableDate", EndEnableDate.Length > 0 ? Convert.ToDateTime(EndEnableDate).ToString("yyyy-MM-dd 23:59:59") : "");

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDisabledDate", @" AND a.DisabledDate >= @StartDisabledDate", StartDisabledDate.Length > 0 ? Convert.ToDateTime(StartDisabledDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDisabledDate", @" AND a.DisabledDate <= @EndDisabledDate", EndDisabledDate.Length > 0 ? Convert.ToDateTime(EndDisabledDate).ToString("yyyy-MM-dd 23:59:59") : "");

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartCreateDate", @" AND a.CreateDate >= @StartCreateDate", StartCreateDate.Length > 0 ? Convert.ToDateTime(StartCreateDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndCreateDate", @" AND a.CreateDate <= @EndCreateDate", EndCreateDate.Length > 0 ? Convert.ToDateTime(EndCreateDate).ToString("yyyy-MM-dd 23:59:59") : "");

                    if (!string.IsNullOrWhiteSpace(Status)) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.DepartmentId, a.DepartmentRateId";
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

        #region //GetUser -- 取得使用者資料 -- Ben Ma 2022.03.31
        public string GetUser(int UserId, int DepartmentId, int CompanyId, string Departments
            , string UserNo, string UserName, string Gender, string Status, string SearchKey
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.UserId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.UserNo, a.UserName, a.Gender, ISNULL(a.Email, '') Email, ISNULL(a.Job, '') Job, ISNULL(a.JobType, '') JobType, a.Status
                        , b.DepartmentId, b.DepartmentNo, b.DepartmentName
                        , c.CompanyId, c.CompanyNo, c.CompanyName, ISNULL(c.LogoIcon, -1) LogoIcon
                        , CASE a.Status 
                            WHEN 'S' THEN a.UserNo + ' ' + a.UserName + '(已離職)'
                            ELSE a.UserNo + ' ' + a.UserName
                        END UserWithNo
                        , CASE a.Status 
                            WHEN 'S' THEN b.DepartmentName + ' ' + a.UserNo + ' ' + a.UserName + '(已離職)'
                            ELSE b.DepartmentName + ' ' + a.UserNo + ' ' + a.UserName
                        END UserWithDepartment";
                    sqlQuery.mainTables =
                        @"FROM BAS.[User] a
                        INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                        INNER JOIN BAS.Company c ON b.CompanyId = c.CompanyId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserId", @" AND a.UserId = @UserId", UserId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DepartmentId", @" AND b.DepartmentId = @DepartmentId", DepartmentId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CompanyId", @" AND c.CompanyId = @CompanyId", CompanyId);
                    if (Departments.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Departments", @" AND b.DepartmentId IN @Departments", Departments.Split(','));
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserNo", @" AND a.UserNo LIKE '%' + @UserNo + '%'", UserNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserName", @" AND a.UserName LIKE '%' + @UserName + '%'", UserName);
                    if (Gender.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Gender", @" AND a.Gender IN @Gender", Gender.Split(','));
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey", @" AND (a.UserNo LIKE '%' + @SearchKey + '%' OR a.UserName LIKE '%' + @SearchKey + '%')", SearchKey);
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

        #region //GetUserByAuthority -- 取得使用者資料(權限) -- Ben Ma 2022.06.29
        public string GetUserByAuthority(int UserId, int DepartmentId, int CompanyId, string Departments
            , string UserNo, string UserName, string Gender, string Status, string SearchKey
            , string FunctionCode, string DetailCode
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.UserId,b.DetailId,b.FunctionId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.UserNo, a.UserName, a.Gender, ISNULL(a.Email, '') Email, ISNULL(a.Job, '') Job, ISNULL(a.JobType, '') JobType, a.Status
                        , e.DepartmentId, e.DepartmentNo, e.DepartmentName
                        , f.CompanyId, f.CompanyNo, f.CompanyName, ISNULL(f.LogoIcon, -1) LogoIcon
                        , CASE a.Status 
                            WHEN 'S' THEN a.UserNo + ' ' + a.UserName + '(已離職)'
                            ELSE a.UserNo + ' ' + a.UserName
                        END UserWithNo
                        , CASE a.Status 
                            WHEN 'S' THEN e.DepartmentName + ' ' + a.UserNo + ' ' + a.UserName + '(已離職)'
                            ELSE e.DepartmentName + ' ' + a.UserNo + ' ' + a.UserName
                        END UserWithDepartment";
                    sqlQuery.mainTables =
                        @"FROM BAS.[User] a
                        OUTER APPLY (
                            SELECT ba.DetailId, ba.FunctionId, ba.DetailCode, ba.[Status]
                            FROM BAS.FunctionDetail ba
                        ) b
                        INNER JOIN BAS.[Function] c ON b.FunctionId = c.FunctionId
                        OUTER APPLY (
                            SELECT ISNULL((
                                SELECT TOP 1 1
                                FROM BAS.RoleFunctionDetail ca
                                WHERE ca.DetailId = b.DetailId
                                AND ca.RoleId IN (
                                    SELECT caa.RoleId
                                    FROM BAS.UserRole caa
                                    INNER JOIN BAS.[Role] cab ON caa.RoleId = cab.RoleId
                                    WHERE 1=1
                                    AND caa.UserId = a.UserId
                                    AND cab.CompanyId = @CurrentCompany
                                )
                            ), 0) Authority
                        ) d
                        INNER JOIN BAS.Department e ON a.DepartmentId = e.DepartmentId
                        INNER JOIN BAS.Company f ON e.CompanyId = f.CompanyId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.[Status] = 'A'
                                            AND b.[Status] = 'A'
                                            AND c.FunctionCode = @FunctionCode
                                            AND b.DetailCode = @DetailCode
                                            AND d.Authority > 0";
                    dynamicParameters.Add("@CurrentCompany", CurrentCompany);
                    dynamicParameters.Add("@FunctionCode", FunctionCode);
                    dynamicParameters.Add("@DetailCode", DetailCode);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserId", @" AND a.UserId = @UserId", UserId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DepartmentId", @" AND e.DepartmentId = @DepartmentId", DepartmentId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CompanyId", @" AND f.CompanyId = @CompanyId", CompanyId);
                    if (Departments.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Departments", @" AND b.DepartmentId IN @Departments", Departments.Split(','));
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserNo", @" AND a.UserNo LIKE '%' + @UserNo + '%'", UserNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserName", @" AND a.UserName LIKE '%' + @UserName + '%'", UserName);
                    if (Gender.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Gender", @" AND a.Gender IN @Gender", Gender.Split(','));
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey", @" AND (a.UserNo LIKE '%' + @SearchKey + '%' OR a.UserName LIKE '%' + @SearchKey + '%')", SearchKey);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "f.CompanyId, a.UserNo";
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

        #region //GetSystem -- 取得系統別資料 -- Ben Ma 2022.05.04
        public string GetSystem(int SystemId, string SystemCode, string SystemName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.SystemId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.SystemCode, a.SystemName, a.ThemeIcon, a.SortNumber, a.Status
                        , a.SystemCode + ' ' + a.SystemName SystemWithCode";
                    sqlQuery.mainTables =
                        @"FROM BAS.[System] a";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SystemId", @" AND a.SystemId = @SystemId", SystemId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SystemCode", @" AND a.SystemCode LIKE '%' + @SystemCode + '%'", SystemCode);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SystemName", @" AND a.SystemName LIKE '%' + @SystemName + '%'", SystemName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
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

        #region //GetModule -- 取得模組別資料 -- Kan 2022.05.05
        public string GetModule(int ModuleId, int SystemId, string ModuleCode, string ModuleName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.ModuleId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ModuleCode, a.ModuleName, a.ThemeIcon, a.SortNumber, a.Status
                        , a.ModuleCode + ' ' + a.ModuleName ModuleWithCode
                        , b.SystemId, b.SystemCode, b.SystemName";
                    sqlQuery.mainTables =
                        @"FROM BAS.Module a
                        INNER JOIN BAS.[System] b ON a.SystemId = b.SystemId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ModuleId", @" AND a.ModuleId = @ModuleId", ModuleId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SystemId", @" AND b.SystemId = @SystemId", SystemId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ModuleCode", @" AND a.ModuleCode = @ModuleCode", ModuleCode);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ModuleName", @" AND a.ModuleName LIKE '%' + @ModuleName + '%'", ModuleName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "b.SortNumber, a.SortNumber";
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

        #region //GetFunction -- 取得功能別資料 -- Kan 2022.05.06
        public string GetFunction(int FunctionId, int ModuleId, int SystemId, string FunctionCode, string FunctionName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.FunctionId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.FunctionCode, a.FunctionName, a.SortNumber, a.UrlTarget, a.[Status]
                        , b.ModuleId, b.ModuleCode, b.ModuleName
                        , c.SystemId, c.SystemCode, c.SystemName";
                    sqlQuery.mainTables =
                        @"FROM BAS.[Function] a
                        INNER JOIN BAS.Module b ON a.ModuleId = b.ModuleId
                        INNER JOIN BAS.[System] c ON b.SystemId = c.SystemId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "FunctionId", @" AND a.FunctionId = @FunctionId", FunctionId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ModuleId", @" AND b.ModuleId = @ModuleId", ModuleId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SystemId", @" AND c.SystemId = @SystemId", SystemId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "FunctionCode", @" AND a.FunctionCode LIKE '%' + @FunctionCode + '%'", FunctionCode);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "FunctionName", @" AND a.FunctionName LIKE '%' + @FunctionName + '%'", FunctionName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "c.SortNumber, b.SortNumber, a.SortNumber";
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

        #region //GetFunctionDetail -- 取得詳細功能別資料 -- Kan 2022.05.10
        public string GetFunctionDetail(int DetailId, int FunctionId, string DetailCode, string DetailName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.DetailId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.DetailCode, a.DetailName, a.SortNumber, a.Status
                        , b.FunctionId, b.FunctionCode, b.FunctionName";
                    sqlQuery.mainTables =
                        @"FROM BAS.FunctionDetail a
                        INNER JOIN BAS.[Function] b ON a.FunctionId = b.FunctionId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DetailId", @" AND a.DetailId = @DetailId", DetailId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "FunctionId", @" AND b.FunctionId = @FunctionId", FunctionId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DetailCode", @" AND a.DetailCode LIKE '%' + @DetailCode + '%'", DetailCode);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DetailName", @" AND a.DetailName LIKE '%' + @DetailName + '%'", DetailName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "b.FunctionId, a.SortNumber";
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

        #region //GetFunctionDetailCode -- 取得功能詳細代碼 -- Ben Ma 2022.06.15
        public string GetFunctionDetailCode()
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT DISTINCT a.DetailCode, a.DetailName
                            FROM BAS.FunctionDetail a
                            WHERE 1=1
                            ORDER BY a.DetailCode";

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

        #region //GetSystemMenu -- 取得系統選單資料 -- Ben Ma 2022.05.18
        public string GetSystemMenu()
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT DISTINCT e.SystemId, e.SystemCode, e.SystemName, e.ThemeIcon, e.SortNumber
                            FROM BAS.RoleFunctionDetail a
                            INNER JOIN BAS.FunctionDetail b ON a.DetailId = b.DetailId
                            INNER JOIN BAS.[Function] c ON b.FunctionId = c.FunctionId
                            INNER JOIN BAS.[Module] d ON c.ModuleId = d.ModuleId
                            INNER JOIN BAS.[System] e ON d.SystemId = e.SystemId
                            WHERE a.RoleId IN (
                                SELECT aa.RoleId
                                FROM BAS.UserRole aa
                                INNER JOIN BAS.[Role] ab ON aa.RoleId = ab.RoleId
                                WHERE aa.UserId = @UserId
                                AND ab.CompanyId = @CompanyId
                            )
                            AND b.DetailCode IN ('read')
                            AND b.[Status] = @Status
                            AND c.[Status] = @Status
                            AND d.[Status] = @Status
                            AND e.[Status] = @Status
                            ORDER BY e.SortNumber";
                    dynamicParameters.Add("UserId", CurrentUser);
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    dynamicParameters.Add("Status", "A");

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

        #region //GetModuleMenu -- 取得模組選單資料 -- Ben Ma 2022.05.18
        public string GetModuleMenu(string SystemCode)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT DISTINCT d.ModuleId, d.ModuleCode, d.ModuleName, d.ThemeIcon, d.SortNumber
                            , (
                                SELECT aa.FunctionId, aa.FunctionCode, aa.FunctionName, aa.UrlTarget
                                FROM BAS.[Function] aa
                                INNER JOIN BAS.[Module] ab ON aa.ModuleId = ab.ModuleId
                                WHERE ab.ModuleId = d.ModuleId
                                AND aa.FunctionId IN (
                                    SELECT y.FunctionId
                                    FROM BAS.RoleFunctionDetail x
                                    INNER JOIN BAS.FunctionDetail y ON x.DetailId = y.DetailId
                                    WHERE x.RoleId IN (
                                        SELECT xa.RoleId
                                        FROM BAS.UserRole xa
                                        INNER JOIN BAS.[Role] xb ON xa.RoleId = xb.RoleId
                                        WHERE xa.UserId = @UserId
                                        AND xb.CompanyId = @CompanyId
                                    )
                                    AND y.DetailCode IN ('read')
                                    AND y.[Status] = @Status
                                )
                                AND aa.[Status] = @Status
                                AND ab.[Status] = @Status
                                ORDER BY aa.SortNumber
                                FOR JSON PATH, ROOT('data')
                            ) FunctionData
                            FROM BAS.RoleFunctionDetail a
                            INNER JOIN BAS.FunctionDetail b ON a.DetailId = b.DetailId
                            INNER JOIN BAS.[Function] c ON b.FunctionId = c.FunctionId
                            INNER JOIN BAS.[Module] d ON c.ModuleId = d.ModuleId
                            INNER JOIN BAS.[System] e ON d.SystemId = e.SystemId
                            WHERE a.RoleId IN (
                                SELECT aa.RoleId
                                FROM BAS.UserRole aa
                                INNER JOIN BAS.[Role] ab ON aa.RoleId = ab.RoleId
                                WHERE aa.UserId = @UserId
                                AND ab.CompanyId = @CompanyId
                            )
                            AND b.DetailCode IN ('read')
                            AND b.[Status] = @Status
                            AND c.[Status] = @Status
                            AND d.[Status] = @Status
                            AND e.[Status] = @Status";
                    dynamicParameters.Add("UserId", CurrentUser);
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    dynamicParameters.Add("Status", "A");
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SystemCode", @" AND e.SystemCode = @SystemCode", SystemCode);

                    sql += @" ORDER BY d.SortNumber";

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

        #region //GetShift -- 取得班次資料 -- Zoey 2022.05.19
        public string GetShift(int ShiftId, string ShiftName, string WorkBeginTime, string WorkEndTime, string WorkHours
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.ShiftId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ShiftName, a.WorkBeginTime, a.WorkEndTime, a.WorkHours";
                    sqlQuery.mainTables =
                        @"FROM BAS.Shift a";
                    sqlQuery.auxTables = "";
                    string queryCondition = " AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ShiftId", @" AND a.ShiftId = @ShiftId", ShiftId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ShiftName", @" AND a.ShiftName LIKE '%' + @ShiftName + '%'", ShiftName);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ShiftId";
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

        #region //GetDepartmentRateDetail 取部門得費用率詳細資料 --Chia Yuan 2023.6.27
        public string GetDepartmentRateDetail(int DepartmentRateId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.DepartmentRateId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.Author, a.ResourceRate, a.OverheadRate";
                    sqlQuery.mainTables =
                        @"FROM BAS.DepartmentRate a
join BAS.Department b on b.DepartmentId = a.DepartmentId";
                    sqlQuery.auxTables = "";
                    string queryCondition = " AND b.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DepartmentRateId", @" AND a.DepartmentRateId = @DepartmentRateId", DepartmentRateId);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.DepartmentRateId";
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

        #region //GetErpData -- 取得Erp資料 -- Zoey 2022.06.16
        public string GetErpData(string UserNo, string ColumnName, string Condition, string OrderBy)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //判斷公司資料是否正確
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 CompanyId, CompanyName, ErpDb
                            FROM BAS.Company
                            WHERE CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0) throw new SystemException("【公司】資料錯誤!");

                    foreach (var item in result)
                    {
                        ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                    }
                    #endregion
                }

                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                {
                    switch (ColumnName)
                    {
                        case "RelatedPerson"://關係人代號
                            sql = @"SELECT LTRIM(RTRIM(MB002)) SubsidiaryName, LTRIM(RTRIM(MB015)) RelatedPerson 
                                    FROM FCSMB
                                    WHERE 1=1";
                            break;
                        case "Currency"://交易幣別
                            sql = @"SELECT LTRIM(RTRIM(MF001)) Currency, LTRIM(RTRIM(MF002)) CurrencyName
                                    FROM CMSMF
                                    WHERE 1=1";
                            break;
                        case "TradeTerm"://交易條件
                            sql = @"SELECT LTRIM(RTRIM(NK001)) TradeTermNo, LTRIM(RTRIM(NK002)) TradeTermName 
                                    FROM CMSNK
                                    WHERE 1=1";
                            break;
                        case "PaymentTerm"://付款條件
                            sql = @"SELECT LTRIM(RTRIM(NA002)) PaymentTermNo, LTRIM(RTRIM(NA003)) PaymentTermName 
                                    FROM CMSNA
                                    WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "Condition", @" AND NA001 = @Condition", Condition);
                            break;
                        case "TaxNo"://稅別碼 課稅別 營業稅率
                            sql = @"SELECT LTRIM(RTRIM(NN001)) TaxNo, LTRIM(RTRIM(NN002)) TaxName, ROUND(LTRIM(RTRIM(NN004)),2) BusinessTaxRate, LTRIM(RTRIM(NN006)) Taxation
                                    FROM CMSNN 
                                    WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "Condition", @" AND NN001 = @Condition", Condition);
                            break;
                        case "InvoiceCount"://發票聯數
                            sql = @"SELECT LTRIM(RTRIM(NM001)) InvoiceCountNo, LTRIM(RTRIM(NM002)) InvoiceCountName 
                                    FROM CMSNM a
                                    INNER JOIN CMSNN b ON a.NM001 = b.NN005
                                    WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "Condition", @" AND b.NN001 = @Condition", Condition);
                            break;
                        case "TransactionObject"://交易對象(國家/地區/路線)
                            sql = @"SELECT LTRIM(RTRIM(MR002)) ClassificationNo, LTRIM(RTRIM(MR003)) ClassificationName
                                    FROM CMSMR
                                    WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "Condition", @" AND MR001 = @Condition", Condition);
                            break;
                        case "Account"://帳款科目
                            sql = @"SELECT LTRIM(RTRIM(MA001)) AccountNo, LTRIM(RTRIM(MA003)) AccountName
                                    FROM ACTMA 
                                    WHERE 1=1";
                            break;
                        case "AccountInvoice"://票據科目
                            sql = @"SELECT LTRIM(RTRIM(MA001)) AccountNo, LTRIM(RTRIM(MA003)) AccountName
                                    FROM ACTMA 
                                    WHERE 1=1 
                                    AND MA001 = '1151' 
                                    OR MA001 = '1172-01'";
                            break;
                        case "AccountOverhead"://加工費用科目
                            sql = @"SELECT LTRIM(RTRIM(MA001)) AccountNo, LTRIM(RTRIM(MA003)) AccountName
                                    FROM ACTMA 
                                    WHERE 1=1 
                                    AND MA001 = '5888-13'";
                            break;
                        case "ShipMethod"://運輸方式
                            sql = @"SELECT LTRIM(RTRIM(NJ001)) ShipMethodNo, LTRIM(RTRIM(NJ002)) ShipMethodName 
                                    FROM CMSNJ 
                                    WHERE 1=1";
                            break;
                        case "ExchangeRate"://匯率
                            string Today = DateTime.Now.ToString("yyyyMMdd");
                            sql = @"SELECT TOP 1 LTRIM(RTRIM(MG001)) Currency, ROUND(LTRIM(RTRIM(MG004)),3) ExchangeRateNameForMwe, ROUND(LTRIM(RTRIM(MG005)),3) ExchangeRateName, LTRIM(RTRIM(MG002)) StartDate
                                    FROM CMSMG 
                                    WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "Condition", @" AND MG001 = @Condition", Condition);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "Today", @" AND MG002 <= @Today", Today);
                            OrderBy = OrderBy.Length > 0 ? OrderBy : "StartDate DESC";
                            break;
                        case "ProduceLineNo"://生產線別
                            sql = @"SELECT LTRIM(RTRIM(MD001)) ProduceLineNo, LTRIM(RTRIM(MD002)) ProduceLineName
                                    FROM CMSMD
                                    LEFT JOIN CMSMB ON MB001=MD003
                                    WHERE 1=1";
                            break;
                        case "Project"://專案代號
                            sql = @"SELECT LTRIM(RTRIM(NB001)) Project, LTRIM(RTRIM(NB002)) ProjectName 
                                    FROM CMSNB
                                    WHERE 1=1";
                            break;
                        case "ProcessNo"://ERP製程代號
                            sql = @"SELECT LTRIM(RTRIM(MW001)) ProcessNo,LTRIM(RTRIM(MW002)) ProcessName  FROM CMSMW a
                                    WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "Condition", @" AND a.MW005 = @Condition", Condition);
                            break;
                        case "ProcessCode"://ERP廠商生產製程資訊
                            sql = @"SELECT LTRIM(RTRIM(a.MW001)) MW001, LTRIM(RTRIM(a.MW002)) MW002
                                    , (LTRIM(RTRIM(a.MW001)) + ' ' + LTRIM(RTRIM(a.MW002))) MWFullNO
                                    FROM CMSMW a
                                    LEFT JOIN CMSMV b ON a.MW007 = b.MV001
                                    WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "Condition", @" AND a.MW005 = @Condition", Condition);
                            OrderBy = OrderBy.Length > 0 ? OrderBy : "MW001";
                            break;
                        case "InvoiceType"://ERP發票聯數(單獨找)
                            sql = @"SELECT LTRIM(RTRIM(NM001)) InvoiceCountNo, LTRIM(RTRIM(NM002)) InvoiceCountName
                                    FROM CMSNM
                                    WHERE 1=1";
                            OrderBy = OrderBy.Length > 0 ? OrderBy : "NM001";
                            break;
                        case "MtlItemType"://類別一、二、三、四
                            sql = @"SELECT LTRIM(RTRIM(MA002)) TypeNo, LTRIM(RTRIM(MA003)) TypeName
                                    FROM INVMA 
                                    WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "Condition", @" AND MA001 = @Condition", Condition);
                            break;
                        case "Department"://部門
                            sql = @"SELECT LTRIM(RTRIM(ME001)) DepartmentNo, LTRIM(RTRIM(ME002)) DepartmentName
                                    FROM CMSME
                                    WHERE 1=1";
                            break;
                        case "CurrencyDecimal"://幣別小數點進位資料
                            sql = @"SELECT LTRIM(RTRIM(MF001)) CurrencyNo, LTRIM(RTRIM(MF002)) CurrencyName
                                    , LTRIM(RTRIM(MF003)) UnitDecimal, LTRIM(RTRIM(MF004)) TotalDecimal
                                    FROM CMSMF
                                    WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "Condition", @" AND MF001 = @Condition", Condition);
                            break;
                        case "WipErpPrefix"://製令單別
                            sql = @"SELECT DISTINCT LTRIM(RTRIM(MQ001)) WoErpPrefix, LTRIM(RTRIM(MQ002)) WoErpPrefixName 
                                    FROM CMSMQ 
                                    LEFT JOIN CMSMU ON MQ001 = MU001 
                                    WHERE (MQ003 = '51' OR MQ003 = '52')  
                                    AND ((MU003 = '" + UserNo + "' AND MQ029 = 'Y')  or MQ029 = 'N' )";
                            break;
                        case "MrErpPrefix"://ERP領退料單單別
                            sql = @"SELECT LTRIM(RTRIM(a.MQ001)) MQ001, LTRIM(RTRIM(a.MQ002)) MQ002
                                    , (LTRIM(RTRIM(a.MQ001)) + ' ' + LTRIM(RTRIM(a.MQ002))) MrErpPrefix
                                    FROM CMSMQ a
                                    LEFT JOIN CMSMU b ON a.MQ001 = b.MU001
                                    WHERE 1=1
                                    AND ((b.MU003 = 'DS' AND a.MQ029 = 'Y') OR a.MQ029 = 'N')
                                    AND a.MQ001 >= N''''";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "Condition", @" AND a.MQ003 = @Condition", Condition);
                            OrderBy = OrderBy.Length > 0 ? OrderBy : "MQ001";
                            break;
                        case "MoErpPrefix"://入庫單別
                            sql = @"SELECT DISTINCT LTRIM(RTRIM(a.MQ001)) MweErpPrefix, LTRIM(RTRIM(a.MQ002)) MweErpPrefixName 
                                    FROM CMSMQ a
                                    LEFT JOIN CMSMU b ON a.MQ001 = b.MU001 
                                    WHERE (a.MQ003 = '59')  
                                    AND ((b.MU003 = '" + UserNo + "' AND a.MQ029 = 'Y')  or a.MQ029 = 'N' )";
                            break;
                        case "PrErpPrefix"://ERP請購單單別
                            sql = @"SELECT LTRIM(RTRIM(a.MQ001)) MQ001, LTRIM(RTRIM(a.MQ002)) MQ002
                                    , (LTRIM(RTRIM(a.MQ001)) + ' ' + LTRIM(RTRIM(a.MQ002))) PrErpPrefix
                                    FROM CMSMQ a
                                    LEFT JOIN CMSMU b ON a.MQ001 = b.MU001
                                    WHERE 1=1
                                    AND ((b.MU003 = 'DS' AND a.MQ029 = 'Y') OR a.MQ029 = 'N')
                                    AND a.MQ001 >= N''''";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "Condition", @" AND a.MQ003 = @Condition", Condition);
                            OrderBy = OrderBy.Length > 0 ? OrderBy : "MQ001";
                            break;
                        case "PoErpPrefix"://ERP採購單單別
                            sql = @"SELECT LTRIM(RTRIM(a.MQ001)) PoErpPrefix, LTRIM(RTRIM(a.MQ002)) PoErpPrefixName
                                    , (LTRIM(RTRIM(a.MQ001)) + ' ' + LTRIM(RTRIM(a.MQ002))) PoFullName
                                    FROM CMSMQ a
                                    LEFT JOIN CMSMU b ON a.MQ001 = b.MU001
                                    WHERE 1=1
                                    AND MQ003 = '33'
                                    AND ((b.MU003 = 'DS' AND a.MQ029 = 'Y') OR a.MQ029 = 'N')
                                    AND a.MQ001 >= N''''";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "Condition", @" AND a.MQ003 = @Condition", Condition);
                            OrderBy = OrderBy.Length > 0 ? OrderBy : "MQ001";
                            break;
                        case "GrErpPrefix": //進貨單別
                            sql = @"SELECT LTRIM(RTRIM(MQ001)) GrErpPrefixNo, LTRIM(RTRIM(MQ002)) GrErpPrefixName
                                    FROM CMSMQ
                                    LEFT JOIN CMSMU ON MQ001 = MU001 
                                    WHERE MQ003 = '34'
                                    AND ((MU003 = 'DS' AND MQ029 ='Y') OR MQ029='N')";
                            break;
                        case "SoErpPrefix"://訂單單別
                            sql = @"SELECT LTRIM(RTRIM(MQ001)) SoErpPrefixNo, LTRIM(RTRIM(MQ002)) SoErpPrefixName 
                                    FROM CMSMQ
                                    LEFT JOIN CMSMU ON MQ001 = MU001 
                                    WHERE 1=1
                                    AND (MQ003 ='22' OR (MQ003 = '27' AND MQ038 = '1'))
                                    AND ((MU003 = 'DS' AND MQ029 ='Y') OR MQ029='N')";
                            break;
                        case "ItErpPrefix": //庫存異動單單別
                            sql = @"SELECT LTRIM(RTRIM(MQ001)) ItErpPrefixNo, LTRIM(RTRIM(MQ002)) ItErpPrefixName
                                    FROM CMSMQ
                                    LEFT JOIN CMSMU ON MQ001 = MU001 
                                    WHERE MQ003 = '11'
                                    AND ((MU003 = 'DS' AND MQ029 ='Y') OR MQ029='N')";
                            break;
                        case "ItfErpPrefix": //調撥單單別
                            sql = @"SELECT LTRIM(RTRIM(MQ001)) ItErpPrefixNo, LTRIM(RTRIM(MQ002)) ItErpPrefixName
                                    FROM CMSMQ
                                    LEFT JOIN CMSMU ON MQ001 = MU001 
                                    WHERE MQ003 = '12'
                                    AND ((MU003 = 'DS' AND MQ029 ='Y') OR MQ029='N')";
                            break;
                        case "TsnErpPrefix": //暫出單別
                            sql = @"SELECT LTRIM(RTRIM(MQ001)) TsnErpPrefixNo, LTRIM(RTRIM(MQ002)) TsnErpPrefixName
                                    FROM CMSMQ
                                    LEFT JOIN CMSMU ON MQ001 = MU001 
                                    WHERE 1=1
                                    AND (MQ003 ='13' OR MQ003 = '14')
                                    AND ((MU003 = 'DS' AND MQ029 ='Y') OR MQ029='N')";
                            break;
                        case "TsrnErpPrefix": //暫出歸還單別
                            sql = @"SELECT LTRIM(RTRIM(MQ001)) TsrnErpPrefixNo, LTRIM(RTRIM(MQ002)) TsrnErpPrefixName
                                    FROM CMSMQ
                                    LEFT JOIN CMSMU ON MQ001 = MU001 
                                    WHERE 1=1
                                    AND (MQ003 ='15' OR MQ003 = '16')
                                    AND ((MU003 = 'DS' AND MQ029 ='Y') OR MQ029='N')";
                            break;
                        case "RoErpPrefix"://銷貨單別
                            sql = @"SELECT LTRIM(RTRIM(a.MQ001)) RoErpPrefix, LTRIM(RTRIM(a.MQ002)) RoErpPrefixName
                                    FROM CMSMQ a
                                    LEFT JOIN CMSMU b ON a.MQ001 = b.MU001 
                                    WHERE 1=1
                                    AND (a.MQ003 ='23' OR (a.MQ003 = '23' AND a.MQ038 = '1'))
                                    AND ((b.MU003 = '" + UserNo + "' AND a.MQ029 = 'Y')  or a.MQ029 = 'N' )";
                            break;
                        case "RtErpPrefix"://銷貨單別
                            sql = @"SELECT LTRIM(RTRIM(a.MQ001)) RtErpPrefix, LTRIM(RTRIM(a.MQ002)) RtErpPrefixName
                                    FROM CMSMQ a
                                    LEFT JOIN CMSMU b ON a.MQ001 = b.MU001 
                                    WHERE 1=1
                                    AND (a.MQ003 ='24' OR (a.MQ003 = '24' AND a.MQ038 = '1'))
                                    AND ((b.MU003 = '" + UserNo + "' AND a.MQ029 = 'Y')  or a.MQ029 = 'N' )";
                            break;
                        case "Processor": //加工廠商
                            sql = @"SELECT LTRIM(RTRIM(MA001)) ProcessorNo, LTRIM(RTRIM(MA002)) ProcessorName
                                    FROM PURMA
                                    WHERE 1=1";
                            break;
                        case "Factory": //廠別
                            sql = @"SELECT LTRIM(RTRIM(MB001)) Factory, LTRIM(RTRIM(MB002)) FactoryName
                                    FROM CMSMB
                                    WHERE 1=1";
                            break;
                    }

                    sql += OrderBy.Length > 0 ? string.Format(@" ORDER BY {0}", OrderBy) : "";
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

        #region //GetErpSysData -- 取得ErpSys資料 -- Zoey 2022.06.21
        public string GetErpSysData(string ColumnName, string Condition, string OrderBy)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(ErpSysConnectionStrings))
                {
                    switch (ColumnName)
                    {
                        case "Bank": //本國銀行
                            sql = @"SELECT LTRIM(RTRIM(MO001)) BankNo, LTRIM(RTRIM(MO006)) BankName 
                                    FROM CMSMO
                                    WHERE MO002 = 1";
                            break;
                    }

                    sql += OrderBy.Length > 0 ? string.Format(@" ORDER BY {0}", OrderBy) : "";
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

        #region //GetAssetsData -- 取得ERP資產資料 -- Ann 2023-05-16
        public string GetAssetsData(string UserNo, string MB001, string MB002, string MB006
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //確認公司別DB
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.ErpNo, a.ErpDb
                            FROM BAS.Company a
                            WHERE CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    var companyResult = sqlConnection.Query(sql, dynamicParameters);

                    if (companyResult.Count() <= 0) throw new SystemException("公司別錯誤!");

                    foreach (var item in companyResult)
                    {
                        ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                    }
                    #endregion

                    using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();
                        sqlQuery.mainKey = "A.MB001";
                        sqlQuery.auxKey = "";
                        sqlQuery.columns =
                            @", B.MC002 部門代號,D.ME002 部門名稱,A.MB006 資產類別,A.MB001 資產編號,A.MB002 資產名稱,A.MB003 資產規格,B.MC003 保管人代號
                            , C.MV001 保管人工號, C.MV002 保管人姓名,B.MC004 數量
                            , A.MB016 取得日期,A.MB020 取得成本";
                        sqlQuery.mainTables =
                            @"FROM ASTMB A
                            LEFT JOIN ASTMC B on A.MB001 = B.MC001
                            LEFT JOIN CMSMV C on B.MC003 = C.MV001
                            LEFT JOIN CMSME D on B.MC002 = D.ME001";
                        sqlQuery.auxTables = "";
                        string queryCondition = @"";
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserNo", @" AND B.MC003 LIKE '%' + @UserNo + '%'", UserNo);
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MB001", @" AND A.MB001 LIKE '%' + @MB001 + '%'", MB001);
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MB002", @" AND A.MB002 LIKE '%' + @MB002 + '%'", MB002);
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MB006", @" AND A.MB006 LIKE '%' + @MB006 + '%'", MB006);
                        sqlQuery.conditions = queryCondition;
                        sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "A.MB001 DESC";
                        sqlQuery.pageIndex = PageIndex;
                        sqlQuery.pageSize = PageSize;

                        var result = BaseHelper.SqlQuery(sqlConnection2, dynamicParameters, sqlQuery);

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            data = result
                        });
                        #endregion
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

        #region //GetUploadWhitelist -- 取得上傳路徑白名單 -- Ann 2024-09-23
        public string GetUploadWhitelist(int ListId, string ListNo, string FolerPath)
        {
            try
            {
                if (ListNo.Length <= 0) throw new SystemException("路徑分類不能為空!!");
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //取得使用者部門
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DepartmentId
                            FROM BAS.[User]
                            WHERE UserId = @UserId";
                    dynamicParameters.Add("UserId", CurrentUser);

                    var UserResult = sqlConnection.Query(sql, dynamicParameters);

                    if (UserResult.Count() <= 0) throw new SystemException("使用者資訊錯誤!!");

                    int DepartmentId = -1;
                    foreach (var item in UserResult)
                    {
                        DepartmentId = item.DepartmentId;
                    }
                    #endregion

                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.ListId, a.ListNo, a.DepartmentId , a.FolderPath
                            FROM BAS.UploadWhitelist a 
                            WHERE a.ListNo = @ListNo
                            AND a.DepartmentId = @DepartmentId";
                    dynamicParameters.Add("ListNo", ListNo);
                    dynamicParameters.Add("DepartmentId", DepartmentId);

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ListId", @" AND a.ListId = @ListId", ListId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "FolerPath", @" AND a.FolerPath = @FolerPath", FolerPath);

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
        #endregion        .

        #region //GetCheckUploadWhitelist -- 確認路徑是否為合法路徑 -- Ann 2024-09-24
        public string GetCheckUploadWhitelist(int ListId, string ListNo, string FolderPath)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //取得使用者部門
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DepartmentId
                            FROM BAS.[User]
                            WHERE UserId = @UserId";
                    dynamicParameters.Add("UserId", CurrentUser);

                    var UserResult = sqlConnection.Query(sql, dynamicParameters);

                    if (UserResult.Count() <= 0) throw new SystemException("使用者資訊錯誤!!");

                    int DepartmentId = -1;
                    foreach (var item in UserResult)
                    {
                        DepartmentId = item.DepartmentId;
                    }
                    #endregion

                    #region //確認白名單資料是否正確
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT FolderPath
                            FROM BAS.UploadWhitelist
                            WHERE ListId = @ListId
                            AND ListNo = @ListNo
                            AND DepartmentId = @DepartmentId";
                    dynamicParameters.Add("ListId", ListId);
                    dynamicParameters.Add("ListNo", ListNo);
                    dynamicParameters.Add("DepartmentId", DepartmentId);

                    var UploadWhitelistResult = sqlConnection.Query(sql, dynamicParameters);

                    if (UploadWhitelistResult.Count() <= 0) throw new SystemException("白名單資料錯誤!!");

                    string WhiteFolderPath = "";
                    foreach (var item in UploadWhitelistResult)
                    {
                        WhiteFolderPath = item.FolderPath;
                    }
                    #endregion

                    if (FolderPath.IndexOf(WhiteFolderPath) == -1) throw new SystemException("此路徑非合法路徑!!");

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = UploadWhitelistResult
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

        #region //GetFileInfo -- 取得檔案相關資訊 -- Ann 2024-09-24
        public string GetFileInfo(string FilePath, string MESFolderPath, string DownloadFlag)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    if (DownloadFlag == "RD")
                    {
                        #region //先確認此下載路徑是否為合法路徑
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT FolderPath
                                FROM BAS.UploadWhitelist
                                WHERE 1=1";
                        var RdWhitelistResult = sqlConnection.Query(sql, dynamicParameters);

                        bool checkWhitelist = false;
                        foreach (var item in RdWhitelistResult)
                        {
                            if (FilePath.IndexOf(item.FolderPath) != -1)
                            {
                                checkWhitelist = true;
                                break;
                            }
                        }

                        if (checkWhitelist == false)
                        {
                            throw new SystemException("此路徑非合法路徑，無法下載!!");
                        }
                        #endregion
                    }
                    else
                    {
                        if (FilePath.IndexOf(MESFolderPath) == -1) throw new SystemException("此路徑非合法路徑，無法下載!!");

                        #region //處理URL特殊符號
                        if (FilePath.IndexOf("+") != -1) FilePath = FilePath.Replace("+", "%2B");
                        if (FilePath.IndexOf("/") != -1) FilePath = FilePath.Replace("/", "%2F");
                        if (FilePath.IndexOf("?") != -1) FilePath = FilePath.Replace("?", "%3F");
                        if (FilePath.IndexOf("#") != -1) FilePath = FilePath.Replace("#", "%23");
                        if (FilePath.IndexOf("&") != -1) FilePath = FilePath.Replace("&", "%26");
                        if (FilePath.IndexOf("=") != -1) FilePath = FilePath.Replace("=", "%3D");
                        #endregion
                    }

                    List<FilePathInfo> fileInfos = new List<FilePathInfo>();
                    if (File.Exists(FilePath))
                    {
                        FilePathInfo cadFileInfo = new FilePathInfo()
                        {
                            FileName = Path.GetFileNameWithoutExtension(FilePath),
                            FileExtension = Path.GetExtension(FilePath),
                            FileByte = File.ReadAllBytes(FilePath),
                        };

                        fileInfos.Add(cadFileInfo);

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            data = fileInfos
                        });
                        #endregion
                    }
                    else
                    {
                        throw new SystemException("檔案路徑錯誤!!");
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

        #region //GetUploadWhitelistManagement -- 取得白名單上傳路徑資料 -- Andrew 2024-10-04
        public string GetUploadWhitelistManagement(int ListId, string ListNo, int DepartmentId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.ListId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ListNo, a.DepartmentId, ISNULL(b.DepartmentName,'') as DepartmentName, a.FolderPath
                         ";
                    sqlQuery.mainTables =
                        @"FROM BAS.UploadWhitelist a 
                           INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND b.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ListId", @" AND a.ListId = @ListId", ListId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ListNo", @" AND a.ListNo LIKE '%' + @ListNo + '%'", ListNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DepartmentId", @" AND a.DepartmentId = @DepartmentId", DepartmentId);

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ListId DESC";
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
        #region //AddCompany -- 公司資料新增 -- Ben Ma 2022.04.18
        public string AddCompany(string CompanyNo, string CompanyName, string Telephone, string Fax, string Address
            , string AddressEn, int LogoIcon)
        {
            try
            {
                if (CompanyNo.Length <= 0) throw new SystemException("【公司編號】不能為空!");
                if (CompanyNo.Length > 10) throw new SystemException("【公司編號】長度錯誤!");
                if (CompanyName.Length <= 0) throw new SystemException("【公司名稱】不能為空!");
                if (CompanyName.Length > 100) throw new SystemException("【公司名稱】長度錯誤!");
                if (Telephone.Length > 30) throw new SystemException("【電話】長度錯誤!");
                if (Fax.Length > 30) throw new SystemException("【傳真】長度錯誤!");
                if (Address.Length > 200) throw new SystemException("【中文地址】長度錯誤!");
                if (AddressEn.Length > 200) throw new SystemException("【英文地址】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷公司編號是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Company
                                WHERE CompanyNo = @CompanyNo";
                        dynamicParameters.Add("CompanyNo", CompanyNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【公司編號】重複，請重新輸入!");
                        #endregion

                        #region //判斷公司名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Company
                                WHERE CompanyName = @CompanyName";
                        dynamicParameters.Add("CompanyName", CompanyName);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【公司名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO BAS.Company (CompanyNo, CompanyName, Telephone, Fax, Address
                                , AddressEn, LogoIcon, SystemStatus, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.CompanyId
                                VALUES (@CompanyNo, @CompanyName, @Telephone, @Fax, @Address
                                , @AddressEn, @LogoIcon, @SystemStatus, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyNo,
                                CompanyName,
                                Telephone,
                                Fax,
                                Address,
                                AddressEn,
                                LogoIcon = LogoIcon > 0 ? LogoIcon : (int?)null,
                                SystemStatus = "M", //手動
                                Status = "A", //啟用
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

        #region //AddDepartment -- 部門資料新增 -- Ben Ma 2022.04.18
        public string AddDepartment(int CompanyId, string DepartmentNo, string DepartmentName)
        {
            try
            {
                if (CompanyId <= 0) throw new SystemException("【所屬公司】不能為空!");
                if (DepartmentNo.Length <= 0) throw new SystemException("【部門編號】不能為空!");
                if (DepartmentNo.Length > 10) throw new SystemException("【部門編號】長度錯誤!");
                if (DepartmentName.Length <= 0) throw new SystemException("【部門名稱】不能為空!");
                if (DepartmentName.Length > 100) throw new SystemException("【部門名稱】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷公司資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Company
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("公司資料錯誤!");
                        #endregion

                        #region //判斷部門編號是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Department
                                WHERE CompanyId = @CompanyId
                                AND DepartmentNo = @DepartmentNo";
                        dynamicParameters.Add("CompanyId", CompanyId);
                        dynamicParameters.Add("DepartmentNo", DepartmentNo);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【部門編號】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO BAS.Department (CompanyId, DepartmentNo, DepartmentName
                                , SystemStatus, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.DepartmentId
                                VALUES (@CompanyId, @DepartmentNo, @DepartmentName
                                , @SystemStatus, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId,
                                DepartmentNo,
                                DepartmentName,
                                SystemStatus = "M", //手動
                                Status = "A", //啟用
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

        #region //AddDepartmentShift -- 部門班次新增 -- Zoey 2022.05.20
        public string AddDepartmentShift(int DepartmentId, string ShiftId)
        {
            try
            {
                if (ShiftId.Length <= 0) throw new SystemException("【班次】不能為空!");

                List<int> shifts = ShiftId.Split(',').Select(int.Parse).ToList();

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷部門資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Department
                                WHERE DepartmentId = @DepartmentId";
                        dynamicParameters.Add("DepartmentId", DepartmentId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【部門】資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        foreach (var shift in shifts)
                        {
                            #region //判斷班次資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.[Shift]
                                    WHERE ShiftId = @ShiftId
                                    AND CompanyId = @CompanyId";
                            dynamicParameters.Add("ShiftId", shift);
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            var result2 = sqlConnection.Query(sql, dynamicParameters);
                            if (result2.Count() <= 0) throw new SystemException("【班次】資料錯誤!");
                            #endregion

                            #region //判斷部門班次是否重複
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.DepartmentShift
                                    WHERE DepartmentId = @DepartmentId
                                    AND ShiftId = @ShiftId";
                            dynamicParameters.Add("DepartmentId", DepartmentId);
                            dynamicParameters.Add("ShiftId", shift);

                            var result3 = sqlConnection.Query(sql, dynamicParameters);
                            if (result3.Count() > 0) throw new SystemException("該部門【班次】重複，請重新輸入!");
                            #endregion

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO BAS.DepartmentShift (DepartmentId, ShiftId
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.ShiftId
                                    VALUES (@DepartmentId, @ShiftId
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    DepartmentId,
                                    ShiftId = shift,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();
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

        #region //AddDepartmentRate --部門費用率新增 -- Chia Yuan 2022.6.27

        public string AddDepartmentRate(int DepartmentId, int AuthorId, decimal ResourceRate, decimal OverheadRate)
        {
            try
            {
                if (AuthorId <= 0) throw new SystemException("【編寫者】不能為空!");
                //todo: 判斷是否為0  待確認
                //ResourceRate
                //OverheadRate

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷編寫者資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[User]
                                WHERE UserId = @AuthorId";
                        dynamicParameters.Add("AuthorId", AuthorId);

                        var result = sqlConnection.Query(sql, dynamicParameters).FirstOrDefault();
                        if (result == null) throw new SystemException("【編寫者】資料錯誤!");
                        #endregion

                        #region //判斷部門資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Department
                                WHERE DepartmentId = @DepartmentId";
                        dynamicParameters.Add("DepartmentId", DepartmentId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters).FirstOrDefault();
                        if (result2 == null) throw new SystemException("【部門】資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO BAS.DepartmentRate (DepartmentId,Author
,ResourceRate,OverheadRate,Status,CreateDate,LastModifiedDate,CreateBy,LastModifiedBy)
OUTPUT INSERTED.DepartmentRateId VALUES(@DepartmentId,@Author,@ResourceRate,@OverheadRate,@Status
,@CreateDate,@LastModifiedDate,@CreateBy,@LastModifiedBy)";

                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                DepartmentId,
                                Author = AuthorId,
                                ResourceRate,
                                OverheadRate,
                                Status = "S",
                                LastModifiedDate,
                                CreateDate,
                                CreateBy,
                                LastModifiedBy = CreateBy
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

        #region //AddUser -- 使用者資料新增 -- Ben Ma 2022.04.13
        public string AddUser(int DepartmentId, string UserNo, string UserName, string Gender, string Email)
        {
            try
            {
                if (DepartmentId <= 0) throw new SystemException("【所屬部門】不能為空!");
                if (UserNo.Length <= 0) throw new SystemException("【使用者編號】不能為空!");
                if (UserNo.Length > 10) throw new SystemException("【使用者編號】長度錯誤!");
                if (UserName.Length <= 0) throw new SystemException("【使用者名稱】不能為空!");
                if (UserName.Length > 30) throw new SystemException("【使用者名稱】長度錯誤!");
                if (Gender.Length <= 0) throw new SystemException("【性別】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷部門資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Department
                                WHERE DepartmentId = @DepartmentId";
                        dynamicParameters.Add("DepartmentId", DepartmentId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【部門】資料錯誤!");
                        #endregion

                        #region //判斷使用者編號是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[User]
                                WHERE UserNo = @UserNo";
                        dynamicParameters.Add("UserNo", UserNo);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【使用者編號】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO BAS.[User] (DepartmentId, UserNo, UserName, Gender, Email
                                , Password, UserStatus, SystemStatus, PasswordStatus, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.UserId
                                VALUES (@DepartmentId, @UserNo, @UserName, @Gender, @Email
                                , @Password, @UserStatus, @SystemStatus, @PasswordStatus, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                DepartmentId,
                                UserNo,
                                UserName,
                                Gender,
                                Email,
                                Password = BaseHelper.Sha256Encrypt(UserNo.ToLower()),
                                UserStatus = "F", //正式員工
                                SystemStatus = "M", //手動
                                PasswordStatus = "Y",
                                Status = "A", //啟用
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

        #region //AddUserLoginLog -- 使用者登入紀錄新增 -- Ben Ma 2022.09.29
        public string AddUserLoginLog(int UserId, string LoginSite)
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
                                FROM BAS.[User]
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", UserId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【使用者】資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO BAS.UserLoginLog (UserId, LoginDate, LoginIP, LoginSite
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.LogId
                                VALUES (@UserId, @LoginDate, @LoginIP, @LoginSite
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                UserId,
                                LoginDate = CreateDate,
                                LoginIP = BaseHelper.ClientIP(),
                                LoginSite,
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

        #region //AddSystem -- 系統別資料新增 -- Ben Ma 2022.05.04
        public string AddSystem(string SystemCode, string SystemName, string ThemeIcon)
        {
            try
            {
                if (SystemCode.Length <= 0) throw new SystemException("【系統代碼】不能為空!");
                if (SystemCode.Length > 10) throw new SystemException("【系統代碼】長度錯誤!");
                if (SystemName.Length <= 0) throw new SystemException("【系統名稱】不能為空!");
                if (SystemName.Length > 100) throw new SystemException("【系統名稱】長度錯誤!");
                if (ThemeIcon.Length > 100) throw new SystemException("【主題Icon】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷系統代碼是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[System]
                                WHERE SystemCode = @SystemCode";
                        dynamicParameters.Add("SystemCode", SystemCode);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【系統代碼】重複，請重新輸入!");
                        #endregion

                        #region //判斷系統名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[System]
                                WHERE SystemName = @SystemName";
                        dynamicParameters.Add("SystemName", SystemName);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【系統名稱】重複，請重新輸入!");
                        #endregion

                        #region //抓取目前最大序號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ISNULL(MAX(SortNumber), 0) MaxSort
                                FROM BAS.[System]";

                        int maxSort = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).MaxSort;
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO BAS.[System] (SystemCode, SystemName, ThemeIcon
                                , SortNumber, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.SystemId
                                VALUES (@SystemCode, @SystemName, @ThemeIcon
                                , @SortNumber, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SystemCode,
                                SystemName,
                                ThemeIcon = ThemeIcon.Length > 0 ? ThemeIcon : null,
                                SortNumber = maxSort + 1,
                                Status = "A", //啟用
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

        #region //AddModule -- 模組別資料新增 -- Kan 2022.05.05
        public string AddModule(int SystemId, string ModuleCode, string ModuleName, string ThemeIcon)
        {
            try
            {
                if (SystemId < 0) throw new SystemException("【所屬系統】不能為空!");
                if (ModuleCode.Length <= 0) throw new SystemException("【模組代碼】不能為空!");
                if (ModuleCode.Length > 50) throw new SystemException("【模組代碼】長度錯誤!");
                if (ModuleName.Length <= 0) throw new SystemException("【模組名稱】不能為空!");
                if (ModuleName.Length > 100) throw new SystemException("【模組名稱】長度錯誤!");
                if (ThemeIcon.Length > 100) throw new SystemException("【主題Icon】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷系統別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[System]
                                WHERE SystemId = @SystemId";
                        dynamicParameters.Add("SystemId", SystemId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【系統別】資料錯誤!");
                        #endregion

                        #region //判斷模組代碼是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Module
                                WHERE SystemId = @SystemId
                                AND ModuleCode = @ModuleCode";
                        dynamicParameters.Add("SystemId", SystemId);
                        dynamicParameters.Add("ModuleCode", ModuleCode);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【模組代碼】重複，請重新輸入!");
                        #endregion

                        #region //判斷模組名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Module
                                WHERE SystemId = @SystemId
                                AND ModuleName = @ModuleName";
                        dynamicParameters.Add("SystemId", SystemId);
                        dynamicParameters.Add("ModuleName", ModuleName);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() > 0) throw new SystemException("【模組名稱】重複，請重新輸入!");
                        #endregion

                        #region //抓取目前最大序號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ISNULL(MAX(SortNumber), 0) MaxSort
                                FROM BAS.Module
                                WHERE SystemId = @SystemId";
                        dynamicParameters.Add("SystemId", SystemId);

                        int maxSort = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).MaxSort;
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO BAS.Module (SystemId, ModuleCode, ModuleName, ThemeIcon
                                , SortNumber, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ModuleId
                                VALUES (@SystemId, @ModuleCode, @ModuleName, @ThemeIcon
                                , @SortNumber, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SystemId,
                                ModuleCode,
                                ModuleName,
                                ThemeIcon = ThemeIcon.Length > 0 ? ThemeIcon : "",
                                SortNumber = maxSort + 1,
                                Status = "A", //啟用
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

        #region //AddFunction -- 功能別資料新增 -- Kan 2022.05.06
        public string AddFunction(int ModuleId, string FunctionCode, string FunctionName, string UrlTarget)
        {
            try
            {
                if (ModuleId < 0) throw new SystemException("【所屬模組】不能為空!");
                if (FunctionCode.Length <= 0) throw new SystemException("【功能代碼】不能為空!");
                if (FunctionCode.Length > 50) throw new SystemException("【功能代碼】長度錯誤!");
                if (FunctionName.Length <= 0) throw new SystemException("【功能名稱】不能為空!");
                if (FunctionName.Length > 100) throw new SystemException("【功能名稱】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷模組別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Module
                                WHERE ModuleId = @ModuleId";
                        dynamicParameters.Add("ModuleId", ModuleId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【模組別】資料錯誤!");
                        #endregion

                        #region //判斷功能代碼是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Function]
                                WHERE ModuleId = @ModuleId
                                AND FunctionCode = @FunctionCode";
                        dynamicParameters.Add("ModuleId", ModuleId);
                        dynamicParameters.Add("FunctionCode", FunctionCode);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【功能代碼】重複，請重新輸入!");
                        #endregion

                        #region //判斷功能名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Function]
                                WHERE ModuleId = @ModuleId
                                AND FunctionName = @FunctionName";
                        dynamicParameters.Add("ModuleId", ModuleId);
                        dynamicParameters.Add("FunctionName", FunctionName);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() > 0) throw new SystemException("【功能名稱】重複，請重新輸入!");
                        #endregion

                        #region //抓取目前最大序號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ISNULL(MAX(SortNumber), 0) MaxSort
                                FROM BAS.[Function]
                                WHERE ModuleId = @ModuleId";
                        dynamicParameters.Add("ModuleId", ModuleId);

                        int maxSort = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).MaxSort;
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO BAS.[Function] (ModuleId, FunctionCode, FunctionName
                                , UrlTarget, SortNumber, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.FunctionId
                                VALUES (@ModuleId, @FunctionCode, @FunctionName
                                , @UrlTarget, @SortNumber, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ModuleId,
                                FunctionCode,
                                FunctionName,
                                UrlTarget,
                                SortNumber = maxSort + 1,
                                Status = "A", //啟用
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

        #region //AddFunctionDetail -- 詳細功能別資料新增 -- Kan 2022.05.10
        public string AddFunctionDetail(int FunctionId, string DetailCode, string DetailName)
        {
            try
            {
                if (FunctionId < 0) throw new SystemException("【所屬功能】不能為空!");
                if (DetailCode.Length <= 0) throw new SystemException("【功能詳細代碼】不能為空!");
                if (DetailCode.Length > 50) throw new SystemException("【功能詳細代碼】長度錯誤!");
                if (DetailName.Length <= 0) throw new SystemException("【功能詳細名稱】不能為空!");
                if (DetailName.Length > 100) throw new SystemException("【功能詳細名稱】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷功能別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Function]
                                WHERE FunctionId = @FunctionId";
                        dynamicParameters.Add("FunctionId", FunctionId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【功能別】資料錯誤!");
                        #endregion

                        #region //判斷功能詳細代碼是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.FunctionDetail
                                WHERE FunctionId = @FunctionId
                                AND DetailCode = @DetailCode";
                        dynamicParameters.Add("FunctionId", FunctionId);
                        dynamicParameters.Add("DetailCode", DetailCode);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【功能詳細代碼】重複，請重新輸入!");
                        #endregion

                        #region //判斷功能詳細名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.FunctionDetail
                                WHERE FunctionId = @FunctionId
                                AND DetailName = @DetailName";
                        dynamicParameters.Add("FunctionId", FunctionId);
                        dynamicParameters.Add("DetailName", DetailName);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() > 0) throw new SystemException("【功能詳細名稱】重複，請重新輸入!");
                        #endregion

                        #region //抓取目前最大序號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ISNULL(MAX(SortNumber), 0) MaxSort
                                FROM BAS.FunctionDetail
                                WHERE FunctionId = @FunctionId";
                        dynamicParameters.Add("FunctionId", FunctionId);

                        int maxSort = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).MaxSort;
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO BAS.FunctionDetail (FunctionId, DetailCode, DetailName
                                , SortNumber, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.DetailId
                                VALUES (@FunctionId, @DetailCode, @DetailName
                                , @SortNumber, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";

                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                FunctionId,
                                DetailCode,
                                DetailName,
                                SortNumber = maxSort + 1,
                                Status = "A", //啟用
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

        #region //AddShift -- 班次資料新增 -- Zoey 2022.05.19
        public string AddShift(string ShiftName, string WorkBeginTime, string WorkEndTime, string WorkHours)
        {
            try
            {
                if (ShiftName.Length <= 0) throw new SystemException("【班次名稱】不能為空!");
                if (ShiftName.Length > 100) throw new SystemException("【班次名稱】長度錯誤!");
                if (WorkBeginTime.Length <= 0) throw new SystemException("【上班時間】不能為空!");
                if (WorkBeginTime.Length > 8) throw new SystemException("【上班時間】長度錯誤!");
                if (WorkEndTime.Length <= 0) throw new SystemException("【下班時間】不能為空!");
                if (WorkEndTime.Length > 8) throw new SystemException("【下班時間】長度錯誤!");
                if (WorkHours.Length <= 0) throw new SystemException("【工作時數】不能為空!");
                if (WorkHours.Length > 100) throw new SystemException("【工作時數】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷班次名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Shift]
                                WHERE CompanyId = @CompanyId
                                AND ShiftName = @ShiftName";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ShiftName", ShiftName);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【班次名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO BAS.[Shift] (CompanyId, ShiftName, WorkBeginTime, WorkEndTime, WorkHours
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ShiftId
                                VALUES (@CompanyId, @ShiftName, @WorkBeginTime, @WorkEndTime, @WorkHours
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                ShiftName,
                                WorkBeginTime,
                                WorkEndTime,
                                WorkHours,
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

        #region //AddUploadWhitelistManagement -- 白名單上傳路徑資料新增 -- Andrew 2024-10-04
        public string AddUploadWhitelistManagement(string ListNo, int DepartmentId, string FolderPath)
        {
            try
            {
                if (ListNo.Length <= 0) throw new SystemException("【白名單類別】不能為空!");
                if (ListNo.Length > 20) throw new SystemException("【白名單類別】長度錯誤!");
                if (DepartmentId <= 0) throw new SystemException("【所屬部門】不能為空!");
                if (FolderPath.Length <= 0) throw new SystemException("【資料夾路徑】不能為空!");
                if (FolderPath.Length > 100) throw new SystemException("【資料夾路徑】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷所屬部門是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1 
                                FROM BAS.Department a 
                                WHERE a.DepartmentId = @DepartmentId
                                AND a.CompanyId = @CompanyId
                                AND a.Status = 'A'";
                        dynamicParameters.Add("DepartmentId", DepartmentId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【所屬部門】不存在，請重新輸入!");
                        #endregion

                        #region //判斷白名單類別、部門名稱及資料夾路徑是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1 
                                FROM BAS.UploadWhitelist a
                                INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                                WHERE a.ListNo = @ListNo
                                AND b.CompanyId = @CompanyId
                                AND a.DepartmentId = @DepartmentId
                                AND a.FolderPath = @FolderPath";
                        dynamicParameters.Add("ListNo", ListNo);
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("DepartmentId", DepartmentId);
                        dynamicParameters.Add("FolderPath", FolderPath);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("已有此筆資料，請透過搜尋查找相關資訊!");
                        #endregion

                        #region //新增SQL
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO BAS.UploadWhitelist (ListNo, DepartmentId, FolderPath,
                                CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ListId
                                VALUES (@ListNo, @DepartmentId, @FolderPath, 
                                @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ListNo,
                                DepartmentId,
                                FolderPath,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = insertResult.Count();
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

        #region //Update
        #region //UpdateNewLogin -- 新密碼登入 -- Ben Ma 2022.11.22
        public string UpdateNewLogin(string Account, string NewPassword, string ConfirmPassword)
        {
            try
            {
                if (NewPassword.Length <= 0) throw new SystemException("【新密碼】不能為空!");
                if (NewPassword != ConfirmPassword) throw new SystemException("【確認密碼】不正確!");                
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //密碼參數設定
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.PasswordExpiration, a.PasswordFormat, a.PasswordWrongCount
                            FROM BAS.PasswordSetting a";

                        var resultPasswordSetting = sqlConnection.Query(sql, dynamicParameters);
                        if (resultPasswordSetting.Count() <= 0) throw new SystemException("【密碼參數設定】資料錯誤!");

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

                        if (!Regex.IsMatch(NewPassword, PasswordFormat, RegexOptions.IgnoreCase)) throw new SystemException("【新密碼】格式錯誤!");

                        #region //判斷使用者資訊是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserId, a.UserName, a.Email, a.Password, a.PasswordStatus, a.Status, c.CompanyId, c.CompanyName
                                FROM BAS.[User] a
                                INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                                INNER JOIN BAS.Company c ON b.CompanyId = c.CompanyId
                                WHERE a.UserNo = @Account
                                AND b.Status = @Status
                                AND c.Status = @Status";
                        dynamicParameters.Add("Account", Account);
                        dynamicParameters.Add("Status", "A");

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("查無此帳號資料：" + Account);

                        int userId = -1;
                        
                        foreach (var item in result)
                        {
                            userId = Convert.ToInt32(item.UserId);                           
                        }
                        #endregion

                        #region //金鑰驗證
                        string key = BaseHelper.StrMid(BaseHelper.Sha256Encrypt(Account), 12, 32);
                        string iv = HttpContext.Current.Session["NewLoginIV"].ToString();
                        string ciphertext = HttpContext.Current.Session["NewLogin"].ToString();

                        if (BaseHelper.AESDecrypt(ciphertext, key, iv) != Account) throw new SystemException("帳號驗證錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE BAS.[User] SET
                                Password = @Password,
                                PasswordStatus = @PasswordStatus,
                                PasswordExpire = @PasswordExpire,
                                PasswordMistake = @PasswordMistake,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE UserId = @UserId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Password = BaseHelper.Sha256Encrypt(NewPassword),
                                PasswordStatus = "N",
                                PasswordExpire = DateTime.Now.AddMinutes(PasswordExpiration),
                                PasswordMistake = 0,
                                LastModifiedDate,
                                LastModifiedBy,
                                UserId = userId
                            });
                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

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

                #region//執行Mamo憑證
                JObject jsonResponseMamo = new JObject();
                string dataRequestMamo = "";
                string account = Account;
                string password = NewPassword;
                // 获取当前时间
                DateTime now = DateTime.UtcNow;
                // 定义Unix时间戳的起始时间（1970年1月1日）
                DateTime unixEpochStart = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
                // 计算从Unix时间戳起始时间到现在的时间差
                TimeSpan timeSpan = now - unixEpochStart;
                // 获取时间戳（总秒数）
                long timestamp = (long)timeSpan.TotalSeconds;
                string ip = BaseHelper.ClientIP();
                string MMQuery = "mm";

                using (HttpClient client = new HttpClient())
                {
                    string apiUrl = "http://192.168.20.46:2536/Mattermost/reset";
                    client.Timeout = TimeSpan.FromMinutes(60);
                    using (HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, apiUrl))
                    {
                        MultipartFormDataContent content = new MultipartFormDataContent
                                {
                                    { new StringContent(account.ToString()), "account" },
                                    { new StringContent(password.ToString()), "password" },
                                    { new StringContent(timestamp.ToString()), "timestamp" },
                                    { new StringContent(ip.ToString()), "ip" },
                                    { new StringContent(MMQuery.ToString()), "MMQuery" }
                                };
                        httpRequestMessage.Content = content;
                        using (HttpResponseMessage httpResponseMessage = client.SendAsync(httpRequestMessage).GetAwaiter().GetResult())
                        {
                            if (httpResponseMessage.IsSuccessStatusCode)
                            {
                                dataRequestMamo = httpResponseMessage.Content.ReadAsStringAsync().Result.ToString();
                                if (JObject.Parse(dataRequestMamo)["status"].ToString() == "success")
                                {
                                    #region //Response
                                    jsonResponseMamo = JObject.FromObject(new
                                    {
                                        status = "success",
                                        msg = ""
                                    });
                                    #endregion
                                }
                                else
                                {
                                    #region //Response
                                    jsonResponseMamo = JObject.FromObject(new
                                    {
                                        status = "error",
                                        msg = JObject.Parse(dataRequestMamo)["message"].ToString()
                                    });
                                    #endregion
                                }
                            }
                            else
                            {
                                throw new SystemException("伺服器連線異常");
                            }
                        }
                    }
                }
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

        #region //UpdatePassword -- 更新使用者密碼 -- Ben Ma 2022.04.06
        public string UpdatePassword(int UserId, string NewPassword, string ConfirmPassword)
        {
            try
            {
                if (NewPassword.Length <= 0) throw new SystemException("【新密碼】不能為空!");
                if (NewPassword != ConfirmPassword) throw new SystemException("【確認密碼】不正確!");
                string UserNo = "";
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //密碼參數設定
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.PasswordExpiration, a.PasswordFormat, a.PasswordWrongCount
                            FROM BAS.PasswordSetting a";

                        var resultPasswordSetting = sqlConnection.Query(sql, dynamicParameters);
                        if (resultPasswordSetting.Count() <= 0) throw new SystemException("【密碼參數設定】資料錯誤!");

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

                        if (!Regex.IsMatch(NewPassword, PasswordFormat, RegexOptions.IgnoreCase)) throw new SystemException("【新密碼】格式錯誤!");

                        #region //判斷使用者資訊是否正確
                        
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT UserNo
                                FROM BAS.[User]
                                WHERE UserId = @UserId
                                AND Status = @Status";
                        dynamicParameters.Add("UserId", UserId);
                        dynamicParameters.Add("Status", "A");

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0)
                        {
                            throw new SystemException("【使用者】資料錯誤!");
                        }
                        else {
                            foreach (var item in result) {
                                UserNo = item.UserNo;
                            }
                        }

                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE BAS.[User] SET
                                Password = @Password,
                                PasswordExpire = @PasswordExpire,
                                PasswordMistake = @PasswordMistake,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE UserId = @UserId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Password = BaseHelper.Sha256Encrypt(NewPassword),
                                PasswordExpire = DateTime.Now.AddMinutes(PasswordExpiration),
                                PasswordMistake = 0,
                                LastModifiedDate,
                                LastModifiedBy,
                                UserId
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

                #region//執行Mamo憑證
                JObject jsonResponseMamo = new JObject();
                string dataRequestMamo = "";
                string account = UserNo;
                string password = NewPassword;
                // 获取当前时间
                DateTime now = DateTime.UtcNow;
                // 定义Unix时间戳的起始时间（1970年1月1日）
                DateTime unixEpochStart = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
                // 计算从Unix时间戳起始时间到现在的时间差
                TimeSpan timeSpan = now - unixEpochStart;
                // 获取时间戳（总秒数）
                long timestamp = (long)timeSpan.TotalSeconds;
                string ip = BaseHelper.ClientIP();
                string MMQuery = "mm";
               
                using (HttpClient client = new HttpClient())
                {
                    string apiUrl = "http://192.168.20.46:2536/Mattermost/reset";
                    client.Timeout = TimeSpan.FromMinutes(60);
                    using (HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, apiUrl))
                    {
                        MultipartFormDataContent content = new MultipartFormDataContent
                                {
                                    { new StringContent(account.ToString()), "account" },
                                    { new StringContent(password.ToString()), "password" },
                                    { new StringContent(timestamp.ToString()), "timestamp" },
                                    { new StringContent(ip.ToString()), "ip" },
                                    { new StringContent(MMQuery.ToString()), "MMQuery" }
                                };
                        httpRequestMessage.Content = content;
                        using (HttpResponseMessage httpResponseMessage = client.SendAsync(httpRequestMessage).GetAwaiter().GetResult())
                        {
                            if (httpResponseMessage.IsSuccessStatusCode)
                            {
                                dataRequestMamo = httpResponseMessage.Content.ReadAsStringAsync().Result.ToString();
                                if (JObject.Parse(dataRequestMamo)["status"].ToString() == "success")
                                {
                                    #region //Response
                                    jsonResponseMamo = JObject.FromObject(new
                                    {
                                        status = "success",
                                        msg = ""
                                    });
                                    #endregion
                                }
                                else
                                {
                                    #region //Response
                                    jsonResponseMamo = JObject.FromObject(new
                                    {
                                        status = "error",
                                        msg = JObject.Parse(dataRequestMamo)["message"].ToString()
                                    });
                                    #endregion
                                }
                            }
                            else
                            {
                                throw new SystemException("伺服器連線異常");
                            }
                        }
                    }
                }
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

        #region //UpdatePasswordReset -- 密碼重置 -- Ben Ma 2022.11.22
        public string UpdatePasswordReset(int UserId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷使用者資訊是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 UserNo
                                FROM BAS.[User]
                                WHERE UserId = @UserId
                                AND Status = @Status";
                        dynamicParameters.Add("UserId", UserId);
                        dynamicParameters.Add("Status", "A");

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【使用者】資料錯誤!");

                        string userNo = "";
                        foreach (var item in result)
                        {
                            userNo = item.UserNo;
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE BAS.[User] SET
                                Password = @Password,
                                PasswordStatus = @PasswordStatus,
                                PasswordMistake = @PasswordMistake,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE UserId = @UserId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Password = BaseHelper.Sha256Encrypt(userNo.ToString().ToLower()),
                                PasswordStatus = "Y",
                                PasswordMistake = 0,
                                LastModifiedDate,
                                LastModifiedBy,
                                UserId
                            });
                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "密碼重置成功!"
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

        #region //UpdatePasswordMistake -- 密碼錯誤累積 -- Ben Ma 2022.12.26
        public string UpdatePasswordMistake(int UserId)
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
                                FROM BAS.[User]
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", UserId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【使用者】資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE BAS.[User] SET
                                PasswordMistake = PasswordMistake + 1,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE UserId = @UserId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                LastModifiedDate,
                                LastModifiedBy,
                                UserId
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

        #region //UpdatePasswordMistakeReset
        public string UpdatePasswordMistakeReset(int UserId)
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
                                FROM BAS.[User]
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", UserId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【使用者】資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE BAS.[User] SET
                                PasswordMistake = @PasswordMistake,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE UserId = @UserId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                PasswordMistake = 0,
                                LastModifiedDate,
                                LastModifiedBy,
                                UserId
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

        #region //UpdateUserLoginKey -- 使用者登入金鑰 -- Zoey 2022.09.30
        public string UpdateUserLoginKey(int UserId, string UserNo)
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
                        sql = @"DELETE BAS.UserLoginKey
                                WHERE UserId = @UserId
                                AND ExpirationDate <= @Now";
                        dynamicParameters.Add("UserId", UserId);
                        dynamicParameters.Add("Now", DateTime.Now);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        if (Login == null || LoginKey == null)
                        {
                            string keyText = BaseHelper.Sha256Encrypt(UserNo + DateTime.Now.ToString("yyyyMMdd") + BaseHelper.ClientIP() + DateTime.Now.ToString("HHmmss"));

                            Login = new HttpCookie("Login")
                            {
                                Value = UserNo,
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
                            //dynamicParameters = new DynamicParameters();
                            //sql = @"DELETE BAS.UserLoginKey
                            //        WHERE LoginIP = @LoginIP";
                            //dynamicParameters.Add("LoginIP", BaseHelper.ClientIP());

                            //rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //新增金鑰
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO BAS.UserLoginKey (UserId, KeyText, ExpirationDate, LoginIP
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.KeyId
                                    VALUES (@UserId, @KeyText, @ExpirationDate, @LoginIP
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    UserId,
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

        #region //UpdateCompany -- 公司資料更新 -- Ben Ma 2022.04.19
        public string UpdateCompany(int CompanyId, string CompanyName, string Telephone, string Fax, string Address
            , string AddressEn, int LogoIcon)
        {
            try
            {
                if (CompanyName.Length <= 0) throw new SystemException("【公司名稱】不能為空!");
                if (CompanyName.Length > 100) throw new SystemException("【公司名稱】長度錯誤!");
                if (Telephone.Length > 30) throw new SystemException("【電話】長度錯誤!");
                if (Fax.Length > 30) throw new SystemException("【傳真】長度錯誤!");
                if (Address.Length > 200) throw new SystemException("【中文地址】長度錯誤!");
                if (AddressEn.Length > 200) throw new SystemException("【英文地址】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷公司資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Company
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【公司】資料錯誤!");
                        #endregion

                        #region //判斷公司名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Company
                                WHERE CompanyName = @CompanyName
                                AND CompanyId != @CompanyId";
                        dynamicParameters.Add("CompanyName", CompanyName);
                        dynamicParameters.Add("CompanyId", CompanyId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【公司名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE BAS.Company SET
                                CompanyName = @CompanyName,
                                Telephone = @Telephone,
                                Fax = @Fax,
                                Address = @Address,
                                AddressEn = @AddressEn,
                                LogoIcon = @LogoIcon,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE CompanyId = @CompanyId";
                        var parametersObject = new
                        {
                            CompanyName,
                            Telephone,
                            Fax,
                            Address,
                            AddressEn,
                            LogoIcon = LogoIcon > 0 ? LogoIcon : (int?)null,
                            LastModifiedDate,
                            LastModifiedBy,
                            CompanyId
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

        #region //UpdateCompanyStatus -- 公司狀態更新 -- Ben Ma 2022.04.19
        public string UpdateCompanyStatus(int CompanyId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷公司資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM BAS.Company
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【公司】資料錯誤!");

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
                        sql = @"UPDATE BAS.Company SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE CompanyId = @CompanyId";
                        var parametersObject = new
                        {
                            Status = status,
                            LastModifiedDate,
                            LastModifiedBy,
                            CompanyId
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

        #region //UpdateDepartment -- 部門資料更新 -- Ben Ma 2022.04.19
        public string UpdateDepartment(int DepartmentId, int CompanyId, string DepartmentName)
        {
            try
            {
                if (CompanyId <= 0) throw new SystemException("【所屬公司】不能為空!");
                if (DepartmentName.Length <= 0) throw new SystemException("【部門名稱】不能為空!");
                if (DepartmentName.Length > 100) throw new SystemException("【部門名稱】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷公司資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Company
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【公司】資料錯誤!");
                        #endregion

                        #region //判斷部門資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Department
                                WHERE DepartmentId = @DepartmentId";
                        dynamicParameters.Add("DepartmentId", DepartmentId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("【部門】資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE BAS.Department SET
                                CompanyId = @CompanyId,
                                DepartmentName = @DepartmentName,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE DepartmentId = @DepartmentId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId,
                                DepartmentName,
                                LastModifiedDate,
                                LastModifiedBy,
                                DepartmentId
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

        #region //UpdateDepartmentStatus -- 部門狀態更新 -- Ben Ma 2022.04.19
        public string UpdateDepartmentStatus(int DepartmentId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷部門資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM BAS.Department
                                WHERE DepartmentId = @DepartmentId";
                        dynamicParameters.Add("DepartmentId", DepartmentId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【部門】資料錯誤!");

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
                        sql = @"UPDATE BAS.Department SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE DepartmentId = @DepartmentId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                DepartmentId
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

        #region //UpdateDepartmentRateStatus -- 部門費用率狀態更新 -- Chia Yuan 2023.6.27
        public string UpdateDepartmentRateStatus(int DepartmentRateId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷部門資料是否正確

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 DepartmentId, EnableDate, DisabledDate, Status
                                FROM BAS.DepartmentRate
                                WHERE DepartmentRateId = @DepartmentRateId";
                        dynamicParameters.Add("DepartmentRateId", DepartmentRateId);

                        var result = sqlConnection.Query(sql, dynamicParameters).FirstOrDefault();
                        if (result == null) throw new SystemException("【部門費用率】資料錯誤!");

                        int departmentId = result.DepartmentId;
                        string status = result.Status;
                        DateTime? enableDate = result.EnableDate;
                        DateTime? disabledDate = result.DisabledDate;

                        #endregion

                        switch (status)
                        {
                            case "A":
                                status = "S";
                                disabledDate = LastModifiedDate;
                                break;
                            case "S":
                                status = "A";
                                enableDate = LastModifiedDate;
                                break;
                        }


                        dynamicParameters = new DynamicParameters();
                        sql = @"
update a set a.LastModifiedDate = @LastModifiedDate, a.LastModifiedBy = @LastModifiedBy
, a.DisabledDate = @EnableDate, a.Status = 'S'
from BAS.DepartmentRate a where a.DepartmentId = @DepartmentId and a.DepartmentRateId <> @DepartmentRateId and a.Status = 'A' 

update a set a.LastModifiedDate = @LastModifiedDate, a.LastModifiedBy = @LastModifiedBy
, a.EnableDate = @EnableDate, a.DisabledDate = @DisabledDate, a.Status = @Status
from BAS.DepartmentRate a where a.DepartmentRateId = @DepartmentRateId";

                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                DepartmentRateId,
                                DepartmentId = departmentId,
                                Status = status,
                                EnableDate = enableDate,
                                DisabledDate = disabledDate,
                                LastModifiedDate,
                                LastModifiedBy,
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

        #region UpdateDepartmentRate -- 部門費用率更新 --Chia Yuan 2023.6.28
        public string UpdateDepartmentRate(int DepartmentRateId, int AuthorId, decimal ResourceRate, decimal OverheadRate)
        {
            try
            {
                if (AuthorId <= 0) throw new SystemException("【編寫者】不能為空!");
                //todo: 判斷是否為0  待確認
                //ResourceRate
                //OverheadRate

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷編寫者資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[User]
                                WHERE UserId = @AuthorId";
                        dynamicParameters.Add("AuthorId", AuthorId);

                        var result = sqlConnection.Query(sql, dynamicParameters).FirstOrDefault();
                        if (result == null) throw new SystemException("【編寫者】資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE a SET 
a.Author = @AuthorId
, a.ResourceRate = @ResourceRate
, a.OverheadRate = @OverheadRate
, a.CreateBy = @CreateBy
, a.LastModifiedBy = @LastModifiedBy
FROM BAS.DepartmentRate a
WHERE a.DepartmentRateId = @DepartmentRateId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                DepartmentRateId,
                                AuthorId,
                                ResourceRate,
                                OverheadRate,
                                CreateBy,
                                LastModifiedBy = CreateBy
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

        #region //UpdateUser -- 使用者資料更新 -- Ben Ma 2022.04.18
        public string UpdateUser(int UserId, int DepartmentId, string UserName, string Gender, string Email)
        {
            try
            {
                if (DepartmentId <= 0) throw new SystemException("【所屬部門】不能為空!");
                if (UserName.Length <= 0) throw new SystemException("【使用者名稱】不能為空!");
                if (UserName.Length > 30) throw new SystemException("【使用者名稱】長度錯誤!");
                if (Gender.Length <= 0) throw new SystemException("【性別】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷部門資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Department
                                WHERE DepartmentId = @DepartmentId";
                        dynamicParameters.Add("DepartmentId", DepartmentId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【部門】資料錯誤!");
                        #endregion

                        #region //判斷使用者資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[User]
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", UserId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("【使用者】資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE BAS.[User] SET
                                DepartmentId = @DepartmentId,
                                UserName = @UserName,
                                Gender = @Gender,
                                Email = @Email,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE UserId = @UserId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                DepartmentId,
                                UserName,
                                Gender,
                                Email,
                                LastModifiedDate,
                                LastModifiedBy,
                                UserId
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

        #region //UpdateUserStatus -- 使用者狀態更新 -- Ben Ma 2022.04.18
        public string UpdateUserStatus(int UserId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷使用者資料是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM BAS.[User]
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", UserId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【使用者】資料錯誤!");

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
                        sql = @"UPDATE BAS.[User] SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE UserId = @UserId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                UserId
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

        #region //UpdateSystem -- 系統別資料更新 -- Ben Ma 2022.05.05
        public string UpdateSystem(int SystemId, string SystemCode, string SystemName, string ThemeIcon)
        {
            try
            {
                if (SystemCode.Length <= 0) throw new SystemException("【系統代碼】不能為空!");
                if (SystemCode.Length > 10) throw new SystemException("【系統代碼】長度錯誤!");
                if (SystemName.Length <= 0) throw new SystemException("【系統名稱】不能為空!");
                if (SystemName.Length > 100) throw new SystemException("【系統名稱】長度錯誤!");
                if (ThemeIcon.Length > 100) throw new SystemException("【主題Icon】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷系統別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[System]
                                WHERE SystemId = @SystemId";
                        dynamicParameters.Add("SystemId", SystemId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【系統別】資料錯誤!");
                        #endregion

                        #region //判斷系統代碼是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[System]
                                WHERE SystemCode = @SystemCode
                                AND SystemId != @SystemId";
                        dynamicParameters.Add("SystemCode", SystemCode);
                        dynamicParameters.Add("SystemId", SystemId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【系統代碼】重複，請重新輸入!");
                        #endregion

                        #region //判斷系統名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[System]
                                WHERE SystemName = @SystemName
                                AND SystemId != @SystemId";
                        dynamicParameters.Add("SystemName", SystemName);
                        dynamicParameters.Add("SystemId", SystemId);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() > 0) throw new SystemException("【系統名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE BAS.[System] SET
                                SystemCode = @SystemCode,
                                SystemName = @SystemName,
                                ThemeIcon = @ThemeIcon,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SystemId = @SystemId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SystemCode,
                                SystemName,
                                ThemeIcon,
                                LastModifiedDate,
                                LastModifiedBy,
                                SystemId
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

        #region //UpdateSystemStatus -- 系統別狀態更新 -- Ben Ma 2022.05.05
        public string UpdateSystemStatus(int SystemId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷系統別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM BAS.[System]
                                WHERE SystemId = @SystemId";
                        dynamicParameters.Add("SystemId", SystemId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【系統別】資料錯誤!");

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
                        sql = @"UPDATE BAS.[System] SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SystemId = @SystemId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                SystemId
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

        #region //UpdateSystemSort -- 系統別順序調整 -- Ben Ma 2022.05.05
        public string UpdateSystemSort(string SystemList)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int totalRowsAffected = 0;
                        #region //更新成負數避免觸發唯一值
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE BAS.[System] SET
                                SortNumber = SortNumber * -1,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy";
                        dynamicParameters.Add("LastModifiedDate", LastModifiedDate);
                        dynamicParameters.Add("LastModifiedBy", LastModifiedBy);

                        totalRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        int rowsAffected = 0;
                        #region //更新順序
                        string[] systemSort = SystemList.Split(',');

                        for (int i = 0; i < systemSort.Length; i++)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE BAS.[System] SET
                                    SortNumber = @SortNumber,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE SystemId = @SystemId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    SortNumber = (i + 1),
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    SystemId = Convert.ToInt32(systemSort[i])
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        }
                        #endregion

                        if (totalRowsAffected != rowsAffected) throw new SystemException("請先清除搜尋條件，再進行排序!");

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

        #region //UpdateModule -- 模組別資料更新 -- Kan 2022.05.05
        public string UpdateModule(int ModuleId, int SystemId, string ModuleCode, string ModuleName, string ThemeIcon)
        {
            try
            {
                if (SystemId < 0) throw new SystemException("【所屬系統】不能為空!");
                if (ModuleCode.Length <= 0) throw new SystemException("【模組代碼】不能為空!");
                if (ModuleCode.Length > 50) throw new SystemException("【模組代碼】長度錯誤!");
                if (ModuleName.Length <= 0) throw new SystemException("【模組名稱】不能為空!");
                if (ModuleName.Length > 100) throw new SystemException("【模組名稱】長度錯誤!");
                if (ThemeIcon.Length > 100) throw new SystemException("【主題Icon】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷模組別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Module
                                WHERE ModuleId = @ModuleId";
                        dynamicParameters.Add("ModuleId", ModuleId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【模組別】資料錯誤!");
                        #endregion

                        #region //判斷系統別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[System]
                                WHERE SystemId = @SystemId";
                        dynamicParameters.Add("SystemId", SystemId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("【系統別】資料錯誤!");
                        #endregion

                        #region //判斷模組代碼是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Module
                                WHERE SystemId = @SystemId
                                AND ModuleCode = @ModuleCode
                                AND ModuleId != @ModuleId";
                        dynamicParameters.Add("SystemId", SystemId);
                        dynamicParameters.Add("ModuleCode", ModuleCode);
                        dynamicParameters.Add("ModuleId", ModuleId);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() > 0) throw new SystemException("【模組代碼】重複，請重新輸入!");
                        #endregion

                        #region //判斷模組名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Module
                                WHERE SystemId = @SystemId
                                AND ModuleName = @ModuleName
                                AND ModuleId != @ModuleId";
                        dynamicParameters.Add("SystemId", SystemId);
                        dynamicParameters.Add("ModuleName", ModuleName);
                        dynamicParameters.Add("ModuleId", ModuleId);

                        var result4 = sqlConnection.Query(sql, dynamicParameters);
                        if (result4.Count() > 0) throw new SystemException("【模組名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE BAS.Module SET
                                SystemId = @SystemId,
                                ModuleCode = @ModuleCode,
                                ModuleName = @ModuleName,
                                ThemeIcon = @ThemeIcon,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ModuleId = @ModuleId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SystemId,
                                ModuleCode,
                                ModuleName,
                                ThemeIcon,
                                LastModifiedDate,
                                LastModifiedBy,
                                ModuleId
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

        #region //UpdateModuleStatus -- 模組別狀態更新 -- Kan 2022.05.05
        public string UpdateModuleStatus(int ModuleId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷模組別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM BAS.Module
                                WHERE ModuleId = @ModuleId";
                        dynamicParameters.Add("ModuleId", ModuleId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【模組別】資料錯誤!");

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
                        sql = @"UPDATE BAS.Module SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ModuleId = @ModuleId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                ModuleId
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

        #region //UpdateModuleSort -- 模組別順序調整 -- Kan 2022.05.05
        public string UpdateModuleSort(int SystemId, string ModuleList)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷系統別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[System]
                                WHERE SystemId = @SystemId";
                        dynamicParameters.Add("SystemId", SystemId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【系統別】資料錯誤!");
                        #endregion

                        int totalRowsAffected = 0;
                        #region //更新成負數避免觸發唯一值
                        sql = @"UPDATE BAS.Module SET
                                SortNumber = SortNumber * -1,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE SystemId = @SystemId";
                        dynamicParameters.Add("LastModifiedDate", LastModifiedDate);
                        dynamicParameters.Add("LastModifiedBy", LastModifiedBy);
                        dynamicParameters.Add("SystemId", SystemId);

                        totalRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        int rowsAffected = 0;
                        #region //更新順序
                        string[] moduleSort = ModuleList.Split(',');

                        for (int i = 0; i < moduleSort.Length; i++)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE BAS.Module SET
                                    SortNumber = @SortNumber,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE ModuleId = @ModuleId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    SortNumber = (i + 1),
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    ModuleId = Convert.ToInt32(moduleSort[i])
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        }
                        #endregion

                        if (totalRowsAffected != rowsAffected) throw new SystemException("請先清除搜尋條件，再進行排序!");

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

        #region //UpdateFunction -- 功能別資料更新 -- Kan 2022.05.06
        public string UpdateFunction(int FunctionId, int ModuleId, string FunctionCode, string FunctionName, string UrlTarget)
        {
            try
            {
                if (ModuleId < 0) throw new SystemException("【所屬模組】不能為空!");
                if (FunctionCode.Length <= 0) throw new SystemException("【功能代碼】不能為空!");
                if (FunctionCode.Length > 50) throw new SystemException("【功能代碼】長度錯誤!");
                if (FunctionName.Length <= 0) throw new SystemException("【功能名稱】不能為空!");
                if (FunctionName.Length > 100) throw new SystemException("【功能名稱】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷模組別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Module
                                WHERE ModuleId = @ModuleId";
                        dynamicParameters.Add("ModuleId", ModuleId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【模組別】資料錯誤!");
                        #endregion

                        #region //判斷功能代碼是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Function]
                                WHERE ModuleId = @ModuleId
                                AND FunctionCode = @FunctionCode
                                AND FunctionId != @FunctionId";
                        dynamicParameters.Add("ModuleId", ModuleId);
                        dynamicParameters.Add("FunctionCode", FunctionCode);
                        dynamicParameters.Add("FunctionId", FunctionId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【功能代碼】重複，請重新輸入!");
                        #endregion

                        #region //判斷功能名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Function]
                                WHERE ModuleId = @ModuleId
                                AND FunctionName = @FunctionName
                                AND FunctionId != @FunctionId";
                        dynamicParameters.Add("ModuleId", ModuleId);
                        dynamicParameters.Add("FunctionName", FunctionName);
                        dynamicParameters.Add("FunctionId", FunctionId);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() > 0) throw new SystemException("【功能名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE BAS.[Function] SET
                                ModuleId = @ModuleId,
                                FunctionCode = @FunctionCode,
                                FunctionName = @FunctionName,
                                UrlTarget = @UrlTarget,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE FunctionId = @FunctionId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ModuleId,
                                FunctionCode,
                                FunctionName,
                                UrlTarget,
                                LastModifiedDate,
                                LastModifiedBy,
                                FunctionId
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

        #region //UpdateFunctionStatus -- 功能別狀態更新 -- Kan 2022.05.06
        public string UpdateFunctionStatus(int FunctionId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷功能別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM BAS.[Function]
                                WHERE FunctionId = @FunctionId";
                        dynamicParameters.Add("FunctionId", FunctionId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【功能別】資料錯誤!");

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
                        sql = @"UPDATE BAS.[Function] SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE FunctionId = @FunctionId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                FunctionId
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

        #region //UpdateFunctionSort -- 功能別順序調整 -- Kan 2022.05.06
        public string UpdateFunctionSort(int ModuleId, string FunctionList)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷模組別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Module
                                WHERE ModuleId = @ModuleId";
                        dynamicParameters.Add("ModuleId", ModuleId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【模組別】資料錯誤!");
                        #endregion

                        int totalRowsAffected = 0;
                        #region //更新成負數避免觸發唯一值
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE BAS.[Function] SET
                                SortNumber = SortNumber * -1,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ModuleId = @ModuleId";
                        dynamicParameters.Add("LastModifiedDate", LastModifiedDate);
                        dynamicParameters.Add("LastModifiedBy", LastModifiedBy);
                        dynamicParameters.Add("ModuleId", ModuleId);

                        totalRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        int rowsAffected = 0;
                        #region //更新順序
                        string[] functionSort = FunctionList.Split(',');

                        for (int i = 0; i < functionSort.Length; i++)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE BAS.[Function] SET
                                    SortNumber = @SortNumber,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE FunctionId = @FunctionId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    SortNumber = (i + 1),
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    FunctionId = Convert.ToInt32(functionSort[i])
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        }
                        #endregion

                        if (totalRowsAffected != rowsAffected) throw new SystemException("請先清除搜尋條件，再進行排序!");

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

        #region //UpdateFunctionDetail -- 詳細功能別資料更新 -- Kan 2022.05.10
        public string UpdateFunctionDetail(int DetailId, int FunctionId, string DetailCode, string DetailName)
        {
            try
            {
                if (FunctionId < 0) throw new SystemException("【所屬功能】不能為空!");
                if (DetailCode.Length <= 0) throw new SystemException("【詳細功能代碼】不能為空!");
                if (DetailCode.Length > 50) throw new SystemException("【詳細功能代碼】長度錯誤!");
                if (DetailName.Length <= 0) throw new SystemException("【詳細功能名稱】不能為空!");
                if (DetailName.Length > 100) throw new SystemException("【詳細功能名稱】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷功能別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Function]
                                WHERE FunctionId = @FunctionId";
                        dynamicParameters.Add("FunctionId", FunctionId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【功能別】資料錯誤!");
                        #endregion

                        #region //判斷詳細功能代碼是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.FunctionDetail
                                WHERE FunctionId = @FunctionId
                                AND DetailCode = @DetailCode
                                AND DetailId != @DetailId";
                        dynamicParameters.Add("FunctionId", FunctionId);
                        dynamicParameters.Add("DetailCode", DetailCode);
                        dynamicParameters.Add("DetailId", DetailId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【詳細功能代碼】重複，請重新輸入!");
                        #endregion

                        #region //判斷詳細功能名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.FunctionDetail
                                WHERE FunctionId = @FunctionId
                                AND DetailName = @DetailName
                                AND DetailId != @DetailId";
                        dynamicParameters.Add("FunctionId", FunctionId);
                        dynamicParameters.Add("DetailName", DetailName);
                        dynamicParameters.Add("DetailId", DetailId);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() > 0) throw new SystemException("【詳細功能名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE BAS.FunctionDetail SET
                                FunctionId = @FunctionId,
                                DetailCode = @DetailCode,
                                DetailName = @DetailName,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE DetailId = @DetailId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                FunctionId,
                                DetailCode,
                                DetailName,
                                LastModifiedDate,
                                LastModifiedBy,
                                DetailId
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

        #region //UpdateFunctionDetailStatus -- 詳細功能別狀態更新 -- Kan 2022.05.10
        public string UpdateFunctionDetailStatus(int DetailId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷詳細功能別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM BAS.FunctionDetail
                                WHERE DetailId = @DetailId";
                        dynamicParameters.Add("DetailId", DetailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【詳細功能別】資料錯誤!");

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
                        sql = @"UPDATE BAS.FunctionDetail SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE DetailId = @DetailId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                DetailId
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

        #region //UpdateFunctionDetailSort -- 詳細功能別順序調整 -- Kan 2022.05.10
        public string UpdateFunctionDetailSort(int FunctionId, string FunctionDetailList)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷功能別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Function]
                                WHERE FunctionId = @FunctionId";
                        dynamicParameters.Add("FunctionId", FunctionId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【功能別】資料錯誤!");
                        #endregion

                        int totalRowsAffected = 0;
                        #region //更新成負數避免觸發唯一值
                        sql = @"UPDATE BAS.FunctionDetail SET
                                SortNumber = SortNumber * -1,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE FunctionId = @FunctionId";
                        dynamicParameters.Add("LastModifiedDate", LastModifiedDate);
                        dynamicParameters.Add("LastModifiedBy", LastModifiedBy);
                        dynamicParameters.Add("FunctionId", FunctionId);

                        totalRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        int rowsAffected = 0;
                        #region //更新順序
                        string[] functionDetailSort = FunctionDetailList.Split(',');

                        for (int i = 0; i < functionDetailSort.Length; i++)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE BAS.FunctionDetail SET
                                    SortNumber = @SortNumber,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE DetailId = @DetailId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    SortNumber = (i + 1),
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    DetailId = Convert.ToInt32(functionDetailSort[i])
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        }
                        #endregion

                        if (totalRowsAffected != rowsAffected) throw new SystemException("請先清除搜尋條件，再進行排序!");

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

        #region //UpdateStatus -- 狀態資料更新 -- Zoey 2022.05.18
        public string UpdateStatus(string StatusSchema, string StatusNo, string StatusName, string StatusDesc)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷狀態資料是否存在
                        bool exist = false;

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Status
                                WHERE StatusSchema = @StatusSchema
                                AND StatusNo = @StatusNo";
                        dynamicParameters.Add("StatusSchema", StatusSchema);
                        dynamicParameters.Add("StatusNo", StatusNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        exist = result.Count() > 0;
                        #endregion

                        int rowsAffected = 0;
                        if (!exist)
                        {
                            #region //新增
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO BAS.Status (StatusSchema, StatusNo, StatusName, StatusDesc
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@StatusSchema, @StatusNo, @StatusName, @StatusDesc
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    StatusSchema,
                                    StatusNo,
                                    StatusName,
                                    StatusDesc,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }
                        else
                        {
                            #region //修改
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE BAS.Status SET
                                    StatusName = @StatusName,
                                    StatusDesc = @StatusDesc,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE StatusSchema = @StatusSchema
                                    AND StatusNo = @StatusNo";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    StatusName,
                                    StatusDesc,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    StatusSchema,
                                    StatusNo,
                                });

                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //UpdateStatusCompany -- 狀態公司對應資料更新 -- Ben Ma 2022.10.17
        public string UpdateStatusCompany(int StatusId, int CompanyId, string StatusName, string StatusDesc)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷狀態資料是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Status]
                                WHERE StatusId = @StatusId";
                        dynamicParameters.Add("StatusId", StatusId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("狀態資料錯誤!");
                        #endregion

                        #region //判斷狀態資料是否存在
                        bool exist = false;

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.StatusCompany
                                WHERE StatusId = @StatusId
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("StatusId", StatusId);
                        dynamicParameters.Add("CompanyId", CompanyId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        exist = result2.Count() > 0;
                        #endregion

                        int rowsAffected = 0;
                        if (!exist)
                        {
                            #region //新增
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO BAS.StatusCompany (StatusId, CompanyId, StatusName, StatusDesc
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@StatusId, @CompanyId, @StatusName, @StatusDesc
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    StatusId,
                                    CompanyId,
                                    StatusName,
                                    StatusDesc,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }
                        else
                        {

                            if (StatusName.Length <= 0 || StatusDesc.Length <= 0)
                            {
                                #region //刪除
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE BAS.StatusCompany
                                        WHERE StatusId = @StatusId
                                        AND CompanyId = @CompanyId";
                                dynamicParameters.Add("StatusId", StatusId);
                                dynamicParameters.Add("CompanyId", CompanyId);

                                rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }
                            else
                            {
                                #region //修改
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE BAS.StatusCompany SET
                                        StatusName = @StatusName,
                                        StatusDesc = @StatusDesc,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE StatusId = @StatusId
                                        AND CompanyId = @CompanyId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        StatusName,
                                        StatusDesc,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        StatusId,
                                        CompanyId
                                    });

                                rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
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

        #region //UpdateType -- 類別資料更新 -- Zoey 2022.05.18
        public string UpdateType(string TypeSchema, string TypeNo, string TypeName, string TypeDesc)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷類別資料是否存在
                        bool exist = false;

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Type
                                WHERE TypeSchema = @TypeSchema
                                AND TypeNo = @TypeNo";
                        dynamicParameters.Add("TypeSchema", TypeSchema);
                        dynamicParameters.Add("TypeNo", TypeNo);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        exist = result3.Count() > 0;
                        #endregion

                        int rowsAffected = 0;

                        if (!exist)
                        {
                            #region //新增
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO BAS.Type (TypeSchema, TypeNo, TypeName, TypeDesc
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@TypeSchema, @TypeNo, @TypeName, @TypeDesc
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    TypeSchema,
                                    TypeNo,
                                    TypeName,
                                    TypeDesc,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }
                        else
                        {
                            #region //修改
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE BAS.Type SET
                                    TypeName = @TypeName,
                                    TypeDesc = @TypeDesc,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE TypeSchema = @TypeSchema
                                    AND TypeNo = @TypeNo";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    TypeName,
                                    TypeDesc,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    TypeSchema,
                                    TypeNo
                                });

                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //UpdateTypeCompany -- 類別公司對應資料更新 -- Ben Ma 2022.10.17
        public string UpdateTypeCompany(int TypeId, int CompanyId, string TypeName, string TypeDesc)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷類別資料是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Type]
                                WHERE TypeId = @TypeId";
                        dynamicParameters.Add("TypeId", TypeId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("類別資料錯誤!");
                        #endregion

                        #region //判斷狀態資料是否存在
                        bool exist = false;

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.TypeCompany
                                WHERE TypeId = @TypeId
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("TypeId", TypeId);
                        dynamicParameters.Add("CompanyId", CompanyId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        exist = result2.Count() > 0;
                        #endregion

                        int rowsAffected = 0;
                        if (!exist)
                        {
                            #region //新增
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO BAS.TypeCompany (TypeId, CompanyId, TypeName, TypeDesc
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@TypeId, @CompanyId, @TypeName, @TypeDesc
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    TypeId,
                                    CompanyId,
                                    TypeName,
                                    TypeDesc,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }
                        else
                        {

                            if (TypeName.Length <= 0 || TypeDesc.Length <= 0)
                            {
                                #region //刪除
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE BAS.TypeCompany
                                        WHERE TypeId = @TypeId
                                        AND CompanyId = @CompanyId";
                                dynamicParameters.Add("TypeId", TypeId);
                                dynamicParameters.Add("CompanyId", CompanyId);

                                rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }
                            else
                            {
                                #region //修改
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE BAS.TypeCompany SET
                                        TypeName = @TypeName,
                                        TypeDesc = @TypeDesc,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE TypeId = @TypeId
                                        AND CompanyId = @CompanyId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        TypeName,
                                        TypeDesc,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        TypeId,
                                        CompanyId
                                    });

                                rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
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

        #region //UpdateShift -- 班次資料更新 -- Zoey 2022.05.19
        public string UpdateShift(int ShiftId, string ShiftName, string WorkBeginTime, string WorkEndTime, string WorkHours)
        {
            try
            {
                if (ShiftName.Length <= 0) throw new SystemException("【班次名稱】不能為空!");
                if (ShiftName.Length > 100) throw new SystemException("【班次名稱】長度錯誤!");
                if (WorkBeginTime.Length <= 0) throw new SystemException("【上班時間】不能為空!");
                if (WorkBeginTime.Length > 8) throw new SystemException("【上班時間】長度錯誤!");
                if (WorkEndTime.Length <= 0) throw new SystemException("【下班時間】不能為空!");
                if (WorkEndTime.Length > 8) throw new SystemException("【下班時間】長度錯誤!");
                if (WorkHours.Length <= 0) throw new SystemException("【工作時數】不能為空!");
                if (WorkHours.Length > 100) throw new SystemException("【工作時數】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷班次資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Shift]
                                WHERE ShiftId = @ShiftId";
                        dynamicParameters.Add("ShiftId", ShiftId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("班次資料錯誤!");
                        #endregion

                        #region //判斷班次名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Shift]
                                WHERE ShiftName = @ShiftName
                                AND ShiftId != @ShiftId";
                        dynamicParameters.Add("ShiftName", ShiftName);
                        dynamicParameters.Add("ShiftId", ShiftId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【班次名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE BAS.[Shift] SET
                                ShiftName = @ShiftName,
                                WorkBeginTime = @WorkBeginTime,
                                WorkEndTime = @WorkEndTime,
                                WorkHours = @WorkHours,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ShiftId = @ShiftId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ShiftName,
                                WorkBeginTime,
                                WorkEndTime,
                                WorkHours,
                                LastModifiedDate,
                                LastModifiedBy,
                                ShiftId
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

        #region //UpdateSendResetMail -- 寄送使用者忘記密碼信件 -- Ann 2024-06-28
        public string UpdateSendResetMail(string UserNo)
        {
            try
            {
                int rowsAffected = 0;
                string userToken = "";
                if (UserNo.Length <= 0) throw new SystemException("【工號】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷使用者資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserId, a.UserName
                                , c.CompanyNo
                                FROM BAS.[User] a 
                                INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                                INNER JOIN BAS.Company c ON b.CompanyId = c.CompanyId
                                WHERE a.UserNo = @UserNo";
                        dynamicParameters.Add("UserNo", UserNo);

                        var UserResult = sqlConnection.Query(sql, dynamicParameters);
                        if (UserResult.Count() <= 0) throw new SystemException("使用者資料錯誤!");

                        int UserId = -1;
                        string UserName = "";
                        string CompanyNo = "";
                        foreach (var item in UserResult)
                        {
                            UserId = item.UserId;
                            UserName = item.UserName;
                            CompanyNo = item.CompanyNo;
                        }
                        #endregion

                        #region //取得此使用者之密鑰資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.TokenId
                                FROM BAS.UserToken a 
                                WHERE a.UserId = @UserId";
                        dynamicParameters.Add("UserId", UserId);

                        var UserTokenResult = sqlConnection.Query(sql, dynamicParameters);

                        CreateGuId();

                        if (UserTokenResult.Count() > 0)
                        {
                            #region //原本已存在密鑰則更新密鑰及日期
                            int tokenId = -1;
                            foreach (var item in UserTokenResult)
                            {
                                tokenId = item.TokenId;
                            }
                            #endregion

                            #region //Update BAS.UserToke
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE BAS.UserToken SET
                                    Token = @Token,
                                    ValidDate = @ValidDate,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE TokenId = @TokenId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    Token = userToken,
                                    ValidDate = DateTime.Now,
                                    LastModifiedDate,
                                    LastModifiedBy = UserId,
                                    TokenId = tokenId
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }
                        else
                        {
                            #region //Insert BAS.UserToken
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO BAS.UserToken (UserId, Token, ValidDate
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.TokenId
                                    VALUES (@UserId, @Token, @ValidDate
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    UserId,
                                    Token = userToken,
                                    ValidDate = DateTime.Now,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy = UserId,
                                    LastModifiedBy = UserId
                                });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();
                            #endregion
                        }
                        #endregion

                        #region //產生密鑰
                        void CreateGuId()
                        {
                            string guId = Guid.NewGuid().ToString("N");

                            // 如果需要的長度大於guidString的長度，則重新生成並連接
                            while (guId.Length < 20)
                            {
                                guId += Guid.NewGuid().ToString("N");
                            }

                            userToken = guId;
                        }
                        #endregion

                        #region //寄送Mamo訊息
                        #region //MAMO推播通知
                        string Content = "";

                        Content = "### 【BM企業管理平台 重新設定密碼認證】\n" +
                                    "您好!系統接收到您申請重新設定密碼之請求，若非您本人所申請請忽略，謝謝!\n" +
                                    "請注意，請勿隨意將下方連結分享給他人，以維護帳號安全!\n" +
                                    "- 請點選下方連結進行密碼重新設定，此連結有效期限為10分鐘:\n" +
                                    "https://bm.zy-tech.com.tw/User/Login?Token=" + userToken + "";
                                    //"http://192.168.134.33:16668/User/Login?Token=" + userToken + "";

                        List<string> Tags = new List<string>();
                        List<int> Files = new List<int>();

                        string MamoResult = mamoHelper.SendMessage(CompanyNo, UserId, "Personal", UserNo, Content, Tags, Files);

                        JObject MamoResultJson = JObject.Parse(MamoResult);
                        if (MamoResultJson["status"].ToString() != "success")
                        {
                            throw new SystemException(MamoResultJson["msg"].ToString());
                        }
                        #endregion
                        #endregion

                        #region //確認公司別DB
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.HrmDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        dynamicParameters.Add("CompanyNo", CompanyNo);

                        var companyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (companyResult.Count() <= 0) throw new SystemException("公司別錯誤!");

                        foreach (var item in companyResult)
                        {
                            HrmConnectionStrings = ConfigurationManager.AppSettings[item.HrmDb];
                        }
                        #endregion

                        #region //寄送私人信箱
                        using (SqlConnection sqlConnection2 = new SqlConnection(HrmConnectionStrings))
                        {
                            #region //取得私人信箱資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT Email 
                                    FROM dbo.Employee a 
                                    WHERE Code = @Code";
                            dynamicParameters.Add("Code", UserNo);

                            var EmployeeResult = sqlConnection2.Query(sql, dynamicParameters);

                            string Email = "";
                            foreach (var item in EmployeeResult)
                            {
                                Email = item.Email;
                            }
                            #endregion

                            if (Email != null && Email.Length > 0)
                            {
                                #region //設定寄送格式
                                string mailSubject = "【中揚光電-企業管理平台】使用者密碼重新設定通知信";
                                string mailContent = "【中揚光電-企業管理平台 重新設定密碼認證】<br>" +
                                                        "您好!系統接收到您申請重新設定密碼之請求，若非您本人所申請請忽略，謝謝!<br>" +
                                                        "<br>" +
                                                        "<b>請注意，請勿隨意將下方連結分享給他人，以維護帳號安全!</b><br>" +
                                                        "<br>" +
                                                        "請點選下方連結進行密碼重新設定，此連結有效期限為10分鐘:<br>" +
                                                        "https://bm.zy-tech.com.tw/User/Login?Token=" + userToken + "<br>" + 
                                                        "<br>" +
                                                        "<i>本通訊及其所有附件所含之資訊均屬機密，僅供指定之收件人使用，未經寄件人許可不得揭露、複製或散布本通訊。若您並非指定之收件人，請勿使用、保存或揭露本通訊之任何部份，並即通知寄件人且完全刪除本通訊及其所有附件。如為指定收件者，應確實保護郵件中本公司之營業機密及個人資料，不得任意傳佈或揭露，並應自行確認本郵件之附檔與超連結之安全性，以共同善盡資訊安全與個資保護責任。寄件人並不保證本網際網路通訊內所載數據資料或其他資訊之完整性及正確性。因此，寄件人對於他人變更、修改、竄改或偽造之本通訊內容，及對於本通訊內容若有不當或不完整之傳遞、傳遞接收之遲延或是對您的電腦造成毀損之情況，恕不負任何責任。網路通訊可能含有病毒，收件人應自行確認本郵件是否安全及未有被病毒入侵、被攔截或是被干擾的情況，若因此造成損害，寄件人恕不負責。</i>";
                                #endregion

                                #region //寄送Mail
                                MailConfig mailConfig = new MailConfig
                                {
                                    Host = "192.168.20.252",
                                    Port = 25,
                                    SendMode = 0,
                                    From = "企業管理平台:jmo-service@zy-tech.com.tw",
                                    Subject = mailSubject,
                                    Account = "jmo-service",
                                    Password = "asdf!QAZ",
                                    MailTo = UserName + ":" + Email,
                                    MailCc = "",
                                    MailBcc = "",
                                    HtmlBody = mailContent,
                                    TextBody = "-"
                                };
                                BaseHelper.MailSend(mailConfig);
                                #endregion
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

        #region //UpdateResetPassword -- 使用者修改密碼登入 -- Ann 2024-06-28
        public string UpdateResetPassword(string Token, string Password)
        {
            try
            {
                int rowsAffected = 0;

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //密碼參數設定
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.PasswordExpiration, a.PasswordFormat, a.PasswordWrongCount
                                FROM BAS.PasswordSetting a";

                        var resultPasswordSetting = sqlConnection.Query(sql, dynamicParameters);
                        if (resultPasswordSetting.Count() <= 0) throw new SystemException("【密碼參數設定】資料錯誤!");

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

                        if (!Regex.IsMatch(Password, PasswordFormat, RegexOptions.IgnoreCase)) throw new SystemException("【新密碼】格式錯誤!");

                        #region //判斷使用者Token資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserId, a.ValidDate
                                , b.UserNo
                                FROM BAS.UserToken a 
                                INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                                WHERE a.Token = @Token ";
                        dynamicParameters.Add("Token", Token);

                        var UserTokenResult = sqlConnection.Query(sql, dynamicParameters);
                        if (UserTokenResult.Count() <= 0) throw new SystemException("使用者Token資料錯誤!");

                        int UserId = -1;
                        string UserNo = "";
                        DateTime ValidDate = new DateTime();
                        foreach (var item in UserTokenResult)
                        {
                            UserId = item.UserId;
                            UserNo = item.UserNo;
                            ValidDate = item.ValidDate;

                            DateTime currentDateTime = DateTime.Now;
                            TimeSpan timeDifference = currentDateTime - ValidDate;
                            if (timeDifference.TotalMinutes > 10)
                            {
                                throw new SystemException("此Token已失效!!");
                            }
                        }
                        #endregion

                        #region //Update BAS.User
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE BAS.[User] SET
                                Password = @Password,
                                PasswordStatus = @PasswordStatus,
                                PasswordExpire = @PasswordExpire,
                                PasswordMistake = @PasswordMistake,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE UserId = @UserId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Password = BaseHelper.Sha256Encrypt(Password),
                                PasswordStatus = "N",
                                PasswordExpire = DateTime.Now.AddMinutes(PasswordExpiration),
                                PasswordMistake = 0,
                                LastModifiedDate,
                                LastModifiedBy,
                                UserId
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //將Token失效
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE BAS.UserToken SET
                                Token = '',
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE UserId = @UserId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                LastModifiedDate,
                                LastModifiedBy,
                                UserId
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            UserId,
                            UserNo
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

        #region //UpdateUploadWhitelistManagement -- 白名單上傳路徑資料更新 -- Andrew 2024-10-04
        public string UpdateUploadWhitelistManagement(int ListId, string ListNo, int DepartmentId, string FolderPath)
        {
            try
            {
                if (ListNo.Length <= 0) throw new SystemException("【白名單類別】不能為空!");
                if (ListNo.Length > 20) throw new SystemException("【白名單類別】長度錯誤!");
                if (DepartmentId <= 0) throw new SystemException("【所屬部門】不能為空!");
                if (FolderPath.Length <= 0) throw new SystemException("【資料夾路徑】不能為空!");
                if (FolderPath.Length > 100) throw new SystemException("【資料夾路徑】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷白名單資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1  
                                FROM BAS.UploadWhitelist a
                                INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                                WHERE a.ListId = @ListId
                                AND b.CompanyId = @CompanyId";
                        dynamicParameters.Add("ListId", ListId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("白名單資料錯誤!");
                        #endregion


                        #region //判斷白名單類別、部門名稱及資料夾路徑是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.DepartmentId 
                                FROM BAS.UploadWhitelist a
                                INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                                WHERE a.ListNo = @ListNo
                                AND b.CompanyId = @CompanyId
                                AND a.DepartmentId = @DepartmentId
                                AND a.FolderPath = @FolderPath
                                AND a.ListId != @ListId";
                        dynamicParameters.Add("ListNo", ListNo);
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("DepartmentId", DepartmentId);
                        dynamicParameters.Add("FolderPath", FolderPath);
                        dynamicParameters.Add("ListId", ListId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("已有此筆資料，請透過搜尋查找相關資訊!");
                        #endregion

                        #region //更新SQL
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE BAS.UploadWhitelist SET
                                ListNo = @ListNo,
                                DepartmentId = @DepartmentId,
                                FolderPath = @FolderPath,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ListId = @ListId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ListNo,
                                DepartmentId,
                                FolderPath,
                                LastModifiedDate,
                                LastModifiedBy,
                                ListId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //Delete
        #region //DeleteStatus -- 狀態資料刪除 -- Zoey 2022.05.18
        public string DeleteStatus(string StatusSchema, string StatusNo)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷狀態資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Status
                                WHERE StatusSchema = @StatusSchema
                                AND StatusNo = @StatusNo";
                        dynamicParameters.Add("StatusSchema", StatusSchema);
                        dynamicParameters.Add("StatusNo", StatusNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("狀態資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE BAS.Status
                                WHERE StatusSchema = @StatusSchema
                                AND StatusNo = @StatusNo";
                        dynamicParameters.Add("StatusSchema", StatusSchema);
                        dynamicParameters.Add("StatusNo", StatusNo);

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

        #region //DeleteType -- 類別資料刪除 -- Zoey 2022.05.18
        public string DeleteType(string TypeSchema, string TypeNo)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷類別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.Type
                                WHERE TypeSchema = @TypeSchema
                                AND TypeNo = @TypeNo";
                        dynamicParameters.Add("TypeSchema", TypeSchema);
                        dynamicParameters.Add("TypeNo", TypeNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("類別資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE BAS.Type
                                WHERE TypeSchema = @TypeSchema
                                AND TypeNo = @TypeNo";
                        dynamicParameters.Add("TypeSchema", TypeSchema);
                        dynamicParameters.Add("TypeNo", TypeNo);

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

        #region //DeleteShift -- 班次資料刪除 -- Zoey 2022.05.19
        public string DeleteShift(int ShiftId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷班次資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Shift]
                                WHERE ShiftId = @ShiftId";
                        dynamicParameters.Add("ShiftId", ShiftId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("班次資料錯誤!");
                        #endregion

                        #region //判斷哪些部門使用此班次資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT (
                                    SELECT b.DepartmentNo + '-' + b.DepartmentName + ','
                                    FROM BAS.DepartmentShift a
                                    INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                                    WHERE a.ShiftId = @ShiftId
                                    ORDER BY b.DepartmentNo
                                    FOR XML PATH('')
                                ) Department
                                WHERE EXISTS (
                                    SELECT TOP 1 1
                                    FROM BAS.DepartmentShift a
                                    WHERE a.ShiftId = @ShiftId
                                )";
                        dynamicParameters.Add("ShiftId", ShiftId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0)
                        {
                            string exceptionMessage = "";

                            foreach (var item in result2)
                            {
                                exceptionMessage += item.Department;
                            }

                            exceptionMessage = exceptionMessage.Remove(exceptionMessage.Length - 1, 1);

                            throw new SystemException("【" + exceptionMessage + "】" + "已選擇此班次，無法刪除!");
                        }
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE BAS.[Shift]
                                WHERE ShiftId = @ShiftId";
                        dynamicParameters.Add("ShiftId", ShiftId);

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

        #region //DeleteDepartmentRateDetail -- 部門費用率詳細資料刪除 --Chia Yuan --230627
        public string DeleteDepartmentRateDetail(int DepartmentRateId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.DepartmentRate
                                WHERE  DepartmentRateId = @DepartmentRateId";
                        dynamicParameters.Add("@DepartmentRateId", DepartmentRateId);

                        var result = sqlConnection.Query(sql, dynamicParameters).FirstOrDefault();
                        if (result == null) throw new SystemException("【部門費用率】資料錯誤!");
                        #endregion

                        #region //判斷資料是否被其他表使用

                        #endregion

                        int rowsAffected = 0;
                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE BAS.DepartmentRate
                                WHERE  DepartmentRateId = @DepartmentRateId";
                        dynamicParameters.Add("DepartmentRateId", DepartmentRateId);

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

        #region //DeleteDepartmentShift -- 部門班次資料刪除 -- Zoey 2022.05.20
        public string DeleteDepartmentShift(int DepartmentId, int ShiftId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷部門班次資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.DepartmentShift
                                WHERE DepartmentId = @DepartmentId
                                AND ShiftId = @ShiftId";
                        dynamicParameters.Add("DepartmentId", DepartmentId);
                        dynamicParameters.Add("ShiftId", ShiftId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("部門班次資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE BAS.DepartmentShift
                                WHERE DepartmentId = @DepartmentId
                                AND ShiftId = @ShiftId";
                        dynamicParameters.Add("DepartmentId", DepartmentId);
                        dynamicParameters.Add("ShiftId", ShiftId);

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

        #region //DeleteUserLoginKey -- 使用者登入金鑰刪除 -- Zoey 2022.09.30
        public string DeleteUserLoginKey(string UserNo, string KeyText)
        {
            try
            {
                if (UserNo.Length <= 0) throw new SystemException("【使用者編號】錯誤!");
                if (KeyText.Length <= 0) throw new SystemException("【金鑰】錯誤!");

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
                                AND a.KeyText = @KeyText";
                        dynamicParameters.Add("UserNo", UserNo);
                        dynamicParameters.Add("KeyText", KeyText);

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
                                AND a.KeyText = @KeyText";
                        dynamicParameters.Add("UserNo", UserNo);
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

        #region //DeleteUploadWhitelistManagement -- 白名單上傳路徑資料刪除 -- Andrew 2024-10-04
        public string DeleteUploadWhitelistManagement(int ListId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷白名單資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.DepartmentId
                                FROM BAS.UploadWhitelist a
                                INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                                WHERE a.ListId = @ListId
                                AND b.CompanyId = @CompanyId";
                        dynamicParameters.Add("ListId", ListId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("白名單資料錯誤!");
                        #endregion

                        #region //刪除SQL - 主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE BAS.UploadWhitelist
                                WHERE ListId = @ListId";
                        dynamicParameters.Add("ListId", ListId);

                        int rowsAffected = 0;
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

        #region //FOR EIP API

        #region //GetTypeEIP -- 取得類別資料 -- GPAI 2024.03.15
        public string GetTypeEIP(string TypeSchema, string TypeNo, string TypeName
            , string OrderBy, int PageIndex, int PageSize, int[] CustomerIds)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {

                    if (CustomerIds == null)
                    {
                        throw new SystemException("客戶資料尚未綁定");

                    }

                    sqlQuery.mainKey = "a.TypeId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.TypeSchema, a.TypeNo, a.TypeName, a.TypeDesc
                        , ISNULL(b.TypeName, a.TypeName) ApplyTypeName, ISNULL(b.TypeDesc, a.TypeDesc) ApplyTypeDesc";
                    sqlQuery.mainTables =
                        @"FROM BAS.[Type] a
                         OUTER APPLY (
                            SELECT ba.TypeName, ba.TypeDesc
                            FROM BAS.TypeCompany ba
							INNER JOIN BAS.Company y on ba.CompanyId = y.CompanyId
                           INNER JOIN SCM.Customer z on y.CompanyId = z.CompanyId
                            WHERE ba.TypeId = a.TypeId
                            --AND ba.CompanyId = @CompanyId
							AND z.CustomerId IN @CustomerIds
                        ) b";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    dynamicParameters.Add("CustomerIds", CustomerIds);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TypeSchema", @" AND a.TypeSchema = @TypeSchema", TypeSchema);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TypeNo", @" AND a.TypeNo = @TypeNo", TypeNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TypeName", @" AND (a.TypeName LIKE '%' + @TypeName + '%' OR a.TypeDesc LIKE '%' + @TypeName + '%')", TypeName);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.TypeSchema, a.TypeNo";
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

        #region //GetStatusEIP -- 取得狀態資料 -- GPAI 2024.03.15
        public string GetStatusEIP(string StatusSchema, string StatusNo, string StatusName
            , int[] CustomerIds
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                if (CustomerIds == null) throw new SystemException("客戶資料尚未綁定");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.StatusId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.StatusSchema, a.StatusNo, a.StatusName, a.StatusDesc
                        , ISNULL(b.StatusName, a.StatusName) ApplyStatusName, ISNULL(b.StatusDesc, a.StatusDesc) ApplyStatusDesc";
                    sqlQuery.mainTables =
                        @"FROM BAS.[Status] a
                        OUTER APPLY (
                            SELECT ba.StatusName, ba.StatusDesc
                            FROM BAS.StatusCompany ba
                            INNER JOIN BAS.Company y on ba.CompanyId = y.CompanyId
                            INNER JOIN SCM.Customer z on y.CompanyId = z.CompanyId
                            WHERE ba.StatusId = a.StatusId
                            --AND ba.CompanyId = @CompanyId
							AND z.CustomerId IN @CustomerIds
                        ) b";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    dynamicParameters.Add("CustomerIds", CustomerIds);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StatusSchema", @" AND a.StatusSchema = @StatusSchema", StatusSchema);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StatusNo", @" AND a.StatusNo = @StatusNo", StatusNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StatusName", @" AND (a.StatusName LIKE '%' + @StatusName + '%' OR a.StatusDesc LIKE '%' + @StatusName + '%')", StatusName);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.StatusSchema, a.StatusNo";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);

                    //sql = @"SELECT a.StatusId, a.StatusSchema, a.StatusNo, a.StatusName, a.StatusDesc
                    //        , ISNULL(b.StatusName, a.StatusName) ApplyStatusName, ISNULL(b.StatusDesc, a.StatusDesc) ApplyStatusDesc
                    //        FROM BAS.[Status] a
                    //        left join BAS.StatusCompany b on a.StatusId = b.StatusId
                    //       ";
                   

                    //BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "StatusSchema", @" AND a.StatusSchema = @StatusSchema", StatusSchema);
                    //BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "StatusNo", @" AND a.StatusNo = @StatusNo", StatusNo);
                    //BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "StatusName", @" AND (a.StatusName LIKE '%' + @StatusName + '%' OR a.StatusDesc LIKE '%' + @StatusName + '%')", StatusName);


                    //var result = sqlConnection.Query(sql, dynamicParameters);

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
    }
}
