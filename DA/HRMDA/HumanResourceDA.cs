using Dapper;
using Helpers;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Transactions;
using System.Web;

namespace HRMDA
{
    public class HumanResourceDA
    {
        public string MainConnectionStrings = "";
        public string ErpConnectionStrings = "";
        public string HrmConnectionStrings = "";
        public string HrmEtergeConnectionStrings = "";

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

        public HumanResourceDA()
        {
            MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];
            ErpConnectionStrings = ConfigurationManager.AppSettings["ErpDb"];
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
        #region//GetDepartmentManagerNames --取得課、部、處主管姓名
        public string GetDepartmentManagerNames(string Company, string UserNo,string DepNo)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int CompanyId = -1;
                    string HrmConnectionStrings = ""
                        , CompanyName = "";

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷公司資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 CompanyId, CompanyName, HrmDb
                                FROM BAS.Company
                                WHERE CompanyNo = @CompanyNo";
                        dynamicParameters.Add("CompanyNo", Company);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【公司】資料錯誤!");

                        foreach (var item in result)
                        {
                            CompanyId = Convert.ToInt32(item.CompanyId);
                            CompanyName = item.CompanyName;
                            HrmConnectionStrings = ConfigurationManager.AppSettings[item.HrmDb];
                        }
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(HrmConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();

                        sql = @"WITH ManagerHierarchy AS (
                            -- 基礎案例：從許智豪開始
                            SELECT 
                                e.EmployeeId,
                                e.Code,
                                e.CnName,
                                e.DirectorId,
                                0 AS Level -- 從0開始，0代表本人
                            FROM 
                                Employee e
                                INNER JOIN Department d ON e.DepartmentId=d.DepartmentId
                            WHERE 
                                (e.Code = @UserNo OR d.Code=@DepNo)
	                          and e.EmployeeId != e.DirectorId
    
                            UNION ALL
    
                            -- 遞迴案例：往上查找主管
                            SELECT 
                                manager.EmployeeId,
                                manager.Code,
                                manager.CnName,
                                manager.DirectorId,
                                mh.Level + 1 AS Level
                            FROM 
                                ManagerHierarchy mh
                                INNER JOIN Employee manager ON mh.DirectorId = manager.EmployeeId
                            WHERE 
                                mh.DirectorId is not null -- 確保有上級主管
		                        and mh.EmployeeId != mh.DirectorId
                        )

                        -- 查詢結果，排除許本人（Level > 0）
                        SELECT 
                            Level,
                            Code,
                            CnName
                        FROM 
                            ManagerHierarchy
                        WHERE
                            Level > 0 -- 只顯示主管，不包括許本人
                        ORDER BY 
                            Level;
                    ";
                        dynamicParameters.Add("UserNo", UserNo);
                        dynamicParameters.Add("DepNo", DepNo);
                        var result = sqlConnection.Query(sql, dynamicParameters);

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

        #region //Add
        #endregion

        #region //Update
        #region //UpdateDepartmentSynchronize -- 部門資料更新 -- Ben Ma 2022.04.21
        public string UpdateDepartmentSynchronize(string CompanyNo, string UpdateDate)
        {
            try
            {
                List<Department> departments = new List<Department>();

                if (CompanyNo.Length <= 0) throw new SystemException("【公司】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int CompanyId = -1;
                    string HrmConnectionStrings = ""
                        , CompanyName = "";

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷公司資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 CompanyId, CompanyName, HrmDb
                                FROM BAS.Company
                                WHERE CompanyNo = @CompanyNo";
                        dynamicParameters.Add("CompanyNo", CompanyNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【公司】資料錯誤!");

                        foreach (var item in result)
                        {
                            CompanyId = Convert.ToInt32(item.CompanyId);
                            CompanyName = item.CompanyName;
                            HrmConnectionStrings = ConfigurationManager.AppSettings[item.HrmDb];
                        }
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(HrmConnectionStrings))
                    {
                        #region //撈取HRM部門資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CODE, a.CODE DepartmentNo, a.Name DepartmentName
                                , CASE a.Flag 
                                    WHEN 0 THEN 'S'
                                    WHEN 1 THEN 'A'
                                END Status
                                FROM Department a
                                INNER JOIN Corporation b ON b.CorporationId = a.CorporationId
                                WHERE 1=1
                                AND LTRIM(RTRIM(b.ShortName)) = @CompanyName";
                        dynamicParameters.Add("CompanyName", CompanyName);
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "UpdateDate", @" AND (a.BeginEndDate_EndDate > @UpdateDate OR a.BeginEndDate_EndDate IS NULL)", UpdateDate);
                        sql += @" ORDER BY a.CODE";

                        departments = sqlConnection.Query<Department>(sql, dynamicParameters).ToList();
                        #endregion
                    }

                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷BAS.Department是否有資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT DepartmentId, DepartmentNo
                                FROM BAS.Department
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<Department> resultDepartment = sqlConnection.Query<Department>(sql, dynamicParameters).ToList();

                        departments = departments.GroupJoin(resultDepartment, x => x.DepartmentNo, y => y.DepartmentNo, (x, y) => { x.DepartmentId = y.FirstOrDefault()?.DepartmentId; return x; }).ToList();
                        #endregion

                        List<Department> addDepartments = departments.Where(x => x.DepartmentId == null).ToList();
                        List<Department> updateDepartments = departments.Where(x => x.DepartmentId != null).ToList();

                        #region //新增
                        if (addDepartments.Count > 0)
                        {
                            addDepartments
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.CompanyId = CompanyId;
                                    x.SystemStatus = "A";
                                    x.CreateDate = CreateDate;
                                    x.LastModifiedDate = LastModifiedDate;
                                    x.CreateBy = CreateBy;
                                    x.LastModifiedBy = LastModifiedBy;
                                });

                            sql = @"INSERT INTO BAS.Department (CompanyId, DepartmentNo, DepartmentName
                                    , SystemStatus, Status
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@CompanyId, @DepartmentNo, @DepartmentName
                                    , @SystemStatus, @Status
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            rowsAffected += sqlConnection.Execute(sql, addDepartments);
                        }
                        #endregion

                        #region //修改
                        if (updateDepartments.Count > 0)
                        {
                            updateDepartments
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.LastModifiedDate = LastModifiedDate;
                                    x.LastModifiedBy = LastModifiedBy;
                                });

                            sql = @"UPDATE BAS.Department SET
                                    DepartmentName = @DepartmentName,
                                    Status = @Status,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE DepartmentId = @DepartmentId";
                            rowsAffected += sqlConnection.Execute(sql, updateDepartments);
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

        #region //UpdateUserSynchronize -- 使用者資料同步 -- Ben Ma 2022.04.21
        public string UpdateUserSynchronize(string CompanyNo, string UpdateDate, string UserNo)
        {
            try
            {
                List<User> users = new List<User>();

                if (CompanyNo.Length <= 0) throw new SystemException("【公司】不能為空!");

                if (UserNo == null) UserNo = "";

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int CompanyId = -1;
                    string HrmConnectionStrings = ""
                        , CompanyName = "";

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷公司資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 CompanyId, CompanyName, HrmDb
                                FROM BAS.Company
                                WHERE CompanyNo = @CompanyNo";
                        dynamicParameters.Add("CompanyNo", CompanyNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【公司】資料錯誤!");

                        foreach (var item in result)
                        {
                            CompanyId = Convert.ToInt32(item.CompanyId);
                            CompanyName = item.CompanyName;
                            HrmConnectionStrings = ConfigurationManager.AppSettings[item.HrmDb];
                        }
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(HrmConnectionStrings))
                    {
                        #region //撈取HRM使用者資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.Code, a.Code UserNo, a.CnName UserName, ISNULL(a.CharacterTest, a.[Weight]) Email,a.MobilePhone
                                , CASE a.EmployeeStateId 
                                    WHEN 'EmployeeState1001' THEN 'T'
                                    WHEN 'EmployeeState2001' THEN 'F'
                                    WHEN 'EmployeeState3001' THEN 'S'
                                    WHEN 'EmployeeState3002' THEN 'R'
                                    WHEN 'EmployeeState3003' THEN 'L'
                                    ELSE 'P'
                                END UserStatus
                                , CASE a.EmployeeStateId 
                                    WHEN 'EmployeeState1001' THEN 'A'
                                    WHEN 'EmployeeState2001' THEN 'A'
                                    WHEN 'EmployeeState3001' THEN 'S'
                                    WHEN 'EmployeeState3002' THEN 'S'
                                    WHEN 'EmployeeState3003' THEN 'S'
                                    ELSE 'A'
                                END Status
                                , b.CODE DepartmentNo
                                , CASE c.ScName
                                    WHEN '女' THEN 'F'
                                    WHEN '男' THEN 'M'
                                    ELSE 'N'
                                END Gender
                                , d.Name Job, e.Name JobType
                                FROM Employee a
                                LEFT JOIN Department b ON b.DepartmentId = a.DepartmentId
                                LEFT JOIN CodeInfo c ON c.CodeInfoId = a.GenderId
                                LEFT JOIN Job d ON d.JobId = a.JobId
                                LEFT JOIN Position e ON e.PositionId = d.PositionId
                                INNER JOIN Corporation f ON a.CorporationId = f.CorporationId
                                WHERE 1=1
                                AND LTRIM(RTRIM(f.ShortName)) = @CompanyName";
                        dynamicParameters.Add("CompanyName", CompanyName);
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "UpdateDate", @" AND a.LastModifiedDate > @UpdateDate", UpdateDate);
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "UserNo", @" AND a.Code = @UserNo", UserNo);
                        sql += @" ORDER BY a.CODE";

                        users = sqlConnection.Query<User>(sql, dynamicParameters).ToList();
                        #endregion
                    }

                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //撈取部門ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT DepartmentId, DepartmentNo
                                FROM BAS.Department
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<User> resultDepartment = sqlConnection.Query<User>(sql, dynamicParameters).ToList();

                        users = users.Join(resultDepartment, x => x.DepartmentNo, y => y.DepartmentNo, (x, y) => { x.DepartmentId = y.DepartmentId; return x; }).ToList();
                        #endregion

                        #region //判斷BAS.User是否有資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT UserId, UserNo
                                FROM BAS.[User] a
                                INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                                WHERE b.CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<User> resultUsers = sqlConnection.Query<User>(sql, dynamicParameters).ToList();

                        users = users.GroupJoin(resultUsers, x => x.UserNo, y => y.UserNo, (x, y) => { x.UserId = y.FirstOrDefault()?.UserId; return x; }).ToList();
                        #endregion

                        List<User> addUsers = users.Where(x => x.UserId == null).ToList();
                        List<User> updateUsers = users.Where(x => x.UserId != null).ToList();

                        #region //新增
                        if (addUsers.Count > 0)
                        {
                            addUsers
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.Password = BaseHelper.Sha256Encrypt(x.UserNo.ToString().ToLower());
                                    x.SystemStatus = "A";
                                    x.PasswordStatus = "Y";
                                    x.CreateDate = CreateDate;
                                    x.LastModifiedDate = LastModifiedDate;
                                    x.CreateBy = CreateBy;
                                    x.LastModifiedBy = LastModifiedBy;
                                });

                            sql = @"INSERT INTO BAS.[User] (DepartmentId, UserNo, UserName, Gender
                                    , Email,MobilePhone, Password, Job, JobType
                                    , UserStatus, SystemStatus, PasswordStatus, Status
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@DepartmentId, @UserNo, @UserName, @Gender
                                    , @Email,@MobilePhone, @Password, @Job, @JobType
                                    , @UserStatus, @SystemStatus, @PasswordStatus, @Status
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            rowsAffected += sqlConnection.Execute(sql, addUsers);

                        }
                        #endregion

                        #region //修改
                        if (updateUsers.Count > 0)
                        {
                            updateUsers
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.LastModifiedDate = LastModifiedDate;
                                    x.LastModifiedBy = LastModifiedBy;
                                });

                            sql = @"UPDATE BAS.[User] SET   
                                    DepartmentId = @DepartmentId,
                                    UserName = @UserName,
                                    Gender = @Gender,
                                    Email = @Email,
                                    MobilePhone=@MobilePhone,
                                    Job = @Job,
                                    JobType = @JobType,
                                    UserStatus = @UserStatus,
                                    Status = @Status,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE UserId = @UserId";
                            rowsAffected += sqlConnection.Execute(sql, updateUsers);
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

        #region //UpdateUserInfoSynchronize -- 新版使用者資料同步 -- Yi 2023.09.06
        public string UpdateUserInfoSynchronize(string CompanyNo, string UpdateDate)
        {
            try
            {
                List<UserInfo> userInfos = new List<UserInfo>();
                List<UserEmployee> userEmployees = new List<UserEmployee>();

                if (CompanyNo.Length <= 0) throw new SystemException("【公司】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int CompanyId = -1;
                    string HrmConnectionStrings = ""
                        , CompanyName = "";

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷公司資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 CompanyId, CompanyName, HrmDb
                                FROM BAS.Company
                                WHERE CompanyNo = @CompanyNo";
                        dynamicParameters.Add("CompanyNo", CompanyNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【公司】資料錯誤!");

                        foreach (var item in result)
                        {
                            CompanyId = Convert.ToInt32(item.CompanyId);
                            CompanyName = item.CompanyName;
                            HrmConnectionStrings = ConfigurationManager.AppSettings[item.HrmDb];
                        }
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(HrmConnectionStrings))
                    {
                        #region //撈取HRM使用者資料(BAS.UserInfo所需欄位)
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.Code InnerCode, a.CnName DisplayName
                                , CASE a.EmployeeStateId 
                                    WHEN 'EmployeeState1001' THEN 'A'
                                    WHEN 'EmployeeState2001' THEN 'A'
                                    WHEN 'EmployeeState3001' THEN 'S'
                                    WHEN 'EmployeeState3002' THEN 'S'
                                    WHEN 'EmployeeState3003' THEN 'S'
                                    ELSE 'A'
                                END Status
                                FROM Employee a
                                WHERE 1=1
                                AND LTRIM(RTRIM(f.ShortName)) = @CompanyName";
                        dynamicParameters.Add("CompanyName", CompanyName);
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "UpdateDate", @" AND a.LastModifiedDate > @UpdateDate", UpdateDate);
                        sql += @" ORDER BY a.Code";

                        userInfos = sqlConnection.Query<UserInfo>(sql, dynamicParameters).ToList();
                        #endregion

                        #region //撈取HRM使用者資料(BAS.UserEmployee所需欄位)
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.Code UserNo, ISNULL(a.CharacterTest, a.[Weight]) Email
                                , CASE a.EmployeeStateId 
                                    WHEN 'EmployeeState1001' THEN 'T'
                                    WHEN 'EmployeeState2001' THEN 'F'
                                    WHEN 'EmployeeState3001' THEN 'S'
                                    WHEN 'EmployeeState3002' THEN 'R'
                                    WHEN 'EmployeeState3003' THEN 'L'
                                    ELSE 'P'
                                END EmployeeStatus
                                , b.CODE DepartmentNo
                                , CASE c.ScName
                                    WHEN '女' THEN 'F'
                                    WHEN '男' THEN 'M'
                                    ELSE 'N'
                                END Gender
                                , d.Name Job, e.Name JobType
                                FROM Employee a
                                LEFT JOIN Department b ON b.DepartmentId = a.DepartmentId
                                LEFT JOIN CodeInfo c ON c.CodeInfoId = a.GenderId
                                LEFT JOIN Job d ON d.JobId = a.JobId
                                LEFT JOIN Position e ON e.PositionId = d.PositionId
                                INNER JOIN Corporation f ON a.CorporationId = f.CorporationId
                                WHERE 1=1
                                AND LTRIM(RTRIM(f.ShortName)) = @CompanyName";
                        dynamicParameters.Add("CompanyName", CompanyName);
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "UpdateDate", @" AND a.LastModifiedDate > @UpdateDate", UpdateDate);
                        sql += @" ORDER BY a.CODE";

                        userEmployees = sqlConnection.Query<UserEmployee>(sql, dynamicParameters).ToList();
                        #endregion
                    }

                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //撈取部門ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT DepartmentId, DepartmentNo
                                FROM BAS.Department
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<UserEmployee> resultDepartment = sqlConnection.Query<UserEmployee>(sql, dynamicParameters).ToList();

                        userEmployees = userEmployees.Join(resultDepartment, x => x.DepartmentNo, y => y.DepartmentNo, (x, y) => { x.DepartmentId = y.DepartmentId; return x; }).ToList();
                        #endregion

                        #region //判斷BAS.UserInfo是否有資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserId, a.InnerCode
                                FROM BAS.UserInfo a
                                INNER JOIN BAS.UserEmployee b ON b.UserId = a.UserId
                                INNER JOIN BAS.Department c ON b.DepartmentId = c.DepartmentId
                                WHERE c.CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<UserInfo> resultUsers = sqlConnection.Query<UserInfo>(sql, dynamicParameters).ToList();

                        userInfos = userInfos.GroupJoin(resultUsers, x => x.InnerCode, y => y.InnerCode, (x, y) => { x.UserId = y.FirstOrDefault()?.UserId; return x; }).ToList();
                        #endregion

                        #region //使用者資料(新增/修改)
                        List<UserInfo> addUserInfos = userInfos.Where(x => x.UserId == null).ToList();
                        List<UserInfo> updateUserInfos = userInfos.Where(x => x.UserId != null).ToList();

                        #region //新增
                        if (addUserInfos.Count > 0)
                        {
                            addUserInfos
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.Password = BaseHelper.Sha256Encrypt(x.InnerCode.ToString().ToLower());
                                    x.PasswordStatus = "Y";
                                    x.CreateDate = CreateDate;
                                    x.LastModifiedDate = LastModifiedDate;
                                    x.CreateBy = CreateBy;
                                    x.LastModifiedBy = LastModifiedBy;
                                });

                            sql = @"INSERT INTO BAS.UserInfo (UserId, UserType, Account, Password, InnerCode, DisplayName
                                    , PasswordStatus, PasswordMistake, Status
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@UserId, 'E', @InnerCode, @Password, @InnerCode, @DisplayName
                                    , @PasswordStatus, @PasswordMistake, @Status
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            rowsAffected += sqlConnection.Execute(sql, addUserInfos);

                        }
                        #endregion

                        #region //修改
                        if (updateUserInfos.Count > 0)
                        {
                            updateUserInfos
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.LastModifiedDate = LastModifiedDate;
                                    x.LastModifiedBy = LastModifiedBy;
                                });

                            sql = @"UPDATE BAS.UserInfo SET
                                    Account = @InnerCode,
                                    InnerCode = @InnerCode,
                                    DisplayName = @UserName,
                                    Status = @Status,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE UserId = @UserId";
                            rowsAffected += sqlConnection.Execute(sql, updateUserInfos);
                        }
                        #endregion

                        #endregion

                        #region //使用者內部員工設定檔資料(新增/修改)

                        #region //撈取BAS.UserInfo ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT UserId, InnerCode
                                FROM BAS.UserInfo
                                ORDER BY UserId";

                        resultUsers = sqlConnection.Query<UserInfo>(sql, dynamicParameters).ToList();

                        userEmployees = userEmployees.GroupJoin(resultUsers, x => x.UserNo, y => y.InnerCode, (x, y) => { x.UserId = y.FirstOrDefault()?.UserId; return x; }).Select(x => x).ToList();
                        #endregion

                        List<UserEmployee> addUserEmployees = userEmployees.Where(x => x.UserId == null).ToList();
                        List<UserEmployee> updateEmployees = userEmployees.Where(x => x.UserId != null).ToList();

                        #region //新增
                        if (addUserEmployees.Count > 0)
                        {
                            addUserEmployees
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.CreateDate = CreateDate;
                                    x.LastModifiedDate = LastModifiedDate;
                                    x.CreateBy = CreateBy;
                                    x.LastModifiedBy = LastModifiedBy;
                                });

                            sql = @"INSERT INTO BAS.UserEmployee (UserId, DepartmentId, Gender, Email
                                    , Job, JobType, EmployeeStatus
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@UserId, @DepartmentId, @Gender, @Email
                                    , @Job, @JobType, @EmployeeStatus
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            rowsAffected += sqlConnection.Execute(sql, addUserEmployees);
                        }
                        #endregion

                        #region //修改
                        if (updateEmployees.Count > 0)
                        {
                            updateEmployees
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.LastModifiedDate = LastModifiedDate;
                                    x.LastModifiedBy = LastModifiedBy;
                                });

                            sql = @"UPDATE BAS.UserEmployee SET
                                    DepartmentId = @DepartmentId,
                                    Gender = @Gender,
                                    Email = @Email,
                                    Job = @Job,
                                    JobType = @JobType,
                                    EmployeeStatus = @EmployeeStatus,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE UserId = @UserId";
                            rowsAffected += sqlConnection.Execute(sql, updateEmployees);
                        }
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
        #endregion

        #region //Delete
        #endregion
    }
}
