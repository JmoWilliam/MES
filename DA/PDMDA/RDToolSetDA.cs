using Dapper;
using Helpers;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Configuration;
using System.Data.SqlClient;
using System.Web;

namespace PDMDA
{
    public class RDToolSetDA
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

        public RDToolSetDA()
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
        #region //GetUser -- 取得使用者資料 -- Andrew 2025.02.13
        public string GetUser(int UserId, int DepartmentId, int CompanyId, string UserNo, string UserName)
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
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserNo", @" AND a.UserNo LIKE '%' + @UserNo + '%'", UserNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserName", @" AND a.UserName LIKE '%' + @UserName + '%'", UserName);
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

        #endregion
    }
}
