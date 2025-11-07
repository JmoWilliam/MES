using Dapper;
using Helpers;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Transactions;
using System.Web;

namespace FINDA
{
    public class DerManagementDA
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

        public DerManagementDA()
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
        #region //GetDepartmentRate -- 取得部門費用率資料 -- Ann 2023-11-29
        public string GetDepartmentRate(int DepartmentRateId, int DepartmentId, int UserId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.DepartmentRateId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.DepartmentId, a.Author, a.ResourceRate, a.OverheadRate, a.EnableDate, a.DisabledDate
                        , b.DepartmentNo + b.DepartmentName DepartmentFullNo
                        , c.UserNo CreateUserNo, c.UserName CreateUserName";
                    sqlQuery.mainTables =
                        @"FROM BAS.DepartmentRate a 
                        INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                        INNER JOIN BAS.[User] c ON a.CreateBy = c.UserId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DepartmentRateId", @" AND a.DepartmentRateId = @DepartmentRateId", DepartmentRateId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DepartmentId", @" AND a.DepartmentId = @DepartmentId", DepartmentId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserId", @" AND a.CreateBy = @UserId", UserId);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.DepartmentRateId DESC";
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
        #region //AddDepartmentRate -- 新增部門費用率資料 -- Ann 2023-11-29
        public string AddDepartmentRate(int DepartmentId, double ResourceRate, double OverheadRate)
        {
            try
            {
                if (ResourceRate < 0) throw new SystemException("【人工費用率】不能為空!");
                if (OverheadRate < 0) throw new SystemException("【製造費用率】不能為空!");

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
                        if (result.Count() <= 0) throw new SystemException("部門資料錯誤!");
                        #endregion

                        #region //INSERT BAS.DepartmentRate
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO BAS.DepartmentRate (DepartmentId, Author, ResourceRate, OverheadRate
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.DepartmentRateId
                                VALUES (@DepartmentId, @Author, @ResourceRate, @OverheadRate
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                DepartmentId,
                                Author = CreateBy,
                                ResourceRate,
                                OverheadRate,
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
        #region //UpdateDepartmentRate -- 更新部門費用率資料 -- Ann 2023-11-29
        public string UpdateDepartmentRate(int DepartmentRateId, int DepartmentId, double ResourceRate, double OverheadRate)
        {
            try
            {
                if (ResourceRate < 0) throw new SystemException("【人工費用率】不能為空!");
                if (OverheadRate < 0) throw new SystemException("【製造費用率】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //檢查部門費用資料是否有誤
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.DepartmentRate a 
                                WHERE DepartmentRateId = @DepartmentRateId";
                        dynamicParameters.Add("DepartmentRateId", DepartmentRateId);

                        var DepartmentRateResult = sqlConnection.Query(sql, dynamicParameters);
                        if (DepartmentRateResult.Count() <= 0) throw new SystemException("部門費用率資料錯誤!!");
                        #endregion

                        #region //UPDATE MES.QcRecord
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE BAS.DepartmentRate SET
                                DepartmentId = @DepartmentId,
                                ResourceRate = @ResourceRate,
                                OverheadRate = @OverheadRate,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE DepartmentRateId = @DepartmentRateId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                DepartmentId,
                                ResourceRate,
                                OverheadRate,
                                LastModifiedDate,
                                LastModifiedBy,
                                DepartmentRateId
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
        #endregion

        #region //Delete
        #region //DeleteDepartmentRate -- 刪除部門費用率資料 -- Ann 20203-11-29
        public string DeleteDepartmentRate(int DepartmentRateId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //檢查部門費用資料是否有誤
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.DepartmentRate a 
                                WHERE DepartmentRateId = @DepartmentRateId";
                        dynamicParameters.Add("DepartmentRateId", DepartmentRateId);

                        var DepartmentRateResult = sqlConnection.Query(sql, dynamicParameters);
                        if (DepartmentRateResult.Count() <= 0) throw new SystemException("部門費用率資料錯誤!!");
                        #endregion

                        #region //DELETE
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE BAS.DepartmentRate
                                WHERE DepartmentRateId = @DepartmentRateId";
                        dynamicParameters.Add("DepartmentRateId", DepartmentRateId);

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
        #endregion

    }
}
