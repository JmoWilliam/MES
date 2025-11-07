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
using System.Text.RegularExpressions;
using System.Transactions;
using System.Web;

namespace QMSDA
{
    public class QmsBasicInformationDA
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

        public QmsBasicInformationDA()
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
        #region //GetDefect -- 取得異常狀態資料 -- Ann 2022-06-14
        public string GetDefect()
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sql = @"WITH MainQueryJson
                            AS
                            (
                                SELECT (a.GroupNo + '：' + a.GroupName) [name]
                                , a.GroupNo code, a.GroupName [call], b.StatusDesc [status], a.GroupDesc reason
                                , '' [owner], ('NgGroup' + CONVERT(varchar(50), a.GroupId)) id, 'Group' parent
                                FROM QMS.DefectGroup a
                                LEFT JOIN BAS.[Status] b ON a.Status = b.StatusNo AND b.StatusSchema = 'Status'
                                WHERE a.CompanyId = @CompanyId
                                UNION
                                SELECT (a.ClassNo + '：' + a.ClassName) [name]
                                , a.ClassNo code, a.ClassName [call], b.StatusDesc [status], a.ClassDesc reason
                                , '' [owner], ('NgClass' + CONVERT(varchar(50), a.ClassId)) id
                                , ('NgGroup' + CONVERT(varchar(50), a.GroupId)) parent
                                FROM QMS.DefectClass a
                                LEFT JOIN BAS.[Status] b ON a.Status = b.StatusNo AND b.StatusSchema = 'Status'
                                UNION
                                SELECT (a.CauseNo + '：' + a.CauseName) [name]
                                , a.CauseNo code, a.CauseName [call], b.StatusDesc [status], a.CauseDesc reason
                                , IIF(a.ResponsibleDepartment = -1,'',CONVERT(varchar(10), c.DepartmentNo) + '_' + c.DepartmentName) [owner]
                                , ('NgCause' + CONVERT(varchar(50), a.CauseId)) id
                                , ('NgClass' + CONVERT(varchar(50), a.ClassId)) parent
                                FROM QMS.DefectCause a
                                LEFT JOIN BAS.[Status] b ON a.Status = b.StatusNo AND b.StatusSchema = 'Status'
                                LEFT JOIN BAS.[Department] c ON a.ResponsibleDepartment = c.DepartmentId
                            )
                            SELECT *
                            FROM
                            (
                                SELECT 
                                (
                                    SELECT * FROM MainQueryJson
                                    FOR JSON PATH, ROOT('dataset')
                                ) as queryJson
                            ) a";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    #region //處理JSON結構
                    var result = sqlConnection.Query(sql, dynamicParameters);
                    var jsonString = "";
                    foreach (var item in result)
                    {
                        jsonString = item.queryJson;
                    }

                    var transferJson = JObject.Parse(jsonString);

                    var datasetCount = transferJson["dataset"].Count();
                    for (var i=0; i<datasetCount; i++)
                    {
                        if (transferJson["dataset"][i]["parent"].ToString() == "Group")
                        {
                            JObject dataFormat = JObject.FromObject(new
                            {
                                name = transferJson["dataset"][i]["name"],
                                code = transferJson["dataset"][i]["code"],
                                call = transferJson["dataset"][i]["call"],
                                status = transferJson["dataset"][i]["status"],
                                reason = transferJson["dataset"][i]["reason"],
                                owner = transferJson["dataset"][i]["owner"],
                                id = transferJson["dataset"][i]["id"]
                            });

                            transferJson["dataset"][i] = dataFormat;
                        }
                    }

                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = transferJson
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

        #region //GetNgGroup -- 取得異常群組 -- Ann 2022-06-15
        public string GetNgGroup(int GroupId, string Status)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.GroupId, a.GroupNo, a.GroupName
                            , (a.GroupNo + '-' + a.GroupName) as NgGroupWithText, a.[Status]
                            FROM QMS.DefectGroup a
                            WHERE 1=1
                            AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "GroupId", @" AND a.GroupId = @GroupId", GroupId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "Status", @" AND a.Status = @Status", Status);
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

        #region //GetNgClass -- 取得異常類別 -- Ann 2022-06-15
        public string GetNgClass(int GroupId, int ClassId, string Status)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.GroupId, a.ClassId, a.ClassNo, a.ClassName
                            , (a.ClassNo + '-' + a.ClassName) as NgClassWithText, a.[Status]
                            FROM QMS.DefectClass a
                            WHERE 1=1
                            AND a.Status = @Status";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "GroupId", @" AND a.GroupId = @GroupId", GroupId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ClassId", @" AND a.ClassId = @ClassId", ClassId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "Status", @" AND a.Status = @Status", Status);
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

        #region //GetNgCause -- 取得異常原因 -- Ann 2022-10-11
        public string GetNgCause(int CauseId, int ClassId, string CauseNo, string Status)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.CauseId, a.ClassId, a.CauseNo, a.CauseName, a.CauseDesc, a.ResponsibleDepartment
                            , a.[Status], (a.CauseNo + '-' + a.CauseName) NgCauseWithText
                            FROM QMS.DefectCause a
                            INNER JOIN QMS.DefectClass b on a.ClassId = b.ClassId
                            INNER JOIN QMS.DefectGroup b1 on b.GroupId = b1.GroupId
                            WHERE 1=1
                            AND b1.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CauseId", @" AND a.CauseId = @CauseId", CauseId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ClassId", @" AND a.ClassId = @ClassId", ClassId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "Status", @" AND a.Status = @Status", Status);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CauseNo", @" AND a.CauseNo = @CauseNo", CauseNo);
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

        #region //GetRepair -- 取得維修狀態資料 -- Ann 2022-06-21
        public string GetRepair()
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sql = @"WITH MainQueryJson
                            AS
                            (
                                SELECT (a.GroupNo + '：' + a.GroupName) [name]
                                , a.GroupNo code, a.GroupName [call], b.StatusDesc [status], a.GroupDesc reason
                                , '' [owner], ('RepairGroup' + CONVERT(varchar(50), a.GroupId)) id, 'Group' parent
                                FROM QMS.RepairGroup a
                                LEFT JOIN BAS.[Status] b ON a.Status = b.StatusNo AND b.StatusSchema = 'Status'
                                WHERE a.CompanyId = @CompanyId
                                UNION
                                SELECT (a.ClassNo + '：' + a.ClassName) [name]
                                , a.ClassNo code, a.ClassName [call], b.StatusDesc [status], a.ClassDesc reason
                                , '' [owner], ('RepairClass' + CONVERT(varchar(50), a.ClassId)) id
                                , ('RepairGroup' + CONVERT(varchar(50), a.GroupId)) parent
                                FROM QMS.RepairClass a
                                LEFT JOIN BAS.[Status] b ON a.Status = b.StatusNo AND b.StatusSchema = 'Status'
                                UNION
                                SELECT (a.CauseNo + '：' + a.CauseName) [name]
                                , a.CauseNo code, a.CauseName [call], b.StatusDesc [status], a.CauseDesc reason
                                , IIF(a.ResponsibleDepartment = -1,'',CONVERT(varchar(10), c.DepartmentNo) + '_' + c.DepartmentName) [owner]
                                , ('RepairCause' + CONVERT(varchar(50), a.CauseId)) id
                                , ('RepairClass' + CONVERT(varchar(50), a.ClassId)) parent
                                FROM QMS.RepairCause a
                                LEFT JOIN BAS.[Status] b ON a.Status = b.StatusNo AND b.StatusSchema = 'Status'
                                LEFT JOIN BAS.[Department] c ON a.ResponsibleDepartment = c.DepartmentId
                            )
                            SELECT *
                            FROM
                            (
                                SELECT 
                                (
                                    SELECT * FROM MainQueryJson
                                    FOR JSON PATH, ROOT('dataset')
                                ) as queryJson
                            ) a";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    #region //處理JSON結構
                    var result = sqlConnection.Query(sql, dynamicParameters);
                    var jsonString = "";
                    foreach (var item in result)
                    {
                        jsonString = item.queryJson;
                    }

                    var transferJson = JObject.Parse(jsonString);

                    var datasetCount = transferJson["dataset"].Count();
                    for (var i = 0; i < datasetCount; i++)
                    {
                        if (transferJson["dataset"][i]["parent"].ToString() == "Group")
                        {
                            JObject dataFormat = JObject.FromObject(new
                            {
                                name = transferJson["dataset"][i]["name"],
                                code = transferJson["dataset"][i]["code"],
                                call = transferJson["dataset"][i]["call"],
                                status = transferJson["dataset"][i]["status"],
                                reason = transferJson["dataset"][i]["reason"],
                                owner = transferJson["dataset"][i]["owner"],
                                id = transferJson["dataset"][i]["id"]
                            });

                            transferJson["dataset"][i] = dataFormat;
                        }
                    }

                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = transferJson
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

        #region //GetRepairGroup -- 取得維修群組資料 -- Ann 2022-06-21
        public string GetRepairGroup(string GroupNo, string GroupName)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.GroupId, a.GroupNo, a.GroupName 
                            , (a.GroupNo + '-' + a.GroupName) as RepairGroupWithText, a.[Status]
                            FROM QMS.RepairGroup a
                            WHERE 1=1
                            AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "GroupNo", @" AND a.GroupNo = @GroupNo", GroupNo);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "GroupName", @" AND a.GroupName = @GroupName", GroupName);
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

        #region //GetRepairClass -- 取得維修類別資料 -- Ann 2022-06-21
        public string GetRepairClass(int GroupId, string ClassNo, string ClassName)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.ClassId, a.ClassNo, a.ClassName
                            , (a.ClassNo + '-' + a.ClassName) as RepairClassWithText, a.[Status]
                            FROM QMS.RepairClass a
                            WHERE 1=1";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "GroupId", @" AND a.GroupId = @GroupId", GroupId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ClassNo", @" AND a.ClassNo = @ClassNo", ClassNo);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ClassName", @" AND a.ClassName = @ClassName", ClassName);
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

        #region //GetRepairCause -- 取得維修原因資料 -- Ann 2022-06-21
        public string GetRepairCause(int ClassId, int CauseId, string CauseNo, string CauseName)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.CauseId, a.CauseNo, a.CauseName , a.CauseDesc
                            , (a.CauseNo + '-' + a.CauseName) as RepairCauseWithText, a.[Status]
                            FROM QMS.RepairCause a
                            INNER JOIN QMS.RepairClass b on a.ClassId = b.ClassId
                            INNER JOIN QMS.RepairGroup b1 on b.GroupId = b1.GroupId
                            WHERE 1=1
                            AND b1.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ClassId", @" AND a.ClassId = @ClassId", ClassId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CauseId", @" AND a.CauseId = @CauseId", CauseId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CauseNo", @" AND a.CauseNo = @CauseNo", CauseNo);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CauseName", @" AND a.CauseName = @CauseName", CauseName);
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

        #region //GetQcGroup -- 取得量測群組資料 -- Shintokuro 2022-10-03
        public string GetQcGroup(string QcGroupNo, string QcGroupName)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.QcGroupId, a.QcGroupNo, a.QcGroupName, a.QcGroupDesc ,(a.QcGroupNo + '-' + a.QcGroupName) as QcGroupWithText, a.[Status]
                            FROM QMS.QcGroup a
                            WHERE a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QcGroupNo", @" AND a.QcGroupNo = @QcGroupNo", QcGroupNo);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QcGroupName", @" AND a.QcGroupName = @QcGroupName", QcGroupName);
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

        #region //GetQcClass -- 取得量測類別資料 -- Shintokuro 2022-10-03
        public string GetQcClass(int QcGroupId, string QcClassNo, string QcClassName)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.QcClassId, a.QcClassNo, a.QcClassName, a.QcClassDesc
                            , (a.QcClassNo + '-' + a.QcClassName) as QcClassWithText, a.[Status]
                            FROM QMS.QcClass a
                            WHERE 1=1";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QcGroupId", @" AND a.QcGroupId = @QcGroupId", QcGroupId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QcClassNo", @" AND a.QcClassNo = @QcClassNo", QcClassNo);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "QcClassName", @" AND a.QcClassName = @QcClassName", QcClassName);
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

        #region //GetQcItemAll -- 取得量測項目資料 -- Ted 2022-10-03
        public string GetQcItemAll()
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sql = @"WITH MainQueryJson
                            AS
                            (
                                SELECT (a.QcGroupNo + '：' + a.QcGroupName) [name], 'Group' parent
                                , a.QcGroupNo code, a.QcGroupName [call], b.StatusDesc [status], a.QcGroupDesc reason
                                , '' [owner], '' [inputway], ('QcGroup' + CONVERT(varchar(50), a.QcGroupId)) id, a.QcGroupId levelId
                                FROM QMS.QcGroup a
                                LEFT JOIN BAS.[Status] b ON a.Status = b.StatusNo AND b.StatusSchema = 'Status'
                                WHERE a.CompanyId = @CompanyId
                                UNION
                                SELECT (a.QcClassNo + '：' + a.QcClassName) [name], ('QcGroup' + CONVERT(varchar(50), a.QcGroupId)) parent
                                , a.QcClassNo code, a.QcClassName [call], b.StatusDesc [status], a.QcClassDesc reason
                                , '' [owner], '' [inputway], ('QcClass' + CONVERT(varchar(50), a.QcClassId)) id, a.QcClassId levelId
                                FROM QMS.QcClass a
                                LEFT JOIN BAS.[Status] b ON a.Status = b.StatusNo AND b.StatusSchema = 'Status'
                                UNION
                                SELECT (a.QcItemNo + '：' + a.QcItemName) [name], ('QcClass' + CONVERT(varchar(50), a.QcClassId)) parent
                                , a.QcItemNo code, a.QcItemName [call], b.StatusDesc [status], a.QcItemDesc reason
                                , (c.TypeNo + '-' +c.TypeName) [owner],(d.TypeNo + '-' +d.TypeName) [inputway], ('QcItem' + CONVERT(varchar(50), a.QcItemId)) id, a.QcItemId levelId
                                FROM QMS.QcItem a
                                LEFT JOIN BAS.[Status] b ON a.Status = b.StatusNo AND b.StatusSchema = 'Status'
                                LEFT JOIN BAS.[Type] c on a.QcType = c.TypeNo AND c.TypeSchema ='QcItem.QcType'
                                LEFT JOIN BAS.[Type] d on a.QcItemType = d.TypeNo AND d.TypeSchema = 'QcItem.QcItemType'

                            )
                            SELECT *
                            FROM
                            (
                                SELECT 
                                (
                                    SELECT * FROM MainQueryJson
                                    FOR JSON PATH, ROOT('dataset')
                                ) as queryJson
                            ) a";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    #region //處理JSON結構
                    var result = sqlConnection.Query(sql, dynamicParameters);
                    var jsonString = "";
                    foreach (var item in result)
                    {
                        jsonString = item.queryJson;
                    }

                    var transferJson = JObject.Parse(jsonString);

                    var datasetCount = transferJson["dataset"].Count();
                    for (var i = 0; i < datasetCount; i++)
                    {
                        if (transferJson["dataset"][i]["parent"].ToString() == "Group")
                        {
                            JObject dataFormat = JObject.FromObject(new
                            {
                                name = transferJson["dataset"][i]["name"],
                                code = transferJson["dataset"][i]["code"],
                                call = transferJson["dataset"][i]["call"],
                                status = transferJson["dataset"][i]["status"],
                                reason = transferJson["dataset"][i]["reason"],
                                owner = transferJson["dataset"][i]["owner"],
                                inputway = transferJson["dataset"][i]["inputway"],
                                id = transferJson["dataset"][i]["id"],
                                levelId = transferJson["dataset"][i]["levelId"]
                            });

                            transferJson["dataset"][i] = dataFormat;
                        }
                    }

                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = transferJson
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

        #region //GetQcItem -- 取得量測項目 -- Ted 2022.09.26
        public string GetQcItem(int QcItemId, string QcItemNo, int QcProdId, string QcType, string Status, string Remark
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.QcItemId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",a.QcItemNo, a.QcItemName,a.QcProdId, a.QcType, a.Status, a.Remark
                          ";
                    sqlQuery.mainTables =
                        @"FROM QMS.QcItem a
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcItemId", @" AND a.QcItemId = @QcItemId", QcItemId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcItemNo", @" AND a.QcItemNo LIKE '%' + @QcItemNo + '%'", QcItemNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcProdId", @" AND a.QcProdId = @QcProdId", QcProdId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcType", @" AND a.QcType = @QcType", QcType);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));


                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.QcItemId DESC";
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

        #region //GetQiDetail -- 取得量測項目機型單身 -- Shintokuro 2022.09.27
        public string GetQiDetail(int QcItemId, int QcMachineModeId, string QcMachineModeNo
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.QiDetailId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",a.QcItemId,a.QcMachineModeId
                          ,b.QcMachineModeNo, b.QcMachineModeName, b.QcMachineModeDesc
                          , (b.QcMachineModeNo + '-' + b.QcMachineModeName) QcMachineFullNo
                          ";
                    sqlQuery.mainTables =
                        @"FROM QMS.QiDetail a
                          INNER JOIN QMS.QcMachineMode b on a.QcMachineModeId = b.QcMachineModeId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND 1=1";
                    
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcItemId", @" AND a.QcItemId = @QcItemId", QcItemId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcMachineModeId", @" AND a.QcMachineModeId = @QcMachineModeId", QcMachineModeId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcMachineModeNo", @" AND b.QcMachineModeNo LIKE  '%' + @QcMachineModeNo + '%'", QcMachineModeNo);

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.QiDetailId DESC";
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

        #region //GetQcMachineMode -- 取得量測機型 -- Ted 2022.09.27
        public string GetQcMachineMode(int QcMachineModeId, string QcMachineModeNo, string QcMachineModeName, string QcMachineModeNumber
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.QcMachineModeId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",a.QcMachineModeNo, a.QcMachineModeName, a.QcMachineModeDesc, a.QcMachineModeNumber, a.ItemNo
                          ,(a.QcMachineModeNo + '-' + a.QcMachineModeName) QcMachineModeWithNoName
                          ";
                    sqlQuery.mainTables =
                        @"FROM QMS.QcMachineMode a
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcMachineModeId", @" AND a.QcMachineModeId = @QcMachineModeId", QcMachineModeId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcMachineModeNo", @" AND a.QcMachineModeNo LIKE '%' + @QcMachineModeNo + '%'", QcMachineModeNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcMachineModeName", @" AND a.QcMachineModeName LIKE '%' + @QcMachineModeName + '%'", QcMachineModeName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcMachineModeNumber", @" AND a.QcMachineModeNumber  = @QcMachineModeNumber", QcMachineModeNumber);

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.QcMachineModeId";
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

        #region //GetQmmDetail -- 取得量測機型 - 機台單身 -- Ted 2022.09.27
        public string GetQmmDetail(int QmmDetailId ,int QcMachineModeId, string MachineNo, string MachineName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.QmmDetailId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",a.QcMachineModeId, a.MachineId, a.Status
                          ,b.MachineNo,b.MachineName,b.MachineDesc
                          ,b1.ShopName
                          ";
                    sqlQuery.mainTables =
                        @"FROM QMS.QmmDetail a
                          INNER JOIN MES.Machine b on a.MachineId = b.MachineId
                          INNER JOIN MES.WorkShop b1 on b.ShopId = b1.ShopId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND b1.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcMachineModeId", @" AND a.QcMachineModeId = @QcMachineModeId", QcMachineModeId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MachineNo", @" AND b.MachineNo LIKE '%' + @MachineNo + '%'", MachineNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MachineName", @" AND b.MachineName LIKE '%' + @MachineName + '%'", MachineName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));


                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.QmmDetailId";
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

        #region //GetQcGroupNew -- 取得量測群組資料(新) -- Shintokuro 2024.05.15
        public string GetQcGroupNew(int QcGroupId, string QcGroupNo, string QcGroupName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.QcGroupId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" , a.QcGroupNo, a.QcGroupName, a.QcGroupDesc, a.Status
                          ,(a.QcGroupNo + '-' + a.QcGroupName) as QcGroupWithText
                        ";
                    sqlQuery.mainTables =
                        @"FROM QMS.QcGroup a
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcGroupId", @" AND a.QcGroupId = @QcGroupId", QcGroupId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcGroupNo", @" AND a.QcGroupNo LIKE '%' + @QcGroupNo + '%'", QcGroupNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcGroupName", @" AND a.QcGroupName LIKE '%' + @QcGroupName + '%'", QcGroupName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));


                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.QcGroupNo ASC";
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

        #region //GetQcClassNew -- 取得量測類別資料(新) -- Shintokuro 2024.05.15
        public string GetQcClassNew(int QcClassId,int QcGroupId, string QcClassNo, string QcClassName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.QcClassId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" , a.QcGroupId, a.QcClassNo, a.QcClassName,a.QcClassDesc, a.Status
                          , (a.QcClassNo + '-' + a.QcClassName) as QcClassWithText
                        ";
                    sqlQuery.mainTables =
                        @"FROM QMS.QcClass a
                          INNER JOIN QMS.QcGroup b on a.QcGroupId = b.QcGroupId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND b.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcClassId", @" AND a.QcClassId = @QcClassId", QcClassId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcGroupId", @" AND a.QcGroupId = @QcGroupId", QcGroupId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcClassNo", @" AND a.QcClassNo LIKE '%' + @QcClassNo + '%'", QcClassNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcClassName", @" AND a.QcClassName LIKE '%' + @QcClassName + '%'", QcClassName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));


                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.QcClassNo ASC";
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

        #region //GetQcItemNew -- 取得量測群組資料(新) -- Shintokuro 2024.05.15
        public string GetQcItemNew(int QcItemId, int QcClassId, int QcGroupId, string QcItemNo, string QcItemName, string QicNo, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.QcItemId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",a.QcItemNo, a.QcItemName, a.QcItemDesc, a.QcType, a.QcItemType, a.Status, a.Remark
                          , (a.QcItemNo + '-' + a.QcItemName) as QcItemWithText
                          ,b.QcClassId ,b.QcGroupId
                          ";
                    sqlQuery.mainTables =
                        @"FROM QMS.QcItem a
                          INNER JOIN QMS.QcClass b on a.QcClassId = b.QcClassId
                          INNER JOIN QMS.QcGroup c on b.QcGroupId = c.QcGroupId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND c.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    if(QicNo.Length > 0)
                    {
                        QicNo = "___" + QicNo + "__";
                    }
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcItemId", @" AND a.QcItemId = @QcItemId", QcItemId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcClassId", @" AND a.QcClassId = @QcClassId", QcClassId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcGroupId", @" AND b.QcGroupId = @QcGroupId", QcGroupId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcItemNo", @" AND a.QcItemNo LIKE '%' + @QcItemNo + '%'", QcItemNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcItemName", @" AND a.QcItemName LIKE '%' + @QcItemName + '%'", QcItemName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    if (QicNo.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QicNo", @" AND a.QcItemNo LIKE @QicNo", QicNo);


                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.QcItemNo ASC";
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

        #region //GetQcClassNoMax -- 取得量測類別最大號 -- Shintokuro 2024.05.16
        public string GetQcClassNoMax(int QcGroupId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    #region //判斷工具目前最大編號是否存在
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 a.QcClassNo
                            FROM QMS.QcClass a
                            INNER JOIN QMS.QcGroup b on a.QcGroupId = b.QcGroupId
                            WHERE 1=1
                            AND a.QcGroupId= @QcGroupId
                            AND b.CompanyId= @CompanyId
                            ORDER BY a.QcClassNo desc";
                    dynamicParameters.Add("QcGroupId", QcGroupId);
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    var result = sqlConnection.Query(sql, dynamicParameters);
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

        #region //GetQcItemPrinciple -- 取得量測項目編碼原則 -- Shintokuro 2024.05.15
        public string GetQcItemPrinciple(int PrincipleId, int QcClassId, int QmmDetailId, string PrincipleNo, string PrincipleDesc
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.PrincipleId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" , a.QcClassId, a.QmmDetailId, a.PrincipleNo, a.PrincipleDesc
                          ,b.QcClassName
                          ,c1.QcMachineModeNo ,c1.QcMachineModeName ,c1.QcMachineModeEnName ,c1.QcMachineModeDesc
                          ,c2.MachineNo ,c2.MachineName ,c2.MachineDesc ,c2.CncMachineNo
                        ";
                    sqlQuery.mainTables =
                        @"FROM QMS.QcItemPrinciple a
                          INNER JOIN QMS.QcClass b on a.QcClassId = b.QcClassId
                          INNER JOIN QMS.QcGroup b1 on b.QcGroupId = b1.QcGroupId
                          INNER JOIN QMS.QmmDetail c on a.QmmDetailId = c.QmmDetailId
                          INNER JOIN QMS.QcMachineMode c1 on c.QcMachineModeId = c1.QcMachineModeId
                          INNER JOIN MES.Machine c2 on c.MachineId = c2.MachineId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND b1.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PrincipleId", @" AND a.PrincipleId = @PrincipleId", PrincipleId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QcClassId", @" AND a.QcClassId = @QcClassId", QcClassId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QmmDetailId", @" AND a.QmmDetailId = @QmmDetailId", QmmDetailId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PrincipleNo", @" AND a.PrincipleNo LIKE '%' + @PrincipleNo + '%'", PrincipleNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PrincipleDesc", @" AND a.PrincipleDesc LIKE '%' + @PrincipleDesc + '%'", PrincipleDesc);


                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.PrincipleNo ASC";
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

        #region //GetPrincipleDetail -- 取得量測項目編碼原則 附加欄位 -- Shintokuro 2024.05.15
        public string GetPrincipleDetail(int PdId, int PrincipleId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.PdId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" , a.PrincipleId, a.PrincipleDesc
                        ";
                    sqlQuery.mainTables =
                        @"FROM QMS.PrincipleDetail a
                          INNER JOIN QMS.QcItemPrinciple b on a.PrincipleId = b.PrincipleId
                          INNER JOIN QMS.QcClassId b1 on b.QcClassId = b1.QcClassId
                          INNER JOIN QMS.QcGroup b2 on b1.QcGroupId = b2.QcGroupId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND b2.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PdId", @" AND a.PdId = @PdId", PdId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PrincipleId", @" AND a.PrincipleId = @PrincipleId", PrincipleId);


                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.PdId DESC";
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

        #region //GetQcItemCoding -- 取得項目編碼規則管理 -- Shintokuro 2024.05.17
        public string GetQcItemCoding(int QicId, string QicNo, string QicName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.QicId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" , a.QicNo, a.QicName, a.QicDesc, a.Status
                            ,(a.QicNo + '-' + a.QicName) QicWithText
                        ";
                    sqlQuery.mainTables =
                        @"FROM QMS.QcItemCoding a
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QicId", @" AND a.QicId = @QicId", QicId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QicNo", @" AND a.QicNo = @QicNo", QicNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QicName", @" AND a.QicName = @QicName", QicName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.QicNo ASC";
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

        #region //GetLotNumberByQc --取得批號資料 -- Chia Yuan 2024-11-20
        public string GetLotNumberByQc(int MoId, int MoProcessId, string ItemValue)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //取得前一站鍍膜的排序碼
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.SortNumber, CONVERT(VARCHAR(20), b.StartDate, 112) StartDate
                            FROM MES.MoProcess a 
                            INNER JOIN MES.BarcodeProcess b ON b.MoProcessId = a.MoProcessId
                            WHERE a.MoId = @MoId
                            AND EXISTS (SELECT TOP 1 1 FROM MES.MoProcess aa 
                            WHERE aa.MoId = a.MoId
                            AND aa.MoProcessId = @MoProcessId
                            AND (aa.SortNumber-1) = a.SortNumber)";
                    dynamicParameters.Add("MoId", MoId);
                    dynamicParameters.Add("MoProcessId", MoProcessId);
                    var result = sqlConnection.Query(sql, dynamicParameters);
                    var SortNumber = result.FirstOrDefault()?.SortNumber ?? -1;
                    #endregion

                    #region //取得批號選單
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DISTINCT lot.LotNumberNo
                            FROM MES.MoProcess a
                            INNER JOIN MES.BarcodeProcess b ON b.MoProcessId = a.MoProcessId
                            INNER JOIN MES.Barcode c ON c.BarcodeId = b.BarcodeId
                            INNER JOIN MES.BarcodeAttribute d ON d.BarcodeId = b.BarcodeId
                            OUTER APPLY (
	                            SELECT ac.LotNumberNo
	                            FROM MES.MoRouting aa 
	                            INNER JOIN MES.RoutingItem ab ON ab.RoutingItemId = aa.RoutingItemId 
	                            INNER JOIN SCM.LotNumber ac ON ac.MtlItemId = ab.MtlItemId 
	                            WHERE aa.MoRoutingId = a.MoRoutingId
                                AND ac.LotNumberNo IN @LotNumberNos
                            ) lot
                            WHERE a.MoId = @MoId --制令
                            AND a.SortNumber = @SortNumber --前一站序號
                            AND d.ItemValue = @ItemValue --鍋次
                            AND c.CurrentProdStatus <> 'S'
                            AND NOT(lot.LotNumberNo IS NULL)
                            ORDER BY lot.LotNumberNo DESC";
                    dynamicParameters.Add("MoId", MoId);
                    dynamicParameters.Add("ItemValue", ItemValue);
                    dynamicParameters.Add("SortNumber", SortNumber);
                    dynamicParameters.Add("LotNumberNos", result.Select(s => s.StartDate).ToArray());
                    result = sqlConnection.Query(sql, dynamicParameters);
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

        #region //GetBarcodeAttributeByQc -- 取得條碼屬性資料 -- Chia Yuan 2024-10-23
        public string GetBarcodeAttributeByQc(int MoId, int MoProcessId, string ItemNo)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.ItemValue
                            , a.ItemValue + ' 條碼數量:' + CONVERT(NVARCHAR(5), COUNT(a.ItemValue)) CountWithItem
                            FROM MES.BarcodeAttribute a
                            WHERE a.MoId = @MoId
                            AND a.ItemNo = @ItemNo --'PotNumber'
                            AND EXISTS (
	                            SELECT TOP 1 1
	                            FROM MES.BarcodeProcess aa
	                            INNER JOIN MES.MoProcess ab ON ab.MoProcessId = aa.MoProcessId AND ab.MoId = aa.MoId
	                            INNER JOIN MES.ManufactureOrder ac ON ac.MoId = ab.MoId
	                            INNER JOIN QMS.QcType ad ON ad.ModeId = ac.ModeId --AND ad.QcTypeNo LIKE N'%CPQC%'
	                            WHERE aa.MoId = a.MoId 
                                AND aa.BarcodeId = a.BarcodeId
                                AND ab.MoProcessId = @MoProcessId
                            )
                            GROUP BY a.ItemValue";
                    dynamicParameters.Add("MoId", MoId);
                    dynamicParameters.Add("ItemNo", ItemNo);
                    dynamicParameters.Add("MoProcessId", MoProcessId);

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

        #region //GetQcItemByQc -- 取得量測項目資料 -- Chia Yuan 2024-11-28
        public string GetQcItemByQc()
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.QcGroupNo, a.QcGroupName, b.QcClassNo, b.QcClassName
                            , c.QcItemId, c.QcItemNo, c.QcItemName, c.QcType + ' ' + c.QcItemName QcItemWithType
                            FROM QMS.QcGroup a
                            INNER JOIN QMS.QcClass b ON b.QcGroupId = a.QcGroupId AND b.QcClassNo = 'B01' AND b.[Status] = 'A'
                            INNER JOIN QMS.QcItem c ON c.QcClassId = b.QcClassId AND c.QcType = 'IPQC' AND c.[Status] = 'A'
                            WHERE a.CompanyId = @CompanyId 
                            AND a.QcGroupNo = 'B' 
                            AND a.[Status] = 'A'
                            AND (c.QcItemNo LIKE '%s001' OR c.QcItemNo LIKE '%s101')
                            ORDER BY a.QcGroupNo, b.QcClassNo, c.QcItemNo";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

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

        #region //GetQcMeasureInTheoryWorkTime -- 取得量測項目預測工時列表 -- GPAI 2024-12-17
        public string GetQcMeasureInTheoryWorkTime(int QmwtId, string ProductType, int QmmDetailId, string QicNo/*, decimal WorkTime*/
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.QmwtId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ProductType, a.MeasureSize, CONVERT(DECIMAL(10,2),a.WorkTime/60.0) WorkTime, a.QmmDetailId
                          , b.MachineNumber
                          , c.MachineNo, c.MachineDesc, c.MachineName
                          , d.QicNo, d.QicName, d.QicId";
                    sqlQuery.mainTables =
                        @"FROM QMS.QcMeasureInTheoryWorkTime a
                          INNER JOIN QMS.QmmDetail b ON a.QmmDetailId = b.QmmDetailId
                          INNER JOIN MES.Machine c ON b.MachineId = c.MachineId
                          INNER JOIN QMS.QcItemCoding d ON a.QicId = d.QicId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND 1 = 1 AND d.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QmwtId", @" AND a.QmwtId = @QmwtId", QmwtId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProductType", @" AND a.ProductType = @ProductType", ProductType);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QmmDetailId", @" AND b.QmmDetailId = @QmmDetailId", QmmDetailId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "QicNo", @" AND d.QicNo = @QicNo", QicNo);

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.QmwtId DESC";
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
        #region //AddNgGroup-- 新增異常群組 -- Ann 2022-06-14
        public string AddNgGroup(string GroupNo, string GroupName, string Status, string GroupDesc)
        {
            try
            {
                if (GroupNo.Length <= 0) throw new SystemException("【群組代碼】不能為空!");
                if (GroupNo.Length > 100) throw new SystemException("【群組代碼】長度錯誤!");
                if (GroupName.Length <= 0) throw new SystemException("【群組名稱】不能為空!");
                if (GroupName.Length > 100) throw new SystemException("【群組名稱】長度錯誤!");
                if (Status.Length <= 0) throw new SystemException("【群組狀態】不能為空!");
                if (Status.Length > 2) throw new SystemException("【群組狀態】長度錯誤!");
                if (GroupDesc.Length <= 0) throw new SystemException("【群組描述】不能為空!");
                if (GroupDesc.Length > 100) throw new SystemException("【群組描述】長度錯誤!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷群組代碼是否已經存在
                        sql = @"SELECT TOP 1 1
                                FROM QMS.DefectGroup a
                                WHERE a.GroupNo = @GroupNo";
                        dynamicParameters.Add("GroupNo", GroupNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【群組代碼】已存在，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO QMS.DefectGroup (CompanyId, GroupNo, GroupName
                                , GroupDesc, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.GroupId
                                VALUES (@CompanyId, @GroupNo, @GroupName, @GroupDesc, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                GroupNo,
                                GroupName,
                                GroupDesc,
                                Status = "A",
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

        #region //AddNgClass-- 新增異常類別 -- Ann 2022-06-15
        public string AddNgClass(int GroupId, string ClassNo, string ClassName, string ClassDesc, string Status)
        {
            try
            {
                if (GroupId <= 0) throw new SystemException("【異常群組】不能為空!");
                if (ClassNo.Length <= 0) throw new SystemException("【類別代碼】不能為空!");
                if (ClassNo.Length > 100) throw new SystemException("【類別代碼】長度錯誤!");
                if (ClassName.Length <= 0) throw new SystemException("【類別名稱】不能為空!");
                if (ClassName.Length > 100) throw new SystemException("【類別名稱】長度錯誤!");
                if (ClassDesc.Length <= 0) throw new SystemException("【類別描述】不能為空!");
                if (ClassDesc.Length > 100) throw new SystemException("【類別描述】長度錯誤!");
                if (Status.Length <= 0) throw new SystemException("【類別狀態】不能為空!");
                if (Status.Length > 2) throw new SystemException("【類別狀態】長度錯誤!");
                

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷同一個群組的類別代碼是否已經存在
                        sql = @"SELECT TOP 1 1
                                FROM QMS.DefectClass a
                                WHERE a.GroupId = @GroupId
                                AND a.ClassNo = @ClassNo";
                        dynamicParameters.Add("GroupId", GroupId);
                        dynamicParameters.Add("ClassNo", ClassNo);                        

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【類別代碼】已存在，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO QMS.DefectClass (GroupId, ClassNo, ClassName
                                , ClassDesc, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ClassId
                                VALUES (@GroupId, @ClassNo, @ClassName, @ClassDesc, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                GroupId,
                                ClassNo,
                                ClassName,
                                ClassDesc,
                                Status = "A",
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

        #region //AddNgCause-- 新增異常原因 -- Ann 2022-06-15
        public string AddNgCause(int ClassId, string CauseNo, string CauseName, string CauseDesc, int ResponsibleDepartment, string Status)
        {
            try
            {
                if (ClassId <= 0) throw new SystemException("【異常類別】不能為空!");
                if (CauseNo.Length <= 0) throw new SystemException("【原因代碼】不能為空!");
                if (CauseNo.Length > 100) throw new SystemException("【原因代碼】長度錯誤!");
                if (CauseName.Length <= 0) throw new SystemException("【原因名稱】不能為空!");
                if (CauseName.Length > 100) throw new SystemException("【原因名稱】長度錯誤!");
                if (CauseDesc.Length <= 0) throw new SystemException("【原因描述】不能為空!");
                if (CauseDesc.Length > 100) throw new SystemException("【原因描述】長度錯誤!");
                if (Status.Length <= 0) throw new SystemException("【原因狀態】不能為空!");
                if (Status.Length > 2) throw new SystemException("【原因狀態】長度錯誤!");


                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷同一個類別的原因代碼是否已經存在
                        sql = @"SELECT TOP 1 1
                                FROM QMS.DefectCause a
                                WHERE a.ClassId = @ClassId
                                AND a.CauseNo = @CauseNo";
                        dynamicParameters.Add("ClassId", ClassId);
                        dynamicParameters.Add("CauseNo", CauseNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【原因代碼】已存在，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO QMS.DefectCause (ClassId, CauseNo, CauseName
                                , CauseDesc, ResponsibleDepartment, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.CauseId
                                VALUES (@ClassId, @CauseNo, @CauseName, @CauseDesc, @ResponsibleDepartment, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ClassId,
                                CauseNo,
                                CauseName,
                                CauseDesc,
                                ResponsibleDepartment,
                                Status = "A",
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

        #region //AddRepairGroup-- 新增維修群組資料 -- Ann 2022-06-21
        public string AddRepairGroup(string GroupNo, string GroupName, string GroupDesc, string Status)
        {
            try
            {
                if (GroupNo.Length <= 0) throw new SystemException("【群組代碼】不能為空!");
                if (GroupNo.Length > 100) throw new SystemException("【群組代碼】長度錯誤!");
                if (GroupName.Length <= 0) throw new SystemException("【群組名稱】不能為空!");
                if (GroupName.Length > 100) throw new SystemException("【群組名稱】長度錯誤!");
                if (GroupDesc.Length <= 0) throw new SystemException("【群組描述】不能為空!");
                if (GroupDesc.Length > 100) throw new SystemException("【群組描述】長度錯誤!");
                if (Status.Length <= 0) throw new SystemException("【群組狀態】不能為空!");
                if (Status.Length > 2) throw new SystemException("【群組狀態】長度錯誤!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷群組代碼是否已經存在
                        sql = @"SELECT TOP 1 1
                                FROM QMS.RepairGroup a
                                WHERE a.GroupNo = @GroupNo";
                        dynamicParameters.Add("GroupNo", GroupNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【群組代碼】已存在，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO QMS.RepairGroup (CompanyId, GroupNo, GroupName
                                , GroupDesc, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.GroupId
                                VALUES (@CompanyId, @GroupNo, @GroupName, @GroupDesc, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                GroupNo,
                                GroupName,
                                GroupDesc,
                                Status = "A",
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

        #region //AddRepairClass-- 新增維修類別資料 -- Ann 2022-06-21
        public string AddRepairClass(int GroupId, string ClassNo, string ClassName, string ClassDesc, string Status)
        {
            try
            {
                if (GroupId <= 0) throw new SystemException("【群組編號】不能為空!");
                if (ClassNo.Length <= 0) throw new SystemException("【類別代碼】不能為空!");
                if (ClassNo.Length > 100) throw new SystemException("【類別代碼】長度錯誤!");
                if (ClassName.Length <= 0) throw new SystemException("【類別名稱】不能為空!");
                if (ClassName.Length > 100) throw new SystemException("【類別名稱】長度錯誤!");
                if (ClassDesc.Length <= 0) throw new SystemException("【類別描述】不能為空!");
                if (ClassDesc.Length > 100) throw new SystemException("【類別描述】長度錯誤!");
                if (Status.Length <= 0) throw new SystemException("【類別狀態】不能為空!");
                if (Status.Length > 2) throw new SystemException("【類別狀態】長度錯誤!");


                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷同一個群組的類別代碼是否已經存在
                        sql = @"SELECT TOP 1 1
                                FROM QMS.RepairClass a
                                WHERE a.GroupId = @GroupId
                                AND a.ClassNo = @ClassNo";
                        dynamicParameters.Add("GroupId", GroupId);
                        dynamicParameters.Add("ClassNo", ClassNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【類別代碼】已存在，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO QMS.RepairClass (GroupId, ClassNo, ClassName
                                , ClassDesc, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ClassId
                                VALUES (@GroupId, @ClassNo, @ClassName, @ClassDesc, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                GroupId,
                                ClassNo,
                                ClassName,
                                ClassDesc,
                                Status = "A",
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

        #region //AddRepairCause-- 新增維修原因資料 -- Ann 2022-06-21
        public string AddRepairCause(int ClassId, string CauseNo, string CauseName, string CauseDesc, int ResponsibleDepartment, string Status)
        {
            try
            {
                if (ClassId <= 0) throw new SystemException("【異常類別】不能為空!");
                if (CauseNo.Length <= 0) throw new SystemException("【原因代碼】不能為空!");
                if (CauseNo.Length > 100) throw new SystemException("【原因代碼】長度錯誤!");
                if (CauseName.Length <= 0) throw new SystemException("【原因名稱】不能為空!");
                if (CauseName.Length > 100) throw new SystemException("【原因名稱】長度錯誤!");
                if (CauseDesc.Length <= 0) throw new SystemException("【原因描述】不能為空!");
                if (CauseDesc.Length > 100) throw new SystemException("【原因描述】長度錯誤!");
                if (ResponsibleDepartment <= 0) throw new SystemException("【權責單位】不能為空!");
                if (Status.Length <= 0) throw new SystemException("【原因狀態】不能為空!");
                if (Status.Length > 2) throw new SystemException("【原因狀態】長度錯誤!");


                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷同一個類別的原因代碼是否已經存在
                        sql = @"SELECT TOP 1 1
                                FROM QMS.RepairCause a
                                WHERE a.ClassId = @ClassId
                                AND a.CauseNo = @CauseNo";
                        dynamicParameters.Add("ClassId", ClassId);
                        dynamicParameters.Add("CauseNo", CauseNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【原因代碼】已存在，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO QMS.RepairCause (ClassId, CauseNo, CauseName
                                , CauseDesc, ResponsibleDepartment, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.CauseId
                                VALUES (@ClassId, @CauseNo, @CauseName, @CauseDesc, @ResponsibleDepartment, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ClassId,
                                CauseNo,
                                CauseName,
                                CauseDesc,
                                ResponsibleDepartment,
                                Status = "A",
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

        #region //AddQcGroup-- 新增量測群組資料 -- Ted 2022-10-03
        public string AddQcGroup(string QcGroupNo, string QcGroupName, string QcGroupDesc, string Status)
        {
            try
            {
                string pattern = @"^[A-Z]$";
                Regex regex = new Regex(pattern);

                if (!regex.IsMatch(QcGroupNo)) throw new SystemException("【群組代碼】格式為1位數的大寫英文!");
                if (QcGroupName.Length <= 0) throw new SystemException("【群組名稱】不能為空!");
                if (QcGroupName.Length > 100) throw new SystemException("【群組名稱】長度錯誤!");
                if (QcGroupDesc.Length <= 0) throw new SystemException("【群組描述】不能為空!");
                if (QcGroupDesc.Length > 100) throw new SystemException("【群組描述】長度錯誤!");
                if (Status.Length <= 0) throw new SystemException("【群組狀態】不能為空!");
                if (Status.Length > 2) throw new SystemException("【群組狀態】長度錯誤!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷群組代碼是否已經存在
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcGroup a
                                WHERE a.QcGroupNo = @QcGroupNo";
                        dynamicParameters.Add("QcGroupNo", QcGroupNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【群組代碼】已存在，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO QMS.QcGroup (CompanyId, QcGroupNo, QcGroupName
                                , QcGroupDesc, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.QcGroupId
                                VALUES (@CompanyId, @QcGroupNo, @QcGroupName, @QcGroupDesc, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                QcGroupNo,
                                QcGroupName,
                                QcGroupDesc,
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

        #region //AddQcClass-- 新增量測類別資料 -- Ted 2022-10-03
        public string AddQcClass(int QcGroupId, string QcClassNo, string QcClassName, string QcClassDesc, string Status)
        {
            try
            {
                if (QcGroupId <= 0) throw new SystemException("【群組編號】不能為空!");
                if (QcClassNo.Length <= 0) throw new SystemException("【類別代碼】不能為空!");
                if (QcClassNo.Length > 100) throw new SystemException("【類別代碼】長度錯誤!");
                if (QcClassName.Length <= 0) throw new SystemException("【類別名稱】不能為空!");
                if (QcClassName.Length > 100) throw new SystemException("【類別名稱】長度錯誤!");
                if (QcClassDesc.Length <= 0) throw new SystemException("【類別描述】不能為空!");
                if (QcClassDesc.Length > 100) throw new SystemException("【類別描述】長度錯誤!");
                if (Status.Length <= 0) throw new SystemException("【類別狀態】不能為空!");
                if (Status.Length > 2) throw new SystemException("【類別狀態】長度錯誤!");


                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷同一個群組的類別代碼是否已經存在
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcClass a
                                WHERE a.QcGroupId = @QcGroupId
                                AND a.QcClassNo = @QcClassNo";
                        dynamicParameters.Add("QcGroupId", QcGroupId);
                        dynamicParameters.Add("QcClassNo", QcClassNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【類別代碼】已存在，請重新輸入!");
                        #endregion


                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO QMS.QcClass (QcGroupId, QcClassNo, QcClassName
                                , QcClassDesc, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.QcClassId
                                VALUES (@QcGroupId, @QcClassNo, @QcClassName, @QcClassDesc, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                QcGroupId,
                                QcClassNo = QcClassNo,
                                QcClassName,
                                QcClassDesc,
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

        #region //AddQcItem-- 新增量測項目 -- Ted 2022-10-03
        public string AddQcItem(int QcClassId, string QcItemNo, string QcItemName, string QcItemDesc, string QcType, string QcItemType, int QcProdId, string Remark, string Status)
        {
            try
            {
                if (QcClassId <= 0) throw new SystemException("【異常類別】不能為空!");
                if (QcItemNo.Length <= 0) throw new SystemException("【原因代碼】不能為空!");
                if (QcItemNo.Length > 100) throw new SystemException("【原因代碼】長度錯誤!");
                if (QcItemName.Length <= 0) throw new SystemException("【原因名稱】不能為空!");
                if (QcItemName.Length > 100) throw new SystemException("【原因名稱】長度錯誤!");
                if (QcItemDesc.Length <= 0) throw new SystemException("【原因描述】不能為空!");
                if (QcItemDesc.Length > 100) throw new SystemException("【原因描述】長度錯誤!");
                if (QcType.Length <= 0) throw new SystemException("【檢驗類別】不能為空!");
                if (QcItemType.Length <= 0) throw new SystemException("【量測項目輸入方式】不能為空!");
                if (Status.Length <= 0) throw new SystemException("【原因狀態】不能為空!");
                if (Status.Length > 2) throw new SystemException("【原因狀態】長度錯誤!");


                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷同一個類別的原因代碼是否已經存在
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcItem a
                                INNER JOIN QMS.QcClass b on a.QcClassId = b.QcClassId
                                INNER JOIN QMS.QcGroup c on b.QcGroupId = c.QcGroupId
                                WHERE a.QcClassId = @QcClassId
                                AND a.QcItemNo = @QcItemNo";
                        dynamicParameters.Add("QcClassId", QcClassId);
                        dynamicParameters.Add("QcItemNo", QcItemNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【原因代碼】已存在，請重新輸入!");
                        #endregion

                        #region //判斷同一個類別的原因代碼是否已經存在
                        sql = @"SELECT a.QcClassNo temporaryQcItemNo
                                FROM QMS.QcClass a
                                WHERE a.QcClassId = @QcClassId";
                        dynamicParameters.Add("QcClassId", QcClassId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        string temporaryQcItemNo = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).temporaryQcItemNo;

                        #endregion



                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO QMS.QcItem (QcClassId, QcItemNo, QcItemName
                                , QcItemDesc, QcType ,QcItemType , Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.QcItemId,INSERTED.QcItemName
                                VALUES (@QcClassId, @QcItemNo, @QcItemName, @QcItemDesc, @QcType, @QcItemType, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                QcClassId,
                                QcItemNo = temporaryQcItemNo + QcItemNo,
                                QcItemName,
                                QcItemDesc,
                                QcType,
                                QcItemType,
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

        #region //AddQcItemBatch-- 新增量測項目(批量) -- Shintokuro 2024-05-17
        public string AddQcItemBatch(string BatchFormData)
        {
            try
            {
                int rowsAffected = 0;
                if (BatchFormData.Length <= 0) throw new SystemException("【批量資料】不能為空!");

                JObject jsonObject = JObject.Parse(BatchFormData);

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        foreach (var item in jsonObject)
                        {
                            string key1 = item.Value.ToString();

                            int QcClassId = Convert.ToInt32(key1.Split(',')[0]);
                            string QcClassNo = "";
                            #region //判斷類別是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 QcClassNo
                                        FROM QMS.QcClass a
                                        WHERE a.QcClassId = @QcClassId";
                            dynamicParameters.Add("QcClassId", QcClassId);

                            var resultQcClass = sqlConnection.Query(sql, dynamicParameters);
                            foreach(var item1 in resultQcClass)
                            {
                                QcClassNo = item1.QcClassNo;
                            }
                            #endregion

                            int key2 = key1.Split(',').Count();
                            for (int i = 1; i < key2; i++)
                            {
                                int QicId = Convert.ToInt32(key1.Split(',')[i].Split(':')[0]);
                                int num = Convert.ToInt32(key1.Split(',')[i].Split(':')[1]);
                                int ItemType = Convert.ToInt32(key1.Split(',')[i].Split(':')[2]);
                                string QicNo = "";
                                string QicName = "";

                                #region //判斷類別是否存在 取得編碼前綴QicNo
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 QicNo,QicName
                                        FROM QMS.QcItemCoding a
                                        WHERE a.QicId = @QicId
                                        AND a.Status = 'A'";
                                dynamicParameters.Add("QicId", QicId);

                                var resultQcprId = sqlConnection.Query(sql, dynamicParameters);
                                foreach(var item1 in resultQcprId)
                                {
                                    QicNo = item1.QicNo;
                                    QicName = item1.QicName;
                                }
                                #endregion

                                #region //判斷QcItem 採用QcClassId 和 QicNo 的最大序號
                                int MaxNo = 0;
                                string keySearch = QcClassNo + QicNo;
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 QcItemNo,QcItemName
                                        FROM QMS.QcItem a
                                        WHERE a.QcClassId = @QcClassId
                                        AND a.QcItemNo LIKE '%"+ keySearch + @"%'
                                        AND a.Status = 'A'
                                        ORDER BY QcItemNo DESC ";
                                dynamicParameters.Add("QcClassId", QcClassId);
                                dynamicParameters.Add("keySearch", keySearch);

                                var resultQcItem = sqlConnection.Query(sql, dynamicParameters);
                                if (resultQcItem.Count() > 0)
                                {
                                    foreach (var item1 in resultQcItem)
                                    {
                                        MaxNo = Convert.ToInt32(item1.QcItemNo.Replace(keySearch, ""));
                                    }
                                }
                                #endregion

                                #region //數量轉換成流水號碼 新增QcItem
                                for (int n = MaxNo+1; n<= num+ MaxNo; n++)
                                {
                                    if(n>=100) throw new SystemException("【項目】最大數量只能99個!請重新確認,目前最大號" + MaxNo);
                                    string Seq = n.ToString().PadLeft(2, '0');
                                    string QcItemNo = keySearch + Seq;
                                    #region //判斷同一個類別的原因代碼是否已經存在
                                    sql = @"SELECT TOP 1 1
                                            FROM QMS.QcItem a
                                            INNER JOIN QMS.QcClass b on a.QcClassId = b.QcClassId
                                            INNER JOIN QMS.QcGroup c on b.QcGroupId = c.QcGroupId
                                            WHERE a.QcClassId = @QcClassId
                                            AND a.QcItemNo = @QcItemNo";
                                    dynamicParameters.Add("QcClassId", QcClassId);
                                    dynamicParameters.Add("QcItemNo", QcItemNo);

                                    var result = sqlConnection.Query(sql, dynamicParameters);
                                    if (result.Count() > 0) throw new SystemException("【原因代碼】已存在，請重新輸入!");
                                    #endregion

                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO QMS.QcItem (QcClassId, QcItemNo, QcItemName
                                            , QcItemDesc, QcItemType, Status
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.QcItemId,INSERTED.QcItemName
                                            VALUES (@QcClassId, @QcItemNo, @QcItemName
                                            , @QcItemDesc, @QcItemType, @Status
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            QcClassId,
                                            QcItemNo,
                                            QcItemName = QicName  + Seq,
                                            QcItemDesc = QicName + Seq,
                                            QcItemType = ItemType,
                                            Status = "A",
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });
                                    var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult.Count();
                                }
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

        #region //AddQiDetail -- 新增量測項目機型單身 -- Ted 2022.09.27
        public string AddQiDetail(int QcItemId, int QcMachineModeId)
        {
            try
            {
                if (QcItemId <= 0) throw new SystemException("【量測項目】不能為空!");
                if (QcMachineModeId <= 0) throw new SystemException("【機型】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷量測項目是否正確
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcItem
                                WHERE QcItemId = @QcItemId";
                        dynamicParameters.Add("QcItemId", QcItemId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() < 0) throw new SystemException("【量測項目】錯誤，請重新輸入!");
                        #endregion

                        #region //判斷量測機型是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcMachineMode
                                WHERE QcMachineModeId = @QcMachineModeId";
                        dynamicParameters.Add("QcMachineModeId", QcMachineModeId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() < 0) throw new SystemException("【量測機型】錯誤，請重新輸入!");
                        #endregion

                        #region //判斷機型是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QiDetail
                                WHERE 1=1
                                AND QcItemId = @QcItemId
                                AND QcMachineModeId = @QcMachineModeId";
                        dynamicParameters.Add("QcItemId", QcItemId);
                        dynamicParameters.Add("QcMachineModeId", QcMachineModeId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【機型】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO QMS.QiDetail (QcItemId, QcMachineModeId
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.QiDetailId
                                VALUES (@QcItemId, @QcMachineModeId
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                QcItemId,
                                QcMachineModeId,
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

        #region //AddQcMachineMode -- 新增量測機型 -- Ted 2022.09.27
        public string AddQcMachineMode(string QcMachineModeNo, string QcMachineModeName, string QcMachineModeDesc, string ItemNo)
        {
            try
            {
                string pattern = @"^[a-zA-Z0-9]+$";
                Regex regex = new Regex(pattern);

                if (QcMachineModeNo.Length <= 0) throw new SystemException("【機型編號】不能為空!");
                if (QcMachineModeName.Length <= 0) throw new SystemException("【機型名稱】不能為空!");
                if (ItemNo.Length != 3) throw new SystemException("【項目編號】格式錯誤,3字元且只能英數字!");
                if (!regex.IsMatch(ItemNo)) throw new SystemException("【項目編號】格式錯誤,3字元且只能英數字!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷機型編號是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcMachineMode
                                WHERE CompanyId = @CompanyId
                                AND QcMachineModeNo = @QcMachineModeNo";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("QcMachineModeNo", QcMachineModeNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【機型編號】重複，請重新輸入!");
                        #endregion

                        #region //判斷機型名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcMachineMode
                                WHERE CompanyId = @CompanyId
                                AND QcMachineModeName = @QcMachineModeName";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("QcMachineModeName", QcMachineModeName);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【機型名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO QMS.QcMachineMode (CompanyId, QcMachineModeNo, QcMachineModeName, QcMachineModeDesc
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.QcMachineModeId
                                VALUES (@CompanyId, @QcMachineModeNo, @QcMachineModeName, @QcMachineModeDesc
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                QcMachineModeNo,
                                QcMachineModeName,
                                QcMachineModeDesc,
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

        #region //AddQiMachineMode -- 取得量測機型 - 機台單身 -- Ted 2022.09.27
        public string AddQiMachineMode(int QcMachineModeId, string MachineId)
        {
            try
            {
                if (QcMachineModeId <= 0) throw new SystemException("【機型】資料錯誤!");
                if (MachineId.Length <= 0) throw new SystemException("【機台】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷機型編號是否重複
                        foreach (var item in MachineId.Split(','))
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT b.MachineNo
                                FROM QMS.QmmDetail a
                                INNER JOIN MES.Machine b on a.MachineId = b.MachineId
                                WHERE a.QcMachineModeId = @QcMachineModeId
                                AND a.MachineId = @MachineId";
                            dynamicParameters.Add("QcMachineModeId", QcMachineModeId);
                            dynamicParameters.Add("MachineId", item);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() > 0)
                            {
                                var MachineNo = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).MachineNo;
                                throw new SystemException("【機台編號:"+ MachineNo + "】重複，請重新輸入!");
                            } 
                        }
                        #endregion

                        int rowsAffected = 0;

                        foreach (var item in MachineId.Split(','))
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO QMS.QmmDetail (QcMachineModeId, MachineId, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                VALUES (@QcMachineModeId, @MachineId, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    QcMachineModeId,
                                    MachineId = item,
                                    Status = "A",
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
                            msg = "(" + rowsAffected + " rows affected)",
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

        #region //AddQcItemPrinciple-- 新增量測項目編碼原則 -- Shintokuro 2024.05.15
        public string AddQcItemPrinciple(int QcClassId, int QmmDetailId, string PrincipleNo, string PrincipleDesc)
        {
            try
            {
                //string pattern = @"^[A-Z]$";
                //Regex regex = new Regex(pattern);

                //if (!regex.IsMatch(QcGroupNo)) throw new SystemException("【群組代碼】格式為1位數的大寫英文!");
                if (QcClassId <= 0) throw new SystemException("【產品類別】不能為空!");
                if (QmmDetailId <= 0) throw new SystemException("【量測機台】不能為空!");
                if (PrincipleNo.Length > 4) throw new SystemException("【編碼】最大4字元!");
                if (PrincipleNo.Length <= 0) throw new SystemException("【編碼】不能為空!");
                if (PrincipleDesc.Length > 150) throw new SystemException("【編碼名稱描述】最大150字元!");
                int PrincipleId = -1;
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷QcClass是否存在
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcClass a
                                WHERE 1 =1
                                AND a.QcClassId = @QcClassId
                                ";
                        dynamicParameters.Add("QcClassId", QcClassId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【產品類別】找不到，請重新輸入!");
                        #endregion

                        #region //判斷QmmDetailId是否存在
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QmmDetail a
                                WHERE 1 =1
                                AND a.QmmDetailId = @QmmDetailId
                                ";
                        dynamicParameters.Add("QmmDetailId", QmmDetailId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【量測機台】找不到，請重新輸入!");
                        #endregion

                        #region //判斷組合是否存在
                        sql = @"SELECT TOP 1 QcItemName
                                FROM QMS.QcItem a
                                WHERE 1 =1
                                AND a.QcClassId = @QcClassId
                                AND a.QcItemNo LIKE '___" + @PrincipleNo + @"'
                                ";
                        dynamicParameters.Add("QcClassId", QcClassId);
                        dynamicParameters.Add("PrincipleNo", PrincipleNo);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【類別+編碼找不到組合】，請先新增QcItemNo!");
                        #endregion

                        #region //判斷組合是否存在
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcItemPrinciple a
                                WHERE 1 =1
                                AND a.QcClassId = @QcClassId
                                AND a.QmmDetailId = @QmmDetailId
                                AND a.PrincipleNo = @PrincipleNo
                                AND a.PrincipleDesc = @PrincipleDesc
                                ";
                        dynamicParameters.Add("QcClassId", QcClassId);
                        dynamicParameters.Add("QmmDetailId", QmmDetailId);
                        dynamicParameters.Add("PrincipleNo", PrincipleNo);
                        dynamicParameters.Add("PrincipleDesc", PrincipleDesc);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【群組代碼】已存在，請重新輸入!");
                        #endregion

                        #region //判斷組合是否存在
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcItemPrinciple a
                                WHERE 1 =1
                                AND a.QcClassId = @QcClassId
                                AND a.QmmDetailId = @QmmDetailId
                                AND a.PrincipleNo = @PrincipleNo
                                ";
                        dynamicParameters.Add("QcClassId", QcClassId);
                        dynamicParameters.Add("QmmDetailId", QmmDetailId);
                        dynamicParameters.Add("PrincipleNo", PrincipleNo);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0)
                        {
                            #region //INSERT QMS.PrincipleDetail
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO QMS.PrincipleDetail (PrincipleId, PrincipleDesc
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        VALUES (@PrincipleId, @PrincipleDesc
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    PrincipleId,
                                    PrincipleDesc,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            int  rowsAffected = insertResult.Count();
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
                        else
                        {
                            #region //INSERT QMS.QcItemPrinciple
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO QMS.QcItemPrinciple (QcClassId, QmmDetailId, PrincipleNo, PrincipleDesc
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.PrincipleId
                                        VALUES (@QcClassId, @QmmDetailId, @PrincipleNo, @PrincipleDesc
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    QcClassId,
                                    QmmDetailId,
                                    PrincipleNo,
                                    PrincipleDesc,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            int rowsAffected = insertResult.Count();
                            foreach (var item in insertResult)
                            {
                                PrincipleId = item.PrincipleId;
                            }
                            #endregion

                            #region //INSERT QMS.PrincipleDetail
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO QMS.PrincipleDetail (PrincipleId, PrincipleDesc
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        VALUES (@PrincipleId, @PrincipleDesc
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    PrincipleId,
                                    PrincipleDesc,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                            insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();
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

        #region //AddQcItemCoding-- 新增項目編碼規則管理 -- Shintokuro 2024.05.17
        public string AddQcItemCoding(string QicNo, string QicName, string QicDesc)
        {
            try
            {
                string pattern = @"^[a-z][0-9]$";
                Regex regex = new Regex(pattern);
                if (!regex.IsMatch(QicNo)) throw new SystemException("【項目編碼代碼】格式為第1位數的小寫英文+數字共2字元!");
                if (QicNo.Length != 2) throw new SystemException("【項目編碼代碼】最大2字元!");
                if (QicName.Length <= 0) throw new SystemException("【項目編碼名稱】不能為空!");
                if (QicName.Length > 100) throw new SystemException("【項目編碼名稱】最大100字元!");
                if (QicDesc.Length > 150) throw new SystemException("【項目編碼描述】最大100字元!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷組合是否存在
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcItemCoding a
                                WHERE 1 =1
                                AND a.QicNo = @QicNo
                                AND a.CompanyId = @CompanyId
                                ";
                        dynamicParameters.Add("QicNo", QicNo);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【項目編碼代碼】已存在，請重新輸入!");
                        #endregion
                        

                        #region //INSERT QMS.QcItemPrinciple
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO QMS.QcItemCoding (CompanyId, QicNo, QicName, QicDesc, Status
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.QicId
                                        VALUES (@CompanyId, @QicNo, @QicName, @QicDesc, @Status
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                QicNo,
                                QicName,
                                QicDesc,
                                Status = "A",
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

        #region //AddQcMeasureInTheoryWorkTime -- 新增量測項目預測工時 -- GPAI 2024-12-17
        public string AddQcMeasureInTheoryWorkTime(int QmmDetailId, string ProductType, int QicId, string MeasureSize, float WorkTime)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {

                    #region //判斷量測機台是否存在
                    sql = @"SELECT TOP 1 1
                                FROM QMS.QmmDetail
                                WHERE QmmDetailId = @QmmDetailId";
                    dynamicParameters.Add("QmmDetailId", QmmDetailId);

                    var result2 = sqlConnection.Query(sql, dynamicParameters);
                    if (result2.Count() <= 0) throw new SystemException("量測機台不存在!");
                    #endregion

                    #region //判斷量測項目是否存在
                    sql = @"SELECT TOP 1 1
                                FROM QMS.QcItemCoding
                                WHERE QicId = @QicId";
                    dynamicParameters.Add("QicId", QicId);

                    var result3 = sqlConnection.Query(sql, dynamicParameters);
                    if (result3.Count() <= 0) throw new SystemException("量測項目不存在!");
                    #endregion


                    dynamicParameters = new DynamicParameters();
                    sql = @"INSERT INTO QMS.QcMeasureInTheoryWorkTime 
                            VALUES (@QmmDetailId, @ProductType, @QicId, @MeasureSize, @WorkTime, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)
                            ";

                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            QmmDetailId,
                                            ProductType,
                                            QicId,
                                            MeasureSize,
                                            WorkTime,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });
                    int rowsAffected = 0;

                    var insertResult = sqlConnection.Query(sql, dynamicParameters);

                    rowsAffected += insertResult.Count();

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = rowsAffected
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

        #region //Update
        #region //UpdateNgGroup -- 更新異常群組 -- Ann 2022.06.15
        public string UpdateNgGroup(int GroupId, string GroupNo, string GroupName, string GroupDesc, string Status)
        {
            try
            {
                if (GroupId <=0 ) throw new SystemException("【群組編號】不能為空!");
                if (GroupNo.Length <= 0) throw new SystemException("【群組代碼】不能為空!");
                if (GroupNo.Length > 100) throw new SystemException("【群組代碼】長度錯誤!");
                if (GroupName.Length <= 0) throw new SystemException("【群組名稱】不能為空!");
                if (GroupName.Length > 100) throw new SystemException("【群組名稱】長度錯誤!");
                if (GroupDesc.Length <= 0) throw new SystemException("【群組描述】不能為空!");
                if (GroupDesc.Length > 100) throw new SystemException("【群組描述】長度錯誤!");
                if (Status.Length <= 0) throw new SystemException("【群組狀態】不能為空!");
                if (Status.Length > 2) throw new SystemException("【群組狀態】長度錯誤!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷異常群組資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.DefectGroup
                                WHERE GroupId = @GroupId";
                        dynamicParameters.Add("GroupId", GroupId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("異常群組資料錯誤!");
                        #endregion

                        #region //判斷異常群組代碼是否已存在
                        sql = @"SELECT TOP 1 1
                                FROM QMS.DefectGroup a
                                WHERE a.GroupNo = @GroupNo
                                AND a.GroupId != @GroupId";
                        dynamicParameters.Add("GroupNo", GroupNo);
                        dynamicParameters.Add("GroupId", GroupId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【異常群組代碼】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE QMS.DefectGroup SET
                                GroupNo = @GroupNo,
                                GroupName = @GroupName,
                                GroupDesc = @GroupDesc,
                                Status = @Status,
                                LastModifiedBy = @LastModifiedBy,
                                LastModifiedDate = @LastModifiedDate
                                WHERE GroupId = @GroupId";
                        var parametersObject = new
                        {
                            GroupNo,
                            GroupName,
                            GroupDesc,
                            Status,
                            LastModifiedBy,
                            LastModifiedDate,
                            GroupId
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

        #region //UpdateNgClass -- 更新異常類別 -- Ann 2022.06.15
        public string UpdateNgClass(int ClassId, int GroupId, string ClassNo, string ClassName, string ClassDesc, string Status)
        {
            try
            {
                if (ClassId <= 0) throw new SystemException("【類別編號】不能為空!");
                if (ClassNo.Length <= 0) throw new SystemException("【類別代碼】不能為空!");
                if (ClassNo.Length > 100) throw new SystemException("【類別代碼】長度錯誤!");
                if (ClassName.Length <= 0) throw new SystemException("【類別名稱】不能為空!");
                if (ClassName.Length > 100) throw new SystemException("【類別名稱】長度錯誤!");
                if (ClassDesc.Length <= 0) throw new SystemException("【類別描述】不能為空!");
                if (ClassDesc.Length > 100) throw new SystemException("【類別描述】長度錯誤!");
                if (Status.Length <= 0) throw new SystemException("【類別狀態】不能為空!");
                if (Status.Length > 2) throw new SystemException("【類別狀態】長度錯誤!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷異常類別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.DefectClass
                                WHERE ClassId = @ClassId";
                        dynamicParameters.Add("ClassId", ClassId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("異常類別資料錯誤!");
                        #endregion

                        #region //判斷異常類別代碼是否已存在
                        sql = @"SELECT TOP 1 1
                                FROM QMS.DefectClass a
                                WHERE a.ClassNo = @ClassNo
                                AND a.ClassId != @ClassId";
                        dynamicParameters.Add("ClassNo", ClassNo);
                        dynamicParameters.Add("ClassId", ClassId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【異常類別代碼】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE QMS.DefectClass SET
                                GroupId = @GroupId,
                                ClassNo = @ClassNo,
                                ClassName = @ClassName,
                                ClassDesc = @ClassDesc,
                                Status = @Status,
                                LastModifiedBy = @LastModifiedBy,
                                LastModifiedDate = @LastModifiedDate
                                WHERE ClassId = @ClassId";
                        var parametersObject = new
                        {
                            GroupId,
                            ClassNo,
                            ClassName,
                            ClassDesc,
                            Status,
                            LastModifiedBy,
                            LastModifiedDate,
                            ClassId
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

        #region //UpdateNgCause -- 更新異常原因 -- Ann 2022.06.15
        public string UpdateNgCause(int CauseId, int ClassId, string CauseNo, string CauseName
            , string DepartmentNo, string CauseDesc, string Status)
        {
            try
            {
                if (CauseId <= 0) throw new SystemException("【原因編號】不能為空!");
                if (CauseNo.Length <= 0) throw new SystemException("【原因代碼】不能為空!");
                if (CauseNo.Length > 100) throw new SystemException("【原因代碼】長度錯誤!");
                if (CauseName.Length <= 0) throw new SystemException("【原因名稱】不能為空!");
                if (CauseName.Length > 100) throw new SystemException("【原因名稱】長度錯誤!");
                if (DepartmentNo.Length <= 0) throw new SystemException("【部門代碼】不能為空!");
                if (DepartmentNo.Length > 100) throw new SystemException("【部門代碼】長度錯誤!");
                if (CauseDesc.Length <= 0) throw new SystemException("【原因描述】不能為空!");
                if (CauseDesc.Length > 100) throw new SystemException("【原因描述】長度錯誤!");
                if (Status.Length <= 0) throw new SystemException("【原因狀態】不能為空!");
                if (Status.Length > 2) throw new SystemException("【原因狀態】長度錯誤!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷異常原因資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.DefectCause
                                WHERE CauseId = @CauseId";
                        dynamicParameters.Add("CauseId", CauseId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("異常原因資料錯誤!");
                        #endregion

                        #region //判斷異常原因代碼是否已存在
                        sql = @"SELECT TOP 1 1
                                FROM QMS.DefectCause a
                                WHERE a.CauseNo = @CauseNo
                                AND a.CauseId != @CauseId";
                        dynamicParameters.Add("CauseNo", CauseNo);
                        dynamicParameters.Add("CauseId", CauseId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【異常原因代碼】重複，請重新輸入!");
                        #endregion

                        #region //查詢部門ID
                        sql = @"SELECT a.DepartmentId
                                FROM BAS.Department a
                                WHERE a.DepartmentNo = @DepartmentNo";
                        dynamicParameters.Add("DepartmentNo", DepartmentNo);

                        var ResponsibleDepartment = 0;
                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() < 0) throw new SystemException("【部門資料】異常，請重新輸入!");
                        else
                        {
                            foreach (var item in result3)
                            {
                                ResponsibleDepartment = item.DepartmentId;
                            }
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE QMS.DefectCause SET
                                ClassId = @ClassId,
                                CauseNo = @CauseNo,
                                CauseName = @CauseName,
                                CauseDesc = @CauseDesc,
                                ResponsibleDepartment = @ResponsibleDepartment,
                                Status = @Status,
                                LastModifiedBy = @LastModifiedBy,
                                LastModifiedDate = @LastModifiedDate
                                WHERE CauseId = @CauseId";
                        var parametersObject = new
                        {
                            ClassId,
                            CauseNo,
                            CauseName,
                            CauseDesc,
                            ResponsibleDepartment,
                            Status,
                            LastModifiedBy,
                            LastModifiedDate,
                            CauseId
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

        #region //UpdateRepairGroup -- 更新維修群組資料 -- Ann 2022.06.21
        public string UpdateRepairGroup(int GroupId, string GroupNo, string GroupName, string GroupDesc, string Status)
        {
            try
            {
                if (GroupId <= 0) throw new SystemException("【群組編號】不能為空!");
                if (GroupNo.Length <= 0) throw new SystemException("【群組代碼】不能為空!");
                if (GroupNo.Length > 100) throw new SystemException("【群組代碼】長度錯誤!");
                if (GroupName.Length <= 0) throw new SystemException("【群組名稱】不能為空!");
                if (GroupName.Length > 100) throw new SystemException("【群組名稱】長度錯誤!");
                if (GroupDesc.Length <= 0) throw new SystemException("【群組描述】不能為空!");
                if (GroupDesc.Length > 100) throw new SystemException("【群組描述】長度錯誤!");
                if (Status.Length <= 0) throw new SystemException("【群組狀態】不能為空!");
                if (Status.Length > 2) throw new SystemException("【群組狀態】長度錯誤!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷維修群組資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.RepairGroup
                                WHERE GroupId = @GroupId";
                        dynamicParameters.Add("GroupId", GroupId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("維修群組資料錯誤!");
                        #endregion

                        #region //判斷維修群組代碼是否已存在
                        sql = @"SELECT TOP 1 1
                                FROM QMS.RepairGroup a
                                WHERE a.GroupNo = @GroupNo
                                AND a.GroupId != @GroupId";
                        dynamicParameters.Add("GroupNo", GroupNo);
                        dynamicParameters.Add("GroupId", GroupId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【維修群組代碼】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE QMS.RepairGroup SET
                                GroupNo = @GroupNo,
                                GroupName = @GroupName,
                                GroupDesc = @GroupDesc,
                                Status = @Status,
                                LastModifiedBy = @LastModifiedBy,
                                LastModifiedDate = @LastModifiedDate
                                WHERE GroupId = @GroupId";
                        var parametersObject = new
                        {
                            GroupNo,
                            GroupName,
                            GroupDesc,
                            Status,
                            LastModifiedBy,
                            LastModifiedDate,
                            GroupId
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

        #region //UpdateRepairClass -- 更新維修類別資料 -- Ann 2022.06.21
        public string UpdateRepairClass(int ClassId, int GroupId, string ClassNo, string ClassName, string ClassDesc, string Status)
        {
            try
            {
                if (ClassId <= 0) throw new SystemException("【類別編號】不能為空!");
                if (ClassNo.Length <= 0) throw new SystemException("【類別代碼】不能為空!");
                if (ClassNo.Length > 100) throw new SystemException("【類別代碼】長度錯誤!");
                if (ClassName.Length <= 0) throw new SystemException("【類別名稱】不能為空!");
                if (ClassName.Length > 100) throw new SystemException("【類別名稱】長度錯誤!");
                if (ClassDesc.Length <= 0) throw new SystemException("【類別描述】不能為空!");
                if (ClassDesc.Length > 100) throw new SystemException("【類別描述】長度錯誤!");
                if (Status.Length <= 0) throw new SystemException("【類別狀態】不能為空!");
                if (Status.Length > 2) throw new SystemException("【類別狀態】長度錯誤!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷維修類別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.RepairClass
                                WHERE ClassId = @ClassId";
                        dynamicParameters.Add("ClassId", ClassId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("維修類別資料錯誤!");
                        #endregion

                        #region //判斷維修類別代碼是否已存在
                        sql = @"SELECT TOP 1 1
                                FROM QMS.RepairClass a
                                WHERE a.ClassNo = @ClassNo
                                AND a.ClassId != @ClassId";
                        dynamicParameters.Add("ClassNo", ClassNo);
                        dynamicParameters.Add("ClassId", ClassId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【維修類別代碼】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE QMS.RepairClass SET
                                GroupId = @GroupId,
                                ClassNo = @ClassNo,
                                ClassName = @ClassName,
                                ClassDesc = @ClassDesc,
                                Status = @Status,
                                LastModifiedBy = @LastModifiedBy,
                                LastModifiedDate = @LastModifiedDate
                                WHERE ClassId = @ClassId";
                        var parametersObject = new
                        {
                            GroupId,
                            ClassNo,
                            ClassName,
                            ClassDesc,
                            Status,
                            LastModifiedBy,
                            LastModifiedDate,
                            ClassId
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

        #region //UpdateRepairCause -- 更新維修原因資料 -- Ann 2022.06.21
        public string UpdateRepairCause(int CauseId, int ClassId, string CauseNo, string CauseName
            , string DepartmentNo, string CauseDesc, string Status)
        {
            try
            {
                if (CauseId <= 0) throw new SystemException("【原因編號】不能為空!");
                if (CauseNo.Length <= 0) throw new SystemException("【原因代碼】不能為空!");
                if (CauseNo.Length > 100) throw new SystemException("【原因代碼】長度錯誤!");
                if (CauseName.Length <= 0) throw new SystemException("【原因名稱】不能為空!");
                if (CauseName.Length > 100) throw new SystemException("【原因名稱】長度錯誤!");
                if (DepartmentNo.Length <= 0) throw new SystemException("【部門代碼】不能為空!");
                if (DepartmentNo.Length > 100) throw new SystemException("【部門代碼】長度錯誤!");
                if (CauseDesc.Length <= 0) throw new SystemException("【原因描述】不能為空!");
                if (CauseDesc.Length > 100) throw new SystemException("【原因描述】長度錯誤!");
                if (Status.Length <= 0) throw new SystemException("【原因狀態】不能為空!");
                if (Status.Length > 2) throw new SystemException("【原因狀態】長度錯誤!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷維修原因資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.RepairCause
                                WHERE CauseId = @CauseId";
                        dynamicParameters.Add("CauseId", CauseId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("異常維修資料錯誤!");
                        #endregion

                        #region //判斷維修原因代碼是否已存在
                        sql = @"SELECT TOP 1 1
                                FROM QMS.RepairCause a
                                WHERE a.CauseNo = @CauseNo
                                AND a.CauseId != @CauseId";
                        dynamicParameters.Add("CauseNo", CauseNo);
                        dynamicParameters.Add("CauseId", CauseId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【維修原因代碼】重複，請重新輸入!");
                        #endregion

                        #region //查詢部門ID
                        sql = @"SELECT a.DepartmentId
                                FROM BAS.Department a
                                WHERE a.DepartmentNo = @DepartmentNo";
                        dynamicParameters.Add("DepartmentNo", DepartmentNo);

                        var ResponsibleDepartment = 0;
                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() < 0) throw new SystemException("【部門資料】異常，請重新輸入!");
                        else
                        {
                            foreach (var item in result3)
                            {
                                ResponsibleDepartment = item.DepartmentId;
                            }
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE QMS.RepairCause SET
                                ClassId = @ClassId,
                                CauseNo = @CauseNo,
                                CauseName = @CauseName,
                                CauseDesc = @CauseDesc,
                                ResponsibleDepartment = @ResponsibleDepartment,
                                Status = @Status,
                                LastModifiedBy = @LastModifiedBy,
                                LastModifiedDate = @LastModifiedDate
                                WHERE CauseId = @CauseId";
                        var parametersObject = new
                        {
                            ClassId,
                            CauseNo,
                            CauseName,
                            CauseDesc,
                            ResponsibleDepartment,
                            Status,
                            LastModifiedBy,
                            LastModifiedDate,
                            CauseId
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

        #region //UpdateQcGroup -- 更新量測群組資料 -- Ted 2022.10.03
        public string UpdateQcGroup(int QcGroupId, string QcGroupNo, string QcGroupName, string QcGroupDesc, string Status)
        {
            try
            {
                if (QcGroupId <= 0) throw new SystemException("【群組編號】不能為空!");
                //if (QcGroupNo.Length <= 0) throw new SystemException("【群組代碼】不能為空!");
                //if (QcGroupNo.Length > 100) throw new SystemException("【群組代碼】長度錯誤!");
                if (QcGroupName.Length <= 0) throw new SystemException("【群組名稱】不能為空!");
                if (QcGroupName.Length > 100) throw new SystemException("【群組名稱】長度錯誤!");
                if (QcGroupDesc.Length <= 0) throw new SystemException("【群組描述】不能為空!");
                if (QcGroupDesc.Length > 100) throw new SystemException("【群組描述】長度錯誤!");
                if (Status.Length <= 0) throw new SystemException("【群組狀態】不能為空!");
                if (Status.Length > 2) throw new SystemException("【群組狀態】長度錯誤!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷量測群組資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcGroup
                                WHERE QcGroupId = @QcGroupId";
                        dynamicParameters.Add("QcGroupId", QcGroupId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("量測群組資料錯誤!");
                        #endregion

                        #region //判斷量測群組代碼是否已存在
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcGroup a
                                WHERE a.QcGroupNo = @QcGroupNo
                                AND a.QcGroupId != @QcGroupId";
                        dynamicParameters.Add("QcGroupNo", QcGroupNo);
                        dynamicParameters.Add("QcGroupId", QcGroupId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【維修群組代碼】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE QMS.QcGroup SET
                                --QcGroupNo = @QcGroupNo,
                                QcGroupName = @QcGroupName,
                                QcGroupDesc = @QcGroupDesc,
                                Status = @Status,
                                LastModifiedBy = @LastModifiedBy,
                                LastModifiedDate = @LastModifiedDate
                                WHERE QcGroupId = @QcGroupId";
                        var parametersObject = new
                        {
                            QcGroupNo,
                            QcGroupName,
                            QcGroupDesc,
                            Status,
                            LastModifiedBy,
                            LastModifiedDate,
                            QcGroupId
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

        #region //UpdateQcClass -- 更新量測類別資料 -- Ted 2022.10.03
        public string UpdateQcClass(int QcClassId, int QcGroupId, string QcClassNo, string QcClassName, string QcClassDesc, string Status)
        {
            try
            {
                if (QcClassId <= 0) throw new SystemException("【類別編號】不能為空!");
                //if (QcClassNo.Length <= 0) throw new SystemException("【類別代碼】不能為空!");
                //if (QcClassNo.Length > 100) throw new SystemException("【類別代碼】長度錯誤!");
                if (QcClassName.Length <= 0) throw new SystemException("【類別名稱】不能為空!");
                if (QcClassName.Length > 100) throw new SystemException("【類別名稱】長度錯誤!");
                if (QcClassDesc.Length <= 0) throw new SystemException("【類別描述】不能為空!");
                if (QcClassDesc.Length > 100) throw new SystemException("【類別描述】長度錯誤!");
                if (Status.Length <= 0) throw new SystemException("【類別狀態】不能為空!");
                if (Status.Length > 2) throw new SystemException("【類別狀態】長度錯誤!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷維修類別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcClass
                                WHERE QcClassId = @QcClassId";
                        dynamicParameters.Add("QcClassId", QcClassId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("維修類別資料錯誤!");
                        #endregion

                        #region //判斷維修類別代碼是否已存在
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcClass a
                                WHERE a.QcClassNo = @QcClassNo
                                AND a.QcClassId != @QcClassId";
                        dynamicParameters.Add("QcClassNo", QcClassNo);
                        dynamicParameters.Add("QcClassId", QcClassId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【維修類別代碼】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE QMS.QcClass SET
                                --QcGroupId = @QcGroupId,
                                --QcClassNo = @QcClassNo,
                                QcClassName = @QcClassName,
                                QcClassDesc = @QcClassDesc,
                                Status = @Status,
                                LastModifiedBy = @LastModifiedBy,
                                LastModifiedDate = @LastModifiedDate
                                WHERE QcClassId = @QcClassId";
                        var parametersObject = new
                        {
                            QcGroupId,
                            QcClassNo,
                            QcClassName,
                            QcClassDesc,
                            Status,
                            LastModifiedBy,
                            LastModifiedDate,
                            QcClassId
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

        #region //UpdateQcItem -- 更新量測項目資料 -- Ted 2022.10.03
        public string UpdateQcItem(int QcItemId, int QcClassId, string QcItemNo, string QcItemName
            , string QcItemDesc, string QcType, string QcItemType, string Remark, string Status)
        {
            try
            {
                if (QcItemId <= 0) throw new SystemException("【項目編號】不能為空!");
                //if (QcItemNo.Length <= 0) throw new SystemException("【項目代碼】不能為空!");
                //if (QcItemNo.Length > 100) throw new SystemException("【項目代碼】長度錯誤!");
                if (QcItemName.Length <= 0) throw new SystemException("【項目名稱】不能為空!");
                if (QcItemName.Length > 100) throw new SystemException("【項目名稱】長度錯誤!");
                if (QcItemDesc.Length <= 0) throw new SystemException("【項目描述】不能為空!");
                if (QcItemDesc.Length > 100) throw new SystemException("【項目描述】長度錯誤!");
                if (QcType.Length <= 0) throw new SystemException("【檢驗類別】不能為空!");
                if (QcItemType.Length <= 0) throw new SystemException("【量測項目輸入方式】不能為空!");
                if (QcItemType.Length > 13) throw new SystemException("【量測項目輸入方式】長度錯誤!");
                if (Status.Length <= 0) throw new SystemException("【項目狀態】不能為空!");
                if (Status.Length > 2) throw new SystemException("【項目狀態】長度錯誤!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        QcItemType = QcItemType.Split('-')[0];


                        #region //判斷項目資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcItem
                                WHERE QcItemId = @QcItemId";
                        dynamicParameters.Add("QcItemId", QcItemId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("項目資料錯誤!");
                        #endregion

                        #region //判斷項目代碼是否已存在
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcItem a
                                WHERE a.QcItemNo = @QcItemNo
                                AND a.QcItemId != @QcItemId";
                        dynamicParameters.Add("QcItemNo", QcItemNo);
                        dynamicParameters.Add("QcItemId", QcItemId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【項目代碼】重複，請重新輸入!");
                        #endregion

                        

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE QMS.QcItem SET
                                --QcClassId = @QcClassId,
                                --QcItemNo = @QcItemNo,
                                QcItemName = @QcItemName,
                                QcItemDesc = @QcItemDesc,
                                QcType = @QcType,
                                QcItemType = @QcItemType,
                                Remark = @Remark,
                                Status = @Status,
                                LastModifiedBy = @LastModifiedBy,
                                LastModifiedDate = @LastModifiedDate
                                WHERE QcItemId = @QcItemId";
                        var parametersObject = new
                        {
                            QcClassId,
                            QcItemNo,
                            QcItemName,
                            QcItemDesc,
                            QcType,
                            QcItemType,
                            Remark,
                            Status,
                            LastModifiedBy,
                            LastModifiedDate,
                            QcItemId
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

        #region //UpdateQiDetail -- 更新量測項目機型單身 -- Ted 2022.09.27
        public string UpdateQiDetail(int QiDetailId, int QcItemId, int QcMachineModeId
            )
        {
            try
            {
                if (QiDetailId <= 0) throw new SystemException("【量測項目】資料錯誤");
                if (QcItemId <= 0) throw new SystemException("【量測項目】不能為空!");
                if (QcMachineModeId <= 0) throw new SystemException("【量測機型】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷量測項目是否正確
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcItem
                                WHERE CompanyId = @CompanyId
                                AND QcItemId = @QcItemId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("QcItemId", QcItemId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() < 0) throw new SystemException("【量測項目】錯誤，請重新輸入!");
                        #endregion

                        #region //判斷量測機型是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcMachineMode
                                WHERE CompanyId = @CompanyId
                                AND QcMachineModeId = @QcMachineModeId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("QcMachineModeId", QcMachineModeId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() < 0) throw new SystemException("【量測機型】錯誤，請重新輸入!");
                        #endregion

                        #region //判斷機型是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QiDetail
                                WHERE 1=1
                                AND QcItemId = @QcItemId
                                AND QcMachineModeId = @QcMachineModeId
                                AND QiDetailId != QiDetailId";
                        dynamicParameters.Add("QcItemId", QcItemId);
                        dynamicParameters.Add("QcMachineModeId", QcMachineModeId);
                        dynamicParameters.Add("QiDetailId", QiDetailId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【機型】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE QMS.QiDetail SET
                                QcMachineModeId = @QcMachineModeId,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE QcItemId = @QcItemId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                QcMachineModeId,
                                LastModifiedDate,
                                LastModifiedBy,
                                QcItemId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #region //UpdateQcMachineMode -- 更新量測機型資料 -- Ted 2022.09.27
        public string UpdateQcMachineMode(int QcMachineModeId, string QcMachineModeNo, string QcMachineModeName, string QcMachineModeDesc,
            string ItemNo)
        {
            try
            {
                string pattern = @"^[a-zA-Z0-9]+$";
                Regex regex = new Regex(pattern);

                if (QcMachineModeNo.Length <= 0) throw new SystemException("【量測編號】不能為空!");
                if (QcMachineModeName.Length <= 0) throw new SystemException("【量測名稱】不能為空!");
                if (ItemNo.Length != 3) throw new SystemException("【項目編號】格式錯誤,3字元且只能英數字!");
                if (!regex.IsMatch(ItemNo)) throw new SystemException("【項目編號】格式錯誤,3字元且只能英數字!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷機型資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcMachineMode
                                WHERE QcMachineModeId = @QcMachineModeId";
                        dynamicParameters.Add("QcMachineModeId", QcMachineModeId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("機型資料錯誤!");
                        #endregion

                        #region //判斷機型編號是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcMachineMode
                                WHERE CompanyId = @CompanyId
                                AND QcMachineModeNo = @QcMachineModeNo
                                AND QcMachineModeId != @QcMachineModeId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("QcMachineModeNo", QcMachineModeNo);
                        dynamicParameters.Add("QcMachineModeId", QcMachineModeId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【機型編號】重複，請重新輸入!");
                        #endregion

                        #region //判斷機型名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcMachineMode
                                WHERE CompanyId = @CompanyId
                                AND QcMachineModeName = @QcMachineModeName
                                AND QcMachineModeId != @QcMachineModeId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("QcMachineModeName", QcMachineModeName);
                        dynamicParameters.Add("QcMachineModeId", QcMachineModeId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【機型名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE QMS.QcMachineMode SET
                                QcMachineModeNo = @QcMachineModeNo,
                                QcMachineModeName = @QcMachineModeName,
                                QcMachineModeDesc = @QcMachineModeDesc,
                                ItemNo = @ItemNo,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE QcMachineModeId = @QcMachineModeId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                QcMachineModeNo,
                                QcMachineModeName,
                                QcMachineModeDesc,
                                ItemNo,
                                LastModifiedDate,
                                LastModifiedBy,
                                QcMachineModeId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #region //UpdateQmmDetailStatus -- 更新量測機型的機台單身狀態 -- Ted 2022.09.27
        public string UpdateQmmDetailStatus(int QmmDetailId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷機台單身資訊是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM QMS.QmmDetail
                                WHERE QmmDetailId = @QmmDetailId";
                        dynamicParameters.Add("QmmDetailId", QmmDetailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("機台單身資料錯誤!");

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
                        sql = @"UPDATE QMS.QmmDetail SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE QmmDetailId = @QmmDetailId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                QmmDetailId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #region //UpdateQcGroupNew -- 更新量測群組資料(新) -- Shintokuro 2024.05.16
        public string UpdateQcGroupNew(int QcGroupId, string QcGroupName, string QcGroupDesc, string Status)
        {
            try
            {
                if (QcGroupId <= 0) throw new SystemException("【群組編號】不能為空!");
                if (QcGroupName.Length <= 0) throw new SystemException("【群組名稱】不能為空!");
                if (QcGroupName.Length > 100) throw new SystemException("【群組名稱】長度錯誤!");
                if (QcGroupDesc.Length <= 0) throw new SystemException("【群組描述】不能為空!");
                if (QcGroupDesc.Length > 100) throw new SystemException("【群組描述】長度錯誤!");
                if (Status.Length <= 0) throw new SystemException("【群組狀態】不能為空!");
                if (Status.Length > 2) throw new SystemException("【群組狀態】長度錯誤!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷量測群組資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcGroup
                                WHERE QcGroupId = @QcGroupId";
                        dynamicParameters.Add("QcGroupId", QcGroupId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("量測群組資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE QMS.QcGroup SET
                                QcGroupName = @QcGroupName,
                                QcGroupDesc = @QcGroupDesc,
                                Status = @Status,
                                LastModifiedBy = @LastModifiedBy,
                                LastModifiedDate = @LastModifiedDate
                                WHERE QcGroupId = @QcGroupId";
                        var parametersObject = new
                        {
                            QcGroupName,
                            QcGroupDesc,
                            Status,
                            LastModifiedBy,
                            LastModifiedDate,
                            QcGroupId
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

        #region //UpdateQcClassNew -- 更新量測類別資料(新) -- Shintokuro 2024.05.16
        public string UpdateQcClassNew(int QcClassId, int QcGroupId, string QcClassName, string QcClassDesc, string Status)
        {
            try
            {
                if (QcClassId <= 0) throw new SystemException("【類別編號】不能為空!");
                if (QcClassName.Length <= 0) throw new SystemException("【類別名稱】不能為空!");
                if (QcClassName.Length > 100) throw new SystemException("【類別名稱】長度錯誤!");
                if (QcClassDesc.Length <= 0) throw new SystemException("【類別描述】不能為空!");
                if (QcClassDesc.Length > 100) throw new SystemException("【類別描述】長度錯誤!");
                if (Status.Length <= 0) throw new SystemException("【類別狀態】不能為空!");
                if (Status.Length > 2) throw new SystemException("【類別狀態】長度錯誤!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷維修類別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcClass
                                WHERE QcClassId = @QcClassId";
                        dynamicParameters.Add("QcClassId", QcClassId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("維修類別資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE QMS.QcClass SET
                                QcGroupId = @QcGroupId,
                                QcClassName = @QcClassName,
                                QcClassDesc = @QcClassDesc,
                                Status = @Status,
                                LastModifiedBy = @LastModifiedBy,
                                LastModifiedDate = @LastModifiedDate
                                WHERE QcClassId = @QcClassId";
                        var parametersObject = new
                        {
                            QcGroupId,
                            QcClassName,
                            QcClassDesc,
                            Status,
                            LastModifiedBy,
                            LastModifiedDate,
                            QcClassId
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

        #region //UpdateQcItemNew -- 更新量測項目資料(新) -- Shintokuro 2024.05.16
        public string UpdateQcItemNew(int QcItemId, string QcItemName
            , string QcItemDesc, string QcType, string QcItemType, string Remark, string Status)
        {
            try
            {
                if (QcItemId <= 0) throw new SystemException("【項目編號】不能為空!");
                if (QcItemName.Length <= 0) throw new SystemException("【項目名稱】不能為空!");
                if (QcItemName.Length > 100) throw new SystemException("【項目名稱】長度錯誤!");
                if (QcItemDesc.Length <= 0) throw new SystemException("【項目描述】不能為空!");
                if (QcItemDesc.Length > 100) throw new SystemException("【項目描述】長度錯誤!");
                if (QcType.Length <= 0) throw new SystemException("【檢驗類別】不能為空!");
                if (QcType.Length > 15) throw new SystemException("【檢驗類別】長度錯誤!");
                if (QcItemType.Length <= 0) throw new SystemException("【量測項目輸入方式】不能為空!");
                if (QcItemType.Length > 13) throw new SystemException("【量測項目輸入方式】長度錯誤!");
                if (Status.Length <= 0) throw new SystemException("【項目狀態】不能為空!");
                if (Status.Length > 2) throw new SystemException("【項目狀態】長度錯誤!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        QcType = QcType.Split('-')[0];
                        QcItemType = QcItemType.Split('-')[0];


                        #region //判斷項目資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcItem
                                WHERE QcItemId = @QcItemId";
                        dynamicParameters.Add("QcItemId", QcItemId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("項目資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE QMS.QcItem SET
                                QcItemName = @QcItemName,
                                QcItemDesc = @QcItemDesc,
                                QcType = @QcType,
                                QcItemType = @QcItemType,
                                Remark = @Remark,
                                Status = @Status,
                                LastModifiedBy = @LastModifiedBy,
                                LastModifiedDate = @LastModifiedDate
                                WHERE QcItemId = @QcItemId";
                        var parametersObject = new
                        {
                            QcItemName,
                            QcItemDesc,
                            QcType,
                            QcItemType,
                            Remark,
                            Status,
                            LastModifiedBy,
                            LastModifiedDate,
                            QcItemId
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

        #region //UpdateQcItemPrinciple-- 更新量測項目編碼原則 -- Shintokuro 2024.05.15
        public string UpdateQcItemPrinciple(int PrincipleId, int QcClassId, int QmmDetailId, string PrincipleNo, string PrincipleDesc)
        {
            try
            {
                //string pattern = @"^[A-Z]$";
                //Regex regex = new Regex(pattern);

                //if (!regex.IsMatch(QcGroupNo)) throw new SystemException("【群組代碼】格式為1位數的大寫英文!");
                if (PrincipleId <= 0) throw new SystemException("【Id】不能為空!");
                if (QcClassId <= 0) throw new SystemException("【產品類別】不能為空!");
                if (QmmDetailId <= 0) throw new SystemException("【量測機台】不能為空!");
                if (PrincipleNo.Length > 4) throw new SystemException("【編碼】最大4字元!");
                if (PrincipleNo.Length <= 0) throw new SystemException("【編碼】不能為空!");
                if (PrincipleDesc.Length > 150) throw new SystemException("【編碼名稱描述】最大150字元!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷組合是否存在
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcItemPrinciple a
                                WHERE 1 =1
                                AND a.QcClassId = @QcClassId
                                AND a.QmmDetailId = @QmmDetailId
                                AND a.PrincipleNo = @PrincipleNo
                                AND a.PrincipleDesc = @PrincipleDesc
                                AND a.PrincipleId != @PrincipleId
                                ";
                        dynamicParameters.Add("QcClassId", QcClassId);
                        dynamicParameters.Add("QmmDetailId", QmmDetailId);
                        dynamicParameters.Add("PrincipleNo", PrincipleNo);
                        dynamicParameters.Add("PrincipleDesc", PrincipleDesc);
                        dynamicParameters.Add("PrincipleId", PrincipleId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【群組代碼】已存在，請重新輸入!");
                        #endregion

                        #region //判斷QcClass是否存在
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcClass a
                                WHERE 1 =1
                                AND a.QcClassId = @QcClassId
                                ";
                        dynamicParameters.Add("QcClassId", QcClassId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【產品類別】找不到，請重新輸入!");
                        #endregion

                        #region //判斷QmmDetailId是否存在
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QmmDetail a
                                WHERE 1 =1
                                AND a.QmmDetailId = @QmmDetailId
                                ";
                        dynamicParameters.Add("QmmDetailId", QmmDetailId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【量測機台】找不到，請重新輸入!");
                        #endregion

                        #region //UPDATE QMS.QcItemPrinciple
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE QMS.QcItemPrinciple SET
                                QcClassId = @QcClassId
                                QmmDetailId = @QmmDetailId
                                PrincipleNo = @PrincipleNo
                                PrincipleDesc = @PrincipleDesc
                                WHERE PrincipleId = @PrincipleId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                QcClassId,
                                QmmDetailId,
                                PrincipleNo,
                                PrincipleDesc,
                                PrincipleId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = insertResult.Count();

                        #endregion

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

        #region //UpdateQcItemCoding-- 更新項目編碼規則管理 -- Shintokuro 2024.05.17
        public string UpdateQcItemCoding(int QicId, string QicName, string QicDesc)
        {
            try
            {
              
                if (QicName.Length <= 0) throw new SystemException("【項目編碼名稱】不能為空!");
                if (QicName.Length > 100) throw new SystemException("【項目編碼名稱】最大100字元!");
                if (QicDesc.Length > 150) throw new SystemException("【項目編碼描述】最大100字元!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //INSERT QMS.QcItemPrinciple
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE QMS.QcItemCoding SET
                                QicName = @QicName,
                                QicDesc = @QicDesc
                                WHERE QicId = @QicId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                QicName,
                                QicDesc,
                                QicId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = insertResult.Count();
                        #endregion


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

        #region //UpdateQcItemCodingStatus -- 更新項目編碼規則管理狀態 -- Shintokru 2024.05.17
        public string UpdateQcItemCodingStatus(int QicId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷車間資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 Status
                                FROM QMS.QcItemCoding
                                WHERE CompanyId = @CompanyId
                                AND QicId = @QicId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("QicId", QicId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("車間資料錯誤!");

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

                        #region //更新SQL
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE QMS.QcItemCoding SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE QicId = @QicId
                                AND CompanyId = @CompanyId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                QicId,
                                CompanyId = CurrentCompany
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

        #region //UpdateQcMeasureInTheoryWorkTime -- 編輯量測項目預測工時 -- GPAI 2024-12-17
        public string UpdateQcMeasureInTheoryWorkTime(int QmwtId, int QmmDetailId, string ProductType, int QicId, string MeasureSize, float WorkTime)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //判斷量測項目預測工時是否存在
                    sql = @"SELECT TOP 1 1
                                FROM QMS.QcMeasureInTheoryWorkTime
                                WHERE QmwtId = @QmwtId";
                    dynamicParameters.Add("QmwtId", QmwtId);

                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0) throw new SystemException("量測項目編碼原則不存在!");
                    #endregion

                    #region //判斷量測機台是否存在
                    sql = @"SELECT TOP 1 1
                                FROM QMS.QmmDetail
                                WHERE QmmDetailId = @QmmDetailId";
                    dynamicParameters.Add("QmmDetailId", QmmDetailId);

                    var result2 = sqlConnection.Query(sql, dynamicParameters);
                    if (result2.Count() <= 0) throw new SystemException("量測機台不存在!");
                    #endregion

                    #region //判斷量測項目是否存在
                    sql = @"SELECT TOP 1 1
                                FROM QMS.QcItemCoding
                                WHERE QicId = @QicId";
                    dynamicParameters.Add("QicId", QicId);

                    var result3 = sqlConnection.Query(sql, dynamicParameters);
                    if (result3.Count() <= 0) throw new SystemException("量測項目不存在!");
                    #endregion

                    dynamicParameters = new DynamicParameters();
                    sql = @"UPDATE QMS.QcMeasureInTheoryWorkTime SET
                                QmmDetailId = @QmmDetailId,
                                ProductType = @ProductType,
                                QicId = @QicId,
                                MeasureSize = @MeasureSize,
                                WorkTime = @WorkTime,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE QmwtId = @QmwtId
                            ";

                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            QmmDetailId,
                                            ProductType,
                                            QicId,
                                            MeasureSize,
                                            WorkTime,
                                            CreateBy,
                                            LastModifiedDate,
                                            LastModifiedBy,
                                            QmwtId
                                        });
                    int rowsAffected = 0;

                    var insertResult = sqlConnection.Query(sql, dynamicParameters);

                    rowsAffected += insertResult.Count();

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = rowsAffected
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

        #region //Delete
        #region //DeleteNgGroup -- 刪除異常群組 -- Ann 2022.06.16
        public string DeleteNgGroup(int GroupId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷異常群組資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM QMS.DefectGroup
                                WHERE GroupId = @GroupId";
                        dynamicParameters.Add("GroupId", GroupId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("異常群組資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除關聯table
                        //QMS.DefectCause
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM QMS.DefectCause a
                                LEFT JOIN QMS.DefectClass b ON a.ClassId = b.ClassId
                                LEFT JOIN QMS.DefectGroup c ON b.GroupId = c.GroupId
                                WHERE b.GroupId = @GroupId";
                        dynamicParameters.Add("GroupId", GroupId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        //QMS.DefectClass
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM  QMS.DefectClass a
                                LEFT JOIN QMS.DefectGroup b ON a.GroupId = b.GroupId
                                WHERE a.GroupId = @GroupId";
                        dynamicParameters.Add("GroupId", GroupId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE FROM QMS.DefectGroup
                                WHERE GroupId = @GroupId";
                        dynamicParameters.Add("GroupId", GroupId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteNgClass -- 刪除異常類別 -- Ann 2022.06.16
        public string DeleteNgClass(int ClassId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷異常類別資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM QMS.DefectClass
                                WHERE ClassId = @ClassId";
                        dynamicParameters.Add("ClassId", ClassId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("異常類別資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除關聯table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE QMS.DefectCause
                                WHERE ClassId = @ClassId";
                        dynamicParameters.Add("ClassId", ClassId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE QMS.DefectClass
                                WHERE ClassId = @ClassId";
                        dynamicParameters.Add("ClassId", ClassId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteNgCause -- 刪除異常原因 -- Ann 2022.06.16
        public string DeleteNgCause(int CauseId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷異常原因資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM QMS.DefectCause
                                WHERE CauseId = @CauseId";
                        dynamicParameters.Add("CauseId", CauseId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("異常原因資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE QMS.DefectCause
                                WHERE CauseId = @CauseId";
                        dynamicParameters.Add("CauseId", CauseId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteRepairGroup -- 刪除維修群組資料 -- Ann 2022.06.21
        public string DeleteRepairGroup(int GroupId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷維修群組資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM QMS.RepairGroup
                                WHERE GroupId = @GroupId";
                        dynamicParameters.Add("GroupId", GroupId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("維修群組資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除關聯table
                        //QMS.DefectCause
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM QMS.RepairCause a
                                LEFT JOIN QMS.RepairClass b ON a.ClassId = b.ClassId
                                LEFT JOIN QMS.RepairGroup c ON b.GroupId = c.GroupId
                                WHERE b.GroupId = @GroupId";
                        dynamicParameters.Add("GroupId", GroupId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        //QMS.DefectClass
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM  QMS.RepairClass a
                                LEFT JOIN QMS.RepairGroup b ON a.GroupId = b.GroupId
                                WHERE a.GroupId = @GroupId";
                        dynamicParameters.Add("GroupId", GroupId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE FROM QMS.RepairGroup
                                WHERE GroupId = @GroupId";
                        dynamicParameters.Add("GroupId", GroupId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteRepairClass -- 刪除維修類別資料 -- Ann 2022.06.21
        public string DeleteRepairClass(int ClassId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷異常類別資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM QMS.RepairClass
                                WHERE ClassId = @ClassId";
                        dynamicParameters.Add("ClassId", ClassId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("維修類別資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除關聯table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE QMS.RepairCause
                                WHERE ClassId = @ClassId";
                        dynamicParameters.Add("ClassId", ClassId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE QMS.RepairClass
                                WHERE ClassId = @ClassId";
                        dynamicParameters.Add("ClassId", ClassId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteRepairCause -- 刪除維修原因資料 -- Ann 2022.06.21
        public string DeleteRepairCause(int CauseId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷異常原因資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM QMS.RepairCause
                                WHERE CauseId = @CauseId";
                        dynamicParameters.Add("CauseId", CauseId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("維修原因資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE QMS.RepairCause
                                WHERE CauseId = @CauseId";
                        dynamicParameters.Add("CauseId", CauseId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteQcGroup -- 刪除量測群組資料 -- Ted 2022.10.03
        public string DeleteQcGroup(int QcGroupId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷量測群組資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcGroup
                                WHERE QcGroupId = @QcGroupId";
                        dynamicParameters.Add("QcGroupId", QcGroupId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("量測資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除關聯table
                        //QMS.QiDetail
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM QMS.QiDetail a
                                LEFT JOIN QMS.QcItem b ON a.QcItemId = b.QcItemId
                                LEFT JOIN QMS.QcClass c ON b.QcClassId = c.QcClassId
                                LEFT JOIN QMS.QcGroup d ON c.QcGroupId = d.QcGroupId
                                WHERE d.QcGroupId = @QcGroupId";
                        dynamicParameters.Add("QcGroupId", QcGroupId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        //QMS.QcItem
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM QMS.QcItem a
                                LEFT JOIN QMS.QcClass b ON a.QcClassId = b.QcClassId
                                LEFT JOIN QMS.QcGroup c ON b.QcGroupId = c.QcGroupId
                                WHERE c.QcGroupId = @QcGroupId";
                        dynamicParameters.Add("QcGroupId", QcGroupId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        //QMS.QcClass
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM  QMS.QcClass a
                                LEFT JOIN QMS.QcGroup b ON a.QcGroupId = b.QcGroupId
                                WHERE b.QcGroupId = @QcGroupId";
                        dynamicParameters.Add("QcGroupId", QcGroupId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        //QMS.QcGroup
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE FROM QMS.QcGroup
                                WHERE QcGroupId = @QcGroupId";
                        dynamicParameters.Add("QcGroupId", QcGroupId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteQcClass -- 刪除量測類別資料 -- Ted 2022.10.03
        public string DeleteQcClass(int QcClassId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷異常類別資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcClass
                                WHERE QcClassId = @QcClassId";
                        dynamicParameters.Add("QcClassId", QcClassId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("維修類別資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除關聯table
                        //QMS.QiDetail
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM QMS.QiDetail a
                                LEFT JOIN QMS.QcItem b ON a.QcItemId = b.QcItemId
                                LEFT JOIN QMS.QcClass c ON b.QcClassId = c.QcClassId
                                WHERE c.QcClassId = @QcClassId";
                        dynamicParameters.Add("QcClassId", QcClassId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        //QMS.QcItem
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a 
                                FROM QMS.QcItem a
                                LEFT JOIN QMS.QcClass b ON a.QcClassId = b.QcClassId
                                WHERE b.QcClassId = @QcClassId";
                        dynamicParameters.Add("QcClassId", QcClassId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        //QMS.QcClass
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE QMS.QcClass
                                WHERE QcClassId = @QcClassId";
                        dynamicParameters.Add("QcClassId", QcClassId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteQcItem -- 刪除量測項目資料 -- Ted 2022.10.03
        public string DeleteQcItem(int QcItemId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷量測項目資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcItem
                                WHERE QcItemId = @QcItemId";
                        dynamicParameters.Add("QcItemId", QcItemId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("量測項目資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除次要table
                        //QMS.QiDetail
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM QMS.QiDetail a
                                LEFT JOIN QMS.QcItem b ON a.QcItemId = b.QcItemId
                                WHERE b.QcItemId = @QcItemId";
                        dynamicParameters.Add("QcItemId", QcItemId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        //QMS.QcItem
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE QMS.QcItem
                                WHERE QcItemId = @QcItemId";
                        dynamicParameters.Add("QcItemId", QcItemId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteQiDetail -- 刪除量測項目機型單身 -- Ted 2022.09.27
        public string DeleteQiDetail(int QiDetailId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷裝置資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QiDetail
                                WHERE QiDetailId = @QiDetailId";
                        dynamicParameters.Add("QiDetailId", QiDetailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("機型資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE QMS.QiDetail
                                WHERE QiDetailId = @QiDetailId";
                        dynamicParameters.Add("QiDetailId", QiDetailId);

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

        #region //DeleteQiMachineMode -- 刪除量測機型 -- Ted 2022.09.27
        public string DeleteQiMachineMode(int QcMachineModeId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷機型資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcMachineMode
                                WHERE CompanyId = @CompanyId
                                AND QcMachineModeId = @QcMachineModeId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("QcMachineModeId", QcMachineModeId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("機型資料錯誤!");
                        #endregion

                        #region //判斷機型是否被使用
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QiDetail
                                WHERE QcMachineModeId = @QcMachineModeId";
                        dynamicParameters.Add("QcMachineModeId", QcMachineModeId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("該機型目前有使用，無法刪除!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除次要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE QMS.QmmDetail
                                WHERE QcMachineModeId = @QcMachineModeId";
                        dynamicParameters.Add("QcMachineModeId", QcMachineModeId);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE QMS.QcMachineMode
                                WHERE CompanyId = @CompanyId
                                AND QcMachineModeId = @QcMachineModeId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("QcMachineModeId", QcMachineModeId);

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

        #region //DeleteQmmDetail -- 刪除量測機型的機台單身 -- Ted 2022.09.27
        public string DeleteQmmDetail(int QmmDetailId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷機台單身資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QmmDetail
                                WHERE QmmDetailId = @QmmDetailId";
                        dynamicParameters.Add("QmmDetailId", QmmDetailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("機台單身資料錯誤!");
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE QMS.QmmDetail
                                WHERE QmmDetailId = @QmmDetailId";
                        dynamicParameters.Add("QmmDetailId", QmmDetailId);

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

        #region //DeleteQcItemPrinciple -- 刪除量測項目編碼原則 -- Shintokuro 2024.05.15
        public string DeleteQcItemPrinciple(int PrincipleId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷量測項目編碼原則是否存在
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcItemPrinciple
                                WHERE PrincipleId = @PrincipleId";
                        dynamicParameters.Add("PrincipleId", PrincipleId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("量測項目編碼原則不存在!");
                        #endregion

                        int rowsAffected = 0;
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE QcItemPrinciple
                                WHERE PrincipleId = @PrincipleId";
                        dynamicParameters.Add("PrincipleId", PrincipleId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

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

        #region //DeleteQcItemCoding -- 刪除項目編碼規則管理 -- Shintokuro 2024.05.15
        public string DeleteQcItemCoding(int QicId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷量測項目編碼原則是否存在
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcItemCoding
                                WHERE QicId = @QicId";
                        dynamicParameters.Add("QicId", QicId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("項目編碼規則不存在!");
                        #endregion

                        int rowsAffected = 0;
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE QMS.QcItemCoding
                                WHERE QicId = @QicId";
                        dynamicParameters.Add("QicId", QicId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

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

        #region //DeleteQcMeasureInTheoryWorkTime -- 刪除量測項目預測工時 -- GPAI 2024-12-17
        public string DeleteQcMeasureInTheoryWorkTime(int QmwtId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    using (TransactionScope transactionScope = new TransactionScope())
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷量測項目預測工時是否存在
                        sql = @"SELECT TOP 1 1
                                FROM QMS.QcMeasureInTheoryWorkTime
                                WHERE QmwtId = @QmwtId";
                        dynamicParameters.Add("QmwtId", QmwtId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("量測項目編碼原則不存在!");
                        #endregion

                        int rowsAffected = 0;
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE QMS.QcMeasureInTheoryWorkTime
                                WHERE QmwtId = @QmwtId";
                        dynamicParameters.Add("QmwtId", QmwtId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

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
