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
using System.Transactions;
using System.Web;

namespace QMSDA
{
    public class CustomerComplaintDA
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

        public CustomerComplaintDA()
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
        #region //GetCCMenuCategory --取得客訴選單類別 --GPai 20230912
        public string GetCCMenuCategory(int McId, string McType, string McNo, string McName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.McId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.CompanyId, a.McType, a.McNo, a.McName, a.McDesc, a.[Status]
                          ";
                    sqlQuery.mainTables =
                        @"FROM QMS.MenuCategory a";
                    sqlQuery.auxTables = "";

                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "McId", @" AND a.McId = @McId", McId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "McType", @" AND a.McType = @McType", McType);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "McNo", @" AND a.McNo = @McNo", McNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "McName", @" AND a.McName = @McName", McName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));


                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.McId ASC";
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

        #region //GetMenuCategorySimple -- 取得客訴選單類別(下拉用) -- Shintokuro 2023.07.10
        public string GetMenuCategorySimple(int McId, string McType)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    #region //取得品DFM單身
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.McId, a.McNo, a.McName
                            FROM QMS.MenuCategory a
                            WHERE a.CompanyId = @CompanyId
                            AND a.Status = 'A'
                            ";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "McType", @" AND a.McType = @McType", McType);

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

        #region //GetCustomerComplaint -- 取得客訴單據資料 -- Shintokuro 2023-09-12
        public string GetCustomerComplaint(int CcId, string CcNo, string MtlItemNo, int CustomerId, string CurrentStatus, string StartDate, string EndDate, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.CcId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.CcNo, a.MtlItemId,a.StageDataSetting
                          , FORMAT(a.FilingDate, 'yyyy-MM-dd') FilingDate
                          , FORMAT(a.DocDate, 'yyyy-MM-dd') DocDate
                          , FORMAT(a.ReplyDate, 'yyyy-MM-dd') ReplyDate
                          , FORMAT(a.ClosureDate, 'yyyy-MM-dd') ClosureDate
                          , a.CustomerId, a.UserId, a.UserName, a.UserEMail, a.IssueDescription
                          , a.ConfirmStatus, a.ConfirmDate, a.CurrentStatus, a.ChangeCustomerCcNo
                          , a.D1ShowSetting, a.D2ShowSetting, a.D3ShowSetting, a.D4ShowSetting
                          , a.D5ShowSetting, a.D6ShowSetting, a.D7ShowSetting, a.D8ShowSetting
                          , b.MtlItemNo, b.MtlItemName, b.MtlItemSpec
                          , c.CustomerNo, c.CustomerShortName
                          ";
                    sqlQuery.mainTables =
                        @"FROM QMS.CustomerComplaint a
                          INNER JOIN PDM.MtlItem b on a.MtlItemId = b.MtlItemId
                          INNER JOIN SCM.Customer c on a.CustomerId = c.CustomerId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CcId", @" AND a.CcId = @CcId", CcId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CcNo", @" AND a.CcNo = @CcNo", CcNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemNo", @" AND a.MtlItemNo LIKE '%' + @MtlItemNo +'%'", MtlItemNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerId", @" AND a.CustomerId = @CustomerId", CustomerId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CurrentStatus", @" AND a.CurrentStatus = @CurrentStatus", CurrentStatus);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.DocDate >= @StartDate ", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.DocDate <= @EndDate ", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ConfirmStatus", @" AND a.ConfirmStatus IN @ConfirmStatus", Status.Split(','));

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.DocDate DESC,a.CcId DESC";
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

        #region //GetCcPdfShowSetting -- 取得客訴單據PDF顯示設定 -- Shintokuro 2023.07.12
        public string GetCcPdfShowSetting(int CcId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    #region //取得品DFM單身
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT CcId,Subject,Score ,StageDataSetting
                            FROM QMS.CustomerComplaint
                            UNPIVOT
                            (
                              Score FOR Subject IN (D1ShowSetting,D2ShowSetting, D3ShowSetting, D4ShowSetting, D5ShowSetting, D6ShowSetting, D7ShowSetting, D8ShowSetting)
                            ) AS UnpivotedData
                            WHERE CcId = @CcId
                            ";
                    dynamicParameters.Add("CcId", CcId);
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

        #region //GetCcStageData -- 取得客訴單據各階段資料 -- Shintokuro 2023-09-08
        public string GetCcStageData(int CcStageDataId, int CcId, string Stage, int McId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.CcStageDataId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.CcId, a.Stage, a.McId, a.DescriptionTextarea, a.DeptId, a.UserId, a.ConfirmStatus,FORMAT(a.ConfirmDate, 'yyyy-MM-dd') ConfirmDate, FORMAT(a.ExpectedDate, 'yyyy-MM-dd') ExpectedDate, FORMAT(a.FinishDate, 'yyyy-MM-dd') FinishDate
                          , b.McName
                          , c.UserNo, c.UserName
                          , d.DepartmentNo, d.DepartmentName
                          ";
                    sqlQuery.mainTables =
                        @"FROM QMS.CcStageData a
                          INNER JOIN QMS.MenuCategory b on a.McId = b.McId
                          INNER JOIN BAS.[User] c on a.UserId = c.UserId
                          INNER JOIN BAS.Department d on a.DeptId = d.DepartmentId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @" AND a.Stage = @Stage";
                    dynamicParameters.Add("Stage", Stage);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CcId", @" AND a.CcId = @CcId", CcId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CcStageDataId", @" AND a.CcStageDataId = @CcStageDataId", CcStageDataId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "McId", @" AND a.McId = @McId", McId);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.CcStageDataId ASC";
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

        #region //GetCcStageDataFile -- 取得客訴單據各階段檔案 -- Shintokuro 2023.09.12
        public string GetCcStageDataFile(int CcStageDataFileId, int CcStageDataId, int CcId, string Stage, string WhereFrom
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.CcStageDataFileId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",a.CcStageDataId, a.FileId
                           ,b.[FileName],b.FileContent,b.FileExtension 
                          ";
                    sqlQuery.mainTables =
                        @"FROM QMS.CcStageDataFile a
                          INNER JOIN BAS.[File] b on a.FileId = b.FileId
                          INNER JOIN QMS.CcStageData c on a.CcStageDataId = c.CcStageDataId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    if (WhereFrom != "List")
                    {
                        WhereFrom = " AND b.FileExtension in ('.Png','.png','.JPEG','.jpeg','.JPG','.jpg')";
                    }
                    else
                    {
                        WhereFrom = " AND 1=1";
                    }
                    string queryCondition = WhereFrom;
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CcStageDataId", @" AND a.CcStageDataId = @CcStageDataId", CcStageDataId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CcId", @" AND c.CcId = @CcId", CcId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Stage", @" AND c.Stage = @Stage", Stage);
                    

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "c.Stage ,a.FileId DESC";
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

        #region //GetCCMember --取得D1小組成員 --GPai 20230914 
        public string GetCCMember(int CcId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.CcId, a.UserId, a.ConfirmStatus
                            , b.UserNo, b.UserName, b.DepartmentId
                            , c.DepartmentNo, c.DepartmentName
                            FROM QMS.CcTeamMembers a
                            INNER JOIN BAS.[User] b on a.UserId = b.UserId
                            INNER JOIN BAS.Department c on b.DepartmentId = c.DepartmentId
                            AND a.CcId = @CcId
                            ORDER BY a.UserId DESC
                            ";
                    dynamicParameters.Add("CcId", CcId);

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

        #region //GetCCMainFile --取得客訴單頭檔案 --GPai 20230922
        public string GetCCMainFile(int CcId, string WhereFrom
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.CcFileId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.CcId, c.CcNo, a.FileType, a.FileId
                          , b.[FileName],b.FileContent,b.FileExtension 
                          ";
                    sqlQuery.mainTables =
                        @"FROM QMS.CcFile a
                          INNER JOIN BAS.[File] b on a.FileId = b.FileId
                          INNER JOIN QMS.CustomerComplaint c on a.CcId = c.CcId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    if(WhereFrom != "List")
                    {
                        WhereFrom = "AND 1=1 AND b.FileExtension in ('.Png','.png','.JPEG','.jpeg','.JPG','.jpg')";
                    }
                    else
                    {
                        WhereFrom = " AND 1=1";
                    }
                    string queryCondition = WhereFrom;
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CcId", @" AND c.CcId = @CcId", CcId);

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

        #region //CustomerComplaintCardPdf -- 取得客訴單據資料(PDF) -- Shintokru 2023.09.14
        public string CustomerComplaintCardPdf(int CcId)
        {
            try
            {

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT 
                            a.CcNo,a.MtlItemId
                            ,D2.StageData StageDataD2
                            ,D3.StageData StageDataD3
                            ,D4.StageData StageDataD4
                            ,D5.StageData StageDataD5
                            ,D6.StageData StageDataD6
                            ,D7.StageData StageDataD7
                            ,D8.StageData StageDataD8
                            ,D8McName.McName D8McName
                            ,a.FilingDate
                            ,a.IssueDescription
                            ,ISNULL(FORMAT(a.FilingDate, 'yyyy/MM/dd'), '') FilingDate
                            ,ISNULL(FORMAT(a.DocDate, 'yyyy/MM/dd'), '') DocDate
                            ,a.D1ShowSetting,a.D2ShowSetting,a.D3ShowSetting,a.D4ShowSetting
                            ,a.D5ShowSetting,a.D6ShowSetting,a.D7ShowSetting,a.D8ShowSetting
                            ,a2.CcTeamMembers
                            ,b.MtlItemNo,b.MtlItemName,b.MtlItemSpec
                            ,c.CustomerName
                            FROM QMS.CustomerComplaint a
                            LEFT JOIN QMS.CcFile a1 on a.CcId = a1.CcId
                            INNER JOIN PDM.MtlItem b on a.MtlItemId = b.MtlItemId
                            INNER JOIN SCM.Customer c on a.CustomerId = c.CustomerId
                            OUTER APPLY(
	                                SELECT STUFF((SELECT ','+ Convert(nvarchar(MAX),(x3.DepartmentName+'-'+x2.UserName)) 
	                                FROM QMS.CcTeamMembers x
                                    INNER JOIN BAS.[User] x2 on x.UserId = x2.UserId
                                     INNER JOIN BAS.Department x3 on x2.DepartmentId = x3.DepartmentId
	                                WHERE a.CcId = x.CcId
                                    FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 1, '') AS CcTeamMembers
                                ) a2
                            OUTER APPLY(
	                            SELECT STUFF((SELECT '(^o^)'+ Convert(nvarchar(MAX),x4.McName+'(ToT)'+d.DescriptionTextarea+'(ToT)'+ x3.DepartmentName +'(ToT)'+ x2.UserName +'(ToT)'+ ISNULL(FORMAT(d.ConfirmDate, 'yyyy/MM/dd'), '')) 
                                FROM QMS.CcStageData d 
                                INNER JOIN BAS.[User] x2 on d.UserId = x2.UserId
                                INNER JOIN BAS.[Department] x3 on d.DeptId = x3.DepartmentId
                                INNER JOIN QMS.MenuCategory x4 on d.McId = x4.McId
                                WHERE a.CcId = d.CcId AND d.Stage = 'D2'
                                FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 5, '') AS StageData
                            ) D2
                            OUTER APPLY(
	                            SELECT STUFF((SELECT '(^o^)'+ Convert(nvarchar(MAX),x4.McName+'(ToT)'+d.DescriptionTextarea+'(ToT)'+ x3.DepartmentName +'(ToT)'+ x2.UserName +'(ToT)'+ ISNULL(FORMAT(d.ConfirmDate, 'yyyy/MM/dd'), '')) 
                                FROM QMS.CcStageData d 
                                INNER JOIN BAS.[User] x2 on d.UserId = x2.UserId
                                INNER JOIN BAS.[Department] x3 on d.DeptId = x3.DepartmentId
                                INNER JOIN QMS.MenuCategory x4 on d.McId = x4.McId
                                WHERE a.CcId = d.CcId AND d.Stage = 'D3'
                                FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 5, '') AS StageData
                            ) D3
                            OUTER APPLY(
	                            SELECT STUFF((SELECT '(^o^)'+ Convert(nvarchar(MAX),x4.McName+'(ToT)'+d.DescriptionTextarea+'(ToT)'+ x3.DepartmentName +'(ToT)'+ x2.UserName +'(ToT)'+ ISNULL(FORMAT(d.ConfirmDate, 'yyyy/MM/dd'), '')) 
                                FROM QMS.CcStageData d 
                                INNER JOIN BAS.[User] x2 on d.UserId = x2.UserId
                                INNER JOIN BAS.[Department] x3 on d.DeptId = x3.DepartmentId
                                INNER JOIN QMS.MenuCategory x4 on d.McId = x4.McId
                                WHERE a.CcId = d.CcId AND d.Stage = 'D4'
                                FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 5, '') AS StageData
                            ) D4
                            OUTER APPLY(
	                            SELECT STUFF((SELECT '(^o^)'+ Convert(nvarchar(MAX),x4.McName+'(ToT)'+d.DescriptionTextarea+'(ToT)'+ x3.DepartmentName +'(ToT)'+ x2.UserName +'(ToT)'+ ISNULL(FORMAT(d.ExpectedDate, 'yyyy/MM/dd'), '')) 
                                FROM QMS.CcStageData d 
                                INNER JOIN BAS.[User] x2 on d.UserId = x2.UserId
                                INNER JOIN BAS.[Department] x3 on d.DeptId = x3.DepartmentId
                                INNER JOIN QMS.MenuCategory x4 on d.McId = x4.McId
                                WHERE a.CcId = d.CcId AND d.Stage = 'D5'
                                FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 5, '') AS StageData
                            ) D5
                            OUTER APPLY(
	                            SELECT STUFF((SELECT '(^o^)'+ Convert(nvarchar(MAX),x4.McName+'(ToT)'+d.DescriptionTextarea+'(ToT)'+ x3.DepartmentName +'(ToT)'+ x2.UserName +'(ToT)'+ ISNULL(FORMAT(d.FinishDate, 'yyyy/MM/dd'), '')) 
                                FROM QMS.CcStageData d 
                                INNER JOIN BAS.[User] x2 on d.UserId = x2.UserId
                                INNER JOIN BAS.[Department] x3 on d.DeptId = x3.DepartmentId
                                INNER JOIN QMS.MenuCategory x4 on d.McId = x4.McId
                                WHERE a.CcId = d.CcId AND d.Stage = 'D6'
                                FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 5, '') AS StageData
                            ) D6
                            OUTER APPLY(
	                            SELECT STUFF((SELECT '(^o^)'+ Convert(nvarchar(MAX),x4.McName+'(ToT)'+d.DescriptionTextarea+'(ToT)'+ x3.DepartmentName +'(ToT)'+ x2.UserName +'(ToT)'+ ISNULL(FORMAT(d.ConfirmDate, 'yyyy/MM/dd'), '')) 
                                FROM QMS.CcStageData d 
                                INNER JOIN BAS.[User] x2 on d.UserId = x2.UserId
                                INNER JOIN BAS.[Department] x3 on d.DeptId = x3.DepartmentId
                                INNER JOIN QMS.MenuCategory x4 on d.McId = x4.McId
                                WHERE a.CcId = d.CcId AND d.Stage = 'D7'
                                FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 5, '') AS StageData
                            ) D7
                            OUTER APPLY(
	                            SELECT STUFF((SELECT '(^o^)'+ Convert(nvarchar(MAX),x4.McName+'(ToT)'+d.DescriptionTextarea+'(ToT)'+ x3.DepartmentName +'(ToT)'+ x2.UserName +'(ToT)'+ ISNULL(FORMAT(d.ConfirmDate, 'yyyy/MM/dd'), '')) 
                                FROM QMS.CcStageData d 
                                INNER JOIN BAS.[User] x2 on d.UserId = x2.UserId
                                INNER JOIN BAS.[Department] x3 on d.DeptId = x3.DepartmentId
                                INNER JOIN QMS.MenuCategory x4 on d.McId = x4.McId
                                WHERE a.CcId = d.CcId AND d.Stage = 'D8'
                                FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 5, '') AS StageData
                            ) D8
                            OUTER APPLY(
                                SELECT x4.McName
                                FROM QMS.CcStageData d
                                INNER JOIN QMS.MenuCategory x4 on d.McId = x4.McId
                                WHERE a.CcId = d.CcId AND d.Stage = 'D8'
                            ) D8McName
                            WHERE a.CcId = @CcId";
                    dynamicParameters.Add("@CcId", CcId);
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

        #region //GetDataConfirmStatus -- 各階段確認狀態(PDF) -- Shintokru 2023.09.14
        public string GetDataConfirmStatus(int CcId)
        {
            try
            {

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT Stage,ConfirmStatus 
							FROM QMS.CcStageData 
							WHERE CcId = @CcId 
							GROUP BY Stage,ConfirmStatus
							UNION 
							SELECT 'D1' Stage, ConfirmStatus 
							FROM QMS.CcTeamMembers
							WHERE CcId = @CcId 
							GROUP BY ConfirmStatus";
                    dynamicParameters.Add("@CcId", CcId);
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

        #region //GetCcStageDataFilePDf -- 取得客訴單據各階段檔案 -- Shintokuro 2023.09.12
        public string GetCcStageDataFilePDf(int CcId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sql = @"SELECT b1.[FileName],b1.FileContent,b1.FileExtension,0 Stage, b.CcFileId FileSore
                            FROM QMS.CustomerComplaint a
                            INNEr JOIN QMS.CcFile b on a.CcId = b.CcId
                            INNER JOIN BAS.[File] b1 on b.FileId = b1.FileId
                            WHERE a.CcId = @CcId
                            AND b1.FileExtension in ('.Png','.png','.JPEG','.jpeg','.JPG','.jpg')
                            UNION
                            SELECT c2.[FileName],c2.FileContent,c2.FileExtension,SUBSTRING(c.Stage, 2, LEN(c.Stage) - 1) ,c1.CcStageDataFileId FileSore
                            FROM QMS.CustomerComplaint a
                            INNER JOIN QMS.CcStageData c on a.CcId = c.CcId
                            INNER JOIN QMS.CcStageDataFile c1 on c.CcStageDataId = c1.CcStageDataId
                            INNER JOIN BAS.[File] c2 on c1.FileId = c2.FileId
                            WHERE a.CcId = @CcId
                            AND c2.FileExtension in ('.Png','.png','.JPEG','.jpeg','.JPG','.jpg')
                            ORDER BY Stage ,FileSore";
                    dynamicParameters.Add("@CcId", CcId);

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

        #region //ShowCustomerComplaint --客訴單檢視 GPai 20230920
        public string ShowCustomerComplaint(int CcId)
        {
            try
            {
                List<CustomerComplaintForm> customerComplaintForm = new List<CustomerComplaintForm>();
                List<CCMember> cCMember = new List<CCMember>();
                List<StageDataDetail> stageDataDetail = new List<StageDataDetail>();
                List<StageDataFile> stageDataFile = new List<StageDataFile>();
                List<CCMainFile> mainFile = new List<CCMainFile>();
                List<CCTrackingData> trackingData = new List<CCTrackingData>();



                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    #region //單頭
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.CcId, a.CcNo, a.MtlItemId
                          , FORMAT(a.FilingDate, 'yyyy-MM-dd') FilingDate
                          , FORMAT(a.DocDate, 'yyyy-MM-dd') DocDate
                          , FORMAT(a.ReplyDate, 'yyyy-MM-dd') ReplyDate
                          , FORMAT(a.ClosureDate, 'yyyy-MM-dd') ClosureDate
                          , a.CustomerId, a.UserId, a.UserName, a.UserEMail, a.IssueDescription
                          , a.ConfirmStatus, a.ConfirmDate, a.CurrentStatus,StageDataSetting
                          , b.MtlItemNo, b.MtlItemName, b.MtlItemSpec
                          , c.CustomerNo, c.CustomerShortName
	                      FROM QMS.CustomerComplaint a
                          INNER JOIN PDM.MtlItem b on a.MtlItemId = b.MtlItemId
                          INNER JOIN SCM.Customer c on a.CustomerId = c.CustomerId
	                      WHERE a.CcId = @CcId AND a.CompanyId = @CompanyId
                            ";
                    dynamicParameters.Add("CcId", CcId);
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    customerComplaintForm = sqlConnection.Query<CustomerComplaintForm>(sql, dynamicParameters).ToList();
                    #endregion

                    #region //單頭檔案
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT c.CcId, a.CcFileId, a.FileType, a.FileId
                           ,b.[FileName]
	                      FROM QMS.CcFile a
                          INNER JOIN BAS.[File] b on a.FileId = b.FileId
                          INNER JOIN QMS.CustomerComplaint c on a.CcId = c.CcId
	                      WHERE c.CcId = @CcId
                          AND b.FileExtension in ('.Png','.png','.JPEG','.jpeg','.JPG','.jpg')
                            ";
                    dynamicParameters.Add("CcId", CcId);

                    mainFile = sqlConnection.Query<CCMainFile>(sql, dynamicParameters).ToList();

                    customerComplaintForm
                        .ToList()
                        .ForEach(x =>
                        {
                            x.CCMainFile = mainFile
                                                    .Where(y => y.CcId == x.CcId)
                                                    .ToList();
                        });
                    #endregion

                    #region //成員(D1)
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.CcId, a.CcTeamMembersId, a.UserId,
                            b.UserNo, b.UserName, b.DepartmentId
                            , c.DepartmentNo, c.DepartmentName
                            FROM QMS.CcTeamMembers a
                            INNER JOIN BAS.[User] b on a.UserId = b.UserId
                            INNER JOIN BAS.Department c on b.DepartmentId = c.DepartmentId
                            AND a.CcId = @CcId
                            ORDER BY a.UserId DESC
                            ";
                    dynamicParameters.Add("CcId", CcId);

                    cCMember = sqlConnection.Query<CCMember>(sql, dynamicParameters).ToList();

                    customerComplaintForm
                        .ToList()
                        .ForEach(x =>
                        {
                            x.D1 = cCMember
                                                    .Where(y => y.CcId == x.CcId)
                                                    .ToList();
                        });
                    #endregion

                    #region //D8附表
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.CCTrackingDataId, a.CcStageDataId, a.Batch, a.ReportNo, a.TrialResult, FORMAT(a.TrackingDate, 'yyyy-MM-dd') TrackingDate, b.CcId
                            FROM QMS.CCTrackingData a
                            INNER JOIN QMS.CcStageData b on a.CcStageDataId = b.CcStageDataId
                            INNER JOIN QMS.MenuCategory c on b.McId = c.McId
                            WHERE b.CcId = @CcId
                            ";
                    dynamicParameters.Add("CcId", CcId);

                    trackingData = sqlConnection.Query<CCTrackingData>(sql, dynamicParameters).ToList();

                    customerComplaintForm
                        .ToList()
                        .ForEach(x =>
                        {
                            x.TrackingData = trackingData
                                                    .Where(y => y.CcId == x.CcId)
                                                    .ToList();
                        });
                    #endregion



                    #region //階段資料(D2~D8)
                    for (int i = 2; i <= 8; i++)
                    {
                        string statge = "D" + i;
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CcStageDataId, a.CcId, a.Stage, a.McId, a.DescriptionTextarea, a.DeptId, a.UserId, a.ConfirmStatus,FORMAT(a.ConfirmDate, 'yyyy-MM-dd') ConfirmDate, FORMAT(a.ExpectedDate, 'yyyy-MM-dd') ExpectedDate, FORMAT(a.FinishDate, 'yyyy-MM-dd') FinishDate
	                        , b.McName
	                        , c.UserNo, c.UserName
	                        , d.DepartmentNo, d.DepartmentName
	                        FROM QMS.CcStageData a
                            INNER JOIN QMS.MenuCategory b on a.McId = b.McId
                            INNER JOIN BAS.[User] c on a.UserId = c.UserId
                            INNER JOIN BAS.Department d on a.DeptId = d.DepartmentId
	                        WHERE a.CcId = @CcId AND a.Stage = @Stage";
                        dynamicParameters.Add("CcId", CcId);
                        dynamicParameters.Add("Stage", statge);

                        stageDataDetail = sqlConnection.Query<StageDataDetail>(sql, dynamicParameters).ToList();

                        #region //檔案
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT c.CcId, a.CcStageDataFileId, a.CcStageDataId, a.FileId, c.Stage
                           ,b.[FileName]
	                      FROM QMS.CcStageDataFile a
                          INNER JOIN BAS.[File] b on a.FileId = b.FileId
                          INNER JOIN QMS.CcStageData c on a.CcStageDataId = c.CcStageDataId
	                      WHERE c.CcId = @CcId AND c.Stage = @Stage AND b.FileExtension in ('.Png','.png','.JPEG','.jpeg','.JPG','.jpg')
                            ";
                        dynamicParameters.Add("CcId", CcId);
                        dynamicParameters.Add("Stage", statge);

                        stageDataFile = sqlConnection.Query<StageDataFile>(sql, dynamicParameters).ToList();

                        stageDataDetail
                            .ToList()
                            .ForEach(x =>
                            {
                                x.DataFile = stageDataFile
                                                        .Where(y => y.Stage == x.Stage)
                                                        .Where(y => y.CcStageDataId == x.CcStageDataId)
                                                        .ToList();
                            });

                        #endregion


                        #region //DATA
                        if (stageDataDetail.Count() > 0)
                        {
                            switch (statge)
                            {
                                case "D2":
                                    customerComplaintForm
                            .ToList()
                            .ForEach(x =>
                            {
                                x.D2 = stageDataDetail
                                                        .Where(y => y.CcId == x.CcId)
                                                        .ToList();
                            });
                                    break;
                                case "D3":
                                    customerComplaintForm
                            .ToList()
                            .ForEach(x =>
                            {
                                x.D3 = stageDataDetail
                                                        .Where(y => y.CcId == x.CcId)
                                                        .ToList();
                            });
                                    break;
                                case "D4":
                                    customerComplaintForm
                            .ToList()
                            .ForEach(x =>
                            {
                                x.D4 = stageDataDetail
                                                        .Where(y => y.CcId == x.CcId)
                                                        .ToList();
                            });
                                    break;
                                case "D5":
                                    customerComplaintForm
                            .ToList()
                            .ForEach(x =>
                            {
                                x.D5 = stageDataDetail
                                                        .Where(y => y.CcId == x.CcId)
                                                        .ToList();
                            });
                                    break;
                                case "D6":
                                    customerComplaintForm
                            .ToList()
                            .ForEach(x =>
                            {
                                x.D6 = stageDataDetail
                                                        .Where(y => y.CcId == x.CcId)
                                                        .ToList();
                            });
                                    break;
                                case "D7":
                                    customerComplaintForm
                            .ToList()
                            .ForEach(x =>
                            {
                                x.D7 = stageDataDetail
                                                        .Where(y => y.CcId == x.CcId)
                                                        .ToList();
                            });
                                    break;
                                case "D8":
                                    customerComplaintForm
                            .ToList()
                            .ForEach(x =>
                            {
                                x.D8 = stageDataDetail
                                                        .Where(y => y.CcId == x.CcId)
                                                        .ToList();
                            });
                                    break;
                            }
                        }
                        #endregion

                    }
                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = customerComplaintForm
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

        #region //GetD8TableData --取得D8附表 --GPai 20230925
        public string GetD8TableData(int CcStageDataId, int CCTrackingDataId, int CcId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.CCTrackingDataId, a.CcStageDataId, a.Batch, a.ReportNo, a.TrialResult, FORMAT(a.TrackingDate, 'yyyy-MM-dd') TrackingDate
                            FROM QMS.CCTrackingData a
                            INNER JOIN QMS.CcStageData b on a.CcStageDataId = b.CcStageDataId
                            INNER JOIN QMS.MenuCategory c on b.McId = c.McId
                            WHERE 1=1
                            ";
                    dynamicParameters.Add("CcStageDataId", CcStageDataId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CcStageDataId", @" AND a.CcStageDataId = @CcStageDataId", CcStageDataId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CCTrackingDataId", @" AND a.CCTrackingDataId = @CCTrackingDataId", CCTrackingDataId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CcId", @" AND b.CcId = @CcId", CcId);

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

        #endregion

        #region //Add

        #region AddCCMenuCategory --新增選單類別 --GPai 20230912
        public string AddCCMenuCategory(string McName, string McNo, string McType, string McDesc)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        DynamicParameters dynamicParameters = new DynamicParameters();
                        #region //是否在該類別有重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.McId, a.CompanyId, a.McType, a.McNo, a.McName, a.McDesc, a.[Status]
                            FROM QMS.MenuCategory a
                            WHERE a.McNo = @McNo AND a.CompanyId = @CompanyId
                            ";
                        dynamicParameters.Add("McNo", McNo);
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("資料重複，請確認資料");
                        #endregion


                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO QMS.MenuCategory (CompanyId, McType, McNo, McName, McDesc,
                                CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.McId, INSERTED.McName
                                VALUES (@CompanyId, @McType, @McNo, @McName, @McDesc,
                                @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                McType,
                                McNo,
                                McName,
                                McDesc,
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

        #region //AddACustomerComplaint -- 新增客訴單據資料 -- Shintokuro 2023-09-12
        public string AddACustomerComplaint(int MtlItemId, string FilingDate, string DocDate, int CustomerId
            , string UserName, string UserEMail, string IssueDescription)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        DynamicParameters dynamicParameters = new DynamicParameters();

                        string CustomerNo = "";
                        string CcNo = "";
                        string Year = (CreateDate.Year % 100).ToString();
                        #region //判斷品號資料是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 MtlItemId
                                FROM PDM.MtlItem
                                WHERE MtlItemId = @MtlItemId
                                AND TransferStatus = 'Y'";
                        dynamicParameters.Add("MtlItemId", MtlItemId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("品號資料錯誤或不存在!");
                        #endregion

                        #region //判斷客戶資料是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 CustomerNo
                                FROM SCM.Customer
                                WHERE CustomerId = @CustomerId";
                        dynamicParameters.Add("CustomerId", CustomerId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("客戶資料錯誤或不存在!");
                        foreach (var item in result)
                        {
                            CustomerNo = item.CustomerNo;
                        }
                        #endregion

                        #region //判斷單據編號是否存在(撈取最大值)
                        CcNo = CustomerNo + Year + "__";
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 RIGHT(CcNo, 2) AS MaxSort
                                FROM QMS.CustomerComplaint
                                WHERE CcNo LIKE  @CcNo
                                ORDER BY CcId DESC";
                        dynamicParameters.Add("CcNo", CcNo);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0)
                        {
                            string serialNum = "";
                            foreach (var item in result)
                            {
                                serialNum = String.Format("{0:D2}", Convert.ToInt32(item.MaxSort) + 1);

                                CcNo = CustomerNo + Year + serialNum;
                            }
                        }
                        else
                        {
                            CcNo = CustomerNo + Year + "01";
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO QMS.CustomerComplaint (CompanyId, CcNo, MtlItemId, FilingDate, DocDate,
                                CustomerId, UserId, UserName, UserEMail, IssueDescription,
                                D1ShowSetting, D2ShowSetting, D3ShowSetting, D4ShowSetting, D5ShowSetting, D6ShowSetting, D7ShowSetting, D8ShowSetting,
                                StageDataSetting, ConfirmStatus, CurrentStatus,
                                CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.CcId, INSERTED.CcNo
                                VALUES (@CompanyId, @CcNo, @MtlItemId, @FilingDate, @DocDate,
                                @CustomerId, @UserId, @UserName, @UserEMail, @IssueDescription,
                                @D1ShowSetting, @D2ShowSetting, @D3ShowSetting, @D4ShowSetting, @D5ShowSetting, @D6ShowSetting, @D7ShowSetting, @D8ShowSetting,
                                @StageDataSetting, @ConfirmStatus, @CurrentStatus,
                                @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                CcNo,
                                MtlItemId,
                                FilingDate,
                                DocDate,
                                CustomerId,
                                UserId = (int?)null,
                                UserName,
                                UserEMail,
                                IssueDescription,
                                D1ShowSetting = "A",
                                D2ShowSetting = "A",
                                D3ShowSetting = "A",
                                D4ShowSetting = "A",
                                D5ShowSetting = "A",
                                D6ShowSetting = "A",
                                D7ShowSetting = "A",
                                D8ShowSetting = "A",
                                StageDataSetting = "YYYYYYYY",
                                ConfirmStatus = "Y",
                                CurrentStatus = "D1",
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

        #region //AddCcStageData -- 新增客訴單據各階段資料 -- Shintokuro 2023-09-08
        public string AddCcStageData(int CcId, string Stage, int McId, int DeptId, int UserId, string DescriptionTextarea, string ExpectedDate, string FinishDate
            )
        {
            try
            {
                string expectedDate = "";
                string finishDate = "";

                if (Stage == "D5") {
                    expectedDate = ExpectedDate;
                    finishDate = null;
                } else if (Stage == "D6") {
                    expectedDate = null;
                    finishDate = FinishDate;
                } else {
                    expectedDate = null;
                    finishDate = null;
                }

                using (TransactionScope transactionScope = new TransactionScope())
                {

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷客訴單據是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT  CcId
                                FROM QMS.CustomerComplaint
                                WHERE CcId = @CcId";
                        dynamicParameters.Add("CcId", CcId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("找不到客訴單據，請重新確認");
                        string CurrentStatus = "";
                        foreach (var item in result)
                        {
                            //if (item.CurrentStatus != Stage) throw new SystemException("上一階段資料尚未維護完成,無法進行新增! 階段: " + item.CurrentStatus);
                        }
                        #endregion

                        #region //判斷客訴單據階段資料是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ConfirmStatus,b.McName
                                FROM QMS.CcStageData a
                                INNER JOIN QMS.MenuCategory b on a.McId = b.McId
                                WHERE CcId = @CcId AND Stage = @Stage ";
                        dynamicParameters.Add("CcId", CcId);
                        dynamicParameters.Add("Stage", Stage);
                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        foreach (var item in result2)
                        {
                            if (item.ConfirmStatus == "Y") throw new SystemException("該階段資料已經確認,無法進行更動!");

                            //if (Stage == "D7") {
                            //    if (item.McName != "無分類") throw new SystemException("該分類已新增! 請重新選取");
                            //}
                            //if (item.McId == McId) throw new SystemException("該分類已新增! 請重新選取");
                        }
                        #endregion

                        #region //判斷客訴選單類別是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT  a.McId
                                FROM QMS.MenuCategory a
                                WHERE McId = @McId";
                        dynamicParameters.Add("McId", McId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("找不到客訴選單類別資料，請重新確認");
                        #endregion

                        #region //判斷部門是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT  a.DepartmentId
                                FROM BAS.Department a
                                WHERE DepartmentId = @DeptId";
                        dynamicParameters.Add("DeptId", DeptId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("找不到部門資料，請重新確認");
                        #endregion

                        #region //判斷人員是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT  a.UserId
                                FROM BAS.[User] a
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", UserId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("找不到客訴人員資料，請重新確認");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO QMS.CcStageData (CcId, Stage, McId, DescriptionTextarea, DeptId, UserId, ConfirmStatus, ExpectedDate, FinishDate,
                                CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.CcStageDataId
                                VALUES (@CcId, @Stage, @McId, @DescriptionTextarea, @DeptId, @UserId, @ConfirmStatus, @ExpectedDate, @FinishDate,
                                @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CcId,
                                Stage,
                                McId,
                                DescriptionTextarea,
                                DeptId,
                                UserId,
                                ConfirmStatus = "N",
                                expectedDate,
                                finishDate,
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

        #region //AddCcStageDataFile-- 新增客訴單據各階段檔案資料 -- Shintokuro 2023-09-12
        public string AddCcStageDataFile(int CcStageDataId, string FileId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        int rowsAffected = 0;

                        #region //品異條碼原因佐證檔案建立

                        foreach (var item in FileId.Split(','))
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO QMS.CcStageDataFile (CcStageDataId, FileId
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.CcStageDataFileId
                                    VALUES (@CcStageDataId, @FileId
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    CcStageDataId,
                                    FileId = item,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();
                        }
                        #endregion

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

        #region //AddCCUser --新增客訴單成員 --GPai 20230915
        public string AddCCUser(int CcId, string UserId)
        {
            try
            {
                string[] UserIdList = UserId.Split(',');
                int rowsAffected = 0;

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();
                        foreach (var userId in UserIdList)
                        {

                            #region //判斷D1是否已確認
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.*,b.UserNo
                                FROM QMS.CcTeamMembers a
                                INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                                WHERE a.CcId = @CcId";
                            dynamicParameters.Add("CcId", CcId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            //if (result.Count() <= 0) throw new SystemException("找不到客訴單成員名單，請重新確認");
                            string CurrentStatus = "";
                            foreach (var item in result)
                            {
                                if (item.ConfirmStatus != "N") throw new SystemException("該階段資料已經確認,無法進行更動!");
                                if (item.UserId == Int64.Parse(userId)) throw new SystemException("該成員已新增! 請重新選取 " + item.UserNo);
                            }
                            #endregion


                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT * FROM BAS.[User] 
                                    WHERE  [Status] = 'A' 
                                    AND UserId = @UserId";
                            dynamicParameters.Add("UserId", userId);

                            var result0 = sqlConnection.Query(sql, dynamicParameters);
                            if (result0.Count() <= 0) throw new SystemException("使用者資料錯誤!" + userId);

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO QMS.CcTeamMembers (CcId, UserId,
                                CreateDate, LastModifiedDate, CreateBy, LastModifiedBy, ConfirmStatus)
                                OUTPUT INSERTED.CcId
                                VALUES (@CcId, @UserId,
                                @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy, @ConfirmStatus)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    CcId,
                                    UserId = Int64.Parse(userId),
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy,
                                    ConfirmStatus = "N"
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

        #region //AddCCMainFile --新增單頭檔案 --GPai 20230922
        public string AddCCMainFile(int CcId, string FileId, string FileType)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        int rowsAffected = 0;

                        if (FileId == "" ) throw new SystemException("請先點擊上傳以利檔案上傳!");

                        #region //檔案建立

                        foreach (var item in FileId.Split(','))
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO QMS.CcFile (CcId, FileId, FileType
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.CcFileId
                                    VALUES (@CcId, @FileId, @FileType
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    CcId,
                                    FileId = item,
                                    FileType,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();
                        }
                        #endregion

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

        #region //AddD8TableData --新增D8附表 --GPai 20230925
        public string AddD8TableData(int CcStageDataId, int Batch, string ReportNo, string TrialResult, string TrackingDate)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        DynamicParameters dynamicParameters = new DynamicParameters();


                        #region //判斷D8是否已確認
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT *
                                FROM QMS.CcStageData
                                WHERE CcStageDataId = @CcStageDataId";
                        dynamicParameters.Add("CcStageDataId", CcStageDataId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        //if (result.Count() <= 0) throw new SystemException("找不到客訴單成員名單，請重新確認");
                        string CurrentStatus = "";
                        foreach (var item in result)
                        {
                            if (item.ConfirmStatus != "N") throw new SystemException("該階段資料已經確認,無法進行更動!");
                        }
                        #endregion

                        #region //判斷D8附表資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT *
                                FROM QMS.CCTrackingData
                                WHERE CcStageDataId = @CcStageDataId";
                        dynamicParameters.Add("CcStageDataId", CcStageDataId);

                        var trackingDataresult = sqlConnection.Query(sql, dynamicParameters);
                        if (trackingDataresult.Count() >= 3) throw new SystemException("附表資料已達上限!");

                        #endregion


                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO QMS.CCTrackingData (CcStageDataId, Batch, ReportNo, TrialResult, TrackingDate,
                                CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.CcStageDataId
                                VALUES (@CcStageDataId, @Batch, @ReportNo, @TrialResult, @TrackingDate,
                                @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CcStageDataId,
                                Batch,
                                ReportNo,
                                TrialResult,
                                TrackingDate,
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

        #endregion

        #region //Update
        #region UpdateCCMenuCategory //編輯選單類別 --GPai 20230912
        public string UpdateCCMenuCategory(int McId, string McName, string McNo, string McType, string McDesc)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷資料是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT  a.McId
                                FROM QMS.MenuCategory a
                                WHERE McId = @McId";
                        dynamicParameters.Add("McId", McId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("找不到客訴選單類別資料，請重新確認");
                        #endregion

                        #region//更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE QMS.MenuCategory SET
                                McDesc = @McDesc,
                                McName = @McName,
                                McNo = @McNo,
                                McType = @McType,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE McId = @McId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                McDesc,
                                McName,
                                McNo,
                                McType,
                                LastModifiedDate,
                                LastModifiedBy,
                                McId
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

        #region //UpdateCustomerComplaint -- 更新客訴單據資料 -- Shintokuro 2023-09-12
        public string UpdateCustomerComplaint(int CcId, string CcNo, int MtlItemId, string FilingDate, string DocDate, int CustomerId
            , string UserName, string UserEMail, string IssueDescription)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷客訴單據是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CcId
                                FROM QMS.CustomerComplaint a
                                WHERE CcId = @CcId";
                        dynamicParameters.Add("CcId", CcId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("找不到客訴單據，請重新確認");
                        #endregion

                        #region //判斷品號資料是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 MtlItemId
                                FROM PDM.MtlItem
                                WHERE MtlItemId = @MtlItemId
                                AND TransferStatus = 'Y'";
                        dynamicParameters.Add("MtlItemId", MtlItemId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("品號資料錯誤或不存在!");
                        #endregion

                        #region//更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE QMS.CustomerComplaint SET
                                MtlItemId = @MtlItemId,
                                FilingDate = @FilingDate,
                                DocDate = @DocDate,
                                CustomerId = @CustomerId,
                                UserName = @UserName,
                                UserEMail = @UserEMail,
                                IssueDescription = @IssueDescription,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE CcId = @CcId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                MtlItemId,
                                FilingDate,
                                DocDate,
                                CustomerId,
                                UserName,
                                UserEMail,
                                IssueDescription,
                                LastModifiedDate,
                                LastModifiedBy,
                                CcId
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

        #region //UpdateChangeCustomer -- 更新客訴單據客戶代碼替換資料 -- Shintokuro 2023-12-20
        public string UpdateChangeCustomer(int CcId, int ChangeCustomerId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        DynamicParameters dynamicParameters = new DynamicParameters();

                        string CustomerNo = "";
                        string CcNo = "";
                        string Year = (CreateDate.Year % 100).ToString();
                        int newCcId = 0;

                        #region //判斷客訴單據是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CcId,a.CustomerId,a.ConfirmStatus
                                FROM QMS.CustomerComplaint a
                                WHERE CcId = @CcId";
                        dynamicParameters.Add("CcId", CcId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("找不到客訴單據，請重新確認");
                        foreach (var item in result)
                        {
                            if (item.CustomerId == ChangeCustomerId) throw new SystemException("欲替換客戶和原單據客戶不可以相同，請重新確認");
                            if (item.ConfirmStatus == "V") throw new SystemException("單據已作廢不能執行替換客戶功能，請重新確認");
                            if (item.ConfirmStatus != "Y") throw new SystemException("單據需處於啟用狀態才能執行替換客戶功能，請重新確認");
                        }
                        #endregion

                        #region //判斷客戶資料是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 CustomerNo
                                FROM SCM.Customer
                                WHERE CustomerId = @ChangeCustomerId";
                        dynamicParameters.Add("ChangeCustomerId", ChangeCustomerId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("客戶資料錯誤或不存在!");
                        foreach (var item in result)
                        {
                            CustomerNo = item.CustomerNo;
                        }
                        #endregion

                        #region //判斷單據編號是否存在(撈取最大值)
                        CcNo = CustomerNo + Year + "__";
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 RIGHT(CcNo, 2) AS MaxSort
                                FROM QMS.CustomerComplaint
                                WHERE CcNo LIKE  @CcNo
                                ORDER BY CcId DESC";
                        dynamicParameters.Add("CcNo", CcNo);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0)
                        {
                            string serialNum = "";
                            foreach (var item in result)
                            {
                                serialNum = String.Format("{0:D2}", Convert.ToInt32(item.MaxSort) + 1);

                                CcNo = CustomerNo + Year + serialNum;
                            }
                        }
                        else
                        {
                            CcNo = CustomerNo + Year + "01";
                        }
                        #endregion

                        #region //撈取更新者名稱+編號
                        string UserNo = "";
                        string UserName = "";
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 UserNo,UserName
                                FROM BAS.[User]
                                WHERE UserId =  @LastModifiedBy";
                        dynamicParameters.Add("LastModifiedBy", LastModifiedBy);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("人員資料錯誤!找不到更新者,請重新整理頁面確認");
                        if (result.Count() > 0)
                        {
                            string serialNum = "";
                            foreach (var item in result)
                            {
                                UserNo = item.UserNo;
                                UserName = item.UserName;
                            }
                        }
                        #endregion

                        #region //客訴單據新增
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO QMS.CustomerComplaint OUTPUT INSERTED.CcId
                                SELECT CompanyId, @CcNo, MtlItemId, FilingDate, DocDate, ReplyDate, ClosureDate, 
                                @CustomerId, UserId, UserName, UserEMail, IssueDescription, D1ShowSetting, D2ShowSetting, D3ShowSetting, 
                                D4ShowSetting, D5ShowSetting, D6ShowSetting, D7ShowSetting, D8ShowSetting,ChangeCustomerCcNo, ConfirmStatus, ConfirmDate, 
                                CurrentStatus, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy
                                FROM   QMS.CustomerComplaint
                                WHERE CcId = @CcId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CcNo,
                                CustomerId = ChangeCustomerId,
                                ConfirmStatus = 'Y',
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy,
                                CcId
                            });


                        result = sqlConnection.Query(sql, dynamicParameters);
                        string[] NewCcNo = { CcNo.ToString() };

                        int rowsAffected = result.Count();

                        foreach (var item in result)
                        {
                            newCcId = Convert.ToInt32(item.CcId);
                        }
                        #endregion

                        #region//原單據的CcFile 替換 成新單據
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE QMS.CcFile SET
                                CcId = @newCcId,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE CcId = @CcId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                newCcId,
                                LastModifiedDate,
                                LastModifiedBy,
                                CcId
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region//原單據的CcTeamMembers 替換 成新單據
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE QMS.CcTeamMembers SET
                                CcId = @newCcId,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE CcId = @CcId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                newCcId,
                                LastModifiedDate,
                                LastModifiedBy,
                                CcId
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region//原單據的CcStageData 替換 成新單據
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE QMS.CcStageData SET
                                CcId = @newCcId,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE CcId = @CcId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                newCcId,
                                LastModifiedDate,
                                LastModifiedBy,
                                CcId
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region//作廢原單據
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE QMS.CustomerComplaint SET
                                IssueDescription = @IssueDescription,
                                ChangeCustomerCcNo = @CcNo,
                                ConfirmStatus = 'V',
                                ConfirmDate = @ConfirmDate,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE CcId = @CcId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                IssueDescription = UserNo + "-" + UserName + ":已將該單據替換成正確客戶,新單據編號【" + CcNo + "】",
                                CcNo,
                                ConfirmDate = LastModifiedDate,
                                LastModifiedDate,
                                LastModifiedBy,
                                CcId
                            });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            data = NewCcNo

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

        #region //UpdateCcStageData -- 更新客訴單據各階段資料 -- Shintokuro 2023-09-08
        public string UpdateCcStageData(int CcStageDataId, int McId, int DeptId, int UserId, string DescriptionTextarea, string ExpectedDate, string FinishDate
            )
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {

                    string expectedDate = "";
                    string finishDate = "";

                    

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷客訴單據階段資料是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 *
                                FROM QMS.CcStageData a
                                WHERE CcStageDataId = @CcStageDataId";
                        dynamicParameters.Add("CcStageDataId", CcStageDataId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("找不到客訴單據，請重新確認");
                        string CurrentStatus = "";
                        foreach (var item in result)
                        {
                            if (item.ConfirmStatus == "Y") throw new SystemException("該階段資料已經確認,無法進行更動!");
                            if (item.Stage == "D5")
                            {
                                expectedDate = ExpectedDate;
                                finishDate = null;
                            }
                            else if (item.Stage == "D6")
                            {
                                expectedDate = null;
                                finishDate = FinishDate;
                            }
                            else
                            {
                                expectedDate = null;
                                finishDate = null;
                            }
                        }
                        #endregion

                        #region //判斷客訴選單類別是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT  a.McId
                                FROM QMS.MenuCategory a
                                WHERE McId = @McId";
                        dynamicParameters.Add("McId", McId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("找不到客訴選單類別資料，請重新確認");
                        #endregion

                        #region //判斷部門是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT  a.DepartmentId
                                FROM BAS.Department a
                                WHERE DepartmentId = @DeptId";
                        dynamicParameters.Add("DeptId", DeptId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("找不到部門資料，請重新確認");
                        #endregion

                        #region //判斷人員是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT  a.UserId
                                FROM BAS.[User] a
                                WHERE UserId = @UserId";
                        dynamicParameters.Add("UserId", UserId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("找不到客訴人員資料，請重新確認");
                        #endregion

                        #region//更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE QMS.CcStageData SET
                                McId = @McId,
                                DescriptionTextarea = @DescriptionTextarea,
                                DeptId = @DeptId,
                                UserId = @UserId,
                                ExpectedDate = @ExpectedDate,
                                FinishDate = @FinishDate,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE CcStageDataId = @CcStageDataId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                McId,
                                DescriptionTextarea,
                                DeptId,
                                UserId,
                                ExpectedDate = expectedDate,
                                FinishDate = finishDate,
                                LastModifiedDate,
                                LastModifiedBy,
                                CcStageDataId
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

        #region //UpdatePdfShowSetting -- 更新PDF顯示狀態設定 -- Shintokru 2023.09.12
        public string UpdatePdfShowSetting(int CcId, string Stage, string Model)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;
                        #region//檢核客訴單據是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 StageDataSetting,
                                D1ShowSetting, D2ShowSetting, D3ShowSetting, D4ShowSetting,
                                D5ShowSetting, D6ShowSetting, D7ShowSetting, D8ShowSetting
                                FROM QMS.CustomerComplaint a
                                WHERE a.CcId=@CcId";
                        dynamicParameters.Add("CcId", CcId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("客訴單據查無資料,請重新確認");
                        #endregion


                        switch (Model)
                        {
                            case "ONE":
                                #region //調整為相反狀態
                                string ChangeStatus = "";
                                string status = "";
                                foreach (var item in result)
                                {
                                    

                                    switch (Stage)
                                    {
                                        case "D1":
                                            status = item.D1ShowSetting;
                                            ChangeStatus = "D1ShowSetting";
                                            break;
                                        case "D2":
                                            status = item.D2ShowSetting;
                                            ChangeStatus = "D2ShowSetting";
                                            break;
                                        case "D3":
                                            status = item.D3ShowSetting;
                                            ChangeStatus = "D3ShowSetting";
                                            break;
                                        case "D4":
                                            status = item.D4ShowSetting;
                                            ChangeStatus = "D4ShowSetting";
                                            break;
                                        case "D5":
                                            status = item.D5ShowSetting;
                                            ChangeStatus = "D5ShowSetting";
                                            break;
                                        case "D6":
                                            status = item.D6ShowSetting;
                                            ChangeStatus = "D6ShowSetting";
                                            break;
                                        case "D7":
                                            status = item.D7ShowSetting;
                                            ChangeStatus = "D7ShowSetting";
                                            break;
                                        case "D8":
                                            status = item.D8ShowSetting;
                                            ChangeStatus = "D8ShowSetting";
                                            break;
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
                                    for (int i = 0; i <= 7; i++)
                                    {
                                        if (Convert.ToInt32(Stage.Split('D')[1]) == (i + 1))
                                        {
                                            var StageDataSetting = item.StageDataSetting;
                                            if (StageDataSetting[i].ToString() == "N" && status == "A") throw new SystemException( Stage + "階段不需要維護,故PDF不須顯示!!");
                                        }
                                    }
                                    #endregion
                                }
                                #endregion

                                #region //更新SQL
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE QMS.CustomerComplaint SET "
                                        + ChangeStatus + @" = @Status,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE CcId = @CcId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        Status = status,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        CcId
                                    });
                                var insertResult2 = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                                break;
                            case "ALL":
                                #region //更新SQL
                                var D1ShowSetting = "";
                                var D2ShowSetting = "";
                                var D3ShowSetting = "";
                                var D4ShowSetting = "";
                                var D5ShowSetting = "";
                                var D6ShowSetting = "";
                                var D7ShowSetting = "";
                                var D8ShowSetting = "";
                                foreach (var item in result)
                                {
                                    var StageDataSetting = item.StageDataSetting;
                                    D1ShowSetting = StageDataSetting[0].ToString() != "N" ? item.D1ShowSetting == "A" ? "S" : "A" : "S";
                                    D2ShowSetting = StageDataSetting[1].ToString() != "N" ? item.D2ShowSetting == "A" ? "S" : "A" : "S";
                                    D3ShowSetting = StageDataSetting[2].ToString() != "N" ? item.D3ShowSetting == "A" ? "S" : "A" : "S";
                                    D4ShowSetting = StageDataSetting[3].ToString() != "N" ? item.D4ShowSetting == "A" ? "S" : "A" : "S";
                                    D5ShowSetting = StageDataSetting[4].ToString() != "N" ? item.D5ShowSetting == "A" ? "S" : "A" : "S";
                                    D6ShowSetting = StageDataSetting[5].ToString() != "N" ? item.D6ShowSetting == "A" ? "S" : "A" : "S";
                                    D7ShowSetting = StageDataSetting[6].ToString() != "N" ? item.D7ShowSetting == "A" ? "S" : "A" : "S";
                                    D8ShowSetting = StageDataSetting[7].ToString() != "N" ? item.D8ShowSetting == "A" ? "S" : "A" : "S";
                                }

                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE QMS.CustomerComplaint SET
                                        D1ShowSetting = @D1ShowSetting,
                                        D2ShowSetting = @D2ShowSetting,
                                        D3ShowSetting = @D3ShowSetting,
                                        D4ShowSetting = @D4ShowSetting,
                                        D5ShowSetting = @D5ShowSetting,
                                        D6ShowSetting = @D6ShowSetting,
                                        D7ShowSetting = @D7ShowSetting,
                                        D8ShowSetting = @D8ShowSetting,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE CcId = @CcId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        D1ShowSetting,
                                        D2ShowSetting,
                                        D3ShowSetting,
                                        D4ShowSetting,
                                        D5ShowSetting,
                                        D6ShowSetting,
                                        D7ShowSetting,
                                        D8ShowSetting,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        CcId
                                    });
                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                                break;
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

        #region //UpdateStageDataSetting -- 更新階段維護設定 -- Shintokru 2023.09.12
        public string UpdateStageDataSetting(int CcId, string Stage)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;
                        string StageDataSetting = "";
                        string status = "";
                        string noeShowSetting = "";
                        #region//檢核客訴單據是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 
                                StageDataSetting
                                FROM QMS.CustomerComplaint a
                                WHERE a.CcId=@CcId";
                        dynamicParameters.Add("CcId", CcId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("客訴單據查無資料,請重新確認");
                        foreach(var item in result)
                        {
                            StageDataSetting = item.StageDataSetting;
                        }

                        int index = Convert.ToInt32(Stage.Split('D')[1])-1;
                        char fifthCharacter = StageDataSetting[index];
                        #endregion

                        #region //調整為相反狀態
                        switch (fifthCharacter.ToString())
                        {
                            case "Y":
                                status = "N";
                                break;
                            case "N":
                                status = "Y";
                                break;
                        }
                        StageDataSetting = StageDataSetting.Substring(0, index) + status + StageDataSetting.Substring(index+1);

                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 
                                D" + (index + 1) + @"ShowSetting  ShowSetting
                                FROM QMS.CustomerComplaint a
                                WHERE a.CcId=@CcId";
                        dynamicParameters.Add("CcId", CcId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        foreach(var item in result)
                        {
                            noeShowSetting = item.ShowSetting;
                        }

                        #region //更新SQL
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE QMS.CustomerComplaint SET
                                StageDataSetting = @StageDataSetting,
                                D"+ (index +1)+ @"ShowSetting = @status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE CcId = @CcId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                StageDataSetting,
                                status =  status == "N" ? noeShowSetting == "A" ? "S" : "S" : "S",
                                LastModifiedDate,
                                LastModifiedBy,
                                CcId
                            });
                        var insertResult2 = sqlConnection.Query(sql, dynamicParameters);

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

        #region //UpdateStageDataStatus --確認/反確認各階段資料 --GPai 20230918
        public string UpdateStageDataStatus(int CcId, string Stage, string StageStatus)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        int rowsAffected = 0;
                        string currentStatus = "";
                        string StageDataSetting = "";
                        string confirmStatus = StageStatus;
                        string replyDate = null;
                        string closureDate = null;

                        #region//檢核客訴單據是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT StageDataSetting
                                FROM QMS.CustomerComplaint a
                                WHERE a.CcId=@CcId";
                        dynamicParameters.Add("CcId", CcId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("查無客訴單據,請重新確認");
                        foreach(var item in result)
                        {
                            StageDataSetting = item.StageDataSetting;
                        }
                        int lastStage = StageDataSetting.LastIndexOf("Y")+1;
                        #endregion

                        #region//檢核客訴單據各階段資料
                        if (Stage != "D1")
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT *
                                FROM QMS.CcStageData
                                WHERE CcId = @CcId AND Stage = @Stage";
                            dynamicParameters.Add("CcId", CcId);
                            dynamicParameters.Add("Stage", Stage);
                            var stageResult = sqlConnection.Query(sql, dynamicParameters);
                            if (stageResult.Count() <= 0) throw new SystemException("該階段無資料，無法進行確認/反確認");
                            if (stageResult.First().ConfirmStatus == "Y" && StageStatus == "Y") throw new SystemException("該階段已確認!");
                            if (stageResult.First().ConfirmStatus == "N" && StageStatus == "N") throw new SystemException("該階段已反確認!");
                        }
                        else
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT *
                                FROM QMS.CcTeamMembers a
                                WHERE a.CcId = @CcId";
                            dynamicParameters.Add("CcId", CcId);
                            var d1Result = sqlConnection.Query(sql, dynamicParameters);
                            if (d1Result.Count() <= 0) throw new SystemException("該階段無資料，無法進行確認/反確認");
                            if (d1Result.First().ConfirmStatus == "Y" && Stage == "Y") throw new SystemException("該階段已確認!");
                            if (d1Result.First().ConfirmStatus == "N" && Stage == "N") throw new SystemException("該階段已反確認!");
                        }

                        #endregion

                        #region//更新單身
                        if (Stage != "D1")
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE QMS.CcStageData SET
                                ConfirmStatus = @ConfirmStatus,
                                ConfirmDate = @ConfirmDate,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE CcId = @CcId AND Stage = @Stage";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    ConfirmStatus = confirmStatus,
                                    ConfirmDate = LastModifiedDate,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    CcId,
                                    Stage
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        }
                        else
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE QMS.CcTeamMembers SET
                                ConfirmStatus = @ConfirmStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE CcId = @CcId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    ConfirmStatus = confirmStatus,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    CcId
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        }
                        #endregion

                        #region//更新單頭
                        #region//客訴單據確認資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT *
                                FROM QMS.CcStageData
                                WHERE CcId = @CcId 
                                AND ConfirmStatus = 'Y'
                                ORDER BY Stage DESC";
                        dynamicParameters.Add("CcId", CcId);
                        var stageConfirmResult1 = sqlConnection.Query(sql, dynamicParameters);
                        #endregion

                        #region//檢核客訴單據有無未確認資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT *
                                FROM QMS.CcStageData
                                WHERE CcId = @CcId 
                                AND ConfirmStatus = 'N'
								AND (Stage != '' AND Stage is not null)
                                ORDER BY Stage";
                        dynamicParameters.Add("CcId", CcId);
                        var stageConfirmResult2 = sqlConnection.Query(sql, dynamicParameters);
                        #endregion

                        #region//檢核客訴單據有無未確認D1資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT *
                                FROM QMS.CcTeamMembers
                                WHERE CcId = @CcId 
                                AND ConfirmStatus = 'N'";
                        dynamicParameters.Add("CcId", CcId);
                        var stageConfirmResult3 = sqlConnection.Query(sql, dynamicParameters);
                        
                        #endregion

                        if (Stage == "D1" && StageStatus == "N")
                        {
                            currentStatus = "D1";
                        }
                        else
                        {
                            if (stageConfirmResult2.Count() == 0)
                            {
                                if (stageConfirmResult1.Count() > 0)
                                {
                                    switch (stageConfirmResult1.First().Stage)
                                    {

                                        case "D2":
                                            currentStatus = "D3";
                                            break;
                                        case "D3":
                                            currentStatus = "D4";
                                            break;
                                        case "D4":
                                            currentStatus = "D5";
                                            break;
                                        case "D5":
                                            currentStatus = "D6";
                                            break;
                                        case "D6":
                                            currentStatus = "D7";
                                            break;
                                        case "D7":
                                            currentStatus = "D8";
                                            break;
                                        case "D8":
                                            currentStatus = "DF";
                                            break;
                                    }

                                    if (stageConfirmResult3.Count() > 0)
                                    {
                                        currentStatus = "D1";
                                    }
                                }
                                else {
                                    currentStatus = "D1";
                                }
                            }
                            else
                            {
                                if (stageConfirmResult3.Count() > 0)
                                {
                                    currentStatus = "D1";
                                }
                                else
                                {
                                    currentStatus = stageConfirmResult2.FirstOrDefault().Stage;
                                }
                            }
                        }
                        if(currentStatus != "DF")
                        {
                            if (Convert.ToInt32(Stage.Split('D')[1]) == lastStage || Convert.ToInt32(currentStatus.Split('D')[1]) > lastStage)
                            {
                                currentStatus = "DF";
                            }
                        }

                        

                        //d7回復*ReplyDate/d8 結案日ClosureDate 
                        if (Stage == "D7" && StageStatus == "Y")
                        {
                            replyDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE QMS.CustomerComplaint SET
                                CurrentStatus = @CurrentStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy,
                                ReplyDate = @ReplyDate
                                WHERE CcId = @CcId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    CurrentStatus = currentStatus,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    ReplyDate = replyDate,
                                    CcId
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        } else if (Stage == "D8" && StageStatus == "Y")
                        {
                            closureDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE QMS.CustomerComplaint SET
                                CurrentStatus = @CurrentStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy,
                                ClosureDate = @ClosureDate
                                WHERE CcId = @CcId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    CurrentStatus = currentStatus,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    ClosureDate = closureDate,
                                    CcId
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        }
                        else
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE QMS.CustomerComplaint SET
                                CurrentStatus = @CurrentStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE CcId = @CcId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    CurrentStatus = currentStatus,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    CcId
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

        #region //UpdateD8TableData --更新D8附表 --GPai 20230925
        public string UpdateD8TableData(int CCTrackingDataId, int Batch, string ReportNo, string TrialResult, string TrackingDate
            )
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {

                    string expectedDate = "";
                    string finishDate = "";



                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷D8附表資料是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT *
                                FROM QMS.CCTrackingData
                                WHERE CCTrackingDataId = @CCTrackingDataId";
                        dynamicParameters.Add("CCTrackingDataId", CCTrackingDataId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("找不到資料，請重新確認");
                        string CurrentStatus = "";
                       
                        #endregion

                        

                        #region//更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE QMS.CCTrackingData SET
                                Batch = @Batch,
                                ReportNo = @ReportNo,
                                TrialResult = @TrialResult,
                                TrackingDate = @TrackingDate,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE CCTrackingDataId = @CCTrackingDataId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Batch,
                                ReportNo,
                                TrialResult,
                                TrackingDate,
                                LastModifiedDate,
                                LastModifiedBy,
                                CCTrackingDataId
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

        #region //UpdateCCMenuCategoryStatus //選單類別狀態
        public string UpdateCCMenuCategoryStatus(int McId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷系統別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.McId, Status
                                FROM QMS.MenuCategory a
                                WHERE McId = @McId";
                        dynamicParameters.Add("McId", McId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("資料錯誤!");

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
                        sql = @"UPDATE QMS.MenuCategory SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE McId = @McId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                McId
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

        #endregion

        #region //Delete

        #region DeleteCCMenuCategory //刪除選單類別 --GPai 20230912
        public string DeleteCCMenuCategory(int McId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        int rowsAffected = 0;

                        #region //判斷資料是否存在
                        sql = @"SELECT TOP 1 1
                                FROM QMS.MenuCategory
                                WHERE McId = @McId";
                        dynamicParameters.Add("McId", McId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("找不到類別資料!請重新確認");
                        #endregion



                        #region //刪除主要table //
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE QMS.MenuCategory
                                WHERE McId = @McId";
                        dynamicParameters.Add("McId", McId);

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

        #region //DeleteCustomerComplaint -- 刪除客訴單據 -- Shintokuro 2023-09-06
        public string DeleteCustomerComplaint(int CcId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        int rowsAffected = 0;

                        #region //判斷資料是否正確
                        sql = @"SELECT TOP 1 ConfirmStatus
                                FROM QMS.CustomerComplaint
                                WHERE CcId = @CcId";
                        dynamicParameters.Add("CcId", CcId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("客訴單據資訊資料錯誤!");
                        foreach (var item in result)
                        {
                            if (item.ConfirmStatus == "V") throw new SystemException("客訴單據已經作廢不可以異動刪除!");
                        }
                        #endregion

                        #region //判斷子資料是否已建立資料
                        sql = @"SELECT CcId FROM QMS.CcFile WHERE CcId = @CcId UNION All
                                SELECT CcId FROM QMS.CcTeamMembers WHERE CcId = @CcId UNION All
                                SELECT CcId FROM QMS.CcStageData WHERE CcId = @CcId";
                        dynamicParameters.Add("CcId", CcId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("客訴單據的子資料已建立資料，不可以刪除!");
                        #endregion

                        #region //檔案資料
                        sql = @"SELECT TOP 1 1
                                FROM QMS.CcFile
                                WHERE CcId = @CcId";
                        dynamicParameters.Add("CcId", CcId);

                        var cccFileresult = sqlConnection.Query(sql, dynamicParameters);
                        #endregion

                        #region //刪除檔案table //
                        if (cccFileresult.Count() > 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE QMS.CcFile
                                WHERE CcId = @CcId";
                            dynamicParameters.Add("CcId", CcId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        }
                        #endregion

                        #region //刪除主要table //
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE QMS.CustomerComplaint
                                WHERE CcId = @CcId";
                        dynamicParameters.Add("CcId", CcId);

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
        #region //DeleteCcStageData -- 刪除客訴單據各階段資料 -- Shintokuro 2023-09-08
        public string DeleteCcStageData(int CcId, int CcStageDataId, string Stage)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        int rowsAffected = 0;

                        #region //判斷客訴單據資料是否存在
                        sql = @"SELECT TOP 1 1
                                FROM QMS.CustomerComplaint
                                WHERE CcId = @CcId";
                        dynamicParameters.Add("CcId", CcId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("找不到客訴單據資訊資料!請重新確認");
                        #endregion

                        #region //判斷客訴單據各階段資料資料是否正確
                        sql = @"SELECT *
                                FROM QMS.CcStageData
                                WHERE CcId = @CcId
                                AND Stage = @Stage";
                        dynamicParameters.Add("CcId", CcId);
                        dynamicParameters.Add("Stage", Stage);
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CcStageDataId", @" AND CcStageDataId = @CcStageDataId", CcStageDataId);

                        var stageDataresult = sqlConnection.Query(sql, dynamicParameters);
                        if (stageDataresult.Count() <= 0) throw new SystemException("找不到客訴單據該階段資料!請重新確認!");
                        if (stageDataresult.First().ConfirmStatus == "Y") throw new SystemException("該階段資料已確認!");
                        #endregion

                        #region //階段資料檔案
                        sql = @"SELECT a.CcStageDataFileId,a.CcStageDataId, a.FileId
                                , b.[FileName]
                                , c.CcId
                                FROM QMS.CcStageDataFile a
                                INNER JOIN BAS.[File] b on a.FileId = b.FileId
                                INNER JOIN QMS.CcStageData c on a.CcStageDataId = c.CcStageDataId
                                WHERE c.CcId = @CcId
                                AND c.Stage = @Stage";
                        dynamicParameters.Add("CcId", CcId);
                        dynamicParameters.Add("Stage", Stage);
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CcStageDataId", @" AND a.CcStageDataId = @CcStageDataId", CcStageDataId);

                        var stageDataFileresult = sqlConnection.Query(sql, dynamicParameters);

                        #endregion

                        #region //D8附表
                        if (Stage == "D8")
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT *
                                    FROM QMS.CCTrackingData a
                                    INNER JOIN QMS.CcStageData b on a.CcStageDataId = b.CcStageDataId
                                    INNER JOIN QMS.MenuCategory c on b.McId = c.McId
                                    WHERE b.CcId = @CcId";
                            dynamicParameters.Add("CcId", CcId);

                            var d8tableresult = sqlConnection.Query(sql, dynamicParameters);
                            if (d8tableresult.Count() > 0) {
                                #region //刪除D8附表
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE QMS.CCTrackingData
                                WHERE CcStageDataId = @CcStageDataId";
                                dynamicParameters.Add("CcStageDataId", d8tableresult.FirstOrDefault().CcStageDataId);

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }
                        }
                        #endregion

                        #region //刪除
                        foreach (var item in stageDataFileresult)
                        {
                            #region //刪除CC檔案
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE QMS.CcStageDataFile
                                WHERE FileId = @FileId";
                            dynamicParameters.Add("FileId", item.FileId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //刪除檔案
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE BAS.[File]
                                WHERE FileId = @FileId";
                            dynamicParameters.Add("FileId", item.FileId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                           
                        }



                        #region //刪除主要table //
                        if (CcStageDataId > 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE QMS.CcStageData
                                WHERE CcStageDataId = @CcStageDataId";
                            dynamicParameters.Add("CcStageDataId", CcStageDataId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        }
                        else
                        {
                            
                            foreach (var item in stageDataresult)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE QMS.CcStageData
                                WHERE CcStageDataId = @CcStageDataId";
                                dynamicParameters.Add("CcStageDataId", item.CcStageDataId);

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            }
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

        #region //DeleteCcStageDataFile -- 刪除客訴單據各階段檔案 -- Shintokuro 2023-09-12
        public string DeleteCcStageDataFile(int CcStageDataFileId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        DynamicParameters dynamicParameters = new DynamicParameters();

                        int FileId = -1;

                        #region //判斷客訴單據各階段檔案資料是否正確
                        sql = @"SELECT TOP 1 FileId
                                FROM QMS.CcStageDataFile
                                WHERE CcStageDataFileId = @CcStageDataFileId";
                        dynamicParameters.Add("CcStageDataFileId", CcStageDataFileId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("品異條碼資料錯誤!");
                        foreach (var item in result)
                        {
                            FileId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).FileId;
                        }
                        #endregion

                        #region //判斷檔案是否存在
                        dynamicParameters = new DynamicParameters();

                        sql = @"SELECT TOP 1 FileId
                                FROM BAS.[File]
                                WHERE FileId = @FileId";
                        dynamicParameters.Add("FileId", FileId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("檔案不存在,無法執行刪除!");
                        #endregion

                        #region //先刪除關聯Table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE FROM QMS.CcStageDataFile
                                WHERE CcStageDataFileId = @CcStageDataFileId";
                        dynamicParameters.Add("CcStageDataFileId", CcStageDataFileId);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE FROM BAS.[File]
                                WHERE FileId = @FileId";
                        dynamicParameters.Add("FileId", FileId);

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

        #region //DeleteCCUser --刪除D1成員
        public string DeleteCCUser(int UserId, int CcId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷D1是否已確認
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT *
                                FROM QMS.CcTeamMembers
                                WHERE CcId = @CcId";
                        dynamicParameters.Add("CcId", CcId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        //if (result.Count() <= 0) throw new SystemException("找不到客訴單成員名單，請重新確認");
                        string CurrentStatus = "";
                        foreach (var item in result)
                        {
                            if (item.ConfirmStatus != "N") throw new SystemException("該階段資料已經確認,無法進行更動!");
                        }
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE FROM QMS.CcTeamMembers
                                WHERE UserId = @UserId AND CcId = @CcId";
                        dynamicParameters.Add("UserId", UserId);
                        dynamicParameters.Add("CcId", CcId);

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

        #region //DeleteCCMainFile --刪除單頭檔案 --GPai 20230922
        public string DeleteCCMainFile(int CcFileId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        DynamicParameters dynamicParameters = new DynamicParameters();

                        int FileId = -1;

                        #region //判斷客訴單據各階段檔案資料是否正確
                        sql = @"SELECT TOP 1 FileId
                                FROM QMS.CcFile
                                WHERE CcFileId = @CcFileId";
                        dynamicParameters.Add("CcFileId", CcFileId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("品異條碼資料錯誤!");
                        foreach (var item in result)
                        {
                            FileId = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).FileId;
                        }
                        #endregion

                        #region //判斷檔案是否存在
                        dynamicParameters = new DynamicParameters();

                        sql = @"SELECT TOP 1 FileId
                                FROM BAS.[File]
                                WHERE FileId = @FileId";
                        dynamicParameters.Add("FileId", FileId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("檔案不存在,無法執行刪除!");
                        #endregion

                        #region //先刪除關聯Table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE FROM QMS.CcFile
                                WHERE CcFileId = @CcFileId";
                        dynamicParameters.Add("CcFileId", CcFileId);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE FROM BAS.[File]
                                WHERE FileId = @FileId";
                        dynamicParameters.Add("FileId", FileId);

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

        #region //DeleteD8TableData --刪除D8附表 --GPai 20230925
        public string DeleteD8TableData(int CCTrackingDataId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        int rowsAffected = 0;

                        #region //判斷客訴單據資料是否存在
                        sql = @"SELECT TOP 1 1
                                FROM QMS.CCTrackingData
                                WHERE CCTrackingDataId = @CCTrackingDataId";
                        dynamicParameters.Add("CCTrackingDataId", CCTrackingDataId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("找不到資料!請重新確認");
                        #endregion

                        #region //刪除

                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE QMS.CCTrackingData
                                WHERE CCTrackingDataId = @CCTrackingDataId";
                        dynamicParameters.Add("CCTrackingDataId", CCTrackingDataId);

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

        #endregion

        #region //暫存
        #endregion

        #region //Model
        public class CustomerComplaintForm
        {
            public int CcId { get; set; }
            public int CompanyId { get; set; }
            public string CcNo { get; set; }
            public int MtlItemId { get; set; }
            public string MtlItemNo { get; set; }
            public string MtlItemName { get; set; }
            public string MtlItemSpec { get; set; }
            public string FilingDate { get; set; }
            public string DocDate { get; set; }
            public string ReplyDate { get; set; }
            public string ClosureDate { get; set; }
            public int CustomerId { get; set; }
            public string CustomerNo { get; set; }
            public string CustomerShortName { get; set; }
            public int UserId { get; set; }
            public string UserName { get; set; }
            public string UserEMail { get; set; }
            public string IssueDescription { get; set; }
            public string CurrentStatus { get; set; }
            public List<CCMainFile> CCMainFile { get; set; }
            public List<CCMember> D1 { get; set; }
            public List<StageDataDetail> D2 { get; set; }
            public List<StageDataDetail> D3 { get; set; }
            public List<StageDataDetail> D4 { get; set; }
            public List<StageDataDetail> D5 { get; set; }
            public List<StageDataDetail> D6 { get; set; }
            public List<StageDataDetail> D7 { get; set; }
            public List<StageDataDetail> D8 { get; set; }
            public List<CCTrackingData> TrackingData { get; set; }

        }

        public class CCMember
        {
            public int CcTeamMembersId { get; set; }
            public int CcId { get; set; }
            public int UserId { get; set; }
            public string UserNo { get; set; }
            public string UserName { get; set; }
            public int DepartmentId { get; set; }
            public string DepartmentNo { get; set; }
            public string DepartmentName { get; set; }
        }

        public class StageDataDetail
        {
            public int CcStageDataId { get; set; }
            public int CcId { get; set; }
            public string Stage { get; set; }
            public int McId { get; set; }
            public string McName { get; set; }
            public string DescriptionTextarea { get; set; }
            public int DeptId { get; set; }
            public string DepartmentNo { get; set; }
            public string DepartmentName { get; set; }
            public int UserId { get; set; }
            public string UserNo { get; set; }
            public string UserName { get; set; }
            public string ExpectedDate { get; set; }
            public string FinishDate { get; set; }
            public List<StageDataFile> DataFile { get; set; }
        }

        public class StageDataFile
        {
            public int CcId { get; set; }
            public int CcStageDataFileId { get; set; }
            public int CcStageDataId { get; set; }
            public string Stage { get; set; }
            public int FileId { get; set; }
            public string FileName { get; set; }
        }

        public class CCMainFile
        {
            public int CcId { get; set; }
            public int CcFileId { get; set; }
            public string FileType { get; set; }
            public int FileId { get; set; }
            public string FileName { get; set; }
        }

        public class CCTrackingData
        {
            public int CcId { get; set; }
            public int CCTrackingDataId { get; set; }
            public int CcStageDataId { get; set; }
            public int Batch { get; set; }
            public string ReportNo { get; set; }
            public string TrialResult { get; set; }
            public string TrackingDate { get; set; }
        }
        #endregion

    }
}
