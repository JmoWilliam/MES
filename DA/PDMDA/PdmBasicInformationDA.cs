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

namespace PDMDA
{
    public class PdmBasicInformationDA
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

        public PdmBasicInformationDA()
        {
            MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];
            ErpConnectionStrings = ConfigurationManager.AppSettings["ErpDb"];
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
        #region //GetMtlElement -- 取得品號元素資料 -- Kan 2022.05.31
        public string GetMtlElement(int ElementId, string ElementNo, string ElementName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.ElementId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ElementNo, a.ElementName, a.SortNumber, a.Status";
                    sqlQuery.mainTables =
                        @"FROM PDM.MtlElement a";
                    sqlQuery.auxTables = "";
                    string queryCondition = " AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ElementId", @" AND a.ElementId = @ElementId", ElementId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ElementNo", @" AND a.ElementNo LIKE '%' + @ElementNo + '%'", ElementNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ElementName", @" AND a.ElementName LIKE '%' + @ElementName + '%'", ElementName);
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

        #region //GetUnitOfMeasure -- 取得單位基本資料 -- Kan 2022.06.15
        public string GetUnitOfMeasure(int UomId, string UomNo, string UomName, string UomType
            , string Status, int RejectUomId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.UomId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.UomNo, a.UomName, a.UomType, a.UomDesc, a.[Status], a.UomNo + ' ' + a.UomName UomWithNo";
                    sqlQuery.mainTables =
                        @"FROM PDM.UnitOfMeasure a";
                    sqlQuery.auxTables = "";
                    string queryCondition = " AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UomId", @" AND a.UomId = @UomId", UomId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UomNo", @" AND a.UomNo LIKE '%' + @UomNo + '%'", UomNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UomName", @" AND a.UomName LIKE '%' + @UomName + '%'", UomName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UomType", @" AND a.UomType LIKE '%' + @UomType + '%'", UomType);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RejectUomId", @" AND a.UomId != @RejectUomId", RejectUomId);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.UomId";
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

        #region //GetUomCalculate -- 取得單位換算資料 -- Kan 2022.06.15
        public string GetUomCalculate(int UomCalculateId, string ConvertUomNo, string ConvertUomName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.UomCalculateId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ConvertUomId, a.[Value], a.[Status]
                        , b.UomId, b.UomNo, b.UomName, b.UomType";
                    sqlQuery.mainTables =
                        @"FROM PDM.UomCalculate a
                        INNER JOIN PDM.UnitOfMeasure b ON a.ConvertUomId = b.UomId";
                    sqlQuery.auxTables = "";
                    string queryCondition = " AND b.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UomCalculateId", @" AND a.UomCalculateId = @UomCalculateId", UomCalculateId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ConvertUomNo", @" AND b.UomNo LIKE '%' + @ConvertUomNo + '%'", ConvertUomNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ConvertUomName", @" AND b.UomName LIKE '%' + @ConvertUomName + '%'", ConvertUomName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.UomCalculateId";
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

        #region //GetMtlModel -- 取得品號機型資料 -- Kan 2022.06.17
        public string GetMtlModel(int MtlModelId, int ParentId, string MtlModelNo, string MtlModelName, string Status
            , string QueryType, int StartParent, string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    string declareSql = @"DECLARE @rowsAdded int

                                        DECLARE @mtlModel TABLE
                                        ( 
                                          MtlModelId int,
                                          ParentId int,
                                          MtlModelLevel int,
                                          MtlModelRoute nvarchar(MAX),
                                          MtlModelSort nvarchar(MAX),
                                          processed int DEFAULT(0)
                                        )

                                        INSERT @mtlModel
                                            SELECT MtlModelId, ParentId, 1 MtlModelLevel
                                            , CAST(ParentId AS nvarchar(MAX)) AS MtlModelRoute
                                            , CAST(MtlModelSort AS nvarchar(MAX)) AS MtlModelSort, 0
                                            FROM PDM.MtlModel
                                            WHERE CompanyId = @CompanyId
                                            AND ParentId = @StartParent

                                        SET @rowsAdded=@@rowcount

                                        WHILE @rowsAdded > 0
                                        BEGIN
                                          UPDATE @mtlModel SET processed = 1 WHERE processed = 0

                                          INSERT @mtlModel
                                              SELECT a.MtlModelId, a.ParentId, ( b.MtlModelLevel + 1 ) MtlModelLevel
                                              , CAST(b.MtlModelRoute + ',' + CAST(a.ParentId AS nvarchar(MAX)) AS nvarchar(MAX)) AS MtlModelRoute
                                              , CAST(b.MtlModelSort + CAST(a.MtlModelSort AS nvarchar(MAX)) AS nvarchar(MAX)) AS MtlModelSort, 0
                                              FROM PDM.MtlModel a
                                              INNER JOIN @mtlModel b ON a.ParentId = b.MtlModelId
                                              WHERE a.ParentId <> a.MtlModelId 
                                              AND b.processed = 1

                                          SET @rowsAdded = @@rowcount

                                          UPDATE @mtlModel SET processed = 2 WHERE processed = 1
                                        END;";

                    switch (QueryType)
                    {
                        case "Query":
                            sqlQuery.mainKey = "a.MtlModelId";
                            sqlQuery.auxKey = "";
                            sqlQuery.columns =
                                @", a.ParentId, a.MtlModelLevel, a.MtlModelRoute, a.MtlModelSort
                                , b.MtlModelNo, b.MtlModelName, b.MtlModelDesc, b.[Status], b.MtlModelNo + ' ' + b.MtlModelName MtlModelWithNo
                                , CASE WHEN c.MtlModelNo IS NOT NULL THEN c.MtlModelNo + '-' + c.MtlModelName ELSE '' END ParentWithNo";
                            sqlQuery.mainTables =
                                @"FROM @mtlModel a
                                INNER JOIN PDM.MtlModel b ON a.MtlModelId = b.MtlModelId
                                LEFT JOIN PDM.MtlModel c ON b.ParentId = c.MtlModelId";
                            sqlQuery.auxTables = "";
                            string queryCondition = " AND b.CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("StartParent", StartParent);
                            BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlModelId", @" AND b.MtlModelId = @MtlModelId", MtlModelId);
                            BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ParentId", @" AND b.ParentId = @ParentId", ParentId, -2);
                            BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlModelNo", @" AND b.MtlModelNo LIKE '%' + @MtlModelNo + '%'", MtlModelNo);
                            BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlModelName", @" AND b.MtlModelName LIKE '%' + @MtlModelName + '%'", MtlModelName);
                            if (Status.Length > 0) BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "Status", @" AND b.Status IN @Status", Status.Split(','));
                            sqlQuery.conditions = queryCondition;
                            sqlQuery.declarePart = declareSql;
                            sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.MtlModelSort";
                            sqlQuery.pageIndex = PageIndex;
                            sqlQuery.pageSize = PageSize;

                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                data = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery)
                            });
                            #endregion
                            break;
                        case "Max":
                            sql = declareSql;
                            sql += @"SELECT MAX(a.MtlModelLevel) MaxLevel
                                     FROM @mtlModel a
                                     INNER JOIN PDM.MtlModel b ON a.MtlModelId = b.MtlModelId
                                     WHERE b.CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("StartParent", StartParent);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlModelId", @" AND b.MtlModelId = @MtlModelId", MtlModelId);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ParentId", @" AND b.ParentId = @ParentId", ParentId, -2);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlModelNo", @" AND b.MtlModelNo LIKE '%' + @MtlModelNo + '%'", MtlModelNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlModelName", @" AND b.MtlModelName LIKE '%' + @MtlModelName + '%'", MtlModelName);
                            if (Status.Length > 0) BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "Status", @" AND b.Status IN @Status", Status.Split(','));

                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                data = sqlConnection.Query(sql, dynamicParameters)
                            });
                            #endregion
                            break;
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

        #region //GetPrinciple -- 取得品號原則 -- Ann 2022.07.07
        public string GetPrinciple(int PrincipleId, int BuildId, string Value, string OrderBy)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.PrincipleId, a.CompanyId, a.PrincipleDesc, a.[Status]
                            , (
                                SELECT b.BuildId, b.ElementId, b.Digits, b.SortNumber, b.Surfix, b.AutoCreate, b.EditAuthority
                                , ba.ElementNo, ba.ElementName
                                , ISNULL(
                                    (
                                        SELECT bb.DataId, bb.BuildId, bb.DataNo, bb.DataName
                                        FROM PDM.MtlPrincipleBuildData bb
                                        WHERE bb.BuildId = b.BuildId
                                        FOR JSON PATH, ROOT('data')
                                    ), 'null'
                                ) MtlPrincipleBuildData
                                , ISNULL(
                                    (
                                        SELECT bc.ConditionId, bc.BuildId, bc.[Value], bc.ToPrinciple, bc.ConditionDesc
                                        FROM PDM.MtlPrincipleCondition bc
                                        WHERE bc.BuildId = b.BuildId
                                        FOR JSON PATH, ROOT('data')
                                    ), 'null'
                                ) MtlPrincipleCondition
                                FROM PDM.MtlPrincipleBuild b
                                INNER JOIN PDM.MtlElement ba ON b.ElementId = ba.ElementId
                                WHERE b.PrincipleId = a.PrincipleId
                                ORDER BY b.SortNumber
                                FOR JSON PATH, ROOT('data')
                            ) MtlPrincipleBuild
                            FROM PDM.MtlPrinciple a
                            WHERE a.CompanyId = @CompanyId
                            AND a.[Status] = 'A'";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PrincipleId", @" AND a.PrincipleId = @PrincipleId", PrincipleId);

                    //sql = @"SELECT a.PrincipleId, a.PrincipleDesc
                    //        , b.BuildId, b.SortNumber, b.Digits, b.Surfix, b.AutoCreate, b.EditAuthority
                    //        , e.ElementNo, e.ElementName
                    //        , ISNULL(c.ConditionId,-1) ConditionId, c.ConditionDesc, c.[Value], c.ToPrinciple
                    //        , ISNULL(d.DataId,-1) DataId, d.DataNo, d.DataName, (d.DataNo + '-' + d.DataName) PrincipleWithText
                    //        FROM PDM.MtlPrinciple a
                    //        LEFT JOIN PDM.MtlPrincipleBuild b ON a.PrincipleId = b.PrincipleId
                    //        LEFT JOIN PDM.MtlPrincipleBuildData d ON b.BuildId = d.BuildId
                    //        LEFT JOIN PDM.MtlPrincipleCondition c ON b.BuildId = c.BuildId AND (c.[Value] = d.DataNo  OR c.[Value] = 'none')
                    //        INNER JOIN PDM.MtlElement e ON b.ElementId = e.ElementId
                    //        WHERE 1=1
                    //        AND a.CompanyId = @CompanyId";
                    //dynamicParameters.Add("CompanyId", CurrentCompany);
                    //BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PrincipleId", @" AND a.PrincipleId = @PrincipleId", PrincipleId);
                    //BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "BuildId", @" AND b.BuildId = @BuildId", BuildId);
                    //BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "Value", @" AND c.Value = @Value", Value);

                    //OrderBy = OrderBy.Length > 0 ? OrderBy : "b.SortNumber";
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

        #region //GetPrincipleCondition -- 取得品號原則條件 -- Ann 2023-02-02
        public string GetPrincipleCondition(int ConditionId, int BuildId, string PrincipleValue)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.ConditionId, a.BuildId, a.[Value], a.ToPrinciple, a.ConditionDesc
                            , c.ElementNo, c.ElementName
                            FROM PDM.MtlPrincipleCondition a
                            INNER JOIN PDM.MtlPrincipleBuild b ON a.BuildId = b.BuildId
                            INNER JOIN PDM.MtlElement c ON b.ElementId = c.ElementId
                            WHERE a.BuildId = @BuildId";
                    dynamicParameters.Add("BuildId", BuildId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ConditionId", @" AND a.ConditionId = @ConditionId", ConditionId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "BuildId", @" AND a.BuildId = @BuildId", BuildId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PrincipleValue", @" AND a.Value = @PrincipleValue", PrincipleValue);

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

        #region //GetItemSegment -- 取得編碼節段資料 -- Shintokuro 2023.03.01
        public string GetItemSegment(int ItemSegmentId, string SegmentNo, string SegmentName, string StartDate, string EndDate, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.ItemSegmentId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" , a.SegmentNo, a.SegmentName, a.SegmentDesc, ISNULL(FORMAT(a.EffectiveDate, 'yyyy-MM-dd'), '') EffectiveDate, ISNULL(FORMAT(a.InactiveDate, 'yyyy-MM-dd'), '') InactiveDate
                           , a.Status
                           , a.SegmentNo + '-'+ a.SegmentName SegmentShow
                          ";
                    sqlQuery.mainTables =
                        @"FROM PDM.ItemSegment a
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ItemSegmentId", @" AND a.ItemSegmentId = @ItemSegmentId", ItemSegmentId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SegmentNo", @" AND a.SegmentNo LIKE '%' + @SegmentNo + '%'", SegmentNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SegmentName", @" AND a.SegmentName LIKE '%' + @SegmentName + '%'", SegmentName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.EffectiveDate >= @StartDate ", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.InactiveDate <= @EndDate ", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));


                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ItemSegmentId DESC";
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

        #region //GetItemSegmentSimple -- 取得編碼節段資料(無換頁) -- Shintokuro 2023.05.01
        public string GetItemSegmentSimple(string Today)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sql = @"SELECT test.* FROM
                            (SELECT  a.ItemSegmentId,a.SegmentNo, a.SegmentName
                            , a.SegmentNo + '-' + a.SegmentName SegmentShow
                            ,(CASE WHEN a.EffectiveDate is not null 
		                            THEN (
			                            case when CONVERT(varchar(10),a.EffectiveDate, 120)  <= @Today
			                            THEN 1
			                            ELSE 2
			                            END ) 
		                            ELSE 3
		                            END) as EffectiveDateStatus
                            ,(CASE WHEN a.InactiveDate is not null 
		                            THEN (
			                            case when CONVERT(varchar(10),a.InactiveDate, 120)  >= @Today
			                            THEN 1 
			                            ELSE 2 
			                            END ) 
		                            ELSE 3
		                            END) as InactiveDateStatus
                            FROM PDM.ItemSegment a
                            WHERE a.CompanyId = @CompanyId AND a.[Status]='A' ) test
                            WHERE test.EffectiveDateStatus !=2
							AND test.InactiveDateStatus !=2";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    dynamicParameters.Add("Today", Today);


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

        #region //GetItemSegmentValue -- 取得編碼節段Value資料 -- Shintokuro 2023.03.01
        public string GetItemSegmentValue(int SegmentValueId, int ItemSegmentId, string SegmentValueNo, string StartDate, string EndDate, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.SegmentValueId,a.ItemSegmentId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" , a.SegmentValueNo, a.SegmentValue, a.SegmentValueDesc, ISNULL(FORMAT(a.EffectiveDate, 'yyyy-MM-dd'), '') EffectiveDate, ISNULL(FORMAT(a.InactiveDate, 'yyyy-MM-dd'), '') InactiveDate
                           , a.Status
                           , a.SegmentValueNo+'-'+a.SegmentValue SegmentValueShow
                           , b.SegmentName + '-' +  a.SegmentValue SegmentName
                          ";
                    sqlQuery.mainTables =
                        @"FROM PDM.ItemSegmentValue a
                          INNER JOIN PDM.ItemSegment b on a.ItemSegmentId = b.ItemSegmentId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"";

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SegmentValueId", @" AND a.SegmentValueId = @SegmentValueId", SegmentValueId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ItemSegmentId", @" AND a.ItemSegmentId = @ItemSegmentId", ItemSegmentId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SegmentValueNo", @" AND a.SegmentValueNo LIKE '%' + @SegmentValueNo + '%'", SegmentValueNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.EffectiveDate >= @StartDate ", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.InactiveDate <= @EndDate ", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));


                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.SegmentValueId ASC";
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

        #region //GetItemSegmentValueSimple -- 取得編碼節段Value資料(無換頁) -- Shintokuro 2023.05.01
        public string GetItemSegmentValueSimple(int ItemSegmentId, string SegmentValueNo, string Today)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT t.* FROM
                            (SELECT  a.SegmentValueId,a.SegmentValueNo, a.SegmentValue
                            , a.SegmentValueNo + '-' + a.SegmentValue SegmentValueShow
                            ,(CASE WHEN a.EffectiveDate is not null 
		                        THEN (
			                        case when CONVERT(varchar(10),a.EffectiveDate, 120)  <= @Today
			                        THEN 1
			                        ELSE 2
			                        END ) 
		                        ELSE 3
		                        END) as EffectiveDateStatus
                            ,(CASE WHEN a.InactiveDate is not null 
		                        THEN (
			                        case when CONVERT(varchar(10),a.InactiveDate, 120)  >= @Today
			                        THEN 1 
			                        ELSE 2 
			                        END ) 
		                        ELSE 3
		                        END) as InactiveDateStatus
                            FROM PDM.ItemSegmentValue a
                            WHERE a.ItemSegmentId = @ItemSegmentId AND a.Status = 'A') t
                            WHERE t.EffectiveDateStatus !=2
							AND t.InactiveDateStatus !=2";
                    dynamicParameters.Add("ItemSegmentId", ItemSegmentId);
                    dynamicParameters.Add("Today", Today);

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SegmentValueNo", @" AND t.SegmentValueNo = @SegmentValueNo", SegmentValueNo);

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

        #region //GetItemType -- 取得品號類別資料 -- Shintokuro 2023.03.06
        public string GetItemType(int ItemTypeId, string ItemTypeNo, string ItemTypeName, string StartDate, string EndDate, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.ItemTypeId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" ,a.ParentItemTypeId, a.ItemTypeNo, a.ItemTypeName, a.ItemTypeDesc, ISNULL(FORMAT(a.EffectiveDate, 'yyyy-MM-dd'), '') EffectiveDate, ISNULL(FORMAT(a.InactiveDate, 'yyyy-MM-dd'), '') InactiveDate
                           , a.Status
                           , a.ItemTypeNo + '-' + a.ItemTypeName ItemTypeShow
                           , b.ItemTypeDefaultId
                          ";
                    sqlQuery.mainTables =
                        @"FROM PDM.ItemType a
                          LEFT JOIN PDM.ItemTypeDefault b on a.ItemTypeId = b.ItemTypeId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ItemTypeId", @" AND a.ItemTypeId = @ItemTypeId", ItemTypeId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ItemTypeNo", @" AND a.ItemTypeNo LIKE '%' + @ItemTypeNo + '%'", ItemTypeNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ItemTypeName", @" AND a.ItemTypeName LIKE '%' + @ItemTypeName + '%'", ItemTypeName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.EffectiveDate <= @StartDate ", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.InactiveDate >= @EndDate ", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));


                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ItemTypeNo ASC";
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

        #region //GetItemTypeSimple -- 取得品號類別資料(無換頁) -- Shintokuro 2023.03.06
        public string GetItemTypeSimple(string Today)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sql = @"SELECT test.* FROM
                            (SELECT  a.ItemTypeId,a.ItemTypeNo, a.ItemTypeName, a.[Status]
                            , a.ItemTypeNo + ':' + a.ItemTypeName ItemTypeShow
                            ,(CASE WHEN a.EffectiveDate is not null 
		                            THEN (
			                            case when CONVERT(varchar(10),a.EffectiveDate, 120)  <= @Today
			                            THEN 1
			                            ELSE 2
			                            END ) 
		                            ELSE 3
		                            END) as EffectiveDateStatus
                            ,(CASE WHEN a.InactiveDate is not null 
		                            THEN (
			                            case when CONVERT(varchar(10),a.InactiveDate, 120)  >= @Today
			                            THEN 1 
			                            ELSE 2 
			                            END ) 
		                            ELSE 3
		                            END) as InactiveDateStatus
                            FROM PDM.ItemType a
                            WHERE a.CompanyId = @CompanyId) test
                            WHERE test.EffectiveDateStatus !=2
							AND test.InactiveDateStatus !=2
                            AND test.[Status]='A'";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    dynamicParameters.Add("Today", Today);


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

        #region //GetItemTypeSegment -- 取得品號類別結構資料 -- Shintokuro 2023.03.07
        public string GetItemTypeSegment(int ItemTypeSegmentId, int ItemTypeId, string SegmentType, string StartDate, string EndDate, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sqlQuery.mainKey = "a.ItemTypeSegmentId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" ,a.SortNumber, a.SegmentType, a.SegmentValue, a.SuffixCode, ISNULL(FORMAT(a.EffectiveDate, 'yyyy-MM-dd'), '') EffectiveDate, ISNULL(FORMAT(a.InactiveDate, 'yyyy-MM-dd'), '') InactiveDate
                           , a.Status
                          ";
                    sqlQuery.mainTables =
                        @"FROM PDM.ItemTypeSegment a
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND a.ItemTypeId = @ItemTypeId ";
                    dynamicParameters.Add("ItemTypeId", ItemTypeId);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ItemTypeSegmentId", @" AND a.ItemTypeSegmentId = @ItemTypeSegmentId", ItemTypeSegmentId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SegmentType", @" AND a.SegmentType = @SegmentType", SegmentType);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.EffectiveDate >= @StartDate ", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.InactiveDate <= @EndDate ", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));


                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.SortNumber,a.ItemTypeSegmentId DESC";
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

        #region //GetItemTypeDefault -- 取得品號類別預設資料-- Shintokuro 2023-04.17
        public string GetItemTypeDefault(int ItemTypeId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    if (ItemTypeId <= 0) throw new SystemException("【品號類別預設資料管理】－品號類別不能為空!!!");

                    DynamicParameters dynamicParameters = new DynamicParameters();

                    sql = @"SELECT  a.ItemTypeId,a.ItemTypeDefaultId, a.TypeOne, a.TypeTwo, a.TypeThree, a.TypeFour, 
                            a.InventoryId, a.RequisitionInventoryId, a.InventoryUomId, a.ItemAttribute, 
                            a.MeasureType, a.PurchaseUomId, a.SaleUomId, a.LotManagement, a.InventoryManagement, 
                            a.MtlModify, a.BondedStore, a.OverReceiptManagement, a.OverDeliveryManagement, 
                            ISNULL(FORMAT(a.EffectiveDate, 'yyyy-MM-dd'), '') EffectiveDate, ISNULL(FORMAT(a.ExpirationDate, 'yyyy-MM-dd'), '') ExpirationDate
                            , a.MtlItemDesc, a.MtlItemRemark
                            FROM PDM.ItemTypeDefault a
                            WHERE a.ItemTypeId = @ItemTypeId";
                    dynamicParameters.Add("ItemTypeId", ItemTypeId);

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

        #region//GetDfmItemCategory --取得DFM項目種類代碼-- Shintokuro 2023-08.28
        public string GetDfmItemCategory(int DfmItemCategoryId, int ModeId, string DfmItemCategoryNo, string DfmItemCategoryName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.DfmItemCategoryId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ModeId, a.DfmItemCategoryNo, a.DfmItemCategoryName, a.DfmItemCategoryDesc, a.Status
                          ,b.ModeName
                         ";
                    sqlQuery.mainTables =
                        @" FROM PDM.DfmItemCategory a
                           INNER JOIN MES.ProdMode b on a.ModeId = b.ModeId
                        ";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND b.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DfmItemCategoryId", @" AND a.DfmItemCategoryId= @DfmItemCategoryId", DfmItemCategoryId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DfmItemCategoryId", @" AND a.ModeId = @ModeId", ModeId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DfmItemCategoryNo", @" AND a.DfmItemCategoryNo LIKE '%' + @DfmItemCategoryNo + '%'", DfmItemCategoryNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DfmItemCategoryName", @" AND a.DfmItemCategoryName LIKE '%' + @DfmItemCategoryName + '%'", DfmItemCategoryName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.DfmItemCategoryId DESC";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;
                    sqlQuery.distinct = false;
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

        #region//GetDfmCategoryTemplate --取得DFM項目種類樣板-- Shintokuro 2023-08.29
        public string GetDfmCategoryTemplate(int DfmCtId, int DfmItemCategoryId, string DataType)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    dynamicParameters = new DynamicParameters();

                    switch (DataType)
                    {
                        case "Material":
                            sql = @"SELECT a.DfmCtId, a.DfmItemCategoryId, a.QiMaterialTemplate
                                    FROM PDM.DfmCategoryTemplate a 
                                    WHERE a.DfmItemCategoryId = @DfmItemCategoryId";
                            dynamicParameters.Add("DfmItemCategoryId", DfmItemCategoryId);
                            break;
                        case "DfmQiOSP":
                            throw new SystemException("【委外加工】尚未開發，敬請期待~~");
                            break;
                        case "DfmQiProcess":
                            sql = @"SELECT a.DfmCtId, a.DfmItemCategoryId, a.QiProcessTemplate
                                    FROM PDM.DfmCategoryTemplate a 
                                    WHERE a.DfmItemCategoryId = @DfmItemCategoryId";
                            dynamicParameters.Add("DfmItemCategoryId", DfmItemCategoryId);
                            break;
                        default:
                            throw new SystemException("資料類別異常，請重新確認!!");
                            break;
                    }
                   
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

        #region//GetDfmMaterial --取得DFM物料資料
        public string GetDfmMaterial(int DfmMaterialId, string DfmMaterialNo, string DfmMaterialName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.DfmMaterialId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.DfmMaterialNo, a.DfmMaterialName, a.DfmMaterialDesc, a.DfmMaterialMoney, a.Status
                          , (a.DfmMaterialNo + '-'+ a.DfmMaterialName+ '-'+ CAST(a.DfmMaterialMoney AS NVARCHAR)) DfmMaterialFullItemNo
                         ";
                    sqlQuery.mainTables =
                        @" FROM PDM.DfmMaterial a
                        ";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DfmMaterialId", @" AND a.DfmMaterialId= @DfmMaterialId", DfmMaterialId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DfmMaterialNo", @" AND a.DfmMaterialNo LIKE '%' + @DfmMaterialNo + '%'", DfmMaterialNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DfmMaterialName", @" AND a.DfmMaterialName LIKE '%' + @DfmMaterialName + '%'", DfmMaterialName);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.DfmMaterialId DESC";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;
                    sqlQuery.distinct = false;
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

        #region//GetDfmCategoryMaterial --取得DFM項目種類物料
        public string GetDfmCategoryMaterial(string DfmMaterialDesc, int DfmMaterialId, int DfmItemCategoryId, string DfmMaterialNo, string DfmMaterialName, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.DfmCategoryMaterialId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",d.CompanyId
                          ,a.DfmMaterialId
                          ,a.DfmItemCategoryId
                          ,c.DfmItemCategoryName
                          ,b.DfmMaterialNo, b.DfmMaterialName, b.DfmMaterialDesc, b.DfmMaterialMoney, a.Status
                          ,(b.DfmMaterialNo + '-'+ b.DfmMaterialName+ '-'+ CAST(b.DfmMaterialMoney AS NVARCHAR)) DfmMaterialFullItemNo";
                    sqlQuery.mainTables =
                        @"FROM PDM.DfmCategoryMaterial a
                          INNER JOIN PDM.DfmMaterial b on a.DfmMaterialId = b.DfmMaterialId
                          INNER JOIN PDM.DfmItemCategory c on c.DfmItemCategoryId = a.DfmItemCategoryId
                          INNER JOIN MES.ProdMode d on c.ModeId = d.ModeId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND d.CompanyId = @CompanyId AND a.DfmItemCategoryId=@DfmItemCategoryId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    dynamicParameters.Add("DfmItemCategoryId", DfmItemCategoryId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DfmMaterialId", @" AND b.DfmMaterialId= @DfmMaterialId", DfmMaterialId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DfmItemCategoryId", @" AND a.DfmItemCategoryId = @DfmItemCategoryId", DfmItemCategoryId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DfmMaterialNo", @" AND b.DfmMaterialNo LIKE '%' + @DfmMaterialNo + '%'", DfmMaterialNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DfmMaterialName", @" AND b.DfmMaterialName LIKE '%' + @DfmMaterialName + '%'", DfmMaterialName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DfmMaterialDesc", @" AND b.DfmMaterialDesc LIKE '%' + @DfmMaterialDesc + '%'", DfmMaterialDesc);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));

                    sqlQuery.conditions = queryCondition;
                    //sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.DfmItemCategoryId DESC";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;
                    sqlQuery.distinct = true;
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

        #region//GetMoldingConditions --取得成型產品生產條件
        public string GetMoldingConditions(int McId, string MtlItemNo, string MtlItemName, string Material
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
                        @",b.MtlItemId,b.MtlItemNo,b.MtlItemName,b.MtlItemSpec
                          ,a.Material,b1.MtlItemNo CellMtlItemNo,b1.MtlItemName CellMtlItemName,b1.MtlItemSpec CellMtlItemSpec
                          ,a.CycleTime,a.MoldWeight,a.Cavity,a.ProcessingFee,a.UnitPrice,a.TheoreticalQty
                          ";
                    sqlQuery.mainTables =
                        @"FROM PDM.MoldingConditions a
                          INNER JOIN PDM.MtlItem b on a.MtlItemId = b.MtlItemId
                          LEFT JOIN PDM.MtlItem b1 on a.CellMtlItemId = b1.MtlItemId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND b.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "McId", @" AND a.McId= @McId", McId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemNo", @" AND b.MtlItemNo LIKE '%' + @MtlItemNo + '%'", MtlItemNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemName", @" AND b.MtlItemName LIKE '%' + @MtlItemName + '%'", MtlItemName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Material", @" AND a.Material LIKE '%' + @Material + '%'", Material);

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.McId";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;
                    //sqlQuery.distinct = true;


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

        #region //GetFileInfoById -- 取得檔案資料 -- Shintokuro 2025-04-29
        public BmFileInfo GetFileInfoById(int FileId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT FileId, FileName, FileContent, FileExtension, FileSize, ClientIP, Source, DeleteStatus
                            FROM BAS.[File]
                            WHERE FileId = @FileId";
                    dynamicParameters.Add("FileId", FileId);

                    var result = sqlConnection.Query<BmFileInfo>(sql, dynamicParameters).FirstOrDefault();

                    return result;
                }
            }
            catch (Exception e)
            {
                logger.Error(e.Message);
            }

            return new BmFileInfo();
        }
        #endregion

        #region //GetProductionStockInDetail -- 取得ERP生產入庫明細表 -- Shintokuro 2025-05-07
        public string GetProductionStockInDetail(string MtlItemNo, string StardDay, string EndDay)
        {
            try
            {
                string MESCompanyNo = "";
                string ErpDb = "";

                StardDay = DateTime.Parse(StardDay).ToString("yyyyMMdd");
                EndDay = DateTime.Parse(EndDay).ToString("yyyyMMdd");

                List<StockInDetail> stockInDetail = new List<StockInDetail>();


                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //公司別資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.ErpNo, a.ErpDb
                            FROM BAS.Company a
                            WHERE CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    var resultCompany = sqlConnection.Query(sql, dynamicParameters);
                    if (resultCompany.Count() <= 0) throw new SystemException("【公司別】資料錯誤!");

                    foreach (var item in resultCompany)
                    {
                        ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        MESCompanyNo = item.ErpNo;
                        ErpDb = item.ErpDb;
                    }
                    #endregion

                }
                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                {
                    #region //撈取線別成本費用
                    // 往前一個月
                    DateTime date = DateTime.ParseExact(StardDay, "yyyyMMdd", null);
                    DateTime previousMonth = date.AddMonths(-1);
                    // 轉為 yyyyMM 格式
                    string yearMonth = previousMonth.ToString("yyyyMM");
                    #endregion

                    string MtlItemStr = "";

                    if (MtlItemNo.Length > 0)
                    {
                        MtlItemStr = $"AND b.TG004 LIKE '%' + @MtlItemNo + '%'";
                    }

                    #region //生產入庫明細表
                    dynamicParameters = new DynamicParameters();
                    sql = $@"
                            SELECT 
                            a.TF003 ReceiptDate,
                            b.TG014+'-'+b.TG015 ErpFull
                            ,b.TG004 MtlItemNo,b.TG005 MtlItemName,b.TG006 MtlItemSpec
                            ,SUM(b.TG011) ReceiptQty
                            ,SUM(b.TG012) ScriptQty
                            ,SUM(b.TG013) AcceptQty
                            ,ROUND((SUM(b.TG013) * 100.0 / NULLIF(SUM(b.TG011), 0)), 1) AS PassRate
                            ,ROUND((SUM(b.TG012) * 100.0 / NULLIF(SUM(b.TG011), 0)), 1) AS DefectRate
                            ,y.ManCost,y.MachineCost
                            FROM MOCTF a
                            INNER JOIN MOCTG b ON a.TF001=b.TG001 AND a.TF002=b.TG002  
                            OUTER APPLY(
                                SELECT  SUM(ROUND(x1.MB005 * y5.UnitManCost, 2)) ManCost, SUM(ROUND(x1.MB006 * y5.UnitMachineCost, 2)) MachineCost
                                FROM CSTMB x1
                                OUTER APPLY(
                                    SELECT  LTRIM(RTRIM(y1.MC026)) UnitManCost,LTRIM(RTRIM(y1.MC027)) UnitMachineCost 
                                    FROM CSTMC y1
                                    WHERE y1.MC002 = @YearMonth
                                    AND y1.MC001 = x1.MB001
                                ) y5
                                WHERE x1.MB003 = b.TG014
                                AND x1.MB004 = b.TG015
                                AND x1.MB002 = a.TF003
                            ) y
                            WHERE 1=1 
                            AND ((a.TF003 Between @StardDay and @EndDay)  AND a.TF006='Y') 
                            {MtlItemStr}
                            GROUP BY 
                            a.TF003,b.TG004,b.TG005,b.TG006,b.TG014+'-'+b.TG015,y.ManCost,y.MachineCost
                            ";
                    dynamicParameters.Add("StardDay", StardDay);
                    dynamicParameters.Add("EndDay", EndDay);
                    dynamicParameters.Add("YearMonth", yearMonth);
                    dynamicParameters.Add("MtlItemNo", MtlItemNo);
                    stockInDetail = sqlConnection.Query<StockInDetail>(sql, dynamicParameters).ToList();
                    #endregion
                }

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //撈取品號ID
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT b.MtlItemNo,b1.MtlItemNo CellMtlItemNo,b1.MtlItemName CellMtlItemName,b1.MtlItemSpec CellMtlItemSpec
                            ,a.Material,a.CycleTime,a.MoldWeight		
                            ,a.Cavity,a.ProcessingFee ,a.UnitPrice,a.TheoreticalQty
                            FROM PDM.MoldingConditions a
                            INNER JOIN PDM.MtlItem b on a.MtlItemId = b.MtlItemId
                            LEFT JOIN PDM.MtlItem b1 on a.CellMtlItemId = b1.MtlItemId
                            WHERE b.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    List<StockInDetail> resultMtlitems = sqlConnection.Query<StockInDetail>(sql, dynamicParameters).ToList();
                    stockInDetail = stockInDetail
                    .Join(resultMtlitems,
                          x => x.MtlItemNo,
                          y => y.MtlItemNo,
                          (x, y) =>
                          {
                              x.CellMtlItemNo = y.CellMtlItemNo;
                              x.CellMtlItemName = y.CellMtlItemName;
                              x.CellMtlItemSpec = y.CellMtlItemSpec;
                              x.Material = y.Material;
                              x.CycleTime = y.CycleTime;
                              x.MoldWeight = y.MoldWeight;
                              x.Cavity = y.Cavity;
                              x.ProcessingFee = y.ProcessingFee;
                              x.UnitPrice = y.UnitPrice;
                              x.TheoreticalQty = y.TheoreticalQty;
                              return x;
                          })
                    .ToList();
                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = stockInDetail
                    });
                    #endregion
                }
            }
            catch (Exception e)
            {
                logger.Error(e.Message);
            }

            return jsonResponse.ToString();

        }
        #endregion

        #region //GetTemporaryOrderOutInDetail -- ERP暫出入單明細表 -- Jean 2025-05-07
        public string GetTemporaryOrderOutInDetail(string StardDay, string EndDay)
        {
            try
            {
                string MESCompanyNo = "";
                string ErpDb = "";

                StardDay = DateTime.Parse(StardDay).ToString("yyyyMMdd");
                EndDay = DateTime.Parse(EndDay).ToString("yyyyMMdd");


                List<ItemCatDept> itemCatDept = new List<ItemCatDept>();
                List<TempOrderDetail> tempOrderDetail = new List<TempOrderDetail>();


                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //公司別資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.ErpNo, a.ErpDb
                            FROM BAS.Company a
                            WHERE CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    var resultCompany = sqlConnection.Query(sql, dynamicParameters);
                    if (resultCompany.Count() <= 0) throw new SystemException("【公司別】資料錯誤!");

                    foreach (var item in resultCompany)
                    {
                        ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        MESCompanyNo = item.ErpNo;
                        ErpDb = item.ErpDb;
                    }
                    #endregion

                    #region //撈取品號類別責任部門資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.TypeTwo,a.TypeThree,a.InventoryNo,a.CatDept
                            FROM PDM.ItemCatDept a ";
                    itemCatDept = sqlConnection.Query<ItemCatDept>(sql, dynamicParameters).ToList();

                    #endregion
                }

                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                {
                    #region //暫出入單明細表
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT FORMAT(CONVERT(DATE, a.TF003, 112), 'yyyy/MM/dd') TF003
                            ,a.TF001+'-'+TF002 as DocNum
                            ,FORMAT(CONVERT(DATE, a.TF024, 112), 'yyyy/MM/dd') TF024
                            ,FORMAT(CONVERT(DATE, a.TF041, 112), 'yyyy/MM/dd') TF041
                            ,a.TF005,a.TF006
                            ,a.TF007,d.ME002,a.TF008,e.MV002,a.TF009,f.MB002,a.TF011
                            ,(CASE WHEN a.TF010='1' THEN N'應稅內含'
                                WHEN a.TF010='2' THEN N'應稅外加'
                                WHEN a.TF010='3' THEN N'零稅率'
                                WHEN a.TF010='4' THEN N'免稅'
                                WHEN a.TF010='9' THEN N'不計稅' END) AS TaxType
                            ,a.TF012,a.TF013,a.TF014,b.TG004,b.TG005,b.TG006,b.TG007,m.MC002 TempOut,n.MC002 TypeTempIn,b.TG008,b.TG027,b.TG035,b.TG036
                            ,CASE WHEN a.TF001 = '1301' THEN b.TG009 ELSE 0 END AS TempOutQty
                            ,CASE WHEN a.TF001 = '1301' THEN b.TG020 ELSE 0 END AS TempOutSaleQty
                            ,CASE WHEN a.TF001 = '1301' THEN b.TG021 ELSE 0 END AS TempOutReturnQty
                            ,CASE WHEN a.TF001 = '1302' THEN b.TG009 ELSE 0 END AS TempInQty
                            ,CASE WHEN a.TF001 = '1302' THEN b.TG020 ELSE 0 END AS TempInSaleQty
                            ,CASE WHEN a.TF001 = '1302' THEN b.TG021 ELSE 0 END AS TempInReturnQty
                            ,CASE WHEN a.TF001 = '1301' THEN b.TG028 ELSE 0 END AS TempOutPackQty
                            ,CASE WHEN a.TF001 = '1301' THEN b.TG029 ELSE 0 END AS TempOutPackSaleQty
                            ,CASE WHEN a.TF001 = '1301' THEN b.TG030 ELSE 0 END AS TempOutPackReturnQty
                            ,CASE WHEN a.TF001 = '1302' THEN b.TG028 ELSE 0 END AS TempInPackQty
                            ,CASE WHEN a.TF001 = '1302' THEN b.TG029 ELSE 0 END AS TempInPackSaleQty
                            ,CASE WHEN a.TF001 = '1302' THEN b.TG030 ELSE 0 END AS TempInPackReturnQty
                            ,CASE WHEN a.TF001 = '1301' THEN b.TG045 ELSE 0 END AS TempOutSpareQty
                            ,CASE WHEN a.TF001 = '1301' THEN b.TG046 ELSE 0 END AS TempOutSpareSaleQty
                            ,CASE WHEN a.TF001 = '1301' THEN b.TG048 ELSE 0 END AS TempOutSpareReturnQty
                            ,CASE WHEN a.TF001 = '1302' THEN b.TG045 ELSE 0 END AS TempInSpareQty
                            ,CASE WHEN a.TF001 = '1302' THEN b.TG046 ELSE 0 END AS TempInSpareSaleQty
                            ,CASE WHEN a.TF001 = '1302' THEN b.TG048 ELSE 0 END AS TempInSpareReturnQty
                            ,CASE WHEN a.TF001 = '1301' THEN b.TG045 ELSE 0 END AS TempOutSparePackQty
                            ,CASE WHEN a.TF001 = '1301' THEN b.TG046 ELSE 0 END AS TempOutSparePackSaleQty
                            ,CASE WHEN a.TF001 = '1301' THEN b.TG048 ELSE 0 END AS TempOutSparePackReturnQty
                            ,CASE WHEN a.TF001 = '1302' THEN b.TG045 ELSE 0 END AS TempInSparePackQty
                            ,CASE WHEN a.TF001 = '1302' THEN b.TG046 ELSE 0 END AS TempInSparePackSaleQty
                            ,CASE WHEN a.TF001 = '1302' THEN b.TG048 ELSE 0 END AS TempInSparePackReturnQty
                            ,b.TG010,b.TG011,b.TG024
                            ,CASE WHEN b.TG024='N' THEN N'未结案' WHEN b.TG024='Y' THEN N'已结案' WHEN b.TG024='N' THEN N'指定结案' END AS CloseCode
                            ,b.TG031,b.TG012
                            ,CASE WHEN a.TF001 = '1301' THEN b.TG013 ELSE 0 END AS TempOutAmount
                            ,CASE WHEN a.TF001 = '1302' THEN b.TG013 ELSE 0 END AS TempInAmount
                            ,b.TG017,b.TG025,b.TG026
                            ,b.TG014+'-'+b.TG015+'-'+b.TG016 AS SourceDocNum
                            ,b.TG019,b.TG018,i.NB002
                            ,k.MB006,k.MB007
                            ,m.MC002 TempOut,n.MC002 TypeTempIn
                            ,x.MA003 TypeTwoName02,y.MA003 TypeTwoName03
                            FROM INVTF AS a
                            LEFT JOIN INVTG AS b ON b.TG001=a.TF001 AND b.TG002=a.TF002 
                            LEFT JOIN CMSMQ AS c ON c.MQ001=a.TF001 
                            LEFT JOIN CMSME AS d ON d.ME001=a.TF007 
                            LEFT JOIN CMSMV AS e ON e.MV001=a.TF008 
                            LEFT JOIN CMSMB AS f ON f.MB001=a.TF009 
                            LEFT JOIN INVMC AS g ON b.TG004=g.MC001 AND b.TG007=g.MC002 
                            LEFT JOIN INVMC AS h ON b.TG004=h.MC001 AND b.TG008=h.MC002 
                            LEFT JOIN CMSNB AS i ON b.TG018=i.NB001 
                            LEFT JOIN CMSMF AS j ON j.MF001= 'NT$'
                            LEFT JOIN INVMB AS k ON k.MB001=b.TG004
                            LEFT JOIN CMSMC AS m ON m.MC001=b.TG007 
                            LEFT JOIN CMSMC AS n ON n.MC001=b.TG008
                            LEFT JOIN INVMA AS x ON x.MA002=k.MB006
                            LEFT JOIN INVMA AS y ON y.MA002=k.MB007
                            WHERE ((c.MQ003='13' OR c.MQ003='14')  and  (a.TF003 Between @StardDay and @EndDay)  AND a.TF020='Y' AND b.TG024='N') 
                            ORDER BY a.TF003,a.TF001,a.TF002,b.TG003  ";
                    dynamicParameters.Add("StardDay", StardDay);
                    dynamicParameters.Add("EndDay", EndDay);
                    tempOrderDetail = sqlConnection.Query<TempOrderDetail>(sql, dynamicParameters).ToList();

                    //tempOrderDetail = tempOrderDetail
                    //.Select(order =>
                    //{
                    //    ItemCatDept matchingCatDept = null;
                    //    var specialMb007Values = new HashSet<string> { "301", "302", "303", "308", "309", "313", "315" };

                    //    if (order.MB006 == "2002")
                    //    {
                    //        当MB006 = 2002时，匹配TypeTwo和InventoryNo
                    //         matchingCatDept = itemCatDept.FirstOrDefault(cd =>
                    //             cd.TypeTwo == order.MB006 &&
                    //             cd.InventoryNo == order.TG007);
                    //    }
                    //    else if (order.MB006 == "2001" || order.MB006 == "2003")
                    //    {
                    //        当MB006 = 2001 / 2003时，根据MB007的值决定匹配条件
                    //        if (specialMb007Values.Contains(order.MB007))
                    //        {
                    //            MB007是特殊值时，匹配TypeTwo和InventoryNo
                    //           matchingCatDept = itemCatDept.FirstOrDefault(cd =>
                    //               cd.TypeThree == order.MB007 &&
                    //               cd.InventoryNo == order.TG007);
                    //        }
                    //        else
                    //        {
                    //            否则仅匹配InventoryNo
                    //           matchingCatDept = itemCatDept.FirstOrDefault(cd =>
                    //               cd.InventoryNo == order.TG007);
                    //        }
                    //    }

                    //    order.CatDept = matchingCatDept?.CatDept;
                    //    return order;
                    //})
                    //.ToList();

                    tempOrderDetail = tempOrderDetail
                    .Select(order =>
                    {
                        ItemCatDept matchingCatDept = null;

                        // 统一按照InventoryNo > TypeThree > TypeTwo的优先级进行匹配，不区分MB006和TG007
                        // 1. 优先匹配三个条件
                        matchingCatDept = itemCatDept.FirstOrDefault(cd =>
                            cd.InventoryNo == order.TG007 &&
                            cd.TypeThree == order.MB007 &&
                            cd.TypeTwo == order.MB006);

                        // 2. 其次匹配InventoryNo + TypeThree
                        if (matchingCatDept == null)
                        {
                            matchingCatDept = itemCatDept.FirstOrDefault(cd =>
                                cd.InventoryNo == order.TG007 &&
                                cd.TypeThree == order.MB007);
                        }

                        // 3. 再次匹配InventoryNo + TypeTwo
                        if (matchingCatDept == null)
                        {
                            matchingCatDept = itemCatDept.FirstOrDefault(cd =>
                                cd.InventoryNo == order.TG007 &&
                                cd.TypeTwo == order.MB006);
                        }

                        // 4. 只匹配InventoryNo
                        if (matchingCatDept == null)
                        {
                            matchingCatDept = itemCatDept.FirstOrDefault(cd =>
                                cd.InventoryNo == order.TG007);
                        }

                        order.CatDept = matchingCatDept?.CatDept;
                        return order;
                    })
                    .ToList();



                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = tempOrderDetail
                    });
                    #endregion


                }
            }
            catch (Exception e)
            {
                logger.Error(e.Message);
            }

            return jsonResponse.ToString();

        }
        #endregion

        #region //GetTempOutReturnRecordDetail -- ERP暫出歸還單明細表 -- Jean 2025-05-07
        public string GetTempOutReturnRecordDetail(string StardDay, string EndDay)
        {
            try
            {
                string MESCompanyNo = "";
                string ErpDb = "";

                StardDay = DateTime.Parse(StardDay).ToString("yyyyMMdd");
                EndDay = DateTime.Parse(EndDay).ToString("yyyyMMdd");


                List<ItemCatDept> itemCatDept = new List<ItemCatDept>();
                List<TempReturnOrderDetail> tempReturnOrderDetail = new List<TempReturnOrderDetail>();


                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //公司別資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.ErpNo, a.ErpDb
                            FROM BAS.Company a
                            WHERE CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    var resultCompany = sqlConnection.Query(sql, dynamicParameters);
                    if (resultCompany.Count() <= 0) throw new SystemException("【公司別】資料錯誤!");

                    foreach (var item in resultCompany)
                    {
                        ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        MESCompanyNo = item.ErpNo;
                        ErpDb = item.ErpDb;
                    }
                    #endregion

                    #region //撈取品號類別責任部門資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.TypeTwo,a.TypeThree,a.InventoryNo,a.CatDept
                            FROM PDM.ItemCatDept a ";
                    itemCatDept = sqlConnection.Query<ItemCatDept>(sql, dynamicParameters).ToList();

                    #endregion
                }

                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                {
                    #region //暫出入單明細表
                    dynamicParameters = new DynamicParameters();
                    sql = @"
                            SELECT FORMAT(CONVERT(DATE, a.TH003, 112), 'yyyy/MM/dd') TH003
                                ,a.TH001+'-'+a.TH002 as DocNum
                                ,FORMAT(CONVERT(DATE, a.TH023, 112), 'yyyy/MM/dd') TH023
                                ,TH005,TH006,a.TH007,d.ME002,a.TH008,e.MV002,a.TH009,f.MB002,a.TH011,a.TH010,a.TH012,a.TH013
                                ,b.TI004,b.TI005,b.TI006,b.TI007,m.MC002 TempOut,b.TI008,n.MC002 TypeTempIn,b.TI026,b.TI027
                                ,b.TI009,b.TI023,b.TI023,b.TI035,b.TI036
                                ,b.TI010,b.TI011,b.TI024,b.TI012,b.TI013,b.TI017,b.TI018,b.TI019
                                ,b.TI014+'-'+b.TI015+'-'+b.TI016 AS SourceDocNum
                                ,b.TI021,b.TI020
                                ,k.MB006,x.MA003 TypeTwoName02,k.MB007,ISNULL(y.MA003,'') TypeTwoName03
                                FROM INVTH AS a
                                LEFT JOIN INVTI AS b ON b.TI001=a.TH001 AND b.TI002=a.TH002 
                                LEFT JOIN CMSMQ AS c ON c.MQ001=a.TH001 
                                LEFT JOIN CMSME AS d ON d.ME001=a.TH007 
                                LEFT JOIN CMSMV AS e ON e.MV001=a.TH008 
                                LEFT JOIN CMSMB AS f ON f.MB001=a.TH009 
                                LEFT JOIN INVMC AS g ON b.TI004=g.MC001 AND b.TI007=g.MC002 
                                LEFT JOIN INVMC AS h ON b.TI004=h.MC001 AND b.TI008=h.MC002 
                                LEFT JOIN CMSNB AS i ON i.NB001=b.TI020 
                                LEFT JOIN CMSMF AS j ON j.MF001= 'NT$'
                                LEFT JOIN INVMB AS k ON k.MB001=b.TI004
                                LEFT JOIN CMSMC AS m ON m.MC001=b.TI007
                                LEFT JOIN CMSMC AS n ON n.MC001=b.TI008
                                LEFT JOIN INVMA AS x ON x.MA002=k.MB006
                                LEFT JOIN INVMA AS y ON y.MA002=k.MB007
                            WHERE ((c.MQ003='15' OR c.MQ003='16')  AND  (a.TH003 Between @StardDay and @EndDay)  AND a.TH020='Y') 
                            ORDER BY a.TH003,a.TH001,a.TH002,b.TI003  ";
                    dynamicParameters.Add("StardDay", StardDay);
                    dynamicParameters.Add("EndDay", EndDay);
                    tempReturnOrderDetail = sqlConnection.Query<TempReturnOrderDetail>(sql, dynamicParameters).ToList();


                    //tempReturnOrderDetail = tempReturnOrderDetail
                    //.Select(order =>
                    //{
                    //    ItemCatDept matchingCatDept = null;
                    //    var specialMb007Values = new HashSet<string> { "301", "302", "303", "308", "309", "313", "315" };

                    //    if (order.MB006 == "2002")
                    //    {
                    //        当MB006 = 2002时，匹配TypeTwo和InventoryNo
                    //         matchingCatDept = itemCatDept.FirstOrDefault(cd =>
                    //             cd.TypeTwo == order.MB006 &&
                    //             cd.InventoryNo == order.TI008);
                    //    }
                    //    else if (order.MB006 == "2001" || order.MB006 == "2003")
                    //    {
                    //        当MB006 = 2001 / 2003时，根据MB007的值决定匹配条件
                    //        if (specialMb007Values.Contains(order.MB007))
                    //        {
                    //            MB007是特殊值时，匹配TypeTwo和InventoryNo
                    //           matchingCatDept = itemCatDept.FirstOrDefault(cd =>
                    //               cd.TypeThree == order.MB007 &&
                    //               cd.InventoryNo == order.TI008);
                    //        }
                    //        else
                    //        {
                    //            否则仅匹配InventoryNo
                    //           matchingCatDept = itemCatDept.FirstOrDefault(cd =>
                    //               cd.InventoryNo == order.TI008);
                    //        }
                    //    }

                    //    order.CatDept = matchingCatDept?.CatDept;
                    //    return order;
                    //})
                    //.ToList();
                        
                    tempReturnOrderDetail = tempReturnOrderDetail
                    .Select(order =>
                    {
                        ItemCatDept matchingCatDept = null;

                        // 统一按照InventoryNo > TypeThree > TypeTwo的优先级进行匹配，不区分MB006和TG007
                        // 1. 优先匹配三个条件
                        matchingCatDept = itemCatDept.FirstOrDefault(cd =>
                            cd.InventoryNo == order.TI008 &&
                            cd.TypeThree == order.MB007 &&
                            cd.TypeTwo == order.MB006);

                        // 2. 其次匹配InventoryNo + TypeThree
                        if (matchingCatDept == null)
                        {
                            matchingCatDept = itemCatDept.FirstOrDefault(cd =>
                                cd.InventoryNo == order.TI008 &&
                                cd.TypeThree == order.MB007);
                        }

                        // 3. 再次匹配InventoryNo + TypeTwo
                        if (matchingCatDept == null)
                        {
                            matchingCatDept = itemCatDept.FirstOrDefault(cd =>
                                cd.InventoryNo == order.TI008 &&
                                cd.TypeTwo == order.MB006);
                        }

                        // 4. 只匹配InventoryNo
                        if (matchingCatDept == null)
                        {
                            matchingCatDept = itemCatDept.FirstOrDefault(cd =>
                                cd.InventoryNo == order.TI008);
                        }

                        order.CatDept = matchingCatDept?.CatDept;
                        return order;
                    })
                    .ToList();

                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = tempReturnOrderDetail
                    });
                    #endregion


                }
            }
            catch (Exception e)
            {
                logger.Error(e.Message);
            }

            return jsonResponse.ToString();

        }
        #endregion

        #region//GetItemCategoryDept --取得品號類別責任部門 --Jean 2025-05-24 --
        public string GetItemCategoryDept(int ItemCatDeptId, string TypeTwo, string TypeThree, string InventoryNo, string CatDept
            , string OrderBy, int PageIndex, int PageSize) 
        {
            try
            {
                    List<ItemCatDepts> itemCatDepts = new List<ItemCatDepts>();
                    List<MtlItemCategory> mtlItemCategory = new List<MtlItemCategory>();
                    List<MtlItemInventory> mtlItemInventorys = new List<MtlItemInventory>();
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //確認公司別DB
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.ErpDb
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

                        #region //撈取類別二、三
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT LTRIM(RTRIM(MA002)) MtlItemTypeNo, LTRIM(RTRIM(MA003)) MtlItemTypeName
                            FROM INVMA
                            WHERE 1=1";

                        mtlItemCategory = sqlConnection2.Query<MtlItemCategory>(sql, dynamicParameters).ToList();

                        #endregion

                        #region //撈取库别
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT LTRIM(RTRIM(MC001)) InventoryNo, LTRIM(RTRIM(MC002)) InventoryName
                                FROM CMSMC 
                                WHERE 1=1";

                        mtlItemInventorys = sqlConnection2.Query<MtlItemInventory>(sql, dynamicParameters).ToList();



                        #endregion
                    }
                    using (SqlConnection sqlConnection2 = new SqlConnection(MainConnectionStrings))
                    {

                        #region //撈取品號類別責任部門資料
                        DynamicParameters dynamicParameters = new DynamicParameters();
                        sqlQuery.mainKey = "a.ItemCatDeptId";
                        sqlQuery.auxKey = "";
                        sqlQuery.columns =
                            @",a.TypeTwo,a.TypeThree,a.InventoryNo,a.CatDept ";
                        sqlQuery.mainTables =
                            @"FROM PDM.ItemCatDept a ";
                        sqlQuery.auxTables = "";
                        string queryCondition = @"AND 1=1";
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ItemCatDeptId", @" AND a.ItemCatDeptId = @ItemCatDeptId", ItemCatDeptId);
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TypeTwo", @" AND a.TypeTwo = @TypeTwo", TypeTwo);
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TypeThree", @" AND a.TypeThree = @TypeThree", TypeThree);
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "InventoryNo", @" AND a.InventoryNo = @InventoryNo", InventoryNo);
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CatDept", @" AND a.CatDept LIKE '%' + @CatDept + '%'", CatDept);

                        sqlQuery.conditions = queryCondition;
                        sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ItemCatDeptId ";
                        sqlQuery.pageIndex = PageIndex;
                        sqlQuery.pageSize = PageSize;
                        //sqlQuery.distinct = true;


                        itemCatDepts = BaseHelper.SqlQuery<ItemCatDepts>(sqlConnection2, dynamicParameters, sqlQuery);

                        // 关联数据（在内存中进行）
                        itemCatDepts = itemCatDepts.GroupJoin(mtlItemCategory,
                            x => x.TypeTwo,
                            y => y.MtlItemTypeNo,
                            (x, y) => { x.TypeTwoName = y.FirstOrDefault()?.MtlItemTypeName ?? ""; return x; })
                            .ToList();

                        itemCatDepts = itemCatDepts.GroupJoin(mtlItemCategory,
                            x => x.TypeThree,
                            y => y.MtlItemTypeNo,
                            (x, y) => { x.TypeThreeName = y.FirstOrDefault()?.MtlItemTypeName ?? ""; return x; })
                            .ToList();

                        itemCatDepts = itemCatDepts.GroupJoin(mtlItemInventorys,
                            x => x.InventoryNo,
                            y => y.InventoryNo,
                            (x, y) => { x.InventoryName = y.FirstOrDefault()?.InventoryName ?? ""; return x; })
                            .ToList();

                        //var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);
                        #endregion


                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            data = itemCatDepts
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

        #endregion

        #region //Add
        #region //AddMtlElement -- 品號元素資料新增 -- Kan 2022.05.31
        public string AddMtlElement(string ElementNo, string ElementName)
        {
            try
            {
                if (ElementNo.Length <= 0) throw new SystemException("【品號元素代碼】不能為空!");
                if (ElementNo.Length > 50) throw new SystemException("【品號元素代碼】長度錯誤!");
                if (ElementName.Length <= 0) throw new SystemException("【品號元素名稱】不能為空!");
                if (ElementName.Length > 100) throw new SystemException("【品號元素名稱】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷品號元素代碼是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.MtlElement
                                WHERE CompanyId = @CompanyId
                                AND ElementNo = @ElementNo";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ElementNo", ElementNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【品號元素代碼】重複，請重新輸入!");
                        #endregion

                        #region //判斷品號元素名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.MtlElement
                                WHERE CompanyId = @CompanyId
                                AND ElementName = @ElementName";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ElementName", ElementName);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【品號元素名稱】重複，請重新輸入!");
                        #endregion

                        #region //抓取目前最大序號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ISNULL(MAX(SortNumber), 0) MaxSort
                                FROM PDM.MtlElement
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        int maxSort = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).MaxSort;
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO PDM.MtlElement (CompanyId, ElementNo, ElementName
                                , SortNumber, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ElementId
                                VALUES (@CompanyId, @ElementNo, @ElementName
                                , @SortNumber, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";

                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                ElementNo,
                                ElementName,
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

        #region //AddUnitOfMeasure -- 單位基本資料新增 -- Kan 2022.06.15
        public string AddUnitOfMeasure(string UomNo, string UomName, string UomType, string UomDesc)
        {
            try
            {
                if (UomNo.Length <= 0) throw new SystemException("【單位代碼】不能為空!");
                if (UomNo.Length > 50) throw new SystemException("【單位代碼】長度錯誤!");
                if (UomName.Length <= 0) throw new SystemException("【單位名稱】不能為空!");
                if (UomName.Length > 100) throw new SystemException("【單位名稱】長度錯誤!");
                if (UomType.Length <= 0) throw new SystemException("【單位類別】不能為空!");
                if (UomDesc.Length > 100) throw new SystemException("【單位描述】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷單位代碼是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.UnitOfMeasure
                                WHERE CompanyId = @CompanyId
                                AND UomNo = @UomNo";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("UomNo", UomNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【單位代碼】重複，請重新輸入!");
                        #endregion

                        #region //判斷單位名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.UnitOfMeasure
                                WHERE CompanyId = @CompanyId
                                AND UomName = @UomName";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("UomName", UomName);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【單位名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO PDM.UnitOfMeasure (CompanyId, UomNo, UomName
                                , UomType, UomDesc, TransferStatus, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.UomId
                                VALUES (@CompanyId, @UomNo, @UomName
                                , @UomType, @UomDesc, @TransferStatus, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";

                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                UomNo,
                                UomName,
                                UomType,
                                UomDesc,
                                TransferStatus = "N", //未拋轉
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

        #region //AddUomCalculate -- 單位換算資料新增 -- Kan 2022.06.15
        public string AddUomCalculate(int UomId, int ConvertUomId, double Value)
        {
            try
            {
                if (UomId <= 0) throw new SystemException("【所屬單位】不能為空!");
                if (ConvertUomId <= 0) throw new SystemException("【換算單位】不能為空!");
                if (Value <= 0) throw new SystemException("【值】不能為0!");
                if (UomId == ConvertUomId) throw new SystemException("【所屬單位】不能與【所屬單位】相同!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷所屬單位資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.UnitOfMeasure
                                WHERE CompanyId = @CompanyId
                                AND UomId = @UomId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("UomId", UomId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("所屬單位資料錯誤!");
                        #endregion

                        #region //判斷換算單位資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.UnitOfMeasure
                                WHERE CompanyId = @CompanyId
                                AND UomId = @ConvertUomId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ConvertUomId", ConvertUomId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("換算單位資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO PDM.UomCalculate (UomId, ConvertUomId, Value, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.UomCalculateId
                                VALUES (@UomId, @ConvertUomId, @Value, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";

                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                UomId,
                                ConvertUomId,
                                Value,
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

        #region //AddMtlModel -- 品號機型資料新增 -- Kan 2022.06.17
        public string AddMtlModel(int ParentId, string MtlModelNo
            , string MtlModelName, string MtlModelDesc)
        {
            try
            {
                if (MtlModelNo.Length <= 0) throw new SystemException("【品號機型代碼】不能為空!");
                if (MtlModelNo.Length > 50) throw new SystemException("【品號機型代碼】長度錯誤!");
                if (MtlModelName.Length <= 0) throw new SystemException("【品號機型名稱】不能為空!");
                if (MtlModelName.Length > 100) throw new SystemException("【品號機型名稱】長度錯誤!");
                if (MtlModelDesc.Length > 100) throw new SystemException("【品號機型描述】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷品號機型代碼是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.MtlModel
                                WHERE CompanyId = @CompanyId
                                AND MtlModelNo = @MtlModelNo";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("MtlModelNo", MtlModelNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【品號機型代碼】重複，請重新輸入!");
                        #endregion

                        #region //判斷品號機型名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.MtlModel
                                WHERE CompanyId = @CompanyId
                                AND MtlModelName = @MtlModelName";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("MtlModelName", MtlModelName);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【品號機型名稱】重複，請重新輸入!");
                        #endregion

                        #region //抓取目前最大序號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ISNULL(MAX(MtlModelSort), 0) MaxSort
                                FROM PDM.MtlModel
                                WHERE CompanyId = @CompanyId
                                AND ParentId = @ParentId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ParentId", ParentId);

                        int maxSort = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).MaxSort;
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO PDM.MtlModel (CompanyId, ParentId, MtlModelSort
                                , MtlModelNo, MtlModelName, MtlModelDesc, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.MtlModelId
                                VALUES (@CompanyId, @ParentId, @MtlModelSort
                                , @MtlModelNo, @MtlModelName, @MtlModelDesc, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";

                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                ParentId,
                                MtlModelSort = maxSort + 1,
                                MtlModelNo,
                                MtlModelName,
                                MtlModelDesc,
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

        #region //AddItemSegment -- 編碼節段新增 -- Shintokuro 2023.03.01
        public string AddItemSegment(string SegmentNo, string SegmentName, string SegmentDesc, string EffectiveDate, string InactiveDate)
        {
            try
            {
                if (SegmentNo.Length <= 0) throw new SystemException("【編碼節段代碼】不能為空!");
                if (SegmentNo.Length > 20) throw new SystemException("【編碼節段代碼】長度錯誤!");
                if (SegmentName.Length <= 0) throw new SystemException("【編碼節段名稱】不能為空!");
                if (SegmentName.Length > 20) throw new SystemException("【編碼節段名稱】長度錯誤!");
                if (SegmentDesc.Length > 200) throw new SystemException("【編碼節段說明】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷編碼節段代碼是否重複
                        sql = @"SELECT TOP 1 1
                                FROM PDM.ItemSegment
                                WHERE SegmentNo = @SegmentNo
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("SegmentNo", SegmentNo);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【編碼節段代碼】重複，請重新輸入!");
                        #endregion

                        #region //生效日失效日笨蛋測試
                        if (EffectiveDate != "" && InactiveDate != "")
                        {
                            DateTime date1 = DateTime.Parse(EffectiveDate);
                            DateTime date2 = DateTime.Parse(InactiveDate);
                            if (date2 < date1)
                            {
                                #region //撈取使用者名稱工號部門
                                sql = @"SELECT UserNo,UserName,DepartmentId
                                        FROM BAS.[User]
                                        WHERE UserId = @CreateBy";
                                dynamicParameters.Add("CreateBy", CreateBy);
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() > 0)
                                {
                                    foreach (var item in result)
                                    {
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO  QMS.FoolListLog (FoolerHumanityId, FoolerHumanityNo, FoolerHumanityName, FoolerHumanityDepartment, FoolDesc, 
                                                CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                                OUTPUT INSERTED.FoolListLogId
                                                VALUES (@FoolerHumanityId, @FoolerHumanityNo, @FoolerHumanityName, @FoolerHumanityDepartment, @FoolDesc, 
                                                @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                FoolerHumanityId = CreateBy,
                                                FoolerHumanityNo = item.UserNo,
                                                FoolerHumanityName = item.UserName,
                                                FoolerHumanityDepartment = item.DepartmentId,
                                                FoolDesc = CreateDate + ":【" + item.UserName + "】操作【編碼節段新增】時,將失效日設置小於生效日",
                                                CreateDate,
                                                LastModifiedDate,
                                                CreateBy,
                                                LastModifiedBy
                                            });
                                        var insertResult1 = sqlConnection.Query(sql, dynamicParameters);
                                    }
                                }
                                #endregion
                            }
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO PDM.ItemSegment (CompanyId, SegmentNo, SegmentName, SegmentDesc, EffectiveDate, InactiveDate, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ItemSegmentId
                                VALUES (@CompanyId, @SegmentNo, @SegmentName, @SegmentDesc, @EffectiveDate, @InactiveDate, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                SegmentNo,
                                SegmentName,
                                SegmentDesc,
                                EffectiveDate = EffectiveDate.Length > 0 ? EffectiveDate : "1900-01-01",
                                InactiveDate = InactiveDate.Length>0? InactiveDate : null,
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

        #region //AddSegmentValue -- 編碼節段Value新增 -- Shintokuro 2023.03.01
        public string AddSegmentValue(int ItemSegmentId ,string SegmentValueNo, string SegmentValue, string SegmentValueDesc, string EffectiveDate, string InactiveDate)
        {
            try
            {
                if (SegmentValueNo.Length <= 0) throw new SystemException("【代碼】不能為空!");
                if (SegmentValueNo.Length > 20) throw new SystemException("【代碼】長度錯誤!");
                if (SegmentValue.Length <= 0) throw new SystemException("【代碼值】不能為空!");
                if (SegmentValue.Length > 200) throw new SystemException("【代碼值】長度錯誤!");
                if (SegmentValueDesc.Length > 200) throw new SystemException("【代碼說明】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷編碼節段代碼是否重複
                        sql = @"SELECT TOP 1 1
                                FROM PDM.ItemSegmentValue
                                WHERE SegmentValueNo = @SegmentValueNo
                                AND ItemSegmentId = @ItemSegmentId";
                        dynamicParameters.Add("SegmentValueNo", SegmentValueNo);
                        dynamicParameters.Add("ItemSegmentId", ItemSegmentId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【代碼】重複，請重新輸入!");
                        #endregion

                        #region //生效日失效日笨蛋測試
                        if (EffectiveDate != "" && InactiveDate != "")
                        {
                            DateTime date1 = DateTime.Parse(EffectiveDate);
                            DateTime date2 = DateTime.Parse(InactiveDate);
                            if (date2 < date1)
                            {
                                #region //撈取使用者名稱工號部門
                                sql = @"SELECT UserNo,UserName,DepartmentId
                                        FROM BAS.[User]
                                        WHERE UserId = @CreateBy";
                                dynamicParameters.Add("CreateBy", CreateBy);
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() > 0)
                                {
                                    foreach (var item in result)
                                    {
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO  QMS.FoolListLog (FoolerHumanityId, FoolerHumanityNo, FoolerHumanityName, FoolerHumanityDepartment, FoolDesc, 
                                                CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                                OUTPUT INSERTED.FoolListLogId
                                                VALUES (@FoolerHumanityId, @FoolerHumanityNo, @FoolerHumanityName, @FoolerHumanityDepartment, @FoolDesc, 
                                                @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                FoolerHumanityId = CreateBy,
                                                FoolerHumanityNo = item.UserNo,
                                                FoolerHumanityName = item.UserName,
                                                FoolerHumanityDepartment = item.DepartmentId,
                                                FoolDesc = CreateDate + ":【" + item.UserName + "】操作【編碼節段Value新增】時,將失效日設置小於生效日",
                                                CreateDate,
                                                LastModifiedDate,
                                                CreateBy,
                                                LastModifiedBy
                                            });
                                        var insertResult1 = sqlConnection.Query(sql, dynamicParameters);
                                    }
                                }
                                #endregion
                            }
                        }
                        #endregion


                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO PDM.ItemSegmentValue (ItemSegmentId, SegmentValueNo, SegmentValue, SegmentValueDesc, EffectiveDate, InactiveDate, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.SegmentValueId
                                VALUES (@ItemSegmentId, @SegmentValueNo, @SegmentValue, @SegmentValueDesc, @EffectiveDate, @InactiveDate, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ItemSegmentId,
                                SegmentValueNo,
                                SegmentValue,
                                SegmentValueDesc,
                                EffectiveDate = EffectiveDate.Length > 0 ? EffectiveDate : "1900-01-01",
                                InactiveDate = InactiveDate.Length > 0 ? InactiveDate : null,
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

        #region //AddItemType -- 品號類別新增 -- Shintokuro 2023.03.06
        public string AddItemType(string ItemTypeNo, string ItemTypeName, string ItemTypeDesc, string EffectiveDate, string InactiveDate)
        {
            try
            {
                if (ItemTypeNo.Length <= 0) throw new SystemException("【品號類別代碼】不能為空!");
                if (ItemTypeNo.Length > 30) throw new SystemException("【品號類別代碼】長度錯誤!");
                if (ItemTypeName.Length <= 0) throw new SystemException("【品號類別名稱】不能為空!");
                if (ItemTypeName.Length > 30) throw new SystemException("【品號類別名稱】長度錯誤!");
                if (ItemTypeDesc.Length > 200) throw new SystemException("【品號類別敘述】長度錯誤!");
                
                

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷品號類別代碼是否重複
                        sql = @"SELECT TOP 1 1
                                FROM PDM.ItemType
                                WHERE ItemTypeNo = @ItemTypeNo
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("ItemTypeNo", ItemTypeNo);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【品號類別代碼】重複，請重新輸入!");
                        #endregion

                        #region //生效日失效日笨蛋測試
                        if (EffectiveDate != "" && InactiveDate != "")
                        {
                            DateTime date1 = DateTime.Parse(EffectiveDate);
                            DateTime date2 = DateTime.Parse(InactiveDate);
                            if (date2 < date1)
                            {
                                #region //撈取使用者名稱工號部門
                                sql = @"SELECT UserNo,UserName,DepartmentId
                                        FROM BAS.[User]
                                        WHERE UserId = @CreateBy";
                                dynamicParameters.Add("CreateBy", CreateBy);
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() > 0)
                                {
                                    foreach (var item in result)
                                    {
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO  QMS.FoolListLog (FoolerHumanityId, FoolerHumanityNo, FoolerHumanityName, FoolerHumanityDepartment, FoolDesc, 
                                                CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                                OUTPUT INSERTED.FoolListLogId
                                                VALUES (@FoolerHumanityId, @FoolerHumanityNo, @FoolerHumanityName, @FoolerHumanityDepartment, @FoolDesc, 
                                                @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                FoolerHumanityId = CreateBy,
                                                FoolerHumanityNo = item.UserNo,
                                                FoolerHumanityName = item.UserName,
                                                FoolerHumanityDepartment = item.DepartmentId,
                                                FoolDesc = CreateDate + ":【" + item.UserName + "】操作【品號類別新增】時,將失效日設置小於生效日",
                                                CreateDate,
                                                LastModifiedDate,
                                                CreateBy,
                                                LastModifiedBy
                                            });
                                        var insertResult1 = sqlConnection.Query(sql, dynamicParameters);
                                    }
                                }
                                #endregion
                            }
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO PDM.ItemType (CompanyId, ItemTypeNo, ItemTypeName, ItemTypeDesc, EffectiveDate, InactiveDate, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ItemTypeId
                                VALUES (@CompanyId, @ItemTypeNo, @ItemTypeName, @ItemTypeDesc, @EffectiveDate, @InactiveDate, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                ItemTypeNo,
                                ItemTypeName,
                                ItemTypeDesc,
                                EffectiveDate = EffectiveDate.Length > 0 ? EffectiveDate : "1900-01-01",
                                InactiveDate = InactiveDate.Length > 0 ? InactiveDate : null,
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

        #region //AddItemTypeSegment -- 品號類別結構新增 -- Shintokuro 2023.03.07
        public string AddItemTypeSegment(int ItemTypeId, string SegmentType, string SegmentValue, string SuffixCode, string EffectiveDate, string InactiveDate)
        {
            try
            {

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷品號類別是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.ItemType
                                WHERE ItemTypeId = @ItemTypeId";
                        dynamicParameters.Add("ItemTypeId", ItemTypeId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【品號類別】不存在，請重新輸入!");
                        #endregion

                        #region //撈取排序值
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT MAX(SortNumber) SortNumberBase
                                FROM PDM.ItemTypeSegment
                                WHERE ItemTypeId = @ItemTypeId";
                        dynamicParameters.Add("ItemTypeId", ItemTypeId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        int? SortNumberBase = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).SortNumberBase;
                        if(SortNumberBase == null || SortNumberBase == -1)
                        {
                            SortNumberBase = 1;
                        }
                        else
                        {
                            SortNumberBase++;
                        }
                        #endregion

                        #region //判斷編碼內容是否符合規範
                        if (SegmentType == "S")
                        {
                            var Sequence = Convert.ToInt32(SegmentValue);
                            if(Sequence <= 0) throw new SystemException("【流水號碼數】不可以小於0，請重新輸入!");
                        }

                        if (SegmentType == "SV")
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM PDM.ItemSegment
                                    WHERE ItemSegmentId = @ItemSegmentId";
                            dynamicParameters.Add("ItemSegmentId", SegmentValue);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【編碼節斷】不存在，請重新輸入!");
                        }
                        #endregion

                        #region //生效日失效日笨蛋測試
                        if (EffectiveDate != "" && InactiveDate != "")
                        {
                            DateTime date1 = DateTime.Parse(EffectiveDate);
                            DateTime date2 = DateTime.Parse(InactiveDate);
                            if (date2 < date1)
                            {
                                #region //撈取使用者名稱工號部門
                                sql = @"SELECT UserNo,UserName,DepartmentId
                                        FROM BAS.[User]
                                        WHERE UserId = @CreateBy";
                                dynamicParameters.Add("CreateBy", CreateBy);
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() > 0)
                                {
                                    foreach (var item in result)
                                    {
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO  QMS.FoolListLog (FoolerHumanityId, FoolerHumanityNo, FoolerHumanityName, FoolerHumanityDepartment, FoolDesc, 
                                                CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                                OUTPUT INSERTED.FoolListLogId
                                                VALUES (@FoolerHumanityId, @FoolerHumanityNo, @FoolerHumanityName, @FoolerHumanityDepartment, @FoolDesc, 
                                                @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                FoolerHumanityId = CreateBy,
                                                FoolerHumanityNo = item.UserNo,
                                                FoolerHumanityName = item.UserName,
                                                FoolerHumanityDepartment = item.DepartmentId,
                                                FoolDesc = CreateDate + ":【" + item.UserName + "】操作【品號類別結構新增】時,將失效日設置小於生效日",
                                                CreateDate,
                                                LastModifiedDate,
                                                CreateBy,
                                                LastModifiedBy
                                            });
                                        var insertResult1 = sqlConnection.Query(sql, dynamicParameters);
                                    }
                                }
                                #endregion
                            }
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO PDM.ItemTypeSegment (ItemTypeId, SortNumber, SegmentType, SegmentValue, SuffixCode, EffectiveDate, InactiveDate, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ItemTypeSegmentId
                                VALUES (@ItemTypeId, @SortNumber, @SegmentType, @SegmentValue, @SuffixCode, @EffectiveDate, @InactiveDate, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ItemTypeId,
                                SortNumber = SortNumberBase,
                                SegmentType,
                                SegmentValue,
                                SuffixCode,
                                EffectiveDate = EffectiveDate.Length > 0 ? EffectiveDate : "1900-01-01",
                                InactiveDate = InactiveDate.Length > 0 ? InactiveDate : null,
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

        #region //AddItemTypeDefault -- 品號類別預設資料新增 -- Shintokuro 2023.04.17
        public string AddItemTypeDefault(int ItemTypeId, string TypeOne, string TypeTwo, string TypeThree, string TypeFour
            , int InventoryId, int RequisitionInventoryId, int InventoryUomId, string ItemAttribute
            , string MeasureType, int PurchaseUomId, int SaleUomId, string LotManagement, string InventoryManagement
            , string MtlModify, string BondedStore, string OverReceiptManagement, string OverDeliveryManagement
            , string EffectiveDate, string ExpirationDate, string MtlItemDesc, string MtlItemRemark
            )
        {
            try
            {
                if (ItemTypeId <= 0) throw new SystemException("【品號類別】不能為空!");
                if (TypeOne.Length <= 0) throw new SystemException("【品號管理】- 【品號分類(一)】不能為空!");
                //if (TypeTwo.Length <= 0) throw new SystemException("【品號管理】- 【品號分類(二)】不能為空!");
                if (TypeThree.Length <= 0) throw new SystemException("【品號管理】- 【品號分類(三)】不能為空!");
                if (TypeFour.Length <= 0) throw new SystemException("【品號管理】- 【品號分類(四)】不能為空!");
                if (InventoryId <= 0) throw new SystemException("【主要庫別】不能為空!");
                if (RequisitionInventoryId <= 0) throw new SystemException("【主要領料庫別】不能為空!");
                //if (InventoryUomId  <= 0) throw new SystemException("【入庫單位】不能為空!");
                if (ItemAttribute.Length <= 0) throw new SystemException("【品號屬性】不能為空!");
                int? nullData = null;
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷品號類別代碼是否存在
                        sql = @"SELECT TOP 1 1
                                FROM PDM.ItemType
                                WHERE ItemTypeId = @ItemTypeId";
                        dynamicParameters.Add("ItemTypeId", ItemTypeId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【品號類別代碼】不存在，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO PDM.ItemTypeDefault (ItemTypeId, TypeOne, TypeTwo, TypeThree, TypeFour, InventoryId, RequisitionInventoryId, InventoryUomId, ItemAttribute,
                                MeasureType, PurchaseUomId, SaleUomId, LotManagement, InventoryManagement,
                                MtlModify, BondedStore, OverReceiptManagement, OverDeliveryManagement, 
                                EffectiveDate, ExpirationDate, MtlItemDesc, MtlItemRemark,
                                CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ItemTypeDefaultId
                                VALUES (@ItemTypeId, @TypeOne, @TypeTwo, @TypeThree, @TypeFour, @InventoryId, @RequisitionInventoryId, @InventoryUomId, @ItemAttribute,
                                @MeasureType, @PurchaseUomId, @SaleUomId, @LotManagement, @InventoryManagement,
                                @MtlModify, @BondedStore, @OverReceiptManagement, @OverDeliveryManagement,
                                @EffectiveDate, @ExpirationDate, @MtlItemDesc, @MtlItemRemark,
                                @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ItemTypeId,
                                TypeOne,
                                TypeTwo,
                                TypeThree,
                                TypeFour,
                                InventoryId,
                                RequisitionInventoryId,
                                InventoryUomId,
                                ItemAttribute,
                                MeasureType = MeasureType != "" ? MeasureType : null,
                                PurchaseUomId = PurchaseUomId != -1 ? PurchaseUomId : nullData,
                                SaleUomId = SaleUomId != -1 ? SaleUomId : nullData,
                                LotManagement = LotManagement != "" ? LotManagement : null,
                                InventoryManagement = InventoryManagement != "" ? InventoryManagement : null,
                                MtlModify = MtlModify != "" ? MtlModify : null,
                                BondedStore = BondedStore != "" ? BondedStore : null,
                                OverReceiptManagement = OverReceiptManagement != "" ? OverReceiptManagement : null,
                                OverDeliveryManagement = OverDeliveryManagement != "" ? OverDeliveryManagement : null,
                                EffectiveDate = EffectiveDate != "" ? EffectiveDate : null,
                                ExpirationDate = ExpirationDate != "" ? ExpirationDate : null,
                                MtlItemDesc = MtlItemDesc != "" ? MtlItemDesc : null,
                                MtlItemRemark = MtlItemRemark != "" ? MtlItemRemark : null,
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

        #region //AddDfmItemCategory -- 新增DFM項目種類代碼 -- Shintokuro 2023.08.28
        public string AddDfmItemCategory(int ModeId, string DfmItemCategoryNo, string DfmItemCategoryName, string DfmItemCategoryDesc)
        {
            try
            {
                if (ModeId <= 0) throw new SystemException("【項目種類】不能為空!");
                if (DfmItemCategoryNo.Length <= 0) throw new SystemException("【項目種類代碼】不能為空!");
                if (DfmItemCategoryNo.Length > 20) throw new SystemException("【項目種類代碼】長度錯誤!");
                if (DfmItemCategoryName.Length <= 0) throw new SystemException("【項目種類名稱】不能為空!");
                if (DfmItemCategoryName.Length > 100) throw new SystemException("【項目種類名稱】長度錯誤!");
                if (DfmItemCategoryDesc.Length <= 0) throw new SystemException("【項目種類描述】不能為空!");
                if (DfmItemCategoryDesc.Length > 100) throw new SystemException("【項目種類描述】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();
                        int rowsAffected = 0;
                        #region //判斷生產模式是否存在
                        sql = @"SELECT TOP 1 1
                                FROM MES.ProdMode
                                WHERE ModeId = @ModeId";
                        dynamicParameters.Add("ModeId", ModeId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【生產模式】不存在，請重新確認!");
                        #endregion

                        #region //判斷項目種類代碼是否重複
                        sql = @"SELECT TOP 1 1
                                FROM PDM.DfmItemCategory a
                                INNER JOIN MES.ProdMode b on a.ModeId = b.ModeId
                                WHERE a.DfmItemCategoryNo = @DfmItemCategoryNo
                                AND b.CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("DfmItemCategoryNo", DfmItemCategoryNo);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【項目種類代碼】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO PDM.DfmItemCategory (ModeId, DfmItemCategoryNo, DfmItemCategoryName, DfmItemCategoryDesc, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.DfmItemCategoryId
                                VALUES (@ModeId, @DfmItemCategoryNo, @DfmItemCategoryName, @DfmItemCategoryDesc, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ModeId,
                                DfmItemCategoryNo,
                                DfmItemCategoryName,
                                DfmItemCategoryDesc,
                                Status = "A", //啟用
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += insertResult.Count();

                        int DfmItemCategoryId = -1;
                        foreach (var item in insertResult)
                        {
                            DfmItemCategoryId = item.DfmItemCategoryId;
                        }
                        if (DfmItemCategoryId>0)
                        {
                            #region //判斷項目種類樣板是否已經有資料
                            sql = @"SELECT TOP 1 1
                                    FROM PDM.DfmCategoryTemplate a
                                    WHERE ａ.DfmItemCategoryId = @DfmItemCategoryId";
                            dynamicParameters.Add("DfmItemCategoryId", DfmItemCategoryId);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() > 0) throw new SystemException("【項目種類樣板】已經存在，請重新輸入!");
                            #endregion

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO PDM.DfmCategoryTemplate (DfmItemCategoryId, QiMaterialTemplate, QiProcessTemplate, Status
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.DfmItemCategoryId
                                    VALUES (@DfmItemCategoryId, @QiMaterialTemplate, @QiProcessTemplate, @Status
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    DfmItemCategoryId,
                                    QiMaterialTemplate = (string)null,
                                    QiProcessTemplate = (string)null,
                                    DfmItemCategoryDesc,
                                    Status = "A", //啟用
                                CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            var insertResult1 = sqlConnection.Query(sql, dynamicParameters);

                        }

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

        #region //AddDfmMaterial -- 新增DFM物料代碼 -- Shintokuro 2023.08.28
        public string AddDfmMaterial(string DfmMaterialNo, string DfmMaterialName, string DfmMaterialDesc, double DfmMaterialMoney)
        {
            try
            {
                if (DfmMaterialNo.Length <= 0) throw new SystemException("【物料代碼】不能為空!");
                if (DfmMaterialNo.Length > 20) throw new SystemException("【物料代碼】長度錯誤!");
                if (DfmMaterialName.Length <= 0) throw new SystemException("【物料名稱】不能為空!");
                if (DfmMaterialName.Length > 100) throw new SystemException("【物料名稱】長度錯誤!");
                if (DfmMaterialDesc.Length > 100) throw new SystemException("【物料描述】長度錯誤!");
                if (DfmMaterialMoney < 0) throw new SystemException("【物料金額】不能為負!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷物料代碼是否重複
                        sql = @"SELECT TOP 1 1
                                FROM PDM.DfmMaterial a
                                WHERE a.CompanyId = @CompanyId
                                AND a.DfmMaterialNo = @DfmMaterialNo";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("DfmMaterialNo", DfmMaterialNo);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【物料代碼】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO PDM.DfmMaterial (CompanyId, DfmMaterialNo, DfmMaterialName, 
                                DfmMaterialDesc, DfmMaterialMoney, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.DfmMaterialId
                                VALUES (@CompanyId, @DfmMaterialNo, @DfmMaterialName, 
                                @DfmMaterialDesc, @DfmMaterialMoney, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                DfmMaterialNo,
                                DfmMaterialName,
                                DfmMaterialDesc,
                                DfmMaterialMoney,
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

        #region //AddDfmCategoryMaterial -- 新增DFM項目種類物料 -- Xuan 2023.09.01
        public string AddDfmCategoryMaterial(int DfmMaterialId, string ProcessData, int DfmItemCategoryId, string DfmMaterialNo, string DfmMaterialName, string DfmMaterialDesc, double DfmMaterialMoney)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        var DfmItemJson = JObject.Parse(ProcessData);
                        int rowsAffected = 0;
                        foreach (var item in DfmItemJson["dfmmaterialInfo"])
                        {
                            #region //判斷物料是否存在種類之中
                            sql = @"SELECT TOP 1 1
                                FROM PDM.DfmCategoryMaterial
                                WHERE DfmMaterialId = @DfmMaterialId
                                AND DfmItemCategoryId = @DfmItemCategoryId";
                            dynamicParameters.Add("DfmMaterialId", Convert.ToInt32(item["DfmMaterialId"]));
                            dynamicParameters.Add("DfmItemCategoryId", DfmItemCategoryId);
                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO PDM.DfmCategoryMaterial (DfmMaterialId,DfmItemCategoryId,DfmMaterialDesc,Status,CreateDate,LastModifiedDate,CreateBy,LastModifiedBy)
                                        OUTPUT INSERTED.DfmCategoryMaterialId
                                        VALUES(@DfmMaterialId,@DfmItemCategoryId,@DfmMaterialDesc,@Status,@CreateDate,@LastModifiedDate,@CreateBy,@LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        DfmMaterialId = Convert.ToInt32(item["DfmMaterialId"]),
                                        DfmItemCategoryId,
                                        DfmMaterialDesc,
                                        Status = "A", //啟用
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

        #region //AddBatchMoldingConditions -- 批量新增成型生產條件 -- Shintokuro 2025-05-07
        public string AddBatchMoldingConditions(List<StockInDetail> mcExcelFormats)
        {
            try
            {
               
                int MtlItemId = -1;
                int rowsAffected = 0;

                string mesErr = "";

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        foreach (var mcExcel in mcExcelFormats)
                        {
                            if (mcExcel.MtlItemNo.Length <= 0) throw new SystemException("【品號】不能為空!");
                            if (mcExcel.Material.Length <= 0) throw new SystemException("【材料】不能為空!");
                            if (mcExcel.CycleTime.Length <= 0) throw new SystemException("【週期】不能為空!");
                            if (mcExcel.MoldWeight.Length <= 0) throw new SystemException("【模重】不能為空!");
                            if (mcExcel.Cavity.Length <= 0) throw new SystemException("【穴數】不能為空!");
                            //if (mcExcel.ProcessingFee.Length <= 0) throw new SystemException("【加共費】不能為空!");
                            //if (mcExcel.TheoreticalQty.Length <= 0) throw new SystemException("【理論生產數】不能為空!");
                            //if (mcExcel.UnitPrice == null) mcExcel.UnitPrice = "0";

                            #region //判斷品號資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.MtlItemId
                                        FROM PDM.MtlItem a 
                                        WHERE a.MtlItemNo = @MtlItemNo";
                            dynamicParameters.Add("MtlItemNo", mcExcel.MtlItemNo);

                            var MtlItemResult = sqlConnection.Query(sql, dynamicParameters);

                            if (MtlItemResult.Count() <= 0)
                            {
                                mesErr += "【" + mcExcel.MtlItemNo + "】";
                                continue;
                            }


                            MtlItemId = MtlItemResult.FirstOrDefault().MtlItemId;
                            #endregion

                            #region //判斷成行條件是否已存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM PDM.MoldingConditions a 
                                    WHERE a.MtlItemId = @MtlItemId";
                            dynamicParameters.Add("MtlItemId", MtlItemId);

                            MtlItemResult = sqlConnection.Query(sql, dynamicParameters);

                            if (MtlItemResult.Count() > 0)
                            {
                                continue;
                            }
                            #endregion

                            #region //INSERT MES.MoldingConditions
                            double CycleTime = Math.Round(Convert.ToDouble(mcExcel.CycleTime), 2);
                            double MoldWeight = Math.Round(Convert.ToDouble(mcExcel.MoldWeight), 2);
                            double temp = double.Parse(mcExcel.TheoreticalQty.ToString());
                            int TheoreticalQty = (int)Math.Floor(temp);  // 無條件捨去

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO PDM.MoldingConditions (
                                        MtlItemId, Material, CycleTime, MoldWeight, Cavity, ProcessingFee,
                                        UnitPrice, TheoreticalQty, CreateDate, LastModifiedDate, CreateBy, LastModifiedBy
                                    )
                                    OUTPUT INSERTED.McId
                                    VALUES (
                                        @MtlItemId, @Material, @CycleTime, @MoldWeight, @Cavity, @ProcessingFee,
                                        @UnitPrice, @TheoreticalQty, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy
                                    )";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                MtlItemId,
                                mcExcel.MtlItemNo,
                                mcExcel.Material,
                                CycleTime,
                                MoldWeight,
                                mcExcel.Cavity,
                                mcExcel.ProcessingFee,
                                TheoreticalQty,
                                UnitPrice = 0,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });

                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();
                            #endregion
                        }

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            err = mesErr
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

        #region //AddItemCategoryDept -- 新增品號類別責任部門 --Jean 2025-05-25 --
        public string AddItemCategoryDept(string TypeTwo, string TypeThree, string InventoryNo, string CatDept)
        {
            try
            {
                if (InventoryNo.Length <= 0) throw new SystemException("【組合三】不能為空!");
                if (CatDept.Length <= 0) throw new SystemException("【責任部門】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO PDM.ItemCatDept ( TypeTwo, TypeThree, InventoryNo, CatDept, CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.ItemCatDeptId
                                VALUES (@TypeTwo, @TypeThree, @InventoryNo, @CatDept, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TypeTwo,
                                TypeThree,
                                InventoryNo,
                                CatDept,
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

        #region //AddMoldingConditions -- 新增成型產品生產條件 --Jean 2025-05-25 --
        public string AddMoldingConditions(string MtlItemId, string Material , string CycleTime , string MoldWeight 
            , string Cavity , string ProcessingFee, string TheoreticalQty )
        {
            try
            {
                if (MtlItemId.Length <= 0) throw new SystemException("【ERP品號】不能為空!");
                if (Material.Length <= 0) throw new SystemException("【材料】不能為空!");
                if (CycleTime.Length <= 0) throw new SystemException("【週期】不能為空!");
                if (MoldWeight.Length <= 0) throw new SystemException("【模重】不能為空!");
                if (Cavity.Length <= 0) throw new SystemException("【穴數】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();


                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO PDM.MoldingConditions ( MtlItemId, Material, CycleTime, MoldWeight, Cavity, ProcessingFee, TheoreticalQty, CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.McId
                                VALUES (@MtlItemId, @Material, @CycleTime, @MoldWeight, @Cavity, ROUND(1800/1.13, 0) , ROUND(24*60*60/@CycleTime*@Cavity, 0), @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                MtlItemId,
                                Material,
                                CycleTime,
                                MoldWeight,
                                Cavity,
                                ProcessingFee,
                                TheoreticalQty,
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
        #region //UpdateMtlElement -- 品號元素資料更新 -- Kan 2022.05.31
        public string UpdateMtlElement(int ElementId, string ElementNo, string ElementName)
        {
            try
            {
                if (ElementNo.Length <= 0) throw new SystemException("【品號元素代碼】不能為空!");
                if (ElementNo.Length > 50) throw new SystemException("【品號元素代碼】長度錯誤!");
                if (ElementName.Length <= 0) throw new SystemException("【品號元素名稱】不能為空!");
                if (ElementName.Length > 100) throw new SystemException("【品號元素名稱】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷品號元素資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.MtlElement
                                WHERE CompanyId = @CompanyId
                                AND ElementId = @ElementId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ElementId", ElementId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("品號元素資料錯誤!");
                        #endregion

                        #region //判斷品號元素代碼是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.MtlElement
                                WHERE CompanyId = @CompanyId
                                AND ElementNo = @ElementNo
                                AND ElementId != @ElementId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ElementNo", ElementNo);
                        dynamicParameters.Add("ElementId", ElementId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【品號元素代碼】重複，請重新輸入!");
                        #endregion

                        #region //判斷品號元素名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.MtlElement
                                WHERE CompanyId = @CompanyId
                                AND ElementName = @ElementName
                                AND ElementId != @ElementId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ElementName", ElementName);
                        dynamicParameters.Add("ElementId", ElementId);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() > 0) throw new SystemException("【品號元素名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PDM.MtlElement SET
                                ElementNo = @ElementNo,
                                ElementName = @ElementName,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ElementId = @ElementId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ElementNo,
                                ElementName,
                                LastModifiedDate,
                                LastModifiedBy,
                                ElementId
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

        #region //UpdateMtlElementStatus -- 品號元素狀態更新 -- Kan 2022.05.31
        public string UpdateMtlElementStatus(int ElementId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷品號元素資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 [Status]
                                FROM PDM.MtlElement
                                WHERE CompanyId = @CompanyId
                                AND ElementId = @ElementId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ElementId", ElementId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("品號元素資料錯誤!");

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
                        sql = @"UPDATE PDM.MtlElement SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ElementId = @ElementId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                ElementId
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

        #region //UpdateMtlElementSort -- 品號元素順序調整 -- Kan 2022.05.31
        public string UpdateMtlElementSort(string MtlElementList)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //更新成負數避免觸發唯一值
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PDM.MtlElement SET
                                SortNumber = SortNumber * -1,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("LastModifiedDate", LastModifiedDate);
                        dynamicParameters.Add("LastModifiedBy", LastModifiedBy);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //更新順序
                        string[] MtlElementSort = MtlElementList.Split(',');

                        for (int i = 0; i < MtlElementSort.Length; i++)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE PDM.MtlElement SET
                                    SortNumber = @SortNumber,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE CompanyId = @CompanyId
                                    AND ElementId = @ElementId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    SortNumber = (i + 1),
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    CompanyId = CurrentCompany,
                                    ElementId = Convert.ToInt32(MtlElementSort[i])
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

        #region //UpdateUnitOfMeasure -- 單位基本資料更新 -- Kan 2022.06.15
        public string UpdateUnitOfMeasure(int UomId, string UomNo, string UomName, string UomType, string UomDesc)
        {
            try
            {
                if (UomNo.Length <= 0) throw new SystemException("【單位代碼】不能為空!");
                if (UomNo.Length > 50) throw new SystemException("【單位代碼】長度錯誤!");
                if (UomName.Length <= 0) throw new SystemException("【單位名稱】不能為空!");
                if (UomName.Length > 100) throw new SystemException("【單位名稱】長度錯誤!");
                if (UomType.Length <= 0) throw new SystemException("【單位類別】不能為空!");
                if (UomDesc.Length > 100) throw new SystemException("【單位描述】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷單位資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.UnitOfMeasure
                                WHERE CompanyId = @CompanyId
                                AND UomId = @UomId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("UomId", UomId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("單位資料錯誤!");
                        #endregion

                        #region //判斷單位代碼是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.UnitOfMeasure
                                WHERE CompanyId = @CompanyId
                                AND UomNo = @UomNo
                                AND UomId != @UomId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("UomNo", UomNo);
                        dynamicParameters.Add("UomId", UomId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【單位代碼】重複，請重新輸入!");
                        #endregion

                        #region //判斷單位名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.UnitOfMeasure
                                WHERE CompanyId = @CompanyId
                                AND UomName = @UomName
                                AND UomId != @UomId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("UomName", UomName);
                        dynamicParameters.Add("UomId", UomId);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() > 0) throw new SystemException("【單位名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PDM.UnitOfMeasure SET
                                CompanyId = @CompanyId,
                                UomNo = @UomNo,
                                UomName = @UomName,
                                UomType = @UomType,
                                UomDesc = @UomDesc,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE UomId = @UomId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                UomNo,
                                UomName,
                                UomType,
                                UomDesc,
                                LastModifiedDate,
                                LastModifiedBy,
                                UomId
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

        #region //UpdateUnitOfMeasureStatus -- 單位基本資料狀態更新 -- Kan 2022.06.15
        public string UpdateUnitOfMeasureStatus(int UomId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷單位資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 [Status]
                                FROM PDM.UnitOfMeasure
                                WHERE CompanyId = @CompanyId
                                AND UomId = @UomId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("UomId", UomId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("單位資料錯誤!");

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
                        sql = @"UPDATE PDM.UnitOfMeasure SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE UomId = @UomId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                UomId
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

        #region //UpdateUomCalculate -- 單位換算資料更新 -- Kan 2022.06.15
        public string UpdateUomCalculate(int UomCalculateId, int UomId, int ConvertUomId, double Value)
        {
            try
            {
                if (UomId <= 0) throw new SystemException("【所屬單位】不能為空!");
                if (ConvertUomId <= 0) throw new SystemException("【換算單位】不能為空!");
                if (Value <= 0) throw new SystemException("【值】不能為0!");
                if (UomId == ConvertUomId) throw new SystemException("【所屬單位】不能與【所屬單位】相同!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷所屬單位資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.UnitOfMeasure
                                WHERE CompanyId = @CompanyId
                                AND UomId = @UomId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("UomId", UomId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("所屬單位資料錯誤!");
                        #endregion

                        #region //判斷換算單位資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.UnitOfMeasure
                                WHERE CompanyId = @CompanyId
                                AND UomId = @ConvertUomId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ConvertUomId", ConvertUomId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("換算單位資料錯誤!");
                        #endregion

                        #region //判斷換算單位資料是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.UomCalculate
                                WHERE UomId = @UomId
                                AND ConvertUomId = @ConvertUomId
                                AND UomCalculateId != @UomCalculateId";
                        dynamicParameters.Add("UomId", UomId);
                        dynamicParameters.Add("ConvertUomId", ConvertUomId);
                        dynamicParameters.Add("UomCalculateId", UomCalculateId);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() > 0) throw new SystemException("單位換算資料重複!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PDM.UomCalculate SET
                                UomId = @UomId,
                                ConvertUomId = @ConvertUomId,
                                Value = @Value,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE UomCalculateId = @UomCalculateId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                UomId,
                                ConvertUomId,
                                Value,
                                LastModifiedDate,
                                LastModifiedBy,
                                UomCalculateId
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

        #region //UpdateUomCalculateStatus -- 單位換算資料狀態更新 -- Kan 2022.06.15
        public string UpdateUomCalculateStatus(int UomCalculateId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷單位換算資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 [Status]
                                FROM PDM.UomCalculate
                                WHERE UomCalculateId = @UomCalculateId";
                        dynamicParameters.Add("UomCalculateId", UomCalculateId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("單位換算資料錯誤!");

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
                        sql = @"UPDATE PDM.UomCalculate SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE UomCalculateId = @UomCalculateId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                UomCalculateId
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

        #region //UpdateUnitOfMeasureSynchronize -- 單位資料同步 -- Kan 2022.06.16
        public string UpdateUnitOfMeasureSynchronize(string CompanyNo, string UpdateDate)
        {
            try
            {
                List<UnitOfMeasure> unitOfMeasures = new List<UnitOfMeasure>();

                if (CompanyNo.Length <= 0) throw new SystemException("【公司】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int CompanyId = -1;
                    string ErpConnectionStrings = "";

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷公司資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 CompanyId, CompanyName, ErpDb
                                FROM BAS.Company
                                WHERE CompanyNo = @CompanyNo";
                        dynamicParameters.Add("CompanyNo", CompanyNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【公司】資料錯誤!");

                        foreach (var item in result)
                        {
                            CompanyId = Convert.ToInt32(item.CompanyId);
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        }
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        #region //撈取ERP單位資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT DISTINCT UPPER(LTRIM(RTRIM(MX001))) UomNo, LTRIM(RTRIM(MX002)) UomName, LTRIM(RTRIM(MX004)) UomDesc
                                , CASE LTRIM(RTRIM(MX003))
                                    WHEN 'Y' THEN 'A'
                                    WHEN 'N' THEN 'S'
                                END Status
                                , CASE WHEN LEN(LTRIM(RTRIM(CREATE_DATE))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_DATE)) as date), 'yyyy-MM-dd') ELSE NULL END TransferDate
                                , CASE WHEN LEN(LTRIM(RTRIM(CREATE_TIME))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_TIME)) as datetime), 'HH:mm:ss') ELSE NULL END TransferTime
                                FROM INVMX
                                WHERE 1=1";
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "UpdateDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME > @UpdateDate OR MODI_DATE + ' ' + MODI_TIME > @UpdateDate)", UpdateDate);
                        sql += @" ORDER BY TransferDate, TransferTime";

                        unitOfMeasures = sqlConnection.Query<UnitOfMeasure>(sql, dynamicParameters).ToList();
                        #endregion
                    }

                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷PDM.UnitOfMeasure是否有資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT UomId, UomNo
                                FROM PDM.UnitOfMeasure
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<UnitOfMeasure> resultUnitOfMeasure = sqlConnection.Query<UnitOfMeasure>(sql, dynamicParameters).ToList();

                        unitOfMeasures = unitOfMeasures.GroupJoin(resultUnitOfMeasure, x => x.UomNo.ToUpper(), y => y.UomNo.ToUpper(), (x, y) => { x.UomId = y.FirstOrDefault()?.UomId; return x; }).ToList();
                        #endregion

                        List<UnitOfMeasure> addUnitOfMeasures = unitOfMeasures.Where(x => x.UomId == null).ToList();
                        List<UnitOfMeasure> updateUnitOfMeasures = unitOfMeasures.Where(x => x.UomId != null).ToList();

                        #region //新增
                        if (addUnitOfMeasures.Count > 0)
                        {
                            addUnitOfMeasures
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.CompanyId = CompanyId;
                                    x.UomType = "I";
                                    x.TransferStatus = "Y";
                                    x.CreateDate = CreateDate;
                                    x.LastModifiedDate = LastModifiedDate;
                                    x.CreateBy = CreateBy;
                                    x.LastModifiedBy = LastModifiedBy;
                                });

                            sql = @"INSERT INTO PDM.UnitOfMeasure (CompanyId, UomNo, UomName, UomType, UomDesc
                                    , TransferStatus, TransferDate, Status
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@CompanyId, @UomNo, @UomName, @UomType, @UomDesc, @TransferStatus, @TransferDate, @Status
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            rowsAffected += sqlConnection.Execute(sql, addUnitOfMeasures);
                        }
                        #endregion

                        #region //修改
                        if (updateUnitOfMeasures.Count > 0)
                        {
                            updateUnitOfMeasures
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.TransferStatus = "Y";
                                    x.LastModifiedDate = LastModifiedDate;
                                    x.LastModifiedBy = LastModifiedBy;
                                });

                            sql = @"UPDATE PDM.UnitOfMeasure SET
                                    UomNo = @UomNo,
                                    UomName = @UomName,
                                    UomDesc = @UomDesc,
                                    TransferStatus = @TransferStatus,
                                    Status = @Status,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE UomId = @UomId";
                            rowsAffected += sqlConnection.Execute(sql, updateUnitOfMeasures);
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

        #region //UpdateMtlModel -- 品號機型資料更新 -- Kan 2022.06.17
        public string UpdateMtlModel(int MtlModelId, string MtlModelNo, string MtlModelName, string MtlModelDesc)
        {
            try
            {
                if (MtlModelNo.Length <= 0) throw new SystemException("【品號機型代碼】不能為空!");
                if (MtlModelNo.Length > 50) throw new SystemException("【品號機型代碼】長度錯誤!");
                if (MtlModelName.Length <= 0) throw new SystemException("【品號機型名稱】不能為空!");
                if (MtlModelName.Length > 100) throw new SystemException("【品號機型名稱】長度錯誤!");
                if (MtlModelDesc.Length > 100) throw new SystemException("【品號機型描述】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷品號機型資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.MtlModel
                                WHERE CompanyId = @CompanyId
                                AND MtlModelId = @MtlModelId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("MtlModelId", MtlModelId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("品號機型資料錯誤!");
                        #endregion

                        #region //判斷品號機型代碼是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.MtlModel
                                WHERE CompanyId = @CompanyId
                                AND MtlModelNo = @MtlModelNo
                                AND MtlModelId != @MtlModelId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("MtlModelNo", MtlModelNo);
                        dynamicParameters.Add("MtlModelId", MtlModelId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【品號機型代碼】重複，請重新輸入!");
                        #endregion

                        #region //判斷品號機型名稱是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.MtlModel
                                WHERE CompanyId = @CompanyId
                                AND MtlModelName = @MtlModelName
                                AND MtlModelId != @MtlModelId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("MtlModelName", MtlModelName);
                        dynamicParameters.Add("MtlModelId", MtlModelId);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() > 0) throw new SystemException("【品號機型名稱】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PDM.MtlModel SET
                                MtlModelNo = @MtlModelNo,
                                MtlModelName = @MtlModelName,
                                MtlModelDesc = @MtlModelDesc,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE MtlModelId = @MtlModelId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                MtlModelNo,
                                MtlModelName,
                                MtlModelDesc,
                                LastModifiedDate,
                                LastModifiedBy,
                                MtlModelId
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

        #region //UpdateMtlModelStatus -- 品號機型狀態更新 -- Kan 2022.06.17
        public string UpdateMtlModelStatus(int MtlModelId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷品號機型資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 [Status]
                                FROM PDM.MtlModel
                                WHERE CompanyId = @CompanyId
                                AND MtlModelId = @MtlModelId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("MtlModelId", MtlModelId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("品號機型資料錯誤!");

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
                        sql = @"UPDATE PDM.MtlModel SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE MtlModelId = @MtlModelId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                MtlModelId
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

        #region //UpdateMtlModelSort -- 品號機型順序調整 -- Kan 2022.06.22
        public string UpdateMtlModelSort(int ParentId, string MtlModelList)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //更新成負數避免觸發唯一值
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PDM.MtlModel SET
                                MtlModelSort = MtlModelSort * -1,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE CompanyId = @CompanyId
                                AND ParentId = @ParentId";
                        dynamicParameters.Add("LastModifiedDate", LastModifiedDate);
                        dynamicParameters.Add("LastModifiedBy", LastModifiedBy);
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ParentId", ParentId);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //更新順序
                        string[] mtlModelSort = MtlModelList.Split(',');

                        for (int i = 0; i < mtlModelSort.Length; i++)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE PDM.MtlModel SET
                                    MtlModelSort = @MtlModelSort,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE CompanyId = @CompanyId
                                    AND MtlModelId = @MtlModelId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MtlModelSort = (i + 1),
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    CompanyId = CurrentCompany,
                                    MtlModelId = Convert.ToInt32(mtlModelSort[i])
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

        #region //UpdateItemSegment -- 編碼節段更新 -- Shintokuro 2023.03.01
        public string UpdateItemSegment(int ItemSegmentId, string SegmentName, string SegmentDesc, string EffectiveDate, string InactiveDate
            )
        {
            try
            {
                if (SegmentName.Length <= 0) throw new SystemException("【編碼節段名稱】不能為空!");
                if (SegmentName.Length > 60) throw new SystemException("【編碼節段名稱】長度錯誤!");
                if (SegmentDesc.Length > 200) throw new SystemException("【編碼節段說明】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷編碼節段資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM PDM.ItemSegment
                                WHERE ItemSegmentId = @ItemSegmentId";
                        dynamicParameters.Add("ItemSegmentId", ItemSegmentId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("裝置資料錯誤!");
                        #endregion

                        #region //生效日失效日測試
                        if (EffectiveDate != "" && InactiveDate != "")
                        {
                            DateTime date1 = DateTime.Parse(EffectiveDate);
                            DateTime date2 = DateTime.Parse(InactiveDate);
                            if (date2 < date1)
                            {
                                #region //撈取使用者名稱工號部門
                                sql = @"SELECT UserNo,UserName,DepartmentId
                                        FROM BAS.[User]
                                        WHERE UserId = @CreateBy";
                                dynamicParameters.Add("CreateBy", CreateBy);
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() > 0)
                                {
                                    foreach (var item in result)
                                    {
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO  QMS.FoolListLog (FoolerHumanityId, FoolerHumanityNo, FoolerHumanityName, FoolerHumanityDepartment, FoolDesc, 
                                                CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                                OUTPUT INSERTED.FoolListLogId
                                                VALUES (@FoolerHumanityId, @FoolerHumanityNo, @FoolerHumanityName, @FoolerHumanityDepartment, @FoolDesc, 
                                                @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                FoolerHumanityId = CreateBy,
                                                FoolerHumanityNo = item.UserNo,
                                                FoolerHumanityName = item.UserName,
                                                FoolerHumanityDepartment = item.DepartmentId,
                                                FoolDesc = CreateDate + ":【" + item.UserName + "】操作【編碼節段更新】時,將失效日設置小於生效日",
                                                CreateDate,
                                                LastModifiedDate,
                                                CreateBy,
                                                LastModifiedBy
                                            });
                                        var insertResult1 = sqlConnection.Query(sql, dynamicParameters);
                                    }
                                }
                                #endregion
                            }
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PDM.ItemSegment SET
                                SegmentName = @SegmentName,
                                SegmentDesc = @SegmentDesc,
                                EffectiveDate = @EffectiveDate,
                                InactiveDate = @InactiveDate,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ItemSegmentId = @ItemSegmentId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SegmentName,
                                SegmentDesc,
                                EffectiveDate = EffectiveDate.Length > 0 ? EffectiveDate : "1900-01-01",
                                InactiveDate = InactiveDate.Length > 0 ? InactiveDate : null,
                                LastModifiedDate,
                                LastModifiedBy,
                                ItemSegmentId
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

        #region //UpdateItemSegmentIdStatus -- 編碼節段狀態更新 -- Shintokru 2023.03.01
        public string UpdateItemSegmentIdStatus(int ItemSegmentId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷裝置資訊是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM PDM.ItemSegment
                                WHERE CompanyId = @CompanyId
                                AND ItemSegmentId = @ItemSegmentId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ItemSegmentId", ItemSegmentId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("編碼節段資料錯誤!");

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
                        sql = @"UPDATE PDM.ItemSegment SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ItemSegmentId = @ItemSegmentId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                ItemSegmentId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                        #region //更新節段的Value狀態
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PDM.ItemSegmentValue SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ItemSegmentId = @ItemSegmentId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                ItemSegmentId
                            });
                        insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #region //UpdateSegmentValue -- 編碼節段Value更新 -- Shintokuro 2023.03.01
        public string UpdateSegmentValue(int ItemSegmentId, int SegmentValueId, string SegmentValue, string SegmentValueDesc, string EffectiveDate, string InactiveDate
            )
        {
            try
            {
                if (SegmentValue.Length <= 0) throw new SystemException("【代碼值】不能為空!");
                if (SegmentValue.Length > 200) throw new SystemException("【代碼值】長度錯誤!");
                if (SegmentValueDesc.Length > 200) throw new SystemException("【代碼說明】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷編碼節段Value資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM PDM.ItemSegmentValue
                                WHERE ItemSegmentId = @ItemSegmentId
                                AND SegmentValueId = @SegmentValueId";
                        dynamicParameters.Add("ItemSegmentId", ItemSegmentId);
                        dynamicParameters.Add("SegmentValueId", SegmentValueId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("編碼節段Value資料錯誤!");
                        #endregion

                        #region //生效日失效日笨蛋測試
                        if (EffectiveDate != "" && InactiveDate != "")
                        {
                            DateTime date1 = DateTime.Parse(EffectiveDate);
                            DateTime date2 = DateTime.Parse(InactiveDate);
                            if (date2 < date1)
                            {
                                #region //撈取使用者名稱工號部門
                                sql = @"SELECT UserNo,UserName,DepartmentId
                                        FROM BAS.[User]
                                        WHERE UserId = @CreateBy";
                                dynamicParameters.Add("CreateBy", CreateBy);
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() > 0)
                                {
                                    foreach (var item in result)
                                    {
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO  QMS.FoolListLog (FoolerHumanityId, FoolerHumanityNo, FoolerHumanityName, FoolerHumanityDepartment, FoolDesc, 
                                                CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                                OUTPUT INSERTED.FoolListLogId
                                                VALUES (@FoolerHumanityId, @FoolerHumanityNo, @FoolerHumanityName, @FoolerHumanityDepartment, @FoolDesc, 
                                                @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                FoolerHumanityId = CreateBy,
                                                FoolerHumanityNo = item.UserNo,
                                                FoolerHumanityName = item.UserName,
                                                FoolerHumanityDepartment = item.DepartmentId,
                                                FoolDesc = CreateDate + ":【" + item.UserName + "】操作【編碼節段Value更新】時,將失效日設置小於生效日",
                                                CreateDate,
                                                LastModifiedDate,
                                                CreateBy,
                                                LastModifiedBy
                                            });
                                        var insertResult1 = sqlConnection.Query(sql, dynamicParameters);
                                    }
                                }
                                #endregion
                            }
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PDM.ItemSegmentValue SET
                                SegmentValue = @SegmentValue,
                                SegmentValueDesc = @SegmentValueDesc,
                                EffectiveDate = @EffectiveDate,
                                InactiveDate = @InactiveDate,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ItemSegmentId = @ItemSegmentId
                                AND SegmentValueId = @SegmentValueId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SegmentValue,
                                SegmentValueDesc,
                                EffectiveDate = EffectiveDate.Length > 0 ? EffectiveDate : "1900-01-01",
                                InactiveDate = InactiveDate.Length > 0 ? InactiveDate : null,
                                LastModifiedDate,
                                LastModifiedBy,
                                ItemSegmentId,
                                SegmentValueId
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

        #region //UpdateItemSegmentValueStatus -- 編碼節段Value狀態更新 -- Shintokru 2023.03.01
        public string UpdateItemSegmentValueStatus(int ItemSegmentId, int SegmentValueId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷編碼節段Value資訊是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM PDM.ItemSegmentValue
                                WHERE ItemSegmentId = @ItemSegmentId
                                AND SegmentValueId = @SegmentValueId";
                        dynamicParameters.Add("ItemSegmentId", ItemSegmentId);
                        dynamicParameters.Add("SegmentValueId", SegmentValueId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("編碼節段Value資料錯誤!");

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
                        sql = @"UPDATE PDM.ItemSegmentValue SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ItemSegmentId = @ItemSegmentId
                                AND SegmentValueId = @SegmentValueId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                ItemSegmentId,
                                SegmentValueId
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

        #region //UpdateItemType -- 品號類別更新 -- Shintokuro 2023.03.06
        public string UpdateItemType(int ItemTypeId,string ItemTypeNo, string ItemTypeName, string ItemTypeDesc, string EffectiveDate, string InactiveDate
            )
        {
            try
            {

                if (ItemTypeNo.Length <= 0) throw new SystemException("【品號類別代碼】不能為空!");
                if (ItemTypeName.Length <= 0) throw new SystemException("【品號類別名稱】不能為空!");
                if (ItemTypeName.Length > 60) throw new SystemException("【品號類別名稱】長度錯誤!");
                if (ItemTypeDesc.Length > 200) throw new SystemException("【品號類別敘述】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷品號類別資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM PDM.ItemType
                                WHERE ItemTypeId = @ItemTypeId";
                        dynamicParameters.Add("ItemTypeId", ItemTypeId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("品號類別資料錯誤!");
                        #endregion

                        #region //生效日失效日測試
                        if (EffectiveDate != "" && InactiveDate != "")
                        {
                            DateTime date1 = DateTime.Parse(EffectiveDate);
                            DateTime date2 = DateTime.Parse(InactiveDate);
                            if (date2 < date1)
                            {
                                #region //撈取使用者名稱工號部門
                                sql = @"SELECT UserNo,UserName,DepartmentId
                                        FROM BAS.[User]
                                        WHERE UserId = @CreateBy";
                                dynamicParameters.Add("CreateBy", CreateBy);
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() > 0)
                                {
                                    foreach (var item in result)
                                    {
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO  QMS.FoolListLog (FoolerHumanityId, FoolerHumanityNo, FoolerHumanityName, FoolerHumanityDepartment, FoolDesc, 
                                                CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                                OUTPUT INSERTED.FoolListLogId
                                                VALUES (@FoolerHumanityId, @FoolerHumanityNo, @FoolerHumanityName, @FoolerHumanityDepartment, @FoolDesc, 
                                                @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                FoolerHumanityId = CreateBy,
                                                FoolerHumanityNo = item.UserNo,
                                                FoolerHumanityName = item.UserName,
                                                FoolerHumanityDepartment = item.DepartmentId,
                                                FoolDesc = CreateDate + ":【" + item.UserName + "】操作【品號類別更新】時,將失效日設置小於生效日",
                                                CreateDate,
                                                LastModifiedDate,
                                                CreateBy,
                                                LastModifiedBy
                                            });
                                        var insertResult1 = sqlConnection.Query(sql, dynamicParameters);
                                    }
                                }
                                #endregion
                            }
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PDM.ItemType SET
                                ItemTypeNo = @ItemTypeNo,
                                ItemTypeName = @ItemTypeName,
                                ItemTypeDesc = @ItemTypeDesc,
                                EffectiveDate = @EffectiveDate,
                                InactiveDate = @InactiveDate,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ItemTypeId = @ItemTypeId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ItemTypeNo,
                                ItemTypeName,
                                ItemTypeDesc,
                                EffectiveDate = EffectiveDate.Length > 0 ? EffectiveDate : "1900-01-01",
                                InactiveDate = InactiveDate.Length > 0 ? InactiveDate : null,
                                LastModifiedDate,
                                LastModifiedBy,
                                ItemTypeId
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

        #region //UpdateItemTypeIdStatus -- 品號類別狀態更新 -- Shintokru 2023.03.06
        public string UpdateItemTypeIdStatus(int ItemTypeId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷品號類別是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM PDM.ItemType
                                WHERE CompanyId = @CompanyId
                                AND ItemTypeId = @ItemTypeId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ItemTypeId", ItemTypeId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("品號類別資料錯誤!");

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
                        sql = @"UPDATE PDM.ItemType SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ItemTypeId = @ItemTypeId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                ItemTypeId
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

        #region //UpdateItemTypeSegment --品號類別結構更新 -- Shintokuro 2023.03.07
        public string UpdateItemTypeSegment(int ItemTypeId, int ItemTypeSegmentId, string SegmentType, string SegmentValue, string SuffixCode, string EffectiveDate, string InactiveDate
            )
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷品號類別是否正確
                        sql = @"SELECT TOP 1 1
                                FROM PDM.ItemType
                                WHERE ItemTypeId = @ItemTypeId";
                        dynamicParameters.Add("ItemTypeId", ItemTypeId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("品號類別資料錯誤!");
                        #endregion

                        #region //判斷品號類別結構是否正確
                        sql = @"SELECT TOP 1 1
                                FROM PDM.ItemTypeSegment
                                WHERE ItemTypeSegmentId = @ItemTypeSegmentId";
                        dynamicParameters.Add("ItemTypeSegmentId", ItemTypeSegmentId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("品號類別結構資料錯誤!");
                        #endregion

                        #region //判斷編碼內容是否符合規範
                        if (SegmentType == "S")
                        {
                            var Sequence = Convert.ToInt32(SegmentValue);
                            if (Sequence <= 0) throw new SystemException("【流水號碼數】不可以小於0，請重新輸入!");
                        }

                        if (SegmentType == "SV")
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM PDM.ItemSegment
                                    WHERE ItemSegmentId = @ItemSegmentId";
                            dynamicParameters.Add("ItemSegmentId", SegmentValue);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【編碼節斷】不存在，請重新輸入!");
                        }
                        #endregion

                        #region //生效日失效日笨蛋測試
                        if (EffectiveDate != "" && InactiveDate != "")
                        {
                            DateTime date1 = DateTime.Parse(EffectiveDate);
                            DateTime date2 = DateTime.Parse(InactiveDate);
                            if (date2 < date1)
                            {
                                #region //撈取使用者名稱工號部門
                                sql = @"SELECT UserNo,UserName,DepartmentId
                                        FROM BAS.[User]
                                        WHERE UserId = @CreateBy";
                                dynamicParameters.Add("CreateBy", CreateBy);
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() > 0)
                                {
                                    foreach (var item in result)
                                    {
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO  QMS.FoolListLog (FoolerHumanityId, FoolerHumanityNo, FoolerHumanityName, FoolerHumanityDepartment, FoolDesc, 
                                                CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                                OUTPUT INSERTED.FoolListLogId
                                                VALUES (@FoolerHumanityId, @FoolerHumanityNo, @FoolerHumanityName, @FoolerHumanityDepartment, @FoolDesc, 
                                                @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                FoolerHumanityId = CreateBy,
                                                FoolerHumanityNo = item.UserNo,
                                                FoolerHumanityName = item.UserName,
                                                FoolerHumanityDepartment = item.DepartmentId,
                                                FoolDesc = CreateDate + ":【" + item.UserName + "】操作【品號類別結構更新】時,將失效日設置小於生效日",
                                                CreateDate,
                                                LastModifiedDate,
                                                CreateBy,
                                                LastModifiedBy
                                            });
                                        var insertResult1 = sqlConnection.Query(sql, dynamicParameters);
                                    }
                                }
                                #endregion
                            }
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PDM.ItemTypeSegment SET
                                SegmentValue = @SegmentValue,
                                SuffixCode = @SuffixCode,
                                EffectiveDate = @EffectiveDate,
                                InactiveDate = @InactiveDate,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ItemTypeId = @ItemTypeId
                                AND ItemTypeSegmentId = @ItemTypeSegmentId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SegmentValue,
                                SuffixCode,
                                EffectiveDate = EffectiveDate.Length > 0 ? EffectiveDate : "1900-01-01",
                                InactiveDate = InactiveDate.Length > 0 ? InactiveDate : null,
                                LastModifiedDate,
                                LastModifiedBy,
                                ItemTypeId,
                                ItemTypeSegmentId
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

        #region //UpdateItemTypeSegmentSort --品號類別結構順序變更 -- Shintokuro 2023.04.04
        public string UpdateItemTypeSegmentSort(string ItemTypeSegmentIdList
            )
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();
                        int rowsAffected = 0;

                        #region //結構順序變更
                        foreach (var item in ItemTypeSegmentIdList.Split(','))
                        {
                            if(Convert.ToInt32(item.Split('❤')[0]) <= 0) throw new SystemException("編碼結構ID不可以小於等於0!");
                            if(Convert.ToInt32(item.Split('❤')[1]) <= 0) throw new SystemException("編碼結構順序不可以小於等於0!");

                            #region //判斷編碼結構ID是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM PDM.ItemTypeSegment
                                    WHERE ItemTypeSegmentId = @ItemTypeSegmentId";
                            dynamicParameters.Add("ItemTypeSegmentId", Convert.ToInt32(item.Split('❤')[0]));
                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("編碼結構ID不存在!");
                            #endregion

                            #region //判斷順序是否有不合規

                            #endregion

                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE PDM.ItemTypeSegment SET
                                    SortNumber = @SortNumber,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE ItemTypeSegmentId = @ItemTypeSegmentId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    SortNumber = Convert.ToInt32(item.Split('❤')[1]),
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    ItemTypeSegmentId = Convert.ToInt32(item.Split('❤')[0])
                                });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #region //UpdateItemTypeSegmentIdStatus -- 品號類別結構狀態更新 -- Shintokru 2023.03.06
        public string UpdateItemTypeSegmentIdStatus(int ItemTypeId, int ItemTypeSegmentId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        int rowsAffected = 0;

                        #region //判斷品號類別是否正確 + 順序更新
                        sql = @"SELECT TOP 1 1
                                FROM PDM.ItemType
                                WHERE ItemTypeId = @ItemTypeId";
                        dynamicParameters.Add("ItemTypeId", ItemTypeId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("品號類別不存在，請重新確認!");
                        #endregion

                        #region //判斷品號類別是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM PDM.ItemTypeSegment
                                WHERE ItemTypeId = @ItemTypeId
                                AND ItemTypeSegmentId = @ItemTypeSegmentId";
                        dynamicParameters.Add("ItemTypeId", ItemTypeId);
                        dynamicParameters.Add("ItemTypeSegmentId", ItemTypeSegmentId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("品號類別結構資料錯誤!");

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
                        sql = @"UPDATE PDM.ItemTypeSegment SET
                                Status = @Status,
                                SortNumber = -1,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ItemTypeId = @ItemTypeId
                                AND ItemTypeSegmentId = @ItemTypeSegmentId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                ItemTypeId,
                                ItemTypeSegmentId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        #region //判斷品號類別是否正確 + 順序更新
                        sql = @"SELECT ItemTypeSegmentId,SortNumber
                                FROM PDM.ItemTypeSegment
                                WHERE ItemTypeId = @ItemTypeId
                                AND Status = 'A'
                                ORDER BY SortNumber";
                        dynamicParameters.Add("ItemTypeId", ItemTypeId);

                        result = sqlConnection.Query(sql, dynamicParameters);

                        if (result.Count() > 0)
                        {
                            int sortNum = 1;
                            foreach (var item in result)
                            {
                                if (status == "A")
                                {
                                    sortNum = result.Count();
                                    if (item.ItemTypeSegmentId == ItemTypeSegmentId)
                                    {
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"UPDATE PDM.ItemTypeSegment SET
                                                SortNumber = @SortNumber,
                                                LastModifiedDate = @LastModifiedDate,
                                                LastModifiedBy = @LastModifiedBy
                                                WHERE ItemTypeSegmentId = @ItemTypeSegmentId";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                SortNumber = sortNum,
                                                LastModifiedDate,
                                                LastModifiedBy,
                                                ItemTypeSegmentId = item.ItemTypeSegmentId
                                            });
                                        var insertResult1 = sqlConnection.Query(sql, dynamicParameters);

                                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    }
                                }
                                else if (status == "S")
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE PDM.ItemTypeSegment SET
                                                SortNumber = @SortNumber,
                                                LastModifiedDate = @LastModifiedDate,
                                                LastModifiedBy = @LastModifiedBy
                                                WHERE ItemTypeSegmentId = @ItemTypeSegmentId";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            SortNumber = sortNum,
                                            LastModifiedDate,
                                            LastModifiedBy,
                                            ItemTypeSegmentId = item.ItemTypeSegmentId
                                        });
                                    var insertResult1 = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    sortNum++;
                                }

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

        #region //UpdateItemTypeDefault --品號類別預設資料更新 -- Shintokuro 2023.04.17
        public string UpdateItemTypeDefault(int ItemTypeDefaultId, int ItemTypeId, string TypeOne, string TypeTwo, string TypeThree, string TypeFour
            , int InventoryId, int RequisitionInventoryId, int InventoryUomId, string ItemAttribute
            , string MeasureType, int PurchaseUomId, int SaleUomId, string LotManagement, string InventoryManagement
            , string MtlModify, string BondedStore, string OverReceiptManagement, string OverDeliveryManagement
            , string EffectiveDate, string ExpirationDate, string MtlItemDesc, string MtlItemRemark
            )
        {
            try
            {
                if (ItemTypeDefaultId <= 0) throw new SystemException("【品號類別預設資料ID】不能為空!");
                if (ItemTypeId <= 0) throw new SystemException("【品號類別】不能為空!");
                if (TypeOne.Length <= 0) throw new SystemException("【品號管理】- 【品號分類(一)】不能為空!");
                //if (TypeTwo.Length <= 0) throw new SystemException("【品號管理】- 【品號分類(二)】不能為空!");
                if (TypeThree.Length <= 0) throw new SystemException("【品號管理】- 【品號分類(三)】不能為空!");
                if (TypeFour.Length <= 0) throw new SystemException("【品號管理】- 【品號分類(四)】不能為空!");
                if (InventoryId <= 0) throw new SystemException("【主要庫別】不能為空!");
                if (RequisitionInventoryId <= 0) throw new SystemException("【主要領料庫別】不能為空!");
                //if (InventoryUomId <= 0) throw new SystemException("【入庫單位】不能為空!");
                if (ItemAttribute.Length <= 0) throw new SystemException("【品號屬性】不能為空!");

                int? nullData = null;
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷品號類別是否正確
                        sql = @"SELECT TOP 1 1
                                FROM PDM.ItemType
                                WHERE ItemTypeId = @ItemTypeId";
                        dynamicParameters.Add("ItemTypeId", ItemTypeId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("品號類別資料錯誤!");
                        #endregion

                        #region //判斷品號類別預設質料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM PDM.ItemTypeDefault
                                WHERE ItemTypeDefaultId = @ItemTypeDefaultId";
                        dynamicParameters.Add("ItemTypeDefaultId", ItemTypeDefaultId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("品號類別預設質料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PDM.ItemTypeDefault SET
                                TypeOne = @TypeOne,
                                TypeTwo = @TypeTwo,
                                TypeThree = @TypeThree,
                                TypeFour = @TypeFour,
                                InventoryId = @InventoryId,
                                RequisitionInventoryId = @RequisitionInventoryId,
                                InventoryUomId = @InventoryUomId,
                                ItemAttribute = @ItemAttribute,
                                MeasureType = @MeasureType,
                                PurchaseUomId = @PurchaseUomId,
                                SaleUomId = @SaleUomId,
                                LotManagement = @LotManagement,
                                InventoryManagement = @InventoryManagement,
                                MtlModify = @MtlModify,
                                BondedStore = @BondedStore,
                                OverReceiptManagement = @OverReceiptManagement,
                                OverDeliveryManagement = @OverDeliveryManagement,
                                EffectiveDate = @EffectiveDate,
                                ExpirationDate = @ExpirationDate,
                                MtlItemDesc = @MtlItemDesc,
                                MtlItemRemark = @MtlItemRemark,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ItemTypeId = @ItemTypeId
                                AND ItemTypeDefaultId = @ItemTypeDefaultId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TypeOne,
                                TypeTwo,
                                TypeThree,
                                TypeFour,
                                InventoryId,
                                RequisitionInventoryId,
                                InventoryUomId,
                                ItemAttribute,
                                MeasureType = MeasureType != "" ? MeasureType : null,
                                PurchaseUomId = PurchaseUomId != -1 ? PurchaseUomId : nullData,
                                SaleUomId = SaleUomId != -1 ? SaleUomId : nullData,
                                LotManagement = LotManagement != "" ? LotManagement : null,
                                InventoryManagement = InventoryManagement != "" ? InventoryManagement : null,
                                MtlModify = MtlModify != "" ? MtlModify : null,
                                BondedStore = BondedStore != "" ? BondedStore : null,
                                OverReceiptManagement = OverReceiptManagement != "" ? OverReceiptManagement : null,
                                OverDeliveryManagement = OverDeliveryManagement != "" ? OverDeliveryManagement : null,
                                EffectiveDate = EffectiveDate != "" ? EffectiveDate : null,
                                ExpirationDate = ExpirationDate != "" ? ExpirationDate : null,
                                MtlItemDesc = MtlItemDesc != "" ? MtlItemDesc : null,
                                MtlItemRemark = MtlItemRemark != "" ? MtlItemRemark : null,
                                LastModifiedDate,
                                LastModifiedBy,
                                ItemTypeId,
                                ItemTypeDefaultId
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

        #region //UpdateDfmItemCategory -- 更新DFM項目種類代碼 -- Shintokuro 2023.08.28
        public string UpdateDfmItemCategory(int DfmItemCategoryId, int ModeId, string DfmItemCategoryName, string DfmItemCategoryDesc)
        {
            try
            {
                if (DfmItemCategoryId <= 0) throw new SystemException("【項目種類代碼】不能為空!");
                if (ModeId <= 0) throw new SystemException("【生產模式】不能為空!");
                if (DfmItemCategoryName.Length <= 0) throw new SystemException("【項目種類名稱】不能為空!");
                if (DfmItemCategoryName.Length > 100) throw new SystemException("【項目種類名稱】長度錯誤!");
                if (DfmItemCategoryDesc.Length <= 0) throw new SystemException("【項目種類描述】不能為空!");
                if (DfmItemCategoryDesc.Length > 100) throw new SystemException("【項目種類描述】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷項目種類代碼是否存在
                        sql = @"SELECT TOP 1 1
                                FROM PDM.DfmItemCategory
                                WHERE DfmItemCategoryId = @DfmItemCategoryId";
                        dynamicParameters.Add("DfmItemCategoryId", DfmItemCategoryId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【項目種類】不存在，請重新確認!");
                        #endregion

                        #region //判斷生產模式是否存在
                        sql = @"SELECT TOP 1 1
                                FROM MES.ProdMode
                                WHERE ModeId = @ModeId";
                        dynamicParameters.Add("ModeId", ModeId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【生產模式】不存在，請重新確認!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PDM.DfmItemCategory SET
                                ModeId = @ModeId,
                                DfmItemCategoryName = @DfmItemCategoryName,
                                DfmItemCategoryDesc = @DfmItemCategoryDesc,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE DfmItemCategoryId = @DfmItemCategoryId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ModeId,
                                DfmItemCategoryName,
                                DfmItemCategoryDesc,
                                LastModifiedDate,
                                LastModifiedBy,
                                DfmItemCategoryId
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

        #region //UpdateDfmMaterial -- 更新DFM物料代碼 -- Shintokuro 2023.08.28
        public string UpdateDfmMaterial(int DfmMaterialId, string DfmMaterialName, string DfmMaterialDesc, double DfmMaterialMoney)
        {
            try
            {
                if (DfmMaterialId <= 0) throw new SystemException("【物料代碼】不能為空!");
                if (DfmMaterialName.Length <= 0) throw new SystemException("【物料名稱】不能為空!");
                if (DfmMaterialName.Length > 100) throw new SystemException("【物料名稱】長度錯誤!");
                if (DfmMaterialDesc.Length > 100) throw new SystemException("【物料描述】長度錯誤!");
                if (DfmMaterialMoney < 0) throw new SystemException("【物料金額】不能為負!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷物料代碼是否存在
                        sql = @"SELECT TOP 1 1
                                FROM PDM.DfmMaterial
                                WHERE CompanyId = @CompanyId
                                AND DfmMaterialId = @DfmMaterialId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("DfmMaterialId", DfmMaterialId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【項目種類】不存在，請重新確認!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PDM.DfmMaterial SET
                                DfmMaterialName = @DfmMaterialName,
                                DfmMaterialDesc = @DfmMaterialDesc,
                                DfmMaterialMoney = @DfmMaterialMoney,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE DfmMaterialId = @DfmMaterialId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                DfmMaterialName,
                                DfmMaterialDesc,
                                DfmMaterialMoney,
                                LastModifiedDate,
                                LastModifiedBy,
                                DfmMaterialId
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

        #region //UpdateDfmItemCategoryStatus -- 更新DFM項目種類代碼狀態 -- Shintokuro 2023.08.28
        public string UpdateDfmItemCategoryStatus(int DfmItemCategoryId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷物料代碼資訊是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM PDM.DfmItemCategory
                                WHERE DfmItemCategoryId = @DfmItemCategoryId";
                        dynamicParameters.Add("DfmItemCategoryId", DfmItemCategoryId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("物料代碼資料錯誤!");

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
                        sql = @"UPDATE PDM.DfmItemCategory SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE DfmItemCategoryId = @DfmItemCategoryId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                DfmItemCategoryId
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

        #region //UpdateDfmCategoryTemplate1 -- 更新DfmCategoryTemplate資料 -- Shintokuro 2023-08-29
        public string UpdateDfmCategoryTemplate1(int DfmCtId, int DfmItemCategoryId, string DataType, string SpreadsheetData)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;
                        #region//檢核DFM【項目種類】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.DfmItemCategory a
                                WHERE a.DfmItemCategoryId=@DfmItemCategoryId";
                        dynamicParameters.Add("DfmItemCategoryId", DfmItemCategoryId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【項目種類】查無資料,請重新確認");
                        #endregion

                        switch (DataType)
                        {
                            case "Material":
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE PDM.DfmCategoryTemplate SET
                                        QiMaterialTemplate = @SpreadsheetData,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE DfmItemCategoryId = @DfmItemCategoryId
                                        AND DfmCtId = @DfmCtId";
                                break;
                            case "DfmQiOSP":
                                throw new SystemException("【委外加工】尚未開發，敬請期待~~");
                                break;
                            case "DfmQiProcess":
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE PDM.DfmCategoryTemplate SET
                                        QiProcessTemplate = @SpreadsheetData,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE DfmItemCategoryId = @DfmItemCategoryId
                                        AND DfmCtId = @DfmCtId";
                                break;
                            default:
                                throw new SystemException("資料類別異常，請重新確認!!");
                                break;
                        }
                        dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        SpreadsheetData,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        DfmItemCategoryId,
                                        DfmCtId
                                    });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

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


        #region //UpdateDfmMaterialStatus -- 更新DFM物料代碼狀態 -- Shintokuro 2023.08.28
        public string UpdateDfmMaterialStatus(int DfmMaterialId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷物料代碼資訊是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM PDM.DfmMaterial
                                WHERE CompanyId = @CompanyId
                                AND DfmMaterialId = @DfmMaterialId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("DfmMaterialId", DfmMaterialId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("物料代碼資料錯誤!");

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
                        sql = @"UPDATE PDM.DfmMaterial SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE DfmMaterialId = @DfmMaterialId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                DfmMaterialId
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

        #region //UpdateDfmCategoryMaterialStatus -- 更新DFM項目種類物料狀態 -- Xuan 2023.09.06
        public string UpdateDfmCategoryMaterialStatus(int DfmCategoryMaterialId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷物料代碼資訊是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM PDM.DfmCategoryMaterial
                                WHERE DfmCategoryMaterialId = @DfmCategoryMaterialId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("DfmCategoryMaterialId", DfmCategoryMaterialId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("項目種類物料資料錯誤!");

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
                        sql = @"UPDATE PDM.DfmCategoryMaterial SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE DfmCategoryMaterialId = @DfmCategoryMaterialId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                DfmCategoryMaterialId
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


        #region //UpdateMoldingConditionsUnitPrice -- 更新成型條件品號採購單價 -- Shintokuro 2025.05.07
        public string UpdateMoldingConditionsUnitPrice(
            )
        {
            try
            {
               
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    string MESCompanyNo = "";
                    string ErpDb = "";
                    List<StockInDetail> stockInDetail = new List<StockInDetail>();

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //公司別資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb
                            FROM BAS.Company a
                            WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var resultCompany = sqlConnection.Query(sql, dynamicParameters);
                        if (resultCompany.Count() <= 0) throw new SystemException("【公司別】資料錯誤!");

                        foreach (var item in resultCompany)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            MESCompanyNo = item.ErpNo;
                            ErpDb = item.ErpDb;
                        }
                        #endregion

                        //#region //撈出成型條件品號
                        //dynamicParameters = new DynamicParameters();
                        //sql = @"SELECT a.MtlItemId,b.MtlItemNo,a.UnitPrice
                        //        FROM PDM.MoldingConditions a
                        //        INNER JOIN PDM.MtlItem b on a.MtlItemId = b.MtlItemId";
                        //dynamicParameters.Add("Company", CurrentCompany);
                        //stockInDetail = sqlConnection.Query<StockInDetail>(sql, dynamicParameters).ToList();
                        //#endregion

                        #region //撈出成型條件品號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MtlItemId,y1.CellMtlItemNo,y1.CellMtlItemId
                                FROM PDM.MoldingConditions a
                                INNER JOIN PDM.MtlItem b on a.MtlItemId = b.MtlItemId
                                OUTER APPLY(
                                    SELECT TOP 1 x3.MtlItemNo CellMtlItemNo,x3.MtlItemId CellMtlItemId
                                    FROM PDM.BillOfMaterials x1
                                    INNER JOIN PDM.BomDetail x2 on x1.BomId =x2.BomId
                                    INNER jOIN PDM.MtlItem x3 on x2.MtlItemId = x3.MtlItemId
                                    INNER jOIN PDM.MtlItem x4 on x1.MtlItemId = x4.MtlItemId
                                    WHERE 1=1
                                   AND (
                                        (x2.EffectiveDate IS NOT NULL AND x2.ExpirationDate IS NOT NULL AND GETDATE() BETWEEN x2.EffectiveDate AND x2.ExpirationDate)
                                        OR
                                        (x2.EffectiveDate IS NULL AND x2.ExpirationDate IS NULL)
                                        OR
                                        (x2.EffectiveDate IS NULL AND x2.ExpirationDate IS NOT NULL AND GETDATE() <= x2.ExpirationDate)
                                        OR
                                        (x2.EffectiveDate IS NOT NULL AND x2.ExpirationDate IS NULL AND GETDATE() >= x2.EffectiveDate)
                                    )
                                    AND a.MtlItemId = x1.MtlItemId
                                ) y1";
                        dynamicParameters.Add("Company", CurrentCompany);
                        stockInDetail = sqlConnection.Query<StockInDetail>(sql, dynamicParameters).ToList();
                        #endregion

                    }

                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        #region //公司別資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"
                                SELECT a.TD004 MtlItemNo,y1.TD010 UnitPrice
                                FROM (
                                SELECT DISTINCT TD004 
                                FROM PURTD ) a
                                OUTER APPLY(
                                    SELECT TOP 1 x1.TD010
                                    FROM PURTD x1
                                    INNER JOIN PURTC x2 on x1.TD001 = x2.TC001 AND x1.TD002 = x2.TC002
                                    WHERE a.TD004 = x1.TD004
                                    AND TC004 != '100C0448' AND TC004 != '100C0086'
                                    AND TC004 != '100A0011' AND TC004 != '100A0037'
                                    ORDER BY x2.TC003 DESC
                                ) y1";
                        List<StockInDetail> resultMtlitems = sqlConnection.Query<StockInDetail>(sql, dynamicParameters).ToList();
                        stockInDetail = stockInDetail
                        .Join(resultMtlitems,
                              x => x.CellMtlItemNo,
                              y => y.MtlItemNo,
                              (x, y) =>
                              {
                                  x.UnitPrice = y.UnitPrice;
                                  return x;
                              })
                        .ToList();
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PDM.MoldingConditions SET
                                UnitPrice = @UnitPrice,
                                CellMtlItemId = @CellMtlItemId,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE MtlItemId = @MtlItemId";
                        int rowsAffected = sqlConnection.Execute(sql, stockInDetail.Select(x => new
                        {
                            UnitPrice = x.UnitPrice,
                            CellMtlItemId = x.CellMtlItemId,
                            LastModifiedDate,
                            LastModifiedBy,
                            MtlItemId = x.MtlItemId
                        }));
                        

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

        #region //UpdateItemCategoryDept -- 更新品號類別責任部門 --Jean 2025-05-25 --
        public string UpdateItemCategoryDept(int ItemCatDeptId, string TypeTwo, string TypeThree, string InventoryNo, string CatDept)
        {
            try
            {
                if (InventoryNo.Length <= 0) throw new SystemException("【組合三】不能為空!");
                if (CatDept.Length <= 0) throw new SystemException("【責任部門】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷品號類別責任部門資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM PDM.ItemCatDept
                                WHERE ItemCatDeptId = @ItemCatDeptId";
                        dynamicParameters.Add("ItemCatDeptId", ItemCatDeptId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("品號類別責任部門資料錯誤!");
                        #endregion
                       

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PDM.ItemCatDept SET
                                TypeTwo = @TypeTwo,
                                TypeThree = @TypeThree,
                                InventoryNo = @InventoryNo,
                                CatDept = @CatDept,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE ItemCatDeptId = @ItemCatDeptId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TypeTwo,
                                TypeThree,
                                InventoryNo,
                                CatDept,
                                LastModifiedDate,
                                LastModifiedBy,
                                ItemCatDeptId
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

        #region //UpdateMoldingConditions -- 更新成型產品生產條件  --Jean 2025-05-25 --
        public string UpdateMoldingConditions(int McId, string MtlItemId, string Material, string CycleTime , string MoldWeight 
            , string Cavity , string ProcessingFee , string TheoreticalQty )
        {
            try
            {

                if (MtlItemId.Length <= 0) throw new SystemException("【ERP品號】不能為空!");
                if (Material.Length <= 0) throw new SystemException("【材料】不能為空!");
                if (CycleTime.Length <= 0) throw new SystemException("【週期】不能為空!");
                if (MoldWeight.Length <= 0) throw new SystemException("【模重】不能為空!");
                if (Cavity.Length <= 0) throw new SystemException("【穴數】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷成型產品生產條件資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM PDM.MoldingConditions
                                WHERE McId = @McId";
                        dynamicParameters.Add("McId", McId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("成型產品生產條件資料錯誤!");
                        #endregion


                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PDM.MoldingConditions SET
                                MtlItemId = @MtlItemId,
                                Material = @Material,
                                CycleTime = @CycleTime,
                                MoldWeight = @MoldWeight,
                                Cavity = @Cavity,
                                ProcessingFee = @ProcessingFee,
                                TheoreticalQty = @TheoreticalQty,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE McId = @McId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                MtlItemId,
                                Material,
                                CycleTime,
                                MoldWeight,
                                Cavity,
                                ProcessingFee,
                                TheoreticalQty,
                                LastModifiedDate,
                                LastModifiedBy,
                                McId
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

        #endregion
         
        #region //Delete
        #region //DeleteMtlModel -- 品號機型資料刪除 -- Kan 2022.06.28
        public string DeleteMtlModel(int MtlModelId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷品號機型資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.MtlModel
                                WHERE CompanyId = @CompanyId
                                AND MtlModelId = @MtlModelId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("MtlModelId", MtlModelId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("品號機型資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除關聯table
                        #region //關聯資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"DECLARE @rowsAdded int

                                DECLARE @mtlModel TABLE
                                ( 
                                    MtlModelId int,
                                    ParentId int,
                                    MtlModelLevel int,
                                    MtlModelRoute nvarchar(MAX),
                                    MtlModelSort nvarchar(MAX),
                                    processed int DEFAULT(0)
                                )

                                INSERT @mtlModel
                                    SELECT MtlModelId, ParentId, 1 MtlModelLevel
                                    , CAST(ParentId AS nvarchar(MAX)) AS MtlModelRoute
                                    , CAST(MtlModelSort AS nvarchar(MAX)) AS MtlModelSort, 0
                                    FROM PDM.MtlModel
                                    WHERE CompanyId = @CompanyId
                                    AND ParentId = @ParentId

                                SET @rowsAdded=@@rowcount

                                WHILE @rowsAdded > 0
                                BEGIN

                                UPDATE @mtlModel SET processed = 1 WHERE processed = 0

                                INSERT @mtlModel
                                    SELECT a.MtlModelId, a.ParentId, ( b.MtlModelLevel + 1 ) MtlModelLevel
                                    , CAST(b.MtlModelRoute + ',' + CAST(a.ParentId AS nvarchar(MAX)) AS nvarchar(MAX)) AS MtlModelRoute
                                    , CAST(b.MtlModelSort + CAST(a.MtlModelSort AS nvarchar(MAX)) AS nvarchar(MAX)) AS MtlModelSort, 0
                                    FROM PDM.MtlModel a
                                    INNER JOIN @mtlModel b ON a.ParentId = b.MtlModelId
                                    WHERE a.ParentId <> a.MtlModelId 
                                    AND b.processed = 1

                                SET @rowsAdded = @@rowcount

                                UPDATE @mtlModel SET processed = 2 WHERE processed = 1

                                END

                                SELECT *
                                FROM @mtlModel";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("ParentId", MtlModelId);

                        var mtlModel = sqlConnection.Query(sql, dynamicParameters).ToList();
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PDM.MtlModel
                                WHERE MtlModelId IN @MtlModel";
                        dynamicParameters.Add("MtlModel", mtlModel.Select(x => x.MtlModelId).ToArray());

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PDM.MtlModel
                                WHERE MtlModelId = @MtlModelId";
                        dynamicParameters.Add("MtlModelId", MtlModelId);

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

        #region //DeleteItemTypeDefault -- 品號類別預設資料刪除 -- Shintokuro 2023.04.17
        public string DeleteItemTypeDefault(int ItemTypeId, int ItemTypeDefaultId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷品號類別是否正確
                        sql = @"SELECT TOP 1 1
                                FROM PDM.ItemType
                                WHERE ItemTypeId = @ItemTypeId";
                        dynamicParameters.Add("ItemTypeId", ItemTypeId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("品號類別資料錯誤!");
                        #endregion

                        #region //判斷品號類別是否正確
                        sql = @"SELECT TOP 1 1
                                FROM PDM.ItemTypeDefault
                                WHERE ItemTypeDefaultId = @ItemTypeDefaultId";
                        dynamicParameters.Add("ItemTypeDefaultId", ItemTypeDefaultId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("品號類別資料錯誤!");
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除關聯table
                        #region //關聯資料
                        #endregion
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PDM.ItemTypeDefault
                                WHERE ItemTypeDefaultId = @ItemTypeDefaultId";
                        dynamicParameters.Add("ItemTypeDefaultId", ItemTypeDefaultId);

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

        #region //DeleteDfmItemCategory -- 刪除DFM項目種類代碼 -- Shintokuro 2023.08.28
        public string DeleteDfmItemCategory(int DfmItemCategoryId)
        {
            try
            {
                if (DfmItemCategoryId <= 0) throw new SystemException("DFM項目種類代碼不可以為空，請重新確認");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;

                        #region //判斷項目種類代碼是否存在
                        sql = @"SELECT TOP 1 1
                                FROM PDM.DfmQuotationItem
                                WHERE DfmItemCategoryId = @DfmItemCategoryId";
                        dynamicParameters.Add("DfmItemCategoryId", DfmItemCategoryId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【項目種類】已經被使用，無法刪除!");
                        #endregion

                        #region //刪除附屬table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PDM.DfmCategoryTemplate
                                WHERE DfmItemCategoryId = @DfmItemCategoryId";
                        dynamicParameters.Add("DfmItemCategoryId", DfmItemCategoryId);
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PDM.DfmItemCategory
                                WHERE DfmItemCategoryId = @DfmItemCategoryId";
                        dynamicParameters.Add("DfmItemCategoryId", DfmItemCategoryId);

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

        #region //DeleteDfmMaterial -- 刪除DFM物料代碼 -- Shintokuro 2023.08.28
        public string DeleteDfmMaterial(int DfmMaterialId)
        {
            try
            {
                if (DfmMaterialId <= 0) throw new SystemException("DFM物料代碼不可以為空，請重新確認");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;

                        #region //判斷物料代碼是否存在
                        sql = @"SELECT TOP 1 1
                                FROM PDM.DfmQiMaterial
                                WHERE DfmMaterialId = @DfmMaterialId";
                        dynamicParameters.Add("DfmMaterialId", DfmMaterialId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【物料代碼】已經被使用，無法刪除!");
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PDM.DfmMaterial
                                WHERE DfmMaterialId = @DfmMaterialId";
                        dynamicParameters.Add("DfmMaterialId", DfmMaterialId);

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

        #region //DeleteDfmCategoryTemplate -- 刪除DfmQiExcelData暫存記錄 -- Shintokuro 2023-08-29
        public string DeleteDfmCategoryTemplate(int DfmItemCategoryId, string DataType)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;
                        #region//檢核DFM【報價項目】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.DfmItemCategory a
                                WHERE a.DfmItemCategoryId=@DfmItemCategoryId";
                        dynamicParameters.Add("DfmItemCategoryId", DfmItemCategoryId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【項目種類】查無資料,請重新確認");
                        #endregion

                        #region //更新PDM.DfmQuotationItem SpreadsheetData
                        switch (DataType)
                        {
                            case "Material":
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE PDM.DfmCategoryTemplate SET
                                        QiMaterialTemplate = null,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE DfmItemCategoryId = @DfmItemCategoryId";

                                break;
                            case "DfmQiOSP":
                                throw new SystemException("【委外加工】尚未開發，敬請期待~~");
                                break;
                            case "DfmQiProcess":
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE PDM.DfmCategoryTemplate SET
                                        QiProcessTemplate = null,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE DfmItemCategoryId = @DfmItemCategoryId";
                                break;
                            default:
                                throw new SystemException("項目種類異常，請重新確認!!");
                                break;
                        }
                        dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        DfmItemCategoryId
                                    });
                        #endregion


                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);


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

        #region //DeleteDfmCategoryMaterial -- 刪除DFM項目種類物料 -- Xuan 2023.09.06
        public string DeleteDfmCategoryMaterial(int DfmCategoryMaterialId)
        {
            try
            {
                
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;

                        #region //判斷物料代碼是否存在
                        sql = @"SELECT TOP 1 1
                                FROM PDM.DfmCategoryMaterial
                                WHERE DfmCategoryMaterialId = @DfmCategoryMaterialId";
                        dynamicParameters.Add("DfmCategoryMaterialId", DfmCategoryMaterialId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("該筆資料不存在，無法刪除!");
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PDM.DfmCategoryMaterial
                                WHERE DfmCategoryMaterialId = @DfmCategoryMaterialId";
                        dynamicParameters.Add("DfmCategoryMaterialId", DfmCategoryMaterialId);

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

        #region //DeleteItemCatDept -- 刪除品號類別責任部門 --Jean 2025-05-25 --
        public string DeleteItemCatDept(int ItemCatDeptId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷托盤資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM PDM.ItemCatDept
                                WHERE ItemCatDeptId = @ItemCatDeptId";
                        dynamicParameters.Add("ItemCatDeptId", ItemCatDeptId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("品號類別責任部門資料錯誤!");
                        #endregion

                        

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PDM.ItemCatDept
                                WHERE ItemCatDeptId = @ItemCatDeptId";
                        dynamicParameters.Add("ItemCatDeptId", ItemCatDeptId);

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

        #region //DeleteMoldingConditions -- 刪除成型產品生產條件 --Jean 2025-05-25 --
        public string DeleteMoldingConditions(int McId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷托盤資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM PDM.MoldingConditions
                                WHERE McId = @McId";
                        dynamicParameters.Add("McId", McId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("品號類別責任部門資料錯誤!");
                        #endregion



                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PDM.MoldingConditions
                                WHERE McId = @McId";
                        dynamicParameters.Add("McId", McId);

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
    }
}
