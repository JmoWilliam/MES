using Dapper;
using Helpers;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Transactions;
using System.Web;

namespace SCMDA
{
    public class RequestForQuotationDA
    {
        public static string MainConnectionStrings = "";
        public static string ErpConnectionStrings = "";
        public static string HrmConnectionStrings = "";

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
        public MamoHelper mamoHelper = new MamoHelper();

        public RequestForQuotationDA()
        {
            MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];
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
        #region //GetRequestForQuotation -- 取得客戶詢價資訊管理(RFQ)資訊 -- Yi 2023.07.07
        public string GetRequestForQuotation(int RfqId, string RfqNo, int MemberId, string MemberName, string AssemblyName, int ProductUseId
            , int SalesId, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.RfqId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.RfqNo, a.AssemblyName, a.MemberId, a.ProductUseId, b.ProductUseName
                        , ISNULL(c.MemberId, d.UserId) MemberUserId, ISNULL(c.MemberName, d.UserName) MemberUserName, c.MemberName
                        , CASE WHEN a.MemberType = 1 THEN '客戶填單' WHEN a.MemberType = 2 THEN '業務填單' END MemberTypeName
                        , a.OrganizaitonType, a.CustomerId, a.CustomerName, a.SupplierId, a.SupplierName, a.ContactPerson, a.ContactPhone
                        , a.Status RfqStatus,f.CustomerName CustomerCompanyName
                        , (
                            SELECT aa.RfqId, aa.RfqDetailId, aa.RfqSequence, ab.RfqProductTypeName, aa.MtlName, aa.CustProdDigram, ISNULL(FORMAT(aa.PlannedOpeningDate, 'yyyy-MM-dd'), '') PlannedOpeningDate
                            , aa.PrototypeQty, ISNULL(ac.TypeName, '') ProtoScheduleName
                            , aa.MassProductionDemand, ah.StatusName MassProductionDemandName
                            , ISNULL(aa.KickOffType, '') KickOffType, aa.PlasticName, aa.OutsideDiameter
                            , aa.ProdLifeCycleStart, aa.ProdLifeCycleEnd, aa.LifeCycleQty, aa.DemandDate
                            , aa.CoatingFlag, aa.Currency, ai.StatusName CoatingFlagName
                            , ISNULL(aa.SalesId, -1) SalesId, ISNULL(ad.UserName, '') SalesName, ISNULL((ag.CompanyName + '-' + ad.UserName), '') SalesInfo
                            , aa.AdditionalFile, aa.[Status] RfqDetailStatus, ae.StatusName RfqDetailStatusName, ISNULL(FORMAT(aa.ConfirmSalesTime, 'yyyy-MM-dd HH:mm:ss'), '') ConfirmSalesTime
                            , ISNULL(FORMAT(aa.ConfirmRdTime, 'yyyy-MM-dd HH:mm:ss'), '') ConfirmRdTime, aa.Description, aa.Edition, aa.DocType, aa.ExpirationDate
                            , aj.TypeName DocTypeName
                            FROM SCM.RfqDetail aa
                            LEFT JOIN SCM.RfqProductType ab ON ab.RfqProTypeId = aa.RfqProTypeId
                            LEFT JOIN BAS.[Type] ac ON ac.TypeNo = aa.ProtoSchedule AND ac.TypeSchema = 'RfqDetail.ProtoSchedule'
                            LEFT JOIN BAS.[User] ad ON ad.UserId = aa.SalesId
                            INNER JOIN BAS.[Status] ae ON ae.StatusNo = aa.[Status] AND ae.StatusSchema = 'RfqDetail.Status'
                            LEFT JOIN BAS.Department af ON af.DepartmentId = ad.DepartmentId
                            LEFT JOIN BAS.Company ag ON ag.CompanyId = af.CompanyId
                            LEFT JOIN BAS.[Status] ah ON ah.StatusNo = aa.MassProductionDemand AND ah.StatusSchema = 'Boolean'
                            INNER JOIN BAS.[Status] ai ON ai.StatusNo = aa.CoatingFlag AND ai.StatusSchema = 'Boolean'
                            LEFT JOIN BAS.[Type] aj ON aa.DocType = aj.TypeNo AND aj.TypeSchema = 'RfqDetail.DocType'
                            WHERE aa.RfqId = a.RfqId";
                    //用於判斷單身多筆不同指派人員，需濾掉其他業務，只能看自己的單據
                    //if (SalesId > 0)
                    //{
                    //    sqlQuery.columns += @" AND EXISTS (
                    //                                SELECT TOP 1 1
                    //                                FROM SCM.RfqDetail ba
                    //                                INNER JOIN SCM.RequestForQuotation bb ON bb.RfqId = ba.RfqId
                    //                                WHERE ba.RfqId = a.RfqId
                    //                                AND ba.SalesId = @SalesId
                    //                           )";
                    //}

                    sqlQuery.columns +=
                        @"  ORDER BY aa.RfqDetailId
                            FOR JSON PATH, ROOT('data')
                        ) RfqDetail";
                    sqlQuery.mainTables =
                        @"FROM SCM.RequestForQuotation a
                        LEFT JOIN SCM.ProductUse b ON b.ProductUseId = a.ProductUseId
                        LEFT JOIN EIP.Member c ON c.MemberId = a.MemberId
                        LEFT JOIN BAS.[User] d ON d.UserId = a.UserId
                        LEFT JOIN BAS.[Type] e ON e.TypeNo = a.MemberType AND e.TypeSchema = 'RequestForQuotation.MemberType'
                        LEFT JOIN SCM.Customer f ON a.CustomerId = f.CustomerId";

                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RfqId", @" AND a.RfqId = @RfqId", RfqId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RfqNo", @" AND a.RfqNo = @RfqNo", RfqNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MemberName", @" AND c.MemberName LIKE '%' + @MemberName + '%'", MemberName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "AssemblyName", @" AND a.AssemblyName LIKE '%' + @AssemblyName + '%'", AssemblyName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ProductUseId", @" AND a.ProductUseId = @ProductUseId", ProductUseId);
                    //BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SalesId", @" AND EXISTS (
                    //                                                                                            SELECT TOP 1 1
                    //                                                                                            FROM SCM.RfqDetail ba
                    //                                                                                            INNER JOIN SCM.RequestForQuotation bb ON bb.RfqId = ba.RfqId
                    //                                                                                            WHERE ba.RfqId = a.RfqId
                    //                                                                                            AND ba.SalesId = @SalesId)", SalesId);
                    if (Status.Length > 0)
                    {
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND EXISTS (
                                                                                                                SELECT TOP 1 1
                                                                                                                FROM SCM.RequestForQuotation aa
                                                                                                                INNER JOIN SCM.RfqDetail ab ON ab.RfqId = aa.RfqId
                                                                                                                WHERE aa.RfqId = a.RfqId
                                                                                                                AND ab.Status IN @Status)", Status.Split(','));
                    }
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.RfqId DESC";
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

        #region //GetRfqDetail -- 取得RFQ單身資訊 -- Yi 2023.07.10
        public string GetRfqDetail(int RfqId, int RfqDetailId,string DocType
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.RfqDetailId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.RfqId, a.RfqSequence, b.RfqNo, b.AssemblyName, b.ProductUseId, c.ProductUseName
                        , ISNULL(d.MemberId, m.UserId) MemberUserId, ISNULL(d.MemberName, m.UserName) MemberUserName
                        , CASE 
                            WHEN b.MemberType = 1 THEN '客戶填單' 
                            WHEN b.MemberType = 2 THEN '業務填單' 
                        END MemberTypeName
                        , b.Status RfqStatus
                        , b.OrganizaitonType, b.CustomerId, b.CustomerName, b.SupplierId, b.SupplierName,b.Status RfqStatus
                        , a.CompanyId, a.RfqSequence, a.RfqProTypeId, e.RfqProductTypeName, a.MtlName, a.CustProdDigram, ISNULL(FORMAT(a.PlannedOpeningDate, 'yyyy-MM-dd'), '') PlannedOpeningDate
                        , a.PrototypeQty, a.ProtoSchedule, ISNULL(f.TypeName, '') ProtoScheduleName
                        , a.MassProductionDemand, k.StatusName MassProductionDemandName
                        , ISNULL(a.KickOffType, '') KickOffType, ISNULL(a.PlasticName, '') PlasticName, ISNULL(a.OutsideDiameter, '') OutsideDiameter
                        , FORMAT(a.ProdLifeCycleStart, 'yyyy-MM-dd') ProdLifeCycleStart, FORMAT(a.ProdLifeCycleEnd, 'yyyy-MM-dd') ProdLifeCycleEnd
                        , a.LifeCycleQty, FORMAT(a.DemandDate, 'yyyy-MM-dd') DemandDate
                        , a.CoatingFlag, a.Currency, a.MonthlyQty, a.SampleQty, a.UomId, l.StatusName CoatingFlagName, a.Status RfqDetailStatus, a.BaseCavities, a.InsertCavities, a.CoreThickness, a.CommonMode
                        , ISNULL(a.SalesId, -1) SalesId, g.UserName SalesName, ISNULL((j.CompanyName + '-' + g.UserName), '') SalesInfo, a.AdditionalFile, h.StatusName RfqDetailStatusName
                        , ISNULL(FORMAT(a.ConfirmSalesTime, 'yyyy-MM-dd HH:mm:ss'), '') ConfirmSalesTime, ISNULL(FORMAT(a.ConfirmRdTime, 'yyyy-MM-dd HH:mm:ss'), '') ConfirmRdTime, a.Description
                        , a.Edition, a.DocType, a.ExpirationDate,a.QuotationStatus, o.TypeName DocTypeName
                        , (
                            SELECT 
                                aa.CompanyId, aa.RfqSequence, aa.RfqProTypeId, ae.RfqProductTypeName, aa.MtlName, aa.CustProdDigram
                                , ISNULL(FORMAT(aa.PlannedOpeningDate, 'yyyy-MM-dd'), '') PlannedOpeningDate, aa.PrototypeQty, aa.ProtoSchedule 
                                , ISNULL(af.TypeName, '') ProtoScheduleName, aa.MassProductionDemand,  aj.StatusName MassProductionDemandName
                                , ISNULL(aa.KickOffType, '') KickOffType, ISNULL(aa.PlasticName, '') PlasticName, ISNULL(aa.OutsideDiameter, '') OutsideDiameter
                                , FORMAT(aa.ProdLifeCycleStart, 'yyyy-MM-dd') ProdLifeCycleStart, FORMAT(aa.ProdLifeCycleEnd, 'yyyy-MM-dd') ProdLifeCycleEnd
                                , aa.LifeCycleQty, FORMAT(aa.DemandDate, 'yyyy-MM-dd') DemandDate, aa.CoatingFlag, ak.StatusName CoatingFlagName, aa.Currency, aa.MonthlyQty, aa.SampleQty, aa.UomId
                                , aa.Status RfqDetailStatus, ah.StatusName RfqDetailStatusName, ISNULL(aa.SalesId, -1) SalesId, ag.UserName SalesName
                                , ISNULL((ai.CompanyName + '-' + ag.UserName), '') SalesInfo,aa.QuotationStatus
                                , aa.AdditionalFile, aa.QuotationFile, ISNULL(FORMAT(aa.ConfirmSalesTime, 'yyyy-MM-dd HH:mm:ss'), '') ConfirmSalesTime
                                , ISNULL(FORMAT(aa.ConfirmRdTime, 'yyyy-MM-dd HH:mm:ss'), '') ConfirmRdTime, aa.Description, aa.Edition, aa.DocType, aa.ExpirationDate
                                , al.TypeName DocTypeName
                            FROM SCM.RfqDetail aa
                            LEFT JOIN BAS.[File] ab ON aa.CustProdDigram = ab.FileId
                            LEFT JOIN BAS.[File] ac ON aa.AdditionalFile = ac.FileId
                            LEFT JOIN BAS.[File] ad ON aa.QuotationFile = ad.FileId
                            LEFT JOIN SCM.RfqProductType ae ON aa.RfqProTypeId = ae.RfqProTypeId
                            LEFT JOIN BAS.[Type] af ON aa.ProtoSchedule = af.TypeNo AND af.TypeSchema = 'RfqDetail.ProtoSchedule'
                            LEFT JOIN BAS.[User] ag ON aa.SalesId = ag.UserId
                            LEFT JOIN BAS.[Status] ah ON aa.[Status] = ah.StatusNo AND ah.StatusSchema = 'RfqDetail.Status'
                            LEFT JOIN BAS.Company ai ON aa.CompanyId = ai.CompanyId
                            LEFT JOIN BAS.[Status] aj ON aa.MassProductionDemand = aj.StatusNo AND aj.StatusSchema = 'Boolean'
                            LEFT JOIN BAS.[Status] ak ON aa.CoatingFlag = ak.StatusNo AND ak.StatusSchema = 'Boolean'
                            LEFT JOIN BAS.[Type] al ON aa.DocType = al.TypeNo AND al.TypeSchema = 'RfqDetail.DocType'
                            WHERE aa.RfqDetailId = a.RfqDetailId
                            FOR JSON PATH, ROOT('data')
                        ) RfqDetailData
                        , (
                            SELECT 
                                bb.RfqPkTypeId, bb.PackagingMethod, ba.SustSupplyStatus, bc.StatusName SustSupplyStatusName, bd.RfqProductClassName
                            FROM SCM.RfqPackage ba
                            INNER JOIN SCM.RfqPackageType bb ON bb.RfqPkTypeId = ba.RfqPkTypeId
                            INNER JOIN BAS.[Status] bc ON bc.StatusNo = ba.SustSupplyStatus AND bc.StatusSchema = 'Boolean'
                            INNER JOIN SCM.RfqProductClass bd ON bd.RfqProClassId = bb.RfqProClassId
                            WHERE ba.RfqDetailId = a.RfqDetailId
                            FOR JSON PATH, ROOT('data')
                        ) PackagingMethod
                        , (
                            SELECT 
                                ca.RfqLineSolutionId, ca.SortNumber, ca.SolutionQty, ca.PeriodicDemandType, cc.TypeName PeriodicDemandTypeName
                            FROM SCM.RfqLineSolution ca
                            INNER JOIN SCM.RfqDetail cb ON ca.RfqDetailId = cb.RfqDetailId
                            INNER JOIN BAS.[Type] cc ON cc.TypeNo = ca.PeriodicDemandType AND cc.TypeSchema = 'RfqDetail.PeriodicDemandType'
                            WHERE ca.RfqDetailId = a.RfqDetailId
                            FOR JSON PATH, ROOT('data')
                        ) RfqLineSolution";
                    sqlQuery.mainTables =
                        @"FROM SCM.RfqDetail a
                        INNER JOIN SCM.RequestForQuotation b ON b.RfqId = a.RfqId
                        LEFT JOIN SCM.ProductUse c ON c.ProductUseId = b.ProductUseId
                        LEFT JOIN EIP.Member d ON d.MemberId = b.MemberId
                        LEFT JOIN SCM.RfqProductType e ON e.RfqProTypeId = a.RfqProTypeId
                        LEFT JOIN BAS.[Type] f ON f.TypeNo = a.ProtoSchedule AND f.TypeSchema = 'RfqDetail.ProtoSchedule'
                        LEFT JOIN BAS.[User] g ON g.UserId = a.SalesId
                        INNER JOIN BAS.[Status] h ON h.StatusNo = a.[Status] AND h.StatusSchema = 'RfqDetail.Status'
                        LEFT JOIN BAS.Department i ON i.DepartmentId = g.DepartmentId
                        LEFT JOIN BAS.Company j ON j.CompanyId = i.CompanyId
                        LEFT JOIN BAS.[Status] k ON k.StatusNo = a.MassProductionDemand AND k.StatusSchema = 'Boolean'
                        INNER JOIN BAS.[Status] l ON l.StatusNo = a.CoatingFlag AND l.StatusSchema = 'Boolean'
                        LEFT JOIN BAS.[User] m ON m.UserId = b.UserId
                        LEFT JOIN BAS.[Type] n ON n.TypeNo = b.MemberType AND n.TypeSchema = 'RequestForQuotation.MemberType'
                        LEFT JOIN BAS.[Type] o ON a.DocType = o.TypeNo AND o.TypeSchema = 'RfqDetail.DocType'
                        ";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"";

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RfqId", @" AND a.RfqId = @RfqId", RfqId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RfqDetailId", @" AND a.RfqDetailId = @RfqDetailId", RfqDetailId);
                    if (DocType.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DocType", @" AND a.DocType IN @DocType", DocType.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.RfqId DESC";
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

        #region //GetSales -- 取得RFQ負責業務(Cmb用) -- Yi 2023.07.10
        public string GetSales(int UserId, string UserName, int CompanyId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.UserId, a.UserName, b.DepartmentName, c.CompanyId, c.CompanyName
                            , (c.CompanyName + '-' + a.UserName) SalesInfo
                            FROM BAS.[User] a
                            INNER JOIN BAS.Department b ON b.DepartmentId = a.DepartmentId
                            INNER JOIN BAS.Company c ON c.CompanyId = b.CompanyId
                            WHERE a.UserStatus != 'S'
                            AND a.UserName != '來賓'
                            AND (b.DepartmentName LIKE '%業務%' OR b.DepartmentName LIKE N'%业务%')
                            AND c.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "UserId", @" AND a.UserId = @UserId", UserId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "UserName", @" AND a.UserName LIKE '%' + @UserName + '%'", UserName);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CompanyId", @" AND c.CompanyId = @CompanyId", CompanyId);
                    sql += @"
                            ORDER BY a.UserId";
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

        #region //GetSalesQuery -- 取得RFQ負責業務(Query用) -- Yi 2023.07.12
        public string GetSalesQuery(int UserId, string UserName)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.UserId, a.UserName, b.DepartmentName, c.CompanyId, c.CompanyName
                            , (c.CompanyName + '-' + a.UserName) SalesInfo
                            FROM BAS.[User] a
                            INNER JOIN BAS.Department b ON b.DepartmentId = a.DepartmentId
                            INNER JOIN BAS.Company c ON c.CompanyId = b.CompanyId
                            WHERE a.UserStatus != 'S'
                            AND a.UserName != '來賓'
                            AND (b.DepartmentName LIKE '%業務%' OR b.DepartmentName LIKE N'%业务%')";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "UserId", @" AND a.UserId = @UserId", UserId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "UserName", @" AND a.UserName LIKE '%' + @UserName + '%'", UserName);
                    sql += @"
                            ORDER BY a.UserId";
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

        #region //GetRfqLineSolution -- 取得RFQ報價方案資料 -- Yi 2023.07.28
        public string GetRfqLineSolution(int RfqDetailId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.RfqLineSolutionId, a.RfqDetailId, a.SortNumber, a.SolutionQty,b.Currency
                            FROM SCM.RfqLineSolution a
                            INNER JOIN SCM.RfqDetail b ON a.RfqDetailId=b.RfqDetailId
                            WHERE 1=1";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "RfqDetailId", @" AND a.RfqDetailId = @RfqDetailId", RfqDetailId);
                    sql += @"
                            ORDER BY a.SolutionQty ASC";
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

        #region //GetDfmQiSolution -- 取得RFQ報價資料 -- Yi 2023.07.24
        public string GetDfmQiSolution(int RfqLineSolutionId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"DECLARE @rowsAdded int

                            DECLARE @DfmQuotationItem TABLE
                            (
	                            DfmQiId int,
	                            ParentDfmQiId int,
	                            SolutionLevel int,
	                            SolutionRoute nvarchar(MAX),
	                            SolutionSort int,
                                processed int DEFAULT(0)
                            )

                            INSERT @DfmQuotationItem
	                            SELECT a.DfmQiId, ISNULL(a.ParentDfmQiId, -1) ParentDfmQiId, 1 SolutionLevel
	                            , CAST(CASE WHEN ParentDfmQiId = -1 THEN DfmQiId ELSE ParentDfmQiId END AS nvarchar(MAX)) AS SolutionRoute
	                            , a.DfmQiId, 0
	                            FROM PDM.DfmQuotationItem a
	                            INNER JOIN PDM.DesignForManufacturing b ON b.DfmId = a.DfmId
	                            INNER JOIN SCM.RfqLineSolution d ON d.RfqDetailId = b.RfqDetailId
	                            WHERE d.RfqLineSolutionId = @RfqLineSolutionId
	                            AND a.ParentDfmQiId = -1

                            SET @rowsAdded=@@rowcount

                            WHILE @rowsAdded > 0
                            BEGIN
                                UPDATE @DfmQuotationItem SET processed = 1 WHERE processed = 0

	                            INSERT @DfmQuotationItem
                                    SELECT b.DfmQiId, b.ParentDfmQiId, ( a.SolutionLevel + 1 ) TaskLevel
                                    , CAST(a.SolutionRoute + ',' + CAST(b.ParentDfmQiId AS nvarchar(MAX)) AS nvarchar(MAX)) AS TaskRoute
                                    , b.DfmQiId, 0
                                    FROM @DfmQuotationItem a
                                    INNER JOIN PDM.DfmQuotationItem b ON a.DfmQiId = b.ParentDfmQiId
                                    WHERE a.processed = 1

                                SET @rowsAdded = @@rowcount

                                UPDATE @DfmQuotationItem SET processed = 2 WHERE processed = 1
                            END;

                            SELECT c.DfmId, a.DfmQiId, a.ParentDfmQiId, b.RfqLineSolutionId, a.SolutionLevel
                            , a.SolutionRoute, a.SolutionSort, a.processed, c.DfmQuotationName
                            , (b.MaterialAmt + b.ResourceAmt + b.OverheadAmt + b.StandardCost) CalculateCost
                            , b.DiscountAmount, b.GrossProfitMargin, b.AfterProfitMargin, b.QuotationAmount
                            , ((b.MaterialAmt + b.ResourceAmt + b.OverheadAmt + b.StandardCost) * b.GrossProfitMargin) CalculateProfitMargin
                            , (((b.MaterialAmt + b.ResourceAmt + b.OverheadAmt + b.StandardCost) * b.GrossProfitMargin) - b.DiscountAmount) CalculateQuotation
                            , e.[Status] RfqDetailStatus
                            FROM @DfmQuotationItem a
                            INNER JOIN PDM.DfmQiSolution b ON b.DfmQiId = a.DfmQiId
                            INNER JOIN PDM.DfmQuotationItem c ON c.DfmQiId = b.DfmQiId
                            INNER JOIN SCM.RfqLineSolution d ON d.RfqLineSolutionId = b.RfqLineSolutionId
                            INNER JOIN SCM.RfqDetail e ON e.RfqDetailId = d.RfqDetailId
                            WHERE b.RfqLineSolutionId = @RfqLineSolutionId
                            ORDER BY a.SolutionRoute, a.SolutionSort";
                    dynamicParameters.Add("RfqLineSolutionId", RfqLineSolutionId);
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

        #region//GetQuotationHead -- 報價單單頭 -- Luca 2023.07.27
        public string GetQuotationHead(int RfqDetailId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT b.CustomerName, b.Contact, b.TelNoFirst, b.GuiNumber, b.RegisterAddressFirst
                            , b.RegisterAddressFirst, b.TelNoFirst, b.FaxNo, b.GuiNumber
                            , (a.RfqNo+'-'+b.CustomerName) QuotationFileName
                            FROM SCM.RequestForQuotation a
                            INNER JOIN SCM.Customer b ON a.CustomerId=b.CustomerId
                            INNER JOIN SCM.RfqDetail c ON a.RfqId=c.RfqId
                            WHERE 1=1";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "RfqDetailId", @" AND c.RfqDetailId = @RfqDetailId", RfqDetailId);
                    sql += @" ORDER BY c.RfqDetailId ";
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

        #region//GetQuotationHead -- 報價單單頭 -- Luca 2023.07.27
        public string GetQuotationLine(int RfqDetailId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT DISTINCT c.RfqDetailId, b.DfmQuotationName,(
                                SELECT z.SolutionQty, x.QuotationAmount
                                FROM PDM.DfmQiSolution x
                                INNER JOIN PDM.DfmQuotationItem y ON x.DfmQiId=y.DfmQiId
                                INNER JOIN SCM.RfqLineSolution z ON x.RfqLineSolutionId=z.RfqLineSolutionId
                                WHERE  y.DfmQuotationName=b.DfmQuotationName AND x.QuotationAmount!=0 
                                AND z.RfqDetailId=c.RfqDetailId
                                ORDER BY z.SolutionQty ASC
                                FOR JSON PATH, ROOT('data')
                            )Detail
                            FROM PDM.DfmQiSolution a
                                INNER JOIN PDM.DfmQuotationItem b ON a.DfmQiId=b.DfmQiId
                                INNER JOIN SCM.RfqLineSolution c ON a.RfqLineSolutionId=c.RfqLineSolutionId
                            WHERE c.RfqDetailId=@RfqDetailId 
                            AND b.QuotationLevel=1
                            ORDER BY b.DfmQuotationName";
                    dynamicParameters.Add("@RfqDetailId", RfqDetailId);                    
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

        #region//GetExchangeRate -- 取得ERP匯率 -- Luca 2023.08.14
        public string GetExchangeRate(string Currency)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT TOP 1 Currency,BankBuyingRate,BankSellingRate,CustomsBuyingRate,CustomsSellingRate,EffectiveDate
                            FROM SCM.ExchangeRate
                            WHERE 1=1";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "Currency", @" AND Currency = @Currency", Currency);
                    sql += @" ORDER BY EffectiveDate DESC ";
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

        #region //GetSalesDefault -- 取得RFQ系統預設客戶對應之業務人員 -- Yi 2023.08.21
        public string GetSalesDefault(int RfqDetailId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.SalesId, a.CompanyId, b.UserName, c.DepartmentName, d.CompanyName
                            , (d.CompanyName + '-' + b.UserName) SalesInfo
                            FROM SCM.RfqDetail a
                            INNER JOIN BAS.[User] b ON b.UserId = a.SalesId
                            INNER JOIN BAS.Department c ON c.DepartmentId = b.DepartmentId
                            INNER JOIN BAS.Company d ON d.CompanyId = c.CompanyId
                            WHERE a.RfqDetailId = @RfqDetailId";
                    dynamicParameters.Add("RfqDetailId", RfqDetailId);
                    sql += @"
                            ORDER BY a.SalesId";
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

        #region //GetQuotationDataModel -- 取得報價單樣板資料 -- Shintokurp --2024-08-12
        public string GetQuotationDataModel(int RfqId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"
                            SELECT a.HtmlColumns,a.HtmlColumns +'. ' +a.HtmlColName Columns,ISNULL(x.DetailName,'') DetailName
                            FROM SCM.QuotationHtmlElement a
                            INNER JOIN SCM.RfqDetail b on a.RfqProTypeId = b.RfqProTypeId
                            OUTER APPLY(
                                SELECT STUFF((SELECT ','+ Convert(nvarchar(MAX),(x1.DetailName)) 
                                FROM SCM.QuotationHtmlElementDetail x1
                                WHERE x1.QeId = a.QeId
                                AND x1.[Status] = 'A' 
                                AND x1.DetailName != 'RMB$'
                                FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 1, '') AS DetailName
                            ) x
                            WHERE b.RfqId = @RfqId
                            AND a.[Status] = 'A'
                            AND b.DocType = 'E'
                            ORDER BY b.RfqDetailId, a.HtmlColumns
                           ";

                    dynamicParameters.Add("RfqId", RfqId);
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

        #region //GetQuotationDetail -- 取得報價單資料 -- Shintokurp --2024-08-12
        public string GetQuotationDetail(int RfqDetailId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"
                            SELECT 
                                a2.RfqNo +'-'+ a1.RfqSequence + '('+ a1.Edition + ')' RfqFullNo,
                                a2.AssemblyName,
                                a3.CustomerName,a3.CustomerNo,
                                a1.RfqId,
                                a.RfqDetailId,
                                a.QiId,
                                a.Charge,
                                a.ConfirmStatus,
                                a.ColRemark,
                                b.UserName ConfirmUserName,
                                a1.RfqProTypeId, c.RfqProductTypeName,
                                a1.MtlName,
                                a1.PlasticName,
                                a1.OutsideDiameter,
                                a1.MassProductionDemand,
                                a1.CoatingFlag,
                                a1.CommonMode,
                                a1.MonthlyQty,
                                a1.BaseCavities,
                                a1.InsertCavities,
                                a1.CoreThickness,
                                a1.UomId, d.UomNo,
                                a1.QuotationStatus,
                                a1.Currency,
                                a1.Description,
                                a1.CustProdDigram,
                                a1.TagList,
                                a1.QuotationRemark,
                                ISNULL(a1.CycleTimeAdoption,'') CycleTimeAdoption,
                                ISNULL(a1.AICycleTime,'') AICycleTime,
                                a.HtmlColumns,
                                a.HtmlColumns + '. ' + a.HtmlColName AS Columns,
                                (
                                    SELECT x1.Sort,QieId,ItemElementNo,ItemElementName,QuoteValue,ColumnSetting,Formula,ElementRemark,x1.Flag
                                    FROM SCM.QuotationItemElement x1
                                    WHERE x1.QiId = a.QiId
                                    ORDER BY x1.Sort
                                    FOR JSON PATH
                                ) Detail
                            FROM SCM.QuotationItem a
                            INNER JOIN SCM.RfqDetail a1 ON a.RfqDetailId = a1.RfqDetailId
                            INNER JOIN SCM.RequestForQuotation a2 ON a1.RfqId = a2.RfqId
                            LEFT JOIN SCM.Customer a3 ON a2.CustomerId = a3.CustomerId
                            LEFT JOIN BAS.[User] b ON a.ConfirmUserId = b.UserId
                            INNER JOIN SCM.RfqProductType c ON a1.RfqProTypeId = c.RfqProTypeId
                            LEFT JOIN PDM.UnitOfMeasure d ON a1.UomId = d.UomId
                            WHERE a1.RfqDetailId = @RfqDetailId
                            AND a1.DocType = 'E'
                            ORDER BY a.RfqDetailId, a.HtmlColumns
                           ";

                    dynamicParameters.Add("RfqDetailId", RfqDetailId);
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

        #region //GetQuotationBargainPrice -- 取得議價資料 -- Shintokurp --2024-09-12
        public string GetQuotationBargainPrice(int RfqDetailId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    string[] FinalItem = new string[2];
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT STRING_AGG(x.ItemElementNo, ',') FlagSetting
                            FROM SCM.QuotationItemElement x
                            INNER JOIN SCM.QuotationItem x1 on x.QiId = x1.QiId
                            WHERE x1.RfqDetailId = @RfqDetailId
                            AND x.Flag = 'FR'";
                    dynamicParameters.Add("RfqDetailId", RfqDetailId);
                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0) throw new SystemException("報價單資料找不到,請重新確認!");
                    foreach (var item in result)
                    {
                        FinalItem[0] = ((item.FlagSetting).Split(','))[0];
                        FinalItem[1] = ((item.FlagSetting).Split(','))[1];
                    }

                    sql = @"
                            SELECT x1.QuoteValue NowGrossMargin,x2.QuoteValue NowFinalPrice
                            , ISNULL(x3.GrossMargin,0) GrossMargin,ISNULL(x3.FinalPrice,0) FinalPrice,ISNULL(x3.ConfirmStatus,'N') ConfirmStatus
                            FROM  SCM.RfqDetail a
                            OUTER APPLY(
                                SELECT x2.QuoteValue 
                                FROM SCM.QuotationItem x1
                                INNER JOIN SCM.QuotationItemElement x2 on x1.QiId = x2.QiId
                                WHERE x1.RfqDetailId = a.RfqDetailId
                                AND x2.ItemElementNo = @FinalItem1
                            ) x1
                            OUTER APPLY(
                                SELECT x2.QuoteValue 
                                FROM SCM.QuotationItem x1
                                INNER JOIN SCM.QuotationItemElement x2 on x1.QiId = x2.QiId
                                WHERE x1.RfqDetailId = a.RfqDetailId
                                AND x2.ItemElementNo = @FinalItem2
                            ) x2
                            OUTER APPLY(
                                SELECT x1.FinalPrice,x1.GrossMargin ,x1.ConfirmStatus
                                FROM SCM.QuotationFinalPrice x1
                                WHERE x1.ConfirmStatus != 'H'
                                AND x1.RfqDetailId = a.RfqDetailId
                            ) x3
                            WHERE a.RfqDetailId = @RfqDetailId
                            AND a.QuotationStatus = 'F'
                            AND DocType = 'E'
                           ";
                    dynamicParameters.Add("RfqDetailId", RfqDetailId);
                    dynamicParameters.Add("FinalItem1", FinalItem[0]);
                    dynamicParameters.Add("FinalItem2", FinalItem[1]);
                    result = sqlConnection.Query(sql, dynamicParameters);

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

        #region //GetQuotationHistory -- 取得報價單歷史資料 -- Shintokurp --2024-08-12
        public string GetQuotationHistory(int RfqDetailId, int RfqProTypeId, string TagList, string PreciseMode)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    string condition = "";
                    if (TagList.Length > 0)
                    {
                        if(PreciseMode != "Y")
                        {
                            condition += " AND (";
                            foreach (var Id in TagList.Split(','))
                            {
                                condition += "',' + a.TagList + ',' LIKE '%," + Id + ",%' OR";
                            }
                            condition = condition.Substring(0, condition.Length - 2);
                            condition += ")";
                        }
                        else
                        {
                            condition += " AND a.TagList = '" + TagList + @"'";
                        }
                    }
                    

                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sql = @"
                            SELECT 
                            a.RfqDetailId,a1.RfqNo +'-'+ a.MtlName RfqFullNo, 'N' 'All'
                            FROM SCM.RfqDetail a
                            INNER JOIN SCM.RequestForQuotation a1 on a.RfqId = a1.RfqId
                            OUTER APPLY(
                                SELECT TOP 1 x1.QiId haveData 
                                FROM SCM.QuotationItem x1
                                WHERE x1.RfqDetailId =a.RfqDetailId
                            ) x
                            WHERE 1=1
                            AND a.RfqProTypeId = @RfqProTypeId
                            AND a.RfqDetailId != @RfqDetailId
                            AND a.DocType = 'E'
                            AND a.QuotationStatus = 'F'
                            AND x.haveData is not null
                            " + condition + @"
                           ";

                    dynamicParameters.Add("RfqDetailId", RfqDetailId);
                    dynamicParameters.Add("RfqProTypeId", RfqProTypeId);
                    var result = sqlConnection.Query(sql, dynamicParameters);

                    if(result.Count() <= 0 && PreciseMode == "N")
                    {
                        dynamicParameters = new DynamicParameters();
                        sql = @"
                            SELECT 
                            a.RfqDetailId,a1.RfqNo +'-'+ a.MtlName RfqFullNo, 'Y' 'All'
                            FROM SCM.RfqDetail a
                            INNER JOIN SCM.RequestForQuotation a1 on a.RfqId = a1.RfqId
                            OUTER APPLY(
                                SELECT TOP 1 x1.QiId haveData 
                                FROM SCM.QuotationItem x1
                                WHERE x1.RfqDetailId =a.RfqDetailId
                            ) x
                            WHERE 1=1
                            AND a.RfqProTypeId = @RfqProTypeId
                            AND a.RfqDetailId != @RfqDetailId
                            AND a.DocType = 'E'
                            AND a.QuotationStatus = 'F'
                            AND x.haveData is not null
                           ";
                        dynamicParameters.Add("RfqDetailId", RfqDetailId);
                        dynamicParameters.Add("RfqProTypeId", RfqProTypeId);
                        result = sqlConnection.Query(sql, dynamicParameters);
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

        #region //GetIncomeStatement -- 取得損益表 -- Shintokurp --2024-08-29
        public string GetIncomeStatement(int RfqId, string TagList, string MtlItemNo, string MtlItemName
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                string departmentNo = "", MESCompanyNo = "", userNo = "", userName = "";
                string CustomerNo = "";

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
                    }
                    #endregion

                    #region //使用者資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 d.UserNo, d.UserName, d.DepartmentNo
                                FROM BAS.FunctionDetail a
                                INNER JOIN BAS.[Function] b ON a.FunctionId = b.FunctionId
                                OUTER APPLY (
                                    SELECT ISNULL((
                                        SELECT TOP 1 1
                                        FROM BAS.RoleFunctionDetail ca
                                        WHERE ca.DetailId = a.DetailId
                                        AND ca.RoleId IN (
                                            SELECT caa.RoleId
                                            FROM BAS.UserRole caa
                                            INNER JOIN BAS.[Role] cab ON caa.RoleId = cab.RoleId
                                            WHERE caa.UserId = @UserId
                                            AND cab.CompanyId = @CompanyId
                                        )
                                    ), 0) Authority
                                ) c
                                OUTER APPLY (
                                    SELECT da.UserNo, da.UserName, db.DepartmentNo
                                    FROM BAS.[User] da
                                    INNER JOIN BAS.Department db ON da.DepartmentId = db.DepartmentId
                                    WHERE da.UserId = @UserId
                                ) d
                                WHERE a.[Status] = @Status
                                AND b.[Status] = @Status
                                AND b.FunctionCode = @FunctionCode
                                AND a.DetailCode = @DetailCode
                                AND c.Authority > 0";
                    dynamicParameters.Add("UserId", CurrentUser);
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    dynamicParameters.Add("Status", "A");
                    dynamicParameters.Add("FunctionCode", "RfqAssignManagment");
                    dynamicParameters.Add("DetailCode", "quote-income");

                    var resultUser = sqlConnection.Query(sql, dynamicParameters);
                    if (resultUser.Count() <= 0) throw new SystemException("使用者資料錯誤!");

                    foreach (var item in resultUser)
                    {
                        userNo = item.UserNo;
                        userName = item.UserName;
                        departmentNo = item.DepartmentNo;
                    }
                    #endregion

                    #region //撈取客戶
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.CustomerId, b.CustomerNo
                                FROM SCM.RequestForQuotation a
                                INNER JOIN SCM.Customer b on a.CustomerId = b.CustomerId
                                WHERE RfqId = @RfqId";
                    dynamicParameters.Add("RfqId", RfqId);

                    var resultCustomer = sqlConnection.Query(sql, dynamicParameters);
                    if (resultCustomer.Count() <= 0) throw new SystemException("【RFQ】資料找不到，請重新確認!");

                    foreach (var item in resultCustomer)
                    {
                        CustomerNo = item.CustomerNo;
                    }
                    #endregion

                    //#region //撈取品號
                    //if(MtlItemNo.Length > 0 || MtlItemName.Length > 0)
                    //{
                    //    dynamicParameters = new DynamicParameters();
                    //    sql = @"SELECT a.MtlItemNo
                    //            FROM PDM.MtlItem a
                    //            WHERE 1 = 1";
                    //    if(MtlItemNo.Length > 0)
                    //    {
                    //        sql += "MtlItemNo = @MtlItemNo";
                    //        dynamicParameters.Add("MtlItemNo", MtlItemNo);
                    //    }
                    //    if (MtlItemName.Length > 0)
                    //    {
                    //        sql += "MtlItemName = @MtlItemName";
                    //        dynamicParameters.Add("MtlItemName", MtlItemName);
                    //    }

                    //    var resultMtlItem = sqlConnection.Query(sql, dynamicParameters);
                    //    if (resultMtlItem.Count() <= 0) throw new SystemException("品號資料找不到，請重新確認!");
                    //}
                    //#endregion
                }

                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();

                    #region //撈取關帳日期
                    string CloseDate = "";
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT LTRIM(RTRIM(MA013)) MA013
                            FROM CMSMA";
                    var cmsmaResult = sqlConnection.Query(sql, dynamicParameters);
                    if (cmsmaResult.Count() < 0) throw new SystemException("EPR關帳日期資料錯誤!!");

                    foreach (var item in cmsmaResult)
                    {
                        CloseDate = item.MA013;
                    }
                    #endregion


                    #region //損益資料表
                    sqlQuery.mainKey = "x1.LA005";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @",x2.MB002,x2.MB003,x1.NetSales,x1.Average,x1.Cost,x1.Gross
                      ";
                    sqlQuery.mainTables =
                        @"FROM(
                            SELECT a.LA005
                            ,SUM(LA017-LA020-LA022) AS NetSales
                            ,Average =CASE WHEN(SUM(LA017-LA020-LA022)=0 OR SUM(LA016-LA019+LA025+LA026)=0) THEN 0 ELSE (SUM(LA017-LA020-LA022)/SUM(LA016-LA019+LA025+LA026)) END
                            ,SUM(LA024) AS Cost 
                            ,SUM(LA017-LA020-LA022-LA024-LA023) AS Gross 
                            FROM SASLA a
                            LEFT JOIN COPMA c ON c.MA001=a.LA006 
                            LEFT JOIN INVMA d  ON a.LA001=d.MA002 AND d.MA001='1'
                            WHERE 1=1  
                            AND a.LA006= @LA006  
                            AND (a.LA015 <= @LA015 OR a.LA015 IS NULL ) 
                            Group by a.LA005
                        ) x1
                        INNER JOIN INVMB x2  ON x1.LA005=x2.MB001";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"";
                    dynamicParameters.Add("LA006", CustomerNo);
                    dynamicParameters.Add("LA015", CloseDate);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemNo", @" AND  x1.LA005 LIKE '%'+ @MtlItemNo + '%' ", MtlItemNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemName", @" AND  x2.MB002 LIKE '%'+ @MtlItemName + '%' ", MtlItemName);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "x1.LA005 DESC";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery);
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

        #region //GetQuotationAuthority -- 取得報價單相關權限 -- Shintokurp --2024-09-06
        public string GetQuotationAuthority(string UserNo)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {

                    #region //使用者資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.FunctionCode,b.DetailCode,b.DetailName ,
                            CASE WHEN x.Authority = 'Y' THEN 'Y' ELSE  'N' END Authority
                            FROM BAS.[Function] a
                            INNER JOIN BAS.FunctionDetail b on a.FunctionId = b.FunctionId
                            OUTER APPLY(
                                SELECT TOP 1 'Y' Authority
                                FROM BAS.[User] x1
                                INNER JOIN BAS.UserRole x2 on x1.UserId = x2.UserId
                                INNER JOIN BAS.[Role] x3 on x2.RoleId = x3.RoleId
                                INNER JOIN BAS.RoleFunctionDetail x4 on x2.RoleId = x4.RoleId
                                INNER JOIN BAS.FunctionDetail x5 on x4.DetailId = x5.DetailId
                                INNER JOIN BAS.[Function] x6 on x5.FunctionId = x6.FunctionId
                                WHERE UserNo = @UserNo
                                AND x6.FunctionCode = a.FunctionCode
                                AND b.DetailId = x5.DetailId
                                AND x3.CompanyId  = @CompanyId
                            ) x
                            WHERE a.FunctionCode = 'RfqAssignManagment'";
                    dynamicParameters.Add("UserNo", UserNo);
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0) throw new SystemException("使用者資料錯誤!");
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

        #region //GetQuotationPdf -- 取得報價單PDF資料 -- Shintokurp --2024-09-24
        public string GetQuotationPdf(int RfqId, int RfqDetailId, string RfqIdList, string ExchangeStatus)
        {
            try
            {
                string PaymentTerm = "";
                string MESCompanyNo = "";
                string DocDate = "";
                string Currency = "";
                string CustomerNo = "";
                //RfqId = 2116;
                //RfqDetailId = -1;
                dynamic result = new ExpandoObject();
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
                    }
                    #endregion
                    string SearchKey = "";
                    if(RfqId > 0)
                    {
                        SearchKey = "AND b.RfqId = @RfqId";
                    }
                    if (RfqDetailId > 0)
                    {
                        SearchKey = "AND b.RfqDetailId = @RfqDetailId";
                    }
                    dynamicParameters = new DynamicParameters();
                    sql = @"
                            SELECT a.RfqId,b.RfqDetailId,b.RfqSequence,b.QuotationNo,a.ContactPerson,a.ContactPhone,b.MtlName,b.MonthlyQty,b.Currency,b.QuotationRemark,c.UomNo,d.CustomerNo,d.CustomerName
                            ,b.InsertCavities,b.CoreThickness,b.QuotationStatus,FORMAT(b.CreateDate,'yyyyMM') DocDate
                            ,x.QuotationDate,d.PaymentTerm
                            ,x.FinalPrice UnitPrice
                            ,ROUND((b.MonthlyQty * x.FinalPrice), 4) Amount
                            ,y1.CurrentUser,y2.SetupCost ,e.UserName CreateUser
                            FROM SCM.RequestForQuotation a
                            INNER JOIN SCM.RfqDetail b on a.RfqId = b.RfqId
                            INNER JOIN PDM.UnitOfMeasure c on b.UomId = c.UomId
                            INNER JOIN SCM.Customer d on a.CustomerId = d.CustomerId
                            INNER JOIN BAS.[User] e on a.UserId = e.UserId
                            OUTER APPLY(
                                SELECT TOP 1 FORMAT(x1.CreateDate,'yyyy-MM-dd') QuotationDate,x1.FinalPrice
                                FROM SCM.QuotationFinalPrice x1
                                WHERE x1.RfqDetailId = b.RfqDetailId
                                ORDER BY x1.QfpId DESC
                            ) x
                            OUTER APPLY(
                                SELECT x1.UserName CurrentUser 
                                FROM BAS.[User] x1
                                WHERE x1.UserId = @CreateBy
                            ) y1
                            OUTER APPLY(
                                SELECT  x1.QuoteValue SetupCost
                                FROM SCM.QuotationItemElement x1
                                INNER JOIN SCM.QuotationItem x2 on x1.QiId = x2.QiId
                                WHERE x2.RfqDetailId = b.RfqDetailId
                                AND x1.Flag = 'SC'
                            ) y2
                            WHERE b.DocType = 'E'
                            " + SearchKey + @"
                           ";

                    dynamicParameters.Add("RfqId", RfqId);
                    dynamicParameters.Add("RfqDetailId", RfqDetailId);
                    dynamicParameters.Add("CreateBy", CurrentUser);
                    result = sqlConnection.Query(sql, dynamicParameters);

                    string NoPass = "";
                    foreach(var item in result)
                    {
                        if(RfqId> 0)
                        {
                            if (item.QuotationStatus != "F")
                            {
                                NoPass += item.RfqSequence + ",";
                                continue;
                            }
                        }
                        else
                        {
                            if (item.QuotationStatus != "F") throw new SystemException("經管中心尚未確認不能列印報價單!!!");
                        }
                        PaymentTerm = item.PaymentTerm;
                        DocDate = item.DocDate;
                        Currency = item.Currency;
                        CustomerNo = item.CustomerNo;
                    }
                    if(NoPass.Length> 0)
                    {
                        NoPass = NoPass.Substring(0, NoPass.Length - 1);
                        throw new SystemException("單據流水編號:" + NoPass + ",經管中心尚未確認不能列印報價單!!!");
                    }
                }

                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();

                    #region //撈取付款條件
                    string PaymentTermName = "";
                    string ExchangeRate = "";
                    string Taxation = "";
                    string TaxRate = "";
                    //DocDate = "20240801";
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT LTRIM(RTRIM(NA002)) PaymentTermNo, LTRIM(RTRIM(NA003)) PaymentTermName 
                            ,y2.Taxation,y3.TaxRate,y3.Currency
                            ,CASE 
                                WHEN y3.Currency = @Currency THEN 1
                                ELSE  y1.ExchangeRate
                            END ExchangeRate
                            FROM CMSNA
                            OUTER APPLY(
                                SELECT MG001,MG002,MG004 ExchangeRate
                                FROM CMSMG
                                WHERE MG002 LIKE '%" + DocDate + @"%'
                                AND MG001 = @Currency
                            ) y1
                            OUTER APPLY(
                                SELECT MA038 Taxation  
                                FROM COPMA 
                                WHERE MA001 = @CustomerNo
                            ) y2
                            OUTER APPLY(
                                SELECT MA004 TaxRate,MA003 Currency
                                FROM CMSMA 
                            ) y3
                            WHERE 1=1
                            AND NA002 = @PaymentTerm";
                    dynamicParameters.Add("PaymentTerm", PaymentTerm);
                    dynamicParameters.Add("Currency", Currency);
                    dynamicParameters.Add("CustomerNo", CustomerNo);

                    var cmsmaResult = sqlConnection.Query(sql, dynamicParameters);
                    if (cmsmaResult.Count() < 0) throw new SystemException("EPR關帳日期資料錯誤!!");

                    foreach (var item in cmsmaResult)
                    {
                        PaymentTermName = item.PaymentTermName;
                        ExchangeRate = item.ExchangeRate.ToString();
                        Taxation = item.Taxation.ToString();
                        TaxRate = item.TaxRate.ToString();
                        Currency = item.Currency.ToString();
                    }
                    foreach (var item in result)
                    {
                        var dictionaryItem = (IDictionary<string, object>)item;
                        dictionaryItem["PaymentTermName"] = PaymentTermName;
                        dictionaryItem["ExchangeRate"] = ExchangeRate;
                        dictionaryItem["Taxation"] = Taxation;
                        dictionaryItem["TaxRate"] = TaxRate;
                        
                        if(ExchangeStatus == "N")
                        {
                            dictionaryItem["Currency"] = Currency;
                        }
                    }
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

        #region //GetQuotationWork -- 取得報價單維護清單(工作平台列表) -- Shintokurp --2024-09-25
        public string GetQuotationWork(int RfqId, string RfqNo, int CustomerId, int RfqProTypeId, string RoleMode, string Status
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    #region //撈取登入者工號
                    string CreateByUserNo = "";
                    sql = @"SELECT TOP 1 UserNo
                                FROM BAS.[User]
                                WHERE UserId = @CurrentUser";
                    dynamicParameters.Add("CurrentUser", CurrentUser);

                    var resultUser = sqlConnection.Query(sql, dynamicParameters);
                    foreach (var item in resultUser)
                    {
                        CreateByUserNo = item.UserNo;
                    }
                    #endregion

                    string RoleCondition = "";
                    switch (RoleMode)
                    {
                        case "1":
                            //RoleCondition = "AND x1.ConfirmStatus = @ConfirmStatus AND (x1.Charge LIKE '%" + CurrentUser + @"%' OR x1.Charge LIKE '%" + CreateByUserNo + @"%') AND x2.QuotationStatus = 'N'";
                            if (Status == "N")
                            {
                                RoleCondition = "AND x1.ConfirmStatus = @ConfirmStatus AND (x1.Charge LIKE '%" + CurrentUser + @"%' OR x1.Charge LIKE '%" + CreateByUserNo + @"%')";
                            }
                            else
                            {
                                RoleCondition = "AND x1.ConfirmStatus = @ConfirmStatus AND (x1.Charge LIKE '%" + CurrentUser + @"%' OR x1.Charge LIKE '%" + CreateByUserNo + @"%')";
                            }
                            break;
                        case "2":
                            //撿核是否為經管中心角
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT x1.UserId,x1.UserNo,x1.UserName
                                    FROM BAS.[User] x1
                                    INNER JOIN BAS.UserRole x2 on x1.UserId = x2.UserId
                                    INNER JOIN BAS.[Role] x3 on x2.RoleId = x3.RoleId
                                    INNER JOIN BAS.RoleFunctionDetail x4 on x2.RoleId = x4.RoleId
                                    INNER JOIN BAS.FunctionDetail x5 on x4.DetailId = x5.DetailId
                                    INNER JOIN BAS.[Function] x6 on x5.FunctionId = x6.FunctionId
                                    WHERE x5.DetailCode = 'quote-review'
                                    AND x6.FunctionCode = 'RfqAssignManagment'
                                    AND x1.UserId  = @UserId
                                    AND x3.CompanyId  = @CompanyId
                                    AND x3.AdminStatus != 'Y'";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("UserId", CurrentUser);
                            var result1 = sqlConnection.Query(sql, dynamicParameters);
                            if (result1.Count() <= 0) throw new SystemException("非報價單-經管中心角色,不能查看!!");

                            if(Status == "N")
                            {
                                RoleCondition = "AND x2.QuotationStatus = 'Y'";
                            }
                            else
                            {
                                RoleCondition = "AND x2.QuotationStatus = 'F'";
                            }
                            break;
                        case "3":
                            //撿核是否為業務角色
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT x1.UserId,x1.UserNo,x1.UserName
                                    FROM BAS.[User] x1
                                    INNER JOIN BAS.UserRole x2 on x1.UserId = x2.UserId
                                    INNER JOIN BAS.[Role] x3 on x2.RoleId = x3.RoleId
                                    INNER JOIN BAS.RoleFunctionDetail x4 on x2.RoleId = x4.RoleId
                                    INNER JOIN BAS.FunctionDetail x5 on x4.DetailId = x5.DetailId
                                    INNER JOIN BAS.[Function] x6 on x5.FunctionId = x6.FunctionId
                                    WHERE x5.DetailCode = 'add'
                                    AND x6.FunctionCode = 'RfqAssignManagment'
                                    AND x1.UserId  = @UserId
                                    AND x3.CompanyId  = @CompanyId
                                    AND x3.AdminStatus != 'Y'";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("UserId", CurrentUser);
                            var result2 = sqlConnection.Query(sql, dynamicParameters);
                            if (result2.Count() <= 0) throw new SystemException("非報價單-擔當角色,不能查看!!");
                            if (Status == "N")
                            {
                                RoleCondition = "AND x2.QuotationStatus != 'F'";
                            }
                            else
                            {
                                RoleCondition = "AND x2.QuotationStatus = 'F'";
                            }
                            break;
                    }

                    sqlQuery.mainKey = "a.RfqDetailId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @" ,a.ConfirmStatus
                           ,b.MtlName,b.PlasticName,b.OutsideDiameter,b.QuotationStatus
                           ,b1.RfqNo,b1.RfqId
                           ,c.RfqProductTypeName
                           ,d.CustomerName
                          ";
                    sqlQuery.mainTables =
                        @"FROM
                          (
                              SELECT DISTINCT x1.RfqDetailId,x1.ConfirmStatus
                              FROM SCM.QuotationItem x1
                              INNER JOIN SCM.RfqDetail x2 on x1.RfqDetailId =x2.RfqDetailId
                              INNER JOIN SCM.RequestForQuotation x3 on x2.RfqId =x3.RfqId
                              WHERE 1=1
                              AND x2.DocType = 'E'
                              " + RoleCondition + @"
                          ) a
                          INNER JOIN SCM.RfqDetail b on a.RfqDetailId = b.RfqDetailId
                          INNER JOIN SCM.RequestForQuotation b1 on b.RfqId = b1.RfqId
                          INNER JOIN SCM.RfqProductType c on b.RfqProTypeId = c.RfqProTypeId
                          INNER JOIN SCM.Customer d on b1.CustomerId = d.CustomerId
                         ";
                    string queryTable = "";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"";
                    dynamicParameters.Add("ConfirmStatus", Status);


                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RfqNo", @" AND b1.DeviceId = @RfqNo", RfqNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerId", @" AND b1.CustomerId = @CustomerId", CustomerId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RfqProTypeId", @" AND b.RfqProTypeId = @RfqProTypeId", RfqProTypeId);


                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.RfqDetailId DESC";
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
        #region //AddQuotationFile-- 新增報價檔案 -- Luca 2023-08-02
        public string AddQuotationFile(byte[] FileContent, string FileName, string FileExtension, int FileSize, string ClientIP, string Source)
        {
            try
            {
                if (FileContent.Length <= 0) throw new SystemException("【檔案Binary】不能為空!");
                if (FileName.Length <= 0) throw new SystemException("【檔案名稱】不能為空!");
                if (FileExtension.Length <= 0) throw new SystemException("【檔案副檔名】不能為空!");
                if (FileSize <= 0) throw new SystemException("【檔案大小】不能為空!");
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO BAS.[File] (CompanyId, FileName, FileContent, FileExtension, FileSize, ClientIP, Source, DeleteStatus
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.FileId
                                VALUES (@CompanyId, @FileName, CONVERT(varbinary(max),@FileContent), @FileExtension, @FileSize, @ClientIP, @Source, @DeleteStatus
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                FileName,
                                FileContent,
                                FileExtension,
                                FileSize,
                                ClientIP,
                                Source,
                                DeleteStatus = "N",
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

        #region //AddRequestForQuotation -- RFQ單頭資料新增 -- Yi 2023.10.11
        public string AddRequestForQuotation(string RfqNo, string MemberName, string AssemblyName, int ProductUseId
            , int MemberId, string OrganizaitonType, int CustomerId, string CustomerName, int SupplierId, string SupplierName
            , string ContactPerson, string ContactPhone)
        {
            try
            {
                //if (MemberName.Length <= 0) throw new SystemException("【客戶名稱】不能為空!");
                if (AssemblyName.Length <= 0) throw new SystemException("【機種名稱】不能為空!");
                if(CustomerId <= 0 && OrganizaitonType == "1") throw new SystemException("【客戶】不能為空!");
                if (ProductUseId <= 0) throw new SystemException("【產品應用】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //RFQ單號取號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT CONVERT(int, RIGHT(ISNULL(MAX(RfqNo), '000'), 3)) + 1 CurrentNum
                                FROM SCM.RequestForQuotation
                                WHERE RfqNo LIKE @RfqNo";
                        dynamicParameters.Add("RfqNo", string.Format("{0}{1}___", "RFQ", DateTime.Now.ToString("yyyyMMdd")));
                        int currentNum = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).CurrentNum;

                        string rfqNo = string.Format("{0}{1}{2}", "RFQ", DateTime.Now.ToString("yyyyMMdd"), string.Format("{0:000}", currentNum));
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.RequestForQuotation (CompanyId, RfqNo, AssemblyName, ProductUseId
                                , MemberType, MemberId, UserId, OrganizaitonType, CustomerId, CustomerName, SupplierId, SupplierName, Status
                                , ContactPerson, ContactPhone
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.RfqId, INSERTED.RfqNo
                                VALUES (@CompanyId, @RfqNo, @AssemblyName, @ProductUseId
                                , @MemberType, @MemberId, @UserId, @OrganizaitonType, @CustomerId, @CustomerName, @SupplierId, @SupplierName, @Status
                                , @ContactPerson, @ContactPhone
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                RfqNo = rfqNo,
                                AssemblyName,
                                ProductUseId,
                                MemberType = 2,
                                MemberId = MemberId == -1 ? (int?)null : MemberId,
                                UserId = CreateBy,
                                OrganizaitonType,
                                CustomerId = CustomerId == -1 ? (int?)null : CustomerId,
                                CustomerName,
                                SupplierId = SupplierId == -1 ? (int?)null : SupplierId,
                                SupplierName,
                                ContactPerson,
                                ContactPhone,
                                Status = 0, //未送出
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

        #region //AddRfqDetail -- RFQ單身資料新增(包含RFQ設計變更) -- Yi 2023.10.12
        public string AddRfqDetail(int RfqId, int CompanyId, int RfqProTypeId, string RfqSequence, string MtlName
            , int CustProdDigram, string PlannedOpeningDate, int PrototypeQty, string ProtoSchedule, string MassProductionDemand
            , string KickOffType, string PlasticName, string OutsideDiameter, string ProdLifeCycleStart, string ProdLifeCycleEnd
            , int LifeCycleQty, string DemandDate, string CoatingFlag, string Currency, int MonthlyQty, int SampleQty, int UomId, string PortOfDelivery, int SalesId
            , int AdditionalFile, int QuotationFile, string Description, string QuotationRemark, string ConfirmVPTime, string ConfirmSalesTime
            , string ConfirmRdTime, string ConfirmCustTime, string Edition, string DocType, string Status
            , string DetailStatus, string PackagingDetails, string SolutionInfos, int BaseCavities, int InsertCavities, string CoreThickness
            , string CommonMode, string DesignChange)
        {
            try
            {
                int rowsAffected = 0;
                int? nullData = null; //null資料
                int RfqDetailIdBase = 1;
                

                bool TestArea = false;
                bool ProTypeChange = false;
                string ChannelNo = "QuoteChannel";
                string Web = "https://bm.zy-tech.com.tw/";
                string RfqProductTypeName = "";
                string MtlFullName = "";
                MtlFullName = MtlName + "/" + PlasticName + "/" + OutsideDiameter + "/";
                #region //判斷資料長度
                if (RfqProTypeId <= 0) throw new SystemException("【產品類別】不能為空!");
                if (KickOffType.Length <= 0) throw new SystemException("【開案型態】不能為空!");
                if (MtlName.Length <= 0) throw new SystemException("【品名】不能為空!");
                if (PlasticName.Length <= 0) throw new SystemException("【塑料品名】不能為空!");
                if (CustProdDigram <= 0) throw new SystemException("【客戶產品圖】不能為空!");
                if (OutsideDiameter.Length <= 0) throw new SystemException("【外徑大小】不能為空!");
                //if (ProtoSchedule.Length <= 0) throw new SystemException("【試作時程】不能為空!");
                //if (PrototypeQty <= 0) throw new SystemException("【試作需求數量】不能為空!");
                if (PlannedOpeningDate.Length <= 0) throw new SystemException("【預計開案日】不能為空!");
                if (DemandDate.Length <= 0) throw new SystemException("【需求日期】不能為空!");
                //if (ProdLifeCycleStart.Length <= 0) throw new SystemException("【生命週期-起】不能為空!");
                //if (ProdLifeCycleEnd.Length <= 0) throw new SystemException("【生命週期-訖】不能為空!");
                //if (LifeCycleQty <= 0) throw new SystemException("【週期數量】不能為空!");
                if (MassProductionDemand.Length <= 0) throw new SystemException("【量產需求】不能為空!");
                if (CoatingFlag.Length <= 0) throw new SystemException("【是否鍍膜】不能為空!");
                if (Currency.Length <= 0) throw new SystemException("【幣別】不能為空!");
                if (PackagingDetails.Length <= 0) throw new SystemException("【包裝種類】不能為空!");
                if (SolutionInfos.Length <= 0) throw new SystemException("【報價依據】不能為空!");
                if (BaseCavities <= 0) throw new SystemException("【模架穴数】不能為空!");
                if (InsertCavities <= 0) throw new SystemException("【模仁穴数】不能為空!");
                if (RfqProTypeId == 2)
                {
                    if (CoreThickness.Length <= 0) throw new SystemException("【芯厚】不能為空!");
                }
                if (CommonMode.Length <= 0) throw new SystemException("【是否共模】不能為空!");
                if (MonthlyQty <= 0) throw new SystemException("【月需求量】不能為空!");
                if (UomId <= 0) throw new SystemException("【單位】不能為空!");
                #endregion

                if (!PackagingDetails.TryParseJson(out JObject tempJObject)) throw new SystemException("包裝種類資料格式錯誤");
                JObject packagingJson = JObject.Parse(PackagingDetails);

                if (!SolutionInfos.TryParseJson(out JObject tempJObject2)) throw new SystemException("報價依據資料格式錯誤");
                JObject solutionJson = JObject.Parse(SolutionInfos);

                List<int> QieIdList = new List<int>();

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷產品類別是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT RfqProductTypeName
                                FROM SCM.RfqProductType
                                WHERE RfqProTypeId = @RfqProTypeId";
                        dynamicParameters.Add("RfqProTypeId", RfqProTypeId);

                        var resultRfqProductType = sqlConnection.Query(sql, dynamicParameters);
                        if (resultRfqProductType.Count() <= 0) throw new SystemException("產品類別找不到或未啟用,請重新確認");
                        foreach (var item in resultRfqProductType)
                        {
                            RfqProductTypeName = item.RfqProductTypeName;
                        }
                        #endregion

                        #region //RFQ單身目前序號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ISNULL(CAST(MAX(RfqSequence) AS INT), 0) + 1 MaxSequence
                                FROM SCM.RfqDetail
                                WHERE RfqId = @RfqId";
                        dynamicParameters.Add("RfqId", RfqId);

                        var resultSequence = sqlConnection.Query(sql, dynamicParameters);

                        int maxSequence = 1;
                        if (resultSequence.Count() > 0)
                        {
                            foreach (var item in resultSequence)
                            {
                                maxSequence = Convert.ToInt32(item.MaxSequence);
                            }
                        }
                        #endregion

                        #region //RFQ設計變更
                        if (RfqSequence != "" && DesignChange == "Y")
                        {
                            #region //撈取當前單身和更新版次
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 RfqDetailId,FORMAT(CAST(Edition AS INT) + 1, '00') AS Edition,RfqProTypeId,CoatingFlag
                                    FROM SCM.RfqDetail
                                    WHERE RfqId = @RfqId
                                    AND RfqSequence = @RfqSequence
                                    ORDER BY Edition DESC";
                            dynamicParameters.Add("RfqId", RfqId);
                            dynamicParameters.Add("RfqSequence", RfqSequence);

                            resultSequence = sqlConnection.Query(sql, dynamicParameters);

                            if (resultSequence.Count() > 0)
                            {
                                int RfqProTypeIdBase = -1;
                                string CoatingFlagBase = "";
                                foreach (var item in resultSequence)
                                {
                                    RfqDetailIdBase = item.RfqDetailId;
                                    Edition = item.Edition;
                                    RfqProTypeIdBase = item.RfqProTypeId;
                                    CoatingFlagBase = item.CoatingFlag;
                                }

                                #region //判斷產品類別是否有改變
                                if (RfqProTypeIdBase != RfqProTypeId)
                                {
                                    ProTypeChange = true;
                                }
                                else
                                {
                                    if(RfqProTypeIdBase == 2 && CoatingFlagBase != CoatingFlag)
                                    {
                                        ProTypeChange = true;
                                    }
                                    else
                                    {
                                        ProTypeChange = false;
                                    }
                                }
                                #endregion
                            }
                            #endregion

                            #region//更新RFQ單身失效
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.RfqDetail SET
                                        DocType = @DocType,
                                        ExpirationDate = @ExpirationDate,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE RfqDetailId = @RfqDetailIdBase
                                        AND DocType != 'F'";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                DocType = "F",
                                ExpirationDate = LastModifiedDate,
                                LastModifiedDate,
                                LastModifiedBy,
                                RfqDetailIdBase
                            });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }
                        #endregion

                        #region //新增RFQ單身資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.RfqDetail (RfqId, CompanyId, RfqProTypeId, RfqSequence, MtlName, CustProdDigram, PlannedOpeningDate, PrototypeQty
                                , ProtoSchedule, MassProductionDemand, KickOffType, PlasticName, OutsideDiameter, ProdLifeCycleStart, ProdLifeCycleEnd, LifeCycleQty
                                , DemandDate, CoatingFlag, Currency, MonthlyQty, SampleQty, UomId, PortOfDelivery, SalesId, AdditionalFile, QuotationFile, Description, QuotationStatus, QuotationRemark
                                , BaseCavities, InsertCavities, CoreThickness, CommonMode, ConfirmVPTime, Edition, DocType, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.RfqDetailId
                                VALUES (@RfqId, @CompanyId, @RfqProTypeId, @RfqSequence, @MtlName, @CustProdDigram, @PlannedOpeningDate, @PrototypeQty
                                , @ProtoSchedule, @MassProductionDemand, @KickOffType, @PlasticName, @OutsideDiameter, @ProdLifeCycleStart, @ProdLifeCycleEnd, @LifeCycleQty
                                , @DemandDate, @CoatingFlag, @Currency, @MonthlyQty, @SampleQty, @UomId, @PortOfDelivery, @SalesId, @AdditionalFile, @QuotationFile, @Description, @QuotationStatus, @QuotationRemark
                                , @BaseCavities, @InsertCavities, @CoreThickness, @CommonMode, @ConfirmVPTime, @Edition, @DocType, @DetailStatus
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RfqId,
                                CompanyId = CurrentCompany,
                                RfqProTypeId,
                                RfqSequence = DesignChange == "N" ? string.Format("{0:0000}", maxSequence) : RfqSequence, //取4位數
                                MtlName,
                                CustProdDigram,
                                PlannedOpeningDate,
                                PrototypeQty,
                                ProtoSchedule,
                                MassProductionDemand,
                                KickOffType,
                                PlasticName,
                                OutsideDiameter,
                                ProdLifeCycleStart,
                                ProdLifeCycleEnd,
                                LifeCycleQty,
                                DemandDate,
                                CoatingFlag,
                                Currency,
                                MonthlyQty,
                                SampleQty = SampleQty > 0 ? SampleQty : 0,
                                UomId,
                                PortOfDelivery,
                                SalesId = CreateBy,
                                AdditionalFile = AdditionalFile == -1 ? (int?)null : AdditionalFile,
                                QuotationFile = QuotationFile == -1 ? (int?)null : QuotationFile,
                                Description,
                                QuotationStatus = "N",
                                QuotationRemark,
                                BaseCavities,
                                InsertCavities,
                                CoreThickness,
                                CommonMode,
                                ConfirmVPTime = CreateDate,
                                Edition = DesignChange == "N" ? "01" : Edition, //取2位數
                                //Edition = string.Format("{0:00}", maxEdition), //取2位數
                                DocType = 'E',
                                DetailStatus = 2,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += insertResult.Count();

                        int RfqDetailId = -1;
                        foreach (var item in insertResult)
                        {
                            RfqDetailId = item.RfqDetailId;
                        }
                        #endregion

                        #region//更新RFQ單頭狀態
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.RequestForQuotation SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RfqId = @RfqId";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            Status = 1,
                            LastModifiedDate,
                            LastModifiedBy,
                            RfqId
                        });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //新增RFQ包裝方式資料
                        for (int i = 0; i < packagingJson["data"].Count(); i++)
                        {
                            #region//新增RFQ包裝方式資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO SCM.RfqPackage (RfqDetailId, RfqPkTypeId, SustSupplyStatus, Status
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.RfqPkId
                                    VALUES (@RfqDetailId, @RfqPkTypeId, @SustSupplyStatus, @Status
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    RfqDetailId,
                                    RfqPkTypeId = Convert.ToInt32(packagingJson["data"][i]["RfqPkTypeId"]),
                                    SustSupplyStatus = packagingJson["data"][i]["SustSupplyStatus"].ToString(),
                                    Status = "A",
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();

                            int RfqPkId = -1;
                            foreach (var item in insertResult)
                            {
                                RfqPkId = item.RfqPkId;
                            }
                            #endregion
                        }
                        #endregion

                        #region //新增RFQ報價依據資料
                        for (int i = 0; i < solutionJson["data"].Count(); i++)
                        {
                            #region//新增RFQ報價依據資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO SCM.RfqLineSolution (RfqDetailId, SortNumber, SolutionQty, PeriodicDemandType
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.RfqLineSolutionId
                                    VALUES (@RfqDetailId, @SortNumber, @SolutionQty, @PeriodicDemandType
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    RfqDetailId,
                                    SortNumber = i + 1,
                                    SolutionQty = Convert.ToInt32(solutionJson["data"][i]["SolutionQty"]),
                                    PeriodicDemandType = "M",
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();

                            int RfqLineSolutionId = -1;
                            foreach (var item in insertResult)
                            {
                                RfqLineSolutionId = item.RfqLineSolutionId;
                            }
                            #endregion
                        }
                        #endregion



                        #region //新增報價單(新)-- Shinotokuro -- 2024-10-19
                        if (RfqSequence != "" && DesignChange == "Y" && ProTypeChange == false)
                        {
                            #region //設計變更採用原報價單
                            List<QuotationItem> QuotationItemData = new List<QuotationItem>();
                            List<QuotationItemElement> QuotationItemElementData = new List<QuotationItemElement>();

                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.*
                                    FROM SCM.QuotationItem a
                                    WHERE a.RfqDetailId = @RfqDetailId
                                    ";
                            dynamicParameters.Add("RfqDetailId", RfqDetailIdBase);

                            var result1 = sqlConnection.Query(sql, dynamicParameters);
                            if (result1.Count() <= 0) throw new SystemException("該產品類別尚未維護報價單項目!!");
                            
                            QuotationItemData = sqlConnection.Query<QuotationItem>(sql, dynamicParameters).ToList();

                            QuotationItemData
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.RfqDetailId = RfqDetailId;
                                    x.ConfirmUserId = nullData;
                                    x.ConfirmStatus = "N";
                                    x.CreateDate = LastModifiedDate;
                                    x.LastModifiedDate = LastModifiedDate;
                                    x.CreateBy = CreateBy;
                                    x.LastModifiedBy = LastModifiedBy;
                                });

                            sql = @"INSERT INTO SCM.QuotationItem (RfqDetailId, HtmlColumns, HtmlColName, Charge, ConfirmUserId
                                    , ConfirmStatus, ColRemark
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@RfqDetailId, @HtmlColumns, @HtmlColName, @Charge, @ConfirmUserId
                                    , @ConfirmStatus, @ColRemark
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            rowsAffected += sqlConnection.Execute(sql, QuotationItemData);

                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.QiId,a.HtmlColumns
                                    FROM SCM.QuotationItem a
                                    WHERE a.RfqDetailId = @RfqDetailId
                                    ";
                            dynamicParameters.Add("RfqDetailId", RfqDetailId);
                            result1 = sqlConnection.Query(sql, dynamicParameters);
                            if (result1.Count() <= 0) throw new SystemException("該產品類別尚未維護報價單項目!!");
                            foreach(var item in result1)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.*
                                        FROM SCM.QuotationItemElement a
                                        INNER JOIN SCM.QuotationItem b on a.QiId = b.QiId
                                        WHERE b.RfqDetailId = @RfqDetailId
                                        AND b.HtmlColumns = @HtmlColumns
                                        ";
                                dynamicParameters.Add("RfqDetailId", RfqDetailIdBase);
                                dynamicParameters.Add("HtmlColumns", item.HtmlColumns);

                                result1 = sqlConnection.Query(sql, dynamicParameters);
                                if (result1.Count() <= 0) throw new SystemException("該產品類別尚未維護報價單項目!!");

                                QuotationItemElementData = sqlConnection.Query<QuotationItemElement>(sql, dynamicParameters).ToList();

                                QuotationItemElementData
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        if (x.Flag == "Material" || x.Flag == "MtAmount" || x.Flag == "MoldingCycle" || x.Flag == "OEE")
                                        {
                                            x.QuoteValue = "";
                                        }
                                        
                                        x.QiId = item.QiId;
                                        x.CreateDate = LastModifiedDate;
                                        x.LastModifiedDate = LastModifiedDate;
                                        x.CreateBy = CreateBy;
                                        x.LastModifiedBy = LastModifiedBy;
                                    });

                                sql = @"INSERT INTO SCM.QuotationItemElement (QiId, Sort, ItemElementNo, ItemElementName, QuoteValue
                                        , ColumnSetting, Formula, DataFrom, ElementRemark, Flag
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        VALUES (@QiId, @Sort, @ItemElementNo, @ItemElementName, @QuoteValue
                                        , @ColumnSetting, @Formula, @DataFrom, @ElementRemark, @Flag
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                rowsAffected += sqlConnection.Execute(sql, QuotationItemElementData);
                            }
                            #endregion
                        }
                        else
                        {
                            #region //全新
                            string QuotationData = "";
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.QuotationData 
                                FROM SCM.QuotationHtmlElement a
                                WHERE a.RfqProTypeId = @RfqProTypeId
                                AND a.CoatingFlag = @CoatingFlag
                                AND a.[Status] = 'A'
                                ";
                            dynamicParameters.Add("CoatingFlag", CoatingFlag);
                            dynamicParameters.Add("RfqProTypeId", RfqProTypeId);

                            var result1 = sqlConnection.Query(sql, dynamicParameters);
                            if (result1.Count() <= 0) throw new SystemException("該產品類別尚未維護報價單項目!!");
                            foreach (var item in result1)
                            {
                                QuotationData = item.QuotationData;
                            }
                            List<QuotationHead> HeadDataJson = JsonConvert.DeserializeObject<List<QuotationHead>>(QuotationData);
                            foreach (var item in HeadDataJson)
                            {
                                string ColNo = item.ColNo;
                                string ColName = item.ColName;
                                string Charge = item.Charge;
                                string ColRemark = item.ColRemark;
                                string UseFlag = item.UseFlag;

                                #region //新增報價單網頁項目
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO SCM.QuotationItem (RfqDetailId, HtmlColumns, HtmlColName, ColRemark, Charge, CreateDate, LastModifiedDate, CreateBy, LastModifiedBy )
	                                        OUTPUT INSERTED.QiId
	                                        VALUES (@RfqDetailId, @HtmlColumns, @HtmlColName, @ColRemark, @Charge, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy )";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        RfqDetailId,
                                        HtmlColumns = ColNo,
                                        HtmlColName = ColName,
                                        ColRemark,
                                        Charge,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                var QuotationitemResult = sqlConnection.Query(sql, dynamicParameters);
                                #endregion

                                int QiId = -1;
                                foreach (var item2 in QuotationitemResult)
                                {
                                    QiId = item2.QiId;
                                }

                                foreach (var detail in item.DetailData)
                                {
                                    string Sort = detail.Sort;
                                    string DetailNo = detail.DetailNo;
                                    string DetailName = detail.DetailName;
                                    string ColumnSetting = detail.ColumnSetting;
                                    string Formula = detail.Formula;
                                    string DataFrom = detail.DataFrom;
                                    string ElementRemark = detail.ElementRemark;
                                    string Flag = detail.Flag;

                                    #region //新增報價單網頁項目元素
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO SCM.QuotationItemElement ( QiId, Sort, ItemElementNo, ItemElementName, QuoteValue, ColumnSetting, Formula, ElementRemark,Flag, CreateDate, LastModifiedDate, CreateBy, LastModifiedBy )
                                            OUTPUT INSERTED.QieId
	                                        VALUES ( @QiId, @Sort, @ItemElementNo, @ItemElementName, @QuoteValue, @ColumnSetting, @Formula, @ElementRemark, @Flag, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy )";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            QiId,
                                            Sort,
                                            ItemElementNo = DetailNo,
                                            ItemElementName = DetailName,
                                            QuoteValue = 0,
                                            ColumnSetting,
                                            Formula,
                                            ElementRemark,
                                            Flag,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });
                                    var QuotationitemelementResult = sqlConnection.Query(sql, dynamicParameters);
                                    #endregion

                                    foreach (var item1 in QuotationitemelementResult)
                                    {
                                        int QieId = item1.QieId;
                                        QieIdList.Add(QieId);
                                    }
                                }
                            }
                            #endregion
                        }
                        #endregion

                        #region //MAMO通知

                        string RfqNo = "";
                        string Content = "";
                        string DetailContent = "";
                        List<string> Tags = new List<string>();

                        #region //確認公司資訊
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CompanyNo,x.UserInfo,a.TestArea
                        FROM BAS.Company a 
                        OUTER APPLY(
                            SELECT x1.UserNo + ' ' + x1.UserName UserInfo FROM BAS.[User] x1 WHERE x1.UserId = @CreateBy
                        ) x
                        WHERE a.CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("CreateBy", CreateBy);

                        var CompanyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (CompanyResult.Count() <= 0) throw new SystemException("公司資料錯誤!!");

                        string CompanyNo = "";
                        string UserInfo = "";
                        foreach (var item in CompanyResult)
                        {
                            CompanyNo = item.CompanyNo;
                            UserInfo = item.UserInfo;
                            if (item.TestArea == "Y")
                            {
                                ChannelNo = "QuoteChannelTest";
                                Web = "http://192.168.20.36:17000/";
                            }
                        }
                        #endregion

                        #region //通知內容
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.HtmlColumns + '. ' +a.HtmlColName Columns,a.Charge,
                            c.RfqNo,c.RfqId
                            FROM SCM.QuotationItem a
                            INNER JOIN SCM.RfqDetail b on a.RfqDetailId = b.RfqDetailId
                            INNER JOIN SCM.RequestForQuotation c on b.RfqId = c.RfqId
                            WHERE a.RfqDetailId = @RfqDetailId";
                        dynamicParameters.Add("RfqDetailId", RfqDetailId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("報價單項目資料找不到,請重新確認!");
                        foreach (var item in result)
                        {
                            string Columns = item.Columns;
                            string Charge = item.Charge;
                            RfqNo = item.RfqNo;
                            string ChargeInfo = "";
                            foreach (var UserNo in Charge.Split(','))
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.UserName, a.UserNo
                                        FROM BAS.[User] a
                                        WHERE a.UserNo = @UserNo OR TRY_CAST(a.UserId AS VARCHAR) = @UserNo";
                                dynamicParameters.Add("UserNo", UserNo);
                                var UserResult = sqlConnection.Query(sql, dynamicParameters);
                                foreach (var item2 in UserResult)
                                {
                                    string UserNoBase = item2.UserNo;
                                    ChargeInfo += "," + item2.UserName;
                                    if (!Tags.Contains(UserNoBase))
                                    {
                                        Tags.Add(UserNoBase);
                                    }
                                }
                            }
                            ChargeInfo = ChargeInfo.Substring(1);

                            DetailContent += "- " + Columns + "項目: " + ChargeInfo + "\n";
                        }
                        #endregion

                        #region //頻道撈取
                        string ChannelId = "";
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ChannelId
                                FROM MAMO.Channels a
                                WHERE a.ChannelNo = @ChannelNo";
                        dynamicParameters.Add("ChannelNo", ChannelNo);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        foreach (var item in result)
                        {
                            ChannelId = item.ChannelId.ToString();
                        }
                        #endregion

                        #region //通知資訊
                        string Url = Web + "RequestForQuotation/RfqAssignManagment?UrlRfqId=" + RfqId + "&UrlRfqDetailId=" + RfqDetailId;
                        string Url2 = Web + "RequestForQuotation/QuotationWorkPlatform";
                        List<int> Files = new List<int>();

                        Content = "### 報價單-【成立通知】\n" +
                                    "##### 產品類別: " + RfqProductTypeName + "\n" +
                                    "##### 品名/塑料品名/外徑大小: " + MtlFullName + "\n" +
                                    DetailContent +
                                    "##### 說明: 請各項目負責人員,協助維護該報價項目資訊\n" +
                                    "##### 傳送連結: " + Url + "\n" +
                                    "##### 工作平台: " + Url2 + "\n" +
                                    "- 發信時間: " + CreateDate + "\n" +
                                    "- 發信人員: " + UserInfo + "\n";
                        #endregion

                        #region //執行
                        string MamoResult = mamoHelper.SendMessage(CompanyNo, CurrentUser, "Channel", ChannelId, Content, Tags, Files);

                        JObject MamoResultJson = JObject.Parse(MamoResult);
                        if (MamoResultJson["status"].ToString() != "success")
                        {
                            throw new SystemException(MamoResultJson["msg"].ToString());
                        }
                        #endregion

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

        #region //AddRfqDetailDocType-- 設變流程(複製RFQ資料功能) -- Yi 2023-11-14
        public string AddRfqDetailDocType(int RfqId, int RfqDetailId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        string currentDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                        #region //判斷RFQ資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.RfqId, a.CompanyId, a.RfqProTypeId, a.RfqSequence, a.MtlName, a.CustProdDigram
                                , ISNULL(FORMAT(a.PlannedOpeningDate, 'yyyy-MM-dd'), '') PlannedOpeningDate, a.PrototypeQty, a.ProtoSchedule
                                , a.MassProductionDemand, a.KickOffType, a.PlasticName, a.OutsideDiameter
                                , FORMAT(a.ProdLifeCycleStart, 'yyyy-MM-dd') ProdLifeCycleStart, FORMAT(a.ProdLifeCycleEnd, 'yyyy-MM-dd') ProdLifeCycleEnd
                                , a.LifeCycleQty, FORMAT(a.DemandDate, 'yyyy-MM-dd') DemandDate, a.CoatingFlag, a.Currency, a.PortOfDelivery
                                , ISNULL(a.SalesId, '') SalesId, ISNULL(a.AdditionalFile, '') AdditionalFile, ISNULL(a.QuotationFile, '') QuotationFile, a.Description, a.QuotationRemark
                                , FORMAT(a.ConfirmVPTime, 'yyyy-MM-dd hh:mm:ss') ConfirmVPTime, FORMAT(a.ConfirmSalesTime, 'yyyy-MM-dd hh:mm:ss') ConfirmSalesTime
                                , FORMAT(a.ConfirmRdTime, 'yyyy-MM-dd hh:mm:ss') ConfirmRdTime, FORMAT(a.ConfirmCustTime, 'yyyy-MM-dd hh:mm:ss') ConfirmCustTime
                                , a.Edition, a.DocType, FORMAT(a.ExpirationDate, 'yyyy-MM-dd hh:mm:ss') ExpirationDate, a.Status
                                FROM SCM.RfqDetail a
                                WHERE a.RfqDetailId = @RfqDetailId";
                        dynamicParameters.Add("RfqDetailId", RfqDetailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("RFQ單身資料錯誤!");

                        int rfqId = -1, companyId = -1, rfqProTypeId = -1, custProdDigram = -1, prototypeQty = -1, outsideDiameter = -1, lifeCycleQty = -1
                            , salesId = -1, additionalFile = -1, quotationFile = -1;

                        string rfqSequence = "", mtlName = "", plannedOpeningDate = "", protoSchedule = ""
                            , massProductionDemand = "", kickOffType = "", plasticName = "", prodLifeCycleStart = "", prodLifeCycleEnd = ""
                            , demandDate = "", coatingFlag = "", currency = "", portOfDelivery = "", description = "", quotationRemark = ""
                            , confirmVPTime = "", confirmSalesTime = "", confirmRdTime = "", confirmCustTime = "", docType = "", edition = ""
                            , expirationDate = "", status = "";

                        foreach (var item in result)
                        {
                            rfqId = item.RfqId;
                            companyId = item.CompanyId;
                            rfqProTypeId = item.RfqProTypeId;
                            rfqSequence = item.RfqSequence;
                            mtlName = item.MtlName;
                            custProdDigram = item.CustProdDigram;
                            plannedOpeningDate = item.PlannedOpeningDate;
                            prototypeQty = item.PrototypeQty;
                            protoSchedule = item.ProtoSchedule;
                            massProductionDemand = item.MassProductionDemand;
                            kickOffType = item.KickOffType;
                            plasticName = item.PlasticName;
                            outsideDiameter = item.OutsideDiameter;
                            prodLifeCycleStart = item.ProdLifeCycleStart;
                            prodLifeCycleEnd = item.ProdLifeCycleEnd;
                            lifeCycleQty = item.LifeCycleQty;
                            demandDate = item.DemandDate;
                            coatingFlag = item.CoatingFlag;
                            currency = item.Currency;
                            portOfDelivery = item.PortOfDelivery;
                            salesId = item.SalesId;
                            additionalFile = item.AdditionalFile;
                            quotationFile = item.QuotationFile;
                            description = item.Description;
                            quotationRemark = item.QuotationRemark;
                            confirmVPTime = item.ConfirmVPTime;
                            confirmSalesTime = item.ConfirmSalesTime;
                            confirmRdTime = item.ConfirmRdTime;
                            confirmCustTime = item.ConfirmCustTime;
                            edition = item.Edition;
                            docType = item.DocType;
                            expirationDate = item.ExpirationDate;
                            status = item.Status;
                        }
                        #endregion

                        #region //判斷單身序號是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.RfqDetail
                                WHERE RfqId = @RfqId
                                AND RfqSequence = @RfqSequence
                                AND RfqDetailId != @RfqDetailId";
                        dynamicParameters.Add("RfqId", rfqId);
                        dynamicParameters.Add("RfqSequence", rfqSequence);
                        dynamicParameters.Add("RfqDetailId", RfqDetailId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("流水號重複，請重新取號!");
                        #endregion

                        #region //RFQ單身目前序號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ISNULL(CAST(MAX(RfqSequence) AS INT), 0) + 1 MaxSequence
                                FROM SCM.RfqDetail
                                WHERE RfqId = @RfqId";
                        dynamicParameters.Add("RfqId", RfqId);

                        var resultSequence = sqlConnection.Query(sql, dynamicParameters);

                        int maxSequence = 1;
                        if (resultSequence.Count() > 0)
                        {
                            foreach (var item in resultSequence)
                            {
                                maxSequence = Convert.ToInt32(item.MaxSequence);
                            }
                        }
                        #endregion

                        #region //RFQ單身目前版次號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ISNULL(CAST(MAX(Edition) AS INT), 0) + 1 MaxEdition
                                FROM SCM.RfqDetail
                                WHERE RfqId = @RfqId";
                        dynamicParameters.Add("RfqId", RfqId);

                        var resultEdition = sqlConnection.Query(sql, dynamicParameters);

                        int maxEdition = 1;
                        if (resultSequence.Count() > 0)
                        {
                            foreach (var item in resultEdition)
                            {
                                maxEdition = Convert.ToInt32(item.MaxEdition);
                            }
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.RfqDetail (RfqId, CompanyId, RfqProTypeId, RfqSequence, MtlName, CustProdDigram, PlannedOpeningDate
                                , PrototypeQty, ProtoSchedule, MassProductionDemand, KickOffType, PlasticName, OutsideDiameter, ProdLifeCycleStart
                                , ProdLifeCycleEnd, LifeCycleQty, DemandDate, CoatingFlag, Currency, PortOfDelivery, SalesId, AdditionalFile
                                , Description, QuotationRemark, ConfirmVPTime, ConfirmSalesTime, ConfirmRdTime, ConfirmCustTime
                                , Edition, DocType, ExpirationDate, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.RfqDetailId
                                VALUES (@RfqId, @CompanyId, @RfqProTypeId, @RfqSequence, @MtlName, @CustProdDigram, @PlannedOpeningDate
                                , @PrototypeQty, @ProtoSchedule, @MassProductionDemand, @KickOffType, @PlasticName, @OutsideDiameter, @ProdLifeCycleStart
                                , @ProdLifeCycleEnd, @LifeCycleQty, @DemandDate, @CoatingFlag, @Currency, @PortOfDelivery, @SalesId, @AdditionalFile
                                , @Description, @QuotationRemark, @ConfirmVPTime, @ConfirmSalesTime, @ConfirmRdTime, @ConfirmCustTime
                                , @Edition, @DocType, @ExpirationDate, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RfqId,
                                CompanyId = companyId,
                                RfqProTypeId = rfqProTypeId,
                                RfqSequence = string.Format("{0:0000}", maxSequence), //取4位數,
                                MtlName = mtlName,
                                CustProdDigram = custProdDigram,
                                PlannedOpeningDate = plannedOpeningDate,
                                PrototypeQty = prototypeQty,
                                ProtoSchedule = protoSchedule,
                                MassProductionDemand = massProductionDemand,
                                KickOffType = kickOffType,
                                PlasticName = plasticName,
                                OutsideDiameter = outsideDiameter,
                                ProdLifeCycleStart = prodLifeCycleStart,
                                ProdLifeCycleEnd = prodLifeCycleEnd,
                                LifeCycleQty = lifeCycleQty,
                                DemandDate = demandDate,
                                CoatingFlag = coatingFlag,
                                Currency = currency,
                                PortOfDelivery = portOfDelivery,
                                SalesId = salesId,
                                AdditionalFile = additionalFile <= 0 ? (int?)null : additionalFile,
                                //QuotationFile = quotationFile,
                                Description = description,
                                QuotationRemark = quotationRemark,
                                ConfirmVPTime = confirmVPTime,
                                ConfirmSalesTime = confirmSalesTime,
                                ConfirmRdTime = confirmRdTime,
                                ConfirmCustTime = confirmCustTime,
                                Edition = string.Format("{0:00}", maxEdition), //取2位數,
                                DocType = docType,
                                ExpirationDate = expirationDate,
                                Status = 2,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });

                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = insertResult.Count();

                        int newRfqDetailId = 0;
                        foreach (var item in insertResult)
                        {
                            newRfqDetailId = item.RfqDetailId;
                        }

                        #region //複製報價方案數量
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.RfqDetailId, a.SortNumber, a.SolutionQty, a.PeriodicDemandType
                                FROM SCM.RfqLineSolution a
                                WHERE RfqDetailId = @RfqDetailId";
                        dynamicParameters.Add("RfqDetailId", RfqDetailId);

                        var resultRs = sqlConnection.Query(sql, dynamicParameters);
                        if (resultRs.Count() > 0)
                        {
                            int sortNumber = -1, solutionQty = -1;
                            string periodicDemandType = "";
                            foreach (var item in resultRs)
                            {
                                sortNumber = item.SortNumber;
                                solutionQty = item.SolutionQty;
                                periodicDemandType = item.PeriodicDemandType;

                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO SCM.RfqLineSolution (RfqDetailId, SortNumber, SolutionQty, PeriodicDemandType
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.RfqLineSolutionId
                                        VALUES (@RfqDetailId, @SortNumber, @SolutionQty, @PeriodicDemandType
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        RfqDetailId = newRfqDetailId,
                                        SortNumber = sortNumber,
                                        SolutionQty = solutionQty,
                                        PeriodicDemandType = periodicDemandType,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });

                                insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();
                            }
                        }
                        #endregion

                        #region //複製包裝方式
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.RfqDetailId, a.RfqPkTypeId, a.SustSupplyStatus, a.Status
                                FROM SCM.RfqPackage a
                                WHERE RfqDetailId = @RfqDetailId";
                        dynamicParameters.Add("RfqDetailId", RfqDetailId);

                        var resultRp = sqlConnection.Query(sql, dynamicParameters);
                        if (resultRp.Count() > 0)
                        {
                            int rfqPkTypeId = -1;
                            string sustSupplyStatus = "", rpStatus = "";
                            foreach (var item in resultRp)
                            {
                                rfqPkTypeId = item.RfqPkTypeId;
                                sustSupplyStatus = item.SustSupplyStatus;
                                rpStatus = item.Status;

                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO SCM.RfqPackage (RfqDetailId, RfqPkTypeId, SustSupplyStatus, Status
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.RfqPkId
                                        VALUES (@RfqDetailId, @RfqPkTypeId, @SustSupplyStatus, @Status
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        RfqDetailId = newRfqDetailId,
                                        RfqPkTypeId = rfqPkTypeId,
                                        SustSupplyStatus = sustSupplyStatus,
                                        Status = rpStatus,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });

                                insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();
                            }
                        }
                        #endregion

                        #region //原Rfq單身單據狀態更改為失效
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.RfqDetail SET
                                DocType = @DocType,
                                ExpirationDate = @ExpirationDate,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RfqDetailId = @RfqDetailId
                                AND RfqSequence = @RfqSequence";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            DocType = 'F',
                            ExpirationDate = currentDate,
                            LastModifiedDate,
                            LastModifiedBy,
                            RfqDetailId,
                            RfqSequence = rfqSequence,
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

        #region //AddQuotationBargain-- 新增報價議價 -- Shinotokuro 2023-08-02
        public string AddQuotationBargain(int RfqDetailId, double FinalPrice, double GrossMargin)
        {
            try
            {
                if (RfqDetailId <= 0) throw new SystemException("【報價單ID】不能為空!");
                if (FinalPrice <= 0) throw new SystemException("【新價格】不能為負!");
                if (GrossMargin <= 0) throw new SystemException("【毛利率】不能為負!");

                int rowsAffected = 0;
                int RfqId = -1;
                int QfpId = -1;

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷報價單是否正確

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.RfqId,b.RfqDetailId,b.QuotationStatus,ISNULL(x.QfpId,-1) QfpId
                                FROM SCM.RequestForQuotation a
                                INNER JOIN SCM.RfqDetail b on a.RfqId = b.RfqId
                                OUTER APPLY(
                                    SELECT TOP 1 QfpId  
                                    FROM SCM.QuotationFinalPrice x1
                                    WHERE x1.RfqDetailId = b.RfqDetailId
                                    AND x1.ConfirmStatus != 'H'
                                ) x
                                WHERE b.RfqDetailId = @RfqDetailId
                                AND b.DocType = 'E'
                                ORDER BY b.RfqDetailId DESC
                                ";
                        dynamicParameters.Add("RfqId", RfqId);
                        dynamicParameters.Add("RfqDetailId", RfqDetailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("報價單資料找不到,請重新確認!");
                        foreach (var item in result)
                        {
                            if (item.QuotationStatus != "F") throw new SystemException("報價單須為最終價格狀態,才可以執行議價動作!!");
                            RfqDetailId = item.RfqDetailId;
                            QfpId = item.QfpId;
                            RfqId = item.RfqId;
                        }
                        #endregion
                        if (QfpId <= -1)
                        {
                            #region //新增議價紀錄資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO SCM.QuotationFinalPrice (RfqDetailId, FinalPrice, GrossMargin, ConfirmStatus, Remark
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.QfpId
                                VALUES (@RfqDetailId, @FinalPrice, @GrossMargin, @ConfirmStatus, @Remark
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    RfqDetailId,
                                    FinalPrice,
                                    GrossMargin, 
                                    ConfirmStatus = "N",
                                    Remark = "",
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();
                            #endregion
                        }
                        else
                        {
                            #region//更新議價紀錄資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.QuotationFinalPrice SET
                                    FinalPrice = @FinalPrice,
                                    GrossMargin = @GrossMargin,
                                    ConfirmStatus = @ConfirmStatus,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE QfpId = @QfpId";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                FinalPrice,
                                GrossMargin,
                                ConfirmStatus = "N",
                                LastModifiedDate,
                                LastModifiedBy,
                                QfpId
                            });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }




                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            data = rowsAffected
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
        #region //UpdateRequestForQuotation -- RFQ單頭資料更新 -- Yi 2023.07.11
        public string UpdateRequestForQuotation(int RfqId, string MemberName, string AssemblyName, int ProductUseId
            , string OrganizaitonType, int CustomerId, string CustomerName, int SupplierId, string SupplierName
            , string ContactPerson, string ContactPhone)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷RFQ資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.RequestForQuotation
                                WHERE RfqId = @RfqId";
                        dynamicParameters.Add("RfqId", RfqId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("RFQ表單資料錯誤!");
                        #endregion

                        #region//更新RFQ明細
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.RequestForQuotation SET
                                AssemblyName = @AssemblyName,
                                ProductUseId = @ProductUseId,
                                OrganizaitonType = @OrganizaitonType,
                                CustomerId = @CustomerId,
                                CustomerName = @CustomerName,
                                SupplierId = @SupplierId,
                                SupplierName = @SupplierName,
                                ContactPerson = @ContactPerson,
                                ContactPhone = @ContactPhone,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RfqId = @RfqId";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            AssemblyName,
                            ProductUseId,
                            OrganizaitonType,
                            CustomerId = CustomerId <= 0 ? (int?)null : CustomerId,
                            CustomerName,
                            SupplierId = SupplierId <= 0 ? (int?)null : SupplierId,
                            SupplierName,
                            ContactPerson,
                            ContactPhone,
                            LastModifiedDate,
                            LastModifiedBy,
                            RfqId
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

        #region //UpdateRfqSales -- 主頁批次指派業務更新 -- Yi 2023.07.10
        public string UpdateRfqSales(int RfqId, int CompanyId, int UserId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    string currentDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    string rfqNo = "", productUseName = "", rfqProductTypeName = "", assemblyName = "", memberName = "";

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷RFQ單頭是否正確
                        sql = @"SELECT TOP 1 RfqId
                                FROM SCM.RequestForQuotation
                                WHERE RfqId = @RfqId";
                        dynamicParameters.Add("RfqId", RfqId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("RFQ單頭明細資料錯誤!");

                        int rfqId = -1;
                        foreach (var item in result)
                        {
                            rfqId = item.RfqId;
                        }
                        #endregion

                        #region //判斷RFQ資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT RfqDetailId
                                FROM SCM.RequestForQuotation a
                                INNER JOIN SCM.RfqDetail b ON b.RfqId = a.RfqId
                                WHERE a.RfqId = @RfqId";
                        dynamicParameters.Add("RfqId", rfqId);

                        var resultDetail = sqlConnection.Query(sql, dynamicParameters);
                        if (resultDetail.Count() <= 0) throw new SystemException("RFQ資料錯誤!");

                        int rfqDetailId = -1;
                        foreach (var item in resultDetail)
                        {
                            rfqDetailId = item.RfqDetailId;
                        }
                        #endregion

                        #region//更新RFQ單身資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.RfqDetail SET
                                CompanyId = @CompanyId,
                                SalesId = @SalesId,
                                ConfirmSalesTime = @ConfirmSalesTime,
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RfqId = @RfqId";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            CompanyId,
                            SalesId = UserId,
                            ConfirmSalesTime = currentDate,
                            Status = 2,
                            LastModifiedDate,
                            LastModifiedBy,
                            RfqId = rfqId
                        });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //發送Mail
                        #region //Mail資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MailSubject, a.MailContent
                                , b.Host, b.Port, b.SendMode, b.Account, b.Password
                                , ISNULL(c.MailFrom, '') MailFrom, ISNULL(d.MailTo, '') MailTo, ISNULL(e.MailCc, '') MailCc, ISNULL(f.MailBcc, '') MailBcc
                                FROM BAS.Mail a
                                LEFT JOIN BAS.MailServer b ON b.ServerId = a.ServerId
                                OUTER APPLY (
                                    SELECT STUFF((
                                        SELECT ';' + CASE WHEN ca.ContactId IS NULL THEN 
                                            CASE cc.JobType 
                                                WHEN '管理制' THEN cd.DepartmentName + cc.Job + '-' + cc.UserName + ':' + cc.Email
                                                ELSE cd.DepartmentName + '-' + cc.UserName + ':' + cc.Email
                                            END
                                        ELSE cb.ContactName + ':' + cb.Email END
                                        FROM BAS.MailUser ca
                                        LEFT JOIN BAS.MailContact cb ON ca.ContactId = cb.ContactId
                                        LEFT JOIN BAS.[User] cc ON ca.UserId = cc.UserId AND cc.[Status] = 'A'
                                        LEFT JOIN BAS.Department cd ON cc.DepartmentId = cd.DepartmentId
                                        WHERE ca.MailId = a.MailId
                                        AND ca.MailUserType = 'F'
                                        FOR XML PATH('')
                                    ), 1, 1, '') MailFrom
                                ) c
                                OUTER APPLY (
                                    SELECT STUFF((
                                        SELECT ';' + CASE WHEN da.ContactId IS NULL THEN 
                                            CASE dc.JobType 
                                                WHEN '管理制' THEN dd.DepartmentName + dc.Job + '-' + dc.UserName + ':' + dc.Email
                                                ELSE dd.DepartmentName + '-' + dc.UserName + ':' + dc.Email
                                            END
                                        ELSE db.ContactName + ':' + db.Email END
                                        FROM BAS.MailUser da
                                        LEFT JOIN BAS.MailContact db ON da.ContactId = db.ContactId
                                        LEFT JOIN BAS.[User] dc ON da.UserId = dc.UserId AND dc.[Status] = 'A'
                                        LEFT JOIN BAS.Department dd ON dc.DepartmentId = dd.DepartmentId
                                        WHERE da.MailId = a.MailId
                                        AND da.MailUserType = 'T'
                                        FOR XML PATH('')
                                    ), 1, 1, '') MailTo
                                ) d
                                OUTER APPLY (
                                    SELECT STUFF((
                                        SELECT ';' + CASE WHEN ea.ContactId IS NULL THEN 
                                            CASE ec.JobType 
                                                WHEN '管理制' THEN ed.DepartmentName + ec.Job + '-' + ec.UserName + ':' + ec.Email
                                                ELSE ed.DepartmentName + '-' + ec.UserName + ':' + ec.Email
                                            END
                                        ELSE eb.ContactName + ':' + eb.Email END
                                        FROM BAS.MailUser ea
                                        LEFT JOIN BAS.MailContact eb ON ea.ContactId = eb.ContactId
                                        LEFT JOIN BAS.[User] ec ON ea.UserId = ec.UserId AND ec.[Status] = 'A'
                                        LEFT JOIN BAS.Department ed ON ec.DepartmentId = ed.DepartmentId
                                        WHERE ea.MailId = a.MailId
                                        AND ea.MailUserType = 'C'
                                        FOR XML PATH('')
                                    ), 1, 1, '') MailCc
                                ) e
                                OUTER APPLY (
                                    SELECT STUFF((
                                        SELECT ';' + CASE WHEN fa.ContactId IS NULL THEN 
                                            CASE fc.JobType 
                                                WHEN '管理制' THEN fd.DepartmentName + fc.Job + '-' + fc.UserName + ':' + fc.Email
                                                ELSE fd.DepartmentName + '-' + fc.UserName + ':' + fc.Email
                                            END
                                        ELSE fb.ContactName + ':' + fb.Email END
                                        FROM BAS.MailUser fa
                                        LEFT JOIN BAS.MailContact fb ON fa.ContactId = fb.ContactId
                                        LEFT JOIN BAS.[User] fc ON fa.UserId = fc.UserId AND fc.[Status] = 'A'
                                        LEFT JOIN BAS.Department fd ON fc.DepartmentId = fd.DepartmentId
                                        WHERE fa.MailId = a.MailId
                                        AND fa.MailUserType = 'B'
                                        FOR XML PATH('')
                                    ), 1, 1, '') MailBcc
                                ) f
                                WHERE a.MailId IN (
                                    SELECT z.MailId
                                    FROM BAS.MailSendSetting z
                                    WHERE z.SettingSchema = @SettingSchema
                                    AND z.SettingNo = @SettingNo
                                )";
                        dynamicParameters.Add("SettingSchema", "RfqAssignSales");
                        dynamicParameters.Add("SettingNo", "Y");

                        var resultMailTemplate = sqlConnection.Query(sql, dynamicParameters);
                        if (resultMailTemplate.Count() <= 0) throw new SystemException("Mail設定錯誤!");
                        #endregion

                        #region //RFQ資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.RfqId, a.RfqNo, a.AssemblyName
                                , a.ProductUseId, c.ProductUseName, b.RfqProTypeId, e.RfqProductTypeName
                                , ISNULL(d.MemberId, f.UserId) MemberUserId, ISNULL(d.MemberName, f.UserName) MemberUserName, a.Status
                                FROM SCM.RequestForQuotation a
                                INNER JOIN SCM.RfqDetail b ON b.RfqId = a.RfqId
                                INNER JOIN SCM.ProductUse c ON c.ProductUseId = a.ProductUseId
                                LEFT JOIN EIP.Member d ON d.MemberId = a.MemberId
                                INNER JOIN SCM.RfqProductType e ON e.RfqProTypeId = b.RfqProTypeId
                                LEFT JOIN BAS.[User] f ON f.UserId = a.UserId
                                WHERE a.RfqId = @RfqId";
                        dynamicParameters.Add("RfqId", RfqId);

                        var resultRfqInfo = sqlConnection.Query(sql, dynamicParameters);
                        if (resultRfqInfo.Count() <= 0) throw new SystemException("RFQ資料錯誤!");

                        foreach (var item in resultRfqInfo)
                        {
                            rfqNo = item.RfqNo;
                            productUseName = item.ProductUseName;
                            rfqProductTypeName = item.RfqProductTypeName;
                            assemblyName = item.AssemblyName;
                            memberName = item.MemberUserName;
                        }
                        #endregion

                        #region //查找UserInfo
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserId, a.UserName, a.Email
                               FROM BAS.[User] a
                               WHERE a.UserId = @UserId";
                        dynamicParameters.Add("UserId", UserId);

                        var resultUserInfo = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUserInfo.Count() <= 0) throw new SystemException("收件者資料錯誤!");

                        string salesName = "";
                        string email = "";
                        foreach (var item in resultUserInfo)
                        {
                            UserId = item.UserId;
                            salesName = item.UserName;
                            email = item.Email;
                        }
                        #endregion

                        foreach (var item in resultMailTemplate)
                        {
                            string mailSubject = item.MailSubject,
                            mailContent = HttpUtility.UrlDecode(item.MailContent);

                            string hyperlink = "http://" + HttpContext.Current.Request.Url.Authority + "/RequestForQuotation/RfqAssignManagment";
                            rfqNo = string.Format("<a href=\"{0}\">{1}</a>", hyperlink, rfqNo);

                            #region //Mail內容
                            mailContent = mailContent.Replace("[RfqNo]", rfqNo);
                            mailContent = mailContent.Replace("[ProductUseName]", productUseName);
                            mailContent = mailContent.Replace("[RfqProductTypeName]", rfqProductTypeName);
                            mailContent = mailContent.Replace("[AssemblyName]", assemblyName);
                            mailContent = mailContent.Replace("[MemberName]", memberName);
                            mailContent = mailContent.Replace("[SalesName]", salesName);
                            #endregion

                            #region //寄送Mail
                            MailConfig mailConfig = new MailConfig
                            {
                                Host = item.Host,
                                Port = Convert.ToInt32(item.Port),
                                SendMode = Convert.ToInt32(item.SendMode),
                                From = item.MailFrom,
                                Subject = mailSubject,
                                Account = item.Account,
                                Password = item.Password,
                                MailTo = item.MailTo,
                                MailCc = item.MailCc,
                                MailBcc = item.MailBcc,
                                HtmlBody = mailContent,
                                TextBody = "-"
                            };

                            BaseHelper.MailSend(mailConfig);
                            #endregion
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
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //UpdateRfqDetail -- RFQ單身明細更新 -- Yi 2023.07.11
        public string UpdateRfqDetail(int RfqId, int RfqProTypeId, int RfqDetailId
            , string MtlName, string PlannedOpeningDate, int PrototypeQty, string ProtoSchedule, string MassProductionDemand
            , string KickOffType, string PlasticName, string OutsideDiameter, string ProdLifeCycleStart, string ProdLifeCycleEnd
            , int LifeCycleQty, string DemandDate, string CoatingFlag, string Currency, int MonthlyQty,int SampleQty, int UomId, int CustProdDigram, int AdditionalFile
            , int QuotationFile, string Description, string PackagingDetails, string SolutionInfos, int BaseCavities, int InsertCavities, string CoreThickness, string CommonMode)
        {
            try
            {
                #region //判斷單身資料長度
                //if (CompanyId <= 0) throw new SystemException("【公司】不能為空!");
                //if (RfqProTypeId <= 0) throw new SystemException("【產品類別】不能為空!");
                if (MtlName.Length <= 0) throw new SystemException("【品名】不能為空!");
                if (PlannedOpeningDate.Length <= 0) throw new SystemException("【預計開案日】不能為空!");
                //if (PrototypeQty <= 0) throw new SystemException("【試作需求數量】不能為空!");
                //if (ProtoSchedule.Length <= 0) throw new SystemException("【試作時程】不能為空!");
                if (MassProductionDemand.Length <= 0) throw new SystemException("【量產需求】不能為空!");
                if (KickOffType.Length <= 0) throw new SystemException("【開案型態】不能為空!");
                if (PlasticName.Length <= 0) throw new SystemException("【塑料品名】不能為空!");
                if (OutsideDiameter.Length <= 0) throw new SystemException("【部品外徑大小】不能為空!");
                //if (ProdLifeCycleStart.Length <= 0) throw new SystemException("【生命週期(起)】不能為空!");
                //if (ProdLifeCycleEnd.Length <= 0) throw new SystemException("【生命週期(迄)】不能為空!");
                //if (LifeCycleQty <= 0) throw new SystemException("【週期數量】不能為空!");
                if (DemandDate.Length <= 0) throw new SystemException("【需求日期】不能為空!");
                if (CoatingFlag.Length <= 0) throw new SystemException("【是否鍍膜】不能為空!");
                if (Currency.Length <= 0) throw new SystemException("【幣別】不能為空!");
                if (CustProdDigram <= 0) throw new SystemException("【客戶產品圖】不能為空!");
                if (PackagingDetails.Length <= 0) throw new SystemException("【包裝種類】不能為空!");
                if (SolutionInfos.Length <= 0) throw new SystemException("【報價依據】不能為空!");
                if (BaseCavities <= 0) throw new SystemException("【模架穴数】不能為空!");
                if (InsertCavities <= 0) throw new SystemException("【模仁穴数】不能為空!");
                if (RfqProTypeId == 2)
                {
                    if (CoreThickness.Length <= 0) throw new SystemException("【芯厚】不能為空!");
                }
                if (CommonMode.Length <= 0) throw new SystemException("【是否共模】不能為空!");
                #endregion

                if (!PackagingDetails.TryParseJson(out JObject tempJObject)) throw new SystemException("包裝種類資料格式錯誤");
                JObject packagingJson = JObject.Parse(PackagingDetails);

                if (!SolutionInfos.TryParseJson(out JObject tempJObject2)) throw new SystemException("報價依據資料格式錯誤");
                JObject solutionJson = JObject.Parse(SolutionInfos);

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷RFQ單頭是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.RequestForQuotation
                                WHERE RfqId = @RfqId";
                        dynamicParameters.Add("RfqId", RfqId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("RFQ單頭明細資料錯誤!");
                        #endregion

                        #region //判斷RFQ明細是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT b.RfqDetailId, b.SalesId
                                FROM SCM.RequestForQuotation a
                                INNER JOIN SCM.RfqDetail b ON b.RfqId = a.RfqId
                                WHERE b.RfqDetailId = @RfqDetailId";
                        dynamicParameters.Add("RfqDetailId", RfqDetailId);

                        var resultDetail = sqlConnection.Query(sql, dynamicParameters);
                        if (resultDetail.Count() <= 0) throw new SystemException("RFQ單身明細資料錯誤!");

                        int rfqDetailId = -1;
                        int salesId = -1;
                        foreach (var item in resultDetail)
                        {
                            rfqDetailId = item.RfqDetailId;
                            salesId = item.SalesId;

                            #region //RFQ包裝方式資料刪除
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE SCM.RfqPackage
                                    WHERE RfqDetailId = @RfqDetailId";
                            dynamicParameters.Add("RfqDetailId", rfqDetailId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }
                        #endregion

                        #region//更新RFQ單身明細
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.RfqDetail SET
                                RfqId = @RfqId,
                                CompanyId = @CompanyId,
                                RfqProTypeId = @RfqProTypeId,
                                MtlName = @MtlName,
                                PlannedOpeningDate = @PlannedOpeningDate,
                                PrototypeQty = @PrototypeQty,
                                ProtoSchedule = @ProtoSchedule,
                                MassProductionDemand = @MassProductionDemand,
                                KickOffType = @KickOffType,
                                PlasticName = @PlasticName,
                                OutsideDiameter = @OutsideDiameter,
                                ProdLifeCycleStart = @ProdLifeCycleStart,
                                ProdLifeCycleEnd = @ProdLifeCycleEnd,
                                LifeCycleQty = @LifeCycleQty,
                                DemandDate = @DemandDate,
                                CoatingFlag = @CoatingFlag,
                                Currency = @Currency,
                                MonthlyQty = @MonthlyQty,
                                SampleQty = @SampleQty,
                                UomId = @UomId,
                                SalesId = @SalesId,
                                CustProdDigram = @CustProdDigram,
                                AdditionalFile = @AdditionalFile,
                                QuotationFile = @QuotationFile,
                                Description = @Description,
                                BaseCavities = @BaseCavities,
                                InsertCavities = @InsertCavities,
                                CoreThickness = @CoreThickness,
                                CommonMode = @CommonMode,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RfqDetailId = @RfqDetailId";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            RfqId,
                            CompanyId = CurrentCompany,
                            RfqProTypeId,
                            MtlName,
                            PlannedOpeningDate,
                            PrototypeQty,
                            ProtoSchedule,
                            MassProductionDemand,
                            KickOffType,
                            PlasticName,
                            OutsideDiameter,
                            ProdLifeCycleStart,
                            ProdLifeCycleEnd,
                            LifeCycleQty,
                            DemandDate,
                            CoatingFlag,
                            Currency,
                            MonthlyQty,
                            SampleQty = SampleQty > 0 ? SampleQty : 0,
                            UomId,
                            SalesId = salesId,
                            CustProdDigram = CustProdDigram == -1 ? (int?)null : CustProdDigram,
                            AdditionalFile = AdditionalFile == -1 ? (int?)null : AdditionalFile,
                            QuotationFile = QuotationFile == -1 ? (int?)null : QuotationFile,
                            Description,
                            BaseCavities,
                            InsertCavities,
                            CoreThickness,
                            CommonMode,
                            LastModifiedDate,
                            LastModifiedBy,
                            RfqDetailId = rfqDetailId
                        });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //新增RFQ包裝方式資料
                        for (int i = 0; i < packagingJson["data"].Count(); i++)
                        {
                            #region//新增RFQ包裝方式資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO SCM.RfqPackage (RfqDetailId, RfqPkTypeId, SustSupplyStatus, Status
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.RfqPkId
                                    VALUES (@RfqDetailId, @RfqPkTypeId, @SustSupplyStatus, @Status
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    RfqDetailId = rfqDetailId,
                                    RfqPkTypeId = Convert.ToInt32(packagingJson["data"][i]["RfqPkTypeId"]),
                                    SustSupplyStatus = packagingJson["data"][i]["SustSupplyStatus"].ToString(),
                                    Status = "A",
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();

                            int RfqPkId = -1;
                            foreach (var item in insertResult)
                            {
                                RfqPkId = item.RfqPkId;
                            }
                            #endregion
                        }
                        #endregion

                        #region //報價方案資料刪除
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.RfqLineSolution
                                WHERE RfqDetailId = @RfqDetailId";
                        dynamicParameters.Add("RfqDetailId", RfqDetailId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //新增RFQ報價方案資料
                        for (int i = 0; i < solutionJson["data"].Count(); i++)
                        {
                            #region//新增RFQ報價方案資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO SCM.RfqLineSolution (RfqDetailId, SortNumber, SolutionQty, PeriodicDemandType
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.RfqLineSolutionId
                                    VALUES (@RfqDetailId, @SortNumber, @SolutionQty, @PeriodicDemandType
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    RfqDetailId = rfqDetailId,
                                    SortNumber = i + 1,
                                    SolutionQty = Convert.ToInt32(solutionJson["data"][i]["SolutionQty"]),
                                    PeriodicDemandType = "M",
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();

                            int RfqLineSolutionId = -1;
                            foreach (var item in insertResult)
                            {
                                RfqLineSolutionId = item.RfqLineSolutionId;
                            }
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

        #region //UpdateRfqDetailSales -- 單身單筆指派業務更新 -- Yi 2023.07.13
        public string UpdateRfqDetailSales(int RfqDetailId, int CompanyId, int UserId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    string currentDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    string rfqNo = "", productUseName = "", rfqProductTypeName = "", assemblyName = "", memberName = "";

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷RFQ資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.RfqDetail a
                                WHERE a.RfqDetailId = @RfqDetailId";
                        dynamicParameters.Add("RfqDetailId", RfqDetailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("RFQ資料錯誤!");
                        #endregion

                        #region//更新RFQ單身資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.RfqDetail SET
                                CompanyId = @CompanyId,
                                SalesId = @SalesId,
                                ConfirmSalesTime = @ConfirmSalesTime,
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RfqDetailId = @RfqDetailId";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            CompanyId,
                            SalesId = UserId,
                            ConfirmSalesTime = currentDate,
                            Status = 2,
                            LastModifiedDate,
                            LastModifiedBy,
                            RfqDetailId
                        });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //發送Mail
                        #region //Mail資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MailSubject, a.MailContent
                                , b.Host, b.Port, b.SendMode, b.Account, b.Password
                                , ISNULL(c.MailFrom, '') MailFrom, ISNULL(d.MailTo, '') MailTo, ISNULL(e.MailCc, '') MailCc, ISNULL(f.MailBcc, '') MailBcc
                                FROM BAS.Mail a
                                LEFT JOIN BAS.MailServer b ON b.ServerId = a.ServerId
                                OUTER APPLY (
                                    SELECT STUFF((
                                        SELECT ';' + CASE WHEN ca.ContactId IS NULL THEN 
                                            CASE cc.JobType 
                                                WHEN '管理制' THEN cd.DepartmentName + cc.Job + '-' + cc.UserName + ':' + cc.Email
                                                ELSE cd.DepartmentName + '-' + cc.UserName + ':' + cc.Email
                                            END
                                        ELSE cb.ContactName + ':' + cb.Email END
                                        FROM BAS.MailUser ca
                                        LEFT JOIN BAS.MailContact cb ON ca.ContactId = cb.ContactId
                                        LEFT JOIN BAS.[User] cc ON ca.UserId = cc.UserId AND cc.[Status] = 'A'
                                        LEFT JOIN BAS.Department cd ON cc.DepartmentId = cd.DepartmentId
                                        WHERE ca.MailId = a.MailId
                                        AND ca.MailUserType = 'F'
                                        FOR XML PATH('')
                                    ), 1, 1, '') MailFrom
                                ) c
                                OUTER APPLY (
                                    SELECT STUFF((
                                        SELECT ';' + CASE WHEN da.ContactId IS NULL THEN 
                                            CASE dc.JobType 
                                                WHEN '管理制' THEN dd.DepartmentName + dc.Job + '-' + dc.UserName + ':' + dc.Email
                                                ELSE dd.DepartmentName + '-' + dc.UserName + ':' + dc.Email
                                            END
                                        ELSE db.ContactName + ':' + db.Email END
                                        FROM BAS.MailUser da
                                        LEFT JOIN BAS.MailContact db ON da.ContactId = db.ContactId
                                        LEFT JOIN BAS.[User] dc ON da.UserId = dc.UserId AND dc.[Status] = 'A'
                                        LEFT JOIN BAS.Department dd ON dc.DepartmentId = dd.DepartmentId
                                        WHERE da.MailId = a.MailId
                                        AND da.MailUserType = 'T'
                                        FOR XML PATH('')
                                    ), 1, 1, '') MailTo
                                ) d
                                OUTER APPLY (
                                    SELECT STUFF((
                                        SELECT ';' + CASE WHEN ea.ContactId IS NULL THEN 
                                            CASE ec.JobType 
                                                WHEN '管理制' THEN ed.DepartmentName + ec.Job + '-' + ec.UserName + ':' + ec.Email
                                                ELSE ed.DepartmentName + '-' + ec.UserName + ':' + ec.Email
                                            END
                                        ELSE eb.ContactName + ':' + eb.Email END
                                        FROM BAS.MailUser ea
                                        LEFT JOIN BAS.MailContact eb ON ea.ContactId = eb.ContactId
                                        LEFT JOIN BAS.[User] ec ON ea.UserId = ec.UserId AND ec.[Status] = 'A'
                                        LEFT JOIN BAS.Department ed ON ec.DepartmentId = ed.DepartmentId
                                        WHERE ea.MailId = a.MailId
                                        AND ea.MailUserType = 'C'
                                        FOR XML PATH('')
                                    ), 1, 1, '') MailCc
                                ) e
                                OUTER APPLY (
                                    SELECT STUFF((
                                        SELECT ';' + CASE WHEN fa.ContactId IS NULL THEN 
                                            CASE fc.JobType 
                                                WHEN '管理制' THEN fd.DepartmentName + fc.Job + '-' + fc.UserName + ':' + fc.Email
                                                ELSE fd.DepartmentName + '-' + fc.UserName + ':' + fc.Email
                                            END
                                        ELSE fb.ContactName + ':' + fb.Email END
                                        FROM BAS.MailUser fa
                                        LEFT JOIN BAS.MailContact fb ON fa.ContactId = fb.ContactId
                                        LEFT JOIN BAS.[User] fc ON fa.UserId = fc.UserId AND fc.[Status] = 'A'
                                        LEFT JOIN BAS.Department fd ON fc.DepartmentId = fd.DepartmentId
                                        WHERE fa.MailId = a.MailId
                                        AND fa.MailUserType = 'B'
                                        FOR XML PATH('')
                                    ), 1, 1, '') MailBcc
                                ) f
                                WHERE a.MailId IN (
                                    SELECT z.MailId
                                    FROM BAS.MailSendSetting z
                                    WHERE z.SettingSchema = @SettingSchema
                                    AND z.SettingNo = @SettingNo
                                )";
                        dynamicParameters.Add("SettingSchema", "RfqAssignSales");
                        dynamicParameters.Add("SettingNo", "Y");

                        var resultMailTemplate = sqlConnection.Query(sql, dynamicParameters);
                        if (resultMailTemplate.Count() <= 0) throw new SystemException("Mail設定錯誤!");
                        #endregion

                        #region //RFQ資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.RfqId, a.RfqNo, a.AssemblyName
                                , a.ProductUseId, c.ProductUseName, b.RfqProTypeId, e.RfqProductTypeName
                                , ISNULL(d.MemberId, f.UserId) MemberUserId, ISNULL(d.MemberName, f.UserName) MemberUserName, a.Status
                                FROM SCM.RequestForQuotation a
                                INNER JOIN SCM.RfqDetail b ON b.RfqId = a.RfqId
                                INNER JOIN SCM.ProductUse c ON c.ProductUseId = a.ProductUseId
                                LEFT JOIN EIP.Member d ON d.MemberId = a.MemberId
                                INNER JOIN SCM.RfqProductType e ON e.RfqProTypeId = b.RfqProTypeId
                                LEFT JOIN BAS.[User] f ON f.UserId = a.UserId
                                WHERE b.RfqDetailId = @RfqDetailId";
                        dynamicParameters.Add("RfqDetailId", RfqDetailId);

                        var resultRfqInfo = sqlConnection.Query(sql, dynamicParameters);
                        if (resultRfqInfo.Count() <= 0) throw new SystemException("RFQ資料錯誤!");

                        foreach (var item in resultRfqInfo)
                        {
                            rfqNo = item.RfqNo;
                            productUseName = item.ProductUseName;
                            rfqProductTypeName = item.RfqProductTypeName;
                            assemblyName = item.AssemblyName;
                            memberName = item.MemberUserName;
                        }
                        #endregion

                        #region //查找UserInfo
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserId, a.UserName, a.Email
                               FROM BAS.[User] a
                               WHERE a.UserId = @UserId";
                        dynamicParameters.Add("UserId", UserId);

                        var resultUserInfo = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUserInfo.Count() <= 0) throw new SystemException("收件者資料錯誤!");

                        string salesName = "";
                        string email = "";
                        foreach (var item in resultUserInfo)
                        {
                            UserId = item.UserId;
                            salesName = item.UserName;
                            email = item.Email;
                        }
                        #endregion

                        foreach (var item in resultMailTemplate)
                        {
                            string mailSubject = item.MailSubject,
                            mailContent = HttpUtility.UrlDecode(item.MailContent);

                            string hyperlink = "http://" + HttpContext.Current.Request.Url.Authority + "/RequestForQuotation/RfqAssignManagment";
                            rfqNo = string.Format("<a href=\"{0}\">{1}</a>", hyperlink, rfqNo);

                            #region //Mail內容
                            mailContent = mailContent.Replace("[RfqNo]", rfqNo);
                            mailContent = mailContent.Replace("[ProductUseName]", productUseName);
                            mailContent = mailContent.Replace("[RfqProductTypeName]", rfqProductTypeName);
                            mailContent = mailContent.Replace("[AssemblyName]", assemblyName);
                            mailContent = mailContent.Replace("[MemberName]", memberName);
                            mailContent = mailContent.Replace("[SalesName]", salesName);
                            #endregion

                            #region //寄送Mail
                            MailConfig mailConfig = new MailConfig
                            {
                                Host = item.Host,
                                Port = Convert.ToInt32(item.Port),
                                SendMode = Convert.ToInt32(item.SendMode),
                                From = item.MailFrom,
                                Subject = mailSubject,
                                Account = item.Account,
                                Password = item.Password,
                                MailTo = item.MailTo,
                                MailCc = item.MailCc,
                                MailBcc = item.MailBcc,
                                HtmlBody = mailContent,
                                TextBody = "-"
                            };

                            BaseHelper.MailSend(mailConfig);
                            #endregion
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
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //UpdateRfqConfirm -- 主頁批次確認狀態更新 -- Yi 2023.07.13
        public string UpdateRfqConfirm(int RfqId, int RfqDetailId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;

                    string currentDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    string rfqNo = "", productUseName = "", rfqProductTypeName = "", assemblyName = "", memberName = "";

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷RFQ單頭是否正確
                        sql = @"SELECT TOP 1 RfqId
                                FROM SCM.RequestForQuotation
                                WHERE RfqId = @RfqId";
                        dynamicParameters.Add("RfqId", RfqId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("RFQ單頭明細資料錯誤!");

                        int rfqId = -1;
                        foreach (var item in result)
                        {
                            rfqId = item.RfqId;
                        }
                        #endregion

                        #region //判斷RFQ資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT b.RfqDetailId
                                FROM SCM.RequestForQuotation a
                                INNER JOIN SCM.RfqDetail b ON b.RfqId = a.RfqId
                                WHERE a.RfqId = @RfqId";
                        dynamicParameters.Add("RfqId", rfqId);

                        var resultDetail = sqlConnection.Query(sql, dynamicParameters);
                        if (resultDetail.Count() <= 0) throw new SystemException("RFQ資料錯誤!");

                        int rfqDetailId = -1;
                        foreach (var item in resultDetail)
                        {
                            rfqDetailId = item.RfqDetailId;
                        }
                        #endregion

                        #region//更新RFQ單身確認狀態
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.RfqDetail SET
                                ConfirmRdTime = @ConfirmRdTime,
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RfqId = @RfqId";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            ConfirmRdTime = currentDate,
                            Status = 3,
                            LastModifiedDate,
                            LastModifiedBy,
                            RfqId = rfqId
                        });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //發送Mail
                        #region //Mail資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MailSubject, a.MailContent
                                , b.Host, b.Port, b.SendMode, b.Account, b.Password
                                , ISNULL(c.MailFrom, '') MailFrom, ISNULL(d.MailTo, '') MailTo, ISNULL(e.MailCc, '') MailCc, ISNULL(f.MailBcc, '') MailBcc
                                FROM BAS.Mail a
                                LEFT JOIN BAS.MailServer b ON b.ServerId = a.ServerId
                                OUTER APPLY (
                                    SELECT STUFF((
                                        SELECT ';' + CASE WHEN ca.ContactId IS NULL THEN 
                                            CASE cc.JobType 
                                                WHEN '管理制' THEN cd.DepartmentName + cc.Job + '-' + cc.UserName + ':' + cc.Email
                                                ELSE cd.DepartmentName + '-' + cc.UserName + ':' + cc.Email
                                            END
                                        ELSE cb.ContactName + ':' + cb.Email END
                                        FROM BAS.MailUser ca
                                        LEFT JOIN BAS.MailContact cb ON ca.ContactId = cb.ContactId
                                        LEFT JOIN BAS.[User] cc ON ca.UserId = cc.UserId AND cc.[Status] = 'A'
                                        LEFT JOIN BAS.Department cd ON cc.DepartmentId = cd.DepartmentId
                                        WHERE ca.MailId = a.MailId
                                        AND ca.MailUserType = 'F'
                                        FOR XML PATH('')
                                    ), 1, 1, '') MailFrom
                                ) c
                                OUTER APPLY (
                                    SELECT STUFF((
                                        SELECT ';' + CASE WHEN da.ContactId IS NULL THEN 
                                            CASE dc.JobType 
                                                WHEN '管理制' THEN dd.DepartmentName + dc.Job + '-' + dc.UserName + ':' + dc.Email
                                                ELSE dd.DepartmentName + '-' + dc.UserName + ':' + dc.Email
                                            END
                                        ELSE db.ContactName + ':' + db.Email END
                                        FROM BAS.MailUser da
                                        LEFT JOIN BAS.MailContact db ON da.ContactId = db.ContactId
                                        LEFT JOIN BAS.[User] dc ON da.UserId = dc.UserId AND dc.[Status] = 'A'
                                        LEFT JOIN BAS.Department dd ON dc.DepartmentId = dd.DepartmentId
                                        WHERE da.MailId = a.MailId
                                        AND da.MailUserType = 'T'
                                        FOR XML PATH('')
                                    ), 1, 1, '') MailTo
                                ) d
                                OUTER APPLY (
                                    SELECT STUFF((
                                        SELECT ';' + CASE WHEN ea.ContactId IS NULL THEN 
                                            CASE ec.JobType 
                                                WHEN '管理制' THEN ed.DepartmentName + ec.Job + '-' + ec.UserName + ':' + ec.Email
                                                ELSE ed.DepartmentName + '-' + ec.UserName + ':' + ec.Email
                                            END
                                        ELSE eb.ContactName + ':' + eb.Email END
                                        FROM BAS.MailUser ea
                                        LEFT JOIN BAS.MailContact eb ON ea.ContactId = eb.ContactId
                                        LEFT JOIN BAS.[User] ec ON ea.UserId = ec.UserId AND ec.[Status] = 'A'
                                        LEFT JOIN BAS.Department ed ON ec.DepartmentId = ed.DepartmentId
                                        WHERE ea.MailId = a.MailId
                                        AND ea.MailUserType = 'C'
                                        FOR XML PATH('')
                                    ), 1, 1, '') MailCc
                                ) e
                                OUTER APPLY (
                                    SELECT STUFF((
                                        SELECT ';' + CASE WHEN fa.ContactId IS NULL THEN 
                                            CASE fc.JobType 
                                                WHEN '管理制' THEN fd.DepartmentName + fc.Job + '-' + fc.UserName + ':' + fc.Email
                                                ELSE fd.DepartmentName + '-' + fc.UserName + ':' + fc.Email
                                            END
                                        ELSE fb.ContactName + ':' + fb.Email END
                                        FROM BAS.MailUser fa
                                        LEFT JOIN BAS.MailContact fb ON fa.ContactId = fb.ContactId
                                        LEFT JOIN BAS.[User] fc ON fa.UserId = fc.UserId AND fc.[Status] = 'A'
                                        LEFT JOIN BAS.Department fd ON fc.DepartmentId = fd.DepartmentId
                                        WHERE fa.MailId = a.MailId
                                        AND fa.MailUserType = 'B'
                                        FOR XML PATH('')
                                    ), 1, 1, '') MailBcc
                                ) f
                                WHERE a.MailId IN (
                                    SELECT z.MailId
                                    FROM BAS.MailSendSetting z
                                    WHERE z.SettingSchema = @SettingSchema
                                    AND z.SettingNo = @SettingNo
                                )";
                        dynamicParameters.Add("SettingSchema", "RfqNotificationRd");
                        dynamicParameters.Add("SettingNo", "Y");

                        var resultMailTemplate = sqlConnection.Query(sql, dynamicParameters);
                        if (resultMailTemplate.Count() <= 0) throw new SystemException("Mail設定錯誤!");
                        #endregion

                        #region //RFQ資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.RfqId, a.RfqNo, a.AssemblyName
                                , a.ProductUseId, c.ProductUseName, b.RfqProTypeId, e.RfqProductTypeName
                                , ISNULL(d.MemberId, f.UserId) MemberUserId, ISNULL(d.MemberName, f.UserName) MemberUserName
                                , b.SalesId, b.RfqDetailId, a.Status
                                FROM SCM.RequestForQuotation a
                                INNER JOIN SCM.RfqDetail b ON b.RfqId = a.RfqId
                                INNER JOIN SCM.ProductUse c ON c.ProductUseId = a.ProductUseId
                                LEFT JOIN EIP.Member d ON d.MemberId = a.MemberId
                                INNER JOIN SCM.RfqProductType e ON e.RfqProTypeId = b.RfqProTypeId
                                LEFT JOIN BAS.[User] f ON f.UserId = a.UserId
                                WHERE a.RfqId = @RfqId";
                        dynamicParameters.Add("RfqId", RfqId);

                        var resultRfqInfo = sqlConnection.Query(sql, dynamicParameters);
                        if (resultRfqInfo.Count() <= 0) throw new SystemException("RFQ資料錯誤!");

                        int salesId = -1;
                        int rfqDetailId = -1;
                        foreach (var item in resultRfqInfo)
                        {
                            rfqNo = item.RfqNo;
                            productUseName = item.ProductUseName;
                            rfqProductTypeName = item.RfqProductTypeName;
                            assemblyName = item.AssemblyName;
                            memberName = item.MemberUserName;
                            salesId = item.SalesId;
                            rfqDetailId = item.RfqDetailId;

                            #region //新增DFM單頭

                            #region //DFM單號取號
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT CONVERT(int, RIGHT(ISNULL(MAX(DfmNo), '000'), 3)) + 1 CurrentNum
                            FROM PDM.DesignForManufacturing
                            WHERE DfmNo LIKE @DfmNo";
                            dynamicParameters.Add("DfmNo", string.Format("{0}{1}___", "DFM", DateTime.Now.ToString("yyyyMMdd")));
                            int currentNum = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).CurrentNum;

                            string dfmNo = string.Format("{0}{1}{2}", "DFM", DateTime.Now.ToString("yyyyMMdd"), string.Format("{0:000}", currentNum));
                            #endregion

                            int rdUserId = -1;
                            string version = "000";
                            int? difficultyLevel = null;
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO PDM.DesignForManufacturing (DfmNo, RfqDetailId, DfmDate, DifficultyLevel, RdUserId, Version, DfmProcessStatus, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                VALUES (@DfmNo, @RfqDetailId, @DfmDate ,@DifficultyLevel, @RdUserId, @Version, @DfmProcessStatus, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    DfmNo = dfmNo,
                                    RfqDetailId = rfqDetailId,
                                    DfmDate = currentDate,
                                    DifficultyLevel = difficultyLevel,
                                    RdUserId = rdUserId <= 0 ? (int?)null : rdUserId,
                                    Version = version,
                                    DfmProcessStatus = 1,
                                    Status = "A", //啟用
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected = insertResult.Count();
                            #endregion
                        }
                        #endregion

                        #region //查找UserInfo
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserId, a.UserName, a.Email
                               FROM BAS.[User] a
                               WHERE a.UserId = @UserId";
                        dynamicParameters.Add("UserId", salesId);

                        var resultUserInfo = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUserInfo.Count() <= 0) throw new SystemException("收件者資料錯誤!");

                        string salesName = "";
                        string email = "";
                        foreach (var item in resultUserInfo)
                        {
                            salesId = item.UserId;
                            salesName = item.UserName;
                            email = item.Email;
                        }
                        #endregion

                        foreach (var item in resultMailTemplate)
                        {
                            string mailSubject = item.MailSubject,
                            mailContent = HttpUtility.UrlDecode(item.MailContent);

                            string hyperlink = "http://" + HttpContext.Current.Request.Url.Authority + "/DfmInformation/DesignForManufacturingManagement";
                            rfqNo = string.Format("<a href=\"{0}\">{1}</a>", hyperlink, rfqNo);

                            #region //Mail內容
                            mailContent = mailContent.Replace("[RfqNo]", rfqNo);
                            mailContent = mailContent.Replace("[ProductUseName]", productUseName);
                            mailContent = mailContent.Replace("[RfqProductTypeName]", rfqProductTypeName);
                            mailContent = mailContent.Replace("[AssemblyName]", assemblyName);
                            mailContent = mailContent.Replace("[MemberName]", memberName);
                            mailContent = mailContent.Replace("[SalesName]", salesName);
                            #endregion

                            #region //寄送Mail
                            MailConfig mailConfig = new MailConfig
                            {
                                Host = item.Host,
                                Port = Convert.ToInt32(item.Port),
                                SendMode = Convert.ToInt32(item.SendMode),
                                From = item.MailFrom,
                                Subject = mailSubject,
                                Account = item.Account,
                                Password = item.Password,
                                MailTo = item.MailTo,
                                MailCc = item.MailCc,
                                MailBcc = item.MailBcc,
                                HtmlBody = mailContent,
                                TextBody = "-"
                            };

                            BaseHelper.MailSend(mailConfig);
                            #endregion
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
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //UpdateRfqDetailConfirm -- 單身單筆確認狀態更新 -- Yi 2023.07.13
        public string UpdateRfqDetailConfirm(int RfqId, int RfqDetailId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;

                    string currentDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    string rfqNo = "", productUseName = "", rfqProductTypeName = "", assemblyName = "", memberName = "";

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷RFQ資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 ISNULL(a.KickOffType, '') KickOffType, ISNULL(a.PlasticName, '') PlasticName
                                , ISNULL(a.OutsideDiameter, 0) OutsideDiameter, ISNULL(a.ProtoSchedule, '') ProtoSchedule, ISNULL(a.PlannedOpeningDate, '') PlannedOpeningDate
                                FROM SCM.RfqDetail a
                                WHERE a.RfqDetailId = @RfqDetailId";
                        dynamicParameters.Add("RfqDetailId", RfqDetailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("RFQ資料錯誤!");

                        string KickOffType = "", PlasticName = "", ProtoSchedule = "", PlannedOpeningDate = "";
                        int OutsideDiameter = -1;
                        foreach (var item in result)
                        {
                            KickOffType = item.KickOffType;
                            PlasticName = item.PlasticName;
                            ProtoSchedule = item.ProtoSchedule;
                            PlannedOpeningDate = Convert.ToDateTime(item.PlannedOpeningDate).ToString("yyyy-MM-dd HH:mm:ss");
                            OutsideDiameter = Convert.ToInt32(item.OutsideDiameter);
                        }
                        #endregion

                        if (KickOffType == "" || PlasticName == "" || ProtoSchedule == "" || PlannedOpeningDate == "" || OutsideDiameter == -1)
                        {
                            throw new SystemException("RFQ單身資料未維護完整，請先填寫完整再確認!");
                        }
                        else
                        {
                            #region//更新RFQ單身確認狀態
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.RfqDetail SET
                                ConfirmRdTime = @ConfirmRdTime,
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RfqDetailId = @RfqDetailId";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmRdTime = currentDate,
                                Status = 3,
                                LastModifiedDate,
                                LastModifiedBy,
                                RfqDetailId
                            });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //發送Mail

                        #region //Mail資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MailSubject, a.MailContent
                                , b.Host, b.Port, b.SendMode, b.Account, b.Password
                                , ISNULL(c.MailFrom, '') MailFrom, ISNULL(d.MailTo, '') MailTo, ISNULL(e.MailCc, '') MailCc, ISNULL(f.MailBcc, '') MailBcc
                                FROM BAS.Mail a
                                LEFT JOIN BAS.MailServer b ON b.ServerId = a.ServerId
                                OUTER APPLY (
                                    SELECT STUFF((
                                        SELECT ';' + CASE WHEN ca.ContactId IS NULL THEN 
                                            CASE cc.JobType 
                                                WHEN '管理制' THEN cd.DepartmentName + cc.Job + '-' + cc.UserName + ':' + cc.Email
                                                ELSE cd.DepartmentName + '-' + cc.UserName + ':' + cc.Email
                                            END
                                        ELSE cb.ContactName + ':' + cb.Email END
                                        FROM BAS.MailUser ca
                                        LEFT JOIN BAS.MailContact cb ON ca.ContactId = cb.ContactId
                                        LEFT JOIN BAS.[User] cc ON ca.UserId = cc.UserId AND cc.[Status] = 'A'
                                        LEFT JOIN BAS.Department cd ON cc.DepartmentId = cd.DepartmentId
                                        WHERE ca.MailId = a.MailId
                                        AND ca.MailUserType = 'F'
                                        FOR XML PATH('')
                                    ), 1, 1, '') MailFrom
                                ) c
                                OUTER APPLY (
                                    SELECT STUFF((
                                        SELECT ';' + CASE WHEN da.ContactId IS NULL THEN 
                                            CASE dc.JobType 
                                                WHEN '管理制' THEN dd.DepartmentName + dc.Job + '-' + dc.UserName + ':' + dc.Email
                                                ELSE dd.DepartmentName + '-' + dc.UserName + ':' + dc.Email
                                            END
                                        ELSE db.ContactName + ':' + db.Email END
                                        FROM BAS.MailUser da
                                        LEFT JOIN BAS.MailContact db ON da.ContactId = db.ContactId
                                        LEFT JOIN BAS.[User] dc ON da.UserId = dc.UserId AND dc.[Status] = 'A'
                                        LEFT JOIN BAS.Department dd ON dc.DepartmentId = dd.DepartmentId
                                        WHERE da.MailId = a.MailId
                                        AND da.MailUserType = 'T'
                                        FOR XML PATH('')
                                    ), 1, 1, '') MailTo
                                ) d
                                OUTER APPLY (
                                    SELECT STUFF((
                                        SELECT ';' + CASE WHEN ea.ContactId IS NULL THEN 
                                            CASE ec.JobType 
                                                WHEN '管理制' THEN ed.DepartmentName + ec.Job + '-' + ec.UserName + ':' + ec.Email
                                                ELSE ed.DepartmentName + '-' + ec.UserName + ':' + ec.Email
                                            END
                                        ELSE eb.ContactName + ':' + eb.Email END
                                        FROM BAS.MailUser ea
                                        LEFT JOIN BAS.MailContact eb ON ea.ContactId = eb.ContactId
                                        LEFT JOIN BAS.[User] ec ON ea.UserId = ec.UserId AND ec.[Status] = 'A'
                                        LEFT JOIN BAS.Department ed ON ec.DepartmentId = ed.DepartmentId
                                        WHERE ea.MailId = a.MailId
                                        AND ea.MailUserType = 'C'
                                        FOR XML PATH('')
                                    ), 1, 1, '') MailCc
                                ) e
                                OUTER APPLY (
                                    SELECT STUFF((
                                        SELECT ';' + CASE WHEN fa.ContactId IS NULL THEN 
                                            CASE fc.JobType 
                                                WHEN '管理制' THEN fd.DepartmentName + fc.Job + '-' + fc.UserName + ':' + fc.Email
                                                ELSE fd.DepartmentName + '-' + fc.UserName + ':' + fc.Email
                                            END
                                        ELSE fb.ContactName + ':' + fb.Email END
                                        FROM BAS.MailUser fa
                                        LEFT JOIN BAS.MailContact fb ON fa.ContactId = fb.ContactId
                                        LEFT JOIN BAS.[User] fc ON fa.UserId = fc.UserId AND fc.[Status] = 'A'
                                        LEFT JOIN BAS.Department fd ON fc.DepartmentId = fd.DepartmentId
                                        WHERE fa.MailId = a.MailId
                                        AND fa.MailUserType = 'B'
                                        FOR XML PATH('')
                                    ), 1, 1, '') MailBcc
                                ) f
                                WHERE a.MailId IN (
                                    SELECT z.MailId
                                    FROM BAS.MailSendSetting z
                                    WHERE z.SettingSchema = @SettingSchema
                                    AND z.SettingNo = @SettingNo
                                )";
                        dynamicParameters.Add("SettingSchema", "RfqNotificationRd");
                        dynamicParameters.Add("SettingNo", "Y");

                        var resultMailTemplate = sqlConnection.Query(sql, dynamicParameters);
                        if (resultMailTemplate.Count() <= 0) throw new SystemException("Mail設定錯誤!");
                        #endregion

                        #region //RFQ資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.RfqId, a.RfqNo, a.AssemblyName
                                , a.ProductUseId, c.ProductUseName, b.RfqProTypeId, e.RfqProductTypeName
                                , ISNULL(d.MemberId, f.UserId) MemberUserId, ISNULL(d.MemberName, f.UserName) MemberUserName
                                , b.SalesId, b.RfqDetailId, a.Status
                                FROM SCM.RequestForQuotation a
                                INNER JOIN SCM.RfqDetail b ON b.RfqId = a.RfqId
                                INNER JOIN SCM.ProductUse c ON c.ProductUseId = a.ProductUseId
                                LEFT JOIN EIP.Member d ON d.MemberId = a.MemberId
                                INNER JOIN SCM.RfqProductType e ON e.RfqProTypeId = b.RfqProTypeId
                                LEFT JOIN BAS.[User] f ON f.UserId = a.UserId
                                WHERE b.RfqDetailId = @RfqDetailId";
                        dynamicParameters.Add("RfqDetailId", RfqDetailId);

                        var resultRfqInfo = sqlConnection.Query(sql, dynamicParameters);
                        if (resultRfqInfo.Count() <= 0) throw new SystemException("RFQ資料錯誤!");

                        int salesId = -1;
                        int rfqDetailId = -1;
                        int rfqProTypeId = -1;
                        foreach (var item in resultRfqInfo)
                        {
                            rfqNo = item.RfqNo;
                            productUseName = item.ProductUseName;
                            rfqProductTypeName = item.RfqProductTypeName;
                            assemblyName = item.AssemblyName;
                            memberName = item.MemberUserName;
                            salesId = item.SalesId;
                            rfqDetailId = item.RfqDetailId;
                            rfqProTypeId = item.RfqProTypeId;
                        }
                        #endregion

                        #region //查找UserInfo
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserId, a.UserName, a.Email
                               FROM BAS.[User] a
                               WHERE a.UserId = @UserId";
                        dynamicParameters.Add("UserId", salesId);

                        var resultUserInfo = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUserInfo.Count() <= 0) throw new SystemException("收件者資料錯誤!");

                        string salesName = "";
                        string selesemail = "";
                        foreach (var item in resultUserInfo)
                        {
                            salesId = item.UserId;
                            salesName = item.UserName;
                            selesemail = item.Email;
                        }
                        #endregion

                        #region //依據產品類別查找欲通知之RD群組
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT DISTINCT b.RoleId, b.UserId, c.UserNo, c.UserName, d.RfqProTypeId
                                , a.DocType, a.SubProdId, a.SubProdType
                                , ISNULL(e.DisplayName, '') DisplayName, ISNULL(e.Email, '') Email
                                FROM EIP.DocNotify a
                                INNER JOIN EIP.NotifyUser b ON b.RoleId = a.RoleId
                                INNER JOIN BAS.[User] c ON c.UserId = b.UserId
                                INNER JOIN SCM.RfqDetail d ON d.RfqProTypeId = a.SubProdId
                                OUTER APPLY (
                                    SELECT ea.Email
                                    , CASE ea.JobType 
                                        WHEN '管理制' THEN eb.DepartmentName + ea.Job + '-' + ea.UserName
                                        ELSE eb.DepartmentName + '-' + ea.UserName 
                                    END DisplayName
                                    FROM BAS.[User] ea
                                    INNER JOIN BAS.Department eb ON ea.DepartmentId = eb.DepartmentId
                                    INNER JOIN BAS.Company ec ON eb.CompanyId = ec.CompanyId
                                    WHERE ea.Status = @Status
                                    AND ea.UserId = c.UserId
                                    AND ea.Email IS NOT NULL
                                    AND ea.Email <> ''
                                ) e
                                WHERE a.DocType = 'RFQ'
                                AND a.SubProdId = @SubProdId";
                        dynamicParameters.Add("Status", 'A');
                        dynamicParameters.Add("SubProdId", rfqProTypeId);
                        var resultRdGroup = sqlConnection.Query(sql, dynamicParameters);
                        if (resultRdGroup.Count() <= 0) throw new SystemException("【通知角色群組】資料錯誤或未設定");

                        int index = 1;
                        string rdName = "";
                        string rdEmail = "";

                        foreach (var item in resultRdGroup)
                        {
                            rdName += item.UserName;
                            rdEmail += item.DisplayName + ':' + item.Email;
                            if (index < resultRdGroup.Count())
                            {
                                rdName += ',';
                                rdEmail += ';';
                            }
                            index++;
                        }
                        #endregion

                        foreach (var item in resultMailTemplate)
                        {
                            string mailSubject = item.MailSubject,
                            mailContent = HttpUtility.UrlDecode(item.MailContent);

                            string hyperlink = "http://" + HttpContext.Current.Request.Url.Authority + "/DfmInformation/DesignForManufacturingManagement";
                            rfqNo = string.Format("<a href=\"{0}\">{1}</a>", hyperlink, rfqNo);

                            #region //Mail內容
                            mailContent = mailContent.Replace("[RfqNo]", rfqNo);
                            mailContent = mailContent.Replace("[ProductUseName]", productUseName);
                            mailContent = mailContent.Replace("[RfqProductTypeName]", rfqProductTypeName);
                            mailContent = mailContent.Replace("[AssemblyName]", assemblyName);
                            mailContent = mailContent.Replace("[MemberName]", memberName);
                            mailContent = mailContent.Replace("[SalesName]", salesName);
                            mailContent = mailContent.Replace("[RdName]", rdName);
                            #endregion

                            #region //寄送Mail
                            MailConfig mailConfig = new MailConfig
                            {
                                Host = item.Host,
                                Port = Convert.ToInt32(item.Port),
                                SendMode = Convert.ToInt32(item.SendMode),
                                From = item.MailFrom,
                                Subject = mailSubject,
                                Account = item.Account,
                                Password = item.Password,
                                MailTo = rdEmail, //變數通知RD群組
                                MailCc = item.MailCc,
                                MailBcc = item.MailBcc,
                                HtmlBody = mailContent,
                                TextBody = "-"
                            };

                            BaseHelper.MailSend(mailConfig);
                            #endregion
                        }
                        #endregion

                        #region //新增DFM單頭

                        #region //DFM單號取號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT CONVERT(int, RIGHT(ISNULL(MAX(DfmNo), '000'), 3)) + 1 CurrentNum
                            FROM PDM.DesignForManufacturing
                            WHERE DfmNo LIKE @DfmNo";
                        dynamicParameters.Add("DfmNo", string.Format("{0}{1}___", "DFM", DateTime.Now.ToString("yyyyMMdd")));
                        int currentNum = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).CurrentNum;

                        string dfmNo = string.Format("{0}{1}{2}", "DFM", DateTime.Now.ToString("yyyyMMdd"), string.Format("{0:000}", currentNum));
                        #endregion

                        #region //取得預設Default SpreadsheetData
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 DefaultSpreadsheetData
                                FROM PDM.DesignForManufacturing
                                WHERE DefaultSpreadsheetData IS NOT NULL";

                        var defaultSpreadsheetResult = sqlConnection.Query(sql, dynamicParameters);

                        string DefaultSpreadsheetData = "";
                        foreach (var item in defaultSpreadsheetResult)
                        {
                            DefaultSpreadsheetData = item.DefaultSpreadsheetData;
                        }
                        #endregion

                        int rdUserId = -1;
                        string version = "000";
                        int? difficultyLevel = null;
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO PDM.DesignForManufacturing (DfmNo, RfqDetailId, DfmDate, DifficultyLevel
                                , RdUserId, Version, DfmProcessStatus, Status, DefaultSpreadsheetData
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                VALUES (@DfmNo, @RfqDetailId, @DfmDate ,@DifficultyLevel
                                , @RdUserId, @Version, @DfmProcessStatus, @Status, @DefaultSpreadsheetData
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                DfmNo = dfmNo,
                                RfqDetailId = rfqDetailId,
                                DfmDate = currentDate,
                                DifficultyLevel = difficultyLevel,
                                RdUserId = rdUserId <= 0 ? (int?)null : rdUserId,
                                Version = version,
                                DfmProcessStatus = 1,
                                Status = "A", //啟用
                                DefaultSpreadsheetData,
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

        #region //UpdateDfmQiSolution -- RFQ單身報價資料更新(單筆儲存) -- Yi 2023.07.31
        public string UpdateDfmQiSolution(int DfmQiId, int RfqLineSolutionId
            , double DiscountAmount, double GrossProfitMargin, double AfterProfitMargin, double QuotationAmount)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷RFQ方案對應之報價是否正確
                        sql = @"SELECT TOP 1 a.DfmQiSolutionId
                                FROM PDM.DfmQiSolution a
                                INNER JOIN SCM.RfqLineSolution b ON b.RfqLineSolutionId = a.RfqLineSolutionId
                                WHERE a.RfqLineSolutionId = @RfqLineSolutionId
                                AND a.DfmQiId = @DfmQiId";
                        dynamicParameters.Add("RfqLineSolutionId", RfqLineSolutionId);
                        dynamicParameters.Add("DfmQiId", DfmQiId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("報價方案資料錯誤!");

                        int dfmQiSolutionId = -1;
                        foreach (var item in result)
                        {
                            dfmQiSolutionId = item.DfmQiSolutionId;
                        }
                        #endregion

                        #region//更新DfmQiSolution報價明細
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PDM.DfmQiSolution SET
                                DfmQiId = @DfmQiId,
                                RfqLineSolutionId = @RfqLineSolutionId,
                                GrossProfitMargin = @GrossProfitMargin,
                                AfterProfitMargin = @AfterProfitMargin,
                                DiscountAmount = @DiscountAmount,
                                QuotationAmount = @QuotationAmount,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE DfmQiSolutionId = @DfmQiSolutionId";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            DfmQiId,
                            RfqLineSolutionId,
                            GrossProfitMargin,
                            AfterProfitMargin,
                            DiscountAmount,
                            QuotationAmount,
                            LastModifiedDate,
                            LastModifiedBy,
                            DfmQiSolutionId = dfmQiSolutionId
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

        #region //UpdateLotDfmQiSolution -- RFQ單身報價資料更新(多筆儲存) -- Yi 2023.08.10
        public string UpdateLotDfmQiSolution(string Quotations)
        {
            try
            {
                if (!Quotations.TryParseJson(out JObject tempJObject)) throw new SystemException("批量儲存報價格式錯誤");

                JObject quotationJson = JObject.Parse(Quotations);
                if (!quotationJson.ContainsKey("data")) throw new SystemException("批量儲存報價資料錯誤");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        for (int i = 0; i < quotationJson["data"].Count(); i++)
                        {
                            int index = Convert.ToInt32(quotationJson["data"][i]["index"]);
                            int rfqLineSolutionId = Convert.ToInt32(quotationJson["data"][i]["rfqLineSolutionId"]);
                            int dfmQiId = Convert.ToInt32(quotationJson["data"][i]["dfmQiId"]);
                            double grossProfitMargin = Convert.ToDouble(quotationJson["data"][i]["grossProfitMargin"]);
                            double afterProfitMargin = Convert.ToDouble(quotationJson["data"][i]["afterProfitMargin"]);
                            double discountAmount = Convert.ToDouble(quotationJson["data"][i]["discountAmount"]);
                            double quotationAmount = Convert.ToDouble(quotationJson["data"][i]["quotationAmount"]);
                            int rowsAffected = 0;

                            #region //判斷RFQ方案對應之報價是否正確
                            sql = @"SELECT TOP 1 a.DfmQiSolutionId
                                    FROM PDM.DfmQiSolution a
                                    INNER JOIN SCM.RfqLineSolution b ON b.RfqLineSolutionId = a.RfqLineSolutionId
                                    WHERE a.RfqLineSolutionId = @RfqLineSolutionId
                                    AND a.DfmQiId = @DfmQiId";
                            dynamicParameters.Add("RfqLineSolutionId", rfqLineSolutionId);
                            dynamicParameters.Add("DfmQiId", dfmQiId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("報價方案資料錯誤!");

                            int dfmQiSolutionId = -1;
                            foreach (var item in result)
                            {
                                dfmQiSolutionId = item.DfmQiSolutionId;
                            }
                            #endregion

                            #region//更新DfmQiSolution報價明細
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE PDM.DfmQiSolution SET
                                DfmQiId = @DfmQiId,
                                RfqLineSolutionId = @RfqLineSolutionId,
                                GrossProfitMargin = @GrossProfitMargin,
                                AfterProfitMargin = @AfterProfitMargin,
                                DiscountAmount = @DiscountAmount,
                                QuotationAmount = @QuotationAmount,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE DfmQiSolutionId = @DfmQiSolutionId";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                DfmQiId = dfmQiId,
                                RfqLineSolutionId = rfqLineSolutionId,
                                GrossProfitMargin = grossProfitMargin,
                                AfterProfitMargin = afterProfitMargin,
                                DiscountAmount = discountAmount,
                                QuotationAmount = quotationAmount,
                                LastModifiedDate,
                                LastModifiedBy,
                                DfmQiSolutionId = dfmQiSolutionId
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

        #region //UpdateDfmQiSolutionConfirm -- RFQ報價確認狀態更新 -- Yi 2023.07.31
        public string UpdateDfmQiSolutionConfirm(int RfqId, int RfqDetailId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    string currentDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    string rfqNo = "", rfqSequence = "", productUseName = "", rfqProductTypeName = "", assemblyName = ""
                        , memberName = "", memberEmail = "";

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷RFQ資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.RfqDetail a
                                INNER JOIN SCM.RfqLineSolution b ON b.RfqDetailId = a.RfqDetailId
                                INNER JOIN PDM.DfmQiSolution c ON c.RfqLineSolutionId = b.RfqLineSolutionId
                                WHERE a.RfqDetailId = @RfqDetailId";
                        dynamicParameters.Add("RfqDetailId", RfqDetailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("RFQ資料錯誤!");
                        #endregion

                        #region//更新RFQ單身確認狀態
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.RfqDetail SET
                                ConfirmCustTime = @ConfirmCustTime,
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RfqDetailId = @RfqDetailId";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            ConfirmCustTime = currentDate,
                            Status = 6,
                            LastModifiedDate,
                            LastModifiedBy,
                            RfqDetailId
                        });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //發送Mail
                        #region //Mail資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MailSubject, a.MailContent
                                , b.Host, b.Port, b.SendMode, b.Account, b.Password
                                , ISNULL(c.MailFrom, '') MailFrom, ISNULL(d.MailTo, '') MailTo, ISNULL(e.MailCc, '') MailCc, ISNULL(f.MailBcc, '') MailBcc
                                FROM BAS.Mail a
                                LEFT JOIN BAS.MailServer b ON b.ServerId = a.ServerId
                                OUTER APPLY (
                                    SELECT STUFF((
                                        SELECT ';' + CASE WHEN ca.ContactId IS NULL THEN 
                                            CASE cc.JobType 
                                                WHEN '管理制' THEN cd.DepartmentName + cc.Job + '-' + cc.UserName + ':' + cc.Email
                                                ELSE cd.DepartmentName + '-' + cc.UserName + ':' + cc.Email
                                            END
                                        ELSE cb.ContactName + ':' + cb.Email END
                                        FROM BAS.MailUser ca
                                        LEFT JOIN BAS.MailContact cb ON ca.ContactId = cb.ContactId
                                        LEFT JOIN BAS.[User] cc ON ca.UserId = cc.UserId AND cc.[Status] = 'A'
                                        LEFT JOIN BAS.Department cd ON cc.DepartmentId = cd.DepartmentId
                                        WHERE ca.MailId = a.MailId
                                        AND ca.MailUserType = 'F'
                                        FOR XML PATH('')
                                    ), 1, 1, '') MailFrom
                                ) c
                                OUTER APPLY (
                                    SELECT STUFF((
                                        SELECT ';' + CASE WHEN da.ContactId IS NULL THEN 
                                            CASE dc.JobType 
                                                WHEN '管理制' THEN dd.DepartmentName + dc.Job + '-' + dc.UserName + ':' + dc.Email
                                                ELSE dd.DepartmentName + '-' + dc.UserName + ':' + dc.Email
                                            END
                                        ELSE db.ContactName + ':' + db.Email END
                                        FROM BAS.MailUser da
                                        LEFT JOIN BAS.MailContact db ON da.ContactId = db.ContactId
                                        LEFT JOIN BAS.[User] dc ON da.UserId = dc.UserId AND dc.[Status] = 'A'
                                        LEFT JOIN BAS.Department dd ON dc.DepartmentId = dd.DepartmentId
                                        WHERE da.MailId = a.MailId
                                        AND da.MailUserType = 'T'
                                        FOR XML PATH('')
                                    ), 1, 1, '') MailTo
                                ) d
                                OUTER APPLY (
                                    SELECT STUFF((
                                        SELECT ';' + CASE WHEN ea.ContactId IS NULL THEN 
                                            CASE ec.JobType 
                                                WHEN '管理制' THEN ed.DepartmentName + ec.Job + '-' + ec.UserName + ':' + ec.Email
                                                ELSE ed.DepartmentName + '-' + ec.UserName + ':' + ec.Email
                                            END
                                        ELSE eb.ContactName + ':' + eb.Email END
                                        FROM BAS.MailUser ea
                                        LEFT JOIN BAS.MailContact eb ON ea.ContactId = eb.ContactId
                                        LEFT JOIN BAS.[User] ec ON ea.UserId = ec.UserId AND ec.[Status] = 'A'
                                        LEFT JOIN BAS.Department ed ON ec.DepartmentId = ed.DepartmentId
                                        WHERE ea.MailId = a.MailId
                                        AND ea.MailUserType = 'C'
                                        FOR XML PATH('')
                                    ), 1, 1, '') MailCc
                                ) e
                                OUTER APPLY (
                                    SELECT STUFF((
                                        SELECT ';' + CASE WHEN fa.ContactId IS NULL THEN 
                                            CASE fc.JobType 
                                                WHEN '管理制' THEN fd.DepartmentName + fc.Job + '-' + fc.UserName + ':' + fc.Email
                                                ELSE fd.DepartmentName + '-' + fc.UserName + ':' + fc.Email
                                            END
                                        ELSE fb.ContactName + ':' + fb.Email END
                                        FROM BAS.MailUser fa
                                        LEFT JOIN BAS.MailContact fb ON fa.ContactId = fb.ContactId
                                        LEFT JOIN BAS.[User] fc ON fa.UserId = fc.UserId AND fc.[Status] = 'A'
                                        LEFT JOIN BAS.Department fd ON fc.DepartmentId = fd.DepartmentId
                                        WHERE fa.MailId = a.MailId
                                        AND fa.MailUserType = 'B'
                                        FOR XML PATH('')
                                    ), 1, 1, '') MailBcc
                                ) f
                                WHERE a.MailId IN (
                                    SELECT z.MailId
                                    FROM BAS.MailSendSetting z
                                    WHERE z.SettingSchema = @SettingSchema
                                    AND z.SettingNo = @SettingNo
                                )";
                        dynamicParameters.Add("SettingSchema", "RfqNotificationCust");
                        dynamicParameters.Add("SettingNo", "Y");

                        var resultMailTemplate = sqlConnection.Query(sql, dynamicParameters);
                        if (resultMailTemplate.Count() <= 0) throw new SystemException("Mail設定錯誤!");
                        #endregion

                        #region //RFQ資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.RfqId, a.RfqNo, a.AssemblyName
                                , a.ProductUseId, c.ProductUseName, b.RfqProTypeId, e.RfqProductTypeName
                                , ISNULL(d.MemberId, f.UserId) MemberUserId, ISNULL(d.MemberName, f.UserName) MemberUserName
                                , b.SalesId, b.RfqDetailId, b.RfqSequence, a.Status
                                FROM SCM.RequestForQuotation a
                                INNER JOIN SCM.RfqDetail b ON b.RfqId = a.RfqId
                                INNER JOIN SCM.ProductUse c ON c.ProductUseId = a.ProductUseId
                                LEFT JOIN EIP.Member d ON d.MemberId = a.MemberId
                                INNER JOIN SCM.RfqProductType e ON e.RfqProTypeId = b.RfqProTypeId
                                LEFT JOIN BAS.[User] f ON f.UserId = a.UserId
                                WHERE b.RfqDetailId = @RfqDetailId";
                        dynamicParameters.Add("RfqDetailId", RfqDetailId);

                        var resultRfqInfo = sqlConnection.Query(sql, dynamicParameters);
                        if (resultRfqInfo.Count() <= 0) throw new SystemException("RFQ資料錯誤!");

                        int salesId = -1;
                        int rfqId = -1;
                        int rfqDetailId = -1;
                        foreach (var item in resultRfqInfo)
                        {
                            rfqNo = item.RfqNo;
                            rfqSequence = item.RfqSequence;
                            productUseName = item.ProductUseName;
                            rfqProductTypeName = item.RfqProductTypeName;
                            assemblyName = item.AssemblyName;
                            memberName = item.MemberUserName;
                            memberEmail = item.MemberEmail;
                            salesId = item.SalesId;
                            rfqId = item.RfqId;
                            rfqDetailId = item.RfqDetailId;
                        }
                        #endregion

                        #region //查找UserInfo
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserId, a.UserName, a.Email
                               FROM BAS.[User] a
                               WHERE a.UserId = @UserId";
                        dynamicParameters.Add("UserId", salesId);

                        var resultUserInfo = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUserInfo.Count() <= 0) throw new SystemException("收件者資料錯誤!");

                        string salesName = "";
                        string email = "";
                        foreach (var item in resultUserInfo)
                        {
                            salesId = item.UserId;
                            salesName = item.UserName;
                            email = item.Email;
                        }
                        #endregion

                        foreach (var item in resultMailTemplate)
                        {
                            string mailSubject = item.MailSubject,
                            mailContent = HttpUtility.UrlDecode(item.MailContent);

                            //string hyperlink = "http://" + HttpContext.Current.Request.Url.Authority + "/RequestForQuotation/RfqLineSolution";
                            //rfqNo = string.Format("<a href=\"{0}\">{1}</a>", hyperlink, rfqNo);

                            #region //Mail內容
                            mailContent = mailContent.Replace("[RfqNo]", rfqNo);
                            mailContent = mailContent.Replace("[RfqSequence]", rfqSequence);
                            mailContent = mailContent.Replace("[ProductUseName]", productUseName);
                            mailContent = mailContent.Replace("[RfqProductTypeName]", rfqProductTypeName);
                            mailContent = mailContent.Replace("[AssemblyName]", assemblyName);
                            mailContent = mailContent.Replace("[MemberName]", memberName);
                            mailContent = mailContent.Replace("[SalesName]", salesName);
                            #endregion

                            #region //SCM.RfqDetail QuotationFile
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.RfqDetailId
                                , b.FileId, a.QuotationFile, b.FileContent, b.[FileName], b.FileExtension
                                , ISNULL(d.MemberName, e.UserName) MemberUserName
                                , (ISNULL(d.MemberName, e.UserName) + '-' + c.AssemblyName + '-' + a.RfqSequence) QuotationFileName
                                FROM SCM.RfqDetail a
                                INNER JOIN BAS.[File] b ON b.FileId = a.QuotationFile
                                INNER JOIN SCM.RequestForQuotation c ON c.RfqId = a.RfqId
                                LEFT JOIN EIP.Member d ON d.MemberId = c.MemberId
                                LEFT JOIN BAS.[User] e ON e.UserId = c.UserId
                                WHERE a.RfqDetailId  = @RfqDetailId";
                            dynamicParameters.Add("RfqDetailId", RfqDetailId);

                            var QuotationResult = sqlConnection.Query(sql, dynamicParameters);

                            List<MailFile> mailFiles = new List<MailFile>();
                            int fileId = -1;
                            string fileName = "";
                            foreach (var quotation in QuotationResult)
                            {
                                MailFile mailFile = new MailFile()
                                {
                                    FileName = quotation.QuotationFileName,
                                    FileExtension = quotation.FileExtension,
                                    FileContent = (byte[])quotation.FileContent
                                };

                                mailFiles.Add(mailFile);
                                fileId = quotation.FileId;
                                fileName = quotation.QuotationFileName;
                            }
                            #endregion

                            #region//更新BAS.File FileName
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE BAS.[File] SET
                                FileName = @FileName,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE FileId = @FileId";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                FileName = fileName,
                                LastModifiedDate,
                                LastModifiedBy,
                                FileId = fileId
                            });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //寄送Mail
                            MailConfig mailConfig = new MailConfig
                            {
                                Host = item.Host,
                                Port = Convert.ToInt32(item.Port),
                                SendMode = Convert.ToInt32(item.SendMode),
                                From = item.MailFrom,
                                Subject = mailSubject,
                                Account = item.Account,
                                Password = item.Password,
                                MailTo = item.MailTo,
                                MailCc = item.MailCc,
                                MailBcc = item.MailBcc,
                                HtmlBody = mailContent,
                                TextBody = "-",
                                FileInfo = mailFiles
                            };

                            BaseHelper.MailSend(mailConfig);
                            #endregion
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
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //UpdateRfqDetailQuotation -- RFQ單身報價更新 -- Luca 2023.08.02
        public string UpdateRfqDetailQuotation(int FileId,int RfqDetailId)
        {
            try
            {
                #region //判斷單身資料長度
                if (RfqDetailId <= 0) throw new SystemException("【RFQ單身ID】不能為空!");
                if (FileId <= 0) throw new SystemException("【報價】不能為空!");
                #endregion

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷RFQ明細是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT b.RfqDetailId, b.SalesId
                                FROM SCM.RequestForQuotation a
                                INNER JOIN SCM.RfqDetail b ON b.RfqId = a.RfqId
                                WHERE b.RfqDetailId = @RfqDetailId";
                        dynamicParameters.Add("RfqDetailId", RfqDetailId);

                        var resultDetail = sqlConnection.Query(sql, dynamicParameters);
                        if (resultDetail.Count() <= 0) throw new SystemException("RFQ單身明細資料錯誤!");                        
                        #endregion

                        #region//更新RFQ單身明細
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.RfqDetail SET
                                QuotationFile = @FileId,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RfqDetailId = @RfqDetailId";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            FileId,
                            LastModifiedDate,
                            LastModifiedBy,
                            RfqDetailId
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

        #region //UpdateRfqQuotationTag -- RFQ報價單標籤更新 -- Shintokuro 2024.08.23
        public string UpdateRfqQuotationTag(int RfqDetailId, string TagList)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    if(RfqDetailId < 0) throw new SystemException("【RFQ單身】不可以為空,請重新確認!");
                    if(TagList.Length <= 0) throw new SystemException("【報價標籤屬性】不可以為空,請重新確認!");

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        #region //判斷RFQ單身是否正確
                        sql = @"SELECT TOP 1 RfqDetailId
                                FROM SCM.RfqDetail
                                WHERE RfqDetailId = @RfqDetailId
                                AND DocType = 'E'";
                        dynamicParameters.Add("RfqDetailId", RfqDetailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【RFQ單身】資料找不到,請重新確認!");
                        #endregion


                        foreach (var QtId in TagList.Split(','))
                        {
                            #region //判斷報價標籤欄位是否存在
                            sql = @"SELECT TOP 1 1
                                    FROM SCM.QuotationTag
                                    WHERE QtId = @QtId";
                            dynamicParameters.Add("QtId", QtId);

                            var result1 = sqlConnection.Query(sql, dynamicParameters);
                            if (result1.Count() <= 0) throw new SystemException("【報價標籤】資料找不到,請重新確認!");
                            #endregion
                        }

                        #region//更新RFQ明細
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.RfqDetail SET
                                    TagList = @TagList,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE RfqDetailId = @RfqDetailId";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            TagList,
                            LastModifiedDate,
                            LastModifiedBy,
                            RfqDetailId
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

        #region //UpdateQuotationRemark -- RFQ報價單備註更新 -- Shintokuro 2024.08.23
        public string UpdateQuotationRemark(int RfqDetailId, string QuotationRemark)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    string dateNow = CreateDate.ToString("yyyyMMdd");
                    string QuotationNo = "";

                    if (RfqDetailId < 0) throw new SystemException("【RFQ單身】不可以為空,請重新確認!");
                    if (QuotationRemark.Length <= 0) throw new SystemException("【報價單備註】不可以為空,請重新確認!");

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        #region //判斷RFQ單身是否正確
                        sql = @"SELECT TOP 1 RfqDetailId
                                FROM SCM.RfqDetail
                                WHERE RfqDetailId = @RfqDetailId
                                AND DocType = 'E'";
                        dynamicParameters.Add("RfqDetailId", RfqDetailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【RFQ單身】資料找不到,請重新確認!");
                        #endregion

                        #region//更新RFQ明細
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.RfqDetail SET
                                QuotationRemark = @QuotationRemark,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RfqDetailId = @RfqDetailId";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            QuotationRemark,
                            LastModifiedDate,
                            LastModifiedBy,
                            RfqDetailId
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

        #region //UpdateQuotationDetail -- 報價單項目資料更新 -- Shintokuro 2024.08.12
        public string UpdateQuotationDetail(int QiId, string Data, string AiData)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    var DataList = JsonConvert.DeserializeObject<Dictionary<string, string>>(Data);
                    var AiDataList = JsonConvert.DeserializeObject<Dictionary<string, string>>(AiData);


                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        bool ChargeAction = false;
                        int RfqDetailId = -1;
                        string HtmlColumns = "";

                        #region //撈取登入者工號
                        string CreateByUserNo = "";
                        sql = @"SELECT TOP 1 UserNo
                                FROM BAS.[User]
                                WHERE UserId = @CreateBy";
                        dynamicParameters.Add("CreateBy", CreateBy);

                        var resultUser = sqlConnection.Query(sql, dynamicParameters);
                        foreach(var item in resultUser)
                        {
                            CreateByUserNo = item.UserNo;
                        }
                        #endregion

                        #region //判斷報價單是否正確
                        sql = @"SELECT TOP 1 ConfirmStatus,Charge,RfqDetailId,HtmlColumns
                                FROM SCM.QuotationItem
                                WHERE QiId = @QiId";
                        dynamicParameters.Add("QiId", QiId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("報價單資料找不到,請重新確認!");
                        foreach(var item in result)
                        {
                            if(item.ConfirmStatus == "Y") throw new SystemException("該項目已確認不可以異動!!");
                            string ChargeId = item.Charge;
                            RfqDetailId = item.RfqDetailId;
                            HtmlColumns = item.HtmlColumns;
                            foreach (var item1 in ChargeId.Split(','))
                            {
                                if(item1 == CreateBy.ToString() || item1 == CreateByUserNo.ToString())
                                {
                                    ChargeAction = true;
                                    break;
                                }
                            }
                            if(!ChargeAction) throw new SystemException("尚無該項目可異動權限!!請重新確認!!");
                        }
                        #endregion

                        if(HtmlColumns == "B")
                        {
                            sql = @"SELECT TOP 1 ConfirmStatus
                                    FROM SCM.QuotationItem
                                    WHERE RfqDetailId = @RfqDetailId
                                    AND HtmlColumns = 'A'";
                            dynamicParameters.Add("RfqDetailId", RfqDetailId);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("報價單資料找不到,請重新確認!");
                            foreach (var item in result)
                            {
                                if (item.ConfirmStatus != "Y") throw new SystemException("研發項目尚未確認,不可以執行!!");
                            }
                            string CycleTimeAdoption = "", AICycleTime = "", MPCycleTime = "";
                            foreach (var item in AiDataList)
                            {
                                switch (item.Key)
                                {
                                    case "CycleTimeAdoption":
                                        CycleTimeAdoption = item.Value;
                                        break;
                                    case "AICycleTime":
                                        AICycleTime = item.Value;
                                        break;
                                    case "MPCycleTime":
                                        MPCycleTime = item.Value;
                                        break;
                                }
                            }

                            #region//更新RFQ AI相關資訊
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.RfqDetail SET
                                    CycleTimeAdoption = @CycleTimeAdoption,
                                    AICycleTime = @AICycleTime,
                                    MPCycleTime = @MPCycleTime,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE RfqDetailId = @RfqDetailId";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                CycleTimeAdoption,
                                AICycleTime,
                                MPCycleTime,
                                LastModifiedDate,
                                LastModifiedBy,
                                RfqDetailId
                            });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                        }


                        foreach (var item in DataList)
                        {
                            int QieId = -1;
                            string Remark = "";
                            string QuoteValue = item.Value;

                            if (item.Key != "Remark")
                            {
                                QieId = Convert.ToInt32(item.Key);
                                if (QuoteValue == "") throw new SystemException("該項目有資料未維護完整不可以執行更新,請重新確認!");
                                #region //判斷項目欄位是否存在
                                sql = @"SELECT TOP 1 1
                                    FROM SCM.QuotationItemElement
                                    WHERE QieId = @QieId";
                                dynamicParameters.Add("QieId", QieId);

                                var result1 = sqlConnection.Query(sql, dynamicParameters);
                                if (result1.Count() <= 0) throw new SystemException("報價單項目資料找不到,請重新確認!");
                                #endregion

                                #region//更新RFQ明細
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE SCM.QuotationItemElement SET
                                    QuoteValue = @QuoteValue,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE QieId = @QieId";
                                dynamicParameters.AddDynamicParams(
                                new
                                {
                                    QuoteValue,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    QieId
                                });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                
                            }
                            else
                            {
                                #region//更新RFQ明細
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE SCM.QuotationItem SET
                                        ColRemark = @QuoteValue,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE QiId = @QiId";
                                dynamicParameters.AddDynamicParams(
                                new
                                {
                                    QuoteValue,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    QiId
                                });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //UpdateQuotationDetailConfirm -- 報價單項目資料確認/反確認 -- Shintokuro 2024.08.13
        public string UpdateQuotationDetailConfirm(int QiId, string Confirm, bool ClassLast, string Data)
        {
            try
            {
                Dictionary<string, string> DataList = new Dictionary<string, string>();
                if (Data != "")
                {
                    DataList = JsonConvert.DeserializeObject<Dictionary<string, string>>(Data);
                }

                int rowsAffected = 0;
                int? nullData = null;
                DateTime? nullDate = null;

                int RfqId = -1;
                string RfqNo = "";
                string Columns = "";
                string UserInfo = "";
                string Charge = "";
                string HtmlColumns = "";
                string SalesUseNo = "";
                bool TestArea = false;
                string ChannelNo = "QuoteChannel";
                string Web = "https://bm.zy-tech.com.tw/";
                string RfqProductTypeName = "";
                string MtlFullName = "";

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認公司資訊
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CompanyNo,x.UserInfo,a.TestArea
                                FROM BAS.Company a 
                                OUTER APPLY(
                                    SELECT x1.UserNo + ' ' + x1.UserName UserInfo FROM BAS.[User] x1 WHERE x1.UserId = @CreateBy
                                ) x
                                WHERE a.CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("CreateBy", CreateBy);


                        var CompanyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (CompanyResult.Count() <= 0) throw new SystemException("公司資料錯誤!!");

                        string CompanyNo = "";
                        foreach (var item in CompanyResult)
                        {
                            CompanyNo = item.CompanyNo;
                            UserInfo = item.UserInfo;
                            if(item.TestArea == "Y")
                            {
                                ChannelNo = "QuoteChannelTest";
                                Web = "http://192.168.20.36:17000/";
                            }
                        }
                        #endregion

                        #region //撈取登入者工號
                        string CreateByUserNo = "";
                        sql = @"SELECT TOP 1 UserNo
                                FROM BAS.[User]
                                WHERE UserId = @CreateBy";
                        dynamicParameters.Add("CreateBy", CreateBy);

                        var resultUser = sqlConnection.Query(sql, dynamicParameters);
                        foreach (var item in resultUser)
                        {
                            CreateByUserNo = item.UserNo;
                        }
                        #endregion

                        #region //判斷報價單是否正確
                        bool ChargeAction = false;
                        int RfqDetailId = -1;

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.RfqDetailId,a.ConfirmStatus,a.Charge,a.HtmlColumns,a.HtmlColumns + '. ' +a.HtmlColName Columns
                                ,c.RfqNo,c.RfqId,d.UserNo SalesUseNo
                                ,b.MtlName,b.PlasticName,b.OutsideDiameter
                                ,e.RfqProductTypeName
                                FROM SCM.QuotationItem a
                                INNER JOIN SCM.RfqDetail b on a.RfqDetailId = b.RfqDetailId
                                INNER JOIN SCM.RequestForQuotation c on b.RfqId = c.RfqId
                                INNER JOIN BAS.[User] d on b.SalesId = d.UserId
                                INNER JOIN SCM.RfqProductType e on e.RfqProTypeId = b.RfqProTypeId
                                WHERE QiId = @QiId";
                        dynamicParameters.Add("QiId", QiId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("報價單資料找不到,請重新確認!");
                        foreach (var item in result)
                        {
                            switch (Confirm)
                            {
                                case "Y":
                                    if (item.ConfirmStatus != "N") throw new SystemException("該項目須處於未確認狀態,才可以執行確認動作!!");

                                    break;
                                case "N":
                                    if (item.ConfirmStatus != "Y") throw new SystemException("該項目須處於已確認狀態,才可以執行反確認動作!!");
                                    break;
                                default:
                                    throw new SystemException("操作狀態異常請重新執行!!");
                                    break;
                            }
                            string ChargeId = item.Charge;
                            foreach (var item1 in ChargeId.Split(','))
                            {
                                if (item1 == CreateBy.ToString() || item1 == CreateByUserNo.ToString())
                                {
                                    ChargeAction = true;
                                    break;
                                }
                            }
                            if (!ChargeAction) throw new SystemException("尚無該項目可異動權限!!請重新確認!!");
                            RfqDetailId = item.RfqDetailId;
                            RfqId = item.RfqId;
                            RfqNo = item.RfqNo;
                            HtmlColumns = item.HtmlColumns;
                            Columns = item.Columns;
                            SalesUseNo = item.SalesUseNo;
                            RfqProductTypeName = item.RfqProductTypeName;
                            string MtlName = item.MtlName;
                            string PlasticName = item.PlasticName;
                            string OutsideDiameter = item.OutsideDiameter;
                            MtlFullName = MtlName + "/" + PlasticName + "/" + OutsideDiameter + "/";

                        }
                        #endregion

                        #region //當成型維護時要判斷研發項目有確認才執行
                        if (Confirm == "Y" && HtmlColumns == "B")
                        {
                            sql = @"SELECT TOP 1 ConfirmStatus
                                        FROM SCM.QuotationItem
                                        WHERE RfqDetailId = @RfqDetailId
                                        AND HtmlColumns = 'A'";
                            dynamicParameters.Add("RfqDetailId", RfqDetailId);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("報價單資料找不到,請重新確認!");
                            foreach (var item in result)
                            {
                                if (item.ConfirmStatus != "Y") throw new SystemException("研發項目尚未確認,不可以執行!!");
                            }
                        }
                        #endregion

                        #region //當研發時反確認須要將成型反確認並清空資料
                        if (Confirm == "N" && HtmlColumns == "A")
                        {
                            sql = @"UPDATE SCM.QuotationItem 
                                    SET ConfirmStatus = 'N',
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE RfqDetailId = @RfqDetailId
                                    AND HtmlColumns = 'B';";
                            dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        RfqDetailId
                                    });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                            sql = @"UPDATE SCM.QuotationItemElement 
                                    SET QuoteValue = '',
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    FROM SCM.QuotationItemElement b
                                    INNER JOIN SCM.QuotationItem a ON a.QiId = b.QiId
                                    WHERE a.RfqDetailId = @RfqDetailId
                                    AND a.HtmlColumns = 'B'
                                    AND b.Flag = 'OEE';";
                            dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        RfqDetailId
                                    });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        }
                        #endregion

                        switch (ClassLast)
                        {
                            case false:
                                #region //判斷報價單是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 QuoteValue
                                        FROM SCM.QuotationItemElement
                                        WHERE QiId = @QiId
                                        AND QuoteValue = ''
                                        ";
                                dynamicParameters.Add("QiId", QiId);

                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() > 0) throw new SystemException("該項目有資料未維護完整不可以執行確認,請重新確認!");
                                #endregion
                                
                                break;
                            case true:
                                #region //判斷項目是否都確認了
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.ConfirmStatus,a.QiId
                                        FROM SCM.QuotationItem a
                                        WHERE a.RfqDetailId = @RfqDetailId";
                                dynamicParameters.Add("RfqDetailId", RfqDetailId);

                                var result1 = sqlConnection.Query(sql, dynamicParameters);
                                if (result1.Count() <= 0) throw new SystemException("報價單項目資料找不到,請重新確認!");
                                foreach (var item1 in result1)
                                {
                                    if (item1.ConfirmStatus != "Y" && item1.QiId != QiId) throw new SystemException("尚有報價單項目未確認,不可以執行報價單確認!!");
                                }
                                #endregion

                                foreach (var item in DataList)
                                {
                                    int QieId = Convert.ToInt32(item.Key);
                                    string QuoteValue = item.Value;
                                    if (QuoteValue == "") throw new SystemException("該項目有資料未維護完整不可以執行確認,請重新確認!");


                                    #region//更新報價單項目確認狀態
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE SCM.QuotationItemElement SET
                                            QuoteValue = @QuoteValue,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE QieId = @QieId";
                                    dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        QuoteValue = Confirm == "Y" ? QuoteValue : "0",
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        QieId
                                    });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }

                                #region//更新RFQ明細
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE SCM.RfqDetail SET
                                        ConfirmCustTime = @ConfirmCustTime,
                                        QuotationStatus = @QuotationStatus,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE RfqDetailId = @RfqDetailId";
                                dynamicParameters.AddDynamicParams(
                                new
                                {
                                    ConfirmCustTime = Confirm == "Y" ? LastModifiedDate : nullDate,
                                    QuotationStatus = Confirm,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    RfqDetailId
                                });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                break;
                        }

                        #region//更新RFQ明細
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.QuotationItem SET
                                        ConfirmUserId = @ConfirmUserId,
                                        ConfirmStatus = @ConfirmStatus,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE QiId = @QiId";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            ConfirmUserId = Confirm == "Y" ? LastModifiedBy : nullData,
                            ConfirmStatus = Confirm,
                            LastModifiedDate,
                            LastModifiedBy,
                            QiId
                        });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        if(Confirm == "Y")
                        {
                            #region //判斷項目是否都完成確認
                            bool finalConfirmStatus = true;
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT DISTINCT ConfirmStatus 
                                FROM SCM.QuotationItem
                                WHERE RfqDetailId = @RfqDetailId";
                            dynamicParameters.Add("RfqDetailId", RfqDetailId);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("資料異常請洽系統開發室!!");
                            foreach (var item in result)
                            {
                                if (item.ConfirmStatus != "Y")
                                {
                                    finalConfirmStatus = false;
                                }
                            }
                            #endregion

                            #region //MAMO通知 
                            string Content = "";
                            string Url = Web + "RequestForQuotation/RfqAssignManagment?UrlRfqId=" + RfqId + "&UrlRfqDetailId=" + RfqDetailId;
                            string Url2 = Web + "RequestForQuotation/QuotationWorkPlatform";
                            List<int> Files = new List<int>();
                            List<string> Tags = new List<string>();

                            #region //頻道撈取 
                            string ChannelId = "";
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.ChannelId
                            FROM MAMO.Channels a
                            WHERE a.ChannelNo = @ChannelNo";
                            dynamicParameters.Add("ChannelNo", ChannelNo);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            foreach (var item in result)
                            {
                                ChannelId = item.ChannelId.ToString();
                            }
                            #endregion

                            #region //通知信格式 
                            if (finalConfirmStatus)
                            {
                                //最終確認通知審核人員
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT x1.UserId,x1.UserNo,x1.UserName
                                    FROM BAS.[User] x1
                                    INNER JOIN BAS.UserRole x2 on x1.UserId = x2.UserId
                                    INNER JOIN BAS.[Role] x3 on x2.RoleId = x3.RoleId
                                    INNER JOIN BAS.RoleFunctionDetail x4 on x2.RoleId = x4.RoleId
                                    INNER JOIN BAS.FunctionDetail x5 on x4.DetailId = x5.DetailId
                                    INNER JOIN BAS.[Function] x6 on x5.FunctionId = x6.FunctionId
                                    WHERE x5.DetailCode = 'quote-review'
                                    AND x6.FunctionCode = 'RfqAssignManagment'
                                    AND x3.CompanyId  = @CompanyId
                                    AND x3.AdminStatus != 'Y'";
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("資料異常請洽系統開發室!!");
                                foreach (var item in result)
                                {
                                    Tags.Add(item.UserNo);
                                }

                                Content = "### 報價單-【請求審核通知】\n" +
                                      "##### RFQ需求單號: " + RfqNo + "\n" +
                                      "##### 產品類別: " + RfqProductTypeName + "\n" +
                                      "##### 品名/塑料品名/外徑大小: " + MtlFullName + "\n" +
                                      "##### 傳送連結: " + Url + "\n" +
                                      "##### 工作平台: " + Url2 + "\n" +
                                      "- 發信時間: " + CreateDate + "\n" +
                                      "- 發信人員: " + UserInfo + "\n";
                            }
                            else
                            {
                                //一般確認通知通知業務
                                Content = "### 報價單-【項目確認通知】\n" +
                                        "##### RFQ需求單號: " + RfqNo + "\n" +
                                        "##### 產品類別: " + RfqProductTypeName + "\n" +
                                        "##### 品名/塑料品名/外徑大小: " + MtlFullName + "\n" +
                                        "##### 報價項目: " + Columns + "\n" +
                                        "##### 傳送連結: " + Url + "\n" +
                                        "##### 工作平台: " + Url2 + "\n" +
                                        "- 發信時間: " + CreateDate + "\n" +
                                        "- 發信人員: " + UserInfo + "\n";

                                Tags.Add(SalesUseNo);

                            }
                            #endregion

                            #region //執行
                            string MamoResult = mamoHelper.SendMessage(CompanyNo, CurrentUser, "Channel", ChannelId, Content, Tags, Files);

                            JObject MamoResultJson = JObject.Parse(MamoResult);
                            if (MamoResultJson["status"].ToString() != "success")
                            {
                                throw new SystemException(MamoResultJson["msg"].ToString());
                            }
                            #endregion
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

        #region //UpdateQuotationFinalConfirm -- 報價單最終確認-- Shintokuro 2024.09.12
        public string UpdateQuotationFinalConfirm(int QiId)
        {
            try
            {
                if (QiId <= 0) throw new SystemException("缺少報價單ID請重新確認!!");

                int rowsAffected = 0;
                bool TestArea = false;
                string ChannelNo = "QuoteChannel";
                string Web = "https://bm.zy-tech.com.tw/";
                string RfqProductTypeName = "";
                string MtlFullName = "";

                int RfqId = -1;
                string RfqNo = "";
                string QuotationNo = "";
                string Columns = "";
                string UserInfo = "";
                string Charge = "";
                string SalesUseNo = "";
                double FinalPrice = 0;
                double GrossMargin = 0;
                string dateNow = CreateDate.ToString("yyyyMMdd");
                string FinalItem = "";


                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認公司資訊
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CompanyNo,x.UserInfo,a.TestArea
                                FROM BAS.Company a 
                                OUTER APPLY(
                                    SELECT x1.UserNo + ' ' + x1.UserName UserInfo FROM BAS.[User] x1 WHERE x1.UserId = @CreateBy
                                ) x
                                WHERE a.CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("CreateBy", CreateBy);


                        var CompanyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (CompanyResult.Count() <= 0) throw new SystemException("公司資料錯誤!!");

                        string CompanyNo = "";
                        foreach (var item in CompanyResult)
                        {
                            CompanyNo = item.CompanyNo;
                            UserInfo = item.UserInfo;
                            if (item.TestArea == "Y")
                            {
                                ChannelNo = "QuoteChannelTest";
                                Web = "http://192.168.20.36:17000/";
                            }
                        }
                        #endregion



                        #region //判斷報價單是否正確
                        bool ChargeAction = false;
                        int RfqDetailId = -1;

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.RfqDetailId,a.ConfirmStatus,a.Charge,a.HtmlColumns + '. ' +a.HtmlColName Columns
                                ,c.RfqNo,c.RfqId,d.UserNo SalesUseNo
                                ,b.MtlName,b.PlasticName,b.OutsideDiameter,b.RfqProTypeId
                                ,e.RfqProductTypeName,x.FlagSetting
                                FROM SCM.QuotationItem a
                                INNER JOIN SCM.RfqDetail b on a.RfqDetailId = b.RfqDetailId
                                INNER JOIN SCM.RequestForQuotation c on b.RfqId = c.RfqId
                                INNER JOIN BAS.[User] d on b.SalesId = d.UserId
                                INNER JOIN SCM.RfqProductType e on e.RfqProTypeId = b.RfqProTypeId
                                OUTER APPLY(
                                    SELECT '(' + STRING_AGG('''' + x1.ItemElementNo + '''', ', ') + ')' AS FlagSetting
                                    FROM SCM.QuotationItemElement x1
                                    WHERE x1.QiId = a.QiId
                                    AND x1.Flag = 'FR'
                                ) x
                                WHERE QiId = @QiId";
                        dynamicParameters.Add("QiId", QiId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("報價單資料找不到,請重新確認!");
                        foreach (var item in result)
                        {
                            if (item.ConfirmStatus != "Y") throw new SystemException("該項目須處於已確認狀態,才可以執行經管中心確認動作!!");
                            string ChargeId = item.Charge;
                            
                            FinalItem = item.FlagSetting;
                            
                            foreach (var item1 in ChargeId.Split(','))
                            {
                                if (item1 == CreateBy.ToString())
                                {
                                    ChargeAction = true;
                                    break;
                                }
                            }

                            //if (!ChargeAction) throw new SystemException("尚無該項目可異動權限!!請重新確認!!");
                            RfqDetailId = item.RfqDetailId;
                            RfqId = item.RfqId;
                            RfqNo = item.RfqNo;
                            Columns = item.Columns;
                            SalesUseNo = item.SalesUseNo;
                            RfqProductTypeName = item.RfqProductTypeName;
                            string MtlName = item.MtlName;
                            string PlasticName = item.PlasticName;
                            string OutsideDiameter = item.OutsideDiameter;
                            MtlFullName = MtlName + "/" + PlasticName + "/" + OutsideDiameter + "/";
                        }
                        #endregion

                        #region //撈取最終報價單 毛利率和金額
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a1.Sort,a1.ItemElementName,a1.QuoteValue 
                                FROM SCM.QuotationItem a
                                INNER JOIN SCM.QuotationItemElement a1 on a.QiId = a1.QiId
                                WHERE a.RfqDetailId = @RfqDetailId
                                AND a1.ItemElementNo in "+ FinalItem + @"
                                ORDER BY a1.Sort ASC";
                        dynamicParameters.Add("RfqDetailId", RfqDetailId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("報價單資料找不到,請重新確認!");
                        

                        GrossMargin = Convert.ToDouble((result.ToList())[0].QuoteValue);
                        FinalPrice = Convert.ToDouble((result.ToList())[1].QuoteValue);
                        #endregion

                        #region //撈取報價單單號最大值
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 QuotationNo +1 MaxQuotationNo ,QuotationNo
                                FROM SCM.RfqDetail 
                                WHERE QuotationNo LIKE @dateNow
                                ORDER BY QuotationNo DESC ";
                        dynamicParameters.Add("dateNow", dateNow + "__");

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0)
                        {
                            foreach (var item in result)
                            {
                                if (item.QuotationNo.Substring(10 - 2) == "99") throw new SystemException("當日報價單流水號已達99,請洽系統開發和經管中心詢問!");
                                QuotationNo = item.MaxQuotationNo.ToString();
                            }
                        }
                        else
                        {
                            QuotationNo = dateNow + "01";
                        }
                        #endregion


                        #region //新增議價紀錄資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.QuotationFinalPrice (RfqDetailId, FinalPrice, GrossMargin, ConfirmStatus, Remark
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.QfpId
                                VALUES (@RfqDetailId, @FinalPrice, @GrossMargin, @ConfirmStatus, @Remark
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RfqDetailId,
                                FinalPrice,
                                GrossMargin, //取4位數
                                ConfirmStatus = "H",
                                Remark = "",
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += insertResult.Count();
                        #endregion

                        #region//更新RFQ明細
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.RfqDetail SET
                                QuotationNo = @QuotationNo,
                                QuotationStatus = @QuotationStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RfqDetailId = @RfqDetailId";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            QuotationNo,
                            QuotationStatus = "F",
                            LastModifiedDate,
                            LastModifiedBy,
                            RfqDetailId
                        });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion
                        
                        #region //MAMO通知
                        string ChannelId = "";
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ChannelId
                            FROM MAMO.Channels a
                            WHERE a.ChannelNo = @ChannelNo";
                        dynamicParameters.Add("ChannelNo", ChannelNo);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        foreach (var item in result)
                        {
                            ChannelId = item.ChannelId.ToString();
                        }

                        string Content = "";
                        string Url = Web + "RequestForQuotation/RfqAssignManagment?UrlRfqId=" + RfqId + "&UrlRfqDetailId=" + RfqDetailId;
                        string Url2 = Web + "RequestForQuotation/QuotationWorkPlatform";

                        Content = "### 報價單-【核准通知】\n" +
                                    "##### RFQ需求單號: " + RfqNo + "\n" +
                                    "##### 產品類別: " + RfqProductTypeName + "\n" +
                                    "##### 品名/塑料品名/外徑大小: " + MtlFullName + "\n" +
                                    "##### 傳送連結: " + Url + "\n" +
                                    "##### 工作平台: " + Url2 + "\n" +
                                    "- 發信時間: " + CreateDate + "\n" +
                                    "- 發信人員: " + UserInfo + "\n";

                        #region //取得標記USER資料
                        List<string> Tags = new List<string>();
                        Tags.Add(SalesUseNo);
                        #endregion

                        List<int> Files = new List<int>();

                        #region //執行
                        string MamoResult = mamoHelper.SendMessage(CompanyNo, CurrentUser, "Channel", ChannelId, Content, Tags, Files);

                        JObject MamoResultJson = JObject.Parse(MamoResult);
                        if (MamoResultJson["status"].ToString() != "success")
                        {
                            throw new SystemException(MamoResultJson["msg"].ToString());
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

        #region //UpdateQuotationBargainConfirm -- 報價單議價確認-- Shintokuro 2024.09.13
        public string UpdateQuotationBargainConfirm(int RfqDetailId)
        {
            try
            {
                if (RfqDetailId <= 0) throw new SystemException("【報價單ID】不能為空!");
                int RfqId = -1;
                int rowsAffected = 0;
                bool TestArea = false;
                string ChannelNo = "QuoteChannel";
                string Web = "https://bm.zy-tech.com.tw/";
                string RfqProductTypeName = "";
                string MtlFullName = "";

                string RfqNo = "";
                string UserInfo = "";
                List<string> Tags = new List<string>();

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認公司資訊
                         dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CompanyNo,x.UserInfo,a.TestArea
                                FROM BAS.Company a 
                                OUTER APPLY(
                                    SELECT x1.UserNo + ' ' + x1.UserName UserInfo FROM BAS.[User] x1 WHERE x1.UserId = @CreateBy
                                ) x
                                WHERE a.CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("CreateBy", CreateBy);
                        var CompanyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (CompanyResult.Count() <= 0) throw new SystemException("公司資料錯誤!!");

                        string CompanyNo = "";
                        foreach (var item in CompanyResult)
                        {
                            CompanyNo = item.CompanyNo;
                            UserInfo = item.UserInfo;
                            if (item.TestArea == "Y")
                            {
                                ChannelNo = "QuoteChannelTest";
                                Web = "http://192.168.20.36:17000/";
                            }
                        }
                        #endregion

                        #region //判斷報價單是否正確
                        int QfpId = -1;

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.RfqId,a.RfqNo,a1.RfqDetailId,ISNULL(x.QfpId, -1) QfpId
                                ,a1.MtlName,a1.PlasticName,a1.OutsideDiameter
                                ,b.RfqProductTypeName
                                FROM SCM.RequestForQuotation a
                                INNER JOIN SCM.RfqDetail a1 on a.RfqId = a1.RfqId
                                INNER JOIN SCM.RfqProductType b on b.RfqProTypeId = a1.RfqProTypeId
                                OUTER APPLY(
                                    SELECT x1.QfpId 
                                    FROM SCM.QuotationFinalPrice x1
                                    INNER JOIN SCM.RfqDetail x2 on x1.RfqDetailId = x2.RfqDetailId
                                    WHERE x1.ConfirmStatus = 'N'
                                ) x
                                WHERE a1.RfqDetailId = @RfqDetailId";
                        dynamicParameters.Add("RfqDetailId", RfqDetailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("報價單資料找不到,請重新確認!");
                        foreach (var item in result)
                        {
                            if(item.QfpId <= 0) throw new SystemException("找不到報價單議價需要確認的資料,或是當前議價已確認待經管中心確認中,請重新確認!");
                            RfqDetailId = item.RfqDetailId;
                            RfqId = item.RfqId;
                            RfqNo = item.RfqNo;
                            QfpId = item.QfpId;
                            RfqProductTypeName = item.RfqProductTypeName;
                            string MtlName = item.MtlName;
                            string PlasticName = item.PlasticName;
                            string OutsideDiameter = item.OutsideDiameter;
                            MtlFullName = MtlName + "/" + PlasticName + "/" + OutsideDiameter + "/";
                        }
                        #endregion

                        #region //撈取核價人
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT x1.UserName,x1.UserNo,x1.UserId
                                FROM BAS.[User] x1
                                INNER JOIN BAS.UserRole x2 on x1.UserId = x2.UserId
                                INNER JOIN BAS.[Role] x3 on x2.RoleId = x3.RoleId
                                INNER JOIN BAS.RoleFunctionDetail x4 on x2.RoleId = x4.RoleId
                                INNER JOIN BAS.FunctionDetail x5 on x4.DetailId = x5.DetailId
                                INNER JOIN BAS.[Function] x6 on x5.FunctionId = x6.FunctionId
                                WHERE x5.DetailCode = 'quote-review'
                                AND x6.FunctionCode = 'RfqAssignManagment'
                                AND x3.CompanyId  = @CompanyId
                                AND x3.AdminStatus != 'Y'";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("找不到核價權限人,請重新確認!");
                        foreach(var item in result)
                        {
                            #region //取得標記USER資料
                            Tags.Add(item.UserNo);
                            #endregion
                        }
                        #endregion

                        #region//更新議價確認
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.QuotationFinalPrice SET
                                ConfirmStatus = @ConfirmStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE QfpId = @QfpId";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            ConfirmStatus = "Y",
                            LastModifiedDate,
                            LastModifiedBy,
                            QfpId
                        });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        
                        #region //MAMO通知
                        string ChannelId = "";
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ChannelId
                            FROM MAMO.Channels a
                            WHERE a.ChannelNo = @ChannelNo";
                        dynamicParameters.Add("ChannelNo", ChannelNo);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        foreach (var item in result)
                        {
                            ChannelId = item.ChannelId.ToString();
                        }

                        string Content = "";
                        string Url = Web + "RequestForQuotation/RfqAssignManagment?UrlRfqId=" + RfqId + "&UrlRfqDetailId=" + RfqDetailId;
                        string Url2 = Web + "RequestForQuotation/QuotationWorkPlatform";

                        Content = "### 報價單-【請求審核議價通知】\n" +
                                  "##### RFQ需求單號: " + RfqNo + "\n" +
                                  "##### 產品類別: " + RfqProductTypeName + "\n" +
                                  "##### 品名/塑料品名/外徑大小: " + MtlFullName + "\n" +
                                  "##### 傳送連結: " + Url + "\n" +
                                  "##### 工作平台: " + Url2 + "\n" +
                                  "- 發信時間: " + CreateDate + "\n" +
                                  "- 發信人員: " + UserInfo + "\n";
                        List<int> Files = new List<int>();

                        #region //執行
                        string MamoResult = mamoHelper.SendMessage(CompanyNo, CurrentUser, "Channel", ChannelId, Content, Tags, Files);

                        JObject MamoResultJson = JObject.Parse(MamoResult);
                        if (MamoResultJson["status"].ToString() != "success")
                        {
                            throw new SystemException(MamoResultJson["msg"].ToString());
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

        #region //UpdateQuotationBargainReview -- 報價單議價審核-- Shintokuro 2024.09.13
        public string UpdateQuotationBargainReview(int RfqDetailId)
        {
            try
            {
                if (RfqDetailId <= 0) throw new SystemException("【報價單ID】不能為空!");
                int RfqId = -1;

                int rowsAffected = 0;
                bool TestArea = false;
                string ChannelNo = "QuoteChannel";
                string Web = "https://bm.zy-tech.com.tw/";
                string RfqProductTypeName = "";
                string MtlFullName = "";
                string FinalItem = "";

                string RfqNo = "";
                string UserInfo = "";
                List<string> Tags = new List<string>();

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認公司資訊
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CompanyNo,x.UserInfo,a.TestArea
                                FROM BAS.Company a 
                                OUTER APPLY(
                                    SELECT x1.UserNo + ' ' + x1.UserName UserInfo FROM BAS.[User] x1 WHERE x1.UserId = @CreateBy
                                ) x
                                WHERE a.CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("CreateBy", CreateBy);
                        var CompanyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (CompanyResult.Count() <= 0) throw new SystemException("公司資料錯誤!!");

                        string CompanyNo = "";
                        foreach (var item in CompanyResult)
                        {
                            CompanyNo = item.CompanyNo;
                            UserInfo = item.UserInfo;
                            if (item.TestArea == "Y")
                            {
                                ChannelNo = "QuoteChannelTest";
                                Web = "http://192.168.20.36:17000/";
                            }
                        }
                        #endregion

                        #region //判斷報價單是否正確
                        int QfpId = -1;
                        double FinalPrice = -1;
                        double GrossMargin = -1;

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.RfqNo,a1.RfqDetailId,b.UserNo
                                ,ISNULL(x.QfpId, -1) QfpId,ISNULL(x.FinalPrice, 0) FinalPrice,ISNULL(x.GrossMargin, 0) GrossMargin
                                ,a1.MtlName,a1.PlasticName,a1.OutsideDiameter,a1.RfqProTypeId
                                ,c.RfqProductTypeName,y.FlagSetting
                                FROM SCM.RequestForQuotation a
                                INNER JOIN SCM.RfqDetail a1 on a.RfqId = a1.RfqId
                                INNER JOIN BAS.[User] b on a1.SalesId = b.UserId
                                INNER JOIN SCM.RfqProductType c on c.RfqProTypeId = a1.RfqProTypeId
                                OUTER APPLY(
                                    SELECT x1.QfpId ,x1.FinalPrice, x1.GrossMargin
                                    FROM SCM.QuotationFinalPrice x1
                                    INNER JOIN SCM.RfqDetail x2 on x1.RfqDetailId = x2.RfqDetailId
                                    WHERE x1.ConfirmStatus = 'Y'
                                ) x
                                OUTER APPLY(
                                    SELECT '(' + STRING_AGG('''' + x1.ItemElementNo + '''', ', ') + ')' AS FlagSetting
                                    FROM SCM.QuotationItemElement x1
                                    INNER JOIN SCM.QuotationItem x2 on x1.QiId = x2.QiId
                                    WHERE x2.RfqDetailId = a1.RfqDetailId
                                    AND x1.Flag = 'FR'
                                ) y
                                WHERE a1.RfqDetailId = @RfqDetailId
                                AND a1.DocType = 'E'";
                        dynamicParameters.Add("RfqDetailId", RfqDetailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("報價單資料找不到,請重新確認!");
                        foreach (var item in result)
                        {
                            if (item.QfpId <= 0) throw new SystemException("報價單議價資料找不到,請重新確認!");
                            RfqDetailId = item.RfqDetailId;
                            RfqNo = item.RfqNo;
                            QfpId = item.QfpId;
                            FinalPrice = item.FinalPrice;
                            GrossMargin = item.GrossMargin;
                            RfqProductTypeName = item.RfqProductTypeName;
                            Tags.Add(item.UserNo);
                            string MtlName = item.MtlName;
                            string PlasticName = item.PlasticName;
                            string OutsideDiameter = item.OutsideDiameter;
                            MtlFullName = MtlName + "/" + PlasticName + "/" + OutsideDiameter + "/";
                            FinalItem = item.FlagSetting;
                        }
                        #endregion

                        #region //判斷報價單是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ROW_NUMBER() OVER (ORDER BY a.Sort) Sort,a.QieId,a.ItemElementNo
                                FROM SCM.QuotationItemElement a
                                INNER JOIN SCM.QuotationItem a1 on a.QiId = a1.QiId
                                WHERE a1.RfqDetailId = @RfqDetailId
                                AND a.ItemElementNo in " + FinalItem + @"
                                ORDER BY a.Sort ASC";
                        dynamicParameters.Add("RfqDetailId", RfqDetailId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        foreach (var item in result)
                        {
                            string QuoteValue = "";
                            string Sort = item.Sort.ToString();
                            switch (Sort)
                            {
                                case "1":
                                    QuoteValue = GrossMargin.ToString();
                                    break;
                                case "2":
                                    QuoteValue = FinalPrice.ToString();
                                    break;
                                case "3":
                                    QuoteValue = Math.Round((FinalPrice * 1.13), 3, MidpointRounding.AwayFromZero).ToString();
                                    break;
                            }
                            #region//更新議價確認
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.QuotationItemElement SET
                                    QuoteValue = @QuoteValue,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE QieId = @QieId";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                QuoteValue,
                                LastModifiedDate,
                                LastModifiedBy,
                                QieId = item.QieId
                            });
                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }
                        #endregion


                        #region//更新議價確認
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.QuotationFinalPrice SET
                                ConfirmStatus = @ConfirmStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE QfpId = @QfpId";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            ConfirmStatus = "H",
                            LastModifiedDate,
                            LastModifiedBy,
                            QfpId
                        });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        
                        #region //MAMO通知
                        string ChannelId = "";
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ChannelId
                            FROM MAMO.Channels a
                            WHERE a.ChannelNo = @ChannelNo";
                        dynamicParameters.Add("ChannelNo", ChannelNo);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        foreach (var item in result)
                        {
                            ChannelId = item.ChannelId.ToString();
                        }

                        string Content = "";
                        string Url = Web + "RequestForQuotation/RfqAssignManagment?UrlRfqId=" + RfqId + "&UrlRfqDetailId=" + RfqDetailId;
                        string Url2 = Web + "RequestForQuotation/QuotationWorkPlatform";

                        Content = "### 報價單-【議價核准通知】\n" +
                                    "##### RFQ需求單號: " + RfqNo + "\n" +
                                    "##### 產品類別: " + RfqProductTypeName + "\n" +
                                    "##### 品名/塑料品名/外徑大小: " + MtlFullName + "\n" +
                                    "##### 傳送連結: " + Url + "\n" +
                                    "##### 工作平台: " + Url2 + "\n" +
                                    "- 發信時間: " + CreateDate + "\n" +
                                    "- 發信人員: " + UserInfo + "\n";
                        List<int> Files = new List<int>();

                        #region //執行
                        string MamoResult = mamoHelper.SendMessage(CompanyNo, CurrentUser, "Channel", ChannelId, Content, Tags, Files);

                        JObject MamoResultJson = JObject.Parse(MamoResult);
                        if (MamoResultJson["status"].ToString() != "success")
                        {
                            throw new SystemException(MamoResultJson["msg"].ToString());
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
        #region //DeleteRequestForQuotation -- 刪除RFQ單頭資料 -- Yi 2023-10-12
        public string DeleteRequestForQuotation(int RfqId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();

                        #region //判斷RFQ單頭資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.RequestForQuotation a
                                WHERE a.RfqId = @RfqId";
                        dynamicParameters.Add("RfqId", RfqId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("RFQ單頭資料錯誤!");
                        #endregion

                        #region //判斷RFQ單身資料是否正確
                        sql = @"SELECT TOP 1 RfqDetailId
                                FROM SCM.RfqDetail
                                WHERE RfqId = @RfqId";
                        dynamicParameters.Add("RfqId", RfqId);

                        var resultRfqDetail = sqlConnection.Query(sql, dynamicParameters);
                        if (resultRfqDetail.Count() > 0) throw new SystemException("RFQ單身資料已存在，不可刪除!");

                        int rfqDetailId = -1;
                        foreach (var item in resultRfqDetail)
                        {
                            rfqDetailId = item.RfqDetailId;
                        }
                        #endregion

                        #region //判斷RFQ報價資訊是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.RfqLineSolution
                                WHERE RfqDetailId = @RfqDetailId";
                        dynamicParameters.Add("RfqDetailId", rfqDetailId);

                        var resultRfqLineSolution = sqlConnection.Query(sql, dynamicParameters);
                        if (resultRfqLineSolution.Count() > 0) throw new SystemException("RFQ報價資訊已存在，不可刪除!");
                        #endregion

                        int rowsAffected = 0;
                        #region //先刪除關聯Table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM SCM.RfqPackage a
                                WHERE a.RfqDetailId IN (
                                            SELECT aa.RfqDetailId 
                                            FROM SCM.RfqDetail aa
                                            INNER JOIN SCM.RequestForQuotation ab ON aa.RfqId = ab.RfqId
                                            WHERE aa.RfqDetailId = a.RfqDetailId
                                            AND aa.RfqId = @RfqId
                                        )";
                        dynamicParameters.Add("RfqDetailId", rfqDetailId);
                        dynamicParameters.Add("RfqId", RfqId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE FROM SCM.RfqDetail
                                WHERE RfqId = @RfqId";
                        dynamicParameters.Add("RfqId", RfqId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.RequestForQuotation
                                WHERE RfqId = @RfqId";
                        dynamicParameters.Add("RfqId", RfqId);

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

        #region //DeleteRfqDetail -- 刪除RFQ單身資料 -- Yi 2023.10.13
        public string DeleteRfqDetail(int RfqDetailId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;

                        #region //判斷RFQ資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.RequestForQuotation a
                                INNER JOIN SCM.RfqDetail b ON a.RfqId = b.RfqId
                                WHERE b.RfqDetailId = @RfqDetailId";
                        dynamicParameters.Add("RfqDetailId", RfqDetailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("RFQ單身資料錯誤!");
                        #endregion

                        //#region //判斷RFQ報價資訊是否存在
                        //dynamicParameters = new DynamicParameters();
                        //sql = @"SELECT TOP 1 1
                        //        FROM SCM.RfqLineSolution
                        //        WHERE RfqDetailId = @RfqDetailId";
                        //dynamicParameters.Add("RfqDetailId", RfqDetailId);

                        //var resultRfqLineSolution = sqlConnection.Query(sql, dynamicParameters);
                        //if (resultRfqLineSolution.Count() > 0) throw new SystemException("RFQ【報價級距】已存在，不可刪除；若要刪除單身，請先刪除報價級距!");
                        //#endregion

                        #region //判斷報價單是否開立
                        sql = @"SELECT a.Status
                                FROM SCM.RfqDetail a
                                INNER JOIN SCM.QuotationItem b ON a.RfqDetailId = b.RfqDetailId
                                WHERE b.RfqDetailId = @RfqDetailId
                                ";
                        dynamicParameters.Add("RfqDetailId", RfqDetailId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0)
                        {
                            foreach(var item in result)
                            {
                                //if(item.Status == "Y") throw new SystemException("報價單已有項目確認不可以刪除!!");
                            }
                            #region //刪除報價單
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE a
                                    FROM SCM.QuotationItemElement a
                                    INNER JOIN SCM.QuotationItem b ON a.QiId = b.QiId
                                    WHERE b.RfqDetailId = @RfqDetailId
                                    ";
                            dynamicParameters.Add("RfqDetailId", RfqDetailId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE SCM.QuotationItem WHERE RfqDetailId = @RfqDetailId";
                            dynamicParameters.Add("RfqDetailId", RfqDetailId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }
                        #endregion

                        #region //刪除關聯table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM SCM.RfqLineSolution a
                                WHERE a.RfqDetailId IN (
                                            SELECT aa.RfqDetailId 
                                            FROM SCM.RfqDetail aa
                                            INNER JOIN SCM.RequestForQuotation ab ON aa.RfqId = ab.RfqId
                                            WHERE aa.RfqDetailId = a.RfqDetailId
                                        )";
                        dynamicParameters.Add("RfqDetailId", RfqDetailId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM SCM.RfqPackage a
                                WHERE a.RfqDetailId IN (
                                            SELECT aa.RfqDetailId 
                                            FROM SCM.RfqDetail aa
                                            INNER JOIN SCM.RequestForQuotation ab ON aa.RfqId = ab.RfqId
                                            WHERE aa.RfqDetailId = a.RfqDetailId
                                        )";
                        dynamicParameters.Add("RfqDetailId", RfqDetailId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE FROM SCM.RfqDetail
                                WHERE RfqDetailId = @RfqDetailId";
                        dynamicParameters.Add("RfqDetailId", RfqDetailId);

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
        #endregion

        #region //API
        #region //NotifyQuotationQuestion -- 報價項目異議通知 -- Shintokuro 2024.08.21
        public string NotifyQuotationQuestion(int QiId, string Remark)
        {
            try
            {
                int RfqId = -1;
                int RfqDetailId = -1;
                string RfqNo = "";
                string Columns = "";
                string UserInfo = "";
                string Charge = "";
                int rowsAffected = 0;
                bool TestArea = false;
                string ChannelNo = "QuoteChannel";
                string Web = "https://bm.zy-tech.com.tw/";
                string RfqProductTypeName = "";
                string MtlFullName = "";

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認公司資訊
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CompanyNo,x.UserInfo,a.TestArea
                                FROM BAS.Company a 
                                OUTER APPLY(
                                    SELECT x1.UserNo + ' ' + x1.UserName UserInfo FROM BAS.[User] x1 WHERE x1.UserId = @CreateBy
                                ) x
                                WHERE a.CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("CreateBy", CreateBy);

                        var CompanyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (CompanyResult.Count() <= 0) throw new SystemException("公司資料錯誤!!");

                        string CompanyNo = "";
                        foreach (var item in CompanyResult)
                        {
                            CompanyNo = item.CompanyNo;
                            UserInfo = item.UserInfo;
                            if (item.TestArea == "Y")
                            {
                                ChannelNo = "QuoteChannelTest";
                                Web = "http://192.168.20.36:17000/";
                            }
                        }
                        #endregion

                        #region //判斷報價單是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.HtmlColumns + '. ' +a.HtmlColName Columns,a.Charge,
                                c.RfqNo,c.RfqId,b.RfqDetailId
                                ,b.MtlName,b.PlasticName,b.OutsideDiameter
                                ,e.RfqProductTypeName
                                FROM SCM.QuotationItem a
                                INNER JOIN SCM.RfqDetail b on a.RfqDetailId = b.RfqDetailId
                                INNER JOIN SCM.RequestForQuotation c on b.RfqId = c.RfqId
                                INNER JOIN SCM.RfqProductType e on e.RfqProTypeId = b.RfqProTypeId
                                WHERE a.QiId = @QiId";
                        dynamicParameters.Add("QiId", QiId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("報價單項目資料找不到,請重新確認!");
                        foreach (var item in result)
                        {
                            RfqId = item.RfqId;
                            RfqDetailId = item.RfqDetailId;
                            RfqNo = item.RfqNo;
                            Columns = item.Columns;
                            Charge = item.Charge;
                            RfqProductTypeName = item.RfqProductTypeName;
                            string MtlName = item.MtlName;
                            string PlasticName = item.PlasticName;
                            string OutsideDiameter = item.OutsideDiameter;
                            MtlFullName = MtlName + "/" + PlasticName + "/" + OutsideDiameter + "/";
                        }
                        #endregion

                        #region //MAMO通知

                        #region //頻道撈取
                        string ChannelId = "";
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ChannelId
                                FROM MAMO.Channels a
                                WHERE a.ChannelNo = @ChannelNo";
                        dynamicParameters.Add("ChannelNo", ChannelNo);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        foreach (var item in result)
                        {
                            ChannelId = item.ChannelId.ToString();
                        }
                        #endregion

                        string Url = Web + "RequestForQuotation/RfqAssignManagment?UrlRfqId=" + RfqId + "&UrlRfqDetailId=" + RfqDetailId;
                        string Url2 = Web + "RequestForQuotation/QuotationWorkPlatform";

                        string Content = "";
                        Content = "### 報價單-【項目異議通知】\n" +
                                  "##### RFQ需求單號: " + RfqNo + "\n" +
                                  "##### 產品類別: " + RfqProductTypeName + "\n" +
                                  "##### 品名/塑料品名/外徑大小: " + MtlFullName + "\n" +
                                  "##### 報價項目: " + Columns + "\n" +
                                  "##### 異議說明: " + Remark + "\n" +
                                  "##### 傳送連結: " + Url + "\n" +
                                  "##### 工作平台: " + Url2 + "\n" +
                                  "- 發信時間: " + CreateDate + "\n" +
                                  "- 發信人員: " + UserInfo + "\n";

                        #region //取得標記USER資料(原送測人員部門)
                        List<string> Tags = new List<string>();
                        foreach (var UserNo in Charge.Split(','))
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.UserName, a.UserNo
                                    FROM BAS.[User] a
                                    WHERE a.UserNo = @UserNo OR TRY_CAST(a.UserId AS VARCHAR) = @UserNo";
                            dynamicParameters.Add("UserNo", UserNo);

                            var UserResult = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item in UserResult)
                            {
                                Tags.Add(item.UserNo);
                            }
                        }

                        #endregion

                        List<int> Files = new List<int>();

                        #region //執行
                        string MamoResult = mamoHelper.SendMessage(CompanyNo, CurrentUser, "Channel", ChannelId, Content, Tags, Files);

                        JObject MamoResultJson = JObject.Parse(MamoResult);
                        if (MamoResultJson["status"].ToString() != "success")
                        {
                            throw new SystemException(MamoResultJson["msg"].ToString());
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
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
