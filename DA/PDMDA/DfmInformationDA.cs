using Dapper;
using Helpers;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Transactions;
using System.Web;
using System.Threading;


namespace PDMDA
{
    public class DfmInformationDA
    {
        public string MainConnectionStrings = "";
        public string ErpConnectionStrings = "";
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

        public DfmInformationDA()
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

        #region//Get
        #region //GetDfmItemUserAuthority -- DFM使用者權限獲取 -- Shintokuro 2023.09.01
        public string GetDfmItemUserAuthority(int DfmId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    List<int> DfmItemCategoryList = new List<int>();
                    List<string> DfmItemCategoryData = new List<string>();

                    #region //判斷DfmId是否存在
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.DfmItemCategoryId
                            FROM PDM.DfmQuotationItem a
                            WHERE a.DfmId = @DfmId
                            ";
                    dynamicParameters.Add("DfmId", DfmId);
                    var result = sqlConnection.Query(sql, dynamicParameters);
                    foreach(var item in result)
                    {
                        DfmItemCategoryList.Add(item.DfmItemCategoryId);
                    }
                    #endregion

                    #region //取得品DFM單身
                    dynamicParameters = new DynamicParameters();
                    foreach(var item in DfmItemCategoryList) {
                        sql = @"SELECT a.SubProdId,a.SubProdType
                                FROM EIP.DocNotify a
                                INNER JOIN EIP.NotifyUser b on a.RoleId = b.RoleId
                                INNER JOIN EIP.NotifyRole c on b.RoleId = c.RoleId
                                WHERE a.DocType = 'DFM'
                                AND b.UserId = @UserId
                                AND a.SubProdId = @SubProdId
                            ";
                        dynamicParameters.Add("SubProdId", item);
                        dynamicParameters.Add("UserId", CurrentUser);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0)
                        {
                            foreach (var item2 in result) {
                                DfmItemCategoryData.Add(item2.SubProdId+","+item2.SubProdType);
                            }
                        }
                    }
                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = DfmItemCategoryData
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


        #region //GetDfmItemSimple -- 取得品DFM項目(下拉用) -- Shintokuro 2023.07.10
        public string GetDfmItemSimple(int DfmItemId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    #region //取得品DFM單身
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.DfmItemId, a.DfmItemNo, a.DfmItemName,
                            (a.DfmItemName) DfmItemFull
                            FROM PDM.DfmItem a
                            WHERE a.CompanyId = @CompanyId
                            AND a.Status = 'A'
                            ";
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

        #region//GetDesignForManufacturing --取得DFM 單頭
        public string GetDesignForManufacturing(int DfmId, string DfmNo, string RfqNo, string DifficultyLevel, int RfqDetailId
            , int ModeId, int RdUserId, string StartCreateDate, string EndCreateDate, string Status, string RFQDocType, string AssemblyName
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.DfmId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.DfmNo, a.RfqDetailId, FORMAT(a.DfmDate, 'yyyy-MM-dd') DfmDate, a.DifficultyLevel, a.RdUserId,a.Version, a.Status, a.DfmProcessStatus
                          ,c.RfqId,c.RfqNo,c.AssemblyName,d.ProductUseName, h.RfqProductClassName,g.RfqProductTypeName
                         ,b.MtlName,b.CustProdDigram,b.PlannedOpeningDate,b.PrototypeQty,b.MassProductionDemand,b.KickOffType,b.PlasticName,b.OutsideDiameter
                         ,b.ProdLifeCycleStart,b.ProdLifeCycleEnd,b.LifeCycleQty,b.DemandDate,b.CoatingFlag,b.DocType
                         ,h.RfqProductClassName
                         ,j.UserNo,j.UserName
                         , a1.UserNo,a1.UserName,a1.Gender
                         ";
                    sqlQuery.mainTables =
                        @"   FROM PDM.DesignForManufacturing a
                             LEFT JOIN SCM.RfqDetail b ON a.RfqDetailId=b.RfqDetailId
                             LEFT JOIN SCM.RequestForQuotation c ON b.RfqId=c.RfqId
                             LEFT JOIN SCM.ProductUse d ON c.ProductUseId =d.ProductUseId
                             LEFT JOIN EIP.Member e ON e.MemberId=c.MemberId
                             LEFT JOIN EIP.MemberOrganization f ON e.MemberId=f.MemberId
                             LEFT JOIN SCM.RfqProductType g ON b.RfqProTypeId=g.RfqProTypeId
                             LEFT JOIN SCM.RfqProductClass h ON g.RfqProClassId=h.RfqProClassId
                             LEFT JOIN BAS.[File] i ON b.CustProdDigram=i.FileId
                             LEFT JOIN BAS.[User] j ON b.SalesId=j.UserId
                             LEFT JOIN BAS.[File] k ON b.AdditionalFile=k.FileId
                             LEFT JOIN BAS.[User] a1 ON a.RdUserId=a1.UserId
                        ";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND b.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RfqDetailId", @" AND b.RfqDetailId= @RfqDetailId", RfqDetailId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DfmId", @" AND a.DfmId = @DfmId", DfmId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DfmNo", @" AND a.DfmNo LIKE '%' + @DfmNo + '%'", DfmNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RfqNo", @" AND c.RfqNo LIKE '%' + @RfqNo + '%'", RfqNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "AssemblyName", @" AND c.AssemblyName LIKE '%' + @AssemblyName + '%'", AssemblyName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RdUserId", @" AND a.RdUserId = @RdUserId", RdUserId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartCreateDate", @" AND a.DfmDate >= @StartCreateDate ", StartCreateDate.Length > 0 ? Convert.ToDateTime(StartCreateDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndCreateDate", @" AND a.DfmDate <= @EndCreateDate ", EndCreateDate.Length > 0 ? Convert.ToDateTime(EndCreateDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    if (DifficultyLevel.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DifficultyLevel", @" AND a.DifficultyLevel IN @DifficultyLevel", DifficultyLevel.Split(','));
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    if (RFQDocType.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RFQDocType", @" AND b.DocType IN @RFQDocType", RFQDocType.Split(','));

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.DfmId DESC";
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

        #region //GetDfmDetail -- 取得DFM單身 -- Shintokuro 2023.07.07
        public string GetDfmDetail(int DfmId, string Version)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    if(Version.Length <= 0)
                    {
                        #region //取得DFM單身
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.DfmDetailId,a.DfmId, a.DfmItemId, a.Standard, a.FinalSpec, a.CustomerSpec, a.SuggestSpec
                            , a.RdRemark, a.CustomerFeedback
                            FROM PDM.DfmDetail a
                            WHERE a.DfmId = @DfmId
                            ";
                        dynamicParameters.Add("DfmId", DfmId);
                        #endregion

                    }
                    else
                    {
                        #region //取得DFM歷史單身
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.DfmDetailLogId DfmDetailId,a.DfmId,a.[Version], a.DfmItemId, a.Standard, a.FinalSpec, a.CustomerSpec, a.SuggestSpec
                            , a.RdRemark, a.CustomerFeedback
                            FROM PDM.DfmDetailLog a
                            WHERE a.DfmId = @DfmId
                            AND a.Version = @Version
                            ";
                        dynamicParameters.Add("DfmId", DfmId);
                        dynamicParameters.Add("Version", Version);
                        #endregion
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

        #region //GetDetailMaxNum -- 取得DFM最大DfmDetailId-- Shintokuro 2023.08.01
        public string GetDetailMaxNum(int DfmId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    #region //取得DFM製程成本
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 DfmDetailId+1 DetailMaxNum FROM PDM.DfmDetail ORDER By DfmDetailId Desc";
                    dynamicParameters.Add("DfmId", DfmId);
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

        #region //GetDfmVersion -- 取得品DFM版本 -- Shintokuro 2023.07.14
        public string GetDfmVersion(int DfmId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    #region //取得品DFM單身
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DISTINCT a.Version, a.DfmId
                            FROM PDM.DfmDetailLog a
                            WHERE a.DfmId = @DfmId
                            ";
                    dynamicParameters.Add("DfmId", DfmId);
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

        #region//GetDfmQuotationItem --取得DFM所需報價項目
        public string GetDfmQuotationItem(int DfmId, int DfmQiId, string DfmNo, string MaterialStatus, string DfmQiProcessStatus, string DataType, int LoginUserId, string RFQDocType
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.DfmQiId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.DfmId,a.DfmItemCategoryId,a.ParentDfmQiId,a.MfgFlag,a.StandardCost, a.TotalManHours, a.TotalMachineHours, a.TotalOspAmount
                          , a.TotalMtlAmount, a.MaterialStatus, a.DfmQiOSPStatus, a.DfmQiProcessStatus, a.[Status], a.[Version], a.DfmQuotationName
                          , a.QuoteModel, a.QuoteNum
                          , a.DefaultMaterialSpreadsheetData,a.MaterialSpreadsheetData
                          , a.DefaultQiProcessSpreadsheetData,a.QiProcessSpreadsheetData
                          , b.ModeId, b.DfmItemCategoryName
                          , ISNULL((x.DfmQuotationName), '最上層') ParentDfmQuotationName
                          , c.DfmNo
                          , d.DocType
                         ";
                    sqlQuery.mainTables =
                        @"  FROM PDM.DfmQuotationItem a 
                            INNER JOIN PDM.DfmItemCategory b on a.DfmItemCategoryId = b.DfmItemCategoryId
                            INNER JOIN MES.ProdMode b1 on b.ModeId = b1.ModeId
                            INNER JOIN PDM.DesignForManufacturing c on a.DfmId = c.DfmId
                            INNER JOIN SCM.RfqDetail d ON c.RfqDetailId=d.RfqDetailId
                            OUTER APPLY(
                                SELECT x1.DfmQuotationName FROM PDM.DfmQuotationItem x1
                                INNER JOIN PDM.DfmItemCategory x2 on x1.DfmItemCategoryId = x2.DfmItemCategoryId
                                INNER JOIN MES.ProdMode x3 on x2.ModeId = x3.ModeId
                                WHERE a.ParentDfmQiId = x1.DfmQiId
                            ) x
                        ";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    switch(DataType)
                    {
                        case "Material":
                            queryCondition = @"AND a.MaterialStatus != 'S'
                                AND EXISTS(
                                    SELECT * FROM EIP.DocNotify x
                                    INNER JOIN EIP.NotifyRole x1 on x.RoleId = x1.RoleId
                                    INNER JOIN EIP.NotifyUser x2 on x.RoleId = x2.RoleId
                                    WHERE a.DfmItemCategoryId = x.SubProdId
                                    AND x.DocType = 'DFM'
                                    AND x.SubProdType = 'Material'
                                    AND x2.UserId = @LoginUserId)
                                ";
                            dynamicParameters.Add("LoginUserId", LoginUserId);

                            break;
                        case "DfmQiOSP":
                            break;
                        case "Process":
                            queryCondition = @"AND a.DfmQiProcessStatus != 'S'
                                AND EXISTS(
                                    SELECT * FROM EIP.DocNotify x
                                    INNER JOIN EIP.NotifyRole x1 on x.RoleId = x1.RoleId
                                    INNER JOIN EIP.NotifyUser x2 on x.RoleId = x2.RoleId
                                    WHERE a.DfmItemCategoryId = x.SubProdId
                                    AND x.DocType = 'DFM'
                                    AND x.SubProdType = 'Process'
                                    AND x2.UserId = @LoginUserId)
                                ";
                            dynamicParameters.Add("LoginUserId", LoginUserId);
                            break;
                    }
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DfmId", @" AND a.DfmId = @DfmId", DfmId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DfmQiId", @" AND a.DfmQiId = @DfmQiId", DfmQiId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DfmNo", @" AND c.DfmNo = @DfmNo", DfmNo);
                    if (RFQDocType.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RFQDocType", @" AND d.DocType IN @RFQDocType", RFQDocType.Split(','));
                    if (MaterialStatus.Length > 0)
                    {
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MaterialStatus", @" AND a.MaterialStatus  IN @MaterialStatus", MaterialStatus.Split(','));
                    }
                    if (DfmQiProcessStatus.Length > 0)
                    {
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DfmQiProcessStatus", @" AND a.DfmQiProcessStatus  IN @DfmQiProcessStatus", DfmQiProcessStatus.Split(','));
                    }

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.DfmQiId DESC";
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

        #region//GetDfmQuotationItemTree --取得DFM所需報價項目(樹狀)
        public string GetDfmQuotationItemTree(int DfmId, int OpenedLevel)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region//檢核DFM單頭
                    string DfmNo = "";
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 a.DfmId, a.Version, a.DfmNo
                                FROM PDM.DesignForManufacturing a
                                WHERE a.DfmId=@DfmId";
                    dynamicParameters.Add("DfmId", DfmId);
                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0) throw new SystemException("【DFM】查無頭,請重新確認");
                    foreach (var item in result)
                    {
                        DfmNo = item.DfmNo;
                    }
                    List<QuotationItem> quotationItems = new List<QuotationItem>();
                    dynamicParameters = new DynamicParameters();
                    sql = @"DECLARE @rowsAdded int
                            
                            DECLARE @dfmquotationItems TABLE
                            ( 
                                DfmQiId int,
                                ParentDfmQiId int,
                                Level int,
                                Route nvarchar(MAX),
                                processed int DEFAULT(0),
                                ModeId int DEFAULT(0),
                                MfgFlag nvarchar(1),
                                QuoteModel nvarchar(1),
                                QuoteNum float DEFAULT(0),
                                MaterialStatus nvarchar(1),
                                DfmQiOSPStatus nvarchar(1),
                                DfmQiProcessStatus nvarchar(1),
                                DfmProcessStatus nvarchar(2),
                                DfmQuotationName nvarchar(20),
                                DfmItemCategoryId int DEFAULT(0)

                            )
                            
                            INSERT @dfmquotationItems
                                SELECT DfmQiId, ISNULL(a.ParentDfmQiId, -1) ParentDfmQiId, 1 Level
                                , CAST(ISNULL(a.ParentDfmQiId, -1) AS nvarchar(MAX)) AS Route
                                , 0, c.ModeId, a.MfgFlag, a.QuoteModel, a.QuoteNum, a.MaterialStatus, a.DfmQiOSPStatus
                                , a.DfmQiProcessStatus, b.DfmProcessStatus, a.DfmQuotationName, a.DfmItemCategoryId
                                FROM PDM.DfmQuotationItem a 
                                INNER JOIN PDM.DesignForManufacturing b on a.DfmId = b.DfmId
                                INNER JOIN PDM.DfmItemCategory c on a.DfmItemCategoryId = c.DfmItemCategoryId
                                WHERE a.DfmId =@DfmId
                                AND a.ParentDfmQiId = -1
                            
                            SET @rowsAdded=@@rowcount
                            
                            WHILE @rowsAdded > 0
                            BEGIN
                                UPDATE @dfmquotationItems SET processed = 1 WHERE processed = 0
                                
                                INSERT @dfmquotationItems
                                    SELECT b.DfmQiId, b.ParentDfmQiId, (a.Level + 1) Level
                                    , CAST(a.Route + ',' + CAST(b.ParentDfmQiId AS nvarchar(MAX)) AS nvarchar(MAX)) AS Route
                                    , 0, d.ModeId, b.MfgFlag, b.QuoteModel, b.QuoteNum, b.MaterialStatus, b.DfmQiOSPStatus
                                    , b.DfmQiProcessStatus, c.DfmProcessStatus, b.DfmQuotationName, b.DfmItemCategoryId
                                    FROM @dfmquotationItems a
                                    INNER JOIN PDM.DfmQuotationItem b on a.DfmQiId = b.ParentDfmQiId
                                    INNER JOIN PDM.DesignForManufacturing c on b.DfmId = c.DfmId
                                    INNER JOIN PDM.DfmItemCategory d on b.DfmItemCategoryId = d.DfmItemCategoryId
                                    WHERE a.processed = 1
                                
                                SET @rowsAdded = @@rowcount

                                UPDATE @dfmquotationItems SET processed = 2 WHERE processed = 1
                            END;
                            
                            SELECT a.DfmQiId, a.ParentDfmQiId, a.Level, a.Route, c.ModeId, a.MfgFlag, a.QuoteModel, a.QuoteNum
                            , a.MaterialStatus, a.DfmQiOSPStatus, a.DfmQiProcessStatus
                            , c.DfmItemCategoryName, d.DfmProcessStatus, a.DfmQuotationName, a.DfmItemCategoryId
                            FROM @dfmquotationItems a 
                            INNER JOIN PDM.DfmQuotationItem b on a.DfmQiId = b.DfmQiId
                            INNER JOIN PDM.DfmItemCategory c on b.DfmItemCategoryId = c.DfmItemCategoryId
                            INNER JOIN PDM.DesignForManufacturing d on b.DfmId = d.DfmId
                            ORDER BY a.Level, a.Route
                           ";
                    dynamicParameters.Add("DfmId", DfmId);
                    quotationItems = sqlConnection.Query<QuotationItem>(sql, dynamicParameters).ToList();
                    #endregion

                    var data = new DfmQuotationTree
                    {
                        value = DfmNo,
                        id = "-1",
                        ModeId = -1,
                        MfgFlag = "",
                        QuoteModel = "",
                        QuoteNum = -1,
                        MaterialStatus = "",
                        DfmQiOSPStatus = "",
                        DfmQiProcessStatus = "",
                        DfmProcessStatus = "",
                        DfmQuotationName = "",
                        DfmItemCategoryId = -1,
                        opened = true
                    };

                    if (quotationItems.Count > 0)
                    {
                        data.items = quotationItems
                            .Where(x => x.Level == 1 && x.ParentDfmQiId == Convert.ToInt32(data.id))
                            .OrderBy(x => x.Level)
                            .ThenBy(x => x.Route)
                            .Select(x => new DfmQuotationTree
                            {
                                value = x.DfmItemCategoryName,
                                id = x.DfmQiId.ToString(),
                                ModeId = x.ModeId,
                                MfgFlag = x.MfgFlag,
                                QuoteModel = x.QuoteModel,
                                QuoteNum = x.QuoteNum,
                                MaterialStatus = x.MaterialStatus,
                                DfmQiOSPStatus = x.DfmQiOSPStatus,
                                DfmQiProcessStatus = x.DfmQiProcessStatus,
                                DfmProcessStatus = x.DfmProcessStatus,
                                DfmQuotationName = x.DfmQuotationName,
                                DfmItemCategoryId = x.DfmItemCategoryId,
                                opened = true,
                                level = (int)x.Level
                            })
                            .ToList();

                        if (data.items.Count > 0) Recursion(data.items);


                    }
                    void Recursion(List<DfmQuotationTree> itemTree)
                    {
                        if (itemTree.Count > 0)
                        {
                            for (int i = 0; i < itemTree.Count; i++)
                            {
                                itemTree[i].items = quotationItems
                                    .Where(x => x.Level == (itemTree[i].level + 1) && x.ParentDfmQiId == Convert.ToInt32(itemTree[i].id))
                                    .OrderBy(x => x.Level)
                                    .ThenBy(x => x.Route)
                                    .Select(x => new DfmQuotationTree
                                    {
                                        value = x.DfmItemCategoryName,
                                        id = x.DfmQiId.ToString(),
                                        ModeId = x.ModeId,
                                        MfgFlag = x.MfgFlag,
                                        QuoteModel = x.QuoteModel,
                                        QuoteNum = x.QuoteNum,
                                        MaterialStatus = x.MaterialStatus,
                                        DfmQiOSPStatus = x.DfmQiOSPStatus,
                                        DfmQiProcessStatus = x.DfmQiProcessStatus,
                                        DfmProcessStatus = x.DfmProcessStatus,
                                        DfmQuotationName = x.DfmQuotationName,
                                        DfmItemCategoryId = x.DfmItemCategoryId,
                                        opened = itemTree[i].level < OpenedLevel ? true : false,
                                        level = (int)x.Level
                                    })
                                    .ToList();

                                if (itemTree[i].items.Count > 0) Recursion(itemTree[i].items);
                            }
                        }
                    }

                    List<DfmQuotationTree> trees = new List<DfmQuotationTree>
                    {
                        data
                    };

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = trees
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

        #region //GetDfmQiMaterial -- 取得DFM使用物料-- Shintokuro 2023.07.07
        public string GetDfmQiMaterial(int DfmQiId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    #region //取得DFM使用物料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.DfmQiMaterialId,a.DfmQiId, a.MaterialNo, a.MaterialName, a.SolutionQtyFlag
                            ,a.Quantity, a.UnitPrice, a.Amount, MAX(a.DfmQiMaterialId) MaxId, b.MaterialStatus
                            FROM PDM.DfmQiMaterial a
                            INNER JOIN PDM.DfmQuotationItem b on a.DfmQiId = b.DfmQiId
                            WHERE a.DfmQiId = @DfmQiId
                            GROUP BY a.DfmQiMaterialId,a.DfmQiId, a.MaterialNo, a.SolutionQtyFlag
                            , a.MaterialName, a.Quantity, a.UnitPrice, a.Amount, b.MaterialStatus
                            ";
                    dynamicParameters.Add("DfmQiId", DfmQiId);
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

        #region //GetDfmQiOSP -- 取得DFM使用委外加工-- Shintokuro 2023.07.07
        public string GetDfmQiOSP(int DfmQiId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    #region //取得DFM使用委外加工
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.DfmQiOSPId,a.DfmQiId, a.OspName, a.Amount, MAX(a.DfmQiOSPId) MaxId
                            FROM PDM.DfmQiOSP a
                            WHERE a.DfmQiId = @DfmQiId
                            GROUP BY a.DfmQiOSPId,a.DfmQiId, a.OspName, a.Amount
                            ";
                    dynamicParameters.Add("DfmQiId", DfmQiId);
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

        #region //GetDfmQiProcess -- 取得DFM製程成本-- Shintokuro 2023.07.12
        public string GetDfmQiProcess(int DfmQiId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    #region //取得DFM製程成本
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.DfmQiProcessId,a.DfmQiId, a.SortNumber, a.ParameterId, a.AllocationType, a.AllocationQty
                            , a.ManHours, a.MachineHours, a.YieldRate
                            , b1.ProcessName
                            , MAX(a.DfmQiProcessId) MaxId, c.DfmQiProcessStatus
                            FROM PDM.DfmQiProcess a
                            INNER JOIN MES.ProcessParameter b on a.ParameterId = b.ParameterId
                            INNER JOIN MES.Process b1 on b.ProcessId = b1.ProcessId
                            INNER JOIN PDM.DfmQuotationItem c on a.DfmQiId = c.DfmQiId
                            WHERE a.DfmQiId = @DfmQiId
                            GROUP BY a.DfmQiProcessId,a.DfmQiId, a.SortNumber, a.ParameterId, a.AllocationType, a.AllocationQty
                            , a.ManHours, a.MachineHours, a.YieldRate, b1.ProcessName, c.DfmQiProcessStatus
                            ";
                    dynamicParameters.Add("DfmQiId", DfmQiId);
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

        #region //GetDfmQiProcessMaxNum -- 取得DFM製程成本最大號-- Shintokuro 2023.08.01
        public string GetDfmQiProcessMaxNum(int DfmId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    #region //取DfmQiProcessMax
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 DfmQiProcessId+1 DfmQiProcessMaxNum FROM PDM.DfmQiProcess ORDER By DfmQiProcessId Desc";
                    dynamicParameters.Add("DfmId", DfmId);
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

        #region//取得DFM物料資訊
        #endregion

        #region//取得DFM所需製程
        #endregion

        #region//取得DFM製程標工成本
        #endregion

        #region//GetDfmTemplateParameter --取得DFM欄位數值
        public string GetDfmTemplateParameter(int RfqProTypeId, int ProductUseId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sql = @" ";
                    //BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MoProcessId", @" AND e.MoProcessId= @MoProcessId", MoProcessId);
                    //BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MachineId", @" AND b.MachineId= @MachineId", MachineId);
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

        #region //GetDesignForManufacturingForSpreadsheer -- 取得DFM資料(Spreadsheet用) -- Ann 2023-08-02
        public string GetDesignForManufacturingForSpreadsheer(int DfmId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.DfmId, a.DefaultSpreadsheetData, a.SpreadsheetData
                            FROM PDM.DesignForManufacturing a 
                            WHERE a.DfmId = @DfmId";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "DfmId", @" AND a.DfmId = @DfmId", DfmId);

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

        #region //GetDfmItem -- 取得DFM項目資料 -- Ann 2023-08-02
        public string GetDfmItem(int DfmItemId, string DfmItemNo, string DfmItemName, string Status)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.DfmItemId, a.DfmItemNo, a.DfmItemName, a.[Status], a.DfmItemNo + ' ' + a.DfmItemName DfmFullItemNo
                            FROM PDM.DfmItem a 
                            WHERE a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "DfmItemId", @" AND a.DfmItemId = @DfmItemId", DfmItemId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "DfmItemNo", @" AND a.DfmItemNo LIKE '%' + @DfmItemNo + '%'", DfmItemNo);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "DfmItemName", @" AND a.DfmItemName LIKE '%' + @DfmItemName + '%'", DfmItemName);
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
        #endregion

        #region//Add
        
        #region //AddDfmHead-- 建立DFM單頭 -- Ding 2023.07.07
        public string AddDfmHead(int RfqDetailId, string DfmDate,string DifficultyLevel, int RdUserId)
        {
            try
            {
                int CompanyId = -1;
                RfqDetailId = 1;
                if(RfqDetailId <= 0) throw new SystemException("DFM【RFQ】不能為空!");
                if (DfmDate.Length <= 0) throw new SystemException("DFM【DFM日期】長度錯誤!");
                if (DifficultyLevel.Length <= 0) throw new SystemException("工具群組【難易度】不能為空!");
                if (RdUserId <= 0) throw new SystemException("DFM【研發人員】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region//檢核RfqDetailId所屬公司別
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.CompanyId
                                FROM SCM.RfqDetail a
                                INNER JOIN BAS.Company b ON a.CompanyId=b.CompanyId
                                WHERE a.RfqDetailId=@RfqDetailId
                                ORDER BY a.CompanyId DESC";
                        dynamicParameters.Add("RfqDetailId", RfqDetailId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() != 1)
                        {
                            throw new SystemException("【RFQ】查無公司別");
                        }
                        else {
                            foreach (var item in result)
                            {
                                CompanyId = item.CompanyId;
                            }
                        }
                        #endregion

                        #region //判斷研發人員是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[User]
                                WHERE UserId = @RdUserId";
                        dynamicParameters.Add("RdUserId", RdUserId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("找不到研發人員資料，請重新確認!");
                        #endregion

                        #region//DFM NO編碼
                        string DfmNo = "";
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 DfmId
                                FROM PDM.DesignForManufacturing
                                ORDER BY DfmId DESC";
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0)
                        {
                            foreach (var item in result) {

                                int DfmSeq = (Convert.ToInt16(item.DfmId)+1);
                                DfmNo = "DFM-" + CreateDate.ToString("yyyyMMdd") + DfmSeq.ToString().PadLeft(4, '0'); ;
                            }                            
                        }
                        else
                        {
                            DfmNo = "DFM-"+ CreateDate.ToString("yyyyMMdd") + "0001";
                        }
                        #endregion

                        #region //取得預設Default SpreadsheetData
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 DefaultSpreadsheetData
                                FROM PDM.DesignForManufacturing
                                WHERE DefaultSpreadsheetData IS NOT NULL
                                --AND DefaultQiProcessSpreadsheetData IS NOT NULL";
                        var defaultSpreadsheetResult = sqlConnection.Query(sql, dynamicParameters);
                        string DefaultSpreadsheetData = "";
                        foreach (var item in defaultSpreadsheetResult)
                        {
                            DefaultSpreadsheetData = item.DefaultSpreadsheetData;
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO PDM.DesignForManufacturing (DfmNo, RfqDetailId, DfmDate,DifficultyLevel,RdUserId, Version, DefaultSpreadsheetData, DfmProcessStatus, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.DfmId,INSERTED.DfmNo, INSERTED.Version, INSERTED.DfmProcessStatus, INSERTED.RfqDetailId
                                VALUES (@DfmNo, @RfqDetailId, @DfmDate ,@DifficultyLevel ,@RdUserId, @Version, @DefaultSpreadsheetData, @DfmProcessStatus, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                DfmNo,
                                RfqDetailId,
                                DfmDate,
                                DifficultyLevel,
                                RdUserId,
                                Version = "000",
                                DefaultSpreadsheetData,
                                DfmProcessStatus = "1",
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

        #region //AddDfmDetail-- DFM單身新增修改版本變更 -- Shintokuro 2023.07.07
        public string AddDfmDetail(int DfmId, string DfmContentList, string VersionControl)
        {
            try
            {
                int rowsAffected = 0;
                string Version = "";
                if (DfmId <= 0) throw new SystemException("DFM【單頭】不能為空!");
                if (DfmContentList.Length <= 0) throw new SystemException("DFM【單身】不能沒有資料!");
                var DfmContentJson = DfmContentList.Split(',');

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region//檢核DFM單頭
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.DfmId, a.Version
                                FROM PDM.DesignForManufacturing a
                                WHERE a.DfmId=@DfmId";
                        dynamicParameters.Add("DfmId", DfmId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【DFM】查無頭,請重新確認");
                        foreach (var item in result)
                        {
                            Version = Convert.ToInt32(item.Version).ToString();
                        }
                        #endregion

                        if (VersionControl == "Y")
                        {
                            Version = (Convert.ToInt32(Version) + 1).ToString().PadLeft(3, '0');

                            #region //DFM單身Log新增
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO PDM.DfmDetailLog OUTPUT INSERTED.DfmDetailLogId
                                    SELECT a.DfmId, a.DfmItemId, a.Standard, a.FinalSpec, a.CustomerSpec, a.SuggestSpec
                                    ,a.RdRemark, a.CustomerFeedback, a.Version
                                    ,@CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy
                                    FROM PDM.DfmDetail a
                                    INNER JOIN PDM.DesignForManufacturing b on a.DfmId = b.DfmId
                                    WHERE a.DfmId = @DfmId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy,
                                    DfmId
                                });

                            var result3 = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += result3.Count();
                            #endregion

                            //#region //DFM報價項目Log新增
                            //dynamicParameters = new DynamicParameters();
                            //sql = @"INSERT INTO PDM.DfmQuotationItemLog OUTPUT INSERTED.DfmQiLogId
                            //        SELECT a.DfmQiId, a.DfmId, a.ModeId, a.MfgFlag, a.ParentDfmQiId, a.StandardCost
                            //        , a.TotalManHours, a.TotalMachineHours, a.TotalOspAmount, a.TotalMtlAmount
                            //        , a.MaterialStatus, a.DfmQiOSPStatus, a.DfmQiProcessStatus, a.Status, a.Version
                            //        ,@CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy
                            //        FROM PDM.DfmQuotationItem a
                            //        INNER JOIN PDM.DesignForManufacturing b on a.DfmId = b.DfmId
                            //        WHERE a.DfmId = @DfmId";
                            //dynamicParameters.AddDynamicParams(
                            //    new
                            //    {
                            //        CreateDate,
                            //        LastModifiedDate,
                            //        CreateBy,
                            //        LastModifiedBy,
                            //        DfmId
                            //    });

                            //result3 = sqlConnection.Query(sql, dynamicParameters);

                            //rowsAffected += result3.Count();
                            //#endregion

                            //#region //報價項目版本更新
                            //dynamicParameters = new DynamicParameters();
                            //sql = @"UPDATE PDM.DfmQuotationItem SET
                            //        Version = @Version,
                            //        LastModifiedDate = @LastModifiedDate,
                            //        LastModifiedBy = @LastModifiedBy
                            //        WHERE DfmId = @DfmId";
                            //dynamicParameters.AddDynamicParams(
                            //    new
                            //    {
                            //        Version,
                            //        LastModifiedDate,
                            //        LastModifiedBy,
                            //        DfmId
                            //    });
                            //result3 = sqlConnection.Query(sql, dynamicParameters);
                            //rowsAffected += result3.Count();
                            //#endregion

                            #region //DFM單頭版本更新
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE PDM.DesignForManufacturing SET
                                    Version = @Version,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE DfmId = @DfmId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    Version,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    DfmId
                                });
                            result3 = sqlConnection.Query(sql, dynamicParameters);
                            rowsAffected += result3.Count();
                            #endregion
                        }

                        foreach (var item in DfmContentJson)
                        {
                            int DfmItemId = Convert.ToInt32(item.Split('❤')[0]);
                            string Standard = item.Split('❤')[1];
                            string FinalSpec = item.Split('❤')[2];
                            string CustomerSpec = item.Split('❤')[3];
                            string SuggestSpec = item.Split('❤')[4];
                            string RdRemark = item.Split('❤')[5];
                            string CustomerFeedback = item.Split('❤')[6];
                            int DfmDetailId = Convert.ToInt32(item.Split('❤')[7]);

                            #region//檢核DFM參數
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                FROM PDM.DfmItem
                                WHERE DfmItemId=@DfmItemId 
                                AND CompanyId=@CompanyId";
                            dynamicParameters.Add("DfmItemId", DfmItemId);
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【DFM】查無此DFM參數,請重新確認");
                            #endregion

                            #region//檢核單身最大序號
                            int SortNumber = 0;
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT COUNT(a.DfmDetailId) MaxSort
                                    FROM PDM.DfmDetail a
                                    WHERE a.DfmId=@DfmId";
                            dynamicParameters.Add("DfmId", DfmId);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() > 0)
                            {
                                foreach (var item1 in result)
                                {
                                    SortNumber = Convert.ToInt32(item1.MaxSort) + 1;
                                }
                            }
                            else
                            {
                                SortNumber = 1;
                            }
                            #endregion

                            #region//新增修改Dfm單身
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                FROM PDM.DfmDetail
                                WHERE DfmDetailId=@DfmDetailId";
                            dynamicParameters.Add("DfmDetailId", DfmDetailId);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO PDM.DfmDetail (DfmId, Version, DfmItemId, Standard, FinalSpec, CustomerSpec
                                        ,SuggestSpec, RdRemark, CustomerFeedback
                                        ,CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.DfmDetailId
                                        VALUES (@DfmId, @Version, @DfmItemId, @Standard, @FinalSpec, @CustomerSpec
                                        ,@SuggestSpec, @RdRemark, @CustomerFeedback
                                        ,@CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        DfmId,
                                        Version,
                                        DfmItemId,
                                        Standard,
                                        FinalSpec,
                                        CustomerSpec,
                                        SuggestSpec,
                                        RdRemark,
                                        CustomerFeedback,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();
                            }
                            else
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE PDM.DfmDetail SET
                                        Version = @Version,
                                        DfmItemId = @DfmItemId,
                                        Standard = @Standard,
                                        FinalSpec = @FinalSpec,
                                        SuggestSpec = @SuggestSpec,
                                        RdRemark = @RdRemark,
                                        CustomerFeedback = @CustomerFeedback,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE DfmDetailId = @DfmDetailId
                                        AND DfmId = @DfmId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        Version,
                                        DfmItemId,
                                        Standard,
                                        FinalSpec,
                                        CustomerSpec,
                                        SuggestSpec,
                                        RdRemark,
                                        CustomerFeedback,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        DfmDetailId,
                                        DfmId
                                    });

                                rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //AddDfmQuotationItem-- 建立DFM報價項目 -- Shintokuro 2023.07.12
        public string AddDfmQuotationItem(int DfmId, string DfmQuotationName, int DfmItemCategoryId, int ParentDfmQiId, string MfgFlag,
            string MaterialStatus, string DfmQiProcessStatus, double StandardCost, string QuoteModel, double QuoteNum)
        {
            try
            {
                int rowsAffected = 0;
                if (DfmId <= 0) throw new SystemException("DFM【單頭】不能為空!");
                if (DfmItemCategoryId <= 0) throw new SystemException("DFM【產品類別】不能沒有資料!");
                if (DfmQuotationName.Length <= 0) throw new SystemException("DFM【項目名稱】不能沒有資料!");
                if (DfmQuotationName.Length > 20) throw new SystemException("DFM【項目名稱】不能超過20字元!");
                if (QuoteModel == "2")
                {
                    if (QuoteNum < 0) throw new SystemException("DFM【報價數量】報價模式選擇固定數量,其數量欄位不可以小於0");
                }
                else
                {
                    QuoteNum = 0;
                }

                int? nullData = null;
                string DfmQiOSPStatus = "S"; //托外管理預設S 暫時不考慮
                //string DfmQiProcessStatus = ""; //製程管理
                //MaterialStatus 物料管理 由前段控制 是否需考慮 S不需要 A需要
                //DfmQiProcessStatus 製程管理 由前段控制 是否需考慮 S不需要 A需要
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region//檢核DFM【單頭】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.DfmId, b.DocType
                                FROM PDM.DesignForManufacturing a
                                INNER JOIN SCM.RfqDetail b on a.RfqDetailId = b.RfqDetailId
                                WHERE a.DfmId=@DfmId";
                        dynamicParameters.Add("DfmId", DfmId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【DFM】單頭查無資料,請重新確認");
                        foreach (var item in result)
                        {
                            if (item.DocType == "F") throw new SystemException("【DFM】RFQ單據已經失效,不能在異動更新");
                        }
                        #endregion

                        #region//檢核DFM【報價項目】是否存在
                        if (ParentDfmQiId > 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.DfmQiId
                                FROM PDM.DfmQuotationItem a
                                WHERE a.DfmQiId=@DfmQiId";
                            dynamicParameters.Add("DfmQiId", ParentDfmQiId);
                             result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【DFM】報價項目查無資料,請重新確認");
                        }
                        #endregion

                        #region//檢核DFM【是否標準建】是否存在
                        if (MfgFlag == "Y")//製造件
                        {
                            if (StandardCost > 0) throw new SystemException("報價項目為自製件【標準金額】不用輸入");
                            StandardCost = 0;
                        }
                        else if (MfgFlag == "N")//標準件
                        {
                            if (StandardCost < 0) throw new SystemException("報價項目為標準件【標準金額】不能為負數");
                        }
                        #endregion

                        #region//檢核DFM【生產模式】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.DfmItemCategoryId
                                FROM PDM.DfmItemCategory a
                                WHERE a.DfmItemCategoryId=@DfmItemCategoryId";
                        dynamicParameters.Add("DfmItemCategoryId", DfmItemCategoryId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【DFM】報價項目表查無資料,請重新確認");
                        #endregion

                        #region //取得預設Default SpreadsheetData
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 DefaultMaterialSpreadsheetData,DefaultQiProcessSpreadsheetData
                                FROM PDM.DfmQuotationItem
                                WHERE DefaultMaterialSpreadsheetData IS NOT NULL
                                AND DefaultQiProcessSpreadsheetData IS NOT NULL";
                        var defaultSpreadsheetResult = sqlConnection.Query(sql, dynamicParameters);
                        string DefaultMaterialSpreadsheetData = "";
                        string DefaultQiProcessSpreadsheetData = "";
                        foreach (var item in defaultSpreadsheetResult)
                        {
                            DefaultMaterialSpreadsheetData = item.DefaultMaterialSpreadsheetData;
                            DefaultQiProcessSpreadsheetData = item.DefaultQiProcessSpreadsheetData;
                        }
                        #endregion

                        #region//新增
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO PDM.DfmQuotationItem (DfmId, DfmQuotationName, DfmItemCategoryId, ParentDfmQiId, MfgFlag, StandardCost, QuoteModel, QuoteNum
                                , TotalManHours,TotalMachineHours, TotalOspAmount, TotalMtlAmount
                                , MaterialStatus, DfmQiOSPStatus, DfmQiProcessStatus, Status, DefaultMaterialSpreadsheetData, DefaultQiProcessSpreadsheetData
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.DfmQiId
                                VALUES (@DfmId, @DfmQuotationName, @DfmItemCategoryId, @ParentDfmQiId, @MfgFlag, @StandardCost, @QuoteModel, @QuoteNum
                                , @TotalManHours, @TotalMachineHours, @TotalOspAmount,@TotalMtlAmount
                                , @MaterialStatus, @DfmQiOSPStatus, @DfmQiProcessStatus, @Status, @DefaultMaterialSpreadsheetData, @DefaultQiProcessSpreadsheetData
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                DfmId,
                                DfmQuotationName,
                                DfmItemCategoryId,
                                ParentDfmQiId,
                                MfgFlag,
                                StandardCost,
                                QuoteModel,
                                QuoteNum,
                                TotalManHours = 0,
                                TotalMachineHours = 0,
                                TotalOspAmount = 0,
                                TotalMtlAmount = 0,
                                OspAmount = 0,
                                MaterialStatus,
                                DfmQiOSPStatus,
                                DfmQiProcessStatus,
                                Status = "S",
                                Version = "000",
                                DefaultMaterialSpreadsheetData,
                                DefaultQiProcessSpreadsheetData,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += insertResult.Count();
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
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

        #region //AddDfmQiMaterial-- 建立/更新DFM使用物料 -- Shintokuro 2023.07.11
        public string AddDfmQiMaterial(int DfmQiId, string DfmqmContentList)
        {
            try
            {
                int rowsAffected = 0;
                if (DfmQiId <= 0) throw new SystemException("DFM【成本表】不能為空!");
                if (DfmqmContentList.Length <= 0) throw new SystemException("DFM【單身】不能沒有資料!");
                var DfmqmContentJson = DfmqmContentList.Split(',');

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region//檢核DFM【成本表】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.DfmQiId
                                FROM PDM.DfmQuotationItem a
                                WHERE a.DfmQiId=@DfmQiId";
                        dynamicParameters.Add("DfmQiId", DfmQiId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【DFM】成本表查無資料,請重新確認");
                        #endregion

                        foreach (var item in DfmqmContentJson)
                        {
                            int DfmQiMaterialId = Convert.ToInt32(item.Split('❤')[0]);
                            string MaterialNo = item.Split('❤')[1];
                            string MaterialName = item.Split('❤')[2];
                            int SolutionQtyFlag = Convert.ToInt32(item.Split('❤')[3]);
                            //Double Quantity = item.Split('❤')[4] != "" ? Convert.ToDouble(item.Split('❤')[4]) : 0;
                            Double UnitPrice = item.Split('❤')[4] != "" ? Convert.ToDouble(item.Split('❤')[4]) : 0;
                            //Double Amount = item.Split('❤')[6] != "" ? Convert.ToDouble(item.Split('❤')[6]) : 0;
                            string DatabaseAction = item.Split('❤')[7] != "" ? item.Split('❤')[7] : throw new SystemException("資料有問題請重新整理!!!");
                            if (MaterialNo.Length <= 0) throw new SystemException("物料代碼不能為空!!!");
                            if (MaterialNo.Length > 20) throw new SystemException("物料代碼字數不能超過20字元!!!");
                            if (MaterialName.Length <= 0) throw new SystemException("物料名稱不能為空!!!");
                            if (MaterialName.Length > 20) throw new SystemException("物料名稱字數不能超過20字元!!!");
                            //if (Quantity < 0) throw new SystemException("數量不能為負數!!!");
                            if (UnitPrice < 0) throw new SystemException("單價不能為負數!!!");
                            //if (Amount <0) throw new SystemException("材料金額不能為負數!!!");

                            #region//檢核單身最大序號
                            int SortNumber = 0;
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT COUNT(a.DfmQiMaterialId) MaxSort
                                    FROM PDM.DfmQiMaterial a
                                    WHERE a.DfmQiId=@DfmQiId";
                            dynamicParameters.Add("DfmQiId", DfmQiId);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() > 0)
                            {
                                foreach (var item1 in result)
                                {
                                    SortNumber = Convert.ToInt32(item1.MaxSort) + 1;
                                }
                            }
                            else
                            {
                                SortNumber = 1;
                            }
                            #endregion

                            #region//檢核DFM 參數
                            
                            if (DatabaseAction == "Create")
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO PDM.DfmQiMaterial (DfmQiId, MaterialNo, MaterialName, SolutionQtyFlag
                                        , UnitPrice, RefSolutionQty, Version
                                        ,CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.DfmQiMaterialId
                                        VALUES (@DfmQiId, @MaterialNo, @MaterialName, @SolutionQtyFlag
                                        , @UnitPrice, @RefSolutionQty, @Version
                                        ,@CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        DfmQiId,
                                        MaterialNo,
                                        MaterialName,
                                        SolutionQtyFlag,
                                        UnitPrice,
                                        RefSolutionQty = "N",
                                        Version = "000",
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();
                            }
                            else if (DatabaseAction == "Update")
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                    FROM PDM.DfmQiMaterial
                                    WHERE DfmQiMaterialId=@DfmQiMaterialId";
                                dynamicParameters.Add("DfmQiMaterialId", DfmQiMaterialId);
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("【DFM】物料成本表查無編號Id為" + DfmQiMaterialId + "這筆資料,請重新確認");

                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE PDM.DfmQiMaterial SET
                                        MaterialNo = @MaterialNo,
                                        MaterialName = @MaterialName,
                                        SolutionQtyFlag = @SolutionQtyFlag,
                                        UnitPrice = @UnitPrice,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE DfmQiMaterialId = @DfmQiMaterialId
                                        AND DfmQiId = @DfmQiId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        MaterialNo,
                                        MaterialName,
                                        SolutionQtyFlag,
                                        UnitPrice,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        DfmQiMaterialId,
                                        DfmQiId
                                    });

                                rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //AddDfmQiOSP-- 建立/更新使用委外加工-- Shintokuro 2023.07.11
        public string AddDfmQiOSP(int DfmQiId, string DfmqiospContentList)
        {
            try
            {
                int rowsAffected = 0;
                if (DfmQiId <= 0) throw new SystemException("DFM【成本表】不能為空!");
                if (DfmqiospContentList.Length <= 0) throw new SystemException("DFM【單身】不能沒有資料!");
                var DfmqiospContentJson = DfmqiospContentList.Split(',');

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region//檢核DFM【成本表】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.DfmQiId
                                FROM PDM.DfmQuotationItem a
                                WHERE a.DfmQiId=@DfmQiId";
                        dynamicParameters.Add("DfmQiId", DfmQiId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【DFM】成本表查無資料,請重新確認");
                        #endregion

                        foreach (var item in DfmqiospContentJson)
                        {
                            int DfmQiOSPId = Convert.ToInt32(item.Split('❤')[0]);
                            string OspName = item.Split('❤')[1];
                            Double Amount = Convert.ToDouble(item.Split('❤')[2]);
                            if (OspName.Length <= 0) throw new SystemException("物料名稱不能為空!!!");
                            if (OspName.Length > 20) throw new SystemException("物料名稱字數不能超過20字元!!!");
                            if (Amount < 0) throw new SystemException("材料金額不能為負數!!!");

                            #region//檢核單身最大序號
                            int SortNumber = 0;
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT COUNT(a.DfmQiOSPId) MaxSort
                                    FROM PDM.DfmQiOSP a
                                    WHERE a.DfmQiId=@DfmQiId";
                            dynamicParameters.Add("DfmQiId", DfmQiId);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() > 0)
                            {
                                foreach (var item1 in result)
                                {
                                    SortNumber = Convert.ToInt32(item1.MaxSort) + 1;
                                }
                            }
                            else
                            {
                                SortNumber = 1;
                            }
                            #endregion

                            #region//檢核DFM 參數
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM PDM.DfmQiOSP
                                    WHERE DfmQiOSPId=@DfmQiOSPId";
                            dynamicParameters.Add("DfmQiOSPId", DfmQiOSPId);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO PDM.DfmQiOSP (DfmQiId, OspName, Amount
                                        ,CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.DfmQiOSPId
                                        VALUES (@DfmQiId, @OspName, @Amount
                                        ,@CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        DfmQiId,
                                        OspName,
                                        Amount,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();
                            }
                            else
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE PDM.DfmQiOSP SET
                                        OspName = @OspName,
                                        Amount = @Amount,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE DfmQiOSPId = @DfmQiOSPId
                                        AND DfmQiId = @DfmQiId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        OspName,
                                        Amount,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        DfmQiOSPId,
                                        DfmQiId
                                    });

                                rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //AddDfmQiProcess-- 建立/更新製程成本-- Shintokuro 2023.07.13
        public string AddDfmQiProcess(int DfmQiId, string DfmqpContentList)
        {
            try
            {
                int rowsAffected = 0;
                if (DfmQiId <= 0) throw new SystemException("DFM【成本表】不能為空!");
                if (DfmqpContentList.Length <= 0) throw new SystemException("DFM【單身】不能沒有資料!");
                var DfmqpContentJson = DfmqpContentList.Split(',');
                Double? nulldate = null;
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region//檢核DFM【成本表】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.DfmQiId
                                FROM PDM.DfmQuotationItem a
                                WHERE a.DfmQiId=@DfmQiId";
                        dynamicParameters.Add("DfmQiId", DfmQiId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【DFM】成本表查無資料,請重新確認");
                        #endregion

                        foreach (var item in DfmqpContentJson)
                        {
                            int DfmQiProcessId = Convert.ToInt32(item.Split('❤')[0]);
                            int ParameterId = Convert.ToInt32(item.Split('❤')[1]);
                            string AllocationType = item.Split('❤')[2];
                            Double? AllocationQty = item.Split('❤')[3] != "" ? Convert.ToDouble(item.Split('❤')[3]) : 0;
                            Double? ManHours = item.Split('❤')[4] != "" ? Convert.ToDouble(item.Split('❤')[4]) : 0;
                            Double? MachineHours = item.Split('❤')[5] != "" ? Convert.ToDouble(item.Split('❤')[5]) : 0;
                            Double? YieldRate = item.Split('❤')[6] != "" ? Convert.ToDouble(item.Split('❤')[6]) : 0;
                            string DatabaseAction = item.Split('❤')[7] != "" ? item.Split('❤')[7] : throw new SystemException("資料有問題請重新整理!!!");
                            if(ManHours <0) throw new SystemException("工時不能為負數!!!");
                            if(MachineHours < 0) throw new SystemException("機時不能為負數!!!");
                            if(YieldRate < 0) throw new SystemException("良率請填寫0~100之間!!!");
                            if(YieldRate > 100) throw new SystemException("良率請填寫0~100之間!!!");


                            if (ParameterId <= 0) throw new SystemException("製程不能為空!!!");
                            if (AllocationType.Length <= 0) throw new SystemException("工時分配類別不能為空!!!");
                            if (AllocationType == "2")
                            {
                                if (AllocationQty <= 0) throw new SystemException("當工時分配類別為【By製程批量】，分配基數數量不能小於0");
                            }
                            else
                            {
                                AllocationQty = 0;
                            }

                            

                            #region//檢核單身最大序號
                            int SortNumber = 0;
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT COUNT(a.DfmQiProcessId) MaxSort
                                    FROM PDM.DfmQiProcess a
                                    WHERE a.DfmQiId=@DfmQiId";
                            dynamicParameters.Add("DfmQiId", DfmQiId);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() > 0)
                            {
                                foreach (var item1 in result)
                                {
                                    SortNumber = Convert.ToInt32(item1.MaxSort) + 1;
                                }
                            }
                            else
                            {
                                SortNumber = 1;
                            }
                            #endregion

                            #region//判斷是新增或修改
                            if (DatabaseAction == "Create")
                            {
                                #region//判斷製程是否重複
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.ParameterId,c.ProcessName
                                    FROM PDM.DfmQiProcess a
                                    INNER JOIN MES.ProcessParameter b on a.ParameterId = b.ParameterId
                                    INNER JOIN MES.Process c on b.ProcessId = c.ProcessId
                                    WHERE a.ParameterId=@ParameterId
                                    and a.DfmQiId=@DfmQiId";
                                dynamicParameters.Add("ParameterId", DfmQiId);
                                dynamicParameters.Add("DfmQiId", DfmQiId);
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() > 0)
                                {
                                    foreach (var item2 in result)
                                    {
                                        throw new SystemException("該製程【" + item2.ProcessName + "】已經存在不能重複新增!!!");
                                    }
                                }
                                #endregion

                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO PDM.DfmQiProcess (DfmQiId, SortNumber, ParameterId
                                        ,AllocationType, AllocationQty, ManHours, MachineHours, YieldRate
                                        ,CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.DfmQiProcessId
                                        VALUES (@DfmQiId, @SortNumber, @ParameterId
                                        ,@AllocationType, @AllocationQty, @ManHours, @MachineHours, @YieldRate
                                        ,@CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        DfmQiId,
                                        SortNumber,
                                        ParameterId,
                                        AllocationType,
                                        AllocationQty,
                                        ManHours,
                                        MachineHours,
                                        YieldRate,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();
                            }
                            else if(DatabaseAction == "Update")
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                    FROM PDM.DfmQiProcess
                                    WHERE DfmQiProcessId=@DfmQiProcessId";
                                dynamicParameters.Add("DfmQiProcessId", DfmQiProcessId);
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("【DFM】製程成本表查無編號Id為"+ DfmQiProcessId + "這筆資料,請重新確認");

                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE PDM.DfmQiProcess SET
                                        SortNumber = @SortNumber,
                                        ParameterId= @ParameterId,
                                        AllocationType = @AllocationType,
                                        AllocationQty = @AllocationQty,
                                        ManHours = @ManHours,
                                        MachineHours = @MachineHours,
                                        YieldRate = @YieldRate,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE DfmQiProcessId = @DfmQiProcessId
                                        AND DfmQiId = @DfmQiId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        SortNumber,
                                        ParameterId,
                                        AllocationType,
                                        AllocationQty,
                                        ManHours,
                                        MachineHours,
                                        YieldRate,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        DfmQiProcessId,
                                        DfmQiId
                                    });

                                rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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
        #endregion

        #region//Update
        
        #region //UpdateDfmHead-- 更新DFM單頭 -- Shintokuro 2023.07.13
        public string UpdateDfmHead(int DfmId, string DfmDate, string DifficultyLevel, int RdUserId)
        {
            try
            {
                int rowsAffected = 0;
                if (DfmDate.Length <= 0) throw new SystemException("DFM【單據日期】不能為空!");
                if (DifficultyLevel.Length <= 0) throw new SystemException("DFM【難易度】不能為空!");
                if (RdUserId <= 0) throw new SystemException("DFM【研發人員】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region//檢核DFM【單頭】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.DfmId, b.DocType
                                FROM PDM.DesignForManufacturing a
                                INNER JOIN SCM.RfqDetail b on a.RfqDetailId = b.RfqDetailId
                                WHERE a.DfmId=@DfmId";
                        dynamicParameters.Add("DfmId", DfmId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【DFM】單頭查無資料,請重新確認");
                        foreach(var item in result)
                        {
                            if(item.DocType == "F") throw new SystemException("【DFM】RFQ單據已經失效,不能在異動更新");
                        }
                        #endregion

                        #region //判斷研發人員是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[User]
                                WHERE UserId = @RdUserId";
                        dynamicParameters.Add("RdUserId", RdUserId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("找不到研發人員資料，請重新確認!");
                        #endregion

                        #region//更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PDM.DesignForManufacturing SET
                                        DfmDate = @DfmDate,
                                        DifficultyLevel = @DifficultyLevel,
                                        RdUserId = @RdUserId,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE DfmId = @DfmId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                DfmDate,
                                DifficultyLevel,
                                RdUserId,
                                LastModifiedDate,
                                LastModifiedBy,
                                DfmId
                            });

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

        #region //UpdateDfmHeadStatus -- 更新DFM單頭狀態 -- Shintokru 2023.07.13
        public string UpdateDfmHeadStatus(int DfmId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        #region//檢核DFM【單頭】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.Status
                                FROM PDM.DesignForManufacturing a
                                WHERE a.DfmId=@DfmId";
                        dynamicParameters.Add("DfmId", DfmId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【DFM】單頭查無資料,請重新確認");
                        #endregion

                        #region //調整為相反狀態
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
                        sql = @"UPDATE PDM.DesignForManufacturing SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE DfmId = @DfmId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                DfmId
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

        #region //UpdateDfmQuotationItem-- 更新DFM報價項目 -- Shintokuro 2023.07.12
        public string UpdateDfmQuotationItem(int DfmQiId, int DfmId, string DfmQuotationName, int DfmItemCategoryId, string MfgFlag,
            string MaterialStatus, string DfmQiProcessStatus, double StandardCost, string QuoteModel, double QuoteNum)
        {
            try
            {
                int rowsAffected = 0;
                if (DfmId <= 0) throw new SystemException("DFM【單頭】不能為空!");
                if (DfmQiId <= 0) throw new SystemException("DFM【報價項目】不能為空!");
                if (DfmItemCategoryId <= 0) throw new SystemException("DFM【項目種類】不能沒有資料!");
                if (DfmQuotationName.Length <= 0) throw new SystemException("DFM【項目名稱】不能沒有資料!");
                if (DfmQuotationName.Length > 20) throw new SystemException("DFM【項目名稱】不能超過20字元!");

                if (MfgFlag == "Y")
                {
                    StandardCost = 0;
                }
                if (QuoteModel == "2")
                {
                    if (QuoteNum < 0) throw new SystemException("DFM【報價數量】報價模式選擇固定數量,其數量欄位不可以小於0");
                }
                else
                {
                    QuoteNum = 0;
                }


                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region//檢核DFM【單頭】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.DfmId, b.DocType
                                FROM PDM.DesignForManufacturing a
                                INNER JOIN SCM.RfqDetail b on a.RfqDetailId = b.RfqDetailId
                                WHERE a.DfmId=@DfmId";
                        dynamicParameters.Add("DfmId", DfmId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【DFM】單頭查無資料,請重新確認");
                        foreach (var item in result)
                        {
                            if (item.DocType == "F") throw new SystemException("【DFM】RFQ單據已經失效,不能在異動更新");
                        }
                        #endregion

                        #region//檢核DFM【所需報價項目】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.DfmQiId
                                FROM PDM.DfmQuotationItem a
                                WHERE a.DfmQiId=@DfmQiId";
                        dynamicParameters.Add("DfmQiId", DfmQiId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【DFM】所需報價項目查無資料,請重新確認");
                        #endregion

                        #region//檢核DFM【項目種類】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.DfmItemCategoryId
                                FROM PDM.DfmItemCategory a
                                WHERE a.DfmItemCategoryId=@DfmItemCategoryId";
                        dynamicParameters.Add("DfmItemCategoryId", DfmItemCategoryId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【DFM】項目種類查無資料,請重新確認");
                        #endregion

                        #region//檢核DFM【是否標準建】是否存在
                        if (MfgFlag == "Y")//製造件
                        {
                            if (StandardCost > 0) throw new SystemException("報價項目為自製件【標準金額】不用輸入");
                            StandardCost = 0;
                        }
                        else if (MfgFlag == "N")//標準件
                        {
                            if (StandardCost < 0) throw new SystemException("報價項目為標準件【標準金額】不能為負數");
                        }
                        #endregion

                        #region//更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PDM.DfmQuotationItem SET
                                DfmQuotationName = @DfmQuotationName,
                                DfmItemCategoryId = @DfmItemCategoryId,
                                MfgFlag = @MfgFlag,
                                QuoteModel = @QuoteModel,
                                QuoteNum = @QuoteNum,
                                StandardCost = @StandardCost,
                                MaterialStatus = @MaterialStatus,
                                DfmQiProcessStatus = @DfmQiProcessStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE DfmQiId = @DfmQiId
                                AND DfmId = @DfmId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                DfmQuotationName,
                                DfmItemCategoryId,
                                MfgFlag,
                                QuoteModel,
                                QuoteNum,
                                StandardCost,
                                MaterialStatus,
                                DfmQiProcessStatus,
                                LastModifiedDate,
                                LastModifiedBy,
                                DfmQiId,
                                DfmId
                            });

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

        #region //updateConfirmDfm-- DFM報價項目確認通知 -- Shintokuro 2023.07.12
        public string updateConfirmDfm(int DfmId)
        {
            try
            {
                int rowsAffected = 0;
                if (DfmId <= 0) throw new SystemException("DFM【單頭】不能為空!");
                int DfmProcessStatus = 2;
                string DfmNo = "";
                int RfqSalesId = -1;
                string RfqFullNo = "";

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region//檢核DFM報價項目是否完成維護
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT (a.MfgFlag + a.MaterialStatus + a.DfmQiProcessStatus + a.DfmQiOSPStatus) FinallyStatus
                                ,b.DfmNo
                                ,c.SalesId,c.DocType
                                ,(c1.RfqNo + c.RfqSequence) RfqFullNo
                                FROM PDM.DfmQuotationItem a
                                INNER JOIN PDM.DesignForManufacturing b on a.DfmId = b.DfmId
                                INNER JOIN SCM.RfqDetail c on b.RfqDetailId = c.RfqDetailId
                                INNER JOIN SCM.RequestForQuotation c1 on c.RfqId = c1.RfqId
                                WHERE a.DfmId = @DfmId";
                        dynamicParameters.Add("DfmId", DfmId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【DFM】查無報價項目資料,請重新確認");

                        foreach (var item in result)
                        {
                            if (item.DocType == "F") throw new SystemException("【DFM】RFQ單據已經失效,不能在異動更新");

                            string FinallyStatus = item.FinallyStatus;
                            DfmNo = item.DfmNo;
                            RfqSalesId = item.SalesId;
                            RfqFullNo = item.RfqFullNo;
                            if (FinallyStatus == "NSSS" || FinallyStatus == "NYSS" || FinallyStatus == "NSYS" || FinallyStatus == "NYYS" ||
                                FinallyStatus == "YSSS" || FinallyStatus == "YYSS" || FinallyStatus == "YSYS" || FinallyStatus == "YYYS")
                            {
                                DfmProcessStatus = 3;
                            }
                            else
                            {
                                DfmProcessStatus = 2;
                                break;
                            }
                        }
                        #endregion

                        #region//更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PDM.DesignForManufacturing SET
                                        DfmProcessStatus = @DfmProcessStatus,
                                        ProcessStatus2Date = @LastModifiedDate,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE DfmId = @DfmId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                DfmProcessStatus,
                                LastModifiedDate,
                                LastModifiedBy,
                                DfmId
                            });

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //發送Mail
                        if(DfmProcessStatus == 2)
                        {
                            #region //關卡2
                            #region //查詢通知者
                            List<string> DfmQmInfoList = new List<string>();
                            List<string> DfmQpInfoList = new List<string>();

                            List<string> DfmQmUserList = new List<string>();
                            List<string> DfmQpUserList = new List<string>();


                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.DfmQiId,a.DfmItemCategoryId,a.MaterialStatus,a.DfmQiProcessStatus 
                                    ,b.DfmItemCategoryName
                                    FROM PDM.DfmQuotationItem a
                                    INNER JOIN PDM.DfmItemCategory b on a.DfmItemCategoryId = b.DfmItemCategoryId
                                    WHERE a.DfmId = @DfmId
                                    ";
                            dynamicParameters.Add("DfmId", DfmId);
                            var resultNotifyInfo = sqlConnection.Query(sql, dynamicParameters);
                            if (resultNotifyInfo.Count() <= 0) throw new SystemException("DFM【單頭】資料找不到,請重新確認!!");
                            foreach(var item in resultNotifyInfo)
                            {
                                if(item.MaterialStatus == "A")
                                {
                                    DfmQmInfoList.Add(item.DfmQiId + "," + item.DfmItemCategoryId + "," + item.DfmItemCategoryName);
                                }
                                if (item.DfmQiProcessStatus == "A")
                                {
                                    DfmQpInfoList.Add(item.DfmQiId + "," + item.DfmItemCategoryId + "," + item.DfmItemCategoryName);
                                }
                            }

                            foreach(var item in DfmQmInfoList) 
                            {
                                int DfmQiId = Convert.ToInt32( item.Split(',')[0]);
                                int DfmItemCategoryId = Convert.ToInt32( item.Split(',')[1]);
                                string DfmItemCategoryName = item.Split(',')[2];

                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT c.UserId
                                    FROM EIP.DocNotify a
                                    INNER JOIN EIP.NotifyRole b on a.RoleId = b.RoleId
                                    INNER JOIN EIP.NotifyUser c on b.RoleId = c.RoleId
                                    WHERE DocType = 'DFM'
                                    AND SubProdType = 'Material'
                                    AND SubProdId = @DfmItemCategory
                                    ";
                                dynamicParameters.Add("DfmItemCategory", DfmItemCategoryId);
                                var resultNotifyUser = sqlConnection.Query(sql, dynamicParameters);
                                if (resultNotifyUser.Count() <= 0) throw new SystemException("物料項目:【"+ DfmItemCategoryName + "】尚未維護發送郵件者");
                                foreach (var item1 in resultNotifyUser)
                                {
                                    string info = DfmQiId.ToString() + ',' + item1.UserId;
                                    DfmQmUserList.Add(info);
                                }
                            }


                            foreach (var item in DfmQpInfoList)
                            {
                                int DfmQiId = Convert.ToInt32(item.Split(',')[0]);
                                int DfmItemCategoryId = Convert.ToInt32(item.Split(',')[1]);
                                string DfmItemCategoryName = item.Split(',')[2];

                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT c.UserId
                                    FROM EIP.DocNotify a
                                    INNER JOIN EIP.NotifyRole b on a.RoleId = b.RoleId
                                    INNER JOIN EIP.NotifyUser c on b.RoleId = c.RoleId
                                    WHERE DocType = 'DFM'
                                    AND SubProdType = 'Process'
                                    AND SubProdId = @DfmItemCategory
                                    ";
                                dynamicParameters.Add("DfmItemCategory", DfmItemCategoryId);
                                var resultNotifyUser = sqlConnection.Query(sql, dynamicParameters);
                                if (resultNotifyUser.Count() <= 0) throw new SystemException("製程項目:【" + DfmItemCategoryName + "】尚未維護發送郵件者");
                                foreach (var item1 in resultNotifyUser)
                                {
                                    string info = DfmQiId.ToString() + ',' + item1.UserId;
                                    DfmQpUserList.Add(info);
                                }
                            }
                            #endregion

                            #region //發送Mail
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
                            dynamicParameters.Add("SettingSchema", "DfmProcessStatus2Notification");
                            dynamicParameters.Add("SettingNo", "Y");

                            var resultMailTemplate = sqlConnection.Query(sql, dynamicParameters);
                            if (resultMailTemplate.Count() <= 0) throw new SystemException("Mail設定錯誤!");
                            #endregion

                            #region //Mail資料(物料)
                            if (DfmQmUserList.Count() > 0)
                            {
                                foreach (var item in DfmQmUserList)
                                {
                                    int DfmQiId = Convert.ToInt32(item.Split(',')[0]);
                                    int UserId = Convert.ToInt32(item.Split(',')[1]);

                                    #region //DFM資料
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT b.DfmItemCategoryName,a.DfmQuotationName
                                            FROM PDM.DfmQuotationItem a
                                            INNER JOIN PDM.DfmItemCategory b on a.DfmItemCategoryId = b.DfmItemCategoryId
                                            INNER JOIN PDM.DesignForManufacturing c on a.DfmId = c.DfmId
                                            WHERE a.DfmId = @DfmId
                                            AND a.DfmQiId = @DfmQiId
                                            ";
                                    dynamicParameters.Add("DfmId", DfmId);
                                    dynamicParameters.Add("DfmQiId", DfmQiId);

                                    var resultDfmInfo = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultDfmInfo.Count() <= 0) throw new SystemException("DFM資料錯誤!");

                                    List<string> modeNameList = new List<string>();
                                    string DfmQuotationName = "";

                                    foreach (var item1 in resultDfmInfo)
                                    {
                                        DfmQuotationName = item1.DfmQuotationName;
                                    }
                                    #endregion

                                    #region //查找UserInfo
                                    string MemberName = "";
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.UserId, a.UserName, a.Email
                                            FROM BAS.[User] a
                                            WHERE a.UserId = @UserId";
                                    dynamicParameters.Add("UserId", UserId);

                                    var resultUserInfo = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultUserInfo.Count() <= 0) throw new SystemException("收件者資料錯誤!");

                                    string salesName = "";
                                    string email = "";
                                    foreach (var item1 in resultUserInfo)
                                    {
                                        MemberName = item1.UserName;
                                        email = item1.Email;
                                    }

                                    foreach (var item1 in resultMailTemplate)
                                    {
                                        string mailSubject = item1.MailSubject,
                                        mailContent = HttpUtility.UrlDecode(item1.MailContent);

                                        string hyperlink = "http://" + HttpContext.Current.Request.Url.Authority + "/DfmInformation/DesignForManufacturingManagement";
                                        hyperlink = string.Format("<a href=\"{0}\">{1}</a>", hyperlink, DfmNo);

                                        #region //Mail內容
                                        mailContent = mailContent.Replace("[Type]", "物料");
                                        mailContent = mailContent.Replace("[DfmNo]", DfmNo);
                                        mailContent = mailContent.Replace("[DfmQuotationName]", DfmQuotationName);
                                        mailContent = mailContent.Replace("[MemberName]", MemberName);
                                        mailContent = mailContent.Replace("[hyperlink]", hyperlink);
                                        #endregion

                                        #region //寄送Mail
                                        MailConfig mailConfig = new MailConfig
                                        {
                                            Host = item1.Host,
                                            Port = Convert.ToInt32(item1.Port),
                                            SendMode = Convert.ToInt32(item1.SendMode),
                                            From = item1.MailFrom,
                                            Subject = mailSubject,
                                            Account = item1.Account,
                                            Password = item1.Password,
                                            MailTo = MemberName+":"+ email,
                                            MailCc = item1.MailCc,
                                            MailBcc = item1.MailBcc,
                                            HtmlBody = mailContent,
                                            TextBody = "-"
                                        };

                                        BaseHelper.MailSend(mailConfig);
                                        Thread.Sleep(300); // 等候一秒（300毫秒）
                                        #endregion
                                    }

                                    #endregion
                                }
                            }
                            #endregion

                            #region //Mail資料(製程)
                            if (DfmQpUserList.Count() > 0)
                            {
                                foreach (var item in DfmQpUserList)
                                {
                                    int DfmQiId = Convert.ToInt32(item.Split(',')[0]);
                                    int UserId = Convert.ToInt32(item.Split(',')[1]);

                                    #region //DFM資料
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT b.DfmItemCategoryName,a.DfmQuotationName
                                            FROM PDM.DfmQuotationItem a
                                            INNER JOIN PDM.DfmItemCategory b on a.DfmItemCategoryId = b.DfmItemCategoryId
                                            INNER JOIN PDM.DesignForManufacturing c on a.DfmId = c.DfmId
                                            WHERE a.DfmId = @DfmId
                                            AND a.DfmQiId = @DfmQiId
                                            ";
                                    dynamicParameters.Add("DfmId", DfmId);
                                    dynamicParameters.Add("DfmQiId", DfmQiId);

                                    var resultDfmInfo = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultDfmInfo.Count() <= 0) throw new SystemException("DFM資料錯誤!");

                                    List<string> modeNameList = new List<string>();
                                    string DfmQuotationName = "";

                                    foreach (var item1 in resultDfmInfo)
                                    {
                                        DfmQuotationName = item1.DfmQuotationName;
                                    }
                                    #endregion

                                    #region //查找UserInfo
                                    string MemberName = "";
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.UserId, a.UserName, a.Email
                                            FROM BAS.[User] a
                                            WHERE a.UserId = @UserId";
                                    dynamicParameters.Add("UserId", UserId);

                                    var resultUserInfo = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultUserInfo.Count() <= 0) throw new SystemException("收件者資料錯誤!");

                                    string salesName = "";
                                    string email = "";
                                    foreach (var item1 in resultUserInfo)
                                    {
                                        MemberName = item1.UserName;
                                        email = item1.Email;
                                    }

                                    foreach (var item1 in resultMailTemplate)
                                    {
                                        string mailSubject = item1.MailSubject,
                                        mailContent = HttpUtility.UrlDecode(item1.MailContent);

                                        string hyperlink = "http://" + HttpContext.Current.Request.Url.Authority + "/DfmInformation/DesignForManufacturingManagement";
                                        hyperlink = string.Format("<a href=\"{0}\">{1}</a>", hyperlink, DfmNo);

                                        #region //Mail內容
                                        mailContent = mailContent.Replace("[Type]", "製程");
                                        mailContent = mailContent.Replace("[DfmNo]", DfmNo);
                                        mailContent = mailContent.Replace("[DfmQuotationName]", DfmQuotationName);
                                        mailContent = mailContent.Replace("[MemberName]", MemberName);
                                        mailContent = mailContent.Replace("[hyperlink]", hyperlink);
                                        #endregion

                                        #region //寄送Mail
                                        MailConfig mailConfig = new MailConfig
                                        {
                                            Host = item1.Host,
                                            Port = Convert.ToInt32(item1.Port),
                                            SendMode = Convert.ToInt32(item1.SendMode),
                                            From = item1.MailFrom,
                                            Subject = mailSubject,
                                            Account = item1.Account,
                                            Password = item1.Password,
                                            MailTo = MemberName + ":" + email,
                                            MailCc = item1.MailCc,
                                            MailBcc = item1.MailBcc,
                                            HtmlBody = mailContent,
                                            TextBody = "-"
                                        };

                                        BaseHelper.MailSend(mailConfig);
                                        Thread.Sleep(300); // 等候一秒（300毫秒）
                                        #endregion
                                    }

                                    #endregion
                                }
                            }
                            #endregion

                            #endregion
                        }
                        else if(DfmProcessStatus == 3)
                        {
                            #region //自動計算成本
                            TxQuotationCalculation(DfmId);
                            #endregion

                            #region //關卡3

                            #region //發送Mail
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
                            dynamicParameters.Add("SettingSchema", "DfmProcessStatus2Notification");
                            dynamicParameters.Add("SettingNo", "Y");

                            var resultMailTemplate = sqlConnection.Query(sql, dynamicParameters);
                            if (resultMailTemplate.Count() <= 0) throw new SystemException("Mail設定錯誤!");
                            #endregion

                            #region //查找UserInfo
                            string MemberName = "";
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.UserId, a.UserName, a.Email
                                    FROM BAS.[User] a
                                    WHERE a.UserId = @UserId";
                            dynamicParameters.Add("UserId", RfqSalesId);

                            var resultUserInfo = sqlConnection.Query(sql, dynamicParameters);
                            if (resultUserInfo.Count() <= 0) throw new SystemException("收件者資料錯誤!");

                            string salesName = "";
                            string email = "";
                            foreach (var item1 in resultUserInfo)
                            {
                                salesName = item1.UserName;
                                email = item1.Email;
                            }

                            foreach (var item1 in resultMailTemplate)
                            {
                                string mailSubject = item1.MailSubject,
                                mailContent = HttpUtility.UrlDecode(item1.MailContent);

                                string hyperlink = "http://" + HttpContext.Current.Request.Url.Authority + "/RequestForQuotation/RfqLineSolution";
                                hyperlink = string.Format("<a href=\"{0}\">{1}</a>", hyperlink, RfqFullNo);

                                #region //Mail內容
                                mailContent = mailContent.Replace("[RfqFullNo]", RfqFullNo);
                                mailContent = mailContent.Replace("[DfmNo]", DfmNo);
                                mailContent = mailContent.Replace("[MemberName]", salesName);
                                mailContent = mailContent.Replace("[hyperlink]", hyperlink);
                                #endregion

                                #region //寄送Mail
                                MailConfig mailConfig = new MailConfig
                                {
                                    Host = item1.Host,
                                    Port = Convert.ToInt32(item1.Port),
                                    SendMode = Convert.ToInt32(item1.SendMode),
                                    From = item1.MailFrom,
                                    Subject = mailSubject,
                                    Account = item1.Account,
                                    Password = item1.Password,
                                    MailTo = salesName+":"+ email+";" + "系統開發室-賴科榜:ted_lai@zy-tech.com.tw",
                                    MailCc = item1.MailCc,
                                    MailBcc = item1.MailBcc,
                                    HtmlBody = mailContent,
                                    TextBody = "-"
                                };

                                BaseHelper.MailSend(mailConfig);
                                Thread.Sleep(300); // 等候一秒（300毫秒）
                                #endregion
                            }

                            #endregion

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

        #region //updateConfirmDfmDataStatus-- 報價項目所需維護資料確認 -- Shintokuro 2023.07.17
        public string updateConfirmDfmDataStatus(int DfmQiId, string Data)
        {
            try
            {
                int rowsAffected = 0;
                if (DfmQiId <= 0) throw new SystemException("DFM【報價項目】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region//檢核DFM【報價項目】是否存在
                        int DfmId = -1;
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.DfmQiId, a.DfmId, c.DocType
                                FROM PDM.DfmQuotationItem a
                                INNER JOIN PDM.DesignForManufacturing  b on a.DfmId = b.DfmId
                                INNER JOIN SCM.RfqDetail c on b.RfqDetailId = c.RfqDetailId
                                WHERE a.DfmQiId=@DfmQiId";
                        dynamicParameters.Add("DfmQiId", DfmQiId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【報價項目】查無資料,請重新確認");
                        foreach(var item in result)
                        {
                            if (item.DocType == "F") throw new SystemException("【DFM】RFQ單據已經失效,不能在異動更新");
                            DfmId = item.DfmId;
                        }
                        #endregion

                        switch (Data)
                        {
                            case "Material":

                                #region//檢核DFM【物料管理】是否有資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 a.DfmQiMaterialId
                                        FROM PDM.DfmQiMaterial a
                                        WHERE a.DfmQiId=@DfmQiId";
                                dynamicParameters.Add("DfmQiId", DfmQiId);
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("【物料管理】查無資料,請重新確認是否有送出資料");
                                #endregion


                                #region//物料更新
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE PDM.DfmQuotationItem SET
                                        MaterialStatus = @MaterialStatus,
                                        MaterialConfirmDate = @LastModifiedDate,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE DfmQiId = @DfmQiId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        MaterialStatus = "Y",
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        DfmQiId
                                    });

                                rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                                break;
                            case "DfmQiOSP":

                                #region//檢核DFM【委外管理】是否有資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 a.DfmQiOSPId
                                        FROM PDM.DfmQiOSP a
                                        WHERE a.DfmQiId=@DfmQiId";
                                dynamicParameters.Add("DfmQiId", DfmQiId);
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("【委外管理】查無資料,請重新確認是否有送出資料");
                                #endregion

                                #region//托外更新
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE PDM.DfmQuotationItem SET
                                        DfmQiOSPStatus = @DfmQiOSPStatus,
                                        QiProcessConfirmDate = @LastModifiedDate,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE DfmQiId = @DfmQiId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        DfmQiOSPStatus = "Y",
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        DfmQiId
                                    });

                                rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                                break;
                            case "DfmQiProcess":
                                #region//檢核DFM【製程管理】是否有資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 a.DfmQiProcessId
                                        FROM PDM.DfmQiProcess a
                                        WHERE a.DfmQiId=@DfmQiId";
                                dynamicParameters.Add("DfmQiId", DfmQiId);
                                result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("【製程管理】查無資料,請重新確認是否有送出資料");
                                #endregion

                                #region//更新
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE PDM.DfmQuotationItem SET
                                        DfmQiProcessStatus = @DfmQiProcessStatus,
                                        DfmQiOSPConfirmDate = @LastModifiedDate,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE DfmQiId = @DfmQiId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        DfmQiProcessStatus = "Y",
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        DfmQiId
                                    });

                                rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                                break;
                            case "QuotationItem":
                                break;
                        }

                        #region//檢核DFM報價項目是否完成維護
                        int DfmProcessStatus = 2;
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT (a.MfgFlag + a.MaterialStatus + a.DfmQiProcessStatus + a.DfmQiOSPStatus) FinallyStatus 
                                FROM PDM.DfmQuotationItem a
                                INNER JOIN PDM.DesignForManufacturing b on a.DfmId = b.DfmId
                                WHERE a.DfmId = @DfmId";
                        dynamicParameters.Add("DfmId", DfmId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【DFM】查無報價項目資料,請重新確認");

                        foreach (var item in result)
                        {
                            string FinallyStatus = item.FinallyStatus;
                            if (FinallyStatus == "NSSS" || FinallyStatus == "NYSS" || FinallyStatus == "NSYS" || FinallyStatus == "NYYS" ||
                                FinallyStatus == "YSSS" || FinallyStatus == "YYSS" || FinallyStatus == "YSYS" || FinallyStatus == "YYYS")
                            {
                                DfmProcessStatus = 3;
                            }
                            else
                            {
                                DfmProcessStatus = 2;
                                break;
                            }
                        }

                        if(DfmProcessStatus == 3)
                        {
                            #region//更新
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE PDM.DesignForManufacturing SET
                                        DfmProcessStatus = @DfmProcessStatus,
                                        ProcessStatus3Date = @LastModifiedDate,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE DfmId = @DfmId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    DfmProcessStatus = 3,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    DfmId
                                });

                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //發送Mail
                            #region //自動計算成本
                            TxQuotationCalculation(DfmId);
                            #endregion

                            #region //關卡3
                            #endregion
                            #endregion
                        }

                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            data = DfmProcessStatus == 3 ? "Finish" : "undone"
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

        #region //UpdateDfmQuotationItemLevel -- 更新DFM報價項目階層 -- Ann 2023-07-31
        public string UpdateDfmQuotationItemLevel(int DfmQiId, int QuotationLevel)
        {
            try
            {
                if (QuotationLevel <= 0) throw new SystemException("階層不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();

                        #region //判斷DFM報價項目資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.DfmQuotationItem a 
                                WHERE a.DfmQiId = @DfmQiId";
                        dynamicParameters.Add("DfmQiId", DfmQiId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("DFM報價項目資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PDM.DfmQuotationItem SET
                                QuotationLevel = @QuotationLevel,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE DfmQiId = @DfmQiId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                QuotationLevel,
                                LastModifiedDate,
                                LastModifiedBy,
                                DfmQiId
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

        #region //UpdateDfmQiExcelData -- 暫存DfmQuotationItemExcel資料 -- Shintokuro 2023-08-25
        public string UpdateDfmQiExcelData(int DfmQiId, string DataType, string SpreadsheetData)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;
                        #region//檢核DFM【報價項目】是否存在
                        string MaterialStatus = "";
                        string DfmQiOSPStatus = "";
                        string DfmQiProcessStatus = "";

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.DfmQiId,a.MaterialStatus,a.DfmQiOSPStatus,a.DfmQiProcessStatus, c.DocType
                                FROM PDM.DfmQuotationItem a
                                INNER JOIN PDM.DesignForManufacturing  b on a.DfmId = b.DfmId
                                INNER JOIN SCM.RfqDetail c on b.RfqDetailId = c.RfqDetailId
                                WHERE a.DfmQiId=@DfmQiId";
                        dynamicParameters.Add("DfmQiId", DfmQiId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【報價項目】查無資料,請重新確認");
                        foreach (var item in result)
                        {
                            if (item.DocType == "F") throw new SystemException("【DFM】RFQ單據已經失效,不能在異動更新");
                            MaterialStatus = item.MaterialStatus;
                            DfmQiOSPStatus = item.DfmQiOSPStatus;
                            DfmQiProcessStatus = item.DfmQiProcessStatus;
                        }
                        #endregion

                        switch (DataType)
                        {
                            case "Material":
                                if (MaterialStatus == "Y") throw new SystemException("資料已經完成確認送出,無法再更改");

                                #region //判斷是否已經有上傳過資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM PDM.DfmQiMaterial
                                        WHERE DfmQiId = @DfmQiId";
                                dynamicParameters.Add("DfmQiId", DfmQiId);

                                var resultDfmQiM = sqlConnection.Query(sql, dynamicParameters);
                                if (resultDfmQiM.Count() > 0) throw new SystemException("此紀錄已上傳，無法修改資料!!");
                                #endregion

                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE PDM.DfmQuotationItem SET
                                        MaterialSpreadsheetData = @SpreadsheetData,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE DfmQiId = @DfmQiId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        SpreadsheetData,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        DfmQiId
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                break;
                            case "DfmQiOSP":
                                break;
                            case "DfmQiProcess":
                                if (DfmQiProcessStatus == "Y") throw new SystemException("資料已經完成確認送出,無法再更改");

                                #region //判斷是否已經有上傳過資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM PDM.DfmQiProcess
                                        WHERE DfmQiId = @DfmQiId";
                                dynamicParameters.Add("DfmQiId", DfmQiId);

                                var resultDfmQiP = sqlConnection.Query(sql, dynamicParameters);
                                if (resultDfmQiP.Count() > 0) throw new SystemException("此紀錄已上傳，無法修改資料!!");
                                #endregion

                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE PDM.DfmQuotationItem SET
                                        QiProcessSpreadsheetData = @SpreadsheetData,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE DfmQiId = @DfmQiId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        SpreadsheetData,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        DfmQiId
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                break;
                            default:
                                throw new SystemException("資料類別異常，請重新確認!!");
                                break;
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

        #region //updateConfirmDfmQiExcelData-- 報價項目所需維護資料確認(Excel) -- Shintokuro 2023.08.26
        public string updateConfirmDfmQiExcelData(int DfmQiId, string DataType, string SpreadsheetJson, string SpreadsheetData)
        {
            try
            {
                int rowsAffected = 0;
                int ModeId = 0;
                if (DfmQiId <= 0) throw new SystemException("DFM【報價項目】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region//檢核DFM【報價項目】是否存在
                        int DfmId = -1;
                        string MaterialStatus = "";
                        string DfmQiOSPStatus = "";
                        string DfmQiProcessStatus = "";

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.DfmQiId, a.DfmId,a.MaterialStatus,a.DfmQiProcessStatus,b.ModeId, d.DocType
                                FROM PDM.DfmQuotationItem a
                                INNER JOIN PDM.DfmItemCategory b on a.DfmItemCategoryId = b.DfmItemCategoryId
                                INNER JOIN PDM.DesignForManufacturing  c on a.DfmId = c.DfmId
                                INNER JOIN SCM.RfqDetail d on c.RfqDetailId = d.RfqDetailId
                                WHERE a.DfmQiId=@DfmQiId";
                        dynamicParameters.Add("DfmQiId", DfmQiId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【報價項目】查無資料,請重新確認");
                        foreach (var item in result)
                        {
                            if (item.DocType == "F") throw new SystemException("【DFM】RFQ單據已經失效,不能在異動更新");
                            DfmId = item.DfmId;
                            MaterialStatus = item.MaterialStatus;
                            DfmQiOSPStatus = item.DfmQiOSPStatus;
                            DfmQiProcessStatus = item.DfmQiProcessStatus;
                            ModeId = item.ModeId;
                        }
                        #endregion

                        var spreadsheetJson = JObject.Parse(SpreadsheetJson);


                        switch (DataType)
                        {
                            #region //物料項目維護
                            case "Material":
                                if (MaterialStatus == "Y") throw new SystemException("資料已經完成確認送出,無法再更改");
                                #region //先儲存目前Spreadsheet資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE PDM.DfmQuotationItem SET
                                        MaterialStatus = 'Y',
                                        MaterialSpreadsheetData = @SpreadsheetData,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE DfmQiId = @DfmQiId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        SpreadsheetData,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        DfmQiId
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region //先刪除目前所有DfmDetail
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE PDM.DfmQiMaterial
                                        WHERE DfmQiId = @DfmQiId";
                                dynamicParameters.Add("DfmQiId", DfmQiId);

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region //解析Spreadsheet Data
                                foreach (var item in spreadsheetJson["spreadsheetInfo"])
                                {
                                    var MaterialNo = item["MaterialNo"] != null ? item["MaterialNo"].ToString() : throw new SystemException("【資料維護不完整】物料欄位資料不可以為空,請重新確認~~");
                                    var MaterialName = item["MaterialName"] != null ? item["MaterialName"].ToString() : null;
                                    //string SolutionQtyFlag = item["SolutionQtyFlag"] != null ? item["SolutionQtyFlag"].ToString() : throw new SystemException("【資料維護不完整】報價方式欄位資料不可以為空,請重新確認~~");
                                    string SolutionQtyFlag = null;
                                    double UnitPrice = item["UnitPrice"] != null ? Convert.ToDouble(item["UnitPrice"]) : throw new SystemException("【資料維護不完整】單價欄位資料不可以為空,請重新確認~~");
                                    if (UnitPrice < 0) throw new SystemException("【資料異常】單價不可以為負數,請重新確認~~");

                                    //string[] SolutionQtyFlagArr = new string[] { "1:考慮方案數量", "2:不考慮方案數量" };
                                    //bool isMatch = SolutionQtyFlagArr.Any(itemType => itemType == SolutionQtyFlag);
                                    //if (isMatch)
                                    //{
                                    //    SolutionQtyFlag = SolutionQtyFlag.Split(':')[0];
                                    //}
                                    //else
                                    //{
                                    //    throw new SystemException("【資料異常】參考報價方案欄位不符合設定,請重新確認!!");
                                    //}


                                    #region //確認報價項目資料是否正確
                                    string DfmMaterialNo = item["MaterialNo"].ToString();
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.DfmMaterialId, (a.DfmMaterialNo + ':' + a.DfmMaterialName) DfmMaterialNoDB
                                            FROM PDM.DfmMaterial a 
                                            WHERE a.DfmMaterialNo = @DfmMaterialNo";
                                    dynamicParameters.Add("DfmMaterialNo", DfmMaterialNo.Split(':')[0]);

                                    var resultDfmMaterial = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultDfmMaterial.Count() <= 0) throw new SystemException("找不到DFM物料代碼,請重新確認~~");
                                    int DfmMaterialId = -1;
                                    foreach (var item2 in resultDfmMaterial)
                                    {
                                        if (item2.DfmMaterialNoDB != DfmMaterialNo) throw new SystemException("【資料異常】物料資料與資料庫不相符,請重新確認");
                                        DfmMaterialId = item2.DfmMaterialId;
                                    }
                                    #endregion

                                    #region //INSERT PDM.DfmDetail
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO PDM.DfmQiMaterial (DfmQiId, DfmMaterialId, MaterialNo, MaterialName, SolutionQtyFlag
                                            , UnitPrice, RefSolutionQty, Version
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.DfmQiMaterialId
                                            VALUES (@DfmQiId, @DfmMaterialId, @MaterialNo, @MaterialName, @SolutionQtyFlag
                                            , @UnitPrice, @RefSolutionQty, @Version
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        DfmQiId,
                                        DfmMaterialId,
                                        MaterialNo,
                                        MaterialName = "測試",
                                        SolutionQtyFlag = 1,
                                        UnitPrice,
                                        RefSolutionQty = "N",
                                        Version = "000",
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });

                                    var insertResult = sqlConnection.Query(sql, dynamicParameters);
                                    #endregion
                                }
                                #endregion
                                break;
                            #endregion

                            #region //托外項目維護
                            case "DfmQiOSP":
                                break;
                            #endregion

                            #region //製程項目維護
                            case "DfmQiProcess":
                                if (DfmQiProcessStatus == "Y") throw new SystemException("資料已經完成確認送出,無法再更改");

                                #region //先儲存目前Spreadsheet資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE PDM.DfmQuotationItem SET
                                        DfmQiProcessStatus = 'Y',
                                        QiProcessSpreadsheetData = @SpreadsheetData,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE DfmQiId = @DfmQiId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        SpreadsheetData,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        DfmQiId
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region //先刪除目前所有DfmQiProcess
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE PDM.DfmQiProcess
                                        WHERE DfmQiId = @DfmQiId";
                                dynamicParameters.Add("DfmQiId", DfmQiId);

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region //解析Spreadsheet Data
                                foreach (var item in spreadsheetJson["spreadsheetInfo"])
                                {
                                    string ParameterNo = item["ParameterNo"] != null ? item["ParameterNo"].ToString() : throw new SystemException("【資料維護不完整】製程欄位資料不可以為空,請重新確認~~");
                                    string AllocationType = item["AllocationType"] != null ? item["AllocationType"].ToString() : throw new SystemException("【資料維護不完整】工時分配類別欄位資料不可以為空,請重新確認~~");
                                    string SolutionQtyFlag = item["SolutionQtyFlag"] != null ? item["SolutionQtyFlag"].ToString() : null;
                                    double AllocationQty = item["AllocationQty"] != null ? Convert.ToDouble(item["AllocationQty"]) : 0;
                                    int ManTimeD = item["ManTimeD"] != null ? Convert.ToInt32(item["ManTimeD"]) : 0;
                                    int ManTimeH = item["ManTimeH"] != null ? Convert.ToInt32(item["ManTimeH"]) : 0;
                                    int ManTimeM = item["ManTimeM"] != null ? Convert.ToInt32(item["ManTimeM"]) : 0;
                                    int ManTimeS = item["ManTimeS"] != null ? Convert.ToInt32(item["ManTimeS"]) : 0;
                                    int MachineTimeD = item["MachineTimeD"] != null ? Convert.ToInt32(item["MachineTimeD"]) : 0;
                                    int MachineTimeH = item["MachineTimeH"] != null ? Convert.ToInt32(item["MachineTimeH"]) : 0;
                                    int MachineTimeM = item["MachineTimeM"] != null ? Convert.ToInt32(item["MachineTimeM"]) : 0;
                                    int MachineTimeS = item["MachineTimeS"] != null ? Convert.ToInt32(item["MachineTimeS"]) : 0;
                                    int YieldRate = item["YieldRate"] != null ? Convert.ToInt32(item["YieldRate"]) : 0;

                                    string[] AllocationTypeArr = new string[] { "1:單PCS工時", "2:By製程批量", "3:By報價數量" };
                                    bool isMatch = AllocationTypeArr.Any(itemType => itemType == AllocationType);
                                    if (isMatch)
                                    {
                                        AllocationType = AllocationType.Split(':')[0];
                                    }
                                    else
                                    {
                                        throw new SystemException("【資料異常】工時分配類別欄位不符合設定,請重新確認!!");
                                    }

                                    if (AllocationType == "2")
                                    {
                                        if (AllocationQty < 0) throw new SystemException("【資料異常】工時分配類別:By製程批量，分配基數數量不能為負");
                                    }
                                    else if (AllocationType == "1" || AllocationType == "3")
                                    {
                                        if (AllocationQty != 0) throw new SystemException("【資料異常】工時分配類別:非By製程批量，分配基數只能為0");
                                    }
                                    else
                                    {
                                        throw new SystemException("【資料異常】工時分配類別欄位異常，請重新確認");
                                    }

                                    if (ManTimeH > 23) throw new SystemException("【資料異常】工時/時 不能大於23");
                                    if (ManTimeH < 0) throw new SystemException("【資料異常】工時/時 不能為負數");
                                    if (ManTimeM > 59) throw new SystemException("【資料異常】工時/分 不能大於59");
                                    if (ManTimeM < 0) throw new SystemException("【資料異常】工時/分 不能為負數");
                                    if (ManTimeS > 59) throw new SystemException("【資料異常】工時/秒 不能大於59");
                                    if (ManTimeS < 0) throw new SystemException("【資料異常】工時/秒 不能為負數");
                                    int ManHours = ManTimeH * 60 * 60 + ManTimeM * 60 + ManTimeS;
                                    if (MachineTimeH > 23) throw new SystemException("【資料異常】工時/時 不能大於23");
                                    if (MachineTimeH < 0) throw new SystemException("【資料異常】工時/時 不能為負數");
                                    if (MachineTimeM > 59) throw new SystemException("【資料異常】工時/分 不能大於59");
                                    if (MachineTimeM < 0) throw new SystemException("【資料異常】工時/分 不能為負數");
                                    if (MachineTimeS > 59) throw new SystemException("【資料異常】工時/秒 不能大於59");
                                    if (MachineTimeS < 0) throw new SystemException("【資料異常】工時/秒 不能為負數");
                                    int MachineHours =  MachineTimeH * 60 * 60 + MachineTimeM * 60 + MachineTimeS;

                                    if (YieldRate > 100) throw new SystemException("【資料異常】良率應該0~100區間");
                                    if (YieldRate < 0) throw new SystemException("【資料異常】良率應該0~100區間");

                                    #region //確認報價項目資料是否正確
                                    string ProcessNo = item["ParameterNo"].ToString().Split(':')[0];
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT b.ParameterId,b.ModeId, (a.ProcessNo + ':' + a.ProcessName) ParameterNoDB
                                            FROM MES.Process a 
                                            INNER JOIN MES.ProcessParameter b ON a.ProcessId = b.ProcessId
                                            WHERE a.ProcessNo = @ProcessNo
                                            AND b.ModeId = @ModeId";
                                    dynamicParameters.Add("ProcessNo", ProcessNo);
                                    dynamicParameters.Add("ModeId", ModeId);
                                    var resultProcessParameter = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultProcessParameter.Count() <= 0) throw new SystemException("找不到生產模式代碼,請重新確認~~");
                                    int ParameterId = -1;
                                    foreach (var item2 in resultProcessParameter)
                                    {
                                        if (item2.ParameterNoDB != ParameterNo) throw new SystemException("【資料異常】製程資料與資料庫不相符,請重新確認");
                                        ParameterId = item2.ParameterId;
                                    }
                                    #endregion

                                    #region //INSERT PDM.DfmQiProcess
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO PDM.DfmQiProcess (DfmQiId, SortNumber, ParameterId
                                            ,AllocationType, AllocationQty, ManHours, MachineHours, YieldRate
                                            ,CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.DfmQiProcessId
                                            VALUES (@DfmQiId, @SortNumber, @ParameterId
                                            ,@AllocationType, @AllocationQty, @ManHours, @MachineHours, @YieldRate
                                            ,@CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            DfmQiId,
                                            SortNumber = 0,
                                            ParameterId,
                                            AllocationType,
                                            AllocationQty,
                                            ManHours,
                                            MachineHours,
                                            YieldRate,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });
                                    var insertResult = sqlConnection.Query(sql, dynamicParameters);
                                    #endregion
                                }
                                #endregion
                                break;
                            #endregion

                            default:
                                throw new SystemException("資料類別異常，請重新確認!!");
                                break;
                        }

                        #region//檢核DFM報價項目是否完成維護
                        int DfmProcessStatus = 2;
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT (a.MfgFlag + a.MaterialStatus + a.DfmQiProcessStatus + a.DfmQiOSPStatus) FinallyStatus 
                                FROM PDM.DfmQuotationItem a
                                INNER JOIN PDM.DesignForManufacturing b on a.DfmId = b.DfmId
                                WHERE a.DfmId = @DfmId";
                        dynamicParameters.Add("DfmId", DfmId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【DFM】查無報價項目資料,請重新確認");

                        foreach (var item in result)
                        {
                            string FinallyStatus = item.FinallyStatus;
                            if (FinallyStatus == "NSSS" || FinallyStatus == "NYSS" || FinallyStatus == "NSYS" || FinallyStatus == "NYYS" ||
                                FinallyStatus == "YSSS" || FinallyStatus == "YYSS" || FinallyStatus == "YSYS" || FinallyStatus == "YYYS")
                            {
                                DfmProcessStatus = 3;
                            }
                            else
                            {
                                DfmProcessStatus = 2;
                                break;
                            }
                        }

                        if (DfmProcessStatus == 3)
                        {
                            #region//更新
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE PDM.DesignForManufacturing SET
                                    DfmProcessStatus = @DfmProcessStatus,
                                    ProcessStatus3Date = @LastModifiedDate,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE DfmId = @DfmId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    DfmProcessStatus = 3,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    DfmId
                                });

                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //發送Mail
                            #region //自動計算成本
                            TxQuotationCalculation(DfmId);
                            #endregion

                            #region //關卡3
                            #endregion
                            #endregion
                        }

                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            data = DfmProcessStatus == 3 ? "Finish" : "undone"
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

        #region //updateReConfirmDfmQiExcelData-- 報價項目所需維護資料取消確認(Excel) -- Shintokuro 2023.11.09
        public string updateReConfirmDfmQiExcelData(int DfmQiId, string DataType)
        {
            try
            {
                int rowsAffected = 0;
                int ModeId = 0;
                if (DfmQiId <= 0) throw new SystemException("DFM【報價項目】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region//檢核DFM【報價項目】是否存在
                        int DfmId = -1;
                        string MaterialStatus = "";
                        string DfmQiOSPStatus = "";
                        string DfmQiProcessStatus = "";

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.DfmQiId, a.DfmId,a.MaterialStatus,a.DfmQiProcessStatus,b.ModeId, d.DocType
                                FROM PDM.DfmQuotationItem a
                                INNER JOIN PDM.DfmItemCategory b on a.DfmItemCategoryId = b.DfmItemCategoryId
                                INNER JOIN PDM.DesignForManufacturing  c on a.DfmId = c.DfmId
                                INNER JOIN SCM.RfqDetail d on c.RfqDetailId = d.RfqDetailId
                                WHERE a.DfmQiId=@DfmQiId";
                        dynamicParameters.Add("DfmQiId", DfmQiId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【報價項目】查無資料,請重新確認");
                        foreach (var item in result)
                        {
                            if (item.DocType == "F") throw new SystemException("【DFM】RFQ單據已經失效,不能在異動更新");
                            DfmId = item.DfmId;
                            MaterialStatus = item.MaterialStatus;
                            DfmQiOSPStatus = item.DfmQiOSPStatus;
                            DfmQiProcessStatus = item.DfmQiProcessStatus;
                        }
                        #endregion

                        #region//清楚報價結果
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PDM.DfmQiSolutionProcess
                                      FROM PDM.DfmQiSolutionProcess a
                                           INNER JOIN PDM.DfmQiProcess b ON a.DfmQiProcessId = b.DfmQiProcessId
	                                       INNER JOIN PDM.DfmQuotationItem c ON b.DfmQiId = c.DfmQiId
	                                       INNER JOIN PDM.DesignForManufacturing d ON c.DfmId = d.DfmId
                                     WHERE d.DfmId = @DfmId";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            DfmId
                        });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion



                        switch (DataType)
                        {
                            #region //物料項目維護
                            case "Material":
                                if (MaterialStatus != "Y") throw new SystemException("資料非確認狀態送出,無法反確認");

                                #region //先儲存目前Spreadsheet資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE PDM.DfmQuotationItem SET
                                        MaterialStatus = 'N',
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE DfmQiId = @DfmQiId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        DfmQiId
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region //先刪除目前所有DfmDetail
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE PDM.DfmQiMaterial
                                        WHERE DfmQiId = @DfmQiId";
                                dynamicParameters.Add("DfmQiId", DfmQiId);

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                break;
                            #endregion

                            #region //托外項目維護
                            case "DfmQiOSP":
                                break;
                            #endregion

                            #region //製程項目維護
                            case "DfmQiProcess":
                                if (DfmQiProcessStatus != "Y") throw new SystemException("資料非確認狀態送出,無法反確認");

                                #region //先儲存目前Spreadsheet資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE PDM.DfmQuotationItem SET
                                        DfmQiProcessStatus = 'N',
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE DfmQiId = @DfmQiId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        DfmQiId
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region //先刪除目前所有DfmQiProcess
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE PDM.DfmQiProcess
                                        WHERE DfmQiId = @DfmQiId";
                                dynamicParameters.Add("DfmQiId", DfmQiId);

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                break;
                            #endregion

                            default:
                                throw new SystemException("資料類別異常，請重新確認!!");
                                break;
                        }

                        #region //先儲存目前Spreadsheet資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PDM.DesignForManufacturing SET
                                        DfmProcessStatus = 2,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE DfmId = @DfmId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                LastModifiedDate,
                                LastModifiedBy,
                                DfmId
                            });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
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

        #region//Delete

        #region //DeleteDfmDetail -- 刪除DFM單身 -- Shintokuro 2023.07.10
        public string DeleteDfmDetail(string DfmDetailIdList)
        {
            try
            {
                if (DfmDetailIdList.Length <= 0) throw new SystemException("DFM單身資料不能為空，請重新確認");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;
                        var DfmContentJson = DfmDetailIdList.Split(',');

                        foreach(var DfmDetailId in DfmContentJson)
                        {
                            #region //判斷DFM單身是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                FROM PDM.DfmDetail
                                WHERE DfmDetailId = @DfmDetailId";
                            dynamicParameters.Add("DfmDetailId", DfmDetailId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("DFM單身不存在，請重新確認");
                            #endregion

                            #region //刪除主要table
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE PDM.DfmDetail
                                WHERE DfmDetailId = @DfmDetailId";
                            dynamicParameters.Add("DfmDetailId", DfmDetailId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteDfmQuotationItem -- 刪除DFM所需報價項目 -- Shintokuro 2023.07.13
        public string DeleteDfmQuotationItem(int DfmQiId, int DfmId)
        {
            try
            {
                if (DfmQiId <= 0)
                {
                    if(DfmId <= 0) throw new SystemException("DFM【報價項目】不能為空!");
                }

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;

                        if (DfmId <= 0)
                        {
                            #region//檢核DFM【所需報價項目】是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.DfmQiId
                                FROM PDM.DfmQuotationItem a
                                WHERE a.DfmQiId=@DfmQiId";
                            dynamicParameters.Add("DfmQiId", DfmQiId);
                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【DFM】所需報價項目查無資料,請重新確認");
                            #endregion

                            #region //刪除附屬資料表
                            #region //刪除物料管理
                            #region //刪除主要table
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE PDM.DfmQiMaterial
                                WHERE DfmQiId = @DfmQiId";
                            dynamicParameters.Add("DfmQiId", DfmQiId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                            #endregion

                            #region //刪除托外項目管理
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE PDM.DfmQiOSP
                                WHERE DfmQiId = @DfmQiId";
                            dynamicParameters.Add("DfmQiId", DfmQiId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //刪除製成管理
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE PDM.DfmQiProcess
                                WHERE DfmQiId = @DfmQiId";
                            dynamicParameters.Add("DfmQiId", DfmQiId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                            #endregion

                            #region //刪除主資料表
                            #region //刪除DFM所需報價項目
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE PDM.DfmQuotationItem
                                WHERE DfmQiId = @DfmQiId";
                            dynamicParameters.Add("DfmQiId", DfmQiId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                            #endregion
                        }
                        else
                        {
                            #region//檢核DFM【單據】是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.DfmId
                                FROM PDM.DfmQuotationItem a
                                WHERE a.DfmId=@DfmId";
                            dynamicParameters.Add("DfmId", DfmId);
                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【DFM】單據查無資料,請重新確認");
                            #endregion

                            #region //刪除附屬資料表
                            #region //刪除物料管理
                            #region //刪除主要table
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE PDM.DfmQiMaterial
                                WHERE DfmQiId IN (SELECT DfmQiId FROM PDM.DfmQuotationItem WHERE DfmId = @DfmId)";
                            dynamicParameters.Add("DfmId", DfmId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                            #endregion

                            #region //刪除托外項目管理
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE PDM.DfmQiOSP
                                WHERE DfmQiId IN (SELECT DfmQiId FROM PDM.DfmQuotationItem WHERE DfmId = @DfmId)";
                            dynamicParameters.Add("DfmId", DfmId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //刪除製成管理
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE PDM.DfmQiProcess
                                WHERE DfmQiId IN (SELECT DfmQiId FROM PDM.DfmQuotationItem WHERE DfmId = @DfmId)";
                            dynamicParameters.Add("DfmId", DfmId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                            #endregion

                            #region //刪除主資料表
                            #region //刪除DFM所需報價項目
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE PDM.DfmQuotationItem
                                WHERE DfmId = @DfmId";
                            dynamicParameters.Add("DfmId", DfmId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteDfmQiMaterial -- 刪除DFM物料管理 -- Shintokuro 2023.07.11
        public string DeleteDfmQiMaterial(string DfmQiMaterialIdList)
        {
            try
            {
                if(DfmQiMaterialIdList.Length <=0) throw new SystemException("DFM物料管理資料不能為空，請重新確認");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;
                        var DfmQiMaterialIdJson = DfmQiMaterialIdList.Split(',');

                        foreach (var DfmQiMaterialId in DfmQiMaterialIdJson)
                        {
                            #region //判斷DFM物料管理是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                FROM PDM.DfmQiMaterial
                                WHERE DfmQiMaterialId = @DfmQiMaterialId";
                            dynamicParameters.Add("DfmQiMaterialId", DfmQiMaterialId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("DFM物料管理不存在，請重新確認");
                            #endregion

                            #region //刪除主要table
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE PDM.DfmQiMaterial
                                WHERE DfmQiMaterialId = @DfmQiMaterialId";
                            dynamicParameters.Add("DfmQiMaterialId", DfmQiMaterialId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteDfmQiOSP -- 刪除DFM委外加工管理 -- Shintokuro 2023.07.11
        public string DeleteDfmQiOSP(string DfmQiOSPIdList)
        {
            try
            {
                if (DfmQiOSPIdList.Length <= 0) throw new SystemException("DFM委外加工管理資料不能為空，請重新確認");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;
                        var DfmQiOSPIdJson = DfmQiOSPIdList.Split(',');

                        foreach (var DfmQiOSPId in DfmQiOSPIdJson)
                        {
                            #region //判斷DFM委外加工管理是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                FROM PDM.DfmQiOSP
                                WHERE DfmQiOSPId = @DfmQiOSPId";
                            dynamicParameters.Add("DfmQiOSPId", DfmQiOSPId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("DFM委外加工管理不存在，請重新確認");
                            #endregion

                            #region //刪除主要table
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE PDM.DfmQiOSP
                                    WHERE DfmQiOSPId = @DfmQiOSPId";
                            dynamicParameters.Add("DfmQiOSPId", DfmQiOSPId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteDfmQiProcess -- 刪除DFM製程成本管理 -- Shintokuro 2023.07.13
        public string DeleteDfmQiProcess(string DfmQiProcessIdList)
        {
            try
            {
                if (DfmQiProcessIdList.Length <= 0) throw new SystemException("DFM製程成本管理資料不能為空，請重新確認");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;
                        var DfmQiProcessIdJson = DfmQiProcessIdList.Split(',');

                        foreach (var DfmQiProcessId in DfmQiProcessIdJson)
                        {
                            #region //判斷DFM委外加工管理是否存在
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                FROM PDM.DfmQiProcess
                                WHERE DfmQiProcessId = @DfmQiProcessId";
                            dynamicParameters.Add("DfmQiProcessId", DfmQiProcessId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("DFM委外加工管理不存在，請重新確認");
                            #endregion

                            #region //刪除主要table
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE PDM.DfmQiProcess
                                    WHERE DfmQiProcessId = @DfmQiProcessId";
                            dynamicParameters.Add("DfmQiProcessId", DfmQiProcessId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //DeleteDfmDetailTempSpreadsheet -- 刪除Dfm詳細資料暫存記錄 -- Ann 2023-08-02
        public string DeleteDfmDetailTempSpreadsheet(int DfmId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region//檢核DFM【單頭】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.DfmId, b.DocType
                                FROM PDM.DesignForManufacturing a
                                INNER JOIN SCM.RfqDetail b on a.RfqDetailId = b.RfqDetailId
                                WHERE a.DfmId=@DfmId";
                        dynamicParameters.Add("DfmId", DfmId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【DFM】單頭查無資料,請重新確認");
                        foreach (var item in result)
                        {
                            if (item.DocType == "F") throw new SystemException("【DFM】RFQ單據已經失效,不能在異動更新");
                        }
                        #endregion

                        #region //判斷是否已經有上傳過資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.DfmDetail
                                WHERE DfmId = @DfmId";
                        dynamicParameters.Add("DfmId", DfmId);

                        var DfmDeatilResult = sqlConnection.Query(sql, dynamicParameters);
                        if (DfmDeatilResult.Count() > 0) throw new SystemException("此紀錄已上傳，無法刪除暫存資料!!");
                        #endregion

                        #region //更新MES.QcRecord SpreadsheetData
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PDM.DesignForManufacturing SET
                                SpreadsheetData = null,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE DfmId = @DfmId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                LastModifiedDate,
                                LastModifiedBy,
                                DfmId
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

        #region //DeleteDfmQiExcelData -- 刪除DfmQiExcelData暫存記錄 -- Shintokuro 2023-08-24
        public string DeleteDfmQiExcelData(int DfmQiId,string DataType)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;
                        #region//檢核DFM【報價項目】是否存在
                        string MaterialStatus = "";
                        string DfmQiOSPStatus = "";
                        string DfmQiProcessStatus = "";

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.DfmQiId,a.MaterialStatus,a.DfmQiOSPStatus,a.DfmQiProcessStatus, c.DocType
                                FROM PDM.DfmQuotationItem a
                                INNER JOIN PDM.DesignForManufacturing  b on a.DfmId = b.DfmId
                                INNER JOIN SCM.RfqDetail c on b.RfqDetailId = c.RfqDetailId
                                WHERE a.DfmQiId=@DfmQiId";
                        dynamicParameters.Add("DfmQiId", DfmQiId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【報價項目】查無資料,請重新確認");
                        foreach (var item in result)
                        {
                            if (item.DocType == "F") throw new SystemException("【DFM】RFQ單據已經失效,不能在異動更新");
                            MaterialStatus = item.MaterialStatus;
                            DfmQiOSPStatus = item.DfmQiOSPStatus;
                            DfmQiProcessStatus = item.DfmQiProcessStatus;
                        }
                        #endregion

                        switch (DataType)
                        {
                            case "Material":
                                if (MaterialStatus == "Y") throw new SystemException("資料已經完成確認送出,無法再更改");

                                #region //判斷是否已經有上傳過資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM PDM.DfmQiMaterial
                                        WHERE DfmQiId = @DfmQiId";
                                dynamicParameters.Add("DfmQiId", DfmQiId);

                                var resultDfmQM = sqlConnection.Query(sql, dynamicParameters);
                                if (resultDfmQM.Count() > 0) throw new SystemException("此紀錄已上傳，無法刪除暫存資料!!");
                                #endregion

                                #region //更新PDM.DfmQuotationItem SpreadsheetData
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE PDM.DfmQuotationItem SET
                                        MaterialSpreadsheetData = null,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE DfmQiId = @DfmQiId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        DfmQiId
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                break;
                            case "DfmQiOSP":
                                break;
                            case "DfmQiProcess":
                                if (DfmQiProcessStatus == "Y") throw new SystemException("資料已經完成確認送出,無法再更改");

                                #region //判斷是否已經有上傳過資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM PDM.DfmQiProcess
                                        WHERE DfmQiId = @DfmQiId";
                                dynamicParameters.Add("DfmQiId", DfmQiId);

                                var resultDfmQP = sqlConnection.Query(sql, dynamicParameters);
                                if (resultDfmQP.Count() > 0) throw new SystemException("此紀錄已上傳，無法刪除暫存資料!!");
                                #endregion

                                #region //更新PDM.DfmQuotationItem SpreadsheetData
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE PDM.DfmQuotationItem SET
                                        QiProcessSpreadsheetData = null,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE DfmQiId = @DfmQiId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        DfmQiId
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                                break;
                            default:
                                throw new SystemException("資料類別異常，請重新確認!!");
                                break;
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

        #endregion

        #region //Transaction
        public string TxQuotationCalculation(int DfmId)
        {
            if (DfmId == -1) throw new Exception("DfmId 不可為空");

            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int rowsAffected = 0;

                        #region default
                        DynamicParameters dynamicParameters = new DynamicParameters();
                        #endregion

                        #region //先清除之前之報價計算結果
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PDM.DfmQiSolutionProcess
                                      FROM PDM.DfmQiSolutionProcess a
                                           INNER JOIN PDM.DfmQiProcess b ON a.DfmQiProcessId = b.DfmQiProcessId
	                                       INNER JOIN PDM.DfmQuotationItem c ON b.DfmQiId = c.DfmQiId
	                                       INNER JOIN PDM.DesignForManufacturing d ON c.DfmId = d.DfmId
                                     WHERE d.DfmId = @DfmId";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            DfmId
                        });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PDM.DfmQiSolution 
                                  FROM PDM.DfmQiSolution a
                                       INNER JOIN PDM.DfmQuotationItem b ON a.DfmQiId = b.DfmQiId
	                                   INNER JOIN PDM.DesignForManufacturing c ON b.DfmId = c.DfmId
                                 WHERE c.DfmId = @DfmId";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            DfmId
                        });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //判斷DfmId是否存在, 且資料是否已維護完整
                        sql = @"SELECT a.DfmNo,b.DfmQuotationName,b.MfgFlag,b.StandardCost,
                                       b.MaterialStatus,b.DfmQiOSPStatus,b.DfmQiProcessStatus,
                                       a.RfqDetailId
                                  FROM PDM.DesignForManufacturing a
                                       INNER JOIN PDM.DfmQuotationItem b ON a.DfmId = b.DfmId
                                 WHERE a.DfmId = @DfmId
                                   AND (b.MaterialStatus = 'A' or b.DfmQiOSPStatus = 'A' or b.DfmQiProcessStatus = 'A')";
                        dynamicParameters.Add("DfmId", DfmId);
                        var unCompleteCnt = sqlConnection.Query(sql, dynamicParameters);
                        if (unCompleteCnt.Count() > 0) throw new SystemException("DFM還未製程或物料尚未維護完成");
                        #endregion

                        #region //取出報價項目
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT b.DfmQiId,a.DfmNo,b.DfmQuotationName,b.MfgFlag,b.StandardCost,
                                       b.MaterialStatus,b.DfmQiOSPStatus,b.DfmQiProcessStatus,
                                       a.RfqDetailId,b.StandardCost
                                  FROM PDM.DesignForManufacturing a
                                       INNER JOIN PDM.DfmQuotationItem b ON a.DfmId = b.DfmId
                                 WHERE a.DfmId = @DfmId";
                        dynamicParameters.Add("DfmId", DfmId);
                        var QuotationItemResult = sqlConnection.Query(sql, dynamicParameters);
                        int QuotationResult = 0;
                        foreach (var item in QuotationItemResult)
                        {
                            int RfqDetailId = Convert.ToInt32(item.RfqDetailId);
                            string MaterialStatus = item.MaterialStatus;
                            string DfmQiOSPStatus = item.DfmQiOSPStatus;
                            string DfmQiProcessStatus = item.DfmQiProcessStatus;
                            int DfmQiId = Convert.ToInt32(item.DfmQiId);
                            int StandardCost = Convert.ToInt32(item.StandardCost);

                            #region //找出報價方案
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.RfqLineSolutionId,a.SortNumber,a.SolutionQty
                                      FROM SCM.RfqLineSolution a
                                     WHERE a.RfqDetailId = @RfqDetailId";
                            dynamicParameters.Add("RfqDetailId", RfqDetailId);
                            var LineSolutionResult = sqlConnection.Query(sql, dynamicParameters);
                            if (LineSolutionResult.Count() <= 0) throw new SystemException("沒有報價方案資料!");
                            foreach (var item3 in LineSolutionResult)
                            {
                                int RfqLineSolutionId = Convert.ToInt32(item3.RfqLineSolutionId);
                                int SolutionSortNumber = Convert.ToInt32(item3.SortNumber);
                                int SolutionQty = Convert.ToInt32(item3.SolutionQty);
                                double ResourceAmt = 0;
                                double OverheadAmt = 0;
                                int DfmQiSolutionId = -1;

                                #region //取得 DfmQiSolutionId
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT DfmQiSolutionId
                                          FROM PDM.DfmQiSolution
                                         WHERE DfmQiId = @DfmQiId
                                           AND RfqLineSolutionId = @RfqLineSolutionId";
                                dynamicParameters.Add("DfmQiId", DfmQiId);
                                dynamicParameters.Add("RfqLineSolutionId", RfqLineSolutionId);
                                var DfmQiSolutionResult = sqlConnection.Query(sql, dynamicParameters);
                                if (DfmQiSolutionResult.Count() > 0)
                                {
                                    DfmQiSolutionId = Convert.ToInt32(sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).DfmQiSolutionId);
                                }
                                else
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO PDM.DfmQiSolution (DfmQiId, RfqLineSolutionId,StandardCost
                                                ,CreateDate, CreateBy
                                                ,LastModifiedDate, LastModifiedBy)
                                                OUTPUT INSERTED.DfmQiSolutionId
                                                VALUES (@DfmQiId, @RfqLineSolutionId,@StandardCost
                                                ,@CreateDate, @CreateBy
                                                ,@LastModifiedDate, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            DfmQiId,
                                            RfqLineSolutionId,
                                            StandardCost,
                                            CreateDate,
                                            CreateBy = CreateBy,
                                            LastModifiedDate,
                                            LastModifiedBy = CreateBy
                                        });
                                    var insertDfmQiSolution = sqlConnection.Query(sql, dynamicParameters);
                                    foreach (var itemQiSolution in insertDfmQiSolution)
                                    {
                                        DfmQiSolutionId = Convert.ToInt32(itemQiSolution.DfmQiSolutionId);
                                    }
                                }

                                #endregion

                                #region //計算物料費用
                                if (MaterialStatus.Equals("Y"))
                                {
                                    QuotationResult++;

                                    #region //更新物料成本
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE PDM.DfmQiSolution
                                               SET MaterialAmt = ISNULL(b.Amount,0),
                                                   LastModifiedDate = @LastModifiedDate,
                                                   LastModifiedBy = @LastModifiedBy
                                              FROM PDM.DfmQiSolution a
                                                   INNER JOIN (SELECT x.DfmQiId,
                                                                      SUM(ISNULL(x.UnitPrice,0) * ISNULL(x1.QuoteNum,1)) Amount
	                                                             FROM PDM.DfmQiMaterial x
                                                                      INNER JOIN PDM.DfmQuotationItem x1 ON x.DfmQiId = x1.DfmQiId
		                                                        GROUP BY x.DfmQiId) b ON a.DfmQiId = b.DfmQiId
                                             WHERE a.DfmQiId = @DfmQiId";
                                    dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        DfmQiId,
                                        LastModifiedDate = DateTime.Now,
                                        LastModifiedBy = CreateBy
                                    });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                                #endregion

                                #region //計算外包費用(先不執行)
                                #endregion

                                #region //計算人工與製造費用
                                if (DfmQiProcessStatus.Equals("Y"))
                                {
                                    QuotationResult++;
                                    #region //找出DfmQiProcess
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.DfmQiProcessId,a.DfmQiId,a.SortNumber,a.ParameterId,a.AllocationType,a.AllocationQty,
                                               a.ManHours,a.MachineHours,a.YieldRate,c.ResourceRate,c.OverheadRate,
	                                           b.DepartmentId,b.ProcessId
                                          FROM PDM.DfmQiProcess a
                                               INNER JOIN MES.ProcessParameter b ON a.ParameterId = b.ParameterId
	                                           LEFT JOIN BAS.DepartmentRate c ON b.DepartmentId = c.DepartmentId
                                         WHERE a.DfmQiId = @DfmQiId
                                           AND c.Status = 'A'
                                         ORDER BY a.SortNumber";
                                    dynamicParameters.Add("DfmQiId", DfmQiId);
                                    var DfmQiProcessResult = sqlConnection.Query(sql, dynamicParameters);
                                    foreach (var item2 in DfmQiProcessResult)
                                    {
                                        #region //取出QiProcess參數
                                        int DfmQiProcessId = Convert.ToInt32(item2.DfmQiProcessId);
                                        int SortNumber = Convert.ToInt32(item2.SortNumber);
                                        int AllocationType = Convert.ToInt32(item2.AllocationType);
                                        int AllocationQty = Convert.ToInt32(item2.AllocationQty);
                                        double ManHours = Convert.ToDouble(item2.ManHours);
                                        double MachineHours = Convert.ToDouble(item2.MachineHours);
                                        double ResourceRate = Convert.ToDouble(item2.ResourceRate);
                                        double OverheadRate = Convert.ToDouble(item2.OverheadRate);
                                        #endregion

                                        #region //計算人工費用與製造費用
                                        if (AllocationType.Equals(1))
                                        {
                                            ResourceAmt = ManHours * ResourceRate;
                                            OverheadAmt = MachineHours * OverheadRate;
                                        }
                                        else if (AllocationType.Equals(2))
                                        {
                                            ResourceAmt = ManHours * ResourceRate / AllocationQty;
                                            OverheadAmt = MachineHours * OverheadRate / AllocationQty;
                                        }
                                        else if (AllocationType.Equals(3))
                                        {
                                            ResourceAmt = ManHours * ResourceRate / SolutionQty;
                                            OverheadAmt = MachineHours * OverheadRate / SolutionQty;
                                        }
                                        #endregion

                                        #region //新增DfmQiSolutionProcess
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO PDM.DfmQiSolutionProcess (DfmQiProcessId,DfmQiSolutionId, ResourceAmt, OverheadAmt
                                                ,CreateDate, CreateBy
                                                ,LastModifiedDate, LastModifiedBy)
                                                OUTPUT INSERTED.DfmQiProcessId
                                                VALUES (@DfmQiProcessId, @DfmQiSolutionId, @ResourceAmt, @OverheadAmt
                                                ,@CreateDate, @CreateBy
                                                ,@LastModifiedDate, @LastModifiedBy)";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                DfmQiProcessId,
                                                DfmQiSolutionId,
                                                ResourceAmt,
                                                OverheadAmt,
                                                CreateDate,
                                                CreateBy = CreateBy,
                                                LastModifiedDate,
                                                LastModifiedBy = CreateBy
                                            });
                                        var insertQisProcess = sqlConnection.Query(sql, dynamicParameters);
                                        #endregion

                                        #region //更新DfmQiSolution的工和費資訊
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"UPDATE PDM.DfmQiSolution
                                                   SET ResourceAmt = ISNULL(ResourceAmt,0) + @ResourceAmt,
                                                       OverheadAmt = ISNULL(OverheadAmt,0) + @OverheadAmt,
                                                       LastModifiedDate = @LastModifiedDate,
                                                       LastModifiedBy = @LastModifiedBy
                                                 WHERE DfmQiSolutionId = @DfmQiSolutionId";
                                        dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            DfmQiSolutionId,
                                            ResourceAmt,
                                            OverheadAmt,
                                            LastModifiedDate = DateTime.Now,
                                            LastModifiedBy = CreateBy
                                        });

                                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                                        #endregion
                                    }
                                    #endregion
                                }
                                #endregion
                            }


                            #endregion


                        }

                        if (QuotationResult <= 0) throw new SystemException("沒有可計算之報價項目!");
                        #endregion

                        #region //將底階的DfmQiSolution Summary到第一階

                        #region //找出最大的QuotationLevl
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT MAX(b.QuotationLevel) QuotationLevel
                                  FROM PDM.DfmQiSolution a
                                       INNER JOIN PDM.DfmQuotationItem b ON a.DfmQiId = b.DfmQiId
                                 WHERE b.DfmId = @DfmId";
                        dynamicParameters.Add("DfmId", DfmId);
                        var QuotationLevelResult = sqlConnection.Query(sql, dynamicParameters);
                        if (QuotationLevelResult.Count() <= 0) throw new SystemException("找不到報價項目!");
                        int QuotationLevel = Convert.ToInt32(sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).QuotationLevel);

                        while(QuotationLevel > 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.RfqLineSolutionId,b.ParentDfmQiId,a.MaterialAmt,a.ResourceAmt, a.OverheadAmt, 
                                           a.StandardCost
                                      FROM PDM.DfmQiSolution a
                                           INNER JOIN PDM.DfmQuotationItem b ON a.DfmQiId = b.DfmQiId
                                     WHERE b.DfmId = @DfmId
                                       AND b.QuotationLevel = @QuotationLevel";
                            dynamicParameters.Add("DfmID", DfmId);
                            dynamicParameters.Add("QuotationLevel", QuotationLevel);

                            var QuotationDetailResult = sqlConnection.Query(sql, dynamicParameters);
                            foreach(var item in QuotationDetailResult)
                            {
                                double QuotationResourceAmt = Convert.ToDouble(item.ResourceAmt);
                                double QuotationOverheadAmt = Convert.ToDouble(item.OverheadAmt);
                                double QuotationMaterialAmt = Convert.ToDouble(item.MaterialAmt);
                                double QuotationStandardCost = Convert.ToDouble(item.StandardCost);
                                int RfqLineSolutionId = Convert.ToInt32(item.RfqLineSolutionId);
                                int ParentDfmQiId = Convert.ToInt32(item.ParentDfmQiId);

                                #region //更新上一層報價金額
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE PDM.DfmQiSolution
                                           SET MaterialAmt = MaterialAmt + @QuotationMaterialAmt,
                                               ResourceAmt = ResourceAmt + @QuotationResourceAmt,
                                               OverheadAmt = OverheadAmt + @QuotationOverheadAmt,
                                               StandardCost = StandardCost + @QuotationStandardCost
                                         WHERE DfmQiId = @ParentDfmQiId
                                           AND RfqLineSolutionId = @RfqLineSolutionId";
                                dynamicParameters.AddDynamicParams(
                                new
                                {
                                    QuotationMaterialAmt,
                                    QuotationResourceAmt,
                                    QuotationOverheadAmt,
                                    QuotationStandardCost,
                                    ParentDfmQiId,
                                    RfqLineSolutionId
                                });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }

                            QuotationLevel--;

                        }
                        #endregion

                        #region //更新RfqDetail.Status
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.RfqDetail
                                   SET Status = '5'
                                  FROM SCM.RfqDetail a
                                       INNER JOIN PDM.DesignForManufacturing b ON a.RfqDetailId = b.RfqDetailId
                                 WHERE b.DfmId = @DfmId";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            DfmId
                        });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //Upload
        #region //UploadDfmDetail -- 解析SpreadSheet及上傳Dfm詳細資料 -- Ann 2023-08-02
        public string UploadDfmDetail(int DfmId, string SpreadsheetJson, string SpreadsheetData)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region//檢核DFM【單頭】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.DfmId, b.DocType
                                FROM PDM.DesignForManufacturing a
                                INNER JOIN SCM.RfqDetail b on a.RfqDetailId = b.RfqDetailId
                                WHERE a.DfmId=@DfmId";
                        dynamicParameters.Add("DfmId", DfmId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【DFM】單頭查無資料,請重新確認");
                        foreach (var item in result)
                        {
                            if (item.DocType == "F") throw new SystemException("【DFM】RFQ單據已經失效,不能在異動更新");
                        }
                        #endregion

                        #region //先儲存目前Spreadsheet資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PDM.DesignForManufacturing SET
                                SpreadsheetData = @SpreadsheetData,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE DfmId = @DfmId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SpreadsheetData,
                                LastModifiedDate,
                                LastModifiedBy,
                                DfmId
                            });

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //先刪除目前所有DfmDetail
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE PDM.DfmDetail
                                WHERE DfmId = @DfmId";
                        dynamicParameters.Add("DfmId", DfmId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //解析Spreadsheet Data
                        var spreadsheetJson = JObject.Parse(SpreadsheetJson);
                        foreach (var item in spreadsheetJson["spreadsheetInfo"])
                        {
                            #region //確認報價項目資料是否正確
                            string DfmItemNo = item["DfmItemFullNo"].ToString().Split(' ')[0];

                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.DfmItemId
                                    FROM PDM.DfmItem a 
                                    WHERE a.DfmItemNo = @DfmItemNo";
                            dynamicParameters.Add("DfmItemNo", DfmItemNo);

                            var dfmItemNoResult = sqlConnection.Query(sql, dynamicParameters);

                            if (dfmItemNoResult.Count() <= 0) throw new SystemException("DFM報價項目資料錯誤!!");

                            int DfmItemId = -1;
                            foreach (var item2 in dfmItemNoResult)
                            {
                                DfmItemId = item2.DfmItemId;
                            }
                            #endregion

                            #region //INSERT PDM.DfmDetail
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO PDM.DfmDetail (DfmId, DfmItemId, Standard, FinalSpec, CustomerSpec, SuggestSpec, RdRemark, CustomerFeedback
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.DfmDetailId
                                    VALUES (@DfmId, @DfmItemId, @Standard, @FinalSpec, @CustomerSpec, @SuggestSpec, @RdRemark, @CustomerFeedback
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                            new
                            {
                                DfmId,
                                DfmItemId,
                                Standard = item["Standard"] != null ? item["Standard"].ToString() : null,
                                FinalSpec = item["FinalSpec"] != null ? item["FinalSpec"].ToString() : null,
                                CustomerSpec = item["CustomerSpec"] != null ? item["CustomerSpec"].ToString() : null,
                                SuggestSpec = item["SuggestSpec"] != null ? item["SuggestSpec"].ToString() : null,
                                RdRemark = item["RdRemark"] != null ? item["RdRemark"].ToString() : null,
                                CustomerFeedback = item["CustomerFeedback"] != null ? item["CustomerFeedback"].ToString() : null,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });

                            var insertResult = sqlConnection.Query(sql, dynamicParameters);
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

        #region //UploadDfmDetailTempSpreadsheetJson -- 暫存Dfm詳細資料SpreadsheetData -- Ann 2023-08-02
        public string UploadDfmDetailTempSpreadsheetJson(int DfmId, string SpreadsheetData)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region//檢核DFM【單頭】是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.DfmId, b.DocType
                                FROM PDM.DesignForManufacturing a
                                INNER JOIN SCM.RfqDetail b on a.RfqDetailId = b.RfqDetailId
                                WHERE a.DfmId=@DfmId";
                        dynamicParameters.Add("DfmId", DfmId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【DFM】單頭查無資料,請重新確認");
                        foreach (var item in result)
                        {
                            if (item.DocType == "F") throw new SystemException("【DFM】RFQ單據已經失效,不能在異動更新");
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PDM.DesignForManufacturing SET
                                SpreadsheetData = @SpreadsheetData,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE DfmId = @DfmId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SpreadsheetData,
                                LastModifiedDate,
                                LastModifiedBy,
                                DfmId
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
    }
}
