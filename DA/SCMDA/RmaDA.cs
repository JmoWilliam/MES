using Dapper;
using Helpers;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Transactions;
using System.Web;

namespace SCMDA
{
    public class RmaDA
    {
        public static string MainConnectionStrings = "";
        public static string ErpConnectionStrings = "";

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

        public RmaDA()
        {
            MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];

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
        #region //GetRandomRmaNo -- 取得隨機退貨單號 -- Ben Ma 2023.03.20
        public string GetRandomRmaNo()
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT CONVERT(int, RIGHT(ISNULL(MAX(RmaNo), '000'), 3)) + 1 CurrentNum
                            FROM SCM.ReturnMerchandiseAuthorization
                            WHERE RmaNo LIKE @RmaNo";
                    dynamicParameters.Add("RmaNo", string.Format("{0}{1}___", "S", DateTime.Now.ToString("yyyyMM")));
                    int currentNum = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).CurrentNum;

                    string rmaNo = string.Format("{0}{1}{2}", "S", DateTime.Now.ToString("yyyyMM"), string.Format("{0:000}", currentNum));

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = rmaNo
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

        #region //GetReturnMerchandiseAuthorization -- 取得退(換)貨資料 -- Ben Ma 2023.03.16
        public string GetReturnMerchandiseAuthorization(int RmaId, string RmaNo, string ErpNo, int CustomerId, string SearchKey
            , string ConfirmStatus, string StartDate, string EndDate
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                List<Rma> rmas = new List<Rma>();

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

                    sqlQuery.mainKey = "a.RmaId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.CompanyId, a.RmaNo, a.CustomerId, FORMAT(a.RmaDate, 'yyyy-MM-dd') RmaDate
                        , a.RmaType, a.RmaRemark, a.DocType, a.ConfirmStatus, a.ConfirmUserId, a.ErpPrefix, a.ErpNo
                        , a.TransferStatus, FORMAT(a.TransferDate, 'yyyy-MM-dd HH:mm') TransferDate, a.TransferMode
                        , ISNULL(b.CustomerNo, '') CustomerNo, ISNULL(b.CustomerShortName, '') CustomerShortName
                        , c.TypeName RmaTypeName
                        , ISNULL(d.UserNo, '') ConfirmUserNo, ISNULL(d.UserName,'') ConfirmUserName, ISNULL(d.Gender, '') ConfirmUserGender
                        , e.StatusName ConfirmName
                        , f.TypeName DocTypeName";
                    sqlQuery.mainTables =
                        @"FROM SCM.ReturnMerchandiseAuthorization a
                        LEFT JOIN SCM.Customer b ON a.CustomerId = b.CustomerId
                        INNER JOIN BAS.[Type] c ON a.RmaType = c.TypeNo AND c.TypeSchema = 'ReturnMerchandiseAuthorization.RmaType'
                        LEFT JOIN BAS.[User] d ON a.ConfirmUserId = d.UserId
                        INNER JOIN BAS.[Status] e ON e.StatusSchema = 'ConfirmStatus' AND e.StatusNo = a.ConfirmStatus
                        INNER JOIN BAS.[Type] f ON a.DocType = f.TypeNo AND f.TypeSchema = 'ReturnMerchandiseAuthorization.DocType'";
                    string queryTable =
                        @"FROM (
                            SELECT a.RmaId, a.CompanyId, a.RmaNo, a.CustomerId, a.RmaDate, a.ConfirmStatus, a.ErpPrefix, a.ErpNo
                            , (
                                SELECT y.MtlItemNo, y.MtlItemName
                                FROM SCM.RmaDetail z
                                INNER JOIN PDM.MtlItem y ON z.MtlItemId = y.MtlItemId
                                WHERE z.RmaId = a.RmaId
                                FOR JSON PATH, ROOT('data')
                            ) RmaDetail
                            FROM SCM.ReturnMerchandiseAuthorization a
                        ) a
                        OUTER APPLY (
                            SELECT TOP 1 x.MtlItemNo, x.MtlItemName
                            FROM OPENJSON(a.RmaDetail, '$.data')
                            WITH (
                                MtlItemNo NVARCHAR(40) N'$.MtlItemNo', 
                                MtlItemName NVARCHAR(120) N'$.MtlItemName'
                            ) x
                        ) b";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RmaId", @" AND a.RmaId = @RmaId", RmaId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RmaNo", @" AND a.RmaNo LIKE '%' + @RmaNo + '%'", RmaNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ErpNo", @" AND (a.ErpPrefix + '-' + a.ErpNo) LIKE '%' + @ErpNo + '%'", ErpNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerId", @" AND a.CustomerId = @CustomerId", CustomerId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey", @" AND (b.MtlItemNo LIKE '%' + @SearchKey + '%' OR b.MtlItemName LIKE '%' + @SearchKey + '%')", SearchKey);
                    if (ConfirmStatus.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ConfirmStatus", @" AND a.ConfirmStatus IN @ConfirmStatus", ConfirmStatus.Split(','));
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.RmaDate >= @StartDate ", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.RmaDate <= @EndDate ", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.RmaDate DESC, a.RmaNo DESC";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    rmas = BaseHelper.SqlQuery<Rma>(sqlConnection, dynamicParameters, sqlQuery);
                }

                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                {
                    List<ErpDocStatus> erpDocStatuses = new List<ErpDocStatus>();

                    #region //客供入料單 && 0成本入庫單
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TA001 ErpPrefix, TA002 ErpNo, TA006 DocStatus
                            FROM INVTA
                            WHERE (TA001 + '-' + TA002) IN @ErpFullNo
                            UNION ALL
                            SELECT TH001 ErpPrefix, TH002 ErpNo, TH020 DocStatus
                            FROM INVTH
                            WHERE (TH001 + '-' + TH002) IN @ErpFullNo";
                    dynamicParameters.Add("ErpFullNo", rmas.Select(x => x.ErpPrefix + '-' + x.ErpNo).ToArray());
                    erpDocStatuses = sqlConnection.Query<ErpDocStatus>(sql, dynamicParameters).ToList();

                    rmas = rmas.GroupJoin(erpDocStatuses, x => x.ErpPrefix + '-' + x.ErpNo, y => y.ErpPrefix + '-' + y.ErpNo, (x, y) => { x.DocStatus = y.FirstOrDefault()?.DocStatus ?? ""; return x; }).ToList();
                    #endregion
                }

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    data = rmas
                });
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

        #region //GetReturnMerchandiseAuthorizationDoc -- 取得退(換)貨單據資料 -- Ben Ma 2023.03.22
        public string GetReturnMerchandiseAuthorizationDoc(int RmaId, string Doc)
        {
            try
            {
                if (!Regex.IsMatch(Doc, "^(rma|temp-shipping-return|inventory-transaction-1106|inventory-transaction-1109)$", RegexOptions.IgnoreCase)) throw new SystemException("【單據種類】錯誤!");

                string companyNo = "", erpPrefix = "", erpNo = "", userName = "";
                switch (Doc)
                {
                    case "rma":
                        using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                        {
                            sql = @"SELECT a.RmaNo, FORMAT(a.RmaDate, 'yyyy-MM-dd') RmaDate, a.RmaType, a.RmaRemark
                            , ISNULL(b.CustomerShortName, '') Customer
                            , c.UserName ConfirmUser
                            , (
                                SELECT z.ItemName, z.RmaQty, z.RmaDesc
                                , ISNULL(x.CustomerMtlItemNo, y.MtlItemName) CustomerMtlItemNo
                                FROM SCM.RmaDetail z
                                INNER JOIN PDM.MtlItem y ON z.MtlItemId = y.MtlItemId
                                LEFT JOIN PDM.CustomerMtlItem x ON z.MtlItemId = x.MtlItemId AND a.CustomerId = x.CustomerId
                                WHERE z.RmaId = a.RmaId
                                FOR JSON PATH, ROOT('data')
                            ) RmaDetail
                            FROM SCM.ReturnMerchandiseAuthorization a
                            LEFT JOIN SCM.Customer b ON a.CustomerId = b.CustomerId
                            INNER JOIN BAS.[User] c ON a.ConfirmUserId = c.UserId
                            WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "RmaId", @" AND a.RmaId = @RmaId", RmaId);

                            var result = sqlConnection.Query(sql, dynamicParameters);

                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                data = result
                            });
                            #endregion
                        }
                        break;
                    case "temp-shipping-return":
                        using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                        {
                            #region //公司別資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.ErpNo, a.ErpDb
                                    FROM BAS.Company a
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            var resultCompany = sqlConnection.Query(sql, dynamicParameters);
                            if (resultCompany.Count() <= 0) throw new SystemException("公司別資料錯誤!");

                            foreach (var item in resultCompany)
                            {
                                ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                                companyNo = item.ErpNo;
                            }
                            #endregion

                            #region //退(換)貨單據資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.ErpPrefix, a.ErpNo, b.UserName
                                    FROM SCM.ReturnMerchandiseAuthorization a
                                    INNER JOIN BAS.[User] b ON a.ConfirmUserId = b.UserId
                                    WHERE RmaId = @RmaId";
                            dynamicParameters.Add("RmaId", RmaId);

                            var resultRma = sqlConnection.Query(sql, dynamicParameters);
                            if (resultRma.Count() <= 0) throw new SystemException("退(換)貨單據資料錯誤!");

                            foreach (var item in resultRma)
                            {
                                erpPrefix = item.ErpPrefix;
                                erpNo = item.ErpNo;
                                userName = item.UserName;
                            }
                            #endregion
                        }

                        using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //INVTH資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.TH001, a.TH002, '客戶' TH004, a.TH011, a.TH002, a.TH005 + ' ' + a.TH006 TH005
                                    , a.TH012, FORMAT(CAST(LTRIM(RTRIM(a.TH023)) as date), 'yyyy/MM/dd') TH023, a.TH015
                                    , FORMAT(a.TH025, 'P0') TH025, a.TH010, a.TH007 + ' ' + b.ME002 TH007, a.TH014
                                    , a.TH008 + ' ' + c.MF002 TH008, a.TH021, a.TH026, a.TH022, a.TH010
                                    FROM INVTH a
                                    INNER JOIN CMSME b ON a.TH007 = b.ME001
                                    INNER JOIN ADMMF c ON a.TH008 = c.MF001
                                    WHERE a.TH001 = @ErpPrefix
                                    AND a.TH002 = @ErpNo";
                            dynamicParameters.Add("ErpPrefix", erpPrefix);
                            dynamicParameters.Add("ErpNo", erpNo);

                            var resultInvth = sqlConnection.Query(sql, dynamicParameters);
                            if (resultInvth.Count() <= 0) throw new SystemException("ERP暫出/入歸還單資料錯誤!");
                            #endregion

                            #region //INVTI資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.TI003 Seq, a.TI004 MtlItemNo, a.TI005 MtlItemName, a.TI006 MtlItemSpec
                                    , a.TI007 OutInvNo, a.TI008 InInvNo, a.TI009 Qty, a.TI010 Uom
                                    , a.TI012 UnitPrice, a.TI013 Amount, a.TI014 + '-' + a.TI015 + '-' + a.TI016 SourceNote
                                    , a.TI021 Remark
                                    FROM INVTI a
                                    WHERE a.TI001 = @ErpPrefix
                                    AND a.TI002 = @ErpNo";
                            dynamicParameters.Add("ErpPrefix", erpPrefix);
                            dynamicParameters.Add("ErpNo", erpNo);

                            var resultInvti = sqlConnection.Query(sql, dynamicParameters);
                            if (resultInvti.Count() <= 0) throw new SystemException("ERP暫出/入歸還單單身資料錯誤!");
                            #endregion

                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                data = resultInvth,
                                dataDetail = resultInvti,
                                user = userName
                            });
                            #endregion
                        }
                        break;
                    case "inventory-transaction-1106":
                    case "inventory-transaction-1109":
                        using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                        {
                            #region //公司別資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.ErpNo, a.ErpDb
                                    FROM BAS.Company a
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            var resultCompany = sqlConnection.Query(sql, dynamicParameters);
                            if (resultCompany.Count() <= 0) throw new SystemException("公司別資料錯誤!");

                            foreach (var item in resultCompany)
                            {
                                ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                                companyNo = item.ErpNo;
                            }
                            #endregion

                            #region //退(換)貨單據資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.ErpPrefix, a.ErpNo, b.UserName
                                    FROM SCM.ReturnMerchandiseAuthorization a
                                    INNER JOIN BAS.[User] b ON a.ConfirmUserId = b.UserId
                                    WHERE RmaId = @RmaId";
                            dynamicParameters.Add("RmaId", RmaId);

                            var resultRma = sqlConnection.Query(sql, dynamicParameters);
                            if (resultRma.Count() <= 0) throw new SystemException("退(換)貨單據資料錯誤!");

                            foreach (var item in resultRma)
                            {
                                erpPrefix = item.ErpPrefix;
                                erpNo = item.ErpNo;
                                userName = item.UserName;
                            }
                            #endregion
                        }

                        using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //INVTA資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.TA001, a.TA002, b.ME002 TA004, a.TA005, a.TA006, a.TA010, a.TA011
                                    , FORMAT(CAST(LTRIM(RTRIM(a.TA014)) as date), 'yyyy/MM/dd') TA014, a.CREATOR
                                    FROM INVTA a
                                    INNER JOIN CMSME b ON a.TA004 = b.ME001
                                    WHERE a.TA001 = @ErpPrefix
                                    AND a.TA002 = @ErpNo";
                            dynamicParameters.Add("ErpPrefix", erpPrefix);
                            dynamicParameters.Add("ErpNo", erpNo);

                            var resultInvta = sqlConnection.Query(sql, dynamicParameters);
                            if (resultInvta.Count() <= 0) throw new SystemException("ERP庫存異動單資料錯誤!");
                            #endregion

                            #region //INVTB資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.TB003 Seq, a.TB004 MtlItemNo, a.TB005 MtlItemName, a.TB006 MtlItemSpec
                                    , a.TB008 Uom, a.TB012 InvNo, b.MC002 InvName, a.TB007 Qty, a.TB017 Remark
                                    FROM INVTB a
                                    INNER JOIN CMSMC b ON a.TB012 = b.MC001
                                    WHERE a.TB001 = @ErpPrefix
                                    AND a.TB002 = @ErpNo
                                    ORDER BY a.TB003";
                            dynamicParameters.Add("ErpPrefix", erpPrefix);
                            dynamicParameters.Add("ErpNo", erpNo);

                            var resultInvtb = sqlConnection.Query(sql, dynamicParameters);
                            if (resultInvtb.Count() <= 0) throw new SystemException("ERP庫存異動單單身資料錯誤!");
                            #endregion

                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                data = resultInvta,
                                dataDetail = resultInvtb,
                                user = userName
                            });
                            #endregion
                        }
                        break;
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //GetRmaDetail -- 取得退(換)貨項目資料 -- Ben Ma 2023.03.20
        public string GetRmaDetail(int RmaDetailId, int RmaId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.RmaDetailId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.RmaId, a.MtlItemId, a.ItemName, a.RmaQty, a.ItemType, ISNULL(a.FreebieOrSpareQty, '0') FreebieOrSpareQty, a.RmaDesc
                        , ISNULL(a.TsnDetailId, -1) TsnDetailId, ISNULL(a.TargetInventory, -1) TargetInventory
                        , b.MtlItemNo, b.MtlItemName, b.MtlItemSpec
                        , ISNULL(c.UomNo, '') UomNo
                        , d.ConfirmStatus
                        , e.TsnErpFullNo, e.SoErpFullNo, ISNULL(e.LastTsnQty, 0) LastTsnQty
                        , CASE 
                            WHEN a.TargetInventory IS NULL THEN '預設'
                            ELSE f.InventoryNo + ' ' + f.InventoryName
                        END TargetInventoryNo
                        , ISNULL(g.TypeName, '') ItemTypeName";
                    sqlQuery.mainTables =
                        @"FROM SCM.RmaDetail a
                        INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
                        LEFT JOIN PDM.UnitOfMeasure c ON b.InventoryUomId = c.UomId
                        INNER JOIN SCM.ReturnMerchandiseAuthorization d ON a.RmaId = d.RmaId
                        OUTER APPLY (
                            SELECT (eb.TsnErpPrefix + '-' + eb.TsnErpNo + '-' + ea.TsnSequence) TsnErpFullNo
                            , (ed.SoErpPrefix + '-' + ed.SoErpNo + '-' + ec.SoSequence) SoErpFullNo
                            , ef.UomNo, (eg.TotalTsn - eg.TotalSale - eg.TotalReturn) LastTsnQty
                            FROM SCM.TsnDetail ea
                            INNER JOIN SCM.TempShippingNote eb ON ea.TsnId = eb.TsnId
                            LEFT JOIN SCM.SoDetail ec ON ea.SoDetailId = ec.SoDetailId
                            INNER JOIN SCM.SaleOrder ed ON ec.SoId = ed.SoId
                            INNER JOIN PDM.MtlItem ee ON ea.MtlItemId = ee.MtlItemId
                            INNER JOIN PDM.UnitOfMeasure ef ON ee.InventoryUomId = ef.UomId
                            OUTER APPLY (
                                SELECT SUM(ega.TsnQty) TotalTsn, SUM(ega.SaleQty) TotalSale, SUM(ega.ReturnQty) TotalReturn
                                FROM SCM.TsnDetail ega
                                WHERE ega.SoDetailId = ec.SoDetailId
                                AND ega.TsnDetailId = ea.TsnDetailId
                                AND ega.ConfirmStatus = @ConfirmStatus
                                AND ega.ClosureStatus = @ClosureStatus
                            ) eg
                            WHERE ea.TsnDetailId = a.TsnDetailId
                        ) e
                        LEFT JOIN SCM.Inventory f ON a.TargetInventory = f.InventoryId
                        LEFT JOIN BAS.[Type] g ON a.ItemType = g.TypeNo AND g.TypeSchema = 'SoDetail.ProductType'";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    dynamicParameters.Add("ConfirmStatus", "Y");
                    dynamicParameters.Add("ClosureStatus", "N");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RmaDetailId", @" AND a.RmaDetailId = @RmaDetailId", RmaDetailId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "RmaId", @" AND a.RmaId = @RmaId", RmaId);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.RmaDetailId";
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

        #region //GetTempShippingNote -- 取得暫出單資料 -- Ben Ma 2023.03.30
        public string GetTempShippingNote(int RmaId, int CustomerId, string SearchKey
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.TsnDetailId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.TsnSequence, a.TsnQty, h.TypeName ProductTypeName
                        , b.TsnErpPrefix, b.TsnErpNo, (b.TsnErpPrefix + '-' + b.TsnErpNo + '-' + a.TsnSequence) TsnErpFullNo
                        , ISNULL(c.SoSequence, '') SoSequence
                        , ISNULL(d.SoErpPrefix, '') SoErpPrefix, ISNULL(d.SoErpNo, '') SoErpNo, (d.SoErpPrefix + '-' + d.SoErpNo + '-' + c.SoSequence) SoErpFullNo
                        , e.MtlItemNo, ISNULL(a.TsnMtlItemName, ISNULL(c.SoMtlItemName, e.MtlItemName)) MtlItemName
                        , ISNULL(a.TsnMtlItemSpec, ISNULL(c.SoMtlItemSpec, e.MtlItemSpec)) MtlItemSpec
                        , f.UomNo
                        , (g.TotalTsn - g.TotalSale - g.TotalReturn) LastTsnQty
                        , CASE WHEN a.ProductType = '1' THEN TsnFreebieQty ELSE TsnSpareQty END FreebieOrSpareQty
                        , (ISNULL(g.TsnFreebieQty, 0) - ISNULL(g.TsnReturnFSFreebieQty, 0)) TsnFreebieQty
                        , (ISNULL(g.TsnSpareQty, 0) - ISNULL(g.TsnReturnFSSpareQty, 0)) TsnSpareQty";
                    sqlQuery.mainTables =
                        @"FROM SCM.TsnDetail a
                        INNER JOIN SCM.TempShippingNote b ON a.TsnId = b.TsnId
                        LEFT JOIN SCM.SoDetail c ON a.SoDetailId = c.SoDetailId
                        INNER JOIN SCM.SaleOrder d ON c.SoId = d.SoId
                        INNER JOIN PDM.MtlItem e ON a.MtlItemId = e.MtlItemId
                        INNER JOIN PDM.UnitOfMeasure f ON e.InventoryUomId = f.UomId
                        OUTER APPLY (
                            SELECT SUM(ga.TsnQty) TotalTsn
                            , CASE WHEN ga.ProductType = '1' THEN SUM(ga.FreebieOrSpareQty) END TsnFreebieQty
                            , CASE WHEN ga.ProductType = '2' THEN SUM(ga.FreebieOrSpareQty) END TsnSpareQty
                            , SUM(ga.SaleQty) TotalSale
                            , CASE WHEN ga.ProductType = '1' THEN SUM(ga.SaleFSQty) END TsnSaleFSFreebieQty
                            , CASE WHEN ga.ProductType = '2' THEN SUM(ga.SaleFSQty) END TsnSaleFSSpareQty
                            , SUM(ga.ReturnQty) TotalReturn
                            , CASE WHEN ga.ProductType = '1' THEN SUM(ga.ReturnFSQty) END TsnReturnFSFreebieQty
                            , CASE WHEN ga.ProductType = '2' THEN SUM(ga.ReturnFSQty) END TsnReturnFSSpareQty
                            FROM SCM.TsnDetail ga
                            WHERE ga.SoDetailId = c.SoDetailId
                            AND ga.TsnDetailId = a.TsnDetailId
                            AND ga.ConfirmStatus = @ConfirmStatus
                            AND ga.ClosureStatus = @ClosureStatus
                            GROUP BY ga.ProductType
                        ) g
                        INNER JOIN BAS.[Type] h ON a.ProductType = h.TypeNo AND h.TypeSchema = 'SoDetail.ProductType'";
                    sqlQuery.auxTables = "";
                    string queryCondition =
                        @" AND (g.TotalTsn - g.TotalSale - g.TotalReturn) > 0
                        AND a.ConfirmStatus = @ConfirmStatus
                        AND b.ConfirmStatus = @ConfirmStatus
                        AND c.ConfirmStatus = @ConfirmStatus
                        AND d.ConfirmStatus = @ConfirmStatus
                        AND a.ClosureStatus = @ClosureStatus
                        AND b.ObjectCustomer = @CustomerId";
                    dynamicParameters.Add("ConfirmStatus", "Y");
                    dynamicParameters.Add("ClosureStatus", "N");
                    dynamicParameters.Add("CustomerId", CustomerId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey",
                        @" AND (
                            (b.TsnErpPrefix + '-' + b.TsnErpNo) LIKE '%' + @SearchKey + '%'
                            OR (d.SoErpPrefix + '-' + d.SoErpNo) LIKE '%' + @SearchKey + '%'
                            OR e.MtlItemNo LIKE '%' + @SearchKey + '%'
                            OR a.TsnMtlItemName LIKE '%' + @SearchKey + '%'
                            OR c.SoMtlItemName LIKE '%' + @SearchKey + '%'
                            OR e.MtlItemName LIKE '%' + @SearchKey + '%'
                        )", SearchKey);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "b.DocDate DESC, b.TsnErpPrefix, a.TsnSequence";
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

        #region //GetRmaTransferLog -- 取得退(換)貨拋轉單據紀錄資料 -- Ben Ma 2023.04.13
        public string GetRmaTransferLog(int RmaId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.ErpPrefix + '-' + a.ErpNo ErpFullNo
                            , FORMAT(a.CreateDate, 'yyyy-MM-dd HH:mm') CreateDate
                            , a.[Status], b.StatusName
                            FROM SCM.RmaTransferLog a
                            INNER JOIN BAS.[Status] b ON a.[Status] = b.StatusNo AND b.StatusSchema = 'Status'
                            WHERE RmaId = @RmaId
                            ORDER BY a.CreateDate DESC";
                    dynamicParameters.Add("RmaId", RmaId);

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
        #region //AddReturnMerchandiseAuthorization -- 退(換)貨資料新增 -- Ben Ma 2023.03.20
        public string AddReturnMerchandiseAuthorization(string RmaNo, int CustomerId, string RmaDate, string RmaType
            , string RmaRemark, string DocType)
        {
            try
            {
                if (RmaNo.Length <= 0) throw new SystemException("【退貨單號】不能為空!");
                if (RmaNo.Length != 10) throw new SystemException("【退貨單號】長度錯誤!");
                if (!DateTime.TryParse(RmaDate, out DateTime tempDate)) throw new SystemException("【退貨日期】格式錯誤!");
                if (RmaType.Length <= 0) throw new SystemException("【退貨種類】不能為空!");
                if (DocType.Length <= 0) throw new SystemException("【單據類型】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷退貨單號是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.ReturnMerchandiseAuthorization
                                WHERE RmaNo = @RmaNo";
                        dynamicParameters.Add("RmaNo", RmaNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【退貨單號】重複，請重新輸入!");
                        #endregion

                        #region //判斷客戶資料是否正確
                        if (CustomerId > 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM SCM.Customer
                                    WHERE CustomerId = @CustomerId";
                            dynamicParameters.Add("CustomerId", CustomerId);

                            var result2 = sqlConnection.Query(sql, dynamicParameters);
                            if (result2.Count() <= 0) throw new SystemException("客戶資料錯誤!");
                        }
                        #endregion

                        #region //判斷退貨種類資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Type]
                                WHERE TypeSchema = @TypeSchema
                                AND TypeNo = @TypeNo";
                        dynamicParameters.Add("TypeSchema", "ReturnMerchandiseAuthorization.RmaType");
                        dynamicParameters.Add("TypeNo", RmaType);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() <= 0) throw new SystemException("退貨種類資料錯誤!");
                        #endregion

                        #region //判斷單據類型資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Type]
                                WHERE TypeSchema = @TypeSchema
                                AND TypeNo = @TypeNo";
                        dynamicParameters.Add("TypeSchema", "ReturnMerchandiseAuthorization.DocType");
                        dynamicParameters.Add("TypeNo", DocType);

                        var result4 = sqlConnection.Query(sql, dynamicParameters);
                        if (result4.Count() <= 0) throw new SystemException("單據類型資料錯誤!");

                        #region //依照單據類別增加判斷
                        switch (DocType)
                        {
                            case "A": //暫出歸還單
                                if (CustomerId <= 0) throw new SystemException("【客戶】不能為空!");
                                break;
                            default:
                                break;
                        }
                        #endregion
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.ReturnMerchandiseAuthorization (CompanyId, RmaNo, CustomerId
                                , RmaDate, RmaType, RmaRemark, DocType, ConfirmStatus, TransferStatus, TransferMode
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.RmaId, INSERTED.TransferStatus, INSERTED.TransferMode
                                VALUES (@CompanyId, @RmaNo, @CustomerId
                                , @RmaDate, @RmaType, @RmaRemark, @DocType, @ConfirmStatus, @TransferStatus, @TransferMode
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                RmaNo,
                                CustomerId = CustomerId <= 0 ? (int?)null : CustomerId,
                                RmaDate,
                                RmaType,
                                RmaRemark,
                                DocType,
                                ConfirmStatus = "N",
                                TransferStatus = "N",
                                TransferMode = "N",
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

        #region //AddRmaDetail -- 退(換)貨項目新增 -- Ben Ma 2023.03.20
        public string AddRmaDetail(int RmaId, int MtlItemId, string ItemName
            , int RmaQty, string ItemType, float FreebieOrSpareQty, string RmaDesc, int TargetInventory)
        {
            try
            {
                if (MtlItemId <= 0) throw new SystemException("【品號】不能為空!");
                if (ItemName.Length > 50) throw new SystemException("【部品名稱】長度錯誤!");
                if (RmaQty <= 0) throw new SystemException("【正常品數量】不能為空!");
                //if (ItemType.Length > 0) throw new SystemException("【類型】不能為空!");
                if (FreebieOrSpareQty < 0) throw new SystemException("【贈/備品】不能為空!");
                if (RmaDesc.Length > 200) throw new SystemException("【不良項目描述】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷退(換)貨資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 DocType, ConfirmStatus
                                FROM SCM.ReturnMerchandiseAuthorization
                                WHERE RmaId = @RmaId";
                        dynamicParameters.Add("RmaId", RmaId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("退(換)貨資料錯誤!");

                        #region //判斷單據狀態
                        string docType = "", confirmStatus = "";
                        foreach (var item in result)
                        {
                            docType = item.DocType;
                            confirmStatus = item.ConfirmStatus;
                        }

                        if (docType == "A") throw new SystemException("暫出歸還單請使用導入暫出單新增!");
                        if (confirmStatus == "Y") throw new SystemException("退貨單已經確認，無法新增!");
                        #endregion
                        #endregion

                        #region //判斷品號資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.MtlItem
                                WHERE MtlItemId = @MtlItemId";
                        dynamicParameters.Add("MtlItemId", MtlItemId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("品號資料錯誤!");
                        #endregion

                        #region //判斷目標庫別資料是否正確
                        if (TargetInventory > 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM SCM.Inventory
                                    WHERE InventoryId = @InventoryId";
                            dynamicParameters.Add("InventoryId", TargetInventory);

                            var result3 = sqlConnection.Query(sql, dynamicParameters);
                            if (result3.Count() <= 0) throw new SystemException("目標庫別資料錯誤!");
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.RmaDetail (RmaId, MtlItemId, ItemName, RmaQty, ItemType, FreebieOrSpareQty
                                , RmaDesc, TargetInventory
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.RmaDetailId
                                VALUES (@RmaId, @MtlItemId, @ItemName, @RmaQty, @ItemType, @FreebieOrSpareQty
                                , @RmaDesc, @TargetInventory
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                RmaId,
                                MtlItemId,
                                ItemName,
                                RmaQty,
                                ItemType,
                                FreebieOrSpareQty,
                                RmaDesc,
                                TargetInventory = TargetInventory <= 0 ? (int?)null : TargetInventory,
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

        #region //AddRmaDetailByTsn -- 退(換)貨項目【暫出單】新增 -- Ben Ma 2023.03.30
        public string AddRmaDetailByTsn(int RmaId, string TsnDetail)
        {
            try
            {
                if (!TsnDetail.TryParseJson(out JObject tempJObject)) throw new SystemException("暫出單格式錯誤");

                JObject tsnJson = JObject.Parse(TsnDetail);
                if (!tsnJson.ContainsKey("data")) throw new SystemException("暫出單資料錯誤");

                JToken tsnData = tsnJson["data"];
                if (tsnData.Count() < 0) throw new SystemException("查無暫出單內容");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷退(換)貨資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.DocType, a.ConfirmStatus, b.CustomerId
                                FROM SCM.ReturnMerchandiseAuthorization a
                                INNER JOIN SCM.Customer b ON a.CustomerId = b.CustomerId
                                WHERE RmaId = @RmaId";
                        dynamicParameters.Add("RmaId", RmaId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("退(換)貨資料錯誤!");

                        #region //判斷單據狀態
                        string docType = "", confirmStatus = "";
                        int customerId = -1;
                        foreach (var item in result)
                        {
                            docType = item.DocType;
                            confirmStatus = item.ConfirmStatus;
                            customerId = Convert.ToInt32(item.CustomerId);
                        }

                        if (docType != "A") throw new SystemException("開立單據類型錯誤，無法新增!");
                        if (confirmStatus == "Y") throw new SystemException("退貨單已經確認，無法新增!");
                        #endregion
                        #endregion

                        int rowsAffected = 0;
                        for (int i = 0; i < tsnData.Count(); i++)
                        {
                            int tsnDetailId = Convert.ToInt32(tsnData[i]["tsnDetailId"]);
                            double freebieOrSpareQty = Convert.ToDouble(tsnData[i]["freebieOrSpareQty"]);
                            double rmaQty = Convert.ToDouble(tsnData[i]["rmaQty"]);

                            #region //判斷暫出單資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.MtlItemId, (d.TotalTsn - d.TotalSale - d.TotalReturn) LastTsnQty, a.ProductType
                                    , CASE WHEN a.ProductType = '1' THEN TsnFreebieQty ELSE TsnSpareQty END FreebieOrSpareQty
                                    , (ISNULL(d.TsnFreebieQty, 0) - ISNULL(d.TsnReturnFSFreebieQty, 0)) TsnFreebieQty
                                    , (ISNULL(d.TsnSpareQty, 0) - ISNULL(d.TsnReturnFSSpareQty, 0)) TsnSpareQty, a.TsnRemark
                                    FROM SCM.TsnDetail a
                                    INNER JOIN SCM.TempShippingNote b ON a.TsnId = b.TsnId
                                    LEFT JOIN SCM.SoDetail c ON a.SoDetailId = c.SoDetailId
                                    OUTER APPLY (
                                        SELECT da.ProductType, SUM(da.TsnQty) TotalTsn
                                        , CASE WHEN da.ProductType = '1' THEN SUM(da.FreebieOrSpareQty) END TsnFreebieQty
                                        , CASE WHEN da.ProductType = '2' THEN SUM(da.FreebieOrSpareQty) END TsnSpareQty
                                        , SUM(da.SaleQty) TotalSale
                                        , CASE WHEN da.ProductType = '1' THEN SUM(da.SaleFSQty) END TsnSaleFSFreebieQty
                                        , CASE WHEN da.ProductType = '2' THEN SUM(da.SaleFSQty) END TsnSaleFSSpareQty
                                        , SUM(da.ReturnQty) TotalReturn
                                        , CASE WHEN da.ProductType = '1' THEN SUM(da.ReturnFSQty) END TsnReturnFSFreebieQty
                                        , CASE WHEN da.ProductType = '2' THEN SUM(da.ReturnFSQty) END TsnReturnFSSpareQty
                                        FROM SCM.TsnDetail da
                                        WHERE da.SoDetailId = c.SoDetailId
                                        AND da.TsnDetailId = a.TsnDetailId
                                        AND da.ConfirmStatus = @ConfirmStatus
                                        GROUP BY da.ProductType
                                    ) d
                                    WHERE a.TsnDetailId = @TsnDetailId
                                    AND a.ConfirmStatus = @ConfirmStatus
                                    AND b.ConfirmStatus = @ConfirmStatus
                                    AND b.ObjectCustomer = @CustomerId";
                            dynamicParameters.Add("TsnDetailId", tsnDetailId);
                            dynamicParameters.Add("ConfirmStatus", "Y");
                            dynamicParameters.Add("CustomerId", customerId);

                            var resultTsn = sqlConnection.Query(sql, dynamicParameters);
                            if (resultTsn.Count() <= 0) throw new SystemException("暫出單資料錯誤!");

                            int mtlItemId = -1;
                            double lastTsnQty = 0, lastFSQty = 0;
                            string tsnRemark = "", productType = "";
                            foreach (var item in resultTsn)
                            {
                                mtlItemId = Convert.ToInt32(item.MtlItemId);
                                lastTsnQty = Convert.ToDouble(item.LastTsnQty);
                                lastFSQty = Convert.ToDouble(item.FreebieOrSpareQty);
                                tsnRemark = item.TsnRemark;
                                productType = item.ProductType;
                            }
                            if (rmaQty > lastTsnQty) throw new SystemException("暫出【正常品】剩餘數量不足!");
                            if (freebieOrSpareQty > lastFSQty) throw new SystemException("暫出【贈/備品】剩餘數量不足!");
                            #endregion

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO SCM.RmaDetail (RmaId, MtlItemId, ItemName, RmaQty, ItemType, FreebieOrSpareQty
                                    , RmaDesc, TsnDetailId
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.RmaDetailId
                                    VALUES (@RmaId, @MtlItemId, @ItemName, @RmaQty, @ItemType, @FreebieOrSpareQty
                                    , @RmaDesc, @TsnDetailId
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    RmaId,
                                    MtlItemId = mtlItemId,
                                    ItemName = "",
                                    RmaQty = rmaQty,
                                    ItemType = productType,
                                    FreebieOrSpareQty = freebieOrSpareQty,
                                    RmaDesc = tsnRemark,
                                    TsnDetailId = tsnDetailId,
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
        #endregion

        #region //Update
        #region //UpdateReturnMerchandiseAuthorization -- 退(換)貨資料更新 -- Ben Ma 2023.03.20
        public string UpdateReturnMerchandiseAuthorization(int RmaId, int CustomerId, string RmaDate
            , string RmaType, string RmaRemark, string DocType, string TransferMode)
        {
            try
            {
                if (!DateTime.TryParse(RmaDate, out DateTime tempDate)) throw new SystemException("【退貨日期】格式錯誤!");
                if (RmaType.Length <= 0) throw new SystemException("【退貨種類】不能為空!");
                if (DocType.Length <= 0) throw new SystemException("【單據類型】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷退(換)貨資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 ISNULL(CustomerId, -1) CustomerId, DocType, ConfirmStatus, TransferStatus
                                FROM SCM.ReturnMerchandiseAuthorization
                                WHERE RmaId = @RmaId";
                        dynamicParameters.Add("RmaId", RmaId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("退(換)貨資料錯誤!");

                        #region //判斷確認過無法再修改
                        int originalCustomerId = -1;
                        string originalDocType = "", confirmStatus = "", transferStatus = "";
                        foreach (var item in result)
                        {
                            originalCustomerId = Convert.ToInt32(item.CustomerId);
                            originalDocType = item.DocType;
                            confirmStatus = item.ConfirmStatus;
                            transferStatus = item.TransferStatus;
                        }

                        if (confirmStatus == "Y") throw new SystemException("退貨單已經確認，無法修改!");
                        if (transferStatus == "Y")
                        {
                            if (TransferMode.Length <= 0) throw new SystemException("【拋轉模式】不能為空!");
                            if (originalCustomerId != CustomerId) throw new SystemException("已拋轉單據無法更換【客戶】!");
                            if (originalDocType != DocType) throw new SystemException("已拋轉單據無法更換【開立單據類型】!");
                        }
                        #endregion
                        #endregion

                        #region //判斷客戶資料是否正確
                        if (CustomerId > 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM SCM.Customer
                                    WHERE CustomerId = @CustomerId";
                            dynamicParameters.Add("CustomerId", CustomerId);

                            var result2 = sqlConnection.Query(sql, dynamicParameters);
                            if (result2.Count() <= 0) throw new SystemException("客戶資料錯誤!");
                        }
                        #endregion

                        #region //判斷退貨種類資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Type]
                                WHERE TypeSchema = @TypeSchema
                                AND TypeNo = @TypeNo";
                        dynamicParameters.Add("TypeSchema", "ReturnMerchandiseAuthorization.RmaType");
                        dynamicParameters.Add("TypeNo", RmaType);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() <= 0) throw new SystemException("退貨種類資料錯誤!");
                        #endregion

                        #region //判斷單據類型資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Type]
                                WHERE TypeSchema = @TypeSchema
                                AND TypeNo = @TypeNo";
                        dynamicParameters.Add("TypeSchema", "ReturnMerchandiseAuthorization.DocType");
                        dynamicParameters.Add("TypeNo", DocType);

                        var result4 = sqlConnection.Query(sql, dynamicParameters);
                        if (result4.Count() <= 0) throw new SystemException("單據類型資料錯誤!");
                        #endregion

                        #region //判斷拋轉模式資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM BAS.[Type]
                                WHERE TypeSchema = @TypeSchema
                                AND TypeNo = @TypeNo";
                        dynamicParameters.Add("TypeSchema", "ReturnMerchandiseAuthorization.TransferMode");
                        dynamicParameters.Add("TypeNo", TransferMode);

                        var result5 = sqlConnection.Query(sql, dynamicParameters);
                        if (result5.Count() <= 0) throw new SystemException("拋轉模式資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.ReturnMerchandiseAuthorization SET
                                CustomerId = @CustomerId,
                                RmaDate = @RmaDate,
                                RmaType = @RmaType,
                                RmaRemark = @RmaRemark,
                                DocType = @DocType,
                                TransferMode = @TransferMode,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RmaId = @RmaId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CustomerId = CustomerId <= 0 ? (int?)null : CustomerId,
                                RmaDate,
                                RmaType,
                                RmaRemark,
                                DocType,
                                TransferMode,
                                LastModifiedDate,
                                LastModifiedBy,
                                RmaId
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

        #region //UpdateReturnMerchandiseAuthorizationConfirm -- 退(換)貨資料確認 -- Ben Ma 2023.03.22
        public string UpdateReturnMerchandiseAuthorizationConfirm(int RmaId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    string companyNo = "", departmentNo = "", userNo = "", userName = "", docType = "", erpPrefix = "", erpNo = "", docDate = ""
                        , transferStatus = "", transferMode = "";

                    string dateNow = DateTime.Now.ToString("yyyyMMdd");
                    string timeNow = DateTime.Now.ToString("HH:mm:ss");

                    List<ErpDocStatus> erpDocStatuses = new List<ErpDocStatus>();

                    List<INVTA> iNVTAs = new List<INVTA>();
                    List<INVTB> iNVTBs = new List<INVTB>();

                    List<INVTH> iNVTHs = new List<INVTH>();
                    List<INVTI> iNVTIs = new List<INVTI>();

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //公司別資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var resultCompany = sqlConnection.Query(sql, dynamicParameters);
                        if (resultCompany.Count() <= 0) throw new SystemException("公司別資料錯誤!");

                        foreach (var item in resultCompany)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            companyNo = item.ErpNo;
                        }
                        #endregion

                        #region //使用者資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserNo, a.UserName, b.DepartmentNo
                                FROM BAS.[User] a
                                INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                                WHERE a.UserId = @userId";
                        dynamicParameters.Add("userId", CurrentUser);

                        var resultUser = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUser.Count() <= 0) throw new SystemException("使用者資料錯誤!");

                        foreach (var item in resultUser)
                        {
                            userNo = item.UserNo;
                            userName = item.UserName;
                            departmentNo = item.DepartmentNo;
                        }
                        #endregion

                        #region //判斷退(換)貨資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 DocType, ConfirmStatus, TransferStatus, TransferMode, ErpPrefix, ErpNo
                                FROM SCM.ReturnMerchandiseAuthorization
                                WHERE RmaId = @RmaId";
                        dynamicParameters.Add("RmaId", RmaId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("退(換)貨資料錯誤!");

                        #region //判斷確認過無法再確認
                        string confirmStatus = "";
                        foreach (var item in result)
                        {
                            docType = item.DocType;
                            confirmStatus = item.ConfirmStatus;
                            transferStatus = item.TransferStatus;
                            transferMode = item.TransferMode;
                            erpPrefix = item.ErpPrefix;
                            erpNo = item.ErpNo;
                        }

                        if (confirmStatus == "Y") throw new SystemException("退貨單已經確認，無法再確認!");
                        #endregion
                        #endregion

                        #region //判斷單身是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.RmaDetail
                                WHERE RmaId = @RmaId";
                        dynamicParameters.Add("RmaId", RmaId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("請先建立退(換)貨項目!");
                        #endregion

                        #region //退(換)貨拋轉單據紀錄
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ErpPrefix, ErpNo
                                FROM SCM.RmaTransferLog
                                WHERE RmaId = @RmaId
                                AND Status = @Status";
                        dynamicParameters.Add("RmaId", RmaId);
                        dynamicParameters.Add("Status", "A");

                        erpDocStatuses = sqlConnection.Query<ErpDocStatus>(sql, dynamicParameters).ToList();
                        #endregion

                        switch (docType)
                        {
                            case "A": //暫出歸還單
                                #region //判斷單身是否有綁定暫出單
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM SCM.RmaDetail
                                        WHERE RmaId = @RmaId
                                        AND TsnDetailId IS NULL";
                                dynamicParameters.Add("RmaId", RmaId);

                                var result3 = sqlConnection.Query(sql, dynamicParameters);
                                if (result3.Count() > 0) throw new SystemException("退(換)貨項目未綁定暫出單!");
                                #endregion

                                #region //暫出歸還單單頭
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT FORMAT(a.RmaDate, 'yyyyMMdd') TH003, '1' TH004, b.CustomerNo TH005, b.CustomerShortName TH006
                                        , b.Taxation TH010, b.Currency TH011, a.RmaRemark TH014, b.CustomerName TH015
                                        , FORMAT(a.RmaDate, 'yyyyMMdd') TH023, b.ShipMethod TH030, b.TaxNo TH038
                                        FROM SCM.ReturnMerchandiseAuthorization a
                                        INNER JOIN SCM.Customer b ON a.CustomerId = b.CustomerId
                                        WHERE a.RmaId = @RmaId";
                                dynamicParameters.Add("RmaId", RmaId);

                                iNVTHs = sqlConnection.Query<INVTH>(sql, dynamicParameters).ToList();
                                if (iNVTHs.Count() <= 0) throw new SystemException("暫出歸還單單頭資料錯誤!");
                                #endregion

                                #region //暫出歸還單單身
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT RIGHT('0000' + CONVERT(NVARCHAR, ROW_NUMBER() OVER (ORDER BY a.RmaDetailId)), 4) TI003
                                        , d.MtlItemNo TI004, d.MtlItemName TI005, d.MtlItemSpec TI006, e.InventoryNo TI007, ISNULL(h.InventoryNo, f.InventoryNo) TI008
                                        , a.RmaQty TI009, a.ItemType TI034, a.FreebieOrSpareQty TI035, g.UomNo TI010, b.UnitPrice TI012, (a.RmaQty * b.UnitPrice) TI013
                                        , c.TsnErpPrefix TI014, c.TsnErpNo TI015
                                        , b.TsnSequence TI016, a.RmaDesc TI021, a.RmaQty TI038, g.UomNo TI039
                                        FROM SCM.RmaDetail a
                                        INNER JOIN SCM.TsnDetail b ON a.TsnDetailId = b.TsnDetailId
                                        INNER JOIN SCM.TempShippingNote c ON b.TsnId = c.TsnId
                                        INNER JOIN PDM.MtlItem d ON a.MtlItemId = d.MtlItemId AND b.MtlItemId = d.MtlItemId
                                        INNER JOIN SCM.Inventory e ON b.TsnInInventory = e.InventoryId
                                        INNER JOIN SCM.Inventory f ON b.TsnOutInventory = f.InventoryId
                                        INNER JOIN PDM.UnitOfMeasure g ON d.InventoryUomId = g.UomId
                                        LEFT JOIN SCM.Inventory h ON a.TargetInventory = h.InventoryId
                                        WHERE a.RmaId = @RmaId";
                                dynamicParameters.Add("RmaId", RmaId);

                                iNVTIs = sqlConnection.Query<INVTI>(sql, dynamicParameters).ToList();
                                if (iNVTIs.Count() <= 0) throw new SystemException("暫出歸還單單身資料錯誤!");
                                #endregion

                                #region //基本資料設定
                                #region //INVTH
                                iNVTHs
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.COMPANY = companyNo;
                                        x.CREATOR = userNo;
                                        x.USR_GROUP = "";
                                        x.CREATE_DATE = dateNow;
                                        x.MODIFIER = "";
                                        x.MODI_DATE = "";
                                        x.FLAG = "1";
                                        x.CREATE_TIME = timeNow;
                                        x.CREATE_AP = userNo + "PC";
                                        x.CREATE_PRID = "BM";
                                        x.MODI_TIME = "";
                                        x.MODI_AP = "";
                                        x.MODI_PRID = "";
                                    });
                                #endregion

                                #region //INVTI
                                iNVTIs
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.COMPANY = companyNo;
                                        x.CREATOR = userNo;
                                        x.USR_GROUP = "";
                                        x.CREATE_DATE = dateNow;
                                        x.MODIFIER = "";
                                        x.MODI_DATE = "";
                                        x.FLAG = "1";
                                        x.CREATE_TIME = timeNow;
                                        x.CREATE_AP = userNo + "PC";
                                        x.CREATE_PRID = "BM";
                                        x.MODI_TIME = "";
                                        x.MODI_AP = "";
                                        x.MODI_PRID = "";
                                    });
                                #endregion
                                #endregion
                                break;
                            default:
                                #region //庫存異動單單頭
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT FORMAT(a.RmaDate, 'yyyyMMdd') TA003, a.RmaRemark TA005, FORMAT(a.RmaDate, 'yyyyMMdd') TA014
                                        FROM SCM.ReturnMerchandiseAuthorization a
                                        WHERE a.RmaId = @RmaId";
                                dynamicParameters.Add("RmaId", RmaId);

                                iNVTAs = sqlConnection.Query<INVTA>(sql, dynamicParameters).ToList();
                                if (iNVTAs.Count() <= 0) throw new SystemException("庫存異動單單頭資料錯誤!");
                                #endregion

                                #region //庫存異動單單身
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT RIGHT('0000' + CONVERT(NVARCHAR, ROW_NUMBER() OVER (ORDER BY a.RmaDetailId)), 4) TB003
                                        , b.MtlItemNo TB004, b.MtlItemName TB005, b.MtlItemSpec TB006, a.RmaQty TB007, c.UomNo TB008
                                        , a.RmaDesc TB017, FORMAT(d.RmaDate, 'yyyyMMdd') TB019
                                        , CASE 
                                            WHEN d.DocType = 'B' THEN ISNULL(e.InventoryNo, 'B01')
                                            WHEN d.DocType = 'C' THEN ISNULL(e.InventoryNo, 'A06')
                                        END TB012
                                        FROM SCM.RmaDetail a
                                        INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
                                        INNER JOIN PDM.UnitOfMeasure c ON b.InventoryUomId = c.UomId
                                        INNER JOIN SCM.ReturnMerchandiseAuthorization d ON a.RmaId = d.RmaId
                                        LEFT JOIN SCM.Inventory e ON a.TargetInventory = e.InventoryId
                                        WHERE a.RmaId = @RmaId";
                                dynamicParameters.Add("RmaId", RmaId);

                                iNVTBs = sqlConnection.Query<INVTB>(sql, dynamicParameters).ToList();
                                if (iNVTBs.Count() <= 0) throw new SystemException("庫存異動單單身資料錯誤!");
                                #endregion

                                #region //基本資料設定
                                #region //INVTA
                                iNVTAs
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.COMPANY = companyNo;
                                        x.CREATOR = userNo;
                                        x.USR_GROUP = "";
                                        x.CREATE_DATE = dateNow;
                                        x.MODIFIER = "";
                                        x.MODI_DATE = "";
                                        x.FLAG = "1";
                                        x.CREATE_TIME = timeNow;
                                        x.CREATE_AP = userNo + "PC";
                                        x.CREATE_PRID = "BM";
                                        x.MODI_TIME = "";
                                        x.MODI_AP = "";
                                        x.MODI_PRID = "";
                                    });
                                #endregion

                                #region //INVTB
                                iNVTBs
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.COMPANY = companyNo;
                                        x.CREATOR = userNo;
                                        x.USR_GROUP = "";
                                        x.CREATE_DATE = dateNow;
                                        x.MODIFIER = "";
                                        x.MODI_DATE = "";
                                        x.FLAG = "1";
                                        x.CREATE_TIME = timeNow;
                                        x.CREATE_AP = userNo + "PC";
                                        x.CREATE_PRID = "BM";
                                        x.MODI_TIME = "";
                                        x.MODI_AP = "";
                                        x.MODI_PRID = "";
                                    });
                                #endregion
                                #endregion
                                break;
                        }
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        int currentNum = 0, yearLength = 0, lineLength = 0;
                        double? totalQty = 0, totalPrice = 0, totalTax = 0;
                        string encode = "", exchangeRateSource = "", factory = "", DocDate = "";
                        DateTime referenceTime = default(DateTime);
                        IEnumerable<dynamic> cmsmaResult;

                        switch (docType)
                        {
                            case "A":
                                #region //暫出歸還單
                                erpPrefix = "1501";
                                docDate = iNVTHs.Select(x => x.TH023).FirstOrDefault();
                                referenceTime = DateTime.ParseExact(docDate, "yyyyMMdd", CultureInfo.InvariantCulture);
                                DocDate = referenceTime.ToString("yyyy-MM-dd");

                                #region //比對ERP關帳日期
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT LTRIM(RTRIM(MA013)) MA013
                                        FROM CMSMA";
                                cmsmaResult = sqlConnection.Query(sql, dynamicParameters);
                                if (cmsmaResult.Count() < 0) throw new SystemException("EPR關帳日期資料錯誤!!");

                                foreach (var item in cmsmaResult)
                                {
                                    string eprDate = item.MA013;
                                    string erpYear = eprDate.Substring(0, 4);
                                    string erpMonth = eprDate.Substring(4, 2);
                                    string erpDay = eprDate.Substring(6, 2);
                                    string erpFullDate = erpYear + "-" + erpMonth + "-" + erpDay;
                                    DateTime erpDateTime = Convert.ToDateTime(erpFullDate);
                                    DateTime DocDateDateTime = Convert.ToDateTime(DocDate);
                                    int compare = DocDateDateTime.CompareTo(erpDateTime);
                                    if (compare <= 0) throw new SystemException("ERP已關帳(" + erpFullDate + ")，無法開立此日期(" + DocDate + ")之單據!!");
                                }
                                #endregion

                                string currency = iNVTHs.Select(x => x.TH011).FirstOrDefault();
                                string taxNo = iNVTHs.Select(x => x.TH038).FirstOrDefault();

                                #region //單據設定
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.MQ004, a.MQ005, a.MQ006, a.MQ044
                                        FROM CMSMQ a
                                        WHERE a.COMPANY = @CompanyNo
                                        AND a.MQ001 = @ErpPrefix";
                                dynamicParameters.Add("CompanyNo", companyNo);
                                dynamicParameters.Add("ErpPrefix", erpPrefix);

                                var result1501DocSetting = sqlConnection.Query(sql, dynamicParameters);
                                if (result1501DocSetting.Count() <= 0) throw new SystemException("ERP單據設定資料不存在!");

                                foreach (var item in result1501DocSetting)
                                {
                                    encode = item.MQ004; //編碼方式
                                    yearLength = Convert.ToInt32(item.MQ005); //年碼數
                                    lineLength = Convert.ToInt32(item.MQ006); //流水號碼數
                                    exchangeRateSource = item.MQ044; //匯率來源
                                }
                                #endregion

                                #region //單號取號
                                if (transferMode == "N")
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT CONVERT(int, RIGHT(ISNULL(MAX(RTRIM(LTRIM(TH002))), '" + new string('0', lineLength) + @"'), " + lineLength + @")) + 1 CurrentNum
                                            FROM INVTH
                                            WHERE TH001 = @ErpPrefix";
                                    dynamicParameters.Add("ErpPrefix", erpPrefix);

                                    #region //編碼方式
                                    string dateFormat = "";
                                    switch (encode)
                                    {
                                        case "1": //日編
                                            dateFormat = new string('y', yearLength) + "MMdd";
                                            sql += @" AND RTRIM(LTRIM(TH002)) LIKE @ReferenceTime + '" + new string('_', lineLength) + @"'";
                                            dynamicParameters.Add("ReferenceTime", referenceTime.ToString(dateFormat));
                                            erpNo = referenceTime.ToString(dateFormat);
                                            break;
                                        case "2": //月編
                                            dateFormat = new string('y', yearLength) + "MM";
                                            sql += @" AND RTRIM(LTRIM(TH002)) LIKE @ReferenceTime + '" + new string('_', lineLength) + @"'";
                                            dynamicParameters.Add("ReferenceTime", referenceTime.ToString(dateFormat));
                                            erpNo = referenceTime.ToString(dateFormat);
                                            break;
                                        case "3": //流水號
                                            break;
                                        case "4": //手動編號
                                            break;
                                        default:
                                            throw new SystemException("編碼方式錯誤!");
                                    }
                                    #endregion

                                    currentNum = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).CurrentNum;
                                    erpNo += string.Format("{0:" + new string('0', lineLength) + "}", currentNum);
                                }
                                #endregion

                                #region //廠別資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 MB001
                                        FROM CMSMB
                                        WHERE COMPANY = @COMPANY";
                                dynamicParameters.Add("COMPANY", companyNo);

                                var result1501Factory = sqlConnection.Query(sql, dynamicParameters);
                                if (result1501Factory.Count() <= 0) throw new SystemException("ERP廠別資料不存在!");

                                foreach (var item in result1501Factory)
                                {
                                    factory = item.MB001; //廠別
                                }
                                #endregion

                                #region //交易幣別設定
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT MF003, MF004, MF005, MF006
                                        FROM CMSMF
                                        WHERE MF001 = @Currency";
                                dynamicParameters.Add("Currency", currency);

                                var resultCurrencySetting = sqlConnection.Query(sql, dynamicParameters);
                                if (resultCurrencySetting.Count() <= 0) throw new SystemException("ERP交易幣別設定不存在!");

                                int unitRound = 0,
                                    totalRound = 0;
                                foreach (var item in resultCurrencySetting)
                                {
                                    unitRound = Convert.ToInt32(item.MF003); //單價取位
                                    totalRound = Convert.ToInt32(item.MF004); //金額取位
                                }
                                #endregion

                                #region //目前匯率
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 MG003, MG004, MG005, MG006
                                        FROM CMSMG
                                        WHERE MG001 = @Currency
                                        AND MG002 <= @DateNow
                                        ORDER BY MG002 DESC";
                                dynamicParameters.Add("Currency", currency);
                                dynamicParameters.Add("DateNow", DateTime.Now.ToString("yyyyMMdd"));

                                var resultExchangeRate = sqlConnection.Query(sql, dynamicParameters);
                                if (resultExchangeRate.Count() <= 0) throw new SystemException("ERP交易幣別匯率不存在!");

                                double exchangeRate = 0;
                                foreach (var item in resultExchangeRate)
                                {
                                    switch (exchangeRateSource)
                                    {
                                        case "I": //銀行買進匯率
                                            exchangeRate = Convert.ToDouble(item.MG003);
                                            break;
                                        case "O": //銀行賣出匯率
                                            exchangeRate = Convert.ToDouble(item.MG004);
                                            break;
                                        case "E": //報關買進匯率
                                            exchangeRate = Convert.ToDouble(item.MG005);
                                            break;
                                        case "W": //報關賣出匯率
                                            exchangeRate = Convert.ToDouble(item.MG006);
                                            break;
                                    }
                                }
                                #endregion

                                #region //稅別碼設定
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT RTRIM(LTRIM(NN001)) NN001, RTRIM(LTRIM(NN002)) NN002, NN004, RTRIM(LTRIM(NN006)) NN006
                                        FROM CMSNN
                                        WHERE NN001 = @TaxNo";
                                dynamicParameters.Add("TaxNo", taxNo);

                                var resultTax = sqlConnection.Query(sql, dynamicParameters);
                                if (resultTax.Count() <= 0) throw new SystemException("ERP稅別碼設定不存在!");

                                double exciseTax = 0;
                                string taxation = "";
                                foreach (var item in resultTax)
                                {
                                    exciseTax = Convert.ToDouble(item.NN004); //營業稅率
                                    taxation = item.NN006; //課稅別
                                }
                                #endregion

                                #region //計算數量與金額
                                totalQty = iNVTIs.Select(x => x.TI009).Sum() + iNVTIs.Select(x => x.TI035).Sum();
                                totalPrice = 0;
                                totalTax = 0;

                                foreach (var item in iNVTIs)
                                {
                                    totalPrice += Math.Round((double)item.TI009 * (double)item.TI012, totalRound);
                                    totalTax += Math.Round((double)totalPrice * exciseTax, totalRound);
                                }

                                switch (taxation)
                                {
                                    case "1":
                                        totalTax = totalPrice - Math.Round((double)totalPrice / (1 + exciseTax), unitRound);
                                        break;
                                    case "2":
                                        break;
                                    case "3":
                                        break;
                                    case "4":
                                        break;
                                    case "9":
                                        break;
                                }
                                #endregion

                                #region //INVTH
                                iNVTHs
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.TH001 = erpPrefix; //異動單別
                                        x.TH002 = erpNo; //異動單號
                                        x.TH007 = departmentNo; //部門代號
                                        x.TH008 = userNo; //員工代號
                                        x.TH009 = factory; //廠別
                                        x.TH012 = exchangeRate; //匯率
                                        x.TH013 = 0; //件數
                                        x.TH016 = "台中市南屯區工業二十二路21號"; //地址一
                                        x.TH017 = ""; //地址二
                                        x.TH018 = ""; //其它備註
                                        x.TH019 = 0; //列印次數
                                        x.TH020 = "N"; //確認碼
                                        x.TH021 = totalQty; //總數量
                                        x.TH022 = totalPrice; //總金額
                                        x.TH024 = ""; //確認者
                                        x.TH025 = exciseTax; //營業稅率
                                        x.TH026 = totalTax; //稅額
                                        x.TH027 = 0; //總包裝數量
                                        x.TH028 = "N"; //簽核狀態碼
                                        x.TH029 = 0; //傳送次數
                                        x.TH031 = ""; //派車單別
                                        x.TH032 = ""; //派車單號
                                        x.TH033 = 0; //預留欄位
                                        x.TH034 = 0; //預留欄位
                                        x.TH035 = ""; //預留欄位
                                        x.TH036 = ""; //預留欄位
                                        x.TH037 = ""; //預留欄位
                                        x.TH039 = "N"; //單身多稅率
                                        x.UDF01 = "";
                                        x.UDF02 = "";
                                        x.UDF03 = "";
                                        x.UDF04 = "";
                                        x.UDF05 = "";
                                        x.UDF06 = 0;
                                        x.UDF07 = 0;
                                        x.UDF08 = 0;
                                        x.UDF09 = 0;
                                        x.UDF10 = 0;
                                    });

                                switch (transferMode)
                                {
                                    case "N":
                                        #region //判斷單號是否重複
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT TOP 1 1
                                                FROM INVTH
                                                WHERE TH001 = @ErpPrefix
                                                AND TH002 = @ErpNo";
                                        dynamicParameters.Add("ErpPrefix", erpPrefix);
                                        dynamicParameters.Add("ErpNo", erpNo);

                                        var result1501RepeatExist = sqlConnection.Query(sql, dynamicParameters);
                                        if (result1501RepeatExist.Count() > 0) throw new SystemException("【暫出/入歸還單號】重複，請重新取號!");
                                        #endregion

                                        #region //開立新單據
                                        sql = @"INSERT INTO INVTH (COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE
                                                , TH001, TH002, TH003, TH004, TH005, TH006, TH007, TH008, TH009, TH010
                                                , TH011, TH012, TH013, TH014, TH015, TH016, TH017, TH018, TH019, TH020
                                                , TH021, TH022, TH023, TH024, TH025, TH026, TH027, TH028, TH029, TH030
                                                , TH031, TH032, TH033, TH034, TH035, TH036, TH037, TH038, TH039
                                                , FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID
                                                , UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10)
                                                VALUES (@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE
                                                , @TH001, @TH002, @TH003, @TH004, @TH005, @TH006, @TH007, @TH008, @TH009, @TH010
                                                , @TH011, @TH012, @TH013, @TH014, @TH015, @TH016, @TH017, @TH018, @TH019, @TH020
                                                , @TH021, @TH022, @TH023, @TH024, @TH025, @TH026, @TH027, @TH028, @TH029, @TH030
                                                , @TH031, @TH032, @TH033, @TH034, @TH035, @TH036, @TH037, @TH038, @TH039
                                                , @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID, @MODI_TIME, @MODI_AP, @MODI_PRID
                                                , @UDF01, @UDF02, @UDF03, @UDF04, @UDF05, @UDF06, @UDF07, @UDF08, @UDF09, @UDF10)";
                                        rowsAffected += sqlConnection.Execute(sql, iNVTHs);
                                        #endregion
                                        break;
                                    case "U":
                                        #region //判斷原單據是否存在
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT TOP 1 TH020
                                                FROM INVTH
                                                WHERE TH001 = @ErpPrefix
                                                AND TH002 = @ErpNo";
                                        dynamicParameters.Add("ErpPrefix", erpPrefix);
                                        dynamicParameters.Add("ErpNo", erpNo);

                                        var result1501Exist = sqlConnection.Query(sql, dynamicParameters);
                                        if (result1501Exist.Count() <= 0) throw new SystemException("原單據不存在，【拋轉模式】請選擇開立新單據!");

                                        foreach (var item in result1501Exist)
                                        {
                                            if (item.TH020 != "N") throw new SystemException("原單據確認碼為無法更動狀態!");
                                        }
                                        #endregion

                                        #region //刪除原單據單身
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"DELETE INVTI
                                                WHERE TI001 = @ErpPrefix
                                                AND TI002 = @ErpNo";
                                        dynamicParameters.Add("ErpPrefix", erpPrefix);
                                        dynamicParameters.Add("ErpNo", erpNo);

                                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                        #endregion

                                        #region //更新原單據
                                        sql = @"UPDATE INVTH SET
                                                TH003 = @TH003,
                                                TH004 = @TH004,
                                                TH005 = @TH005,
                                                TH006 = @TH006,
                                                TH007 = @TH007,
                                                TH008 = @TH008,
                                                TH009 = @TH009,
                                                TH010 = @TH010,
                                                TH011 = @TH011,
                                                TH012 = @TH012,
                                                TH013 = @TH013,
                                                TH014 = @TH014,
                                                TH015 = @TH015,
                                                TH016 = @TH016,
                                                TH017 = @TH017,
                                                TH018 = @TH018,
                                                TH019 = @TH019,
                                                TH020 = @TH020,
                                                TH021 = @TH021,
                                                TH022 = @TH022,
                                                TH023 = @TH023,
                                                TH024 = @TH024,
                                                TH025 = @TH025,
                                                TH026 = @TH026,
                                                TH027 = @TH027,
                                                TH028 = @TH028,
                                                TH029 = @TH029,
                                                TH030 = @TH030,
                                                TH031 = @TH031,
                                                TH032 = @TH032,
                                                TH033 = @TH033,
                                                TH034 = @TH034,
                                                TH035 = @TH035,
                                                TH036 = @TH036,
                                                TH037 = @TH037,
                                                TH038 = @TH038,
                                                TH039 = @TH039,
                                                UDF01 = @UDF01,
                                                UDF02 = @UDF02,
                                                UDF03 = @UDF03,
                                                UDF04 = @UDF04,
                                                UDF05 = @UDF05,
                                                UDF06 = @UDF06,
                                                UDF07 = @UDF07,
                                                UDF08 = @UDF08,
                                                UDF09 = @UDF09,
                                                UDF10 = @UDF10
                                                WHERE TH001 = @TH001
                                                AND TH002 = @TH002";
                                        rowsAffected += sqlConnection.Execute(sql, iNVTHs);
                                        #endregion
                                        break;
                                    default:
                                        throw new SystemException("拋轉模式錯誤!");
                                }
                                #endregion

                                #region //INVTI
                                iNVTIs
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.TI001 = erpPrefix; //異動單別
                                        x.TI002 = erpNo; //異動單號
                                        x.TI011 = ""; //小單位
                                        x.TI017 = ""; //批號
                                        x.TI018 = ""; //有效日期
                                        x.TI019 = ""; //複檢日期
                                        x.TI020 = ""; //專案代號
                                        x.TI022 = "N"; //確認碼
                                        x.TI023 = 0; //包裝數量
                                        x.TI024 = ""; //包裝單位
                                        x.TI025 = 0; //產品序號數量
                                        x.TI026 = ""; //轉出儲位
                                        x.TI027 = ""; //轉入儲位
                                        x.TI028 = 0; //預留欄位
                                        x.TI029 = 0; //預留欄位
                                        x.TI030 = ""; //預留欄位
                                        x.TI031 = ""; //預留欄位
                                        x.TI032 = ""; //預留欄位
                                        x.TI033 = exciseTax; //營業稅率
                                        x.TI036 = 0; //贈/備品包裝量
                                        x.TI037 = ""; //稅別碼
                                        x.TI500 = ""; //刻號/BIN記錄
                                        x.TI501 = ""; //刻號管理
                                        x.TI502 = ""; //DATECODE
                                        x.TI503 = ""; //產品系列
                                        x.UDF01 = "";
                                        x.UDF02 = "";
                                        x.UDF03 = "";
                                        x.UDF04 = "";
                                        x.UDF05 = "";
                                        x.UDF06 = 0;
                                        x.UDF07 = 0;
                                        x.UDF08 = 0;
                                        x.UDF09 = 0;
                                        x.UDF10 = 0;
                                    });

                                sql = @"INSERT INTO INVTI (COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE
                                        , TI001, TI002, TI003, TI004, TI005, TI006, TI007, TI008, TI009, TI010
                                        , TI011, TI012, TI013, TI014, TI015, TI016, TI017, TI018, TI019, TI020
                                        , TI021, TI022, TI023, TI024, TI025, TI026, TI027, TI028, TI029, TI030
                                        , TI031, TI032, TI033, TI034, TI035, TI036, TI037, TI038, TI039
                                        , TI500, TI501, TI502, TI503
                                        , FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID
                                        , UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10)
                                        VALUES (@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE
                                        , @TI001, @TI002, @TI003, @TI004, @TI005, @TI006, @TI007, @TI008, @TI009, @TI010
                                        , @TI011, @TI012, @TI013, @TI014, @TI015, @TI016, @TI017, @TI018, @TI019, @TI020
                                        , @TI021, @TI022, @TI023, @TI024, @TI025, @TI026, @TI027, @TI028, @TI029, @TI030
                                        , @TI031, @TI032, @TI033, @TI034, @TI035, @TI036, @TI037, @TI038, @TI039
                                        , @TI500, @TI501, @TI502, @TI503
                                        , @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID, @MODI_TIME, @MODI_AP, @MODI_PRID
                                        , @UDF01, @UDF02, @UDF03, @UDF04, @UDF05, @UDF06, @UDF07, @UDF08, @UDF09, @UDF10)";
                                rowsAffected += sqlConnection.Execute(sql, iNVTIs);
                                #endregion
                                #endregion
                                break;
                            case "B":
                                #region //客供入料單
                                erpPrefix = "1106";
                                docDate = iNVTAs.Select(x => x.TA014).FirstOrDefault();
                                referenceTime = DateTime.ParseExact(docDate, "yyyyMMdd", CultureInfo.InvariantCulture);
                                DocDate = referenceTime.ToString("yyyy-MM-dd");

                                #region //比對ERP關帳日期
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT LTRIM(RTRIM(MA013)) MA013
                                        FROM CMSMA";
                                cmsmaResult = sqlConnection.Query(sql, dynamicParameters);
                                if (cmsmaResult.Count() < 0) throw new SystemException("EPR關帳日期資料錯誤!!");

                                foreach (var item in cmsmaResult)
                                {
                                    string eprDate = item.MA013;
                                    string erpYear = eprDate.Substring(0, 4);
                                    string erpMonth = eprDate.Substring(4, 2);
                                    string erpDay = eprDate.Substring(6, 2);
                                    string erpFullDate = erpYear + "-" + erpMonth + "-" + erpDay;
                                    DateTime erpDateTime = Convert.ToDateTime(erpFullDate);
                                    DateTime DocDateDateTime = Convert.ToDateTime(DocDate);
                                    int compare = DocDateDateTime.CompareTo(erpDateTime);
                                    if (compare <= 0) throw new SystemException("ERP已關帳(" + erpFullDate + ")，無法開立此日期(" + DocDate + ")之單據!!");
                                }
                                #endregion

                                #region //單據設定
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.MQ004, a.MQ005, a.MQ006, a.MQ044
                                        FROM CMSMQ a
                                        WHERE a.COMPANY = @CompanyNo
                                        AND a.MQ001 = @ErpPrefix";
                                dynamicParameters.Add("CompanyNo", companyNo);
                                dynamicParameters.Add("ErpPrefix", erpPrefix);

                                var result1106DocSetting = sqlConnection.Query(sql, dynamicParameters);
                                if (result1106DocSetting.Count() <= 0) throw new SystemException("ERP單據設定資料不存在!");

                                foreach (var item in result1106DocSetting)
                                {
                                    encode = item.MQ004; //編碼方式
                                    yearLength = Convert.ToInt32(item.MQ005); //年碼數
                                    lineLength = Convert.ToInt32(item.MQ006); //流水號碼數
                                }
                                #endregion

                                #region //單號取號
                                if (transferMode == "N")
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT CONVERT(int, RIGHT(ISNULL(MAX(RTRIM(LTRIM(TA002))), '" + new string('0', lineLength) + @"'), " + lineLength + @")) + 1 CurrentNum
                                            FROM INVTA
                                            WHERE TA001 = @ErpPrefix";
                                    dynamicParameters.Add("ErpPrefix", erpPrefix);

                                    #region //編碼方式
                                    string dateFormat = "";
                                    switch (encode)
                                    {
                                        case "1": //日編
                                            dateFormat = new string('y', yearLength) + "MMdd";
                                            sql += @" AND RTRIM(LTRIM(TA002)) LIKE @ReferenceTime + '" + new string('_', lineLength) + @"'";
                                            dynamicParameters.Add("ReferenceTime", referenceTime.ToString(dateFormat));
                                            erpNo = referenceTime.ToString(dateFormat);
                                            break;
                                        case "2": //月編
                                            dateFormat = new string('y', yearLength) + "MM";
                                            sql += @" AND RTRIM(LTRIM(TA002)) LIKE @ReferenceTime + '" + new string('_', lineLength) + @"'";
                                            dynamicParameters.Add("ReferenceTime", referenceTime.ToString(dateFormat));
                                            erpNo = referenceTime.ToString(dateFormat);
                                            break;
                                        case "3": //流水號
                                            break;
                                        case "4": //手動編號
                                            break;
                                        default:
                                            throw new SystemException("編碼方式錯誤!");
                                    }
                                    #endregion

                                    currentNum = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).CurrentNum;
                                    erpNo += string.Format("{0:" + new string('0', lineLength) + "}", currentNum);
                                }
                                #endregion

                                #region //廠別資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 MB001
                                        FROM CMSMB
                                        WHERE COMPANY = @COMPANY";
                                dynamicParameters.Add("COMPANY", companyNo);

                                var result1106Factory = sqlConnection.Query(sql, dynamicParameters);
                                if (result1106Factory.Count() <= 0) throw new SystemException("ERP廠別資料不存在!");

                                foreach (var item in result1106Factory)
                                {
                                    factory = item.MB001; //廠別
                                }
                                #endregion

                                #region //計算數量
                                totalQty = iNVTBs.Select(x => x.TB007).Sum();
                                #endregion

                                #region //INVTA
                                iNVTAs
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.TA001 = erpPrefix; //單別
                                        x.TA002 = erpNo; //單號
                                        x.TA004 = departmentNo; //部門代號
                                        x.TA006 = "N"; //確認碼
                                        x.TA007 = 0; //列印次數
                                        x.TA008 = factory; //廠別代號
                                        x.TA009 = "11"; //單據性質碼 > 一般單據
                                        x.TA010 = 0; //件數
                                        x.TA011 = totalQty; //總數量
                                        x.TA012 = totalPrice; //總金額
                                        x.TA013 = ""; //產生分錄碼
                                        x.TA015 = ""; //確認者
                                        x.TA016 = 0; //總包裝數量
                                        x.TA017 = "N"; //簽核狀態碼
                                        x.TA018 = "0"; //保稅碼 > 依品號預設
                                        x.TA019 = 0; //傳送次數
                                        x.TA020 = ""; //運輸方式
                                        x.TA021 = ""; //派車單別
                                        x.TA022 = ""; //派車單號
                                        x.TA023 = 0; //預留欄位
                                        x.TA024 = 0; //預留欄位
                                        x.TA025 = ""; //來源
                                        x.TA026 = ""; //預留欄位
                                        x.TA027 = ""; //預留欄位
                                        x.TA028 = ""; //海關手冊
                                        x.UDF01 = "";
                                        x.UDF02 = "";
                                        x.UDF03 = "";
                                        x.UDF04 = "";
                                        x.UDF05 = "";
                                        x.UDF06 = 0;
                                        x.UDF07 = 0;
                                        x.UDF08 = 0;
                                        x.UDF09 = 0;
                                        x.UDF10 = 0;
                                    });

                                switch (transferMode)
                                {
                                    case "N":
                                        #region //判斷單號是否重複
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT TOP 1 1
                                                FROM INVTA
                                                WHERE TA001 = @ErpPrefix
                                                AND TA002 = @ErpNo";
                                        dynamicParameters.Add("ErpPrefix", erpPrefix);
                                        dynamicParameters.Add("ErpNo", erpNo);

                                        var result1106RepeatExist = sqlConnection.Query(sql, dynamicParameters);
                                        if (result1106RepeatExist.Count() > 0) throw new SystemException("【庫存異動單單號】重複，請重新取號!");
                                        #endregion

                                        #region //開立新單據
                                        sql = @"INSERT INTO INVTA (COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE
                                                , TA001, TA002, TA003, TA004, TA005, TA006, TA007, TA008, TA009, TA010
                                                , TA011, TA012, TA013, TA014, TA015, TA016, TA017, TA018, TA019, TA020
                                                , TA021, TA022, TA023, TA024, TA025, TA026, TA027, TA028
                                                , FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID
                                                , UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10)
                                                VALUES (@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE
                                                , @TA001, @TA002, @TA003, @TA004, @TA005, @TA006, @TA007, @TA008, @TA009, @TA010
                                                , @TA011, @TA012, @TA013, @TA014, @TA015, @TA016, @TA017, @TA018, @TA019, @TA020
                                                , @TA021, @TA022, @TA023, @TA024, @TA025, @TA026, @TA027, @TA028
                                                , @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID, @MODI_TIME, @MODI_AP, @MODI_PRID
                                                , @UDF01, @UDF02, @UDF03, @UDF04, @UDF05, @UDF06, @UDF07, @UDF08, @UDF09, @UDF10)";
                                        rowsAffected += sqlConnection.Execute(sql, iNVTAs);
                                        #endregion
                                        break;
                                    case "U":
                                        #region //判斷原單據是否存在
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT TOP 1 TA006
                                                FROM INVTA
                                                WHERE TA001 = @ErpPrefix
                                                AND TA002 = @ErpNo";
                                        dynamicParameters.Add("ErpPrefix", erpPrefix);
                                        dynamicParameters.Add("ErpNo", erpNo);

                                        var result1106Exist = sqlConnection.Query(sql, dynamicParameters);
                                        if (result1106Exist.Count() <= 0) throw new SystemException("原單據不存在，【拋轉模式】請選擇開立新單據!");

                                        foreach (var item in result1106Exist)
                                        {
                                            if (item.TA006 != "N") throw new SystemException("原單據確認碼為無法更動狀態!");
                                        }
                                        #endregion

                                        #region //刪除原單據單身
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"DELETE INVTB
                                                WHERE TB001 = @ErpPrefix
                                                AND TB002 = @ErpNo";
                                        dynamicParameters.Add("ErpPrefix", erpPrefix);
                                        dynamicParameters.Add("ErpNo", erpNo);

                                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                        #endregion

                                        #region //更新原單據
                                        sql = @"UPDATE INVTA SET
                                                TA003 = @TA003,
                                                TA004 = @TA004,
                                                TA005 = @TA005,
                                                TA006 = @TA006,
                                                TA007 = @TA007,
                                                TA008 = @TA008,
                                                TA009 = @TA009,
                                                TA010 = @TA010,
                                                TA011 = @TA011,
                                                TA012 = @TA012,
                                                TA013 = @TA013,
                                                TA014 = @TA014,
                                                TA015 = @TA015,
                                                TA016 = @TA016,
                                                TA017 = @TA017,
                                                TA018 = @TA018,
                                                TA019 = @TA019,
                                                TA020 = @TA020,
                                                TA021 = @TA021,
                                                TA022 = @TA022,
                                                TA023 = @TA023,
                                                TA024 = @TA024,
                                                TA025 = @TA025,
                                                TA026 = @TA026,
                                                TA027 = @TA027,
                                                TA028 = @TA028,
                                                UDF01 = @UDF01,
                                                UDF02 = @UDF02,
                                                UDF03 = @UDF03,
                                                UDF04 = @UDF04,
                                                UDF05 = @UDF05,
                                                UDF06 = @UDF06,
                                                UDF07 = @UDF07,
                                                UDF08 = @UDF08,
                                                UDF09 = @UDF09,
                                                UDF10 = @UDF10
                                                WHERE TA001 = @TA001
                                                AND TA002 = @TA002";
                                        rowsAffected += sqlConnection.Execute(sql, iNVTAs);
                                        #endregion
                                        break;
                                    default:
                                        throw new SystemException("拋轉模式錯誤!");
                                }
                                #endregion

                                #region //INVTB
                                iNVTBs
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.TB001 = erpPrefix; //單別
                                        x.TB002 = erpNo; //單號
                                        x.TB009 = 0; //庫存數量
                                        x.TB010 = 0; //單位成本
                                        x.TB011 = 0; //金額
                                        x.TB013 = ""; //轉入庫
                                        x.TB014 = ""; //批號
                                        x.TB015 = ""; //有效日期
                                        x.TB016 = ""; //複檢日期
                                        x.TB018 = "N"; //確認碼
                                        x.TB020 = ""; //小單位
                                        x.TB021 = ""; //專案代號
                                        x.TB022 = 0; //包裝數量
                                        x.TB023 = ""; //包裝單位
                                        x.TB024 = "N"; //保稅碼
                                        x.TB025 = 0; //產品序號數量
                                        x.TB026 = ""; //轉出儲位
                                        x.TB027 = ""; //轉入儲位
                                        x.TB028 = 0; //預留欄位
                                        x.TB029 = 0; //預留欄位
                                        x.TB030 = ""; //預留欄位
                                        x.TB031 = ""; //預留欄位
                                        x.TB032 = ""; //預留欄位
                                        x.TB500 = ""; //刻號/BIN記錄
                                        x.TB501 = ""; //刻號管理
                                        x.TB502 = ""; //DATECODE
                                        x.UDF01 = "";
                                        x.UDF02 = "";
                                        x.UDF03 = "";
                                        x.UDF04 = "";
                                        x.UDF05 = "";
                                        x.UDF06 = 0;
                                        x.UDF07 = 0;
                                        x.UDF08 = 0;
                                        x.UDF09 = 0;
                                        x.UDF10 = 0;
                                    });

                                sql = @"INSERT INTO INVTB (COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE
                                        , TB001, TB002, TB003, TB004, TB005, TB006, TB007, TB008, TB009, TB010
                                        , TB011, TB012, TB013, TB014, TB015, TB016, TB017, TB018, TB019, TB020
                                        , TB021, TB022, TB023, TB024, TB025, TB026, TB027, TB028, TB029, TB030
                                        , TB031, TB032, TB500, TB501, TB502
                                        , FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID
                                        , UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10)
                                        VALUES (@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE
                                        , @TB001, @TB002, @TB003, @TB004, @TB005, @TB006, @TB007, @TB008, @TB009, @TB010
                                        , @TB011, @TB012, @TB013, @TB014, @TB015, @TB016, @TB017, @TB018, @TB019, @TB020
                                        , @TB021, @TB022, @TB023, @TB024, @TB025, @TB026, @TB027, @TB028, @TB029, @TB030
                                        , @TB031, @TB032, @TB500, @TB501, @TB502
                                        , @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID, @MODI_TIME, @MODI_AP, @MODI_PRID
                                        , @UDF01, @UDF02, @UDF03, @UDF04, @UDF05, @UDF06, @UDF07, @UDF08, @UDF09, @UDF10)";
                                rowsAffected += sqlConnection.Execute(sql, iNVTBs);
                                #endregion
                                #endregion
                                break;
                            case "C":
                                #region //0成本入庫單
                                erpPrefix = "1109";
                                docDate = iNVTAs.Select(x => x.TA014).FirstOrDefault();
                                referenceTime = DateTime.ParseExact(docDate, "yyyyMMdd", CultureInfo.InvariantCulture);
                                DocDate = referenceTime.ToString("yyyy-MM-dd");

                                #region //比對ERP關帳日期
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT LTRIM(RTRIM(MA013)) MA013
                                        FROM CMSMA";
                                cmsmaResult = sqlConnection.Query(sql, dynamicParameters);
                                if (cmsmaResult.Count() < 0) throw new SystemException("EPR關帳日期資料錯誤!!");

                                foreach (var item in cmsmaResult)
                                {
                                    string eprDate = item.MA013;
                                    string erpYear = eprDate.Substring(0, 4);
                                    string erpMonth = eprDate.Substring(4, 2);
                                    string erpDay = eprDate.Substring(6, 2);
                                    string erpFullDate = erpYear + "-" + erpMonth + "-" + erpDay;
                                    DateTime erpDateTime = Convert.ToDateTime(erpFullDate);
                                    DateTime DocDateDateTime = Convert.ToDateTime(DocDate);
                                    int compare = DocDateDateTime.CompareTo(erpDateTime);
                                    if (compare <= 0) throw new SystemException("ERP已關帳(" + erpFullDate + ")，無法開立此日期(" + DocDate + ")之單據!!");
                                }
                                #endregion

                                #region //單據設定
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.MQ004, a.MQ005, a.MQ006, a.MQ044
                                        FROM CMSMQ a
                                        WHERE a.COMPANY = @CompanyNo
                                        AND a.MQ001 = @ErpPrefix";
                                dynamicParameters.Add("CompanyNo", companyNo);
                                dynamicParameters.Add("ErpPrefix", erpPrefix);

                                var result1109DocSetting = sqlConnection.Query(sql, dynamicParameters);
                                if (result1109DocSetting.Count() <= 0) throw new SystemException("ERP單據設定資料不存在!");

                                foreach (var item in result1109DocSetting)
                                {
                                    encode = item.MQ004; //編碼方式
                                    yearLength = Convert.ToInt32(item.MQ005); //年碼數
                                    lineLength = Convert.ToInt32(item.MQ006); //流水號碼數
                                }
                                #endregion

                                #region //單號取號
                                if (transferMode == "N")
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT CONVERT(int, RIGHT(ISNULL(MAX(RTRIM(LTRIM(TA002))), '" + new string('0', lineLength) + @"'), " + lineLength + @")) + 1 CurrentNum
                                            FROM INVTA
                                            WHERE TA001 = @ErpPrefix";
                                    dynamicParameters.Add("ErpPrefix", erpPrefix);

                                    #region //編碼方式
                                    string dateFormat = "";
                                    switch (encode)
                                    {
                                        case "1": //日編
                                            dateFormat = new string('y', yearLength) + "MMdd";
                                            sql += @" AND RTRIM(LTRIM(TA002)) LIKE @ReferenceTime + '" + new string('_', lineLength) + @"'";
                                            dynamicParameters.Add("ReferenceTime", referenceTime.ToString(dateFormat));
                                            erpNo = referenceTime.ToString(dateFormat);
                                            break;
                                        case "2": //月編
                                            dateFormat = new string('y', yearLength) + "MM";
                                            sql += @" AND RTRIM(LTRIM(TA002)) LIKE @ReferenceTime + '" + new string('_', lineLength) + @"'";
                                            dynamicParameters.Add("ReferenceTime", referenceTime.ToString(dateFormat));
                                            erpNo = referenceTime.ToString(dateFormat);
                                            break;
                                        case "3": //流水號
                                            break;
                                        case "4": //手動編號
                                            break;
                                        default:
                                            throw new SystemException("編碼方式錯誤!");
                                    }
                                    #endregion

                                    currentNum = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).CurrentNum;
                                    erpNo += string.Format("{0:" + new string('0', lineLength) + "}", currentNum);
                                }
                                #endregion

                                #region //廠別資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 MB001
                                        FROM CMSMB
                                        WHERE COMPANY = @COMPANY";
                                dynamicParameters.Add("COMPANY", companyNo);

                                var result1109Factory = sqlConnection.Query(sql, dynamicParameters);
                                if (result1109Factory.Count() <= 0) throw new SystemException("ERP廠別資料不存在!");

                                foreach (var item in result1109Factory)
                                {
                                    factory = item.MB001; //廠別
                                }
                                #endregion

                                #region //計算數量
                                totalQty = iNVTBs.Select(x => x.TB007).Sum();
                                #endregion

                                #region //INVTA
                                iNVTAs
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.TA001 = erpPrefix; //單別
                                        x.TA002 = erpNo; //單號
                                        x.TA004 = departmentNo; //部門代號
                                        x.TA006 = "N"; //確認碼
                                        x.TA007 = 0; //列印次數
                                        x.TA008 = factory; //廠別代號
                                        x.TA009 = "11"; //單據性質碼 > 一般單據
                                        x.TA010 = 0; //件數
                                        x.TA011 = totalQty; //總數量
                                        x.TA012 = totalPrice; //總金額
                                        x.TA013 = ""; //產生分錄碼
                                        x.TA015 = ""; //確認者
                                        x.TA016 = 0; //總包裝數量
                                        x.TA017 = "N"; //簽核狀態碼
                                        x.TA018 = "0"; //保稅碼 > 依品號預設
                                        x.TA019 = 0; //傳送次數
                                        x.TA020 = ""; //運輸方式
                                        x.TA021 = ""; //派車單別
                                        x.TA022 = ""; //派車單號
                                        x.TA023 = 0; //預留欄位
                                        x.TA024 = 0; //預留欄位
                                        x.TA025 = ""; //來源
                                        x.TA026 = ""; //預留欄位
                                        x.TA027 = ""; //預留欄位
                                        x.TA028 = ""; //海關手冊
                                        x.UDF01 = "";
                                        x.UDF02 = "";
                                        x.UDF03 = "";
                                        x.UDF04 = "";
                                        x.UDF05 = "";
                                        x.UDF06 = 0;
                                        x.UDF07 = 0;
                                        x.UDF08 = 0;
                                        x.UDF09 = 0;
                                        x.UDF10 = 0;
                                    });

                                switch (transferMode)
                                {
                                    case "N":
                                        #region //判斷單號是否重複
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT TOP 1 1
                                                FROM INVTA
                                                WHERE TA001 = @ErpPrefix
                                                AND TA002 = @ErpNo";
                                        dynamicParameters.Add("ErpPrefix", erpPrefix);
                                        dynamicParameters.Add("ErpNo", erpNo);

                                        var result1109RepeatExist = sqlConnection.Query(sql, dynamicParameters);
                                        if (result1109RepeatExist.Count() > 0) throw new SystemException("【庫存異動單單號】重複，請重新取號!");
                                        #endregion

                                        #region //開立新單據
                                        sql = @"INSERT INTO INVTA (COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE
                                                , TA001, TA002, TA003, TA004, TA005, TA006, TA007, TA008, TA009, TA010
                                                , TA011, TA012, TA013, TA014, TA015, TA016, TA017, TA018, TA019, TA020
                                                , TA021, TA022, TA023, TA024, TA025, TA026, TA027, TA028
                                                , FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID
                                                , UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10)
                                                VALUES (@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE
                                                , @TA001, @TA002, @TA003, @TA004, @TA005, @TA006, @TA007, @TA008, @TA009, @TA010
                                                , @TA011, @TA012, @TA013, @TA014, @TA015, @TA016, @TA017, @TA018, @TA019, @TA020
                                                , @TA021, @TA022, @TA023, @TA024, @TA025, @TA026, @TA027, @TA028
                                                , @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID, @MODI_TIME, @MODI_AP, @MODI_PRID
                                                , @UDF01, @UDF02, @UDF03, @UDF04, @UDF05, @UDF06, @UDF07, @UDF08, @UDF09, @UDF10)";
                                        rowsAffected += sqlConnection.Execute(sql, iNVTAs);
                                        #endregion
                                        break;
                                    case "U":
                                        #region //判斷原單據是否存在
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT TOP 1 TA006
                                                FROM INVTA
                                                WHERE TA001 = @ErpPrefix
                                                AND TA002 = @ErpNo";
                                        dynamicParameters.Add("ErpPrefix", erpPrefix);
                                        dynamicParameters.Add("ErpNo", erpNo);

                                        var result1109Exist = sqlConnection.Query(sql, dynamicParameters);
                                        if (result1109Exist.Count() <= 0) throw new SystemException("原單據不存在，【拋轉模式】請選擇開立新單據!");

                                        foreach (var item in result1109Exist)
                                        {
                                            if (item.TA006 != "N") throw new SystemException("原單據確認碼為無法更動狀態!");
                                        }
                                        #endregion

                                        #region //刪除原單據單身
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"DELETE INVTB
                                                WHERE TB001 = @ErpPrefix
                                                AND TB002 = @ErpNo";
                                        dynamicParameters.Add("ErpPrefix", erpPrefix);
                                        dynamicParameters.Add("ErpNo", erpNo);

                                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                        #endregion

                                        #region //更新原單據
                                        sql = @"UPDATE INVTA SET
                                                TA003 = @TA003,
                                                TA004 = @TA004,
                                                TA005 = @TA005,
                                                TA006 = @TA006,
                                                TA007 = @TA007,
                                                TA008 = @TA008,
                                                TA009 = @TA009,
                                                TA010 = @TA010,
                                                TA011 = @TA011,
                                                TA012 = @TA012,
                                                TA013 = @TA013,
                                                TA014 = @TA014,
                                                TA015 = @TA015,
                                                TA016 = @TA016,
                                                TA017 = @TA017,
                                                TA018 = @TA018,
                                                TA019 = @TA019,
                                                TA020 = @TA020,
                                                TA021 = @TA021,
                                                TA022 = @TA022,
                                                TA023 = @TA023,
                                                TA024 = @TA024,
                                                TA025 = @TA025,
                                                TA026 = @TA026,
                                                TA027 = @TA027,
                                                TA028 = @TA028,
                                                UDF01 = @UDF01,
                                                UDF02 = @UDF02,
                                                UDF03 = @UDF03,
                                                UDF04 = @UDF04,
                                                UDF05 = @UDF05,
                                                UDF06 = @UDF06,
                                                UDF07 = @UDF07,
                                                UDF08 = @UDF08,
                                                UDF09 = @UDF09,
                                                UDF10 = @UDF10
                                                WHERE TA001 = @TA001
                                                AND TA002 = @TA002";
                                        rowsAffected += sqlConnection.Execute(sql, iNVTAs);
                                        #endregion
                                        break;
                                    default:
                                        throw new SystemException("拋轉模式錯誤!");
                                }
                                #endregion

                                #region //INVTB
                                iNVTBs
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.TB001 = erpPrefix; //單別
                                        x.TB002 = erpNo; //單號
                                        x.TB009 = 0; //庫存數量
                                        x.TB010 = 0; //單位成本
                                        x.TB011 = 0; //金額
                                        x.TB013 = ""; //轉入庫
                                        x.TB014 = ""; //批號
                                        x.TB015 = ""; //有效日期
                                        x.TB016 = ""; //複檢日期
                                        x.TB018 = "N"; //確認碼
                                        x.TB020 = ""; //小單位
                                        x.TB021 = ""; //專案代號
                                        x.TB022 = 0; //包裝數量
                                        x.TB023 = ""; //包裝單位
                                        x.TB024 = "N"; //保稅碼
                                        x.TB025 = 0; //產品序號數量
                                        x.TB026 = ""; //轉出儲位
                                        x.TB027 = ""; //轉入儲位
                                        x.TB028 = 0; //預留欄位
                                        x.TB029 = 0; //預留欄位
                                        x.TB030 = ""; //預留欄位
                                        x.TB031 = ""; //預留欄位
                                        x.TB032 = ""; //預留欄位
                                        x.TB500 = ""; //刻號/BIN記錄
                                        x.TB501 = ""; //刻號管理
                                        x.TB502 = ""; //DATECODE
                                        x.UDF01 = "";
                                        x.UDF02 = "";
                                        x.UDF03 = "";
                                        x.UDF04 = "";
                                        x.UDF05 = "";
                                        x.UDF06 = 0;
                                        x.UDF07 = 0;
                                        x.UDF08 = 0;
                                        x.UDF09 = 0;
                                        x.UDF10 = 0;
                                    });

                                sql = @"INSERT INTO INVTB (COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE
                                        , TB001, TB002, TB003, TB004, TB005, TB006, TB007, TB008, TB009, TB010
                                        , TB011, TB012, TB013, TB014, TB015, TB016, TB017, TB018, TB019, TB020
                                        , TB021, TB022, TB023, TB024, TB025, TB026, TB027, TB028, TB029, TB030
                                        , TB031, TB032, TB500, TB501, TB502
                                        , FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID
                                        , UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10)
                                        VALUES (@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE
                                        , @TB001, @TB002, @TB003, @TB004, @TB005, @TB006, @TB007, @TB008, @TB009, @TB010
                                        , @TB011, @TB012, @TB013, @TB014, @TB015, @TB016, @TB017, @TB018, @TB019, @TB020
                                        , @TB021, @TB022, @TB023, @TB024, @TB025, @TB026, @TB027, @TB028, @TB029, @TB030
                                        , @TB031, @TB032, @TB500, @TB501, @TB502
                                        , @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID, @MODI_TIME, @MODI_AP, @MODI_PRID
                                        , @UDF01, @UDF02, @UDF03, @UDF04, @UDF05, @UDF06, @UDF07, @UDF08, @UDF09, @UDF10)";
                                rowsAffected += sqlConnection.Execute(sql, iNVTBs);
                                #endregion
                                #endregion
                                break;
                            default:
                                throw new SystemException("單據類型錯誤!");
                        }

                        #region //如開新單據，作廢原單據
                        if (transferMode == "N")
                        {
                            if (erpDocStatuses.Count() > 0)
                            {
                                switch (docType)
                                {
                                    case "A":
                                        #region //暫出歸還單
                                        foreach (var doc in erpDocStatuses)
                                        {
                                            #region //判斷原單據是否存在
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"SELECT TOP 1 TH020
                                                    FROM INVTH
                                                    WHERE TH001 = @ErpPrefix
                                                    AND TH002 = @ErpNo";
                                            dynamicParameters.Add("ErpPrefix", doc.ErpPrefix);
                                            dynamicParameters.Add("ErpNo", doc.ErpNo);

                                            var result1501Exist = sqlConnection.Query(sql, dynamicParameters);

                                            if (result1501Exist.Count() > 0)
                                            {
                                                foreach (var item in result1501Exist)
                                                {
                                                    if (item.TH020 != "N") throw new SystemException("原單據確認碼為無法更動狀態!");
                                                }

                                                int tempRowsAffected = 0;
                                                #region //作廢
                                                dynamicParameters = new DynamicParameters();
                                                sql = @"UPDATE INVTH SET
                                                        TH020 = @AbandonStatus
                                                        WHERE TH001 = @ErpPrefix
                                                        AND TH002 = @ErpNo";
                                                dynamicParameters.AddDynamicParams(
                                                    new
                                                    {
                                                        AbandonStatus = "V",
                                                        doc.ErpPrefix,
                                                        doc.ErpNo
                                                    });

                                                tempRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                                if (tempRowsAffected != 1) throw new SystemException(string.Format("{0}-{1} 作廢失敗!", erpPrefix, erpNo));
                                                rowsAffected += tempRowsAffected;
                                                #endregion
                                            }
                                            #endregion
                                        }
                                        #endregion
                                        break;
                                    case "B":
                                    case "C":
                                        #region //客供入料單 && 0成本入庫單
                                        foreach (var doc in erpDocStatuses)
                                        {
                                            #region //判斷原單據是否存在
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"SELECT TOP 1 TA006
                                                    FROM INVTA
                                                    WHERE TA001 = @ErpPrefix
                                                    AND TA002 = @ErpNo";
                                            dynamicParameters.Add("ErpPrefix", doc.ErpPrefix);
                                            dynamicParameters.Add("ErpNo", doc.ErpNo);

                                            var result110XExist = sqlConnection.Query(sql, dynamicParameters);

                                            if (result110XExist.Count() > 0)
                                            {
                                                foreach (var item in result110XExist)
                                                {
                                                    if (item.TA006 != "N") throw new SystemException("原單據確認碼為無法更動狀態!");
                                                }

                                                int tempRowsAffected = 0;
                                                #region //作廢
                                                dynamicParameters = new DynamicParameters();
                                                sql = @"UPDATE INVTA SET
                                                        TA006 = @AbandonStatus
                                                        WHERE TA001 = @ErpPrefix
                                                        AND TA002 = @ErpNo";
                                                dynamicParameters.AddDynamicParams(
                                                    new
                                                    {
                                                        AbandonStatus = "V",
                                                        doc.ErpPrefix,
                                                        doc.ErpNo
                                                    });

                                                tempRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                                if (tempRowsAffected != 1) throw new SystemException(string.Format("{0}-{1} 作廢失敗!", erpPrefix, erpNo));
                                                rowsAffected += tempRowsAffected;
                                                #endregion
                                            }
                                            #endregion
                                        }
                                        #endregion
                                        break;
                                    default:
                                        throw new SystemException("單據類型錯誤!");
                                }
                            }
                        }
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //退(換)貨資料更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.ReturnMerchandiseAuthorization SET
                                ConfirmStatus = @ConfirmStatus,
                                ConfirmUserId = @ConfirmUserId,
                                TransferStatus = @TransferStatus,
                                TransferDate = @TransferDate,
                                ErpPrefix = @ErpPrefix,
                                ErpNo = @ErpNo,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RmaId = @RmaId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmStatus = "Y",
                                ConfirmUserId = CurrentUser,
                                TransferStatus = "Y",
                                TransferDate = CreateDate,
                                ErpPrefix = erpPrefix,
                                ErpNo = erpNo,
                                LastModifiedDate,
                                LastModifiedBy,
                                RmaId
                            });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //退(換)貨拋轉單據紀錄新增
                        if (transferMode == "N")
                        {
                            #region //原紀錄停用
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.RmaTransferLog SET
                                    Status = @Status
                                    WHERE RmaId = @RmaId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    Status = "S",
                                    RmaId
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //新增紀錄
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO SCM.RmaTransferLog (RmaId, ErpPrefix, ErpNo, Status
                                    , CreateDate, CreateBy)
                                    OUTPUT INSERTED.RmaTransferLogId
                                    VALUES (@RmaId, @ErpPrefix, @ErpNo, @Status
                                    , @CreateDate, @CreateBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    RmaId,
                                    ErpPrefix = erpPrefix,
                                    ErpNo = erpNo,
                                    Status = "A",
                                    CreateDate,
                                    CreateBy
                                });

                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();
                            #endregion
                        }
                        #endregion

                        #region //發送Mail
                        switch (docType)
                        {
                            case "C":
                                #region //0成本入庫單
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
                                dynamicParameters.Add("SettingSchema", "ReturnMerchandiseAuthorization");
                                dynamicParameters.Add("SettingNo", "A");

                                var resultMailTemplate = sqlConnection.Query(sql, dynamicParameters);
                                if (resultMailTemplate.Count() <= 0) throw new SystemException("Mail設定錯誤!");
                                #endregion

                                #region //庫存調增資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT FORMAT(b.RmaDate, 'M/d') RmaDate, c.MtlItemNo, c.MtlItemName
                                        , a.RmaQty, a.ItemType, a.FreebieOrSpareQty
                                        , d.UomNo, a.RmaDesc, b.RmaRemark
                                        , ISNULL(e.InventoryNo, 'A06') InventoryNo
                                        FROM SCM.RmaDetail a
                                        INNER JOIN SCM.ReturnMerchandiseAuthorization b ON a.RmaId = b.RmaId
                                        INNER JOIN PDM.MtlItem c ON a.MtlItemId = c.MtlItemId
                                        INNER JOIN PDM.UnitOfMeasure d ON c.InventoryUomId = d.UomId
                                        LEFT JOIN SCM.Inventory e ON a.TargetInventory = e.InventoryId
                                        WHERE a.RmaId = @RmaId";
                                dynamicParameters.Add("RmaId", RmaId);

                                var resultRmaDetail = sqlConnection.Query(sql, dynamicParameters);
                                if (resultRmaDetail.Count() <= 0) throw new SystemException("退(換)貨資料錯誤!");
                                #endregion

                                foreach (var item in resultMailTemplate)
                                {
                                    string mailSubject = item.MailSubject,
                                        mailContent = HttpUtility.UrlDecode(item.MailContent);

                                    #region //Mail內容
                                    string replaceHtml = @"<table border='0' cellpadding='0' cellspacing='0' width='100%'>
                                                             <tr>
                                                               <td bgcolor='#FFFFFF' style='padding: 0 0 20px 0;'>
                                                                 <table align='left' border='1' cellpadding='0' cellspacing='0' width='1325'>
                                                                   <tr>
                                                                     <td align='center' bgcolor='#CCFFFF' style='font-family: 微軟正黑體, sans-serif;' width='50'>日期</td>
                                                                     <td align='center' bgcolor='#CCFFFF' style='font-family: 微軟正黑體, sans-serif;' width='175'>單別</td>
                                                                     <td align='center' bgcolor='#CCFFFF' style='font-family: 微軟正黑體, sans-serif;' width='75'>庫位</td>
                                                                     <td align='center' bgcolor='#CCFFFF' style='font-family: 微軟正黑體, sans-serif;' width='75'>調整</td>
                                                                     <td align='center' bgcolor='#CCFFFF' style='font-family: 微軟正黑體, sans-serif;' width='225'>品號</td>
                                                                     <td align='center' bgcolor='#CCFFFF' style='font-family: 微軟正黑體, sans-serif;' width='275'>品名</td>
                                                                     <td align='center' bgcolor='#CCFFFF' style='font-family: 微軟正黑體, sans-serif;' width='75'>數量</td>
                                                                     <td align='center' bgcolor='#CCFFFF' style='font-family: 微軟正黑體, sans-serif;' width='75'>單位</td>
                                                                     <td align='center' bgcolor='#CCFFFF' style='font-family: 微軟正黑體, sans-serif;' width='300'>備註</td>
                                                                   </tr>";
                                    foreach (var rma in resultRmaDetail)
                                    {
                                        replaceHtml += @"          <tr>
                                                                     <td align='center' style='font-family: 微軟正黑體, sans-serif;' width='50'>" + rma.RmaDate + @"</td>
                                                                     <td align='center' style='font-family: 微軟正黑體, sans-serif; color: #FF0000;' width='175'>#1109(0成本入庫單)</td>
                                                                     <td align='center' style='font-family: 微軟正黑體, sans-serif;' width='75'>" + rma.InventoryNo + @"</td>
                                                                     <td align='center' style='font-family: 微軟正黑體, sans-serif;' width='75'>調增</td>
                                                                     <td align='center' style='font-family: 微軟正黑體, sans-serif;' width='225'>" + rma.MtlItemNo + @"</td>
                                                                     <td align='center' style='font-family: 微軟正黑體, sans-serif;' width='275'>" + rma.MtlItemName + @"</td>
                                                                     <td align='center' style='font-family: 微軟正黑體, sans-serif;' width='75'>" + rma.RmaQty + @"</td>
                                                                     <td align='center' style='font-family: 微軟正黑體, sans-serif;' width='75'>" + rma.UomNo + @"</td>
                                                                     <td align='center' style='font-family: 微軟正黑體, sans-serif;' width='300'>" + rma.RmaDesc + @"</td>
                                                                   </tr>";

                                        mailSubject = mailSubject.Replace("[Date]", rma.RmaDate);
                                        mailContent = mailContent.Replace("[Reason]", rma.RmaRemark);
                                    }

                                    replaceHtml += @"             </table>
                                                               </td>
                                                             </tr>
                                                           </table>";

                                    mailContent = mailContent.Replace("[CustomContent]", replaceHtml);
                                    mailContent = mailContent.Replace("[ErpFullNo]", erpPrefix + "-" + erpNo);
                                    mailContent = mailContent.Replace("[ApplyUser]", userNo + " " + userName);
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
                                break;
                            default:
                                break;
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

        #region //UpdateReturnMerchandiseAuthorizationReConfirm -- 退(換)貨資料反確認 -- Ben Ma 2023.03.22
        public string UpdateReturnMerchandiseAuthorizationReConfirm(int RmaId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0, tempRowsAffected = 0, totalRows = 0;
                    string companyNo = "", docType = "", erpPrefix = "", erpNo = "", docDate = "";
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //公司別資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var resultCompany = sqlConnection.Query(sql, dynamicParameters);
                        if (resultCompany.Count() <= 0) throw new SystemException("公司別資料錯誤!");

                        foreach (var item in resultCompany)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            companyNo = item.ErpNo;
                        }
                        #endregion

                        #region //判斷退(換)貨資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 DocType, ConfirmStatus, ErpPrefix, ErpNo, FORMAT(RmaDate, 'yyyyMMdd') RmaDate
                                FROM SCM.ReturnMerchandiseAuthorization
                                WHERE RmaId = @RmaId";
                        dynamicParameters.Add("RmaId", RmaId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("退(換)貨資料錯誤!");

                        #region //判斷確認過無法再確認
                        string confirmStatus = "";
                        foreach (var item in result)
                        {
                            docType = item.DocType;
                            confirmStatus = item.ConfirmStatus;
                            erpPrefix = item.ErpPrefix;
                            erpNo = item.ErpNo;
                            docDate = item.RmaDate;
                        }

                        if (confirmStatus == "N") throw new SystemException("退貨單尚未確認，無法反確認!");
                        #endregion
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        switch (docType)
                        {
                            case "A":
                                #region //暫出歸還單
                                #region //判斷單據是否已經確認
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM INVTH
                                        WHERE TH001 = @ErpPrefix
                                        AND TH002 = @ErpNo
                                        AND TH020 != @ConfirmStatus";
                                dynamicParameters.Add("ErpPrefix", erpPrefix);
                                dynamicParameters.Add("ErpNo", erpNo);
                                dynamicParameters.Add("ConfirmStatus", "N");

                                var resultInvthConfirm = sqlConnection.Query(sql, dynamicParameters);
                                if (resultInvthConfirm.Count() > 0) throw new SystemException("ERP單據已經確認或作廢!");
                                #endregion

                                #region //單身數量
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT COUNT(1) TotalRows
                                        FROM INVTI
                                        WHERE TI001 = @ErpPrefix
                                        AND TI002 = @ErpNo";
                                dynamicParameters.Add("ErpPrefix", erpPrefix);
                                dynamicParameters.Add("ErpNo", erpNo);

                                var resultInvti = sqlConnection.Query(sql, dynamicParameters);
                                if (resultInvti.Count() <= 0) throw new SystemException("ERP單據單身數量錯誤!");

                                foreach (var item in resultInvti)
                                {
                                    totalRows = Convert.ToInt32(item.TotalRows);
                                }
                                #endregion

                                #region //判斷是否有接續單據
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM INVTH
                                        WHERE TH001 = @ErpPrefix
                                        AND TH002 > @ErpNo
                                        AND TH023 = @DocDate";
                                dynamicParameters.Add("ErpPrefix", erpPrefix);
                                dynamicParameters.Add("ErpNo", erpNo);
                                dynamicParameters.Add("DocDate", docDate);

                                var resultInvthContinue = sqlConnection.Query(sql, dynamicParameters);
                                #endregion

                                #region //刪除或是作廢ERP單據
                                //if (resultInvthContinue.Count() <= 0)
                                //{
                                //    #region //刪除
                                //    dynamicParameters = new DynamicParameters();
                                //    sql = @"DELETE INVTH
                                //            WHERE TH001 = @ErpPrefix
                                //            AND TH002 = @ErpNo";
                                //    dynamicParameters.Add("ErpPrefix", erpPrefix);
                                //    dynamicParameters.Add("ErpNo", erpNo);

                                //    tempRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                //    if (tempRowsAffected != 1) throw new SystemException(string.Format("{0}-{1} 刪除失敗!", erpPrefix, erpNo));
                                //    rowsAffected += tempRowsAffected;

                                //    dynamicParameters = new DynamicParameters();
                                //    sql = @"DELETE INVTI
                                //            WHERE TI001 = @ErpPrefix
                                //            AND TI002 = @ErpNo";
                                //    dynamicParameters.Add("ErpPrefix", erpPrefix);
                                //    dynamicParameters.Add("ErpNo", erpNo);

                                //    tempRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                //    if (tempRowsAffected != totalRows) throw new SystemException(string.Format("{0}-{1} 單身刪除失敗!", erpPrefix, erpNo));
                                //    rowsAffected += tempRowsAffected;
                                //    #endregion
                                //}
                                //else
                                //{
                                //    #region //作廢
                                //    dynamicParameters = new DynamicParameters();
                                //    sql = @"UPDATE INVTH SET
                                //            TH020 = @AbandonStatus
                                //            WHERE TH001 = @ErpPrefix
                                //            AND TH002 = @ErpNo";
                                //    dynamicParameters.AddDynamicParams(
                                //        new
                                //        {
                                //            AbandonStatus = "V",
                                //            ErpPrefix = erpPrefix,
                                //            ErpNo = erpNo
                                //        });

                                //    tempRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                //    if (tempRowsAffected != 1) throw new SystemException(string.Format("{0}-{1} 作廢失敗!", erpPrefix, erpNo));
                                //    rowsAffected += tempRowsAffected;
                                //    #endregion
                                //}
                                #endregion
                                #endregion
                                break;
                            case "B":
                            case "C":
                                #region //客供入料單 && 0成本入庫單
                                #region //判斷單據是否已經確認
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM INVTA
                                        WHERE TA001 = @ErpPrefix
                                        AND TA002 = @ErpNo
                                        AND TA006 != @ConfirmStatus";
                                dynamicParameters.Add("ErpPrefix", erpPrefix);
                                dynamicParameters.Add("ErpNo", erpNo);
                                dynamicParameters.Add("ConfirmStatus", "N");

                                var resultInvtaConfirm = sqlConnection.Query(sql, dynamicParameters);
                                if (resultInvtaConfirm.Count() > 0) throw new SystemException("ERP單據已經確認或作廢!");
                                #endregion

                                #region //單身數量
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT COUNT(1) TotalRows
                                        FROM INVTB
                                        WHERE TB001 = @ErpPrefix
                                        AND TB002 = @ErpNo";
                                dynamicParameters.Add("ErpPrefix", erpPrefix);
                                dynamicParameters.Add("ErpNo", erpNo);

                                var resultInvtb = sqlConnection.Query(sql, dynamicParameters);
                                if (resultInvtb.Count() <= 0) throw new SystemException("ERP單據單身數量錯誤!");

                                foreach (var item in resultInvtb)
                                {
                                    totalRows = Convert.ToInt32(item.TotalRows);
                                }
                                #endregion

                                #region //判斷是否有接續單據
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM INVTA
                                        WHERE TA001 = @ErpPrefix
                                        AND TA002 > @ErpNo
                                        AND TA014 = @DocDate";
                                dynamicParameters.Add("ErpPrefix", erpPrefix);
                                dynamicParameters.Add("ErpNo", erpNo);
                                dynamicParameters.Add("DocDate", docDate);

                                var resultInvtaContinue = sqlConnection.Query(sql, dynamicParameters);
                                #endregion

                                #region //刪除或是作廢ERP單據
                                //if (resultInvtaContinue.Count() <= 0)
                                //{
                                //    #region //刪除
                                //    dynamicParameters = new DynamicParameters();
                                //    sql = @"DELETE INVTA
                                //            WHERE TA001 = @ErpPrefix
                                //            AND TA002 = @ErpNo";
                                //    dynamicParameters.Add("ErpPrefix", erpPrefix);
                                //    dynamicParameters.Add("ErpNo", erpNo);

                                //    tempRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                //    if (tempRowsAffected != 1) throw new SystemException(string.Format("{0}-{1} 刪除失敗!", erpPrefix, erpNo));
                                //    rowsAffected += tempRowsAffected;

                                //    dynamicParameters = new DynamicParameters();
                                //    sql = @"DELETE INVTB
                                //            WHERE TB001 = @ErpPrefix
                                //            AND TB002 = @ErpNo";
                                //    dynamicParameters.Add("ErpPrefix", erpPrefix);
                                //    dynamicParameters.Add("ErpNo", erpNo);

                                //    tempRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                //    if (tempRowsAffected != totalRows) throw new SystemException(string.Format("{0}-{1} 單身刪除失敗!", erpPrefix, erpNo));
                                //    rowsAffected += tempRowsAffected;
                                //    #endregion
                                //}
                                //else
                                //{
                                //    #region //作廢
                                //    dynamicParameters = new DynamicParameters();
                                //    sql = @"UPDATE INVTA SET
                                //            TA006 = @AbandonStatus
                                //            WHERE TA001 = @ErpPrefix
                                //            AND TA002 = @ErpNo";
                                //    dynamicParameters.AddDynamicParams(
                                //        new
                                //        {
                                //            AbandonStatus = "V",
                                //            ErpPrefix = erpPrefix,
                                //            ErpNo = erpNo
                                //        });

                                //    tempRowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                //    if (tempRowsAffected != 1) throw new SystemException(string.Format("{0}-{1} 作廢失敗!", erpPrefix, erpNo));
                                //    rowsAffected += tempRowsAffected;
                                //    #endregion
                                //}
                                #endregion
                                #endregion
                                break;
                            default:
                                throw new SystemException("單據類型錯誤!");
                        }
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //退(換)貨資料更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.ReturnMerchandiseAuthorization SET
                                ConfirmStatus = @ConfirmStatus,
                                ConfirmUserId = @ConfirmUserId,
                                TransferMode = @TransferMode,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RmaId = @RmaId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmStatus = "N",
                                ConfirmUserId = CurrentUser,
                                TransferMode = "U",
                                LastModifiedDate,
                                LastModifiedBy,
                                RmaId
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

        #region //UpdateRmaDetail -- 退(換)貨項目資料更新 -- Ben Ma 2023.03.20
        public string UpdateRmaDetail(int RmaDetailId, int MtlItemId, string ItemName
            , int RmaQty, float FreebieOrSpareQty, string RmaDesc, int TargetInventory)
        {
            try
            {
                if (MtlItemId <= 0) throw new SystemException("【品號】不能為空!");
                if (ItemName.Length > 50) throw new SystemException("【部品名稱】長度錯誤!");
                if (RmaQty <= 0) throw new SystemException("【正常品數量】不能為空!");
                if (FreebieOrSpareQty < 0) throw new SystemException("【贈/備品】不能為空!");
                if (RmaDesc.Length > 200) throw new SystemException("【不良項目描述】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷退(換)貨項目資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 ISNULL(a.TsnDetailId, -1) TsnDetailId, b.DocType, b.ConfirmStatus
                                FROM SCM.RmaDetail a
                                INNER JOIN SCM.ReturnMerchandiseAuthorization b ON a.RmaId = b.RmaId
                                WHERE a.RmaDetailId = @RmaDetailId";
                        dynamicParameters.Add("RmaDetailId", RmaDetailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("退(換)貨項目資料錯誤!");

                        #region //判斷單據狀態
                        int tsnDetailId = -1;
                        string docType = "", confirmStatus = "";
                        foreach (var item in result)
                        {
                            tsnDetailId = Convert.ToInt32(item.TsnDetailId);
                            docType = item.DocType;
                            confirmStatus = item.ConfirmStatus;
                        }

                        if (confirmStatus == "Y") throw new SystemException("退貨單已經確認，無法刪除!");
                        #endregion
                        #endregion

                        #region //判斷品號資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM PDM.MtlItem
                                WHERE MtlItemId = @MtlItemId";
                        dynamicParameters.Add("MtlItemId", MtlItemId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("品號資料錯誤!");
                        #endregion

                        #region //判斷目標庫別資料是否正確
                        if (TargetInventory > 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM SCM.Inventory
                                    WHERE InventoryId = @InventoryId";
                            dynamicParameters.Add("InventoryId", TargetInventory);

                            var result3 = sqlConnection.Query(sql, dynamicParameters);
                            if (result3.Count() <= 0) throw new SystemException("目標庫別資料錯誤!");
                        }
                        #endregion

                        #region //暫出歸還單據數量控管
                        if (docType == "A")
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.MtlItemId, (d.TotalTsn - d.TotalSale - d.TotalReturn) LastTsnQty, a.ProductType
                                    , CASE WHEN a.ProductType = '1' THEN TsnFreebieQty ELSE TsnSpareQty END FreebieOrSpareQty
                                    , (ISNULL(d.TsnFreebieQty, 0) - ISNULL(d.TsnReturnFSFreebieQty, 0)) TsnFreebieQty
                                    , (ISNULL(d.TsnSpareQty, 0) - ISNULL(d.TsnReturnFSSpareQty, 0)) TsnSpareQty
                                    FROM SCM.TsnDetail a
                                    INNER JOIN SCM.TempShippingNote b ON a.TsnId = b.TsnId
                                    LEFT JOIN SCM.SoDetail c ON a.SoDetailId = c.SoDetailId
                                    OUTER APPLY (
                                        SELECT da.ProductType, SUM(da.TsnQty) TotalTsn
                                        , CASE WHEN da.ProductType = '1' THEN SUM(da.FreebieOrSpareQty) END TsnFreebieQty
                                        , CASE WHEN da.ProductType = '2' THEN SUM(da.FreebieOrSpareQty) END TsnSpareQty
                                        , SUM(da.SaleQty) TotalSale
                                        , CASE WHEN da.ProductType = '1' THEN SUM(da.SaleFSQty) END TsnSaleFSFreebieQty
                                        , CASE WHEN da.ProductType = '2' THEN SUM(da.SaleFSQty) END TsnSaleFSSpareQty
                                        , SUM(da.ReturnQty) TotalReturn
                                        , CASE WHEN da.ProductType = '1' THEN SUM(da.ReturnFSQty) END TsnReturnFSFreebieQty
                                        , CASE WHEN da.ProductType = '2' THEN SUM(da.ReturnFSQty) END TsnReturnFSSpareQty
                                        FROM SCM.TsnDetail da
                                        WHERE da.SoDetailId = c.SoDetailId
                                        AND da.TsnDetailId = a.TsnDetailId
                                        AND da.ConfirmStatus = @ConfirmStatus
                                        AND da.ClosureStatus = @ClosureStatus
                                        GROUP BY da.ProductType
                                    ) d
                                    WHERE a.TsnDetailId = @TsnDetailId
                                    AND a.ConfirmStatus = @ConfirmStatus
                                    AND b.ConfirmStatus = @ConfirmStatus";
                            dynamicParameters.Add("TsnDetailId", tsnDetailId);
                            dynamicParameters.Add("ConfirmStatus", "Y");
                            dynamicParameters.Add("ClosureStatus", "N");

                            var resultTsn = sqlConnection.Query(sql, dynamicParameters);
                            if (resultTsn.Count() <= 0) throw new SystemException("暫出單資料錯誤!");

                            int mtlItemId = -1;
                            double lastTsnQty = 0, freebieOrSpareQty = 0;
                            foreach (var item in resultTsn)
                            {
                                mtlItemId = Convert.ToInt32(item.MtlItemId);
                                lastTsnQty = Convert.ToDouble(item.LastTsnQty);
                                freebieOrSpareQty = Convert.ToDouble(item.FreebieOrSpareQty);
                            }
                            if (MtlItemId != mtlItemId) throw new SystemException("暫出單與退貨品號不同!");
                            if (RmaQty > lastTsnQty) throw new SystemException("暫出【正常品】剩餘數量不足!");
                            if (FreebieOrSpareQty > freebieOrSpareQty) throw new SystemException("暫出【贈/備品】剩餘數量不足!");
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.RmaDetail SET
                                MtlItemId = @MtlItemId,
                                ItemName = @ItemName,
                                RmaQty = @RmaQty,
                                FreebieOrSpareQty = @FreebieOrSpareQty,
                                RmaDesc = @RmaDesc,
                                TargetInventory = @TargetInventory,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE RmaDetailId = @RmaDetailId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                MtlItemId,
                                ItemName,
                                RmaQty,
                                FreebieOrSpareQty,
                                RmaDesc,
                                TargetInventory = TargetInventory <= 0 ? (int?)null : TargetInventory,
                                LastModifiedDate,
                                LastModifiedBy,
                                RmaDetailId
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
        #region //DeleteReturnMerchandiseAuthorization -- 退(換)貨資料刪除 -- Ben Ma 2023.03.20
        public string DeleteReturnMerchandiseAuthorization(int RmaId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷退(換)貨資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 ConfirmStatus, TransferStatus
                                FROM SCM.ReturnMerchandiseAuthorization
                                WHERE RmaId = @RmaId";
                        dynamicParameters.Add("RmaId", RmaId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("退(換)貨資料錯誤!");

                        #region //判斷確認過無法再確認
                        string confirmStatus = "", transferStatus = "";
                        foreach (var item in result)
                        {
                            confirmStatus = item.ConfirmStatus;
                            transferStatus = item.TransferStatus;
                        }

                        if (confirmStatus == "Y") throw new SystemException("退貨單已經確認，無法刪除!");
                        if (transferStatus == "Y") throw new SystemException("退貨單已經拋轉，無法刪除!");
                        #endregion
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除關聯table
                        #region //退(換)貨物件刪除
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE a
                                FROM SCM.RmaItem a
                                INNER JOIN SCM.RmaDetail b ON a.RmaDetailId = b.RmaDetailId
                                WHERE b.RmaId = @RmaId";
                        dynamicParameters.Add("RmaId", RmaId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //退(換)貨項目刪除
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.RmaDetail
                                WHERE RmaId = @RmaId";
                        dynamicParameters.Add("RmaId", RmaId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.ReturnMerchandiseAuthorization
                                WHERE RmaId = @RmaId";
                        dynamicParameters.Add("RmaId", RmaId);

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

        #region //DeleteRmaDetail -- 退(換)貨項目資料刪除 -- Ben Ma 2023.03.20
        public string DeleteRmaDetail(int RmaDetailId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷退(換)貨項目資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 b.ConfirmStatus, b.TransferStatus
                                FROM SCM.RmaDetail a
                                INNER JOIN SCM.ReturnMerchandiseAuthorization b ON a.RmaId = b.RmaId
                                WHERE a.RmaDetailId = @RmaDetailId";
                        dynamicParameters.Add("RmaDetailId", RmaDetailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("退(換)貨項目資料錯誤!");

                        #region //判斷確認過無法再刪除
                        string confirmStatus = "", transferStatus = "";
                        foreach (var item in result)
                        {
                            confirmStatus = item.ConfirmStatus;
                            transferStatus = item.TransferStatus;
                        }

                        if (confirmStatus == "Y") throw new SystemException("退貨單已經確認，無法刪除!");
                        if (transferStatus == "Y") throw new SystemException("退貨單已經拋轉，無法刪除!");
                        #endregion
                        #endregion

                        int rowsAffected = 0;
                        #region //刪除關聯table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.RmaItem
                                WHERE RmaDetailId = @RmaDetailId";
                        dynamicParameters.Add("RmaDetailId", RmaDetailId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.RmaDetail
                                WHERE RmaDetailId = @RmaDetailId";
                        dynamicParameters.Add("RmaDetailId", RmaDetailId);

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
    }
}
