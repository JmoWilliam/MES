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
using System.Text.RegularExpressions;
using System.Transactions;
using System.Web;
using static Dapper.SqlMapper;

namespace SCMDA
{
    public class PurchaseOrderDA
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

        public PurchaseOrderDA()
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
                if (HttpContext.Current.Session["UserLock"] == null)
                {
                    HttpCookie Login = HttpContext.Current.Request.Cookies.Get("Login");
                    HttpCookie LoginKey = HttpContext.Current.Request.Cookies.Get("LoginKey");

                    if (Login == null || LoginKey == null)
                    {
                        CurrentCompany = -1;
                        CurrentUser = -1;
                        CreateBy = -1;
                        LastModifiedBy = -1;
                    }
                    else
                    {
                        using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                        {
                            sql = @"SELECT a.UserId
                                    , d.CompanyId
                                    FROM BAS.UserLoginKey a
                                    INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                                    INNER JOIN BAS.Department c ON b.DepartmentId = c.DepartmentId
                                    INNER JOIN BAS.Company d ON c.CompanyId = d.CompanyId
                                    WHERE 1=1
                                    AND b.[Status] = @Status
                                    AND c.[Status] = @Status
                                    AND d.[Status] = @Status
                                    AND a.LoginIP = @LoginIP
                                    AND a.KeyText = @KeyText
                                    AND b.UserNo = @UserNo";
                            dynamicParameters.Add("Status", "A");
                            dynamicParameters.Add("LoginIP", BaseHelper.ClientIP());
                            dynamicParameters.Add("KeyText", LoginKey.Value);
                            dynamicParameters.Add("UserNo", Login.Value);
                            var result = sqlConnection.Query(sql, dynamicParameters);

                            if (result.Count() < 0) throw new SystemException("使用者資料錯誤!");
                            foreach (var item in result)
                            {
                                CurrentCompany = Convert.ToInt32(item.CompanyId);
                                CurrentUser = Convert.ToInt32(item.UserId);
                                CreateBy = Convert.ToInt32(item.UserId);
                                LastModifiedBy = Convert.ToInt32(item.UserId);

                                HttpContext.Current.Session["UserLock"] = true;
                                HttpContext.Current.Session["UserCompany"] = Convert.ToInt32(item.CompanyId);
                                HttpContext.Current.Session["UserId"] = Convert.ToInt32(item.UserId);
                                HttpContext.Current.Session["UserNo"] = Login.Value;
                            }
                        }
                    }
                }
                else
                {
                    CurrentCompany = Convert.ToInt32(HttpContext.Current.Session["UserCompany"]);
                    CurrentUser = Convert.ToInt32(HttpContext.Current.Session["UserId"]);
                    CreateBy = Convert.ToInt32(HttpContext.Current.Session["UserId"]);
                    LastModifiedBy = Convert.ToInt32(HttpContext.Current.Session["UserId"]);
                }

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
        #region //GetMaterial -- 取得材料資料 -- Shintokuro 2024-08-19
        public string GetMaterial(int MtId, string MaterialName, string MtlItemNo, int SupplierId, string Status, string CanUse,
            string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.MtId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.MaterialName, a.SupplierId, a.MtlItemId, a.Remark, a.Status
                          , b.SupplierName
                          , ISNULL(c.MtlItemNo,'') MtlItemNo, c.MtlItemName, c.MtlItemSpec
                          , x.ConfirmStatus
                        ";
                    sqlQuery.mainTables =
                        @"FROM SCM.MaterialManagement a 
                        LEFT JOIN SCM.Supplier b ON a.SupplierId = b.SupplierId
                        LEFT JOIN PDM.MtlItem c ON a.MtlItemId = c.MtlItemId
                        OUTER APPLY(
                            SELECT TOP 1 'Y' ConfirmStatus
                            FROM SCM.MtAmount x1
                            WHERE a.MtId = x1.MtId
                            AND ConfirmStatus != 'N'
                        ) x
                        ";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtId", @" AND a.MtId = @MtId", MtId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MaterialName", @" AND a.MaterialName LIKE '%' + @MaterialName + '%'", MaterialName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemNo", @" AND a.MtlItemNo LIKE @MtlItemNo", MtlItemNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SupplierId", @" AND a.SupplierId = @SupplierId", SupplierId);
                    if (Status.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status IN @Status", Status.Split(','));
                    if (CanUse == "Y") BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CanUse", @" AND x.ConfirmStatus is not null", CanUse);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.MtId DESC";
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

        #region //GetMtAmount -- 取得材料資料 -- Shintokuro 2024-08-19
        public string GetMtAmount(int MtId, int MaId, string ConfirmStatus,
            string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.MaId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.Amount,a.UnitQuantity, a.ConfirmStatus, ISNULL(FORMAT(a.ConfirmDate, 'yyyy-MM-dd'), '') ConfirmDate
                          , ISNULL(b.UserName, '')  ConfirmUserName
                        ";
                    sqlQuery.mainTables =
                        @"FROM SCM.MtAmount a 
                        INNER JOIN SCM.MaterialManagement a1 ON a.MtId = a1.MtId
                        LEFT JOIN SCM.Supplier a2 ON a1.SupplierId = a2.SupplierId
                        LEFT JOIN BAS.[User] b ON a.ConfirmUserId = b.UserId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a1.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtId", @" AND a.MtId = @MtId", MtId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MaId", @" AND a.MaId = @MaId", MaId);
                    if (ConfirmStatus.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ConfirmStatus", @" AND a.ConfirmStatus IN @ConfirmStatus", ConfirmStatus.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.MaId DESC";
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

        #region //GetMtAmountERP -- 取得採購材料金額(ERP) -- Shintokuro 2024-08-20
        public string GetMtAmountERP(string MaterialName)
        {
            try
            {
                if (MaterialName.Length <= 0) throw new SystemException("【材料】不能為空!");

                string ErpNo = "";
                string ErpDbName = "";
                string MtlItemNo = "";
                double AmountMes = -1;
                double AmountErp = -1;
                List<dynamic> Amount = new List<dynamic>();
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {

                    DynamicParameters dynamicParameters = new DynamicParameters();
                    int MtlItemId = -1;
                    #region //判斷材料資訊是否正確
                    sql = @"SELECT TOP 1 a.Status,ISNULL(a.MtlItemId,-1) MtlItemId,ISNULL(x.Amount,0) Amount,x.Amount AmountStatus
                            FROM SCM.MaterialManagement a
                            OUTER APPLY(
                                SELECT TOP 1 x1.Amount
                                FROM SCM.MtAmount x1 
                                WHERE x1.MtId = a.MtId
                                AND x1.ConfirmStatus = 'Y'
                                ORDER BY x1.MaId DESC
                            ) x
                            WHERE a.MaterialName = @MaterialName";
                    dynamicParameters.Add("MaterialName", MaterialName);

                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0) throw new SystemException("【材料】找不到，請重新確認!");
                    foreach(var item in result)
                    {
                        if(item.Status != "A") throw new SystemException("【材料】已停用，請重新確認!");
                        if(item.AmountStatus == null) throw new SystemException("【材料】尚未維護金額，請重新確認!");
                        MtlItemId = item.MtlItemId;
                        AmountMes = Convert.ToDouble(item.Amount);
                    }
                    #endregion

                    if (MtlItemId > 0)
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
                            ErpDbName = ErpConnectionStrings.Split('=')[2].Split(';')[0];
                            ErpNo = item.ErpNo;
                        }
                        #endregion

                        #region //判斷品號是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MtlItemNo
                                FROM PDM.MtlItem a
                                WHERE MtlItemId = @MtlItemId";
                        dynamicParameters.Add("MtlItemId", MtlItemId);

                        var mtlItemResult = sqlConnection.Query(sql, dynamicParameters);

                        if (mtlItemResult.Count() <= 0) throw new SystemException("【品號】找不到，請重新確認!");

                        foreach (var item in mtlItemResult)
                        {
                            MtlItemNo = item.MtlItemNo;
                        }
                        #endregion
                    }
                    else
                    {
                        Amount.Add(AmountMes);
                        Amount.Add(AmountErp);

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            data = Amount
                        });
                        #endregion

                        return jsonResponse.ToString();
                    }
                }

                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();

                    #region //近期採購金額撈取
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.TD010 Amount
                            FROM PURTD a
                            INNER JOIN PURTC b on a.TD001 = b.TC001 AND a.TD002 = b.TC002
                            WHERE a.TD004 = @MtlItemNo
                            AND a.TD018 = 'Y'
                            ORDER BY b.TC024 DESC ";
                    dynamicParameters.Add("MtlItemNo", MtlItemNo);
                    var purtdresult = sqlConnection.Query(sql, dynamicParameters);
                    if (purtdresult.Count() > 0)
                    {
                        foreach (var item in purtdresult)
                        {

                            AmountErp = Convert.ToDouble(item.Amount);
                        }
                    }
                    #endregion

                    Amount.Add(AmountMes);
                    Amount.Add(AmountErp);

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = Amount
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

        #region //GetPurchaseOrder -- 取得採購單資料 -- Ann 2024-03-04
        public string GetPurchaseOrder(int PoId, string PoErpFullNo, int SupplierId, int PoUserId, string SearchKey, string ConfirmStatus, string ClosureStatus, string StartDate, string EndDate
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.PoId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.CompanyId, a.PoErpPrefix, a.PoErpNo, FORMAT(a.PoDate, 'yyyy-MM-dd') PoDate, a.SupplierId, a.CurrencyCode, a.Exchange, a.PaymentTerm, a.Remark
                        , a.PoUserId, a.ConfirmStatus, a.Taxation, a.PoPrice, a.TaxAmount, a.FirstAddress, a.SecondAddress, a.Quantity, FORMAT(a.DocDate, 'yyyy-MM-dd') DocDate
                        , a.ConfirmUserId, a.TaxRate, a.PaymentTermNo, a.ApproveStatus, a.Edition, a.DepositPartial, a.TaxNo, a.TradeTerm, a.DetailMultiTax
                        , a.ContactUser, a.TransferStatus, FORMAT(a.TransferDate, 'yyyy-MM-dd'), a.PoErpPrefix + '-' + a.PoErpNo PoErpFullNo
                        , b.SupplierNo, b.SupplierName, b.SupplierNo + ' ' + b.SupplierName SupplierFullNo
                        , c.UserNo PoUserNo, c.UserName PoUserName, c.UserNo + ' ' + c.UserName PoUserFullNo
                        , (
                            SELECT TOP 1 1
                            FROM SCM.PoDetail x 
                            INNER JOIN PDM.MtlItem xa ON x.MtlItemId = xa.MtlItemId
                            WHERE x.PoId = a.PoId
                            AND xa.LotManagement IN ('Y', 'T')
                        ) LotManagement";
                    sqlQuery.mainTables =
                        @"FROM SCM.PurchaseOrder a 
                        INNER JOIN SCM.Supplier b ON a.SupplierId = b.SupplierId
                        INNER JOIN BAS.[User] c ON a.PoUserId = c.UserId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PoId", @" AND a.PoId = @PoId", PoId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SupplierId", @" AND a.SupplierId = @SupplierId", SupplierId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PoUserId", @" AND a.PoUserId = @PoUserId", PoUserId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey", @" AND EXISTS (
                                                                                                            SELECT TOP 1 1
                                                                                                            FROM SCM.PoDetail x
                                                                                                            LEFT JOIN PDM.MtlItem xa ON x.MtlItemId = xa.MtlItemId
                                                                                                            WHERE x.PoId = a.PoId
                                                                                                            AND (xa.MtlItemNo LIKE '%' + @SearchKey + '%' OR x.PoMtlItemName LIKE '%' + @SearchKey + '%')
                                                                                                       )", SearchKey);
                    if (ConfirmStatus.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ConfirmStatus", @" AND a.ConfirmStatus IN @ConfirmStatus", ConfirmStatus.Split(','));
                    if (ClosureStatus.Length > 0)
                    {
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ClosureStatus", @" AND EXISTS (
                                                                                                                    SELECT TOP 1 1
                                                                                                                    FROM SCM.PoDetail x
                                                                                                                    WHERE x.PoId = a.PoId
                                                                                                                    AND x.ClosureStatus IN @ClosureStatus
                                                                                                               )", ClosureStatus.Split(','));
                    }
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PoErpFullNo", @" AND (a.PoErpPrefix + '-' + a.PoErpNo) LIKE '%' + @PoErpFullNo + '%'", PoErpFullNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.DocDate >= @StartDate", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.DocDate <= @EndDate", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.PoId DESC";
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

        #region //GetPoDetail -- 取得採購單身資料 -- Ann 2024-03-06
        public string GetPoDetail(int PoDetailId, int PoId, string PoErpFullNo, int SupplierId, int ConfirmUserId, string SearchKey, string ConfirmStatus, string ClosureStatus, string StartDate, string EndDate, string PaymentTermNo, string CurrencyCode
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.PoDetailId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.PoId, a.PoErpPrefix, a.PoErpNo, a.PoSeq, a.MtlItemId, a.PoMtlItemName, a.PoMtlItemSpec, a.InventoryId, a.Quantity 
                        , a.UomId, a.PoUnitPrice, a.PoPrice, FORMAT(a.PromiseDate, 'yyyy-MM-dd') PromiseDate, a.ReferencePrefix, a.Remark, a.SiQty
                        , a.ClosureStatus, a.ConfirmStatus, a.InventoryQty, a.SmallUnit, a.ReferenceNo, a.Project, a.ReferenceSeq, a.UrgentMtl
                        , a.PrErpPrefix, a.PrErpNo, a.PrSequence, a.FromDocType, a.MtlItemType, a.PartialSeq, a.OriPromiseDate, a.DeliveryDate
                        , a.PoPriceQty, a.PoPriceUomId, a.SiPriceQty, a.DiscountRate, a.DiscountAmount
                        , b.PaymentTerm, b.PaymentTermNo, b.CurrencyCode, b.Exchange
                        , c.SupplierNo, c.SupplierName
                        , d.UserNo PoUserNo, d.UserName PoUserName
                        , e.MtlItemNo, e.LotManagement
                        , f.StatusNo QcStatus, f.StatusName QcStatusName";
                    sqlQuery.mainTables =
                        @"FROM SCM.PoDetail a 
                        INNER JOIN SCM.PurchaseOrder b ON a.PoId = b.PoId
                        INNER JOIN SCM.Supplier c ON b.SupplierId = c.SupplierId
                        INNER JOIN BAS.[User] d ON b.PoUserId = d.UserId
                        INNER JOIN PDM.MtlItem e ON a.MtlItemId = e.MtlItemId
                        INNER JOIN BAS.[Status] f ON CASE WHEN e.MeasureType = 0 THEN 0 ELSE 1 END = f.StatusNo AND f.StatusSchema = 'GrDetail.QcStatus'";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND b.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PoId", @" AND a.PoId = @PoId", PoId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SupplierId", @" AND b.SupplierId = @SupplierId", SupplierId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PaymentTermNo", @" AND b.PaymentTermNo = @PaymentTermNo", PaymentTermNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ConfirmStatus", @" AND a.ConfirmStatus = @ConfirmStatus", ConfirmStatus);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ConfirmUserId", @" AND b.ConfirmUserId = @ConfirmUserId", ConfirmUserId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CurrencyCode", @" AND b.CurrencyCode = @CurrencyCode", CurrencyCode);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey", @" AND (e.MtlItemNo LIKE '%' + @SearchKey + '%' OR a.PoMtlItemName LIKE '%' + @SearchKey + '%' OR a.PoMtlItemSpec LIKE '%' + @SearchKey + '%')", SearchKey);
                    if (ConfirmStatus.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ConfirmStatus", @" AND b.ConfirmStatus IN @ConfirmStatus", ConfirmStatus.Split(','));
                    if (ClosureStatus.Length > 0)
                    {
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ClosureStatus", @" AND a.ClosureStatus IN @ClosureStatus", ClosureStatus.Split(','));
                    }
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PoErpFullNo", @" AND (b.PoErpPrefix + '-' + b.PoErpNo) LIKE '%' + @PoErpFullNo + '%'", PoErpFullNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND b.DocDate >= @StartDate", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND b.DocDate <= @EndDate", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.PoDetailId,a.PoSeq";
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

        #region //GetPurchaseOrderFromERP -- Chia Yuan -- 2024.04.16
        public string GetPurchaseOrderFromERP(int CompanyId, int CustomerId, string PoErpPrefix, string PoErpNo, string SearchKey
            , string StartDate, string EndDate
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
                    dynamicParameters.Add("CompanyId", CompanyId);

                    var companyResult = sqlConnection.Query(sql, dynamicParameters);

                    if (companyResult.Count() <= 0) throw new SystemException("請選擇公司別!");

                    foreach (var item in companyResult)
                    {
                        ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                    }
                    #endregion

                    PageSize = 10;
                    List<dynamic> resultErp = new List<dynamic>();
                    using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        if (SearchKey.Length > 0)
                        {
                            #region //取得單據筆數
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 COUNT(a.TD003) AS DetailCount
                                    FROM PURTD a
                                    WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PoErpPrefix", @" AND LTRIM(RTRIM(a.TD001)) = @PoErpPrefix", PoErpPrefix.Trim());
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PoErpNo", @" AND LTRIM(RTRIM(a.TD002)) = @PoErpNo", PoErpNo.Trim());
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SearchKey", @" AND a.TD001 + '-' + a.TD002 + '-' + a.TD003 LIKE '%' + @SearchKey + '%'", SearchKey.Trim());
                            sql += " GROUP BY a.TD001 + a.TD002";
                            var resultCount = erpConnection.Query(sql, dynamicParameters);
                            foreach (var item in resultCount)
                            {
                                PageSize = item.DetailCount;
                            }
                            #endregion
                        }

                        #region //取得客戶採購單
                        dynamicParameters = new DynamicParameters();
                        sqlQuery.mainKey = "c.MainKey";
                        sqlQuery.auxKey = "";
                        sqlQuery.columns =
                            @", a.COMPANY
                            , LTRIM(RTRIM(a.TD001)) AS TD001, LTRIM(RTRIM(a.TD002)) AS TD002, LTRIM(RTRIM(a.TD003)) AS TD003
                            , LTRIM(RTRIM(a.TD001)) + '-' + LTRIM(RTRIM(a.TD002)) AS PoErpPrefixNo
                            , LTRIM(RTRIM(a.TD001)) + '-' + LTRIM(RTRIM(a.TD002)) + '-' + LTRIM(RTRIM(a.TD003)) AS PoErpPrefixSno
                            , LTRIM(RTRIM(a.TD004)) + ' ' + LTRIM(RTRIM(a.TD005)) AS MtlItemNameWithNo
                            , LTRIM(RTRIM(b.ML002)) AS ML002, LTRIM(RTRIM(b.ML003)) AS ML003
                            , LTRIM(RTRIM(a.TD004)) AS TD004, LTRIM(RTRIM(a.TD005)) AS TD005, LTRIM(RTRIM(a.TD006)) AS TD006
                            , a.TD007, a.TD008, a.TD009, a.TD010, a.TD011
                            , a.TD012, a.TD045, a.TD046, a.TD015
                            , a.TD024, a.TD057, a.TD058, a.TD059, a.TD060
                            , a.TD061, a.TD062, a.TD063
                            , SUBSTRING(a.TD045, 1, 4) + '-' + SUBSTRING(a.TD045, 5, 2) + '-' + SUBSTRING(a.TD045, 7, 8) AS PromiseDate";
                        sqlQuery.mainTables =
                            @"FROM PURTD a
                            INNER JOIN CMSML b ON b.COMPANY = a.COMPANY
                            OUTER APPLY(SELECT TOP 1 LTRIM(RTRIM(TD001)) + LTRIM(RTRIM(TD002)) + LTRIM(RTRIM(TD003)) AS MainKey FROM PURTD WHERE TD001 = a.TD001 AND TD002 = a.TD002 AND TD003 = a.TD003) AS c";
                        sqlQuery.auxTables = "";
                        string queryCondition = "";

                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PoErpPrefix", @" AND LTRIM(RTRIM(a.TD001)) = @PoErpPrefix", PoErpPrefix.Trim());
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PoErpNo", @" AND LTRIM(RTRIM(a.TD002)) = @PoErpNo", PoErpNo.Trim());

                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.TD045 >= @StartDate", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyyMMdd") : string.Empty);
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.TD045 <= @EndDate", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).AddDays(1).ToString("yyyyMMdd") : string.Empty);
                        if (SearchKey.Length > 0)
                        {
                            BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey", @" AND (a.TD001 + '-' + a.TD002 + '-' + a.TD003 LIKE '%' + @SearchKey + '%'
                                                                                                OR a.TD004 LIKE '%' + @SearchKey + '%'
                                                                                                OR a.TD005 LIKE '%' + @SearchKey + '%')", SearchKey.Trim());
                        }
                        sqlQuery.conditions = queryCondition;
                        sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.TD001, a.TD002 DESC, a.TD003";
                        sqlQuery.pageIndex = PageIndex;
                        sqlQuery.pageSize = PageSize;

                        resultErp = BaseHelper.SqlQuery(erpConnection, dynamicParameters, sqlQuery).ToList();
                        #endregion
                    }

                    #region //取得客戶部番
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.*
                            FROM (
	                            SELECT a.CustomerMtlItemId, a.CustomerMtlItemNo, a.CustomerMtlItemName, a.[Status] AS CustMtlItemStatus
	                            , a.TransferStatus AS CustMtlItemTransferStatus, s1.StatusName AS CustMtlItemTransferStatusName
	                            , b.MtlItemId, ISNULL(b.MtlItemNo, '') AS MtlItemNo, ISNULL(b.MtlItemName, '') AS MtlItemName, ISNULL(b.MtlItemDesc, '') AS MtlItemDesc, b.[Status] AS MtlItemStatus
	                            , b.TransferStatus AS MtlItemTransferStatus, s2.StatusName AS MtlItemTransferStatusName
                                , ISNULL(u.UomNo, '') AS SaleUomNo, ISNULL(u.UomName, '') AS SaleUomName
	                            , ROW_NUMBER() OVER (PARTITION BY a.CustomerMtlItemNo ORDER BY a.CustomerMtlItemId DESC) AS SortCustMtlItem
	                            FROM PDM.CustomerMtlItem a 
	                            INNER JOIN PDM.MtlItem b ON b.MtlItemId = a.MtlItemId AND b.[Status] = a.[Status]
                                LEFT JOIN PDM.UnitOfMeasure u ON u.UomId = b.SaleUomId
	                            INNER JOIN BAS.[Status] s1 ON s1.StatusNo = a.TransferStatus AND s1.StatusSchema = 'Boolean'
	                            INNER JOIN BAS.[Status] s2 ON s2.StatusNo = b.TransferStatus AND s2.StatusSchema = 'Boolean'
	                            WHERE (a.CustomerMtlItemNo IN @CustomerMtlItems OR a.CustomerMtlItemName IN @CustomerMtlItems)
	                            AND b.[Status] = @Status
                                AND a.CustomerId = @CustomerId
                            ) AS a
                            WHERE a.SortCustMtlItem = 1";
                    dynamicParameters.Add("Status", "A");
                    dynamicParameters.Add("CustomerId", CustomerId);
                    dynamicParameters.Add("CustomerMtlItems", resultErp.Select(s => s.TD004).Concat(resultErp.Select(s => s.TD005)).Distinct().ToList());
                    var resultCustMtlItem = sqlConnection.Query(sql, dynamicParameters).ToList();
                    #endregion

                    #region //資料集合處理
                    var result = (from s1 in resultErp
                                  join s2 in resultCustMtlItem
                                    on s1.TD004 equals s2.CustomerMtlItemNo into a1
                                  from s2 in a1.DefaultIfEmpty()
                                  select new
                                  {
                                      CompanyId,
                                      s1.MainKey,
                                      s1.PoErpPrefixNo,
                                      s1.MtlItemNameWithNo,
                                      s1.COMPANY,
                                      s1.TD001,
                                      s1.TD002,
                                      s1.TD003,
                                      s1.ML002,
                                      s1.ML003,
                                      s1.TD004,
                                      s1.TD005,
                                      s1.TD006,
                                      s1.TD007,
                                      s1.TD008,
                                      s1.TD009,
                                      s1.TD010,
                                      s1.TD011,
                                      s1.TD012,
                                      s1.TD045,
                                      s1.TD046,
                                      s1.TD015,
                                      s1.TD024,
                                      s1.TD057,
                                      s1.TD058,
                                      s1.TD059,
                                      s1.TD060,
                                      s1.TD061,
                                      s1.TD062,
                                      s1.TD063,
                                      s1.PromiseDate,
                                      CustomerMtlItemId = s2?.CustomerMtlItemId ?? "",
                                      CustomerMtlItemNo = s2?.CustomerMtlItemNo ?? "",
                                      CustomerMtlItemName = s2?.CustomerMtlItemName ?? "",
                                      s2?.CustMtlItemStatus,
                                      s2?.CustMtlItemTransferStatus,
                                      s2?.CustMtlItemTransferStatusName,
                                      s2?.MtlItemId,
                                      MtlItemNo = s2?.MtlItemNo ?? "",
                                      MtlItemName = s2?.MtlItemName ?? "",
                                      MtlItemDesc = s2?.MtlItemDesc ?? "",
                                      s2?.MtlItemStatus,
                                      s2?.MtlItemTransferStatus,
                                      s2?.MtlItemTransferStatusName,
                                      SaleUomNo = s2?.SaleUomNo ?? "",
                                      s2?.SaleUomName,
                                      s2?.SortCustMtlItem,
                                      s1.TotalCount,
                                      PageSize
                                  })
                                  .OrderBy(o => o.TD001).ThenByDescending(t => t.TD002).ThenBy(t => t.TD003);
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

        #region GetPurchaseOrderTypeFromERP -- Chia Yuan -- 2024.04.16
        public string GetPurchaseOrderTypeFromERP(int CompanyId, string SearchKey
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
                    dynamicParameters.Add("CompanyId", CompanyId);

                    var companyResult = sqlConnection.Query(sql, dynamicParameters);

                    if (companyResult.Count() <= 0) throw new SystemException("請選擇公司別!");

                    foreach (var item in companyResult)
                    {
                        ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                    }
                    #endregion

                    using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        #region //取得客戶採購單
                        dynamicParameters = new DynamicParameters();
                        sqlQuery.mainKey = "c.MainKey";
                        sqlQuery.auxKey = "";
                        sqlQuery.columns =
                            @", a.COMPANY
                            , LTRIM(RTRIM(a.MQ001)) + ' ' + LTRIM(RTRIM(a.MQ002)) AS NameWithNo
                            , LTRIM(RTRIM(a.MQ001)) AS MQ001
                            , LTRIM(RTRIM(a.MQ002)) AS MQ002
                            , LTRIM(RTRIM(b.ML002)) AS ML002, LTRIM(RTRIM(b.ML003)) AS ML003";
                        sqlQuery.mainTables =
                            @"FROM CMSMQ a
                            INNER JOIN CMSML b ON b.COMPANY = a.COMPANY
                            OUTER APPLY(SELECT DISTINCT LTRIM(RTRIM(MQ001)) AS MainKey FROM CMSMQ WHERE MQ001 = a.MQ001) c";
                        sqlQuery.auxTables = "";
                        string queryCondition = " AND MQ003 = @MQ003";
                        dynamicParameters.Add("MQ003", "33");

                        if (SearchKey.Length > 0)
                        {
                            BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey", @" AND (a.MQ001 LIKE '%' + @SearchKey + '%'
                                                                                                OR a.MQ002 LIKE '%' + @SearchKey + '%')", SearchKey.Trim());
                        }
                        sqlQuery.conditions = queryCondition;
                        sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.MQ001";
                        sqlQuery.pageIndex = PageIndex;
                        sqlQuery.pageSize = PageSize;

                        var result = BaseHelper.SqlQuery(erpConnection, dynamicParameters, sqlQuery);
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
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
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
        #region //AddMaterial -- 新增材料資料 -- Shinotokuro 2024.08.19
        public string AddMaterial(string MaterialName, int SupplierId, string Remark, int MtlItemId)
        {
            try
            {
                if (MaterialName.Length <= 0) throw new SystemException("【材料】不能為空!");
                if (MaterialName.Length > 100) throw new SystemException("【材料】長度錯誤!");
                if (SupplierId <= 0) throw new SystemException("【廠商】不能為空!");
                if (Remark.Length > 200) throw new SystemException("【備註】長度不能超過200!");
                int? nullData = null;
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷廠商是否存在
                        sql = @"SELECT TOP 1 1
                                FROM SCM.Supplier
                                WHERE SupplierId = @SupplierId";
                        dynamicParameters.Add("SupplierId", SupplierId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【廠商】找不到，請重新確認!");
                        #endregion

                        #region //判斷材料名稱是否重複
                        sql = @"SELECT TOP 1 1
                                FROM SCM.MaterialManagement
                                WHERE SupplierId = @SupplierId
                                AND MaterialName = @MaterialName";
                        dynamicParameters.Add("SupplierId", SupplierId);
                        dynamicParameters.Add("MaterialName", MaterialName);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【材料名稱+廠商】組合重複，請重新確認!");
                        #endregion

                        

                        #region //品號相關判斷
                        if (MtlItemId > 0)
                        {
                            #region //判斷品號是否存在
                            sql = @"SELECT TOP 1 1
                                    FROM PDM.MtlItem
                                    WHERE MtlItemId = @MtlItemId";
                            dynamicParameters.Add("MtlItemId", MtlItemId);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【品號】找不到，請重新確認!");
                            #endregion

                            #region //判斷品號是否有被綁定
                            sql = @"SELECT TOP 1 1
                                    FROM SCM.MaterialManagement
                                    WHERE MtlItemId = @MtlItemId";
                            dynamicParameters.Add("MtlItemId", MtlItemId);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() > 0) throw new SystemException("【品號】已經被綁定了，請重新確認!");
                            #endregion
                        }
                        #endregion


                        #region //資料新增
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.MaterialManagement (CompanyId,MaterialName, SupplierId, MtlItemId, Remark, Status
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.MtId
                                VALUES (@CompanyId,@MaterialName, @SupplierId, @MtlItemId, @Remark, @Status
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                CompanyId = CurrentCompany,
                                MaterialName,
                                SupplierId,
                                MtlItemId = MtlItemId != -1 ? MtlItemId : nullData,
                                Remark,
                                Status = "A", //啟用
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

        #region //AddMtAmount -- 新增材料金額資料 -- Shinotokuro 2024.08.19
        public string AddMtAmount(int MtId, float Amount, float UnitQuantity)
        {
            try
            {
                if (MtId <= 0) throw new SystemException("【材料】不能為空!");
                if (Amount < -1) throw new SystemException("【單位金額】不能為負!");
                if (UnitQuantity < -1) throw new SystemException("【單位量】不能為負!");
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷材料是否存在
                        sql = @"SELECT TOP 1 1
                                FROM SCM.MaterialManagement
                                WHERE MtId = @MtId";
                        dynamicParameters.Add("MtId", MtId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【材料】找不到，請重新確認!");
                        #endregion

                        #region //資料新增
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.MtAmount (MtId, Amount, UnitQuantity,
                                CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.MaId
                                VALUES (@MtId, @Amount, @UnitQuantity,
                                @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                MtId,
                                Amount,
                                UnitQuantity,
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

        #region//AddPurchaseOrder --採購單新增 -- Ding 2023.06.05
        public string AddPurchaseOrder(string CompanyNo,string UserNo)
        {
            try
            {
                if(CompanyNo.Length <= 0) throw new SystemException("【公司】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope(TransactionScopeOption.Required, new TimeSpan(0, 180, 0)))
                {
                    int CompanyId = -1;
                    string ErpConnectionStrings = "", ErpNo="";
                    List<string> purt = new List<string>(); //採購單號List

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷公司資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 CompanyId, CompanyName, ErpDb,ErpNo
                                FROM BAS.Company
                                WHERE CompanyNo = @CompanyNo";
                        dynamicParameters.Add("CompanyNo", CompanyNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() != 1) throw new SystemException("【公司】資料錯誤!");

                        foreach (var item in result)
                        {
                            CompanyId = Convert.ToInt32(item.CompanyId);
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            ErpNo = item.ErpNo;
                        }
                        #endregion

                        #region//判斷使用者是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserNo
                                    FROM BAS.[User] a
                                    WHERE a.UserNo = @UserNo";
                        dynamicParameters.Add("UserNo", UserNo);
                        var UserResult = sqlConnection.Query(sql, dynamicParameters);
                        if (UserResult.Count() != 1) throw new SystemException("採購單【使用者】資料錯誤!");
                        foreach (var item in UserResult)
                        {
                            UserNo = item.UserNo;
                        }
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        #region //判斷ERP使用者是否存在
                        DynamicParameters dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 MF004, MF005
                                FROM ADMMF
                                WHERE COMPANY = @CompanyNo
                                AND MF001 = @userNo";
                        dynamicParameters.Add("CompanyNo", ErpNo);
                        dynamicParameters.Add("userNo", UserNo);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (!result.Any()) throw new SystemException("【ERP使用者】不存在!");
                        #endregion

                        #region//判斷ERP使用者是否具備開立採購單權限
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 MG006
                                    FROM ADMMG
                                    WHERE COMPANY = @CompanyNo
                                    AND MG001 = @UserNo
                                    AND MG002 = @Function
                                    AND MG004 = 'Y'
                                    AND MG006 LIKE @Auth";
                        dynamicParameters.Add("CompanyNo", ErpNo);
                        dynamicParameters.Add("UserNo", UserNo);
                        dynamicParameters.Add("Function", "PURI07");
                        dynamicParameters.Add("Auth", "YYNYYYYYYYYY"); //修改/查詢/新增
                        var resultUserAuthExist = sqlConnection.Query(sql, dynamicParameters);
                        if (!resultUserAuthExist.Any()) throw new SystemException("月度採購拋轉-【ERP使用者】無權限!");
                        #endregion

                        #region//取得供應商清單
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT DISTINCT x.Supplier
                                FROM (
                                    SELECT T1.Item, T1.ItemName, T1.ItemSpec, T1.Supplier, T1.Price, T1.EffectiveDate, T1.ExpiryDate, T1.Currency,
                                        T1.MA026, T1.MA025
                                    FROM (
                                        SELECT TM.TM004 AS Item, TM.TM005 AS ItemName, TM.TM006 AS ItemSpec, TL.TL004 AS Supplier, TL.TL005 AS Currency,
                                            TM.TM010 AS Price, TM.TM014 AS EffectiveDate, TM.TM015 AS ExpiryDate,
                                            ROW_NUMBER() OVER (PARTITION BY TM.TM004 ORDER BY TM.TM014 DESC) AS RowNum,
                                            MA.MA026, MA.MA025
                                        FROM PURTL TL
                                            LEFT JOIN PURTM TM ON TL.TL001 = TM.TM001 AND TL.TL002 = TM.TM002
                                            LEFT JOIN INVMB MB ON MB.MB001 = TM.TM004
                                            LEFT JOIN CMSMG MG ON MG.MG001=TL.TL005
                                            LEFT JOIN PURMA MA ON MA.MA001=TL.TL004
                                            LEFT JOIN CMSNN NN ON NN.NN001=MA.MA076
                                        WHERE TL.TL006 = 'Y'
                                            AND TM.TM014 <= GETDATE()
                                            AND (TM.TM015 >= GETDATE() OR TM.TM015 = '')
                                            AND TM.TM004 IN (
                                            SELECT MB001
                                            FROM INVMB
                                            WHERE MB008 = '491'
                                                AND (MB030 <= GETDATE() OR MB030 = '')
                                                AND (MB031 >= GETDATE() OR MB031 = '')
                                        )
                                    ) AS T1
                                    WHERE T1.RowNum = 1
                                ) x
                                    INNER JOIN (
                                    SELECT a.TD004, (a.Avg- SUM(MC.MC007)) AS FinCount
                                    FROM (
                                        SELECT TD.TD004, CAST(ROUND(SUM(TD.TD008)/6, 0) AS DECIMAL(10,3)) AS Avg
                                        FROM dbo.PURTC TC
                                            LEFT JOIN dbo.PURTD TD ON TC.TC001 = TD.TD001 AND TC.TC002 = TD.TD002
                                        WHERE TD.TD004 IN (
                                            SELECT MB001
                                            FROM INVMB
                                            WHERE MB008 = '491'
                                                AND (MB030 <= GETDATE() OR MB030 = '')
                                                AND (MB031 >= GETDATE() OR MB031 = '')
                                        )
                                            AND TC.TC024 >= DATEADD(MONTH, -6, GETDATE())
                                            AND TC.TC014 = 'Y'
                                        GROUP BY TD.TD004
                                    ) a
                                        INNER JOIN dbo.INVMC MC ON a.TD004 = MC.MC001                                        
                                    GROUP BY a.TD004,a.Avg
                                    HAVING (a.Avg - SUM(MC.MC007)) > 0    
                                ) y ON x.Item = y.TD004
                                WHERE  1=1 
                                AND FinCount IS NOT NULL
                                GROUP BY x.Supplier,x.Item,x.Price";
                        var SupplierResult = sqlConnection.Query(sql, dynamicParameters);
                        if (SupplierResult.Count() <= 0) throw new SystemException("【供應商】資料錯誤!");
                        foreach (var item in SupplierResult)
                        {
                            string Supplier = item.Supplier;
                            DateTime now = DateTime.Now;

                            #region //審核ERP 採購單單頭 權限
                            string UsrGroupId = BaseHelper.CheckErpAuthority(UserNo, sqlConnection, "PURI07", "CREATE");
                            #endregion

                            #region//單據設定
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.MQ004, a.MQ005, a.MQ006, a.MQ044
                                FROM CMSMQ a
                                WHERE  a.MQ001 = '3309'";
                            dynamicParameters.Add("CompanyNo", ErpNo);

                            var resultDocSetting = sqlConnection.Query(sql, dynamicParameters);
                            if (resultDocSetting.Count() <= 0) throw new SystemException("ERP單據設定資料不存在!");
                            string encode = "";
                            int currentNum = 0, yearLength = 0, lineLength = 0;
                            foreach (var b in resultDocSetting)
                            {
                                encode = b.MQ004; //編碼方式
                                yearLength = Convert.ToInt32(b.MQ005); //年碼數
                                lineLength = Convert.ToInt32(b.MQ006); //流水號碼數
                            }
                            #endregion

                            #region //單號取號
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT CONVERT(int, RIGHT(ISNULL(MAX(RTRIM(LTRIM(TC002))), '" + new string('0', lineLength) + @"'), " + lineLength + @")) + 1 CurrentNum
                                FROM PURTC
                                WHERE TC001 = '3309' ";

                            #region //編碼方式
                            string dateFormat = "";
                            string PoNo = "";
                            
                            switch (encode)
                            {
                                case "1": //日編
                                    dateFormat = new string('y', yearLength) + "MMdd";
                                    sql += @" AND TC002 LIKE @ReferenceTime + '" + new string('_', lineLength) + @"'";
                                    dynamicParameters.Add("ReferenceTime", now.ToString(dateFormat));
                                    PoNo = now.ToString(dateFormat);
                                    break;
                                case "2": //月編
                                    dateFormat = new string('y', yearLength) + "MM";
                                    sql += @" AND TC002 LIKE @ReferenceTime + '" + new string('_', lineLength) + @"'";
                                    dynamicParameters.Add("ReferenceTime", now.ToString(dateFormat));
                                    PoNo = now.ToString(dateFormat);
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
                            PoNo += string.Format("{0:" + new string('0', lineLength) + "}", currentNum);
                            #endregion

                            #region//取得幣別
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 T1.Currency,T1.Supplier
                                    FROM (
                                        SELECT TM.TM004 AS Item, TM.TM005 AS ItemName, TM.TM006 AS ItemSpec, TL.TL004 AS Supplier,TL.TL005 AS Currency,
                                            TM.TM010 AS Price, TM.TM014 AS EffectiveDate, TM.TM015 AS ExpiryDate,
                                            ROW_NUMBER() OVER (PARTITION BY TM.TM004 ORDER BY TM.TM014 DESC) AS RowNum,
                                            MA.MA026,MA.MA025
                                        FROM PURTL TL
                                        LEFT JOIN PURTM TM ON TL.TL001 = TM.TM001 AND TL.TL002 = TM.TM002
                                        LEFT JOIN INVMB MB ON MB.MB001 = TM.TM004
                                        LEFT JOIN CMSMG MG ON MG.MG001=TL.TL005
                                        LEFT JOIN PURMA MA ON MA.MA001=TL.TL004
                                        LEFT JOIN CMSNN NN ON NN.NN001=MA.MA076
                                        WHERE TL.TL006 = 'Y'
                                            AND TM.TM014 <= GETDATE()
                                            AND (TM.TM015 >= GETDATE() OR TM.TM015 = '')
                                            AND TM.TM004 IN (
                                                SELECT MB001
                                                FROM INVMB
                                                WHERE MB008 = '491'
                                                AND (MB030 <= GETDATE() OR MB030 = '')
                                                AND (MB031 >= GETDATE() OR MB031 = '')
                                            )
                                    ) AS T1
                                    WHERE T1.RowNum = 1
                                    AND T1.Supplier = @Supplier";
                            dynamicParameters.Add("Supplier", Supplier);
                            var CurrencyResult = sqlConnection.Query(sql, dynamicParameters);
                            if (CurrencyResult.Count() != 1) throw new SystemException("【幣別】資料錯誤!");
                            string Currency = "";
                            foreach (var k in CurrencyResult)
                            {
                                Currency = k.Currency;
                            }
                            #endregion

                            #region//取得幣別條件
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT MF003,MF004,MF005,MF006 
                                    FROM CMSMF
                                    WHERE MF001=@Currency";
                            dynamicParameters.Add("Currency", Currency);
                            var CMSMFResult = sqlConnection.Query(sql, dynamicParameters);
                            if (CMSMFResult.Count() != 1) throw new SystemException("【幣別條件】資料錯誤!");
                            int UnitPriceDP = 0;//單價取位
                            int AmountDP = 0;//金額取位
                            int UnitCostDP = 0;//單位成本取位
                            int CostAmountDP = 0;//成本金額取位

                            foreach (var a in CMSMFResult)
                            {
                                UnitPriceDP = Convert.ToInt32(a.MF003);
                                AmountDP = Convert.ToInt32(a.MF004);
                                UnitCostDP = Convert.ToInt32(a.MF005);
                                CostAmountDP = Convert.ToInt32(a.MF006);
                            }
                            #endregion

                            #region//取得匯率
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT MG004 
                                    FROM CMSMG
                                    WHERE MG001=@Currency
                                    AND MG002<=@MG002";
                            dynamicParameters.Add("Currency", Currency);
                            dynamicParameters.Add("MG002", string.Format("{0:0000}{1:00}{2:00}", now.Year, now.Month, now.Day));
                            var CMSMGResult = sqlConnection.Query(sql, dynamicParameters);
                            if (CMSMGResult.Count() != 1) throw new SystemException("【匯率】資料錯誤!");
                            Double ExchangeRate = 0;
                            foreach (var a in CMSMGResult)
                            {
                                ExchangeRate = Convert.ToDouble(a.MG004);
                            }
                            #endregion

                            #region//取得價格條件
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT MA026 
                                    FROM PURMA
                                    WHERE MA001=@Supplier";
                            dynamicParameters.Add("Supplier", Supplier);
                            var PURMAResult = sqlConnection.Query(sql, dynamicParameters);
                            if (PURMAResult.Count() != 1) throw new SystemException("【價格條件】資料錯誤!");
                            string PriceConditions = "";
                            foreach (var a in PURMAResult)
                            {
                                PriceConditions = a.MA026;
                            }
                            #endregion

                            #region//取得付款條件
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT MA025 
                                    FROM PURMA
                                    WHERE MA001=@Supplier";
                            dynamicParameters.Add("Supplier", Supplier);
                            PURMAResult = sqlConnection.Query(sql, dynamicParameters);
                            if (PURMAResult.Count() != 1) throw new SystemException("【付款條件】資料錯誤!");
                            string PayConditions = "";
                            foreach (var b in PURMAResult)
                            {
                                PayConditions = b.MA025;
                            }
                            #endregion

                            #region //查詢廠別
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MB001)) MB001 FROM CMSMB";
                            var CMSMBResult = sqlConnection.Query(sql, dynamicParameters);
                            if (CMSMBResult.Count() > 1) throw new SystemException("廠別數有多個，請與資訊人員確認!!");
                            string FactoryNo = "";
                            foreach (var j in CMSMBResult)
                            {
                                FactoryNo = j.MB001;
                            }
                            #endregion

                            #region//取得稅別碼
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT MA076 
                                    FROM PURMA
                                    WHERE MA001=@Supplier";
                            dynamicParameters.Add("Supplier", Supplier);
                            PURMAResult = sqlConnection.Query(sql, dynamicParameters);
                            if (PURMAResult.Count() != 1) throw new SystemException("【稅別碼】資料錯誤!");
                            string TaxCode = "";
                            foreach (var b in PURMAResult)
                            {
                                TaxCode = b.MA076;
                            }
                            #endregion

                            #region//課稅別
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT NN006 
                                    FROM CMSNN
                                    WHERE NN001=@TaxCode";
                            dynamicParameters.Add("TaxCode", TaxCode);
                            var CMSNNResult = sqlConnection.Query(sql, dynamicParameters);
                            if (CMSNNResult.Count() != 1) throw new SystemException("【課稅別】資料錯誤!");
                            string TaxType = "";
                            foreach (var b in CMSNNResult)
                            {
                                TaxType = b.NN006;
                            }
                            #endregion

                            #region//總計金額
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT ROUND(SUM(x.Price * ROUND(y.FinCount,0)),@AmountDP) FinPrice
                                    FROM (
                                        SELECT T1.Item, T1.ItemName, T1.ItemSpec, T1.Supplier, T1.Price, T1.EffectiveDate, T1.ExpiryDate, T1.Currency,
                                            T1.MA026, T1.MA025
                                        FROM (
                                            SELECT TM.TM004 AS Item, TM.TM005 AS ItemName, TM.TM006 AS ItemSpec, TL.TL004 AS Supplier, TL.TL005 AS Currency,
                                                TM.TM010 AS Price, TM.TM014 AS EffectiveDate, TM.TM015 AS ExpiryDate,
                                                ROW_NUMBER() OVER (PARTITION BY TM.TM004 ORDER BY TM.TM014 DESC) AS RowNum,
                                                MA.MA026, MA.MA025
                                            FROM PURTL TL
                                                LEFT JOIN PURTM TM ON TL.TL001 = TM.TM001 AND TL.TL002 = TM.TM002
                                                LEFT JOIN INVMB MB ON MB.MB001 = TM.TM004
                                                LEFT JOIN CMSMG MG ON MG.MG001=TL.TL005
                                                LEFT JOIN PURMA MA ON MA.MA001=TL.TL004
                                                LEFT JOIN CMSNN NN ON NN.NN001=MA.MA076
                                            WHERE TL.TL006 = 'Y'
                                                AND TM.TM014 <= GETDATE()
                                                AND (TM.TM015 >= GETDATE() OR TM.TM015 = '')
                                                AND TM.TM004 IN (
                                                SELECT MB001
                                                FROM INVMB
                                                WHERE MB008 = '491'
                                                    AND (MB030 <= GETDATE() OR MB030 = '')
                                                    AND (MB031 >= GETDATE() OR MB031 = '')
                                            )
                                        ) AS T1
                                        WHERE T1.RowNum = 1
                                    ) x
                                        LEFT JOIN (
                                        SELECT a.TD004, (a.Avg- SUM(MC.MC007)) AS FinCount
                                        FROM (
                                            SELECT TD.TD004, (SUM(TD.TD008)/6)  AS Avg
                                            FROM dbo.PURTC TC
                                                LEFT JOIN dbo.PURTD TD ON TC.TC001 = TD.TD001 AND TC.TC002 = TD.TD002
                                            WHERE TD.TD004 IN (
                                                SELECT MB001
                                                FROM INVMB
                                                WHERE MB008 = '491'
                                                    AND (MB030 <= GETDATE() OR MB030 = '')
                                                    AND (MB031 >= GETDATE() OR MB031 = '')
                                            )
                                                AND TC.TC024 >= DATEADD(MONTH, -6, GETDATE())
                                                AND TC.TC014 = 'Y'
                                            GROUP BY TD.TD004
                                        ) a
                                            INNER JOIN dbo.INVMC MC ON a.TD004 = MC.MC001
                                        WHERE  1=1
                                        GROUP BY a.TD004,a.Avg
                                        HAVING (a.Avg - SUM(MC.MC007)) > 0    
                                    ) y ON x.Item = y.TD004
                                    WHERE x.Supplier=@Supplier";
                            dynamicParameters.Add("Supplier", Supplier);
                            dynamicParameters.Add("AmountDP", AmountDP);
                            var FinPriceResult = sqlConnection.Query(sql, dynamicParameters);

                            Double FinPrice=0;
                            foreach (var b in FinPriceResult)
                            {
                                FinPrice = Convert.ToDouble(b.FinPrice);
                            }
                            #endregion

                            #region//營業稅率
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT NN004 
                                    FROM CMSNN
                                    WHERE NN001=@TaxCode";
                            dynamicParameters.Add("TaxCode", TaxCode);
                            CMSNNResult = sqlConnection.Query(sql, dynamicParameters);
                            if (CMSNNResult.Count() != 1) throw new SystemException("【營業稅率】資料錯誤!");
                            Double BusinessTaxRate =0;
                            foreach (var b in CMSNNResult)
                            {
                                BusinessTaxRate = Convert.ToDouble(b.NN004);
                            }
                            #endregion

                            #region//採購金額
                            /*
                              由單頭課稅別換算而得
                                (1).若課稅別為「應稅外加」，則稅額=金額×稅率。
                                (2).若課稅別為「應稅內含」，則稅額=金額-金額/(1+稅率)。
                             */                            
                            Double PurchaseAmount = 0;
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT MA053 
                                    FROM PURMA
                                    WHERE MA001=@Supplier";
                            dynamicParameters.Add("Supplier", Supplier);
                            PURMAResult = sqlConnection.Query(sql, dynamicParameters);
                            if (PURMAResult.Count() != 1) throw new SystemException("【客戶稅額設定】資料錯誤!");
                            string CustTaxSumConfiguration = "";
                            foreach (var b in PURMAResult)
                            {
                                CustTaxSumConfiguration = b.MA053;
                            }

                            if (CustTaxSumConfiguration == "1")
                            {

                                if (TaxType == "1")
                                {                                   
                                    PurchaseAmount = Math.Round((Convert.ToDouble(FinPrice) / (1 + Convert.ToDouble(BusinessTaxRate))), AmountDP);                                   
                                }
                                else if (TaxType == "2")
                                {
                                    PurchaseAmount = Math.Round((Convert.ToDouble(FinPrice)+ (Convert.ToDouble(FinPrice) * Convert.ToDouble(BusinessTaxRate))), AmountDP);
                                }
                                else
                                {
                                    //throw new SystemException("【採購金額】不可空值!");
                                    PurchaseAmount = 0;
                                }
                            }
                            else if (CustTaxSumConfiguration == "2")
                            {
                                //(單身單筆資料計算) --> 稅額  = (單身每筆金額 * 稅率) 再加總
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT ROUND((x.Price* ROUND(y.FinCount,0)),@AmountDP) LineFinCount
                                    FROM (
                                        SELECT T1.Item, T1.ItemName, T1.ItemSpec, T1.Supplier, T1.Price, T1.EffectiveDate, T1.ExpiryDate, T1.Currency,
                                            T1.MA026, T1.MA025
                                        FROM (
                                            SELECT TM.TM004 AS Item, TM.TM005 AS ItemName, TM.TM006 AS ItemSpec, TL.TL004 AS Supplier, TL.TL005 AS Currency,
                                                TM.TM010 AS Price, TM.TM014 AS EffectiveDate, TM.TM015 AS ExpiryDate,
                                                ROW_NUMBER() OVER (PARTITION BY TM.TM004 ORDER BY TM.TM014 DESC) AS RowNum,
                                                MA.MA026, MA.MA025
                                            FROM PURTL TL
                                                LEFT JOIN PURTM TM ON TL.TL001 = TM.TM001 AND TL.TL002 = TM.TM002
                                                LEFT JOIN INVMB MB ON MB.MB001 = TM.TM004
                                                LEFT JOIN CMSMG MG ON MG.MG001=TL.TL005
                                                LEFT JOIN PURMA MA ON MA.MA001=TL.TL004
                                                LEFT JOIN CMSNN NN ON NN.NN001=MA.MA076
                                            WHERE TL.TL006 = 'Y'
                                                AND TM.TM014 <= GETDATE()
                                                AND (TM.TM015 >= GETDATE() OR TM.TM015 = '')
                                                AND TM.TM004 IN (
                                                SELECT MB001
                                                FROM INVMB
                                                WHERE MB008 = '491'
                                                    AND (MB030 <= GETDATE() OR MB030 = '')
                                                    AND (MB031 >= GETDATE() OR MB031 = '')
                                            )
                                        ) AS T1
                                        WHERE T1.RowNum = 1
                                    ) x
                                        LEFT JOIN (
                                        SELECT a.TD004, (a.Avg- SUM(MC.MC007)) AS FinCount
                                        FROM (
                                            SELECT TD.TD004, (SUM(TD.TD008)/6)  AS Avg
                                            FROM dbo.PURTC TC
                                                LEFT JOIN dbo.PURTD TD ON TC.TC001 = TD.TD001 AND TC.TC002 = TD.TD002
                                            WHERE TD.TD004 IN (
                                                SELECT MB001
                                                FROM INVMB
                                                WHERE MB008 = '491'
                                                    AND (MB030 <= GETDATE() OR MB030 = '')
                                                    AND (MB031 >= GETDATE() OR MB031 = '')
                                            )
                                                AND TC.TC024 >= DATEADD(MONTH, -6, GETDATE())
                                                AND TC.TC014 = 'Y'
                                            GROUP BY TD.TD004
                                        ) a
                                            INNER JOIN dbo.INVMC MC ON a.TD004 = MC.MC001
                                        WHERE  1=1
                                        GROUP BY a.TD004,a.Avg
                                        HAVING (a.Avg - SUM(MC.MC007)) > 0    
                                    ) y ON x.Item = y.TD004
                                    WHERE x.Supplier=@Supplier";
                                dynamicParameters.Add("Supplier", Supplier);
                                dynamicParameters.Add("AmountDP", AmountDP);
                                var LineFinCountResult = sqlConnection.Query(sql, dynamicParameters);                               
                                foreach (var b in LineFinCountResult)
                                {
                                    if (TaxType == "1")
                                    {
                                        Double a = Math.Round((Convert.ToDouble(b.LineFinCount) / (1 + Convert.ToDouble(BusinessTaxRate))), AmountDP);
                                        PurchaseAmount = PurchaseAmount + a;
                                    }
                                    else if (TaxType == "2")
                                    {
                                        Double a = Math.Round((Convert.ToDouble(b.LineFinCount) + (Convert.ToDouble(b.LineFinCount) * Convert.ToDouble(BusinessTaxRate))), AmountDP);
                                        PurchaseAmount = PurchaseAmount+a;
                                    }
                                }
                            }
                            else {
                                throw new SystemException("【客戶稅額設定】資料設定錯誤!");
                            }

                            #endregion

                            #region//稅額
                            Double TaxSum = 0;
                            TaxSum = Convert.ToDouble(FinPrice) - PurchaseAmount;                                                    
                            #endregion

                            #region//送貨地址(一)
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT MB005 
                                    FROM CMSMB WHERE MB001=@MB001"; 
                            dynamicParameters.Add("MB001", FactoryNo);
                            CMSMBResult = sqlConnection.Query(sql, dynamicParameters);
                            if (CMSMBResult.Count() != 1) throw new SystemException("【送貨地址】資料錯誤!");
                            string DeliveryAddress1 = "";
                            foreach (var b in CMSMBResult)
                            {
                                DeliveryAddress1 = b.MB005;
                            }
                            #endregion

                            #region//聯絡人
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT MA013 
                                    FROM PURMA
                                    WHERE MA001=@Supplier";
                            dynamicParameters.Add("Supplier", Supplier);
                            PURMAResult = sqlConnection.Query(sql, dynamicParameters);
                            if (PURMAResult.Count() != 1) throw new SystemException("【聯絡人】資料錯誤!");
                            string ContactPerson = "";
                            foreach (var b in PURMAResult)
                            {
                                ContactPerson = b.MA013;
                            }
                            #endregion

                            #region//數量合計
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT x.Supplier,SUM(y.FinCount)FinCount
                                    FROM (
                                        SELECT T1.Item, T1.ItemName, T1.ItemSpec, T1.Supplier, T1.Price, T1.EffectiveDate, T1.ExpiryDate, T1.Currency,
                                            T1.MA026, T1.MA025
                                        FROM (
                                            SELECT TM.TM004 AS Item, TM.TM005 AS ItemName, TM.TM006 AS ItemSpec, TL.TL004 AS Supplier, TL.TL005 AS Currency,
                                                TM.TM010 AS Price, TM.TM014 AS EffectiveDate, TM.TM015 AS ExpiryDate,
                                                ROW_NUMBER() OVER (PARTITION BY TM.TM004 ORDER BY TM.TM014 DESC) AS RowNum,
                                                MA.MA026, MA.MA025
                                            FROM PURTL TL
                                                LEFT JOIN PURTM TM ON TL.TL001 = TM.TM001 AND TL.TL002 = TM.TM002
                                                LEFT JOIN INVMB MB ON MB.MB001 = TM.TM004
                                                LEFT JOIN CMSMG MG ON MG.MG001=TL.TL005
                                                LEFT JOIN PURMA MA ON MA.MA001=TL.TL004
                                                LEFT JOIN CMSNN NN ON NN.NN001=MA.MA076
                                            WHERE TL.TL006 = 'Y'
                                                AND TM.TM014 <= GETDATE()
                                                AND (TM.TM015 >= GETDATE() OR TM.TM015 = '')
                                                AND TM.TM004 IN (
                                                SELECT MB001
                                                FROM INVMB
                                                WHERE MB008 = '491'
                                                    AND (MB030 <= GETDATE() OR MB030 = '')
                                                    AND (MB031 >= GETDATE() OR MB031 = '')
                                            )
                                        ) AS T1
                                        WHERE T1.RowNum = 1
                                    ) x
                                        LEFT JOIN (
                                        SELECT a.TD004, (a.Avg- SUM(MC.MC007)) AS FinCount
                                        FROM (
                                            SELECT TD.TD004, CAST(ROUND(SUM(TD.TD008)/6, 0) AS DECIMAL(10,3)) AS Avg
                                            FROM dbo.PURTC TC
                                                LEFT JOIN dbo.PURTD TD ON TC.TC001 = TD.TD001 AND TC.TC002 = TD.TD002
                                            WHERE TD.TD004 IN (
                                                SELECT MB001
                                                FROM INVMB
                                                WHERE MB008 = '491'
                                                    AND (MB030 <= GETDATE() OR MB030 = '')
                                                    AND (MB031 >= GETDATE() OR MB031 = '')
                                            )
                                                AND TC.TC024 >= DATEADD(MONTH, -6, GETDATE())
                                                AND TC.TC014 = 'Y'
                                            GROUP BY TD.TD004
                                        ) a
                                            INNER JOIN dbo.INVMC MC ON a.TD004 = MC.MC001
                                        WHERE  1=1
                                        GROUP BY a.TD004,a.Avg
                                        HAVING (a.Avg - SUM(MC.MC007)) > 0    
                                    ) y ON x.Item = y.TD004
                                    WHERE x.Supplier=@Supplier
                                    GROUP BY x.Supplier
                                    ORDER BY x.Supplier";
                            dynamicParameters.Add("Supplier", Supplier);
                            var FinCountResult = sqlConnection.Query(sql, dynamicParameters);

                            Double FinCount = 0;
                            foreach (var b in FinCountResult)
                            {
                                FinCount = Convert.ToDouble(b.FinCount);
                            }
                            #endregion

                            #region//交易條件
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT MA077 
                                    FROM PURMA
                                    WHERE MA001=@Supplier";
                            dynamicParameters.Add("Supplier", Supplier);
                            PURMAResult = sqlConnection.Query(sql, dynamicParameters);
                            if (PURMAResult.Count() != 1) throw new SystemException("【交易條件】資料錯誤!");
                            string TradingConditions = "";
                            foreach (var b in PURMAResult)
                            {
                                TradingConditions = b.MA077;
                            }
                            #endregion

                            #region//付款條件代號
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT MA055 
                                    FROM PURMA
                                    WHERE MA001=@Supplier";
                            dynamicParameters.Add("Supplier", Supplier);
                            PURMAResult = sqlConnection.Query(sql, dynamicParameters);
                            if (PURMAResult.Count() != 1) throw new SystemException("【付款條件代號】資料錯誤!");
                            string PaymentNo = "";
                            foreach (var b in PURMAResult)
                            {
                                PaymentNo = b.MA055;
                            }
                            #endregion

                            #region//訂金比率
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT MA058 
                                    FROM PURMA
                                    WHERE MA001=@Supplier";
                            dynamicParameters.Add("Supplier", Supplier);
                            PURMAResult = sqlConnection.Query(sql, dynamicParameters);
                            if (PURMAResult.Count() != 1) throw new SystemException("【訂金比率】資料錯誤!");
                            Double DepositRatio = 0.0;
                            foreach (var b in PURMAResult)
                            {
                                DepositRatio = Convert.ToDouble(b.MA055);
                            }
                            #endregion

                            #region//Get單頭所需資訊  PURTC                          
                            string COMPANY = ErpNo;
                            string CREATOR = UserNo;
                            string USR_GROUP = UsrGroupId;
                            string CREATE_DATE = string.Format("{0:0000}{1:00}{2:00}", now.Year, now.Month, now.Day);
                            string MODIFIER = "";
                            string MODI_DATE = "";
                            string FLAG = "1";
                            string CREATE_TIME = string.Format("{0:00}:{1:00}:{2:00}", now.Hour, now.Minute, now.Second);
                            string CREATE_AP = "";
                            string CREATE_PRID = "PURI07";
                            string MODI_TIME = "";
                            string MODI_AP = "";
                            string MODI_PRID = "PURI07";
                            string TC001 = "3309";
                            string TC002 = PoNo;
                            string TC003 = string.Format("{0:0000}{1:00}{2:00}", now.Year, now.Month, now.Day);
                            string TC004 = Supplier;
                            string TC005 = Currency;
                            Double TC006 = Math.Round(ExchangeRate,6);
                            string TC007 = PriceConditions;
                            string TC008 = PayConditions;
                            string TC009 = "";
                            string TC010 = FactoryNo;
                            string TC011 = UserNo;
                            string TC012 = "1";
                            int TC013 = 0;
                            string TC014 = "N";
                            string TC015 = "";
                            string TC016 = "";
                            string TC017 = "";
                            string TC018 = TaxType;
                            Decimal TC019 = Convert.ToDecimal(Math.Round(PurchaseAmount, AmountDP));
                            Decimal TC020 = Convert.ToDecimal(Math.Round(TaxSum, AmountDP));
                            string TC021 = DeliveryAddress1;
                            string TC022 = "";
                            Double TC023 = Math.Round(FinCount,3);
                            string TC024 = string.Format("{0:0000}{1:00}{2:00}", now.Year, now.Month, now.Day);
                            string TC025 = UserNo;
                            Double TC026 = Math.Round(BusinessTaxRate,2);
                            string TC027 = PaymentNo;
                            Double TC028 = Math.Round(DepositRatio,4);
                            Double TC029 = 0.000;
                            string TC030 = "N";
                            int TC031 = 0;
                            string TC032 = "";
                            string TC033 = "N";
                            string TC034 = "";
                            string TC035 = "N";
                            string TC036 = "";
                            string TC037 = "";
                            string TC038 = "N";
                            string TC039 = "0000";
                            string TC040 = "N";
                            string TC041 = "";
                            int TC042 = 0;
                            int TC043 = 0;
                            string TC044 = "";
                            string TC045 = "";
                            string TC046 = "";
                            string TC047 = TaxCode;
                            string TC048 = TradingConditions;
                            string TC049 = "";
                            string TC050 = "";
                            string TC051 = "N";
                            string TC052 = ContactPerson;
                            string TC500 = "";
                            string TC502 = "";
                            string TC503 = "";
                            string TC550 = "";
                            #endregion

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO PURTC(COMPANY,CREATOR,USR_GROUP,CREATE_DATE,MODIFIER,MODI_DATE,FLAG,CREATE_TIME,CREATE_AP,CREATE_PRID,MODI_TIME,MODI_AP,MODI_PRID
                                ,TC001,TC002,TC003,TC004,TC005,TC006,TC007,TC008,TC009,TC010,TC011,TC012
                                ,TC013,TC014,TC015,TC016,TC017,TC018,TC019,TC020,TC021,TC022,TC023,TC024,TC025
                                ,TC026,TC027,TC028,TC029,TC030,TC031,TC032,TC033,TC034,TC035,TC036,TC037,TC038
                                ,TC039,TC040,TC041,TC042,TC043,TC044,TC045,TC046,TC047,TC048,TC049,TC050,TC051
                                ,TC052,TC500,TC502,TC503,TC550)
                                VALUES (@COMPANY,@CREATOR,@USR_GROUP,@CREATE_DATE,@MODIFIER,@MODI_DATE,@FLAG,@CREATE_TIME,@CREATE_AP,@CREATE_PRID,@MODI_TIME,@MODI_AP,@MODI_PRID
                                ,@TC001,@TC002,@TC003,@TC004,@TC005,@TC006,@TC007,@TC008,@TC009,@TC010,@TC011,@TC012
                                ,@TC013,@TC014,@TC015,@TC016,@TC017,@TC018,@TC019,@TC020,@TC021,@TC022,@TC023,@TC024,@TC025
                                ,@TC026,@TC027,@TC028,@TC029,@TC030,@TC031,@TC032,@TC033,@TC034,@TC035,@TC036,@TC037,@TC038
                                ,@TC039,@TC040,@TC041,@TC042,@TC043,@TC044,@TC045,@TC046,@TC047,@TC048,@TC049,@TC050,@TC051
                                ,@TC052,@TC500,@TC502,@TC503,@TC550
                                )";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    COMPANY,
                                    CREATOR,
                                    USR_GROUP,
                                    CREATE_DATE,
                                    MODIFIER,
                                    MODI_DATE,
                                    FLAG,
                                    CREATE_TIME,
                                    CREATE_AP,
                                    CREATE_PRID,
                                    MODI_TIME,
                                    MODI_AP,
                                    MODI_PRID,
                                    TC001,
                                    TC002,
                                    TC003,
                                    TC004,
                                    TC005,
                                    TC006,
                                    TC007,
                                    TC008,
                                    TC009,
                                    TC010,
                                    TC011,
                                    TC012,
                                    TC013,
                                    TC014,
                                    TC015,
                                    TC016,
                                    TC017,
                                    TC018,
                                    TC019,
                                    TC020,
                                    TC021,
                                    TC022,
                                    TC023,
                                    TC024,
                                    TC025,
                                    TC026,
                                    TC027,
                                    TC028,
                                    TC029,
                                    TC030,
                                    TC031,
                                    TC032,
                                    TC033,
                                    TC034,
                                    TC035,
                                    TC036,
                                    TC037,
                                    TC038,
                                    TC039,
                                    TC040,
                                    TC041,
                                    TC042,
                                    TC043,
                                    TC044,
                                    TC045,
                                    TC046,
                                    TC047,
                                    TC048,
                                    TC049,
                                    TC050,
                                    TC051,
                                    TC052,
                                    TC500,
                                    TC502,
                                    TC503,
                                    TC550
                                });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            #region//取單身 PURTD
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT x.Supplier, x.Item,x.Price,SUM(y.FinCount) FinCount,(x.Price*SUM(y.FinCount)) FinPrice
                                    FROM (
                                        SELECT T1.Item, T1.ItemName, T1.ItemSpec, T1.Supplier, T1.Price, T1.EffectiveDate, T1.ExpiryDate, T1.Currency,
                                            T1.MA026, T1.MA025
                                        FROM (
                                            SELECT TM.TM004 AS Item, TM.TM005 AS ItemName, TM.TM006 AS ItemSpec, TL.TL004 AS Supplier, TL.TL005 AS Currency,
                                                TM.TM010 AS Price, TM.TM014 AS EffectiveDate, TM.TM015 AS ExpiryDate,
                                                ROW_NUMBER() OVER (PARTITION BY TM.TM004 ORDER BY TM.TM014 DESC) AS RowNum,
                                                MA.MA026, MA.MA025
                                            FROM PURTL TL
                                                LEFT JOIN PURTM TM ON TL.TL001 = TM.TM001 AND TL.TL002 = TM.TM002
                                                LEFT JOIN INVMB MB ON MB.MB001 = TM.TM004
                                                LEFT JOIN CMSMG MG ON MG.MG001=TL.TL005
                                                LEFT JOIN PURMA MA ON MA.MA001=TL.TL004
                                                LEFT JOIN CMSNN NN ON NN.NN001=MA.MA076
                                            WHERE TL.TL006 = 'Y'
                                                AND TM.TM014 <= GETDATE()
                                                AND (TM.TM015 >= GETDATE() OR TM.TM015 = '')
                                                AND TM.TM004 IN (
                                                SELECT MB001
                                                FROM INVMB
                                                WHERE MB008 = '491'
                                                    AND (MB030 <= GETDATE() OR MB030 = '')
                                                    AND (MB031 >= GETDATE() OR MB031 = '')
                                            )
                                        ) AS T1
                                        WHERE T1.RowNum = 1
                                    ) x
                                        INNER JOIN (
                                        SELECT a.TD004, (a.Avg- SUM(MC.MC007)) AS FinCount
                                        FROM (
                                            SELECT TD.TD004, CAST(ROUND(SUM(TD.TD008)/6, 0) AS DECIMAL(10,3)) AS Avg
                                            FROM dbo.PURTC TC
                                                LEFT JOIN dbo.PURTD TD ON TC.TC001 = TD.TD001 AND TC.TC002 = TD.TD002
                                            WHERE TD.TD004 IN (
                                                SELECT MB001
                                                FROM INVMB
                                                WHERE MB008 = '491'
                                                    AND (MB030 <= GETDATE() OR MB030 = '')
                                                    AND (MB031 >= GETDATE() OR MB031 = '')
                                            )
                                                AND TC.TC024 >= DATEADD(MONTH, -6, GETDATE())
                                                AND TC.TC014 = 'Y'
                                            GROUP BY TD.TD004
                                        ) a
                                            INNER JOIN dbo.INVMC MC ON a.TD004 = MC.MC001                                        
                                        GROUP BY a.TD004,a.Avg
                                        HAVING (a.Avg - SUM(MC.MC007)) > 0    
                                    ) y ON x.Item = y.TD004
                                    WHERE  1=1 
                                    AND FinCount IS NOT NULL
                                    AND x.Supplier=@Supplier
                                    GROUP BY x.Supplier,x.Item,x.Price";
                            dynamicParameters.Add("Supplier", Supplier);
                            var DetailResult = sqlConnection.Query(sql, dynamicParameters);
                            if (DetailResult.Count()>0) {
                                for (int i = 0; i < DetailResult.Count(); i++)
                                {
                                    #region//品號資訊
                                    var MtlNo = DetailResult.ElementAt(i).Item;
                                    string MtlName = "", MtlSpec = "", InventoryNo = "", Uom = "";
                                    Double MtlInventoryQty = 0.0;
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT MB002,MB003,MB004,MB017,MB064
                                        FROM INVMB
                                        WHERE MB001 = @MtlNo";
                                    dynamicParameters.Add("MtlNo", MtlNo);
                                    var MtlResult = sqlConnection.Query(sql, dynamicParameters);
                                    if (MtlResult.Count() != 1) throw new SystemException("【品號資訊】資料錯誤!");
                                    foreach (var a in MtlResult)
                                    {
                                        MtlName = a.MB002;
                                        MtlSpec = a.MB003;
                                        Uom = a.MB004;
                                        InventoryNo = a.MB017;
                                        MtlInventoryQty = Math.Round(Convert.ToDouble(a.MB064), 3);
                                    }
                                    #endregion

                                    #region //審核ERP 採購單單身 權限
                                    string UsrLineGroupId = BaseHelper.CheckErpAuthority(UserNo, sqlConnection, "PURI09", "CREATE");
                                    #endregion

                                    COMPANY = ErpNo;
                                    CREATOR = UserNo;
                                    USR_GROUP = UsrLineGroupId;
                                    CREATE_DATE = string.Format("{0:0000}{1:00}{2:00}", now.Year, now.Month, now.Day);
                                    MODIFIER = "";
                                    MODI_DATE = "";
                                    FLAG = "1";
                                    CREATE_TIME = string.Format("{0:00}:{1:00}:{2:00}", now.Hour, now.Minute, now.Second);
                                    CREATE_AP = "";
                                    CREATE_PRID = "PURI09";
                                    MODI_TIME = "";
                                    MODI_AP = "";
                                    MODI_PRID = "PURI09";
                                    string TD001 = TC001;
                                    string TD002 = TC002;
                                    string TD003 = (i + 1).ToString().PadLeft(4, '0');
                                    string TD004 = MtlNo;
                                    string TD005 = MtlName;
                                    string TD006 = MtlSpec;
                                    string TD007 = InventoryNo;
                                    Double TD008 = Math.Round(Convert.ToDouble(DetailResult.ElementAt(i).FinCount), 3);
                                    string TD009 = Uom;
                                    Double TD010 = Math.Round(Convert.ToDouble(DetailResult.ElementAt(i).Price), UnitPriceDP);
                                    Double TD011 = Math.Round(Convert.ToDouble(DetailResult.ElementAt(i).FinPrice), AmountDP);
                                    string TD012 = CREATE_DATE;
                                    string TD013 = "****";
                                    string TD014 = "";
                                    Double TD015 = 0;
                                    string TD016 = "N";
                                    string TD017 = "";
                                    string TD018 = "N";
                                    Double TD019 = Math.Round(MtlInventoryQty, 3);
                                    string TD020 = "";
                                    string TD021 = "***********";
                                    string TD022 = "";
                                    string TD023 = "****";
                                    string TD024 = "";
                                    string TD025 = "N";
                                    string TD026 = "****";
                                    string TD027 = "***********";
                                    string TD028 = "****";
                                    string TD029 = "";
                                    Double TD030 = 0;
                                    Double TD031 = 0;
                                    string TD032 = "";
                                    string TD033 = "";
                                    string TD034 = "9";
                                    string TD035 = "";
                                    string TD036 = "";
                                    string TD037 = "";
                                    string TD038 = "2";
                                    string TD039 = "";
                                    string TD040 = "";
                                    string TD041 = "****";
                                    string TD042 = "";
                                    string TD043 = "";
                                    string TD044 = "";
                                    string TD045 = string.Format("{0:0000}{1:00}{2:00}", now.Year, now.Month, now.Day);
                                    string TD046 = string.Format("{0:0000}{1:00}{2:00}", now.Year, now.Month, now.Day);
                                    string TD047 = "";
                                    Double TD048 = 0;
                                    Double TD049 = 0;
                                    string TD050 = "";
                                    string TD051 = "";
                                    string TD052 = "";
                                    string TD053 = "";
                                    string TD054 = "";
                                    string TD055 = "";
                                    string TD056 = "";
                                    Double TD057 = Math.Round(BusinessTaxRate,2);
                                    Double TD058 = Math.Round(TD008, 3);
                                    string TD059 = TD009;
                                    Double TD060 = 0;
                                    string TD061 = "";
                                    Double TD062 = 0;
                                    Double TD063 = 0;
                                    string TD500 = "";
                                    string TD501 = "";
                                    string TD502 = "";
                                    string TD503 = "";
                                    string TD550 = "";
                                    Double TD551 = 0;
                                    string TD552 = "";
                                    string TD553 = "";
                                    string TD554 = "";
                                    string TD555 = "";

                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO PURTD(COMPANY,CREATOR,USR_GROUP,CREATE_DATE,MODIFIER,MODI_DATE,FLAG,CREATE_TIME,CREATE_AP,CREATE_PRID,MODI_TIME,MODI_AP,MODI_PRID
                                            ,TD001,TD002,TD003,TD004,TD005,TD006,TD007,TD008,TD009,TD010,TD011,TD012
                                            ,TD013,TD014,TD015,TD016,TD017,TD018,TD019,TD020,TD021,TD022,TD023,TD024,TD025
                                            ,TD026,TD027,TD028,TD029,TD030,TD031,TD032,TD033,TD034,TD035,TD036,TD037,TD038
                                            ,TD039,TD040,TD041,TD042,TD043,TD044,TD045,TD046,TD047,TD048,TD049,TD050,TD051
                                            ,TD052,TD053,TD054,TD055,TD056,TD057,TD058,TD059,TD060,TD061,TD062,TD063,TD500,TD501,TD502,TD503
                                            ,TD550,TD551,TD552,TD553,TD554,TD555)
                                            VALUES (@COMPANY,@CREATOR,@USR_GROUP,@CREATE_DATE,@MODIFIER,@MODI_DATE,@FLAG,@CREATE_TIME,@CREATE_AP,@CREATE_PRID,@MODI_TIME,@MODI_AP,@MODI_PRID
                                            ,@TD001,@TD002,@TD003,@TD004,@TD005,@TD006,@TD007,@TD008,@TD009,@TD010,@TD011,@TD012
                                            ,@TD013,@TD014,@TD015,@TD016,@TD017,@TD018,@TD019,@TD020,@TD021,@TD022,@TD023,@TD024,@TD025
                                            ,@TD026,@TD027,@TD028,@TD029,@TD030,@TD031,@TD032,@TD033,@TD034,@TD035,@TD036,@TD037,@TD038
                                            ,@TD039,@TD040,@TD041,@TD042,@TD043,@TD044,@TD045,@TD046,@TD047,@TD048,@TD049,@TD050,@TD051
                                            ,@TD052,@TD053,@TD054,@TD055,@TD056,@TD057,@TD058,@TD059,@TD060,@TD061,@TD062,@TD063,@TD500,@TD501,@TD502,@TD503
                                            ,@TD550,@TD551,@TD552,@TD553,@TD554,@TD555)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            COMPANY,
                                            CREATOR,
                                            USR_GROUP,
                                            CREATE_DATE,
                                            MODIFIER,
                                            MODI_DATE,
                                            FLAG,
                                            CREATE_TIME,
                                            CREATE_AP,
                                            CREATE_PRID,
                                            MODI_TIME,
                                            MODI_AP,
                                            MODI_PRID,
                                            TD001,
                                            TD002,
                                            TD003,
                                            TD004,
                                            TD005,
                                            TD006,
                                            TD007,
                                            TD008,
                                            TD009,
                                            TD010,
                                            TD011,
                                            TD012,
                                            TD013,
                                            TD014,
                                            TD015,
                                            TD016,
                                            TD017,
                                            TD018,
                                            TD019,
                                            TD020,
                                            TD021,
                                            TD022,
                                            TD023,
                                            TD024,
                                            TD025,
                                            TD026,
                                            TD027,
                                            TD028,
                                            TD029,
                                            TD030,
                                            TD031,
                                            TD032,
                                            TD033,
                                            TD034,
                                            TD035,
                                            TD036,
                                            TD037,
                                            TD038,
                                            TD039,
                                            TD040,
                                            TD041,
                                            TD042,
                                            TD043,
                                            TD044,
                                            TD045,
                                            TD046,
                                            TD047,
                                            TD048,
                                            TD049,
                                            TD050,
                                            TD051,
                                            TD052,
                                            TD053,
                                            TD054,
                                            TD055,
                                            TD056,
                                            TD057,
                                            TD058,
                                            TD059,
                                            TD060,
                                            TD061,
                                            TD062,
                                            TD063,
                                            TD500,
                                            TD501,
                                            TD502,
                                            TD503,
                                            TD550,
                                            TD551,
                                            TD552,
                                            TD553,
                                            TD554,
                                            TD555
                                        });
                                    insertResult = sqlConnection.Query(sql, dynamicParameters);
                                }                               
                            }
                            #endregion

                            purt.Add(TC001+"-"+TC002);
                        }
                        #endregion                       
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //信件通知
                        #region //Mail資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MailFrom, a.MailTo, a.MailCc, a.MailBcc, a.MailSubject, a.MailContent
                            , b.Host, b.Port, b.SendMode, b.Account, b.Password
                            FROM BAS.Mail a
                            LEFT JOIN BAS.MailServer b ON b.ServerId = a.ServerId
                            WHERE a.MailId IN (
                                SELECT z.MailId
                                FROM BAS.MailSendSetting z
                                WHERE z.SettingSchema = @SettingSchema
                                AND z.SettingNo = @SettingNo
                            )";
                        dynamicParameters.Add("SettingSchema", "PurchaseOrderContentMailAdvice");
                        dynamicParameters.Add("SettingNo", "Y");

                        var resultMailTemplate = sqlConnection.Query(sql, dynamicParameters);
                        if (resultMailTemplate.Count() <= 0) throw new SystemException("Mail設定錯誤!");
                        #endregion

                        foreach (var item in resultMailTemplate)
                        {
                            string mailSubject = item.MailSubject,
                                mailContent = HttpUtility.UrlDecode(item.MailContent);

                            #region //Mail內容
                            string replaceHtml = @"<table border='0' cellpadding='0' cellspacing='0' width='100%'>
                                                 <tr>
                                                   <td bgcolor='#FFFFFF' style='padding: 0 0 20px 0;'>
                                                     <table align='left' border='1' cellpadding='0' cellspacing='0' width='1625'>
                                                       <tr>
                                                         <td align='center' bgcolor='#CCFFFF' style='font-family: 微軟正黑體, sans-serif;' width='100'>採購單編號</td>
                                                       </tr>";
                            foreach (var z in purt) {
                                replaceHtml += @"
                                                <tr><td align='center' style='font-family: 微軟正黑體, sans-serif;' width='100'>"+z+"</td></tr>";
                            }
                            replaceHtml += @"            </table>
                                                   </td>
                                                 </tr>
                                               </table>";

                            mailSubject = mailSubject.Replace("[Date]", DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd"));
                            mailContent = mailContent.Replace("[Date]", DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd"));
                            mailContent = mailContent.Replace("[PurchaseOrderContent]", replaceHtml);
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
                        msg = "(1 rows affected)"
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
        #endregion

        #region //Update
        #region //UpdateMaterial -- 更新材料資料 -- Shinotokuro 2024.08.19
        public string UpdateMaterial(int MtId, string MaterialName, int SupplierId, string Remark, int MtlItemId)
        {
            try
            {
                if (MtId <= 0) throw new SystemException("【材料】不能為空!");
                if (MaterialName.Length <= 0) throw new SystemException("【材料】不能為空!");
                if (MaterialName.Length > 100) throw new SystemException("【材料】長度錯誤!");
                if (SupplierId <= 0) throw new SystemException("【廠商】不能為空!");
                if (Remark.Length > 200) throw new SystemException("【備註】長度不能超過200!");
                int? nullData = null;
                string ConfirmStatus = "";
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷材料是否存在
                        sql = @"SELECT TOP 1 
                                CASE 
                                    WHEN x.ConfirmStatus = 'Y' OR x.ConfirmStatus = 'H'  THEN 'Y'
                                    ELSE  'N'
                                END ConfirmStatus
                                FROM SCM.MaterialManagement a
                                OUTER APPLY(
                                    SELECT TOP 1 x1.ConfirmStatus 
                                    FROM SCM.MtAmount x1
                                    WHERE a.MtId = x1.MtId
                                    AND x1.ConfirmStatus != 'N'
                                ) x
                                WHERE a.MtId = @MtId";
                        dynamicParameters.Add("MtId", MtId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【材料】找不到，請重新確認!");
                        foreach(var item in result)
                        {
                            ConfirmStatus = item.ConfirmStatus;
                            //if(item.Status == "A") throw new SystemException("【材料】已啟用不可以更改");
                        }
                        #endregion

                        #region //判斷廠商是否存在
                        sql = @"SELECT TOP 1 1
                                FROM SCM.Supplier
                                WHERE SupplierId = @SupplierId";
                        dynamicParameters.Add("SupplierId", SupplierId);
                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【廠商】找不到，請重新確認!");
                        #endregion

                        #region //品號相關判斷
                        if (MtlItemId > 0)
                        {
                            #region //判斷品號是否存在
                            sql = @"SELECT TOP 1 1
                                    FROM PDM.MtlItem
                                    WHERE MtlItemId = @MtlItemId";
                            dynamicParameters.Add("MtlItemId", MtlItemId);
                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("【品號】找不到，請重新確認!");
                            #endregion

                            #region //判斷品號是否有被綁定
                            sql = @"SELECT TOP 1 1
                                    FROM SCM.MaterialManagement
                                    WHERE MtlItemId = @MtlItemId
                                    AND MtId != @MtId";
                            dynamicParameters.Add("MtlItemId", MtlItemId);
                            dynamicParameters.Add("MtId", MtId);

                            result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() > 0) throw new SystemException("【品號】已經被綁定了，請重新確認!");
                            #endregion
                        }
                        #endregion

                        #region //判斷材料名稱是否重複
                        sql = @"SELECT TOP 1 1
                                FROM SCM.MaterialManagement
                                WHERE SupplierId = @SupplierId
                                AND MaterialName = @MaterialName
                                AND MtId != @MtId";
                        dynamicParameters.Add("SupplierId", SupplierId);
                        dynamicParameters.Add("MaterialName", MaterialName);
                        dynamicParameters.Add("MtId", MtId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【材料名稱+廠商】組合重複，請重新確認!");
                        #endregion


                        #region //資料更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.MaterialManagement SET " +
                                (ConfirmStatus == "N" ? 
                              @"MaterialName = @MaterialName,
                                SupplierId = @SupplierId," : "")
                                +
                              @"MtlItemId = @MtlItemId,
                                Remark = @Remark,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE MtId = @MtId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                MaterialName,
                                SupplierId,
                                MtlItemId = MtlItemId != -1 ? MtlItemId : nullData,
                                Remark,
                                LastModifiedDate,
                                LastModifiedBy,
                                MtId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //UpdateMaterialStatus -- 更新材料狀態 -- Shinotokuro 2024.08.19
        public string UpdateMaterialStatus(int MtId)
        {
            try
            {
                if (MtId <= 0) throw new SystemException("【材料】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();
                        string status = "";

                        #region //判斷材料資訊是否正確
                        sql = @"SELECT TOP 1 Status
                                FROM SCM.MaterialManagement
                                WHERE MtId = @MtId";
                        dynamicParameters.Add("MtId", MtId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【材料】找不到，請重新確認!");
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

                        #region //資料更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.MaterialManagement SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE MtId = @MtId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = status,
                                LastModifiedDate,
                                LastModifiedBy,
                                MtId
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

        #region //UpdateMtAmount -- 更新材料金額資料 -- Shinotokuro 2024.08.19
        public string UpdateMtAmount(int MaId, double Amount, double UnitQuantity)
        {
            try
            {
                if (MaId <= 0) throw new SystemException("【材料】不能為空!");
                if (Amount < -1) throw new SystemException("【單位金額】不能為負!");
                if (UnitQuantity < -1) throw new SystemException("【單位量】不能為負!");
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷材料金額是否存在
                        sql = @"SELECT TOP 1 ConfirmStatus
                                FROM SCM.MtAmount
                                WHERE MaId = @MaId";
                        dynamicParameters.Add("MaId", MaId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【材料金額】找不到，請重新確認!");
                        foreach (var item in result)
                        {
                            if (item.ConfirmStatus != "N") throw new SystemException("【材料金額】非未確認狀態不可以修改");
                        }
                        #endregion


                        #region //資料更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.MtAmount SET
                                Amount = @Amount,
                                UnitQuantity = @UnitQuantity,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE MaId = @MaId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Amount,
                                UnitQuantity,
                                LastModifiedDate,
                                LastModifiedBy,
                                MaId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //UpdateMtAmountConfirm -- 更新確認材料金額資料 -- Shinotokuro 2024.08.19
        public string UpdateMtAmountConfirm(int MaId)
        {
            try
            {
                if (MaId <= 0) throw new SystemException("【材料金額】不能為空!");
                int MtId = -1;
                int rowsAffected = 0;

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();
                        
                        #region //判斷材料金額資訊是否正確
                        sql = @"SELECT TOP 1 ConfirmStatus,MtId
                                FROM SCM.MtAmount
                                WHERE MaId = @MaId";
                        dynamicParameters.Add("MaId", MaId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【材料金額】找不到，請重新確認!");
                        foreach (var item in result)
                        {
                            if (item.ConfirmStatus == "Y") throw new SystemException("【材料】已確認不可以更改");
                            if (item.ConfirmStatus == "H") throw new SystemException("【材料】已成為歷史不可以更改");
                            MtId = item.MtId;
                        }
                        #endregion

                        #region //資料更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.MtAmount SET
                                ConfirmStatus = @ConfirmStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE MtId = @MtId
                                AND ConfirmStatus = 'Y'";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmStatus = 'H',
                                LastModifiedDate,
                                LastModifiedBy,
                                MtId
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //資料更新
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.MtAmount SET
                                ConfirmDate = @ConfirmDate,
                                ConfirmUserId = @ConfirmUserId,
                                ConfirmStatus = @ConfirmStatus,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE MaId = @MaId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ConfirmDate = LastModifiedDate,
                                ConfirmUserId = LastModifiedBy,
                                ConfirmStatus = 'Y',
                                LastModifiedDate,
                                LastModifiedBy,
                                MaId
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

        #region //UpdatePurchaseOrderSynchronize -- 採購單資料同步 -- Ann 2024-03-01
        public string UpdatePurchaseOrderSynchronize(string CompanyNo, string UpdateDate, string PoErpFullNo)
        {
            try
            {
                if (PoErpFullNo == null) PoErpFullNo = "";
                List<PurchaseOrder> purchaseOrders = new List<PurchaseOrder>();
                List<PoDetail> poDetails = new List<PoDetail>();

                using (TransactionScope transactionScope = new TransactionScope(TransactionScopeOption.Required, new TimeSpan(0, 180, 0)))
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
                        #region //撈取ERP採購單頭資料
                        sql = @"SELECT LTRIM(RTRIM(TC001)) PoErpPrefix, LTRIM(RTRIM(TC002)) PoErpNo
                                , CASE WHEN LEN(LTRIM(RTRIM(TC003))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TC003)) as date), 'yyyy-MM-dd') ELSE NULL END PoDate
                                , LTRIM(RTRIM(TC004)) SupplierNo, LTRIM(RTRIM(TC005)) CurrencyCode, TC006 Exchange, LTRIM(RTRIM(TC008)) PaymentTerm
                                , LTRIM(RTRIM(TC009)) Remark, LTRIM(RTRIM(TC011)) PoUserNo, LTRIM(RTRIM(TC014)) ConfirmStatus, LTRIM(RTRIM(TC018)) Taxation
                                , TC019 PoPrice, TC020 TaxAmount, LTRIM(RTRIM(TC021)) FirstAddress, LTRIM(RTRIM(TC022)) SecondAddress, TC023 Quantity
                                , CASE WHEN LEN(LTRIM(RTRIM(TC024))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TC024)) as date), 'yyyy-MM-dd') ELSE NULL END DocDate
                                , LTRIM(RTRIM(TC025)) ConfirmUserNo, TC026 TaxRate, LTRIM(RTRIM(TC027)) PaymentTermNo, LTRIM(RTRIM(TC030)) ApproveStatus
                                , LTRIM(RTRIM(TC039)) Edition, LTRIM(RTRIM(TC040)) DepositPartial, LTRIM(RTRIM(TC047)) TaxNo, LTRIM(RTRIM(TC048)) TradeTerm
                                , LTRIM(RTRIM(TC051)) DetailMultiTax, LTRIM(RTRIM(TC052)) ContactUser
                                , CASE WHEN LEN(LTRIM(RTRIM(CREATE_DATE))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_DATE)) as date), 'yyyy-MM-dd') ELSE NULL END TransferDate
                                , CASE WHEN LEN(LTRIM(RTRIM(CREATE_TIME))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_TIME)) as datetime), 'HH:mm:ss') ELSE NULL END TransferTime
                                FROM PURTC
                                WHERE 1=1";
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "UpdateDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME > @UpdateDate OR MODI_DATE + ' ' + MODI_TIME > @UpdateDate)", UpdateDate);
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PoErpFullNo", @" AND TC001 + '-' + TC002 LIKE '%' + @PoErpFullNo + '%'", PoErpFullNo);
                        sql += @" ORDER BY TransferDate, TransferTime";

                        purchaseOrders = sqlConnection.Query<PurchaseOrder>(sql, dynamicParameters).ToList();
                        #endregion

                        #region //撈取ERP採購單身資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT LTRIM(RTRIM(TD001)) PoErpPrefix, LTRIM(RTRIM(TD002)) PoErpNo, LTRIM(RTRIM(TD003)) PoSeq
                                , LTRIM(RTRIM(TD004)) MtlItemNo, LTRIM(RTRIM(TD005)) PoMtlItemName, LTRIM(RTRIM(TD006)) PoMtlItemSpec, LTRIM(RTRIM(TD007)) InventoryNo
                                , TD008 Quantity, LTRIM(RTRIM(TD009)) UomNo, TD010 PoUnitPrice, TD011 PoPrice
                                , CASE WHEN LEN(LTRIM(RTRIM(TD012))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TD012)) as date), 'yyyy-MM-dd') ELSE NULL END PromiseDate
                                , LTRIM(RTRIM(TD013)) ReferencePrefix, LTRIM(RTRIM(TD014)) Remark, TD015 SiQty, LTRIM(RTRIM(TD016)) ClosureStatus
                                , LTRIM(RTRIM(TD018)) ConfirmStatus, TD019 InventoryQty, LTRIM(RTRIM(TD020)) SmallUnit, LTRIM(RTRIM(TD021)) ReferenceNo
                                , LTRIM(RTRIM(TD022)) Project, LTRIM(RTRIM(TD023)) ReferenceSeq, LTRIM(RTRIM(TD025)) UrgentMtl
                                , LTRIM(RTRIM(TD026)) PrErpPrefix, LTRIM(RTRIM(TD027)) PrErpNo, LTRIM(RTRIM(TD028)) PrSequence, LTRIM(RTRIM(TD034)) FromDocType
                                , LTRIM(RTRIM(TD038)) MtlItemType, LTRIM(RTRIM(TD041)) PartialSeq
                                , CASE WHEN LEN(LTRIM(RTRIM(TD045))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TD045)) as date), 'yyyy-MM-dd') ELSE NULL END OriPromiseDate
                                , CASE WHEN LEN(LTRIM(RTRIM(TD046))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TD046)) as date), 'yyyy-MM-dd') ELSE NULL END DeliveryDate
                                , TD058 PoPriceQty, LTRIM(RTRIM(TD059)) PoPriceUomNo, TD060 SiPriceQty, ISNULL(TD062, 0) DiscountRate, ISNULL(TD063, 0) DiscountAmount
                                , CASE WHEN LEN(LTRIM(RTRIM(CREATE_DATE))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_DATE)) as date), 'yyyy-MM-dd') ELSE NULL END TransferDate
                                , CASE WHEN LEN(LTRIM(RTRIM(CREATE_TIME))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_TIME)) as datetime), 'HH:mm:ss') ELSE NULL END TransferTime
                                FROM PURTD
                                WHERE 1=1";
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "UpdateDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME > @UpdateDate OR MODI_DATE + ' ' + MODI_TIME > @UpdateDate)", UpdateDate);
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PoErpFullNo", @" AND TD001 + '-' + TD002 LIKE '%' + @PoErpFullNo + '%'", PoErpFullNo);
                        sql += @" ORDER BY TransferDate, TransferTime";

                        poDetails = sqlConnection.Query<PoDetail>(sql, dynamicParameters).ToList();
                        #endregion
                    }

                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //取得MES供應商資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT SupplierId, SupplierNo
                                FROM SCM.Supplier
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<Supplier> resultSuppliers = sqlConnection.Query<Supplier>(sql, dynamicParameters).ToList();

                        purchaseOrders = purchaseOrders.Join(resultSuppliers, x => x.SupplierNo, y => y.SupplierNo, (x, y) => { x.SupplierId = y.SupplierId; return x; }).ToList();
                        #endregion

                        #region //撈取單位ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT UomId, UomNo
                                FROM PDM.UnitOfMeasure
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<Uom> resultSoPriceUomrNos = sqlConnection.Query<Uom>(sql, dynamicParameters).ToList();

                        poDetails = poDetails.Join(resultSoPriceUomrNos, x => x.UomNo.ToUpper(), y => y.UomNo, (x, y) => { x.UomId = y.UomId; return x; }).ToList();
                        poDetails = poDetails.Join(resultSoPriceUomrNos, x => x.PoPriceUomNo.ToUpper(), y => y.UomNo, (x, y) => { x.PoPriceUomId = y.UomId; return x; }).ToList();
                        #endregion

                        #region //撈取品號ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT MtlItemId, MtlItemNo
                                FROM PDM.MtlItem
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<MtlItem> resultMtlitems = sqlConnection.Query<MtlItem>(sql, dynamicParameters).ToList();

                        poDetails = poDetails.Join(resultMtlitems, x => x.MtlItemNo, y => y.MtlItemNo, (x, y) => { x.MtlItemId = y.MtlItemId; return x; }).Select(x => x).ToList();
                        #endregion

                        #region //撈取庫別ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT InventoryId, InventoryNo
                                FROM SCM.Inventory
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<Inventory> resultInventories = sqlConnection.Query<Inventory>(sql, dynamicParameters).ToList();

                        poDetails = poDetails.Join(resultInventories, x => x.InventoryNo, y => y.InventoryNo, (x, y) => { x.InventoryId = y.InventoryId; return x; }).ToList();
                        #endregion

                        #region //撈取使用者ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserId , a.UserNo 
                                FROM BAS.[User] a";

                        List<User> resultConfirmUsers = sqlConnection.Query<User>(sql, dynamicParameters).ToList();

                        purchaseOrders = purchaseOrders.GroupJoin(resultConfirmUsers, x => x.ConfirmUserNo, y => y.UserNo, (x, y) => { x.ConfirmUserId = y.FirstOrDefault()?.UserId; return x; }).Select(x => x).ToList();
                        purchaseOrders = purchaseOrders.Join(resultConfirmUsers, x => x.PoUserNo, y => y.UserNo, (x, y) => { x.PoUserId = y.UserId; return x; }).ToList();
                        #endregion

                        #region //判斷SCM.PurchaseOrder是否有資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT PoId, PoErpPrefix, PoErpNo
                                FROM SCM.PurchaseOrder
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<PurchaseOrder> resultPurchaseOrder = sqlConnection.Query<PurchaseOrder>(sql, dynamicParameters).ToList();

                        purchaseOrders = purchaseOrders.GroupJoin(resultPurchaseOrder, x => new { x.PoErpPrefix, x.PoErpNo }, y => new { y.PoErpPrefix, y.PoErpNo }, (x, y) => { x.PoId = y.FirstOrDefault()?.PoId; return x; }).Select(x => x).ToList();
                        #endregion

                        #region //進貨單單頭(新增/修改)
                        List<PurchaseOrder> addPurchaseOrders = purchaseOrders.Where(x => x.PoId == null).ToList();
                        List<PurchaseOrder> updatePurchaseOrders = purchaseOrders.Where(x => x.PoId != null).ToList();

                        #region //新增
                        dynamicParameters = new DynamicParameters();
                        if (addPurchaseOrders.Count > 0)
                        {
                            addPurchaseOrders
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.CompanyId = CompanyId;
                                    x.TransferStatus = "Y";
                                    x.CreateDate = CreateDate;
                                    x.LastModifiedDate = LastModifiedDate;
                                    x.CreateBy = CreateBy;
                                    x.LastModifiedBy = LastModifiedBy;
                                });

                            sql = @"INSERT INTO SCM.PurchaseOrder (CompanyId, PoErpPrefix, PoErpNo, PoDate, SupplierId, CurrencyCode, Exchange, PaymentTerm, Remark
                                    , PoUserId, ConfirmStatus, Taxation, PoPrice, TaxAmount, FirstAddress, SecondAddress, Quantity, DocDate
                                    , ConfirmUserId, TaxRate, PaymentTermNo, ApproveStatus, Edition, DepositPartial, TaxNo, TradeTerm, DetailMultiTax
                                    , ContactUser, TransferStatus, TransferDate
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@CompanyId, @PoErpPrefix, @PoErpNo, @PoDate, @SupplierId, @CurrencyCode, @Exchange, @PaymentTerm, @Remark
                                    , @PoUserId, @ConfirmStatus, @Taxation, @PoPrice, @TaxAmount, @FirstAddress, @SecondAddress, @Quantity, @DocDate
                                    , @ConfirmUserId, @TaxRate, @PaymentTermNo, @ApproveStatus, @Edition, @DepositPartial, @TaxNo, @TradeTerm, @DetailMultiTax
                                    , @ContactUser, @TransferStatus, @TransferDate
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            rowsAffected += sqlConnection.Execute(sql, addPurchaseOrders);
                        }
                        #endregion

                        #region //修改
                        dynamicParameters = new DynamicParameters();
                        if (updatePurchaseOrders.Count > 0)
                        {
                            updatePurchaseOrders
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.LastModifiedDate = LastModifiedDate;
                                    x.LastModifiedBy = LastModifiedBy;
                                });

                            sql = @"UPDATE SCM.PurchaseOrder SET
                                    PoDate = @PoDate,
                                    SupplierId = @SupplierId,
                                    CurrencyCode = @CurrencyCode,
                                    Exchange = @Exchange,
                                    PaymentTerm = @PaymentTerm,
                                    Remark = @Remark,
                                    PoUserId = @PoUserId,
                                    ConfirmStatus = @ConfirmStatus,
                                    Taxation = @Taxation,
                                    PoPrice = @PoPrice,
                                    TaxAmount = @TaxAmount,
                                    FirstAddress = @FirstAddress,
                                    SecondAddress = @SecondAddress,
                                    Quantity = @Quantity,
                                    DocDate = @DocDate,
                                    ConfirmUserId = @ConfirmUserId,
                                    TaxRate = @TaxRate,
                                    PaymentTermNo = @PaymentTermNo,
                                    ApproveStatus = @ApproveStatus,
                                    Edition = @Edition,
                                    DepositPartial = @DepositPartial,
                                    TaxNo = @TaxNo,
                                    TradeTerm = @TradeTerm,
                                    DetailMultiTax = @DetailMultiTax,
                                    ContactUser = @ContactUser,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE PoId = @PoId";
                            rowsAffected += sqlConnection.Execute(sql, updatePurchaseOrders);
                        }
                        #endregion
                        #endregion

                        #region //進貨單單身(新增/修改)
                        #region //撈取進貨單單頭ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT PoId, PoErpPrefix, PoErpNo
                                FROM  SCM.PurchaseOrder
                                WHERE CompanyId = @CompanyId
                                ORDER BY PoId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        resultPurchaseOrder = sqlConnection.Query<PurchaseOrder>(sql, dynamicParameters).ToList();

                        poDetails = poDetails.Join(resultPurchaseOrder, x => new { x.PoErpPrefix, x.PoErpNo }, y => new { y.PoErpPrefix, y.PoErpNo }, (x, y) => { x.PoId = y.PoId; return x; }).Select(x => x).ToList();
                        #endregion

                        #region //判斷SCM.PoDetail是否有資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.PoDetailId, a.PoId, a.PoSeq
                                FROM SCM.PoDetail a
                                INNER JOIN SCM.PurchaseOrder b ON a.PoId = b.PoId
                                WHERE b.CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<PoDetail> resultPoDetail = sqlConnection.Query<PoDetail>(sql, dynamicParameters).ToList();

                        poDetails = poDetails.GroupJoin(resultPoDetail, x => new { x.PoId, x.PoSeq }, y => new { y.PoId, y.PoSeq }, (x, y) => { x.PoDetailId = y.FirstOrDefault()?.PoDetailId; return x; }).ToList();
                        #endregion

                        List<PoDetail> addPoDetail = poDetails.Where(x => x.PoDetailId == null).ToList();
                        List<PoDetail> updatePoDetail = poDetails.Where(x => x.PoDetailId != null).ToList();

                        #region //新增
                        dynamicParameters = new DynamicParameters();
                        if (addPoDetail.Count > 0)
                        {
                            addPoDetail
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.CreateDate = CreateDate;
                                    x.LastModifiedDate = LastModifiedDate;
                                    x.CreateBy = CreateBy;
                                    x.LastModifiedBy = LastModifiedBy;
                                });

                            sql = @"INSERT INTO SCM.PoDetail (PoId, PoErpPrefix, PoErpNo, PoSeq, MtlItemId, PoMtlItemName, PoMtlItemSpec, InventoryId, Quantity
                                    , UomId, PoUnitPrice, PoPrice, PromiseDate, ReferencePrefix, Remark, SiQty, ClosureStatus, ConfirmStatus
                                    , InventoryQty, SmallUnit, ReferenceNo, Project, ReferenceSeq, UrgentMtl, PrErpPrefix, PrErpNo, PrSequence
                                    , FromDocType, MtlItemType, PartialSeq, OriPromiseDate, DeliveryDate, PoPriceQty, PoPriceUomId, SiPriceQty
                                    , DiscountRate, DiscountAmount, CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@PoId, @PoErpPrefix, @PoErpNo, @PoSeq, @MtlItemId, @PoMtlItemName, @PoMtlItemSpec, @InventoryId, @Quantity
                                    , @UomId, @PoUnitPrice, @PoPrice, @PromiseDate, @ReferencePrefix, @Remark, @SiQty, @ClosureStatus, @ConfirmStatus
                                    , @InventoryQty, @SmallUnit, @ReferenceNo, @Project, @ReferenceSeq, @UrgentMtl, @PrErpPrefix, @PrErpNo, @PrSequence
                                    , @FromDocType, @MtlItemType, @PartialSeq, @OriPromiseDate, @DeliveryDate, @PoPriceQty, @PoPriceUomId, @SiPriceQty
                                    , @DiscountRate, @DiscountAmount, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            rowsAffected += sqlConnection.Execute(sql, addPoDetail);
                        }
                        #endregion

                        #region //修改
                        dynamicParameters = new DynamicParameters();
                        if (updatePoDetail.Count > 0)
                        {
                            updatePoDetail
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.LastModifiedDate = LastModifiedDate;
                                    x.LastModifiedBy = LastModifiedBy;
                                });

                            sql = @"UPDATE SCM.PoDetail SET
                                    MtlItemId = @MtlItemId,
                                    PoMtlItemName = @PoMtlItemName,
                                    PoMtlItemSpec = @PoMtlItemSpec,
                                    InventoryId = @InventoryId,
                                    Quantity = @Quantity,
                                    UomId = @UomId,
                                    PoUnitPrice = @PoUnitPrice,
                                    PoPrice = @PoPrice,
                                    PromiseDate = @PromiseDate,
                                    ReferencePrefix = @ReferencePrefix,
                                    Remark = @Remark,
                                    SiQty = @SiQty,
                                    ClosureStatus = @ClosureStatus,
                                    ConfirmStatus = @ConfirmStatus,
                                    InventoryQty = @InventoryQty,
                                    SmallUnit = @SmallUnit,
                                    ReferenceNo = @ReferenceNo,
                                    Project = @Project,
                                    ReferenceSeq = @ReferenceSeq,
                                    UrgentMtl = @UrgentMtl,
                                    PrErpPrefix = @PrErpPrefix,
                                    PrErpNo = @PrErpNo,
                                    PrSequence = @PrSequence,
                                    FromDocType = @FromDocType,
                                    MtlItemType = @MtlItemType,
                                    PartialSeq = @PartialSeq,
                                    OriPromiseDate = @OriPromiseDate,
                                    DeliveryDate = @DeliveryDate,
                                    PoPriceQty = @PoPriceQty,
                                    PoPriceUomId = @PoPriceUomId,
                                    SiPriceQty = @SiPriceQty,
                                    DiscountRate = @DiscountRate,
                                    DiscountAmount = @DiscountAmount,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE PoDetailId = @PoDetailId";
                            rowsAffected += sqlConnection.Execute(sql, updatePoDetail);
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
        #region //DeleteMaterial -- 刪除材料 -- Shinotokuro 2024.08.19
        public string DeleteMaterial(int MtId)
        {
            try
            {
                if (MtId <= 0) throw new SystemException("【材料】不能為空!");
                int rowsAffected = 0;

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷材料資訊是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.MaterialManagement
                                WHERE MtId = @MtId";
                        dynamicParameters.Add("MtId", MtId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【材料】找不到，請重新確認!");
                        #endregion

                        #region //判斷材料金額是否有資料
                        sql = @"SELECT TOP 1 1
                                FROM SCM.MtAmount
                                WHERE MtId = @MtId
                                AND ConfirmStatus = 'Y'";
                        dynamicParameters.Add("MtId", MtId);

                        result = sqlConnection.Query(sql, dynamicParameters);
                        if(result.Count() > 0) throw new SystemException("【材料金額】已經有資料，不可以刪除!");
                        #endregion

                        #region //資料刪除
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.MtAmount
                                WHERE MtId = @MtId";
                        dynamicParameters.Add("MtId", MtId);

                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.MaterialManagement
                                WHERE MtId = @MtId";
                        dynamicParameters.Add("MtId", MtId);

                        insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #region //DeleteMtAmount -- 刪除材料金額 -- Shinotokuro 2024.08.19
        public string DeleteMtAmount(int MaId)
        {
            try
            {
                if (MaId <= 0) throw new SystemException("【材料金額】不能為空!");
                int rowsAffected = 0;

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        DynamicParameters dynamicParameters = new DynamicParameters();

                        #region //判斷材料資訊是否正確
                        sql = @"SELECT TOP 1 ConfirmStatus
                                FROM SCM.MtAmount
                                WHERE MaId = @MaId";
                        dynamicParameters.Add("MaId", MaId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【材料金額】找不到，請重新確認!");
                        foreach(var item in result)
                        {
                            if(item.ConfirmStatus != "N") throw new SystemException("【材料金額】只有未確認狀態下才可以執行刪除!");
                        }
                        #endregion
                        

                        #region //資料刪除
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.MtAmount
                                WHERE MaId = @MaId";
                        dynamicParameters.Add("MaId", MaId);

                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

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

        #region //API For Supplier
        #region //GetCompanyKanbanInfo -- 取得採購端看板資訊 -- Chia Yuan 2025.3.27
        public string GetCompanyKanbanInfo(string CompanyNo, string NewCompanyNo, string SupplierNos)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //確認公司別DB
                    sql = @"SELECT a.ErpNo, a.ErpDb
                            FROM BAS.Company a
                            WHERE CompanyNo = @CompanyNo";
                    var resultCorp = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo = CompanyNo == "ETG" ? "Eterge" : CompanyNo }) ?? throw new SystemException("【公司別】資料錯誤!");
                    ErpConnectionStrings = ConfigurationManager.AppSettings[resultCorp.ErpDb];
                    #endregion

                    using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        sql = @"DECLARE @ThreeDaysAgo DATETIME
                                SET @ThreeDaysAgo = DATEADD(DAY, -3, GETDATE())
                                CREATE TABLE #TempStats (StatType VARCHAR(20), StatCount INT)
                                -- PR 統計
                                INSERT INTO #TempStats (StatType, StatCount)
                                SELECT 'prNotToPo', COUNT(1)
                                FROM PURTB aa
                                INNER JOIN PURTA ac ON ac.TA001 = aa.TB001 AND ac.TA002 = aa.TB002
                                INNER JOIN CMSML cmp ON cmp.COMPANY = aa.COMPANY
                                WHERE RTRIM(cmp.ML001) = @ErpNo
                                    AND aa.TB010 IN @SupplierNos
                                    AND aa.TB025 = 'Y'
                                    AND ac.TA007 = 'Y'
                                    AND NOT EXISTS (
                                        SELECT 1 FROM PURTD ab 
                                        WHERE ab.TD026 = aa.TB001 AND ab.TD027 = aa.TB002 AND ab.TD028 = aa.TB003
                                    )
                                INSERT INTO #TempStats (StatType, StatCount)
                                SELECT 'prNotToPo3Day', COUNT(1)
                                FROM PURTB aa
                                INNER JOIN PURTA ac ON ac.TA001 = aa.TB001 AND ac.TA002 = aa.TB002
                                INNER JOIN CMSML cmp ON cmp.COMPANY = aa.COMPANY
                                WHERE RTRIM(cmp.ML001) = @ErpNo
                                    AND aa.TB010 IN @SupplierNos
                                    AND aa.TB025 = 'Y'
                                    AND ac.TA007 = 'Y'
                                    AND CONVERT(DATETIME, ac.TA003, 112) <= @ThreeDaysAgo
                                    AND NOT EXISTS (
                                        SELECT 1 FROM PURTD ab 
                                        WHERE ab.TD026 = aa.TB001 AND ab.TD027 = aa.TB002 AND ab.TD028 = aa.TB003
                                    )
                                -- PO 統計
                                INSERT INTO #TempStats (StatType, StatCount)
                                SELECT 'poNotClose', COUNT(1)
                                FROM PURTD aa
                                INNER JOIN PURTC ac ON ac.TC001 = aa.TD001 AND ac.TC002 = aa.TD002
                                INNER JOIN CMSML cmp ON cmp.COMPANY = aa.COMPANY
                                WHERE RTRIM(cmp.ML001) = @ErpNo
                                    AND ac.TC004 IN @SupplierNos
                                    AND aa.TD016 = 'N'
                                INSERT INTO #TempStats (StatType, StatCount)
                                SELECT 'poNotClose3Day', COUNT(1)
                                FROM PURTD aa
                                INNER JOIN PURTC ac ON ac.TC001 = aa.TD001 AND ac.TC002 = aa.TD002
                                INNER JOIN CMSML cmp ON cmp.COMPANY = aa.COMPANY
                                WHERE RTRIM(cmp.ML001) = @ErpNo
                                    AND ac.TC004 IN @SupplierNos
                                    AND aa.TD016 = 'N'
                                    AND CONVERT(DATETIME, ac.TC003, 112) <= @ThreeDaysAgo
                                -- PU 統計
                                INSERT INTO #TempStats (StatType, StatCount)
                                SELECT 'poNotToPu', COUNT(1)
                                FROM PURTD aa
                                INNER JOIN PURTC ac ON ac.TC001 = aa.TD001 AND ac.TC002 = aa.TD002
                                INNER JOIN CMSML cmp ON cmp.COMPANY = aa.COMPANY
                                WHERE RTRIM(cmp.ML001) = @ErpNo
                                    AND ac.TC004 IN @SupplierNos
                                    AND aa.TD018 = 'Y'
                                    AND ac.TC014 = 'Y'
                                    AND NOT EXISTS (
                                        SELECT 1 FROM PURTH ab 
                                        WHERE ab.TH011 = aa.TD001 AND ab.TH012 = aa.TD002 AND ab.TH013 = aa.TD003
                                    )
                                INSERT INTO #TempStats (StatType, StatCount)
                                SELECT 'poNotToPu3Day', COUNT(1)
                                FROM PURTD aa
                                INNER JOIN PURTC ac ON ac.TC001 = aa.TD001 AND ac.TC002 = aa.TD002
                                INNER JOIN CMSML cmp ON cmp.COMPANY = aa.COMPANY
                                WHERE RTRIM(cmp.ML001) = @ErpNo
                                    AND ac.TC004 IN @SupplierNos
                                    AND aa.TD018 = 'Y'
                                    AND ac.TC014 = 'Y'
                                    AND CONVERT(DATETIME, ac.TC003, 112) <= @ThreeDaysAgo
                                    AND NOT EXISTS (
                                        SELECT 1 FROM PURTH ab 
                                        WHERE ab.TH011 = aa.TD001 AND ab.TH012 = aa.TD002 AND ab.TH013 = aa.TD003
                                    )
                                -- 最終結果
                                SELECT TOP 1 
                                    @NewCompanyNo CompanyNo,
                                    RTRIM(a.ML002) CompanyName,
                                    ISNULL((SELECT StatCount FROM #TempStats WHERE StatType = 'prNotToPo'), 0) prNotToPo,
                                    ISNULL((SELECT StatCount FROM #TempStats WHERE StatType = 'prNotToPo3Day'), 0) prNotToPo3Day,
                                    ISNULL((SELECT StatCount FROM #TempStats WHERE StatType = 'poNotClose'), 0) poNotClose,
                                    ISNULL((SELECT StatCount FROM #TempStats WHERE StatType = 'poNotClose3Day'), 0) poNotClose3Day,
                                    ISNULL((SELECT StatCount FROM #TempStats WHERE StatType = 'poNotToPu'), 0) poNotToPu,
                                    ISNULL((SELECT StatCount FROM #TempStats WHERE StatType = 'poNotToPu3Day'), 0) poNotToPu3Day
                                FROM CMSML a
                                WHERE RTRIM(a.ML001) = @ErpNo
                                DROP TABLE #TempStats";
                        var result = erpConnection.QueryFirstOrDefault(sql, new { NewCompanyNo, resultCorp.ErpNo, SupplierNos = SupplierNos.Split(',') });

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            data = result,
                            status = "success"
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

        #region //GetTaxTypeERP -- 取得ERP稅別資料 -- Chia Yuan 2025.03.07
        public string GetTaxTypeERP(List<SupplierCompany> SupplierCompany
            , string TaxCode, string TaxName, string TypeNo, string Taxation, string SearchKey
            , bool IsOptions, int draw
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                List<dynamic> result = new List<dynamic>();

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    SupplierCompany.Distinct().ToList().ForEach(x =>
                    {
                        #region //確認公司別DB
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        var resultCorp = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo = x.CompanyNo == "ETG" ? "Eterge" : x.CompanyNo }) ?? throw new SystemException("【公司別】資料錯誤!");
                        ErpConnectionStrings = ConfigurationManager.AppSettings[resultCorp.ErpDb];
                        #endregion

                        using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //條件資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(NN001)) TaxNo
                                    , LTRIM(RTRIM(NN002)) TaxName
                                    , CASE NN003 WHEN 1 THEN '進項' ELSE '銷項' END TypeName
                                    , ROUND(LTRIM(RTRIM(NN004)), 2) BusinessTaxRate
                                    , NN005 InvoiceType
                                    , LTRIM(RTRIM(NN006)) Taxation
                                    , LTRIM(RTRIM(NN001)) + ' ' + LTRIM(RTRIM(NN002)) TaxWithNo
                                    , @NewCompanyNo CompanyNo
                                    FROM CMSNN a
                                    WHERE 1=1";
                            dynamicParameters.Add("NewCompanyNo", x.NewCompanyNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TaxNo", @" AND LTRIM(RTRIM(a.NN001)) = @TaxNo", TaxCode);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TaxName", @" AND LTRIM(RTRIM(a.NN002)) = @TaxName", TaxName);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TypeNo", @" AND LTRIM(RTRIM(a.NN003)) = @TypeNo", TypeNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "Taxation", @" AND LTRIM(RTRIM(a.NN006)) = @Taxation", Taxation);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SearchKey",
                                @" AND (LTRIM(RTRIM(a.NN001)) LIKE N'%' + @SearchKey + '%' 
                                    OR LTRIM(RTRIM(a.NN002)) LIKE N'%' + @SearchKey + '%')", SearchKey);

                            result.AddRange(erpConnection.Query<dynamic>(sql, dynamicParameters).ToList());
                            #endregion
                        }
                    });

                    ApplyDynamicOrdering(ref result, OrderBy, "TaxNo");

                    if (PageSize > 0) result = result.Skip(Math.Max(PageIndex, 0) * PageSize).Take(PageSize).ToList(); //符合條件的品號(與庫存條件交集)

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        draw,
                        recordsTotal = result.Count,
                        recordsFiltered = result.Count,
                        data = result,
                        status = "success"
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

        #region //GetTradeTermERP --取得ERP交易條件 -- Shintokuro 2025-03-17
        public string GetTradeTermERP(List<SupplierCompany> SupplierCompany
            , bool IsOptions, int draw
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                List<dynamic> result = new List<dynamic>();

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    SupplierCompany.Distinct().ToList().ForEach(x =>
                    {
                        #region //確認公司別DB
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        var resultCorp = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo = x.CompanyNo == "ETG" ? "Eterge" : x.CompanyNo }) ?? throw new SystemException("【公司別】資料錯誤!");
                        ErpConnectionStrings = ConfigurationManager.AppSettings[resultCorp.ErpDb];
                        #endregion

                        using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //取得交易條件
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(NK001)) TradeTermNo, LTRIM(RTRIM(NK002)) TradeTermName
                                    , @NewCompanyNo CompanyNo
                                    FROM CMSNK
                                    WHERE 1=1";
                            dynamicParameters.Add("NewCompanyNo", x.NewCompanyNo);
                            result.AddRange(erpConnection.Query<dynamic>(sql, dynamicParameters).ToList());
                            #endregion
                        }
                    });

                    ApplyDynamicOrdering(ref result, OrderBy, "TradeTermNo");

                    if (PageSize > 0) result = result.Skip(Math.Max(PageIndex, 0) * PageSize).Take(PageSize).ToList(); //符合條件的品號(與庫存條件交集)

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        draw,
                        recordsTotal = result.Count,
                        recordsFiltered = result.Count,
                        data = result,
                        status = "success"
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

        #region //GetPaymentTermERP 取得ERP付款條件資料 2024.09.15
        public string GetPaymentTermERP(List<SupplierCompany> SupplierCompany
            , string PaymentType, string PaymentTermCode, string PaymentTermCodes, string SearchKey
            , bool IsOptions, int draw
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                List<dynamic> result = new List<dynamic>();

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    SupplierCompany.Distinct().ToList().ForEach(x =>
                    {
                        #region //確認公司別DB
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        var resultCorp = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo = x.CompanyNo == "ETG" ? "Eterge" : x.CompanyNo }) ?? throw new SystemException("【公司別】資料錯誤!");
                        ErpConnectionStrings = ConfigurationManager.AppSettings[resultCorp.ErpDb];
                        #endregion

                        using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //取得交易條件
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(a.NA001)) PaymentType, LTRIM(RTRIM(a.NA002)) PaymentTermCode
                                    , LTRIM(RTRIM(a.NA003)) PaymentTermName, a.NA002 + ' ' + a.NA003 PaymentTermWithCode
                                    , @NewCompanyNo CompanyNo
                                    FROM CMSNA a
                                    WHERE 1=1";
                            dynamicParameters.Add("NewCompanyNo", x.NewCompanyNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "NA001", @" AND LTRIM(RTRIM(a.NA001)) = @NA001", PaymentType);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "NA002", @" AND LTRIM(RTRIM(a.NA002)) = @NA002", PaymentTermCode);
                            if (PaymentTermCodes.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "NA002s", @" AND LTRIM(RTRIM(a.NA002)) IN @NA002s", PaymentTermCodes.Split(','));
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SearchKey",
                                @" AND (a.NA002 LIKE N'%' + @SearchKey + '%' 
                                    OR a.NA003 LIKE N'%' + @SearchKey + '%')", SearchKey);
                            result.AddRange(erpConnection.Query<dynamic>(sql, dynamicParameters).ToList());
                            #endregion
                        }
                    });

                    ApplyDynamicOrdering(ref result, OrderBy, "TradeTermNo");

                    if (PageSize > 0) result = result.Skip(Math.Max(PageIndex, 0) * PageSize).Take(PageSize).ToList(); //符合條件的品號(與庫存條件交集)

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        draw,
                        recordsTotal = result.Count,
                        recordsFiltered = result.Count,
                        data = result,
                        status = "success"
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

        #region //GetCurrencyERP -- 取得ERP幣別資料 -- Chia Yuan 2025.03.05
        public string GetCurrencyInfo(List<SupplierCompany> SupplierCompany
            , string Currency, string CurrencyName, string SearchKey
            , bool IsOptions, int draw
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    List<dynamic> result = new List<dynamic>();

                    SupplierCompany.Distinct().ToList().ForEach(x =>
                    {
                        #region //確認公司別DB
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        var resultCorp = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo = x.CompanyNo == "ETG" ? "Eterge" : x.CompanyNo }) ?? throw new SystemException("【公司別】資料錯誤!");
                        ErpConnectionStrings = ConfigurationManager.AppSettings[resultCorp.ErpDb];
                        #endregion

                        using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //條件資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(a.MF001)) Currency
                                    , LTRIM(RTRIM(a.MF002)) CurrencyName, LTRIM(RTRIM(a.MF001)) + ' ' + LTRIM(RTRIM(a.MF002)) CurrencyWithNo
                                    , a.MF003 UnitPriceRound, a.MF004 AmountRound, a.MF005 UnitCostRound, a.MF006 CostRound
                                    , @NewCompanyNo CompanyNo
                                    FROM CMSMF a
                                    WHERE 1=1";
                            dynamicParameters.Add("NewCompanyNo", x.NewCompanyNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MF001", @" AND LTRIM(RTRIM(a.MF001)) = @MF001", Currency);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MF002", @" AND a.MF002 = @MF002", CurrencyName);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SearchKey", @" AND (a.MF001 LIKE N'%' + @SearchKey + '%' OR a.MF002 LIKE N'%' + @SearchKey + '%')", SearchKey);
                            result.AddRange(erpConnection.Query<dynamic>(sql, dynamicParameters).ToList());
                            #endregion
                        }
                    });

                    ApplyDynamicOrdering(ref result, OrderBy, "Currency");

                    if (PageSize > 0) result = result.Skip(Math.Max(PageIndex, 0) * PageSize).Take(PageSize).ToList(); //符合條件的品號(與庫存條件交集)

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        draw,
                        recordsTotal = result.Count,
                        recordsFiltered = result.Count,
                        data = result,
                        status = "success"
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

        #region //GetExchangeRate --取得ERP匯率資料 -- Chia Yuan 2025.03.05
        public string GetExchangeRate(List<SupplierCompany> SupplierCompany
            , string Currency, string StartDate, string SearchKey
            , bool IsOptions, int draw, bool TopOne
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    List<dynamic> result = new List<dynamic>();

                    SupplierCompany.Distinct().ToList().ForEach(x =>
                    {
                        #region //確認公司別DB
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        var resultCorp = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo = x.CompanyNo == "ETG" ? "Eterge" : x.CompanyNo }) ?? throw new SystemException("【公司別】資料錯誤!");
                        ErpConnectionStrings = ConfigurationManager.AppSettings[resultCorp.ErpDb];
                        #endregion

                        using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            var topOne = TopOne ? "TOP 1" : "";

                            #region //條件資料
                            dynamicParameters = new DynamicParameters();
                            sql = $@"SELECT {topOne} LTRIM(RTRIM(a.MG001)) Currency, CONVERT(DATETIME, CASE a.MG002 WHEN '' THEN NULL ELSE a.MG002 END) EffectiveDate
                                    , ROUND(LTRIM(RTRIM(a.MG003)), 4) BankBuyingRate
                                    , ROUND(LTRIM(RTRIM(a.MG004)), 4) BankSellingRate
                                    , ROUND(LTRIM(RTRIM(a.MG005)), 4) CustomsBuyingRate
                                    , ROUND(LTRIM(RTRIM(a.MG006)), 4) CustomsSellingRate
                                    , b.MF003 UnitPriceDecimal, b.MF004 AmountDecimal, b.MF005 UnitCostDecimal, b.MF006 CostAmountDecimal
                                    , @NewCompanyNo CompanyNo
                                    FROM CMSMG a
                                    INNER JOIN CMSMF b ON b.MF001 = a.MG001
                                    WHERE 1=1";
                            dynamicParameters.Add("NewCompanyNo", x.NewCompanyNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MG001", @" AND LTRIM(RTRIM(a.MG001)) = @MG001", Currency);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MG002", @" AND a.MG002 <= @MG002",
                                StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyyMMdd") : string.Empty);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SearchKey", @" AND a.MG001 LIKE N'%' + @SearchKey + '%'", SearchKey);

                            sql += OrderBy.Length > 0 ? "  ORDER BY " + OrderBy : " ORDER BY Currency, EffectiveDate DESC";
                            result.AddRange(erpConnection.Query<dynamic>(sql, dynamicParameters).ToList());
                            #endregion
                        }
                    });

                    ApplyDynamicOrdering(ref result, OrderBy, "Currency");

                    if (PageSize > 0) result = result.Skip(Math.Max(PageIndex, 0) * PageSize).Take(PageSize).ToList(); //符合條件的品號(與庫存條件交集)

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        draw,
                        recordsTotal = result.Count,
                        recordsFiltered = result.Count,
                        data = result,
                        status = "success"
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

        #region //GetItemUnitConvertERP -- 取得ERP單位換算資料 -- Chia Yuan 2025.04.10
        public string GetItemUnitConvertERP(List<SupplierCompany> SupplierCompany
            , string MtlItemNo, string MtlItemNos, string ConvertUnit, string Unit, string SearchKey
            , bool IsOptions, int draw
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    List<dynamic> result = new List<dynamic>();

                    SupplierCompany.Distinct().ToList().ForEach(x =>
                    {
                        #region //確認公司別DB
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        var resultCorp = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo = x.CompanyNo == "ETG" ? "Eterge" : x.CompanyNo }) ?? throw new SystemException("【公司別】資料錯誤!");
                        ErpConnectionStrings = ConfigurationManager.AppSettings[resultCorp.ErpDb];
                        #endregion

                        using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //條件資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(a.MD001)) MtlItemNo
                                    , LTRIM(RTRIM(a.MD002)) ConversionUnit, a.MD003 SwapNumerator, a.MD004 SwapDenominator, LTRIM(RTRIM(b.MB004)) Unit
                                    , STUFF(STUFF(a.CREATE_DATE + ' ' + a.CREATE_TIME, 5, 0, '-'), 8, 0, '-') CreateDate
                                    , STUFF(STUFF(a.MODI_DATE + ' ' + a.MODI_TIME, 5, 0, '-'), 8, 0, '-') LastModifiedDate
                                    , LTRIM(RTRIM(a.CREATOR)) CreateBy, LTRIM(RTRIM(a.MODIFIER)) LastModifiedBy
                                    , @NewCompanyNo CompanyNo
                                    FROM INVMD a
                                    INNER JOIN INVMB b ON b.MB001 = a.MD001
                                    WHERE 1=1";
                            dynamicParameters.Add("NewCompanyNo", x.NewCompanyNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MD001", @" AND LTRIM(RTRIM(a.MD001)) = @MD001", MtlItemNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MD002", @" AND LTRIM(RTRIM(a.MD002)) = @MD002", ConvertUnit);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MB004", @" AND LTRIM(RTRIM(b.MB004)) = @MB004", Unit);
                            if (MtlItemNos.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MD001s", @" AND LTRIM(RTRIM(a.MD001)) IN @MD001s", MtlItemNos.Split(','));
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SearchKey",
                                @" AND (LTRIM(RTRIM(a.MD001)) LIKE N'%' + @SearchKey + '%'
                                    OR a.MD002 LIKE N'%' + @SearchKey + '%'
                                    OR b.MB004 LIKE N'%' + @SearchKey + '%')", SearchKey);

                            //sql += OrderBy.Length > 0 ? "  ORDER BY " + OrderBy : " ORDER BY CompanyName";
                            result.AddRange(erpConnection.Query<dynamic>(sql, dynamicParameters).ToList());
                            #endregion
                        }
                    });

                    ApplyDynamicOrdering(ref result, OrderBy, "MtlItemNo");

                    if (PageSize > 0) result = result.Skip(Math.Max(PageIndex, 0) * PageSize).Take(PageSize).ToList(); //符合條件的品號(與庫存條件交集)

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        draw,
                        recordsTotal = result.Count,
                        recordsFiltered = result.Count,
                        data = result,
                        status = "success"
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

        #region //GetSupplierERP -- 取得ERP廠商資料 -- Chia Yuan 2025.01.24
        public string GetSupplierERP(List<SupplierCompany> SupplierCompany
            , string LastTradingStartDate, string LastTradingEndDate, string SearchKey
            , bool IsOptions, int draw
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    List<(string, string)> Nos = new List<(string, string)>();

                    SupplierCompany.Distinct().ToList().ForEach(x =>
                    {
                        #region //確認公司別DB
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        var resultCorp = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo = x.CompanyNo == "ETG" ? "Eterge" : x.CompanyNo }) ?? throw new SystemException("【公司別】資料錯誤!");
                        ErpConnectionStrings = ConfigurationManager.AppSettings[resultCorp.ErpDb];
                        #endregion

                        using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //條件資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(a.MA001)) SupplierNo
                                    FROM PURMA a
                                    WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MA001", @" AND a.MA001 = @MA001", x.SupplierNo);
                            if (x.SupplierNos.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MA001s", @" AND a.MA001 IN @MA001s", x.SupplierNos.Split(','));
                            if (LastTradingStartDate.Length > 0 && DateTime.TryParse(LastTradingStartDate, out DateTime lastTradingStartDate))
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "LastTradingStartDate", @" AND a.MA023 >= @LastTradingStartDate", lastTradingStartDate.ToString("yyyyMMdd"));
                            if (LastTradingEndDate.Length > 0 && DateTime.TryParse(LastTradingEndDate, out DateTime lastTradingEndDate))
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "LastTradingEndDate", @" AND a.MA023 <= @LastTradingEndDate", lastTradingEndDate.AddDays(1).ToString("yyyyMMdd"));
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SearchKey",
                                @" AND (a.MA001 LIKE N'%' + @SearchKey + '%' 
                                    OR a.MA002 LIKE N'%' + @SearchKey + '%' 
                                    OR a.MA003 LIKE N'%' + @SearchKey + '%' 
                                    OR a.MA088 LIKE N'%' + @SearchKey + '%')", SearchKey);
                            sql += OrderBy.Length > 0 ? " ORDER BY " + OrderBy : " ORDER BY SupplierNo";

                            Nos.AddRange(erpConnection.Query(sql, dynamicParameters).Select(s => ((string)resultCorp.ErpNo, string.Format("{0}", (string)s.SupplierNo))).Distinct().ToList());
                            #endregion
                        }
                    });

                    List<(string, string)> PagedNos = new List<(string, string)>();
                    if (PageSize > 0) PagedNos = Nos.Skip(Math.Max(PageIndex, 0) * PageSize).Take(PageSize).ToList(); //符合條件的品號(與庫存條件交集)
                    else PagedNos = Nos;

                    List<dynamic> result = new List<dynamic>();

                    SupplierCompany.Distinct().ToList().ForEach(x =>
                    {
                        #region //確認公司別DB
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        var resultCorp = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo = x.CompanyNo == "ETG" ? "Eterge" : x.CompanyNo }) ?? throw new SystemException("【公司別】資料錯誤!");
                        ErpConnectionStrings = ConfigurationManager.AppSettings[resultCorp.ErpDb];
                        #endregion

                        string[] KeyNos = PagedNos.Where(w => w.Item1 == resultCorp.ErpNo).Select(s => s.Item2).ToArray();

                        if (KeyNos.Any())
                        {
                            using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                            {
                                #region //取得資料
                                sql = IsOptions ?
                                    @"SELECT LTRIM(RTRIM(a.MA001)) SupplierNo
                                    , LTRIM(RTRIM(a.MA001)) + ' ' + LTRIM(RTRIM(a.MA002)) SupplierWithNo
                                    , LTRIM(RTRIM(a.MA002)) SupplierName
                                    , LTRIM(RTRIM(a.MA003)) SupplierFullName
                                    , LTRIM(RTRIM(a.MA088)) SupplierEnglishName
                                    , LTRIM(RTRIM(c.ML003)) CompanyName, @NewCompanyNo CompanyNo" :
                                    @"SELECT LTRIM(RTRIM(a.MA001)) SupplierNo
                                    , LTRIM(RTRIM(a.MA001)) + ' ' + LTRIM(RTRIM(a.MA002)) SupplierWithNo
                                    , LTRIM(RTRIM(a.MA002)) SupplierName
                                    , LTRIM(RTRIM(a.MA003)) SupplierFullName
                                    , LTRIM(RTRIM(a.MA088)) SupplierEnglishName
                                    , a.MA004 SupplierType, a.MA005 GuiNumber, a.MA006 Country, a.MA007 Region, a.MA008 TelNoFirst, a.MA009 TelNoSecond
                                    , a.MA010 FaxNo, a.MA011 Email, a.MA012 ResponsiblePerson, a.MA013 ContactPersonFirst, a.MA014 AddressFirst, a.MA015 AddressSecond
                                    , a.MA016 PermitStatus, a.MA017 OpeningDate, a.MA018 Capital, a.MA019 HeadCount, a.MA020 PoDeliver
                                    , a.MA021 Currency, a.MA022 FirstTradingDay, a.MA023 LastTradingDay, a.MA024 PaymentType, a.MA025 PaymentTermName
                                    , a.MA026 PriceTerm, LTRIM(RTRIM(a.MA027)) RemitBank, LTRIM(RTRIM(a.MA028)) RemitAccount, a.MA029 ReceiptReceive, a.MA030 InvoiceType
                                    , a.MA031 SupplierLevel, a.MA032 DeliveryRating, a.MA033 QualityRating, a.MA034 AccountMonth, a.MA035 AccountDay
                                    , a.MA036 PaymentMonth, a.MA037 PaymentDay, a.MA038 InvoiceMonth, a.MA039 InvoiceDay, a.MA040 SupplierRemark
                                    , a.MA041 AccountPayable, a.MA042 AccountOverhead, a.MA043 AccountInvoice, a.MA044 Taxation, a.MA045 PermitPartialDelivery
                                    , a.MA046 ZipCodeFirst, a.MA047 PurchaseUser, a.MA048 ContactPersonSecond, a.MA049 ContactPersonThird, a.MA050 ZipCodeSecond
                                    , a.MA051 BillAddressFirst, a.MA052 BillAddressSecond, a.MA053 TaxAmountCalculateType, a.MA055 PaymentTerm
                                    , a.MA056 InvocieAttachedStatus, a.MA057 CertificateFormatType, a.MA058 DepositRate, a.MA059 FaxNoAccounting
                                    , a.MA066 TradeItem, a.MA067 HeadOffice, a.MA068 HeadOfficeReceive, a.MA069 'Version', a.MA070 PermitDate
                                    , a.MA076 TaxNo, a.MA077 TradeTerm, a.MA085 RelatedPerson
                                    , ISNULL(b.MB002, '') RelatedPersonName
                                    , LTRIM(RTRIM(c.ML002)) CompanyShortName, LTRIM(RTRIM(c.ML003)) CompanyName
                                    , LTRIM(RTRIM(CASE WHEN CHARINDEX('-', u.MF002) > 0 THEN LEFT(u.MF002, CHARINDEX('-', u.MF002) - 1) ELSE u.MF002 END)) CreatorName
                                    , STUFF(STUFF(a.CREATE_DATE + ' ' + a.CREATE_TIME, 5, 0, '-'), 8, 0, '-') CreateDate
                                    , STUFF(STUFF(a.MODI_DATE + ' ' + a.MODI_TIME, 5, 0, '-'), 8, 0, '-') LastModifiedDate
                                    , LTRIM(RTRIM(a.CREATOR)) CreateBy, LTRIM(RTRIM(a.MODIFIER)) LastModifiedBy
                                    , ISNULL(b1.MO006, '') RemitBankName
                                    , ISNULL(c1.MA003, '') AccountOverheadName
                                    , ISNULL(c2.MA003, '') AccountPayableName
                                    , ISNULL(c3.MA003, '') AccountInvoiceName
                                    , ISNULL(f.NK002, '') TradeTermName
                                    , ISNULL(g.MF002, '') CurrencyName
                                    , ISNULL(h.NN002, '') TaxName
                                    , @NewCompanyNo CompanyNo";
                                sql += @" FROM PURMA a
                                    LEFT JOIN DSCSYS.dbo.CMSMO b1 ON b1.MO001 = a.MA027
                                    LEFT JOIN FCSMB b ON b.MB001 = a.MA085
                                    INNER JOIN CMSML c ON c.COMPANY = a.COMPANY
                                    LEFT JOIN ACTMA c1 ON c1.MA001 = a.MA042
                                    LEFT JOIN ACTMA c2 ON c2.MA001 = a.MA041
                                    LEFT JOIN ACTMA c3 ON c3.MA001 = a.MA043
                                    LEFT JOIN CMSNK f ON f.NK001 = a.MA077
                                    LEFT JOIN CMSMF g ON g.MF001 = a.MA014
                                    LEFT JOIN CMSNN h ON h.NN001 = a.MA076
                                    LEFT JOIN ADMMF u ON LTRIM(RTRIM(u.MF001)) = LTRIM(RTRIM(a.CREATOR))
                                    WHERE LTRIM(RTRIM(a.MA001)) IN @KeyNos";

                                const int batchSize = 2000;
                                for (int i = 0; i < PagedNos.Count; i += batchSize)
                                {
                                    result.AddRange(erpConnection.Query<dynamic>(sql, new { x.NewCompanyNo, KeyNos = KeyNos.Skip(i).Take(batchSize).ToArray() }).ToList());
                                }
                                #endregion
                            }
                        }
                    });

                    ApplyDynamicOrdering(ref result, OrderBy, "SupplierNo");

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        draw,
                        recordsTotal = Nos.Count,
                        recordsFiltered = Nos.Count,
                        data = result,
                        status = "success"
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

        #region //GetCompanyERP -- 取得ERP公司資料 -- Chia Yuan 2025.03.05
        public string GetCompanyERP(List<SupplierCompany> SupplierCompany
            , string SearchKey
            , bool IsOptions, int draw
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    List<dynamic> result = new List<dynamic>();

                    SupplierCompany.Distinct().ToList().ForEach(x =>
                    {
                        #region //確認公司別DB
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        var resultCorp = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo = x.CompanyNo == "ETG" ? "Eterge" : x.CompanyNo }) ?? throw new SystemException("【公司別】資料錯誤!");
                        ErpConnectionStrings = ConfigurationManager.AppSettings[resultCorp.ErpDb];
                        #endregion

                        using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //取得資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(a.ML002)) CompanyShortName, LTRIM(RTRIM(a.ML003)) CompanyName
                                    , LTRIM(RTRIM(a.ML012)) AddressFirst, LTRIM(RTRIM(a.ML013)) AddressSecond
                                    , LTRIM(RTRIM(a.ML014)) CompanyEnglishName, LTRIM(RTRIM(a.ML015)) EnglishAddressFirst, LTRIM(RTRIM(a.ML016)) EnglishAddressSecond
                                    , @NewCompanyNo CompanyNo
                                    FROM CMSML a
                                    WHERE 1=1";
                            dynamicParameters.Add("NewCompanyNo", x.NewCompanyNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SearchKey",
                                @" AND (LTRIM(RTRIM(a.ML002)) LIKE N'%' + @SearchKey + '%'
                                    OR  LTRIM(RTRIM(a.ML003)) LIKE N'%' + @SearchKey + '%' 
                                    OR  LTRIM(RTRIM(a.ML014)) LIKE N'%' + @SearchKey + '%')", SearchKey);

                            //sql += OrderBy.Length > 0 ? "  ORDER BY " + OrderBy : " ORDER BY CompanyName";
                            result.AddRange(erpConnection.Query<dynamic>(sql, dynamicParameters).ToList());
                            #endregion
                        }
                    });

                    ApplyDynamicOrdering(ref result, OrderBy, "CompanyName");

                    if (PageSize > 0) result = result.Skip(Math.Max(PageIndex, 0) * PageSize).Take(PageSize).ToList(); //符合條件的品號(與庫存條件交集)

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        draw,
                        recordsTotal = result.Count,
                        recordsFiltered = result.Count,
                        data = result,
                        status = "success"
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

        #region //GetMtlItemERP -- 取得ERP品號資料 -- Chia Yuan 2025.03.05
        public string GetMtlItemERP(List<SupplierCompany> SupplierCompany
            , string MtlItemNo, string MtlItemNos, string SearchKey
            , bool IsOptions, int draw
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    List<dynamic> result = new List<dynamic>();

                    SupplierCompany.Distinct().ToList().ForEach(x =>
                    {
                        #region //確認公司別DB
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        var resultCorp = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo = x.CompanyNo == "ETG" ? "Eterge" : x.CompanyNo }) ?? throw new SystemException("【公司別】資料錯誤!");
                        ErpConnectionStrings = ConfigurationManager.AppSettings[resultCorp.ErpDb];
                        #endregion

                        using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //取得資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(a.MB001)) MtlItemNo, LTRIM(RTRIM(a.MB002)) MtlItemName, LTRIM(RTRIM(a.MB003)) MtlItemSpec, a.MB004 Unit
                                    , a.MB005 TypeOne, a.MB006 TypeTwo, a.MB007 TypeThree, a.MB008 TypeFour
                                    , a.MB009 MtlItemDesc, a.MB015 WeightUnit, a.MB017 Inventory, a.MB019 InventoryManagement, a.MB020 BondedStore, a.MB022 LotManagement
                                    , CONVERT(INT, a.MB023) EfficientDays, CONVERT(INT, a.MB024) RetestDays, a.MB025 ItemAttribute, a.MB028 MtlItemRemark
                                    , STUFF(STUFF(a.MB030, 5, 0, '-'), 8, 0, '-') EffectiveDate
                                    , STUFF(STUFF(a.MB031, 5, 0, '-'), 8, 0, '-') ExpirationDate
                                    , a.MB034 ReplenishmentPolicy, a.MB043 MeasureType, a.MB044 OverReceiptManagement, a.MB045 OverReceiptRatio
                                    , a.MB064 Quantity, a.MB065 Amount, a.MB066 MtlModify, a.MB068 ProductionLine, a.MB070 PriceSix, a.MB087 OverDeliveryManagement, a.MB088 OverDeliveryRatio
                                    , a.MB155 PoUnit, a.MB156 SoUnit, a.MB157 RequisitionInventory, a.MB165 [Version]
                                    , STUFF(STUFF(a.CREATE_DATE + ' ' + a.CREATE_TIME, 5, 0, '-'), 8, 0, '-') CreateDate
                                    , STUFF(STUFF(a.MODI_DATE + ' ' + a.MODI_TIME, 5, 0, '-'), 8, 0, '-') LastModifiedDate
                                    , LTRIM(RTRIM(a.CREATOR)) CreateBy, LTRIM(RTRIM(a.MODIFIER)) LastModifiedBy
                                    , @NewCompanyNo CompanyNo
                                    FROM INVMB a 
                                    LEFT JOIN ADMMF u ON LTRIM(RTRIM(u.MF001)) = LTRIM(RTRIM(a.CREATOR))
                                    WHERE 1=1";
                            dynamicParameters.Add("NewCompanyNo", x.NewCompanyNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MB001", @" AND LTRIM(RTRIM(a.MB001)) = @MB001", MtlItemNo);
                            if (MtlItemNos.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MB001s", @" AND LTRIM(RTRIM(a.MB001)) IN @MB001s", MtlItemNos.Split(','));
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SearchKey",
                                @" AND (LTRIM(RTRIM(a.MB001)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.MB002)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.MB003)) LIKE N'%' + @SearchKey + '%')", SearchKey);

                            result.AddRange(erpConnection.Query<dynamic>(sql, dynamicParameters).ToList());
                            #endregion
                        }
                    });

                    ApplyDynamicOrdering(ref result, OrderBy, "MtlItemNo");

                    if (PageSize > 0) result = result.Skip(Math.Max(PageIndex, 0) * PageSize).Take(PageSize).ToList(); //符合條件的品號(與庫存條件交集)

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        draw,
                        recordsTotal = result.Count,
                        recordsFiltered = result.Count,
                        data = result,
                        status = "success"
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

        #region //GetPurchaseRequisitionERP -- 取得ERP請購單頭資料 -- Chia Yuan 2025.3.26
        public string GetPurchaseRequisitionERP(List<SupplierCompany> SupplierCompany
            , string PrErpPrefix, string PrErpNo, string PrErpFullNo
            , string SupplierNo, string SupplierNos
            , string PrErpPrefixs, string PrErpNos
            , string PrUser, string MtlItemNo, string MtlItemNos
            , string PrStartDate, string PrEndDate
            , string DocStartDate, string DocEndDate
            , string ConfirmStatus, string ClosureStatus, string SearchKey
            , bool IsOptions, int draw
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    List<(string, string)> Nos = new List<(string, string)>();

                    SupplierCompany.Distinct().ToList().ForEach(x =>
                    {
                        #region //確認公司別DB
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        var resultCorp = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo = x.CompanyNo == "ETG" ? "Eterge" : x.CompanyNo }) ?? throw new SystemException("【公司別】資料錯誤!");
                        ErpConnectionStrings = ConfigurationManager.AppSettings[resultCorp.ErpDb];
                        #endregion

                        using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //條件資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(a.TA001)) PrErpPrefix, LTRIM(RTRIM(a.TA002)) PrErpNo
                                    FROM PURTA a
                                    WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TA001", @" AND LTRIM(RTRIM(a.TA001)) = @TA001", PrErpPrefix);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TA002", @" AND LTRIM(RTRIM(a.TA002)) = @TA002", PrErpNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PoErpFullNo", @" AND LTRIM(RTRIM(a.TA001)) + '-' + LTRIM(RTRIM(a.TA002)) = @PoErpFullNo", PrErpFullNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TA012", @" AND a.TA012 = @TA012", PrUser);
                            if (PrErpPrefixs.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TA001s", @" AND LTRIM(RTRIM(a.TA001)) IN @TA001s", PrErpPrefixs.Split(','));
                            if (PrErpNos.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TA002s", @" AND LTRIM(RTRIM(a.TA002)) IN @TA002s", PrErpNos.Split(','));
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PoStartDate", @" AND a.TA003 >= @PoStartDate",
                                PrStartDate.Length > 0 ? Convert.ToDateTime(PrStartDate).ToString("yyyyMMdd") : string.Empty);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PoEndDate", @" AND a.TA003 <= @PoEndDate",
                                PrEndDate.Length > 0 ? Convert.ToDateTime(PrEndDate).AddDays(1).ToString("yyyyMMdd") : string.Empty);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "DocStartDate", @" AND a.TA013 >= @DocStartDate",
                                DocStartDate.Length > 0 ? Convert.ToDateTime(DocStartDate).ToString("yyyyMMdd") : string.Empty);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "DocEndDate", @" AND a.TA013 <= @DocEndDate",
                                DocEndDate.Length > 0 ? Convert.ToDateTime(DocEndDate).AddDays(1).ToString("yyyyMMdd") : string.Empty);

                            if (MtlItemNo.Length + MtlItemNos.Length + SupplierNo.Length + SupplierNos.Length > 0)
                            {
                                sql += @" AND EXISTS (SELECT TOP 1 1 FROM PURTB aa WHERE LTRIM(RTRIM(aa.TB001)) = LTRIM(RTRIM(a.TA001)) AND LTRIM(RTRIM(aa.TB002)) = LTRIM(RTRIM(a.TA002))";
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TB004", @" AND LTRIM(RTRIM(aa.TB004)) = @TB004", MtlItemNo);
                                if (MtlItemNos.Length > 0)
                                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TB004s", @" AND LTRIM(RTRIM(aa.TB004)) IN @TB004s", MtlItemNos.Split(',').Distinct());
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TB010", @" AND LTRIM(RTRIM(aa.TB010)) = @TB010", MtlItemNo);
                                if (SupplierNos.Length > 0)
                                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TB010s", @" AND LTRIM(RTRIM(aa.TB010)) IN @TB010s", SupplierNos.Split(',').Distinct());
                                sql += ")";
                            }
                            if (ConfirmStatus.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TA007", @" AND a.TA007 IN @TA007", ConfirmStatus.Split(','));
                            if (Regex.IsMatch(ClosureStatus, "^(Y|N)$", RegexOptions.IgnoreCase))
                            {
                                if (ClosureStatus == "N")
                                {
                                    sql += @" AND a.TA007 != 'V'
                                            AND EXISTS (SELECT TOP 1 1 FROM PURTB aa 
                                                WHERE LTRIM(RTRIM(aa.TB001)) = LTRIM(RTRIM(a.TA001)) AND LTRIM(RTRIM(aa.TB002)) = LTRIM(RTRIM(a.TA002)) 
                                                AND aa.TB039 = 'N'
                                                AND aa.TB025 != 'V')";
                                }
                                if (ClosureStatus == "Y") 
                                {
                                    sql += @" AND a.TA007 = 'Y'
                                            AND EXISTS (SELECT TOP 1 1 FROM PURTB aa 
                                                WHERE LTRIM(RTRIM(aa.TB001)) = LTRIM(RTRIM(a.TA001)) AND LTRIM(RTRIM(aa.TB002)) = LTRIM(RTRIM(a.TA002))
                                                AND aa.TB039 != 'N')";
                                }
                            }
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SearchKey",
                                @" AND (LTRIM(RTRIM(a.TA001)) + '-' + LTRIM(RTRIM(a.TA002)) LIKE N'%' + @SearchKey + '%'
                                OR a.TA001 LIKE N'%' + @SearchKey + '%' 
                                OR a.TA002 LIKE N'%' + @SearchKey + '%'
                                OR EXISTS (
                                    SELECT TOP 1 1 
                                    FROM PURTB aa 
                                    WHERE LTRIM(RTRIM(aa.TB001)) = LTRIM(RTRIM(a.TA001)) AND LTRIM(RTRIM(aa.TB002)) = LTRIM(RTRIM(a.TA002)) 
                                    AND (LTRIM(RTRIM(aa.TB004)) LIKE N'%' + @SearchKey + '%'
                                        OR LTRIM(RTRIM(aa.TB005)) LIKE N'%' + @SearchKey + '%')))", SearchKey);
                            sql += OrderBy.Length > 0 ? "  ORDER BY " + OrderBy : " ORDER BY PrErpNo DESC, PrErpPrefix";

                            Nos.AddRange(erpConnection.Query(sql, dynamicParameters)
                                .Select(s => ((string)resultCorp.ErpNo, string.Format("{0}{1}", (string)s.PrErpPrefix, (string)s.PrErpNo)))
                                .ToList()); //取得識別單號
                            #endregion
                        }
                    });

                    List<(string, string)> PagedNos = new List<(string, string)>();
                    if (PageSize > 0) PagedNos = Nos.Skip(Math.Max(PageIndex, 0) * PageSize).Take(PageSize).ToList(); //符合條件的品號(與庫存條件交集)
                    else PagedNos = Nos;

                    List<dynamic> result = new List<dynamic>();

                    SupplierCompany.Distinct().ToList().ForEach(x =>
                    {
                        #region //確認公司別DB
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        var resultCorp = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo = x.CompanyNo == "ETG" ? "Eterge" : x.CompanyNo }) ?? throw new SystemException("【公司別】資料錯誤!");
                        ErpConnectionStrings = ConfigurationManager.AppSettings[resultCorp.ErpDb];
                        #endregion

                        string[] KeyNos = PagedNos.Where(w => w.Item1 == resultCorp.ErpNo).Select(s => s.Item2).ToArray();

                        if (KeyNos.Any())
                        {
                            using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                            {
                                #region //取得資料
                                sql = IsOptions ?
                                    @"SELECT LTRIM(RTRIM(a.TA001)) PrErpPrefix, LTRIM(RTRIM(a.TA002)) PrErpNo, LTRIM(RTRIM(a.TA001)) + '-' + LTRIM(RTRIM(a.TA002)) PrErpFullNo
                                    , @NewCompanyNo CompanyNo" :
                                    @"SELECT LTRIM(RTRIM(a.TA001)) PrErpPrefix, LTRIM(RTRIM(a.TA002)) PrErpNo, LTRIM(RTRIM(a.TA001)) + '-' + LTRIM(RTRIM(a.TA002)) PrErpFullNo
                                    , STUFF(STUFF(a.TA003, 5, 0, '-'), 8, 0, '-') PrDate
                                    , a.TA004 PrDepartmentNo, a.TA006 Remark
                                    , a.TA007 ConfirmStatus, a.TA008 PrintCount, a.TA011 Quantity, LTRIM(RTRIM(a.TA012)) PrUser, LTRIM(RTRIM(u1.MF002)) PrUserName
                                    , STUFF(STUFF(a.TA013, 5, 0, '-'), 8, 0, '-') DocDate
                                    , a.TA014 ConfirmUser, a.TA015 PackageQty, a.TA016 ApproveStatus
                                    , a.TA017 TransmissionCount, a.TA020 TotalAmount, a.TA021 Edition, a.TA550 LockStaus
                                    , LTRIM(RTRIM(c.ML002)) CompanyShortName, LTRIM(RTRIM(c.ML003)) CompanyName
                                    , LTRIM(RTRIM(CASE WHEN CHARINDEX('-', u.MF002) > 0 THEN LEFT(u.MF002, CHARINDEX('-', u.MF002) - 1) ELSE u.MF002 END)) CreatorName
                                    , STUFF(STUFF(a.CREATE_DATE + ' ' + a.CREATE_TIME, 5, 0, '-'), 8, 0, '-') CreateDate
                                    , STUFF(STUFF(a.MODI_DATE + ' ' + a.MODI_TIME, 5, 0, '-'), 8, 0, '-') LastModifiedDate
                                    , LTRIM(RTRIM(a.CREATOR)) CreateBy, LTRIM(RTRIM(a.MODIFIER)) LastModifiedBy
                                    , @NewCompanyNo CompanyNo";
                                sql += @" FROM PURTA a
                                        LEFT JOIN ADMMF u1 ON LTRIM(RTRIM(u1.MF001)) = LTRIM(RTRIM(a.TA012))
                                        LEFT JOIN CMSML c ON c.COMPANY = a.COMPANY
                                        LEFT JOIN ADMMF u ON LTRIM(RTRIM(u.MF001)) = LTRIM(RTRIM(a.CREATOR))
                                        WHERE LTRIM(RTRIM(a.TA001)) + LTRIM(RTRIM(a.TA002)) IN @KeyNos";

                                const int batchSize = 2000;
                                for (int i = 0; i < PagedNos.Count; i += batchSize)
                                {
                                    result.AddRange(erpConnection.Query<dynamic>(sql, new { x.NewCompanyNo, KeyNos = KeyNos.Skip(i).Take(batchSize).ToArray() }).ToList());
                                }
                                #endregion
                            }
                        }
                    });

                    ApplyDynamicOrdering(ref result, OrderBy, "PrErpNo DESC, PrErpPrefix");

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        draw,
                        recordsTotal = Nos.Count,
                        recordsFiltered = Nos.Count,
                        data = result,
                        status = "success"
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

        #region //GetPrDetailERP -- 取得ERP請購單身資料 -- Chia Yuan 2025.3.26
        public string GetPrDetailERP(List<SupplierCompany> SupplierCompany
            , string SupplierNo, string SupplierNos, string MtlItemNo, string InventoryNos, string PoCurrency
            , string PrErpPrefix, string PrErpNo, string PrSeq
            , string PrErpPrefixs, string PrErpNos, string PrSeqs
            , string PrUser, string PoUsers, string DepartmentNo, string LockStatus
            , string PrStartDate, string PrEndDate
            , string DemandStartDate, string DemandEndDate
            , string DeliveryStartDate, string DeliveryEndDate
            , string ConfirmStatus, string ClosureStatus, bool? DocTransfer, string SearchKey
            , bool IsOptions, int draw
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    List<(string, string)> Nos = new List<(string, string)>();

                    SupplierCompany.Distinct().ToList().ForEach(x =>
                    {
                        #region //確認公司別DB
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        var resultCorp = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo = x.CompanyNo == "ETG" ? "Eterge" : x.CompanyNo }) ?? throw new SystemException("【公司別】資料錯誤!");
                        ErpConnectionStrings = ConfigurationManager.AppSettings[resultCorp.ErpDb];
                        #endregion

                        using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //條件資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(a.TB001)) PrErpPrefix, LTRIM(RTRIM(a.TB002)) PrErpNo, LTRIM(RTRIM(a.TB003)) PrSeq
                                    FROM PURTB a
                                    INNER JOIN PURTA b ON LTRIM(RTRIM(b.TA001)) = a.TB001 AND LTRIM(RTRIM(b.TA002)) = a.TB002
                                    WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TB001", @" AND LTRIM(RTRIM(a.TB001)) = @TB001", PrErpPrefix);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TB002", @" AND LTRIM(RTRIM(a.TB002)) = @TB002", PrErpNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TB003", @" AND LTRIM(RTRIM(a.TB003)) = @TB003", PrSeq);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemNo", @" AND LTRIM(RTRIM(a.TB004)) = @MtlItemNo", MtlItemNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TA012", @" AND LTRIM(RTRIM(b.TA012)) = @TA012", PrUser);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TA004", @" AND LTRIM(RTRIM(b.TA004)) = @TA004", DepartmentNo);
                            if (DocTransfer != null)
                            {
                                if (DocTransfer.Value)
                                    sql += @" AND EXISTS (
                                        SELECT TOP 1 1 
                                        FROM PURTD ab 
                                        WHERE LTRIM(RTRIM(ab.TD026)) = LTRIM(RTRIM(a.TB001)) 
                                        AND LTRIM(RTRIM(ab.TD027)) = LTRIM(RTRIM(a.TB002)) 
                                        AND LTRIM(RTRIM(ab.TD028)) = LTRIM(RTRIM(a.TB003)))";
                                else
                                    sql += @" AND NOT EXISTS (
                                        SELECT TOP 1 1 
                                        FROM PURTD ab 
                                        WHERE LTRIM(RTRIM(ab.TD026)) = LTRIM(RTRIM(a.TB001)) 
                                        AND LTRIM(RTRIM(ab.TD027)) = LTRIM(RTRIM(a.TB002)) 
                                        AND LTRIM(RTRIM(ab.TD028)) = LTRIM(RTRIM(a.TB003)))
	                                    AND a.TB025 = 'Y'
	                                    AND b.TA007 = 'Y'";
                            }
                            if (PrErpPrefixs.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TB001s", @" AND LTRIM(RTRIM(a.TB001)) IN @TB001s", PrErpPrefixs.Split(','));
                            if (PrErpNos.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TB002s", @" AND LTRIM(RTRIM(a.TB002)) IN @TB002s", PrErpNos.Split(','));
                            if (PrSeqs.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TB003s", @" AND LTRIM(RTRIM(a.TB003)) IN @TB003s", PrSeqs.Split(','));
                            if (ConfirmStatus.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TB025", @" AND a.TB025 IN @TB025", ConfirmStatus.Split(','));
                            if (Regex.IsMatch(ClosureStatus, "^(Y|N)$", RegexOptions.IgnoreCase))
                            {
                                if (ClosureStatus == "N") sql += @" AND b.TA007 <> 'V' AND a.TB039 = 'N' AND a.TB025 <> 'V'";
                                if (ClosureStatus == "Y") sql += @" AND b.TA007 = 'Y' AND a.TB039 <> 'N'";
                            }
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TB010", @" AND LTRIM(RTRIM(a.TB010)) = @TB010", SupplierNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TB016", @" AND TB016 = @TB016", PoCurrency);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PrStartDate", @" AND b.TA003 >= @PrStartDate",
                                PrStartDate.Length > 0 ? Convert.ToDateTime(PrStartDate).ToString("yyyyMMdd") : string.Empty);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PrEndDate", @" AND b.TA003<= @PrEndDate",
                                PrEndDate.Length > 0 ? Convert.ToDateTime(PrEndDate).AddDays(1).ToString("yyyyMMdd") : string.Empty);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "DemandStartDate", @" AND a.TB011 >= @DemandStartDate",
                                DemandStartDate.Length > 0 ? Convert.ToDateTime(DemandStartDate).ToString("yyyyMMdd") : string.Empty);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "DemandEndDate", @" AND a.TB011 <= @DemandEndDate",
                                DemandEndDate.Length > 0 ? Convert.ToDateTime(DemandEndDate).AddDays(1).ToString("yyyyMMdd") : string.Empty);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "DeliveryStartDate", @" AND a.TB019 >= @DeliveryStartDate",
                                DeliveryStartDate.Length > 0 ? Convert.ToDateTime(DeliveryStartDate).ToString("yyyyMMdd") : string.Empty);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "DeliveryEndDate", @" AND a.TB019 <= @DeliveryEndDate",
                                DeliveryEndDate.Length > 0 ? Convert.ToDateTime(DeliveryEndDate).AddDays(1).ToString("yyyyMMdd") : string.Empty);
                            if (InventoryNos.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TB008s", @" AND LTRIM(RTRIM(TB008)) IN @TB008s", InventoryNos.Split(','));
                            if (SupplierNos.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TB010s", @" AND LTRIM(RTRIM(a.TB010)) IN @TB010s", SupplierNos.Split(','));
                            if (PoUsers.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TB013s", @" AND LTRIM(RTRIM(a.TB013)) IN @TB013s", PoUsers.Split(','));
                            if (LockStatus.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TB020", @" AND TB020 IN @TB020", LockStatus.Split(','));
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SearchKey",
                                @" AND (LTRIM(RTRIM(a.TB001)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TB002)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TB001)) + '-' + LTRIM(RTRIM(a.TB002)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TB001)) + '-' + LTRIM(RTRIM(a.TB002)) + '-' + LTRIM(RTRIM(a.TB003)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TB004)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TB005)) LIKE N'%' + @SearchKey + '%'
                                    OR EXISTS (SELECT TOP 1 1 FROM INVMB aa WHERE LTRIM(RTRIM(aa.MB001)) = LTRIM(RTRIM(a.TB004)) AND aa.MB002 LIKE N'%' + @SearchKey + '%'))", SearchKey);
                            sql += OrderBy.Length > 0 ? " ORDER BY " + OrderBy : " ORDER BY PrErpNo DESC, PrSeq DESC, PrErpPrefix";

                            Nos.AddRange(erpConnection.Query(sql, dynamicParameters)
                                .Select(s => ((string)resultCorp.ErpNo, string.Format("{0}{1}{2}", (string)s.PrErpPrefix, (string)s.PrErpNo, (string)s.PrSeq)))
                                .ToList()); //取得識別單號
                            #endregion
                        }
                    });

                    List<(string, string)> PagedNos = new List<(string, string)>();
                    if (PageSize > 0) PagedNos = Nos.Skip(Math.Max(PageIndex, 0) * PageSize).Take(PageSize).ToList(); //符合條件的品號(與庫存條件交集)
                    else PagedNos = Nos;

                    List<dynamic> result = new List<dynamic>();

                    SupplierCompany.Distinct().ToList().ForEach(x =>
                    {
                        #region //確認公司別DB
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        var resultCorp = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo = x.CompanyNo == "ETG" ? "Eterge" : x.CompanyNo }) ?? throw new SystemException("【公司別】資料錯誤!");
                        ErpConnectionStrings = ConfigurationManager.AppSettings[resultCorp.ErpDb];
                        #endregion

                        string[] KeyNos = PagedNos.Where(w => w.Item1 == resultCorp.ErpNo).Select(s => s.Item2).ToArray();

                        if (KeyNos.Any())
                        {
                            using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                            {
                                #region //取得資料
                                sql = IsOptions ?
                                    @"SELECT LTRIM(RTRIM(a.TB001)) PrErpPrefix, LTRIM(RTRIM(a.TB002)) PrErpNo, LTRIM(RTRIM(a.TB003)) PrSeq
                                    , LTRIM(RTRIM(a.TB001)) + '-' + LTRIM(RTRIM(a.TB002)) + '-' + LTRIM(RTRIM(a.TB003)) PrErpPrefixNo
                                    , LTRIM(RTRIM(a.TB004)) + ' ' + LTRIM(RTRIM(a.TB005)) MtlItemWithNo
                                    , LTRIM(RTRIM(a.TB004)) MtlItemNo, LTRIM(RTRIM(a.TB005)) MtlItemName
                                    , @NewCompanyNo CompanyNo" :
                                    @"SELECT LTRIM(RTRIM(a.TB001)) PrErpPrefix, LTRIM(RTRIM(a.TB002)) PrErpNo, LTRIM(RTRIM(a.TB003)) PrSeq
                                    , LTRIM(RTRIM(a.TB001)) + '-' + LTRIM(RTRIM(a.TB002)) + '-' + LTRIM(RTRIM(a.TB003)) PrErpPrefixNo
                                    , LTRIM(RTRIM(a.TB004)) + ' ' + LTRIM(RTRIM(a.TB005)) MtlItemWithNo
                                    , LTRIM(RTRIM(a.TB004)) MtlItemNo, LTRIM(RTRIM(a.TB005)) MtlItemName, LTRIM(RTRIM(a.TB006)) MtlItemSpec
                                    , a.TB007 PrUnit, a.TB008 Inventory, a.TB009 PrQty, a.TB010 SupplierNo
                                    , STUFF(STUFF(a.TB011, 5, 0, '-'), 8, 0, '-') DemandDate
                                    , a.TB012 PrRemark, a.TB013 PoUser, u3.MF002 PoUserName, a.TB014 PoQty
                                    , a.TB015 PoUnit, a.TB016 PoCurrency, a.TB017 PoUnitPrice, a.TB018 PoPrice
                                    , STUFF(STUFF(a.TB019, 5, 0, '-'), 8, 0, '-') DeliveryDate
                                    , a.TB020 LockStaus, a.TB021 PoStatus, a.TB022 PoErpPrefixNo
                                    , a.TB024 PoRemark, a.TB025 ConfirmStatus, a.TB026 Taxation, a.TB032 UrgentMtl, a.TB033 PrPackageQty
                                    , a.TB034 PoPackageQty, a.TB039 ClosureStatus, a.TB040 InquiryStatus, a.TB042 PoType, a.TB043 ProjectCode
                                    , a.TB044 PrExchangeRate, a.TB045 PrPriceLocal, a.TB046 PoPartial, a.TB047 BudgetNo
                                    , a.TB048 BudgetSubject, a.TB049 PrUnitPrice, a.TB050 PrCurrency, a.TB051 PrPrice
                                    , a.TB057 TaxCode, a.TB058 TradeTerm, a.TB063 TaxRate, a.TB064 DetailMultiTax
                                    , a.TB065 PrPriceQty, a.TB066 PrPriceUnit, a.TB067 DiscountRate, a.TB068 DiscountAmount
                                    , i.MC002 InventoryName, LTRIM(RTRIM(t.NN002)) TaxName, t.NN005 InvoiceType, LTRIM(RTRIM(m.MB004)) Unit, m.MB022 LotManagement, m.MB043 MeasureType
                                    , STUFF(STUFF(b.TA003, 5, 0, '-'), 8, 0, '-') PrDate
                                    , STUFF(STUFF(b.TA013, 5, 0, '-'), 8, 0, '-') DocDate
                                    , LTRIM(RTRIM(b.TA004)) PrDepartmentNo, LTRIM(RTRIM(h.ME002))  PrDepartmentName
                                    , b.TA012 PrUser, u2.MF002 PrUserName, d.MA002 SupplierName
                                    , LTRIM(RTRIM(c.ML002)) CompanyShortName, LTRIM(RTRIM(c.ML003)) CompanyName
                                    , LTRIM(RTRIM(CASE WHEN CHARINDEX('-', u.MF002) > 0 THEN LEFT(u.MF002, CHARINDEX('-', u.MF002) - 1) ELSE u.MF002 END)) CreatorName
                                    , STUFF(STUFF(a.CREATE_DATE + ' ' + a.CREATE_TIME, 5, 0, '-'), 8, 0, '-') CreateDate
                                    , STUFF(STUFF(a.MODI_DATE + ' ' + a.MODI_TIME, 5, 0, '-'), 8, 0, '-') LastModifiedDate
                                    , LTRIM(RTRIM(a.CREATOR)) CreateBy, LTRIM(RTRIM(a.MODIFIER)) LastModifiedBy
                                    , @NewCompanyNo CompanyNo";
                                sql += @" FROM PURTB a
                                        INNER JOIN PURTA b ON a.TB001 = b.TA001 AND a.TB002 = b.TA002
                                        LEFT JOIN CMSML c ON c.COMPANY = a.COMPANY
                                        LEFT JOIN PURMA d ON d.MA001 = a.TB010
                                        INNER JOIN CMSMC i ON i.MC001 = a.TB008
                                        INNER JOIN INVMB m ON LTRIM(RTRIM(m.MB001)) = LTRIM(RTRIM(a.TB004))
                                        LEFT JOIN CMSME h ON b.TA004 = h.ME001
                                        LEFT JOIN CMSNN t ON t.NN001 = a.TB057
                                        LEFT JOIN ADMMF u ON LTRIM(RTRIM(u.MF001)) = LTRIM(RTRIM(a.CREATOR))
                                        LEFT JOIN ADMMF u2 ON u2.MF001 = b.TA012
                                        LEFT JOIN ADMMF u3 ON u3.MF001 = a.TB013
                                        WHERE LTRIM(RTRIM(a.TB001)) + LTRIM(RTRIM(a.TB002)) + LTRIM(RTRIM(a.TB003)) IN @KeyNos";

                                const int batchSize = 2000;
                                for (int i = 0; i < PagedNos.Count; i += batchSize)
                                {
                                    result.AddRange(erpConnection.Query<dynamic>(sql, new { x.NewCompanyNo, KeyNos = KeyNos.Skip(i).Take(batchSize).ToArray() }).ToList());
                                }

                                ApplyDynamicOrdering(ref result, OrderBy, "PrErpNo DESC, PrSeq DESC, PrErpPrefix");
                                #endregion
                            }
                        }
                    });

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        draw,
                        recordsTotal = Nos.Count,
                        recordsFiltered = Nos.Count,
                        data = result,
                        status = "success"
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

        #region //GetPurchaseOrderERP -- 取得ERP採購單頭資料 -- Chia Yuan 2025.01.17
        public string GetPurchaseOrderERP(List<SupplierCompany> SupplierCompany
            , string SupplierNo, string SupplierNos
            , string PoErpPrefix, string PoErpNo, string PoErpFullNo
            , string PoErpPrefixs, string PoErpNos, string PoErpFullNos
            , string PoUser, string MtlItemNo, string MtlItemNos, string Currency
            , string PoStartDate, string PoEndDate
            , string DocStartDate, string DocEndDate
            , string ConfirmStatus, string ClosureStatus, string SearchKey
            , bool IsOptions, int draw
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    List<(string, string)> Nos = new List<(string, string)>();

                    SupplierCompany.Distinct().ToList().ForEach(x =>
                    {
                        #region //確認公司別DB
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        var resultCorp = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo = x.CompanyNo == "ETG" ? "Eterge" : x.CompanyNo }) ?? throw new SystemException("【公司別】資料錯誤!");
                        ErpConnectionStrings = ConfigurationManager.AppSettings[resultCorp.ErpDb];
                        #endregion

                        using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //條件資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(a.TC001)) PoErpPrefix, LTRIM(RTRIM(a.TC002)) PoErpNo
                                    FROM PURTC a
                                    WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TC001", @" AND LTRIM(RTRIM(a.TC001)) = @TC001", PoErpPrefix);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TC002", @" AND LTRIM(RTRIM(a.TC002)) = @TC002", PoErpNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PoErpFullNo", @" AND LTRIM(RTRIM(a.TC001)) + '-' + LTRIM(RTRIM(a.TC002)) = @PoErpFullNo", PoErpFullNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TC004", @" AND a.TC004 = @TC004", SupplierNo); //支持接換查詢
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TC005", @" AND a.TC005 = @TC005", Currency);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TC011", @" AND a.TC011 = @TC011", PoUser);
                            if (PoErpFullNos.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PoErpFullNos", @" AND LTRIM(RTRIM(a.TC001)) + '-' + LTRIM(RTRIM(a.TC002)) IN @PoErpFullNos", PoErpFullNos.Split(','));
                            if (PoErpPrefixs.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TC001s", @" AND LTRIM(RTRIM(a.TC001)) IN @TC001s", PoErpPrefixs.Split(','));
                            if (PoErpNos.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TC002s", @" AND LTRIM(RTRIM(a.TC002)) IN @TC002s", PoErpNos.Split(','));
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PoStartDate", @" AND a.TC003 >= @PoStartDate",
                                PoStartDate.Length > 0 ? Convert.ToDateTime(PoStartDate).ToString("yyyyMMdd") : string.Empty);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PoEndDate", @" AND a.TC003 <= @PoEndDate",
                                PoEndDate.Length > 0 ? Convert.ToDateTime(PoEndDate).AddDays(1).ToString("yyyyMMdd") : string.Empty);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "DocStartDate", @" AND a.TC024 >= @DocStartDate",
                                DocStartDate.Length > 0 ? Convert.ToDateTime(DocStartDate).ToString("yyyyMMdd") : string.Empty);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "DocEndDate", @" AND a.TC024 <= @DocEndDate",
                                DocEndDate.Length > 0 ? Convert.ToDateTime(DocEndDate).AddDays(1).ToString("yyyyMMdd") : string.Empty);
                            if (MtlItemNo.Length > 0 || MtlItemNos.Length > 0)
                            { 
                                sql += @" AND EXISTS (SELECT TOP 1 1 FROM PURTD aa WHERE LTRIM(RTRIM(aa.TD001)) = LTRIM(RTRIM(a.TC001)) AND LTRIM(RTRIM(aa.TD002)) = LTRIM(RTRIM(a.TC002))";
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TD004", @" AND LTRIM(RTRIM(aa.TD004)) = @TD004", MtlItemNo);
                                if (MtlItemNos.Length > 0)
                                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TD004s", @" AND LTRIM(RTRIM(aa.TD004)) IN @TD004s", MtlItemNos.Split(',').Distinct());
                                sql += ")";
                            }
                            if (SupplierNos.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TC004s", @" AND LTRIM(RTRIM(a.TC004)) IN @TC004s", SupplierNos.Split(','));
                            if (ConfirmStatus.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TC014", @" AND a.TC014 IN @TC014", ConfirmStatus.Split(','));
                            if (Regex.IsMatch(ClosureStatus, "^(Y|N)$", RegexOptions.IgnoreCase))
                            {
                                if (ClosureStatus == "N")
                                {
                                    sql += @" AND a.TC014 != 'V'
                                        AND EXISTS (
                                            SELECT TOP 1 1
                                            FROM PURTD aa 
                                            WHERE LTRIM(RTRIM(aa.TD001)) = LTRIM(RTRIM(a.TC001)) AND LTRIM(RTRIM(aa.TD002)) = LTRIM(RTRIM(a.TC002))
                                            AND aa.TD016 = 'N' AND aa.TD018 != 'V')";
                                }
                                if (ClosureStatus == "Y")
                                {
                                    sql += @" AND a.TC014 = 'Y'
                                        AND EXISTS (
                                            SELECT TOP 1 1
                                            FROM PURTD aa 
                                            WHERE LTRIM(RTRIM(aa.TD001)) = LTRIM(RTRIM(a.TC001)) AND LTRIM(RTRIM(aa.TD002)) = LTRIM(RTRIM(a.TC002))
                                            AND aa.TD016 != 'N')";
                                }
                            }
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SearchKey",
                                @" AND (LTRIM(RTRIM(a.TC001)) + '-' + LTRIM(RTRIM(a.TC002)) LIKE N'%' + @SearchKey + '%'
                                OR a.TC001 LIKE N'%' + @SearchKey + '%' 
                                OR a.TC002 LIKE N'%' + @SearchKey + '%' 
                                OR a.TC004 LIKE N'%' + @SearchKey + '%'
                                OR EXISTS (
                                    SELECT TOP 1 1 FROM PURTD aa WHERE aa.TD001 = a.TC001 AND aa.TD002 = a.TC002
                                    AND (aa.TD004 LIKE N'%' + @SearchKey + '%'
                                        OR aa.TD005 LIKE N'%' + @SearchKey + '%')))", SearchKey);
                            sql += OrderBy.Length > 0 ? "  ORDER BY " + OrderBy : " ORDER BY PoErpNo DESC, PoErpPrefix";

                            Nos.AddRange(erpConnection.Query(sql, dynamicParameters)
                                .Select(s => ((string)resultCorp.ErpNo, string.Format("{0}{1}", (string)s.PoErpPrefix, (string)s.PoErpNo)))
                                .ToList()); //取得識別單號
                            #endregion
                        }
                    });

                    List<(string, string)> PagedNos = new List<(string, string)>();
                    if (PageSize > 0) PagedNos = Nos.Skip(Math.Max(PageIndex, 0) * PageSize).Take(PageSize).ToList(); //符合條件的品號(與庫存條件交集)
                    else PagedNos = Nos;

                    List<dynamic> result = new List<dynamic>();

                    SupplierCompany.Distinct().ToList().ForEach(x =>
                    {
                        #region //確認公司別DB
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        var resultCorp = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo = x.CompanyNo == "ETG" ? "Eterge" : x.CompanyNo }) ?? throw new SystemException("【公司別】資料錯誤!");
                        ErpConnectionStrings = ConfigurationManager.AppSettings[resultCorp.ErpDb];
                        #endregion

                        string[] KeyNos = PagedNos.Where(w => w.Item1 == resultCorp.ErpNo).Select(s => s.Item2).ToArray();

                        if (KeyNos.Any())
                        {
                            using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                            {
                                #region //取得資料
                                sql = IsOptions ?
                                    @"SELECT LTRIM(RTRIM(a.TC001)) PoErpPrefix, LTRIM(RTRIM(a.TC002)) PoErpNo, LTRIM(RTRIM(a.TC001)) + '-' + LTRIM(RTRIM(a.TC002)) PoErpFullNo
                                    , LTRIM(RTRIM(a.TC004)) SupplierNo, @NewCompanyNo CompanyNo" :
                                    @"SELECT LTRIM(RTRIM(a.TC001)) PoErpPrefix, LTRIM(RTRIM(a.TC002)) PoErpNo, LTRIM(RTRIM(a.TC001)) + '-' + LTRIM(RTRIM(a.TC002)) PoErpFullNo
                                    , LTRIM(RTRIM(a.TC004)) SupplierNo
                                    , STUFF(STUFF(a.TC003, 5, 0, '-'), 8, 0, '-') PoDate
                                    , a.TC005 Currency, f.MF002 CurrencyName, a.TC006 PoExchangeRate, a.TC008 PaymentTerm
                                    , a.TC009 Remark, a.TC010 Factory, LTRIM(RTRIM(a.TC011)) PoUser, LTRIM(RTRIM(u1.MF002)) PoUserName, a.TC012 PrintFormat
                                    , a.TC013 PrintCount, a.TC014 ConfirmStatus
                                    , STUFF(STUFF(a.TC015, 5, 0, '-'), 8, 0, '-') PiDate
                                    , a.TC016 PiNo, a.TC017 ShipMethodName, a.TC018 Taxation
                                    , a.TC019 TotalAmount, a.TC020 TaxAmount, a.TC021 FirstAddress, a.TC022 SecondAddress, a.TC023 Quantity
                                    , STUFF(STUFF(a.TC024, 5, 0, '-'), 8, 0, '-') DocDate
                                    , LTRIM(RTRIM(a.TC025)) ConfirmUser, LTRIM(RTRIM(u2.MF002)) ConfirmUserName, a.TC026 TaxRate
                                    , LTRIM(RTRIM(a.TC027)) PaymentTermCode, a.TC028 DepositRate, a.TC029 PackageQty, a.TC030 ApproveStatus
                                    , a.TC031 TransmissionCount, a.TC032 ProcessCode, a.TC033 TransferStatusERP, a.TC039 Edition
                                    , a.TC040 DepositPartial, a.TC041 ShipMethod, a.TC047 TaxCode, a.TC048 TradeTerm
                                    , a.TC049 Manufacturer, a.TC050 LockStatus, a.TC051 DetailMultiTax, a.TC052 ContactUser
                                    , LTRIM(RTRIM(p.MA002)) SupplierName, p.MA024 PaymentType
                                    , LTRIM(RTRIM(t.NN002)) TaxName, t.NN005 InvoiceType
                                    , LTRIM(RTRIM(c.ML002)) CompanyShortName, LTRIM(RTRIM(c.ML003)) CompanyName
                                    , LTRIM(RTRIM(CASE WHEN CHARINDEX('-', u.MF002) > 0 THEN LEFT(u.MF002, CHARINDEX('-', u.MF002) - 1) ELSE u.MF002 END)) CreatorName
                                    , STUFF(STUFF(a.CREATE_DATE + ' ' + a.CREATE_TIME, 5, 0, '-'), 8, 0, '-') CreateDate
                                    , STUFF(STUFF(a.MODI_DATE + ' ' + a.MODI_TIME, 5, 0, '-'), 8, 0, '-') LastModifiedDate
                                    , LTRIM(RTRIM(a.CREATOR)) CreateBy, LTRIM(RTRIM(a.MODIFIER)) LastModifiedBy
                                    , @NewCompanyNo CompanyNo";
                                sql += @" FROM PURTC a
                                        LEFT JOIN CMSML c ON c.COMPANY = a.COMPANY
                                        INNER JOIN CMSMF f ON f.MF001 = a.TC005
                                        INNER JOIN PURMA p ON p.MA001 = a.TC004
                                        INNER JOIN CMSNN t ON t.NN001 = a.TC047
                                        LEFT JOIN ADMMF u ON LTRIM(RTRIM(u.MF001)) = LTRIM(RTRIM(a.CREATOR))
                                        LEFT JOIN ADMMF u1 ON LTRIM(RTRIM(u1.MF001)) = LTRIM(RTRIM(a.TC011))
                                        LEFT JOIN ADMMF u2 ON LTRIM(RTRIM(u2.MF001)) = LTRIM(RTRIM(a.TC025))
                                        WHERE LTRIM(RTRIM(a.TC001)) + LTRIM(RTRIM(a.TC002)) IN @KeyNos";
                                //sql += OrderBy.Length > 0 ? "  ORDER BY " + OrderBy : " ORDER BY PoErpNo DESC, PoErpPrefix";

                                const int batchSize = 2000;
                                for (int i = 0; i < PagedNos.Count; i += batchSize)
                                {
                                    result.AddRange(erpConnection.Query<dynamic>(sql, new { x.NewCompanyNo, KeyNos = KeyNos.Skip(i).Take(batchSize).ToArray() }).ToList());
                                }

                                ApplyDynamicOrdering(ref result, OrderBy, "PoErpNo DESC, PoErpPrefix");
                                #endregion
                            }
                        }
                    });

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        draw,
                        recordsTotal = Nos.Count,
                        recordsFiltered = Nos.Count,
                        data = result,
                        status = "success"
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

        #region //GetPoDetailERP --取得採購單身資料 -- Chia Yuan 2024-08-13
        public string GetPoDetailERP(List<SupplierCompany> SupplierCompany
            , string SupplierNo, string SupplierNos
            , string PoErpPrefix, string PoErpNo, string PoSeq
            , string PoErpPrefixs, string PoErpNos, string PoSeqs
            , string PoErpFullNo, string PoErpFullNos
            , string PoErpPrefixNo, string PoErpPrefixNos
            , string PromiseStartDate, string PromiseEndDate
            , string OriPromiseStartDate, string OriPromiseEndDate
            , string DeliveryStartDate, string DeliveryEndDate
            , string ConfirmStatus, string ClosureStatus, string SearchKey
            , bool IsOptions, int draw
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    List<(string, string)> Nos = new List<(string, string)>();

                    SupplierCompany.Distinct().ToList().ForEach(x =>
                    {
                        #region //確認公司別DB
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        var resultCorp = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo = x.CompanyNo == "ETG" ? "Eterge" : x.CompanyNo }) ?? throw new SystemException("【公司別】資料錯誤!");
                        ErpConnectionStrings = ConfigurationManager.AppSettings[resultCorp.ErpDb];
                        #endregion

                        using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //條件資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(a.TD001)) PoErpPrefix, LTRIM(RTRIM(a.TD002)) PoErpNo, LTRIM(RTRIM(a.TD003)) PoSeq
                                    FROM PURTD a
                                    INNER JOIN PURTC b ON LTRIM(RTRIM(b.TC001)) = a.TD001 AND LTRIM(RTRIM(b.TC002)) = a.TD002
                                    WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PoErpPrefix", @" AND LTRIM(RTRIM(a.TD001)) = @PoErpPrefix", PoErpPrefix);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PoErpNo", @" AND LTRIM(RTRIM(a.TD002)) = @PoErpNo", PoErpNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PoSeq", @" AND LTRIM(RTRIM(a.TD003)) = @PoSeq", PoSeq);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PoErpPrefixNo", @" AND LTRIM(RTRIM(a.TD001)) + '-' + LTRIM(RTRIM(a.TD002)) + '-' + LTRIM(RTRIM(a.TD003)) = @PoErpPrefixNo", PoErpPrefixNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PoErpFullNo", @" AND LTRIM(RTRIM(a.TD001)) + '-' + LTRIM(RTRIM(a.TD002)) = @PoErpFullNo", PoErpFullNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TC004", @" AND LTRIM(RTRIM(b.TC004)) = @TC004", SupplierNo);
                            if (PoErpPrefixs.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TD001s", @" AND LTRIM(RTRIM(a.TD001)) IN @TD001s", PoErpPrefixs.Split(','));
                            if (PoErpNos.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TD002s", @" AND LTRIM(RTRIM(a.TD002)) IN @TD002s", PoErpNos.Split(','));
                            if (PoSeqs.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TD003s", @" AND LTRIM(RTRIM(a.TD003)) IN @TD003s", PoSeqs.Split(','));
                            if (PoErpFullNos.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PoErpFullNos", @" AND LTRIM(RTRIM(a.TD001)) + '-' + LTRIM(RTRIM(a.TD002)) IN @PoErpFullNos", PoErpFullNos.Split(','));
                            if (PoErpPrefixNos.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PoErpPrefixNos", @" AND LTRIM(RTRIM(a.TD001)) + '-' + LTRIM(RTRIM(a.TD002)) + '-' + LTRIM(RTRIM(a.TD003)) IN @PoErpPrefixNos", PoErpPrefixNos.Split(','));
                            if (ConfirmStatus.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TD018", @" AND a.TD018 IN @TD018", ConfirmStatus.Split(','));
                            if (Regex.IsMatch(ClosureStatus, "^(Y|N)$", RegexOptions.IgnoreCase))
                            {
                                if (ClosureStatus == "N")
                                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ClosureStatus", @" AND b.TC014 != 'V' AND a.TD016 = 'N' AND a.TD018 != 'V'", ClosureStatus);

                                if (ClosureStatus == "Y")
                                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ClosureStatus", @" AND b.TC014 = 'Y' AND a.TD016 != 'N'", ClosureStatus);
                            }
                            if (SupplierNos.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TC004s", @" AND b.TC004 IN @TC004s", SupplierNos.Split(','));
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PromiseStartDate", @" AND a.TD012 >= @PromiseStartDate",
                                PromiseStartDate.Length > 0 ? Convert.ToDateTime(PromiseStartDate).ToString("yyyyMMdd") : string.Empty);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PromiseEndDate", @" AND a.TD012 <= @PromiseEndDate",
                                PromiseEndDate.Length > 0 ? Convert.ToDateTime(PromiseEndDate).AddDays(1).ToString("yyyyMMdd") : string.Empty);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "OriPromiseStartDate", @" AND a.TD045 >= @OriPromiseStartDate",
                                OriPromiseStartDate.Length > 0 ? Convert.ToDateTime(OriPromiseStartDate).ToString("yyyyMMdd") : string.Empty);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "OriPromiseEndDate", @" AND a.TD045 <= @OriPromiseEndDate",
                                OriPromiseEndDate.Length > 0 ? Convert.ToDateTime(OriPromiseEndDate).AddDays(1).ToString("yyyyMMdd") : string.Empty);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "DeliveryStartDate", @" AND a.TD046 >= @DeliveryStartDate",
                                DeliveryStartDate.Length > 0 ? Convert.ToDateTime(DeliveryStartDate).ToString("yyyyMMdd") : string.Empty);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "DeliveryEndDate", @" AND a.TD046 <= @DeliveryEndDate",
                                DeliveryEndDate.Length > 0 ? Convert.ToDateTime(DeliveryEndDate).AddDays(1).ToString("yyyyMMdd") : string.Empty);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SearchKey",
                                @" AND (LTRIM(RTRIM(a.TD001)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TD002)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TD001)) + '-' + LTRIM(RTRIM(a.TD002)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TD001)) + '-' + LTRIM(RTRIM(a.TD002)) + '-' + LTRIM(RTRIM(a.TD003)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TD004)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TD005)) LIKE N'%' + @SearchKey + '%')", SearchKey);
                            sql += OrderBy.Length > 0 ? " ORDER BY " + OrderBy : " ORDER BY PoErpNo DESC, PoSeq DESC, PoErpPrefix";

                            Nos.AddRange(erpConnection.Query(sql, dynamicParameters)
                                .Select(s => ((string)resultCorp.ErpNo, string.Format("{0}{1}{2}", (string)s.PoErpPrefix, (string)s.PoErpNo, (string)s.PoSeq)))
                                .ToList()); //取得識別單號
                            #endregion
                        }
                    });

                    List<(string, string)> PagedNos = new List<(string, string)>();
                    if (PageSize > 0) PagedNos = Nos.Skip(Math.Max(PageIndex, 0) * PageSize).Take(PageSize).ToList(); //符合條件的品號(與庫存條件交集)
                    else PagedNos = Nos;

                    List<dynamic> result = new List<dynamic>();

                    SupplierCompany.Distinct().ToList().ForEach(x => 
                    {
                        #region //確認公司別DB
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        var resultCorp = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo = x.CompanyNo == "ETG" ? "Eterge" : x.CompanyNo }) ?? throw new SystemException("【公司別】資料錯誤!");
                        ErpConnectionStrings = ConfigurationManager.AppSettings[resultCorp.ErpDb];
                        #endregion

                        string[] KeyNos = PagedNos.Where(w => w.Item1 == resultCorp.ErpNo).Select(s => s.Item2).ToArray();

                        using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //取得資料
                            sql = IsOptions ?
                                @"SELECT LTRIM(RTRIM(a.TD001)) PoErpPrefix, LTRIM(RTRIM(a.TD002)) PoErpNo, LTRIM(RTRIM(a.TD003)) PoSeq
                                , LTRIM(RTRIM(a.TD001)) + '-' + LTRIM(RTRIM(a.TD002)) PoErpFullNo
                                , LTRIM(RTRIM(a.TD001)) + '-' + LTRIM(RTRIM(a.TD002)) + '-' + LTRIM(RTRIM(a.TD003)) PoErpPrefixNo
                                , LTRIM(RTRIM(a.TD004)) + ' ' + LTRIM(RTRIM(a.TD005)) MtlItemWithNo
                                , LTRIM(RTRIM(a.TD004)) MtlItemNo, LTRIM(RTRIM(a.TD005)) MtlItemName
                                , LTRIM(RTRIM(b.TC004)) SupplierNo, @NewCompanyNo CompanyNo" :
                                @"SELECT LTRIM(RTRIM(a.TD001)) PoErpPrefix, LTRIM(RTRIM(a.TD002)) PoErpNo, LTRIM(RTRIM(a.TD003)) PoSeq
                                , LTRIM(RTRIM(a.TD001)) + '-' + LTRIM(RTRIM(a.TD002)) PoErpFullNo
                                , LTRIM(RTRIM(a.TD001)) + '-' + LTRIM(RTRIM(a.TD002)) + '-' + LTRIM(RTRIM(a.TD003)) PoErpPrefixNo
                                , LTRIM(RTRIM(a.TD004)) + ' ' + LTRIM(RTRIM(a.TD005)) MtlItemWithNo
                                , LTRIM(RTRIM(a.TD004)) MtlItemNo, LTRIM(RTRIM(a.TD005)) MtlItemName, LTRIM(RTRIM(a.TD006)) MtlItemSpec
                                , a.TD007 Inventory, a.TD008 PoQty, a.TD009 PoUnit, a.TD010 PoUnitPrice, a.TD011 PoPrice
                                , STUFF(STUFF(a.TD012, 5, 0, '-'), 8, 0, '-') PromiseDate
                                , a.TD013 ReferencePrefix, a.TD014 Remark, a.TD015 SiQty, a.TD016 ClosureStatus, a.TD017 Manufacturer, a.TD018 ConfirmStatus
                                , a.TD019 InventoryQty, a.TD020 SmallUnit, a.TD021 ReferenceNo, a.TD022 ProjectCode, a.TD023 ReferenceSeq, a.TD024 FormErpNo
                                , a.TD025 UrgentMtl, a.TD026 PrErpPrefix, a.TD027 PrErpNo, a.TD028 PrSeq, a.TD029 Model, a.TD030 PackageQty
                                , a.TD031 SiPackageQty, a.TD032 PackageUnit, a.TD033 OriCustomer, a.TD034 FromDocType, a.TD035 MtlItemType, a.TD039 DrawingNo, a.TD040 DrawingSeq, a.TD041 PartialSeq
                                , STUFF(STUFF(a.TD045, 5, 0, '-'), 8, 0, '-') OriPromiseDate
                                , STUFF(STUFF(a.TD046, 5, 0, '-'), 8, 0, '-') DeliveryDate
                                , a.TD057 TaxRate, a.TD058 PoPriceQty, a.TD059 PoPriceUnit, a.TD060 SiPriceQty
                                , a.TD061 TaxCode, a.TD062 DiscountRate, a.TD063 DiscountAmount
                                , i.MC002 InventoryName, LTRIM(RTRIM(t.NN002)) TaxName, t.NN005 InvoiceType, LTRIM(RTRIM(m.MB004)) Unit, m.MB022 LotManagement, m.MB043 MeasureType
                                , STUFF(STUFF(b.TC003, 5, 0, '-'), 8, 0, '-') PoDate
                                , STUFF(STUFF(b.TC024, 5, 0, '-'), 8, 0, '-') DocDate
                                , LTRIM(RTRIM(c.ML002)) CompanyShortName, LTRIM(RTRIM(c.ML003)) CompanyName
                                , LTRIM(RTRIM(CASE WHEN CHARINDEX('-', u.MF002) > 0 THEN LEFT(u.MF002, CHARINDEX('-', u.MF002) - 1) ELSE u.MF002 END)) CreatorName
                                , STUFF(STUFF(a.CREATE_DATE + ' ' + a.CREATE_TIME, 5, 0, '-'), 8, 0, '-') CreateDate
                                , STUFF(STUFF(a.MODI_DATE + ' ' + a.MODI_TIME, 5, 0, '-'), 8, 0, '-') LastModifiedDate
                                , LTRIM(RTRIM(a.CREATOR)) CreateBy, LTRIM(RTRIM(a.MODIFIER)) LastModifiedBy
                                , LTRIM(RTRIM(b.TC004)) SupplierNo, @NewCompanyNo CompanyNo";
                            sql += @" FROM PURTD a
                                    INNER JOIN PURTC b ON LTRIM(RTRIM(b.TC001)) = LTRIM(RTRIM(a.TD001)) AND LTRIM(RTRIM(b.TC002)) = LTRIM(RTRIM(a.TD002))
                                    LEFT JOIN CMSML c ON c.COMPANY = a.COMPANY
                                    INNER JOIN INVMB m ON LTRIM(RTRIM(m.MB001)) = LTRIM(RTRIM(a.TD004))
                                    INNER JOIN CMSMC i ON i.MC001 = a.TD007
                                    LEFT JOIN CMSNN t ON t.NN001 = a.TD061
                                    LEFT JOIN ADMMF u ON LTRIM(RTRIM(u.MF001)) = LTRIM(RTRIM(a.CREATOR))
                                    WHERE LTRIM(RTRIM(a.TD001)) + LTRIM(RTRIM(a.TD002)) + LTRIM(RTRIM(a.TD003)) IN @KeyNos";
                            //sql += OrderBy.Length > 0 ? " ORDER BY " + OrderBy : " ORDER BY PoErpNo DESC, PoSeq DESC, PoErpPrefix";

                            const int batchSize = 2000;
                            for (int i = 0; i < PagedNos.Count; i += batchSize)
                            {
                                result.AddRange(erpConnection.Query<dynamic>(sql, new { x.NewCompanyNo, KeyNos = KeyNos.Skip(i).Take(batchSize).ToArray() }).ToList());
                            }

                            ApplyDynamicOrdering(ref result, OrderBy, "PoErpNo DESC, PoSeq DESC, PoErpPrefix");
                            #endregion
                        }
                    });

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        draw,
                        recordsTotal = Nos.Count,
                        recordsFiltered = Nos.Count,
                        data = result,
                        status = "success"
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

        #region //GetPriceCheckingERP --取得ERP核價單頭資料 -- Chia Yuan 2025-4-4
        public string GetPriceCheckingERP(List<SupplierCompany> SupplierCompany
            , string PcErpPrefix, string PcErpNo
            , string PcErpPrefixs, string PcErpNos
            , string Currency, string TaxIncluded
            , string PcStartDate, string PcEndDate
            , string DocStartDate, string DocEndDate
            , string ConfirmStatus, string DocStatus, string SearchKey
            , bool IsOptions, int draw
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    List<(string, string)> Nos = new List<(string, string)>();

                    SupplierCompany.Distinct().ToList().ForEach(x =>
                    {
                        #region //確認公司別DB
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        var resultCorp = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo = x.CompanyNo == "ETG" ? "Eterge" : x.CompanyNo }) ?? throw new SystemException("【公司別】資料錯誤!");
                        ErpConnectionStrings = ConfigurationManager.AppSettings[resultCorp.ErpDb];
                        #endregion

                        using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //條件資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(a.TL001)) PcErpPrefix, LTRIM(RTRIM(a.TL002)) PcErpNo
                                    FROM PURTL a
                                    WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TL001", @" AND LTRIM(RTRIM(a.TL001)) = @TL001", PcErpPrefix);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TL002", @" AND LTRIM(RTRIM(a.TL002)) = @TL002", PcErpNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TL005", @" AND a.TL005 = @TL005", Currency);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TL003", @" AND a.TL003 >= @TL003",
                                PcStartDate.Length > 0 ? Convert.ToDateTime(PcStartDate).ToString("yyyyMMdd") : string.Empty);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TL003", @" AND a.TL003 <= @TL003",
                                PcEndDate.Length > 0 ? Convert.ToDateTime(PcEndDate).AddDays(1).ToString("yyyyMMdd") : string.Empty);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TL010", @" AND a.TL010 >= @TL010",
                                DocStartDate.Length > 0 ? Convert.ToDateTime(DocStartDate).ToString("yyyyMMdd") : string.Empty);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TL010", @" AND a.TL010 <= @TL010",
                                DocEndDate.Length > 0 ? Convert.ToDateTime(DocEndDate).AddDays(1).ToString("yyyyMMdd") : string.Empty);
                            if (PcErpPrefixs.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TL001s", @" AND LTRIM(RTRIM(a.TL001)) IN @TL001s", PcErpPrefixs.Split(','));
                            if (PcErpNos.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TL002s", @" AND LTRIM(RTRIM(a.TL002)) IN @TL002s", PcErpNos.Split(','));
                            if (TaxIncluded.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TL008s", @" AND a.TL008 IN @TL008s", TaxIncluded.Split(','));
                            if (DocStatus.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TL012s", @" AND a.TL012 IN @TL012s", DocStatus.Split(','));
                            if (ConfirmStatus.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TL006s", @" AND a.TL006 IN @TL006s", ConfirmStatus.Split(','));
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SearchKey",
                                @" AND (LTRIM(RTRIM(a.TL001)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TL002)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TL001)) + '-' + LTRIM(RTRIM(a.TL002)) LIKE N'%' + @SearchKey + '%'
                                    OR EXISTS (
                                        SELECT TOP 1 1 
                                        FROM PURTM aa 
                                        WHERE LTRIM(RTRIM(a.TM001)) = LTRIM(RTRIM(a.TL001)) 
                                        AND LTRIM(RTRIM(a.TM002)) = LTRIM(RTRIM(a.TL002))
                                        AND (LTRIM(RTRIM(a.TL004)) LIKE N'%' + @SearchKey + '%' 
                                            OR LTRIM(RTRIM(a.TL005)) LIKE N'%' + @SearchKey + '%')))", SearchKey);
                            sql += OrderBy.Length > 0 ? " ORDER BY " + OrderBy : " ORDER BY PcErpNo DESC, PcErpPrefix";

                            Nos.AddRange(erpConnection.Query(sql, dynamicParameters)
                                .Select(s => ((string)resultCorp.ErpNo, string.Format("{0}{1}", (string)s.PcErpPrefix, (string)s.PcErpNo)))
                                .ToList()); //取得識別單號
                            #endregion
                        }
                    });

                    List<(string, string)> PagedNos = new List<(string, string)>();
                    if (PageSize > 0) PagedNos = Nos.Skip(Math.Max(PageIndex, 0) * PageSize).Take(PageSize).ToList(); //符合條件的品號(與庫存條件交集)
                    else PagedNos = Nos;

                    List<dynamic> result = new List<dynamic>();

                    SupplierCompany.Distinct().ToList().ForEach(x =>
                    {
                        #region //確認公司別DB
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        var resultCorp = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo = x.CompanyNo == "ETG" ? "Eterge" : x.CompanyNo }) ?? throw new SystemException("【公司別】資料錯誤!");
                        ErpConnectionStrings = ConfigurationManager.AppSettings[resultCorp.ErpDb];
                        #endregion

                        string[] KeyNos = PagedNos.Where(w => w.Item1 == resultCorp.ErpNo).Select(s => s.Item2).ToArray();

                        using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //取得資料
                            sql = IsOptions ?
                                @"SELECT LTRIM(RTRIM(a.TL001)) PcErpPrefix, LTRIM(RTRIM(a.TL002)) PcErpNo
                                , LTRIM(RTRIM(a.TL001)) + '-' + LTRIM(RTRIM(a.TL002)) PcErpFullNo, @NewCompanyNo CompanyNo" :
                                @"SELECT LTRIM(RTRIM(a.TL001)) PcErpPrefix, LTRIM(RTRIM(a.TL002)) PcErpNo
                                , LTRIM(RTRIM(a.TL001)) + '-' + LTRIM(RTRIM(a.TL002)) PcErpFullNo
                                , STUFF(STUFF(a.TL003, 5, 0, '-'), 8, 0, '-') PcDate
                                , a.TL004 SupplierNo
                                , a.TL005 Currency, f.MF002 CurrencyName
                                , a.TL006 ConfirmStatus, a.TL007 Remark
                                , a.TL008 TaxIncluded, a.TL009 PrintCount
                                , STUFF(STUFF(a.TL010, 5, 0, '-'), 8, 0, '-') DocDate
                                , a.TL011 ConfirmUser
                                , a.TL012 DocStatus
                                , a.TL013 TransmissionCount
                                , LTRIM(RTRIM(c.ML002)) CompanyShortName, LTRIM(RTRIM(c.ML003)) CompanyName
                                , LTRIM(RTRIM(CASE WHEN CHARINDEX('-', u.MF002) > 0 THEN LEFT(u.MF002, CHARINDEX('-', u.MF002) - 1) ELSE u.MF002 END)) CreatorName
                                , STUFF(STUFF(a.CREATE_DATE + ' ' + a.CREATE_TIME, 5, 0, '-'), 8, 0, '-') CreateDate
                                , STUFF(STUFF(a.MODI_DATE + ' ' + a.MODI_TIME, 5, 0, '-'), 8, 0, '-') LastModifiedDate
                                , LTRIM(RTRIM(a.CREATOR)) CreateBy, LTRIM(RTRIM(a.MODIFIER)) LastModifiedBy
                                , @NewCompanyNo CompanyNo";
                            sql += @" FROM PURTL a
                                    LEFT JOIN CMSML c ON c.COMPANY = a.COMPANY
                                    INNER JOIN CMSMF f ON f.MF001 = a.TL005
                                    LEFT JOIN ADMMF u ON LTRIM(RTRIM(u.MF001)) = LTRIM(RTRIM(a.CREATOR))
                                    WHERE LTRIM(RTRIM(a.TL001)) + LTRIM(RTRIM(a.TL002)) IN @KeyNos";
                            sql += OrderBy.Length > 0 ? " ORDER BY " + OrderBy : " ORDER BY PcErpNo DESC, PcErpPrefix";

                            const int batchSize = 2000;
                            for (int i = 0; i < PagedNos.Count; i += batchSize)
                            {
                                result.AddRange(erpConnection.Query<dynamic>(sql, new { x.NewCompanyNo, KeyNos = KeyNos.Skip(i).Take(batchSize).ToArray() }).ToList());
                            }
                            #endregion
                        }
                    });

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        draw,
                        recordsTotal = Nos.Count,
                        recordsFiltered = Nos.Count,
                        data = result,
                        status = "success"
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

        #region //GetPcDetailERP --取得ERP核價單身資料 -- Chia Yuan 2025-4-7
        public string GetPcDetailERP(List<SupplierCompany> SupplierCompany
            , string PcErpPrefix, string PcErpNo, string PcSeq
            , string PcErpPrefixs, string PcErpNos, string PcSeqs
            , string MtlItemNo, string MtlItemNos, string MtlItemName
            , string SupplierItemNo, string SupplierItemNos, string UnitPricing
            , string EffectiveStartDate, string EffectiveEndDate
            , string ExpirationStartDate, string ExpirationEndDate
            , string ConfirmStatus, string SearchKey
            , bool IsOptions, int draw
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    List<(string, string)> Nos = new List<(string, string)>();

                    SupplierCompany.Distinct().ToList().ForEach(x =>
                    {
                        #region //確認公司別DB
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        var resultCorp = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo = x.CompanyNo == "ETG" ? "Eterge" : x.CompanyNo }) ?? throw new SystemException("【公司別】資料錯誤!");
                        ErpConnectionStrings = ConfigurationManager.AppSettings[resultCorp.ErpDb];
                        #endregion

                        using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //條件資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(a.TM001)) PcErpPrefix, LTRIM(RTRIM(a.TM002)) PcErpNo, LTRIM(RTRIM(a.TM003)) PcSeq
                                    FROM PURTM a
                                    INNER JOIN PURTL b ON LTRIM(RTRIM(b.TL001)) = LTRIM(RTRIM(a.TM001)) AND LTRIM(RTRIM(b.TL002)) = LTRIM(RTRIM(a.TM002))
                                    WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TM001", @" AND LTRIM(RTRIM(a.TM001)) = @TM001", PcErpPrefix);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TM002", @" AND LTRIM(RTRIM(a.TM002)) = @TM002", PcErpNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TM003", @" AND LTRIM(RTRIM(a.TM003)) = @TM003", PcSeq);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TM004", @" AND LTRIM(RTRIM(a.TM004)) = @TM004", MtlItemNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TM005", @" AND LTRIM(RTRIM(a.TM005)) = @TM005", MtlItemName);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TM007", @" AND LTRIM(RTRIM(a.TM007)) = @TM007", SupplierItemNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TM008", @" AND LTRIM(RTRIM(a.TM008)) = @TM008", UnitPricing);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "EffectiveStartDate", @" AND a.TM014 >= @EffectiveStartDate",
                                EffectiveStartDate.Length > 0 ? Convert.ToDateTime(EffectiveStartDate).ToString("yyyyMMdd") : string.Empty);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "EffectiveEndDate", @" AND a.TM014 <= @EffectiveEndDate",
                                EffectiveEndDate.Length > 0 ? Convert.ToDateTime(EffectiveEndDate).AddDays(1).ToString("yyyyMMdd") : string.Empty);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ExpirationStartDate", @" AND a.TM015 >= @ExpirationStartDate",
                                ExpirationStartDate.Length > 0 ? Convert.ToDateTime(ExpirationStartDate).ToString("yyyyMMdd") : string.Empty);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ExpirationEndDate", @" AND a.TM015 <= @ExpirationEndDate",
                                ExpirationEndDate.Length > 0 ? Convert.ToDateTime(ExpirationEndDate).AddDays(1).ToString("yyyyMMdd") : string.Empty);
                            if (PcErpPrefixs.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TM001s", @" AND LTRIM(RTRIM(a.TM001)) IN @TM001s", PcErpPrefixs.Split(','));
                            if (PcErpNos.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TM002s", @" AND LTRIM(RTRIM(a.TM002)) IN @TM002s", PcErpNos.Split(','));
                            if (PcSeqs.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TM003s", @" AND LTRIM(RTRIM(a.TM003)) IN @TM003s", PcSeqs.Split(','));
                            if (SupplierItemNos.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TM007s", @" AND LTRIM(RTRIM(a.TM007)) IN @TM007s", SupplierItemNos.Split(','));
                            if (ConfirmStatus.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TM011s", @" AND a.TM011 IN @TM011s", ConfirmStatus.Split(','));

                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SearchKey",
                                @" AND (LTRIM(RTRIM(a.TM001)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TM002)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TM003)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TM001)) + '-' + LTRIM(RTRIM(a.TM002)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TM001)) + '-' + LTRIM(RTRIM(a.TM002)) + '-' + LTRIM(RTRIM(a.TM003)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TM004)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TM005)) LIKE N'%' + @SearchKey + '%')", SearchKey);
                            sql += OrderBy.Length > 0 ? " ORDER BY " + OrderBy : " ORDER BY PcErpNo DESC, PcSeq DESC, PcErpPrefix";

                            Nos.AddRange(erpConnection.Query(sql, dynamicParameters)
                                .Select(s => ((string)resultCorp.ErpNo, string.Format("{0}{1}{2}", (string)s.PcErpPrefix, (string)s.PcErpNo, (string)s.PcSeq)))
                                .ToList()); //取得識別單號
                            #endregion
                        }
                    });

                    List<(string, string)> PagedNos = new List<(string, string)>();
                    if (PageSize > 0) PagedNos = Nos.Skip(Math.Max(PageIndex, 0) * PageSize).Take(PageSize).ToList(); //符合條件的品號(與庫存條件交集)
                    else PagedNos = Nos;

                    List<dynamic> result = new List<dynamic>();

                    SupplierCompany.Distinct().ToList().ForEach(x =>
                    {
                        #region //確認公司別DB
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        var resultCorp = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo = x.CompanyNo == "ETG" ? "Eterge" : x.CompanyNo }) ?? throw new SystemException("【公司別】資料錯誤!");
                        ErpConnectionStrings = ConfigurationManager.AppSettings[resultCorp.ErpDb];
                        #endregion

                        string[] KeyNos = PagedNos.Where(w => w.Item1 == resultCorp.ErpNo).Select(s => s.Item2).ToArray();

                        using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //取得資料
                            sql = IsOptions ?
                                @"SELECT LTRIM(RTRIM(a.TM001)) PcErpPrefix, LTRIM(RTRIM(a.TM002)) PcErpNo, LTRIM(RTRIM(a.TM003)) PcSeq
                                , LTRIM(RTRIM(a.TM001)) + '-' + LTRIM(RTRIM(a.TM002)) + '-' + LTRIM(RTRIM(a.TM003)) PcErpPrefixNo
                                , @NewCompanyNo CompanyNo" :
                                @"SELECT LTRIM(RTRIM(a.TM001)) PcErpPrefix, LTRIM(RTRIM(a.TM002)) PcErpNo, LTRIM(RTRIM(a.TM003)) PcSeq
                                , LTRIM(RTRIM(a.TM001)) + '-' + LTRIM(RTRIM(a.TM002)) + '-' + LTRIM(RTRIM(a.TM003)) PcErpPrefixNo
                                , LTRIM(RTRIM(a.TM004)) MtlItemNo,LTRIM(RTRIM(a.TM005)) MtlItemName, a.TM006 MtlItemSpec
                                , LTRIM(RTRIM(a.TM007)) SupplierItemNo
                                , a.TM008 UnitPricing, a.TM009 PcPriceUnit, a.TM010 UnitPrice, a.TM011 ConfirmStatus, a.TM012 Remark
                                , STUFF(STUFF(a.TM014, 5, 0, '-'), 8, 0, '-') EffectiveDate
                                , STUFF(STUFF(a.TM015, 5, 0, '-'), 8, 0, '-') ExpirationDate
                                , a.TM016 SmallUnit, a.TM017 OriEffectiveDate, a.TM018 OriUnitPrice
                                , LTRIM(RTRIM(c.ML002)) CompanyShortName, LTRIM(RTRIM(c.ML003)) CompanyName
                                , LTRIM(RTRIM(CASE WHEN CHARINDEX('-', u.MF002) > 0 THEN LEFT(u.MF002, CHARINDEX('-', u.MF002) - 1) ELSE u.MF002 END)) CreatorName
                                , STUFF(STUFF(a.CREATE_DATE + ' ' + a.CREATE_TIME, 5, 0, '-'), 8, 0, '-') CreateDate
                                , STUFF(STUFF(a.MODI_DATE + ' ' + a.MODI_TIME, 5, 0, '-'), 8, 0, '-') LastModifiedDate
                                , LTRIM(RTRIM(a.CREATOR)) CreateBy, LTRIM(RTRIM(a.MODIFIER)) LastModifiedBy
                                , @NewCompanyNo CompanyNo";
                            sql += @" FROM PURTM a
                                    INNER JOIN PURTL b ON LTRIM(RTRIM(b.TL001)) = LTRIM(RTRIM(a.TM001)) AND LTRIM(RTRIM(b.TL002)) = LTRIM(RTRIM(a.TM002))
                                    LEFT JOIN CMSML c ON c.COMPANY = a.COMPANY
                                    LEFT JOIN ADMMF u ON LTRIM(RTRIM(u.MF001)) = LTRIM(RTRIM(a.CREATOR))
                                    WHERE LTRIM(RTRIM(a.TM001)) + LTRIM(RTRIM(a.TM002)) + LTRIM(RTRIM(a.TM003)) IN @KeyNos";
                            sql += OrderBy.Length > 0 ? " ORDER BY " + OrderBy : " ORDER BY PcErpNo DESC, PcSeq DESC, PcErpPrefix";

                            const int batchSize = 2000;
                            for (int i = 0; i < PagedNos.Count; i += batchSize)
                            {
                                result.AddRange(erpConnection.Query<dynamic>(sql, new { x.NewCompanyNo, KeyNos = KeyNos.Skip(i).Take(batchSize).ToArray() }).ToList());
                            }
                            #endregion
                        }
                    });

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        draw,
                        recordsTotal = Nos.Count,
                        recordsFiltered = Nos.Count,
                        data = result,
                        status = "success"
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

        #region //GetGrDetail --取得ERP進貨單身 --Chia Yuan 2025-4-29
        public string GetGrDetailERP(List<SupplierCompany> SupplierCompany
            , string GrErpPrefix, string GrErpNo, string GrSeq, string GrErpFullNo, string GrErpPrefixNo
            , string GrErpFullNos, string GrErpPrefixNos
            , string PoErpPrefix, string PoErpNo, string PoSeq, string PoErpFullNo, string PoErpPrefixNo
            , string PoErpFullNos, string PoErpPrefixNos
            , string MtlItemNo, string MtlItemNos, string MtlItemName
            , string AcceptanceStartDate, string AcceptanceEndDate
            , string EffectiveStartDate, string EffectiveEndDate
            , string ConfirmStatus, string SearchKey
            , bool IsOptions, int draw
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    List<(string, string)> Nos = new List<(string, string)>();

                    SupplierCompany.Distinct().ToList().ForEach(x =>
                    {
                        #region //確認公司別DB
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        var resultCorp = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo = x.CompanyNo == "ETG" ? "Eterge" : x.CompanyNo }) ?? throw new SystemException("【公司別】資料錯誤!");
                        ErpConnectionStrings = ConfigurationManager.AppSettings[resultCorp.ErpDb];
                        #endregion

                        using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //條件資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(a.TH001)) GrErpPrefix, LTRIM(RTRIM(a.TH002)) GrErpNo, LTRIM(RTRIM(a.TH003)) GrSeq
                                    FROM PURTH a
                                    INNER JOIN PURTG b ON LTRIM(RTRIM(b.TG001)) = LTRIM(RTRIM(a.TH001)) AND LTRIM(RTRIM(b.TG002)) = LTRIM(RTRIM(a.TH002))
                                    WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TH001", @" AND LTRIM(RTRIM(a.TH001)) = @TH001", GrErpPrefix);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TH002", @" AND LTRIM(RTRIM(a.TH002)) = @TH002", GrErpNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TH003", @" AND LTRIM(RTRIM(a.TH003)) = @TH003", GrSeq);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "GrErpFullNo", @" AND LTRIM(RTRIM(a.TH001)) + '-' + LTRIM(RTRIM(a.TH002)) = @GrErpFullNo", GrErpFullNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "GrErpPrefixNo", @" AND LTRIM(RTRIM(a.TH001)) + '-' + LTRIM(RTRIM(a.TH002)) + '-' + LTRIM(RTRIM(a.TH003)) = @GrErpPrefixNo", GrErpPrefixNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TH004", @" AND LTRIM(RTRIM(a.TH004)) = @TH004", MtlItemNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TH005", @" AND LTRIM(RTRIM(a.TH005)) = @TH005", MtlItemName);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TH011", @" AND LTRIM(RTRIM(a.TH011)) = @TH011", PoErpPrefix);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TH012", @" AND LTRIM(RTRIM(a.TH012)) = @TH012", PoErpNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TH013", @" AND LTRIM(RTRIM(a.TH013)) = @TH013", PoSeq);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PoErpFullNo", @" AND a.TH011 + '-' + a.TH012 = @PoErpFullNo", PoErpFullNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PoErpPrefixNo", @" AND a.TH011 + '-' + a.TH012 + '-' + a.TH013 = @PoErpPrefixNo", PoErpPrefixNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "AcceptanceStartDate", @" AND a.TH014 >= @AcceptanceStartDate",
                                AcceptanceStartDate.Length > 0 ? Convert.ToDateTime(AcceptanceStartDate).ToString("yyyyMMdd") : string.Empty);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "AcceptanceEndDate", @" AND a.TH014 <= @AcceptanceEndDate",
                                AcceptanceEndDate.Length > 0 ? Convert.ToDateTime(AcceptanceEndDate).AddDays(1).ToString("yyyyMMdd") : string.Empty);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "EffectiveStartDate", @" AND a.TH036 >= @EffectiveStartDate",
                                EffectiveStartDate.Length > 0 ? Convert.ToDateTime(EffectiveStartDate).ToString("yyyyMMdd") : string.Empty);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "EffectiveEndDate", @" AND a.TH036 <= @EffectiveEndDate",
                                EffectiveEndDate.Length > 0 ? Convert.ToDateTime(EffectiveEndDate).AddDays(1).ToString("yyyyMMdd") : string.Empty);
                            if (GrErpFullNos.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "GrErpFullNos", @" AND LTRIM(RTRIM(a.TH001)) + '-' + LTRIM(RTRIM(a.TH002)) IN @GrErpFullNos", GrErpFullNos.Split(','));
                            if (GrErpPrefixNos.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "GrErpPrefixNos", @" AND LTRIM(RTRIM(a.TH001)) + '-' + LTRIM(RTRIM(a.TH002)) + '-' + LTRIM(RTRIM(a.TH003)) IN @GrErpPrefixNos", GrErpPrefixNos.Split(','));
                            if (MtlItemNos.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TH004s", @" AND a.TH004 IN @TH004s", MtlItemNos.Split(','));
                            if (PoErpFullNos.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PoErpFullNos", @" AND a.TH011 + '-' + a.TH012 IN @PoErpFullNos", PoErpFullNos.Split(','));
                            if (PoErpPrefixNos.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PoErpPrefixNos", @" AND a.TH011 + '-' + a.TH012 + '-' + a.TH013 IN @PoErpPrefixNos", PoErpPrefixNos.Split(','));
                            if (ConfirmStatus.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TH030s", @" AND a.TH030 IN @TH030s", ConfirmStatus.Split(','));

                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SearchKey",
                                @" AND (LTRIM(RTRIM(a.TH001)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TH002)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TH003)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TH001)) + '-' + LTRIM(RTRIM(a.TH002)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TH001)) + '-' + LTRIM(RTRIM(a.TH002)) + '-' + LTRIM(RTRIM(a.TH003)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TH011)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TH012)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TH013)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TH011)) + '-' + LTRIM(RTRIM(a.TH012)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TH011)) + '-' + LTRIM(RTRIM(a.TH012)) + '-' + LTRIM(RTRIM(a.TH013)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TH004)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TH005)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(b.TG005)) LIKE N'%' + @SearchKey + '%')", SearchKey);
                            sql += OrderBy.Length > 0 ? " ORDER BY " + OrderBy : " ORDER BY GrErpNo DESC, GrSeq DESC, GrErpPrefix";

                            Nos.AddRange(erpConnection.Query(sql, dynamicParameters)
                                .Select(s => ((string)resultCorp.ErpNo, string.Format("{0}{1}{2}", (string)s.GrErpPrefix, (string)s.GrErpNo, (string)s.GrSeq)))
                                .ToList()); //取得識別單號
                            #endregion
                        }
                    });

                    List<(string, string)> PagedNos = new List<(string, string)>();
                    if (PageSize > 0) PagedNos = Nos.Skip(Math.Max(PageIndex, 0) * PageSize).Take(PageSize).ToList(); //符合條件的品號(與庫存條件交集)
                    else PagedNos = Nos;

                    List<dynamic> result = new List<dynamic>();

                    SupplierCompany.Distinct().ToList().ForEach(x =>
                    {
                        #region //確認公司別DB
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        var resultCorp = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo = x.CompanyNo == "ETG" ? "Eterge" : x.CompanyNo }) ?? throw new SystemException("【公司別】資料錯誤!");
                        ErpConnectionStrings = ConfigurationManager.AppSettings[resultCorp.ErpDb];
                        #endregion

                        string[] KeyNos = PagedNos.Where(w => w.Item1 == resultCorp.ErpNo).Select(s => s.Item2).ToArray();

                        using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //取得資料
                            sql = IsOptions ?
                                @"SELECT LTRIM(RTRIM(a.TH001)) GrErpPrefix, LTRIM(RTRIM(a.TH002)) GrErpNo, LTRIM(RTRIM(a.TH003)) GrSeq
                                , LTRIM(RTRIM(a.TH001)) + '-' + LTRIM(RTRIM(a.TH002)) GrErpFullNo
                                , LTRIM(RTRIM(a.TH001)) + '-' + LTRIM(RTRIM(a.TH002)) + '-' + LTRIM(RTRIM(a.TH003)) GrErpPrefixNo
                                , @NewCompanyNo CompanyNo" :
                                @"SELECT LTRIM(RTRIM(a.TH001)) GrErpPrefix, LTRIM(RTRIM(a.TH002)) GrErpNo, LTRIM(RTRIM(a.TH003)) GrSeq
                                , LTRIM(RTRIM(a.TH001)) + '-' + LTRIM(RTRIM(a.TH002)) GrErpFullNo
                                , LTRIM(RTRIM(a.TH001)) + '-' + LTRIM(RTRIM(a.TH002)) + '-' + LTRIM(RTRIM(a.TH003)) GrErpPrefixNo
                                , LTRIM(RTRIM(a.TH004)) MtlItemNo, LTRIM(RTRIM(a.TH005)) MtlItemName, a.TH006 MtlItemSpec, a.TH007 GrQty, a.TH008 Unit, a.TH009 Inventory
                                , a.TH010 LotNumber, a.TH011 PoErpPrefix, a.TH012 PoErpNo, a.TH013 PoSeq
                                , STUFF(STUFF(a.TH014, 5, 0, '-'), 8, 0, '-') AcceptanceDate
                                , a.TH015 AcceptQty, a.TH016 GrPriceQty, a.TH017 ReturnQty
                                , a.TH018 OriUnitPrice, a.TH019 OriAmount, a.TH020 OriDiscountAmount
                                , a.TH024 ReceiptExpense, a.TH025 DiscountRemark
                                , a.TH027 Overdue, a.TH028 QcStatus, a.TH029 ReturnCode
                                , a.TH030 ConfirmStatus, a.TH031 CheckOutCode, a.TH032 UpdateCode
                                , a.TH033 Remark, a.TH034 InventoryQty, a.TH035 SmallUnit
                                , STUFF(STUFF(a.TH036, 5, 0, '-'), 8, 0, '-') EffectiveDate
                                , STUFF(STUFF(a.TH037, 5, 0, '-'), 8, 0, '-') ReCheckDate
                                , a.TH038 ConfirmUser
                                , a.TH039 ApvErpPrefix, a.TH040 ApvErpNo, a.TH041 ApvSeq
                                , a.TH042 ProjectCode
                                , a.TH045 OriPreTaxAmount, a.TH046 OriTaxAmount, a.TH047 PreTaxAmount, a.TH048 TaxAmount
                                , a.TH053 PackageUnit, a.TH054 BondedCode, a.TH058 ApproveStatus
                                , a.TH073 TaxRate, a.TH088 TaxCode, a.TH089 DiscountRate, a.TH090 DiscountAmount
                                , LTRIM(RTRIM(c.ML002)) CompanyShortName, LTRIM(RTRIM(c.ML003)) CompanyName
                                , LTRIM(RTRIM(CASE WHEN CHARINDEX('-', u.MF002) > 0 THEN LEFT(u.MF002, CHARINDEX('-', u.MF002) - 1) ELSE u.MF002 END)) CreatorName
                                , STUFF(STUFF(a.CREATE_DATE + ' ' + a.CREATE_TIME, 5, 0, '-'), 8, 0, '-') CreateDate
                                , STUFF(STUFF(a.MODI_DATE + ' ' + a.MODI_TIME, 5, 0, '-'), 8, 0, '-') LastModifiedDate
                                , LTRIM(RTRIM(a.CREATOR)) CreateBy, LTRIM(RTRIM(a.MODIFIER)) LastModifiedBy
                                , @NewCompanyNo CompanyNo";
                            sql += @" FROM PURTH a
                                    LEFT JOIN CMSML c ON c.COMPANY = a.COMPANY
                                    LEFT JOIN ADMMF u ON LTRIM(RTRIM(u.MF001)) = LTRIM(RTRIM(a.CREATOR))
                                    WHERE LTRIM(RTRIM(a.TH001)) + LTRIM(RTRIM(a.TH002)) + LTRIM(RTRIM(a.TH003)) IN @KeyNos";
                            sql += OrderBy.Length > 0 ? " ORDER BY " + OrderBy : " ORDER BY GrErpNo DESC, GrSeq DESC, GrErpPrefix";

                            const int batchSize = 2000;
                            for (int i = 0; i < PagedNos.Count; i += batchSize)
                            {
                                result.AddRange(erpConnection.Query<dynamic>(sql, new { x.NewCompanyNo, KeyNos = KeyNos.Skip(i).Take(batchSize).ToArray() }).ToList());
                            }
                            #endregion
                        }
                    });

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        draw,
                        recordsTotal = Nos.Count,
                        recordsFiltered = Nos.Count,
                        data = result,
                        status = "success"
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

        #region //GetRgDetailERP --取得ERP退貨單身資料 -- Chia Yuan 2025-4-18
        public string GetRgDetailERP(List<SupplierCompany> SupplierCompany
            , string RgErpPrefix, string RgErpNo, string RgSeq
            , string RgErpFullNo, string RgErpPrefixNo
            , string RgErpFullNos, string RgErpPrefixNos
            , string PoErpPrefix, string PoErpNo, string PoSeq
            , string PoErpFullNo, string PoErpPrefixNo
            , string PoErpFullNos, string PoErpPrefixNos
            , string MtlItemNo, string MtlItemName, string SupplierNo
            , string ConfirmStatus, string CheckOutCode, string SearchKey
            , bool IsOptions, int draw
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    List<(string, string)> Nos = new List<(string, string)>();

                    SupplierCompany.Distinct().ToList().ForEach(x =>
                    {
                        #region //確認公司別DB
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        var resultCorp = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo = x.CompanyNo == "ETG" ? "Eterge" : x.CompanyNo }) ?? throw new SystemException("【公司別】資料錯誤!");
                        ErpConnectionStrings = ConfigurationManager.AppSettings[resultCorp.ErpDb];
                        #endregion

                        using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //條件資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(a.TJ001)) RgErpPrefix, LTRIM(RTRIM(a.TJ002)) RgErpNo, LTRIM(RTRIM(a.TJ003)) RgSeq
                                    FROM PURTJ a
                                    WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TJ001", @" AND LTRIM(RTRIM(a.TJ001)) = @TJ001", RgErpPrefix);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TJ002", @" AND LTRIM(RTRIM(a.TJ002)) = @TJ002", RgErpNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TJ003", @" AND LTRIM(RTRIM(a.TJ003)) = @TJ003", RgSeq);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "RgErpFullNo", @" AND LTRIM(RTRIM(a.TJ001)) + '-' + LTRIM(RTRIM(a.TJ002)) = @RgErpFullNo", RgErpFullNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "RgErpPrefixNo", @" AND LTRIM(RTRIM(a.TJ001)) + '-' + LTRIM(RTRIM(a.TJ002)) + '-' + LTRIM(RTRIM(a.TJ003)) = @RgErpPrefixNo", RgErpPrefixNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TJ004", @" AND LTRIM(RTRIM(a.TJ004)) = @TJ004", MtlItemNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TJ005", @" AND LTRIM(RTRIM(a.TJ005)) = @TJ005", MtlItemName);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TJ016", @" AND LTRIM(RTRIM(a.TJ016)) = @TJ016", PoErpPrefix);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TJ017", @" AND LTRIM(RTRIM(a.TJ017)) = @TJ017", PoErpNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TJ018", @" AND LTRIM(RTRIM(a.TJ018)) = @TJ018", PoSeq);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PoErpFullNo", @" AND a.TJ016 + '-' + a.TJ017 = @PoErpFullNo", PoErpFullNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PoErpPrefixNo", @" AND a.TJ016 + '-' + a.TJ017 + '-' + a.TJ018 = @PoErpPrefixNo", PoErpPrefixNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TJ008",
                                @" AND EXISTS (
                                SELECT TOP 1 1 
                                FROM PURTI aa 
                                WHERE LTRIM(RTRIM(aa.TI001)) = LTRIM(RTRIM(a.TJ001)) 
                                AND LTRIM(RTRIM(aa.TI002)) = LTRIM(RTRIM(a.TJ002))
                                AND LTRIM(RTRIM(aa.TI004)) = @SupplierNo)", SupplierNo);
                            if (RgErpFullNos.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "RgErpFullNos", @" AND LTRIM(RTRIM(a.TJ001)) + '-' + LTRIM(RTRIM(a.TJ002)) IN @RgErpFullNos", RgErpFullNos.Split(','));
                            if (RgErpPrefixNos.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "RgErpPrefixNos", @" AND LTRIM(RTRIM(a.TJ001)) + '-' + LTRIM(RTRIM(a.TJ002)) + '-' + LTRIM(RTRIM(a.TJ003)) IN @RgErpPrefixNos", RgErpPrefixNos.Split(','));       
                            if (PoErpFullNos.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PoErpFullNos", @" AND a.TJ016 + '-' + a.TJ017 IN @PoErpFullNos", PoErpFullNos.Split(','));
                            if (PoErpPrefixNos.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PoErpPrefixNos", @" AND a.TJ016 + '-' + a.TJ017 + '-' + a.TJ018 IN @PoErpPrefixNos", PoErpPrefixNos.Split(','));
                            if (ConfirmStatus.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TJ020s", @" AND LTRIM(RTRIM(a.TJ020)) = @TJ020s", ConfirmStatus.Split(','));
                            if (CheckOutCode.Length > 0)
                                BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TJ021s", @" AND LTRIM(RTRIM(a.TJ021)) = @TJ021s", CheckOutCode.Split(','));
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SearchKey",
                                @" AND (LTRIM(RTRIM(a.TJ001)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TJ002)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TJ003)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TJ001)) + '-' + LTRIM(RTRIM(a.TJ002)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TJ001)) + '-' + LTRIM(RTRIM(a.TJ002)) + '-' + LTRIM(RTRIM(a.TJ003)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TJ016)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TJ017)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TJ018)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TJ016)) + '-' + LTRIM(RTRIM(a.TJ017)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TJ016)) + '-' + LTRIM(RTRIM(a.TJ017)) + '-' + LTRIM(RTRIM(a.TJ018)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TJ004)) LIKE N'%' + @SearchKey + '%'
                                    OR LTRIM(RTRIM(a.TJ005)) LIKE N'%' + @SearchKey + '%')", SearchKey);
                            sql += OrderBy.Length > 0 ? " ORDER BY " + OrderBy : " ORDER BY RgErpNo DESC, RgSeq DESC, RgErpPrefix";

                            Nos.AddRange(erpConnection.Query(sql, dynamicParameters)
                                .Select(s => ((string)resultCorp.ErpNo, string.Format("{0}{1}{2}", (string)s.RgErpPrefix, (string)s.RgErpNo, (string)s.RgSeq)))
                                .ToList()); //取得識別單號
                            #endregion
                        }
                    });

                    List<(string, string)> PagedNos = new List<(string, string)>();
                    if (PageSize > 0) PagedNos = Nos.Skip(Math.Max(PageIndex, 0) * PageSize).Take(PageSize).ToList(); //符合條件的品號(與庫存條件交集)
                    else PagedNos = Nos;

                    List<dynamic> result = new List<dynamic>();

                    SupplierCompany.Distinct().ToList().ForEach(x =>
                    {
                        #region //確認公司別DB
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        var resultCorp = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo = x.CompanyNo == "ETG" ? "Eterge" : x.CompanyNo }) ?? throw new SystemException("【公司別】資料錯誤!");
                        ErpConnectionStrings = ConfigurationManager.AppSettings[resultCorp.ErpDb];
                        #endregion

                        string[] KeyNos = PagedNos.Where(w => w.Item1 == resultCorp.ErpNo).Select(s => s.Item2).ToArray();

                        using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //取得資料
                            sql = IsOptions ?
                                @"LTRIM(RTRIM(a.TJ001)) RgErpPrefix, LTRIM(RTRIM(a.TJ002)) RgErpNo, LTRIM(RTRIM(a.TJ003)) RgSeq
                                , LTRIM(RTRIM(a.TJ001)) + '-' + LTRIM(RTRIM(a.TJ002)) RgErpFullNo
                                , LTRIM(RTRIM(a.TJ001)) + '-' + LTRIM(RTRIM(a.TJ002)) + '-' + LTRIM(RTRIM(a.TJ003)) RgErpPrefixNo
                                , @NewCompanyNo CompanyNo" :
                                @"SELECT LTRIM(RTRIM(a.TJ001)) RgErpPrefix, LTRIM(RTRIM(a.TJ002)) RgErpNo, LTRIM(RTRIM(a.TJ003)) RgSeq
                                , LTRIM(RTRIM(a.TJ001)) + '-' + LTRIM(RTRIM(a.TJ002)) RgErpFullNo
                                , LTRIM(RTRIM(a.TJ001)) + '-' + LTRIM(RTRIM(a.TJ002)) + '-' + LTRIM(RTRIM(a.TJ003)) RgErpPrefixNo
                                , LTRIM(RTRIM(a.TJ004)) MtlItemNo, LTRIM(RTRIM(a.TJ005)) MtlItemName, a.TJ006 MtlItemSpec, a.TJ007 Unit
                                , a.TJ008 RgUnitPrice, a.TJ009 RgQty, a.TJ010 RgPrice
                                , a.TJ011 Inventory, a.TJ012 LotNumber
                                , LTRIM(RTRIM(a.TJ013)) GrErpPrefix, LTRIM(RTRIM(a.TJ014)) GrErpNo, LTRIM(RTRIM(a.TJ015)) GrSeq
                                , LTRIM(RTRIM(a.TJ016)) PoErpPrefix, LTRIM(RTRIM(a.TJ017)) PoErpNo, LTRIM(RTRIM(a.TJ018)) PoSeq
                                , TJ019 Remark, TJ020 ConfirmStatus, TJ021 CheckOutCode
                                , TJ022 InventoryQty, TJ023 SmallUnit, TJ028 UpdateCode
                                , a.TJ030 OriPreTaxAmount, a.TJ031 OriTaxAmount, a.TJ032 PreTaxAmount, a.TJ033 TaxAmount, a.TJ034 PackageUnit
                                , a.TJ035 ReturnType, a.TJ036 BondedCode, a.TJ052 TaxRate, a.TJ060 TaxCode
                                , a.TJ061 DiscountRate, a.TJ062 DiscountAmount
                                , LTRIM(RTRIM(c.ML002)) CompanyShortName, LTRIM(RTRIM(c.ML003)) CompanyName
                                , LTRIM(RTRIM(CASE WHEN CHARINDEX('-', u.MF002) > 0 THEN LEFT(u.MF002, CHARINDEX('-', u.MF002) - 1) ELSE u.MF002 END)) CreatorName
                                , STUFF(STUFF(a.CREATE_DATE + ' ' + a.CREATE_TIME, 5, 0, '-'), 8, 0, '-') CreateDate
                                , STUFF(STUFF(a.MODI_DATE + ' ' + a.MODI_TIME, 5, 0, '-'), 8, 0, '-') LastModifiedDate
                                , LTRIM(RTRIM(a.CREATOR)) CreateBy, LTRIM(RTRIM(a.MODIFIER)) LastModifiedBy
                                , @NewCompanyNo CompanyNo";
                            sql += @" FROM PURTJ a
                                    LEFT JOIN CMSML c ON c.COMPANY = a.COMPANY
                                    LEFT JOIN ADMMF u ON LTRIM(RTRIM(u.MF001)) = LTRIM(RTRIM(a.CREATOR))
                                    WHERE LTRIM(RTRIM(a.TJ001)) + LTRIM(RTRIM(a.TJ002)) + LTRIM(RTRIM(a.TJ003)) IN @KeyNos";
                            sql += OrderBy.Length > 0 ? " ORDER BY " + OrderBy : " ORDER BY RgErpNo DESC, RgSeq DESC, RgErpPrefix";

                            const int batchSize = 2000;
                            for (int i = 0; i < PagedNos.Count; i += batchSize)
                            {
                                result.AddRange(erpConnection.Query<dynamic>(sql, new { x.NewCompanyNo, KeyNos = KeyNos.Skip(i).Take(batchSize).ToArray() }).ToList());
                            }
                            #endregion
                        }
                    });

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        draw,
                        recordsTotal = Nos.Count,
                        recordsFiltered = Nos.Count,
                        data = result,
                        status = "success"
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

        #region//UpdatePurchaseOrderBatch --更新ERP採購單身 -- Chia Yuan 2025.5.16
        public string UpdatePurchaseOrderBatch(PurchaseOrderMaster model)
        {
            try
            {
                if (model.CompanyNo.Length <= 0) throw new SystemException("【公司別】資料錯誤!");

                using (TransactionScope transactionScope = new TransactionScope(TransactionScopeOption.Required, new TimeSpan(0, 180, 0)))
                {
                    string ErpConnectionStrings = "", ErpNo = "";
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認公司別DB
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { model.CompanyNo }) ?? throw new SystemException("【公司別】資料錯誤!");
                        #endregion

                        ErpConnectionStrings = ConfigurationManager.AppSettings[result.ErpDb];
                        ErpNo = result.ErpNo;

                        #region//判斷使用者是否存在
                        sql = @"SELECT TOP 1 a.UserNo
                                FROM BAS.[User] a
                                WHERE a.UserNo = @CurrentUser";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { model.CurrentUser }) ?? throw new SystemException("【使用者】資料錯誤!");
                        #endregion
                    }

                    using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        #region //判斷ERP使用者是否存在
                        sql = @"SELECT TOP 1 MF001, MF004, MF005
                                FROM ADMMF
                                WHERE COMPANY = @ErpNo
                                AND MF001 = @CurrentUser";
                        var result = erpConnection.QueryFirstOrDefault(sql, new { ErpNo, model.CurrentUser }) ?? throw new SystemException("【ERP使用者】資料錯誤!");
                        #endregion

                        #region //判斷ERP使用者是否存在
                        sql = @"SELECT TOP 1 MF001, MF004, MF005
                                FROM ADMMF
                                WHERE COMPANY = @ErpNo
                                AND MF001 = @PoUser";
                        result = erpConnection.QueryFirstOrDefault(sql, new { ErpNo, model.PoUser });
                        if (!string.IsNullOrWhiteSpace(model.PoUser) && result == null) throw new SystemException("【ERP採購人】資料錯誤!");
                        model.PoUser = result?.MF001 ?? "";
                        #endregion

                        var PoErpFullNos = model.PoDetailItems.Select(s => string.Format("{0}{1}", s.PoErpPrefix, s.PoErpNo)).Distinct().ToArray();

                        #region //取得ERP採購單
                        sql = @"SELECT TOP 1 LTRIM(RTRIM(a.TC001)) PoErpPrefix, LTRIM(RTRIM(a.TC002)) PoErpNo
                                , a.TC005, a.TC014, a.TC018, a.TC026, a.TC050, a.TC051
                                FROM PURTC a
                                WHERE LTRIM(RTRIM(a.TC001)) = @PoErpPrefix AND LTRIM(RTRIM(a.TC002)) = @PoErpNo";
                        var PURTC = erpConnection.QueryFirstOrDefault(sql, new { model.PoErpPrefix, model.PoErpNo }) ?? throw new SystemException("【ERP採購單】資料錯誤!");
                        #endregion

                        if (PURTC.TC014 != "N") throw new SystemException("【採購單】狀態錯誤<br/>");
                        if (PURTC.TC050 == "Y") throw new SystemException("【採購單】已鎖定，不得更改<br/>");

                        #region //取得ERP採購單身
                        sql = @"SELECT LTRIM(RTRIM(a.TD001)) PoErpPrefix, LTRIM(RTRIM(a.TD002)) PoErpNo, LTRIM(RTRIM(a.TD003)) PoSeq
                                , LTRIM(RTRIM(a.TD004)) MtlItemNo, STUFF(STUFF(a.TD012, 5, 0, '-'), 8, 0, '-') PromiseDate
                                , a.TD016, a.TD018, a.TD057
                                , LTRIM(RTRIM(m.MB004)) Unit, m.MB022 LotManagement
                                FROM PURTD a
                                INNER JOIN INVMB m ON m.MB001 = a.TD004
                                WHERE LTRIM(RTRIM(a.TD001)) + LTRIM(RTRIM(a.TD002)) IN @PoErpFullNos";
                        var resultPURTD = erpConnection.Query(sql, new { PoErpFullNos }).ToList();
                        
                        if (!resultPURTD.Any()) throw new SystemException("【ERP採購單身】資料錯誤!");
                        //if (resultPURTD.Count != PoErpPrefixNos.Length) throw new SystemException("【ERP採購單身】資料錯誤!");
                        if (resultPURTD.Any(x => x.TD016 != "N")) throw new SystemException("【ERP採購單身】已結案，不得更改!");
                        #endregion

                        #region //取得ERP單位換算
                        sql = @"SELECT LTRIM(RTRIM(a.MD001)) MtlItemNo
                                , LTRIM(RTRIM(a.MD002)) ConversionUnit, a.MD003 SwapNumerator, a.MD004 SwapDenominator, LTRIM(RTRIM(b.MB004)) Unit
                                FROM INVMD a 
                                INNER JOIN INVMB b ON b.MB001 = a.MD001
                                WHERE LTRIM(RTRIM(a.MD001)) IN @MD001s";
                        var resultINVMD = erpConnection.Query(sql, new { MD001s = resultPURTD.Select(s => s.MtlItemNo).Distinct().ToArray() }).ToList();
                        #endregion

                        #region //取得ERP幣別
                        sql = @"SELECT LTRIM(RTRIM(a.MF001)) Currency
                                , LTRIM(RTRIM(a.MF002)) CurrencyName, LTRIM(RTRIM(a.MF001)) + ' ' + LTRIM(RTRIM(a.MF002)) CurrencyWithNo
                                , a.MF003 UnitPriceRound, a.MF004 AmountRound, a.MF005 UnitCostRound, a.MF006 CostRound
                                FROM CMSMF a WHERE LTRIM(RTRIM(a.MF001)) = @Currency";
                        var resultCMSMF = erpConnection.QueryFirstOrDefault(sql, new { model.Currency }) ?? throw new SystemException("【幣別】資料錯誤!"); ;
                        #endregion

                        if (!DateTime.TryParse(model.PoDate, out DateTime PoDate)) throw new SystemException("【採購日期】資料錯誤!"); ;

                        double BankSellingRate = 1;
                        int UnitPriceDecimal = 0, AmountDecimal = 0, UnitCostDecimal = 0, CostAmountDecimal = 0;

                        #region //外幣處理
                        var QueryCMSMG = GetExchangeRate(new List<SupplierCompany>() { new SupplierCompany { CompanyNo = model.CompanyNo } }, model.Currency, model.PoDate, "", false, -1, true, "", -1, -1);
                        var resultStatus = JObject.Parse(QueryCMSMG)["status"]?.ToString();
                        if (resultStatus == "success")
                        {
                            JObject.Parse(QueryCMSMG)["data"].ToList().ForEach(x =>
                            {
                                double.TryParse(x["BankSellingRate"].ToString(), out BankSellingRate);
                                int.TryParse(x["UnitPriceDecimal"].ToString(), out UnitPriceDecimal);
                                int.TryParse(x["AmountDecimal"].ToString(), out AmountDecimal);
                                int.TryParse(x["UnitCostDecimal"].ToString(), out UnitCostDecimal);
                                int.TryParse(x["CostAmountDecimal"].ToString(), out CostAmountDecimal);
                            });
                        }
                        #endregion

                        #region //取得交貨庫別
                        sql = @"SELECT LTRIM(RTRIM(MC001)) InventoryNo
                                , LTRIM(RTRIM(MC002)) InventoryName
                                , LTRIM(RTRIM(MC001)) + ' ' + LTRIM(RTRIM(MC002)) InventoryWithNo
                                FROM CMSMC WHERE LTRIM(RTRIM(MC001)) IN @InventoryNos";
                        var resultCMSMC = erpConnection.Query(sql, new { InventoryNos = model.PoDetailItems.Select(s => s.Inventory).Distinct().ToArray() }).ToList();
                        #endregion

                        #region //取得單位資料
                        sql = @"SELECT LTRIM(RTRIM(MX001)) UomNo, LTRIM(RTRIM(MX002)) UomName,
                                MX001+' '+ MX002 UomWithNo
                                FROM INVMX WHERE LTRIM(RTRIM(MX001)) IN @PoUnits";
                        var resultINVMX = erpConnection.Query(sql, new { PoUnits = model.PoDetailItems.Select(s => s.PoUnit).Union(model.PoDetailItems.Select(s=>s.PoPriceUnit)).Distinct().ToArray() }).ToList();
                        #endregion

                        #region //取得運輸方式
                        sql = @"SELECT NJ001 ShipMethod, NJ002 ShipMethodName FROM CMSNJ WHERE NJ001 = @ShipMethod";
                        result = erpConnection.QueryFirstOrDefault(sql, new { model.ShipMethod });
                        if (!string.IsNullOrWhiteSpace(model.ShipMethod) && result == null) throw new SystemException("【運輸方式】資料錯誤!");
                        model.ShipMethod = result?.ShipMethod ?? "";
                        model.ShipMethodName = result?.ShipMethodName ?? "";
                        #endregion

                        string msg = string.Empty;
                        model.PoDetailItems.ForEach(x =>
                        {
                            var PURTD = resultPURTD.FirstOrDefault(f => f.PoErpPrefix == x.PoErpPrefix && f.PoErpNo == x.PoErpNo && f.PoSeq == x.PoSeq);
                            var CMSMC = resultCMSMC.FirstOrDefault(f => f.InventoryNo == x.Inventory);
                            var INVMX1 = resultINVMX.FirstOrDefault(f => f.UomNo == x.PoUnit);
                            var INVMX2 = resultINVMX.FirstOrDefault(f => f.UomNo == x.PoPriceUnit);

                            x.PoErpPrefixNo = PURTD.PoErpPrefix + "-" + PURTD.PoErpNo + "-" + PURTD.PoSeq;

                            if (PURTD.TD018 != "N") x.Msg += "【採購單身】狀態錯誤<br/>";
                            if (PURTD.TD016 != "N") x.Msg += "【採購單身】已結案，不得更改<br/>";
                            if (CMSMC == null) x.Msg += "【交貨庫別】選取錯誤<br/>";
                            if (INVMX1 == null || INVMX2 == null) x.Msg += "【單位】選取錯誤<br/>";
                            //if (x.PoUnit != x.PoPriceUnit) 
                            //{
                            //    var INVMD = resultINVMD.FirstOrDefault(f => f.MtlItemNo == PURTD?.MtlItemNo && f.ConversionUnit == x.PoPriceUnit); //ERP單位換算
                            //    x.Msg += "【單位】資料錯誤<br/>";
                            //}
                            if (!DateTime.TryParse(x.PromiseDate, out DateTime PromiseDate)) x.Msg += "【預交日】格式錯誤<br/>";
                            if (PromiseDate == DateTime.MinValue) x.Msg += "【預交日】格式錯誤<br/>";
                            if (!Regex.IsMatch(x.UrgentMtl, "^(Y|N)$", RegexOptions.IgnoreCase)) x.Msg += "【急料】選取錯誤<br/>";
                        });

                        if (model.PoDetailItems.Any(a => !string.IsNullOrWhiteSpace(a.Msg)))
                            throw new SystemException(model.CompanyNo + ":" + string.Join("<br/>", model.PoDetailItems.Where(w => !string.IsNullOrWhiteSpace(w.Msg)).Select(s => s.PoErpPrefixNo + "<br/>" + s.Msg)));

                        decimal totalAmount = 0m;
                        decimal totalTax = 0m;
                        double recordCount = 0d;
                        decimal taxRate = PURTC.TC026; //營業稅率

                        model.PoDetailItems.ForEach(x =>
                        {
                            //TC051 單身多稅率、TC026 營業稅率、TD057 營業稅率
                            var PURTD = resultPURTD.FirstOrDefault(f => f.PoErpPrefix == x.PoErpPrefix && f.PoErpNo == x.PoErpNo && f.PoSeq == x.PoSeq);

                            decimal tax = 0m; //, amount = 0m;
                            if (PURTC.TC051 == "Y") taxRate = PURTD.TD057; //單身多稅率 使用單身稅率

                            x.MODI_PRID = "SRM";
                            x.MODI_DATE = CreateDate.ToString("yyyyMMdd");
                            x.MODI_TIME = CreateDate.ToString("HH:mm:ss");
                            x.MODI_AP = model.CurrentUser + "PC";
                            x.MODIFIER = model.CurrentUser;

                            x.PoQty = Math.Round(x.PoQty, 3, MidpointRounding.AwayFromZero); //採購數量 四捨五入至小數點3位
                            x.PoPriceQty = Math.Round(x.PoPriceQty, 3, MidpointRounding.AwayFromZero); //計價數量 四捨五入至小數點3位
                            x.PoUnitPrice = Math.Round(x.PoUnitPrice, UnitPriceDecimal, MidpointRounding.AwayFromZero); //採購單價
                            x.PoPrice = Math.Round(x.PoUnitPrice * (decimal)x.PoPriceQty, AmountDecimal, MidpointRounding.AwayFromZero); //採購金額
                            x.PoPriceSiQty = x.SiQty; //已交數量

                            if (x.PoUnit != x.PoPriceUnit)
                            {
                                double convertRate = 1d;

                                if (x.PoPriceUnit != PURTD.Unit)
                                {
                                    var _unitConvert = resultINVMD.FirstOrDefault(f => f.ConversionUnit == x.PoPriceUnit); //取得單位轉換
                                    double SwapDenominator = _unitConvert?.SwapNumerator == 0 ? 1 : (double)_unitConvert?.SwapNumerator;
                                    convertRate = (double)_unitConvert?.SwapDenominator / SwapDenominator; //先轉回庫存單位
                                }
                                if (x.PoUnit != PURTD.Unit)
                                {
                                    var _unitConvert = resultINVMD.FirstOrDefault(f => f.ConversionUnit == x.PoUnit); //取得單位轉換
                                    double SwapDenominator = _unitConvert?.SwapDenominator == 0 ? 1 : (double)_unitConvert?.SwapDenominator;
                                    convertRate = convertRate * (double)_unitConvert?.SwapNumerator / SwapDenominator;
                                }

                                x.PoPriceSiQty = Math.Round(x.SiQty * convertRate, UnitPriceDecimal, MidpointRounding.AwayFromZero); //已交計價數量
                            }

                            #region //稅額計算
                            if (x.PoPrice > 0)
                            {
                                switch (model.Taxation)
                                {
                                    case "1":
                                        tax = Math.Round(x.PoPrice - x.PoPrice / (1 + Convert.ToDecimal(taxRate)), AmountDecimal, MidpointRounding.AwayFromZero);
                                        //x.PoPrice = x.PoPrice - tax; //採購金額
                                        break;
                                    case "2":
                                        tax = Math.Round(x.PoPrice * Convert.ToDecimal(taxRate), AmountDecimal, MidpointRounding.AwayFromZero);
                                        break;
                                }
                            }
                            #endregion

                            totalAmount += x.PoPrice; //採購金額
                            totalTax += tax; //稅額
                            recordCount += x.PoQty; //採購數量

                            DateTime.TryParse(x.PromiseDate, out DateTime PromiseDate);
                            x.PromiseDate = PromiseDate.ToString("yyyyMMdd");
                            x.Remark = x.Remark ?? "";

                            #region //更新採購單身
                            sql = @"UPDATE a SET
                                    a.MODI_PRID = @MODI_PRID,
                                    a.MODI_DATE = @MODI_DATE,
                                    a.MODI_TIME = @MODI_TIME,
                                    a.MODI_AP = @MODI_AP,
                                    a.MODIFIER = @MODIFIER,
                                    a.FLAG = FLAG + 1,
                                    a.TD007 = @Inventory,
                                    a.TD008 = @PoQty,
                                    a.TD009 = @PoUnit,
                                    a.TD010 = @PoUnitPrice,
                                    a.TD011 = @PoPrice,
                                    a.TD012 = @PromiseDate,
                                    a.TD014 = @Remark,
                                    a.TD015 = @SiQty,
                                    a.TD025 = @UrgentMtl,
                                    a.TD058 = @PoPriceQty,
                                    a.TD059 = @PoPriceUnit,
                                    a.TD060 = @PoPriceSiQty
                                    FROM PURTD a 
                                    WHERE LTRIM(RTRIM(a.TD001)) = @PoErpPrefix 
                                    AND LTRIM(RTRIM(a.TD002)) = @PoErpNo
                                    AND LTRIM(RTRIM(a.TD003)) = @PoSeq";
                            rowsAffected += erpConnection.Execute(sql, x);
                            #endregion
                        });

                        if (model.Taxation == "1") totalAmount = totalAmount - totalTax; //採購金額

                        model.PoDate = PoDate.ToString("yyyyMMdd");
                        model.TotalAmount = Math.Round(totalAmount, AmountDecimal); //總金額
                        model.TaxAmount = Math.Round(totalTax, AmountDecimal); //總稅額

                        #region //更新採購單頭
                        sql = @"UPDATE a SET
                                a.MODI_PRID = @MODI_PRID,
                                a.MODI_DATE = @MODI_DATE,
                                a.MODI_TIME = @MODI_TIME,
                                a.MODI_AP = @MODI_AP,
                                a.MODIFIER = @MODIFIER,
                                a.FLAG = FLAG + 1,
                                a.TC003 = @PoDate,
                                a.TC004 = @SupplierNo,
                                a.TC005 = @Currency,
                                a.TC006 = @PoExchangeRate,
                                a.TC008 = @PaymentTerm,
                                a.TC009 = @Remark,
                                a.TC011 = @PoUser,
                                a.TC017 = @ShipMethodName,
                                a.TC018 = @Taxation,
                                a.TC019 = @TotalAmount,
                                a.TC020 = @TaxAmount,
                                a.TC021 = @FirstAddress,
                                a.TC026 = @TaxRate,
                                a.TC027 = @PaymentTermCode,
                                a.TC041 = @ShipMethod,
                                a.TC047 = @TaxCode,
                                a.TC052 = @ContactUser
                                FROM PURTC a 
                                WHERE LTRIM(RTRIM(a.TC001)) = @PoErpPrefix AND LTRIM(RTRIM(a.TC002)) = @PoErpNo";
                        rowsAffected += erpConnection.Execute(sql, new
                        {
                            model.PoErpPrefix,
                            model.PoErpNo,
                            MODI_PRID = "SRM",
                            MODI_DATE = CreateDate.ToString("yyyyMMdd"),
                            MODI_TIME = CreateDate.ToString("HH:mm:ss"),
                            MODI_AP = model.CurrentUser + "PC",
                            MODIFIER = model.CurrentUser,
                            model.PoDate,
                            model.SupplierNo,
                            model.Currency,
                            model.PoExchangeRate,
                            model.PaymentTerm,
                            Remark = model.Remark ?? "",
                            model.PoUser,
                            model.ShipMethodName,                            
                            model.Taxation,
                            model.TotalAmount,
                            model.TaxAmount,
                            model.FirstAddress,
                            TaxRate = taxRate,
                            model.PaymentTermCode,
                            model.ShipMethod,
                            model.TaxCode,
                            model.ContactUser
                        });
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

        #region //API公用函數
        private void ApplyDynamicOrdering(ref List<dynamic> items, string orderByExpression, string defOrderByExpression)
        {
            if (string.IsNullOrEmpty(orderByExpression))
            {
                // 默認排序
                //items.OrderBy(item => ((IDictionary<string, object>)item)[defOrderByExpression]);
                orderByExpression = defOrderByExpression;
            }

            // 解析排序表達式，格式例如："CompanyName ASC, CompanyShortName DESC"
            var orderBySegments = orderByExpression.Split(',');
            IOrderedEnumerable<dynamic> orderedItems = null;

            foreach (var segment in orderBySegments)
            {
                var trimmedSegment = segment.Trim();
                var parts = trimmedSegment.Split(' ');

                var propertyName = parts[0].Trim();
                var descending = parts.Length > 1 && parts[1].Trim().Equals("DESC", StringComparison.OrdinalIgnoreCase);

                if (orderedItems == null)
                {
                    // 第一個排序條件
                    orderedItems = descending
                        ? items.OrderByDescending(item => GetPropertyValue(item, propertyName))
                        : items.OrderBy(item => GetPropertyValue(item, propertyName));
                }
                else
                {
                    // 後續排序條件
                    orderedItems = descending
                        ? orderedItems.ThenByDescending(item => GetPropertyValue(item, propertyName))
                        : orderedItems.ThenBy(item => GetPropertyValue(item, propertyName));
                }
            }

            //return orderedItems ?? items; // 如果沒有有效的排序條件，返回原始集合
        }

        private object GetPropertyValue(dynamic item, string propertyName)
        {
            // 處理 dynamic 類型，將項目視為 IDictionary<string, object>
            var dictionary = (IDictionary<string, object>)item;

            if (dictionary.TryGetValue(propertyName, out var value))
            {
                return value ?? string.Empty; // 處理 null 值
            }

            return string.Empty; // 屬性不存在時返回空字符串
        }
        #endregion

        #region //Shintokuro
        #region //GetPurchaseRequisitionBM --取得請購單(BM) -- Shintokuro 2025-02-05
        public string GetPurchaseRequisitionBM(string CompanyNo, string PrErpPrefix, string PrErpNo
            , string PrId, string PurchaseOrde, string MtlItemNo, string MtlItemName, string UserNo, string SupplierNo
            , string SearchKey, string OrderBy, string PageIndex, string PageSize
            , int draw)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //確認公司別DB
                    sql = @"SELECT a.ErpNo, a.ErpDb, a.CompanyId
                            FROM BAS.Company a
                            WHERE CompanyNo = @CompanyNo";
                    var resultCorp = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo = CompanyNo == "ETG" ? "Eterge" : CompanyNo }) ?? throw new SystemException("【公司別】資料錯誤!");
                    #endregion

                    ErpConnectionStrings = ConfigurationManager.AppSettings[resultCorp.ErpDb];
                    int CompanyId = resultCorp.CompanyId;
                    string ErpNo = resultCorp.ErpNo;

                    #region //取得資料
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.PrId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.PrNo, a.CompanyId, a.DepartmentId, a.UserId, a.BpmNo, a.PrErpPrefix, a.PrErpNo, a.Edition
                        , FORMAT(a.PrDate, 'yyyy-MM-dd') PrDate, FORMAT(a.DocDate, 'yyyy-MM-dd') DocDate, a.PrRemark
                        , a.TotalQty, a.Amount, a.BudgetDepartmentId, a.PrStatus, a.SignupStaus, a.LockStaus
                        , a.ConfirmStatus, a.ConfirmUserId, a.BpmTransferStatus, a.BpmTransferUserId, a.BomType
                        , a.BpmTransferDate, a.TransferStatus, a.TransferDate
                        , FORMAT(a.CreateDate, 'yyyy-MM-dd HH:mm:ss') CreateDate
                        , b.UserNo, b.UserName
                        , c.DepartmentNo, c.DepartmentName
                        , d.StatusNo, d.StatusName
                        , e.TypeNo Priority
                        , (
                            SELECT x.FileId
                            FROM SCM.PrFile x
                            WHERE PrId = a.PrId
                            AND PrDetailId IS NULL
                            FOR JSON PATH, ROOT('data')
                        ) PrFile";
                    sqlQuery.mainTables =
                        $@"FROM SCM.PurchaseRequisition a
                        INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                        INNER JOIN BAS.Department c ON a.DepartmentId = c.DepartmentId
                        INNER JOIN BAS.[Status] d ON a.PrStatus = d.StatusNo AND d.StatusSchema = 'PurchaseRequisition.PrStatus'
                        INNER JOIN BAS.[Type] e ON a.Priority = e.TypeNo AND e.TypeSchema = 'PR.Priority'";
                    sqlQuery.auxTables = "";
                    string queryCondition = @" AND a.CompanyId = @CompanyId AND PrStatus = 'Y'";
                    dynamicParameters.Add("CompanyId", CompanyId);
                    if (MtlItemNo.Length > 0 || MtlItemName.Length > 0)
                    {
                        queryCondition += @" AND EXISTS (
                            SELECT TOP 1 1 
                            FROM SCM.PrDetail aa 
                            INNER JOIN PDM.MtlItem bb ON aa.MtlItemId = bb.MtlItemId
                            WHERE aa.PrId = a.PrId";
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemNo", @" AND bb.MtlItemNo = @MtlItemNo", MtlItemNo);
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemName", @" AND bb.MtlItemName = @MtlItemName", MtlItemName);
                        queryCondition += ")";
                    }
                    if (SupplierNo.Length > 0)
                    {
                        queryCondition += @" AND EXISTS (
                            SELECT TOP 1 1
                            FROM SCM.PrDetail aa
                            INNER JOIN SCM.Supplier bb ON bb.SupplierId = aa.SupplierId
                            WHERE aa.PrId = a.PrId";
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SupplierNo", @" AND bb.SupplierNo = @SupplierNo", SupplierNo);
                        queryCondition += ")";
                    }
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PrId", @" AND a.PrId = @PrId", PrId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PrErpPrefix", @" AND a.PrErpPrefix = @PrErpPrefix", PrErpPrefix);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PrErpNo", @" AND a.PrErpNo = @PrErpNo", PrErpNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PurchaseOrde",
                        @" AND (a.PrErpPrefix + '-' + a.PrErpNo) = @PurchaseOrde", PurchaseOrde);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserNo", @" AND b.UserNo = @UserNo", UserNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey",
                        @" AND EXISTS (
                            SELECT TOP 1 1
                            FROM SCM.PrDetail aa
                            INNER JOIN PDM.MtlItem bb ON aa.MtlItemId = bb.MtlItemId
                            INNER JOIN SCM.Supplier cc ON aa.SupplierId = cc.SupplierId
                            WHERE aa.PrId = a.PrId
                            AND (bb.MtlItemNo LIKE N'%' + @SearchKey + '%'
                                OR bb.MtlItemName LIKE N'%' + @SearchKey + '%'
                                OR cc.SupplierNo LIKE N'%' + @SearchKey + '%'))", SearchKey);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.PrId DESC";
                    sqlQuery.pageIndex = Convert.ToInt32(PageIndex);
                    sqlQuery.pageSize = Convert.ToInt32(PageSize);
                    var result = BaseHelper.SqlQuery(sqlConnection, dynamicParameters, sqlQuery).ToList();
                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        draw,
                        recordsTotal = result.Select(x => x.TotalCount).FirstOrDefault(),
                        recordsFiltered = result.Select(x => x.TotalCount).FirstOrDefault(),
                        data = result,
                        status = "success"
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

        #region //GetPrDepartmentERP --取得請購部門 -- Chia Yuan 2025-4-14
        public string GetPrDepartmentERP(string CompanyNo, string DepartmentNo, string PrUser, string PoUser, string SearchKey)
        {
            try
            {
                string ErpNo = string.Empty;

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //確認公司別DB
                    sql = @"SELECT a.ErpNo, a.ErpDb
                            FROM BAS.Company a
                            WHERE CompanyNo = @CompanyNo";
                    var resultCorp = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo = CompanyNo == "ETG" ? "Eterge" : CompanyNo }) ?? throw new SystemException("【公司別】資料錯誤!");
                    #endregion

                    ErpConnectionStrings = ConfigurationManager.AppSettings[resultCorp.ErpDb];
                    ErpNo = resultCorp.ErpNo;
                }

                using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                {
                    #region //取得採購人
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DISTINCT LTRIM(RTRIM(a.TB013)) PoUser
                            , LTRIM(RTRIM(f.MF002)) PoUserName
                            , LTRIM(RTRIM(f.MF001)) + ' ' + LTRIM(RTRIM(f.MF002)) PoUserWithNo
                            , LTRIM(RTRIM(b.TA004)) DepartmentNo
                            , LTRIM(RTRIM(e.ME002)) DepartmentName
                            , LTRIM(RTRIM(e.ME001)) + ' ' + LTRIM(RTRIM(e.ME002)) DepartmentWithNo
                            FROM PURTB a
                            INNER JOIN PURTA b ON a.TB001 = b.TA001 AND a.TB002 = b.TA002
                            INNER JOIN CMSME e ON e.ME001 = b.TA004
                            INNER JOIN ADMMF f ON LTRIM(RTRIM(f.MF001)) = LTRIM(RTRIM(a.TB013))
                            WHERE 1=1";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TA004", " AND LTRIM(RTRIM(b.TA004)) = @TA004", DepartmentNo);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TA012", " AND LTRIM(RTRIM(b.TA012)) = @TA012", PrUser);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TB013", " AND LTRIM(RTRIM(a.TB013)) = @TB013", PoUser);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SearchKey",
                        @" AND (b.TA004 LIKE N'%' + @SearchKey + '%'
                            OR e.ME002 LIKE N'%' + @SearchKey + '%')", SearchKey);
                    var result = erpConnection.Query(sql, dynamicParameters).ToList();
                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        data = result,
                        status = "success"
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

        #region //GetPrUserERP --取得請購人員 -- Chia Yuan 2025-4-14
        public string GetPrUserERP(string CompanyNo, string UserNo, string LockStatus, string SearchKey)
        {
            try
            {
                string ErpNo = string.Empty;

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //確認公司別DB
                    sql = @"SELECT a.ErpNo, a.ErpDb
                            FROM BAS.Company a
                            WHERE CompanyNo = @CompanyNo";
                    var resultCorp = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo = CompanyNo == "ETG" ? "Eterge" : CompanyNo }) ?? throw new SystemException("【公司別】資料錯誤!");
                    #endregion

                    ErpConnectionStrings = ConfigurationManager.AppSettings[resultCorp.ErpDb];
                    ErpNo = resultCorp.ErpNo;
                }

                using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                {
                    #region //取得資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DISTINCT LTRIM(RTRIM(a.TA012)) PrUser
                            , LTRIM(RTRIM(f.MF002)) PrUserName
                            , LTRIM(RTRIM(a.TA012)) + ' ' + LTRIM(RTRIM(f.MF002)) PrUserWithNo
                            FROM PURTA a
                            INNER JOIN ADMMF f ON LTRIM(RTRIM(f.MF001)) = LTRIM(RTRIM(a.TA012))
                            INNER JOIN DSCSYS.dbo.DSCMA g ON LTRIM(RTRIM(g.MA001)) = LTRIM(RTRIM(a.TA012))
                            WHERE 1=1";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MA008", " AND g.MA008 = @MA008", LockStatus);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TA012", " AND LTRIM(RTRIM(a.TA012)) = @TA012", UserNo);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SearchKey",
                        @" AND (a.TA012 LIKE N'%' + @SearchKey + '%'
                            OR f.MF002 LIKE N'%' + @SearchKey + '%')", SearchKey);
                    var result = erpConnection.Query(sql, dynamicParameters).ToList();
                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        data = result,
                        status = "success"
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

        #region //GetPoUserERP --取得採購人員 -- Shintokuro 2025-03-17
        public string GetPoUserERP(string CompanyNo, string UserNo, string LockStatus, string SearchKey)
        {
            try
            {
                string ErpNo = string.Empty;

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //確認公司別DB
                    sql = @"SELECT a.ErpNo, a.ErpDb
                            FROM BAS.Company a
                            WHERE CompanyNo = @CompanyNo";
                    var resultCorp = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo = CompanyNo == "ETG" ? "Eterge" : CompanyNo }) ?? throw new SystemException("【公司別】資料錯誤!");
                    #endregion

                    ErpConnectionStrings = ConfigurationManager.AppSettings[resultCorp.ErpDb];
                    ErpNo = resultCorp.ErpNo;
                }

                using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                {
                    #region //取得採購人
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DISTINCT LTRIM(RTRIM(a.TB013)) PoUser
                            , LTRIM(RTRIM(f.MF002)) PoUserName
                            , LTRIM(RTRIM(a.TB013)) + ' ' + LTRIM(RTRIM(f.MF002)) PoUserWithNo
                            FROM PURTB a
                            INNER JOIN ADMMF f ON LTRIM(RTRIM(f.MF001)) = LTRIM(RTRIM(a.TB013))
                            INNER JOIN DSCSYS.dbo.DSCMA g ON LTRIM(RTRIM(g.MA001)) = LTRIM(RTRIM(a.TB013))
                            WHERE 1=1";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MA008", " AND g.MA008 = @MA008", LockStatus);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TB013", " AND LTRIM(RTRIM(a.TB013)) = @TB013", UserNo);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SearchKey",
                        @" AND (a.TB013 LIKE N'%' + @SearchKey + '%'
                            OR f.MF002 LIKE N'%' + @SearchKey + '%')", SearchKey);
                    var result = erpConnection.Query(sql, dynamicParameters).ToList();
                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        data = result,
                        status = "success"
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

        #region //GetPrSupplierErp --取得廠商 -- Shintokuro 2025-03-17
        public string GetPrSupplierErp(string SupplierNo, string SupplierName, string CompanyNo)
        {
            try
            {
                string ErpNo = string.Empty;

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //確認公司別DB
                    sql = @"SELECT a.ErpNo, a.ErpDb
                            FROM BAS.Company a
                            WHERE CompanyNo = @CompanyNo";
                    var resultCorp = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo = CompanyNo == "ETG" ? "Eterge" : CompanyNo }) ?? throw new SystemException("【公司別】資料錯誤!");
                    #endregion

                    ErpConnectionStrings = ConfigurationManager.AppSettings[resultCorp.ErpDb];
                    ErpNo = resultCorp.ErpNo;
                }

                using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                {
                    #region //取得資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DISTINCT LTRIM(RTRIM(a.TB010)) SupplierNo 
                            , LTRIM(RTRIM(a.TB010)) + ' ' + LTRIM(RTRIM(f.MA003)) SupplierWithNo
                            , LTRIM(RTRIM(f.MA013)) ContactPersonFirst
                            FROM PURTB a
                            INNER JOIN PURMA f ON LTRIM(RTRIM(f.MA001)) = LTRIM(RTRIM(a.TB010))
                            WHERE 1=1";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TB010", " AND LTRIM(RTRIM(a.TB010)) = @TB010", SupplierNo);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MA003", " AND LTRIM(RTRIM(f.MA003)) = @MA003", SupplierName);
                    var result = erpConnection.Query(sql, dynamicParameters).ToList();
                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        data = result,
                        status = "success"
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

        #region //GetPoSupplierERP --取得廠商 -- Shintokuro 2025-03-17
        public string GetPoSupplierERP(string SupplierNo, string SupplierName, string CompanyNo)
        {
            try
            {
                string ErpNo = string.Empty;

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //確認公司別DB
                    sql = @"SELECT a.ErpNo, a.ErpDb
                            FROM BAS.Company a
                            WHERE CompanyNo = @CompanyNo";
                    var resultCorp = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo = CompanyNo == "ETG" ? "Eterge" : CompanyNo }) ?? throw new SystemException("【公司別】資料錯誤!");
                    #endregion

                    ErpConnectionStrings = ConfigurationManager.AppSettings[resultCorp.ErpDb];
                    ErpNo = resultCorp.ErpNo;
                }

                using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                {
                    #region //取得資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT DISTINCT LTRIM(RTRIM(a.TC004)) SupplierNo 
                            , LTRIM(RTRIM(a.TC004)) + ' ' + LTRIM(RTRIM(f.MA003)) SupplierWithNo
                            , LTRIM(RTRIM(f.MA013)) ContactPersonFirst
                            FROM PURTC a
                            INNER JOIN PURMA f ON LTRIM(RTRIM(f.MA001)) = LTRIM(RTRIM(a.TC004))
                            WHERE 1=1";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TC004", " AND LTRIM(RTRIM(a.TC004)) = @TC004", SupplierNo);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MA003", " AND LTRIM(RTRIM(f.MA003)) = @MA003", SupplierName);
                    var result = erpConnection.Query(sql, dynamicParameters).ToList();
                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        data = result,
                        status = "success"
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

        #region //GetUnitERP --取得單位 -- Shintokuro 2025-03-17
        public string GetUnitERP(string UomNo, string UomName, string UomType, string CompanyNo)
        {
            try
            {
                string ErpNo = string.Empty;

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //確認公司別DB
                    sql = @"SELECT a.ErpNo, a.ErpDb
                            FROM BAS.Company a
                            WHERE CompanyNo = @CompanyNo";
                    var resultCorp = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo = CompanyNo == "ETG" ? "Eterge" : CompanyNo }) ?? throw new SystemException("【公司別】資料錯誤!");
                    #endregion

                    ErpConnectionStrings = ConfigurationManager.AppSettings[resultCorp.ErpDb];
                    ErpNo = resultCorp.ErpNo;
                }

                using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                {
                    #region //取得資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT LTRIM(RTRIM(MX001)) UomNo, LTRIM(RTRIM(MX002)) UomName,
                            MX001 + ' ' + MX002 UnitWithNo
                            FROM INVMX 
                            WHERE 1=1";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MX001", " AND LTRIM(RTRIM(MX001)) = @MX001", UomNo);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MX002", " AND LTRIM(RTRIM(MX002)) = @MX002", UomName);
                    var result = erpConnection.Query(sql, dynamicParameters).ToList();
                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        data = result,
                        status = "success"
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

        #region //GetInventoryERP --取得庫別 -- Shintokuro 2025-03-17
        public string GetInventoryERP(string InventoryNo, string InventoryName, string CompanyNo, string InventoryType)
        {
            try
            {
                string ErpNo = string.Empty;

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //確認公司別DB
                    sql = @"SELECT a.ErpNo, a.ErpDb
                            FROM BAS.Company a
                            WHERE CompanyNo = @CompanyNo";
                    var resultCorp = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo = CompanyNo == "ETG" ? "Eterge" : CompanyNo }) ?? throw new SystemException("【公司別】資料錯誤!");
                    #endregion

                    ErpConnectionStrings = ConfigurationManager.AppSettings[resultCorp.ErpDb];
                    ErpNo = resultCorp.ErpNo;
                }

                using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                {
                    #region //確認公司別DB
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT LTRIM(RTRIM(MC001)) InventoryNo
                            , LTRIM(RTRIM(MC002)) InventoryName
                            , LTRIM(RTRIM(MC001)) + ' ' + LTRIM(RTRIM(MC002)) InventoryWithNo
                            FROM CMSMC
                            WHERE 1=1";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MC001", " AND LTRIM(RTRIM(MC001)) = @MC001", InventoryNo);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MC002", " AND LTRIM(RTRIM(MC002)) = @MC002", InventoryName);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MC004", " AND MC004 = @MC004", InventoryType);
                    var result = erpConnection.Query(sql, dynamicParameters).ToList();
                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        data = result,
                        status = "success"
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

        #region //GetCurrencyErp --取得幣別 -- Shintokuro 2025-03-17
        public string GetCurrencyErp(string CompanyNo)
        {
            try
            {
                string ErpNo = string.Empty;

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //確認公司別DB
                    sql = @"SELECT a.ErpNo, a.ErpDb
                            FROM BAS.Company a
                            WHERE CompanyNo = @CompanyNo";
                    var resultCorp = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo = CompanyNo == "ETG" ? "Eterge" : CompanyNo }) ?? throw new SystemException("【公司別】資料錯誤!");
                    #endregion

                    ErpConnectionStrings = ConfigurationManager.AppSettings[resultCorp.ErpDb];
                    ErpNo = resultCorp.ErpNo;
                }

                using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                {
                    #region //取得幣別
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT LTRIM(RTRIM(MF001)) Currency, LTRIM(RTRIM(MF002)) CurrencyName
                            FROM CMSMF";
                    var result = erpConnection.Query(sql, dynamicParameters).ToList();
                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        data = result,
                        status = "success"
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

        #region //GetPoErpPrefix --取得採購單別 -- Shintokuro 2025-03-21
        public string GetPoErpPrefix(string CompanyNo)
        {
            try
            {
                string ErpNo = string.Empty;

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //確認公司別DB
                    sql = @"SELECT a.ErpNo, a.ErpDb
                            FROM BAS.Company a
                            WHERE CompanyNo = @CompanyNo";
                    var resultCorp = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo = CompanyNo == "ETG" ? "Eterge" : CompanyNo }) ?? throw new SystemException("【公司別】資料錯誤!");
                    #endregion

                    ErpConnectionStrings = ConfigurationManager.AppSettings[resultCorp.ErpDb];
                    ErpNo = resultCorp.ErpNo;
                }

                using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                {
                    #region //取得資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT LTRIM(RTRIM(MQ001)) PoErpPrefix,LTRIM(RTRIM(MQ001)) + ' ' +LTRIM(RTRIM(MQ002)) PoErpPrefixName
                            FROM CMSMQ
                            WHERE MQ003 = '33'";
                    var result = erpConnection.Query(sql, dynamicParameters).ToList();
                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        data = result,
                        status = "success"
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

        #region//AddPoByPrBatch --請購轉採購(批量) -- Chia Yuan 2025.4.15
        public string AddPoByPrBatch(string CompanyNo
            , string PoErpPrefix, string PoDate, string SupplierNo
            //, string PoUser, string Currency, string CurrencyLocal, string Inventory, string LockStatus
            , List<PrDetail> PrItems, string CurrentUser)
        {
            try
            {
                if (CompanyNo.Length <= 0) throw new SystemException("【公司】不能為空!");
                if (CurrentUser.Length <= 0) throw new SystemException("【使用者】不能為空!");
                if (PoErpPrefix.Length <= 0) throw new SystemException("【採購單別】不能為空!");
                if (PoDate.Length <= 0) throw new SystemException("【採購日】不能為空!");
                if (!DateTime.TryParse(PoDate, out DateTime poDate))
                    throw new SystemException("【採購日】格式錯誤!");

                PoDate = poDate.ToString("yyyyMMdd"); //轉為ERP日期格式
                string dateNow = CreateDate.ToString("yyyyMMdd"), timeNow = CreateDate.ToString("HH:mm:ss");

                List<PURTC> PurchaseOrder = new List<PURTC>();
                List<PURTD> PoDetails = new List<PURTD>();

                int rowsAffected = 0;

                using (TransactionScope transactionScope = new TransactionScope(TransactionScopeOption.Required, new TimeSpan(0, 180, 0)))
                {
                    string ErpNo = "";
                    List<string> purt = new List<string>(); //採購單號List

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認公司別DB
                        sql = @"SELECT TOP 1 a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo = CompanyNo == "ETG" ? "Eterge" : CompanyNo }) ?? throw new SystemException("【公司別】格式錯誤!");
                        #endregion

                        ErpConnectionStrings = ConfigurationManager.AppSettings[result.ErpDb];
                        ErpNo = result.ErpNo;

                        #region//判斷使用者是否存在
                        sql = @"SELECT a.UserNo
                                FROM BAS.[User] a
                                WHERE a.UserNo = @CurrentUser";
                        result = sqlConnection.Query(sql, new { CurrentUser }) ?? throw new SystemException("採購單【使用者】資料錯誤!");
                        #endregion
                    }

                    using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        #region //判斷ERP使用者是否存在
                        sql = @"SELECT TOP 1 MF004, MF005
                                FROM ADMMF
                                WHERE COMPANY = @ErpNo
                                AND MF001 = @CurrentUser";
                        var result = erpConnection.QueryFirstOrDefault(sql, new { ErpNo, CurrentUser }) ?? throw new SystemException("【ERP使用者】不存在!");
                        #endregion

                        string USR_GROUP = result.MF004;

                        #region //判斷採購單別是否存在
                        sql = @"SELECT TOP 1 1
                                FROM CMSMQ
                                WHERE MQ003 = '33'
                                AND MQ001 = @PoErpPrefix";
                        result = erpConnection.QueryFirstOrDefault(sql, new { PoErpPrefix }) ?? throw new SystemException("【採購單別】不存在!");
                        #endregion

                        #region //判斷採購日期是否超過結帳日
                        DateTime referenceTime = poDate;
                        sql = @"SELECT TOP 1 MA011,MA012,MA013
                                FROM CMSMA";
                        result = erpConnection.QueryFirstOrDefault(sql);

                        if (!DateTime.TryParseExact(result.MA013, "yyyyMMdd", null, DateTimeStyles.None, out DateTime closeDateBase))
                            throw new SystemException("【ERP結帳日】格式錯誤!");

                        if (poDate.CompareTo(closeDateBase) <= 0)
                        {
                            throw new SystemException("【銷貨單】ERP已經結帳,不可以在新增【" + closeDateBase.ToString("yyyy-MM-dd") + "】之前的單據");
                        }
                        #endregion

                        string ContactUser = string.Empty;
                        if (SupplierNo.Length > 0) 
                        {
                            sql = @"SELECT TOP 1 MA013 FROM PURMA WHERE MA001 = @SupplierNo";
                            result = erpConnection.QueryFirstOrDefault(sql, new { SupplierNo }) ?? throw new SystemException("【供應商】資料錯誤!"); ;
                            ContactUser = result.MA013; //聯絡人
                        }

                        //撈取單據編號編碼格式設定
                        string WoType = "";
                        string encode = ""; // 編碼格式
                        int yearLength = 0; // 年碼數
                        int lineLength = 0; // 流水碼數

                        #region //撈取ERP單據設定
                        sql = @"SELECT a.MQ004, a.MQ005, a.MQ006, a.MQ044, a.MQ017
                                FROM CMSMQ a
                                WHERE MQ001 = @PoErpPrefix";
                        result = erpConnection.QueryFirstOrDefault(sql, new { PoErpPrefix });
                        encode = result.MQ004; //編碼方式
                        yearLength = Convert.ToInt32(result.MQ005); //年碼數
                        lineLength = Convert.ToInt32(result.MQ006); //流水號碼數
                        WoType = result.MQ017;
                        #endregion

                        #region //單號自動取號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT CONVERT(int, RIGHT(ISNULL(MAX(RTRIM(LTRIM(TC002))), '" + new string('0', lineLength) + @"'), " + lineLength + @")) + 1 CurrentNum
                                FROM PURTC
                                WHERE TC001 = @PoErpPrefix";
                        dynamicParameters.Add("PoErpPrefix", PoErpPrefix);
                        #endregion

                        #region //編碼格式相關
                        string dateFormat = "";
                        string PoErpNo = "";
                        string PoErpNoBase = "";
                        switch (encode)
                        {
                            case "1": //日編
                                if (yearLength < 0 || yearLength > 4) throw new SystemException("【ERP單據性質】年碼數不能小於等於0或大於4");
                                if ((lineLength + yearLength + 4) > 11) throw new SystemException("【ERP單據性質】編碼格式碼數不能大於11碼");
                                dateFormat = new string('y', yearLength) + "MMdd";
                                sql += @" AND RTRIM(LTRIM(TC002)) LIKE @ReferenceTime  + '" + new string('_', lineLength) + @"'";
                                dynamicParameters.Add("ReferenceTime", referenceTime.ToString(dateFormat));

                                string tedstNo = referenceTime.ToString(dateFormat);
                                PoErpNo = referenceTime.ToString(dateFormat);
                                break;
                            case "2": //月編
                                if (yearLength < 0 || yearLength > 4) throw new SystemException("【ERP單據性質】年碼數不能小於等於0或大於4");
                                if ((lineLength + yearLength + 2) > 11) throw new SystemException("【ERP單據性質】編碼格式碼數不能大於11碼");
                                dateFormat = new string('y', yearLength) + "MM";
                                sql += @" AND RTRIM(LTRIM(TC002))  LIKE @ReferenceTime + '" + new string('_', lineLength) + @"'";
                                dynamicParameters.Add("ReferenceTime", referenceTime.ToString(dateFormat));
                                PoErpNo = referenceTime.ToString(dateFormat);
                                break;
                            case "3": //流水號
                                if (yearLength == 0) throw new SystemException("【ERP單據性質】年碼數必須等於0");
                                if (lineLength <= 0 || lineLength > 11) throw new SystemException("【ERP單據性質】流水編碼碼數必須大於0小於等於11");
                                break;
                            case "4": //手動編號
                                break;
                            default:
                                throw new SystemException("編碼方式錯誤!");
                        }
                        #endregion

                        int currentNum = erpConnection.QueryFirstOrDefault(sql, dynamicParameters).CurrentNum;
                        PoErpNoBase = PoErpNo;
                        PoErpNo += string.Format("{0:" + new string('0', lineLength) + "}", currentNum);

                        #region //判斷ERP採購單號是否重複 (停用)
                        //sql = @"SELECT TOP 1 1
                        //        FROM PURTC
                        //        WHERE TC001 = @PoErpPrefix
                        //        AND TC002 = @PoErpNo";
                        //result = erpConnection.QueryFirstOrDefault(sql, new { PoErpPrefix, PoErpNo });
                        //if (result != null) throw new SystemException("【採購單號】重複，請重新取號!");
                        #endregion

                        var PrKey = PrItems.Select(x => x.PrErpPrefix + x.PrErpNo + x.PrSeq).Distinct().ToArray();

                        //TB010 供應廠商
                        //TB013 採購人員
                        //TB016 採購幣別
                        //TB017 採購單價
                        //TB008 交貨庫別
                        //TB026 課稅別 1.應稅內含、2.應稅外加、3.零稅率、4.免稅、9.不計稅
                        //TB044 請購匯率
                        //TB050	請購幣別
                        //TB063	營業稅率
                        //TB064 單身多稅率

                        //TC004 供應廠商
                        //TC005 交易幣別
                        //TC006 匯率
                        //TC011 採購人員
                        //TC018 課稅別 1.應稅內含、2.應稅外加、3.零稅率、4.免稅、9.不計稅
                        //TC019 採購金額
                        //TC020 稅額
                        //TC026 營業稅率
                        //TC047 稅別碼
                        //TC051 單身多稅率

                        //TD010 採購單價
                        //TD011 採購金額
                        //TD063 折扣金額

                        #region //撈取請購單資料
                        sql = @"SELECT a.TB010 TC004, a.TB016 TC005
                                , (
                                    CASE WHEN a.TB016 ='NT%' THEN 1 
                                    ELSE  (
                                        SELECT TOP 1 MG004 FROM CMSMG
                                        WHERE MG001= a.TB016 AND MG002<=@PoDate
                                    ) END
                                ) TC006
                                , b.MA026 TC007, c.NA003 TC008, x.TC010
                                , a.TB013 TC011, b.MA057 TC012, x.TC021
                                , a.TB026 TC018, a.TB063 TC026, c.NA002 TC027
                                , b.MA058 TC028, a.TB057 TC047, a.TB058 TC048
                                , a.TB064 TC051, b.MA013 TC052
                                , a.TB004 TD004, a.TB005 TD005, a.TB006 TD006
                                , a.TB015 TD009, a.TB008 TD007, a.TB014 TD008
                                , a.TB017 TD010, a.TB018 TD011, a.TB011 TD012
                                , a.TB012 TD014, a.TB043 TD022, a.TB032 TD025, a.TB001 TD026
                                , a.TB002 TD027, a.TB003 TD028, a.TB035 TD030, a.TB038 TD032
                                , '0000' TD041, a.TB011 TD045, a.TB011 TD046, a.TB047 TD042, a.TB048 TD043
                                , a.TB063 TD057, a.TB065 TD058, a.TB066 TD059, a.TB057 TD061
                                FROM PURTB a
                                INNER JOIN PURMA b ON a.TB010 = b.MA001
                                INNER JOIN CMSNA c ON b.MA055 = c.NA002
                                OUTER APPLY(
                                    SELECT LTRIM(RTRIM(MB001)) TC010,LTRIM(RTRIM(MB005)) TC021 FROM CMSMB
                                ) x
                                WHERE LTRIM(RTRIM(a.TB001)) + LTRIM(RTRIM(a.TB002)) + LTRIM(RTRIM(a.TB003)) IN @PrKey
                                ORDER BY a.TB010";
                        var resultPURTC = erpConnection.Query<PURTC>(sql, new { PrKey, PoDate }).ToList();
                        if (!resultPURTC.Any()) throw new SystemException("請購單資料找不到，請重新確認!");
                        PoDetails = erpConnection.Query<PURTD>(sql, new { PrKey, PoDate }).ToList();

                        // TC004 供應廠商, TC005 交易幣別, TC011 採購人員, TC047 稅別碼, TC048 交易條件

                        foreach (var item in resultPURTC.GroupBy(g => new { g.TC004, g.TC005, g.TC011, g.TC047, g.TC048, g.TC051 }).Select(s => new { s.Key }))
                        {
                            var PURTC = resultPURTC
                                .FirstOrDefault(f => f.TC004 == item.Key.TC004 && f.TC005 == item.Key.TC005 && f.TC011 == item.Key.TC011
                                    && f.TC047 == item.Key.TC047 && f.TC048 == item.Key.TC048 && f.TC051 == item.Key.TC051);

                            double BankSellingRate = 1;
                            int UnitPriceDecimal = 0, AmountDecimal = 0, UnitCostDecimal = 0, CostAmountDecimal = 0;

                            #region //外幣處理
                            var QueryCMSMG = GetExchangeRate(new List<SupplierCompany>() { new SupplierCompany { CompanyNo = CompanyNo } }, item.Key.TC005, poDate.ToString("yyyy-MM-dd"), "", false, -1, true, "", -1, -1);
                            var resultStatus = JObject.Parse(QueryCMSMG)["status"]?.ToString();
                            if (resultStatus == "success")
                            {
                                JObject.Parse(QueryCMSMG)["data"].ToList().ForEach(x =>
                                {
                                    double.TryParse(x["BankSellingRate"].ToString(), out BankSellingRate);
                                    int.TryParse(x["UnitPriceDecimal"].ToString(), out UnitPriceDecimal);
                                    int.TryParse(x["AmountDecimal"].ToString(), out AmountDecimal);
                                    int.TryParse(x["UnitCostDecimal"].ToString(), out UnitCostDecimal);
                                    int.TryParse(x["CostAmountDecimal"].ToString(), out CostAmountDecimal);
                                });
                            }
                            #endregion

                            int num = 1;
                            decimal totalAmount = 0m;
                            decimal totalTax = 0m;
                            double totalCount = 0d;
                            //double taxRate = PURTC.TC026;

                            PoDetails.Where(w => w.TC004 == item.Key.TC004 && w.TC005 == item.Key.TC005 && w.TC011 == item.Key.TC011 && w.TC047 == item.Key.TC047 && w.TC048 == item.Key.TC048).ToList()
                                .ForEach(x =>
                                {
                                    x.TD001 = PoErpPrefix;
                                    x.TD002 = PoErpNo;
                                    x.TD003 = string.Format("{0:0000}", num);

                                    decimal tax = 0m, amount = 0m;
                                    if (PURTC.TC051 == "Y") //單身多稅率
                                    {
                                        x.TD057 = PURTC.TC051 == "Y" ? x.TD057 : PURTC.TC026; //稅率                                        
                                        x.TD061 = PURTC.TC051 == "Y" ? x.TD061 : PURTC.TC047; //稅別碼
                                    }

                                    //double exchangeRate = x.TC006 / BankSellingRate;
                                    //decimal poUnitPriceLocal = Math.Round(x.TD010 * (decimal)x.TD008 * (decimal)x.TC006, UnitPriceDecimal, MidpointRounding.AwayFromZero);
                                    //decimal poPriceLocal = Math.Round(x.TD011 * (decimal)x.TC006, AmountDecimal, MidpointRounding.AwayFromZero);
                                    //x.TD010 = Math.Round(poUnitPriceLocal / (decimal)exchangeRate, UnitPriceDecimal, MidpointRounding.AwayFromZero);
                                    //x.TD011 = Math.Round(poPriceLocal / (decimal)exchangeRate, AmountDecimal, MidpointRounding.AwayFromZero);

                                    //x.TD010 = Math.Round(x.TD010 * (decimal)x.TD008, UnitPriceDecimal, MidpointRounding.AwayFromZero); //採購單價
                                    //x.TD011 = Math.Round(x.TD010 * (decimal)x.TD008, AmountDecimal, MidpointRounding.AwayFromZero); //採購金額

                                    #region //稅額計算
                                    if (x.TD011 > 0)
                                    {
                                        switch (PURTC.TC018)
                                        {
                                            case "1":
                                                tax = Math.Round(x.TD011 - x.TD011 / (1 + Convert.ToDecimal(x.TD057)), AmountDecimal, MidpointRounding.AwayFromZero);
                                                amount = x.TD011 - tax; //採購金額
                                                break;
                                            case "2":
                                                tax = Math.Round(x.TD011 * Convert.ToDecimal(x.TD057), AmountDecimal, MidpointRounding.AwayFromZero);
                                                amount = x.TD011; //採購金額
                                                break;
                                        }
                                    }
                                    #endregion

                                    totalAmount += amount; //採購金額
                                    totalTax += tax; //稅額
                                    totalCount += x.TD008; //採購數量

                                    #region //請購的採購資料更新 PURTB
                                    sql = @"UPDATE PURTB SET 
                                            MODIFIER = @CurrentUser,
                                            MODI_DATE = @dateNow,
                                            MODI_TIME = @timeNow,
                                            MODI_AP = @MODI_AP,
                                            MODI_PRID = @MODI_PRID,
                                            FLAG = FLAG + 1,
                                            TB021 = 'Y',
                                            TB022 = @TB022,
                                            TB039 = 'Y'
                                            WHERE TB001 = @TD026
                                            AND TB002 = @TD027
                                            AND TB003 = @TD028";
                                    rowsAffected += erpConnection.Execute(sql,
                                        new
                                        {
                                            CurrentUser,
                                            dateNow,
                                            timeNow,
                                            MODI_AP = CurrentUser + "PC",
                                            MODI_PRID = "SRM",
                                            TB022 = x.TD001 + "-" + x.TD002 + "-" + x.TD003,
                                            x.TD026,//請購單別
                                            x.TD027,//請購單號
                                            x.TD028 //請購序號
                                        });
                                    #endregion

                                    #region //請購的採購資料更新 PURTY
                                    sql = @"UPDATE PURTY SET 
                                            MODIFIER = @CurrentUser,
                                            MODI_DATE = @dateNow,
                                            MODI_TIME = @timeNow,
                                            MODI_AP = @MODI_AP,
                                            MODI_PRID = @MODI_PRID,
                                            FLAG = FLAG + 1,
                                            TY013 = @TY013,
                                            TY021 = 'Y',
                                            TY022 = 'Y'
                                            WHERE TY001 = @TD026
                                            AND TY002 = @TD027
                                            AND TY003 = @TD028";
                                    rowsAffected += erpConnection.Execute(sql,
                                        new
                                        {
                                            CurrentUser,
                                            dateNow,
                                            timeNow,
                                            MODI_AP = CurrentUser + "PC",
                                            MODI_PRID = "SRM",
                                            TY013 = x.TD001 + "-" + x.TD002 + "-" + x.TD003,
                                            x.TD026,//請購單別
                                            x.TD027,//請購單號
                                            x.TD028 //請購序號
                                        });
                                    #endregion

                                    num++;
                                });

                            #region //加入採購單頭資料 PURTC
                            PurchaseOrder.Add(new PURTC
                            {
                                TC001 = PoErpPrefix,
                                TC002 = PoErpNo,
                                TC003 = PoDate,
                                TC004 = SupplierNo.Length > 0 ? SupplierNo : PURTC.TC004 , //供應商,
                                TC005 = PURTC.TC005, //幣別
                                TC006 = BankSellingRate, //匯率
                                TC007 = PURTC.TC007, //價格條件
                                TC008 = PURTC.TC008, //付款條件
                                TC009 = "", //備註
                                TC010 = PURTC.TC010, //廠別
                                TC011 = PURTC.TC011, //採購人員
                                TC012 = PURTC.TC012, //列印格式
                                TC013 = 0, //列印次數
                                TC014 = "N", //確認碼
                                TC015 = "", //P/I日期
                                TC016 = "", //P/I單號
                                TC017 = "", //運輸方式
                                TC018 = PURTC.TC018, //課稅別 1.應稅內含、2.應稅外加、3.零稅率、4.免稅、9.不計稅
                                TC019 = totalAmount, //採購總金額
                                TC020 = totalTax, //稅額
                                TC021 = PURTC.TC021, //送貨地址(一)
                                TC022 = "", //送貨地址(二)
                                TC023 = totalCount, //數量合計
                                TC024 = dateNow, //單據日期
                                TC025 = "", //確認者
                                TC026 = PURTC.TC026, //營業稅率
                                TC027 = PURTC.TC027, //付款條件代號
                                TC028 = PURTC.TC028, //訂金比率
                                TC029 = 0, //包裝數量合計
                                TC030 = "N", //簽核狀態
                                TC031 = 0, //傳送次數
                                TC032 = "", //流程代號
                                TC033 = "N", //拋轉狀態
                                TC034 = "", //下游廠商
                                TC035 = "N", //EBC確認碼
                                TC036 = "", //EBC採購單號
                                TC037 = "", //EBC採購版次
                                TC038 = "N", //匯至EBC
                                TC039 = "0000", //版次
                                TC040 = "N", //訂金分批
                                TC041 = "", //運輸方式代號
                                TC042 = 0, //預留欄位
                                TC043 = 0, //預留欄位
                                TC044 = "", //預留欄位
                                TC045 = "", //預留欄位
                                TC046 = "", //預留欄位
                                TC047 = PURTC.TC047, //稅別碼
                                TC048 = PURTC.TC048, //交易條件
                                TC049 = "", //製造廠商
                                TC050 = "", //鎖定碼
                                TC051 = PURTC.TC051, //單身多稅率
                                TC052 = SupplierNo.Length > 0 ? ContactUser : PURTC.TC052, //聯絡人
                                TC500 = "",
                                TC501 = "",
                                TC502 = "",
                                TC503 = "",
                                TC550 = ""
                            });
                            #endregion

                            PoErpNoBase += string.Format("{0:" + new string('0', lineLength) + "}", currentNum + 1);
                            PoErpNo = PoErpNoBase;
                        }
                        #endregion

                        #region //單頭資料新增 PURTC
                        sql = @"INSERT INTO PURTC(COMPANY,CREATOR,USR_GROUP,CREATE_DATE,MODIFIER,MODI_DATE,FLAG,CREATE_TIME,CREATE_AP,CREATE_PRID,MODI_TIME,MODI_AP,MODI_PRID
                                , TC001,TC002,TC003,TC004,TC005,TC006,TC007,TC008,TC009,TC010,TC011,TC012
                                , TC013,TC014,TC015,TC016,TC017,TC018,TC019,TC020,TC021,TC022,TC023,TC024,TC025
                                , TC026,TC027,TC028,TC029,TC030,TC031,TC032,TC033,TC034,TC035,TC036,TC037,TC038
                                , TC039,TC040,TC041,TC042,TC043,TC044,TC045,TC046,TC047,TC048,TC049,TC050,TC051
                                , TC052,TC500,TC502,TC503,TC550)
                                VALUES (@COMPANY,@CREATOR,@USR_GROUP,@CREATE_DATE,@MODIFIER,@MODI_DATE,@FLAG,@CREATE_TIME,@CREATE_AP,@CREATE_PRID,@MODI_TIME,@MODI_AP,@MODI_PRID
                                , @TC001,@TC002,@TC003,@TC004,@TC005,@TC006,@TC007,@TC008,@TC009,@TC010,@TC011,@TC012
                                , @TC013,@TC014,@TC015,@TC016,@TC017,@TC018,@TC019,@TC020,@TC021,@TC022,@TC023,@TC024,@TC025
                                , @TC026,@TC027,@TC028,@TC029,@TC030,@TC031,@TC032,@TC033,@TC034,@TC035,@TC036,@TC037,@TC038
                                , @TC039,@TC040,@TC041,@TC042,@TC043,@TC044,@TC045,@TC046,@TC047,@TC048,@TC049,@TC050,@TC051
                                , @TC052,@TC500,@TC502,@TC503,@TC550)";
                        PurchaseOrder
                        .ToList()
                        .ForEach(x =>
                        {
                            x.COMPANY = ErpNo;
                            x.USR_GROUP = USR_GROUP;
                            x.FLAG = "1";
                            x.CREATOR = CurrentUser;
                            x.CREATE_DATE = dateNow;
                            x.CREATE_TIME = timeNow;
                            x.CREATE_AP = CurrentUser + "PC";
                            x.CREATE_PRID = "SRM";
                            x.MODIFIER = CurrentUser;
                            x.MODI_DATE = dateNow;
                            x.MODI_TIME = timeNow;
                            x.MODI_AP = CurrentUser + "PC";
                            x.MODI_PRID = "SRM";
                        });
                        rowsAffected += erpConnection.Execute(sql, PurchaseOrder);
                        #endregion

                        #region //單身新增 PURTD
                        sql = @"INSERT INTO PURTD(COMPANY,CREATOR,USR_GROUP,CREATE_DATE,MODIFIER,MODI_DATE,FLAG,CREATE_TIME,CREATE_AP,CREATE_PRID,MODI_TIME,MODI_AP,MODI_PRID
                                , TD001,TD002,TD003,TD004,TD005,TD006,TD007,TD008,TD009,TD010,TD011,TD012
                                , TD013,TD014,TD015,TD016,TD017,TD018,TD019,TD020,TD021,TD022,TD023,TD024,TD025
                                , TD026,TD027,TD028,TD029,TD030,TD031,TD032,TD033,TD034,TD035,TD036,TD037,TD038
                                , TD039,TD040,TD041,TD042,TD043,TD044,TD045,TD046,TD047,TD048,TD049,TD050,TD051
                                , TD052,TD053,TD054,TD055,TD056,TD057,TD058,TD059,TD060,TD061,TD062,TD063,TD500,TD501,TD502,TD503
                                , TD550,TD551,TD552,TD553,TD554,TD555)
                                VALUES (@COMPANY,@CREATOR,@USR_GROUP,@CREATE_DATE,@MODIFIER,@MODI_DATE,@FLAG,@CREATE_TIME,@CREATE_AP,@CREATE_PRID,@MODI_TIME,@MODI_AP,@MODI_PRID
                                , @TD001,@TD002,@TD003,@TD004,@TD005,@TD006,@TD007,@TD008,@TD009,@TD010,@TD011,@TD012
                                , @TD013,@TD014,@TD015,@TD016,@TD017,@TD018,@TD019,@TD020,@TD021,@TD022,@TD023,@TD024,@TD025
                                , @TD026,@TD027,@TD028,@TD029,@TD030,@TD031,@TD032,@TD033,@TD034,@TD035,@TD036,@TD037,@TD038
                                , @TD039,@TD040,@TD041,@TD042,@TD043,@TD044,@TD045,@TD046,@TD047,@TD048,@TD049,@TD050,@TD051
                                , @TD052,@TD053,@TD054,@TD055,@TD056,@TD057,@TD058,@TD059,@TD060,@TD061,@TD062,@TD063,@TD500,@TD501,@TD502,@TD503
                                , @TD550,@TD551,@TD552,@TD553,@TD554,@TD555)";
                        PoDetails
                        .ToList()
                        .ForEach(x =>
                        {
                            x.COMPANY = ErpNo;
                            x.USR_GROUP = USR_GROUP;
                            x.FLAG = "1";
                            x.CREATOR = CurrentUser;
                            x.CREATE_DATE = dateNow;
                            x.CREATE_TIME = timeNow;
                            x.CREATE_AP = CurrentUser + "PC";
                            x.CREATE_PRID = "SRM";
                            x.MODIFIER = CurrentUser;
                            x.MODI_DATE = dateNow;
                            x.MODI_TIME = timeNow;
                            x.MODI_AP = CurrentUser + "PC";
                            x.MODI_PRID = "SRM";
                            x.TD013 = ""; //參考單別
                            //x.TD014 = ""; //備註
                            x.TD015 = 0; //已交數量
                            x.TD016 = "N"; //結案碼
                            x.TD017 = ""; //製造商
                            x.TD018 = "N"; //確認碼
                            x.TD019 = 0; //庫存數量
                            x.TD020 = ""; //小單位
                            x.TD021 = ""; //參考單號
                            //x.TD022 = ""; //專案代號
                            x.TD023 = ""; //參考序號
                            x.TD024 = ""; //來源單號
                            x.TD029 = ""; //承認型號
                            //x.TD030 = 0; //採購包裝數量
                            x.TD031 = 0; //已交包裝數量
                            //x.TD032 = ""; //包裝單位
                            x.TD033 = ""; //原始客戶
                            x.TD034 = "1"; //來源單據 1.請購 2.LRP 3.MRP 4.訂單 5.合約採購 6.採購變更 9.其他
                            x.TD035 = ""; //APS規劃採購號碼
                            x.TD036 = ""; //EBC採購單號
                            x.TD037 = ""; //EBC採購版次
                            x.TD038 = "2"; //類型 1.工程品號 2.正式品號
                            x.TD039 = ""; //產品圖號
                            x.TD040 = ""; //圖號版次
                            x.TD041 = "0001"; //分批序號
                            //x.TD042 = ""; //預算編號
                            //x.TD043 = ""; //預算科目
                            x.TD044 = ""; //預算部門
                            x.TD047 = ""; //預留欄位
                            x.TD048 = 0;  //預留欄位
                            x.TD049 = 0;  //預留欄位
                            x.TD050 = ""; //預留欄位
                            x.TD051 = ""; //預留欄位
                            x.TD052 = ""; //預留欄位
                            x.TD053 = ""; //預留欄位
                            x.TD054 = ""; //預留欄位
                            x.TD055 = ""; //預留欄位
                            x.TD056 = ""; //預留欄位
                            //x.TD057 = 0; //稅率
                            x.TD060 = 0; //已交計價數量
                            //x.TD061 = ""; //稅別碼
                            x.TD062 = 0;
                            x.TD063 = 0;
                            x.TD500 = "";
                            x.TD501 = "";
                            x.TD502 = "";
                            x.TD503 = "";
                            x.TD550 = "";
                            x.TD551 = 0;
                            x.TD552 = "";
                            x.TD553 = "";
                            x.TD554 = "";
                            x.TD555 = "";
                        });
                        rowsAffected += erpConnection.Execute(sql, PoDetails);
                        #endregion
                    }

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "(1 rows affected)"
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

        #region//UpdatePrDetailERP --更新ERP請購單身 -- Shinotokuro 2025.02.04
        public string UpdatePrDetailERP(string PrErpPrefix, string PrErpNo, string PrSeq
            , string UrgentMtl, string SupplierNo, string TradeTerm, string TaxCode, string Taxation, string PoCurrency, string PoCurrencyLocal
            , string PoUser, string Inventory, int PoQty, string PoUnit
            , int PoPriceQty, string PoPriceUnit, double PoUnitPrice, double PoPrice, string PoRemark
            , string CompanyNo, string CurrentUser)
        {
            try
            {
                if (PrErpPrefix.Length <= 0) throw new SystemException("【請購單別】不能為空!");
                if (PrErpNo.Length <= 0) throw new SystemException("【請購單號】不能為空!");
                if (PrSeq.Length <= 0) throw new SystemException("【請購序號】不能為空!");
                if (!Regex.IsMatch(UrgentMtl, "^(Y|N)$", RegexOptions.IgnoreCase)) throw new SystemException("【急料】核選錯誤!");
                string dateNow = CreateDate.ToString("yyyyMMdd"), timeNow = CreateDate.ToString("HH:mm:ss");
                int rowsAffected = 0;

                using (TransactionScope transactionScope = new TransactionScope(TransactionScopeOption.Required, new TimeSpan(0, 180, 0)))
                {
                    string ErpConnectionStrings = "", ErpNo = "";
                    List<string> purt = new List<string>(); //採購單號List

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認公司別DB
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo }) ?? throw new SystemException("請選擇公司別!");
                        #endregion

                        ErpConnectionStrings = ConfigurationManager.AppSettings[result.ErpDb];
                        ErpNo = result.ErpNo;

                        #region//判斷使用者是否存在
                        sql = @"SELECT TOP 1 a.UserNo
                                FROM BAS.[User] a
                                WHERE a.UserNo = @CurrentUser";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { CurrentUser }) ?? throw new SystemException("請購單【使用者】資料錯誤!");
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        #region //判斷ERP使用者是否存在
                        sql = @"SELECT TOP 1 MF004, MF005
                                FROM ADMMF
                                WHERE COMPANY = @ErpNo
                                AND MF001 = @CurrentUser";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { ErpNo, CurrentUser }) ?? throw new SystemException("【ERP使用者】不存在!");
                        #endregion

                        #region //判斷請購單是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 STUFF(STUFF(b.TA003, 5, 0, '-'), 8, 0, '-') PrDate
                                , a.TB025, a.TB020, a.TB021, a.TB039
                                FROM PURTB a
                                INNER JOIN PURTA b ON a.TB001 = b.TA001 AND a.TB002 = b.TA002
                                WHERE a.TB001 = @PrErpPrefix
                                AND a.TB002 = @PrErpNo
                                AND a.TB003 = @PrSeq";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { PrErpPrefix, PrErpNo, PrSeq }) ?? throw new SystemException("請購單查不到,請重新確認");
                        if (result.TB025 != "Y") throw new SystemException("請購單非核單狀態,請重新確認");
                        if (result.TB020 == "Y") throw new SystemException("請購單鎖定碼已鎖定,請重新確認");
                        if (result.TB021 == "Y") throw new SystemException("請購單採購碼已鎖定,請重新確認");
                        if (result.TB039 != "N") throw new SystemException("請購單已結案,請重新確認");
                        #endregion

                        string PrDate = result.PrDate;
                        double PrPriceLocal = PoPrice, BankSellingRate = 1;
                        int UnitPriceDecimal = 0, AmountDecimal = 0, UnitCostDecimal = 0, CostAmountDecimal = 0;

                        #region //外幣處理
                        var QueryCMSMG = GetExchangeRate(new List<SupplierCompany>() { new SupplierCompany { CompanyNo = CompanyNo} }, PoCurrency, PrDate, "", false, -1, true, "", -1, -1);
                        var resultStatus = JObject.Parse(QueryCMSMG)["status"]?.ToString();
                        if (resultStatus == "success") 
                        {
                            JObject.Parse(QueryCMSMG)["data"].ToList().ForEach(x =>
                            {
                                double.TryParse(x["BankSellingRate"].ToString(), out BankSellingRate);
                                int.TryParse(x["UnitPriceDecimal"].ToString(), out UnitPriceDecimal);
                                int.TryParse(x["AmountDecimal"].ToString(), out AmountDecimal);
                                int.TryParse(x["UnitCostDecimal"].ToString(), out UnitCostDecimal);
                                int.TryParse(x["CostAmountDecimal"].ToString(), out CostAmountDecimal);
                            });
                        }
                        #endregion

                        #region //本幣處理
                        var _Local = GetExchangeRate(new List<SupplierCompany>() { new SupplierCompany { CompanyNo = CompanyNo } }, PoCurrencyLocal, PrDate, "", false, -1, true, "", -1, -1);
                        if (JObject.Parse(_Local)["status"]?.ToString() == "success")
                        {
                            JObject.Parse(_Local)["data"].ToList().ForEach(x => 
                            {
                                int.TryParse(x["AmountDecimal"].ToString(), out AmountDecimal);
                                PrPriceLocal = Math.Round(PoPrice * BankSellingRate, AmountDecimal, MidpointRounding.AwayFromZero);
                            });
                        }
                        #endregion

                        #region //請購的採購資料更新 PURTB
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PURTB SET 
                                MODIFIER = @MODIFIER,
                                MODI_DATE = @dateNow,
                                MODI_TIME = @timeNow,
                                MODI_AP = @MODI_AP,
                                MODI_PRID = @MODI_PRID,
                                FLAG = FLAG + 1,
                                TB008 = @Inventory,
                                TB010 = @SupplierNo,
                                TB013 = @PoUser,
                                TB014 = @PoQty,
                                TB015 = @PoUnit,
                                TB016 = @PoCurrency,
                                TB017 = @PoUnitPrice,
                                TB018 = @PoPrice,
                                TB024 = @PoRemark,
                                TB026 = @Taxation,
                                TB032 = @UrgentMtl,
                                TB045 = @PrPriceLocal,
                                TB057 = @TaxCode,
                                TB058 = @TradeTerm,
                                TB065 = @PoPriceQty,
                                TB066 = @PoPriceUnit
                                WHERE TB001 = @PrErpPrefix
                                AND TB002 = @PrErpNo
                                AND TB003 = @PrSeq";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            MODIFIER = CurrentUser,
                            dateNow,
                            timeNow,
                            MODI_AP = CurrentUser + "PC",
                            MODI_PRID = "SRM",
                            SupplierNo,
                            TradeTerm,
                            TaxCode,
                            Taxation,
                            UrgentMtl,
                            PoCurrency,
                            PoUser,
                            Inventory,
                            PoQty,
                            PoUnit,
                            PoPriceQty,
                            PoPriceUnit,
                            PoUnitPrice,
                            PoPrice,
                            PoRemark,
                            PrPriceLocal,
                            PrErpPrefix,
                            PrErpNo,
                            PrSeq
                        });

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //請購的採購資料更新 PURTY
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PURTY SET 
                                MODIFIER = @MODIFIER,
                                MODI_DATE = @dateNow,
                                MODI_TIME = @timeNow,
                                MODI_AP = @MODI_AP,
                                MODI_PRID = @MODI_PRID,
                                FLAG = FLAG + 1,
                                TY005 = @Inventory,
                                TY006 = @SupplierNo,
                                TY007 = @PoQty,
                                TY008 = @PoCurrency,
                                TY009 = @PoUnitPrice,
                                TY010 = @PoPrice,
                                TY015 = @PoRemark,
                                TY016 = @PoUser,
                                TY017 = @Taxation,
                                TY018 = @UrgentMtl,
                                TY031 = @TaxCode,
                                TY032 = @TradeTerm,
                                TY041 = @PoPriceQty
                                WHERE TY001 = @PrErpPrefix
                                AND TY002 = @PrErpNo
                                AND TY003 = @PrSeq";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            MODIFIER = CurrentUser,
                            dateNow,
                            timeNow,
                            MODI_AP = CurrentUser + "PC",
                            MODI_PRID = "SRM",
                            SupplierNo,
                            TradeTerm,
                            TaxCode,
                            Taxation,
                            UrgentMtl,
                            PoCurrency,
                            PoUser,
                            Inventory,
                            PoQty,
                            PoPriceQty,
                            PoUnitPrice,
                            PoPrice,
                            PoRemark,
                            PrErpPrefix,
                            PrErpNo,
                            PrSeq
                        });

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                    }

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "(1 rows affected)"
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

        #region//UpdatePrDetailCloseERP --更新ERP請購單指定結案 -- Shinotokuro 2025.03.20
        public string UpdatePrDetailCloseERP(string PrErpPrefix, string PrErpNo, string PrSeq
            , string CompanyNo, string CurrentUser)
        {
            try
            {
                if (PrErpPrefix.Length <= 0) throw new SystemException("【請購單別】不能為空!");
                if (PrErpNo.Length <= 0) throw new SystemException("【請購單號】不能為空!");
                if (PrSeq.Length <= 0) throw new SystemException("【請購序號】不能為空!");

                string dateNow = CreateDate.ToString("yyyyMMdd"), timeNow = CreateDate.ToString("HH:mm:ss"), errorMsg = "";
                int rowsAffected = 0;

                using (TransactionScope transactionScope = new TransactionScope(TransactionScopeOption.Required, new TimeSpan(0, 180, 0)))
                {
                    string ErpConnectionStrings = "", ErpNo = "";
                    List<string> purt = new List<string>(); //採購單號List

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認公司別DB
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo }) ?? throw new SystemException("請選擇公司別!");
                        #endregion

                        ErpConnectionStrings = ConfigurationManager.AppSettings[result.ErpDb];
                        ErpNo = result.ErpNo;

                        #region//判斷使用者是否存在
                        sql = @"SELECT TOP 1 a.UserNo
                                FROM BAS.[User] a
                                WHERE a.UserNo = @CurrentUser";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { CurrentUser }) ?? throw new SystemException("請購單【使用者】資料錯誤!");
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        #region //判斷ERP使用者是否存在
                        sql = @"SELECT TOP 1 MF004, MF005
                                FROM ADMMF
                                WHERE COMPANY = @ErpNo
                                AND MF001 = @CurrentUser";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { ErpNo, CurrentUser }) ?? throw new SystemException("【ERP使用者】不存在!");
                        #endregion

                        #region //判斷請購單是否存在
                        sql = @"SELECT TOP 1 a.TB025,a.TB020,a.TB021, a.TB039
                                FROM PURTB a
                                INNER JOIN PURTA b ON a.TB001 = b.TA001 AND a.TB002 = b.TA002
                                WHERE a.TB001 = @PrErpPrefix
                                AND a.TB002 = @PrErpNo
                                AND a.TB003 = @PrSeq";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { PrErpPrefix, PrErpNo, PrSeq }) ?? throw new SystemException("請購單查不到,請重新確認");
                        if (result.TB025 != "Y") throw new SystemException("請購單非核單狀態,請重新確認");
                        if (result.TB020 == "Y") throw new SystemException("請購單鎖定碼已鎖定,請重新確認");
                        if (result.TB021 == "Y") throw new SystemException("請購單採購碼已鎖定,請重新確認");
                        if (result.TB039 != "N") throw new SystemException("請購單已結案,請重新確認");
                        #endregion

                        #region //請購的採購資料更新 PURTB
                        sql = @"UPDATE PURTB SET 
                                MODIFIER = @CurrentUser,
                                MODI_DATE = @MODI_DATE,
                                MODI_TIME = @MODI_TIME,
                                MODI_AP = @MODI_AP,
                                MODI_PRID = @MODI_PRID,
                                FLAG = FLAG + 1,
                                TB039 = 'y'
                                WHERE TB001 = @PrErpPrefix
                                AND TB002 = @PrErpNo
                                AND TB003 = @PrSeq";
                        rowsAffected = sqlConnection.Execute(sql,
                            new
                            {
                                CurrentUser,
                                MODI_DATE = dateNow,
                                MODI_TIME = timeNow,
                                MODI_AP = CurrentUser + "PC",
                                MODI_PRID = "SRM",
                                PrErpPrefix,
                                PrErpNo,
                                PrSeq
                            });
                        #endregion

                        #region //請購的採購資料更新 PURTY
                        sql = @"UPDATE PURTY SET 
                                MODIFIER = @CurrentUser,
                                MODI_DATE = @MODI_DATE,
                                MODI_TIME = @MODI_TIME,
                                MODI_AP = @MODI_AP,
                                MODI_PRID = @MODI_PRID,
                                FLAG = FLAG + 1,
                                TY021= 'y'
                                WHERE TY001 = @PrErpPrefix
                                AND TY002 = @PrErpNo
                                AND TY003 = @PrSeq";
                        rowsAffected = sqlConnection.Execute(sql,
                            new
                            {
                                CurrentUser,
                                MODI_DATE = dateNow,
                                MODI_TIME = timeNow,
                                MODI_AP = CurrentUser + "PC",
                                MODI_PRID = "SRM",
                                PrErpPrefix,
                                PrErpNo,
                                PrSeq
                            });
                        #endregion
                    }

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "(1 rows affected)"
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

        #region//UpdatePrDetailReCloseERP --更新ERP請購單反確認指定結案 -- Shinotokuro 2025.03.20
        public string UpdatePrDetailReCloseERP(string PrErpPrefix, string PrErpNo, string PrSeq
            , string CompanyNo, string CurrentUser)
        {
            try
            {
                if (PrErpPrefix.Length <= 0) throw new SystemException("【請購單別】不能為空!");
                if (PrErpNo.Length <= 0) throw new SystemException("【請購單號】不能為空!");
                if (PrSeq.Length <= 0) throw new SystemException("【請購序號】不能為空!");

                string dateNow = CreateDate.ToString("yyyyMMdd"), timeNow = CreateDate.ToString("HH:mm:ss"), errorMsg = "";
                int rowsAffected = 0;

                using (TransactionScope transactionScope = new TransactionScope(TransactionScopeOption.Required, new TimeSpan(0, 180, 0)))
                {
                    string ErpConnectionStrings = "", ErpNo = "";
                    List<string> purt = new List<string>(); //採購單號List

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認公司別DB
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo }) ?? throw new SystemException("請選擇公司別!");
                        #endregion

                        ErpConnectionStrings = ConfigurationManager.AppSettings[result.ErpDb];
                        ErpNo = result.ErpNo;

                        #region//判斷使用者是否存在
                        sql = @"SELECT TOP 1 a.UserNo
                                FROM BAS.[User] a
                                WHERE a.UserNo = @CurrentUser";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { CurrentUser }) ?? throw new SystemException("請購單【使用者】資料錯誤!");
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        #region //判斷ERP使用者是否存在
                        sql = @"SELECT TOP 1 MF004, MF005
                                FROM ADMMF
                                WHERE COMPANY = @ErpNo
                                AND MF001 = @CurrentUser";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { ErpNo, CurrentUser }) ?? throw new SystemException("【ERP使用者】不存在!");
                        #endregion

                        #region //判斷請購單是否存在
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.TB025,a.TB020,a.TB021, a.TB039
                                FROM PURTB a
                                INNER JOIN PURTA b ON a.TB001 = b.TA001 AND a.TB002 = b.TA002
                                WHERE a.TB001 = @PrErpPrefix
                                AND a.TB002 = @PrErpNo
                                AND a.TB003 = @PrSeq";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { PrErpPrefix, PrErpNo, PrSeq }) ?? throw new SystemException("請購單查不到,請重新確認");
                        if (result.TB025 != "Y") throw new SystemException("請購單非核單狀態,請重新確認");
                        if (result.TB020 == "Y") throw new SystemException("請購單鎖定碼已鎖定,請重新確認");
                        if (result.TB021 == "Y") throw new SystemException("請購單採購碼已鎖定,請重新確認");
                        if (result.TB039 != "N") throw new SystemException("請購單已結案,請重新確認");
                        #endregion

                        #region //請購的採購資料更新 PURTB
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE PURTB SET 
                                MODIFIER = @CurrentUser,
                                MODI_DATE = @MODI_DATE,
                                MODI_TIME = @MODI_TIME,
                                MODI_AP = @MODI_AP,
                                MODI_PRID = @MODI_PRID,
                                FLAG = FLAG + 1,
                                TB039 = 'N'
                                WHERE TB001 = @PrErpPrefix
                                AND TB002 = @PrErpNo
                                AND TB003 = @PrSeq";
                        rowsAffected = sqlConnection.Execute(sql,
                            new
                            {
                                CurrentUser,
                                MODI_DATE = dateNow,
                                MODI_TIME = timeNow,
                                MODI_AP = CurrentUser + "PC",
                                MODI_PRID = "SRM",
                                PrErpPrefix,
                                PrErpNo,
                                PrSeq
                            });
                        #endregion

                        #region //請購的採購資料更新 PURTY
                        sql = @"UPDATE PURTY SET 
                                MODIFIER = @CurrentUser,
                                MODI_DATE = @MODI_DATE,
                                MODI_TIME = @MODI_TIME,
                                MODI_AP = @MODI_AP,
                                MODI_PRID = @MODI_PRID,
                                FLAG = FLAG + 1,
                                TY021 = 'N'
                                WHERE TY001 = @PrErpPrefix
                                AND TY002 = @PrErpNo
                                AND TY003 = @PrSeq";
                        rowsAffected = sqlConnection.Execute(sql,
                            new
                            {
                                CurrentUser,
                                MODI_DATE = dateNow,
                                MODI_TIME = timeNow,
                                MODI_AP = CurrentUser + "PC",
                                MODI_PRID = "SRM",
                                PrErpPrefix,
                                PrErpNo,
                                PrSeq
                            });
                        #endregion
                    }

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "(1 rows affected)"
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

        #region//UpdatePoConfirmERP --更新ERP請購單確認 -- Shinotokuro 2025.03.20
        public string UpdatePoConfirmERP(string PoErpPrefix, string PoErpNo
            , string CompanyNo, string CurrentUser)
        {
            try
            {
                if (CompanyNo.Length <= 0) throw new SystemException("【公司】不能為空!");
                if (CurrentUser.Length <= 0) throw new SystemException("【使用者】不能為空!");

                string dateNow = CreateDate.ToString("yyyyMMdd"), timeNow = CreateDate.ToString("HH:mm:ss");

                using (TransactionScope transactionScope = new TransactionScope(TransactionScopeOption.Required, new TimeSpan(0, 180, 0)))
                {
                    string ErpNo = "";
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認公司別DB
                        sql = @"SELECT TOP 1 a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo = CompanyNo == "ETG" ? "Eterge" : CompanyNo }) ??  throw new SystemException("【公司別】資料錯誤!");
                        #endregion

                        ErpConnectionStrings = ConfigurationManager.AppSettings[result.ErpDb];
                        ErpNo = result.ErpNo;

                        #region//判斷使用者是否存在
                        sql = @"SELECT TOP 1 a.UserNo
                                FROM BAS.[User] a
                                WHERE a.UserNo = @CurrentUser";
                        dynamicParameters.Add("UserNo", CurrentUser);
                        result = sqlConnection.QueryFirstOrDefault(sql, new { CurrentUser }) ?? throw new SystemException("採購單【使用者】資料錯誤!");
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        //PURTC TC014='Y'確認碼 TC024=''(日期) TC025=''(確認者)
                        //PURTD TD018='Y'
                        //PURMB 第一次新增資料表

                        #region //判斷ERP使用者是否存在
                        sql = @"SELECT TOP 1 MF004, MF005
                                FROM ADMMF
                                WHERE COMPANY = @ErpNo
                                AND MF001 = @CurrentUser";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { ErpNo, CurrentUser }) ??  throw new SystemException("【ERP使用者】不存在!");
                        #endregion

                        #region //判斷採購單是否存在
                        sql = @"SELECT TOP 1 a.TC004 SupplierNo, a.TC005 Currency,a.TC048 TradeTerm,a.TC018 Taxation
                                ,b.TD004 MtlItemNo, b.TD059 PoPriceUnit,b.TD010 PoUnitPrice
                                , x.HavePurmb
                                FROM PURTC a
                                INNER JOIN PURTD b on a.TC001 = b.TD001 AND a.TC002 = b.TD002
                                INNER JOIN PURMA c on a.TC004 = c.MA001
                                OUTER APPLY(
                                    SELECT TOP 1 1 HavePurmb FROM PURMB x1 WHERE x1.MB001 = b.TD004 AND x1.MB002 = a.TC004
                                ) x
                                WHERE a.TC001 = @PoErpPrefix
                                AND a.TC002 = @PoErpNo";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { PoErpPrefix, PoErpNo }) ?? throw new SystemException("【採購單】不存在!");

                        if (result.HavePurmb != 1)
                        {
                            #region //新增PURMB
                            sql = @"INSERT INTO PURMB (MB001, MB002, MB003, MB004, MB005
                                    , MB007, MB008, MB009, MB010, MB011, MB012, MB013, MB014
                                    , MB015, MB016, MB017, MB018, MB019, MB020, MB021, MB022)
                                    VALUES (@MB001, @MB002, @MB003, @MB004, @MB005
                                    , @MB007, @MB008, @MB009, @MB010, @MB011, @MB012, @MB013, @MB014
                                    , @MB015, @MB016, @MB017, @MB018, @MB019, @MB020, @MB021, @MB022)";
                            rowsAffected += sqlConnection.Execute(sql,
                                new
                                {
                                    MB001 = result.MtlItemNo,
                                    MB002 = result.SupplierNo,
                                    MB003 = result.Currency,
                                    MB004 = result.PoPriceUnit,
                                    MB005 = "",
                                    MB007 = "",
                                    MB008 = dateNow,
                                    MB009 = "",
                                    MB010 = "N",
                                    MB011 = result.PoUnitPrice,
                                    MB012 = "",
                                    MB013 = result.Taxation == "1" ? "Y" : "N",
                                    MB014 = dateNow,
                                    MB015 = "",
                                    MB016 = "",
                                    MB017 = 0,
                                    MB018 = 0,
                                    MB019 = "",
                                    MB020 = "",
                                    MB021 = "",
                                    MB022 = result.TradeTerm,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            #endregion
                        }
                        #endregion

                        #region //請購的採購資料更新 PURTC
                        sql = @"UPDATE PURTC SET 
                                MODIFIER = @CurrentUser,
                                MODI_DATE = @MODI_DATE,
                                MODI_TIME = @MODI_TIME,
                                MODI_AP = @MODI_AP,
                                MODI_PRID = @MODI_PRID,
                                FLAG = FLAG + 1,
                                TC014 = 'Y',
                                TC024 = @TC024,
                                TC025 = @TC025
                                WHERE TC001 = @PoErpPrefix
                                AND TC002 = @PoErpNo";
                        rowsAffected += sqlConnection.Execute(sql,
                            new
                            {
                                CurrentUser,
                                MODI_DATE = dateNow,
                                MODI_TIME = timeNow,
                                MODI_AP = CurrentUser + "PC",
                                MODI_PRID = "SRM",
                                TC024 = dateNow,
                                TC025 = CurrentUser,
                                PoErpPrefix,
                                PoErpNo
                            });
                        #endregion

                        #region //請購的採購資料更新 PURTD
                        sql = @"UPDATE PURTD SET
                                MODIFIER = @CurrentUser,
                                MODI_DATE = @MODI_DATE,
                                MODI_TIME = @MODI_TIME,
                                MODI_AP = @MODI_AP,
                                MODI_PRID = @MODI_PRID,
                                FLAG = FLAG + 1,
                                TD018 = 'Y'
                                WHERE TD001 = @PoErpPrefix
                                AND TD002 = @PoErpNo";
                        rowsAffected += sqlConnection.Execute(sql,
                            new
                            {
                                CurrentUser,
                                MODI_DATE = dateNow,
                                MODI_TIME = timeNow,
                                MODI_AP = CurrentUser + "PC",
                                MODI_PRID = "SRM",
                                PoErpPrefix,
                                PoErpNo
                            });
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

        #region//UpdatePoReConfirmERP --更新ERP採購單反確認 -- Shinotokuro 2025.03.20
        public string UpdatePoReConfirmERP(string PoErpPrefix, string PoErpNo
            , string CompanyNo, string CurrentUser)
        {
            try
            {
                if (CompanyNo.Length <= 0) throw new SystemException("【公司】不能為空!");
                if (CurrentUser.Length <= 0) throw new SystemException("【使用者】不能為空!");

                string dateNow = CreateDate.ToString("yyyyMMdd"), timeNow = CreateDate.ToString("HH:mm:ss"), errorMsg = "";

                using (TransactionScope transactionScope = new TransactionScope(TransactionScopeOption.Required, new TimeSpan(0, 180, 0)))
                {
                    string ErpNo = "";
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認公司別DB
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { CompanyNo }) ?? throw new SystemException("請選擇公司別!");
                        #endregion

                        ErpConnectionStrings = ConfigurationManager.AppSettings[result.ErpDb];
                        ErpNo = result.ErpNo;

                        #region//判斷使用者是否存在
                        sql = @"SELECT TOP 1 a.UserNo
                                FROM BAS.[User] a
                                WHERE a.UserNo = @CurrentUser";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { CurrentUser }) ?? throw new SystemException("採購單【使用者】資料錯誤!");
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        //PURTC TC014='Y'確認碼 TC024=''(日期) TC025=''(確認者)
                        //PURTD TD018='Y'
                        //PURMB 第一次新增資料表

                        #region //判斷ERP使用者是否存在
                        sql = @"SELECT TOP 1 MF004, MF005
                                FROM ADMMF
                                WHERE COMPANY = @ErpNo
                                AND MF001 = @CurrentUser";
                        var result = sqlConnection.QueryFirstOrDefault(sql, new { ErpNo, CurrentUser }) ?? throw new SystemException("【ERP使用者】不存在!");
                        #endregion

                        #region //判斷採購單是否存在
                        sql = @"SELECT TOP 1 1
                                FROM PURTC a
                                INNER JOIN PURTD b on a.TC001 = b.TD001 AND a.TC002 = b.TD002
                                WHERE a.TC001 = @PoErpPrefix
                                AND a.TC002 = @PoErpNo";
                        result = sqlConnection.QueryFirstOrDefault(sql, new { PoErpPrefix, PoErpNo }) ?? throw new SystemException("【採購單】不存在!");
                        #endregion

                        #region //請購的採購資料更新 PURTC
                        sql = @"UPDATE PURTC SET 
                                MODIFIER = @CurrentUser,
                                MODI_DATE = @MODI_DATE,
                                MODI_TIME = @MODI_TIME,
                                MODI_AP = @MODI_AP,
                                MODI_PRID = @MODI_PRID,
                                FLAG = FLAG + 1,
                                TC014 = 'N',
                                TC024 = @TC024,
                                TC025 = @TC025
                                WHERE TC001 = @PoErpPrefix
                                AND TC002 = @PoErpNo";
                        rowsAffected += sqlConnection.Execute(sql,
                            new
                            {
                                CurrentUser,
                                MODI_DATE = dateNow,
                                MODI_TIME = timeNow,
                                MODI_AP = CurrentUser + "PC",
                                MODI_PRID = "SRM",
                                TC024 = dateNow,
                                TC025 = "",
                                PoErpPrefix,
                                PoErpNo
                            });
                        #endregion

                        #region //請購的採購資料更新 PURTD
                        sql = @"UPDATE PURTD SET 
                                MODIFIER = @CurrentUser,
                                MODI_DATE = @MODI_DATE,
                                MODI_TIME = @MODI_TIME,
                                MODI_AP = @MODI_AP,
                                MODI_PRID = @MODI_PRID,
                                FLAG = FLAG + 1,
                                TD018= 'N'
                                WHERE TD001 = @PoErpPrefix
                                AND TD002 = @PoErpNo";
                        rowsAffected += sqlConnection.Execute(sql,
                            new
                            {
                                CurrentUser,
                                MODI_DATE = dateNow,
                                MODI_TIME = timeNow,
                                MODI_AP = CurrentUser + "PC",
                                MODI_PRID = "SRM",
                                PoErpPrefix,
                                PoErpNo
                            });
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
        #endregion
        #endregion
    }
}
