using Dapper;
using Helpers;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Transactions;
using System.Web;

namespace SCMDA
{
    public class ScmInventoryDA
    {
        public string MainConnectionStrings = "";
        public string ErpConnectionStrings = "";

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

        public ScmInventoryDA()
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
        #region //GetInventoryTransaction -- 取得庫存異動資料 -- Ben Ma 2023.03.16
        public string GetInventoryTransaction(int ItId, string ItErpPrefix, string ItErpNo, string SearchKey
            , string ConfirmStatus, string StartDate, string EndDate
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                List<InventoryTransaction> its = new List<InventoryTransaction>();

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

                    sqlQuery.mainKey = "a.ItId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.CompanyId, a.ItErpPrefix, a.ItErpNo, FORMAT(a.DocDate, 'yyyy-MM-dd') DocDate
                        , a.DepartmentId, a.Remark
                        , a.ConfirmStatus, a.ConfirmUserId, a.TransferStatus, FORMAT(a.TransferDate, 'yyyy-MM-dd') TransferDate
                        , a.ItErpPrefix + '-' + a.ItErpNo ItErpFullNo
                        , b.DepartmentNo, b.DepartmentName
                        , (
                            SELECT aa.ItDetailId, aa.ItSequence, aa.MtlItemId, aa.ItMtlItemName, aa.ItMtlItemSpec
                            , aa.ItQty, aa.UnitCost, aa.Amount, aa.InventoryId, aa.ItRemark
                            , ab.MtlItemNo
                            , ac.InventoryNo InventoryNo, ac.InventoryName InventoryName
                            FROM SCM.ItDetail aa
                            INNER JOIN PDM.MtlItem ab ON aa.MtlItemId = ab.MtlItemId
                            INNER JOIN SCM.Inventory ac ON aa.InventoryId = ac.InventoryId
                            WHERE aa.ItId = a.ItId
                            ORDER BY aa.ItSequence
                            FOR JSON PATH, ROOT('data')
                        ) ItDetail
                        , ISNULL(c.UserNo, '') ConfirmUserNo, ISNULL(c.UserName,'') ConfirmUserName, ISNULL(c.Gender, '') ConfirmUserGender";
                    sqlQuery.mainTables =
                        @"FROM SCM.InventoryTransaction a
                        INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                        LEFT JOIN BAS.[User] c ON a.ConfirmUserId = c.UserId";
                    string queryTable =
                        @"FROM (
                            SELECT a.ItId, a.CompanyId, a.ItErpPrefix, a.ItErpNo, a.DocDate, a.ConfirmStatus
                            , (
                                SELECT y.MtlItemNo, y.MtlItemName
                                FROM SCM.ItDetail z
                                INNER JOIN PDM.MtlItem y ON z.MtlItemId = y.MtlItemId
                                WHERE z.ItId = a.ItId
                                FOR JSON PATH, ROOT('data')
                            ) ItDetail
                            FROM SCM.InventoryTransaction a
                        ) a
                        OUTER APPLY (
                            SELECT TOP 1 x.MtlItemNo, x.MtlItemName
                            FROM OPENJSON(a.ItDetail, '$.data')
                            WITH (
                                MtlItemNo NVARCHAR(40) N'$.MtlItemNo', 
                                MtlItemName NVARCHAR(120) N'$.MtlItemName'
                            ) x
                        ) b";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ItId", @" AND a.ItId = @ItId", ItId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ItErpPrefix", @" AND a.ItErpPrefix = @ItErpPrefix", ItErpPrefix);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ItErpNo", @" AND a.ItErpNo LIKE '%' + @ItErpNo + '%'", ItErpNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey", @" AND (b.MtlItemNo LIKE '%' + @SearchKey + '%' OR b.MtlItemName LIKE '%' + @SearchKey + '%')", SearchKey);
                    if (ConfirmStatus.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ConfirmStatus", @" AND a.ConfirmStatus IN @ConfirmStatus", ConfirmStatus.Split(','));
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.DocDate >= @StartDate ", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.DocDate <= @EndDate ", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.DocDate DESC, a.ItErpPrefix, a.ItErpNo";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    its = BaseHelper.SqlQuery<InventoryTransaction>(sqlConnection, dynamicParameters, sqlQuery);
                }

                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                {
                    List<ErpDocStatus> erpDocStatuses = new List<ErpDocStatus>();

                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TA001 ErpPrefix, TA002 ErpNo, TA006 DocStatus
                            FROM INVTA
                            WHERE (TA001 + '-' + TA002) IN @ErpFullNo";
                    dynamicParameters.Add("ErpFullNo", its.Select(x => x.ItErpPrefix + '-' + x.ItErpNo).ToArray());
                    erpDocStatuses = sqlConnection.Query<ErpDocStatus>(sql, dynamicParameters).ToList();

                    its = its.GroupJoin(erpDocStatuses, x => x.ItErpPrefix + '-' + x.ItErpNo, y => y.ErpPrefix + '-' + y.ErpNo, (x, y) => { x.DocStatus = y.FirstOrDefault()?.DocStatus ?? ""; return x; }).ToList();
                    its = its.OrderByDescending(x => x.DocDate).ToList();
                }

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    data = its
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

        #region //GetTempShippingNote -- 取得暫出單資料 -- Ben Ma 2023.05.03
        public string GetTempShippingNote(int TsnId, string TsnErpPrefix, string TsnErpNo, string SearchKey
            , int ObjectCustomer, string ConfirmStatus, string StartDate, string EndDate
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                List<TempShippingNote> tsns = new List<TempShippingNote>();

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

                    sqlQuery.mainKey = "a.TsnId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.CompanyId, LTRIM(RTRIM(a.TsnErpPrefix)) TsnErpPrefix, LTRIM(RTRIM(a.TsnErpNo)) TsnErpNo, FORMAT(a.DocDate, 'yyyy-MM-dd') DocDate
                        , a.ToObject, ISNULL(a.ObjectCustomer, -1) ObjectCustomer, ISNULL(a.ObjectSupplier, -1) ObjectSupplier
                        , ISNULL(a.ObjectUser, -1) ObjectUser, a.ObjectOther, a.DepartmentId, a.UserId, a.Remark, a.OtherRemark
                        , a.ConfirmStatus, a.ConfirmUserId, a.TransferStatus, FORMAT(a.TransferDate, 'yyyy-MM-dd') TransferDate
                        , a.TsnErpPrefix + '-' + a.TsnErpNo TsnErpFullNo
                        , b.TypeName ToObjectName
                        , CASE 
                            WHEN a.ToObject = '1' THEN c.CustomerShortName
                            WHEN a.ToObject = '2' THEN d.SupplierName
                            WHEN a.ToObject = '3' THEN e.UserName
                        ELSE a.ObjectOther END ObjectName
                        , (
                            SELECT aa.TsnDetailId, aa.TsnSequence, aa.MtlItemId, aa.TsnMtlItemName, aa.TsnMtlItemSpec
                            , aa.TsnOutInventory, aa.TsnInInventory, aa.TsnQty
                            , aa.ProductType, ISNULL(aa.FreebieOrSpareQty, '0') FreebieOrSpareQty, ag.TypeName ProductTypeName
                            , aa.SoDetailId, aa.TsnRemark, ISNULL(aa.LotNumber, '') LotNumber
                            , ab.MtlItemNo
                            , ac.InventoryNo TsnOutInventoryNo, ac.InventoryName TsnOutInventoryName
                            , ad.InventoryNo TsnInInventoryNo, ad.InventoryName TsnInInventoryName
                            , ae.SoSequence
                            , af.SoErpPrefix, af.SoErpNo
                            , af.SoErpPrefix + '-' + af.SoErpNo + '-' + ae.SoSequence SoErpFullNo
                            FROM SCM.TsnDetail aa
                            INNER JOIN PDM.MtlItem ab ON aa.MtlItemId = ab.MtlItemId
                            INNER JOIN SCM.Inventory ac ON aa.TsnOutInventory = ac.InventoryId
                            INNER JOIN SCM.Inventory ad ON aa.TsnInInventory = ad.InventoryId
                            LEFT JOIN SCM.SoDetail ae ON aa.SoDetailId = ae.SoDetailId
                            INNER JOIN SCM.SaleOrder af ON ae.SoId = af.SoId
                            INNER JOIN BAS.[Type] ag ON aa.ProductType = ag.TypeNo AND ag.TypeSchema = 'SoDetail.ProductType'
                            WHERE aa.TsnId = a.TsnId
                            ORDER BY aa.TsnSequence
                            FOR JSON PATH, ROOT('data')
                        ) TsnDetail
                        , ISNULL(f.UserNo, '') ConfirmUserNo, ISNULL(f.UserName,'') ConfirmUserName, ISNULL(f.Gender, '') ConfirmUserGender
                        , ISNULL(g.UserNo, '') UserNo, ISNULL(g.UserName,'') UserName, ISNULL(g.Gender, '') UserGender";
                    sqlQuery.mainTables =
                        @"FROM SCM.TempShippingNote a
                        INNER JOIN BAS.[Type] b ON a.ToObject = b.TypeNo AND b.TypeSchema = 'TempShippingNote.ToObject'
                        LEFT JOIN SCM.Customer c ON a.ObjectCustomer = c.CustomerId
                        LEFT JOIN SCM.Supplier d ON a.ObjectSupplier = d.SupplierId
                        LEFT JOIN BAS.[User] e ON a.ObjectUser = e.UserId
                        LEFT JOIN BAS.[User] f ON a.ConfirmUserId = f.UserId
                        LEFT JOIN BAS.[User] g ON a.UserId = g.UserId";
                    string queryTable =
                        @"FROM (
                            SELECT a.TsnId, a.CompanyId, a.TsnErpPrefix, a.TsnErpNo, FORMAT(a.DocDate, 'yyyy-MM-dd') DocDate, a.ConfirmStatus, a.ObjectCustomer
                            ,(
                                SELECT aa.TsnSequence, ab.MtlItemNo, aa.TsnMtlItemName MtlItemName
                                FROM SCM.TsnDetail aa
                                LEFT JOIN PDM.MtlItem ab ON aa.MtlItemId = ab.MtlItemId
                                WHERE aa.TsnId = a.TsnId
                                FOR JSON PATH, ROOT('data')
                            ) TsnDetail
                            FROM SCM.TempShippingNote a
                        ) a 
                        OUTER APPLY (
                            SELECT TOP 1 x.MtlItemNo, x.MtlItemName
                            FROM OPENJSON(a.TsnDetail, '$.data')
                            WITH (
                                MtlItemNo NVARCHAR(40) N'$.MtlItemNo', 
                                MtlItemName NVARCHAR(120) N'$.MtlItemName'
                            ) x
                        ) b";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TsnId", @" AND a.TsnId = @TsnId", TsnId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TsnErpPrefix", @" AND a.TsnErpPrefix = @TsnErpPrefix", TsnErpPrefix);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TsnErpNo", @" AND a.TsnErpNo LIKE '%' + @TsnErpNo + '%'", TsnErpNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ObjectCustomer", @" AND a.ObjectCustomer = @ObjectCustomer", ObjectCustomer);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey", @" AND (b.MtlItemNo LIKE '%' + @SearchKey + '%' OR b.MtlItemName LIKE '%' + @SearchKey + '%')", SearchKey);
                    if (ConfirmStatus.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ConfirmStatus", @" AND a.ConfirmStatus IN  @ConfirmStatus ", ConfirmStatus.Split(','));
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.DocDate >= @StartDate ", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.DocDate <= @EndDate ", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd") : "");
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.DocDate DESC ,a.ConfirmStatus, a.TsnErpPrefix, a.TsnErpNo";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    tsns = BaseHelper.SqlQuery<TempShippingNote>(sqlConnection, dynamicParameters, sqlQuery);
                }

                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                {
                    List<ErpDocStatus> erpDocStatuses = new List<ErpDocStatus>();

                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT LTRIM(RTRIM(TF001)) ErpPrefix, LTRIM(RTRIM(TF002)) ErpNo, TF020 DocStatus
                            FROM INVTF
                            WHERE (LTRIM(RTRIM(TF001)) + '-' + LTRIM(RTRIM(TF002))) IN @ErpFullNo";
                    dynamicParameters.Add("ErpFullNo", tsns.Select(x => x.TsnErpPrefix + '-' + x.TsnErpNo).ToArray());
                    erpDocStatuses = sqlConnection.Query<ErpDocStatus>(sql, dynamicParameters).ToList();

                    tsns = tsns.GroupJoin(erpDocStatuses, x => x.TsnErpPrefix + '-' + x.TsnErpNo, y => y.ErpPrefix + '-' + y.ErpNo, (x, y) => { x.DocStatus = y.FirstOrDefault()?.DocStatus ?? ""; return x; }).ToList();
                    tsns = tsns.OrderByDescending(x => x.DocDate).ToList();
                }

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    data = tsns
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

        #region //GetDeliveryToTsn -- 取得出貨轉暫出資料 -- Ben Ma 2023.05.09
        public string GetDeliveryToTsn(string SearchKey, string StartDate, string EndDate
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.DoDetailId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", FORMAT(b.DoDate, 'yyyy-MM-dd hh:mm:ss') DoDate, c.DcShortName
                        , a.DoQty RegularQty, ISNULL(a.FreebieQty, 0) FreebieQty, ISNULL(a.SpareQty, 0) SpareQty
                        , b.[Status], d.StatusName DoStatus
                        , e.InventoryId, e1.InventoryNo
                        , g.MtlItemNo, ISNULL(e.SoMtlItemName, g.MtlItemName) MtlItemName, ISNULL(e.SoMtlItemSpec, g.MtlItemSpec) MtlItemSpec
                        , f.SoErpPrefix + '-' + f.SoErpNo SoErpFullNo, e.SoSequence,f.CustomerId
                        , (ISNULL(i.TsnQty, 0) - ISNULL(i.ReturnQty, 0)) TsnQty, ISNULL(i.SaleQty, 0) SaleQty
                        , e.SoQty, ISNULL(h.PickQty, 0) PickQty
                        , j.CustomerNo, j.CustomerShortName
                        , e.SoQty SoRegularQty, ISNULL(e.FreebieQty, 0) SoFreebieQty, ISNULL(e.SpareQty, 0) SoSpareQty
                        , (ISNULL(o.TsnFreebieQty, 0) - ISNULL(o.TsnReturnFSFreebieQty, 0)) TsnFreebieQty, ISNULL(o.TsnSaleFSFreebieQty, 0) TsnSaleFSFreebieQty
                        , (ISNULL(p.TsnSpareQty, 0) - ISNULL(p.TsnReturnFSSpareQty, 0)) TsnSpareQty, ISNULL(p.TsnSaleFSSpareQty, 0) TsnSaleFSSpareQty
                        , ISNULL(k.PickRegularQty, 0) PickRegularQty, ISNULL(l.PickFreebieQty, 0) PickFreebieQty, ISNULL(m.PickSpareQty, 0) PickSpareQty
                        , ISNULL(n.UserNo, '') SalesmenNo, ISNULL(n.UserName, '') SalesmenName, ISNULL(n.Gender, '') SalesmenGender
                        , a.PcDoDetailRemark ";
                    sqlQuery.mainTables =
                        @"FROM SCM.DoDetail a
                        INNER JOIN SCM.DeliveryOrder b ON a.DoId = b.DoId
                        INNER JOIN SCM.DeliveryCustomer c ON b.DcId = c.DcId
                        INNER JOIN BAS.[Status] d ON b.[Status] = d.StatusNo AND d.StatusSchema = 'DeliveryOrder.Status'
                        INNER JOIN SCM.SoDetail e ON a.SoDetailId = e.SoDetailId
                        INNER JOIN SCM.Inventory e1 ON e.InventoryId = e1.InventoryId
                        INNER JOIN SCM.SaleOrder f ON e.SoId = f.SoId
                        INNER JOIN PDM.MtlItem g ON e.MtlItemId = g.MtlItemId
                        OUTER APPLY (
                            SELECT SUM(x.ItemQty) PickQty
                            FROM SCM.PickingItem x
                            WHERE x.SoDetailId = a.SoDetailId
                            AND x.DoId = a.DoId
                        ) h
                        OUTER APPLY(
                            SELECT SUM(x.TsnQty) TsnQty, SUM(x.SaleQty) SaleQty, SUM(x.ReturnQty) ReturnQty
                            FROM SCM.TsnDetail x
                            INNER JOIN SCM.TempShippingNote y ON x.TsnId = y.TsnId
                            WHERE x.SoDetailId = a.SoDetailId
                            AND x.ConfirmStatus = 'Y'
                            AND y.ConfirmStatus = 'Y'
                        ) i
                        INNER JOIN SCM.Customer j ON f.CustomerId = j.CustomerId
                        OUTER APPLY (
                            SELECT SUM(za.ItemQty) PickRegularQty
                            FROM SCM.PickingItem za
                            WHERE za.SoDetailId = a.SoDetailId
                            AND za.ItemType = 1
                            AND za.DoId = a.DoId
                        ) k
                        OUTER APPLY (
                            SELECT SUM(zb.ItemQty) PickFreebieQty
                            FROM SCM.PickingItem zb
                            WHERE zb.SoDetailId = a.SoDetailId
                            AND zb.ItemType = 2
                            AND zb.DoId = a.DoId
                        ) l
                        OUTER APPLY (
                            SELECT SUM(zc.ItemQty) PickSpareQty
                            FROM SCM.PickingItem zc
                            WHERE zc.SoDetailId = a.SoDetailId
                            AND zc.ItemType = 3
                            AND zc.DoId = a.DoId
                        ) m
                        INNER JOIN BAS.[User] n ON f.SalesmenId = n.UserId
                        OUTER APPLY (
                            SELECT SUM(x.FreebieOrSpareQty) TsnFreebieQty, SUM(x.SaleFSQty) TsnSaleFSFreebieQty, SUM(x.ReturnFSQty) TsnReturnFSFreebieQty
                            FROM SCM.TsnDetail x
                            INNER JOIN SCM.TempShippingNote y ON x.TsnId = y.TsnId
                            WHERE x.SoDetailId = a.SoDetailId
                            AND x.ConfirmStatus = 'Y'
                            AND y.ConfirmStatus = 'Y'
                            AND x.ProductType = '1'
                        ) o
                        OUTER APPLY (
                            SELECT SUM(x.FreebieOrSpareQty) TsnSpareQty, SUM(x.SaleFSQty) TsnSaleFSSpareQty, SUM(x.ReturnFSQty) TsnReturnFSSpareQty
                            FROM SCM.TsnDetail x
                            INNER JOIN SCM.TempShippingNote y ON x.TsnId = y.TsnId
                            WHERE x.SoDetailId = a.SoDetailId
                            AND x.ConfirmStatus = 'Y'
                            AND y.ConfirmStatus = 'Y'
                            AND x.ProductType = '2'
                        ) p";
                    sqlQuery.auxTables = "";
                    string queryCondition = @" AND b.CompanyId = @CompanyId
                                            AND b.[Status] = 'S'";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey",
                        @" AND (g.MtlItemNo LIKE '%' + @SearchKey + '%' 
                            OR g.MtlItemName LIKE '%' + @SearchKey + '%'
                            OR f.SoErpPrefix + '-' + f.SoErpNo LIKE '%' + @SearchKey + '%')", SearchKey);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND b.DoDate >= @StartDate ", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND b.DoDate <= @EndDate ", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "FORMAT(b.DoDate, 'yyyy-MM-dd') DESC, b.DcId, f.SoErpPrefix, f.SoErpNo, e.SoSequence";
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

        #region //GetTempShippingReturnNote -- 取得暫出歸還單資料 -- Ben Ma 2023.05.22
        public string GetTempShippingReturnNote(int TsrnId, string TsrnErpPrefix, string TsrnErpNo, string SearchKey
            , string ConfirmStatus, string StartDate, string EndDate
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                List<TempShippingReturnNote> tsrns = new List<TempShippingReturnNote>();

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

                    sqlQuery.mainKey = "a.TsrnId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.CompanyId, a.TsrnErpPrefix, a.TsrnErpNo, FORMAT(a.DocDate, 'yyyy-MM-dd') DocDate
                        , a.ToObject, ISNULL(a.ObjectCustomer, -1) ObjectCustomer, ISNULL(a.ObjectSupplier, -1) ObjectSupplier
                        , ISNULL(a.ObjectUser, -1) ObjectUser, a.ObjectOther, a.DepartmentId, a.UserId, a.Remark, a.OtherRemark
                        , a.ConfirmStatus, a.ConfirmUserId, a.TransferStatus, FORMAT(a.TransferDate, 'yyyy-MM-dd') TransferDate
                        , a.TsrnErpPrefix + '-' + a.TsrnErpNo TsrnErpFullNo
                        , b.TypeName ToObjectName
                        , CASE 
                            WHEN a.ToObject = '1' THEN c.CustomerShortName
                            WHEN a.ToObject = '2' THEN d.SupplierName
                            WHEN a.ToObject = '3' THEN e.UserName
                        ELSE a.ObjectOther END ObjectName
                        , (
                            SELECT aa.TsrnDetailId, aa.TsrnSequence, aa.MtlItemId, aa.TsrnMtlItemName, aa.TsrnMtlItemSpec
                            , aa.TsrnOutInventory, aa.TsrnInInventory, aa.TsrnQty, aa.ProductType, ISNULL(aa.FreebieOrSpareQty, '0') FreebieOrSpareQty, aa.TsnDetailId, aa.TsrnRemark
                            , ab.MtlItemNo
                            , ac.InventoryNo TsrnOutInventoryNo, ac.InventoryName TsrnOutInventoryName
                            , ad.InventoryNo TsrnInInventoryNo, ad.InventoryName TsrnInInventoryName
                            , ae.TsnSequence
                            , af.TsnErpPrefix, af.TsnErpNo
                            , af.TsnErpPrefix + '-' + af.TsnErpNo + '-' + ae.TsnSequence TsnErpFullNo
                            , ag.TypeName ProductTypeName
                            FROM SCM.TsrnDetail aa
                            INNER JOIN PDM.MtlItem ab ON aa.MtlItemId = ab.MtlItemId
                            INNER JOIN SCM.Inventory ac ON aa.TsrnOutInventory = ac.InventoryId
                            INNER JOIN SCM.Inventory ad ON aa.TsrnInInventory = ad.InventoryId
                            INNER JOIN SCM.TsnDetail ae ON aa.TsnDetailId = ae.TsnDetailId
                            INNER JOIN SCM.TempShippingNote af ON ae.TsnId = af.TsnId
                            INNER JOIN BAS.[Type] ag ON aa.ProductType = ag.TypeNo AND ag.TypeSchema = 'SoDetail.ProductType'
                            WHERE aa.TsrnId = a.TsrnId
                            ORDER BY aa.TsrnSequence
                            FOR JSON PATH, ROOT('data')
                        ) TsrnDetail
                        , ISNULL(f.UserNo, '') ConfirmUserNo, ISNULL(f.UserName,'') ConfirmUserName, ISNULL(f.Gender, '') ConfirmUserGender";
                    sqlQuery.mainTables =
                        @"FROM SCM.TempShippingReturnNote a
                        INNER JOIN BAS.[Type] b ON a.ToObject = b.TypeNo AND b.TypeSchema = 'TempShippingNote.ToObject'
                        LEFT JOIN SCM.Customer c ON a.ObjectCustomer = c.CustomerId
                        LEFT JOIN SCM.Supplier d ON a.ObjectSupplier = d.SupplierId
                        LEFT JOIN BAS.[User] e ON a.ObjectUser = e.UserId
                        LEFT JOIN BAS.[User] f ON a.ConfirmUserId = f.UserId";
                    string queryTable =
                        @"FROM (
                            SELECT a.TsrnId, a.CompanyId, a.TsrnErpPrefix, a.TsrnErpNo, FORMAT(a.DocDate, 'yyyy-MM-dd') DocDate, a.ConfirmStatus
                            ,(
                                SELECT aa.TsrnSequence, ab.MtlItemNo, aa.TsrnMtlItemName MtlItemName
                                FROM SCM.TsrnDetail aa
                                LEFT JOIN PDM.MtlItem ab ON aa.MtlItemId = ab.MtlItemId
                                WHERE aa.TsrnId = a.TsrnId
                                FOR JSON PATH, ROOT('data')
                            ) TsrnDetail
                            FROM SCM.TempShippingReturnNote a
                        ) a 
                        OUTER APPLY (
                            SELECT TOP 1 x.MtlItemNo, x.MtlItemName
                            FROM OPENJSON(a.TsrnDetail, '$.data')
                            WITH (
                                MtlItemNo NVARCHAR(40) N'$.MtlItemNo', 
                                MtlItemName NVARCHAR(120) N'$.MtlItemName'
                            ) x
                        ) b";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TsnId", @" AND a.TsrnId = @TsrnId", TsrnId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TsrnErpPrefix", @" AND EXISTS (
                                                                                                            SELECT TOP 1 1
                                                                                                            FROM SCM.TsrnDetail aa
                                                                                                            LEFT JOIN PDM.MtlItem ab ON aa.MtlItemId = ab.MtlItemId
                                                                                                            WHERE aa.TsrnId = a.TsrnId
                                                                                                            AND a.TsrnErpPrefix = @TsrnErpPrefix
                                                                                                       )", TsrnErpPrefix);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TsrnErpNo", @" AND EXISTS (
                                                                                                            SELECT TOP 1 1
                                                                                                            FROM SCM.TsrnDetail aa
                                                                                                            LEFT JOIN PDM.MtlItem ab ON aa.MtlItemId = ab.MtlItemId
                                                                                                            WHERE aa.TsrnId = a.TsrnId
                                                                                                            AND a.TsrnErpNo = @TsrnErpNo
                                                                                                       )", TsrnErpNo);


                    //BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TsnErpPrefix", @" AND a.TsrnErpPrefix = @TsrnErpPrefix", TsrnErpPrefix);
                    //BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TsnErpNo", @" AND a.TsrnErpNo LIKE '%' + @TsrnErpNo + '%'", TsrnErpNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey", @" AND (b.MtlItemNo LIKE '%' + @SearchKey + '%' OR b.MtlItemName LIKE '%' + @SearchKey + '%')", SearchKey);
                    if (ConfirmStatus.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ConfirmStatus", @" AND a.ConfirmStatus IN @ConfirmStatus", ConfirmStatus.Split(','));
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.DocDate >= @StartDate ", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.DocDate <= @EndDate ", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.DocDate DESC, a.TsrnErpPrefix, a.TsrnErpNo";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    tsrns = BaseHelper.SqlQuery<TempShippingReturnNote>(sqlConnection, dynamicParameters, sqlQuery);
                }

                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                {
                    List<ErpDocStatus> erpDocStatuses = new List<ErpDocStatus>();

                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT LTRIM(RTRIM(TH001)) ErpPrefix, LTRIM(RTRIM(TH002)) ErpNo, TH020 DocStatus
                            FROM INVTH
                            WHERE (LTRIM(RTRIM(TH001)) + '-' + LTRIM(RTRIM(TH002))) IN @ErpFullNo";
                    dynamicParameters.Add("ErpFullNo", tsrns.Select(x => x.TsrnErpPrefix + '-' + x.TsrnErpNo).ToArray());
                    erpDocStatuses = sqlConnection.Query<ErpDocStatus>(sql, dynamicParameters).ToList();

                    tsrns = tsrns.GroupJoin(erpDocStatuses, x => x.TsrnErpPrefix + '-' + x.TsrnErpNo, y => y.ErpPrefix + '-' + y.ErpNo, (x, y) => { x.DocStatus = y.FirstOrDefault()?.DocStatus ?? ""; return x; }).ToList();
                    tsrns = tsrns.OrderByDescending(x => x.DocDate).ToList();
                }

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    data = tsrns
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

        #region //GetInventoryQuery -- 庫存整合查詢 -- Ben Ma 2023.05.31
        public string GetInventoryQuery(string SearchKey 
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                List<InventoryMtlItem> inventoryMtlItems = new List<InventoryMtlItem>();

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

                    sqlQuery.mainKey = "a.MtlItemId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.MtlItemNo, a.MtlItemName, a.MtlItemSpec";
                    sqlQuery.mainTables =
                        @"FROM PDM.MtlItem a";
                    string queryTable =
                        @"FROM (
                            SELECT a.MtlItemId, a.CompanyId, a.MtlItemNo, a.MtlItemName, a.MtlItemSpec
                            , (
                                SELECT aa.CustomerMtlItemNo
                                FROM PDM.CustomerMtlItem aa
                                WHERE aa.MtlItemId = a.MtlItemId
                                FOR JSON PATH, ROOT('data')
                            ) CustomerMtlItem
                            FROM PDM.MtlItem a
                        ) a 
                        OUTER APPLY (
                            SELECT TOP 1 x.CustomerMtlItemNo
                            FROM OPENJSON(a.CustomerMtlItem, '$.data')
                            WITH (
                                CustomerMtlItemNo NVARCHAR(50) N'$.CustomerMtlItemNo'
                            ) x
                        ) b";
                    sqlQuery.auxTables = queryTable;
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey", @" AND (a.MtlItemNo LIKE '%' + @SearchKey + '%' 
                                                                                                        OR a.MtlItemName LIKE '%' + @SearchKey + '%' 
                                                                                                        OR b.CustomerMtlItemNo LIKE '%' + @SearchKey + '%')", SearchKey);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.MtlItemId DESC";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    inventoryMtlItems = BaseHelper.SqlQuery<InventoryMtlItem>(sqlConnection, dynamicParameters, sqlQuery);
                }

                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                {
                    List<InventoryData> inventoryDatas = new List<InventoryData>();
                    List<InventoryAvailable> inventoryAvailables = new List<InventoryAvailable>();

                    #region //各庫庫存量
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT LTRIM(RTRIM(x.MC001)) MtlItemNo, LTRIM(RTRIM(y.MC001)) InventoryNo, LTRIM(RTRIM(y.MC002)) InventoryName, x.MC007 InventoryQty
                            FROM INVMC x
                            INNER JOIN CMSMC y ON x.MC002 = y.MC001
                            WHERE LTRIM(RTRIM(x.MC001)) IN @MtlItemNos
                            AND x.MC007 > 0
                            ORDER BY y.MC001";
                    dynamicParameters.Add("MtlItemNos", inventoryMtlItems.Select(x => x.MtlItemNo).ToArray());

                    inventoryDatas = sqlConnection.Query<InventoryData>(sql, dynamicParameters).ToList();

                    inventoryMtlItems
                        .ToList()
                        .ForEach(x =>
                        {
                            x.InventoryDatas = inventoryDatas
                                                    .Where(y => y.MtlItemNo == x.MtlItemNo)
                                                    .Select(y =>
                                                    new InventoryData
                                                    {
                                                        MtlItemNo = y.MtlItemNo,
                                                        InventoryNo = y.InventoryNo,
                                                        InventoryName = y.InventoryName,
                                                        InventoryQty = y.InventoryQty
                                                    }).ToList();
                        });
                    #endregion

                    #region //庫存可用量
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT LTRIM(RTRIM(a.MB001)) MtlItemNo, b.ScheduledPurchase, c.ScheduledIn
                            , d.ScheduledProduce, e.ShippedQuantity, f.AvailableInventory
                            FROM INVMB a
                            OUTER APPLY (
                                SELECT ISNULL(SUM(bb.TB009), 0) ScheduledPurchase
                                FROM PURTA ba
                                INNER JOIN PURTB bb ON ba.TA001 = bb.TB001 AND ba.TA002 = bb.TB002
                                WHERE 1=1
                                AND ba.TA007 = 'Y'
                                AND bb.TB025 = 'Y'
                                AND bb.TB039 NOT IN('Y','y')
                                AND LTRIM(RTRIM(bb.TB004)) = LTRIM(RTRIM(a.MB001))
                            ) b
                            OUTER APPLY (
                                SELECT ISNULL(SUM(cb.TD008 - cb.TD015), 0) ScheduledIn
                                FROM PURTC ca
                                INNER JOIN PURTD cb ON ca.TC001 = cb.TD001 AND ca.TC002 = cb.TD002
                                WHERE 1 = 1
                                AND ca.TC014 = 'Y'
                                AND cb.TD016 NOT IN('Y', 'y')
                                AND cb.TD018 = 'Y'
                                AND LTRIM(RTRIM(cb.TD004)) = LTRIM(RTRIM(a.MB001))
                            ) c
                            OUTER APPLY (
                                SELECT ISNULL(SUM(da.TA015 - da.TA017 - da.TA018), 0) ScheduledProduce
                                FROM MOCTA da
                                WHERE da.TA011 NOT IN('Y', 'y')
                                AND da.TA013 = 'Y'
                                AND LTRIM(RTRIM(da.TA006)) = LTRIM(RTRIM(a.MB001))
                            ) d
                            OUTER APPLY (
                                SELECT ISNULL(SUM(eb.TG009), 0) ShippedQuantity
                                FROM INVTF ea
                                INNER JOIN INVTG eb ON ea.TF001 = eb.TG001 AND ea.TF002 = eb.TG002
                                WHERE ea.TF001 = '1301'
                                AND ea.TF020 = 'Y'
                                AND eb.TG022 = 'Y'
                                AND eb.TG024 = 'N'
                                AND LTRIM(RTRIM(eb.TG004)) = LTRIM(RTRIM(a.MB001))
                            ) e
                            OUTER APPLY (
                                SELECT ISNULL(SUM(fa.MC007), 0) AvailableInventory
                                FROM INVMC fa
                                INNER JOIN CMSMC fb ON fa.MC002 = fb.MC001
                                WHERE fa.MC007 > 0
                                AND fb.MC001 LIKE 'A__%'
                                AND fb.MC001 NOT IN('A05', 'A07')
                                AND LTRIM(RTRIM(fa.MC001)) = LTRIM(RTRIM(a.MB001))
                            ) f
                            WHERE LTRIM(RTRIM(a.MB001)) IN @MtlItemNos";
                    dynamicParameters.Add("MtlItemNos", inventoryMtlItems.Select(x => x.MtlItemNo).ToArray());

                    inventoryAvailables = sqlConnection.Query<InventoryAvailable>(sql, dynamicParameters).ToList();

                    inventoryMtlItems = inventoryMtlItems.GroupJoin(inventoryAvailables, x => x.MtlItemNo, y => y.MtlItemNo
                                            , (x, y) => {
                                                x.ScheduledPurchase = y.FirstOrDefault()?.ScheduledPurchase ?? 0;
                                                x.ScheduledIn = y.FirstOrDefault()?.ScheduledIn ?? 0;
                                                x.ScheduledProduce = y.FirstOrDefault()?.ScheduledProduce ?? 0;
                                                x.ShippedQuantity = y.FirstOrDefault()?.ShippedQuantity ?? 0;
                                                x.AvailableInventory = y.FirstOrDefault()?.AvailableInventory ?? 0;
                                                return x;
                                            }).ToList();
                    #endregion
                }

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    data = inventoryMtlItems
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

        #region //GetTempShippingNoteDoc -- 取得暫出單憑證資訊 -- Yi 2023.10.27
        public string GetTempShippingNoteDoc(int TsnId, string LotSetting)
        {
            try
            {
                string companyNo = "", erpPrefix = "", erpNo = "";

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

                    #region //暫出單據資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.TsnErpPrefix, a.TsnErpNo
                            FROM SCM.TempShippingNote a
                            WHERE TsnId = @TsnId";
                    dynamicParameters.Add("TsnId", TsnId);

                    var resultTsn = sqlConnection.Query(sql, dynamicParameters);
                    if (resultTsn.Count() <= 0) throw new SystemException("暫出單單據資料錯誤!");

                    foreach (var item in resultTsn)
                    {
                        erpPrefix = item.TsnErpPrefix;
                        erpNo = item.TsnErpNo;
                    }
                    #endregion
                }

                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                {
                    #region //INVTF資料 
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT b.MQ002
                            , a.TF001 + ' ' + b.MQ002 TF001Doc
                            , a.TF001, a.TF002
                            , FORMAT(CAST(LTRIM(RTRIM(a.TF024)) as date), 'yyyy/MM/dd') TF024Date
                            , CASE a.TF004
                                WHEN 1 THEN '客戶'
                                WHEN 2 THEN '廠商'
                                WHEN 3 THEN '人員'
                                WHEN 9 THEN '其他'
                            END TF004
                            , a.TF015, a.TF016, a.TF018
                            , a.TF007 + ' ' + d.ME002 TF007Dep, a.TF008 + ' ' + e.MF002 TF008Staff
                            , LTRIM(RTRIM(a.TF009)) + ' ' + c.MB002 TF009Fac, a.TF005
                            , a.TF011
                            , CASE a.TF010
                                WHEN 1 THEN '內含'
                                WHEN 2 THEN '外加'
                                WHEN 3 THEN '零稅率'
                                WHEN 4 THEN '免稅'
                                WHEN 5 THEN '不計稅'
                            END TF010
                            , a.TF013, a.TF014
                            , CONVERT(real, a.TF012) TF012Rate
                            , a.TF026, a.TF020
                            , FORMAT(a.TF022, '#,#.00') TF022Qua
                            , FORMAT(a.TF023, '#,#') TF023Amount
                            , FORMAT(a.TF027, '#,#') TF027Tax
                            , FORMAT(a.TF023 + a.TF027, '#,#') TotalAmount
                            FROM INVTF a
                            INNER JOIN CMSMQ b ON b.MQ001 = a.TF001
                            INNER JOIN CMSMB c ON c.MB001 = a.TF009
                            INNER JOIN CMSME d ON d.ME001 = a.TF007
                            INNER JOIN ADMMF e ON e.MF001 = a.TF008
                            WHERE a.TF001 = @ErpPrefix
                            AND a.TF002 = @ErpNo";
                    dynamicParameters.Add("ErpPrefix", erpPrefix);
                    dynamicParameters.Add("ErpNo", erpNo);

                    var resultCoptc = sqlConnection.Query(sql, dynamicParameters);
                    if (resultCoptc.Count() <= 0) throw new SystemException("ERP暫出單資料錯誤!");
                    #endregion

                    #region //INVTG資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT 
                            a.TG003 Seq
                            , a.TG004 MtlItemNo
                            , a.TG005 MtlItemName
                            , a.TG006 MtlItemSpec
                            , a.TG007 OutStock
                            , a.TG008 InStock
                            , a.TG027 ReturnDate
                            , CONVERT(real, a.TG009) Qty
                            , CONVERT(real, a.TG020) PsQty
                            , CONVERT(real, a.TG044) GpQty
                            , CONVERT(real, a.TG021) RQty
                            , a.TG010 Uom, a.TG011 Muom
                            , a.TG024 Closeout
                            , CAST(a.TG012 as float) UnPri
                            , CAST(a.TG013 as float) Amount
                            , a.TG017 LotNo
                            , a.TG025 ExpDate
                            , a.TG026 RiDate
                            , a.TG014 + '-' + a.TG015 + '-' + a.TG016 SourceDoc
                            , a.TG018 Project
                            , a.TG019 Remark
                            FROM INVTG a
                            WHERE a.TG001 = @ErpPrefix
                            AND a.TG002 = @ErpNo
                            ORDER BY a.TG003";
                    if(LotSetting == "Y")
                    {
                        sql = @"SELECT 
                             FORMAT(ROW_NUMBER() OVER (ORDER BY a.TG004), '0000') AS Seq
                            , a.TG004 MtlItemNo
                            , a.TG005 MtlItemName
                            , a.TG006 MtlItemSpec
                            , SUM(CONVERT(real, a.TG009)) AS Qty
                            , SUM(CONVERT(real, a.TG020)) AS PsQty
                            , SUM(CONVERT(real, a.TG044)) AS GpQty
                            , SUM(CONVERT(real, a.TG021)) AS RQty
                            , CAST(a.TG012 as float) UnPri
                            , CAST(SUM(a.TG013) as float) Amount
                            , a.TG007 OutStock
                            , a.TG008 InStock
                            , a.TG027 ReturnDate
                            , a.TG010 Uom
                            , a.TG011 Muom
                            , a.TG024 Closeout
                            , '' LotNo
                            , '' ExpDate
                            , '' RiDate
                            , a.TG014 + '-' + a.TG015 + '-' + a.TG016 SourceDoc
                            , a.TG018 Project
                            , a.TG019 Remark
                            FROM INVTG a
                            WHERE a.TG001 = @ErpPrefix
                            AND a.TG002 = @ErpNo
                            GROUP BY a.TG004,a.TG005,a.TG006,a.TG007,a.TG008, a.TG027,a.TG010,a.TG011,a.TG024,a.TG012,a.TG014,a.TG015,a.TG016,a.TG018,a.TG019";
                    }
                    dynamicParameters.Add("ErpPrefix", erpPrefix);
                    dynamicParameters.Add("ErpNo", erpNo);

                    var resultCoptd = sqlConnection.Query(sql, dynamicParameters);
                    if (resultCoptd.Count() <= 0) throw new SystemException("ERP暫出單單身資料錯誤!");
                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = resultCoptc,
                        dataDetail = resultCoptd
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

        #region //GetTsnDocDate -- 取得暫出單單據日期資料 -- Yi 2023.11.15
        public string GetTsnDocDate(int TsnId, string ConfirmStatus)
        {
            try
            {
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

                    sql = @"SELECT a.TsnId, a.CompanyId, a.TsnErpPrefix, a.TsnErpNo, FORMAT(a.DocDate, 'yyyy-MM-dd') DocDate
                            , a.ConfirmStatus, a.ConfirmUserId
                            , a.TsnErpPrefix + '-' + a.TsnErpNo TsnErpFullNo
                            FROM SCM.TempShippingNote a
                            WHERE a.CompanyId = @CompanyId
                            AND a.ConfirmStatus = @ConfirmStatus";

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TsnId", @" AND a.TsnId = @TsnId", TsnId);
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    dynamicParameters.Add("ConfirmStatus", 'N');

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

        #region //GetSoSalesInfo -- 取得訂單及暫出單人員資料 -- Yi 2023.11.21
        public string GetSoSalesInfo(string DoDetailId)
        {
            try
            {
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

                    sql = @"SELECT TOP 1 a.DoDetailId, a.DoId, a.SoDetailId, b.SoId,d.SalesmenId,d.DepartmentId
                            FROM SCM.DoDetail a
                            INNER JOIN SCM.SoDetail b ON a.SoDetailId = b.SoDetailId
                            INNER JOIN SCM.DeliveryOrder c ON a.DoId = c.DoId
                            INNER JOIN SCM.SaleOrder d ON b.SoId=d.SoId
                            WHERE a.DoDetailId IN @DoDetailId
                            AND c.CompanyId = @CompanyId";
                    dynamicParameters.Add("DoDetailId", DoDetailId.Split(','));
                    //if (DoDetailId.Length > 0) BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "DoDetailId", @" AND a.DoDetailId IN ( @DoDetailId )", DoDetailId.Split(','));
                    dynamicParameters.Add("CompanyId", CurrentCompany);

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

        #region //GetInventoryAgingReport -- 取得存貨庫齡分析表 -- kitty cat  2024.12.26
        public string GetInventoryAgingReport(string MtlItemNo , string AgingDate, string ChangeType, string DataDate
            , int PageIndex, int PageSize)
        {
            try
            {
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

                    using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                    {
                        #region //新舊品號
                        string NewMtlItemNo = "", NewMtlItemName = "", NewMtlItemSpec = "";
                        string OriMtlItemNo = "", OriMtlItemName = "", OriMtlItemSpec = "";
                        if (MtlItemNo.Length >0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"
                                        SELECT a.NewMtlItemNo, a.NewMtlItemName, a.NewMtlItemSpec
                                        , b.MtlItemNo OriMtlItemNo, b.MtlItemName OriMtlItemName, b.MtlItemSpec OriMtlItemSpec
                                        FROM PDM.MtlItemNoIntegration a
                                        INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
                                        WHERE a.NewMtlItemNo = @MtlItemNo";
                            dynamicParameters.Add("MtlItemNo", MtlItemNo);

                            var MtlItemNoIntegrationResult = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item in MtlItemNoIntegrationResult)
                            {
                                MtlItemNo = item.OriMtlItemNo;
                                NewMtlItemNo = item.NewMtlItemNo;
                                NewMtlItemName = item.NewMtlItemName;
                                NewMtlItemSpec = item.NewMtlItemSpec;
                                OriMtlItemNo = item.OriMtlItemNo;
                                OriMtlItemName = item.OriMtlItemName;
                                OriMtlItemSpec = item.OriMtlItemSpec;
                            }
                        }
                        #endregion

                        string mainSql = "";
                        string middleSql = "";
                        if (ChangeType == "Y")
                        {
                            if (NewMtlItemNo.Length >0 )
                            {
                                middleSql += @"--品號+庫存(新+舊)
                                                                WITH BaseData AS (
                                                                SELECT 
                                                                    LA.LA001,
                                                                    LA.LA009,
                                                                    MC.MC005,
                                                                    (SELECT MAX(LA004) 
                                                                     FROM INVLA AS Sub
                                                                     WHERE Sub.LA001 = LA.LA001
                                                                       AND Sub.LA009 = LA.LA009
                                                                       AND Sub.LA006 IN ('1101', '1102', '1105', '1109', '1106', '1112', 
                                                                                        '1199', '2301', '2302', '2303', '2304', '2305',
                                                                                        '2306', '3411', '3412', '3413', '3414', '3415',
                                                                                        '3416', '3417', '3418', '3419', '5401', '5402',
                                                                                        '5403', '5404', '5405', '5406', '5501', '5502',
                                                                                        '5503', '5901', '5902', '5903', '5904', '5905',
                                                                                        '5906', '5907')    
                                                                    ) AS nLA004,
                                                                    SUM(LA011 * LA005) AS INV_QTY,
                                                                    SUM(LA013 * LA005) AS INV_AMT,
                                                                    SUM(LA021 * LA005) AS PKG_QTY
                                                                FROM 
                                                                    INVLA LA
                                                                JOIN 
                                                                    CMSMC AS MC ON MC.MC001 = LA.LA009
                                                                WHERE 
                                                                    LA.LA001 IN (@MtlItemNo, @NewMtlItemNo)  
                                                                        AND LA.LA004 < @DataDate
                                                                GROUP BY 
                                                                    LA.LA001, LA.LA009, MC.MC005
                                                            ),
                                                            LA_MaxDate AS (
                                                                SELECT 
                                                                    M.first_LA001 AS LA001,
                                                                    B.LA009,
                                                                    B.MC005,
                                                                    M.max_date AS nLA004,
                                                                    M.total_INV_QTY AS INV_QTY,
                                                                    M.total_INV_AMT AS INV_AMT,
                                                                    M.total_PKG_QTY AS PKG_QTY
                                                                FROM BaseData B
                                                                JOIN (
                                                                    SELECT 
                                                                        LA009,
                                                                        MAX(nLA004) as max_date,
                                                                        SUM(INV_QTY) as total_INV_QTY,
                                                                        SUM(INV_AMT) as total_INV_AMT,
                                                                        SUM(PKG_QTY) as total_PKG_QTY,
                                                                        MIN(LA001) as first_LA001
                                                                    FROM BaseData
                                                                    GROUP BY LA009
                                                                ) M ON B.LA009 = M.LA009 AND B.LA001 = M.first_LA001
                                                            )
                                                            SELECT 
                                                                LA.LA001,
                                                                MB.MB002,
                                                                MB.MB003,
                                                                MB.MB004,
                                                                LA.LA009,
                                                                MC.MC002,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) < 30 THEN SUM(LA.LA011 * LA.LA005) ELSE 0 END AS INV_QTY_30D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) < 30 THEN SUM(LA.LA013 * LA.LA005) ELSE 0 END AS INV_AMT_30D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) < 30 THEN SUM(LA.LA021 * LA.LA005) ELSE 0 END AS PKG_QTY_30D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 30 AND 90 THEN SUM(LA.LA011 * LA.LA005) ELSE 0 END AS INV_QTY_30_90D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 30 AND 90 THEN SUM(LA.LA013 * LA.LA005) ELSE 0 END AS INV_AMT_30_90D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 30 AND 90 THEN SUM(LA.LA021 * LA.LA005) ELSE 0 END AS PKG_QTY_30_90D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 90 AND 180 THEN SUM(LA.LA011 * LA.LA005) ELSE 0 END AS INV_QTY_90_180D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 90 AND 180 THEN SUM(LA.LA013 * LA.LA005) ELSE 0 END AS INV_AMT_90_180D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 90 AND 180 THEN SUM(LA.LA021 * LA.LA005) ELSE 0 END AS PKG_QTY_90_180D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 180 AND 270 THEN SUM(LA.LA011 * LA.LA005) ELSE 0 END AS INV_QTY_180_270D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 180 AND 270 THEN SUM(LA.LA013 * LA.LA005) ELSE 0 END AS INV_AMT_180_270D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 180 AND 270 THEN SUM(LA.LA021 * LA.LA005) ELSE 0 END AS PKG_QTY_180_270D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 270 AND 365 THEN SUM(LA.LA011 * LA.LA005) ELSE 0 END AS INV_QTY_270_365D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 270 AND 365 THEN SUM(LA.LA013 * LA.LA005) ELSE 0 END AS INV_AMT_270_365D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 270 AND 365 THEN SUM(LA.LA021 * LA.LA005) ELSE 0 END AS PKG_QTY_270_365D,
                                                                CASE WHEN LA_MaxDate.nLA004 IS NULL OR DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) > 365 THEN SUM(LA.LA011 * LA.LA005) ELSE 0 END AS INV_QTY_365D_PLUS,
                                                                CASE WHEN LA_MaxDate.nLA004 IS NULL OR DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) > 365 THEN SUM(LA.LA013 * LA.LA005) ELSE 0 END AS INV_AMT_365D_PLUS,
                                                                CASE WHEN LA_MaxDate.nLA004 IS NULL OR DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) > 365 THEN SUM(LA.LA021 * LA.LA005) ELSE 0 END AS PKG_QTY_365D_PLUS
                                                            FROM 
                                                                LA_MaxDate
                                                            JOIN 
                                                                INVLA AS LA ON LA.LA001 = LA_MaxDate.LA001 AND LA.LA009 = LA_MaxDate.LA009
                                                            JOIN 
                                                                INVMB AS MB ON MB.MB001 = LA.LA001
                                                            JOIN 
                                                                CMSMC AS MC ON MC.MC001 = LA.LA009
                                                            WHERE 
                                                                (LA_MaxDate.INV_QTY > 0 OR LA_MaxDate.INV_AMT > 0 OR LA_MaxDate.PKG_QTY > 0) 
                                                                AND MC.MC005 IN ('Y', 'N') 
                                                                AND LA.LA009 NOT LIKE 'B%'
                                                            GROUP BY 
                                                                LA.LA001, MB.MB002, MB.MB003, MB.MB004, LA.LA009, MC.MC002, LA_MaxDate.nLA004
                                                            HAVING 
                                                                SUM(LA.LA011 * LA.LA005) > 0
                                                            ORDER BY 
                                                                LA.LA001;";

                                sql = middleSql;
                            }
                            else
                            {
                                mainSql += @"--品號+庫存
                                                            WITH LA_MaxDate AS (
                                                                    SELECT 
                                                                        LA.LA001,
                                                                        LA.LA009,
                                                                        MC.MC005,
                                                                        (SELECT MAX(LA004) 
                                                                         FROM INVLA AS Sub
                                                                         WHERE Sub.LA001 = LA.LA001
                                                                           AND Sub.LA009 = LA.LA009
                                                                           AND Sub.LA006 IN ('1101', '1102', '1105', '1109', '1106', '1112', 
                                                                                             '1199', '2301', '2302', '2303', '2304', '2305',
                                                                                             '2306', '3411', '3412', '3413', '3414', '3415',
                                                                                             '3416', '3417', '3418', '3419', '5401', '5402',
                                                                                             '5403', '5404', '5405', '5406', '5501', '5502',
                                                                                             '5503', '5901', '5902', '5903', '5904', '5905',
                                                                                             '5906', '5907')
                                                                        ) AS nLA004,
                                                                        SUM(LA011 * LA005) AS INV_QTY,
                                                                        SUM(LA013 * LA005) AS INV_AMT,
                                                                        SUM(LA021 * LA005) AS PKG_QTY
                                                                    FROM 
                                                                        INVLA LA
                                                                    JOIN 
                                                                        CMSMC AS MC ON MC.MC001 = LA.LA009
                                                                    WHERE LA.LA004 < @DataDate";
                                if (MtlItemNo.Length > 0)
                                {
                                    mainSql += "  AND LA.LA001 = @MtlItemNo";
                                }

                                mainSql += @" GROUP BY 
                                                                    LA.LA001, LA.LA009, MC.MC005
                                                            )";
                                middleSql = @" SELECT 
                                                                LA.LA001,
                                                                MB.MB002,
                                                                MB.MB003,
                                                                MB.MB004,
                                                                LA.LA009,
                                                                MC.MC002,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) < 30 THEN SUM(LA.LA011 * LA.LA005) ELSE 0 END AS INV_QTY_30D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) < 30 THEN SUM(LA.LA013 * LA.LA005) ELSE 0 END AS INV_AMT_30D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) < 30 THEN SUM(LA.LA021 * LA.LA005) ELSE 0 END AS PKG_QTY_30D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 30 AND 90 THEN SUM(LA.LA011 * LA.LA005) ELSE 0 END AS INV_QTY_30_90D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 30 AND 90 THEN SUM(LA.LA013 * LA.LA005) ELSE 0 END AS INV_AMT_30_90D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 30 AND 90 THEN SUM(LA.LA021 * LA.LA005) ELSE 0 END AS PKG_QTY_30_90D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 90 AND 180 THEN SUM(LA.LA011 * LA.LA005) ELSE 0 END AS INV_QTY_90_180D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 90 AND 180 THEN SUM(LA.LA013 * LA.LA005) ELSE 0 END AS INV_AMT_90_180D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 90 AND 180 THEN SUM(LA.LA021 * LA.LA005) ELSE 0 END AS PKG_QTY_90_180D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 180 AND 270 THEN SUM(LA.LA011 * LA.LA005) ELSE 0 END AS INV_QTY_180_270D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 180 AND 270 THEN SUM(LA.LA013 * LA.LA005) ELSE 0 END AS INV_AMT_180_270D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 180 AND 270 THEN SUM(LA.LA021 * LA.LA005) ELSE 0 END AS PKG_QTY_180_270D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 270 AND 365 THEN SUM(LA.LA011 * LA.LA005) ELSE 0 END AS INV_QTY_270_365D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 270 AND 365 THEN SUM(LA.LA013 * LA.LA005) ELSE 0 END AS INV_AMT_270_365D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 270 AND 365 THEN SUM(LA.LA021 * LA.LA005) ELSE 0 END AS PKG_QTY_270_365D,
                                                                CASE WHEN LA_MaxDate.nLA004 IS NULL OR DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) > 365 THEN SUM(LA.LA011 * LA.LA005) ELSE 0 END AS INV_QTY_365D_PLUS,
                                                                CASE WHEN LA_MaxDate.nLA004 IS NULL OR DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) > 365 THEN SUM(LA.LA013 * LA.LA005) ELSE 0 END AS INV_AMT_365D_PLUS,
                                                                CASE WHEN LA_MaxDate.nLA004 IS NULL OR DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) > 365 THEN SUM(LA.LA021 * LA.LA005) ELSE 0 END AS PKG_QTY_365D_PLUS

                                                            FROM 
                                                                LA_MaxDate
                                                            JOIN 
                                                                INVLA AS LA ON LA.LA001 = LA_MaxDate.LA001 AND LA.LA009 = LA_MaxDate.LA009
                                                            JOIN 
                                                                INVMB AS MB ON MB.MB001 = LA.LA001
                                                            JOIN 
                                                                CMSMC AS MC ON MC.MC001 = LA.LA009
                                                            WHERE 
                                                                (LA_MaxDate.INV_QTY > 0 OR LA_MaxDate.INV_AMT > 0 OR LA_MaxDate.PKG_QTY > 0) 
                                                                AND MC.MC005 IN ('Y', 'N') 
                                                                AND LA.LA009 NOT LIKE 'B%'
                                                            GROUP BY 
                                                                LA.LA001, MB.MB002, MB.MB003, MB.MB004, LA.LA009, MC.MC002, LA_MaxDate.nLA004
                                                            HAVING 
                                                                SUM(LA.LA011 * LA.LA005) > 0";
                                sql = mainSql + middleSql + @" ORDER BY LA.LA001 OFFSET (@PageIndex - 1) * @PageSize ROWS FETCH NEXT @PageSize ROWS ONLY; ";
                            }
                        }
                        else
                        {
                            if (NewMtlItemNo.Length > 0)
                            {
                                middleSql += @"--品號(不含轉撥) (新+舊)
                                                                WITH AggregatedData AS (
                                                                    SELECT 
                                                                        LA.LA001,
                                                                        (SELECT MAX(Sub.LA004)
                                                                         FROM INVLA AS Sub
                                                                         WHERE Sub.LA001 = LA.LA001
                                                                           AND Sub.LA006 IN ('1101', '1102', '1105', '1109', '1106', '1112',
                                                                                              '1199', '2301', '2302', '2303', '2304', '2305',
                                                                                              '2306', '3411', '3412', '3413', '3414', '3415',
                                                                                              '3416', '3417', '3418', '3419', '5401', '5402',
                                                                                              '5403', '5404', '5405', '5406', '5501', '5502',
                                                                                              '5503', '5901', '5902', '5903', '5904', '5905',
                                                                                              '5906', '5907')
                                                                        ) AS nLA004,
                                                                        SUM(LA.LA011 * LA.LA005) AS INV_QTY,
                                                                        SUM(LA.LA013 * LA.LA005) AS INV_AMT,
                                                                        SUM(LA.LA021 * LA.LA005) AS PKG_QTY
                                                                    FROM 
                                                                        INVLA LA
                                                                    WHERE 
                                                                        LA.LA001 IN (@MtlItemNo, @NewMtlItemNo)
                                                                            AND LA.LA004 < @DataDate
                                                                    GROUP BY 
                                                                        LA.LA001
                                                                ),
                                                                LA_MaxDate AS (
                                                                    SELECT 
                                                                        (SELECT TOP 1 LA001 FROM AggregatedData ORDER BY LA001) AS LA001,
                                                                        (SELECT MAX(nLA004) FROM AggregatedData) AS nLA004,
                                                                        (SELECT TOP 1 INV_QTY FROM AggregatedData 
                                                                         ORDER BY CASE 
                                                                                    WHEN INV_QTY > 0 THEN 1
                                                                                    WHEN INV_QTY < 0 THEN 2
                                                                                    ELSE 3 
                                                                                  END) AS INV_QTY,
                                                                        (SELECT TOP 1 INV_AMT FROM AggregatedData 
                                                                         ORDER BY CASE 
                                                                                    WHEN INV_AMT > 0 THEN 1
                                                                                    WHEN INV_AMT < 0 THEN 2
                                                                                    ELSE 3 
                                                                                  END) AS INV_AMT,
                                                                        (SELECT TOP 1 PKG_QTY FROM AggregatedData 
                                                                         ORDER BY CASE 
                                                                                    WHEN PKG_QTY > 0 THEN 1
                                                                                    WHEN PKG_QTY < 0 THEN 2
                                                                                    ELSE 3 
                                                                                  END) AS PKG_QTY
                                                                )
                                                                SELECT 
                                                                    LA_MaxDate.LA001,
                                                                    MAX(MB.MB002) AS MB002,
                                                                    MAX(MB.MB003) AS MB003,
                                                                    MAX(MB.MB004) AS MB004,
                                                                    SUM(CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) < 30 THEN LA.LA011 * LA.LA005 ELSE 0 END) AS INV_QTY_30D,
                                                                    SUM(CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) < 30 THEN LA.LA013 * LA.LA005 ELSE 0 END) AS INV_AMT_30D,
                                                                    SUM(CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) < 30 THEN LA.LA021 * LA.LA005 ELSE 0 END) AS PKG_QTY_30D,
                                                                    SUM(CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 30 AND 90 THEN LA.LA011 * LA.LA005 ELSE 0 END) AS INV_QTY_30_90D,
                                                                    SUM(CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 30 AND 90 THEN LA.LA013 * LA.LA005 ELSE 0 END) AS INV_AMT_30_90D,
                                                                    SUM(CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 30 AND 90 THEN LA.LA021 * LA.LA005 ELSE 0 END) AS PKG_QTY_30_90D,
                                                                    SUM(CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 90 AND 180 THEN LA.LA011 * LA.LA005 ELSE 0 END) AS INV_QTY_90_180D,
                                                                    SUM(CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 90 AND 180 THEN LA.LA013 * LA.LA005 ELSE 0 END) AS INV_AMT_90_180D,
                                                                    SUM(CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 90 AND 180 THEN LA.LA021 * LA.LA005 ELSE 0 END) AS PKG_QTY_90_180D,
                                                                    SUM(CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 180 AND 270 THEN LA.LA011 * LA.LA005 ELSE 0 END) AS INV_QTY_180_270D,
                                                                    SUM(CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 180 AND 270 THEN LA.LA013 * LA.LA005 ELSE 0 END) AS INV_AMT_180_270D,
                                                                    SUM(CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 180 AND 270 THEN LA.LA021 * LA.LA005 ELSE 0 END) AS PKG_QTY_180_270D,
                                                                    SUM(CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 270 AND 365 THEN LA.LA011 * LA.LA005 ELSE 0 END) AS INV_QTY_270_365D,
                                                                    SUM(CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 270 AND 365 THEN LA.LA013 * LA.LA005 ELSE 0 END) AS INV_AMT_270_365D,
                                                                    SUM(CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 270 AND 365 THEN LA.LA021 * LA.LA005 ELSE 0 END) AS PKG_QTY_270_365D,
                                                                    SUM(CASE WHEN LA_MaxDate.nLA004 IS NULL OR DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) > 365 THEN LA.LA011 * LA.LA005 ELSE 0 END) AS INV_QTY_365D_PLUS,
                                                                    SUM(CASE WHEN LA_MaxDate.nLA004 IS NULL OR DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) > 365 THEN LA.LA013 * LA.LA005 ELSE 0 END) AS INV_AMT_365D_PLUS,
                                                                    SUM(CASE WHEN LA_MaxDate.nLA004 IS NULL OR DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) > 365 THEN LA.LA021 * LA.LA005 ELSE 0 END) AS PKG_QTY_365D_PLUS
                                                                FROM 
                                                                    INVLA AS LA
                                                                CROSS JOIN 
                                                                    LA_MaxDate 
                                                                JOIN 
                                                                    INVMB AS MB ON MB.MB001 = LA.LA001
                                                                WHERE 
                                                                    LA.LA001 IN ('1M00-0790-2185-2', '304STAV-0790218501-0')
                                                                    AND (LA_MaxDate.INV_QTY > 0 OR LA_MaxDate.INV_AMT > 0 OR LA_MaxDate.PKG_QTY > 0) 
                                                                    AND LA.LA009 NOT LIKE 'B%'
                                                                GROUP BY 
                                                                    LA_MaxDate.LA001
                                                                HAVING 
                                                                    SUM(LA.LA011 * LA.LA005) > 0;";

                                sql = middleSql;
                            }
                            else
                            {
                                mainSql = @"--品號(不含轉撥)
                                                    WITH LA_MaxDate AS (
                                                        SELECT 
                                                            LA.LA001,
                                                            (SELECT MAX(LA004) 
                                                             FROM INVLA AS Sub
                                                             WHERE Sub.LA001 = LA.LA001
                                                               AND Sub.LA006 IN ('1101', '1102', '1105', '1109', '1106', '1112', 
                                                                                 '1199', '2301', '2302', '2303', '2304', '2305',
                                                                                 '2306', '3411', '3412', '3413', '3414', '3415',
                                                                                 '3416', '3417', '3418', '3419', '5401', '5402',
                                                                                 '5403', '5404', '5405', '5406', '5501', '5502',
                                                                                 '5503', '5901', '5902', '5903', '5904', '5905',
                                                                                 '5906', '5907')    
                                                            ) AS nLA004,
                                                            SUM(LA011 * LA005) AS INV_QTY,
                                                            SUM(LA013 * LA005) AS INV_AMT,
                                                            SUM(LA021 * LA005) AS PKG_QTY
                                                        FROM 
                                                            INVLA LA
                                                                WHERE LA.LA004 < @DataDate";
                                if (MtlItemNo.Length > 0)
                                {
                                    mainSql += " AND  LA.LA001 = @MtlItemNo";
                                }
                                mainSql += @" GROUP BY 
                                                                    LA.LA001
                                                            )";
                                middleSql += @" SELECT 
                                                                LA.LA001,
                                                                MB.MB002,
                                                                MB.MB003,
                                                                MB.MB004,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) < 30 THEN SUM(LA.LA011 * LA.LA005) ELSE 0 END AS INV_QTY_30D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) < 30 THEN SUM(LA.LA013 * LA.LA005) ELSE 0 END AS INV_AMT_30D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) < 30 THEN SUM(LA.LA021 * LA.LA005) ELSE 0 END AS PKG_QTY_30D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 30 AND 90 THEN SUM(LA.LA011 * LA.LA005) ELSE 0 END AS INV_QTY_30_90D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 30 AND 90 THEN SUM(LA.LA013 * LA.LA005) ELSE 0 END AS INV_AMT_30_90D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 30 AND 90 THEN SUM(LA.LA021 * LA.LA005) ELSE 0 END AS PKG_QTY_30_90D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 90 AND 180 THEN SUM(LA.LA011 * LA.LA005) ELSE 0 END AS INV_QTY_90_180D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 90 AND 180 THEN SUM(LA.LA013 * LA.LA005) ELSE 0 END AS INV_AMT_90_180D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 90 AND 180 THEN SUM(LA.LA021 * LA.LA005) ELSE 0 END AS PKG_QTY_90_180D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 180 AND 270 THEN SUM(LA.LA011 * LA.LA005) ELSE 0 END AS INV_QTY_180_270D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 180 AND 270 THEN SUM(LA.LA013 * LA.LA005) ELSE 0 END AS INV_AMT_180_270D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 180 AND 270 THEN SUM(LA.LA021 * LA.LA005) ELSE 0 END AS PKG_QTY_180_270D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 270 AND 365 THEN SUM(LA.LA011 * LA.LA005) ELSE 0 END AS INV_QTY_270_365D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 270 AND 365 THEN SUM(LA.LA013 * LA.LA005) ELSE 0 END AS INV_AMT_270_365D,
                                                                CASE WHEN DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) BETWEEN 270 AND 365 THEN SUM(LA.LA021 * LA.LA005) ELSE 0 END AS PKG_QTY_270_365D,
                                                                CASE WHEN LA_MaxDate.nLA004 IS NULL OR DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) > 365 THEN SUM(LA.LA011 * LA.LA005) ELSE 0 END AS INV_QTY_365D_PLUS,
                                                                CASE WHEN LA_MaxDate.nLA004 IS NULL OR DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) > 365 THEN SUM(LA.LA013 * LA.LA005) ELSE 0 END AS INV_AMT_365D_PLUS,
                                                                CASE WHEN LA_MaxDate.nLA004 IS NULL OR DATEDIFF(DAY, LA_MaxDate.nLA004, GETDATE()) > 365 THEN SUM(LA.LA021 * LA.LA005) ELSE 0 END AS PKG_QTY_365D_PLUS
                                                            FROM 
                                                                INVLA AS LA
                                                            JOIN 
                                                                LA_MaxDate ON LA.LA001 = LA_MaxDate.LA001
                                                            JOIN 
                                                                INVMB AS MB ON MB.MB001 = LA.LA001
                                                            WHERE 
                                                                (LA_MaxDate.INV_QTY > 0 OR LA_MaxDate.INV_AMT > 0 OR LA_MaxDate.PKG_QTY > 0) 
                                                                AND LA.LA009 NOT LIKE 'B%'";

                                middleSql += @" GROUP BY 
                                                                    LA.LA001, MB.MB002, MB.MB003, MB.MB004, LA_MaxDate.nLA004
                                                                HAVING 
                                                                    SUM(LA.LA011 * LA.LA005) > 0";
                                sql = mainSql + middleSql + @"  
                                                                ORDER BY 
                                                                    LA.LA001
                                                                OFFSET (@PageIndex - 1) * @PageSize ROWS FETCH NEXT @PageSize ROWS ONLY;";
                            }
                        }
                        
                        dynamicParameters.Add("MtlItemNo", MtlItemNo);
                        dynamicParameters.Add("NewMtlItemNo", NewMtlItemNo);
                        dynamicParameters.Add("DataDate", DataDate);
                        //if (DoDetailId.Length > 0) BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "DoDetailId", @" AND a.DoDetailId IN ( @DoDetailId )", DoDetailId.Split(','));
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("PageIndex", PageIndex);
                        dynamicParameters.Add("PageSize", PageSize);

                        List<InventoryAgingReport> inventoryAgingReports = sqlConnection2.Query<InventoryAgingReport>(sql, dynamicParameters).ToList();

                        #region //總數量計算
                        int TotalCount = 0;
                        if (NewMtlItemNo.Length > 0)
                        {
                            TotalCount = inventoryAgingReports.Count();
                        }
                        else
                        {
                            sql = mainSql + @"SELECT COUNT(1) AS TotalCount
                                                            FROM (" + middleSql + ") AS TotalQuery";

                            var TotalCountResult = sqlConnection2.Query(sql, dynamicParameters);

                            foreach (var item in TotalCountResult)
                            {
                                TotalCount = item.TotalCount;
                            }
                        }
                        #endregion

                        int i = 0;
                        foreach (var item in inventoryAgingReports)
                        {
                            inventoryAgingReports[i].LA001 = OriMtlItemNo.Length > 0 ? OriMtlItemNo : inventoryAgingReports[i].LA001;
                            inventoryAgingReports[i].MB002 = OriMtlItemName.Length > 0 ? OriMtlItemName : inventoryAgingReports[i].MB002;
                            inventoryAgingReports[i].MB003 = OriMtlItemSpec.Length > 0 ? OriMtlItemSpec : inventoryAgingReports[i].MB003;
                            inventoryAgingReports[i].NewMtlItemNo = NewMtlItemNo;
                            inventoryAgingReports[i].NewMtlItemNo = NewMtlItemNo;
                            inventoryAgingReports[i].NewMtlItemNo = NewMtlItemNo;
                            inventoryAgingReports[i].NewMtlItemName = NewMtlItemName;
                            inventoryAgingReports[i].NewMtlItemSpec = NewMtlItemSpec;
                            i++;
                        }
                        
                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            data = inventoryAgingReports,
                            TotalCount
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
        #region //AddDeliveryToTsn -- 出貨轉暫出資料新增 -- Ben Ma 2023.05.10
        public string AddDeliveryToTsn(string TsnErpPrefix, int SalesmenId, int DepartmentId, string DoDetails, string Remark)

        {
            try
            {
                if (TsnErpPrefix.Length <= 0) throw new SystemException("【單別】不能為空!");
                if (SalesmenId <= 0) throw new SystemException("【業務人員】不能為空!");
                if (DepartmentId <= 0) throw new SystemException("【所屬部門】不能為空!");

                if (!DoDetails.TryParseJson(out JObject tempJObject)) throw new SystemException("出貨資料格式錯誤");

                JObject doJson = JObject.Parse(DoDetails);
                if (!doJson.ContainsKey("data")) throw new SystemException("出貨資料錯誤");

                JToken doData = doJson["data"];
                if (doData.Count() < 0) throw new SystemException("查無出貨資料內容");

                List<DeliveryToTsn> deliveryToTsns = new List<DeliveryToTsn>();

                List<TempShippingNote> tempShippingNotes = new List<TempShippingNote>();
                List<TsnDetail> tsnDetails = new List<TsnDetail>();

                List<INVTF> iNVTFs = new List<INVTF>();
                List<INVTG> iNVTGs = new List<INVTG>();

                //訂單
                List<SaleOrder> saleOrder = new List<SaleOrder>();
                List<SoDetail> soDetail = new List<SoDetail>();

                int rowsAffected = 0, doId = -1;
                string companyNo = "", departmentNo = "", userNo = "", userName = ""
                    , salesDepartmentNo = "", salesUserNo = "", salesUserName = ""
                    , erpPrefix = "", erpNo = "", docDate = "";

                string dateNow = DateTime.Now.ToString("yyyyMMdd");
                string timeNow = DateTime.Now.ToString("HH:mm:ss");
                string CurrentCompanyNo = "";
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷出貨日要同一天
                        List<int> doDetail = new List<int>();

                        for (int i = 0; i < doData.Count(); i++)
                        {
                            doDetail.Add(Convert.ToInt32(doData[i]["doDetailId"]));
                        }

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT DISTINCT b.DoDate
                                FROM SCM.DoDetail a
                                INNER JOIN SCM.DeliveryOrder b ON a.DoId = b.DoId
                                WHERE a.DoDetailId IN @DoDetail
                                AND b.[Status] = @Status
                                AND b.CompanyId = @CompanyId";
                        dynamicParameters.Add("DoDetail", doDetail.ToArray());
                        dynamicParameters.Add("Status", "S");
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var resultDateVerify = sqlConnection.Query(sql, dynamicParameters);
                        if (resultDateVerify.Count() <= 0) throw new SystemException("出貨資料錯誤!");

                        if (resultDateVerify.Count() > 1) throw new SystemException("出貨日期不同!");
                        #endregion

                        for (int i = 0; i < doData.Count(); i++)
                        {
                            #region //判斷出貨資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 b.DoId, CASE WHEN PickFreebieQty > 0 THEN '1' ELSE '2' END ItemType
                                    , ISNULL(c.PickQty, '0') PickQty, ISNULL(d.PickFreebieQty, '0') PickFreebieQty
                                    , ISNULL(e.PickSpareQty, '0') PickSpareQty, f.SoSequence, g.SoErpPrefix, g.SoErpNo
                                    , h.LotManagement, a.SoDetailId, h.MtlItemNo
                                    FROM SCM.DoDetail a
                                    INNER JOIN SCM.DeliveryOrder b ON a.DoId = b.DoId
                                    OUTER APPLY (
                                        SELECT x.ItemType, SUM(x.ItemQty) PickQty
                                        FROM SCM.PickingItem x
                                        WHERE x.SoDetailId = a.SoDetailId
                                        AND x.DoId = b.DoId
                                        AND x.ItemType = 1
                                        GROUP BY x.ItemType
                                    ) c
                                    OUTER APPLY (
                                        SELECT y.ItemType, SUM(y.ItemQty) PickFreebieQty
                                        FROM SCM.PickingItem y
                                        WHERE y.SoDetailId = a.SoDetailId
                                        AND y.DoId = b.DoId
                                        AND y.ItemType = 2
                                        GROUP BY y.ItemType
                                    ) d
                                    OUTER APPLY (
                                        SELECT z.ItemType, SUM(z.ItemQty) PickSpareQty
                                        FROM SCM.PickingItem z
                                        WHERE z.SoDetailId = a.SoDetailId
                                        AND z.DoId = b.DoId
                                        AND z.ItemType = 3
                                        GROUP BY z.ItemType
                                    ) e
                                    INNER JOIN SCM.SoDetail f ON a.SoDetailId = f.SoDetailId
                                    INNER JOIN SCM.SaleOrder g ON f.SoId = g.SoId
                                    INNER JOIN PDM.MtlItem h ON f.MtlItemId = h.MtlItemId
                                    WHERE a.DoDetailId = @DoDetailId
                                    AND b.[Status] = @Status
                                    AND b.CompanyId = @CompanyId";
                            dynamicParameters.Add("DoDetailId", Convert.ToInt32(doData[i]["doDetailId"]));
                            dynamicParameters.Add("Status", "S");
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            var resultDoDetail = sqlConnection.Query(sql, dynamicParameters);
                            if (resultDoDetail.Count() <= 0) throw new SystemException("出貨資料錯誤");

                            string productType = "", lotManagement = "", mtlItemNo = "";
                            int pickRegularQty = -1, pickFreebieQty = -1, pickSpareQty = -1, soDetailId = -1;
                            foreach (var item in resultDoDetail)
                            {
                                productType = item.ItemType;
                                pickRegularQty = item.PickQty;
                                pickFreebieQty = item.PickFreebieQty;
                                pickSpareQty = item.PickSpareQty;
                                lotManagement = item.LotManagement;
                                soDetailId = item.SoDetailId;
                                mtlItemNo = item.MtlItemNo;
                            }
                            #endregion

                            #region //判斷轉出轉入庫別是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.InventoryNo
                                    FROM SCM.Inventory a
                                    WHERE a.InventoryId = @InventoryId
                                    AND a.CompanyId = @CompanyId";
                            dynamicParameters.Add("InventoryId", Convert.ToInt32(doData[i]["tsnOutInventory"]));
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            var resultOutInventory = sqlConnection.Query(sql, dynamicParameters);
                            if (resultOutInventory.Count() <= 0) throw new SystemException("轉出庫資料錯誤");

                            string tsnOutInventoryNo = "";
                            foreach (var item in resultOutInventory)
                            {
                                tsnOutInventoryNo = item.InventoryNo;
                            }

                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 a.InventoryNo
                                    FROM SCM.Inventory a
                                    WHERE a.InventoryId = @InventoryId
                                    AND a.CompanyId = @CompanyId";
                            dynamicParameters.Add("InventoryId", Convert.ToInt32(doData[i]["tsnInInventory"]));
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            var resultInInventory = sqlConnection.Query(sql, dynamicParameters);
                            if (resultInInventory.Count() <= 0) throw new SystemException("轉入庫資料錯誤");

                            string tsnInInventoryNo = "";
                            foreach (var item in resultInInventory)
                            {
                                tsnInInventoryNo = item.InventoryNo;
                            }
                            #endregion

                            foreach (var item in resultDoDetail)
                            {
                                doId = Convert.ToInt32(item.DoId);

                                deliveryToTsns.Add(
                                    new DeliveryToTsn
                                    {
                                        DoDetailId = Convert.ToInt32(doData[i]["doDetailId"]),
                                        SoErpPrefix = item.SoErpPrefix,
                                        SoErpNo = item.SoErpNo,
                                        SoSequence = item.SoSequence,
                                        TsnOutInventory = Convert.ToInt32(doData[i]["tsnOutInventory"]),
                                        TsnOutInventoryNo = tsnOutInventoryNo,
                                        TsnInInventory = Convert.ToInt32(doData[i]["tsnInInventory"]),
                                        TsnInInventoryNo = tsnInInventoryNo,
                                        //PickQty = Convert.ToInt32(doData[i]["tsnQty"]) > Convert.ToInt32(item.PickQty) ? Convert.ToInt32(item.PickQty) : Convert.ToInt32(doData[i]["tsnQty"]),
                                        PickQty = pickRegularQty,
                                        ProductType = productType,
                                        FreebieOrSpareQty = pickFreebieQty > 0 ? pickFreebieQty : pickSpareQty,
                                        LotManagement = lotManagement,
                                        SoDetailId = soDetailId,
                                        MtlItemNo = mtlItemNo
                                    });
                            }
                        }

                        #region //公司別資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb,a.CompanyNo
                                FROM BAS.Company a
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var resultCompany = sqlConnection.Query(sql, dynamicParameters);
                        if (resultCompany.Count() <= 0) throw new SystemException("公司別資料錯誤!");

                        foreach (var item in resultCompany)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            companyNo = item.ErpNo;
                            CurrentCompanyNo = item.CompanyNo;
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

                        #region //業務資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserNo, a.UserName, b.DepartmentNo
                                FROM BAS.[User] a
                                INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                                WHERE a.UserId = @userId";
                        dynamicParameters.Add("userId", SalesmenId);

                        var resultSales = sqlConnection.Query(sql, dynamicParameters);
                        if (resultSales.Count() <= 0) throw new SystemException("業務資料錯誤!");

                        foreach (var item in resultSales)
                        {
                            salesUserNo = item.UserNo;
                            salesUserName = item.UserName;
                            salesDepartmentNo = item.DepartmentNo;
                        }
                        #endregion

                        #region //出貨資料
                        #region //INVTF
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT FORMAT(a.DocDate, 'yyyyMMdd') TF003, '1' TF004, b.CustomerNo TF005
                                , FORMAT(a.DocDate, 'yyyyMMdd') TF024, FORMAT(a.DocDate, 'yyyyMMdd') TF041
                                FROM SCM.DeliveryOrder a
                                INNER JOIN SCM.Customer b ON a.CustomerId = b.CustomerId
                                WHERE a.DoId = @DoId
                                AND a.[Status] = @Status
                                AND a.CompanyId = @CompanyId";
                        dynamicParameters.Add("DoId", doId);
                        dynamicParameters.Add("Status", "S");
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        iNVTFs = sqlConnection.Query<INVTF>(sql, dynamicParameters).ToList();
                        #endregion

                        #region //INVTG
                        foreach (var item in deliveryToTsns)
                        {
                            #region //若撿貨物料有批號，則先找出所有批號，再逐步解析
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.LotNumber
                                    FROM SCM.PickingItem a
                                    WHERE a.SoDetailId = @SoDetailId
                                    GROUP BY a.LotNumber";
                            dynamicParameters.Add("SoDetailId", item.SoDetailId);

                            var LotNumberResult = sqlConnection.Query(sql, dynamicParameters);

                            if (LotNumberResult.Count() <= 0) throw new SystemException("撿貨資料錯誤!!");

                            List<string> lotNumbers = new List<string>();
                            foreach (var item2 in LotNumberResult)
                            {
                                if (item.LotManagement != "N" && item2.LotNumber == null) throw new SystemException("此品號【" + item.MtlItemNo + "】需做批號管理!!");

                                if (item2.LotNumber != null && item2.LotNumber != "")
                                {
                                    lotNumbers.Add(item2.LotNumber);
                                }
                            }
                            #endregion

                            if (lotNumbers.Count() > 0)
                            {
                                foreach (var lotNumber in lotNumbers)
                                {
                                    #region //新增iNVTGs
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT RIGHT('0000' + CONVERT(NVARCHAR, ROW_NUMBER() OVER (ORDER BY a.DoDetailId)), 4) TG003
                                            , e.PickQty TG009
                                            , CASE WHEN PickFreebieQty > 0 THEN '1' ELSE '2' END TG043
                                            , ISNULL(CASE WHEN PickFreebieQty > 0 THEN f.PickFreebieQty ELSE g.PickSpareQty END, '') TG044
                                            , e.PickQty TG052, d.SoErpPrefix TG014
                                            , d.SoErpNo TG015, c.SoSequence TG016
                                            , e.LotNumber TG017
                                            FROM SCM.DoDetail a
                                            INNER JOIN SCM.DeliveryOrder b ON a.DoId = b.DoId
                                            INNER JOIN SCM.SoDetail c ON a.SoDetailId = c.SoDetailId
                                            INNER JOIN SCM.SaleOrder d ON c.SoId = d.SoId
                                            OUTER APPLY (
                                                SELECT SUM(ea.ItemQty) PickQty, ea.LotNumber
                                                FROM SCM.PickingItem ea
                                                WHERE ea.SoDetailId = a.SoDetailId
                                                AND ea.ItemType = 1
                                                AND ea.DoId = a.DoId
                                                AND ea.LotNumber = @LotNumber
                                                GROUP BY ea.ItemQty, ea.LotNumber
                                            ) e
                                            OUTER APPLY (
                                                SELECT SUM(fa.ItemQty) PickFreebieQty
                                                FROM SCM.PickingItem fa
                                                WHERE fa.SoDetailId = a.SoDetailId
                                                AND fa.ItemType = 2
                                                AND fa.DoId = a.DoId
                                                AND fa.LotNumber = @LotNumber
                                            ) f
                                            OUTER APPLY (
                                                SELECT SUM(ga.ItemQty) PickSpareQty
                                                FROM SCM.PickingItem ga
                                                WHERE ga.SoDetailId = a.SoDetailId
                                                AND ga.ItemType = 3
                                                AND ga.DoId = a.DoId
                                                AND ga.LotNumber = @LotNumber
                                            ) g
                                            LEFT JOIN SCM.Inventory h ON a.TransInInventoryId = h.InventoryId
                                            WHERE a.DoDetailId = @DoDetailId
                                            AND b.[Status] = @Status
                                            AND d.CompanyId = @CompanyId
                                            ORDER BY d.SoErpPrefix, d.SoErpNo, c.SoSequence";
                                    dynamicParameters.Add("DoDetailId", item.DoDetailId);
                                    dynamicParameters.Add("Status", "S");
                                    dynamicParameters.Add("CompanyId", CurrentCompany);
                                    dynamicParameters.Add("LotNumber", lotNumber);

                                    var DoDetailResult = sqlConnection.Query<INVTG>(sql, dynamicParameters);
                                    iNVTGs.AddRange(DoDetailResult);
                                    #endregion
                                }
                            }
                            else
                            {
                                #region //新增iNVTGs
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT RIGHT('0000' + CONVERT(NVARCHAR, ROW_NUMBER() OVER (ORDER BY a.DoDetailId)), 4) TG003
                                        , e.PickQty TG009
                                        , CASE WHEN PickFreebieQty > 0 THEN '1' ELSE '2' END TG043
                                        , ISNULL(CASE WHEN PickFreebieQty > 0 THEN f.PickFreebieQty ELSE g.PickSpareQty END, '') TG044
                                        , e.PickQty TG052, d.SoErpPrefix TG014
                                        , d.SoErpNo TG015, c.SoSequence TG016
                                        FROM SCM.DoDetail a
                                        INNER JOIN SCM.DeliveryOrder b ON a.DoId = b.DoId
                                        INNER JOIN SCM.SoDetail c ON a.SoDetailId = c.SoDetailId
                                        INNER JOIN SCM.SaleOrder d ON c.SoId = d.SoId
                                        OUTER APPLY (
                                            SELECT SUM(ea.ItemQty) PickQty
                                            FROM SCM.PickingItem ea
                                            WHERE ea.SoDetailId = a.SoDetailId
                                            AND ea.ItemType = 1
                                            AND ea.DoId = a.DoId
                                        ) e
                                        OUTER APPLY (
                                            SELECT SUM(fa.ItemQty) PickFreebieQty
                                            FROM SCM.PickingItem fa
                                            WHERE fa.SoDetailId = a.SoDetailId
                                            AND fa.ItemType = 2
                                            AND fa.DoId = a.DoId
                                        ) f
                                        OUTER APPLY (
                                            SELECT SUM(ga.ItemQty) PickSpareQty
                                            FROM SCM.PickingItem ga
                                            WHERE ga.SoDetailId = a.SoDetailId
                                            AND ga.ItemType = 3
                                            AND ga.DoId = a.DoId
                                        ) g
                                        LEFT JOIN SCM.Inventory h ON a.TransInInventoryId = h.InventoryId
                                        WHERE a.DoDetailId = @DoDetailId
                                        AND b.[Status] = @Status
                                        AND d.CompanyId = @CompanyId
                                        ORDER BY d.SoErpPrefix, d.SoErpNo, c.SoSequence";
                                dynamicParameters.Add("DoDetailId", item.DoDetailId);
                                dynamicParameters.Add("Status", "S");
                                dynamicParameters.Add("CompanyId", CurrentCompany);

                                var DoDetailResult = sqlConnection.Query<INVTG>(sql, dynamicParameters);
                                iNVTGs.AddRange(DoDetailResult);
                                #endregion
                            }

                            //#region //新增iNVTGs
                            //dynamicParameters = new DynamicParameters();
                            //sql = @"SELECT RIGHT('0000' + CONVERT(NVARCHAR, ROW_NUMBER() OVER (ORDER BY a.DoDetailId)), 4) TG003
                            //        , e.PickQty TG009
                            //        , CASE WHEN PickFreebieQty > 0 THEN '1' ELSE '2' END TG043
                            //        , ISNULL(CASE WHEN PickFreebieQty > 0 THEN f.PickFreebieQty ELSE g.PickSpareQty END, '') TG044
                            //        , e.PickQty TG052, d.SoErpPrefix TG014
                            //        , d.SoErpNo TG015, c.SoSequence TG016
                            //        FROM SCM.DoDetail a
                            //        INNER JOIN SCM.DeliveryOrder b ON a.DoId = b.DoId
                            //        INNER JOIN SCM.SoDetail c ON a.SoDetailId = c.SoDetailId
                            //        INNER JOIN SCM.SaleOrder d ON c.SoId = d.SoId
                            //        OUTER APPLY (
                            //            SELECT SUM(ea.ItemQty) PickQty
                            //            FROM SCM.PickingItem ea
                            //            WHERE ea.SoDetailId = a.SoDetailId
                            //            AND ea.ItemType = 1
                            //            AND ea.DoId = a.DoId
                            //        ) e
                            //        OUTER APPLY (
                            //            SELECT SUM(fa.ItemQty) PickFreebieQty
                            //            FROM SCM.PickingItem fa
                            //            WHERE fa.SoDetailId = a.SoDetailId
                            //            AND fa.ItemType = 2
                            //            AND fa.DoId = a.DoId
                            //        ) f
                            //        OUTER APPLY (
                            //            SELECT SUM(ga.ItemQty) PickSpareQty
                            //            FROM SCM.PickingItem ga
                            //            WHERE ga.SoDetailId = a.SoDetailId
                            //            AND ga.ItemType = 3
                            //            AND ga.DoId = a.DoId
                            //        ) g
                            //        LEFT JOIN SCM.Inventory h ON a.TransInInventoryId = h.InventoryId
                            //        WHERE a.DoDetailId IN @DoDetailId
                            //        AND b.[Status] = @Status
                            //        AND d.CompanyId = @CompanyId
                            //        ORDER BY d.SoErpPrefix, d.SoErpNo, c.SoSequence";
                            //dynamicParameters.Add("DoDetailId", deliveryToTsns.Select(x => x.DoDetailId).ToArray());
                            //dynamicParameters.Add("Status", "S");
                            //dynamicParameters.Add("CompanyId", CurrentCompany);

                            //iNVTGs = sqlConnection.Query<INVTG>(sql, dynamicParameters).ToList();
                            //#endregion
                        }
                        #endregion
                        #endregion

                        #region //基本資料設定
                        #region //INVTF
                        iNVTFs
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

                        #region //INVTG
                        int counter = 1;
                        iNVTGs
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
                                x.TG003 = counter.ToString("D4");
                                counter++;
                            });
                        #endregion
                        #endregion

                        #region //複製未出完定板
                        if (CurrentCompany == 4)
                        {
                            #region //轉暫出定版單頭
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT DISTINCT a.DoId
                                    FROM SCM.DoDetail a
                                    WHERE a.DoDetailId IN @DoDetail";
                            dynamicParameters.Add("DoDetail", doDetail.ToArray());

                            var resultDoList = sqlConnection.Query(sql, dynamicParameters);
                            #endregion

                            foreach (var DoId in resultDoList)
                            {
                                #region //該出貨單所有單身
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.DoDetailId, a.SoDetailId, a.DoQty, a.FreebieQty, a.SpareQty
                                        , ISNULL(b.PickQty,0) PickQty, ISNULL(c.PickFreebieQty,0) PickFreebieQty, ISNULL(d.PickSpareQty,0) PickSpareQty
                                        , e.DoDate
                                        FROM SCM.DoDetail a
                                        OUTER APPLY (
                                            SELECT SUM(ba.ItemQty) PickQty
                                            FROM SCM.PickingItem ba
                                            WHERE ba.SoDetailId = a.SoDetailId
                                            AND ba.ItemType = 1
                                            AND ba.DoId = a.DoId
                                        ) b
                                        OUTER APPLY (
                                            SELECT SUM(ca.ItemQty) PickFreebieQty
                                            FROM SCM.PickingItem ca
                                            WHERE ca.SoDetailId = a.SoDetailId
                                            AND ca.ItemType = 2
                                            AND ca.DoId = a.DoId
                                        ) c
                                        OUTER APPLY (
                                            SELECT SUM(da.ItemQty) PickSpareQty
                                            FROM SCM.PickingItem da
                                            WHERE da.SoDetailId = a.SoDetailId
                                            AND da.ItemType = 3
                                            AND da.DoId = a.DoId
                                        ) d
                                        INNER JOIN SCM.DeliveryOrder e ON a.DoId = e.DoId
                                        WHERE a.DoId = @DoId";
                                dynamicParameters.Add("DoId", DoId.DoId);

                                var resultDoDetail = sqlConnection.Query(sql, dynamicParameters);
                                #endregion

                                foreach (var DoDetail in resultDoDetail)
                                {
                                    string modifier = userNo, soErpPrefix = "", soErpNo = "", soSequence = "";

                                    double UnQty = DoDetail.DoQty - DoDetail.PickQty;
                                    double UnFreebieQty = DoDetail.FreebieQty - DoDetail.PickFreebieQty;
                                    double UnSpareQty = DoDetail.SpareQty - DoDetail.PickSpareQty;
                                    DateTime DoDate = Convert.ToDateTime(DoDetail.DoDate).AddDays(1);

                                    if (UnQty > 0 || UnFreebieQty > 0 || UnSpareQty > 0)
                                    {
                                        #region //交期歷史紀錄新增
                                        #region //撈取原排定交貨日
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT PcPromiseDate
                                                FROM SCM.SoDetail
                                                WHERE SoDetailId = @SoDetailId";
                                        dynamicParameters.Add("SoDetailId", DoDetail.SoDetailId);

                                        var resultPcPromiseDate = sqlConnection.Query(sql, dynamicParameters);

                                        DateTime pcPromiseDate = new DateTime();
                                        foreach (var item in resultPcPromiseDate)
                                        {
                                            pcPromiseDate = item.PcPromiseDate;
                                        }
                                        #endregion

                                        #region //撈取原修改紀錄ID
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT ISNULL(MAX(DeliveryDateLogId), -1) AS ParentLogId
                                                FROM SCM.DeliveryDateLog
                                                WHERE SoDetailId = @SoDetailId";
                                        dynamicParameters.Add("SoDetailId", DoDetail.SoDetailId);

                                        var resultLogId = sqlConnection.Query(sql, dynamicParameters);

                                        int ParentLogId = -1;
                                        foreach (var item in resultLogId)
                                        {
                                            ParentLogId = item.ParentLogId;
                                        }
                                        #endregion

                                        #region //新增交期修改紀錄
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO SCM.DeliveryDateLog (ParentLogId, SoDetailId, PcPromiseDate, DepartmentId
                                                , SupervisorId, CauseType, CauseDescription
                                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                                OUTPUT INSERTED.DeliveryDateLogId
                                                VALUES (@ParentLogId, @SoDetailId, @PcPromiseDate, @DepartmentId
                                                , @SupervisorId, @CauseType, @CauseDescription
                                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                ParentLogId = ParentLogId > 0 ? (int?)ParentLogId : null,
                                                DoDetail.SoDetailId,
                                                PcPromiseDate = pcPromiseDate,
                                                CauseType = (int?)null,
                                                DepartmentId = (int?)null,
                                                SupervisorId = (int?)null,
                                                CauseDescription = "分批出貨",
                                                CreateDate,
                                                LastModifiedDate,
                                                CreateBy,
                                                LastModifiedBy
                                            });
                                        var logResult = sqlConnection.Query(sql, dynamicParameters);

                                        rowsAffected += logResult.Count();
                                        #endregion

                                        #region //修改BM訂單排定交貨日
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"UPDATE SCM.SoDetail SET
                                                PcPromiseDate = @PcPromiseDate,
                                                LastModifiedDate = @LastModifiedDate,
                                                LastModifiedBy = @LastModifiedBy
                                                WHERE SoDetailId = @SoDetailId";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                PcPromiseDate = DoDate,
                                                LastModifiedDate,
                                                LastModifiedBy,
                                                DoDetail.SoDetailId
                                            });

                                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                        #endregion

                                        #region //撈取單別/單號/流水號
                                        sql = @"SELECT b.SoErpPrefix, b.SoErpNo, a.SoSequence 
                                                FROM SCM.SoDetail a
                                                INNER JOIN SCM.SaleOrder b ON b.SoId = a.SoId
                                                WHERE SoDetailId = @SoDetailId";
                                        dynamicParameters.Add("SoDetailId", DoDetail.SoDetailId);

                                        var resultSoDetail = sqlConnection.Query(sql, dynamicParameters);

                                        foreach (var item in resultSoDetail)
                                        {
                                            soErpPrefix = item.SoErpPrefix;
                                            soErpNo = item.SoErpNo;
                                            soSequence = item.SoSequence;
                                        }
                                        #endregion

                                        using (SqlConnection erpSqlConnection = new SqlConnection(ErpConnectionStrings))
                                        {
                                            #region //修改ERP訂單排定交貨日
                                            sql = @"UPDATE COPTD SET
                                                    TD048 = @TD048,
                                                    MODI_DATE = @MODI_DATE,
                                                    MODI_TIME = @MODI_TIME,
                                                    MODIFIER = @MODIFIER
                                                    WHERE TD001 = @TD001
                                                    AND TD002 = @TD002
                                                    AND TD003 = @TD003";
                                            rowsAffected += erpSqlConnection.Execute(sql,
                                                new
                                                {
                                                    TD048 = DoDate.ToString("yyyyMMdd"),
                                                    MODI_DATE = dateNow,
                                                    MODI_TIME = timeNow,
                                                    MODIFIER = modifier,
                                                    TD001 = soErpPrefix,
                                                    TD002 = soErpNo,
                                                    TD003 = soSequence
                                                });
                                            #endregion
                                        }
                                        #endregion

                                        #region //判斷新時段是否有同出貨客戶的出貨單
                                        sql = @"SELECT TOP 1 a.DoId
                                                FROM SCM.DeliveryOrder a
                                                WHERE a.DoDate = @DoDate
                                                AND a.DcId = (
                                                    SELECT y.DcId
                                                    FROM SCM.DoDetail z
                                                    INNER JOIN SCM.DeliveryOrder y ON z.DoId = y.DoId
                                                    WHERE z.DoDetailId = @DoDetailId
                                                )";
                                        dynamicParameters.Add("DoDate", DoDate);
                                        dynamicParameters.Add("DoDetailId", DoDetail.DoDetailId);

                                        var resultExistDeliveryOrder = sqlConnection.Query(sql, dynamicParameters);

                                        int existDoId = -1;
                                        foreach (var item in resultExistDeliveryOrder)
                                        {
                                            existDoId = item.DoId;
                                        }
                                        #endregion

                                        if (existDoId > 0)
                                        {
                                            #region //查詢出貨單流水號
                                            sql = @"SELECT REPLACE(STR(ISNULL(MAX(DoSequence),0) + 1, 4),' ','0') DoSequence 
                                                    FROM SCM.DoDetail
                                                    WHERE DoId = @DoId";
                                            dynamicParameters.Add("DoId", existDoId);

                                            var resultDoSequence = sqlConnection.Query(sql, dynamicParameters);

                                            string DoSequence = "";
                                            foreach (var item in resultDoSequence)
                                            {
                                                DoSequence = item.DoSequence;
                                            }
                                            #endregion

                                            #region //複製出貨單單身
                                            sql = @"INSERT INTO SCM.DoDetail
                                                    SELECT @DoId, a.SoDetailId, @DoSequence, a.TransInInventoryId, @DoQty, @FreebieQty, @SpareQty
                                                    , a.UnitPrice, a.Amount, a.DoDetailRemark, a.WareHouseDoDetailRemark
                                                    , a.PcDoDetailRemark, a.DeliveryProcess, a.DeliveryRoutine, a.DeliveryDocument
                                                    , a.OrderSituation, a.DeliveryMethod, a.ConfirmStatus, a.ClosureStatus
                                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy
                                                    FROM SCM.DoDetail a
                                                    WHERE a.DoDetailId = @DoDetailId";
                                            rowsAffected += sqlConnection.Execute(sql,
                                                new
                                                {
                                                    DoId = existDoId,
                                                    DoSequence,
                                                    TransInInventoryId = -1,
                                                    DoQty = UnQty,
                                                    FreebieQty = UnFreebieQty,
                                                    SpareQty = UnSpareQty,
                                                    CreateDate,
                                                    LastModifiedDate,
                                                    CreateBy,
                                                    LastModifiedBy,
                                                    DoDetail.DoDetailId
                                                });
                                            #endregion
                                        }
                                        else
                                        {
                                            #region //複製出貨單單頭
                                            sql = @"INSERT INTO SCM.DeliveryOrder
                                                    OUTPUT INSERTED.DoId
                                                    SELECT a.CompanyId, a.DepartmentId, a.UserId, a.DoErpPrefix, @DoErpNo
                                                    , @DoDate, @DocDate, a.CustomerId, a.DcId, a.WayBill, a.Traffic, a.ShipMethod
                                                    , a.DoAddressFirst, a.DoAddressSecond, a.TotalQty, a.Amount, a.TaxAmount
                                                    , a.DoRemark, a.WareHouseDoRemark, a.MeasureMailStatus, a.ConfirmStatus, a.ConfirmUserId
                                                    , a.TransferStatus, a.TransferDate, a.Status
                                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy
                                                    FROM SCM.DeliveryOrder a
                                                    INNER JOIN SCM.DoDetail b ON b.DoId = a.DoId
                                                    WHERE DoDetailId = @DoDetailId";
                                            var insertResult = sqlConnection.Query(sql,
                                                new
                                                {
                                                    DoErpNo = BaseHelper.RandomCode(11),
                                                    DoDate,
                                                    DocDate = CreateDate,
                                                    CreateDate,
                                                    LastModifiedDate,
                                                    CreateBy,
                                                    LastModifiedBy,
                                                    DoDetail.DoDetailId
                                                });

                                            rowsAffected += insertResult.Count();

                                            int newDoId = insertResult.Select(x => x.DoId).FirstOrDefault();
                                            #endregion

                                            #region //複製出貨單單身
                                            sql = @"INSERT INTO SCM.DoDetail
                                                    SELECT @DoId, a.SoDetailId, @DoSequence, a.TransInInventoryId, @DoQty, @FreebieQty, @SpareQty
                                                    , a.UnitPrice, a.Amount, a.DoDetailRemark, a.WareHouseDoDetailRemark
                                                    , a.PcDoDetailRemark, a.DeliveryProcess, a.DeliveryRoutine, a.DeliveryDocument
                                                    , a.OrderSituation, a.DeliveryMethod, a.ConfirmStatus, a.ClosureStatus
                                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy
                                                    FROM SCM.DoDetail a
                                                    WHERE a.DoDetailId = @DoDetailId";
                                            rowsAffected += sqlConnection.Execute(sql,
                                                new
                                                {
                                                    DoId = newDoId,
                                                    DoSequence = "0001",
                                                    TransInInventoryId = -1,
                                                    DoQty = UnQty,
                                                    FreebieQty = UnFreebieQty,
                                                    SpareQty = UnSpareQty,
                                                    CreateDate,
                                                    LastModifiedDate,
                                                    CreateBy,
                                                    LastModifiedBy,
                                                    DoDetail.DoDetailId
                                                });
                                            #endregion
                                        }
                                    }
                                }
                            }
                        }
                        #endregion
                    }

                    #region//暫出單 數量檢核 

                    //檢核條件
                    // 1: 訂單正常數<=暫出單正常數+撿貨數-暫出單歸還數
                    // 2: 訂單贈品數<=暫出單贈品數+撿貨贈品數-暫出歸還贈品數
                    // 3: 訂單備品數<=暫出單備品數+撿貨備品數-暫出單備品數
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        for (int i = 0; i < doData.Count(); i++)
                        {
                            #region//正常數卡控
                            int NormalOrderQty = 0; //訂單數
                            int NormalTemporaryOrderQty = 0; //暫出單正常數     
                            int NormalTemporaryReturnOrderQty = 0; //暫出單歸還數
                            int PickQty = 0; //撿貨數

                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT c.SoErpPrefix,c.SoErpNo,b.SoSequence
                                    FROM SCM.DoDetail a
                                    INNER JOIN SCM.SoDetail b ON a.SoDetailId=b.SoDetailId
                                    INNER JOIN SCM.SaleOrder  c ON b.SoId=c.SoId  
                                    WHERE a.DoDetailId=@DoDetailId";
                            dynamicParameters.Add("DoDetailId", Convert.ToInt32(doData[i]["doDetailId"]));
                            var resultSoDetail = sqlConnection.Query(sql, dynamicParameters);
                            if (resultSoDetail.Count() > 0)
                            {
                                foreach (var item in resultSoDetail)
                                {
                                    string SoErpPrefix = item.SoErpPrefix;//訂單單別
                                    string SoErpNo = item.SoErpNo;//訂單單號
                                    string SoSequence = item.SoSequence;//訂單流水號

                                    #region//取出ERP暫出單資訊                                  
                                    using (SqlConnection sqlConnectionErp = new SqlConnection(ErpConnectionStrings))
                                    {
                                        //1.檢查訂單單身數量
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT b.TD001,b.TD002,b.TD003,b.TD008,b.TD009
                                                FROM COPTC a
                                                INNER JOIN COPTD b ON a.TC001=b.TD001 AND a.TC002=b.TD002
                                                WHERE a.TC027 !='V'
                                                AND b.TD001=@SoErpPrefix
                                                AND b.TD002=@SoErpNo
                                                AND b.TD003=@SoSequence";
                                        dynamicParameters.Add("SoErpPrefix", SoErpPrefix);
                                        dynamicParameters.Add("SoErpNo", SoErpNo);
                                        dynamicParameters.Add("SoSequence", SoSequence);
                                        var resultCOPTC = sqlConnectionErp.Query(sql, dynamicParameters);
                                        if (resultCOPTC.Count() > 0)
                                        {
                                            foreach (var itemCOPTC in resultCOPTC)
                                            {
                                                NormalOrderQty = Convert.ToInt32(itemCOPTC.TD008);
                                            }
                                        }

                                        //2.檢查這張訂單有沒有開過暫出單
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT b.TG001,b.TG002,b.TG003,b.TG014,b.TG015,b.TG016,b.TG009
                                                FROM INVTF a
                                                INNER JOIN INVTG b ON a.TF001=b.TG001 AND a.TF002=b.TG002
                                                WHERE a.TF020 !='V'
                                                AND b.TG014=@SoErpPrefix
                                                AND b.TG015=@SoErpNo
                                                AND b.TG016=@SoSequence";
                                        dynamicParameters.Add("SoErpPrefix", SoErpPrefix);
                                        dynamicParameters.Add("SoErpNo", SoErpNo);
                                        dynamicParameters.Add("SoSequence", SoSequence);
                                        var resultINVTF = sqlConnectionErp.Query(sql, dynamicParameters);
                                        if (resultINVTF.Count() > 0)
                                        {
                                            foreach (var itemINVTF in resultINVTF)
                                            {
                                                //3.檢查所開的暫出單有沒有開過【暫出歸還單】
                                                dynamicParameters = new DynamicParameters();
                                                sql = @"SELECT a.TI009
                                                FROM INVTI a
                                                INNER JOIN INVTG b ON a.TI014 = b.TG001 AND a.TI015=b.TG002 AND a.TI016=b.TG003
                                                WHERE a.TI022 !='V'
                                                AND b.TG001=@TsnErpPrefix
                                                AND b.TG002=@TsnErpNo
                                                AND b.TG003=@TsnSequence";
                                                dynamicParameters.Add("TsnErpPrefix", itemINVTF.TG001);
                                                dynamicParameters.Add("TsnErpNo", itemINVTF.TG002);
                                                dynamicParameters.Add("TsnSequence", itemINVTF.TG003);
                                                var resultINVTI = sqlConnectionErp.Query(sql, dynamicParameters);
                                                if (resultINVTI.Count() > 0)
                                                {
                                                    foreach (var itemINVTI in resultINVTI)
                                                    {
                                                        NormalTemporaryReturnOrderQty += Convert.ToInt32(itemINVTI.TI009); //暫出單歸還數
                                                    }
                                                }
                                                else
                                                {
                                                    NormalTemporaryReturnOrderQty = 0;
                                                }
                                                NormalTemporaryOrderQty += Convert.ToInt32(itemINVTF.TG009);
                                            }
                                        }
                                        else
                                        {
                                            NormalTemporaryOrderQty = 0;
                                        }
                                    }
                                    #endregion

                                    #region//本次撿貨數量
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT  ISNULL(e.PickQty,0) PickQty
                                            , ISNULL(f.PickFreebieQty,0) PickFreebieQty
                                            ,ISNULL(g.PickSpareQty,0)PickSpareQty
                                        FROM SCM.DoDetail a
                                            INNER JOIN SCM.DeliveryOrder b ON a.DoId = b.DoId
                                            INNER JOIN SCM.SoDetail c ON a.SoDetailId = c.SoDetailId
                                            INNER JOIN SCM.SaleOrder d ON c.SoId = d.SoId
                                            OUTER APPLY (
                                                SELECT SUM(ea.ItemQty) PickQty
                                                FROM SCM.PickingItem ea
                                                WHERE ea.SoDetailId = a.SoDetailId
                                                AND ea.ItemType = 1
                                                AND ea.DoId = a.DoId
                                            ) e
                                            OUTER APPLY (
                                                SELECT SUM(fa.ItemQty) PickFreebieQty
                                                FROM SCM.PickingItem fa
                                                WHERE fa.SoDetailId = a.SoDetailId
                                                AND fa.ItemType = 2
                                                AND fa.DoId = a.DoId
                                            ) f
                                            OUTER APPLY (
                                                SELECT SUM(ga.ItemQty) PickSpareQty
                                                FROM SCM.PickingItem ga
                                                WHERE ga.SoDetailId = a.SoDetailId
                                                AND ga.ItemType = 3
                                                AND ga.DoId = a.DoId
                                            ) g
                                            LEFT JOIN SCM.Inventory h ON a.TransInInventoryId = h.InventoryId
                                    WHERE a.DoDetailId = @DoDetailId
                                    AND b.[Status] = @Status
                                    AND d.CompanyId = @CompanyId
                                    ORDER BY d.SoErpPrefix, d.SoErpNo, c.SoSequence";
                                    dynamicParameters.Add("DoDetailId", Convert.ToInt32(doData[i]["doDetailId"]));
                                    dynamicParameters.Add("Status", "S");
                                    dynamicParameters.Add("CompanyId", CurrentCompany);
                                    var resultPickingItem = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultPickingItem.Count() > 0)
                                    {
                                        foreach (var itemPickingItem in resultPickingItem)
                                        {
                                            PickQty = Convert.ToInt32(itemPickingItem.PickQty); //撿貨數
                                        }
                                    }
                                    else
                                    {
                                        PickQty = 0;
                                    }
                                    #endregion

                                    if (NormalOrderQty < NormalTemporaryOrderQty + PickQty - NormalTemporaryReturnOrderQty)
                                    {
                                        int Available = NormalOrderQty - NormalTemporaryOrderQty + NormalTemporaryReturnOrderQty;
                                        throw new SystemException("銷貨量超交，本次最大可銷貨量為:[" + Available + "]");
                                    }
                                }

                            }
                            #endregion

                            #region//贈品數卡控
                            int NormalOrderFreebieQty = 0; //訂單贈品數量
                            int TemporaryOrderGiftQty = 0; //暫出單贈品數
                            int TemporaryReturnOrderGiftQty = 0; //暫出歸還贈品數
                            int PickFreebieQty = 0; //撿貨贈品數

                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT c.SoErpPrefix,c.SoErpNo,b.SoSequence
                                    FROM SCM.DoDetail a
                                    INNER JOIN SCM.SoDetail b ON a.SoDetailId=b.SoDetailId
                                    INNER JOIN SCM.SaleOrder  c ON b.SoId=c.SoId  
                                    WHERE a.DoDetailId=@DoDetailId";
                            dynamicParameters.Add("DoDetailId", Convert.ToInt32(doData[i]["doDetailId"]));
                            var resultFreebie = sqlConnection.Query(sql, dynamicParameters);
                            if (resultFreebie.Count() > 0)
                            {
                                foreach (var item in resultFreebie)
                                {
                                    string SoErpPrefix = item.SoErpPrefix;//訂單單別
                                    string SoErpNo = item.SoErpNo;//訂單單號
                                    string SoSequence = item.SoSequence;//訂單流水號

                                    #region//取出ERP暫出單資訊
                                    using (SqlConnection sqlConnectionErp = new SqlConnection(ErpConnectionStrings))
                                    {
                                        //1.檢查訂單單身數量
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT b.TD024
                                                FROM COPTC a
                                                INNER JOIN COPTD b ON a.TC001=b.TD001 AND a.TC002=b.TD002
                                                WHERE a.TC027 !='V'
                                                AND b.TD001=@SoErpPrefix
                                                AND b.TD002=@SoErpNo
                                                AND b.TD003=@SoSequence";
                                        dynamicParameters.Add("SoErpPrefix", SoErpPrefix);
                                        dynamicParameters.Add("SoErpNo", SoErpNo);
                                        dynamicParameters.Add("SoSequence", SoSequence);
                                        var resultCOPTC = sqlConnectionErp.Query(sql, dynamicParameters);
                                        if (resultCOPTC.Count() > 0)
                                        {
                                            foreach (var itemCOPTC in resultCOPTC)
                                            {
                                                NormalOrderFreebieQty = Convert.ToInt32(itemCOPTC.TD024);
                                            }
                                        }

                                        //2.檢查這張訂單沒有開過暫出單
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT b.TG001,b.TG002,b.TG003,b.TG014,b.TG015,b.TG016,b.TG044
                                                FROM INVTF a
                                                INNER JOIN INVTG b ON a.TF001=b.TG001 AND a.TF002=b.TG002
                                                WHERE a.TF020 !='V' AND b.TG043='1'
                                                AND b.TG014=@SoErpPrefix
                                                AND b.TG015=@SoErpNo
                                                AND b.TG016=@SoSequence";
                                        dynamicParameters.Add("SoErpPrefix", SoErpPrefix);
                                        dynamicParameters.Add("SoErpNo", SoErpNo);
                                        dynamicParameters.Add("SoSequence", SoSequence);
                                        var resultINVTF = sqlConnectionErp.Query(sql, dynamicParameters);
                                        if (resultINVTF.Count() > 0)
                                        {
                                            foreach (var itemINVTF in resultINVTF)
                                            {
                                                //3.檢查所開的暫出單有沒有開過【暫出歸還單】
                                                dynamicParameters = new DynamicParameters();
                                                sql = @"SELECT a.TI035
                                                FROM INVTI a
                                                INNER JOIN INVTG b ON a.TI014 = b.TG001 AND a.TI015=b.TG002 AND a.TI016=b.TG003
                                                WHERE a.TI022 !='V' AND a.TI034='1'
                                                AND b.TG001=@TsnErpPrefix
                                                AND b.TG002=@TsnErpNo
                                                AND b.TG003=@TsnSequence";
                                                dynamicParameters.Add("TsnErpPrefix", itemINVTF.TG001);
                                                dynamicParameters.Add("TsnErpNo", itemINVTF.TG002);
                                                dynamicParameters.Add("TsnSequence", itemINVTF.TG003);
                                                var resultINVTI = sqlConnectionErp.Query(sql, dynamicParameters);
                                                if (resultINVTI.Count() > 0)
                                                {
                                                    foreach (var itemINVTI in resultINVTI)
                                                    {
                                                        TemporaryReturnOrderGiftQty += Convert.ToInt32(itemINVTI.TI035); //暫出單歸還數
                                                    }
                                                }
                                                else
                                                {
                                                    TemporaryReturnOrderGiftQty = 0;
                                                }
                                                TemporaryOrderGiftQty += Convert.ToInt32(itemINVTF.TG044);
                                            }
                                        }
                                        else
                                        {
                                            TemporaryOrderGiftQty = 0;
                                        }
                                    }
                                    #endregion

                                    #region//本次撿貨數量
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT  ISNULL(e.PickQty,0) PickQty
                                                , ISNULL(f.PickFreebieQty,0) PickFreebieQty
                                                ,ISNULL(g.PickSpareQty,0)PickSpareQty
                                            FROM SCM.DoDetail a
                                                INNER JOIN SCM.DeliveryOrder b ON a.DoId = b.DoId
                                                INNER JOIN SCM.SoDetail c ON a.SoDetailId = c.SoDetailId
                                                INNER JOIN SCM.SaleOrder d ON c.SoId = d.SoId
                                                OUTER APPLY (
                                                    SELECT SUM(ea.ItemQty) PickQty
                                                    FROM SCM.PickingItem ea
                                                    WHERE ea.SoDetailId = a.SoDetailId
                                                    AND ea.ItemType = 1
                                                    AND ea.DoId = a.DoId
                                                ) e
                                                OUTER APPLY (
                                                    SELECT SUM(fa.ItemQty) PickFreebieQty
                                                    FROM SCM.PickingItem fa
                                                    WHERE fa.SoDetailId = a.SoDetailId
                                                    AND fa.ItemType = 2
                                                    AND fa.DoId = a.DoId
                                                ) f
                                                OUTER APPLY (
                                                    SELECT SUM(ga.ItemQty) PickSpareQty
                                                    FROM SCM.PickingItem ga
                                                    WHERE ga.SoDetailId = a.SoDetailId
                                                    AND ga.ItemType = 3
                                                    AND ga.DoId = a.DoId
                                                ) g
                                                LEFT JOIN SCM.Inventory h ON a.TransInInventoryId = h.InventoryId
                                        WHERE a.DoDetailId = @DoDetailId                                       
                                        AND b.[Status] = @Status
                                        AND d.CompanyId = @CompanyId
                                        AND d.CompanyId = @CompanyId
                                        ORDER BY d.SoErpPrefix, d.SoErpNo, c.SoSequence";
                                    dynamicParameters.Add("DoDetailId", Convert.ToInt32(doData[i]["doDetailId"]));
                                    dynamicParameters.Add("Status", "S");
                                    dynamicParameters.Add("CompanyId", CurrentCompany);
                                    var resultPickingItem = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultPickingItem.Count() > 0)
                                    {
                                        foreach (var itemPickingItem in resultPickingItem)
                                        {
                                            PickFreebieQty = Convert.ToInt32(item.PickFreebieQty); //撿貨贈品數
                                        }
                                    }
                                    else
                                    {
                                        PickFreebieQty = 0;
                                    }
                                    #endregion\                                   

                                    if (NormalOrderFreebieQty < TemporaryOrderGiftQty + PickFreebieQty - TemporaryReturnOrderGiftQty)
                                    {
                                        int Available = NormalOrderFreebieQty - TemporaryOrderGiftQty + TemporaryReturnOrderGiftQty;
                                        throw new SystemException("贈品量超交，本次最大可贈品量為:[" + Available + "]");
                                    }

                                }
                            }
                            #endregion

                            #region//備品數卡控
                            int NormalOrderSpareQty = 0; //訂單備品數量
                            int TemporaryQrderSparePartsQty = 0; //暫出單備品數
                            int TemporaryReturnQrderSparePartsQty = 0; //暫出歸還備品數 
                            int PickSpareQty = 0;//撿貨備品數

                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT c.SoErpPrefix,c.SoErpNo,b.SoSequence
                                    FROM SCM.DoDetail a
                                    INNER JOIN SCM.SoDetail b ON a.SoDetailId=b.SoDetailId
                                    INNER JOIN SCM.SaleOrder  c ON b.SoId=c.SoId  
                                    WHERE a.DoDetailId=@DoDetailId";
                            dynamicParameters.Add("DoDetailId", Convert.ToInt32(doData[i]["doDetailId"]));
                            var resultSpare = sqlConnection.Query(sql, dynamicParameters);
                            if (resultSpare.Count() > 0)
                            {
                                foreach (var item in resultSpare)
                                {
                                    string SoErpPrefix = item.SoErpPrefix;//訂單單別
                                    string SoErpNo = item.SoErpNo;//訂單單號
                                    string SoSequence = item.SoSequence;//訂單流水號

                                    #region//取出ERP暫出歸還數量                                    
                                    using (SqlConnection sqlConnectionErp = new SqlConnection(ErpConnectionStrings))
                                    {
                                        //1.檢查訂單單身數量
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT b.TD050
                                                FROM COPTC a
                                                INNER JOIN COPTD b ON a.TC001=b.TD001 AND a.TC002=b.TD002
                                                WHERE a.TC027 !='V'
                                                AND b.TD001=@SoErpPrefix
                                                AND b.TD002=@SoErpNo
                                                AND b.TD003=@SoSequence";
                                        dynamicParameters.Add("SoErpPrefix", SoErpPrefix);
                                        dynamicParameters.Add("SoErpNo", SoErpNo);
                                        dynamicParameters.Add("SoSequence", SoSequence);
                                        var resultCOPTC = sqlConnectionErp.Query(sql, dynamicParameters);
                                        if (resultCOPTC.Count() > 0)
                                        {
                                            foreach (var itemCOPTC in resultCOPTC)
                                            {
                                                NormalOrderSpareQty = Convert.ToInt32(itemCOPTC.TD050);
                                            }
                                        }

                                        //2.檢查這張訂單沒有開過暫出單
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT b.TG001,b.TG002,b.TG003,b.TG014,b.TG015,b.TG016,b.TG044
                                                FROM INVTF a
                                                INNER JOIN INVTG b ON a.TF001=b.TG001 AND a.TF002=b.TG002
                                                WHERE a.TF020 !='V' AND b.TG043='2'
                                                AND b.TG014=@SoErpPrefix
                                                AND b.TG015=@SoErpNo
                                                AND b.TG016=@SoSequence";
                                        dynamicParameters.Add("SoErpPrefix", SoErpPrefix);
                                        dynamicParameters.Add("SoErpNo", SoErpNo);
                                        dynamicParameters.Add("SoSequence", SoSequence);
                                        var resultINVTF = sqlConnectionErp.Query(sql, dynamicParameters);
                                        if (resultINVTF.Count() > 0)
                                        {
                                            foreach (var itemINVTF in resultINVTF)
                                            {
                                                //3.檢查所開的暫出單有沒有開過【暫出歸還單】
                                                dynamicParameters = new DynamicParameters();
                                                sql = @"SELECT a.TI035
                                                FROM INVTI a
                                                INNER JOIN INVTG b ON a.TI014 = b.TG001 AND a.TI015=b.TG002 AND a.TI016=b.TG003
                                                WHERE a.TI022 !='V' AND a.TI034='2'
                                                AND b.TG001=@TsnErpPrefix
                                                AND b.TG002=@TsnErpNo
                                                AND b.TG003=@TsnSequence";
                                                dynamicParameters.Add("TsnErpPrefix", itemINVTF.TG001);
                                                dynamicParameters.Add("TsnErpNo", itemINVTF.TG002);
                                                dynamicParameters.Add("TsnSequence", itemINVTF.TG003);
                                                var resultINVTI = sqlConnectionErp.Query(sql, dynamicParameters);
                                                if (resultINVTI.Count() > 0)
                                                {
                                                    foreach (var itemINVTI in resultINVTI)
                                                    {
                                                        TemporaryReturnQrderSparePartsQty += Convert.ToInt32(itemINVTI.TI035); //暫出單歸還數
                                                    }
                                                }
                                                else
                                                {
                                                    TemporaryReturnQrderSparePartsQty = 0;
                                                }
                                                TemporaryQrderSparePartsQty += Convert.ToInt32(itemINVTF.TG044);
                                            }
                                        }
                                        else
                                        {
                                            TemporaryQrderSparePartsQty = 0;
                                        }

                                    }
                                    #endregion

                                    #region//本次撿貨數量
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT  ISNULL(e.PickQty,0) PickQty
                                                , ISNULL(f.PickFreebieQty,0) PickFreebieQty
                                                ,ISNULL(g.PickSpareQty,0)PickSpareQty
                                            FROM SCM.DoDetail a
                                                INNER JOIN SCM.DeliveryOrder b ON a.DoId = b.DoId
                                                INNER JOIN SCM.SoDetail c ON a.SoDetailId = c.SoDetailId
                                                INNER JOIN SCM.SaleOrder d ON c.SoId = d.SoId
                                                OUTER APPLY (
                                                    SELECT SUM(ea.ItemQty) PickQty
                                                    FROM SCM.PickingItem ea
                                                    WHERE ea.SoDetailId = a.SoDetailId
                                                    AND ea.ItemType = 1
                                                    AND ea.DoId = a.DoId
                                                ) e
                                                OUTER APPLY (
                                                    SELECT SUM(fa.ItemQty) PickFreebieQty
                                                    FROM SCM.PickingItem fa
                                                    WHERE fa.SoDetailId = a.SoDetailId
                                                    AND fa.ItemType = 2
                                                    AND fa.DoId = a.DoId
                                                ) f
                                                OUTER APPLY (
                                                    SELECT SUM(ga.ItemQty) PickSpareQty
                                                    FROM SCM.PickingItem ga
                                                    WHERE ga.SoDetailId = a.SoDetailId
                                                    AND ga.ItemType = 3
                                                    AND ga.DoId = a.DoId
                                                ) g
                                                LEFT JOIN SCM.Inventory h ON a.TransInInventoryId = h.InventoryId
                                        WHERE a.DoDetailId = @DoDetailId
                                        AND b.[Status] = @Status
                                        AND d.CompanyId = @CompanyId
                                        ORDER BY d.SoErpPrefix, d.SoErpNo, c.SoSequence";
                                    dynamicParameters.Add("DoDetailId", Convert.ToInt32(doData[i]["doDetailId"]));
                                    dynamicParameters.Add("Status", "S");
                                    dynamicParameters.Add("CompanyId", CurrentCompany);
                                    var resultPickingItem = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultPickingItem.Count() > 0)
                                    {
                                        foreach (var itemPickingItem in resultPickingItem)
                                        {
                                            PickSpareQty = Convert.ToInt32(itemPickingItem.PickSpareQty); //撿貨數
                                        }
                                    }
                                    else
                                    {
                                        PickSpareQty = 0;
                                    }
                                    #endregion

                                    if (NormalOrderSpareQty < TemporaryQrderSparePartsQty + PickSpareQty - TemporaryReturnQrderSparePartsQty)
                                    {
                                        int Available = NormalOrderSpareQty - TemporaryQrderSparePartsQty + TemporaryReturnQrderSparePartsQty;
                                        throw new SystemException("備品量超交，本次最大可備品量為:[" + Available + "]");
                                    }
                                }
                            }
                            #endregion
                        }
                    }
                    #endregion

                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        //erpPrefix = "1301";
                        erpPrefix = TsnErpPrefix;
                        docDate = iNVTFs.Select(x => x.TF024).FirstOrDefault();
                        DateTime referenceTime = DateTime.ParseExact(docDate, "yyyyMMdd", CultureInfo.InvariantCulture);

                        #region //單據設定
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MQ004, a.MQ005, a.MQ006, a.MQ044
                                FROM CMSMQ a
                                WHERE a.COMPANY = @CompanyNo
                                AND a.MQ001 = @ErpPrefix";
                        dynamicParameters.Add("CompanyNo", companyNo);
                        dynamicParameters.Add("ErpPrefix", erpPrefix);

                        var resultDocSetting = sqlConnection.Query(sql, dynamicParameters);
                        if (resultDocSetting.Count() <= 0) throw new SystemException("ERP單據設定資料不存在!");

                        string encode = "", exchangeRateSource = "";
                        int yearLength = 0, lineLength = 0;
                        foreach (var item in resultDocSetting)
                        {
                            encode = item.MQ004; //編碼方式
                            yearLength = Convert.ToInt32(item.MQ005); //年碼數
                            lineLength = Convert.ToInt32(item.MQ006); //流水號碼數
                            exchangeRateSource = item.MQ044; //匯率來源
                        }
                        #endregion

                        #region //單號取號
                        int currentNum = 0;

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT CONVERT(int, RIGHT(ISNULL(MAX(RTRIM(LTRIM(TF002))), '" + new string('0', lineLength) + @"'), " + lineLength + @")) + 1 CurrentNum
                                FROM INVTF
                                WHERE TF001 = @ErpPrefix";
                        dynamicParameters.Add("ErpPrefix", erpPrefix);

                        #region //編碼方式
                        string dateFormat = "";
                        switch (encode)
                        {
                            case "1": //日編
                                dateFormat = new string('y', yearLength) + "MMdd";
                                sql += @" AND RTRIM(LTRIM(TF002)) LIKE @ReferenceTime + '" + new string('_', lineLength) + @"'";
                                dynamicParameters.Add("ReferenceTime", referenceTime.ToString(dateFormat));
                                erpNo = referenceTime.ToString(dateFormat);
                                break;
                            case "2": //月編
                                dateFormat = new string('y', yearLength) + "MM";
                                sql += @" AND RTRIM(LTRIM(TF002)) LIKE @ReferenceTime + '" + new string('_', lineLength) + @"'";
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
                        #endregion

                        #region //廠別資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 MB001
                                FROM CMSMB
                                WHERE COMPANY = @COMPANY";
                        dynamicParameters.Add("COMPANY", companyNo);

                        var resultFactory = sqlConnection.Query(sql, dynamicParameters);
                        if (resultFactory.Count() <= 0) throw new SystemException("ERP廠別資料不存在!");

                        string factory = "";
                        foreach (var item in resultFactory)
                        {
                            factory = item.MB001; //廠別
                        }
                        #endregion

                        #region //客戶資料
                        string customerNo = iNVTFs.Select(x => x.TF005).FirstOrDefault();

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT MA002, MA003, MA014, MA027, MA038, MA064, MA118
                                FROM COPMA
                                WHERE MA001 = @CustomerNo";
                        dynamicParameters.Add("CustomerNo", customerNo);

                        var resultCustomer = sqlConnection.Query(sql, dynamicParameters);
                        if (resultCustomer.Count() <= 0) throw new SystemException("ERP客戶資料不存在!");

                        string currency = "", taxation = "", taxNo = "";
                        foreach (var item in resultCustomer)
                        {
                            currency = item.MA014; //交易幣別
                            taxation = item.MA038; //課稅別
                            taxNo = item.MA118; //稅別碼

                            iNVTFs
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.TF006 = item.MA002; //對象簡稱
                                    x.TF010 = taxation; //課稅別
                                    x.TF011 = currency; //幣別
                                    x.TF015 = item.MA003; //對象全名
                                    x.TF016 = item.MA027; //地址一
                                    x.TF017 = item.MA064; //地址二
                                });
                        }
                        #endregion

                        #region //交易幣別設定
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT MF003, MF004
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
                        sql = @"SELECT NN004
                                FROM CMSNN
                                WHERE NN001 = @TaxNo";
                        dynamicParameters.Add("TaxNo", taxNo);

                        var resultTax = sqlConnection.Query(sql, dynamicParameters);
                        if (resultTax.Count() <= 0) throw new SystemException("ERP稅別碼設定不存在!");

                        double exciseTax = 0;
                        foreach (var item in resultTax)
                        {
                            exciseTax = Convert.ToDouble(item.NN004); //營業稅率
                        }
                        #endregion

                        #region //判斷訂單是否存在
                        foreach (var so in iNVTGs)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(b.MB001)) MB001, a.TD005, a.TD006, a.TD010, a.TD011
                                    , a.TD020, a.TD027
                                    FROM COPTD a
                                    INNER JOIN INVMB b ON a.TD004 = b.MB001
                                    WHERE a.TD001 = @ErpPrefix
                                    AND a.TD002 = @ErpNo
                                    AND a.TD003 = @ErpSequence";
                            dynamicParameters.Add("ErpPrefix", so.TG014);
                            dynamicParameters.Add("ErpNo", so.TG015);
                            dynamicParameters.Add("ErpSequence", so.TG016);

                            var resultSaleOrder = sqlConnection.Query(sql, dynamicParameters);
                            if (resultSaleOrder.Count() <= 0) throw new SystemException("ERP訂單資料不存在!");

                            foreach (var item in resultSaleOrder)
                            {
                                #region//檢核品號於轉出庫的數量是否足夠
                                using (SqlConnection sqlConnectionErp = new SqlConnection(ErpConnectionStrings))
                                {
                                    string checkOutInventoryNo = deliveryToTsns.Where(y => y.SoErpPrefix == so.TG014 && y.SoErpNo == so.TG015 && y.SoSequence == so.TG016).Select(y => y.TsnOutInventoryNo).FirstOrDefault();
                                    string checkMtlItemNo = item.MB001;
                                }
                                #endregion

                                #region //設定 INVTG
                                iNVTGs
                                .Where(x => x.TG014 == so.TG014 && x.TG015 == so.TG015 && x.TG016 == so.TG016)
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.TG001 = erpPrefix; //異動單別
                                    x.TG002 = erpNo; //異動單號
                                    x.TG004 = item.MB001; //品號
                                    x.TG005 = item.TD005; //品名
                                    x.TG006 = item.TD006; //規格
                                    x.TG007 = deliveryToTsns.Where(y => y.SoErpPrefix == so.TG014 && y.SoErpNo == so.TG015 && y.SoSequence == so.TG016).Select(y => y.TsnOutInventoryNo).FirstOrDefault(); //轉出庫別
                                    x.TG008 = deliveryToTsns.Where(y => y.SoErpPrefix == so.TG014 && y.SoErpNo == so.TG015 && y.SoSequence == so.TG016).Select(y => y.TsnInInventoryNo).FirstOrDefault(); //轉入庫別
                                                                                                                                                                                                          //x.TG009 = deliveryToTsns.Where(y => y.SoErpPrefix == so.TG014 && y.SoErpNo == so.TG015 && y.SoSequence == so.TG016).Select(y => y.PickQty).FirstOrDefault(); //數量
                                    x.TG010 = item.TD010; //單位
                                    x.TG011 = ""; //小單位
                                    x.TG012 = Math.Round(Convert.ToDouble(item.TD011), unitRound); //單價
                                    x.TG013 = Math.Round(Math.Round(Convert.ToDouble(item.TD011), unitRound) * x.TG009, totalRound); //金額
                                    x.TG018 = item.TD027; //專案代號
                                    x.TG019 = item.TD020; //備註
                                    x.TG020 = 0; //轉進銷量
                                    x.TG021 = 0; //歸還量
                                    x.TG022 = "N"; //確認碼
                                    x.TG023 = "N"; //更新碼
                                    x.TG024 = "N"; //結案碼
                                    x.TG025 = ""; //有效日期
                                    x.TG026 = ""; //複檢日期
                                    x.TG027 = ""; //預計歸還日
                                    x.TG028 = 0; //包裝數量
                                    x.TG029 = 0; //轉進銷包裝量
                                    x.TG030 = 0; //歸還包裝量
                                    x.TG031 = ""; //包裝單位
                                    x.TG032 = ""; //預留欄位
                                    x.TG033 = 0; //產品序號數量
                                    x.TG034 = 0; //預留欄位
                                    x.TG035 = ""; //轉出儲位
                                    x.TG036 = ""; //轉入儲位
                                    x.TG037 = 0; //預留欄位
                                    x.TG038 = 0; //預留欄位
                                    x.TG039 = ""; //預留欄位
                                    x.TG040 = ""; //最終客戶代號
                                    x.TG041 = ""; //預留欄位
                                    x.TG042 = exciseTax; //營業稅率
                                                         //x.TG043 = deliveryToTsns.Where(y => y.SoErpPrefix == so.TG014 && y.SoErpNo == so.TG015 && y.SoSequence == so.TG016).Select(y => y.ProductType).FirstOrDefault(); //類型
                                                         //x.TG044 = deliveryToTsns.Where(y => y.SoErpPrefix == so.TG014 && y.SoErpNo == so.TG015 && y.SoSequence == so.TG016).Select(y => y.FreebieOrSpareQty).FirstOrDefault(); //贈/備品量
                                    x.TG045 = 0; //贈/備品包裝量
                                    x.TG046 = 0; //轉銷贈/備品量
                                    x.TG047 = 0; //轉銷贈/備品包裝量
                                    x.TG048 = 0; //歸還贈/備品量
                                    x.TG049 = 0; //歸還贈/備品包裝量
                                    x.TG050 = ""; //業務品號
                                    x.TG500 = ""; //刻號/BIN記錄
                                    x.TG501 = ""; //刻號管理
                                    x.TG502 = ""; //DATECODE
                                    x.TG503 = ""; //產品系列
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
                                    x.TG051 = ""; //稅別碼
                                                  //x.TG052 = x.TG009; //計價數量
                                    x.TG053 = item.TD010; //計價單位
                                    x.TG054 = 0;
                                    x.TG055 = 0;
                                });
                                #endregion
                            }
                        }
                        #endregion

                        #region //計算數量與金額
                        int? totalQty = iNVTGs.Sum(i => i.TG009) + iNVTGs.Sum(i => i.TG044);
                        double? totalPrice = iNVTGs.Sum(i => i.TG013);
                        double? totalTax = Math.Round((double)totalPrice * exciseTax, totalRound);

                        switch (taxation)
                        {
                            case "1": //應稅內含
                                totalTax = totalPrice - Math.Round((double)totalPrice / (1 + exciseTax), totalRound);
                                totalPrice = Math.Round((double)totalPrice / (1 + exciseTax), totalRound);
                                break;
                            case "2": //應稅外加
                                break;
                            case "3": //零稅率
                                break;
                            case "4": //免稅
                                break;
                            case "9": //不計稅
                                break;
                        }
                        #endregion

                        #region //設定 INVTF
                        #region //判斷單號是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM INVTF
                                WHERE TF001 = @ErpPrefix
                                AND TF002 = @ErpNo";
                        dynamicParameters.Add("ErpPrefix", erpPrefix);
                        dynamicParameters.Add("ErpNo", erpNo);

                        var resultRepeatExist = sqlConnection.Query(sql, dynamicParameters);
                        if (resultRepeatExist.Count() > 0) throw new SystemException("【異動單號】重複，請重新取號!");
                        #endregion

                        iNVTFs
                            .ToList()
                            .ForEach(x =>
                            {
                                x.TF001 = erpPrefix; //異動單別
                                x.TF002 = erpNo; //異動單號
                                x.TF007 = salesDepartmentNo; //部門代號
                                x.TF008 = salesUserNo; //員工代號
                                x.TF009 = factory; //廠別
                                x.TF012 = exchangeRate; //匯率
                                x.TF013 = 0; //件數
                                x.TF014 = Remark; //備註
                                x.TF018 = ""; //其它備註
                                x.TF019 = 0; //列印次數
                                x.TF020 = "N"; //確認碼
                                x.TF021 = "N"; //更新碼
                                x.TF022 = totalQty; //總數量
                                x.TF023 = totalPrice; //總金額
                                x.TF025 = ""; //確認者
                                x.TF026 = exciseTax; //營業稅率
                                x.TF027 = totalTax; //稅額
                                x.TF028 = 0; //總包裝數量
                                x.TF029 = "N"; //簽核狀態碼
                                x.TF030 = 0; //傳送次數
                                x.TF031 = ""; //運輸方式
                                x.TF032 = ""; //派車單別
                                x.TF033 = ""; //派車單號
                                x.TF034 = 0; //預留欄位
                                x.TF035 = 0; //預留欄位
                                x.TF036 = ""; //來源
                                x.TF037 = ""; //銷貨單單價別
                                x.TF038 = ""; //預留欄位
                                x.TF039 = taxNo; //稅別碼
                                x.TF040 = "N"; //單身多稅率
                                x.TF042 = ""; //轉銷日
                                x.TF043 = "N"; //出通單更新碼
                                x.TF044 = "N"; //不控管信用額度
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
                                x.TF045 = "";
                                x.TF046 = "";
                            });
                        #endregion

                        #region//暫出單 ERP信用額度卡控
                        string CompanyNo = CurrentCompanyNo;
                        string CustomerNo = iNVTFs.Select(x => x.TF005).FirstOrDefault(); //客戶代碼
                        string Currency = currency;
                        string DocType = "TempShippingNote";
                        decimal Amount = 0, TotalAmount = Convert.ToDecimal(totalPrice);

                        #region 呼叫信用額度卡控API
                        try
                        {
                            using (HttpClient client = new HttpClient())
                            {
                                string apiUrl = "https://bm.zy-tech.com.tw/api/BM/CheckCreditLimit";
                                client.Timeout = TimeSpan.FromMinutes(60);
                                using (HttpRequestMessage httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, apiUrl))
                                {
                                    MultipartFormDataContent formData = new MultipartFormDataContent
                                {
                                    { new StringContent(CustomerNo.ToString()), "CustomerNo" },
                                    { new StringContent(Currency.ToString()), "Currency" },
                                    { new StringContent(DocType.ToString()), "DocType" },
                                    { new StringContent(Amount.ToString()), "Amount" },
                                    { new StringContent(TotalAmount.ToString()), "TotalAmount" },
                                    { new StringContent(CompanyNo.ToString()), "CompanyNo" }
                                };
                                    httpRequestMessage.Content = formData;
                                    using (HttpResponseMessage httpResponseMessage = client.SendAsync(httpRequestMessage).GetAwaiter().GetResult())
                                    {
                                        if (httpResponseMessage.IsSuccessStatusCode)
                                        {
                                            string Result = httpResponseMessage.Content.ReadAsStringAsync().Result.ToString();
                                            if (JObject.Parse(Result)["status"].ToString() == "errorForDA")
                                            {
                                                return httpResponseMessage.Content.ReadAsStringAsync().Result.ToString();
                                            }
                                        }
                                        else
                                        {
                                            return "伺服器連線異常";
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception e)
                        {
                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "error",
                                msg = e.Message
                            });
                            #endregion

                            logger.Error(e.Message);
                        }
                        #endregion

                        #endregion

                        #region //拋轉ERP
                        #region //INVTF
                        sql = @"INSERT INTO INVTF (COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE
                                , TF001, TF002, TF003, TF004, TF005, TF006, TF007, TF008, TF009, TF010
                                , TF011, TF012, TF013, TF014, TF015, TF016, TF017, TF018, TF019, TF020
                                , TF021, TF022, TF023, TF024, TF025, TF026, TF027, TF028, TF029, TF030
                                , TF031, TF032, TF033, TF034, TF035, TF036, TF037, TF038, TF039, TF040
                                , TF041, TF042, TF043, TF044, TF045, TF046
                                , FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID
                                , UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10)
                                VALUES (@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE
                                , @TF001, @TF002, @TF003, @TF004, @TF005, @TF006, @TF007, @TF008, @TF009, @TF010
                                , @TF011, @TF012, @TF013, @TF014, @TF015, @TF016, @TF017, @TF018, @TF019, @TF020
                                , @TF021, @TF022, @TF023, @TF024, @TF025, @TF026, @TF027, @TF028, @TF029, @TF030
                                , @TF031, @TF032, @TF033, @TF034, @TF035, @TF036, @TF037, @TF038, @TF039, @TF040
                                , @TF041, @TF042, @TF043, @TF044, @TF045, @TF046
                                , @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID, @MODI_TIME, @MODI_AP, @MODI_PRID
                                , @UDF01, @UDF02, @UDF03, @UDF04, @UDF05, @UDF06, @UDF07, @UDF08, @UDF09, @UDF10)";
                        rowsAffected += sqlConnection.Execute(sql, iNVTFs);
                        #endregion

                        #region //INVTG
                        sql = @"INSERT INTO INVTG (COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE
                                , TG001, TG002, TG003, TG004, TG005, TG006, TG007, TG008, TG009, TG010
                                , TG011, TG012, TG013, TG014, TG015, TG016, TG017, TG018, TG019, TG020
                                , TG021, TG022, TG023, TG024, TG025, TG026, TG027, TG028, TG029, TG030
                                , TG031, TG032, TG033, TG034, TG035, TG036, TG037, TG038, TG039, TG040
                                , TG041, TG042, TG043, TG044, TG045, TG046, TG047, TG048, TG049, TG050
                                , TG051, TG052, TG053, TG054, TG055
                                , TG500, TG501, TG502, TG503
                                , FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID
                                , UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10)
                                VALUES (@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE
                                , @TG001, @TG002, @TG003, @TG004, @TG005, @TG006, @TG007, @TG008, @TG009, @TG010
                                , @TG011, @TG012, @TG013, @TG014, @TG015, @TG016, @TG017, @TG018, @TG019, @TG020
                                , @TG021, @TG022, @TG023, @TG024, @TG025, @TG026, @TG027, @TG028, @TG029, @TG030
                                , @TG031, @TG032, @TG033, @TG034, @TG035, @TG036, @TG037, @TG038, @TG039, @TG040
                                , @TG041, @TG042, @TG043, @TG044, @TG045, @TG046, @TG047, @TG048, @TG049, @TG050
                                , @TG051, @TG052, @TG053, @TG054, @TG055
                                , @TG500, @TG501, @TG502, @TG503
                                , @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID, @MODI_TIME, @MODI_AP, @MODI_PRID
                                , @UDF01, @UDF02, @UDF03, @UDF04, @UDF05, @UDF06, @UDF07, @UDF08, @UDF09, @UDF10)";
                        rowsAffected += sqlConnection.Execute(sql, iNVTGs);
                        #endregion
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //暫出單單頭
                        foreach (var item in iNVTFs)
                        {
                            tempShippingNotes.Add(
                                new TempShippingNote
                                {
                                    TsnErpPrefix = item.TF001,
                                    TsnErpNo = item.TF002,
                                    TsnDate = DateTime.ParseExact(item.TF003, "yyyyMMdd", CultureInfo.InvariantCulture),
                                    DocDate = DateTime.ParseExact(item.TF024, "yyyyMMdd", CultureInfo.InvariantCulture),
                                    ToObject = item.TF004,
                                    ObjectOther = item.TF005,
                                    DepartmentNo = item.TF007,
                                    UserNo = item.TF008,
                                    Remark = item.TF014,
                                    OtherRemark = item.TF018,
                                    TotalQty = item.TF022,
                                    Amount = item.TF023,
                                    TaxAmount = item.TF027,
                                    ConfirmStatus = item.TF020,
                                    ConfirmUserNo = item.TF025,
                                    CompanyId = CurrentCompany,
                                    TransferStatus = "Y",
                                    CreateDate = CreateDate,
                                    LastModifiedDate = LastModifiedDate,
                                    CreateBy = CreateBy,
                                    LastModifiedBy = LastModifiedBy,
                                });
                        }
                        #endregion

                        #region //暫出單單身
                        foreach (var item in iNVTGs)
                        {
                            tsnDetails.Add(
                                new TsnDetail
                                {
                                    TsnErpPrefix = item.TG001,
                                    TsnErpNo = item.TG002,
                                    TsnSequence = item.TG003,
                                    MtlItemNo = item.TG004,
                                    TsnMtlItemName = item.TG005,
                                    TsnMtlItemSpec = item.TG006,
                                    TsnOutInventoryNo = item.TG007,
                                    TsnInInventoryNo = item.TG008,
                                    TsnQty = item.TG009,
                                    ProductType = item.TG043,
                                    FreebieOrSpareQty = item.TG044,
                                    UnitPrice = item.TG012,
                                    Amount = item.TG013,
                                    SoErpPrefix = item.TG014,
                                    SoErpNo = item.TG015,
                                    SoSequence = item.TG016,
                                    LotNumber = item.TG017,
                                    TsnRemark = item.TG019,
                                    ConfirmStatus = item.TG022,
                                    ClosureStatus = item.TG024,
                                    SaleQty = item.TG020,
                                    ReturnQty = item.TG021,
                                    CreateDate = CreateDate,
                                    LastModifiedDate = LastModifiedDate,
                                    CreateBy = CreateBy,
                                    LastModifiedBy = LastModifiedBy
                                });
                        }
                        #endregion

                        #region //撈取部門ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT DepartmentId, DepartmentNo
                                FROM BAS.Department
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        List<Department> resultDepartments = sqlConnection.Query<Department>(sql, dynamicParameters).ToList();

                        tempShippingNotes = tempShippingNotes.Join(resultDepartments, x => x.DepartmentNo, y => y.DepartmentNo, (x, y) => { x.DepartmentId = y.DepartmentId; return x; }).ToList();
                        #endregion

                        #region //撈取人員ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserId, a.UserNo
                                FROM BAS.[User] a
                                LEFT JOIN BAS.Department b ON b.DepartmentId = a.DepartmentId
                                WHERE b.CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        List<User> resultUsers = sqlConnection.Query<User>(sql, dynamicParameters).ToList();

                        tempShippingNotes = tempShippingNotes.Join(resultUsers, x => x.UserNo, y => y.UserNo, (x, y) => { x.UserId = y.UserId; return x; }).ToList();
                        tempShippingNotes = tempShippingNotes.GroupJoin(resultUsers, x => x.ConfirmUserNo, y => y.UserNo, (x, y) => { x.ConfirmUserId = y.FirstOrDefault()?.UserId; return x; }).ToList();
                        List<TempShippingNote> toObjectUser = tempShippingNotes.Where(x => x.ToObject == "3").Join(resultUsers, x => x.ObjectOther, y => y.UserNo, (x, y) => { x.ObjectUser = y.UserId; return x; }).Select(x => x).ToList();
                        #endregion

                        #region //撈取異動對象ID
                        #region //撈取客戶ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT CustomerId, CustomerNo 
                                FROM SCM.Customer
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        List<Customer> resultCustomers = sqlConnection.Query<Customer>(sql, dynamicParameters).ToList();

                        List<TempShippingNote> toObjectCustomer = tempShippingNotes.Where(x => x.ToObject == "1").Join(resultCustomers, x => x.ObjectOther, y => y.CustomerNo, (x, y) => { x.ObjectCustomer = y.CustomerId; return x; }).Select(x => x).ToList();
                        #endregion

                        #region //撈取供應商ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT SupplierId, SupplierNo 
                                FROM SCM.Supplier
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        List<Supplier> resultSuppliers = sqlConnection.Query<Supplier>(sql, dynamicParameters).ToList();

                        List<TempShippingNote> toObjectSupplier = tempShippingNotes.Where(x => x.ToObject == "2").Join(resultSuppliers, x => x.ObjectOther, y => y.SupplierNo, (x, y) => { x.ObjectSupplier = y.SupplierId; return x; }).Select(x => x).ToList();
                        #endregion

                        #region //其他
                        List<TempShippingNote> toObjectOther = tempShippingNotes.Where(x => x.ToObject == "9").Select(x => x).ToList();
                        #endregion

                        tempShippingNotes = toObjectCustomer.Union(toObjectSupplier).Union(toObjectUser).Union(toObjectOther).ToList();
                        #endregion

                        #region //撈取品號ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT MtlItemId, MtlItemNo
                                FROM PDM.MtlItem
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        List<MtlItem> resultMtlitems = sqlConnection.Query<MtlItem>(sql, dynamicParameters).ToList();

                        tsnDetails = tsnDetails.Join(resultMtlitems, x => x.MtlItemNo, y => y.MtlItemNo, (x, y) => { x.MtlItemId = y.MtlItemId; return x; }).Select(x => x).ToList();
                        #endregion

                        #region //撈取庫別ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT InventoryId, InventoryNo
                                FROM SCM.Inventory
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        List<Inventory> resultInventories = sqlConnection.Query<Inventory>(sql, dynamicParameters).ToList();

                        tsnDetails = tsnDetails.Join(resultInventories, x => x.TsnOutInventoryNo, y => y.InventoryNo, (x, y) => { x.TsnOutInventory = y.InventoryId; return x; }).ToList();
                        tsnDetails = tsnDetails.Join(resultInventories, x => x.TsnInInventoryNo, y => y.InventoryNo, (x, y) => { x.TsnInInventory = y.InventoryId; return x; }).ToList();
                        #endregion

                        #region //撈取訂單單身ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.SoId, a.SoErpPrefix, a.SoErpNo, b.SoDetailId, b.SoSequence
                                FROM SCM.SaleOrder a
                                LEFT JOIN SCM.SoDetail b ON b.SoId = a.SoId
                                WHERE a.CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        List<SoDetail> resultSoDetails = sqlConnection.Query<SoDetail>(sql, dynamicParameters).ToList();

                        tsnDetails = tsnDetails.GroupJoin(resultSoDetails, x => new { x.SoErpPrefix, x.SoErpNo, x.SoSequence }, y => new { y.SoErpPrefix, y.SoErpNo, y.SoSequence }, (x, y) => { x.SoDetailId = y.FirstOrDefault()?.SoDetailId; return x; }).Select(x => x).ToList();
                        #endregion

                        #region //暫出單單頭BM新增
                        sql = @"INSERT INTO SCM.TempShippingNote (CompanyId, TsnErpPrefix, TsnErpNo, TsnDate
                                , DocDate, ToObject, ObjectCustomer, ObjectSupplier, ObjectUser, ObjectOther
                                , DepartmentId, UserId, Remark, OtherRemark, TotalQty, Amount, TaxAmount
                                , ConfirmStatus, ConfirmUserId, TransferStatus, TransferDate
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                VALUES (@CompanyId, @TsnErpPrefix, @TsnErpNo, @TsnDate
                                , @DocDate, @ToObject, @ObjectCustomer, @ObjectSupplier, @ObjectUser, @ObjectOther
                                , @DepartmentId, @UserId, @Remark, @OtherRemark, @TotalQty, @Amount, @TaxAmount
                                , @ConfirmStatus, @ConfirmUserId, @TransferStatus, @TransferDate
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        rowsAffected += sqlConnection.Execute(sql, tempShippingNotes);
                        #endregion

                        #region //撈取暫出單單頭ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TsnId, TsnErpPrefix, TsnErpNo
                                FROM  SCM.TempShippingNote
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var resultTempShippingNotes = sqlConnection.Query<TempShippingNote>(sql, dynamicParameters).ToList();

                        tsnDetails = tsnDetails.Join(resultTempShippingNotes, x => new { x.TsnErpPrefix, x.TsnErpNo }, y => new { y.TsnErpPrefix, y.TsnErpNo }, (x, y) => { x.TsnId = y.TsnId; return x; }).Select(x => x).ToList();
                        #endregion

                        #region //暫出單單身BM新增
                        sql = @"INSERT INTO SCM.TsnDetail (TsnId, TsnSequence, MtlItemId, TsnMtlItemName
                                , TsnMtlItemSpec, TsnOutInventory, TsnInInventory, TsnQty, ProductType, FreebieOrSpareQty
                                , UnitPrice, Amount, SoDetailId, LotNumber, TsnRemark, ConfirmStatus, ClosureStatus
                                , SaleQty, ReturnQty
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                VALUES (@TsnId, @TsnSequence, @MtlItemId, @TsnMtlItemName
                                , @TsnMtlItemSpec, @TsnOutInventory, @TsnInInventory, @TsnQty, @ProductType, @FreebieOrSpareQty
                                , @UnitPrice, @Amount, @SoDetailId, @LotNumber, @TsnRemark, @ConfirmStatus, @ClosureStatus
                                , @SaleQty, @ReturnQty
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        rowsAffected += sqlConnection.Execute(sql, tsnDetails);
                        #endregion
                    }

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "(" + rowsAffected + " rows affected)"
                    });
                    #endregion

                    //transactionScope.Complete();
                }
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //AddDeliveryToTsn1 -- 出貨轉暫出資料新增 -- Ben Ma 2023.05.10
        public string AddDeliveryToTsn1(string TsnErpPrefix, int SalesmenId, int DepartmentId, string DoDetails, string Remark)

        {
            try
            {
                if (TsnErpPrefix.Length <= 0) throw new SystemException("【單別】不能為空!");
                if (SalesmenId <= 0) throw new SystemException("【業務人員】不能為空!");
                if (DepartmentId <= 0) throw new SystemException("【所屬部門】不能為空!");


                JArray jsonArray = JArray.Parse(DoDetails);

                string dateNow = DateTime.Now.ToString("yyyyMMdd");
                string timeNow = DateTime.Now.ToString("HH:mm:ss");

                List<DeliveryToTsn> deliveryToTsns = new List<DeliveryToTsn>();

                //ERP資料寫入用
                List<INVTF> iNVTFs = new List<INVTF>();
                List<INVTG> iNVTGs = new List<INVTG>();

                //MES資料寫入用
                List<TempShippingNote> tempShippingNotes = new List<TempShippingNote>();
                List<TsnDetail> tsnDetails = new List<TsnDetail>();

                int rowsAffected = 0;
                string companyNo = "", departmentNo = "", userNo = "", userName = ""
                    , salesUserNo = "", salesUserName = ""
                    , erpPrefix = "", erpNo = "", docDate = "";
                string CurrentCompanyNo = "";

                int CustomerId = -1;

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int DoId = -1;
                        //#region //判斷出貨日要同一天
                        List<int> doDetail = new List<int>();

                        #region //公司別資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb,a.CompanyNo
                                FROM BAS.Company a
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var resultCompany = sqlConnection.Query(sql, dynamicParameters);
                        if (resultCompany.Count() <= 0) throw new SystemException("公司別資料錯誤!");

                        foreach (var item in resultCompany)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            companyNo = item.ErpNo;
                            CurrentCompanyNo = item.CompanyNo;
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
                        }
                        #endregion

                        #region //部門資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.DepartmentNo
                                FROM BAS.Department a
                                WHERE a.DepartmentId = @DepartmentId";
                        dynamicParameters.Add("DepartmentId", DepartmentId);

                        var resultDepartment = sqlConnection.Query(sql, dynamicParameters);
                        if (resultDepartment.Count() <= 0) throw new SystemException("部門資料錯誤!");

                        foreach (var item in resultDepartment)
                        {
                            departmentNo = item.DepartmentNo;
                        }
                        #endregion

                        #region //業務資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserNo, a.UserName, b.DepartmentNo
                                FROM BAS.[User] a
                                INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                                WHERE a.UserId = @userId";
                        dynamicParameters.Add("userId", SalesmenId);

                        var resultSales = sqlConnection.Query(sql, dynamicParameters);
                        if (resultSales.Count() <= 0) throw new SystemException("業務資料錯誤!");

                        foreach (var item in resultSales)
                        {
                            salesUserNo = item.UserNo;
                            salesUserName = item.UserName;
                        }
                        #endregion

                        #region //單據解析
                        int counter = 1;

                        foreach (var item in jsonArray)
                        {
                            foreach (var item2 in item.Children<JProperty>())//單據層
                            {

                                int doDetailId = Convert.ToInt32(item2.Name); //取得doDetailId
                                doDetail.Add(doDetailId);

                                #region //判斷出貨資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 b.DoId, CASE WHEN PickFreebieQty > 0 THEN '1' ELSE '2' END ItemType
                                        , ISNULL(c.PickQty, '0') PickQty, ISNULL(d.PickFreebieQty, '0') PickFreebieQty
                                        , ISNULL(e.PickSpareQty, '0') PickSpareQty, f.SoSequence, g.SoErpPrefix, g.SoErpNo
                                        , h.LotManagement, a.SoDetailId, h.MtlItemNo,g.CustomerId
                                        FROM SCM.DoDetail a
                                        INNER JOIN SCM.DeliveryOrder b ON a.DoId = b.DoId
                                        OUTER APPLY (
                                            SELECT x.ItemType, SUM(x.ItemQty) PickQty
                                            FROM SCM.PickingItem x
                                            WHERE x.SoDetailId = a.SoDetailId
                                            AND x.DoId = b.DoId
                                            AND x.ItemType = 1
                                            GROUP BY x.ItemType
                                        ) c
                                        OUTER APPLY (
                                            SELECT y.ItemType, SUM(y.ItemQty) PickFreebieQty
                                            FROM SCM.PickingItem y
                                            WHERE y.SoDetailId = a.SoDetailId
                                            AND y.DoId = b.DoId
                                            AND y.ItemType = 2
                                            GROUP BY y.ItemType
                                        ) d
                                        OUTER APPLY (
                                            SELECT z.ItemType, SUM(z.ItemQty) PickSpareQty
                                            FROM SCM.PickingItem z
                                            WHERE z.SoDetailId = a.SoDetailId
                                            AND z.DoId = b.DoId
                                            AND z.ItemType = 3
                                            GROUP BY z.ItemType
                                        ) e
                                        INNER JOIN SCM.SoDetail f ON a.SoDetailId = f.SoDetailId
                                        INNER JOIN SCM.SaleOrder g ON f.SoId = g.SoId
                                        INNER JOIN PDM.MtlItem h ON f.MtlItemId = h.MtlItemId
                                        WHERE a.DoDetailId = @DoDetailId
                                        AND b.[Status] = @Status
                                        AND b.CompanyId = @CompanyId";
                                dynamicParameters.Add("DoDetailId", doDetailId);
                                dynamicParameters.Add("Status", "S");
                                dynamicParameters.Add("CompanyId", CurrentCompany);

                                var resultDoDetail = sqlConnection.Query(sql, dynamicParameters);
                                if (resultDoDetail.Count() <= 0) throw new SystemException("出貨資料錯誤");

                                string productType = "", lotManagement = "", mtlItemNo = "", SoErpPrefix = "", SoErpNo = "", SoSequence = "";
                                int pickRegularQty = -1, pickFreebieQty = -1, pickSpareQty = -1, soDetailId = -1;
                                foreach (var item3 in resultDoDetail)
                                {
                                    SoErpPrefix = item3.SoErpPrefix;
                                    SoErpNo = item3.SoErpNo;
                                    SoSequence = item3.SoSequence;
                                    productType = item3.ItemType;
                                    pickRegularQty = item3.PickQty;
                                    pickFreebieQty = item3.PickFreebieQty;
                                    pickSpareQty = item3.PickSpareQty;
                                    lotManagement = item3.LotManagement;
                                    soDetailId = item3.SoDetailId;
                                    mtlItemNo = item3.MtlItemNo;
                                    DoId = item3.DoId;
                                    if (CustomerId == -1)
                                    {
                                        CustomerId = item3.CustomerId;
                                    }
                                    else
                                    {
                                        if (CustomerId != item3.CustomerId) throw new SystemException("出貨客戶不同");
                                    }
                                }
                                #endregion

                                var lotList = item2.Value.Children<JProperty>(); //批號層
                                foreach (var item3 in lotList)
                                {
                                    string innerKey = item3.Name; // 取得批號

                                    var details = item3.Value; //批號物件
                                    string LotNumber = details["LotNumber"].ToString();
                                    int ItemQty1 = Convert.ToInt32(details["ItemQty1"].ToString());
                                    int ItemQty2 = Convert.ToInt32(details["ItemQty2"].ToString());
                                    int ItemQty3 = Convert.ToInt32(details["ItemQty3"].ToString());
                                    int OutInventory = Convert.ToInt32(details["OutInventory"]);
                                    int InInventory = Convert.ToInt32(details["InInventory"]);

                                    #region //判斷轉出轉入庫別是否正確
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 a.InventoryNo
                                            FROM SCM.Inventory a
                                            WHERE a.InventoryId = @InventoryId
                                            AND a.CompanyId = @CompanyId";
                                    dynamicParameters.Add("InventoryId", OutInventory);
                                    dynamicParameters.Add("CompanyId", CurrentCompany);

                                    var resultOutInventory = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultOutInventory.Count() <= 0) throw new SystemException("轉出庫資料錯誤");

                                    string tsnOutInventoryNo = "";
                                    foreach (var item4 in resultOutInventory)
                                    {
                                        tsnOutInventoryNo = item4.InventoryNo;
                                    }

                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 a.InventoryNo
                                            FROM SCM.Inventory a
                                            WHERE a.InventoryId = @InventoryId
                                            AND a.CompanyId = @CompanyId";
                                    dynamicParameters.Add("InventoryId", InInventory);
                                    dynamicParameters.Add("CompanyId", CurrentCompany);

                                    var resultInInventory = sqlConnection.Query(sql, dynamicParameters);
                                    if (resultInInventory.Count() <= 0) throw new SystemException("轉入庫資料錯誤");

                                    string tsnInInventoryNo = "";
                                    foreach (var item4 in resultInInventory)
                                    {
                                        tsnInInventoryNo = item4.InventoryNo;
                                    }
                                    #endregion

                                    iNVTGs.Add(
                                    new INVTG
                                    {
                                        COMPANY = companyNo,
                                        CREATOR = userNo,
                                        USR_GROUP = "",
                                        CREATE_DATE = dateNow,
                                        MODIFIER = "",
                                        MODI_DATE = "",
                                        FLAG = "1",
                                        CREATE_TIME = timeNow,
                                        CREATE_AP = userNo + "PC",
                                        CREATE_PRID = "BM",
                                        MODI_TIME = "",
                                        MODI_AP = "",
                                        MODI_PRID = "",
                                        TG003 = counter.ToString("D4"),
                                        TG004 = mtlItemNo,
                                        TG007 = tsnOutInventoryNo,
                                        TG008 = tsnInInventoryNo,
                                        TG009 = ItemQty1,
                                        TG014 = SoErpPrefix,
                                        TG015 = SoErpNo,
                                        TG016 = SoSequence,
                                        TG017 = LotNumber,
                                        TG043 = productType,
                                        TG044 = ItemQty2 != 0 ? ItemQty2 : ItemQty3,
                                        TG052 = ItemQty1
                                    });
                                    counter++;

                                    deliveryToTsns.Add(
                                    new DeliveryToTsn
                                    {
                                        DoDetailId = doDetailId,
                                        SoErpPrefix = SoErpPrefix,
                                        SoErpNo = SoErpNo,
                                        SoSequence = SoSequence,
                                        TsnOutInventory = OutInventory,
                                        TsnOutInventoryNo = tsnOutInventoryNo,
                                        TsnInInventory = InInventory,
                                        TsnInInventoryNo = tsnInInventoryNo,
                                        PickQty = ItemQty1,
                                        ProductType = productType,
                                        FreebieOrSpareQty = ItemQty2 != 0 ? ItemQty2 : ItemQty3,
                                        LotManagement = lotManagement != "N" ? lotManagement : "",
                                        SoDetailId = soDetailId,
                                        MtlItemNo = mtlItemNo
                                    });
                                }
                            }
                        }
                        #endregion

                        #region //判斷出貨日要同一天
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT DISTINCT b.DoDate
                                FROM SCM.DoDetail a
                                INNER JOIN SCM.DeliveryOrder b ON a.DoId = b.DoId
                                WHERE a.DoDetailId IN @DoDetail
                                AND b.[Status] = @Status
                                AND b.CompanyId = @CompanyId";
                        dynamicParameters.Add("DoDetail", doDetail.ToArray());
                        dynamicParameters.Add("Status", "S");
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var resultDateVerify = sqlConnection.Query(sql, dynamicParameters);
                        if (resultDateVerify.Count() <= 0) throw new SystemException("出貨資料錯誤!");

                        if (resultDateVerify.Count() > 1) throw new SystemException("出貨日期不同!");
                        #endregion

                        #region //出貨資料-編寫ERP暫出單

                        #region //INVTF
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT FORMAT(a.DocDate, 'yyyyMMdd') TF003, '1' TF004, b.CustomerNo TF005
                                , FORMAT(a.DocDate, 'yyyyMMdd') TF024, FORMAT(a.DocDate, 'yyyyMMdd') TF041
                                FROM SCM.DeliveryOrder a
                                INNER JOIN SCM.Customer b ON a.CustomerId = b.CustomerId
                                WHERE a.DoId = @DoId
                                AND a.[Status] = @Status
                                AND a.CompanyId = @CompanyId";
                        dynamicParameters.Add("DoId", DoId);
                        dynamicParameters.Add("Status", "S");
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        iNVTFs = sqlConnection.Query<INVTF>(sql, dynamicParameters).ToList();
                        #endregion

                        #endregion

                        #region //複製未出完定板
                        if (CurrentCompany == 4)
                        {
                            #region //轉暫出定版單頭
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT DISTINCT a.DoId
                                    FROM SCM.DoDetail a
                                    WHERE a.DoDetailId IN @DoDetail";
                            dynamicParameters.Add("DoDetail", doDetail.ToArray());

                            var resultDoList = sqlConnection.Query(sql, dynamicParameters);
                            #endregion

                            foreach (var Key in resultDoList)
                            {
                                #region //該出貨單所有單身
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.DoDetailId, a.SoDetailId, a.DoQty, a.FreebieQty, a.SpareQty
                                        , ISNULL(b.PickQty,0) PickQty, ISNULL(c.PickFreebieQty,0) PickFreebieQty, ISNULL(d.PickSpareQty,0) PickSpareQty
                                        , e.DoDate
                                        FROM SCM.DoDetail a
                                        OUTER APPLY (
                                            SELECT SUM(ba.ItemQty) PickQty
                                            FROM SCM.PickingItem ba
                                            WHERE ba.SoDetailId = a.SoDetailId
                                            AND ba.ItemType = 1
                                            AND ba.DoId = a.DoId
                                        ) b
                                        OUTER APPLY (
                                            SELECT SUM(ca.ItemQty) PickFreebieQty
                                            FROM SCM.PickingItem ca
                                            WHERE ca.SoDetailId = a.SoDetailId
                                            AND ca.ItemType = 2
                                            AND ca.DoId = a.DoId
                                        ) c
                                        OUTER APPLY (
                                            SELECT SUM(da.ItemQty) PickSpareQty
                                            FROM SCM.PickingItem da
                                            WHERE da.SoDetailId = a.SoDetailId
                                            AND da.ItemType = 3
                                            AND da.DoId = a.DoId
                                        ) d
                                        INNER JOIN SCM.DeliveryOrder e ON a.DoId = e.DoId
                                        WHERE a.DoId = @DoId";
                                dynamicParameters.Add("DoId", Key.DoId);

                                var resultDoDetail = sqlConnection.Query(sql, dynamicParameters);
                                #endregion

                                foreach (var DoDetail in resultDoDetail)
                                {
                                    string modifier = userNo, soErpPrefix = "", soErpNo = "", soSequence = "";

                                    double UnQty = DoDetail.DoQty - DoDetail.PickQty;
                                    double UnFreebieQty = DoDetail.FreebieQty - DoDetail.PickFreebieQty;
                                    double UnSpareQty = DoDetail.SpareQty - DoDetail.PickSpareQty;
                                    DateTime DoDate = Convert.ToDateTime(DoDetail.DoDate).AddDays(1);

                                    if (UnQty > 0 || UnFreebieQty > 0 || UnSpareQty > 0)
                                    {
                                        #region //交期歷史紀錄新增
                                        #region //撈取原排定交貨日
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT PcPromiseDate
                                                FROM SCM.SoDetail
                                                WHERE SoDetailId = @SoDetailId";
                                        dynamicParameters.Add("SoDetailId", DoDetail.SoDetailId);

                                        var resultPcPromiseDate = sqlConnection.Query(sql, dynamicParameters);

                                        DateTime pcPromiseDate = new DateTime();
                                        foreach (var item in resultPcPromiseDate)
                                        {
                                            pcPromiseDate = item.PcPromiseDate;
                                        }
                                        #endregion

                                        #region //撈取原修改紀錄ID
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT ISNULL(MAX(DeliveryDateLogId), -1) AS ParentLogId
                                                FROM SCM.DeliveryDateLog
                                                WHERE SoDetailId = @SoDetailId";
                                        dynamicParameters.Add("SoDetailId", DoDetail.SoDetailId);

                                        var resultLogId = sqlConnection.Query(sql, dynamicParameters);

                                        int ParentLogId = -1;
                                        foreach (var item in resultLogId)
                                        {
                                            ParentLogId = item.ParentLogId;
                                        }
                                        #endregion

                                        #region //新增交期修改紀錄
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO SCM.DeliveryDateLog (ParentLogId, SoDetailId, PcPromiseDate, DepartmentId
                                                , SupervisorId, CauseType, CauseDescription
                                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                                OUTPUT INSERTED.DeliveryDateLogId
                                                VALUES (@ParentLogId, @SoDetailId, @PcPromiseDate, @DepartmentId
                                                , @SupervisorId, @CauseType, @CauseDescription
                                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                ParentLogId = ParentLogId > 0 ? (int?)ParentLogId : null,
                                                DoDetail.SoDetailId,
                                                PcPromiseDate = pcPromiseDate,
                                                CauseType = (int?)null,
                                                DepartmentId = (int?)null,
                                                SupervisorId = (int?)null,
                                                CauseDescription = "分批出貨",
                                                CreateDate,
                                                LastModifiedDate,
                                                CreateBy,
                                                LastModifiedBy
                                            });
                                        var logResult = sqlConnection.Query(sql, dynamicParameters);

                                        rowsAffected += logResult.Count();
                                        #endregion

                                        #region //修改BM訂單排定交貨日
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"UPDATE SCM.SoDetail SET
                                                PcPromiseDate = @PcPromiseDate,
                                                LastModifiedDate = @LastModifiedDate,
                                                LastModifiedBy = @LastModifiedBy
                                                WHERE SoDetailId = @SoDetailId";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                PcPromiseDate = DoDate,
                                                LastModifiedDate,
                                                LastModifiedBy,
                                                DoDetail.SoDetailId
                                            });

                                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                        #endregion

                                        #region //撈取單別/單號/流水號
                                        sql = @"SELECT b.SoErpPrefix, b.SoErpNo, a.SoSequence 
                                                FROM SCM.SoDetail a
                                                INNER JOIN SCM.SaleOrder b ON b.SoId = a.SoId
                                                WHERE SoDetailId = @SoDetailId";
                                        dynamicParameters.Add("SoDetailId", DoDetail.SoDetailId);

                                        var resultSoDetail = sqlConnection.Query(sql, dynamicParameters);

                                        foreach (var item in resultSoDetail)
                                        {
                                            soErpPrefix = item.SoErpPrefix;
                                            soErpNo = item.SoErpNo;
                                            soSequence = item.SoSequence;
                                        }
                                        #endregion

                                        using (SqlConnection erpSqlConnection = new SqlConnection(ErpConnectionStrings))
                                        {
                                            #region //修改ERP訂單排定交貨日
                                            sql = @"UPDATE COPTD SET
                                                    TD048 = @TD048,
                                                    MODI_DATE = @MODI_DATE,
                                                    MODI_TIME = @MODI_TIME,
                                                    MODIFIER = @MODIFIER
                                                    WHERE TD001 = @TD001
                                                    AND TD002 = @TD002
                                                    AND TD003 = @TD003";
                                            rowsAffected += erpSqlConnection.Execute(sql,
                                                new
                                                {
                                                    TD048 = DoDate.ToString("yyyyMMdd"),
                                                    MODI_DATE = dateNow,
                                                    MODI_TIME = timeNow,
                                                    MODIFIER = modifier,
                                                    TD001 = soErpPrefix,
                                                    TD002 = soErpNo,
                                                    TD003 = soSequence
                                                });
                                            #endregion
                                        }
                                        #endregion

                                        #region //判斷新時段是否有同出貨客戶的出貨單
                                        sql = @"SELECT TOP 1 a.DoId
                                                FROM SCM.DeliveryOrder a
                                                WHERE a.DoDate = @DoDate
                                                AND a.DcId = (
                                                    SELECT y.DcId
                                                    FROM SCM.DoDetail z
                                                    INNER JOIN SCM.DeliveryOrder y ON z.DoId = y.DoId
                                                    WHERE z.DoDetailId = @DoDetailId
                                                )";
                                        dynamicParameters.Add("DoDate", DoDate);
                                        dynamicParameters.Add("DoDetailId", DoDetail.DoDetailId);

                                        var resultExistDeliveryOrder = sqlConnection.Query(sql, dynamicParameters);

                                        int existDoId = -1;
                                        foreach (var item in resultExistDeliveryOrder)
                                        {
                                            existDoId = item.DoId;
                                        }
                                        #endregion

                                        if (existDoId > 0)
                                        {
                                            #region //查詢出貨單流水號
                                            sql = @"SELECT REPLACE(STR(ISNULL(MAX(DoSequence),0) + 1, 4),' ','0') DoSequence 
                                                    FROM SCM.DoDetail
                                                    WHERE DoId = @DoId";
                                            dynamicParameters.Add("DoId", existDoId);

                                            var resultDoSequence = sqlConnection.Query(sql, dynamicParameters);

                                            string DoSequence = "";
                                            foreach (var item in resultDoSequence)
                                            {
                                                DoSequence = item.DoSequence;
                                            }
                                            #endregion

                                            #region //複製出貨單單身
                                            sql = @"INSERT INTO SCM.DoDetail
                                                    SELECT @DoId, a.SoDetailId, @DoSequence, a.TransInInventoryId, @DoQty, @FreebieQty, @SpareQty
                                                    , a.UnitPrice, a.Amount, a.DoDetailRemark, a.WareHouseDoDetailRemark
                                                    , a.PcDoDetailRemark, a.DeliveryProcess, a.DeliveryRoutine, a.DeliveryDocument
                                                    , a.OrderSituation, a.DeliveryMethod, a.ConfirmStatus, a.ClosureStatus
                                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy
                                                    FROM SCM.DoDetail a
                                                    WHERE a.DoDetailId = @DoDetailId";
                                            rowsAffected += sqlConnection.Execute(sql,
                                                new
                                                {
                                                    DoId = existDoId,
                                                    DoSequence,
                                                    TransInInventoryId = -1,
                                                    DoQty = UnQty,
                                                    FreebieQty = UnFreebieQty,
                                                    SpareQty = UnSpareQty,
                                                    CreateDate,
                                                    LastModifiedDate,
                                                    CreateBy,
                                                    LastModifiedBy,
                                                    DoDetail.DoDetailId
                                                });
                                            #endregion
                                        }
                                        else
                                        {
                                            #region //複製出貨單單頭
                                            sql = @"INSERT INTO SCM.DeliveryOrder
                                                    OUTPUT INSERTED.DoId
                                                    SELECT a.CompanyId, a.DepartmentId, a.UserId, a.DoErpPrefix, @DoErpNo
                                                    , @DoDate, @DocDate, a.CustomerId, a.DcId, a.WayBill, a.Traffic, a.ShipMethod
                                                    , a.DoAddressFirst, a.DoAddressSecond, a.TotalQty, a.Amount, a.TaxAmount
                                                    , a.DoRemark, a.WareHouseDoRemark, a.MeasureMailStatus, a.ConfirmStatus, a.ConfirmUserId
                                                    , a.TransferStatus, a.TransferDate, a.Status
                                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy
                                                    FROM SCM.DeliveryOrder a
                                                    INNER JOIN SCM.DoDetail b ON b.DoId = a.DoId
                                                    WHERE DoDetailId = @DoDetailId";
                                            var insertResult = sqlConnection.Query(sql,
                                                new
                                                {
                                                    DoErpNo = BaseHelper.RandomCode(11),
                                                    DoDate,
                                                    DocDate = CreateDate,
                                                    CreateDate,
                                                    LastModifiedDate,
                                                    CreateBy,
                                                    LastModifiedBy,
                                                    DoDetail.DoDetailId
                                                });

                                            rowsAffected += insertResult.Count();

                                            int newDoId = insertResult.Select(x => x.DoId).FirstOrDefault();
                                            #endregion

                                            #region //複製出貨單單身
                                            sql = @"INSERT INTO SCM.DoDetail
                                                    SELECT @DoId, a.SoDetailId, @DoSequence, a.TransInInventoryId, @DoQty, @FreebieQty, @SpareQty
                                                    , a.UnitPrice, a.Amount, a.DoDetailRemark, a.WareHouseDoDetailRemark
                                                    , a.PcDoDetailRemark, a.DeliveryProcess, a.DeliveryRoutine, a.DeliveryDocument
                                                    , a.OrderSituation, a.DeliveryMethod, a.ConfirmStatus, a.ClosureStatus
                                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy
                                                    FROM SCM.DoDetail a
                                                    WHERE a.DoDetailId = @DoDetailId";
                                            rowsAffected += sqlConnection.Execute(sql,
                                                new
                                                {
                                                    DoId = newDoId,
                                                    DoSequence = "0001",
                                                    TransInInventoryId = -1,
                                                    DoQty = UnQty,
                                                    FreebieQty = UnFreebieQty,
                                                    SpareQty = UnSpareQty,
                                                    CreateDate,
                                                    LastModifiedDate,
                                                    CreateBy,
                                                    LastModifiedBy,
                                                    DoDetail.DoDetailId
                                                });
                                            #endregion
                                        }
                                    }
                                }
                            }
                        }
                        #endregion
                    }

                    #region//暫出單 數量檢核 

                    //檢核條件
                    // 1: 訂單正常數<=暫出單正常數+撿貨數-暫出單歸還數
                    // 2: 訂單贈品數<=暫出單贈品數+撿貨贈品數-暫出歸還贈品數
                    // 3: 訂單備品數<=暫出單備品數+撿貨備品數-暫出單備品數
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        foreach (var item in jsonArray)
                        {
                            foreach (var item2 in item.Children<JProperty>())//單據層
                            {
                                int doDetailId = Convert.ToInt32(item2.Name); //取得doDetailId

                                #region//正常數卡控
                                int NormalOrderQty = 0; //訂單數
                                int NormalTemporaryOrderQty = 0; //暫出單正常數     
                                int NormalTemporaryReturnOrderQty = 0; //暫出單歸還數
                                int PickQty = 0; //撿貨數

                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT c.SoErpPrefix,c.SoErpNo,b.SoSequence
                                        FROM SCM.DoDetail a
                                        INNER JOIN SCM.SoDetail b ON a.SoDetailId=b.SoDetailId
                                        INNER JOIN SCM.SaleOrder  c ON b.SoId=c.SoId  
                                        WHERE a.DoDetailId=@DoDetailId";
                                dynamicParameters.Add("DoDetailId", doDetailId);
                                var resultSoDetail = sqlConnection.Query(sql, dynamicParameters);
                                if (resultSoDetail.Count() > 0)
                                {
                                    foreach (var item3 in resultSoDetail)
                                    {
                                        string SoErpPrefix = item3.SoErpPrefix;//訂單單別
                                        string SoErpNo = item3.SoErpNo;//訂單單號
                                        string SoSequence = item3.SoSequence;//訂單流水號

                                        #region//取出ERP暫出單資訊                                  
                                        using (SqlConnection sqlConnectionErp = new SqlConnection(ErpConnectionStrings))
                                        {
                                            //1.檢查訂單單身數量
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"SELECT b.TD001,b.TD002,b.TD003,b.TD008,b.TD009
                                                FROM COPTC a
                                                INNER JOIN COPTD b ON a.TC001=b.TD001 AND a.TC002=b.TD002
                                                WHERE a.TC027 !='V'
                                                AND b.TD001=@SoErpPrefix
                                                AND b.TD002=@SoErpNo
                                                AND b.TD003=@SoSequence";
                                            dynamicParameters.Add("SoErpPrefix", SoErpPrefix);
                                            dynamicParameters.Add("SoErpNo", SoErpNo);
                                            dynamicParameters.Add("SoSequence", SoSequence);
                                            var resultCOPTC = sqlConnectionErp.Query(sql, dynamicParameters);
                                            if (resultCOPTC.Count() > 0)
                                            {
                                                foreach (var itemCOPTC in resultCOPTC)
                                                {
                                                    NormalOrderQty = Convert.ToInt32(itemCOPTC.TD008);
                                                }
                                            }

                                            //2.檢查這張訂單有沒有開過暫出單
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"SELECT b.TG001,b.TG002,b.TG003,b.TG014,b.TG015,b.TG016,b.TG009
                                                FROM INVTF a
                                                INNER JOIN INVTG b ON a.TF001=b.TG001 AND a.TF002=b.TG002
                                                WHERE a.TF020 !='V'
                                                AND b.TG014=@SoErpPrefix
                                                AND b.TG015=@SoErpNo
                                                AND b.TG016=@SoSequence";
                                            dynamicParameters.Add("SoErpPrefix", SoErpPrefix);
                                            dynamicParameters.Add("SoErpNo", SoErpNo);
                                            dynamicParameters.Add("SoSequence", SoSequence);
                                            var resultINVTF = sqlConnectionErp.Query(sql, dynamicParameters);
                                            if (resultINVTF.Count() > 0)
                                            {
                                                foreach (var itemINVTF in resultINVTF)
                                                {
                                                    //3.檢查所開的暫出單有沒有開過【暫出歸還單】
                                                    dynamicParameters = new DynamicParameters();
                                                    sql = @"SELECT a.TI009
                                                            FROM INVTI a
                                                            INNER JOIN INVTG b ON a.TI014 = b.TG001 AND a.TI015=b.TG002 AND a.TI016=b.TG003
                                                            WHERE a.TI022 !='V'
                                                            AND b.TG001=@TsnErpPrefix
                                                            AND b.TG002=@TsnErpNo
                                                            AND b.TG003=@TsnSequence";
                                                    dynamicParameters.Add("TsnErpPrefix", itemINVTF.TG001);
                                                    dynamicParameters.Add("TsnErpNo", itemINVTF.TG002);
                                                    dynamicParameters.Add("TsnSequence", itemINVTF.TG003);
                                                    var resultINVTI = sqlConnectionErp.Query(sql, dynamicParameters);
                                                    if (resultINVTI.Count() > 0)
                                                    {
                                                        foreach (var itemINVTI in resultINVTI)
                                                        {
                                                            NormalTemporaryReturnOrderQty += Convert.ToInt32(itemINVTI.TI009); //暫出單歸還數
                                                        }
                                                    }
                                                    else
                                                    {
                                                        NormalTemporaryReturnOrderQty = 0;
                                                    }
                                                    NormalTemporaryOrderQty += Convert.ToInt32(itemINVTF.TG009);
                                                }
                                            }
                                            else
                                            {
                                                NormalTemporaryOrderQty = 0;
                                            }
                                        }
                                        #endregion

                                        #region//本次撿貨數量
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT  ISNULL(e.PickQty,0) PickQty
                                                    , ISNULL(f.PickFreebieQty,0) PickFreebieQty
                                                    ,ISNULL(g.PickSpareQty,0)PickSpareQty
                                                FROM SCM.DoDetail a
                                                    INNER JOIN SCM.DeliveryOrder b ON a.DoId = b.DoId
                                                    INNER JOIN SCM.SoDetail c ON a.SoDetailId = c.SoDetailId
                                                    INNER JOIN SCM.SaleOrder d ON c.SoId = d.SoId
                                                    OUTER APPLY (
                                                        SELECT SUM(ea.ItemQty) PickQty
                                                        FROM SCM.PickingItem ea
                                                        WHERE ea.SoDetailId = a.SoDetailId
                                                        AND ea.ItemType = 1
                                                        AND ea.DoId = a.DoId
                                                    ) e
                                                    OUTER APPLY (
                                                        SELECT SUM(fa.ItemQty) PickFreebieQty
                                                        FROM SCM.PickingItem fa
                                                        WHERE fa.SoDetailId = a.SoDetailId
                                                        AND fa.ItemType = 2
                                                        AND fa.DoId = a.DoId
                                                    ) f
                                                    OUTER APPLY (
                                                        SELECT SUM(ga.ItemQty) PickSpareQty
                                                        FROM SCM.PickingItem ga
                                                        WHERE ga.SoDetailId = a.SoDetailId
                                                        AND ga.ItemType = 3
                                                        AND ga.DoId = a.DoId
                                                    ) g
                                                    LEFT JOIN SCM.Inventory h ON a.TransInInventoryId = h.InventoryId
                                            WHERE a.DoDetailId = @DoDetailId
                                            AND b.[Status] = @Status
                                            AND d.CompanyId = @CompanyId
                                            ORDER BY d.SoErpPrefix, d.SoErpNo, c.SoSequence";
                                        dynamicParameters.Add("DoDetailId", doDetailId);
                                        dynamicParameters.Add("Status", "S");
                                        dynamicParameters.Add("CompanyId", CurrentCompany);
                                        var resultPickingItem = sqlConnection.Query(sql, dynamicParameters);
                                        if (resultPickingItem.Count() > 0)
                                        {
                                            foreach (var itemPickingItem in resultPickingItem)
                                            {
                                                PickQty = Convert.ToInt32(itemPickingItem.PickQty); //撿貨數
                                            }
                                        }
                                        else
                                        {
                                            PickQty = 0;
                                        }
                                        #endregion

                                        if (NormalOrderQty < NormalTemporaryOrderQty + PickQty - NormalTemporaryReturnOrderQty)
                                        {
                                            int Available = NormalOrderQty - NormalTemporaryOrderQty + NormalTemporaryReturnOrderQty;
                                            throw new SystemException("銷貨量超交，本次最大可銷貨量為:[" + Available + "]");
                                        }
                                    }

                                }
                                #endregion

                                #region//贈品數卡控
                                int NormalOrderFreebieQty = 0; //訂單贈品數量
                                int TemporaryOrderGiftQty = 0; //暫出單贈品數
                                int TemporaryReturnOrderGiftQty = 0; //暫出歸還贈品數
                                int PickFreebieQty = 0; //撿貨贈品數

                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT c.SoErpPrefix,c.SoErpNo,b.SoSequence
                                    FROM SCM.DoDetail a
                                    INNER JOIN SCM.SoDetail b ON a.SoDetailId=b.SoDetailId
                                    INNER JOIN SCM.SaleOrder  c ON b.SoId=c.SoId  
                                    WHERE a.DoDetailId=@DoDetailId";
                                dynamicParameters.Add("DoDetailId", doDetailId);
                                var resultFreebie = sqlConnection.Query(sql, dynamicParameters);
                                if (resultFreebie.Count() > 0)
                                {
                                    foreach (var item3 in resultFreebie)
                                    {
                                        string SoErpPrefix = item3.SoErpPrefix;//訂單單別
                                        string SoErpNo = item3.SoErpNo;//訂單單號
                                        string SoSequence = item3.SoSequence;//訂單流水號

                                        #region//取出ERP暫出單資訊
                                        using (SqlConnection sqlConnectionErp = new SqlConnection(ErpConnectionStrings))
                                        {
                                            //1.檢查訂單單身數量
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"SELECT b.TD024
                                                    FROM COPTC a
                                                    INNER JOIN COPTD b ON a.TC001=b.TD001 AND a.TC002=b.TD002
                                                    WHERE a.TC027 !='V'
                                                    AND b.TD001=@SoErpPrefix
                                                    AND b.TD002=@SoErpNo
                                                    AND b.TD003=@SoSequence";
                                            dynamicParameters.Add("SoErpPrefix", SoErpPrefix);
                                            dynamicParameters.Add("SoErpNo", SoErpNo);
                                            dynamicParameters.Add("SoSequence", SoSequence);
                                            var resultCOPTC = sqlConnectionErp.Query(sql, dynamicParameters);
                                            if (resultCOPTC.Count() > 0)
                                            {
                                                foreach (var itemCOPTC in resultCOPTC)
                                                {
                                                    NormalOrderFreebieQty = Convert.ToInt32(itemCOPTC.TD024);
                                                }
                                            }

                                            //2.檢查這張訂單沒有開過暫出單
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"SELECT b.TG001,b.TG002,b.TG003,b.TG014,b.TG015,b.TG016,b.TG044
                                                FROM INVTF a
                                                INNER JOIN INVTG b ON a.TF001=b.TG001 AND a.TF002=b.TG002
                                                WHERE a.TF020 !='V' AND b.TG043='1'
                                                AND b.TG014=@SoErpPrefix
                                                AND b.TG015=@SoErpNo
                                                AND b.TG016=@SoSequence";
                                            dynamicParameters.Add("SoErpPrefix", SoErpPrefix);
                                            dynamicParameters.Add("SoErpNo", SoErpNo);
                                            dynamicParameters.Add("SoSequence", SoSequence);
                                            var resultINVTF = sqlConnectionErp.Query(sql, dynamicParameters);
                                            if (resultINVTF.Count() > 0)
                                            {
                                                foreach (var itemINVTF in resultINVTF)
                                                {
                                                    //3.檢查所開的暫出單有沒有開過【暫出歸還單】
                                                    dynamicParameters = new DynamicParameters();
                                                    sql = @"SELECT a.TI035
                                                            FROM INVTI a
                                                            INNER JOIN INVTG b ON a.TI014 = b.TG001 AND a.TI015=b.TG002 AND a.TI016=b.TG003
                                                            WHERE a.TI022 !='V' AND a.TI034='1'
                                                            AND b.TG001=@TsnErpPrefix
                                                            AND b.TG002=@TsnErpNo
                                                            AND b.TG003=@TsnSequence";
                                                    dynamicParameters.Add("TsnErpPrefix", itemINVTF.TG001);
                                                    dynamicParameters.Add("TsnErpNo", itemINVTF.TG002);
                                                    dynamicParameters.Add("TsnSequence", itemINVTF.TG003);
                                                    var resultINVTI = sqlConnectionErp.Query(sql, dynamicParameters);
                                                    if (resultINVTI.Count() > 0)
                                                    {
                                                        foreach (var itemINVTI in resultINVTI)
                                                        {
                                                            TemporaryReturnOrderGiftQty += Convert.ToInt32(itemINVTI.TI035); //暫出單歸還數
                                                        }
                                                    }
                                                    else
                                                    {
                                                        TemporaryReturnOrderGiftQty = 0;
                                                    }
                                                    TemporaryOrderGiftQty += Convert.ToInt32(itemINVTF.TG044);
                                                }
                                            }
                                            else
                                            {
                                                TemporaryOrderGiftQty = 0;
                                            }
                                        }
                                        #endregion

                                        #region//本次撿貨數量
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT  ISNULL(e.PickQty,0) PickQty
                                                        , ISNULL(f.PickFreebieQty,0) PickFreebieQty
                                                        ,ISNULL(g.PickSpareQty,0)PickSpareQty
                                                    FROM SCM.DoDetail a
                                                        INNER JOIN SCM.DeliveryOrder b ON a.DoId = b.DoId
                                                        INNER JOIN SCM.SoDetail c ON a.SoDetailId = c.SoDetailId
                                                        INNER JOIN SCM.SaleOrder d ON c.SoId = d.SoId
                                                        OUTER APPLY (
                                                            SELECT SUM(ea.ItemQty) PickQty
                                                            FROM SCM.PickingItem ea
                                                            WHERE ea.SoDetailId = a.SoDetailId
                                                            AND ea.ItemType = 1
                                                            AND ea.DoId = a.DoId
                                                        ) e
                                                        OUTER APPLY (
                                                            SELECT SUM(fa.ItemQty) PickFreebieQty
                                                            FROM SCM.PickingItem fa
                                                            WHERE fa.SoDetailId = a.SoDetailId
                                                            AND fa.ItemType = 2
                                                            AND fa.DoId = a.DoId
                                                        ) f
                                                        OUTER APPLY (
                                                            SELECT SUM(ga.ItemQty) PickSpareQty
                                                            FROM SCM.PickingItem ga
                                                            WHERE ga.SoDetailId = a.SoDetailId
                                                            AND ga.ItemType = 3
                                                            AND ga.DoId = a.DoId
                                                        ) g
                                                        LEFT JOIN SCM.Inventory h ON a.TransInInventoryId = h.InventoryId
                                                WHERE a.DoDetailId = @DoDetailId                                       
                                                AND b.[Status] = @Status
                                                AND d.CompanyId = @CompanyId
                                                AND d.CompanyId = @CompanyId
                                                ORDER BY d.SoErpPrefix, d.SoErpNo, c.SoSequence";
                                        dynamicParameters.Add("DoDetailId", doDetailId);
                                        dynamicParameters.Add("Status", "S");
                                        dynamicParameters.Add("CompanyId", CurrentCompany);
                                        var resultPickingItem = sqlConnection.Query(sql, dynamicParameters);
                                        if (resultPickingItem.Count() > 0)
                                        {
                                            foreach (var itemPickingItem in resultPickingItem)
                                            {
                                                PickFreebieQty = Convert.ToInt32(itemPickingItem.PickFreebieQty); //撿貨贈品數
                                            }
                                        }
                                        else
                                        {
                                            PickFreebieQty = 0;
                                        }
                                        #endregion\                                   

                                        if (NormalOrderFreebieQty < TemporaryOrderGiftQty + PickFreebieQty - TemporaryReturnOrderGiftQty)
                                        {
                                            int Available = NormalOrderFreebieQty - TemporaryOrderGiftQty + TemporaryReturnOrderGiftQty;
                                            throw new SystemException("贈品量超交，本次最大可贈品量為:[" + Available + "]");
                                        }

                                    }
                                }
                                #endregion

                                #region//備品數卡控
                                int NormalOrderSpareQty = 0; //訂單備品數量
                                int TemporaryQrderSparePartsQty = 0; //暫出單備品數
                                int TemporaryReturnQrderSparePartsQty = 0; //暫出歸還備品數 
                                int PickSpareQty = 0;//撿貨備品數

                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT c.SoErpPrefix,c.SoErpNo,b.SoSequence
                                    FROM SCM.DoDetail a
                                    INNER JOIN SCM.SoDetail b ON a.SoDetailId=b.SoDetailId
                                    INNER JOIN SCM.SaleOrder  c ON b.SoId=c.SoId  
                                    WHERE a.DoDetailId=@DoDetailId";
                                dynamicParameters.Add("DoDetailId", doDetailId);
                                var resultSpare = sqlConnection.Query(sql, dynamicParameters);
                                if (resultSpare.Count() > 0)
                                {
                                    foreach (var item3 in resultSpare)
                                    {
                                        string SoErpPrefix = item3.SoErpPrefix;//訂單單別
                                        string SoErpNo = item3.SoErpNo;//訂單單號
                                        string SoSequence = item3.SoSequence;//訂單流水號

                                        #region//取出ERP暫出歸還數量                                    
                                        using (SqlConnection sqlConnectionErp = new SqlConnection(ErpConnectionStrings))
                                        {
                                            //1.檢查訂單單身數量
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"SELECT b.TD050
                                                FROM COPTC a
                                                INNER JOIN COPTD b ON a.TC001=b.TD001 AND a.TC002=b.TD002
                                                WHERE a.TC027 !='V'
                                                AND b.TD001=@SoErpPrefix
                                                AND b.TD002=@SoErpNo
                                                AND b.TD003=@SoSequence";
                                            dynamicParameters.Add("SoErpPrefix", SoErpPrefix);
                                            dynamicParameters.Add("SoErpNo", SoErpNo);
                                            dynamicParameters.Add("SoSequence", SoSequence);
                                            var resultCOPTC = sqlConnectionErp.Query(sql, dynamicParameters);
                                            if (resultCOPTC.Count() > 0)
                                            {
                                                foreach (var itemCOPTC in resultCOPTC)
                                                {
                                                    NormalOrderSpareQty = Convert.ToInt32(itemCOPTC.TD050);
                                                }
                                            }

                                            //2.檢查這張訂單沒有開過暫出單
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"SELECT b.TG001,b.TG002,b.TG003,b.TG014,b.TG015,b.TG016,b.TG044
                                                FROM INVTF a
                                                INNER JOIN INVTG b ON a.TF001=b.TG001 AND a.TF002=b.TG002
                                                WHERE a.TF020 !='V' AND b.TG043='2'
                                                AND b.TG014=@SoErpPrefix
                                                AND b.TG015=@SoErpNo
                                                AND b.TG016=@SoSequence";
                                            dynamicParameters.Add("SoErpPrefix", SoErpPrefix);
                                            dynamicParameters.Add("SoErpNo", SoErpNo);
                                            dynamicParameters.Add("SoSequence", SoSequence);
                                            var resultINVTF = sqlConnectionErp.Query(sql, dynamicParameters);
                                            if (resultINVTF.Count() > 0)
                                            {
                                                foreach (var itemINVTF in resultINVTF)
                                                {
                                                    //3.檢查所開的暫出單有沒有開過【暫出歸還單】
                                                    dynamicParameters = new DynamicParameters();
                                                    sql = @"SELECT a.TI035
                                                FROM INVTI a
                                                INNER JOIN INVTG b ON a.TI014 = b.TG001 AND a.TI015=b.TG002 AND a.TI016=b.TG003
                                                WHERE a.TI022 !='V' AND a.TI034='2'
                                                AND b.TG001=@TsnErpPrefix
                                                AND b.TG002=@TsnErpNo
                                                AND b.TG003=@TsnSequence";
                                                    dynamicParameters.Add("TsnErpPrefix", itemINVTF.TG001);
                                                    dynamicParameters.Add("TsnErpNo", itemINVTF.TG002);
                                                    dynamicParameters.Add("TsnSequence", itemINVTF.TG003);
                                                    var resultINVTI = sqlConnectionErp.Query(sql, dynamicParameters);
                                                    if (resultINVTI.Count() > 0)
                                                    {
                                                        foreach (var itemINVTI in resultINVTI)
                                                        {
                                                            TemporaryReturnQrderSparePartsQty += Convert.ToInt32(itemINVTI.TI035); //暫出單歸還數
                                                        }
                                                    }
                                                    else
                                                    {
                                                        TemporaryReturnQrderSparePartsQty = 0;
                                                    }
                                                    TemporaryQrderSparePartsQty += Convert.ToInt32(itemINVTF.TG044);
                                                }
                                            }
                                            else
                                            {
                                                TemporaryQrderSparePartsQty = 0;
                                            }

                                        }
                                        #endregion

                                        #region//本次撿貨數量
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT  ISNULL(e.PickQty,0) PickQty
                                                , ISNULL(f.PickFreebieQty,0) PickFreebieQty
                                                ,ISNULL(g.PickSpareQty,0)PickSpareQty
                                            FROM SCM.DoDetail a
                                                INNER JOIN SCM.DeliveryOrder b ON a.DoId = b.DoId
                                                INNER JOIN SCM.SoDetail c ON a.SoDetailId = c.SoDetailId
                                                INNER JOIN SCM.SaleOrder d ON c.SoId = d.SoId
                                                OUTER APPLY (
                                                    SELECT SUM(ea.ItemQty) PickQty
                                                    FROM SCM.PickingItem ea
                                                    WHERE ea.SoDetailId = a.SoDetailId
                                                    AND ea.ItemType = 1
                                                    AND ea.DoId = a.DoId
                                                ) e
                                                OUTER APPLY (
                                                    SELECT SUM(fa.ItemQty) PickFreebieQty
                                                    FROM SCM.PickingItem fa
                                                    WHERE fa.SoDetailId = a.SoDetailId
                                                    AND fa.ItemType = 2
                                                    AND fa.DoId = a.DoId
                                                ) f
                                                OUTER APPLY (
                                                    SELECT SUM(ga.ItemQty) PickSpareQty
                                                    FROM SCM.PickingItem ga
                                                    WHERE ga.SoDetailId = a.SoDetailId
                                                    AND ga.ItemType = 3
                                                    AND ga.DoId = a.DoId
                                                ) g
                                                LEFT JOIN SCM.Inventory h ON a.TransInInventoryId = h.InventoryId
                                        WHERE a.DoDetailId = @DoDetailId
                                        AND b.[Status] = @Status
                                        AND d.CompanyId = @CompanyId
                                        ORDER BY d.SoErpPrefix, d.SoErpNo, c.SoSequence";
                                        dynamicParameters.Add("DoDetailId", doDetailId);
                                        dynamicParameters.Add("Status", "S");
                                        dynamicParameters.Add("CompanyId", CurrentCompany);
                                        var resultPickingItem = sqlConnection.Query(sql, dynamicParameters);
                                        if (resultPickingItem.Count() > 0)
                                        {
                                            foreach (var itemPickingItem in resultPickingItem)
                                            {
                                                PickSpareQty = Convert.ToInt32(itemPickingItem.PickSpareQty); //撿貨數
                                            }
                                        }
                                        else
                                        {
                                            PickSpareQty = 0;
                                        }
                                        #endregion

                                        if (NormalOrderSpareQty < TemporaryQrderSparePartsQty + PickSpareQty - TemporaryReturnQrderSparePartsQty)
                                        {
                                            int Available = NormalOrderSpareQty - TemporaryQrderSparePartsQty + TemporaryReturnQrderSparePartsQty;
                                            throw new SystemException("備品量超交，本次最大可備品量為:[" + Available + "]");
                                        }
                                    }
                                }
                                #endregion


                            }
                        }
                    }
                    #endregion

                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        erpPrefix = TsnErpPrefix;
                        docDate = iNVTFs.Select(x => x.TF024).FirstOrDefault();
                        DateTime referenceTime = DateTime.ParseExact(docDate, "yyyyMMdd", CultureInfo.InvariantCulture);
                        string DocDate = referenceTime.ToString("yyyy-MM-dd");

                        #region //比對ERP關帳日期
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT LTRIM(RTRIM(MA013)) MA013
                                    FROM CMSMA";
                        var cmsmaResult = sqlConnection.Query(sql, dynamicParameters);
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

                        var resultDocSetting = sqlConnection.Query(sql, dynamicParameters);
                        if (resultDocSetting.Count() <= 0) throw new SystemException("ERP單據設定資料不存在!");

                        string encode = "", exchangeRateSource = "";
                        int yearLength = 0, lineLength = 0;
                        foreach (var item in resultDocSetting)
                        {
                            encode = item.MQ004; //編碼方式
                            yearLength = Convert.ToInt32(item.MQ005); //年碼數
                            lineLength = Convert.ToInt32(item.MQ006); //流水號碼數
                            exchangeRateSource = item.MQ044; //匯率來源
                        }
                        #endregion

                        #region //單號取號
                        int currentNum = 0;

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT CONVERT(int, RIGHT(ISNULL(MAX(RTRIM(LTRIM(TF002))), '" + new string('0', lineLength) + @"'), " + lineLength + @")) + 1 CurrentNum
                                FROM INVTF
                                WHERE TF001 = @ErpPrefix";
                        dynamicParameters.Add("ErpPrefix", erpPrefix);

                        #region //編碼方式
                        string dateFormat = "";
                        switch (encode)
                        {
                            case "1": //日編
                                dateFormat = new string('y', yearLength) + "MMdd";
                                sql += @" AND RTRIM(LTRIM(TF002)) LIKE @ReferenceTime + '" + new string('_', lineLength) + @"'";
                                dynamicParameters.Add("ReferenceTime", referenceTime.ToString(dateFormat));
                                erpNo = referenceTime.ToString(dateFormat);
                                break;
                            case "2": //月編
                                dateFormat = new string('y', yearLength) + "MM";
                                sql += @" AND RTRIM(LTRIM(TF002)) LIKE @ReferenceTime + '" + new string('_', lineLength) + @"'";
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
                        #endregion

                        #region //廠別資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 MB001
                                FROM CMSMB
                                WHERE COMPANY = @COMPANY";
                        dynamicParameters.Add("COMPANY", companyNo);

                        var resultFactory = sqlConnection.Query(sql, dynamicParameters);
                        if (resultFactory.Count() <= 0) throw new SystemException("ERP廠別資料不存在!");

                        string factory = "";
                        foreach (var item in resultFactory)
                        {
                            factory = item.MB001; //廠別
                        }
                        #endregion

                        #region //客戶資料
                        string customerNo = iNVTFs.Select(x => x.TF005).FirstOrDefault();

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT MA002, MA003, MA014, MA027, MA038, MA064, MA118
                                FROM COPMA
                                WHERE MA001 = @CustomerNo";
                        dynamicParameters.Add("CustomerNo", customerNo);

                        var resultCustomer = sqlConnection.Query(sql, dynamicParameters);
                        if (resultCustomer.Count() <= 0) throw new SystemException("ERP客戶資料不存在!");

                        foreach (var item in resultCustomer)
                        {
                           
                            iNVTFs
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.TF006 = item.MA002; //對象簡稱
                                    x.TF015 = item.MA003; //對象全名
                                    x.TF016 = item.MA027; //地址一
                                    x.TF017 = item.MA064; //地址二
                                });
                        }
                        #endregion

                        #region //訂單取得交易幣別、課稅別、稅別碼
                        string firstSoErpPrefix = iNVTGs.Select(x => x.TG014).FirstOrDefault();
                        string firstSoErpNo = iNVTGs.Select(x => x.TG015).FirstOrDefault();

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TC008, TC016, TC078
                                FROM COPTC
                                WHERE TC001 = @SoErpPrefix AND TC002 = @SoErpNo";
                        dynamicParameters.Add("SoErpPrefix", firstSoErpPrefix);
                        dynamicParameters.Add("SoErpNo", firstSoErpNo);

                        var resultOrder = sqlConnection.Query(sql, dynamicParameters);
                        if (resultOrder.Count() <= 0) throw new SystemException("ERP訂單資料不存在!");

                        string currency = "", taxation = "", taxNo = "";
                        foreach (var item in resultOrder)
                        {
                            currency = item.TC008;   // 交易幣別
                            taxation = item.TC016;   // 課稅別
                            taxNo = item.TC078;      // 稅別碼

                            iNVTFs
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.TF010 = taxation; //課稅別
                                    x.TF011 = currency; //幣別
                                });
                        }
                    
                        #endregion

                        #region //交易幣別設定
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT MF003, MF004
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
                        sql = @"SELECT NN004
                                FROM CMSNN
                                WHERE NN001 = @TaxNo";
                        dynamicParameters.Add("TaxNo", taxNo);

                        var resultTax = sqlConnection.Query(sql, dynamicParameters);
                        if (resultTax.Count() <= 0) throw new SystemException("ERP稅別碼設定不存在!");

                        double exciseTax = 0;
                        foreach (var item in resultTax)
                        {
                            exciseTax = Convert.ToDouble(item.NN004); //營業稅率
                        }
                        #endregion

                        #region //判斷訂單是否存在
                        foreach (var so in iNVTGs)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(b.MB001)) MB001, a.TD005, a.TD006, a.TD010, a.TD011
                                    , a.TD020, a.TD027
                                    FROM COPTD a
                                    INNER JOIN INVMB b ON a.TD004 = b.MB001
                                    WHERE a.TD001 = @ErpPrefix
                                    AND a.TD002 = @ErpNo
                                    AND a.TD003 = @ErpSequence";
                            dynamicParameters.Add("ErpPrefix", so.TG014);
                            dynamicParameters.Add("ErpNo", so.TG015);
                            dynamicParameters.Add("ErpSequence", so.TG016);

                            var resultSaleOrder = sqlConnection.Query(sql, dynamicParameters);
                            if (resultSaleOrder.Count() <= 0) throw new SystemException("ERP訂單資料不存在!");

                            foreach (var item in resultSaleOrder)
                            {
                                #region//檢核品號於轉出庫的數量是否足夠
                                using (SqlConnection sqlConnectionErp = new SqlConnection(ErpConnectionStrings))
                                {
                                    string checkOutInventoryNo = deliveryToTsns.Where(y => y.SoErpPrefix == so.TG014 && y.SoErpNo == so.TG015 && y.SoSequence == so.TG016).Select(y => y.TsnOutInventoryNo).FirstOrDefault();
                                    string checkMtlItemNo = item.MB001;
                                }
                                #endregion

                                #region //設定 INVTG
                                iNVTGs
                                .Where(x => x.TG014 == so.TG014 && x.TG015 == so.TG015 && x.TG016 == so.TG016)
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.TG001 = erpPrefix; //異動單別
                                    x.TG002 = erpNo; //異動單號
                                    x.TG004 = item.MB001; //品號
                                    x.TG005 = item.TD005; //品名
                                    x.TG006 = item.TD006; //規格
                                    x.TG010 = item.TD010; //單位
                                    x.TG011 = ""; //小單位
                                    x.TG012 = Math.Round(Convert.ToDouble(item.TD011), unitRound); //單價
                                    x.TG013 = Math.Round(Math.Round(Convert.ToDouble(item.TD011), unitRound) * x.TG009, totalRound); //金額
                                    x.TG018 = item.TD027; //專案代號
                                    x.TG019 = item.TD020; //備註
                                    x.TG020 = 0; //轉進銷量
                                    x.TG021 = 0; //歸還量
                                    x.TG022 = "N"; //確認碼
                                    x.TG023 = "N"; //更新碼
                                    x.TG024 = "N"; //結案碼
                                    x.TG025 = ""; //有效日期
                                    x.TG026 = ""; //複檢日期
                                    x.TG027 = ""; //預計歸還日
                                    x.TG028 = 0; //包裝數量
                                    x.TG029 = 0; //轉進銷包裝量
                                    x.TG030 = 0; //歸還包裝量
                                    x.TG031 = ""; //包裝單位
                                    x.TG032 = ""; //預留欄位
                                    x.TG033 = 0; //產品序號數量
                                    x.TG034 = 0; //預留欄位
                                    x.TG035 = ""; //轉出儲位
                                    x.TG036 = ""; //轉入儲位
                                    x.TG037 = 0; //預留欄位
                                    x.TG038 = 0; //預留欄位
                                    x.TG039 = ""; //預留欄位
                                    x.TG040 = ""; //最終客戶代號
                                    x.TG041 = ""; //預留欄位
                                    x.TG042 = exciseTax; //營業稅率
                                    x.TG045 = 0; //贈/備品包裝量
                                    x.TG046 = 0; //轉銷贈/備品量
                                    x.TG047 = 0; //轉銷贈/備品包裝量
                                    x.TG048 = 0; //歸還贈/備品量
                                    x.TG049 = 0; //歸還贈/備品包裝量
                                    x.TG050 = ""; //業務品號
                                    x.TG500 = ""; //刻號/BIN記錄
                                    x.TG501 = ""; //刻號管理
                                    x.TG502 = ""; //DATECODE
                                    x.TG503 = ""; //產品系列
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
                                    x.TG051 = ""; //稅別碼
                                    x.TG053 = item.TD010; //計價單位
                                    x.TG054 = 0;
                                    x.TG055 = 0;
                                });
                                #endregion
                            }
                        }
                        #endregion

                        #region //計算數量與金額
                        int? totalQty = iNVTGs.Sum(i => i.TG009) + iNVTGs.Sum(i => i.TG044);
                        double? totalPrice = iNVTGs.Sum(i => i.TG013);
                        double? totalTax = Math.Round((double)totalPrice * exciseTax, totalRound);

                        switch (taxation)
                        {
                            case "1": //應稅內含
                                totalTax = totalPrice - Math.Round((double)totalPrice / (1 + exciseTax), totalRound);
                                totalPrice = Math.Round((double)totalPrice / (1 + exciseTax), totalRound);
                                break;
                            case "2": //應稅外加
                                break;
                            case "3": //零稅率
                                break;
                            case "4": //免稅
                                break;
                            case "9": //不計稅
                                break;
                        }
                        #endregion

                        #region //設定 INVTF
                        #region //判斷單號是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM INVTF
                                WHERE TF001 = @ErpPrefix
                                AND TF002 = @ErpNo";
                        dynamicParameters.Add("ErpPrefix", erpPrefix);
                        dynamicParameters.Add("ErpNo", erpNo);

                        var resultRepeatExist = sqlConnection.Query(sql, dynamicParameters);
                        if (resultRepeatExist.Count() > 0) throw new SystemException("【異動單號】重複，請重新取號!");
                        #endregion

                        iNVTFs
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
                                x.TF001 = erpPrefix; //異動單別
                                x.TF002 = erpNo; //異動單號
                                x.TF007 = departmentNo; //部門代號
                                x.TF008 = salesUserNo; //員工代號
                                x.TF009 = factory; //廠別
                                x.TF012 = exchangeRate; //匯率
                                x.TF013 = 0; //件數
                                x.TF014 = Remark; //備註
                                x.TF018 = ""; //其它備註
                                x.TF019 = 0; //列印次數
                                x.TF020 = "N"; //確認碼
                                x.TF021 = "N"; //更新碼
                                x.TF022 = totalQty; //總數量
                                x.TF023 = totalPrice; //總金額
                                x.TF025 = ""; //確認者
                                x.TF026 = exciseTax; //營業稅率
                                x.TF027 = totalTax; //稅額
                                x.TF028 = 0; //總包裝數量
                                x.TF029 = "N"; //簽核狀態碼
                                x.TF030 = 0; //傳送次數
                                x.TF031 = ""; //運輸方式
                                x.TF032 = ""; //派車單別
                                x.TF033 = ""; //派車單號
                                x.TF034 = 0; //預留欄位
                                x.TF035 = 0; //預留欄位
                                x.TF036 = ""; //來源
                                x.TF037 = ""; //銷貨單單價別
                                x.TF038 = ""; //預留欄位
                                x.TF039 = taxNo; //稅別碼
                                x.TF040 = "N"; //單身多稅率
                                x.TF042 = ""; //轉銷日
                                x.TF043 = "N"; //出通單更新碼
                                x.TF044 = "N"; //不控管信用額度
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
                                x.TF045 = "";
                                x.TF046 = "";
                            });
                        #endregion

                        #region//暫出單 ERP信用額度卡控
                        string CompanyNo = CurrentCompanyNo;
                        string CustomerNo = iNVTFs.Select(x => x.TF005).FirstOrDefault(); //客戶代碼
                        string Currency = currency;
                        string DocType = "TempShippingNote";
                        decimal Amount = 0, TotalAmount = Convert.ToDecimal(totalPrice);

                        #region //信用額度檢核
                        //string domainUrl = "http://192.168.20.136:16668/";
                        string domainUrl = "https://bm.zy-tech.com.tw/";


                        string targetDate = DateTime.Now.AddDays(-2).ToString("yyyyMMdd HH:mm:ss"),
                            targetUrl = string.Format("{0}{1}", domainUrl, "api/BM/CheckCreditLimit");

                        var postData = new List<KeyValuePair<string, string>>
                        {
                            new KeyValuePair<string, string>("CustomerNo", CustomerNo),
                            new KeyValuePair<string, string>("Currency", Currency),
                            new KeyValuePair<string, string>("TotalAmount", TotalAmount.ToString()),
                            new KeyValuePair<string, string>("DocType", "ShippingOrder"),
                            new KeyValuePair<string, string>("Amount", "0"),
                            new KeyValuePair<string, string>("CompanyNo", CompanyNo)
                        };

                        string response = BaseHelper.PostWebRequest(targetUrl, postData);

                        if (response.TryParseJson(out JObject tempJObject))
                        {
                            JObject resultJson = JObject.Parse(response);

                            Console.WriteLine("狀態：" + resultJson["status"].ToString());
                            Console.WriteLine("回傳訊息：" + resultJson["msg"].ToString());

                            if (resultJson["status"].ToString() != "ok")
                            {
                                throw new SystemException(resultJson["msg"].ToString());
                            }
                        }
                        else
                        {
                            logger.Error(response);
                            throw new SystemException(response);
                        }
                        #endregion
                        #endregion

                        #region //拋轉ERP
                        #region //INVTF
                        sql = @"INSERT INTO INVTF (COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE
                                , TF001, TF002, TF003, TF004, TF005, TF006, TF007, TF008, TF009, TF010
                                , TF011, TF012, TF013, TF014, TF015, TF016, TF017, TF018, TF019, TF020
                                , TF021, TF022, TF023, TF024, TF025, TF026, TF027, TF028, TF029, TF030
                                , TF031, TF032, TF033, TF034, TF035, TF036, TF037, TF038, TF039, TF040
                                , TF041, TF042, TF043, TF044, TF045, TF046
                                , FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID
                                , UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10)
                                VALUES (@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE
                                , @TF001, @TF002, @TF003, @TF004, @TF005, @TF006, @TF007, @TF008, @TF009, @TF010
                                , @TF011, @TF012, @TF013, @TF014, @TF015, @TF016, @TF017, @TF018, @TF019, @TF020
                                , @TF021, @TF022, @TF023, @TF024, @TF025, @TF026, @TF027, @TF028, @TF029, @TF030
                                , @TF031, @TF032, @TF033, @TF034, @TF035, @TF036, @TF037, @TF038, @TF039, @TF040
                                , @TF041, @TF042, @TF043, @TF044, @TF045, @TF046
                                , @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID, @MODI_TIME, @MODI_AP, @MODI_PRID
                                , @UDF01, @UDF02, @UDF03, @UDF04, @UDF05, @UDF06, @UDF07, @UDF08, @UDF09, @UDF10)";
                        rowsAffected += sqlConnection.Execute(sql, iNVTFs);
                        #endregion

                        #region //INVTG
                        sql = @"INSERT INTO INVTG (COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE
                                , TG001, TG002, TG003, TG004, TG005, TG006, TG007, TG008, TG009, TG010
                                , TG011, TG012, TG013, TG014, TG015, TG016, TG017, TG018, TG019, TG020
                                , TG021, TG022, TG023, TG024, TG025, TG026, TG027, TG028, TG029, TG030
                                , TG031, TG032, TG033, TG034, TG035, TG036, TG037, TG038, TG039, TG040
                                , TG041, TG042, TG043, TG044, TG045, TG046, TG047, TG048, TG049, TG050
                                , TG051, TG052, TG053, TG054, TG055
                                , TG500, TG501, TG502, TG503
                                , FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID
                                , UDF01, UDF02, UDF03, UDF04, UDF05, UDF06, UDF07, UDF08, UDF09, UDF10)
                                VALUES (@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE
                                , @TG001, @TG002, @TG003, @TG004, @TG005, @TG006, @TG007, @TG008, @TG009, @TG010
                                , @TG011, @TG012, @TG013, @TG014, @TG015, @TG016, @TG017, @TG018, @TG019, @TG020
                                , @TG021, @TG022, @TG023, @TG024, @TG025, @TG026, @TG027, @TG028, @TG029, @TG030
                                , @TG031, @TG032, @TG033, @TG034, @TG035, @TG036, @TG037, @TG038, @TG039, @TG040
                                , @TG041, @TG042, @TG043, @TG044, @TG045, @TG046, @TG047, @TG048, @TG049, @TG050
                                , @TG051, @TG052, @TG053, @TG054, @TG055
                                , @TG500, @TG501, @TG502, @TG503
                                , @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID, @MODI_TIME, @MODI_AP, @MODI_PRID
                                , @UDF01, @UDF02, @UDF03, @UDF04, @UDF05, @UDF06, @UDF07, @UDF08, @UDF09, @UDF10)";
                        rowsAffected += sqlConnection.Execute(sql, iNVTGs);
                        #endregion
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //暫出單單頭
                        foreach (var item in iNVTFs)
                        {
                            tempShippingNotes.Add(
                                new TempShippingNote
                                {
                                    TsnErpPrefix = item.TF001,
                                    TsnErpNo = item.TF002,
                                    TsnDate = DateTime.ParseExact(item.TF003, "yyyyMMdd", CultureInfo.InvariantCulture),
                                    DocDate = DateTime.ParseExact(item.TF024, "yyyyMMdd", CultureInfo.InvariantCulture),
                                    ToObject = item.TF004,
                                    ObjectOther = item.TF005,
                                    DepartmentNo = item.TF007,
                                    UserNo = item.TF008,
                                    Remark = item.TF014,
                                    OtherRemark = item.TF018,
                                    TotalQty = item.TF022,
                                    Amount = item.TF023,
                                    TaxAmount = item.TF027,
                                    ConfirmStatus = item.TF020,
                                    ConfirmUserNo = item.TF025,
                                    CompanyId = CurrentCompany,
                                    TransferStatus = "Y",
                                    CreateDate = CreateDate,
                                    LastModifiedDate = LastModifiedDate,
                                    CreateBy = CreateBy,
                                    LastModifiedBy = LastModifiedBy,
                                });
                        }
                        #endregion

                        #region //暫出單單身
                        foreach (var item in iNVTGs)
                        {
                            tsnDetails.Add(
                                new TsnDetail
                                {
                                    TsnErpPrefix = item.TG001,
                                    TsnErpNo = item.TG002,
                                    TsnSequence = item.TG003,
                                    MtlItemNo = item.TG004,
                                    TsnMtlItemName = item.TG005,
                                    TsnMtlItemSpec = item.TG006,
                                    TsnOutInventoryNo = item.TG007,
                                    TsnInInventoryNo = item.TG008,
                                    TsnQty = item.TG009,
                                    ProductType = item.TG043,
                                    FreebieOrSpareQty = item.TG044,
                                    UnitPrice = item.TG012,
                                    Amount = item.TG013,
                                    SoErpPrefix = item.TG014,
                                    SoErpNo = item.TG015,
                                    SoSequence = item.TG016,
                                    LotNumber = item.TG017,
                                    TsnRemark = item.TG019,
                                    ConfirmStatus = item.TG022,
                                    ClosureStatus = item.TG024,
                                    SaleQty = item.TG020,
                                    ReturnQty = item.TG021,
                                    CreateDate = CreateDate,
                                    LastModifiedDate = LastModifiedDate,
                                    CreateBy = CreateBy,
                                    LastModifiedBy = LastModifiedBy
                                });
                        }
                        #endregion

                        #region //撈取部門ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT DepartmentId, DepartmentNo
                                FROM BAS.Department
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        List<Department> resultDepartments = sqlConnection.Query<Department>(sql, dynamicParameters).ToList();

                        tempShippingNotes = tempShippingNotes.Join(resultDepartments, x => x.DepartmentNo, y => y.DepartmentNo, (x, y) => { x.DepartmentId = y.DepartmentId; return x; }).ToList();
                        #endregion

                        #region //撈取人員ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserId, a.UserNo
                                FROM BAS.[User] a
                                LEFT JOIN BAS.Department b ON b.DepartmentId = a.DepartmentId
                                WHERE b.CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        List<User> resultUsers = sqlConnection.Query<User>(sql, dynamicParameters).ToList();

                        tempShippingNotes = tempShippingNotes.Join(resultUsers, x => x.UserNo, y => y.UserNo, (x, y) => { x.UserId = y.UserId; return x; }).ToList();
                        tempShippingNotes = tempShippingNotes.GroupJoin(resultUsers, x => x.ConfirmUserNo, y => y.UserNo, (x, y) => { x.ConfirmUserId = y.FirstOrDefault()?.UserId; return x; }).ToList();
                        List<TempShippingNote> toObjectUser = tempShippingNotes.Where(x => x.ToObject == "3").Join(resultUsers, x => x.ObjectOther, y => y.UserNo, (x, y) => { x.ObjectUser = y.UserId; return x; }).Select(x => x).ToList();
                        #endregion

                        #region //撈取異動對象ID
                        #region //撈取客戶ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT CustomerId, CustomerNo 
                                FROM SCM.Customer
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        List<Customer> resultCustomers = sqlConnection.Query<Customer>(sql, dynamicParameters).ToList();

                        List<TempShippingNote> toObjectCustomer = tempShippingNotes.Where(x => x.ToObject == "1").Join(resultCustomers, x => x.ObjectOther, y => y.CustomerNo, (x, y) => { x.ObjectCustomer = y.CustomerId; return x; }).Select(x => x).ToList();
                        #endregion

                        #region //撈取供應商ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT SupplierId, SupplierNo 
                                FROM SCM.Supplier
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        List<Supplier> resultSuppliers = sqlConnection.Query<Supplier>(sql, dynamicParameters).ToList();

                        List<TempShippingNote> toObjectSupplier = tempShippingNotes.Where(x => x.ToObject == "2").Join(resultSuppliers, x => x.ObjectOther, y => y.SupplierNo, (x, y) => { x.ObjectSupplier = y.SupplierId; return x; }).Select(x => x).ToList();
                        #endregion

                        #region //其他
                        List<TempShippingNote> toObjectOther = tempShippingNotes.Where(x => x.ToObject == "9").Select(x => x).ToList();
                        #endregion

                        tempShippingNotes = toObjectCustomer.Union(toObjectSupplier).Union(toObjectUser).Union(toObjectOther).ToList();
                        #endregion

                        #region //撈取品號ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT MtlItemId, MtlItemNo
                                FROM PDM.MtlItem
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        List<MtlItem> resultMtlitems = sqlConnection.Query<MtlItem>(sql, dynamicParameters).ToList();

                        tsnDetails = tsnDetails.Join(resultMtlitems, x => x.MtlItemNo, y => y.MtlItemNo, (x, y) => { x.MtlItemId = y.MtlItemId; return x; }).Select(x => x).ToList();
                        #endregion

                        #region //撈取庫別ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT InventoryId, InventoryNo
                                FROM SCM.Inventory
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        List<Inventory> resultInventories = sqlConnection.Query<Inventory>(sql, dynamicParameters).ToList();

                        tsnDetails = tsnDetails.Join(resultInventories, x => x.TsnOutInventoryNo, y => y.InventoryNo, (x, y) => { x.TsnOutInventory = y.InventoryId; return x; }).ToList();
                        tsnDetails = tsnDetails.Join(resultInventories, x => x.TsnInInventoryNo, y => y.InventoryNo, (x, y) => { x.TsnInInventory = y.InventoryId; return x; }).ToList();
                        #endregion

                        #region //撈取訂單單身ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.SoId, a.SoErpPrefix, a.SoErpNo, b.SoDetailId, b.SoSequence
                                FROM SCM.SaleOrder a
                                LEFT JOIN SCM.SoDetail b ON b.SoId = a.SoId
                                WHERE a.CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        List<SoDetail> resultSoDetails = sqlConnection.Query<SoDetail>(sql, dynamicParameters).ToList();

                        tsnDetails = tsnDetails.GroupJoin(resultSoDetails, x => new { x.SoErpPrefix, x.SoErpNo, x.SoSequence }, y => new { y.SoErpPrefix, y.SoErpNo, y.SoSequence }, (x, y) => { x.SoDetailId = y.FirstOrDefault()?.SoDetailId; return x; }).Select(x => x).ToList();
                        #endregion

                        #region //暫出單單頭BM新增
                        sql = @"INSERT INTO SCM.TempShippingNote (CompanyId, TsnErpPrefix, TsnErpNo, TsnDate
                                , DocDate, ToObject, ObjectCustomer, ObjectSupplier, ObjectUser, ObjectOther
                                , DepartmentId, UserId, Remark, OtherRemark, TotalQty, Amount, TaxAmount
                                , ConfirmStatus, ConfirmUserId, TransferStatus, TransferDate
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                VALUES (@CompanyId, @TsnErpPrefix, @TsnErpNo, @TsnDate
                                , @DocDate, @ToObject, @ObjectCustomer, @ObjectSupplier, @ObjectUser, @ObjectOther
                                , @DepartmentId, @UserId, @Remark, @OtherRemark, @TotalQty, @Amount, @TaxAmount
                                , @ConfirmStatus, @ConfirmUserId, @TransferStatus, @TransferDate
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        rowsAffected += sqlConnection.Execute(sql, tempShippingNotes);
                        #endregion

                        #region //撈取暫出單單頭ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TsnId, TsnErpPrefix, TsnErpNo
                                FROM  SCM.TempShippingNote
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var resultTempShippingNotes = sqlConnection.Query<TempShippingNote>(sql, dynamicParameters).ToList();

                        tsnDetails = tsnDetails.Join(resultTempShippingNotes, x => new { x.TsnErpPrefix, x.TsnErpNo }, y => new { y.TsnErpPrefix, y.TsnErpNo }, (x, y) => { x.TsnId = y.TsnId; return x; }).Select(x => x).ToList();
                        #endregion

                        #region //暫出單單身BM新增
                        sql = @"INSERT INTO SCM.TsnDetail (TsnId, TsnSequence, MtlItemId, TsnMtlItemName
                                , TsnMtlItemSpec, TsnOutInventory, TsnInInventory, TsnQty, ProductType, FreebieOrSpareQty
                                , UnitPrice, Amount, SoDetailId, LotNumber, TsnRemark, ConfirmStatus, ClosureStatus
                                , SaleQty, ReturnQty
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                VALUES (@TsnId, @TsnSequence, @MtlItemId, @TsnMtlItemName
                                , @TsnMtlItemSpec, @TsnOutInventory, @TsnInInventory, @TsnQty, @ProductType, @FreebieOrSpareQty
                                , @UnitPrice, @Amount, @SoDetailId, @LotNumber, @TsnRemark, @ConfirmStatus, @ClosureStatus
                                , @SaleQty, @ReturnQty
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        rowsAffected += sqlConnection.Execute(sql, tsnDetails);
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

        #region //Update
        #region //UpdateInventoryTransactionSynchronize -- 庫存異動資料同步 -- Ben Ma 2023.03.15
        public string UpdateInventoryTransactionSynchronize(string CompanyNo, string UpdateDate)
        {
            try
            {
                List<InventoryTransaction> inventoryTransactions = new List<InventoryTransaction>();
                List<ItDetail> itDetails = new List<ItDetail>();

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

                        var resultCompany = sqlConnection.Query(sql, dynamicParameters);
                        if (resultCompany.Count() <= 0) throw new SystemException("【公司】資料錯誤!");

                        foreach (var item in resultCompany)
                        {
                            CompanyId = Convert.ToInt32(item.CompanyId);
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        }
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        #region //撈取ERP庫存異動資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT LTRIM(RTRIM(TA001)) ItErpPrefix, LTRIM(RTRIM(TA002)) ItErpNo
                                , CASE WHEN LEN(LTRIM(RTRIM(TA003))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TA003)) as date), 'yyyy-MM-dd') ELSE NULL END ItDate
                                , CASE WHEN LEN(LTRIM(RTRIM(TA014))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TA014)) as date), 'yyyy-MM-dd') ELSE NULL END DocDate
                                , LTRIM(RTRIM(TA004)) DepartmentNo, LTRIM(RTRIM(TA005)) Remark
                                , TA011 TotalQty, TA012 Amount
                                , LTRIM(RTRIM(TA006)) ConfirmStatus, LTRIM(RTRIM(TA015)) ConfirmUserNo
                                , CASE WHEN LEN(LTRIM(RTRIM(CREATE_DATE))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_DATE)) as date), 'yyyy-MM-dd') ELSE NULL END TransferDate
                                , CASE WHEN LEN(LTRIM(RTRIM(CREATE_TIME))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_TIME)) as datetime), 'HH:mm:ss') ELSE NULL END TransferTime
                                FROM INVTA
                                WHERE 1=1";
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "UpdateDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME > @UpdateDate OR MODI_DATE + ' ' + MODI_TIME > @UpdateDate)", UpdateDate);
                        sql += @" ORDER BY TransferDate, TransferTime";

                        inventoryTransactions = sqlConnection.Query<InventoryTransaction>(sql, dynamicParameters).ToList();
                        #endregion

                        #region //撈取ERP庫存異動單身資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT LTRIM(RTRIM(TB001)) ItErpPrefix, LTRIM(RTRIM(TB002)) ItErpNo, LTRIM(RTRIM(TB003)) ItSequence
                                , LTRIM(RTRIM(TB004)) MtlItemNo, LTRIM(RTRIM(TB005)) ItMtlItemName, LTRIM(RTRIM(TB006)) ItMtlItemSpec
                                , LTRIM(RTRIM(CAST(TB007 AS decimal(16,3)))) ItQty, LTRIM(RTRIM(CAST(TB009 AS decimal(16,3)))) InvQty
                                , LTRIM(RTRIM(CAST(TB010 AS decimal(21,6)))) UnitCost, LTRIM(RTRIM(CAST(TB011 AS decimal(21,6)))) Amount
                                , LTRIM(RTRIM(TB012)) InventoryNo, LTRIM(RTRIM(TB017)) ItRemark
                                FROM INVTB
                                WHERE 1=1";
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "UpdateDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME > @UpdateDate OR MODI_DATE + ' ' + MODI_TIME > @UpdateDate)", UpdateDate);

                        itDetails = sqlConnection.Query<ItDetail>(sql, dynamicParameters).ToList();
                        #endregion
                    }

                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //撈取部門ID
                        sql = @"SELECT DepartmentId, DepartmentNo
                                FROM BAS.Department
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<Department> resultDepartments = sqlConnection.Query<Department>(sql, dynamicParameters).ToList();

                        inventoryTransactions = inventoryTransactions.Join(resultDepartments, x => x.DepartmentNo, y => y.DepartmentNo, (x, y) => { x.DepartmentId = y.DepartmentId; return x; }).ToList();
                        #endregion

                        #region //撈取確認者ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserId, a.UserNo 
                                FROM BAS.[User] a";

                        List<User> resultConfirmUsers = sqlConnection.Query<User>(sql, dynamicParameters).ToList();

                        inventoryTransactions = inventoryTransactions.GroupJoin(resultConfirmUsers, x => x.ConfirmUserNo, y => y.UserNo, (x, y) => { x.ConfirmUserId = y.FirstOrDefault()?.UserId; return x; }).ToList();
                        #endregion

                        #region //撈取品號ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT MtlItemId, MtlItemNo
                                FROM PDM.MtlItem
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<MtlItem> resultMtlitems = sqlConnection.Query<MtlItem>(sql, dynamicParameters).ToList();

                        itDetails = itDetails.Join(resultMtlitems, x => x.MtlItemNo, y => y.MtlItemNo, (x, y) => { x.MtlItemId = y.MtlItemId; return x; }).Select(x => x).ToList();
                        #endregion

                        #region //撈取庫別ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT InventoryId, InventoryNo
                                FROM SCM.Inventory
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<Inventory> resultInventories = sqlConnection.Query<Inventory>(sql, dynamicParameters).ToList();

                        itDetails = itDetails.Join(resultInventories, x => x.InventoryNo, y => y.InventoryNo, (x, y) => { x.InventoryId = y.InventoryId; return x; }).ToList();
                        #endregion

                        #region //判斷SCM.InventoryTransaction是否有資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ItId, ItErpPrefix, ItErpNo
                                FROM SCM.InventoryTransaction
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<InventoryTransaction> resultInvTran = sqlConnection.Query<InventoryTransaction>(sql, dynamicParameters).ToList();

                        inventoryTransactions = inventoryTransactions.GroupJoin(resultInvTran, x => new { x.ItErpPrefix, x.ItErpNo }, y => new { y.ItErpPrefix, y.ItErpNo }, (x, y) => { x.ItId = y.FirstOrDefault()?.ItId; return x; }).Select(x => x).ToList();
                        #endregion

                        #region //庫存異動單頭(新增/修改)
                        List<InventoryTransaction> addInvTran = inventoryTransactions.Where(x => x.ItId == null).ToList();
                        List<InventoryTransaction> updateInvTran = inventoryTransactions.Where(x => x.ItId != null).ToList();

                        #region //新增
                        dynamicParameters = new DynamicParameters();
                        if (addInvTran.Count > 0)
                        {
                            addInvTran
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

                            sql = @"INSERT INTO SCM.InventoryTransaction (CompanyId, ItErpPrefix, ItErpNo, ItDate
                                    , DocDate, DepartmentId, Remark, TotalQty, Amount
                                    , ConfirmStatus, ConfirmUserId, TransferStatus, TransferDate
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@CompanyId, @ItErpPrefix, @ItErpNo, @ItDate
                                    , @DocDate, @DepartmentId, @Remark, @TotalQty, @Amount
                                    , @ConfirmStatus, @ConfirmUserId, @TransferStatus, @TransferDate
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            rowsAffected += sqlConnection.Execute(sql, addInvTran);
                        }
                        #endregion

                        #region //修改
                        dynamicParameters = new DynamicParameters();
                        if (updateInvTran.Count > 0)
                        {
                            updateInvTran
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.LastModifiedDate = LastModifiedDate;
                                    x.LastModifiedBy = LastModifiedBy;
                                });

                            sql = @"UPDATE SCM.InventoryTransaction SET
                                    ItErpPrefix = @ItErpPrefix,
                                    ItErpNo = @ItErpNo,
                                    ItDate = @ItDate,
                                    DocDate = @DocDate,
                                    DepartmentId = @DepartmentId,
                                    Remark = @Remark,
                                    TotalQty = @TotalQty,
                                    Amount = @Amount,
                                    ConfirmStatus = @ConfirmStatus,
                                    ConfirmUserId = @ConfirmUserId,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE ItId = @ItId";
                            rowsAffected += sqlConnection.Execute(sql, updateInvTran);
                        }
                        #endregion
                        #endregion

                        #region //庫存異動單身(新增/修改)
                        #region //撈取庫存異動單頭ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ItId, ItErpPrefix, ItErpNo
                                FROM  SCM.InventoryTransaction
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        resultInvTran = sqlConnection.Query<InventoryTransaction>(sql, dynamicParameters).ToList();

                        itDetails = itDetails.Join(resultInvTran, x => new { x.ItErpPrefix, x.ItErpNo }, y => new { y.ItErpPrefix, y.ItErpNo }, (x, y) => { x.ItId = y.ItId; return x; }).Select(x => x).ToList();
                        #endregion

                        #region //判斷SCM.ItDetail是否有資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ItDetailId, a.ItId, a.ItSequence
                                FROM SCM.ItDetail a
                                INNER JOIN SCM.InventoryTransaction b ON a.ItId = b.ItId
                                WHERE b.CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<ItDetail> resultItDetail = sqlConnection.Query<ItDetail>(sql, dynamicParameters).ToList();

                        itDetails = itDetails.GroupJoin(resultItDetail, x => new { x.ItId, x.ItSequence }, y => new { y.ItId, y.ItSequence }, (x, y) => { x.ItDetailId = y.FirstOrDefault()?.ItDetailId; return x; }).ToList();
                        #endregion

                        List<ItDetail> addItDetail = itDetails.Where(x => x.ItDetailId == null).ToList();
                        List<ItDetail> updateItDetail = itDetails.Where(x => x.ItDetailId != null).ToList();

                        #region //新增
                        dynamicParameters = new DynamicParameters();
                        if (addItDetail.Count > 0)
                        {
                            addItDetail
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.CreateDate = CreateDate;
                                    x.LastModifiedDate = LastModifiedDate;
                                    x.CreateBy = CreateBy;
                                    x.LastModifiedBy = LastModifiedBy;
                                });

                            sql = @"INSERT INTO SCM.ItDetail (ItId, ItSequence, MtlItemId, ItMtlItemName
                                    , ItMtlItemSpec, ItQty, InvQty, UnitCost, Amount, InventoryId, ItRemark
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@ItId, @ItSequence, @MtlItemId, @ItMtlItemName, @ItMtlItemSpec
                                    , @ItQty, @InvQty, @UnitCost, @Amount, @InventoryId, @ItRemark
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            rowsAffected += sqlConnection.Execute(sql, addItDetail);
                        }
                        #endregion

                        #region //修改
                        dynamicParameters = new DynamicParameters();
                        if (updateItDetail.Count > 0)
                        {
                            updateItDetail
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.LastModifiedDate = LastModifiedDate;
                                    x.LastModifiedBy = LastModifiedBy;
                                });

                            sql = @"UPDATE SCM.ItDetail SET
                                    MtlItemId = @MtlItemId,
                                    ItMtlItemName = @ItMtlItemName,
                                    ItMtlItemSpec = @ItMtlItemSpec,
                                    ItQty = @ItQty,
                                    InvQty = @InvQty,
                                    UnitCost = @UnitCost,
                                    Amount = @Amount,
                                    InventoryId = @InventoryId,
                                    ItRemark = @ItRemark,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE ItDetailId = @ItDetailId";
                            rowsAffected += sqlConnection.Execute(sql, updateItDetail);
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

        #region //UpdateInventoryTransactionManualSynchronize -- 庫存異動單資料手動同步 -- Ben Ma 2023.05.22
        public string UpdateInventoryTransactionManualSynchronize(string ItErpFullNo, string SyncStartDate, string SyncEndDate
            , string NormalSync, string TranSync)
        {
            try
            {
                List<InventoryTransaction> inventoryTransactions = new List<InventoryTransaction>();
                List<ItDetail> itDetails = new List<ItDetail>();

                int mainAffected = 0, detailAffected = 0, mainDelAffected = 0, detailDelAffected = 0; ;
                string companyNo = "";

                using (TransactionScope transactionScope = new TransactionScope(TransactionScopeOption.Required, new TimeSpan(0, 180, 0)))
                {
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
                    }

                    #region //正常同步
                    if (NormalSync == "Y")
                    {
                        using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //判斷ERP庫存異動單資料是否存在
                            if (ItErpFullNo.Length > 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM INVTA
                                        WHERE (LTRIM(RTRIM(TA001)) + '-' + LTRIM(RTRIM(TA002))) LIKE '%' + @ItErpFullNo + '%'";
                                dynamicParameters.Add("ItErpFullNo", ItErpFullNo);

                                var resultTsn = sqlConnection.Query(sql, dynamicParameters);
                                if (resultTsn.Count() <= 0) throw new SystemException("ERP庫存異動單資料錯誤");
                            }
                            #endregion

                            #region //撈取ERP庫存異動資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(TA001)) ItErpPrefix, LTRIM(RTRIM(TA002)) ItErpNo
                                    , CASE WHEN LEN(LTRIM(RTRIM(TA003))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TA003)) as date), 'yyyy-MM-dd') ELSE NULL END ItDate
                                    , CASE WHEN LEN(LTRIM(RTRIM(TA014))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TA014)) as date), 'yyyy-MM-dd') ELSE NULL END DocDate
                                    , LTRIM(RTRIM(TA004)) DepartmentNo, LTRIM(RTRIM(TA005)) Remark
                                    , TA011 TotalQty, TA012 Amount
                                    , LTRIM(RTRIM(TA006)) ConfirmStatus, LTRIM(RTRIM(TA015)) ConfirmUserNo
                                    , CASE WHEN LEN(LTRIM(RTRIM(CREATE_DATE))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_DATE)) as date), 'yyyy-MM-dd') ELSE NULL END TransferDate
                                    , CASE WHEN LEN(LTRIM(RTRIM(CREATE_TIME))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_TIME)) as datetime), 'HH:mm:ss') ELSE NULL END TransferTime
                                    FROM INVTA
                                    WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ItErpFullNo", @" AND (LTRIM(RTRIM(TA001)) + '-' + LTRIM(RTRIM(TA002))) LIKE '%' + @ItErpFullNo + '%'", ItErpFullNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SyncStartDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME >= @SyncStartDate OR MODI_DATE + ' ' + MODI_TIME >= @SyncStartDate)", SyncStartDate.Length > 0 ? Convert.ToDateTime(SyncStartDate).ToString("yyyyMMdd 00:00:00") : "");
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SyncEndDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME <= @SyncEndDate OR MODI_DATE + ' ' + MODI_TIME <= @SyncEndDate)", SyncEndDate.Length > 0 ? Convert.ToDateTime(SyncEndDate).ToString("yyyyMMdd 23:59:59") : "");
                            sql += @" ORDER BY TransferDate, TransferTime";

                            inventoryTransactions = sqlConnection.Query<InventoryTransaction>(sql, dynamicParameters).ToList();
                            #endregion

                            #region //撈取ERP庫存異動單身資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(TB001)) ItErpPrefix, LTRIM(RTRIM(TB002)) ItErpNo, LTRIM(RTRIM(TB003)) ItSequence
                                    , LTRIM(RTRIM(TB004)) MtlItemNo, LTRIM(RTRIM(TB005)) ItMtlItemName, LTRIM(RTRIM(TB006)) ItMtlItemSpec
                                    , LTRIM(RTRIM(CAST(TB007 AS decimal(16,3)))) ItQty, LTRIM(RTRIM(TB008)) UomNo, LTRIM(RTRIM(CAST(TB009 AS decimal(16,3)))) InvQty
                                    , LTRIM(RTRIM(CAST(TB010 AS decimal(21,6)))) UnitCost, LTRIM(RTRIM(CAST(TB011 AS decimal(21,6)))) Amount
                                    , LTRIM(RTRIM(TB012)) InventoryNo, LTRIM(RTRIM(TB013)) ToInventoryNo, LTRIM(RTRIM(TB017)) ItRemark
                                    FROM INVTB
                                    WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ItErpFullNo", @" AND (LTRIM(RTRIM(TB001)) + '-' + LTRIM(RTRIM(TB002))) LIKE '%' + @ItErpFullNo + '%'", ItErpFullNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SyncStartDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME >= @SyncStartDate OR MODI_DATE + ' ' + MODI_TIME >= @SyncStartDate)", SyncStartDate.Length > 0 ? Convert.ToDateTime(SyncStartDate).ToString("yyyyMMdd 00:00:00") : "");
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SyncEndDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME <= @SyncEndDate OR MODI_DATE + ' ' + MODI_TIME <= @SyncEndDate)", SyncEndDate.Length > 0 ? Convert.ToDateTime(SyncEndDate).ToString("yyyyMMdd 23:59:59") : "");

                            itDetails = sqlConnection.Query<ItDetail>(sql, dynamicParameters).ToList();
                            #endregion
                        }

                        using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                        {
                            #region //撈取部門ID
                            sql = @"SELECT DepartmentId, DepartmentNo
                                    FROM BAS.Department
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<Department> resultDepartments = sqlConnection.Query<Department>(sql, dynamicParameters).ToList();

                            inventoryTransactions = inventoryTransactions.Join(resultDepartments, x => x.DepartmentNo, y => y.DepartmentNo, (x, y) => { x.DepartmentId = y.DepartmentId; return x; }).ToList();
                            #endregion

                            #region //撈取確認者ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.UserId, a.UserNo 
                                    FROM BAS.[User] a";

                            List<User> resultConfirmUsers = sqlConnection.Query<User>(sql, dynamicParameters).ToList();

                            inventoryTransactions = inventoryTransactions.GroupJoin(resultConfirmUsers, x => x.ConfirmUserNo, y => y.UserNo, (x, y) => { x.ConfirmUserId = y.FirstOrDefault()?.UserId; return x; }).ToList();
                            #endregion

                            #region //撈取品號ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT MtlItemId, MtlItemNo
                                    FROM PDM.MtlItem
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<MtlItem> resultMtlitems = sqlConnection.Query<MtlItem>(sql, dynamicParameters).ToList();

                            itDetails = itDetails.Join(resultMtlitems, x => x.MtlItemNo, y => y.MtlItemNo, (x, y) => { x.MtlItemId = y.MtlItemId; return x; }).Select(x => x).ToList();
                            #endregion

                            #region //撈取庫別ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT InventoryId, InventoryNo
                                    FROM SCM.Inventory
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<Inventory> resultInventories = sqlConnection.Query<Inventory>(sql, dynamicParameters).ToList();

                            itDetails = itDetails.Join(resultInventories, x => x.InventoryNo, y => y.InventoryNo, (x, y) => { x.InventoryId = y.InventoryId; return x; }).ToList();
                            #endregion

                            #region //若為轉撥單，額外找轉入庫ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT InventoryId, InventoryNo
                                    FROM SCM.Inventory
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<Inventory> resultInventories2 = sqlConnection.Query<Inventory>(sql, dynamicParameters).ToList();

                            itDetails = itDetails.GroupJoin(resultInventories2, x => x.ToInventoryNo, y => y.InventoryNo, (x, y) => { x.ToInventoryId = y.FirstOrDefault()?.InventoryId; return x; }).ToList();
                            #endregion

                            #region //撈取單位ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT UomId, UomNo
                                    FROM PDM.UnitOfMeasure
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<Uom> resultUom = sqlConnection.Query<Uom>(sql, dynamicParameters).ToList();

                            itDetails = itDetails.Join(resultUom, x => x.UomNo, y => y.UomNo, (x, y) => { x.UomId = y.UomId; return x; }).ToList();
                            #endregion

                            #region //判斷SCM.InventoryTransaction是否有資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT ItId, ItErpPrefix, ItErpNo
                                    FROM SCM.InventoryTransaction
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<InventoryTransaction> resultInvTran = sqlConnection.Query<InventoryTransaction>(sql, dynamicParameters).ToList();

                            inventoryTransactions = inventoryTransactions.GroupJoin(resultInvTran, x => new { x.ItErpPrefix, x.ItErpNo }, y => new { y.ItErpPrefix, y.ItErpNo }, (x, y) => { x.ItId = y.FirstOrDefault()?.ItId; return x; }).Select(x => x).ToList();
                            #endregion

                            #region //庫存異動單頭(新增/修改)
                            List<InventoryTransaction> addInvTran = inventoryTransactions.Where(x => x.ItId == null).ToList();
                            List<InventoryTransaction> updateInvTran = inventoryTransactions.Where(x => x.ItId != null).ToList();

                            #region //新增
                            dynamicParameters = new DynamicParameters();
                            if (addInvTran.Count > 0)
                            {
                                addInvTran
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.CompanyId = CurrentCompany;
                                        x.TransferStatus = "Y";
                                        x.CreateDate = CreateDate;
                                        x.LastModifiedDate = LastModifiedDate;
                                        x.CreateBy = CreateBy;
                                        x.LastModifiedBy = LastModifiedBy;
                                    });

                                sql = @"INSERT INTO SCM.InventoryTransaction (CompanyId, ItErpPrefix, ItErpNo, ItDate
                                        , DocDate, DepartmentId, Remark, TotalQty, Amount
                                        , ConfirmStatus, ConfirmUserId, TransferStatus, TransferDate
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        VALUES (@CompanyId, @ItErpPrefix, @ItErpNo, @ItDate
                                        , @DocDate, @DepartmentId, @Remark, @TotalQty, @Amount
                                        , @ConfirmStatus, @ConfirmUserId, @TransferStatus, @TransferDate
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                int addMain = sqlConnection.Execute(sql, addInvTran);
                                mainAffected += addMain;
                            }
                            #endregion

                            #region //修改
                            dynamicParameters = new DynamicParameters();
                            if (updateInvTran.Count > 0)
                            {
                                updateInvTran
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.LastModifiedDate = LastModifiedDate;
                                        x.LastModifiedBy = LastModifiedBy;
                                    });

                                sql = @"UPDATE SCM.InventoryTransaction SET
                                        ItErpPrefix = @ItErpPrefix,
                                        ItErpNo = @ItErpNo,
                                        ItDate = @ItDate,
                                        DocDate = @DocDate,
                                        DepartmentId = @DepartmentId,
                                        Remark = @Remark,
                                        TotalQty = @TotalQty,
                                        Amount = @Amount,
                                        ConfirmStatus = @ConfirmStatus,
                                        ConfirmUserId = @ConfirmUserId,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE ItId = @ItId";
                                int updateMain = sqlConnection.Execute(sql, updateInvTran);
                                mainAffected += updateMain;
                            }
                            #endregion
                            #endregion

                            #region //庫存異動單身(新增/修改)
                            #region //撈取庫存異動單頭ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT ItId, ItErpPrefix, ItErpNo
                                    FROM SCM.InventoryTransaction
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            resultInvTran = sqlConnection.Query<InventoryTransaction>(sql, dynamicParameters).ToList();

                            itDetails = itDetails.Join(resultInvTran, x => new { x.ItErpPrefix, x.ItErpNo }, y => new { y.ItErpPrefix, y.ItErpNo }, (x, y) => { x.ItId = y.ItId; return x; }).Select(x => x).ToList();
                            #endregion

                            #region //判斷SCM.ItDetail是否有資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.ItDetailId, a.ItId, a.ItSequence
                                    FROM SCM.ItDetail a
                                    INNER JOIN SCM.InventoryTransaction b ON a.ItId = b.ItId
                                    WHERE b.CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<ItDetail> resultItDetail = sqlConnection.Query<ItDetail>(sql, dynamicParameters).ToList();

                            itDetails = itDetails.GroupJoin(resultItDetail, x => new { x.ItId, x.ItSequence }, y => new { y.ItId, y.ItSequence }, (x, y) => { x.ItDetailId = y.FirstOrDefault()?.ItDetailId; return x; }).ToList();
                            #endregion

                            List<ItDetail> addItDetail = itDetails.Where(x => x.ItDetailId == null).ToList();
                            List<ItDetail> updateItDetail = itDetails.Where(x => x.ItDetailId != null).ToList();

                            #region //新增
                            dynamicParameters = new DynamicParameters();
                            if (addItDetail.Count > 0)
                            {
                                addItDetail
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.CreateDate = CreateDate;
                                        x.LastModifiedDate = LastModifiedDate;
                                        x.CreateBy = CreateBy;
                                        x.LastModifiedBy = LastModifiedBy;
                                    });

                                sql = @"INSERT INTO SCM.ItDetail (ItId, ItSequence, MtlItemId, ItMtlItemName
                                        , ItMtlItemSpec, ItQty, UomId, InvQty, UnitCost, Amount, InventoryId, ToInventoryId, ItRemark
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        VALUES (@ItId, @ItSequence, @MtlItemId, @ItMtlItemName, @ItMtlItemSpec
                                        , @ItQty, @UomId, @InvQty, @UnitCost, @Amount, @InventoryId, @ToInventoryId, @ItRemark
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                int addDetail = sqlConnection.Execute(sql, addItDetail);
                                detailAffected += addDetail;
                            }
                            #endregion

                            #region //修改
                            dynamicParameters = new DynamicParameters();
                            if (updateItDetail.Count > 0)
                            {
                                updateItDetail
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.LastModifiedDate = LastModifiedDate;
                                        x.LastModifiedBy = LastModifiedBy;
                                    });

                                sql = @"UPDATE SCM.ItDetail SET
                                        MtlItemId = @MtlItemId,
                                        ItMtlItemName = @ItMtlItemName,
                                        ItMtlItemSpec = @ItMtlItemSpec,
                                        ItQty = @ItQty,
                                        UomId = @UomId,
                                        InvQty = @InvQty,
                                        UnitCost = @UnitCost,
                                        Amount = @Amount,
                                        InventoryId = @InventoryId,
                                        ToInventoryId = @ToInventoryId,
                                        ItRemark = @ItRemark,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE ItDetailId = @ItDetailId";
                                int updateDetail = sqlConnection.Execute(sql, updateItDetail);
                                detailAffected += updateDetail;
                            }
                            #endregion
                            #endregion
                        }
                    }
                    #endregion

                    #region //異動同步
                    if (TranSync == "Y")
                    {
                        using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //ERP庫存異動單單頭資料
                            sql = @"SELECT TA001 ItErpPrefix, TA002 ItErpNo
                                    FROM INVTA
                                    WHERE 1=1
                                    ORDER BY TA001, TA002";
                            var resultErpIt = erpConnection.Query<InventoryTransaction>(sql, dynamicParameters);
                            #endregion

                            using (SqlConnection bmConnection = new SqlConnection(MainConnectionStrings))
                            {
                                #region //BM庫存異動單單頭資料
                                sql = @"SELECT ItErpPrefix, ItErpNo
                                        FROM SCM.InventoryTransaction
                                        WHERE CompanyId = @CompanyId
                                        ORDER BY ItErpPrefix, ItErpNo";
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                var resultBmIt = bmConnection.Query<InventoryTransaction>(sql, dynamicParameters);
                                #endregion

                                #region //篩選ERP不存在但BM還存在的庫存異動單單頭
                                var dictionaryErpIt = resultErpIt.ToDictionary(x => x.ItErpPrefix + "-" + x.ItErpNo, x => x.ItErpPrefix + "-" + x.ItErpNo);
                                var dictionaryBmIt = resultBmIt.ToDictionary(x => x.ItErpPrefix + "-" + x.ItErpNo, x => x.ItErpPrefix + "-" + x.ItErpNo);
                                var changeIt = dictionaryBmIt.Where(x => !dictionaryErpIt.ContainsKey(x.Key)).ToList();
                                var changeItList = changeIt.Select(x => x.Value).ToList();
                                #endregion

                                #region //刪除異動庫存異動單單頭
                                if (changeItList.Count > 0)
                                {
                                    #region //單身刪除
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE SCM.ItDetail
                                            WHERE ItId IN (
                                                SELECT ItId
                                                FROM SCM.InventoryTransaction
                                                WHERE ItErpPrefix + '-' + ItErpNo IN @ItErpFullNo
                                            )";
                                    dynamicParameters.Add("ItErpFullNo", changeItList.Select(x => x).ToArray());
                                    mainDelAffected += bmConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //單頭刪除
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE SCM.InventoryTransaction
                                            WHERE ItErpPrefix + '-' + ItErpNo IN @ItErpFullNo";
                                    dynamicParameters.Add("ItErpFullNo", changeItList.Select(x => x).ToArray());
                                    detailDelAffected += bmConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                                #endregion
                            }

                            #region //ERP庫存異動單單身資料
                            sql = @"SELECT TB001 ItErpPrefix, TB002 ItErpNo, RIGHT('0000' + TB003, 4) ItSequence
                                    FROM INVTB
                                    WHERE 1=1
                                    ORDER BY TB001, TB002, TB003";
                            var resultErpItDetail = erpConnection.Query<ItDetail>(sql, dynamicParameters);
                            #endregion

                            using (SqlConnection bmConnection = new SqlConnection(MainConnectionStrings))
                            {
                                #region //BM庫存異動單單身資料
                                sql = @"SELECT b.ItErpPrefix, b.ItErpNo, RIGHT('0000' + a.ItSequence, 4) ItSequence
                                        FROM SCM.ItDetail a
                                        INNER JOIN SCM.InventoryTransaction b ON a.ItId = b.ItId
                                        WHERE b.CompanyId = @CompanyId
                                        ORDER BY b.ItErpPrefix, b.ItErpNo, a.ItSequence";
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                var resultBmItDetail = bmConnection.Query<ItDetail>(sql, dynamicParameters);
                                #endregion

                                #region //篩選ERP不存在但BM還存在的庫存異動單單身
                                var dictionaryErpItDetail = resultErpItDetail.ToDictionary(x => x.ItErpPrefix + "-" + x.ItErpNo + "-" + x.ItSequence, x => x.ItErpPrefix + "-" + x.ItErpNo + "-" + x.ItSequence);
                                var dictionaryBmItDetail = resultBmItDetail.ToDictionary(x => x.ItErpPrefix + "-" + x.ItErpNo + "-" + x.ItSequence, x => x.ItErpPrefix + "-" + x.ItErpNo + "-" + x.ItSequence);
                                var changeItDetail = dictionaryBmItDetail.Where(x => !dictionaryErpItDetail.ContainsKey(x.Key)).ToList();
                                var changeItDetailList = changeItDetail.Select(x => x.Value).ToList();
                                #endregion

                                #region //刪除異動暫出單單身
                                if (changeItDetailList.Count > 0)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE a
                                            FROM SCM.ItDetail a
                                            INNER JOIN SCM.InventoryTransaction b ON a.ItId = b.ItId
                                            WHERE b.ItErpPrefix + '-' + b.ItErpNo + '-' + RIGHT('0000' + a.ItSequence, 4) IN @ItErpFullNo";
                                    dynamicParameters.Add("ItErpFullNo", changeItDetailList.Select(x => x).ToArray());
                                    detailDelAffected += bmConnection.Execute(sql, dynamicParameters, null, 300);
                                }
                                #endregion
                            }
                        }
                    }
                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "(" + mainAffected + detailAffected + " rows affected)",
                        data = "已更新資料<br />【" + mainAffected + "】筆單頭<br />【" + detailAffected + "】筆單身<br />刪除<br />【" + mainDelAffected + "】筆單頭<br />【" + detailDelAffected + "】筆單身"
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

        #region //UpdateTempShippingNoteSynchronize -- 暫出單資料同步 -- Ben Ma 2023.03.28
        public string UpdateTempShippingNoteSynchronize(string CompanyNo, string UpdateDate)
        {
            try
            {
                //暫出單
                List<TempShippingNote> tempShippingNotes = new List<TempShippingNote>();
                List<TsnDetail> tsnDetails = new List<TsnDetail>();

                //暫出歸還單
                List<TempShippingReturnNote> tempShippingReturnNotes = new List<TempShippingReturnNote>();
                List<TsrnDetail> tsrnDetails = new List<TsrnDetail>();

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

                        var resultCompany = sqlConnection.Query(sql, dynamicParameters);
                        if (resultCompany.Count() <= 0) throw new SystemException("【公司】資料錯誤!");

                        foreach (var item in resultCompany)
                        {
                            CompanyId = Convert.ToInt32(item.CompanyId);
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        }
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        #region //撈取ERP暫出單資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT LTRIM(RTRIM(TF001)) TsnErpPrefix, LTRIM(RTRIM(TF002)) TsnErpNo
                                , CASE WHEN LEN(LTRIM(RTRIM(TF003))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TF003)) as date), 'yyyy-MM-dd') ELSE NULL END TsnDate
                                , CASE WHEN LEN(LTRIM(RTRIM(TF024))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TF024)) as date), 'yyyy-MM-dd') ELSE NULL END DocDate
                                , LTRIM(RTRIM(TF004)) ToObject, LTRIM(RTRIM(TF005)) ObjectOther, LTRIM(RTRIM(TF007)) DepartmentNo
                                , LTRIM(RTRIM(TF008)) UserNo, LTRIM(RTRIM(TF014)) Remark, LTRIM(RTRIM(TF018)) OtherRemark
                                , TF022 TotalQty, TF023 Amount, TF027 TaxAmount
                                , LTRIM(RTRIM(TF020)) ConfirmStatus, LTRIM(RTRIM(TF025)) ConfirmUserNo
                                , CASE WHEN LEN(LTRIM(RTRIM(CREATE_DATE))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_DATE)) as date), 'yyyy-MM-dd') ELSE NULL END TransferDate
                                , CASE WHEN LEN(LTRIM(RTRIM(CREATE_TIME))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_TIME)) as datetime), 'HH:mm:ss') ELSE NULL END TransferTime
                                FROM INVTF
                                WHERE 1=1";
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "UpdateDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME > @UpdateDate OR MODI_DATE + ' ' + MODI_TIME > @UpdateDate)", UpdateDate);
                        sql += @" ORDER BY TransferDate, TransferTime";

                        tempShippingNotes = sqlConnection.Query<TempShippingNote>(sql, dynamicParameters).ToList();
                        #endregion

                        #region //撈取ERP暫出單單身資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT LTRIM(RTRIM(TG001)) TsnErpPrefix, LTRIM(RTRIM(TG002)) TsnErpNo, LTRIM(RTRIM(TG003)) TsnSequence
                                , LTRIM(RTRIM(TG004)) MtlItemNo, LTRIM(RTRIM(TG005)) TsnMtlItemName, LTRIM(RTRIM(TG006)) TsnMtlItemSpec
                                , LTRIM(RTRIM(TG007)) TsnOutInventoryNo, LTRIM(RTRIM(TG008)) TsnInInventoryNo
                                , LTRIM(RTRIM(CAST(TG009 AS decimal(16,3)))) TsnQty, LTRIM(RTRIM(CAST(TG012 AS decimal(21,6)))) UnitPrice
                                , LTRIM(RTRIM(CAST(TG013 AS decimal(21,6)))) Amount
                                , LTRIM(RTRIM(TG043)) ProductType, LTRIM(RTRIM(CAST(TG044 AS INT))) FreebieOrSpareQty, LTRIM(RTRIM(CAST(TG052 AS INT))) TsnPriceQty
                                , LTRIM(RTRIM(TG053)) TsnPriceUomrNo, LTRIM(RTRIM(TG042)) BusinessTaxRate
                                , CASE WHEN LEN(LTRIM(RTRIM(TG027))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TG027)) as date), 'yyyy-MM-dd') ELSE NULL END ExpectedReturnDate
                                , LTRIM(RTRIM(TG014)) SoErpPrefix, LTRIM(RTRIM(TG015)) SoErpNo, LTRIM(RTRIM(TG016)) SoSequence
                                , LTRIM(RTRIM(TG019)) TsnRemark, LTRIM(RTRIM(TG022)) ConfirmStatus, LTRIM(RTRIM(TG024)) ClosureStatus
                                , LTRIM(RTRIM(CAST(TG020 AS decimal(16,3)))) SaleQty, LTRIM(RTRIM(CAST(TG046 AS decimal(16,3)))) SaleFSQty
                                , LTRIM(RTRIM(CAST(TG021 AS decimal(16,3)))) ReturnQty, LTRIM(RTRIM(CAST(TG048 AS decimal(16,3)))) ReturnFSQty
                                FROM INVTG
                                WHERE 1=1";
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "UpdateDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME > @UpdateDate OR MODI_DATE + ' ' + MODI_TIME > @UpdateDate)", UpdateDate);

                        tsnDetails = sqlConnection.Query<TsnDetail>(sql, dynamicParameters).ToList();
                        #endregion

                        #region //撈取ERP暫出歸還單資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT LTRIM(RTRIM(TH001)) TsrnErpPrefix, LTRIM(RTRIM(TH002)) TsrnErpNo
                                , CASE WHEN LEN(LTRIM(RTRIM(TH003))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TH003)) as date), 'yyyy-MM-dd') ELSE NULL END TsrnDate
                                , CASE WHEN LEN(LTRIM(RTRIM(TH023))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TH023)) as date), 'yyyy-MM-dd') ELSE NULL END DocDate
                                , LTRIM(RTRIM(TH004)) ToObject, LTRIM(RTRIM(TH005)) ObjectOther, LTRIM(RTRIM(TH007)) DepartmentNo
                                , LTRIM(RTRIM(TH008)) UserNo, LTRIM(RTRIM(TH014)) Remark, LTRIM(RTRIM(TH018)) Remark
                                , TH021 TotalQty, TH022 Amount, TH026 TaxAmount
                                , LTRIM(RTRIM(TH020)) ConfirmStatus, LTRIM(RTRIM(TH024)) ConfirmUserNo
                                , CASE WHEN LEN(LTRIM(RTRIM(CREATE_DATE))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_DATE)) as date), 'yyyy-MM-dd') ELSE NULL END TransferDate
                                , CASE WHEN LEN(LTRIM(RTRIM(CREATE_TIME))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_TIME)) as datetime), 'HH:mm:ss') ELSE NULL END TransferTime
                                FROM INVTH
                                WHERE 1=1";
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "UpdateDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME > @UpdateDate OR MODI_DATE + ' ' + MODI_TIME > @UpdateDate)", UpdateDate);
                        sql += @" ORDER BY TransferDate, TransferTime";

                        tempShippingReturnNotes = sqlConnection.Query<TempShippingReturnNote>(sql, dynamicParameters).ToList();
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

                        List<Department> resultDepartments = sqlConnection.Query<Department>(sql, dynamicParameters).ToList();

                        tempShippingNotes = tempShippingNotes.Join(resultDepartments, x => x.DepartmentNo, y => y.DepartmentNo, (x, y) => { x.DepartmentId = y.DepartmentId; return x; }).ToList();
                        #endregion

                        #region //撈取人員ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserId, a.UserNo
                                FROM BAS.[User] a";

                        List<User> resultUsers = sqlConnection.Query<User>(sql, dynamicParameters).ToList();

                        tempShippingNotes = tempShippingNotes.Join(resultUsers, x => x.UserNo, y => y.UserNo, (x, y) => { x.UserId = y.UserId; return x; }).ToList();
                        tempShippingNotes = tempShippingNotes.GroupJoin(resultUsers, x => x.ConfirmUserNo, y => y.UserNo, (x, y) => { x.ConfirmUserId = y.FirstOrDefault()?.UserId; return x; }).ToList();
                        List<TempShippingNote> toObjectUser = tempShippingNotes.Where(x => x.ToObject == "3").Join(resultUsers, x => x.ObjectOther, y => y.UserNo, (x, y) => { x.ObjectUser = y.UserId; return x; }).Select(x => x).ToList();
                        #endregion

                        #region //撈取異動對象ID
                        #region //撈取客戶ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT CustomerId, CustomerNo 
                                FROM SCM.Customer
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<Customer> resultCustomers = sqlConnection.Query<Customer>(sql, dynamicParameters).ToList();

                        List<TempShippingNote> toObjectCustomer = tempShippingNotes.Where(x => x.ToObject == "1").Join(resultCustomers, x => x.ObjectOther, y => y.CustomerNo, (x, y) => { x.ObjectCustomer = y.CustomerId; return x; }).Select(x => x).ToList();
                        #endregion

                        #region //撈取供應商ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT SupplierId, SupplierNo 
                                FROM SCM.Supplier
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<Supplier> resultSuppliers = sqlConnection.Query<Supplier>(sql, dynamicParameters).ToList();

                        List<TempShippingNote> toObjectSupplier = tempShippingNotes.Where(x => x.ToObject == "2").Join(resultSuppliers, x => x.ObjectOther, y => y.SupplierNo, (x, y) => { x.ObjectSupplier = y.SupplierId; return x; }).Select(x => x).ToList();
                        #endregion

                        #region //其他
                        List<TempShippingNote> toObjectOther = tempShippingNotes.Where(x => x.ToObject == "9").Select(x => x).ToList();
                        #endregion

                        tempShippingNotes = toObjectCustomer.Union(toObjectSupplier).Union(toObjectUser).Union(toObjectOther).ToList();
                        #endregion

                        #region //撈取品號ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT MtlItemId, MtlItemNo
                                FROM PDM.MtlItem
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<MtlItem> resultMtlitems = sqlConnection.Query<MtlItem>(sql, dynamicParameters).ToList();

                        tsnDetails = tsnDetails.Join(resultMtlitems, x => x.MtlItemNo, y => y.MtlItemNo, (x, y) => { x.MtlItemId = y.MtlItemId; return x; }).Select(x => x).ToList();
                        #endregion

                        #region //撈取庫別ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT InventoryId, InventoryNo
                                FROM SCM.Inventory
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<Inventory> resultInventories = sqlConnection.Query<Inventory>(sql, dynamicParameters).ToList();

                        tsnDetails = tsnDetails.Join(resultInventories, x => x.TsnOutInventoryNo, y => y.InventoryNo, (x, y) => { x.TsnOutInventory = y.InventoryId; return x; }).ToList();
                        tsnDetails = tsnDetails.Join(resultInventories, x => x.TsnInInventoryNo, y => y.InventoryNo, (x, y) => { x.TsnInInventory = y.InventoryId; return x; }).ToList();
                        #endregion

                        #region //撈取訂單單身ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.SoId, a.SoErpPrefix, a.SoErpNo, b.SoDetailId, b.SoSequence
                                FROM SCM.SaleOrder a
                                LEFT JOIN SCM.SoDetail b ON b.SoId = a.SoId
                                WHERE a.CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<SoDetail> resultSoDetails = sqlConnection.Query<SoDetail>(sql, dynamicParameters).ToList();

                        tsnDetails = tsnDetails.GroupJoin(resultSoDetails, x => new { x.SoErpPrefix, x.SoErpNo, x.SoSequence }, y => new { y.SoErpPrefix, y.SoErpNo, y.SoSequence }, (x, y) => { x.SoDetailId = y.FirstOrDefault()?.SoDetailId; return x; }).Select(x => x).ToList();
                        #endregion

                        #region //判斷SCM.TempShippingNote是否有資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TsnId, TsnErpPrefix, TsnErpNo 
                                FROM SCM.TempShippingNote
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<TempShippingNote> resultTempShippingNotes = sqlConnection.Query<TempShippingNote>(sql, dynamicParameters).ToList();

                        tempShippingNotes = tempShippingNotes.GroupJoin(resultTempShippingNotes, x => new { x.TsnErpPrefix, x.TsnErpNo }, y => new { y.TsnErpPrefix, y.TsnErpNo }, (x, y) => { x.TsnId = y.FirstOrDefault()?.TsnId; return x; }).Select(x => x).ToList();
                        #endregion

                        #region //暫出單單頭(新增/修改)
                        List<TempShippingNote> addTempShippingNotes = tempShippingNotes.Where(x => x.TsnId == null).ToList();
                        List<TempShippingNote> updateTempShippingNotes = tempShippingNotes.Where(x => x.TsnId != null).ToList();

                        #region //新增
                        dynamicParameters = new DynamicParameters();
                        if (addTempShippingNotes.Count > 0)
                        {
                            addTempShippingNotes
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

                            sql = @"INSERT INTO SCM.TempShippingNote (CompanyId, TsnErpPrefix, TsnErpNo, TsnDate
                                    , DocDate, ToObject, ObjectCustomer, ObjectSupplier, ObjectUser, ObjectOther
                                    , DepartmentId, UserId, Remark, OtherRemark, TotalQty, Amount, TaxAmount
                                    , ConfirmStatus, ConfirmUserId, TransferStatus, TransferDate
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@CompanyId, @TsnErpPrefix, @TsnErpNo, @TsnDate
                                    , @DocDate, @ToObject, @ObjectCustomer, @ObjectSupplier, @ObjectUser, @ObjectOther
                                    , @DepartmentId, @UserId, @Remark, @OtherRemark, @TotalQty, @Amount, @TaxAmount
                                    , @ConfirmStatus, @ConfirmUserId, @TransferStatus, @TransferDate
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            rowsAffected += sqlConnection.Execute(sql, addTempShippingNotes);
                        }
                        #endregion

                        #region //修改
                        dynamicParameters = new DynamicParameters();
                        if (updateTempShippingNotes.Count > 0)
                        {
                            updateTempShippingNotes
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.LastModifiedDate = LastModifiedDate;
                                    x.LastModifiedBy = LastModifiedBy;
                                });

                            sql = @"UPDATE SCM.TempShippingNote SET
                                    TsnErpPrefix = @TsnErpPrefix,
                                    TsnErpNo = @TsnErpNo,
                                    TsnDate = @TsnDate,
                                    DocDate = @DocDate,
                                    ToObject = @ToObject,
                                    ObjectCustomer = @ObjectCustomer,
                                    ObjectSupplier = @ObjectSupplier,
                                    ObjectUser = @ObjectUser,
                                    ObjectOther = @ObjectOther,
                                    DepartmentId = @DepartmentId,
                                    UserId = @UserId,
                                    Remark = @Remark,
                                    OtherRemark = @OtherRemark,
                                    TotalQty = @TotalQty,
                                    Amount = @Amount,
                                    TaxAmount = @TaxAmount,
                                    ConfirmStatus = @ConfirmStatus,
                                    ConfirmUserId = @ConfirmUserId,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE TsnId = @TsnId";
                            rowsAffected += sqlConnection.Execute(sql, updateTempShippingNotes);
                        }
                        #endregion
                        #endregion

                        #region //暫出單單身(新增/修改)
                        #region //撈取暫出單單頭ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TsnId, TsnErpPrefix, TsnErpNo
                                FROM  SCM.TempShippingNote
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        resultTempShippingNotes = sqlConnection.Query<TempShippingNote>(sql, dynamicParameters).ToList();

                        tsnDetails = tsnDetails.Join(resultTempShippingNotes, x => new { x.TsnErpPrefix, x.TsnErpNo }, y => new { y.TsnErpPrefix, y.TsnErpNo }, (x, y) => { x.TsnId = y.TsnId; return x; }).Select(x => x).ToList();
                        #endregion

                        #region //判斷SCM.TsnDetail是否有資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.TsnDetailId, a.TsnId, a.TsnSequence
                                FROM SCM.TsnDetail a
                                INNER JOIN SCM.TempShippingNote b ON a.TsnId = b.TsnId
                                WHERE b.CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<TsnDetail> resultItDetail = sqlConnection.Query<TsnDetail>(sql, dynamicParameters).ToList();

                        tsnDetails = tsnDetails.GroupJoin(resultItDetail, x => new { x.TsnId, x.TsnSequence }, y => new { y.TsnId, y.TsnSequence }, (x, y) => { x.TsnDetailId = y.FirstOrDefault()?.TsnDetailId; return x; }).ToList();
                        #endregion

                        List<TsnDetail> addTsnDetail = tsnDetails.Where(x => x.TsnDetailId == null).ToList();
                        List<TsnDetail> updateTsnDetail = tsnDetails.Where(x => x.TsnDetailId != null).ToList();

                        #region //新增
                        if (addTsnDetail.Count > 0)
                        {
                            addTsnDetail
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.CreateDate = CreateDate;
                                    x.LastModifiedDate = LastModifiedDate;
                                    x.CreateBy = CreateBy;
                                    x.LastModifiedBy = LastModifiedBy;
                                });

                            sql = @"INSERT INTO SCM.TsnDetail (TsnId, TsnSequence, MtlItemId, TsnMtlItemName
                                    , TsnMtlItemSpec, TsnOutInventory, TsnInInventory, TsnQty, UnitPrice
                                    , ProductType, FreebieOrSpareQty, TsnPriceQty, TsnPriceUomId, BusinessTaxRate
                                    , ExpectedReturnDate, Amount, SoDetailId, TsnRemark, ConfirmStatus, ClosureStatus
                                    , SaleQty, SaleFSQty, ReturnQty, ReturnFSQty
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@TsnId, @TsnSequence, @MtlItemId, @TsnMtlItemName
                                    , @TsnMtlItemSpec, @TsnOutInventory, @TsnInInventory, @TsnQty, @UnitPrice
                                    , @ProductType, @FreebieOrSpareQty, @TsnPriceQty, @TsnPriceUomId, @BusinessTaxRate
                                    , @ExpectedReturnDate, @Amount, @SoDetailId, @TsnRemark, @ConfirmStatus, @ClosureStatus
                                    , @SaleQty, @SaleFSQty, @ReturnQty, @ReturnFSQty
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            rowsAffected += sqlConnection.Execute(sql, addTsnDetail);
                        }
                        #endregion

                        #region //修改
                        if (updateTsnDetail.Count > 0)
                        {
                            updateTsnDetail
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.LastModifiedDate = LastModifiedDate;
                                    x.LastModifiedBy = LastModifiedBy;
                                });

                            sql = @"UPDATE SCM.TsnDetail SET
                                    TsnId = @TsnId,
                                    TsnSequence = @TsnSequence,
                                    MtlItemId = @MtlItemId,
                                    TsnMtlItemName = @TsnMtlItemName,
                                    TsnMtlItemSpec = @TsnMtlItemSpec,
                                    TsnOutInventory = @TsnOutInventory,
                                    TsnInInventory = @TsnInInventory,
                                    TsnQty = @TsnQty,
                                    UnitPrice = @UnitPrice,
                                    ProductType = @ProductType,
                                    FreebieOrSpareQty = @FreebieOrSpareQty,
                                    TsnPriceQty = @TsnPriceQty,
                                    TsnPriceUomId = @TsnPriceUomId,
                                    BusinessTaxRate = @BusinessTaxRate,
                                    ExpectedReturnDate = @ExpectedReturnDate,
                                    Amount = @Amount,
                                    SoDetailId = @SoDetailId,
                                    TsnRemark = @TsnRemark,
                                    ConfirmStatus = @ConfirmStatus,
                                    ClosureStatus = @ClosureStatus,
                                    SaleQty = @SaleQty,
                                    SaleFSQty = @SaleFSQty,
                                    ReturnQty = @ReturnQty,
                                    ReturnFSQty = @ReturnFSQty,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE TsnDetailId = @TsnDetailId";
                            rowsAffected += sqlConnection.Execute(sql, updateTsnDetail);
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

        #region //UpdateTempShippingNoteManualSynchronize -- 暫出單資料手動同步 -- Ben Ma 2023.05.16
        public string UpdateTempShippingNoteManualSynchronize(string TsnErpFullNo, string SyncStartDate, string SyncEndDate
            , string NormalSync, string TranSync)
        {
            try
            {
                List<TempShippingNote> tempShippingNotes = new List<TempShippingNote>();
                List<TsnDetail> tsnDetails = new List<TsnDetail>();

                int mainAffected = 0, detailAffected = 0, mainDelAffected = 0, detailDelAffected = 0; ;
                string companyNo = "";

                using (TransactionScope transactionScope = new TransactionScope(TransactionScopeOption.Required, new TimeSpan(0, 180, 0)))
                {
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
                    }

                    #region //正常同步
                    if (NormalSync == "Y")
                    {
                        using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //判斷ERP暫出單資料是否存在
                            if (TsnErpFullNo.Length > 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM INVTF
                                        WHERE (LTRIM(RTRIM(TF001)) + '-' + LTRIM(RTRIM(TF002))) LIKE '%' + @TsnErpFullNo + '%'";
                                dynamicParameters.Add("TsnErpFullNo", TsnErpFullNo);

                                var resultTsn = sqlConnection.Query(sql, dynamicParameters);
                                if (resultTsn.Count() <= 0) throw new SystemException("ERP暫出單資料錯誤");
                            }
                            #endregion

                            #region //撈取ERP暫出單資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(TF001)) TsnErpPrefix, LTRIM(RTRIM(TF002)) TsnErpNo
                                    , CASE WHEN LEN(LTRIM(RTRIM(TF003))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TF003)) as date), 'yyyy-MM-dd') ELSE NULL END TsnDate
                                    , CASE WHEN LEN(LTRIM(RTRIM(TF024))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TF024)) as date), 'yyyy-MM-dd') ELSE NULL END DocDate
                                    , LTRIM(RTRIM(TF004)) ToObject, LTRIM(RTRIM(TF005)) ObjectOther, LTRIM(RTRIM(TF007)) DepartmentNo
                                    , LTRIM(RTRIM(TF008)) UserNo, LTRIM(RTRIM(TF014)) Remark, LTRIM(RTRIM(TF018)) OtherRemark
                                    , TF022 TotalQty, TF023 Amount, TF027 TaxAmount
                                    , LTRIM(RTRIM(TF020)) ConfirmStatus, LTRIM(RTRIM(TF025)) ConfirmUserNo
                                    , CASE WHEN LEN(LTRIM(RTRIM(CREATE_DATE))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_DATE)) as date), 'yyyy-MM-dd') ELSE NULL END TransferDate
                                    , CASE WHEN LEN(LTRIM(RTRIM(CREATE_TIME))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_TIME)) as datetime), 'HH:mm:ss') ELSE NULL END TransferTime
                                    FROM INVTF
                                    WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TsnErpFullNo", @" AND (LTRIM(RTRIM(TF001)) + '-' + LTRIM(RTRIM(TF002))) LIKE '%' + @TsnErpFullNo + '%'", TsnErpFullNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SyncStartDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME >= @SyncStartDate OR MODI_DATE + ' ' + MODI_TIME >= @SyncStartDate)", SyncStartDate.Length > 0 ? Convert.ToDateTime(SyncStartDate).ToString("yyyyMMdd 00:00:00") : "");
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SyncEndDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME <= @SyncEndDate OR MODI_DATE + ' ' + MODI_TIME <= @SyncEndDate)", SyncEndDate.Length > 0 ? Convert.ToDateTime(SyncEndDate).ToString("yyyyMMdd 23:59:59") : "");
                            sql += @" ORDER BY TransferDate, TransferTime";

                            tempShippingNotes = sqlConnection.Query<TempShippingNote>(sql, dynamicParameters).ToList();
                            #endregion

                            #region //撈取ERP暫出單單身資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(TG001)) TsnErpPrefix, LTRIM(RTRIM(TG002)) TsnErpNo, LTRIM(RTRIM(TG003)) TsnSequence
                                    , LTRIM(RTRIM(TG004)) MtlItemNo, LTRIM(RTRIM(TG005)) TsnMtlItemName, LTRIM(RTRIM(TG006)) TsnMtlItemSpec
                                    , LTRIM(RTRIM(TG007)) TsnOutInventoryNo, LTRIM(RTRIM(TG008)) TsnInInventoryNo
                                    , LTRIM(RTRIM(CAST(TG009 AS decimal(16,3)))) TsnQty, LTRIM(RTRIM(CAST(TG012 AS decimal(21,6)))) UnitPrice
                                    , LTRIM(RTRIM(CAST(TG013 AS decimal(21,6)))) Amount
                                    , LTRIM(RTRIM(TG043)) ProductType, LTRIM(RTRIM(CAST(TG044 AS INT))) FreebieOrSpareQty, LTRIM(RTRIM(CAST(TG052 AS INT))) TsnPriceQty
                                    , LTRIM(RTRIM(TG053)) TsnPriceUomrNo, LTRIM(RTRIM(TG042)) BusinessTaxRate
                                    , CASE WHEN LEN(LTRIM(RTRIM(TG027))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TG027)) as date), 'yyyy-MM-dd') ELSE NULL END ExpectedReturnDate
                                    , LTRIM(RTRIM(TG014)) SoErpPrefix, LTRIM(RTRIM(TG015)) SoErpNo, LTRIM(RTRIM(TG016)) SoSequence
                                    , LTRIM(RTRIM(TG019)) TsnRemark, LTRIM(RTRIM(TG022)) ConfirmStatus, LTRIM(RTRIM(TG024)) ClosureStatus
                                    , LTRIM(RTRIM(CAST(TG020 AS decimal(16,3)))) SaleQty, LTRIM(RTRIM(CAST(TG021 AS decimal(16,3)))) ReturnQty
                                    FROM INVTG
                                    WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TsnErpFullNo", @" AND (LTRIM(RTRIM(TG001)) + '-' + LTRIM(RTRIM(TG002))) LIKE '%' + @TsnErpFullNo + '%'", TsnErpFullNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SyncStartDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME >= @SyncStartDate OR MODI_DATE + ' ' + MODI_TIME >= @SyncStartDate)", SyncStartDate.Length > 0 ? Convert.ToDateTime(SyncStartDate).ToString("yyyyMMdd 00:00:00") : "");
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SyncEndDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME <= @SyncEndDate OR MODI_DATE + ' ' + MODI_TIME <= @SyncEndDate)", SyncEndDate.Length > 0 ? Convert.ToDateTime(SyncEndDate).ToString("yyyyMMdd 23:59:59") : "");

                            tsnDetails = sqlConnection.Query<TsnDetail>(sql, dynamicParameters).ToList();
                            #endregion
                        }

                        using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                        {
                            #region //撈取部門ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT DepartmentId, DepartmentNo
                                    FROM BAS.Department
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<Department> resultDepartments = sqlConnection.Query<Department>(sql, dynamicParameters).ToList();

                            tempShippingNotes = tempShippingNotes.Join(resultDepartments, x => x.DepartmentNo, y => y.DepartmentNo, (x, y) => { x.DepartmentId = y.DepartmentId; return x; }).ToList();
                            #endregion

                            #region //撈取人員ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.UserId, a.UserNo
                                    FROM BAS.[User] a";

                            List<User> resultUsers = sqlConnection.Query<User>(sql, dynamicParameters).ToList();

                            tempShippingNotes = tempShippingNotes.Join(resultUsers, x => x.UserNo, y => y.UserNo, (x, y) => { x.UserId = y.UserId; return x; }).ToList();
                            tempShippingNotes = tempShippingNotes.GroupJoin(resultUsers, x => x.ConfirmUserNo, y => y.UserNo, (x, y) => { x.ConfirmUserId = y.FirstOrDefault()?.UserId; return x; }).ToList();
                            List<TempShippingNote> toObjectUser = tempShippingNotes.Where(x => x.ToObject == "3").Join(resultUsers, x => x.ObjectOther, y => y.UserNo, (x, y) => { x.ObjectUser = y.UserId; return x; }).Select(x => x).ToList();
                            #endregion

                            #region //撈取異動對象ID
                            #region //撈取客戶ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT CustomerId, CustomerNo 
                                    FROM SCM.Customer
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<Customer> resultCustomers = sqlConnection.Query<Customer>(sql, dynamicParameters).ToList();

                            List<TempShippingNote> toObjectCustomer = tempShippingNotes.Where(x => x.ToObject == "1").Join(resultCustomers, x => x.ObjectOther, y => y.CustomerNo, (x, y) => { x.ObjectCustomer = y.CustomerId; return x; }).Select(x => x).ToList();
                            #endregion

                            #region //撈取供應商ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT SupplierId, SupplierNo 
                                    FROM SCM.Supplier
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<Supplier> resultSuppliers = sqlConnection.Query<Supplier>(sql, dynamicParameters).ToList();

                            List<TempShippingNote> toObjectSupplier = tempShippingNotes.Where(x => x.ToObject == "2").Join(resultSuppliers, x => x.ObjectOther, y => y.SupplierNo, (x, y) => { x.ObjectSupplier = y.SupplierId; return x; }).Select(x => x).ToList();
                            #endregion

                            #region //其他
                            List<TempShippingNote> toObjectOther = tempShippingNotes.Where(x => x.ToObject == "9").Select(x => x).ToList();
                            #endregion

                            tempShippingNotes = toObjectCustomer.Union(toObjectSupplier).Union(toObjectUser).Union(toObjectOther).ToList();
                            #endregion

                            #region //撈取品號ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT MtlItemId, MtlItemNo
                                    FROM PDM.MtlItem
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<MtlItem> resultMtlitems = sqlConnection.Query<MtlItem>(sql, dynamicParameters).ToList();

                            tsnDetails = tsnDetails.Join(resultMtlitems, x => x.MtlItemNo, y => y.MtlItemNo, (x, y) => { x.MtlItemId = y.MtlItemId; return x; }).Select(x => x).ToList();
                            #endregion

                            #region //撈取庫別ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT InventoryId, InventoryNo
                                    FROM SCM.Inventory
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<Inventory> resultInventories = sqlConnection.Query<Inventory>(sql, dynamicParameters).ToList();

                            tsnDetails = tsnDetails.Join(resultInventories, x => x.TsnOutInventoryNo, y => y.InventoryNo, (x, y) => { x.TsnOutInventory = y.InventoryId; return x; }).ToList();
                            tsnDetails = tsnDetails.Join(resultInventories, x => x.TsnInInventoryNo, y => y.InventoryNo, (x, y) => { x.TsnInInventory = y.InventoryId; return x; }).ToList();
                            #endregion

                            #region //撈取訂單單身ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.SoId, a.SoErpPrefix, a.SoErpNo, b.SoDetailId, b.SoSequence
                                    FROM SCM.SaleOrder a
                                    LEFT JOIN SCM.SoDetail b ON b.SoId = a.SoId
                                    WHERE a.CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<SoDetail> resultSoDetails = sqlConnection.Query<SoDetail>(sql, dynamicParameters).ToList();

                            tsnDetails = tsnDetails.GroupJoin(resultSoDetails, x => new { x.SoErpPrefix, x.SoErpNo, x.SoSequence }, y => new { y.SoErpPrefix, y.SoErpNo, y.SoSequence }, (x, y) => { x.SoDetailId = y.FirstOrDefault()?.SoDetailId; return x; }).Select(x => x).ToList();
                            #endregion

                            #region //判斷SCM.TempShippingNote是否有資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TsnId, TsnErpPrefix, TsnErpNo 
                                    FROM SCM.TempShippingNote
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<TempShippingNote> resultTempShippingNotes = sqlConnection.Query<TempShippingNote>(sql, dynamicParameters).ToList();

                            tempShippingNotes = tempShippingNotes.GroupJoin(resultTempShippingNotes, x => new { x.TsnErpPrefix, x.TsnErpNo }, y => new { y.TsnErpPrefix, y.TsnErpNo }, (x, y) => { x.TsnId = y.FirstOrDefault()?.TsnId; return x; }).Select(x => x).ToList();
                            #endregion

                            #region //暫出單單頭(新增/修改)
                            List<TempShippingNote> addTempShippingNotes = tempShippingNotes.Where(x => x.TsnId == null).ToList();
                            List<TempShippingNote> updateTempShippingNotes = tempShippingNotes.Where(x => x.TsnId != null).ToList();

                            #region //新增
                            dynamicParameters = new DynamicParameters();
                            if (addTempShippingNotes.Count > 0)
                            {
                                addTempShippingNotes
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.CompanyId = CurrentCompany;
                                        x.TransferStatus = "Y";
                                        x.CreateDate = CreateDate;
                                        x.LastModifiedDate = LastModifiedDate;
                                        x.CreateBy = CreateBy;
                                        x.LastModifiedBy = LastModifiedBy;
                                    });

                                sql = @"INSERT INTO SCM.TempShippingNote (CompanyId, TsnErpPrefix, TsnErpNo, TsnDate
                                        , DocDate, ToObject, ObjectCustomer, ObjectSupplier, ObjectUser, ObjectOther
                                        , DepartmentId, UserId, Remark, OtherRemark, TotalQty, Amount, TaxAmount
                                        , ConfirmStatus, ConfirmUserId, TransferStatus, TransferDate
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        VALUES (@CompanyId, @TsnErpPrefix, @TsnErpNo, @TsnDate
                                        , @DocDate, @ToObject, @ObjectCustomer, @ObjectSupplier, @ObjectUser, @ObjectOther
                                        , @DepartmentId, @UserId, @Remark, @OtherRemark, @TotalQty, @Amount, @TaxAmount
                                        , @ConfirmStatus, @ConfirmUserId, @TransferStatus, @TransferDate
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                int addMain = sqlConnection.Execute(sql, addTempShippingNotes);
                                mainAffected += addMain;
                            }
                            #endregion

                            #region //修改
                            dynamicParameters = new DynamicParameters();
                            if (updateTempShippingNotes.Count > 0)
                            {
                                updateTempShippingNotes
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.LastModifiedDate = LastModifiedDate;
                                        x.LastModifiedBy = LastModifiedBy;
                                    });

                                sql = @"UPDATE SCM.TempShippingNote SET
                                        TsnErpPrefix = @TsnErpPrefix,
                                        TsnErpNo = @TsnErpNo,
                                        TsnDate = @TsnDate,
                                        DocDate = @DocDate,
                                        ToObject = @ToObject,
                                        ObjectCustomer = @ObjectCustomer,
                                        ObjectSupplier = @ObjectSupplier,
                                        ObjectUser = @ObjectUser,
                                        ObjectOther = @ObjectOther,
                                        DepartmentId = @DepartmentId,
                                        UserId = @UserId,
                                        Remark = @Remark,
                                        OtherRemark = @OtherRemark,
                                        TotalQty = @TotalQty,
                                        Amount = @Amount,
                                        TaxAmount = @TaxAmount,
                                        ConfirmStatus = @ConfirmStatus,
                                        ConfirmUserId = @ConfirmUserId,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE TsnId = @TsnId";
                                int updateMain = sqlConnection.Execute(sql, updateTempShippingNotes);
                                mainAffected += updateMain;
                            }
                            #endregion
                            #endregion

                            #region //暫出單單身(新增/修改)
                            #region //撈取暫出單單頭ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TsnId, TsnErpPrefix, TsnErpNo
                                    FROM SCM.TempShippingNote
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            resultTempShippingNotes = sqlConnection.Query<TempShippingNote>(sql, dynamicParameters).ToList();

                            tsnDetails = tsnDetails.Join(resultTempShippingNotes, x => new { x.TsnErpPrefix, x.TsnErpNo }, y => new { y.TsnErpPrefix, y.TsnErpNo }, (x, y) => { x.TsnId = y.TsnId; return x; }).Select(x => x).ToList();
                            #endregion

                            #region //判斷SCM.TsnDetail是否有資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.TsnDetailId, a.TsnId, a.TsnSequence
                                    FROM SCM.TsnDetail a
                                    INNER JOIN SCM.TempShippingNote b ON a.TsnId = b.TsnId
                                    WHERE b.CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<TsnDetail> resultItDetail = sqlConnection.Query<TsnDetail>(sql, dynamicParameters).ToList();

                            tsnDetails = tsnDetails.GroupJoin(resultItDetail, x => new { x.TsnId, x.TsnSequence }, y => new { y.TsnId, y.TsnSequence }, (x, y) => { x.TsnDetailId = y.FirstOrDefault()?.TsnDetailId; return x; }).ToList();
                            #endregion

                            List<TsnDetail> addTsnDetail = tsnDetails.Where(x => x.TsnDetailId == null).ToList();
                            List<TsnDetail> updateTsnDetail = tsnDetails.Where(x => x.TsnDetailId != null).ToList();

                            #region //新增
                            if (addTsnDetail.Count > 0)
                            {
                                addTsnDetail
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.CreateDate = CreateDate;
                                        x.LastModifiedDate = LastModifiedDate;
                                        x.CreateBy = CreateBy;
                                        x.LastModifiedBy = LastModifiedBy;
                                    });

                                sql = @"INSERT INTO SCM.TsnDetail (TsnId, TsnSequence, MtlItemId, TsnMtlItemName
                                        , TsnMtlItemSpec, TsnOutInventory, TsnInInventory, TsnQty, UnitPrice
                                        , ProductType, FreebieOrSpareQty, TsnPriceQty, TsnPriceUomId, BusinessTaxRate
                                        , ExpectedReturnDate, SaleFSQty, ReturnFSQty
                                        , Amount, SoDetailId, TsnRemark, ConfirmStatus, ClosureStatus
                                        , SaleQty, ReturnQty
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        VALUES (@TsnId, @TsnSequence, @MtlItemId, @TsnMtlItemName
                                        , @TsnMtlItemSpec, @TsnOutInventory, @TsnInInventory, @TsnQty, @UnitPrice
                                        , @ProductType, @FreebieOrSpareQty, @TsnPriceQty, @TsnPriceUomId, @BusinessTaxRate
                                        , @ExpectedReturnDate, @SaleFSQty, @ReturnFSQty
                                        , @Amount, @SoDetailId, @TsnRemark, @ConfirmStatus, @ClosureStatus
                                        , @SaleQty, @ReturnQty
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                int addDetail = sqlConnection.Execute(sql, addTsnDetail);
                                detailAffected += addDetail;
                            }
                            #endregion

                            #region //修改
                            if (updateTsnDetail.Count > 0)
                            {
                                updateTsnDetail
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.LastModifiedDate = LastModifiedDate;
                                        x.LastModifiedBy = LastModifiedBy;
                                    });

                                sql = @"UPDATE SCM.TsnDetail SET
                                        TsnId = @TsnId,
                                        TsnSequence = @TsnSequence,
                                        MtlItemId = @MtlItemId,
                                        TsnMtlItemName = @TsnMtlItemName,
                                        TsnMtlItemSpec = @TsnMtlItemSpec,
                                        TsnOutInventory = @TsnOutInventory,
                                        TsnInInventory = @TsnInInventory,
                                        TsnQty = @TsnQty,
                                        UnitPrice = @UnitPrice,
                                        ProductType = @ProductType,
                                        FreebieOrSpareQty = @FreebieOrSpareQty,
                                        TsnPriceQty = @TsnPriceQty,
                                        TsnPriceUomId = @TsnPriceUomId,
                                        BusinessTaxRate = @BusinessTaxRate,
                                        ExpectedReturnDate = @ExpectedReturnDate,
                                        SaleFSQty = @SaleFSQty,
                                        ReturnFSQty = @ReturnFSQty,
                                        Amount = @Amount,
                                        SoDetailId = @SoDetailId,
                                        TsnRemark = @TsnRemark,
                                        ConfirmStatus = @ConfirmStatus,
                                        ClosureStatus = @ClosureStatus,
                                        SaleQty = @SaleQty,
                                        ReturnQty = @ReturnQty,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE TsnDetailId = @TsnDetailId";
                                int updateDetail = sqlConnection.Execute(sql, updateTsnDetail);
                                detailAffected += updateDetail;
                            }
                            #endregion
                            #endregion
                        }
                    }
                    #endregion
                    
                    #region //異動同步
                    if (TranSync == "Y")
                    {
                        using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //ERP暫出單單頭資料
                            sql = @"SELECT TF001 TsnErpPrefix, TF002 TsnErpNo
                                    FROM INVTF
                                    WHERE 1=1
                                    ORDER BY TF001, TF002";
                            var resultErpTsn = erpConnection.Query<TempShippingNote>(sql, dynamicParameters);
                            #endregion

                            using (SqlConnection bmConnection = new SqlConnection(MainConnectionStrings))
                            {
                                #region //BM暫出單單頭資料
                                sql = @"SELECT TsnErpPrefix, TsnErpNo
                                        FROM SCM.TempShippingNote
                                        WHERE CompanyId = @CompanyId
                                        ORDER BY TsnErpPrefix, TsnErpNo";
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                var resultBmTsn = bmConnection.Query<TempShippingNote>(sql, dynamicParameters);
                                #endregion

                                #region //篩選ERP不存在但BM還存在的暫出單單頭
                                var dictionaryErpTsn = resultErpTsn.ToDictionary(x => x.TsnErpPrefix + "-" + x.TsnErpNo, x => x.TsnErpPrefix + "-" + x.TsnErpNo);
                                var dictionaryBmTsn = resultBmTsn.ToDictionary(x => x.TsnErpPrefix + "-" + x.TsnErpNo, x => x.TsnErpPrefix + "-" + x.TsnErpNo);
                                var changeTsn = dictionaryBmTsn.Where(x => !dictionaryErpTsn.ContainsKey(x.Key)).ToList();
                                var changeTsnList = changeTsn.Select(x => x.Value).ToList();
                                #endregion

                                #region //刪除異動暫出單單頭
                                if (changeTsnList.Count > 0)
                                {
                                    #region //單身刪除
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE SCM.TsnDetail
                                            WHERE TsnId IN (
                                                SELECT TsnId
                                                FROM SCM.TempShippingNote
                                                WHERE TsnErpPrefix + '-' + TsnErpNo IN @TsnErpFullNo
                                            )";
                                    dynamicParameters.Add("TsnErpFullNo", changeTsnList.Select(x => x).ToArray());
                                    mainDelAffected += bmConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //單頭刪除
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE SCM.TempShippingNote
                                            WHERE TsnErpPrefix + '-' + TsnErpNo IN @TsnErpFullNo";
                                    dynamicParameters.Add("TsnErpFullNo", changeTsnList.Select(x => x).ToArray());
                                    detailDelAffected += bmConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                                #endregion
                            }

                            #region //ERP暫出單單身資料
                            sql = @"SELECT TG001 TsnErpPrefix, TG002 TsnErpNo, TG003 TsnSequence
                                    FROM INVTG
                                    WHERE 1=1
                                    ORDER BY TG001, TG002, TG003";
                            var resultErpTsnDetail = erpConnection.Query<TsnDetail>(sql, dynamicParameters);
                            #endregion

                            using (SqlConnection bmConnection = new SqlConnection(MainConnectionStrings))
                            {
                                #region //BM暫出單單身資料
                                sql = @"SELECT b.TsnErpPrefix, b.TsnErpNo, a.TsnSequence
                                        FROM SCM.TsnDetail a
                                        INNER JOIN SCM.TempShippingNote b ON a.TsnId = b.TsnId
                                        WHERE b.CompanyId = @CompanyId
                                        ORDER BY b.TsnErpPrefix, b.TsnErpNo, a.TsnSequence";
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                var resultBmTsnDetail = bmConnection.Query<TsnDetail>(sql, dynamicParameters);
                                #endregion

                                #region //篩選ERP不存在但BM還存在的暫出單單身
                                var dictionaryErpTsnDetail = resultErpTsnDetail.ToDictionary(x => x.TsnErpPrefix + "-" + x.TsnErpNo + "-" + x.TsnSequence, x => x.TsnErpPrefix + "-" + x.TsnErpNo + "-" + x.TsnSequence);
                                var dictionaryBmTsnDetail = resultBmTsnDetail.ToDictionary(x => x.TsnErpPrefix + "-" + x.TsnErpNo + "-" + x.TsnSequence, x => x.TsnErpPrefix + "-" + x.TsnErpNo + "-" + x.TsnSequence);
                                var changeTsnDetail = dictionaryBmTsnDetail.Where(x => !dictionaryErpTsnDetail.ContainsKey(x.Key)).ToList();
                                var changeTsnDetailList = changeTsnDetail.Select(x => x.Value).ToList();
                                #endregion

                                #region //刪除異動暫出單單身
                                if (changeTsnDetailList.Count > 0)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE a
                                            FROM SCM.TsnDetail a
                                            INNER JOIN SCM.TempShippingNote b ON a.TsnId = b.TsnId
                                            WHERE b.TsnErpPrefix + '-' + b.TsnErpNo + '-' + a.TsnSequence IN @TsnErpFullNo";
                                    dynamicParameters.Add("TsnErpFullNo", changeTsnDetailList.Select(x => x).ToArray());
                                    detailDelAffected += bmConnection.Execute(sql, dynamicParameters);
                                }
                                #endregion
                            }
                        }
                    }
                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "(" + mainAffected + detailAffected + " rows affected)",
                        data = "已更新資料<br />【" + mainAffected + "】筆單頭<br />【" + detailAffected + "】筆單身<br />刪除<br />【" + mainDelAffected + "】筆單頭<br />【" + detailDelAffected + "】筆單身"
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

        #region //UpdateTempShippingReturnNoteSynchronize -- 暫出歸還單資料同步 -- Ben Ma 2023.03.29
        public string UpdateTempShippingReturnNoteSynchronize(string CompanyNo, string UpdateDate)
        {
            try
            {
                List<TempShippingReturnNote> tempShippingReturnNotes = new List<TempShippingReturnNote>();
                List<TsrnDetail> tsrnDetails = new List<TsrnDetail>();

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

                        var resultCompany = sqlConnection.Query(sql, dynamicParameters);
                        if (resultCompany.Count() <= 0) throw new SystemException("【公司】資料錯誤!");

                        foreach (var item in resultCompany)
                        {
                            CompanyId = Convert.ToInt32(item.CompanyId);
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        }
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        #region //撈取ERP暫出歸還單資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT LTRIM(RTRIM(TH001)) TsrnErpPrefix, LTRIM(RTRIM(TH002)) TsrnErpNo
                                , CASE WHEN LEN(LTRIM(RTRIM(TH003))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TH003)) as date), 'yyyy-MM-dd') ELSE NULL END TsrnDate
                                , CASE WHEN LEN(LTRIM(RTRIM(TH023))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TH023)) as date), 'yyyy-MM-dd') ELSE NULL END DocDate
                                , LTRIM(RTRIM(TH004)) ToObject, LTRIM(RTRIM(TH005)) ObjectOther, LTRIM(RTRIM(TH007)) DepartmentNo
                                , LTRIM(RTRIM(TH008)) UserNo, LTRIM(RTRIM(TH014)) Remark, LTRIM(RTRIM(TH018)) Remark
                                , TH021 TotalQty, TH022 Amount, TH026 TaxAmount
                                , LTRIM(RTRIM(TH020)) ConfirmStatus, LTRIM(RTRIM(TH024)) ConfirmUserNo
                                , CASE WHEN LEN(LTRIM(RTRIM(CREATE_DATE))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_DATE)) as date), 'yyyy-MM-dd') ELSE NULL END TransferDate
                                , CASE WHEN LEN(LTRIM(RTRIM(CREATE_TIME))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_TIME)) as datetime), 'HH:mm:ss') ELSE NULL END TransferTime
                                FROM INVTH
                                WHERE 1=1";
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "UpdateDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME > @UpdateDate OR MODI_DATE + ' ' + MODI_TIME > @UpdateDate)", UpdateDate);
                        sql += @" ORDER BY TransferDate, TransferTime";

                        tempShippingReturnNotes = sqlConnection.Query<TempShippingReturnNote>(sql, dynamicParameters).ToList();
                        #endregion

                        #region //撈取ERP暫出歸還單單身資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT LTRIM(RTRIM(TI001)) TsrnErpPrefix, LTRIM(RTRIM(TI002)) TsrnErpNo, LTRIM(RTRIM(TI003)) TsrnSequence
                                , LTRIM(RTRIM(TI004)) MtlItemNo, LTRIM(RTRIM(TI005)) TsrnMtlItemName, LTRIM(RTRIM(TI006)) TsrnMtlItemSpec
                                , LTRIM(RTRIM(TI007)) TsrnOutInventoryNo, LTRIM(RTRIM(TI008)) TsrnInInventoryNo
                                , LTRIM(RTRIM(CAST(TI009 AS decimal(16,3)))) TsrnQty, LTRIM(RTRIM(CAST(TI012 AS decimal(21,6)))) UnitPrice
                                , LTRIM(RTRIM(CAST(TI013 AS decimal(21,6)))) Amount
                                , LTRIM(RTRIM(TI034)) ProductType, LTRIM(RTRIM(CAST(TI035 AS INT))) FreebieOrSpareQty, LTRIM(RTRIM(CAST(TI038 AS INT))) TsrnPriceQty
                                , LTRIM(RTRIM(TI039)) TsrnPriceUomrNo, LTRIM(RTRIM(TI033)) BusinessTaxRate
                                , LTRIM(RTRIM(TI014)) TsnErpPrefix, LTRIM(RTRIM(TI015)) TsnErpNo, LTRIM(RTRIM(TI016)) TsnSequence
                                , LTRIM(RTRIM(TI021)) TsrnRemark, LTRIM(RTRIM(TI022)) ConfirmStatus
                                FROM INVTI
                                WHERE 1=1";
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "UpdateDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME > @UpdateDate OR MODI_DATE + ' ' + MODI_TIME > @UpdateDate)", UpdateDate);

                        tsrnDetails = sqlConnection.Query<TsrnDetail>(sql, dynamicParameters).ToList();
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

                        List<Department> resultDepartments = sqlConnection.Query<Department>(sql, dynamicParameters).ToList();

                        tempShippingReturnNotes = tempShippingReturnNotes.Join(resultDepartments, x => x.DepartmentNo, y => y.DepartmentNo, (x, y) => { x.DepartmentId = y.DepartmentId; return x; }).ToList();
                        #endregion

                        #region //撈取人員ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserId, a.UserNo
                                FROM BAS.[User] a";

                        List<User> resultUsers = sqlConnection.Query<User>(sql, dynamicParameters).ToList();

                        tempShippingReturnNotes = tempShippingReturnNotes.Join(resultUsers, x => x.UserNo, y => y.UserNo, (x, y) => { x.UserId = y.UserId; return x; }).ToList();
                        tempShippingReturnNotes = tempShippingReturnNotes.GroupJoin(resultUsers, x => x.ConfirmUserNo, y => y.UserNo, (x, y) => { x.ConfirmUserId = y.FirstOrDefault()?.UserId; return x; }).ToList();
                        List<TempShippingReturnNote> toObjectUser = tempShippingReturnNotes.Where(x => x.ToObject == "3").Join(resultUsers, x => x.ObjectOther, y => y.UserNo, (x, y) => { x.ObjectUser = y.UserId; return x; }).Select(x => x).ToList();
                        #endregion

                        #region //撈取異動對象ID
                        #region //撈取客戶ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT CustomerId, CustomerNo 
                                FROM SCM.Customer
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<Customer> resultCustomers = sqlConnection.Query<Customer>(sql, dynamicParameters).ToList();

                        List<TempShippingReturnNote> toObjectCustomer = tempShippingReturnNotes.Where(x => x.ToObject == "1").Join(resultCustomers, x => x.ObjectOther, y => y.CustomerNo, (x, y) => { x.ObjectCustomer = y.CustomerId; return x; }).Select(x => x).ToList();
                        #endregion

                        #region //撈取供應商ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT SupplierId, SupplierNo 
                                FROM SCM.Supplier
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<Supplier> resultSuppliers = sqlConnection.Query<Supplier>(sql, dynamicParameters).ToList();

                        List<TempShippingReturnNote> toObjectSupplier = tempShippingReturnNotes.Where(x => x.ToObject == "2").Join(resultSuppliers, x => x.ObjectOther, y => y.SupplierNo, (x, y) => { x.ObjectSupplier = y.SupplierId; return x; }).Select(x => x).ToList();
                        #endregion

                        #region //其他
                        List<TempShippingReturnNote> toObjectOther = tempShippingReturnNotes.Where(x => x.ToObject == "9").Select(x => x).ToList();
                        #endregion

                        tempShippingReturnNotes = toObjectCustomer.Union(toObjectSupplier).Union(toObjectUser).Union(toObjectOther).ToList();
                        #endregion

                        #region //撈取品號ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT MtlItemId, MtlItemNo
                                FROM PDM.MtlItem
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<MtlItem> resultMtlitems = sqlConnection.Query<MtlItem>(sql, dynamicParameters).ToList();

                        tsrnDetails = tsrnDetails.Join(resultMtlitems, x => x.MtlItemNo, y => y.MtlItemNo, (x, y) => { x.MtlItemId = y.MtlItemId; return x; }).Select(x => x).ToList();
                        #endregion

                        #region //撈取庫別ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT InventoryId, InventoryNo
                                FROM SCM.Inventory
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<Inventory> resultInventories = sqlConnection.Query<Inventory>(sql, dynamicParameters).ToList();

                        tsrnDetails = tsrnDetails.Join(resultInventories, x => x.TsrnOutInventoryNo, y => y.InventoryNo, (x, y) => { x.TsrnOutInventory = y.InventoryId; return x; }).ToList();
                        tsrnDetails = tsrnDetails.Join(resultInventories, x => x.TsrnInInventoryNo, y => y.InventoryNo, (x, y) => { x.TsrnInInventory = y.InventoryId; return x; }).ToList();
                        #endregion

                        #region //撈取暫出單單身ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.TsnId, a.TsnErpPrefix, a.TsnErpNo, b.TsnDetailId, b.TsnSequence
                                FROM SCM.TempShippingNote a
                                LEFT JOIN SCM.TsnDetail b ON b.TsnId = a.TsnId
                                WHERE a.CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<TsnDetail> resultTsnDetails = sqlConnection.Query<TsnDetail>(sql, dynamicParameters).ToList();

                        tsrnDetails = tsrnDetails.GroupJoin(resultTsnDetails, x => new { x.TsnErpPrefix, x.TsnErpNo, x.TsnSequence }, y => new { y.TsnErpPrefix, y.TsnErpNo, y.TsnSequence }, (x, y) => { x.TsnDetailId = y.FirstOrDefault()?.TsnDetailId; return x; }).Select(x => x).ToList();
                        #endregion

                        #region //判斷SCM.TempShippingReturnNote是否有資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TsrnId, TsrnErpPrefix, TsrnErpNo 
                                FROM SCM.TempShippingReturnNote
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<TempShippingReturnNote> resultTempShippingReturnNotes = sqlConnection.Query<TempShippingReturnNote>(sql, dynamicParameters).ToList();

                        tempShippingReturnNotes = tempShippingReturnNotes.GroupJoin(resultTempShippingReturnNotes, x => new { x.TsrnErpPrefix, x.TsrnErpNo }, y => new { y.TsrnErpPrefix, y.TsrnErpNo }, (x, y) => { x.TsrnId = y.FirstOrDefault()?.TsrnId; return x; }).Select(x => x).ToList();
                        #endregion

                        #region //暫出歸還單單頭(新增/修改)
                        List<TempShippingReturnNote> addTempShippingReturnNotes = tempShippingReturnNotes.Where(x => x.TsrnId == null).ToList();
                        List<TempShippingReturnNote> updateTempShippingReturnNotes = tempShippingReturnNotes.Where(x => x.TsrnId != null).ToList();

                        #region //新增
                        dynamicParameters = new DynamicParameters();
                        if (addTempShippingReturnNotes.Count > 0)
                        {
                            addTempShippingReturnNotes
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

                            sql = @"INSERT INTO SCM.TempShippingReturnNote (CompanyId, TsrnErpPrefix, TsrnErpNo, TsrnDate
                                    , DocDate, ToObject, ObjectCustomer, ObjectSupplier, ObjectUser, ObjectOther
                                    , DepartmentId, UserId, Remark, OtherRemark, TotalQty, Amount, TaxAmount
                                    , ConfirmStatus, ConfirmUserId, TransferStatus, TransferDate
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@CompanyId, @TsrnErpPrefix, @TsrnErpNo, @TsrnDate
                                    , @DocDate, @ToObject, @ObjectCustomer, @ObjectSupplier, @ObjectUser, @ObjectOther
                                    , @DepartmentId, @UserId, @Remark, @OtherRemark, @TotalQty, @Amount, @TaxAmount
                                    , @ConfirmStatus, @ConfirmUserId, @TransferStatus, @TransferDate
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            rowsAffected += sqlConnection.Execute(sql, addTempShippingReturnNotes);
                        }
                        #endregion

                        #region //修改
                        dynamicParameters = new DynamicParameters();
                        if (updateTempShippingReturnNotes.Count > 0)
                        {
                            updateTempShippingReturnNotes
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.LastModifiedDate = LastModifiedDate;
                                    x.LastModifiedBy = LastModifiedBy;
                                });

                            sql = @"UPDATE SCM.TempShippingReturnNote SET
                                    TsrnErpPrefix = @TsrnErpPrefix,
                                    TsrnErpNo = @TsrnErpNo,
                                    TsrnDate = @TsrnDate,
                                    DocDate = @DocDate,
                                    ToObject = @ToObject,
                                    ObjectCustomer = @ObjectCustomer,
                                    ObjectSupplier = @ObjectSupplier,
                                    ObjectUser = @ObjectUser,
                                    ObjectOther = @ObjectOther,
                                    DepartmentId = @DepartmentId,
                                    UserId = @UserId,
                                    Remark = @Remark,
                                    OtherRemark = @OtherRemark,
                                    TotalQty = @TotalQty,
                                    Amount = @Amount,
                                    TaxAmount = @TaxAmount,
                                    ConfirmStatus = @ConfirmStatus,
                                    ConfirmUserId = @ConfirmUserId,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE TsrnId = @TsrnId";
                            rowsAffected += sqlConnection.Execute(sql, updateTempShippingReturnNotes);
                        }
                        #endregion
                        #endregion

                        #region //暫出歸還單單身(新增/修改)
                        #region //撈取暫出歸還單單頭ID
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TsrnId, TsrnErpPrefix, TsrnErpNo
                                FROM  SCM.TempShippingReturnNote
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        resultTempShippingReturnNotes = sqlConnection.Query<TempShippingReturnNote>(sql, dynamicParameters).ToList();

                        tsrnDetails = tsrnDetails.Join(resultTempShippingReturnNotes, x => new { x.TsrnErpPrefix, x.TsrnErpNo }, y => new { y.TsrnErpPrefix, y.TsrnErpNo }, (x, y) => { x.TsrnId = y.TsrnId; return x; }).Select(x => x).ToList();
                        #endregion

                        #region //判斷SCM.TsrnDetail是否有資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.TsrnDetailId, a.TsrnId, a.TsrnSequence
                                FROM SCM.TsrnDetail a
                                INNER JOIN SCM.TempShippingReturnNote b ON a.TsrnId = b.TsrnId
                                WHERE b.CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CompanyId);

                        List<TsrnDetail> resultItDetail = sqlConnection.Query<TsrnDetail>(sql, dynamicParameters).ToList();

                        tsrnDetails = tsrnDetails.GroupJoin(resultItDetail, x => new { x.TsrnId, x.TsrnSequence }, y => new { y.TsrnId, y.TsrnSequence }, (x, y) => { x.TsrnDetailId = y.FirstOrDefault()?.TsrnDetailId; return x; }).ToList();
                        #endregion

                        List<TsrnDetail> addTsrnDetail = tsrnDetails.Where(x => x.TsrnDetailId == null).ToList();
                        List<TsrnDetail> updateTsrnDetail = tsrnDetails.Where(x => x.TsrnDetailId != null).ToList();

                        #region //新增
                        if (addTsrnDetail.Count > 0)
                        {
                            addTsrnDetail
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.CreateDate = CreateDate;
                                    x.LastModifiedDate = LastModifiedDate;
                                    x.CreateBy = CreateBy;
                                    x.LastModifiedBy = LastModifiedBy;
                                });

                            sql = @"INSERT INTO SCM.TsrnDetail (TsrnId, TsrnSequence, MtlItemId, TsrnMtlItemName
                                    , TsrnMtlItemSpec, TsrnOutInventory, TsrnInInventory, TsrnQty, UnitPrice
                                    , ProductType, FreebieOrSpareQty, TsrnPriceQty, TsrnPriceUomId, BusinessTaxRate
                                    , Amount, TsnDetailId, TsrnRemark, ConfirmStatus
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@TsrnId, @TsrnSequence, @MtlItemId, @TsrnMtlItemName
                                    , @TsrnMtlItemSpec, @TsrnOutInventory, @TsrnInInventory, @TsrnQty, @UnitPrice
                                    , @ProductType, @FreebieOrSpareQty, @TsrnPriceQty, @TsrnPriceUomId, @BusinessTaxRate
                                    , @Amount, @TsnDetailId, @TsrnRemark, @ConfirmStatus
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            rowsAffected += sqlConnection.Execute(sql, addTsrnDetail);
                        }
                        #endregion

                        #region //修改
                        if (updateTsrnDetail.Count > 0)
                        {
                            updateTsrnDetail
                                .ToList()
                                .ForEach(x =>
                                {
                                    x.LastModifiedDate = LastModifiedDate;
                                    x.LastModifiedBy = LastModifiedBy;
                                });

                            sql = @"UPDATE SCM.TsrnDetail SET
                                    TsrnId = @TsrnId,
                                    TsrnSequence = @TsrnSequence,
                                    MtlItemId = @MtlItemId,
                                    TsrnMtlItemName = @TsrnMtlItemName,
                                    TsrnMtlItemSpec = @TsrnMtlItemSpec,
                                    TsrnOutInventory = @TsrnOutInventory,
                                    TsrnInInventory = @TsrnInInventory,
                                    TsrnQty = @TsrnQty,
                                    UnitPrice = @UnitPrice,
                                    ProductType = @ProductType,
                                    FreebieOrSpareQty = @FreebieOrSpareQty,
                                    TsrnPriceQty = @TsrnPriceQty,
                                    TsrnPriceUomId = @TsrnPriceUomId,
                                    BusinessTaxRate = @BusinessTaxRate,
                                    Amount = @Amount,
                                    TsnDetailId = @TsnDetailId,
                                    TsrnRemark = @TsrnRemark,
                                    ConfirmStatus = @ConfirmStatus,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE TsrnDetailId = @TsrnDetailId";
                            rowsAffected += sqlConnection.Execute(sql, updateTsrnDetail);
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

        #region //UpdateTempShippingReturnNoteManualSynchronize -- 暫出歸還單資料手動同步 -- Ben Ma 2023.05.23
        public string UpdateTempShippingReturnNoteManualSynchronize(string TsrnErpFullNo, string SyncStartDate, string SyncEndDate
            , string NormalSync, string TranSync)
        {
            try
            {
                List<TempShippingReturnNote> tempShippingReturnNotes = new List<TempShippingReturnNote>();
                List<TsrnDetail> tsrnDetails = new List<TsrnDetail>();

                int mainAffected = 0, detailAffected = 0, mainDelAffected = 0, detailDelAffected = 0; ;
                string companyNo = "";

                using (TransactionScope transactionScope = new TransactionScope(TransactionScopeOption.Required, new TimeSpan(0, 180, 0)))
                {
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
                    }

                    #region //正常同步
                    if (NormalSync == "Y")
                    {
                        using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //判斷ERP暫出歸還單資料是否存在
                            if (TsrnErpFullNo.Length > 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM INVTH
                                        WHERE (LTRIM(RTRIM(TH001)) + '-' + LTRIM(RTRIM(TH002))) LIKE '%' + @TsrnErpFullNo + '%'";
                                dynamicParameters.Add("TsrnErpFullNo", TsrnErpFullNo);

                                var resultTsn = sqlConnection.Query(sql, dynamicParameters);
                                if (resultTsn.Count() <= 0) throw new SystemException("ERP暫出歸還單資料錯誤");
                            }
                            #endregion

                            #region //撈取ERP暫出歸還單資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(TH001)) TsrnErpPrefix, LTRIM(RTRIM(TH002)) TsrnErpNo
                                    , CASE WHEN LEN(LTRIM(RTRIM(TH003))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TH003)) as date), 'yyyy-MM-dd') ELSE NULL END TsrnDate
                                    , CASE WHEN LEN(LTRIM(RTRIM(TH023))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(TH023)) as date), 'yyyy-MM-dd') ELSE NULL END DocDate
                                    , LTRIM(RTRIM(TH004)) ToObject, LTRIM(RTRIM(TH005)) ObjectOther, LTRIM(RTRIM(TH007)) DepartmentNo
                                    , LTRIM(RTRIM(TH008)) UserNo, LTRIM(RTRIM(TH014)) Remark, LTRIM(RTRIM(TH018)) Remark
                                    , TH021 TotalQty, TH022 Amount, TH026 TaxAmount
                                    , LTRIM(RTRIM(TH020)) ConfirmStatus, LTRIM(RTRIM(TH024)) ConfirmUserNo
                                    , CASE WHEN LEN(LTRIM(RTRIM(CREATE_DATE))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_DATE)) as date), 'yyyy-MM-dd') ELSE NULL END TransferDate
                                    , CASE WHEN LEN(LTRIM(RTRIM(CREATE_TIME))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(CREATE_TIME)) as datetime), 'HH:mm:ss') ELSE NULL END TransferTime
                                    FROM INVTH
                                    WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TsrnErpFullNo", @" AND (LTRIM(RTRIM(TH001)) + '-' + LTRIM(RTRIM(TH002))) LIKE '%' + @TsrnErpFullNo + '%'", TsrnErpFullNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SyncStartDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME >= @SyncStartDate OR MODI_DATE + ' ' + MODI_TIME >= @SyncStartDate)", SyncStartDate.Length > 0 ? Convert.ToDateTime(SyncStartDate).ToString("yyyyMMdd 00:00:00") : "");
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SyncEndDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME <= @SyncEndDate OR MODI_DATE + ' ' + MODI_TIME <= @SyncEndDate)", SyncEndDate.Length > 0 ? Convert.ToDateTime(SyncEndDate).ToString("yyyyMMdd 23:59:59") : "");
                            sql += @" ORDER BY TransferDate, TransferTime";

                            tempShippingReturnNotes = sqlConnection.Query<TempShippingReturnNote>(sql, dynamicParameters).ToList();
                            #endregion

                            #region //撈取ERP暫出歸還單單身資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(TI001)) TsrnErpPrefix, LTRIM(RTRIM(TI002)) TsrnErpNo, LTRIM(RTRIM(TI003)) TsrnSequence
                                    , LTRIM(RTRIM(TI004)) MtlItemNo, LTRIM(RTRIM(TI005)) TsrnMtlItemName, LTRIM(RTRIM(TI006)) TsrnMtlItemSpec
                                    , LTRIM(RTRIM(TI007)) TsrnOutInventoryNo, LTRIM(RTRIM(TI008)) TsrnInInventoryNo
                                    , LTRIM(RTRIM(CAST(TI009 AS decimal(16,3)))) TsrnQty, LTRIM(RTRIM(CAST(TI012 AS decimal(21,6)))) UnitPrice
                                    , LTRIM(RTRIM(CAST(TI013 AS decimal(21,6)))) Amount
                                    , LTRIM(RTRIM(TI034)) ProductType, LTRIM(RTRIM(CAST(TI035 AS INT))) FreebieOrSpareQty, LTRIM(RTRIM(CAST(TI038 AS INT))) TsrnPriceQty
                                    , LTRIM(RTRIM(TI039)) TsrnPriceUomrNo, LTRIM(RTRIM(TI033)) BusinessTaxRate
                                    , LTRIM(RTRIM(TI014)) TsnErpPrefix, LTRIM(RTRIM(TI015)) TsnErpNo, LTRIM(RTRIM(TI016)) TsnSequence
                                    , LTRIM(RTRIM(TI021)) TsrnRemark, LTRIM(RTRIM(TI022)) ConfirmStatus
                                    FROM INVTI
                                    WHERE 1=1";
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "TsrnErpFullNo", @" AND (LTRIM(RTRIM(TI001)) + '-' + LTRIM(RTRIM(TI002))) LIKE '%' + @TsrnErpFullNo + '%'", TsrnErpFullNo);
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SyncStartDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME >= @SyncStartDate OR MODI_DATE + ' ' + MODI_TIME >= @SyncStartDate)", SyncStartDate.Length > 0 ? Convert.ToDateTime(SyncStartDate).ToString("yyyyMMdd 00:00:00") : "");
                            BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SyncEndDate", @" AND (CREATE_DATE + ' ' + CREATE_TIME <= @SyncEndDate OR MODI_DATE + ' ' + MODI_TIME <= @SyncEndDate)", SyncEndDate.Length > 0 ? Convert.ToDateTime(SyncEndDate).ToString("yyyyMMdd 23:59:59") : "");

                            tsrnDetails = sqlConnection.Query<TsrnDetail>(sql, dynamicParameters).ToList();
                            #endregion
                        }

                        using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                        {
                            #region //撈取部門ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT DepartmentId, DepartmentNo
                                    FROM BAS.Department
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<Department> resultDepartments = sqlConnection.Query<Department>(sql, dynamicParameters).ToList();

                            tempShippingReturnNotes = tempShippingReturnNotes.Join(resultDepartments, x => x.DepartmentNo, y => y.DepartmentNo, (x, y) => { x.DepartmentId = y.DepartmentId; return x; }).ToList();
                            #endregion

                            #region //撈取人員ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.UserId, a.UserNo
                                    FROM BAS.[User] a";

                            List<User> resultUsers = sqlConnection.Query<User>(sql, dynamicParameters).ToList();

                            tempShippingReturnNotes = tempShippingReturnNotes.Join(resultUsers, x => x.UserNo, y => y.UserNo, (x, y) => { x.UserId = y.UserId; return x; }).ToList();
                            tempShippingReturnNotes = tempShippingReturnNotes.GroupJoin(resultUsers, x => x.ConfirmUserNo, y => y.UserNo, (x, y) => { x.ConfirmUserId = y.FirstOrDefault()?.UserId; return x; }).ToList();
                            List<TempShippingReturnNote> toObjectUser = tempShippingReturnNotes.Where(x => x.ToObject == "3").Join(resultUsers, x => x.ObjectOther, y => y.UserNo, (x, y) => { x.ObjectUser = y.UserId; return x; }).Select(x => x).ToList();
                            #endregion

                            #region //撈取異動對象ID
                            #region //撈取客戶ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT CustomerId, CustomerNo 
                                    FROM SCM.Customer
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<Customer> resultCustomers = sqlConnection.Query<Customer>(sql, dynamicParameters).ToList();

                            List<TempShippingReturnNote> toObjectCustomer = tempShippingReturnNotes.Where(x => x.ToObject == "1").Join(resultCustomers, x => x.ObjectOther, y => y.CustomerNo, (x, y) => { x.ObjectCustomer = y.CustomerId; return x; }).Select(x => x).ToList();
                            #endregion

                            #region //撈取供應商ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT SupplierId, SupplierNo 
                                    FROM SCM.Supplier
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<Supplier> resultSuppliers = sqlConnection.Query<Supplier>(sql, dynamicParameters).ToList();

                            List<TempShippingReturnNote> toObjectSupplier = tempShippingReturnNotes.Where(x => x.ToObject == "2").Join(resultSuppliers, x => x.ObjectOther, y => y.SupplierNo, (x, y) => { x.ObjectSupplier = y.SupplierId; return x; }).Select(x => x).ToList();
                            #endregion

                            #region //其他
                            List<TempShippingReturnNote> toObjectOther = tempShippingReturnNotes.Where(x => x.ToObject == "9").Select(x => x).ToList();
                            #endregion

                            tempShippingReturnNotes = toObjectCustomer.Union(toObjectSupplier).Union(toObjectUser).Union(toObjectOther).ToList();
                            #endregion

                            #region //撈取品號ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT MtlItemId, MtlItemNo
                                    FROM PDM.MtlItem
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<MtlItem> resultMtlitems = sqlConnection.Query<MtlItem>(sql, dynamicParameters).ToList();

                            tsrnDetails = tsrnDetails.Join(resultMtlitems, x => x.MtlItemNo, y => y.MtlItemNo, (x, y) => { x.MtlItemId = y.MtlItemId; return x; }).Select(x => x).ToList();
                            #endregion

                            #region //撈取庫別ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT InventoryId, InventoryNo
                                    FROM SCM.Inventory
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<Inventory> resultInventories = sqlConnection.Query<Inventory>(sql, dynamicParameters).ToList();

                            tsrnDetails = tsrnDetails.Join(resultInventories, x => x.TsrnOutInventoryNo, y => y.InventoryNo, (x, y) => { x.TsrnOutInventory = y.InventoryId; return x; }).ToList();
                            tsrnDetails = tsrnDetails.Join(resultInventories, x => x.TsrnInInventoryNo, y => y.InventoryNo, (x, y) => { x.TsrnInInventory = y.InventoryId; return x; }).ToList();
                            #endregion

                            #region //撈取暫出單單身ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.TsnId, a.TsnErpPrefix, a.TsnErpNo, b.TsnDetailId, b.TsnSequence
                                    FROM SCM.TempShippingNote a
                                    LEFT JOIN SCM.TsnDetail b ON b.TsnId = a.TsnId
                                    WHERE a.CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<TsnDetail> resultTsnDetails = sqlConnection.Query<TsnDetail>(sql, dynamicParameters).ToList();

                            tsrnDetails = tsrnDetails.GroupJoin(resultTsnDetails, x => new { x.TsnErpPrefix, x.TsnErpNo, x.TsnSequence }, y => new { y.TsnErpPrefix, y.TsnErpNo, y.TsnSequence }, (x, y) => { x.TsnDetailId = y.FirstOrDefault()?.TsnDetailId; return x; }).Select(x => x).ToList();
                            #endregion

                            #region //判斷SCM.TempShippingReturnNote是否有資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TsrnId, TsrnErpPrefix, TsrnErpNo 
                                    FROM SCM.TempShippingReturnNote
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<TempShippingReturnNote> resultTempShippingReturnNotes = sqlConnection.Query<TempShippingReturnNote>(sql, dynamicParameters).ToList();

                            tempShippingReturnNotes = tempShippingReturnNotes.GroupJoin(resultTempShippingReturnNotes, x => new { x.TsrnErpPrefix, x.TsrnErpNo }, y => new { y.TsrnErpPrefix, y.TsrnErpNo }, (x, y) => { x.TsrnId = y.FirstOrDefault()?.TsrnId; return x; }).Select(x => x).ToList();
                            #endregion

                            #region //暫出歸還單單頭(新增/修改)
                            List<TempShippingReturnNote> addTempShippingReturnNotes = tempShippingReturnNotes.Where(x => x.TsrnId == null).ToList();
                            List<TempShippingReturnNote> updateTempShippingReturnNotes = tempShippingReturnNotes.Where(x => x.TsrnId != null).ToList();

                            #region //新增
                            dynamicParameters = new DynamicParameters();
                            if (addTempShippingReturnNotes.Count > 0)
                            {
                                addTempShippingReturnNotes
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.CompanyId = CurrentCompany;
                                        x.TransferStatus = "Y";
                                        x.CreateDate = CreateDate;
                                        x.LastModifiedDate = LastModifiedDate;
                                        x.CreateBy = CreateBy;
                                        x.LastModifiedBy = LastModifiedBy;
                                    });

                                sql = @"INSERT INTO SCM.TempShippingReturnNote (CompanyId, TsrnErpPrefix, TsrnErpNo, TsrnDate
                                        , DocDate, ToObject, ObjectCustomer, ObjectSupplier, ObjectUser, ObjectOther
                                        , DepartmentId, UserId, Remark, OtherRemark, TotalQty, Amount, TaxAmount
                                        , ConfirmStatus, ConfirmUserId, TransferStatus, TransferDate
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        VALUES (@CompanyId, @TsrnErpPrefix, @TsrnErpNo, @TsrnDate
                                        , @DocDate, @ToObject, @ObjectCustomer, @ObjectSupplier, @ObjectUser, @ObjectOther
                                        , @DepartmentId, @UserId, @Remark, @OtherRemark, @TotalQty, @Amount, @TaxAmount
                                        , @ConfirmStatus, @ConfirmUserId, @TransferStatus, @TransferDate
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                int addMain = sqlConnection.Execute(sql, addTempShippingReturnNotes);
                                mainAffected += addMain;
                            }
                            #endregion

                            #region //修改
                            dynamicParameters = new DynamicParameters();
                            if (updateTempShippingReturnNotes.Count > 0)
                            {
                                updateTempShippingReturnNotes
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.LastModifiedDate = LastModifiedDate;
                                        x.LastModifiedBy = LastModifiedBy;
                                    });

                                sql = @"UPDATE SCM.TempShippingReturnNote SET
                                        TsrnErpPrefix = @TsrnErpPrefix,
                                        TsrnErpNo = @TsrnErpNo,
                                        TsrnDate = @TsrnDate,
                                        DocDate = @DocDate,
                                        ToObject = @ToObject,
                                        ObjectCustomer = @ObjectCustomer,
                                        ObjectSupplier = @ObjectSupplier,
                                        ObjectUser = @ObjectUser,
                                        ObjectOther = @ObjectOther,
                                        DepartmentId = @DepartmentId,
                                        UserId = @UserId,
                                        Remark = @Remark,
                                        OtherRemark = @OtherRemark,
                                        TotalQty = @TotalQty,
                                        Amount = @Amount,
                                        TaxAmount = @TaxAmount,
                                        ConfirmStatus = @ConfirmStatus,
                                        ConfirmUserId = @ConfirmUserId,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE TsrnId = @TsrnId";
                                int updateMain = sqlConnection.Execute(sql, updateTempShippingReturnNotes);
                                mainAffected += updateMain;
                            }
                            #endregion
                            #endregion

                            #region //暫出歸還單單身(新增/修改)
                            #region //撈取暫出歸還單單頭ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TsrnId, TsrnErpPrefix, TsrnErpNo
                                    FROM SCM.TempShippingReturnNote
                                    WHERE CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            resultTempShippingReturnNotes = sqlConnection.Query<TempShippingReturnNote>(sql, dynamicParameters).ToList();

                            tsrnDetails = tsrnDetails.Join(resultTempShippingReturnNotes, x => new { x.TsrnErpPrefix, x.TsrnErpNo }, y => new { y.TsrnErpPrefix, y.TsrnErpNo }, (x, y) => { x.TsrnId = y.TsrnId; return x; }).Select(x => x).ToList();
                            #endregion

                            #region //判斷SCM.TsrnDetail是否有資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.TsrnDetailId, a.TsrnId, a.TsrnSequence, a.TsnDetailId
                                    FROM SCM.TsrnDetail a
                                    INNER JOIN SCM.TempShippingReturnNote b ON a.TsrnId = b.TsrnId
                                    WHERE b.CompanyId = @CompanyId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            List<TsrnDetail> resultItDetail = sqlConnection.Query<TsrnDetail>(sql, dynamicParameters).ToList();

                            tsrnDetails = tsrnDetails.GroupJoin(resultItDetail, x => new { x.TsrnId, x.TsrnSequence }, y => new { y.TsrnId, y.TsrnSequence }, (x, y) => { x.TsrnDetailId = y.FirstOrDefault()?.TsrnDetailId; return x; }).ToList();
                            #endregion

                            List<TsrnDetail> addTsrnDetail = tsrnDetails.Where(x => x.TsrnDetailId == null).ToList();
                            List<TsrnDetail> updateTsrnDetail = tsrnDetails.Where(x => x.TsrnDetailId != null).ToList();

                            #region //新增
                            if (addTsrnDetail.Count > 0)
                            {
                                addTsrnDetail
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.CreateDate = CreateDate;
                                        x.LastModifiedDate = LastModifiedDate;
                                        x.CreateBy = CreateBy;
                                        x.LastModifiedBy = LastModifiedBy;
                                    });

                                sql = @"INSERT INTO SCM.TsrnDetail (TsrnId, TsrnSequence, MtlItemId, TsrnMtlItemName
                                        , TsrnMtlItemSpec, TsrnOutInventory, TsrnInInventory, TsrnQty, UnitPrice
                                        , ProductType, FreebieOrSpareQty, TsrnPriceQty, TsrnPriceUomId, BusinessTaxRate
                                        , Amount, TsnDetailId, TsrnRemark, ConfirmStatus
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        VALUES (@TsrnId, @TsrnSequence, @MtlItemId, @TsrnMtlItemName
                                        , @TsrnMtlItemSpec, @TsrnOutInventory, @TsrnInInventory, @TsrnQty, @UnitPrice
                                        , @ProductType, @FreebieOrSpareQty, @TsrnPriceQty, @TsrnPriceUomId, @BusinessTaxRate
                                        , @Amount, @TsnDetailId, @TsrnRemark, @ConfirmStatus
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                int addDetail = sqlConnection.Execute(sql, addTsrnDetail);
                                detailAffected += addDetail;
                            }
                            #endregion

                            #region //修改
                            if (updateTsrnDetail.Count > 0)
                            {
                                updateTsrnDetail
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.LastModifiedDate = LastModifiedDate;
                                        x.LastModifiedBy = LastModifiedBy;
                                    });

                                sql = @"UPDATE SCM.TsrnDetail SET
                                        TsrnId = @TsrnId,
                                        TsrnSequence = @TsrnSequence,
                                        MtlItemId = @MtlItemId,
                                        TsrnMtlItemName = @TsrnMtlItemName,
                                        TsrnMtlItemSpec = @TsrnMtlItemSpec,
                                        TsrnOutInventory = @TsrnOutInventory,
                                        TsrnInInventory = @TsrnInInventory,
                                        TsrnQty = @TsrnQty,
                                        UnitPrice = @UnitPrice,
                                        ProductType = @ProductType,
                                        FreebieOrSpareQty = @FreebieOrSpareQty,
                                        TsrnPriceQty = @TsrnPriceQty,
                                        TsrnPriceUomId = @TsrnPriceUomId,
                                        BusinessTaxRate = @BusinessTaxRate,
                                        Amount = @Amount,
                                        TsnDetailId = @TsnDetailId,
                                        TsrnRemark = @TsrnRemark,
                                        ConfirmStatus = @ConfirmStatus,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE TsrnDetailId = @TsrnDetailId";
                                int updateDetail = sqlConnection.Execute(sql, updateTsrnDetail);
                                detailAffected += updateDetail;
                            }
                            #endregion
                            #endregion
                        }
                    }
                    #endregion

                    #region //異動同步
                    if (TranSync == "Y")
                    {
                        using (SqlConnection erpConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //ERP暫出歸還單單頭資料
                            sql = @"SELECT TH001 TsrnErpPrefix, TH002 TsrnErpNo
                                    FROM INVTH
                                    WHERE 1=1
                                    ORDER BY TH001, TH002";
                            var resultErpTsrn = erpConnection.Query<TempShippingReturnNote>(sql, dynamicParameters);
                            #endregion

                            using (SqlConnection bmConnection = new SqlConnection(MainConnectionStrings))
                            {
                                #region //BM暫出歸還單單頭資料
                                sql = @"SELECT TsrnErpPrefix, TsrnErpNo
                                        FROM SCM.TempShippingReturnNote
                                        WHERE CompanyId = @CompanyId
                                        ORDER BY TsrnErpPrefix, TsrnErpNo";
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                var resultBmTsrn = bmConnection.Query<TempShippingReturnNote>(sql, dynamicParameters);
                                #endregion

                                #region //篩選ERP不存在但BM還存在的暫出歸還單單頭
                                var dictionaryErpTsrn = resultErpTsrn.ToDictionary(x => x.TsrnErpPrefix + "-" + x.TsrnErpNo, x => x.TsrnErpPrefix + "-" + x.TsrnErpNo);
                                var dictionaryBmTsrn = resultBmTsrn.ToDictionary(x => x.TsrnErpPrefix + "-" + x.TsrnErpNo, x => x.TsrnErpPrefix + "-" + x.TsrnErpNo);
                                var changeTsrn = dictionaryBmTsrn.Where(x => !dictionaryErpTsrn.ContainsKey(x.Key)).ToList();
                                var changeTsrnList = changeTsrn.Select(x => x.Value).ToList();
                                #endregion

                                #region //刪除異動暫出歸還單單頭
                                if (changeTsrnList.Count > 0)
                                {
                                    #region //單身刪除
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE SCM.TsrnDetail
                                            WHERE TsrnId IN (
                                                SELECT TsrnId
                                                FROM SCM.TempShippingReturnNote
                                                WHERE TsrnErpPrefix + '-' + TsrnErpNo IN @TsrnErpFullNo
                                            )";
                                    dynamicParameters.Add("TsrnErpFullNo", changeTsrnList.Select(x => x).ToArray());
                                    mainDelAffected += bmConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //單頭刪除
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE SCM.TempShippingReturnNote
                                            WHERE TsrnErpPrefix + '-' + TsrnErpNo IN @TsrnErpFullNo";
                                    dynamicParameters.Add("TsrnErpFullNo", changeTsrnList.Select(x => x).ToArray());
                                    detailDelAffected += bmConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                                #endregion
                            }

                            #region //ERP暫出歸還單單身資料
                            sql = @"SELECT TI001 TsrnErpPrefix, TI002 TsrnErpNo, TI003 TsrnSequence
                                    FROM INVTI
                                    WHERE 1=1
                                    ORDER BY TI001, TI002, TI003";
                            var resultErpTsrnDetail = erpConnection.Query<TsrnDetail>(sql, dynamicParameters);
                            #endregion

                            using (SqlConnection bmConnection = new SqlConnection(MainConnectionStrings))
                            {
                                #region //BM暫出歸還單單身資料
                                sql = @"SELECT b.TsrnErpPrefix, b.TsrnErpNo, a.TsrnSequence
                                        FROM SCM.TsrnDetail a
                                        INNER JOIN SCM.TempShippingReturnNote b ON a.TsrnId = b.TsrnId
                                        WHERE b.CompanyId = @CompanyId
                                        ORDER BY b.TsrnErpPrefix, b.TsrnErpNo, a.TsrnSequence";
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                var resultBmTsrnDetail = bmConnection.Query<TsrnDetail>(sql, dynamicParameters);
                                #endregion

                                #region //篩選ERP不存在但BM還存在的暫出歸還單單身
                                var dictionaryErpTsrnDetail = resultErpTsrnDetail.ToDictionary(x => x.TsrnErpPrefix + "-" + x.TsrnErpNo + "-" + x.TsrnSequence, x => x.TsrnErpPrefix + "-" + x.TsrnErpNo + "-" + x.TsrnSequence);
                                var dictionaryBmTsrnDetail = resultBmTsrnDetail.ToDictionary(x => x.TsrnErpPrefix + "-" + x.TsrnErpNo + "-" + x.TsrnSequence, x => x.TsrnErpPrefix + "-" + x.TsrnErpNo + "-" + x.TsrnSequence);
                                var changeTsrnDetail = dictionaryBmTsrnDetail.Where(x => !dictionaryErpTsrnDetail.ContainsKey(x.Key)).ToList();
                                var changeTsrnDetailList = changeTsrnDetail.Select(x => x.Value).ToList();
                                #endregion

                                #region //刪除異動暫出歸還單單身
                                if (changeTsrnDetailList.Count > 0)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE a
                                            FROM SCM.TsrnDetail a
                                            INNER JOIN SCM.TempShippingReturnNote b ON a.TsrnId = b.TsrnId
                                            WHERE b.TsrnErpPrefix + '-' + b.TsrnErpNo + '-' + a.TsrnSequence IN @TsrnErpFullNo";
                                    dynamicParameters.Add("TsrnErpFullNo", changeTsrnDetailList.Select(x => x).ToArray());
                                    detailDelAffected += bmConnection.Execute(sql, dynamicParameters);
                                }
                                #endregion
                            }
                        }
                    }
                    #endregion

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "(" + mainAffected + detailAffected + " rows affected)",
                        data = "已更新資料<br />【" + mainAffected + "】筆單頭<br />【" + detailAffected + "】筆單身<br />刪除<br />【" + mainDelAffected + "】筆單頭<br />【" + detailDelAffected + "】筆單身"
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

        #region //UpdateTsnDocDate -- 暫出單單據日期資料更新 -- Yi 2023.11.15
        public string UpdateTsnDocDate(int TsnId, string DocDate)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷暫出單資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.TsnErpPrefix, a.TsnErpNo, FORMAT(a.DocDate, 'yyyy-MM-dd') DocDate
                                FROM SCM.TempShippingNote a
                                WHERE a.TsnId = @TsnId";
                        dynamicParameters.Add("TsnId", TsnId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("暫出單資料錯誤!");

                        string tsnErpPrefix = "", tsnErpNo = "";
                        foreach (var item in result)
                        {
                            tsnErpPrefix = item.TsnErpPrefix;
                            tsnErpNo = item.TsnErpNo;
                        }
                        #endregion

                        #region//更新暫出單單據日期
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.TempShippingNote SET
                                DocDate = @DocDate,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE TsnId = @TsnId
                                AND TsnErpPrefix = @TsnErpPrefix
                                AND TsnErpNo = @TsnErpNo";
                        dynamicParameters.AddDynamicParams(
                        new
                        {
                            DocDate,
                            LastModifiedDate,
                            LastModifiedBy,
                            TsnId,
                            TsnErpPrefix = tsnErpPrefix,
                            TsnErpNo = tsnErpNo
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
        #endregion

        #region //Delete
        #endregion
    }
}
