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
using System.Text.RegularExpressions;
using System.Transactions;
using System.Web;

namespace SCMDA
{
    public class DeliveryDA
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

        public DeliveryDA()
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
        #region //GetDeliverySchedule -- 取得出貨排程資料 -- Zoey 2022.08.31
        public string GetDeliverySchedule(string DeliveryStatus, string SoIds, string SoErpFullNo, string CustomerMtlItemNo, int CustomerId
            , string MtlItemNo, int SalesmenId, string StartDate, string EndDate
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                List<DeliverySchedule> deliverySchedules = new List<DeliverySchedule>();

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

                    sqlQuery.mainKey = "a.SoDetailId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.SoId, a.SoSequence, a.MtlItemId, a.InventoryId, a.UomId, a.SoQty, a.SiQty, a.SoDetailRemark
                        , FORMAT(a.PcPromiseDate, 'yyyy-MM-dd') PcPromiseDate
                        , FORMAT(a.PcPromiseDate, 'HH:mm:ss') PcPromiseTime
                        , a.ProductType, ISNULL(a.FreebieQty, 0) FreebieQty, ISNULL(a.SpareQty, 0) SpareQty
                        , ISNULL(i.PickRegularQty, 0) PickRegularQty, ISNULL(j.PickFreebieQty, 0) PickFreebieQty, ISNULL(k.PickSpareQty, 0) PickSpareQty
                        , b.SoErpPrefix, b.SoErpNo, b.SoErpPrefix + '-' + b.SoErpNo SoErpFullNo, b.CustomerId
                        , c.CustomerNo, c.CustomerName, c.CustomerShortName, c.CustomerNo + ' ' + c.CustomerShortName CustomerWithNo
                        , ISNULL(d.MtlItemNo, '') MtlItemNo, ISNULL(a.SoMtlItemName, ISNULL(d.MtlItemName, '')) MtlItemName, ISNULL(a.SoMtlItemSpec, ISNULL(d.MtlItemSpec, '')) MtlItemSpec
                        , (ISNULL(e.ItemQty, 0) - ISNULL(h.TotalRmaQty, 0)) PickQty
                        , (ISNULL(f.TsnQty, 0) - ISNULL(f.ReturnQty, 0)) TsnQty, ISNULL(f.SaleQty, 0) SaleQty
                        , g.LastBindStatus";
                    sqlQuery.mainTables =
                        @"FROM SCM.SoDetail a
                        INNER JOIN SCM.SaleOrder b ON b.SoId = a.SoId
                        INNER JOIN SCM.Customer c ON c.CustomerId = b.CustomerId
                        LEFT JOIN PDM.MtlItem d ON a.MtlItemId = d.MtlItemId
                        OUTER APPLY(
                            SELECT x.SoDetailId, SUM(x.ItemQty) ItemQty
                            FROM SCM.PickingItem x
                            WHERE x.SoDetailId = a.SoDetailId
                            GROUP BY x.SoDetailId
                        ) e
                        OUTER APPLY(
                            SELECT SUM(x.TsnQty) TsnQty, SUM(x.ReturnQty) ReturnQty, SUM(x.SaleQty) SaleQty
                            FROM SCM.TsnDetail x
                            INNER JOIN SCM.TempShippingNote y ON x.TsnId = y.TsnId
                            WHERE x.SoDetailId = a.SoDetailId
                            AND x.ConfirmStatus = 'Y'
                            AND y.ConfirmStatus = 'Y'
                        ) f
                        OUTER APPLY(
                            SELECT TOP 1 x.BindStatus LastBindStatus
                            FROM SCM.WipLinkLog x
                            WHERE x.SoErpPrefix = b.SoErpPrefix
                            AND x.SoErpNo = b.SoErpNo
                            AND x.SoSequence = a.SoSequence
                            ORDER BY x.CreateDate DESC
                        ) g
                        OUTER APPLY (
                            SELECT SUM(ha.RmaQty) TotalRmaQty
                            FROM SCM.RmaDetail ha
                            INNER JOIN SCM.TsnDetail hb ON ha.TsnDetailId = hb.TsnDetailId
                            INNER JOIN SCM.ReturnMerchandiseAuthorization hc ON ha.RmaId = hc.RmaId
                            WHERE hb.SoDetailId = a.SoDetailId
                            AND hc.ConfirmStatus = 'Y'
                        ) h
                        OUTER APPLY (
                            SELECT SUM(za.ItemQty) PickRegularQty
                            FROM SCM.PickingItem za
                            WHERE za.SoDetailId = a.SoDetailId
                            AND za.ItemType = 1
                        ) i
                        OUTER APPLY (
                            SELECT SUM(zb.ItemQty) PickFreebieQty
                            FROM SCM.PickingItem zb
                            WHERE zb.SoDetailId = a.SoDetailId
                            AND zb.ItemType = 2
                        ) j
                        OUTER APPLY (
                            SELECT SUM(zc.ItemQty) PickSpareQty
                            FROM SCM.PickingItem zc
                            WHERE zc.SoDetailId = a.SoDetailId
                            AND zc.ItemType = 3
                        ) k";
                    sqlQuery.auxTables = "";
                    string queryCondition = @" AND b.CompanyId = @CompanyId 
                                            AND a.ConfirmStatus = 'Y'";
                    switch (DeliveryStatus)
                    {
                        case "tracked":
                            queryCondition += @" AND a.ClosureStatus = 'N'
                                                AND (a.SoQty - ISNULL(e.ItemQty, 0)) > 0";
                            break;
                        case "triTrade":
                            queryCondition += @"AND d.MtlItemNo LIKE '5%'";
                            break;
                        default:
                            break;
                    }
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    if (SoIds.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SoIds", @" AND a.SoId IN @SoIds", SoIds.Split(','));
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerId", @" AND b.CustomerId = @CustomerId", CustomerId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemNo", @" AND (d.MtlItemNo LIKE '%' + @MtlItemNo + '%' OR d.MtlItemName LIKE '%' + @MtlItemNo + '%' OR a.CustomerMtlItemNo LIKE '%' + @MtlItemNo + '%')", MtlItemNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SalesmenId", @" AND b.SalesmenId = @SalesmenId", SalesmenId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.PcPromiseDate >= @StartDate ", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.PcPromiseDate <= @EndDate ", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.SoId DESC, a.SoSequence";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    deliverySchedules = BaseHelper.SqlQuery<DeliverySchedule>(sqlConnection, dynamicParameters, sqlQuery);
                }

                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                {
                    List<ErpAccounting> erpAccountings = new List<ErpAccounting>();

                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT LTRIM(RTRIM(a.TD001)) + '-' + LTRIM(RTRIM(a.TD002)) SoErpFullNo, a.TD003 SoSequence
                            , b1.TG009 erp_tsnQty, b1.TG021 erp_tsrnQty
                            ,CASE WHEN a.TD049 = '1' THEN b1.TG044  ELSE  0 END erp_tsnFreebieQty
                            ,CASE WHEN a.TD049 = '2' THEN b1.TG044  ELSE  0 END erp_tsnSpareQty
                            ,CASE WHEN a.TD049 = '1' THEN b1.TG048  ELSE  0 END erp_tsrnFreebieQty
                            ,CASE WHEN a.TD049 = '2' THEN b1.TG048  ELSE  0 END erp_tsrnSpareQty
                            , b2.TG009 erp_un_tsnQty
                            ,CASE WHEN a.TD049 = '1' THEN b2.TG044  ELSE  0 END erp_un_tsnFreebieQty
                            ,CASE WHEN a.TD049 = '2' THEN b2.TG044  ELSE  0 END erp_un_tsnSpareQty
                            , b3.TI009 erp_un_tsrnQty
                            ,CASE WHEN a.TD049 = '1' THEN b3.TI035  ELSE  0 END erp_un_tsrnFreebieQty
                            ,CASE WHEN a.TD049 = '2' THEN b3.TI035  ELSE  0 END erp_un_tsrnSpareQty
                            , c1.TH008 erp_un_snQty
                            ,CASE WHEN a.TD049 = '1' THEN c1.TH024  ELSE  0 END erp_un_snFreebieQty
                            ,CASE WHEN a.TD049 = '2' THEN c1.TH024  ELSE  0 END erp_un_snSpareQty
                            , c2.TJ007 erp_un_srnQty
                            ,CASE WHEN a.TD049 = '1' THEN c2.TJ042  ELSE  0 END erp_un_srnFreebieQty
                            ,CASE WHEN a.TD049 = '2' THEN c2.TJ042  ELSE  0 END erp_un_srnSpareQty
                            , c3.TH008 erp_snQty
                            ,CASE WHEN a.TD049 = '1' THEN c3.TH024  ELSE  0 END erp_snFreebieQty
                            ,CASE WHEN a.TD049 = '2' THEN c3.TH024  ELSE  0 END erp_snSpareQty
                            , c4.TJ007 erp_srnQty
                            ,CASE WHEN a.TD049 = '1' THEN c4.TJ042  ELSE  0 END erp_srnFreebieQty
                            ,CASE WHEN a.TD049 = '2' THEN c4.TJ042  ELSE  0 END erp_srnSpareQty
                            FROM COPTD a
                            OUTER APPLY(
                                --暫出單單身確認
                                SELECT SUM(x1.TG009) TG009, SUM(x1.TG044) TG044, SUM(x1.TG020) TG020, SUM(x1.TG021) TG021, SUM(x1.TG046) TG046, SUM(x1.TG048) TG048
                                FROM INVTG x1
                                WHERE LTRIM(RTRIM(x1.TG014)) = LTRIM(RTRIM(a.TD001))
                                AND LTRIM(RTRIM(x1.TG015)) = LTRIM(RTRIM(a.TD002))
                                AND LTRIM(RTRIM(x1.TG016)) = LTRIM(RTRIM(a.TD003))
                                AND x1.TG022 = 'Y'
                            ) b1
                            OUTER APPLY(
                                --暫出單單身未確認
                                SELECT SUM(x1.TG009) TG009 ,SUM(x1.TG044)  TG044
                                FROM  INVTG x1
                                WHERE LTRIM(RTRIM(x1.TG014)) = LTRIM(RTRIM(a.TD001))
                                AND LTRIM(RTRIM(x1.TG015)) = LTRIM(RTRIM(a.TD002))
                                AND LTRIM(RTRIM(x1.TG016)) = LTRIM(RTRIM(a.TD003))
                                AND x1.TG022 = 'N'
                            ) b2
                            OUTER APPLY(
                                --暫出歸還單單身未確認
                                SELECT SUM(x1.TI009) TI009 ,SUM(x1.TI035)  TI035
                                FROM  INVTI x1
                                INNER JOIN INVTG x2 on x2.TG001 = x1.TI014 AND x2.TG002 = x1.TI015 AND x2.TG003 = x1.TI016 
                                WHERE LTRIM(RTRIM(x2.TG014)) = LTRIM(RTRIM(a.TD001))
                                AND LTRIM(RTRIM(x2.TG015)) = LTRIM(RTRIM(a.TD002))
                                AND LTRIM(RTRIM(x2.TG016)) = LTRIM(RTRIM(a.TD003))
                                AND x1.TI022 = 'N'
                            ) b3
                            OUTER APPLY(
                                --銷貨未確認
                                SELECT SUM(x1.TH008) TH008 ,SUM(x1.TH024)  TH024
                                FROM  COPTH x1
                                WHERE LTRIM(RTRIM(x1.TH014)) = LTRIM(RTRIM(a.TD001))
                                AND x1.TH015 = LTRIM(RTRIM(a.TD002))
                                AND x1.TH016 = LTRIM(RTRIM(a.TD003))
                                AND x1.TH020 = 'N'
                            ) c1
                            OUTER APPLY(
                                --銷退未確認
                                SELECT SUM(x1.TJ007) TJ007 ,SUM(x1.TJ042)  TJ042
                                FROM  COPTJ x1
                                WHERE x1.TJ018 = LTRIM(RTRIM(a.TD001))
                                AND x1.TJ019 = LTRIM(RTRIM(a.TD002))
                                AND x1.TJ020 = LTRIM(RTRIM(a.TD003))
                                AND x1.TJ021 = 'N'
                            ) c2
                            OUTER APPLY(
                                --銷貨確認
                                SELECT SUM(x1.TH008) TH008 ,SUM(x1.TH024)  TH024
                                FROM  COPTH x1
                                WHERE x1.TH014 = LTRIM(RTRIM(a.TD001))
                                AND x1.TH015 = LTRIM(RTRIM(a.TD002))
                                AND x1.TH016 = LTRIM(RTRIM(a.TD003))
                                AND x1.TH020 = 'Y'
                            ) c3
                            OUTER APPLY(
                                --銷退確認
                                SELECT SUM(x1.TJ007) TJ007 ,SUM(x1.TJ042)  TJ042
                                FROM  COPTJ x1
                                WHERE x1.TJ018 = LTRIM(RTRIM(a.TD001))
                                AND x1.TJ019 = LTRIM(RTRIM(a.TD002))
                                AND x1.TJ020 = LTRIM(RTRIM(a.TD003))
                                AND x1.TJ021 = 'Y'
                            ) c4
                            WHERE (LTRIM(RTRIM(a.TD001)) + '-' + LTRIM(RTRIM(a.TD002)) + '-' + LTRIM(RTRIM(a.TD003))) IN @ErpFullNo";
                    dynamicParameters.Add("ErpFullNo", deliverySchedules.Select(x => x.SoErpFullNo + '-' + x.SoSequence).ToArray());
                    erpAccountings = sqlConnection.Query<ErpAccounting>(sql, dynamicParameters).ToList();

                    deliverySchedules = deliverySchedules.GroupJoin(erpAccountings, 
                        x => x.SoErpFullNo + '-' + x.SoSequence, 
                        y => y.SoErpFullNo + '-' + y.SoSequence, 
                        (x, y) => {
                            x.erp_tsnQty = y.FirstOrDefault()?.erp_tsnQty ?? 0;
                            x.erp_un_tsnQty = y.FirstOrDefault()?.erp_un_tsnQty ?? 0;
                            x.erp_tsrnQty = y.FirstOrDefault()?.erp_tsrnQty ?? 0;
                            x.erp_un_tsrnQty = y.FirstOrDefault()?.erp_un_tsrnQty ?? 0;
                            x.erp_tsnFreebieQty = y.FirstOrDefault()?.erp_tsnFreebieQty ?? 0;
                            x.erp_un_tsnFreebieQty = y.FirstOrDefault()?.erp_un_tsnFreebieQty ?? 0;
                            x.erp_tsrnFreebieQty = y.FirstOrDefault()?.erp_tsrnFreebieQty ?? 0;
                            x.erp_un_tsrnFreebieQty = y.FirstOrDefault()?.erp_un_tsrnFreebieQty ?? 0;
                            x.erp_tsnSpareQty = y.FirstOrDefault()?.erp_tsnSpareQty ?? 0;
                            x.erp_un_tsnSpareQty = y.FirstOrDefault()?.erp_un_tsnSpareQty ?? 0;
                            x.erp_tsrnSpareQty = y.FirstOrDefault()?.erp_tsrnSpareQty ?? 0;
                            x.erp_un_tsrnSpareQty = y.FirstOrDefault()?.erp_un_tsrnSpareQty ?? 0;
                            x.erp_snQty = y.FirstOrDefault()?.erp_snQty ?? 0;
                            x.erp_un_snQty = y.FirstOrDefault()?.erp_un_snQty ?? 0;
                            x.erp_srnQty = y.FirstOrDefault()?.erp_srnQty ?? 0;
                            x.erp_un_srnQty = y.FirstOrDefault()?.erp_un_srnQty ?? 0;
                            x.erp_snFreebieQty = y.FirstOrDefault()?.erp_snFreebieQty ?? 0;
                            x.erp_un_snFreebieQty = y.FirstOrDefault()?.erp_un_snFreebieQty ?? 0;
                            x.erp_srnFreebieQty = y.FirstOrDefault()?.erp_srnFreebieQty ?? 0;
                            x.erp_un_srnFreebieQty = y.FirstOrDefault()?.erp_un_srnFreebieQty ?? 0;
                            x.erp_snSpareQty = y.FirstOrDefault()?.erp_snSpareQty ?? 0;
                            x.erp_un_snSpareQty = y.FirstOrDefault()?.erp_un_snSpareQty ?? 0;
                            x.erp_srnSpareQty = y.FirstOrDefault()?.erp_srnSpareQty ?? 0;
                            x.erp_un_srnSpareQty = y.FirstOrDefault()?.erp_un_srnSpareQty ?? 0;
                            return x;
                        }).ToList();
                    deliverySchedules = deliverySchedules.OrderByDescending(x => x.SoId).ThenBy(x => x.SoSequence).ToList();
                }

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    data = deliverySchedules
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

        #region //GetDeliveryWipDetail -- 取得出貨排程製令資料 -- Zoey 2022.11.23
        public string GetDeliveryWipDetail(int SoDetailId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {

                    sql = @"SELECT b.MtlItemNo, b.MtlItemName, b.MtlItemSpec, a.WoErpPrefix + '-' + a.WoErpNo ErpFullNo
                            FROM MES.WipOrder a
                            INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
                            WHERE a.MtlItemId IN (
                                SELECT z.MtlItemId
                                FROM SCM.SoDetail z
                                WHERE SoDetailId = @SoDetailId
                            )
                            AND a.WoStatus NOT IN ('4', '5')
                            AND a.ConfirmStatus = 'Y'
                            ORDER BY a.DocDate DESC";
                    dynamicParameters.Add("SoDetailId", SoDetailId);

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

        #region //GetDeliveryFinalize -- 取得出貨定版資料 -- Zoey 2022.09.02
        public string GetDeliveryFinalize(string StartDate, string EndDate,string ShipmentCustomer,string MtlItemNo,string MtlItemName
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
           {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.DoDetailId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.SoDetailId, a.DoSequence, a.DoQty RegularQty, a.PcDoDetailRemark
                        , ISNULL(a.FreebieQty, 0) FreebieQty, ISNULL(a.SpareQty, 0) SpareQty
                        , CASE WHEN a.DeliveryProcess = '-1' THEN '' ELSE ISNULL(a.DeliveryProcess,'') END DeliveryProcess
                        , CASE WHEN a.DeliveryMethod = '-1' THEN '' ELSE ISNULL(a.DeliveryMethod,'') END DeliveryMethod
                        , CASE WHEN a.OrderSituation = '-1' THEN '' ELSE ISNULL(a.OrderSituation,'') END OrderSituation
                        , b.DoId, b.CompanyId, b.DepartmentId, b.UserId, b.DoErpPrefix, b.DoErpNo
                        , FORMAT(b.DoDate, 'yyyy-MM-dd') DoDate, FORMAT(b.DoDate, 'HH:mm:ss') DoTime
                        , b.DocDate, b.CustomerId, b.DcId, b.WayBill, b.Traffic, b.DoAddressFirst, b.DoAddressSecond
                        , b.DoRemark, b.WareHouseDoRemark, b.MeasureMailStatus, b.ConfirmStatus, b.ConfirmUserId
                        , b.TransferStatus, b.TransferDate, b.Status, b.DoErpPrefix + '-' + b.DoErpNo DoErpFullNo
                        , c.SoQty SoRegularQty, ISNULL(c.FreebieQty, '0') SoFreebieQty, ISNULL(c.SpareQty, '0') SoSpareQty, c.SoSequence
                        , d.SoErpPrefix + '-' + d.SoErpNo SoErpFullNo, ISNULL(d.CustomerPurchaseOrder, '') CustomerPurchaseOrder
                        , e.MtlItemNo, e.MtlItemName
                        , ISNULL(f.DcName,'') CustomerName, ISNULL(f.DcShortName,'') CustomerShortName
                        , ISNULL(g.TypeName, '') DeliveryProcessName   
                        , ISNULL(h.PickQty, 0) ExistPickQty
                        , ISNULL(i.PickRegularQty, 0) PickRegularQty
                        , ISNULL(j.PickFreebieQty, 0) PickFreebieQty
                        , ISNULL(k.PickSpareQty, 0) PickSpareQty";
                    sqlQuery.mainTables =
                          @"FROM SCM.DoDetail a
                            INNER JOIN SCM.DeliveryOrder b ON b.DoId = a.DoId
                            INNER JOIN SCM.SoDetail c ON c.SoDetailId = a.SoDetailId
                            INNER JOIN SCM.SaleOrder d ON d.SoId = c.SoId
                            LEFT JOIN PDM.MtlItem e ON e.MtlItemId = c.MtlItemId
                            LEFT JOIN SCM.DeliveryCustomer f ON f.DcId = b.DcId
                            LEFT JOIN BAS.[Type] g ON a.DeliveryProcess = g.TypeNo AND g.TypeSchema = 'DeliveryProcess'
                            OUTER APPLY(
                                SELECT SUM(x.ItemQty) PickQty
                                FROM SCM.PickingItem x
                                INNER JOIN SCM.DeliveryOrder y ON x.DoId = y.DoId
                                WHERE x.SoDetailId = a.SoDetailId
                            ) h
                            OUTER APPLY(
                                SELECT SUM(za.ItemQty) PickRegularQty
                                FROM SCM.PickingItem za
                                WHERE za.SoDetailId = c.SoDetailId
                                AND za.ItemType = 1
                                AND za.DoId = a.DoId
                            ) i
                            OUTER APPLY(
                                SELECT SUM(zb.ItemQty) PickFreebieQty
                                FROM SCM.PickingItem zb
                                WHERE zb.SoDetailId = c.SoDetailId
                                AND zb.ItemType = 2
                                AND zb.DoId = a.DoId
                            ) j
                            OUTER APPLY(
                                SELECT SUM(zc.ItemQty) PickSpareQty
                                FROM SCM.PickingItem zc
                                WHERE zc.SoDetailId = c.SoDetailId
                                AND zc.ItemType = 3
                                AND zc.DoId = a.DoId
                            ) k";
                    sqlQuery.auxTables = "";
                    string queryCondition = @" AND b.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    if (StartDate.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND b.DoDate >= @StartDate ", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    if (EndDate.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND b.DoDate <= @EndDate ", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ShipmentCustomer", @" AND f.DcShortName LIKE '%' +  @ShipmentCustomer +'%'", ShipmentCustomer);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemNo", @" AND e.MtlItemNo LIKE '%' +  @MtlItemNo +'%'", MtlItemNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemName", @" AND e.MtlItemName LIKE '%' +  @MtlItemName +'%'", MtlItemName);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "b.DoDate DESC, b.DoId";
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

        #region //GetDeliveryDateLog -- 取得交期歷史紀錄 -- Zoey 2022.09.15
        public string GetDeliveryDateLog(int SoDetailId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    string declareSql = @"DECLARE @rowsAdded int

                                          DECLARE @deliveryDateLog TABLE
                                          ( 
                                            DeliveryDateLogId int,
                                            ParentLogId int,
                                            SoDetailId int,
                                            PcPromiseDate DATETIME,
                                            PcPromiseTime DATETIME,
                                            OriginalDate DATETIME,
                                            OriginalTime DATETIME,
                                            DepartmentId int,
                                            SupervisorId int,
                                            CauseType int,
                                            CauseDescription nvarchar(255),
                                            processed int DEFAULT(0)
                                          )

                                          INSERT @deliveryDateLog
                                              SELECT a.DeliveryDateLogId, ISNULL(a.ParentLogId, -1) ParentLogId, a.SoDetailId
                                              , b.PcPromiseDate, b.PcPromiseDate PcPromiseTime 
                                              , a.PcPromiseDate OriginalDate, a.PcPromiseDate OriginalTime
                                              , a.DepartmentId, a.SupervisorId, a.CauseType, a.CauseDescription, 0
                                              FROM SCM.DeliveryDateLog a
                                              INNER JOIN SCM.SoDetail b ON b.SoDetailId = a.SoDetailId
                                              WHERE a.DeliveryDateLogId = (
                                                  SELECT ISNULL(MAX(x.DeliveryDateLogId), -1) AS MaxDeliveryDateLogId
                                                  FROM SCM.DeliveryDateLog x
                                                  WHERE x.SoDetailId = @SoDetailId
                                              )

                                          SET @rowsAdded=@@rowcount

                                          WHILE @rowsAdded > 0
                                          BEGIN

                                            UPDATE @deliveryDateLog SET processed = 1 WHERE processed = 0

                                            INSERT @deliveryDateLog
                                              SELECT b.DeliveryDateLogId, ISNULL(b.ParentLogId, -1) ParentLogId, a.SoDetailId
                                              , a.OriginalDate PcPromiseDate,  a.OriginalTime PcPromiseTime
                                              , b.PcPromiseDate OriginalDate,  b.PcPromiseDate OriginalTime
                                              , b.DepartmentId, b.SupervisorId, b.CauseType, b.CauseDescription, 0
                                              FROM @deliveryDateLog a
                                              INNER JOIN SCM.DeliveryDateLog b ON b.DeliveryDateLogId = a.ParentLogId
                                              WHERE 1=1
                                              AND a.processed = 1

                                            SET @rowsAdded = @@rowcount

                                            UPDATE @deliveryDateLog SET processed = 2 WHERE processed = 1


                                          END;";

                    sqlQuery.mainKey = "a.DeliveryDateLogId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ParentLogId, a.SoDetailId
                          , FORMAT(a.PcPromiseDate,'yyyy-MM-dd') PcPromiseDate, FORMAT(a.OriginalDate,'yyyy-MM-dd') OriginalDate
                          , FORMAT(a.PcPromiseTime,'HH:mm:ss') PcPromiseTime, FORMAT(a.OriginalTime,'HH:mm:ss') OriginalTime
                          , a.DepartmentId, a.SupervisorId, a.CauseType, ISNULL(a.CauseDescription,'') CauseDescription
                          , ISNULL(b.DepartmentName,'') DepartmentName
                          , ISNULL(c.UserNo + ' ' + c.UserName ,'') SupervisorName
                          , ISNULL(d.TypeName,'') CauseTypeName";
                    sqlQuery.mainTables =
                          @"FROM @deliveryDateLog a
                            LEFT JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                            LEFT JOIN BAS.[User] c ON a.SupervisorId = c.UserId
                            LEFT JOIN BAS.[Type] d ON a.CauseType = d.TypeNo AND d.TypeSchema = 'Delivery.CauseType'";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    sqlQuery.declarePart = declareSql;
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SoDetailId", @" AND a.SoDetailId = @SoDetailId", SoDetailId);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.DeliveryDateLogId";
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

        #region //GetDeliveryHistory -- 取得出貨歷史紀錄 -- Zoey 2022.09.27
        public string GetDeliveryHistory(string StartDate, string EndDate, string CustomerIds, string DcIds
            , string SoErpFullNo, string ItemValue,string MtlItemNo,string MtlItemName
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.DoDetailId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", FORMAT(b.DoDate, 'yyyy-MM-dd') DoDate, c.DcShortName, b.DcId
                        , b.[Status], g.StatusName DoStatus
                        , ISNULL(d.SoMtlItemName, e.MtlItemName) MtlItemName, ISNULL(d.SoMtlItemSpec, e.MtlItemSpec) MtlItemSpec
                        , e.MtlItemNo,e.LotManagement
                        , i.TypeName DeliveryProcess
                        , ISNULL(i1.PickRegularQty, 0) PickRegularQty, ISNULL(i2.PickRegularBarcodeQty, 0) PickRegularBarcodeQty
                        , ISNULL(i3.PickRegularNonBarcodeQty, 0) PickRegularNonBarcodeQty, ISNULL(i4.PickRegularSubstituteQty, 0) PickRegularSubstituteQty

                        , ISNULL(j1.PickFreebieQty, 0) PickFreebieQty, ISNULL(j2.PickFreebieBarcodeQty, 0) PickFreebieBarcodeQty
                        , ISNULL(j3.PickFreebieNonBarcodeQty, 0) PickFreebieNonBarcodeQty, ISNULL(j4.PickFreebieSubstituteQty, 0) PickFreebieSubstituteQty
                        , ISNULL(k1.PickSpareQty, 0) PickSpareQty, ISNULL(k2.PickSpareBarcodeQty, 0) PickSpareBarcodeQty
                        , ISNULL(k3.PickSpareNonBarcodeQty, 0) PickSpareNonBarcodeQty, ISNULL(k4.PickSpareSubstituteQty, 0) PickSpareSubstituteQty

                        , a.DoQty RegularQty, ISNULL(a.FreebieQty, '0') FreebieQty, ISNULL(a.SpareQty, '0') SpareQty, a.PcDoDetailRemark, d.SoDetailRemark
                        , h.SoErpPrefix + '-' + h.SoErpNo + '-' + d.SoSequence SoErpFullNo, ISNULL(h.CustomerPurchaseOrder, '') CustomerPurchaseOrder
                        , CASE WHEN a.OrderSituation = '-1' THEN '' ELSE ISNULL(a.OrderSituation,'') END OrderSituation
                        ";
                    sqlQuery.mainTables =
                        @"FROM SCM.DoDetail a
                        INNER JOIN SCM.DeliveryOrder b ON b.DoId = a.DoId
                        INNER JOIN SCM.DeliveryCustomer c ON b.DcId = c.DcId
                        INNER JOIN SCM.SoDetail d ON a.SoDetailId = d.SoDetailId
                        INNER JOIN PDM.MtlItem e ON d.MtlItemId = e.MtlItemId
                        INNER JOIN BAS.[Status] g ON b.[Status] = g.StatusNo AND g.StatusSchema = 'DeliveryOrder.Status'
                        INNER JOIN SCM.SaleOrder h ON d.SoId = h.SoId
                        LEFT JOIN BAS.[Type] i ON a.DeliveryProcess = i.TypeNo AND i.TypeSchema = 'DeliveryProcess'
                        OUTER APPLY (
                            SELECT SUM(x.ItemQty) PickRegularQty
                            FROM SCM.PickingItem x
                            WHERE x.SoDetailId = a.SoDetailId
                            AND x.DoId = a.DoId
                            AND x.ItemType = 1
                        ) i1
                        OUTER APPLY (
                            SELECT SUM(x.ItemQty) PickRegularBarcodeQty
                            FROM SCM.PickingItem x
                            WHERE x.SoDetailId = a.SoDetailId
                            AND x.DoId = a.DoId
                            AND x.ItemStatus = 'B'
                            AND x.ItemType = 1
                        ) i2
                        OUTER APPLY (
                            SELECT SUM(x.ItemQty) PickRegularNonBarcodeQty
                            FROM SCM.PickingItem x
                            WHERE x.SoDetailId = a.SoDetailId
                            AND x.DoId = a.DoId
                            AND x.ItemStatus = 'N'
                            AND x.ItemType = 1
                        ) i3
                        OUTER APPLY (
                            SELECT SUM(x.ItemQty) PickRegularSubstituteQty
                            FROM SCM.PickingItem x
                            WHERE x.SoDetailId = a.SoDetailId
                            AND x.DoId = a.DoId
                            AND x.ItemStatus = 'S'
                            AND x.ItemType = 1
                        ) i4


                        OUTER APPLY (
                            SELECT SUM(x.ItemQty) PickFreebieQty
                            FROM SCM.PickingItem x
                            WHERE x.SoDetailId = a.SoDetailId
                            AND x.DoId = a.DoId
                            AND x.ItemType = 2
                        ) j1
                        OUTER APPLY (
                            SELECT SUM(x.ItemQty) PickFreebieBarcodeQty
                            FROM SCM.PickingItem x
                            WHERE x.SoDetailId = a.SoDetailId
                            AND x.DoId = a.DoId
                            AND x.ItemStatus = 'B'
                            AND x.ItemType = 2
                        ) j2
                        OUTER APPLY (
                            SELECT SUM(x.ItemQty) PickFreebieNonBarcodeQty
                            FROM SCM.PickingItem x
                            WHERE x.SoDetailId = a.SoDetailId
                            AND x.DoId = a.DoId
                            AND x.ItemStatus = 'N'
                            AND x.ItemType = 2
                        ) j3
                        OUTER APPLY (
                            SELECT SUM(x.ItemQty) PickFreebieSubstituteQty
                            FROM SCM.PickingItem x
                            WHERE x.SoDetailId = a.SoDetailId
                            AND x.DoId = a.DoId
                            AND x.ItemStatus = 'S'
                            AND x.ItemType = 2
                        ) j4


                        OUTER APPLY (
                            SELECT SUM(x.ItemQty) PickSpareQty
                            FROM SCM.PickingItem x
                            WHERE x.SoDetailId = a.SoDetailId
                            AND x.DoId = a.DoId
                            AND x.ItemType = 3
                        ) k1
                        OUTER APPLY (
                            SELECT SUM(x.ItemQty) PickSpareBarcodeQty
                            FROM SCM.PickingItem x
                            WHERE x.SoDetailId = a.SoDetailId
                            AND x.DoId = a.DoId
                            AND x.ItemStatus = 'B'
                            AND x.ItemType = 3
                        ) k2
                        OUTER APPLY (
                            SELECT SUM(x.ItemQty) PickSpareNonBarcodeQty
                            FROM SCM.PickingItem x
                            WHERE x.SoDetailId = a.SoDetailId
                            AND x.DoId = a.DoId
                            AND x.ItemStatus = 'N'
                            AND x.ItemType = 3
                        ) k3
                        OUTER APPLY (
                            SELECT SUM(x.ItemQty) PickSpareSubstituteQty
                            FROM SCM.PickingItem x
                            WHERE x.SoDetailId = a.SoDetailId
                            AND x.DoId = a.DoId
                            AND x.ItemStatus = 'S'
                            AND x.ItemType = 3
                        ) k4";
                    sqlQuery.auxTables = "";
                    string queryCondition = @" AND b.CompanyId = @CompanyId ";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND b.DoDate >= @StartDate ", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND b.DoDate <= @EndDate ", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SoErpFullNo", @" AND (h.SoErpPrefix + '-' + h.SoErpNo + '-' + d.SoSequence) = @SoErpFullNo ", SoErpFullNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ItemValue", @" AND EXISTS (
                                                                                                                SELECT TOP 1 1
                                                                                                                FROM SCM.PickingItem m
                                                                                                                INNER JOIN MES.Barcode n ON n.BarcodeId = m.BarcodeId
                                                                                                                INNER JOIN MES.BarcodeAttribute o ON o.BarcodeId = n.BarcodeId
                                                                                                                WHERE m.SoDetailId = a.SoDetailId 
                                                                                                                AND o.ItemValue LIKE '%' + @ItemValue + '%')", ItemValue);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemNo", @" AND e.MtlItemNo LIKE '%' + @MtlItemNo +'%' ", MtlItemNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemName", @" AND e.MtlItemName LIKE '%' +  @MtlItemName +'%'", MtlItemName);
                    if (CustomerIds.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerIds", @" AND b.CustomerId IN @CustomerIds", CustomerIds.Split(','));
                    if (DcIds.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DcIds", @" AND b.DcId IN @DcIds", DcIds.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "b.DoDate DESC, b.DcId";
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

        #region //GetWipLink -- 取得製令綁定資料 -- Ben Ma 2023.04.26
        public string GetWipLink(int SoDetailId, string SearchKey)
        {
            try
            {
                string mtlItemNo = "", soErpPrefix = "", soErpNo = "", soSequence = "";
                List<WipLink> wipLinks = new List<WipLink>();

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

                    #region //判斷訂單資料是否正確
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT c.MtlItemNo, b.SoErpPrefix, b.SoErpNo, a.SoSequence
                            FROM SCM.SoDetail a
                            INNER JOIN SCM.SaleOrder b ON a.SoId = b.SoId
                            INNER JOIN PDM.MtlItem c ON a.MtlItemId = c.MtlItemId
                            WHERE a.ConfirmStatus = @ConfirmStatus
                            AND b.ConfirmStatus = @ConfirmStatus
                            AND a.SoDetailId = @SoDetailId";
                    dynamicParameters.Add("ConfirmStatus", "Y");
                    dynamicParameters.Add("SoDetailId", SoDetailId);

                    var result2 = sqlConnection.Query(sql, dynamicParameters);
                    if (result2.Count() <= 0) throw new SystemException("【訂單】資料錯誤!");

                    foreach (var item in result2)
                    {
                        mtlItemNo = item.MtlItemNo;
                        soErpPrefix = item.SoErpPrefix;
                        soErpNo = item.SoErpNo;
                        soSequence = item.SoSequence;
                    }
                    #endregion
                }

                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                {
                    #region //製令資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.TA001 WoErpPrefix, a.TA002 WoErpNo, a.TA015 PlanQty
                            , ISNULL(a.TA026, '') SoErpPrefix, ISNULL(a.TA027, '') SoErpNo
                            , ISNULL(a.TA028, '') SoSequence, 'N' LinkStatus
                            , CASE WHEN LEN(a.TA026 + a.TA027 + a.TA028) > 0 
                                THEN 'Y'
                                ELSE 'N' END BindSoStatus
                            FROM MOCTA a
                            WHERE a.TA013 = @ConfirmStatus
                            AND a.TA006 = @MtlItemNo
                            ORDER BY a.TA002 DESC, a.TA001";
                    dynamicParameters.Add("ConfirmStatus", "Y");
                    dynamicParameters.Add("MtlItemNo", mtlItemNo);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SearchKey", @" AND (a.TA001 + '-' + a.TA002) LIKE '%' + @SearchKey + '%'", SearchKey);

                    wipLinks = sqlConnection.Query<WipLink>(sql, dynamicParameters).ToList();

                    wipLinks
                        .Where(x => x.SoErpPrefix == soErpPrefix && x.SoErpNo == soErpNo && x.SoSequence == soSequence)
                        .ToList()
                        .ForEach(x =>
                        {
                            x.LinkStatus = "Y";
                        });
                    #endregion
                }

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    data = wipLinks
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

        #region //GetDailyDeliveryWipUnLinkDetail -- 每日出貨未綁定製令明細 -- Ben Ma 2023.07.04
        public string GetDailyDeliveryWipUnLinkDetail(string CompanyNo, string StartDate, string EndDate)
        {
            try
            {
                int CompanyId = -1;
                List<DailyDeliveryWipUnLinkDetail> dailyDeliveryWipUnLinkDetails = new List<DailyDeliveryWipUnLinkDetail>();
                List<WipLink> wipLinks = new List<WipLink>();

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //判斷公司資料是否正確
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 CompanyId, ErpDb
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

                    #region //出貨未綁定製令明細
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT FORMAT(b.DoDate, 'yyyy-MM-dd') DoDate, c.DcShortName
                            , e.SoErpPrefix + '-' + e.SoErpNo + '-' + d.SoSequence SoErpFullNo
                            , d.SoQty
                            , f.MtlItemNo, ISNULL(d.SoMtlItemName, f.MtlItemName) MtlItemName, ISNULL(d.SoMtlItemSpec, f.MtlItemSpec) MtlItemSpec
                            , g.UserNo PcUserNo, g.UserName PcUserName
                            , ISNULL(h.PickQty, 0) PickQty
                            , i.DepartmentName
                            FROM SCM.DoDetail a
                            INNER JOIN SCM.DeliveryOrder b ON a.DoId = b.DoId
                            INNER JOIN SCM.DeliveryCustomer c ON b.DcId = c.DcId
                            INNER JOIN SCM.SoDetail d ON a.SoDetailId = d.SoDetailId
                            INNER JOIN SCM.SaleOrder e ON d.SoId = e.SoId
                            INNER JOIN PDM.MtlItem f ON d.MtlItemId = f.MtlItemId
                            INNER JOIN BAS.[User] g ON b.CreateBy = g.UserId
                            OUTER APPLY (
                                SELECT SUM(x.ItemQty) PickQty
                                FROM SCM.PickingItem x
                                WHERE x.SoDetailId = a.SoDetailId
                                AND x.DoId = a.DoId
                            ) h
                            INNER JOIN BAS.[Department] i ON g.DepartmentId = i.DepartmentId
                            WHERE b.CompanyId = @CompanyId
                            AND b.DoDate >= @StartDate
                            AND b.DoDate <= @EndDate
                            AND c.DcId NOT IN @DcId";
                    dynamicParameters.Add("CompanyId", CompanyId);
                    dynamicParameters.Add("StartDate", Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00"));
                    dynamicParameters.Add("EndDate", Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59"));
                    dynamicParameters.Add("DcId", new string[] { "16" });

                    dailyDeliveryWipUnLinkDetails = sqlConnection.Query<DailyDeliveryWipUnLinkDetail>(sql, dynamicParameters).ToList();
                    #endregion
                }

                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                {
                    #region //ERP製令綁訂定單資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.TA001 WoErpPrefix, a.TA002 WoErpNo
                            , ISNULL(a.TA026, '') SoErpPrefix, ISNULL(a.TA027, '') SoErpNo
                            , ISNULL(a.TA028, '') SoSequence
                            , CASE WHEN LEN(a.TA026 + a.TA027 + a.TA028) > 0 
                                THEN 'Y'
                                ELSE 'N' END BindSoStatus
                            FROM MOCTA a
                            WHERE a.TA026 + '-' + a.TA027 + '-' + a.TA028 IN @SoErpFullNo
                            ORDER BY a.TA002 DESC, a.TA001";
                    dynamicParameters.Add("SoErpFullNo", dailyDeliveryWipUnLinkDetails.Select(x => x.SoErpFullNo).ToArray());

                    wipLinks = sqlConnection.Query<WipLink>(sql, dynamicParameters).ToList();
                    dailyDeliveryWipUnLinkDetails = dailyDeliveryWipUnLinkDetails.GroupJoin(wipLinks, x => x.SoErpFullNo, y => y.SoErpPrefix + '-' + y.SoErpNo + '-' + y.SoSequence, (x, y) => { x.WipLink = y.FirstOrDefault()?.BindSoStatus == "Y" ? true : false; return x; }).ToList();
                    #endregion
                }

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    data = dailyDeliveryWipUnLinkDetails
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

        #region //GetDeliveryLetteringInfo -- 取得出貨歷史查詢之刻號資料 -- Yi 2023.07.19
        public string GetDeliveryLetteringInfo(int DoDetailId, string ItemType)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT DISTINCT a.PickingCartonId, a.SoDetailId, f.DoDetailId, a.ItemType
                            , b.CartonName, b.CartonBarcode
                            , d.MtlItemNo, d.MtlItemName
                            , ISNULL((
                                SELECT aa.ItemType, aa.ItemQty, ISNULL(ab.BarcodeNo, '') ItemBarcode, ISNULL(ac.ItemValue, '') LetteringNo
                                , ad.TypeName ItemTypeName
                                , ISNULL(aa.LotNumber, '') LotNumber
								, ISNULL(ae.TrayNo, '') TrayNo
                                FROM SCM.PickingItem aa
                                LEFT JOIN MES.Barcode ab ON ab.BarcodeId = aa.BarcodeId
                                OUTER APPLY (
                                    SELECT aca.ItemValue
                                    FROM MES.BarcodeAttribute aca
                                    WHERE aca.ItemNo = 'Lettering'
                                    AND aca.BarcodeId = aa.BarcodeId
                                ) ac
                                INNER JOIN BAS.[Type] ad ON aa.ItemType = ad.TypeNo AND ad.TypeSchema = 'PickingItem.ProductType'
                                LEFT JOIN MES.Tray ae ON ab.BarcodeNo = ae.BarcodeNo
                                WHERE 1=1
                                AND ISNULL(aa.PickingCartonId, -999) = ISNULL(a.PickingCartonId, -999)
                                AND aa.SoDetailId = a.SoDetailId
                                AND aa.ItemType = 1
                                FOR JSON PATH, ROOT('data')
                            ) ,'') RegularItemDetail
                            , ISNULL((
                                SELECT aa.ItemType, aa.ItemQty, ISNULL(ab.BarcodeNo, '') ItemBarcode, ISNULL(ac.ItemValue, '') LetteringNo
                                , ad.TypeName ItemTypeName
                                FROM SCM.PickingItem aa
                                LEFT JOIN MES.Barcode ab ON ab.BarcodeId = aa.BarcodeId
                                OUTER APPLY (
                                    SELECT aca.ItemValue
                                    FROM MES.BarcodeAttribute aca
                                    WHERE aca.ItemNo = 'Lettering'
                                    AND aca.BarcodeId = aa.BarcodeId
                                ) ac
                                INNER JOIN BAS.[Type] ad ON aa.ItemType = ad.TypeNo AND ad.TypeSchema = 'PickingItem.ProductType'
                                WHERE 1=1
                                AND ISNULL(aa.PickingCartonId, -999) = ISNULL(a.PickingCartonId, -999)
                                AND aa.SoDetailId = a.SoDetailId
                                AND aa.ItemType = 2
                                FOR JSON PATH, ROOT('data')
                            ) ,'') FreebieItemDetail
                            , ISNULL((
                                SELECT aa.ItemType, aa.ItemQty, ISNULL(ab.BarcodeNo, '') ItemBarcode, ISNULL(ac.ItemValue, '') LetteringNo
                                , ad.TypeName ItemTypeName
                                FROM SCM.PickingItem aa
                                LEFT JOIN MES.Barcode ab ON ab.BarcodeId = aa.BarcodeId
                                OUTER APPLY (
                                    SELECT aca.ItemValue
                                    FROM MES.BarcodeAttribute aca
                                    WHERE aca.ItemNo = 'Lettering'
                                    AND aca.BarcodeId = aa.BarcodeId
                                ) ac
                                INNER JOIN BAS.[Type] ad ON aa.ItemType = ad.TypeNo AND ad.TypeSchema = 'PickingItem.ProductType'
                                WHERE 1=1
                                AND ISNULL(aa.PickingCartonId, -999) = ISNULL(a.PickingCartonId, -999)
                                AND aa.SoDetailId = a.SoDetailId
                                AND aa.ItemType = 3
                                FOR JSON PATH, ROOT('data')
                            ) ,'') SpareItemDetail
                            FROM SCM.PickingItem a
                            INNER JOIN SCM.DoDetail f ON f.SoDetailId = a.SoDetailId AND f.DoId = a.DoId
                            LEFT JOIN SCM.PickingCarton b ON b.PickingCartonId = a.PickingCartonId
                            LEFT JOIN SCM.SoDetail c ON c.SoDetailId = a.SoDetailId
                            LEFT JOIN PDM.MtlItem d ON d.MtlItemId = c.MtlItemId
                            WHERE 1=1";

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "DoDetailId", @" AND f.DoDetailId = @DoDetailId", DoDetailId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ItemType", @" AND a.ItemType = @ItemType", ItemType);
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

        #region //GetPickingCartonBarcode -- 取得箱號內的條碼 -- Shintokuro 2024.11.27
        public string GetPickingCartonBarcode(string CartonBarcode)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT c.BarcodeNo, 
                                   CAST(c.BarcodeQty AS INT) AS BarcodeQty
                            FROM SCM.PickingCarton a
                            INNER JOIN SCM.PickingItem b ON a.PickingCartonId = b.PickingCartonId
                            INNER JOIN MES.Barcode c ON b.BarcodeId = c.BarcodeId
                            WHERE a.CartonBarcode = @CartonBarcode
                            ORDER BY BarcodeNo ASC";
                    dynamicParameters.Add("CartonBarcode", CartonBarcode);

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

        #region //GetDoDetail -- 取得出貨排程資料 -- Ann 2025-05-19
        public string GetDoDetail(string DeliveryStatus, string DoErpFullNo, string SoIds, string SoErpFullNo, string CustomerMtlItemNo, int CustomerId
            , string MtlItemNo, int SalesmenId, string StartDate, string EndDate
            , string OrderBy, int PageIndex, int PageSize)
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

                    var CompanyResult = sqlConnection.Query(sql, dynamicParameters);
                    if (CompanyResult.Count() <= 0) throw new SystemException("【公司】資料錯誤!");

                    foreach (var item in CompanyResult)
                    {
                        ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                    }
                    #endregion

                    sqlQuery.mainKey = "a.DoDetailId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.DoId, a.SoDetailId, a.DoSequence, a.TransInInventoryId, a.DoQty, ISNULL(a.FreebieQty, 0) FreebieQty, ISNULL(a.SpareQty, 0) SpareQty, a.UnitPrice, a.Amount
                        , a.DoDetailRemark, a.WareHouseDoDetailRemark, a.PcDoDetailRemark, a.DeliveryProcess, a.DeliveryRoutine, a.DeliveryDocument, a.OrderSituation
                        , a.DeliveryMethod, a.ConfirmStatus, a.ClosureStatus
                        , b.DoErpPrefix, b.DoErpNo, FORMAT(b.DoDate, 'yyyy-MM-dd') DoDate, FORMAT(b.DocDate, 'yyyy-MM-dd') DocDate, b.Status DoStatus
                        , c.SoQty, c.SoSequence
                        , d.SoErpPrefix, d.SoErpNo
                        , e.CustomerNo, e.CustomerName, e.CustomerShortName
                        , f.MtlItemNo, f.MtlItemName
                        , (ISNULL(g.ItemQty, 0) - ISNULL(j.TotalRmaQty, 0)) PickQty
                        , ISNULL(k.PickRegularQty, 0) PickRegularQty
                        , ISNULL(l.PickFreebieQty, 0) PickFreebieQty
                        , ISNULL(m.PickSpareQty, 0) PickSpareQty";
                    sqlQuery.mainTables =
                        @"FROM SCM.DoDetail a 
                        INNER JOIN SCM.DeliveryOrder b ON a.DoId = b.DoId
                        INNER JOIN SCM.SoDetail c ON a.SoDetailId = c.SoDetailId
                        INNER JOIN SCM.SaleOrder d ON c.SoId = d.SoId
                        INNER JOIN SCM.Customer e ON b.CustomerId = e.CustomerId
                        INNER JOIN PDM.MtlItem f ON c.MtlItemId = f.MtlItemId
                        OUTER APPLY(
                            SELECT x.SoDetailId, SUM(x.ItemQty) ItemQty
                            FROM SCM.PickingItem x
                            INNER JOIN SCM.DoDetail xa ON x.DoId = xa.DoId AND x.SoDetailId = xa.SoDetailId
                            WHERE xa.DoDetailId = a.DoDetailId
                            GROUP BY x.SoDetailId
                        ) g
                        OUTER APPLY(
                            SELECT SUM(x.TsnQty) TsnQty, SUM(x.ReturnQty) ReturnQty, SUM(x.SaleQty) SaleQty
                            FROM SCM.TsnDetail x
                            INNER JOIN SCM.TempShippingNote y ON x.TsnId = y.TsnId
                            WHERE x.SoDetailId = c.SoDetailId
                            AND x.ConfirmStatus = 'Y'
                            AND y.ConfirmStatus = 'Y'
                        ) h
                        OUTER APPLY(
                            SELECT TOP 1 x.BindStatus LastBindStatus
                            FROM SCM.WipLinkLog x
                            WHERE x.SoErpPrefix = d.SoErpPrefix
                            AND x.SoErpNo = d.SoErpNo
                            AND x.SoSequence = c.SoSequence
                            ORDER BY x.CreateDate DESC
                        ) i
                        OUTER APPLY (
                            SELECT SUM(ha.RmaQty) TotalRmaQty
                            FROM SCM.RmaDetail ha
                            INNER JOIN SCM.TsnDetail hb ON ha.TsnDetailId = hb.TsnDetailId
                            INNER JOIN SCM.ReturnMerchandiseAuthorization hc ON ha.RmaId = hc.RmaId
                            INNER JOIN SCM.TempShippingNote hd ON hb.TsnId = hd.TsnId
                            INNER JOIN SCM.DoDetail he ON he.DoId = hd.DoId AND hb.SoDetailId = he.SoDetailId
                            WHERE he.DoDetailId = a.DoDetailId
                            AND hc.ConfirmStatus = 'Y'
                        ) j
                        OUTER APPLY (
                            SELECT SUM(za.ItemQty) PickRegularQty
                            FROM SCM.PickingItem za
                            INNER JOIN SCM.DoDetail zb ON za.DoId = zb.DoId AND za.SoDetailId = zb.SoDetailId
                            WHERE zb.DoDetailId = a.DoDetailId 
                            AND za.ItemType = 1
                        ) k
                        OUTER APPLY (
                            SELECT SUM(zb.ItemQty) PickFreebieQty
                            FROM SCM.PickingItem zb
                            INNER JOIN SCM.DoDetail zc ON zb.DoId = zc.DoId AND zb.SoDetailId = zc.SoDetailId
                            WHERE zc.DoDetailId = a.DoDetailId
                            AND zb.ItemType = 2
                        ) l
                        OUTER APPLY (
                            SELECT SUM(zc.ItemQty) PickSpareQty
                            FROM SCM.PickingItem zc
                            INNER JOIN SCM.DoDetail zd ON zc.DoId = zd.DoId AND zc.SoDetailId = zd.SoDetailId
                            WHERE zd.DoDetailId = a.DoDetailId
                            AND zc.ItemType = 3
                        ) m";
                    sqlQuery.auxTables = "";
                    string queryCondition = @" AND b.CompanyId = @CompanyId AND c.ClosureStatus = 'N'";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    if (SoIds.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SoIds", @" AND c.SoId IN @SoIds", SoIds.Split(','));
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DoErpFullNo", @" AND b.DoErpPrefix + '-' + b.DoErpNo LIKE '%' + @DoErpFullNo + '%'", DoErpFullNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerId", @" AND b.CustomerId = @CustomerId", CustomerId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemNo", @" AND (f.MtlItemNo LIKE '%' + @MtlItemNo + '%' OR f.MtlItemName LIKE '%' + @MtlItemNo + '%' OR c.CustomerMtlItemNo LIKE '%' + @MtlItemNo + '%')", MtlItemNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SalesmenId", @" AND d.SalesmenId = @SalesmenId", SalesmenId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND b.DoDate >= @StartDate ", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND b.DoDate <= @EndDate ", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "CASE WHEN b.Status != 'S' THEN 0 ELSE 1 END, b.DoDate DESC, a.DoId, a.DoSequence";
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

        #region //GetSoDetail -- 取得訂單單身資料 -- Ann 2025-05-22
        public string GetSoDetail(int SoDetailId, int SoId, int MtlItemId, string SoErpFullNo, string CustomerMtlItemNo, string TransferStatus, string SearchKey, string MtlItemIdIsNull, int CompanyId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.SoDetailId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.SoId, a.SoSequence, a.MtlItemId, ISNULL(a.SoMtlItemName, d.MtlItemName) MtlItemName
                          , ISNULL(a.SoMtlItemSpec, d.MtlItemSpec) MtlItemSpec, a.InventoryId, a.UomId, a.SoQty
                          , a.SiQty, a.ProductType, a.FreebieQty, a.FreebieSiQty, a.SpareQty, a.SpareSiQty
                          , a.UnitPrice, a.Amount, a.Project
                          , ISNULL(FORMAT(a.PromiseDate, 'yyyy-MM-dd'), '') PromiseDate
                          , ISNULL(FORMAT(a.PcPromiseDate, 'yyyy-MM-dd'), '') PcPromiseDate
                          , ISNULL(FORMAT(a.PcPromiseDate, 'HH:mm:ss'), '') PcPromiseTime
                          , a.SoDetailRemark, a.SoPriceQty, a.SoPriceUomId, a.TaxNo
                          , a.BusinessTaxRate, a.DiscountRate, a.DiscountAmount, a.ConfirmStatus, a.ClosureStatus,a.QuotationErp
                          , b.SoErpPrefix, b.SoErpNo, b.TransferStatus, b.SoErpPrefix + '-' + b.SoErpNo SoErpFullNo, b.CustomerId
                          , c.CustomerNo, c.CustomerShortName, c.CustomerName
                          , ISNULL(d.MtlItemNo, '') MtlItemNo
                          , e.UomNo, f.UomNo SoPriceUomNo
                          , h.SoDetailTempId
                          , ISNULL(g.CustomerMtlItemNo, a.CustomerMtlItemNo) CustomerMtlItemNo
                          , (SELECT ISNULL(SUM(x.DoQty), 0) FROM SCM.DoDetail x WHERE x.SoDetailId = a.SoDetailId) TotalDoQty
                          , (SELECT ISNULL(SUM(x.FreebieQty), 0) FROM SCM.DoDetail x WHERE x.SoDetailId = a.SoDetailId) TotalFreebieQty
                          , (SELECT ISNULL(SUM(x.SpareQty), 0) FROM SCM.DoDetail x WHERE x.SoDetailId = a.SoDetailId) TotalSpareQty
                          , ISNULL(i.PickRegularQty, 0) PickRegularQty
                          , ISNULL(i.TotalPickDoQty, 0) TotalPickDoQty
                          , ISNULL(j.PickFreebieQty, 0) PickFreebieQty
                          , ISNULL(j.TotalPickFreebieQty, 0) TotalPickFreebieQty
                          , ISNULL(k.PickSpareQty, 0) PickSpareQty
                          , ISNULL(k.TotalPickSpareQty, 0) TotalPickSpareQty";
                    sqlQuery.mainTables =
                          @"FROM SCM.SoDetail a
                            INNER JOIN SCM.SaleOrder b ON a.SoId = b.SoId
                            INNER JOIN SCM.Customer c ON b.CustomerId = c.CustomerId
                            LEFT JOIN PDM.MtlItem d ON a.MtlItemId = d.MtlItemId
                            LEFT JOIN PDM.UnitOfMeasure e ON a.UomId = e.UomId
                            LEFT JOIN PDM.UnitOfMeasure f ON a.SoPriceUomId = f.UomId
                            LEFT JOIN SCM.SoDetailTemp h ON h.SoDetailId = a.SoDetailId
                            OUTER APPLY (
                                SELECT TOP 1 ga.CustomerMtlItemNo
                                FROM PDM.CustomerMtlItem ga
                                WHERE ga.CustomerId = c.CustomerId 
                                AND ga.MtlItemId = d.MtlItemId
                                ORDER BY ga.LastModifiedDate DESC
                            ) g
                            OUTER APPLY (
                                SELECT SUM(za.ItemQty) PickRegularQty
                                , SUM(zc.DoQty) TotalPickDoQty
                                FROM SCM.PickingItem za 
                                    INNER JOIN SCM.DeliveryOrder zb ON za.DoId = zb.DoId
                                    INNER JOIN SCM.DoDetail zc ON zb.DoId = zc.DoId AND za.SoDetailId = zc.SoDetailId
                                WHERE za.SoDetailId = a.SoDetailId
                                    AND za.ItemType = 1
                                    AND zb.Status = 'S'
                            ) i
                            OUTER APPLY (
                                SELECT SUM(za.ItemQty) PickFreebieQty
                                , SUM(zc.FreebieQty) TotalPickFreebieQty
                                FROM SCM.PickingItem za 
                                    INNER JOIN SCM.DeliveryOrder zb ON za.DoId = zb.DoId
                                    INNER JOIN SCM.DoDetail zc ON zb.DoId = zc.DoId AND za.SoDetailId = zc.SoDetailId
                                WHERE za.SoDetailId = a.SoDetailId
                                    AND za.ItemType = 2
                                    AND zb.Status = 'S'
                            ) j
                            OUTER APPLY (
                                SELECT SUM(za.ItemQty) PickSpareQty
                                , SUM(zc.SpareQty) TotalPickSpareQty
                                FROM SCM.PickingItem za 
                                    INNER JOIN SCM.DeliveryOrder zb ON za.DoId = zb.DoId
                                    INNER JOIN SCM.DoDetail zc ON zb.DoId = zc.DoId AND za.SoDetailId = zc.SoDetailId
                                WHERE za.SoDetailId = a.SoDetailId
                                    AND za.ItemType = 3
                                    AND zb.Status = 'S'
                            ) k";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    if (MtlItemIdIsNull == "Y") queryCondition += @" AND a.MtlItemId IS NULL";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SoDetailId", @" AND a.SoDetailId = @SoDetailId", SoDetailId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SoId", @" AND a.SoId = @SoId", SoId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemId", @" AND a.MtlItemId = @MtlItemId", MtlItemId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SoErpFullNo", @" AND (b.SoErpPrefix + '-' + b.SoErpNo) LIKE '%' + @SoErpFullNo + '%'", SoErpFullNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CustomerMtlItemNo", @" AND a.CustomerMtlItemNo LIKE '%' + @CustomerMtlItemNo + '%'", CustomerMtlItemNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TransferStatus", @" AND b.TransferStatus = @TransferStatus", TransferStatus);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CompanyId", @" AND b.CompanyId = @CompanyId", CompanyId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "SearchKey",
                          @" AND (d.MtlItemNo LIKE '%' + @SearchKey + '%' 
                          OR d.MtlItemName LIKE '%' + @SearchKey + '%' 
                          OR a.SoMtlItemName LIKE '%' + @SearchKey + '%' 
                          OR a.CustomerMtlItemNo LIKE '%' + @SearchKey + '%' 
                          OR g.CustomerMtlItemNo LIKE '%' + @SearchKey + '%')", SearchKey);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.SoId DESC, a.SoSequence";
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

        #region //GetSoDetailForDeliveryOrder -- 取得訂單單身資料(取得可定版資料用) -- Ann 2025-05-22
        public string GetSoDetailForDeliveryOrder(string SoErpFullNo)
        {
            try
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
                    }
                    #endregion

                    using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                    {
                        sql = @"SELECT a.TD001 SoErpPrefix, a.TD002 SoErpNo, a.TD003 SoSequence, a.TD008 SoQty, a.TD009 SiQty
                                , CASE WHEN LEN(LTRIM(RTRIM(a.TD048))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(a.TD048)) as date), 'yyyy-MM-dd') ELSE NULL END PcPromiseDate
                                , a.TD004 MtlItemNo, a.TD005 MtlItemName, a.TD006 MtlItemSpec, a.TD024 FreebieQty, a.TD050 SpareQty, a.TD020 SoDetailRemark
                                , a.TD001 + '-' + a.TD002 SoErpFullNo, a.TD021 ConfirmStatus, a.TD016 ClosureStatus
                                , b.TC004 CustomerNo
                                , c.MA002 CustomerShortName
                                , xa.TempShippingQty, xb.ReturnTempShippingQty
                                , xc.ShippingQty, xd.ReturnShippingQty
                                FROM COPTD a 
                                INNER JOIN COPTC b ON a.TD001 = b.TC001 AND a.TD002 = b.TC002
                                INNER JOIN COPMA c ON b.TC004 =  c.MA001
                                OUTER APPLY (
                                    SELECT ISNULL(SUM(TG009), 0) TempShippingQty
                                    FROM INVTG 
                                    WHERE TG014 = a.TD001
                                    AND TG015 = a.TD002
                                    AND TG016 = a.TD003
                                    AND TG022 = 'Y'
                                ) xa
                                OUTER APPLY (
                                    SELECT ISNULL(SUM(TI009), 0) ReturnTempShippingQty
                                    FROM INVTI
                                    WHERE TI014 = a.TD001
                                    AND TI015 = a.TD002
                                    AND TI016 = a.TD003
                                    AND TI022 = 'Y'
                                ) xb
                                OUTER APPLY (
                                    SELECT SUM(TH008) ShippingQty
                                    FROM COPTH 
                                    WHERE TH014 = a.TD001
                                    AND TH015 = a.TD002
                                    AND TH016 = a.TD003
                                    AND TH020 = 'Y'
                                ) xc
                                OUTER APPLY (
                                    SELECT ISNULL(SUM(TJ007), 0) ReturnShippingQty
                                    FROM COPTJ 
                                    WHERE TJ018 = a.TD001
                                    AND TJ019 = a.TD002
                                    AND TJ020 = a.TD003
                                    AND TJ021 = 'Y'
                                ) xd
                                WHERE a.TD001 + '-' + a.TD002 = @SoErpFullNo";

                        dynamicParameters.Add("SoErpFullNo", SoErpFullNo);

                        List<SoDetailForDelivery> result = sqlConnection2.Query<SoDetailForDelivery>(sql, dynamicParameters).ToList();

                        #region //取得BM資料
                        foreach (var item in result)
                        {
                            #region //取得訂單單身ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.SoDetailId
                                    FROM SCM.SoDetail a 
                                    INNER JOIN SCM.SaleOrder b ON a.SoId = b.SoId
                                    WHERE b.SoErpPrefix = @SoErpPrefix
                                    AND b.SoErpNo = @SoErpNo
                                    AND a.SoSequence = @SoSequence";
                            dynamicParameters.Add("SoErpPrefix", item.SoErpPrefix);
                            dynamicParameters.Add("SoErpNo", item.SoErpNo);
                            dynamicParameters.Add("SoSequence", item.SoSequence);

                            var soDetailIdResult = sqlConnection.Query(sql, dynamicParameters);

                            if (soDetailIdResult.Count() <= 0)
                            {
                                throw new SystemException($"查無訂單:{item.SoErpPrefix}-{item.SoErpNo}({item.SoSequence})!!");
                            }

                            item.SoDetailId = soDetailIdResult.FirstOrDefault().SoDetailId;
                            #endregion

                            #region //取得客戶ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.CustomerId
                                    FROM SCM.Customer a 
                                    WHERE a.CustomerNo = @CustomerNo
                                    AND a.CompanyId = @CompanyId";
                            dynamicParameters.Add("CustomerNo", item.CustomerNo);
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            var customerResult = sqlConnection.Query(sql, dynamicParameters);

                            if (customerResult.Count() <= 0)
                            {
                                throw new SystemException($"查無客戶:{item.CustomerNo}!!");
                            }

                            item.CustomerId = customerResult.FirstOrDefault().CustomerId;
                            #endregion

                            #region //取得此訂單單身的出貨單(單頭狀態未出貨)
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT ISNULL(SUM(a.DoQty), 0) TotalDoQty
                                    FROM SCM.DoDetail a 
                                    INNER JOIN SCM.DeliveryOrder b ON a.DoId = b.DoId
                                    WHERE a.SoDetailId = @SoDetailId
                                    AND b.[Status] != 'S'";
                            dynamicParameters.Add("SoDetailId", item.SoDetailId);

                            var doDetailQtyResult = sqlConnection.Query(sql, dynamicParameters);

                            if (doDetailQtyResult.Count() > 0)
                            {
                                item.TotalDoQty = doDetailQtyResult.FirstOrDefault().TotalDoQty;
                            }
                            #endregion
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
            }
            catch (Exception e)
            {
                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "errorForDA",
                    msg = e.Message
                });
                #endregion

                logger.Error(e.Message);
            }

            return jsonResponse.ToString();
        }
        #endregion

        #region //GetDeliveryOrderDateLog -- 取得出貨日歷史紀錄 -- Ann 2025-05-22
        public string GetDeliveryOrderDateLog(int DoDetailId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.DeliveryDateLogId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ParentLogId, a.SoDetailId, a.DoDetailId, a.PcPromiseDate, a.DoQty, a.OriDoQty
                        , FORMAT(a.DoDate, 'yyyy-MM-dd') DoDate, FORMAT(a.OriDoDate, 'yyyy-MM-dd') OriDoDate, a.DepartmentId, a.SupervisorId, a.CauseType, a.CauseDescription
                        , b.DepartmentNo, ISNULL(b.DepartmentName,'') DepartmentName
                        , ISNULL(c.UserNo + ' ' + c.UserName ,'') SupervisorName
                        , FORMAT(a.CreateDate, 'yyyy-MM-dd HH:mm:ss') CreateDate
                        , ISNULL(d.TypeName,'') CauseTypeName
                        , e.DoQty CurrentDoQty";
                    sqlQuery.mainTables =
                          @"FROM SCM.DeliveryDateLog a 
                            LEFT JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                            LEFT JOIN BAS.[User] c ON a.SupervisorId = c.UserId
                            LEFT JOIN BAS.[Type] d ON a.CauseType = d.TypeNo AND d.TypeSchema = 'Delivery.CauseType'
                            INNER JOIN SCM.DoDetail e ON a.DoDetailId = e.DoDetailId";
                    sqlQuery.auxTables = "";
                    string queryCondition = "";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DoDetailId", @" AND a.DoDetailId = @DoDetailId", DoDetailId);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.CreateDate DESC";
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

        #region //GetPendingOrders -- 取得出貨日歷史紀錄 -- Ann 2025-05-22
        public string GetPendingOrders(string SoErpFullNo, string SearchKey
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
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
                    }
                    #endregion

                    using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                    {
                        string mainSql = @"select case a.TC027
                                             when 'Y' then '已確認'
                                             when 'N' then '未確認'
                                             when 'V' then '作廢'
                                           end AS OrderStatus,
                                           a.TC001 AS SoErpPrefix,
                                           a.TC002 AS SoErpNo,
                                           a.TC001 + '-' + a.TC002 AS SoErpFullNo,
                                           CASE WHEN LEN(LTRIM(RTRIM(a.TC003))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(a.TC003)) as date), 'yyyy-MM-dd') ELSE NULL END OrderDate,
                                           h.MA002 AS UserName,
                                           FORMAT(convert(date,a.TC003), 'yyyyMM') AS OrderDateYearMonth,
                                           e.MA002 AS CustomerShortName,
                                           a.TC008 AS Currency,
                                           a.TC009 AS Exchange,
                                           b.TD003 AS SoSeq,
                                           b.TD004 AS MtlItemNo,
                                           b.TD005 AS MtlItemName,
                                           b.TD006 AS MtlItemSpec,
                                           CASE WHEN LEN(LTRIM(RTRIM(b.TD048))) > 0 THEN FORMAT(CAST(LTRIM(RTRIM(b.TD048)) as date), 'yyyy-MM-dd') ELSE NULL END PcPromiseDate,
                                           b.TD008 AS SoQty,
                                           b.TD011 AS UnitPrice,
                                           b.TD008 * a.TC009 * b.TD011 AS Amount,
                                           b.TD009 AS SiQty,
                                           (b.TD008 - b.TD009) * b.TD011 * a.TC009 AS NotSiAmount,
                                           c.ONHAND AS InventoryQty,
                                           ISNULL(c.ONHAND,0) * ISNULL(g.ItemCost,0) AS ItemCost,
                                           ISNULL(f.WipAmt,0) AS ProcessItemCost,
                                           xa.TempShippingQty - xb.ReturnTempShippingQty AS　TempShippingQty
                                      from COPTC a
                                           INNER JOIN COPTD b ON a.TC001 = b.TD001 and a.TC002 = b.TD002
                                           LEFT JOIN (select LA001,SUM(LA005 * LA011) ONHAND
                                                         from INVLA 
                                                        group by LA001) c ON b.TD004 = c.LA001
                                           LEFT JOIN (select TA006,sum(TA015 - TA017) WipQty
                                                        from MOCTA
                                                       where TA011 not in ('Y','y')
                                                       group by TA006) d ON b.TD004 = d.TA006
                                           INNER JOIN COPMA e ON a.TC004 = e.MA001
                                           LEFT JOIN (select a.TA006 AS AssemblyNo,
                                                             sum((b.TB005 - ((a.TA017 + a.TA018) * 
                                                               case when a.TA006 = b.TB003 then 1 else ISNULL(c.MD006,1) / 
                                                                case when c.MD007 is null then 1
                                                                     when c.MD007 = 0 then 1
                                                                     else ISNULL(c.MD007,1) end end
                                                               )) * 
                                                             e.ItemCost) WipAmt
                                                        from MOCTA a
                                                             INNER JOIN MOCTB b ON a.TA001 = b.TB001 AND a.TA002 = b.TB002
                                                             LEFT JOIN BOMMD c ON c.MD001 = a.TA006 AND c.MD003 = b.TB003
                                                             INNER JOIN INVMB d ON b.TB003 = d.MB001
                                                             INNER JOIN (SELECT LB001,
                                                                                   LB002,
                                                                                LB010 AS ItemCost
                                                                           FROM INVLB aa
                                                                          WHERE LB002 = '202504'
                                                                          --LB002 = (
                                                                             --  SELECT MAX(LB002)
                                                                             --  FROM INVLB aa1
                                                                             --  WHERE aa.LB001 = aa1.LB001
                                                                          -- )
                                                                           ) e ON e.LB001 = b.TB003
                                                       where a.TA011 not in('Y','y')
                                                       group by a.TA006) f ON f.AssemblyNo = b.TD004
                                           LEFT JOIN (SELECT LB001,
                                                             LB002,
                                                             LB010 AS ItemCost
                                                        FROM INVLB bb
                                                       WHERE LB002 = '202504'
                                                      -- LB002 = (
                                                         --    SELECT MAX(LB002)
                                                            --   FROM INVLB bb1
                                                            --  WHERE bb.LB001 = bb1.LB001
                                                      --)
                                                      ) g ON b.TD004 = g.LB001
                                           INNER JOIN DSCSYS.dbo.DSCMA h ON a.TC006 = h.MA001
                                            OUTER APPLY (
                                                SELECT ISNULL(SUM(TG009), 0) TempShippingQty
                                                FROM INVTG 
                                                WHERE TG014 = b.TD001
                                                AND TG015 = b.TD002
                                                AND TG016 = b.TD003
                                                AND TG022 = 'Y'
                                            ) xa
                                            OUTER APPLY (
                                                SELECT ISNULL(SUM(TI009), 0) ReturnTempShippingQty
                                                FROM INVTI
                                                WHERE TI014 = b.TD001
                                                AND TI015 = b.TD002
                                                AND TI016 = b.TD003
                                                AND TI022 = 'Y'
                                            ) xb
                                     where (b.TD008 - b.TD009) * b.TD011 * a.TC009 > 0
                                       and b.TD008 - b.TD009 > 0
                                       and b.TD016 not in ('Y','y')
                                       and a.TC027 in ('Y')";

                        string whereClause = string.Empty;

                        if (!string.IsNullOrEmpty(SoErpFullNo))
                        {
                            whereClause += " AND (b.TD001 + '-' + b.TD002 +  '(' + b.TD003 + ')') LIKE @SoErpFullNo";
                            dynamicParameters.Add("SoErpFullNo", $"%{SoErpFullNo}%");
                        }

                        if (!string.IsNullOrEmpty(SearchKey))
                        {
                            whereClause += " AND (b.TD004 LIKE '%' + @SearchKey + '%' OR b.TD005 LIKE '%' + @SearchKey + '%')";
                            dynamicParameters.Add("SearchKey", SearchKey);
                        }

                        mainSql += whereClause;

                        string sql = mainSql + @" ORDER BY a.TC003 DESC
                                                OFFSET @PageSize * (@PageIndex - 1) ROWS
                                                FETCH NEXT @PageSize ROWS ONLY;";

                        dynamicParameters.Add("PageSize", PageSize);
                        dynamicParameters.Add("PageIndex", PageIndex);

                        List<SoDetailForErpDelivery> result = sqlConnection2.Query<SoDetailForErpDelivery>(sql, dynamicParameters).ToList();

                        // 計算總筆數
                        string countSql = "SELECT COUNT(1) FROM (" + mainSql + ") AS CountQuery";
                        var TotalCount = sqlConnection2.QuerySingle<int>(countSql, dynamicParameters);

                        #region //計算BM已定版但未轉暫出數量
                        for (int i = result.Count - 1; i >= 0; i--)
                        {
                            var item = result[i];

                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT ISNULL(SUM(a.DoQty), 0) TotalDoQty
                                    FROM SCM.DoDetail a 
                                    INNER JOIN SCM.DeliveryOrder b ON a.DoId = b.DoId
                                    INNER JOIN SCM.SoDetail c ON a.SoDetailId = c.SoDetailId
                                    INNER JOIN SCM.SaleOrder d ON c.SoId = d.SoId
                                    WHERE d.SoErpPrefix = @SoErpPrefix 
                                    AND d.SoErpNo = @SoErpNo
                                    AND c.SoSequence = @SoSeq
                                    AND b.[Status] != 'S'";

                            dynamicParameters.Add("SoErpPrefix", item.SoErpPrefix);
                            dynamicParameters.Add("SoErpNo", item.SoErpNo);
                            dynamicParameters.Add("SoSeq", item.SoSeq);

                            var doDetailResult = sqlConnection.Query(sql, dynamicParameters);

                            if (doDetailResult.Any())
                            {
                                item.TotalDoQty = doDetailResult.First().TotalDoQty;
                            }

                            double currentQty = item.SiQty > item.TempShippingQty
                                ? item.SoQty - item.SiQty - item.TotalDoQty
                                : item.SoQty - item.TempShippingQty - item.TotalDoQty;

                            if (currentQty <= 0)
                            {
                                result.RemoveAt(i);
                            }
                        }
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            result,
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
        #region //AddDeliveryDateLog -- 交期歷史紀錄新增 -- Zoey 2022.09.15
        public string AddDeliveryDateLog(int DoDetailId, int SoDetailId, string PcPromiseDate, string PcPromiseTime
            , int CauseType, int DepartmentId, int SupervisorId, string CauseDescription)
        {
            try
            {
                if (SoDetailId <= 0) throw new SystemException("【訂單資料錯誤!】");
                if (PcPromiseDate.Length <= 0) throw new SystemException("【交貨日期】不能為空!");
                if (PcPromiseTime.Length <= 0) throw new SystemException("【交貨時間】不能為空!");

                string dateNow = DateTime.Now.ToString("yyyyMMdd");
                string timeNow = DateTime.Now.ToString("HH:mm:ss");

                int rowsAffected = 0;

                DateTime PcPromiseDateTime = Convert.ToDateTime(PcPromiseDate + " " + PcPromiseTime);

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    string modifier = "", soErpPrefix = "", soErpNo = "", soSequence = "";

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

                        #region //判斷出貨單是否已產出暫出單
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ToId
                                FROM SCM.TemporaryOrder a
                                INNER JOIN SCM.DeliveryOrder b ON b.DoId = a.DoId 
                                INNER JOIN SCM.DoDetail c ON c.DoId = b.DoId
                                WHERE c.DoDetailId = @DoDetailId";
                        dynamicParameters.Add("DoDetailId", DoDetailId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        #endregion

                        if (result2.Count() <= 0)
                        {

                            #region //判斷交期修改次數
                            sql = @"SELECT DeliveryDateLogId
                                    FROM SCM.DeliveryDateLog
                                    WHERE SoDetailId = @SoDetailId";
                            dynamicParameters.Add("SoDetailId", SoDetailId);

                            var result3 = sqlConnection.Query(sql, dynamicParameters);
                            if (result3.Count() >= 2) {
                                if (CauseType <= 0) throw new SystemException("【原因類別】不能為空!");
                                if (DepartmentId <= 0) throw new SystemException("【責任部門】不能為空!");
                                if (SupervisorId <= 0) throw new SystemException("【責任主管】不能為空!");
                                if (CauseDescription.Length <= 0) throw new SystemException("【延遲原因】不能為空!");
                                if (CauseDescription.Length > 255) throw new SystemException("【延遲原因】長度錯誤!");
                            }
                            #endregion

                            #region //撈取原排定交貨日
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT PcPromiseDate
                                    FROM SCM.SoDetail
                                    WHERE SoDetailId = @SoDetailId";
                            dynamicParameters.Add("SoDetailId", SoDetailId);

                            var result4 = sqlConnection.Query(sql, dynamicParameters);

                            DateTime pcPromiseDate = new DateTime();
                            foreach (var item in result4)
                            {
                                pcPromiseDate = item.PcPromiseDate;
                            }
                            #endregion

                            #region //撈取原修改紀錄ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT ISNULL(MAX(DeliveryDateLogId), -1) AS ParentLogId
                                    FROM SCM.DeliveryDateLog
                                    WHERE SoDetailId = @SoDetailId";
                            dynamicParameters.Add("SoDetailId", SoDetailId);

                            var result5 = sqlConnection.Query(sql, dynamicParameters);

                            int ParentLogId = -1;
                            foreach (var item in result5)
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
                                    SoDetailId,
                                    PcPromiseDate = pcPromiseDate,
                                    CauseType = CauseType > 0 ? (int?)CauseType : null,
                                    DepartmentId = DepartmentId > 0 ? (int?)DepartmentId : null,
                                    SupervisorId = SupervisorId > 0 ? (int?)SupervisorId : null,
                                    CauseDescription = CauseDescription.Length > 0 ? CauseDescription : null,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            var logResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += logResult.Count();
                            #endregion

                            if (DoDetailId > 0)
                            {
                                #region //撈取原出貨單Id
                                sql = @"SELECT DoId
                                        FROM SCM.DeliveryOrder
                                        WHERE DoId = (
                                            SELECT DoId
                                            FROM SCM.DoDetail
                                            WHERE DoDetailId = @DoDetailId
                                        )";
                                dynamicParameters.Add("DoDetailId", DoDetailId);

                                var result6 = sqlConnection.Query(sql, dynamicParameters);
                                int oriDoId = -1;
                                foreach (var item in result6)
                                {
                                    oriDoId = Convert.ToInt32(item.DoId);
                                }
                                #endregion

                                #region //判斷新時段是否有同出貨客戶的出貨單
                                sql = @"SELECT a.DoId
                                        FROM SCM.DeliveryOrder a
                                        WHERE a.DoDate = @DoDate
                                        AND a.DcId = (
                                            SELECT y.DcId
                                            FROM SCM.DoDetail z
                                            INNER JOIN SCM.DeliveryOrder y ON z.DoId = y.DoId
                                            WHERE z.DoDetailId = @DoDetailId
                                        )";
                                dynamicParameters.Add("DoDate", PcPromiseDateTime);
                                dynamicParameters.Add("DoDetailId", DoDetailId);

                                var result7 = sqlConnection.Query(sql, dynamicParameters);

                                int doId = -1;
                                foreach (var item in result7)
                                {
                                    doId = item.DoId;
                                }
                                #endregion

                                if (doId < 0)
                                {
                                    #region //複製出貨單
                                    dynamicParameters = new DynamicParameters();
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
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            DoErpNo = BaseHelper.RandomCode(11),
                                            DoDate = PcPromiseDateTime,
                                            DocDate = CreateDate,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy,
                                            DoDetailId
                                        });
                                    var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                    foreach (var item in insertResult)
                                    {
                                        doId = item.DoId;
                                    }

                                    rowsAffected += insertResult.Count();
                                    #endregion
                                }

                                #region //修改出貨單身的單頭
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE SCM.DoDetail SET
                                        DoId = @DoId,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE SoDetailId = @SoDetailId
                                        AND DoId = @oriDoId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        DoId = doId,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        SoDetailId,
                                        oriDoId
                                    });

                                rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region //修改揀貨DoId
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE SCM.PickingItem SET
                                        DoId = @DoId,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE SoDetailId = @SoDetailId
                                        AND DoId = @oriDoId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        DoId = doId,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        SoDetailId,
                                        oriDoId
                                    });

                                rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region //判斷原出貨單是否還有單身
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM SCM.DoDetail
                                        WHERE DoId = @DoId";
                                dynamicParameters.Add("DoId", oriDoId);

                                var result8 = sqlConnection.Query(sql, dynamicParameters);

                                if (result8.Count() <= 0)
                                {

                                    #region //刪除原出貨單
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE SCM.DeliveryOrder
                                            WHERE DoId = @DoId";
                                    dynamicParameters.Add("DoId", oriDoId);

                                    rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                                #endregion
                            }

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
                                    PcPromiseDate = PcPromiseDateTime,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    SoDetailId
                                });

                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //撈取修改人員No
                            sql = @"SELECT UserNo
                                    FROM BAS.[User]
                                    WHERE UserId = @UserId";
                            dynamicParameters.Add("UserId", LastModifiedBy);

                            var result9 = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item in result9)
                            {
                                modifier = item.UserNo;
                            }
                            #endregion

                            #region //撈取單別/單號/流水號
                            sql = @"SELECT b.SoErpPrefix, b.SoErpNo, a.SoSequence 
                                        FROM SCM.SoDetail a
                                        INNER JOIN SCM.SaleOrder b ON b.SoId = a.SoId
                                        WHERE SoDetailId = @SoDetailId";
                            dynamicParameters.Add("SoDetailId", SoDetailId);

                            var result10 = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item in result10)
                            {
                                soErpPrefix = item.SoErpPrefix;
                                soErpNo = item.SoErpNo;
                                soSequence = item.SoSequence;
                            }
                            #endregion

                        }
                        else throw new SystemException("該出貨單已產出暫出單，無法修改交期!");

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)"
                        });
                        #endregion
                    }
                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
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
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TD048 = DateTime.ParseExact(PcPromiseDate, "yyyy-MM-dd", null).ToString("yyyyMMdd"),
                                MODI_DATE = dateNow,
                                MODI_TIME = timeNow,
                                MODIFIER = modifier,
                                TD001 = soErpPrefix,
                                TD002 = soErpNo,
                                TD003 = soSequence
                            });

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //AddDeliveryFinalize -- 出貨定版資料新增 -- Zoey 2022.09.19
        public string AddDeliveryFinalize(string Delieverys)
        {
            try
            {
                if (!Delieverys.TryParseJson(out JObject tempJObject)) throw new SystemException("定版資料格式錯誤");

                JObject deliveryJson = JObject.Parse(Delieverys);
                if (!deliveryJson.ContainsKey("data")) throw new SystemException("定版資料錯誤");

                JToken deliveryData = deliveryJson["data"];
                if (deliveryData.Count() < 0) throw new SystemException("查無出貨單定版內容");

                List<DeliveryFinalize> deliveryFinalizes = new List<DeliveryFinalize>();
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
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
                        }
                        #endregion

                        for (int i = 0; i < deliveryData.Count(); i++)
                        {
                            #region //判斷出貨客戶是否符合訂單客戶
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM SCM.DeliveryCustomer
                                    WHERE DcId = @DcId
                                    AND CustomerId = @CustomerId";
                            dynamicParameters.Add("DcId", deliveryData[i]["deliveryCustomer"].ToString());
                            dynamicParameters.Add("CustomerId", deliveryData[i]["customerId"].ToString());

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("該【出貨客戶】不屬於此【訂單客戶】!");
                            #endregion

                            #region //查找訂單資料
                            string soErpPrefix = "", soErpNo = "", soSequence = "";
                            int SoQty = 0, FreebieQty = 0, SpareQty = 0;

                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT b.SoErpPrefix, b.SoErpNo, a.SoSequence
                                    , SUM(SoQty) SoQty, SUM(FreebieQty) FreebieQty, SUM(SpareQty) SpareQty
                                    FROM SCM.SoDetail a
                                    INNER JOIN SCM.SaleOrder b ON a.SoId = b.SoId
                                    WHERE a.SoDetailId = @SoDetailId
                                    GROUP BY b.SoErpPrefix, b.SoErpNo, a.SoSequence";
                            dynamicParameters.Add("SoDetailId", Convert.ToInt32(deliveryData[i]["soDetailId"]));
                            
                            var resultSoInfo = sqlConnection.Query(sql, dynamicParameters);
                            if (resultSoInfo.Count() > 0)
                            {
                                foreach (var item in resultSoInfo)
                                {
                                    soErpPrefix = item.SoErpPrefix;
                                    soErpNo = item.SoErpNo;
                                    soSequence = item.SoSequence;

                                    SoQty = Convert.ToInt32(item.SoQty);
                                    FreebieQty = Convert.ToInt32(item.FreebieQty);
                                    SpareQty = Convert.ToInt32(item.SpareQty);
                                }
                            }
                            #endregion

                            #region //判斷定版數量與相關Table數量
                            int finalize_soQty = Convert.ToInt32(deliveryData[i]["soQty"]);
                            int finalize_freebieQty = Convert.ToInt32(deliveryData[i]["freebieQty"]);
                            int finalize_spareQty = Convert.ToInt32(deliveryData[i]["spareQty"]);

                            int tsnQty = 0,
                                un_tsnQty = 0,
                                tsrnQty = 0,
                                un_tsrnQty = 0,
                                tsnFreebieQty = 0,
                                un_tsnFreebieQty = 0,
                                tsrnFreebieQty = 0,
                                un_tsrnFreebieQty = 0,
                                tsnSpareQty = 0,
                                un_tsnSpareQty = 0,
                                tsrnSpareQty = 0,
                                un_tsrnSpareQty = 0,
                                snQty = 0,
                                un_snQty = 0,
                                srnQty = 0,
                                un_srnQty = 0,
                                snFreebieQty = 0,
                                un_snFreebieQty = 0,
                                srnFreebieQty = 0,
                                un_srnFreebieQty = 0,
                                snSpareQty = 0,
                                un_snSpareQty = 0,
                                srnSpareQty = 0,
                                un_srnSpareQty = 0;

                            using (SqlConnection erpSqlConnection = new SqlConnection(ErpConnectionStrings))
                            {
                                #region //ERP 相關帳務
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.TD008 soQty, a.TD024 freebieQty, a.TD050 spareQty
                                        , b1.TG009 tsnQty, b1.TG021 tsrnQty
                                        ,CASE WHEN a.TD049 = '1' THEN b1.TG044  ELSE  0 END tsnFreebieQty
                                        ,CASE WHEN a.TD049 = '2' THEN b1.TG044  ELSE  0 END tsnSpareQty
                                        ,CASE WHEN a.TD049 = '1' THEN b1.TG048  ELSE  0 END tsrnFreebieQty
                                        ,CASE WHEN a.TD049 = '2' THEN b1.TG048  ELSE  0 END tsrnSpareQty
                                        , b2.TG009 un_tsnQty
                                        ,CASE WHEN a.TD049 = '1' THEN b2.TG044  ELSE  0 END un_tsnFreebieQty
                                        ,CASE WHEN a.TD049 = '2' THEN b2.TG044  ELSE  0 END un_tsnSpareQty
                                        , b3.TI009 un_tsrnQty
                                        ,CASE WHEN a.TD049 = '1' THEN b3.TI035  ELSE  0 END un_tsrnFreebieQty
                                        ,CASE WHEN a.TD049 = '2' THEN b3.TI035  ELSE  0 END un_tsrnSpareQty
                                        , c1.TH008 un_snQty
                                        ,CASE WHEN a.TD049 = '1' THEN c1.TH024  ELSE  0 END un_snFreebieQty
                                        ,CASE WHEN a.TD049 = '2' THEN c1.TH024  ELSE  0 END un_snSpareQty
                                        , c2.TJ007 un_srnQty
                                        ,CASE WHEN a.TD049 = '1' THEN c2.TJ042  ELSE  0 END un_srnFreebieQty
                                        ,CASE WHEN a.TD049 = '2' THEN c2.TJ042  ELSE  0 END un_srnSpareQty
                                        , c3.TH008 snQty
                                        ,CASE WHEN a.TD049 = '1' THEN c3.TH024  ELSE  0 END snFreebieQty
                                        ,CASE WHEN a.TD049 = '2' THEN c3.TH024  ELSE  0 END snSpareQty
                                        , c4.TJ007 srnQty
                                        ,CASE WHEN a.TD049 = '1' THEN c4.TJ042  ELSE  0 END srnFreebieQty
                                        ,CASE WHEN a.TD049 = '2' THEN c4.TJ042  ELSE  0 END srnSpareQty
                                        FROM COPTD a
                                        OUTER APPLY(
                                            --暫出單單身確認
                                            SELECT SUM(x1.TG009) TG009,SUM(x1.TG044)  TG044,SUM(x1.TG020)  TG020 ,SUM(x1.TG021)  TG021,SUM(x1.TG046)  TG046,SUM(x1.TG048)  TG048
                                            FROM  INVTG x1
                                            WHERE x1.TG014 = LTRIM(RTRIM(a.TD001))
                                            AND x1.TG015 = LTRIM(RTRIM(a.TD002))
                                            AND x1.TG016 = LTRIM(RTRIM(a.TD003))
                                            AND x1.TG022 = 'Y'
                                        ) b1
                                        OUTER APPLY(
                                            --暫出單單身未確認
                                            SELECT SUM(x1.TG009) TG009 ,SUM(x1.TG044)  TG044
                                            FROM  INVTG x1
                                            WHERE x1.TG014 = LTRIM(RTRIM(a.TD001))
                                            AND x1.TG015 = LTRIM(RTRIM(a.TD002))
                                            AND x1.TG016 = LTRIM(RTRIM(a.TD003))
                                            AND x1.TG022 = 'N'
                                        ) b2
                                        OUTER APPLY(
                                            --暫出歸還單單身未確認
                                            SELECT SUM(x1.TI009) TI009 ,SUM(x1.TI035)  TI035
                                            FROM  INVTI x1
                                            INNER JOIN INVTG x2 on x2.TG001 = x1.TI014 AND x2.TG002 = x1.TI015 AND x2.TG003 = x1.TI016 
                                            WHERE x2.TG014 = LTRIM(RTRIM(a.TD001))
                                            AND x2.TG015 = LTRIM(RTRIM(a.TD002))
                                            AND x2.TG016 = LTRIM(RTRIM(a.TD003))
                                            AND x1.TI022 = 'N'
                                        ) b3
                                        OUTER APPLY(
                                            --銷貨未確認
                                            SELECT SUM(x1.TH008) TH008 ,SUM(x1.TH024)  TH024
                                            FROM  COPTH x1
                                            WHERE x1.TH014 = LTRIM(RTRIM(a.TD001))
                                            AND x1.TH015 = LTRIM(RTRIM(a.TD002))
                                            AND x1.TH016 = LTRIM(RTRIM(a.TD003))
                                            AND x1.TH020 = 'N'
                                        ) c1
                                        OUTER APPLY(
                                            --銷退未確認
                                            SELECT SUM(x1.TJ007) TJ007 ,SUM(x1.TJ042)  TJ042
                                            FROM  COPTJ x1
                                            WHERE x1.TJ018 = LTRIM(RTRIM(a.TD001))
                                            AND x1.TJ019 = LTRIM(RTRIM(a.TD002))
                                            AND x1.TJ020 = LTRIM(RTRIM(a.TD003))
                                            AND x1.TJ021 = 'N'
                                        ) c2
                                        OUTER APPLY(
                                            --銷貨確認
                                            SELECT SUM(x1.TH008) TH008 ,SUM(x1.TH024)  TH024
                                            FROM  COPTH x1
                                            WHERE x1.TH014 = LTRIM(RTRIM(a.TD001))
                                            AND x1.TH015 = LTRIM(RTRIM(a.TD002))
                                            AND x1.TH016 = LTRIM(RTRIM(a.TD003))
                                            AND x1.TH020 = 'Y'
                                        ) c3
                                        OUTER APPLY(
                                            --銷退確認
                                            SELECT SUM(x1.TJ007) TJ007 ,SUM(x1.TJ042)  TJ042
                                            FROM  COPTJ x1
                                            WHERE x1.TJ018 = LTRIM(RTRIM(a.TD001))
                                            AND x1.TJ019 = LTRIM(RTRIM(a.TD002))
                                            AND x1.TJ020 = LTRIM(RTRIM(a.TD003))
                                            AND x1.TJ021 = 'Y'
                                        ) c4
                                        WHERE LTRIM(RTRIM(a.TD001)) = @SoErpPrefix
                                        AND LTRIM(RTRIM(a.TD002)) = @SoErpNo
                                        AND LTRIM(RTRIM(a.TD003)) = @SoSequence
                                        AND a.TD021 = 'Y'";
                                dynamicParameters.Add("SoErpPrefix", soErpPrefix);
                                dynamicParameters.Add("SoErpNo", soErpNo);
                                dynamicParameters.Add("SoSequence", soSequence);

                                var resultErpAccounting = erpSqlConnection.Query(sql, dynamicParameters);

                                if (resultErpAccounting.Count() > 0)
                                {
                                    foreach (var item in resultErpAccounting)
                                    {
                                        tsnQty = Convert.ToInt32(item.tsnQty);
                                        un_tsnQty = Convert.ToInt32(item.un_tsnQty);
                                        tsrnQty = Convert.ToInt32(item.tsrnQty);
                                        un_tsrnQty = Convert.ToInt32(item.un_tsrnQty);
                                        tsnFreebieQty = Convert.ToInt32(item.tsnFreebieQty);
                                        un_tsnFreebieQty = Convert.ToInt32(item.un_tsnFreebieQty);
                                        tsrnFreebieQty = Convert.ToInt32(item.tsrnFreebieQty);
                                        un_tsrnFreebieQty = Convert.ToInt32(item.un_tsrnFreebieQty);
                                        tsnSpareQty = Convert.ToInt32(item.tsnSpareQty);
                                        un_tsnSpareQty = Convert.ToInt32(item.un_tsnSpareQty);
                                        tsrnSpareQty = Convert.ToInt32(item.tsrnSpareQty);
                                        un_tsrnSpareQty = Convert.ToInt32(item.un_tsrnSpareQty);
                                        snQty = Convert.ToInt32(item.snQty);
                                        un_snQty = Convert.ToInt32(item.un_snQty);
                                        srnQty = Convert.ToInt32(item.srnQty);
                                        un_srnQty = Convert.ToInt32(item.un_srnQty);
                                        snFreebieQty = Convert.ToInt32(item.snFreebieQty);
                                        un_snFreebieQty = Convert.ToInt32(item.un_snFreebieQty);
                                        srnFreebieQty = Convert.ToInt32(item.srnFreebieQty);
                                        un_srnFreebieQty = Convert.ToInt32(item.un_srnFreebieQty);
                                        snSpareQty = Convert.ToInt32(item.snSpareQty);
                                        un_snSpareQty = Convert.ToInt32(item.un_snSpareQty);
                                        srnSpareQty = Convert.ToInt32(item.srnSpareQty);
                                        un_srnSpareQty = Convert.ToInt32(item.un_srnSpareQty);
                                    }
                                }
                                #endregion
                            }

                            #region //正常品檢查
                            int regularQty = tsnQty + un_tsnQty - (tsrnQty + un_tsrnQty) - (srnQty + un_srnQty);

                            if ((tsnQty + un_tsnQty) <= 0 && (snQty + un_snQty) > 0)
                            {
                                regularQty = snQty + un_snQty - (srnQty + un_srnQty);
                            }

                            if (finalize_soQty > (SoQty - regularQty)) throw new SystemException("【正常品】撿貨數不足!");
                            #endregion

                            #region //贈品檢查
                            int freebieQty = tsnFreebieQty + un_tsnFreebieQty - (tsrnFreebieQty + un_tsrnFreebieQty) - (srnFreebieQty + un_srnFreebieQty);

                            if ((tsnFreebieQty + un_tsnFreebieQty) <= 0 && (snFreebieQty + un_snFreebieQty) > 0)
                            {
                                freebieQty = snFreebieQty + un_snFreebieQty - (srnFreebieQty + un_srnFreebieQty);
                            }

                            if (finalize_freebieQty > (FreebieQty - freebieQty)) throw new SystemException("【贈品】撿貨數不足!");
                            #endregion

                            #region //備品檢查
                            int spareQty = tsnSpareQty + un_tsnSpareQty - (tsrnSpareQty + un_tsrnSpareQty) - (srnSpareQty + un_srnSpareQty);

                            if ((tsnSpareQty + un_tsnSpareQty) <= 0 && (snSpareQty + un_snSpareQty) > 0)
                            {
                                spareQty = snSpareQty + un_snSpareQty - (srnSpareQty + un_srnSpareQty);
                            }

                            if (finalize_spareQty > (SpareQty - spareQty)) throw new SystemException("【備品】撿貨數不足!");
                            #endregion

                            #region //改為帳務計算停用(2024.08.02)
                            /*
                            #region //查找暫出單身數量
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT (ISNULL(SUM(TsnQty), 0) - ISNULL(SUM(ReturnQty), 0)) TsnQty
                                    FROM SCM.TsnDetail
                                    WHERE SoDetailId = @SoDetailId
                                    AND ConfirmStatus = @ConfirmStatus";
                            dynamicParameters.Add("SoDetailId", Convert.ToInt32(deliveryData[i]["soDetailId"]));
                            dynamicParameters.Add("ConfirmStatus", "Y");

                            int TsnQty = 0;
                            var result2 = sqlConnection.Query(sql, dynamicParameters);
                            if (result2.Count() > 0)
                            {
                                foreach (var item in result2)
                                {
                                    TsnQty = Convert.ToInt32(item.TsnQty);
                                }
                            }
                            #endregion

                            #region //查找揀貨物件【正常品】數量
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT SUM(a.ItemQty) RegularItemQty
                                    FROM SCM.PickingItem a
                                    WHERE a.SoDetailId = @SoDetailId
                                    AND a.ItemType = 1
                                    AND a.DoId IS NOT NULL";
                            dynamicParameters.Add("SoDetailId", Convert.ToInt32(deliveryData[i]["soDetailId"]));

                            int RegularItemQty = 0;
                            var resultPickingRegular = sqlConnection.Query(sql, dynamicParameters);
                            if (resultPickingRegular.Count() > 0)
                            {
                                foreach (var item in resultPickingRegular)
                                {
                                    RegularItemQty = Convert.ToInt32(item.RegularItemQty);
                                }
                            }
                            #endregion

                            #region //查找揀貨物件【贈品】數量
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT SUM(a.ItemQty) FreebieItemQty
                                    FROM SCM.PickingItem a
                                    WHERE a.SoDetailId = @SoDetailId
                                    AND a.ItemType = 2
                                    AND a.DoId IS NOT NULL";
                            dynamicParameters.Add("SoDetailId", Convert.ToInt32(deliveryData[i]["soDetailId"]));

                            int FreebieItemQty = 0;
                            var resultPickingFreebie = sqlConnection.Query(sql, dynamicParameters);
                            if (resultPickingFreebie.Count() > 0)
                            {
                                foreach (var item in resultPickingFreebie)
                                {
                                    FreebieItemQty = Convert.ToInt32(item.FreebieItemQty);
                                }
                            }
                            #endregion

                            #region //查找揀貨物件【備品】數量
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT SUM(a.ItemQty) SpareItemQty
                                    FROM SCM.PickingItem a
                                    WHERE a.SoDetailId = @SoDetailId
                                    AND a.ItemType = 3
                                    AND a.DoId IS NOT NULL";
                            dynamicParameters.Add("SoDetailId", Convert.ToInt32(deliveryData[i]["soDetailId"]));

                            int SpareItemQty = 0;
                            var resultPickingSpare = sqlConnection.Query(sql, dynamicParameters);
                            if (resultPickingSpare.Count() > 0)
                            {
                                foreach (var item in resultPickingSpare)
                                {
                                    SpareItemQty = Convert.ToInt32(item.SpareItemQty);
                                }
                            }
                            #endregion

                            #region //查找退貨【正常品】數量
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT SUM(a.RmaQty) TotalRegularRmaQty
                                    FROM SCM.RmaDetail a
                                    INNER JOIN SCM.TsnDetail b ON a.TsnDetailId = b.TsnDetailId
                                    INNER JOIN SCM.ReturnMerchandiseAuthorization c ON a.RmaId = c.RmaId
                                    WHERE b.SoDetailId = @SoDetailId
                                    AND a.ItemType = 1
                                    AND c.ConfirmStatus = 'Y'";
                            dynamicParameters.Add("SoDetailId", Convert.ToInt32(deliveryData[i]["soDetailId"]));

                            int TotalRegularRmaQty = 0;
                            var resultRmaDtlRegular = sqlConnection.Query(sql, dynamicParameters);
                            if (resultRmaDtlRegular.Count() > 0)
                            {
                                foreach (var item in resultRmaDtlRegular)
                                {
                                    TotalRegularRmaQty = Convert.ToInt32(item.TotalRegularRmaQty);
                                }
                            }
                            #endregion

                            #region //查找退貨【贈品】數量
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT SUM(a.RmaQty) TotalFreebieRmaQty
                                    FROM SCM.RmaDetail a
                                    INNER JOIN SCM.TsnDetail b ON a.TsnDetailId = b.TsnDetailId
                                    INNER JOIN SCM.ReturnMerchandiseAuthorization c ON a.RmaId = c.RmaId
                                    WHERE b.SoDetailId = @SoDetailId
                                    AND a.ItemType = 2
                                    AND c.ConfirmStatus = 'Y'";
                            dynamicParameters.Add("SoDetailId", Convert.ToInt32(deliveryData[i]["soDetailId"]));

                            int TotalFreebieRmaQty = 0;
                            var resultRmaDtlFreebie = sqlConnection.Query(sql, dynamicParameters);
                            if (resultRmaDtlFreebie.Count() > 0)
                            {
                                foreach (var item in resultRmaDtlFreebie)
                                {
                                    TotalFreebieRmaQty = Convert.ToInt32(item.TotalFreebieRmaQty);
                                }
                            }
                            #endregion

                            #region //查找退貨【備品】數量
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT SUM(a.RmaQty) TotalSpareRmaQty
                                    FROM SCM.RmaDetail a
                                    INNER JOIN SCM.TsnDetail b ON a.TsnDetailId = b.TsnDetailId
                                    INNER JOIN SCM.ReturnMerchandiseAuthorization c ON a.RmaId = c.RmaId
                                    WHERE b.SoDetailId = @SoDetailId
                                    AND a.ItemType = 3
                                    AND c.ConfirmStatus = 'Y'";
                            dynamicParameters.Add("SoDetailId", Convert.ToInt32(deliveryData[i]["soDetailId"]));

                            int TotalSpareRmaQty = 0;
                            var resultRmaDtlSpare = sqlConnection.Query(sql, dynamicParameters);
                            if (resultRmaDtlSpare.Count() > 0)
                            {
                                foreach (var item in resultRmaDtlSpare)
                                {
                                    TotalSpareRmaQty = Convert.ToInt32(item.TotalSpareRmaQty);
                                }
                            }
                            #endregion

                            //if (Convert.ToInt32(deliveryData[i]["soQty"]) > (SoQty - TsnQty)) throw new SystemException("撿貨數不足!");
                            if (Convert.ToInt32(deliveryData[i]["soQty"]) > (SoQty - (RegularItemQty - TotalRegularRmaQty))) throw new SystemException("【正常品】撿貨數不足!");
                            if (Convert.ToInt32(deliveryData[i]["freebieQty"]) > (FreebieQty - (FreebieItemQty - TotalFreebieRmaQty))) throw new SystemException("【贈品】撿貨數不足!");
                            if (Convert.ToInt32(deliveryData[i]["spareQty"]) > (SpareQty - (SpareItemQty - TotalSpareRmaQty))) throw new SystemException("【備品】撿貨數不足!");
                            */
                            #endregion
                            #endregion

                            deliveryFinalizes.Add(
                                new DeliveryFinalize
                                {
                                    index = Convert.ToInt32(deliveryData[i]["index"]),
                                    soDetailId = Convert.ToInt32(deliveryData[i]["soDetailId"]),
                                    pcPromiseDate = deliveryData[i]["pcPromiseDate"].ToString(),
                                    pcPromiseTime = deliveryData[i]["pcPromiseTime"].ToString(),
                                    customerNo = deliveryData[i]["customerNo"].ToString(),
                                    customerName = deliveryData[i]["customerName"].ToString(),
                                    soNo = deliveryData[i]["soNo"].ToString(),
                                    soSeq = deliveryData[i]["soSeq"].ToString(),
                                    mtlItemName = deliveryData[i]["mtlItemName"].ToString(),
                                    soQty = Convert.ToDouble(deliveryData[i]["soQty"]),
                                    freebieQty = Convert.ToDouble(deliveryData[i]["freebieQty"]),
                                    spareQty = Convert.ToDouble(deliveryData[i]["spareQty"]),
                                    deliveryCustomer = Convert.ToInt32(deliveryData[i]["deliveryCustomer"]),
                                    deliveryProcess = Convert.ToInt32(deliveryData[i]["deliveryProcess"]),
                                    orderSituation = deliveryData[i]["orderSituation"].ToString(),
                                    deliveryMethod = deliveryData[i]["deliveryMethod"].ToString(),
                                    pcDoDetailRemark = deliveryData[i]["pcDoDetailRemark"].ToString()
                                });
                        }

                        var deliveryCustomers = deliveryFinalizes
                                .GroupBy(x => new
                                {
                                    x.customerNo,
                                    x.customerName,
                                    x.deliveryCustomer,
                                    x.pcPromiseDate,
                                    x.pcPromiseTime
                                })
                                .Select(x => new DeliveryFinalize
                                {
                                    customerNo = x.Key.customerNo,
                                    customerName = x.Key.customerName,
                                    deliveryCustomer = x.Key.deliveryCustomer,
                                    pcPromiseDate = x.Key.pcPromiseDate,
                                    pcPromiseTime = x.Key.pcPromiseTime
                                })
                                .ToList();

                        int rowsAffected = 0;
                        foreach (var dc in deliveryCustomers)
                        {
                            string DoErpNo = BaseHelper.RandomCode(11);

                            var deliverys = deliveryFinalizes
                            .Where(x => x.deliveryCustomer == dc.deliveryCustomer
                                    && x.pcPromiseDate == dc.pcPromiseDate
                                    && x.pcPromiseTime == dc.pcPromiseTime)
                            .OrderBy(x => x.soNo)
                            .ThenBy(x => x.soSeq);

                            #region //撈取所屬部門ID
                            sql = @"SELECT DepartmentId 
                                    FROM BAS.[User]
                                    WHERE UserId = @UserId";
                            dynamicParameters.Add("UserId", CurrentUser);

                            var result = sqlConnection.Query(sql, dynamicParameters);

                            int DepartmentId = -1;
                            foreach (var item in result)
                            {
                                DepartmentId = item.DepartmentId;
                            }
                            #endregion

                            #region //撈取客戶資料
                            sql = @"SELECT a.CustomerId, a.DeliveryAddressFirst
                                    FROM SCM.Customer a
                                    WHERE CompanyId = @CompanyId
                                    AND CustomerNo = @CustomerNo";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("CustomerNo", dc.customerNo);

                            var result2 = sqlConnection.Query(sql, dynamicParameters);

                            int CustomerId = -1;
                            string DoAddressFirst = "";
                            foreach (var item in result2)
                            {
                                CustomerId = item.CustomerId;
                                DoAddressFirst = item.DeliveryAddressFirst;
                            }
                            #endregion

                            #region //判斷出貨單是否存在/撈取DoId
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT DoId, Status
                                    FROM SCM.DeliveryOrder
                                    WHERE DcId = @DcId
                                    AND DoDate = @DoDate";
                            dynamicParameters.Add("DcId", dc.deliveryCustomer);
                            dynamicParameters.Add("DoDate", dc.pcPromiseDate + " " + dc.pcPromiseTime);

                            var result3 = sqlConnection.Query(sql, dynamicParameters);
                            int DoId = -1;
                            string DoStatus = "";
                            foreach (var item in result3)
                            {
                                DoId = item.DoId;
                                DoStatus = item.Status;
                            }
                            #endregion

                            if (DoId < 0)
                            {
                                #region //新增出貨單頭
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO SCM.DeliveryOrder (CompanyId, DepartmentId, UserId, DoErpPrefix, DoErpNo, DoDate, DocDate
                                        , CustomerId, DcId, WayBill, Traffic, ShipMethod, DoAddressFirst, DoAddressSecond, TotalQty, Amount, TaxAmount
                                        , DoRemark, WareHouseDoRemark, MeasureMailStatus, ConfirmStatus, ConfirmUserId, TransferStatus
                                        , TransferDate, Status, CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.DoId
                                        VALUES (@CompanyId, @DepartmentId, @UserId, @DoErpPrefix, @DoErpNo, @DoDate, @DocDate, @CustomerId
                                        , @DcId, @WayBill, @Traffic, @ShipMethod, @DoAddressFirst, @DoAddressSecond, @TotalQty, @Amount, @TaxAmount
                                        , @DoRemark, @WareHouseDoRemark, @MeasureMailStatus, @ConfirmStatus, @ConfirmUserId, @TransferStatus
                                        , @TransferDate, @Status, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        CompanyId = CurrentCompany,
                                        DepartmentId,
                                        UserId = CurrentUser,
                                        DoErpPrefix = 1301,
                                        DoErpNo,
                                        DoDate = dc.pcPromiseDate + " " + dc.pcPromiseTime,
                                        DocDate = dc.pcPromiseDate + " " + dc.pcPromiseTime,
                                        CustomerId,
                                        DcId = dc.deliveryCustomer,
                                        WayBill = "",
                                        Traffic = "",
                                        ShipMethod = "",
                                        DoAddressFirst,
                                        DoAddressSecond = "",
                                        TotalQty = 0,
                                        Amount = 0,
                                        TaxAmount = 0,
                                        DoRemark = "",
                                        WareHouseDoRemark = "",
                                        MeasureMailStatus = "N",
                                        ConfirmStatus = "N",
                                        ConfirmUserId = (int?)null,
                                        TransferStatus = "N",
                                        TransferDate = "",
                                        Status = "N",
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });

                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();

                                foreach (var item in insertResult)
                                {
                                    DoId = Convert.ToInt32(item.DoId);
                                }
                                #endregion

                                foreach (var delivery in deliverys)
                                {
                                    if ((delivery.soQty + delivery.freebieQty + delivery.spareQty) <= 0) throw new SystemException("出貨數量不可需大於0!");

                                    #region //判斷訂單資料是否存在/判斷訂單狀態、數量/撈取訂單單身ID
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.ConfirmStatus, a.ClosureStatus, a.SoDetailId
                                            , a.SoQty, a.FreebieQty, a.SpareQty
                                            FROM SCM.SoDetail a
                                            LEFT JOIN SCM.SaleOrder b ON b.SoId = a.SoId
                                            WHERE b.SoErpPrefix +'-'+ b.SoErpNo = @SoFullNo
                                            AND a.SoSequence = @SoSequence
                                            AND b.CompanyId = @CompanyId";
                                    dynamicParameters.Add("SoFullNo", delivery.soNo);
                                    dynamicParameters.Add("SoSequence", delivery.soSeq);
                                    dynamicParameters.Add("CompanyId", CurrentCompany);

                                    var result4 = sqlConnection.Query(sql, dynamicParameters);

                                    string ConfirmStatus = "";
                                    string ClosureStatus = "";
                                    int SoDetailId = -1;
                                    int SoQty = -1;
                                    int FreebieQty = -1;
                                    int SpareQty = -1;
                                    foreach (var item in result4)
                                    {
                                        ConfirmStatus = item.ConfirmStatus;
                                        ClosureStatus = item.ClosureStatus;
                                        SoDetailId = item.SoDetailId;
                                        SoQty = Convert.ToInt32(item.SoQty);
                                        FreebieQty = Convert.ToInt32(item.FreebieQty);
                                        SpareQty = Convert.ToInt32(item.SpareQty);
                                    }

                                    if (result4.Count() <= 0) throw new SystemException("查無 " + delivery.soNo + "(" + delivery.soSeq + ") 訂單資料!");

                                    if (SoQty < delivery.soQty) throw new SystemException("出貨數量不可大於訂單數量!");
                                    if (ConfirmStatus == "N") throw new SystemException("訂單：" + delivery.soNo + "(" + delivery.soSeq + ") 尚未確認，無法進行建立作業!");
                                    if (ConfirmStatus == "V") throw new SystemException("訂單：" + delivery.soNo + "(" + delivery.soSeq + ") 已作廢，無法進行建立作業!");
                                    if (ClosureStatus == "Y") throw new SystemException("訂單：" + delivery.soNo + "(" + delivery.soSeq + ") 已自動結案，無法進行建立作業!");
                                    if (ClosureStatus == "y") throw new SystemException("訂單：" + delivery.soNo + "(" + delivery.soSeq + ") 已指定結案，無法進行建立作業!");
                                    #endregion

                                    #region //查詢出貨單流水號
                                    sql = @"SELECT REPLACE(STR(ISNULL(MAX(DoSequence),0) + 1,4),' ','0') DoSequence 
                                            FROM SCM.DoDetail
                                            WHERE DoId = @DoId";
                                    dynamicParameters.Add("DoId", DoId);

                                    var result5 = sqlConnection.Query(sql, dynamicParameters);

                                    string DoSequence = "";
                                    foreach (var item in result5)
                                    {
                                        DoSequence = item.DoSequence;
                                    }
                                    #endregion

                                    #region //新增出貨單身
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO SCM.DoDetail (DoId, SoDetailId, DoSequence, TransInInventoryId, DoQty, FreebieQty, SpareQty
                                            , UnitPrice, Amount, DoDetailRemark, WareHouseDoDetailRemark, PcDoDetailRemark
                                            , DeliveryProcess, DeliveryRoutine, DeliveryDocument, OrderSituation, DeliveryMethod
                                            , ConfirmStatus, ClosureStatus, CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.DoDetailId
                                            VALUES (@DoId, @SoDetailId, @DoSequence, @TransInInventoryId, @DoQty, @FreebieQty, @SpareQty
                                            , @UnitPrice, @Amount, @DoDetailRemark, @WareHouseDoDetailRemark
                                            , @PcDoDetailRemark, @DeliveryProcess, @DeliveryRoutine, @DeliveryDocument, @OrderSituation, @DeliveryMethod
                                            , @ConfirmStatus, @ClosureStatus, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            DoId,
                                            SoDetailId,
                                            DoSequence,
                                            TransInInventoryId = (int?)null,
                                            DoQty = delivery.soQty,
                                            FreebieQty = delivery.freebieQty,
                                            SpareQty = delivery.spareQty,
                                            UnitPrice = 0,
                                            Amount = 0,
                                            DoDetailRemark = "",
                                            WareHouseDoDetailRemark = "",
                                            PcDoDetailRemark = delivery.pcDoDetailRemark,
                                            DeliveryProcess = delivery.deliveryProcess,
                                            DeliveryRoutine = "",
                                            DeliveryDocument = "",
                                            OrderSituation = delivery.orderSituation,
                                            DeliveryMethod = delivery.deliveryMethod,
                                            ConfirmStatus = "N",
                                            ClosureStatus = "N",
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    insertResult = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult.Count();

                                    #endregion
                                }
                            }
                            else
                            {
                                if (DoStatus == "R" || DoStatus == "S") throw new SystemException("該出貨時間已截止，請選擇其他出貨時間!");

                                #region //修改出貨單頭
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE SCM.DeliveryOrder SET
                                        DoDate = @DoDate,
                                        DocDate = @DocDate,
                                        DcId = @DcId,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE DoId = @DoId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        DoDate = dc.pcPromiseDate + " " + dc.pcPromiseTime,
                                        DocDate = dc.pcPromiseDate + " " + dc.pcPromiseTime,
                                        DcId = dc.deliveryCustomer,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        DoId
                                    });

                                rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                foreach (var delivery in deliverys)
                                {
                                    if (delivery.soQty <= 0) throw new SystemException("出貨數量不可需大於0!");

                                    #region //判斷訂單資料是否存在/判斷訂單狀態、數量/撈取訂單單身ID
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.ConfirmStatus, a.ClosureStatus, a.SoDetailId
                                            , a.SoQty, a.FreebieQty, a.SpareQty
                                            FROM SCM.SoDetail a
                                            LEFT JOIN SCM.SaleOrder b ON b.SoId = a.SoId
                                            WHERE b.SoErpPrefix +'-'+ b.SoErpNo = @SoFullNo
                                            AND a.SoSequence = @SoSequence
                                            AND b.CompanyId = @CompanyId";
                                    dynamicParameters.Add("SoFullNo", delivery.soNo);
                                    dynamicParameters.Add("SoSequence", delivery.soSeq);
                                    dynamicParameters.Add("CompanyId", CurrentCompany);

                                    var result4 = sqlConnection.Query(sql, dynamicParameters);

                                    string ConfirmStatus = "";
                                    string ClosureStatus = "";
                                    int SoDetailId = -1;
                                    int SoQty = -1;
                                    int FreebieQty = -1;
                                    int SpareQty = -1;
                                    foreach (var item in result4)
                                    {
                                        ConfirmStatus = item.ConfirmStatus;
                                        ClosureStatus = item.ClosureStatus;
                                        SoDetailId = item.SoDetailId;
                                        SoQty = Convert.ToInt32(item.SoQty);
                                        FreebieQty = Convert.ToInt32(item.FreebieQty);
                                        SpareQty = Convert.ToInt32(item.SpareQty);
                                    }

                                    if (result4.Count() <= 0) throw new SystemException("查無 " + delivery.soNo + "(" + delivery.soSeq + ") 訂單資料!");

                                    if (SoQty < delivery.soQty) throw new SystemException("出貨數量不可大於訂單數量!");
                                    if (ConfirmStatus == "N") throw new SystemException("訂單：" + delivery.soNo + "(" + delivery.soSeq + ") 尚未確認，無法進行建立作業!");
                                    if (ConfirmStatus == "V") throw new SystemException("訂單：" + delivery.soNo + "(" + delivery.soSeq + ") 已作廢，無法進行建立作業!");
                                    if (ClosureStatus == "Y") throw new SystemException("訂單：" + delivery.soNo + "(" + delivery.soSeq + ") 已自動結案，無法進行建立作業!");
                                    if (ClosureStatus == "y") throw new SystemException("訂單：" + delivery.soNo + "(" + delivery.soSeq + ") 已指定結案，無法進行建立作業!");
                                    #endregion

                                    #region //判斷出貨單身是否存在/撈取DoDetailId
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT DoDetailId
                                            FROM SCM.DoDetail
                                            WHERE DoId = @DoId
                                            AND SoDetailId = @SoDetailId";
                                    dynamicParameters.Add("DoId", DoId);
                                    dynamicParameters.Add("SoDetailId", SoDetailId);

                                    var result5 = sqlConnection.Query(sql, dynamicParameters);
                                    int DoDetailId = -1;
                                    foreach (var item in result5)
                                    {
                                        DoDetailId = item.DoDetailId;
                                    }
                                    #endregion

                                    if (DoDetailId < 0)
                                    {
                                        #region //查詢出貨單流水號
                                        sql = @"SELECT REPLACE(STR(ISNULL(MAX(DoSequence),0) + 1,4),' ','0') DoSequence 
                                                FROM SCM.DoDetail
                                                WHERE DoId = @DoId";
                                        dynamicParameters.Add("DoId", DoId);

                                        var result6 = sqlConnection.Query(sql, dynamicParameters);

                                        string DoSequence = "";
                                        foreach (var item in result6)
                                        {
                                            DoSequence = item.DoSequence;
                                        }
                                        #endregion

                                        #region //新增出貨單身
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO SCM.DoDetail (DoId, SoDetailId, DoSequence, TransInInventoryId, DoQty, FreebieQty, SpareQty
                                                , UnitPrice, Amount, DoDetailRemark, WareHouseDoDetailRemark, PcDoDetailRemark
                                                , DeliveryProcess, DeliveryRoutine, DeliveryDocument, OrderSituation, DeliveryMethod
                                                , ConfirmStatus, ClosureStatus, CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                                OUTPUT INSERTED.DoDetailId
                                                VALUES (@DoId, @SoDetailId, @DoSequence, @TransInInventoryId, @DoQty, @FreebieQty, @SpareQty
                                                , @UnitPrice, @Amount, @DoDetailRemark, @WareHouseDoDetailRemark
                                                , @PcDoDetailRemark, @DeliveryProcess, @DeliveryRoutine, @DeliveryDocument, @OrderSituation, @DeliveryMethod
                                                , @ConfirmStatus, @ClosureStatus, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                DoId,
                                                SoDetailId,
                                                DoSequence,
                                                TransInInventoryId = (int?)null,
                                                DoQty = delivery.soQty,
                                                FreebieQty = delivery.freebieQty,
                                                SpareQty = delivery.spareQty,
                                                UnitPrice = 0,
                                                Amount = 0,
                                                DoDetailRemark = "",
                                                WareHouseDoDetailRemark = "",
                                                PcDoDetailRemark = delivery.pcDoDetailRemark,
                                                DeliveryProcess = delivery.deliveryProcess,
                                                DeliveryRoutine = "",
                                                DeliveryDocument = "",
                                                OrderSituation = delivery.orderSituation,
                                                DeliveryMethod = delivery.deliveryMethod,
                                                ConfirmStatus = "N",
                                                ClosureStatus = "N",
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
                                        #region //判斷【正常品】出貨數量是否大於已揀數量
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT ISNULL(SUM(ItemQty), 0) RegularItemQty
                                                FROM SCM.PickingItem
                                                WHERE DoId = @DoId
                                                AND ItemType = @ItemType
                                                AND SoDetailId = @SoDetailId";
                                        dynamicParameters.Add("DoId", DoId);
                                        dynamicParameters.Add("ItemType", 1);
                                        dynamicParameters.Add("SoDetailId", SoDetailId);

                                        var resultPickingRegular = sqlConnection.Query(sql, dynamicParameters);
                                        int RegularItemQty = -1;
                                        foreach (var item in resultPickingRegular)
                                        {
                                            RegularItemQty = item.RegularItemQty;
                                        }
                                        if (RegularItemQty > delivery.soQty) throw new SystemException("【正常品】出貨數量不可小於已揀數量!");
                                        #endregion

                                        #region //判斷【贈品】出貨數量是否大於已揀數量
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT ISNULL(SUM(ItemQty), 0) FreebieItemQty
                                                FROM SCM.PickingItem
                                                WHERE DoId = @DoId
                                                AND ItemType = @ItemType
                                                AND SoDetailId = @SoDetailId";
                                        dynamicParameters.Add("DoId", DoId);
                                        dynamicParameters.Add("ItemType", 2);
                                        dynamicParameters.Add("SoDetailId", SoDetailId);

                                        var resultPickingFreebie = sqlConnection.Query(sql, dynamicParameters);
                                        int FreebieItemQty = -1;
                                        foreach (var item in resultPickingFreebie)
                                        {
                                            FreebieItemQty = item.FreebieItemQty;
                                        }
                                        if (FreebieItemQty > delivery.soQty) throw new SystemException("【贈品】出貨數量不可小於已揀數量!");
                                        #endregion

                                        #region //判斷【備品】出貨數量是否大於已揀數量
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT ISNULL(SUM(ItemQty), 0) SpareItemQty
                                                FROM SCM.PickingItem
                                                WHERE DoId = @DoId
                                                AND ItemType = @ItemType
                                                AND SoDetailId = @SoDetailId";
                                        dynamicParameters.Add("DoId", DoId);
                                        dynamicParameters.Add("ItemType", 3);
                                        dynamicParameters.Add("SoDetailId", SoDetailId);

                                        var resultPickingSpare = sqlConnection.Query(sql, dynamicParameters);
                                        int SpareItemQty = -1;
                                        foreach (var item in resultPickingSpare)
                                        {
                                            SpareItemQty = item.SpareItemQty;
                                        }
                                        if (SpareItemQty > delivery.soQty) throw new SystemException("【備品】出貨數量不可小於已揀數量!");
                                        #endregion

                                        #region //修改出貨單身
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"UPDATE SCM.DoDetail SET
                                                DoQty = @DoQty,
                                                FreebieQty = @FreebieQty,
                                                SpareQty = @SpareQty,
                                                PcDoDetailRemark = @PcDoDetailRemark,
                                                DeliveryProcess = @DeliveryProcess,
                                                OrderSituation = @OrderSituation,
                                                DeliveryMethod = @DeliveryMethod,
                                                LastModifiedDate = @LastModifiedDate,
                                                LastModifiedBy = @LastModifiedBy
                                                WHERE DoDetailId = @DoDetailId";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                DoQty = delivery.soQty,
                                                FreebieQty = delivery.freebieQty,
                                                SpareQty = delivery.spareQty,
                                                PcDoDetailRemark = delivery.pcDoDetailRemark,
                                                DeliveryProcess = delivery.deliveryProcess,
                                                OrderSituation = delivery.orderSituation,
                                                DeliveryMethod = delivery.deliveryMethod,
                                                LastModifiedDate,
                                                LastModifiedBy,
                                                DoDetailId
                                            });

                                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                        #endregion
                                    }
                                }
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

        #region //AddDeliveryFinalizeOrder -- 新增出貨單詳細資料 -- Ann 2025-05-16
        public string AddDeliveryFinalizeOrder(string Delieverys)
        {
            try
            {
                if (!Delieverys.TryParseJson(out JObject tempJObject)) throw new SystemException("定版資料格式錯誤");

                JObject deliveryJson = JObject.Parse(Delieverys);
                if (!deliveryJson.ContainsKey("data")) throw new SystemException("定版資料錯誤");

                JToken deliveryData = deliveryJson["data"];
                if (deliveryData.Count() < 0) throw new SystemException("查無出貨單定版內容");

                var groupedBySoNo = deliveryData
                    .GroupBy(item => new
                    {
                        SoNo = (string)item["soNo"],
                        SoSeq = (string)item["soSeq"]
                    })
                    .Select(group => new
                    {
                        SoNo = group.Key,
                        TotalSoQty = group.Sum(item => decimal.Parse((string)item["soQty"])),
                        TotalFreebieQty = group.Sum(item => decimal.Parse((string)item["freebieQty"])),
                        TotalSpareQty = group.Sum(item => decimal.Parse((string)item["spareQty"])),
                        Items = group.ToList()
                    });

                List<DeliveryFinalize> deliveryFinalizes = new List<DeliveryFinalize>();
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
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
                        }
                        #endregion

                        for (int i = 0; i < deliveryData.Count(); i++)
                        {
                            #region //判斷出貨客戶是否符合訂單客戶
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM SCM.DeliveryCustomer
                                    WHERE DcId = @DcId
                                    AND CustomerId = @CustomerId";
                            dynamicParameters.Add("DcId", deliveryData[i]["deliveryCustomer"].ToString());
                            dynamicParameters.Add("CustomerId", deliveryData[i]["customerId"].ToString());

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("該【出貨客戶】不屬於此【訂單客戶】!");
                            #endregion

                            #region //查找訂單資料
                            string soErpPrefix = "", soErpNo = "", soSequence = "";
                            int SoQty = 0, FreebieQty = 0, SpareQty = 0;

                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT b.SoErpPrefix, b.SoErpNo, a.SoSequence
                                    , SUM(SoQty) SoQty, SUM(FreebieQty) FreebieQty, SUM(SpareQty) SpareQty
                                    FROM SCM.SoDetail a
                                    INNER JOIN SCM.SaleOrder b ON a.SoId = b.SoId
                                    WHERE a.SoDetailId = @SoDetailId
                                    GROUP BY b.SoErpPrefix, b.SoErpNo, a.SoSequence";
                            dynamicParameters.Add("SoDetailId", Convert.ToInt32(deliveryData[i]["soDetailId"]));

                            var resultSoInfo = sqlConnection.Query(sql, dynamicParameters);
                            if (resultSoInfo.Count() > 0)
                            {
                                foreach (var item in resultSoInfo)
                                {
                                    soErpPrefix = item.SoErpPrefix;
                                    soErpNo = item.SoErpNo;
                                    soSequence = item.SoSequence;

                                    SoQty = Convert.ToInt32(item.SoQty);
                                    FreebieQty = Convert.ToInt32(item.FreebieQty);
                                    SpareQty = Convert.ToInt32(item.SpareQty);
                                }
                            }
                            #endregion

                            #region //判斷定版數量與相關Table數量
                            int finalize_soQty = 0;
                            int finalize_freebieQty = 0;
                            int finalize_spareQty = 0;

                            if (groupedBySoNo.Any(g => g.SoNo.ToString() == deliveryData[i]["soNo"].ToString()))
                            {
                                finalize_soQty = Convert.ToInt32(groupedBySoNo.First(g => g.SoNo.ToString() == deliveryData[i]["soNo"].ToString()).TotalSoQty);
                                finalize_freebieQty = Convert.ToInt32(groupedBySoNo.First(g => g.SoNo.ToString() == deliveryData[i]["soNo"].ToString()).TotalFreebieQty);
                                finalize_spareQty = Convert.ToInt32(groupedBySoNo.First(g => g.SoNo.ToString() == deliveryData[i]["soNo"].ToString()).TotalSpareQty);
                            }
                            else
                            {
                                finalize_soQty = Convert.ToInt32(deliveryData[i]["soQty"]);
                                finalize_freebieQty = Convert.ToInt32(deliveryData[i]["freebieQty"]);
                                finalize_spareQty = Convert.ToInt32(deliveryData[i]["spareQty"]);
                            }

                            int tsnQty = 0,
                                un_tsnQty = 0,
                                tsrnQty = 0,
                                un_tsrnQty = 0,
                                tsnFreebieQty = 0,
                                un_tsnFreebieQty = 0,
                                tsrnFreebieQty = 0,
                                un_tsrnFreebieQty = 0,
                                tsnSpareQty = 0,
                                un_tsnSpareQty = 0,
                                tsrnSpareQty = 0,
                                un_tsrnSpareQty = 0,
                                snQty = 0,
                                un_snQty = 0,
                                srnQty = 0,
                                un_srnQty = 0,
                                snFreebieQty = 0,
                                un_snFreebieQty = 0,
                                srnFreebieQty = 0,
                                un_srnFreebieQty = 0,
                                snSpareQty = 0,
                                un_snSpareQty = 0,
                                srnSpareQty = 0,
                                un_srnSpareQty = 0;

                            using (SqlConnection erpSqlConnection = new SqlConnection(ErpConnectionStrings))
                            {
                                #region //ERP 相關帳務
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.TD008 soQty, a.TD024 freebieQty, a.TD050 spareQty
                                        , b1.TG009 tsnQty, b1.TG021 tsrnQty
                                        ,CASE WHEN a.TD049 = '1' THEN b1.TG044  ELSE  0 END tsnFreebieQty
                                        ,CASE WHEN a.TD049 = '2' THEN b1.TG044  ELSE  0 END tsnSpareQty
                                        ,CASE WHEN a.TD049 = '1' THEN b1.TG048  ELSE  0 END tsrnFreebieQty
                                        ,CASE WHEN a.TD049 = '2' THEN b1.TG048  ELSE  0 END tsrnSpareQty
                                        , b2.TG009 un_tsnQty
                                        ,CASE WHEN a.TD049 = '1' THEN b2.TG044  ELSE  0 END un_tsnFreebieQty
                                        ,CASE WHEN a.TD049 = '2' THEN b2.TG044  ELSE  0 END un_tsnSpareQty
                                        , b3.TI009 un_tsrnQty
                                        ,CASE WHEN a.TD049 = '1' THEN b3.TI035  ELSE  0 END un_tsrnFreebieQty
                                        ,CASE WHEN a.TD049 = '2' THEN b3.TI035  ELSE  0 END un_tsrnSpareQty
                                        , c1.TH008 un_snQty
                                        ,CASE WHEN a.TD049 = '1' THEN c1.TH024  ELSE  0 END un_snFreebieQty
                                        ,CASE WHEN a.TD049 = '2' THEN c1.TH024  ELSE  0 END un_snSpareQty
                                        , c2.TJ007 un_srnQty
                                        ,CASE WHEN a.TD049 = '1' THEN c2.TJ042  ELSE  0 END un_srnFreebieQty
                                        ,CASE WHEN a.TD049 = '2' THEN c2.TJ042  ELSE  0 END un_srnSpareQty
                                        , c3.TH008 snQty
                                        ,CASE WHEN a.TD049 = '1' THEN c3.TH024  ELSE  0 END snFreebieQty
                                        ,CASE WHEN a.TD049 = '2' THEN c3.TH024  ELSE  0 END snSpareQty
                                        , c4.TJ007 srnQty
                                        ,CASE WHEN a.TD049 = '1' THEN c4.TJ042  ELSE  0 END srnFreebieQty
                                        ,CASE WHEN a.TD049 = '2' THEN c4.TJ042  ELSE  0 END srnSpareQty
                                        FROM COPTD a
                                        OUTER APPLY(
                                            --暫出單單身確認
                                            SELECT SUM(x1.TG009) TG009,SUM(x1.TG044)  TG044,SUM(x1.TG020)  TG020 ,SUM(x1.TG021)  TG021,SUM(x1.TG046)  TG046,SUM(x1.TG048)  TG048
                                            FROM  INVTG x1
                                            WHERE x1.TG014 = LTRIM(RTRIM(a.TD001))
                                            AND x1.TG015 = LTRIM(RTRIM(a.TD002))
                                            AND x1.TG016 = LTRIM(RTRIM(a.TD003))
                                            AND x1.TG022 = 'Y'
                                        ) b1
                                        OUTER APPLY(
                                            --暫出單單身未確認
                                            SELECT SUM(x1.TG009) TG009 ,SUM(x1.TG044)  TG044
                                            FROM  INVTG x1
                                            WHERE x1.TG014 = LTRIM(RTRIM(a.TD001))
                                            AND x1.TG015 = LTRIM(RTRIM(a.TD002))
                                            AND x1.TG016 = LTRIM(RTRIM(a.TD003))
                                            AND x1.TG022 = 'N'
                                        ) b2
                                        OUTER APPLY(
                                            --暫出歸還單單身未確認
                                            SELECT SUM(x1.TI009) TI009 ,SUM(x1.TI035)  TI035
                                            FROM  INVTI x1
                                            INNER JOIN INVTG x2 on x2.TG001 = x1.TI014 AND x2.TG002 = x1.TI015 AND x2.TG003 = x1.TI016 
                                            WHERE x2.TG014 = LTRIM(RTRIM(a.TD001))
                                            AND x2.TG015 = LTRIM(RTRIM(a.TD002))
                                            AND x2.TG016 = LTRIM(RTRIM(a.TD003))
                                            AND x1.TI022 = 'N'
                                        ) b3
                                        OUTER APPLY(
                                            --銷貨未確認
                                            SELECT SUM(x1.TH008) TH008 ,SUM(x1.TH024)  TH024
                                            FROM  COPTH x1
                                            WHERE x1.TH014 = LTRIM(RTRIM(a.TD001))
                                            AND x1.TH015 = LTRIM(RTRIM(a.TD002))
                                            AND x1.TH016 = LTRIM(RTRIM(a.TD003))
                                            AND x1.TH020 = 'N'
                                        ) c1
                                        OUTER APPLY(
                                            --銷退未確認
                                            SELECT SUM(x1.TJ007) TJ007 ,SUM(x1.TJ042)  TJ042
                                            FROM  COPTJ x1
                                            WHERE x1.TJ018 = LTRIM(RTRIM(a.TD001))
                                            AND x1.TJ019 = LTRIM(RTRIM(a.TD002))
                                            AND x1.TJ020 = LTRIM(RTRIM(a.TD003))
                                            AND x1.TJ021 = 'N'
                                        ) c2
                                        OUTER APPLY(
                                            --銷貨確認
                                            SELECT SUM(x1.TH008) TH008 ,SUM(x1.TH024)  TH024
                                            FROM  COPTH x1
                                            WHERE x1.TH014 = LTRIM(RTRIM(a.TD001))
                                            AND x1.TH015 = LTRIM(RTRIM(a.TD002))
                                            AND x1.TH016 = LTRIM(RTRIM(a.TD003))
                                            AND x1.TH020 = 'Y'
                                        ) c3
                                        OUTER APPLY(
                                            --銷退確認
                                            SELECT SUM(x1.TJ007) TJ007 ,SUM(x1.TJ042)  TJ042
                                            FROM  COPTJ x1
                                            WHERE x1.TJ018 = LTRIM(RTRIM(a.TD001))
                                            AND x1.TJ019 = LTRIM(RTRIM(a.TD002))
                                            AND x1.TJ020 = LTRIM(RTRIM(a.TD003))
                                            AND x1.TJ021 = 'Y'
                                        ) c4
                                        WHERE LTRIM(RTRIM(a.TD001)) = @SoErpPrefix
                                        AND LTRIM(RTRIM(a.TD002)) = @SoErpNo
                                        AND LTRIM(RTRIM(a.TD003)) = @SoSequence
                                        AND a.TD021 = 'Y'";
                                dynamicParameters.Add("SoErpPrefix", soErpPrefix);
                                dynamicParameters.Add("SoErpNo", soErpNo);
                                dynamicParameters.Add("SoSequence", soSequence);

                                var resultErpAccounting = erpSqlConnection.Query(sql, dynamicParameters);

                                if (resultErpAccounting.Count() > 0)
                                {
                                    foreach (var item in resultErpAccounting)
                                    {
                                        tsnQty = Convert.ToInt32(item.tsnQty);
                                        un_tsnQty = Convert.ToInt32(item.un_tsnQty);
                                        tsrnQty = Convert.ToInt32(item.tsrnQty);
                                        un_tsrnQty = Convert.ToInt32(item.un_tsrnQty);
                                        tsnFreebieQty = Convert.ToInt32(item.tsnFreebieQty);
                                        un_tsnFreebieQty = Convert.ToInt32(item.un_tsnFreebieQty);
                                        tsrnFreebieQty = Convert.ToInt32(item.tsrnFreebieQty);
                                        un_tsrnFreebieQty = Convert.ToInt32(item.un_tsrnFreebieQty);
                                        tsnSpareQty = Convert.ToInt32(item.tsnSpareQty);
                                        un_tsnSpareQty = Convert.ToInt32(item.un_tsnSpareQty);
                                        tsrnSpareQty = Convert.ToInt32(item.tsrnSpareQty);
                                        un_tsrnSpareQty = Convert.ToInt32(item.un_tsrnSpareQty);
                                        snQty = Convert.ToInt32(item.snQty);
                                        un_snQty = Convert.ToInt32(item.un_snQty);
                                        srnQty = Convert.ToInt32(item.srnQty);
                                        un_srnQty = Convert.ToInt32(item.un_srnQty);
                                        snFreebieQty = Convert.ToInt32(item.snFreebieQty);
                                        un_snFreebieQty = Convert.ToInt32(item.un_snFreebieQty);
                                        srnFreebieQty = Convert.ToInt32(item.srnFreebieQty);
                                        un_srnFreebieQty = Convert.ToInt32(item.un_srnFreebieQty);
                                        snSpareQty = Convert.ToInt32(item.snSpareQty);
                                        un_snSpareQty = Convert.ToInt32(item.un_snSpareQty);
                                        srnSpareQty = Convert.ToInt32(item.srnSpareQty);
                                        un_srnSpareQty = Convert.ToInt32(item.un_srnSpareQty);
                                    }
                                }
                                #endregion
                            }

                            #region //正常品檢查
                            int regularQty = tsnQty + un_tsnQty - (tsrnQty + un_tsrnQty) - (srnQty + un_srnQty);

                            if ((tsnQty + un_tsnQty) <= 0 && (snQty + un_snQty) > 0)
                            {
                                regularQty = snQty + un_snQty - (srnQty + un_srnQty);
                            }

                            if (finalize_soQty > (SoQty - regularQty)) throw new SystemException($"訂單【{deliveryData[i]["soNo"].ToString()}】【正常品】撿貨數不足!");
                            #endregion

                            #region //贈品檢查
                            int freebieQty = tsnFreebieQty + un_tsnFreebieQty - (tsrnFreebieQty + un_tsrnFreebieQty) - (srnFreebieQty + un_srnFreebieQty);

                            if ((tsnFreebieQty + un_tsnFreebieQty) <= 0 && (snFreebieQty + un_snFreebieQty) > 0)
                            {
                                freebieQty = snFreebieQty + un_snFreebieQty - (srnFreebieQty + un_srnFreebieQty);
                            }

                            if (finalize_freebieQty > (FreebieQty - freebieQty)) throw new SystemException($"訂單【{deliveryData[i]["soNo"].ToString()}】【贈品】撿貨數不足!");
                            #endregion

                            #region //備品檢查
                            int spareQty = tsnSpareQty + un_tsnSpareQty - (tsrnSpareQty + un_tsrnSpareQty) - (srnSpareQty + un_srnSpareQty);

                            if ((tsnSpareQty + un_tsnSpareQty) <= 0 && (snSpareQty + un_snSpareQty) > 0)
                            {
                                spareQty = snSpareQty + un_snSpareQty - (srnSpareQty + un_srnSpareQty);
                            }

                            if (finalize_spareQty > (SpareQty - spareQty)) throw new SystemException($"訂單【{deliveryData[i]["soNo"].ToString()}】【備品】撿貨數不足!");
                            #endregion
                            #endregion

                            deliveryFinalizes.Add(
                                new DeliveryFinalize
                                {
                                    index = Convert.ToInt32(deliveryData[i]["index"]),
                                    soDetailId = Convert.ToInt32(deliveryData[i]["soDetailId"]),
                                    pcPromiseDate = deliveryData[i]["pcPromiseDate"].ToString(),
                                    pcPromiseTime = deliveryData[i]["pcPromiseTime"].ToString(),
                                    customerNo = deliveryData[i]["customerNo"].ToString(),
                                    customerName = deliveryData[i]["customerName"].ToString(),
                                    soNo = deliveryData[i]["soNo"].ToString(),
                                    soSeq = deliveryData[i]["soSeq"].ToString(),
                                    mtlItemName = deliveryData[i]["mtlItemName"].ToString(),
                                    soQty = Convert.ToDouble(deliveryData[i]["soQty"]),
                                    freebieQty = Convert.ToDouble(deliveryData[i]["freebieQty"]),
                                    spareQty = Convert.ToDouble(deliveryData[i]["spareQty"]),
                                    deliveryCustomer = Convert.ToInt32(deliveryData[i]["deliveryCustomer"]),
                                    deliveryProcess = Convert.ToInt32(deliveryData[i]["deliveryProcess"]),
                                    orderSituation = deliveryData[i]["orderSituation"].ToString(),
                                    deliveryMethod = deliveryData[i]["deliveryMethod"].ToString(),
                                    pcDoDetailRemark = deliveryData[i]["pcDoDetailRemark"].ToString()
                                });
                        }

                        var deliveryCustomers = deliveryFinalizes
                                .GroupBy(x => new
                                {
                                    x.customerNo,
                                    x.customerName,
                                    x.deliveryCustomer,
                                    x.pcPromiseDate,
                                    x.pcPromiseTime
                                })
                                .Select(x => new DeliveryFinalize
                                {
                                    customerNo = x.Key.customerNo,
                                    customerName = x.Key.customerName,
                                    deliveryCustomer = x.Key.deliveryCustomer,
                                    pcPromiseDate = x.Key.pcPromiseDate,
                                    pcPromiseTime = x.Key.pcPromiseTime
                                })
                                .ToList();

                        int rowsAffected = 0;
                        foreach (var dc in deliveryCustomers)
                        {
                            string DoErpNo = BaseHelper.RandomCode(11);

                            var deliverys = deliveryFinalizes
                            .Where(x => x.deliveryCustomer == dc.deliveryCustomer
                                    && x.pcPromiseDate == dc.pcPromiseDate
                                    && x.pcPromiseTime == dc.pcPromiseTime)
                            .OrderBy(x => x.soNo)
                            .ThenBy(x => x.soSeq);

                            #region //撈取所屬部門ID
                            sql = @"SELECT DepartmentId 
                                    FROM BAS.[User]
                                    WHERE UserId = @UserId";
                            dynamicParameters.Add("UserId", CurrentUser);

                            var result = sqlConnection.Query(sql, dynamicParameters);

                            int DepartmentId = -1;
                            foreach (var item in result)
                            {
                                DepartmentId = item.DepartmentId;
                            }
                            #endregion

                            #region //撈取客戶資料
                            sql = @"SELECT a.CustomerId, a.DeliveryAddressFirst
                                    FROM SCM.Customer a
                                    WHERE CompanyId = @CompanyId
                                    AND CustomerNo = @CustomerNo";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("CustomerNo", dc.customerNo);

                            var result2 = sqlConnection.Query(sql, dynamicParameters);

                            int CustomerId = -1;
                            string DoAddressFirst = "";
                            foreach (var item in result2)
                            {
                                CustomerId = item.CustomerId;
                                DoAddressFirst = item.DeliveryAddressFirst;
                            }
                            #endregion

                            #region //判斷出貨單是否存在/撈取DoId
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT DoId, Status
                                    FROM SCM.DeliveryOrder
                                    WHERE DcId = @DcId
                                    AND DoDate = @DoDate";
                            dynamicParameters.Add("DcId", dc.deliveryCustomer);
                            dynamicParameters.Add("DoDate", dc.pcPromiseDate + " " + dc.pcPromiseTime);

                            var result3 = sqlConnection.Query(sql, dynamicParameters);
                            int DoId = -1;
                            string DoStatus = "";
                            foreach (var item in result3)
                            {
                                DoId = item.DoId;
                                DoStatus = item.Status;
                            }
                            #endregion

                            if (DoId < 0)
                            {
                                #region //新增出貨單頭
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO SCM.DeliveryOrder (CompanyId, DepartmentId, UserId, DoErpPrefix, DoErpNo, DoDate, DocDate
                                        , CustomerId, DcId, WayBill, Traffic, ShipMethod, DoAddressFirst, DoAddressSecond, TotalQty, Amount, TaxAmount
                                        , DoRemark, WareHouseDoRemark, MeasureMailStatus, ConfirmStatus, ConfirmUserId, TransferStatus
                                        , TransferDate, Status, CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.DoId
                                        VALUES (@CompanyId, @DepartmentId, @UserId, @DoErpPrefix, @DoErpNo, @DoDate, @DocDate, @CustomerId
                                        , @DcId, @WayBill, @Traffic, @ShipMethod, @DoAddressFirst, @DoAddressSecond, @TotalQty, @Amount, @TaxAmount
                                        , @DoRemark, @WareHouseDoRemark, @MeasureMailStatus, @ConfirmStatus, @ConfirmUserId, @TransferStatus
                                        , @TransferDate, @Status, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        CompanyId = CurrentCompany,
                                        DepartmentId,
                                        UserId = CurrentUser,
                                        DoErpPrefix = 1301,
                                        DoErpNo,
                                        DoDate = dc.pcPromiseDate + " " + dc.pcPromiseTime,
                                        DocDate = dc.pcPromiseDate + " " + dc.pcPromiseTime,
                                        CustomerId,
                                        DcId = dc.deliveryCustomer,
                                        WayBill = "",
                                        Traffic = "",
                                        ShipMethod = "",
                                        DoAddressFirst,
                                        DoAddressSecond = "",
                                        TotalQty = 0,
                                        Amount = 0,
                                        TaxAmount = 0,
                                        DoRemark = "",
                                        WareHouseDoRemark = "",
                                        MeasureMailStatus = "N",
                                        ConfirmStatus = "N",
                                        ConfirmUserId = (int?)null,
                                        TransferStatus = "N",
                                        TransferDate = "",
                                        Status = "N",
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });

                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();

                                foreach (var item in insertResult)
                                {
                                    DoId = Convert.ToInt32(item.DoId);
                                }
                                #endregion

                                foreach (var delivery in deliverys)
                                {
                                    if (delivery.soQty <= 0) throw new SystemException("出貨數量不可需大於0!");

                                    #region //判斷訂單資料是否存在/判斷訂單狀態、數量/撈取訂單單身ID
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.ConfirmStatus, a.ClosureStatus, a.SoDetailId
                                            , a.SoQty, a.FreebieQty, a.SpareQty
                                            FROM SCM.SoDetail a
                                            LEFT JOIN SCM.SaleOrder b ON b.SoId = a.SoId
                                            WHERE b.SoErpPrefix +'-'+ b.SoErpNo = @SoFullNo
                                            AND a.SoSequence = @SoSequence
                                            AND b.CompanyId = @CompanyId";
                                    dynamicParameters.Add("SoFullNo", delivery.soNo);
                                    dynamicParameters.Add("SoSequence", delivery.soSeq);
                                    dynamicParameters.Add("CompanyId", CurrentCompany);

                                    var result4 = sqlConnection.Query(sql, dynamicParameters);

                                    string ConfirmStatus = "";
                                    string ClosureStatus = "";
                                    int SoDetailId = -1;
                                    int SoQty = -1;
                                    int FreebieQty = -1;
                                    int SpareQty = -1;
                                    foreach (var item in result4)
                                    {
                                        ConfirmStatus = item.ConfirmStatus;
                                        ClosureStatus = item.ClosureStatus;
                                        SoDetailId = item.SoDetailId;
                                        SoQty = Convert.ToInt32(item.SoQty);
                                        FreebieQty = Convert.ToInt32(item.FreebieQty);
                                        SpareQty = Convert.ToInt32(item.SpareQty);
                                    }

                                    if (result4.Count() <= 0) throw new SystemException("查無 " + delivery.soNo + "(" + delivery.soSeq + ") 訂單資料!");

                                    if (SoQty < delivery.soQty) throw new SystemException("出貨數量不可大於訂單數量!");
                                    if (ConfirmStatus == "N") throw new SystemException("訂單：" + delivery.soNo + "(" + delivery.soSeq + ") 尚未確認，無法進行建立作業!");
                                    if (ConfirmStatus == "V") throw new SystemException("訂單：" + delivery.soNo + "(" + delivery.soSeq + ") 已作廢，無法進行建立作業!");
                                    if (ClosureStatus == "Y") throw new SystemException("訂單：" + delivery.soNo + "(" + delivery.soSeq + ") 已自動結案，無法進行建立作業!");
                                    if (ClosureStatus == "y") throw new SystemException("訂單：" + delivery.soNo + "(" + delivery.soSeq + ") 已指定結案，無法進行建立作業!");
                                    #endregion

                                    #region //確認出貨數量卡控
                                    #region //查看目前MES出貨單總數量
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT ISNULL(SUM(a.DoQty), 0) TotalDoQty
                                            FROM SCM.DoDetail a 
                                            WHERE a.SoDetailId = @SoDetailId";
                                    dynamicParameters.Add("SoDetailId", SoDetailId);

                                    int totalDoQty = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).TotalDoQty;
                                    #endregion

                                    #region //確認目前此訂單單身的所有檢貨數量
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT ISNULL(SUM(za.ItemQty), 0) TotalItemQty
                                                , ISNULL(SUM(zc.DoQty), 0) TotalPickDoQty
                                                FROM SCM.PickingItem za 
                                                    INNER JOIN SCM.DeliveryOrder zb ON za.DoId = zb.DoId
                                                    INNER JOIN SCM.DoDetail zc ON zb.DoId = zc.DoId AND za.SoDetailId = zc.SoDetailId
                                                WHERE za.SoDetailId = @SoDetailId
                                                    AND za.ItemType = 1
                                                    AND zb.Status = 'S'";
                                    dynamicParameters.Add("SoDetailId", SoDetailId);

                                    int totalItemQty = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).TotalItemQty;
                                    int totalPickDoQty = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).TotalPickDoQty;
                                    #endregion

                                    if (delivery.soQty > ((SoQty - totalDoQty) + (totalPickDoQty - totalItemQty)))
                                    {
                                        throw new SystemException($"訂單：{delivery.soNo}({delivery.soSeq})數量已大於剩餘可出貨數量【{(SoQty - totalDoQty) + (totalPickDoQty - totalItemQty)}】!");
                                    }
                                    #endregion

                                    #region //確認贈品數量卡控
                                    #region //查看目前MES出貨單總贈品數量
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT ISNULL(SUM(a.FreebieQty), 0) TotalFreebieQty
                                            FROM SCM.DoDetail a 
                                            WHERE a.SoDetailId = @SoDetailId";
                                    dynamicParameters.Add("SoDetailId", SoDetailId);

                                    var TotalFreebieQtyResult = sqlConnection.Query(sql, dynamicParameters);

                                    double totalFreebieQty = 0;
                                    if (TotalFreebieQtyResult.Count() > 0)
                                    {
                                        totalFreebieQty = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).TotalFreebieQty;
                                    }
                                    #endregion

                                    #region //確認目前此訂單單身的所有贈品檢貨數量
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT ISNULL(SUM(za.ItemQty), 0) PickFreebieQty
                                                , ISNULL(SUM(zc.FreebieQty), 0) TotalPickFreebieQty
                                                FROM SCM.PickingItem za 
                                                    INNER JOIN SCM.DeliveryOrder zb ON za.DoId = zb.DoId
                                                    INNER JOIN SCM.DoDetail zc ON zb.DoId = zc.DoId AND za.SoDetailId = zc.SoDetailId
                                                WHERE za.SoDetailId = @SoDetailId
                                                    AND za.ItemType = 2
                                                    AND zb.Status = 'S'";
                                    dynamicParameters.Add("SoDetailId", SoDetailId);

                                    var PickingItemQtyResult = sqlConnection.Query(sql, dynamicParameters);

                                    double pickFreebieQty = 0;
                                    double totalPickFreebieQty = 0;
                                    if (PickingItemQtyResult.Count() > 0)
                                    {
                                        pickFreebieQty = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).PickFreebieQty;
                                        totalPickFreebieQty = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).TotalPickFreebieQty;
                                    }
                                    #endregion

                                    if (delivery.freebieQty > ((FreebieQty - totalFreebieQty) + (totalPickFreebieQty - pickFreebieQty)))
                                    {
                                        throw new SystemException($"訂單：{delivery.soNo}({delivery.soSeq})贈品數量已大於剩餘可出貨贈品數量【{((FreebieQty - totalFreebieQty) + (totalPickFreebieQty - pickFreebieQty))}】!");
                                    }
                                    #endregion

                                    #region //確認備品數量卡控
                                    #region //查看目前MES出貨單總備品數量
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT ISNULL(SUM(a.SpareQty), 0) TotalSpareQty
                                            FROM SCM.DoDetail a 
                                            WHERE a.SoDetailId = @SoDetailId";
                                    dynamicParameters.Add("SoDetailId", SoDetailId);

                                    double totalSpareQty = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).TotalSpareQty;
                                    #endregion

                                    #region //確認目前此訂單單身的所有贈品檢貨數量
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT ISNULL(SUM(za.ItemQty), 0) PickSpareQty
                                                , ISNULL(SUM(zc.FreebieQty), 0) TotalPickSpareQty
                                                FROM SCM.PickingItem za 
                                                    INNER JOIN SCM.DeliveryOrder zb ON za.DoId = zb.DoId
                                                    INNER JOIN SCM.DoDetail zc ON zb.DoId = zc.DoId AND za.SoDetailId = zc.SoDetailId
                                                WHERE za.SoDetailId = @SoDetailId
                                                    AND za.ItemType = 3
                                                    AND zb.Status = 'S'";
                                    dynamicParameters.Add("SoDetailId", SoDetailId);

                                    double pickSpareQty = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).PickSpareQty;
                                    double totalPickSpareQty = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).TotalPickSpareQty;
                                    #endregion

                                    if (delivery.spareQty > ((SpareQty - totalSpareQty) + (totalPickSpareQty - pickSpareQty)))
                                    {
                                        throw new SystemException($"訂單：{delivery.soNo}({delivery.soSeq})贈品數量已大於剩餘可出貨贈品數量【{((SpareQty - totalSpareQty) + (totalPickSpareQty - pickSpareQty))}】!");
                                    }
                                    #endregion

                                    #region //查詢出貨單流水號
                                    sql = @"SELECT REPLACE(STR(ISNULL(MAX(DoSequence),0) + 1,4),' ','0') DoSequence 
                                            FROM SCM.DoDetail
                                            WHERE DoId = @DoId";
                                    dynamicParameters.Add("DoId", DoId);

                                    var result5 = sqlConnection.Query(sql, dynamicParameters);

                                    string DoSequence = "";
                                    foreach (var item in result5)
                                    {
                                        DoSequence = item.DoSequence;
                                    }
                                    #endregion

                                    #region //新增出貨單身
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO SCM.DoDetail (DoId, SoDetailId, DoSequence, TransInInventoryId, DoQty, FreebieQty, SpareQty
                                            , UnitPrice, Amount, DoDetailRemark, WareHouseDoDetailRemark, PcDoDetailRemark
                                            , DeliveryProcess, DeliveryRoutine, DeliveryDocument, OrderSituation, DeliveryMethod
                                            , ConfirmStatus, ClosureStatus, CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.DoDetailId
                                            VALUES (@DoId, @SoDetailId, @DoSequence, @TransInInventoryId, @DoQty, @FreebieQty, @SpareQty
                                            , @UnitPrice, @Amount, @DoDetailRemark, @WareHouseDoDetailRemark
                                            , @PcDoDetailRemark, @DeliveryProcess, @DeliveryRoutine, @DeliveryDocument, @OrderSituation, @DeliveryMethod
                                            , @ConfirmStatus, @ClosureStatus, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            DoId,
                                            SoDetailId,
                                            DoSequence,
                                            TransInInventoryId = (int?)null,
                                            DoQty = delivery.soQty,
                                            FreebieQty = delivery.freebieQty,
                                            SpareQty = delivery.spareQty,
                                            UnitPrice = 0,
                                            Amount = 0,
                                            DoDetailRemark = "",
                                            WareHouseDoDetailRemark = "",
                                            PcDoDetailRemark = delivery.pcDoDetailRemark,
                                            DeliveryProcess = delivery.deliveryProcess,
                                            DeliveryRoutine = "",
                                            DeliveryDocument = "",
                                            OrderSituation = delivery.orderSituation,
                                            DeliveryMethod = delivery.deliveryMethod,
                                            ConfirmStatus = "N",
                                            ClosureStatus = "N",
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    insertResult = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult.Count();

                                    #endregion
                                }
                            }
                            else
                            {
                                if (DoStatus == "R" || DoStatus == "S") throw new SystemException("該出貨時間已截止，請選擇其他出貨時間!");

                                #region //修改出貨單頭
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE SCM.DeliveryOrder SET
                                        DoDate = @DoDate,
                                        DocDate = @DocDate,
                                        DcId = @DcId,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE DoId = @DoId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        DoDate = dc.pcPromiseDate + " " + dc.pcPromiseTime,
                                        DocDate = dc.pcPromiseDate + " " + dc.pcPromiseTime,
                                        DcId = dc.deliveryCustomer,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        DoId
                                    });

                                rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                foreach (var delivery in deliverys)
                                {
                                    if (delivery.soQty <= 0) throw new SystemException("出貨數量不可需大於0!");

                                    #region //判斷訂單資料是否存在/判斷訂單狀態、數量/撈取訂單單身ID
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.ConfirmStatus, a.ClosureStatus, a.SoDetailId
                                            , a.SoQty, a.FreebieQty, a.SpareQty
                                            FROM SCM.SoDetail a
                                            LEFT JOIN SCM.SaleOrder b ON b.SoId = a.SoId
                                            WHERE b.SoErpPrefix +'-'+ b.SoErpNo = @SoFullNo
                                            AND a.SoSequence = @SoSequence
                                            AND b.CompanyId = @CompanyId";
                                    dynamicParameters.Add("SoFullNo", delivery.soNo);
                                    dynamicParameters.Add("SoSequence", delivery.soSeq);
                                    dynamicParameters.Add("CompanyId", CurrentCompany);

                                    var result4 = sqlConnection.Query(sql, dynamicParameters);

                                    string ConfirmStatus = "";
                                    string ClosureStatus = "";
                                    int SoDetailId = -1;
                                    int SoQty = -1;
                                    int FreebieQty = -1;
                                    int SpareQty = -1;
                                    foreach (var item in result4)
                                    {
                                        ConfirmStatus = item.ConfirmStatus;
                                        ClosureStatus = item.ClosureStatus;
                                        SoDetailId = item.SoDetailId;
                                        SoQty = Convert.ToInt32(item.SoQty);
                                        FreebieQty = Convert.ToInt32(item.FreebieQty);
                                        SpareQty = Convert.ToInt32(item.SpareQty);
                                    }

                                    if (result4.Count() <= 0) throw new SystemException("查無 " + delivery.soNo + "(" + delivery.soSeq + ") 訂單資料!");

                                    if (SoQty < delivery.soQty) throw new SystemException("出貨數量不可大於訂單數量!");
                                    if (ConfirmStatus == "N") throw new SystemException("訂單：" + delivery.soNo + "(" + delivery.soSeq + ") 尚未確認，無法進行建立作業!");
                                    if (ConfirmStatus == "V") throw new SystemException("訂單：" + delivery.soNo + "(" + delivery.soSeq + ") 已作廢，無法進行建立作業!");
                                    if (ClosureStatus == "Y") throw new SystemException("訂單：" + delivery.soNo + "(" + delivery.soSeq + ") 已自動結案，無法進行建立作業!");
                                    if (ClosureStatus == "y") throw new SystemException("訂單：" + delivery.soNo + "(" + delivery.soSeq + ") 已指定結案，無法進行建立作業!");
                                    #endregion

                                    #region //判斷出貨單身是否存在/撈取DoDetailId
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT DoDetailId
                                            FROM SCM.DoDetail
                                            WHERE DoId = @DoId
                                            AND SoDetailId = @SoDetailId";
                                    dynamicParameters.Add("DoId", DoId);
                                    dynamicParameters.Add("SoDetailId", SoDetailId);

                                    var result5 = sqlConnection.Query(sql, dynamicParameters);
                                    int DoDetailId = -1;
                                    foreach (var item in result5)
                                    {
                                        DoDetailId = item.DoDetailId;
                                    }
                                    #endregion

                                    if (DoDetailId < 0)
                                    {
                                        #region //查詢出貨單流水號
                                        sql = @"SELECT REPLACE(STR(ISNULL(MAX(DoSequence),0) + 1,4),' ','0') DoSequence 
                                                FROM SCM.DoDetail
                                                WHERE DoId = @DoId";
                                        dynamicParameters.Add("DoId", DoId);

                                        var result6 = sqlConnection.Query(sql, dynamicParameters);

                                        string DoSequence = "";
                                        foreach (var item in result6)
                                        {
                                            DoSequence = item.DoSequence;
                                        }
                                        #endregion

                                        #region //新增出貨單身
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO SCM.DoDetail (DoId, SoDetailId, DoSequence, TransInInventoryId, DoQty, FreebieQty, SpareQty
                                                , UnitPrice, Amount, DoDetailRemark, WareHouseDoDetailRemark, PcDoDetailRemark
                                                , DeliveryProcess, DeliveryRoutine, DeliveryDocument, OrderSituation, DeliveryMethod
                                                , ConfirmStatus, ClosureStatus, CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                                OUTPUT INSERTED.DoDetailId
                                                VALUES (@DoId, @SoDetailId, @DoSequence, @TransInInventoryId, @DoQty, @FreebieQty, @SpareQty
                                                , @UnitPrice, @Amount, @DoDetailRemark, @WareHouseDoDetailRemark
                                                , @PcDoDetailRemark, @DeliveryProcess, @DeliveryRoutine, @DeliveryDocument, @OrderSituation, @DeliveryMethod
                                                , @ConfirmStatus, @ClosureStatus, @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                DoId,
                                                SoDetailId,
                                                DoSequence,
                                                TransInInventoryId = (int?)null,
                                                DoQty = delivery.soQty,
                                                FreebieQty = delivery.freebieQty,
                                                SpareQty = delivery.spareQty,
                                                UnitPrice = 0,
                                                Amount = 0,
                                                DoDetailRemark = "",
                                                WareHouseDoDetailRemark = "",
                                                PcDoDetailRemark = delivery.pcDoDetailRemark,
                                                DeliveryProcess = delivery.deliveryProcess,
                                                DeliveryRoutine = "",
                                                DeliveryDocument = "",
                                                OrderSituation = delivery.orderSituation,
                                                DeliveryMethod = delivery.deliveryMethod,
                                                ConfirmStatus = "N",
                                                ClosureStatus = "N",
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
                                        #region //判斷【正常品】出貨數量是否大於已揀數量
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT ISNULL(SUM(ItemQty), 0) RegularItemQty
                                                FROM SCM.PickingItem
                                                WHERE DoId = @DoId
                                                AND ItemType = @ItemType
                                                AND SoDetailId = @SoDetailId";
                                        dynamicParameters.Add("DoId", DoId);
                                        dynamicParameters.Add("ItemType", 1);
                                        dynamicParameters.Add("SoDetailId", SoDetailId);

                                        var resultPickingRegular = sqlConnection.Query(sql, dynamicParameters);
                                        int RegularItemQty = -1;
                                        foreach (var item in resultPickingRegular)
                                        {
                                            RegularItemQty = item.RegularItemQty;
                                        }
                                        if (RegularItemQty > delivery.soQty) throw new SystemException("【正常品】出貨數量不可小於已揀數量!");
                                        #endregion

                                        #region //判斷【贈品】出貨數量是否大於已揀數量
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT ISNULL(SUM(ItemQty), 0) FreebieItemQty
                                                FROM SCM.PickingItem
                                                WHERE DoId = @DoId
                                                AND ItemType = @ItemType
                                                AND SoDetailId = @SoDetailId";
                                        dynamicParameters.Add("DoId", DoId);
                                        dynamicParameters.Add("ItemType", 2);
                                        dynamicParameters.Add("SoDetailId", SoDetailId);

                                        var resultPickingFreebie = sqlConnection.Query(sql, dynamicParameters);
                                        int FreebieItemQty = -1;
                                        foreach (var item in resultPickingFreebie)
                                        {
                                            FreebieItemQty = item.FreebieItemQty;
                                        }
                                        if (FreebieItemQty > delivery.soQty) throw new SystemException("【贈品】出貨數量不可小於已揀數量!");
                                        #endregion

                                        #region //判斷【備品】出貨數量是否大於已揀數量
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT ISNULL(SUM(ItemQty), 0) SpareItemQty
                                                FROM SCM.PickingItem
                                                WHERE DoId = @DoId
                                                AND ItemType = @ItemType
                                                AND SoDetailId = @SoDetailId";
                                        dynamicParameters.Add("DoId", DoId);
                                        dynamicParameters.Add("ItemType", 3);
                                        dynamicParameters.Add("SoDetailId", SoDetailId);

                                        var resultPickingSpare = sqlConnection.Query(sql, dynamicParameters);
                                        int SpareItemQty = -1;
                                        foreach (var item in resultPickingSpare)
                                        {
                                            SpareItemQty = item.SpareItemQty;
                                        }
                                        if (SpareItemQty > delivery.soQty) throw new SystemException("【備品】出貨數量不可小於已揀數量!");
                                        #endregion

                                        #region //修改出貨單身
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"UPDATE SCM.DoDetail SET
                                                DoQty = @DoQty,
                                                FreebieQty = @FreebieQty,
                                                SpareQty = @SpareQty,
                                                PcDoDetailRemark = @PcDoDetailRemark,
                                                DeliveryProcess = @DeliveryProcess,
                                                OrderSituation = @OrderSituation,
                                                DeliveryMethod = @DeliveryMethod,
                                                LastModifiedDate = @LastModifiedDate,
                                                LastModifiedBy = @LastModifiedBy
                                                WHERE DoDetailId = @DoDetailId";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                DoQty = delivery.soQty,
                                                FreebieQty = delivery.freebieQty,
                                                SpareQty = delivery.spareQty,
                                                PcDoDetailRemark = delivery.pcDoDetailRemark,
                                                DeliveryProcess = delivery.deliveryProcess,
                                                OrderSituation = delivery.orderSituation,
                                                DeliveryMethod = delivery.deliveryMethod,
                                                LastModifiedDate,
                                                LastModifiedBy,
                                                DoDetailId
                                            });

                                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                        #endregion
                                    }
                                }
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

        #region //AddDeliveryOrderDateLog -- 交期歷史紀錄新增(另一版本) -- Ann 2025-05-19
        public string AddDeliveryOrderDateLog(int DoDetailId, int SoDetailId, string PcPromiseDate, string PcPromiseTime
            , int CauseType, int DepartmentId, int SupervisorId, string CauseDescription)
        {
            try
            {
                if (SoDetailId <= 0) throw new SystemException("【訂單資料錯誤!】");
                if (PcPromiseDate.Length <= 0) throw new SystemException("【交貨日期】不能為空!");
                if (PcPromiseTime.Length <= 0) throw new SystemException("【交貨時間】不能為空!");

                string dateNow = DateTime.Now.ToString("yyyyMMdd");
                string timeNow = DateTime.Now.ToString("HH:mm:ss");

                int rowsAffected = 0;

                DateTime PcPromiseDateTime = Convert.ToDateTime(PcPromiseDate + " " + PcPromiseTime);

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    string modifier = "", soErpPrefix = "", soErpNo = "", soSequence = "";

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

                        #region //判斷出貨單是否已產出暫出單
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ToId
                                FROM SCM.TemporaryOrder a
                                INNER JOIN SCM.DeliveryOrder b ON b.DoId = a.DoId 
                                INNER JOIN SCM.DoDetail c ON c.DoId = b.DoId
                                WHERE c.DoDetailId = @DoDetailId";
                        dynamicParameters.Add("DoDetailId", DoDetailId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        #endregion

                        if (result2.Count() <= 0)
                        {
                            #region //判斷交期修改次數
                            sql = @"SELECT DeliveryDateLogId
                                    FROM SCM.DeliveryDateLog
                                    WHERE SoDetailId = @SoDetailId";
                            dynamicParameters.Add("SoDetailId", SoDetailId);

                            var result3 = sqlConnection.Query(sql, dynamicParameters);
                            if (result3.Count() >= 2)
                            {
                                if (CauseType <= 0) throw new SystemException("【原因類別】不能為空!");
                                if (DepartmentId <= 0) throw new SystemException("【責任部門】不能為空!");
                                if (SupervisorId <= 0) throw new SystemException("【責任主管】不能為空!");
                                if (CauseDescription.Length <= 0) throw new SystemException("【延遲原因】不能為空!");
                                if (CauseDescription.Length > 255) throw new SystemException("【延遲原因】長度錯誤!");
                            }
                            #endregion

                            #region //撈取原排定交貨日
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT PcPromiseDate
                                    FROM SCM.SoDetail
                                    WHERE SoDetailId = @SoDetailId";
                            dynamicParameters.Add("SoDetailId", SoDetailId);

                            var result4 = sqlConnection.Query(sql, dynamicParameters);

                            DateTime? pcPromiseDate = new DateTime();
                            foreach (var item in result4)
                            {
                                pcPromiseDate = item.PcPromiseDate;
                            }
                            #endregion

                            #region //撈取原修改紀錄ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT ISNULL(MAX(DeliveryDateLogId), -1) AS ParentLogId
                                    FROM SCM.DeliveryDateLog
                                    WHERE SoDetailId = @SoDetailId";
                            dynamicParameters.Add("SoDetailId", SoDetailId);

                            var result5 = sqlConnection.Query(sql, dynamicParameters);

                            int ParentLogId = -1;
                            foreach (var item in result5)
                            {
                                ParentLogId = item.ParentLogId;
                            }
                            #endregion

                            #region //取得原出貨日期
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT b.DoDate
                                    FROM SCM.DoDetail a 
                                    INNER JOIN SCM.DeliveryOrder b ON a.DoId = b.DoId
                                    WhERE a.DoDetailId = @DoDetailId";
                            dynamicParameters.Add("DoDetailId", DoDetailId);

                            DateTime OriDoDate = sqlConnection.QueryFirst(sql, dynamicParameters).DoDate;
                            #endregion

                            #region //新增交期修改紀錄
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO SCM.DeliveryDateLog (ParentLogId, SoDetailId, DoDetailId, PcPromiseDate, DoDate, OriDoDate, DepartmentId
                                    , SupervisorId, CauseType, CauseDescription
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.DeliveryDateLogId
                                    VALUES (@ParentLogId, @SoDetailId, @DoDetailId, @PcPromiseDate, @DoDate, @OriDoDate, @DepartmentId
                                    , @SupervisorId, @CauseType, @CauseDescription
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    ParentLogId = ParentLogId > 0 ? (int?)ParentLogId : null,
                                    SoDetailId,
                                    DoDetailId,
                                    PcPromiseDate = pcPromiseDate != null ? pcPromiseDate : OriDoDate,
                                    DoDate = PcPromiseDateTime,
                                    OriDoDate,
                                    CauseType = CauseType > 0 ? (int?)CauseType : null,
                                    DepartmentId = DepartmentId > 0 ? (int?)DepartmentId : null,
                                    SupervisorId = SupervisorId > 0 ? (int?)SupervisorId : null,
                                    CauseDescription = CauseDescription.Length > 0 ? CauseDescription : null,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            var logResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += logResult.Count();
                            #endregion

                            if (DoDetailId > 0)
                            {
                                #region //撈取原出貨單Id
                                sql = @"SELECT DoId
                                        FROM SCM.DeliveryOrder
                                        WHERE DoId = (
                                            SELECT DoId
                                            FROM SCM.DoDetail
                                            WHERE DoDetailId = @DoDetailId
                                        )";
                                dynamicParameters.Add("DoDetailId", DoDetailId);

                                var result6 = sqlConnection.Query(sql, dynamicParameters);
                                int oriDoId = -1;
                                foreach (var item in result6)
                                {
                                    oriDoId = Convert.ToInt32(item.DoId);
                                }
                                #endregion

                                #region //判斷新時段是否有同出貨客戶的出貨單
                                sql = @"SELECT a.DoId
                                        FROM SCM.DeliveryOrder a
                                        WHERE a.DoDate = @DoDate
                                        AND a.DcId = (
                                            SELECT y.DcId
                                            FROM SCM.DoDetail z
                                            INNER JOIN SCM.DeliveryOrder y ON z.DoId = y.DoId
                                            WHERE z.DoDetailId = @DoDetailId
                                        )";
                                dynamicParameters.Add("DoDate", PcPromiseDateTime);
                                dynamicParameters.Add("DoDetailId", DoDetailId);

                                var result7 = sqlConnection.Query(sql, dynamicParameters);

                                int doId = -1;
                                foreach (var item in result7)
                                {
                                    doId = item.DoId;
                                }
                                #endregion

                                if (doId < 0)
                                {
                                    #region //複製出貨單
                                    dynamicParameters = new DynamicParameters();
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
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            DoErpNo = BaseHelper.RandomCode(11),
                                            DoDate = PcPromiseDateTime,
                                            DocDate = CreateDate,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy,
                                            DoDetailId
                                        });
                                    var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                    foreach (var item in insertResult)
                                    {
                                        doId = item.DoId;
                                    }

                                    rowsAffected += insertResult.Count();
                                    #endregion
                                }

                                #region //修改出貨單身的單頭
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE SCM.DoDetail SET
                                        DoId = @DoId,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE SoDetailId = @SoDetailId
                                        AND DoId = @oriDoId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        DoId = doId,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        SoDetailId,
                                        oriDoId
                                    });

                                rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region //修改揀貨DoId
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE SCM.PickingItem SET
                                        DoId = @DoId,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE SoDetailId = @SoDetailId
                                        AND DoId = @oriDoId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        DoId = doId,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        SoDetailId,
                                        oriDoId
                                    });

                                rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region //判斷原出貨單是否還有單身
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM SCM.DoDetail
                                        WHERE DoId = @DoId";
                                dynamicParameters.Add("DoId", oriDoId);

                                var result8 = sqlConnection.Query(sql, dynamicParameters);

                                if (result8.Count() <= 0)
                                {
                                    #region //刪除原出貨單
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE SCM.DeliveryOrder
                                            WHERE DoId = @DoId";
                                    dynamicParameters.Add("DoId", oriDoId);

                                    rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                                #endregion
                            }
                        }
                        else throw new SystemException("該出貨單已產出暫出單，無法修改交期!");

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

        #region //AddDeliveryQtyLog -- 新增更改出貨數量紀錄 -- Ann 2025-05-23
        public string AddDeliveryQtyLog(int DoDetailId, int DoQty
            , int CauseType, int DepartmentId, int SupervisorId, string CauseDescription)
        {
            try
            {
                if (DoQty <= 0) throw new SystemException("出貨數量需至少大於0!!");

                int rowsAffected = 0;

                using (TransactionScope transactionScope = new TransactionScope())
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

                        #region //確認出貨單資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.SoDetailId, a.DoQty
                                , b.SoQty
                                , c.DoDate
                                FROM SCM.DoDetail a 
                                INNER JOIN SCM.SoDetail b ON a.SoDetailId = b.SoDetailId
                                INNER JOIN SCM.DeliveryOrder c ON a.DoId = c.DoId
                                WHERE a.DoDetailId = @DoDetailId";
                        dynamicParameters.Add("DoDetailId", DoDetailId);

                        var DoDetailResult = sqlConnection.Query(sql, dynamicParameters);

                        if (DoDetailResult.Count() <= 0) throw new SystemException("【出貨單】資料錯誤!");

                        int SoDetailId = DoDetailResult.FirstOrDefault().SoDetailId;
                        int OriDoQty = DoDetailResult.FirstOrDefault().DoQty;
                        DateTime OriDoDate = DoDetailResult.FirstOrDefault().DoDate;
                        int SoQty = Convert.ToInt32(DoDetailResult.FirstOrDefault().SoQty);
                        #endregion

                        #region //判斷出貨單是否已產出暫出單
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ToId
                                FROM SCM.TemporaryOrder a
                                INNER JOIN SCM.DeliveryOrder b ON b.DoId = a.DoId 
                                INNER JOIN SCM.DoDetail c ON c.DoId = b.DoId
                                WHERE c.DoDetailId = @DoDetailId";
                        dynamicParameters.Add("DoDetailId", DoDetailId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        #endregion

                        if (result2.Count() <= 0)
                        {
                            #region //判斷交期修改次數
                            sql = @"SELECT DeliveryDateLogId
                                    FROM SCM.DeliveryDateLog
                                    WHERE SoDetailId = @SoDetailId";
                            dynamicParameters.Add("SoDetailId", SoDetailId);

                            var result3 = sqlConnection.Query(sql, dynamicParameters);
                            if (result3.Count() >= 2)
                            {
                                if (CauseType <= 0) throw new SystemException("【原因類別】不能為空!");
                                if (DepartmentId <= 0) throw new SystemException("【責任部門】不能為空!");
                                if (SupervisorId <= 0) throw new SystemException("【責任主管】不能為空!");
                                if (CauseDescription.Length <= 0) throw new SystemException("【延遲原因】不能為空!");
                                if (CauseDescription.Length > 255) throw new SystemException("【延遲原因】長度錯誤!");
                            }
                            #endregion

                            #region //撈取原排定交貨日
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT PcPromiseDate
                                    FROM SCM.SoDetail
                                    WHERE SoDetailId = @SoDetailId";
                            dynamicParameters.Add("SoDetailId", SoDetailId);

                            var result4 = sqlConnection.Query(sql, dynamicParameters);

                            DateTime pcPromiseDate = new DateTime();
                            foreach (var item in result4)
                            {
                                pcPromiseDate = item.PcPromiseDate;
                            }
                            #endregion

                            #region //撈取原修改紀錄ID
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT ISNULL(MAX(DeliveryDateLogId), -1) AS ParentLogId
                                    FROM SCM.DeliveryDateLog
                                    WHERE DoDetailId = @DoDetailId";
                            dynamicParameters.Add("DoDetailId", DoDetailId);

                            var result5 = sqlConnection.Query(sql, dynamicParameters);

                            int ParentLogId = -1;
                            foreach (var item in result5)
                            {
                                ParentLogId = item.ParentLogId;
                            }
                            #endregion

                            #region //確認修改數量是否有超過剩餘可出貨數量(正常品)
                            //檢貨數量(全部)
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT ISNULL(SUM(x.ItemQty), 0) PickRegularQty
                                    FROM SCM.PickingItem x
                                    INNER JOIN SCM.DoDetail xa ON x.DoId = xa.DoId AND x.SoDetailId = xa.SoDetailId
                                    INNER JOIN SCM.DeliveryOrder xb ON xa.DoId = xb.DoId
                                    WHERE x.SoDetailId = @SoDetailId
                                    AND x.ItemType = 1
                                    AND xb.Status = 'S'";
                            dynamicParameters.Add("SoDetailId", SoDetailId);

                            int PickRegularQty = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).PickRegularQty;

                            //檢貨數量(此張出貨單)
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT ISNULL(SUM(x.ItemQty), 0) PickRegularQty
                                    FROM SCM.PickingItem x
                                    INNER JOIN SCM.DoDetail xa ON x.DoId = xa.DoId AND x.SoDetailId = xa.SoDetailId
                                    WHERE xa.DoDetailId = @DoDetailId
                                    AND x.ItemType = 1";
                            dynamicParameters.Add("DoDetailId", DoDetailId);

                            int currentPickRegularQty = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).PickRegularQty;

                            if (DoQty < currentPickRegularQty)
                            {
                                throw new SystemException($"更改數量不可超過已檢貨數量!!");
                            }

                            //出貨總數
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT ISNULL(SUM(a.DoQty), 0) TotalDoQty
                                    FROM SCM.DoDetail a 
                                    WHERE a.SoDetailId = @SoDetailId
                                    AND a.DoDetailId != @DoDetailId";
                            dynamicParameters.Add("SoDetailId", SoDetailId);
                            dynamicParameters.Add("DoDetailId", DoDetailId);

                            int TotalDoQty = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).TotalDoQty;

                            if (DoQty > (SoQty - TotalDoQty - PickRegularQty))
                            {
                                throw new SystemException($"該數量超過剩餘可出貨數量!<br> 剩餘數量: 訂單總數量{SoQty} - 目前總出貨單數量{TotalDoQty} - 目前總檢貨數量{PickRegularQty} = {SoQty - TotalDoQty - PickRegularQty}");
                            }

                            if (DoQty < currentPickRegularQty)
                            {
                                throw new SystemException($"該數量超過已檢貨數量【{currentPickRegularQty}】!");
                            }
                            #endregion

                            #region //新增交期修改紀錄
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO SCM.DeliveryDateLog (ParentLogId, SoDetailId, DoDetailId, PcPromiseDate, DoQty, OriDoQty, DoDate, OriDoDate, DepartmentId
                                    , SupervisorId, CauseType, CauseDescription
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.DeliveryDateLogId
                                    VALUES (@ParentLogId, @SoDetailId, @DoDetailId, @PcPromiseDate, @DoQty, @OriDoQty, @DoDate, @OriDoDate, @DepartmentId
                                    , @SupervisorId, @CauseType, @CauseDescription
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    ParentLogId = ParentLogId > 0 ? (int?)ParentLogId : null,
                                    SoDetailId,
                                    DoDetailId,
                                    PcPromiseDate = pcPromiseDate,
                                    DoQty,
                                    OriDoQty,
                                    DoDate = OriDoDate,
                                    OriDoDate,
                                    CauseType = CauseType > 0 ? (int?)CauseType : null,
                                    DepartmentId = DepartmentId > 0 ? (int?)DepartmentId : null,
                                    SupervisorId = SupervisorId > 0 ? (int?)SupervisorId : null,
                                    CauseDescription = CauseDescription.Length > 0 ? CauseDescription : null,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });
                            var logResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += logResult.Count();
                            #endregion

                            #region //更改出貨單數量
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.DoDetail
                                    SET DoQty = @DoQty
                                    WHERE DoDetailId = @DoDetailId";
                            dynamicParameters.Add("DoQty", DoQty);
                            dynamicParameters.Add("DoDetailId", DoDetailId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }
                        else throw new SystemException("該出貨單已產出暫出單，無法修改數量!");

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
        #region //UpdateWipLink -- 訂單製令綁定 -- Ben Ma 2023.04.27
        public string UpdateWipLink(int SoDetailId, string WipNo, string LinkStatus)
        {
            try
            {
                if (WipNo.Length <= 0) throw new SystemException("【製令】不能為空!");
                if (!Regex.IsMatch(LinkStatus, "^(N|Y)$", RegexOptions.IgnoreCase)) throw new SystemException("【綁定狀態】錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0;
                    string mtlItemNo = "", soErpPrefix = "", soErpNo = "", soSequence = "";

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

                        #region //判斷訂單資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT c.MtlItemNo, b.SoErpPrefix, b.SoErpNo, a.SoSequence
                                FROM SCM.SoDetail a
                                INNER JOIN SCM.SaleOrder b ON a.SoId = b.SoId
                                INNER JOIN PDM.MtlItem c ON a.MtlItemId = c.MtlItemId
                                WHERE a.ConfirmStatus = @ConfirmStatus
                                AND b.ConfirmStatus = @ConfirmStatus
                                AND a.SoDetailId = @SoDetailId";
                        dynamicParameters.Add("ConfirmStatus", "Y");
                        dynamicParameters.Add("SoDetailId", SoDetailId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("【訂單】資料錯誤!");

                        foreach (var item in result2)
                        {
                            mtlItemNo = item.MtlItemNo;
                            soErpPrefix = item.SoErpPrefix;
                            soErpNo = item.SoErpNo;
                            soSequence = item.SoSequence;
                        }
                        #endregion

                        #region //新增綁定&解綁紀錄
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.WipLinkLog (SoErpPrefix, SoErpNo, SoSequence
                                , WoErpPrefix, WoErpNo, BindStatus
                                , CreateDate, CreateBy)
                                OUTPUT INSERTED.WipLinkLogId
                                VALUES (@SoErpPrefix, @SoErpNo, @SoSequence
                                , @WoErpPrefix, @WoErpNo, @BindStatus
                                , @CreateDate, @CreateBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SoErpPrefix = soErpPrefix,
                                SoErpNo = soErpNo,
                                SoSequence = soSequence,
                                WoErpPrefix = WipNo.Split('-')[0],
                                WoErpNo = WipNo.Split('-')[1],
                                BindStatus = LinkStatus == "Y" ? "B" : "U",
                                CreateDate,
                                CreateBy
                            });
                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += insertResult.Count();
                        #endregion
                    }

                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        #region //判斷製令資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MOCTA
                                WHERE TA013 = @ConfirmStatus
                                AND TA001 = @WoErpPrefix
                                AND TA002 = @WoErpNo";
                        dynamicParameters.Add("ConfirmStatus", "Y");
                        dynamicParameters.Add("WoErpPrefix", WipNo.Split('-')[0]);
                        dynamicParameters.Add("WoErpNo", WipNo.Split('-')[1]);

                        var resultWip = sqlConnection.Query(sql, dynamicParameters);
                        if (resultWip.Count() <= 0) throw new SystemException("【ERP製令】資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MOCTA SET
                                TA026 = @SoErpPrefix,
                                TA027 = @SoErpNo,
                                TA028 = @SoSequence
                                WHERE TA001 = @WoErpPrefix
                                AND TA002 = @WoErpNo";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                SoErpPrefix = LinkStatus == "Y" ? soErpPrefix : "",
                                SoErpNo = LinkStatus == "Y" ? soErpNo : "",
                                SoSequence = LinkStatus == "Y" ? soSequence : "",
                                WoErpPrefix = WipNo.Split('-')[0],
                                WoErpNo = WipNo.Split('-')[1]
                            });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //UpdateLotDeliveryFinalize -- 出貨定版異動 -- Ben Ma 2024.07.31
        public string UpdateLotDeliveryFinalize(string DoDetails, string PcPromiseDate, string PcPromiseTime
            , int CauseType, int DepartmentId, int SupervisorId, string CauseDescription)
        {
            try
            {
                if (PcPromiseDate.Length <= 0) throw new SystemException("【交貨日期】不能為空!");
                if (PcPromiseTime.Length <= 0) throw new SystemException("【交貨時間】不能為空!");

                string dateNow = DateTime.Now.ToString("yyyyMMdd");
                string timeNow = DateTime.Now.ToString("HH:mm:ss");

                int rowsAffected = 0;

                DateTime PcPromiseDateTime = Convert.ToDateTime(PcPromiseDate + " " + PcPromiseTime);

                using (TransactionScope transactionScope = new TransactionScope())
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

                        List<int> DoDetailList = DoDetails.Split(',').Select(x => Convert.ToInt32(x)).ToList();

                        foreach (var DoDetailId in DoDetailList)
                        {
                            string modifier = "", soErpPrefix = "", soErpNo = "", soSequence = "";

                            #region //判斷定版資料是否正確
                            int DoId = -1, SoDetailId = -1;

                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 DoId, SoDetailId
                                    FROM SCM.DoDetail
                                    WHERE DoDetailId = @DoDetailId";
                            dynamicParameters.Add("DoDetailId", DoDetailId);

                            var resultDoDetail = sqlConnection.Query(sql, dynamicParameters);
                            if (resultDoDetail.Count() <= 0) throw new SystemException("定版資料錯誤!");

                            foreach (var item in resultDoDetail)
                            {
                                DoId = Convert.ToInt32(item.DoId);
                                SoDetailId = Convert.ToInt32(item.SoDetailId);
                            }
                            #endregion

                            #region //判斷是否已有撿貨資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT COUNT(a.ItemQty) PickingQty
                                    FROM SCM.PickingItem a
                                    OUTER APPLY (
                                        SELECT ba.DoId, ba.SoDetailId
                                        FROM SCM.DoDetail ba
                                        WHERE ba.DoDetailId = @DoDetailId
                                    ) b
                                    WHERE a.DoId = b.DoId
                                    AND a.SoDetailId = b.SoDetailId";
                            dynamicParameters.Add("DoDetailId", DoDetailId);

                            var resultPickingItem = sqlConnection.Query(sql, dynamicParameters);
                            if (resultPickingItem.Count() >= 0)
                            {
                                foreach (var item in resultPickingItem)
                                {
                                    if (Convert.ToInt32(item.PickingQty) > 0) throw new SystemException("該物件已經撿貨，無法異動!");
                                }
                            }
                            #endregion

                            #region //交期歷史紀錄新增
                            #region //判斷交期修改次數
                            sql = @"SELECT DeliveryDateLogId
                                    FROM SCM.DeliveryDateLog
                                    WHERE SoDetailId = @SoDetailId";
                            dynamicParameters.Add("SoDetailId", SoDetailId);

                            var resultLog = sqlConnection.Query(sql, dynamicParameters);
                            if (resultLog.Count() >= 2)
                            {
                                if (CauseType <= 0) throw new SystemException("【原因類別】不能為空!");
                                if (DepartmentId <= 0) throw new SystemException("【責任部門】不能為空!");
                                if (SupervisorId <= 0) throw new SystemException("【責任主管】不能為空!");
                                if (CauseDescription.Length <= 0) throw new SystemException("【延遲原因】不能為空!");
                                if (CauseDescription.Length > 255) throw new SystemException("【延遲原因】長度錯誤!");
                            }
                            #endregion

                            #region //撈取原排定交貨日
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT PcPromiseDate
                                    FROM SCM.SoDetail
                                    WHERE SoDetailId = @SoDetailId";
                            dynamicParameters.Add("SoDetailId", SoDetailId);

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
                            dynamicParameters.Add("SoDetailId", SoDetailId);

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
                                    SoDetailId,
                                    PcPromiseDate = pcPromiseDate,
                                    CauseType = CauseType > 0 ? (int?)CauseType : null,
                                    DepartmentId = DepartmentId > 0 ? (int?)DepartmentId : null,
                                    SupervisorId = SupervisorId > 0 ? (int?)SupervisorId : null,
                                    CauseDescription = CauseDescription.Length > 0 ? CauseDescription : null,
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
                                    PcPromiseDate = PcPromiseDateTime,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    SoDetailId
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //撈取修改人員No
                            sql = @"SELECT UserNo
                                    FROM BAS.[User]
                                    WHERE UserId = @UserId";
                            dynamicParameters.Add("UserId", LastModifiedBy);

                            var resultUser = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item in resultUser)
                            {
                                modifier = item.UserNo;
                            }
                            #endregion

                            #region //撈取單別/單號/流水號
                            sql = @"SELECT b.SoErpPrefix, b.SoErpNo, a.SoSequence 
                                    FROM SCM.SoDetail a
                                    INNER JOIN SCM.SaleOrder b ON b.SoId = a.SoId
                                    WHERE SoDetailId = @SoDetailId";
                            dynamicParameters.Add("SoDetailId", SoDetailId);

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
                                        TD048 = DateTime.ParseExact(PcPromiseDate, "yyyy-MM-dd", null).ToString("yyyyMMdd"),
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
                            dynamicParameters.Add("DoDate", PcPromiseDateTime);
                            dynamicParameters.Add("DoDetailId", DoDetailId);

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

                                #region //移動出貨單單身
                                sql = @"UPDATE SCM.DoDetail SET
                                        DoId = @NewDoId,
                                        DoSequence = @DoSequence,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE DoDetailId = @DoDetailId";
                                rowsAffected += sqlConnection.Execute(sql,
                                    new
                                    {
                                        NewDoId = existDoId,
                                        DoSequence,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        DoDetailId
                                    });
                                #endregion

                                #region //複製出貨單單身
                                //sql = @"INSERT INTO SCM.DoDetail
                                //        SELECT @DoId, a.SoDetailId, @DoSequence, a.TransInInventoryId, a.DoQty, a.FreebieQty, a.SpareQty
                                //        , a.UnitPrice, a.Amount, a.DoDetailRemark, a.WareHouseDoDetailRemark
                                //        , a.PcDoDetailRemark, a.DeliveryProcess, a.DeliveryRoutine, a.DeliveryDocument
                                //        , a.OrderSituation, a.DeliveryMethod, a.ConfirmStatus, a.ClosureStatus
                                //        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy
                                //        FROM SCM.DoDetail a
                                //        WHERE a.DoDetailId = @DoDetailId";
                                //rowsAffected += sqlConnection.Execute(sql,
                                //    new
                                //    {
                                //        DoId = existDoId,
                                //        DoSequence,
                                //        TransInInventoryId = -1,
                                //        CreateDate,
                                //        LastModifiedDate,
                                //        CreateBy,
                                //        LastModifiedBy,
                                //        DoDetailId
                                //    });
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
                                        DoDate = PcPromiseDateTime,
                                        DocDate = CreateDate,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy,
                                        DoDetailId
                                    });

                                rowsAffected += insertResult.Count();

                                int newDoId = insertResult.Select(x => x.DoId).FirstOrDefault();
                                #endregion

                                #region //移動出貨單單身
                                sql = @"UPDATE SCM.DoDetail SET
                                        DoId = @NewDoId,
                                        DoSequence = @DoSequence,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE DoDetailId = @DoDetailId";
                                rowsAffected += sqlConnection.Execute(sql,
                                    new
                                    {
                                        NewDoId = newDoId,
                                        DoSequence = "0001",
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        DoDetailId
                                    });
                                #endregion

                                #region //複製出貨單單身
                                //sql = @"INSERT INTO SCM.DoDetail
                                //        SELECT @DoId, a.SoDetailId, @DoSequence, a.TransInInventoryId, a.DoQty, a.FreebieQty, a.SpareQty
                                //        , a.UnitPrice, a.Amount, a.DoDetailRemark, a.WareHouseDoDetailRemark
                                //        , a.PcDoDetailRemark, a.DeliveryProcess, a.DeliveryRoutine, a.DeliveryDocument
                                //        , a.OrderSituation, a.DeliveryMethod, a.ConfirmStatus, a.ClosureStatus
                                //        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy
                                //        FROM SCM.DoDetail a
                                //        WHERE a.DoDetailId = @DoDetailId";
                                //rowsAffected += sqlConnection.Execute(sql,
                                //    new
                                //    {
                                //        DoId = newDoId,
                                //        DoSequence = "0001",
                                //        TransInInventoryId = -1,
                                //        CreateDate,
                                //        LastModifiedDate,
                                //        CreateBy,
                                //        LastModifiedBy,
                                //        DoDetailId
                                //    });
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
        #endregion

        #region //Delete
        #region //DeleteDeliveryFinalize -- 出貨定版資料刪除 -- Zoey 2022.09.19
        public string DeleteDeliveryFinalize(int DoDetailId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷定版資料是否正確
                        int DoId = -1;

                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 DoId
                                FROM SCM.DoDetail
                                WHERE DoDetailId = @DoDetailId";
                        dynamicParameters.Add("DoDetailId", DoDetailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("定版資料錯誤!");

                        foreach (var item in result)
                        {
                            DoId = Convert.ToInt32(item.DoId);
                        }
                        #endregion

                        #region //判斷是否已有撿貨資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT COUNT(a.ItemQty) PickingQty
                                FROM SCM.PickingItem a
                                OUTER APPLY (
                                    SELECT ba.DoId, ba.SoDetailId
                                    FROM SCM.DoDetail ba
                                    WHERE ba.DoDetailId = @DoDetailId
                                ) b
                                WHERE a.DoId = b.DoId
                                AND a.SoDetailId = b.SoDetailId";
                        dynamicParameters.Add("DoDetailId", DoDetailId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() >= 0)
                        {
                            foreach (var item in result2)
                            {
                                if (Convert.ToInt32(item.PickingQty) > 0) throw new SystemException("該物件已經撿貨，無法刪除!");
                            }
                        }
                        #endregion

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.DoDetail
                                WHERE DoDetailId = @DoDetailId";
                        dynamicParameters.Add("DoDetailId", DoDetailId);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //判斷出貨單沒有單身，即刪除單頭
                        #region //目前單身數量
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT COUNT(1) TotalDetail
                                FROM SCM.DoDetail
                                WHERE DoId = @DoId";
                        dynamicParameters.Add("DoId", DoId);

                        var resultDoDetail = sqlConnection.Query(sql, dynamicParameters);

                        int TotalDetail = 0;
                        foreach (var item in resultDoDetail)
                        {
                            TotalDetail = Convert.ToInt32(item.TotalDetail);
                        }
                        #endregion

                        #region //刪除單頭
                        if (TotalDetail <= 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE SCM.DeliveryOrder
                                    WHERE DoId = @DoId";
                            dynamicParameters.Add("DoId", DoId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //MailAdvice
        #region //DailyDeliveryWipUnLinkDetailMailAdvice -- 每日出貨未綁定製令明細 -- Ben Ma 2023.05.24
        public string DailyDeliveryWipUnLinkDetailMailAdvice(string CompanyNo)
        {
            try
            {
                int CompanyId = -1;
                List<DailyDeliveryWipUnLinkDetail> dailyDeliveryWipUnLinkDetails = new List<DailyDeliveryWipUnLinkDetail>();
                List<WipLink> wipLinks = new List<WipLink>();

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //判斷公司資料是否正確
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 CompanyId, ErpDb
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

                    #region //出貨未綁定製令明細
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT FORMAT(b.DoDate, 'yyyy-MM-dd') DoDate, c.DcShortName
                            , e.SoErpPrefix + '-' + e.SoErpNo + '-' + d.SoSequence SoErpFullNo
                            , d.SoQty
                            , f.MtlItemNo, ISNULL(d.SoMtlItemName, f.MtlItemName) MtlItemName, ISNULL(d.SoMtlItemSpec, f.MtlItemSpec) MtlItemSpec
                            , g.UserNo PcUserNo, g.UserName PcUserName
                            , ISNULL(h.PickQty, 0) PickQty
                            , i.DepartmentName
                            FROM SCM.DoDetail a
                            INNER JOIN SCM.DeliveryOrder b ON a.DoId = b.DoId
                            INNER JOIN SCM.DeliveryCustomer c ON b.DcId = c.DcId
                            INNER JOIN SCM.SoDetail d ON a.SoDetailId = d.SoDetailId
                            INNER JOIN SCM.SaleOrder e ON d.SoId = e.SoId
                            INNER JOIN PDM.MtlItem f ON d.MtlItemId = f.MtlItemId
                            INNER JOIN BAS.[User] g ON b.CreateBy = g.UserId
                            OUTER APPLY (
                                SELECT SUM(x.ItemQty) PickQty
                                FROM SCM.PickingItem x
                                WHERE x.SoDetailId = a.SoDetailId
                                AND x.DoId = a.DoId
                            ) h
                            INNER JOIN BAS.[Department] i ON g.DepartmentId = i.DepartmentId
                            WHERE b.CompanyId = @CompanyId
                            AND b.DoDate >= @StartDate
                            AND b.DoDate <= @EndDate
                            AND c.DcId NOT IN @DcId";
                    dynamicParameters.Add("CompanyId", CompanyId);
                    dynamicParameters.Add("StartDate", DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 00:00:00"));
                    dynamicParameters.Add("EndDate", DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 23:59:59"));
                    dynamicParameters.Add("DcId", new string[] { "16" });

                    dailyDeliveryWipUnLinkDetails = sqlConnection.Query<DailyDeliveryWipUnLinkDetail>(sql, dynamicParameters).ToList();
                    #endregion
                }

                using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                {
                    #region //ERP製令綁訂定單資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.TA001 WoErpPrefix, a.TA002 WoErpNo
                            , ISNULL(a.TA026, '') SoErpPrefix, ISNULL(a.TA027, '') SoErpNo
                            , ISNULL(a.TA028, '') SoSequence
                            , CASE WHEN LEN(a.TA026 + a.TA027 + a.TA028) > 0 
                                THEN 'Y'
                                ELSE 'N' END BindSoStatus
                            FROM MOCTA a
                            WHERE a.TA026 + '-' + a.TA027 + '-' + a.TA028 IN @SoErpFullNo
                            ORDER BY a.TA002 DESC, a.TA001";
                    dynamicParameters.Add("SoErpFullNo", dailyDeliveryWipUnLinkDetails.Select(x => x.SoErpFullNo).ToArray());

                    wipLinks = sqlConnection.Query<WipLink>(sql, dynamicParameters).ToList();
                    dailyDeliveryWipUnLinkDetails = dailyDeliveryWipUnLinkDetails.GroupJoin(wipLinks, x => x.SoErpFullNo, y => y.SoErpPrefix + '-' + y.SoErpNo + '-' + y.SoSequence, (x, y) => { x.WipLink = y.FirstOrDefault()?.BindSoStatus == "Y" ? true : false; return x; }).ToList();
                    #endregion
                }

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //信件通知
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
                    dynamicParameters.Add("SettingSchema", "DailyDeliveryWipUnLinkDetailMailAdvice");
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
                                                         <td align='center' bgcolor='#CCFFFF' style='font-family: 微軟正黑體, sans-serif;' width='100'>出貨日期</td>
                                                         <td align='center' bgcolor='#CCFFFF' style='font-family: 微軟正黑體, sans-serif;' width='175'>出貨客戶</td>
                                                         <td align='center' bgcolor='#CCFFFF' style='font-family: 微軟正黑體, sans-serif;' width='75'>出貨數量</td>
                                                         <td align='center' bgcolor='#CCFFFF' style='font-family: 微軟正黑體, sans-serif;' width='200'>訂單</td>
                                                         <td align='center' bgcolor='#CCFFFF' style='font-family: 微軟正黑體, sans-serif;' width='75'>訂單數量</td>
                                                         <td align='center' bgcolor='#CCFFFF' style='font-family: 微軟正黑體, sans-serif;' width='250'>品號</td>
                                                         <td align='center' bgcolor='#CCFFFF' style='font-family: 微軟正黑體, sans-serif;' width='300'>品名</td>
                                                         <td align='center' bgcolor='#CCFFFF' style='font-family: 微軟正黑體, sans-serif;' width='200'>規格</td>
                                                         <td align='center' bgcolor='#CCFFFF' style='font-family: 微軟正黑體, sans-serif;' width='250'>生管</td>
                                                       </tr>";

                        var datas = dailyDeliveryWipUnLinkDetails.Where(x => x.WipLink == false).Select(x => x).ToList();
                        foreach (var data in datas)
                        {
                            replaceHtml += @"          <tr>
                                                         <td align='center' style='font-family: 微軟正黑體, sans-serif;' width='100'>" + data.DoDate + @"</td>
                                                         <td align='center' style='font-family: 微軟正黑體, sans-serif;' width='175'>" + data.DcShortName + @"</td>
                                                         <td align='center' style='font-family: 微軟正黑體, sans-serif;' width='75'>" + data.PickQty + @"</td>
                                                         <td align='center' style='font-family: 微軟正黑體, sans-serif;' width='200'>" + data.SoErpFullNo + @"</td>
                                                         <td align='center' style='font-family: 微軟正黑體, sans-serif;' width='75'>" + data.SoQty + @"</td>
                                                         <td align='center' style='font-family: 微軟正黑體, sans-serif;' width='250'>" + data.MtlItemNo + @"</td>
                                                         <td align='center' style='font-family: 微軟正黑體, sans-serif;' width='300'>" + data.MtlItemName + @"</td>
                                                         <td align='center' style='font-family: 微軟正黑體, sans-serif;' width='200'>" + data.MtlItemSpec + @"</td>
                                                         <td align='center' style='font-family: 微軟正黑體, sans-serif;' width='250'>" + data.DepartmentName + data.PcUserNo + data.PcUserName + @"</td>
                                                       </tr>";
                        }

                        replaceHtml += @"            </table>
                                                   </td>
                                                 </tr>
                                               </table>";

                        mailSubject = mailSubject.Replace("[Date]", DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd"));
                        mailContent = mailContent.Replace("[Date]", DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd"));
                        mailContent = mailContent.Replace("[CustomContent]", replaceHtml);
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

                        if (datas.Count > 0) BaseHelper.MailSend(mailConfig);
                        #endregion
                    }
                    #endregion
                }

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    msg = "信件發送成功!"
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

        #region //DailyDeliveryMailAdvice -- 每日出貨明細 -- Ben Ma 2023.07.05
        public string DailyDeliveryMailAdvice(string CompanyNo)
        {
            try
            {
                int CompanyId = -1;

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //判斷公司資料是否正確
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT TOP 1 CompanyId
                            FROM BAS.Company
                            WHERE CompanyNo = @CompanyNo";
                    dynamicParameters.Add("CompanyNo", CompanyNo);

                    var resultCompany = sqlConnection.Query(sql, dynamicParameters);
                    if (resultCompany.Count() <= 0) throw new SystemException("【公司】資料錯誤!");

                    foreach (var item in resultCompany)
                    {
                        CompanyId = Convert.ToInt32(item.CompanyId);
                    }
                    #endregion

                    #region //出貨明細
                    sql = @"SELECT FORMAT(b.DoDate, 'yyyy-MM-dd') DoDate, c.DcShortName, a.DoQty
                            , e.SoErpPrefix + '-' + e.SoErpNo + '-' + d.SoSequence SoErpFullNo
                            , d.SoQty
                            , f.MtlItemNo, ISNULL(d.SoMtlItemName, f.MtlItemName) MtlItemName, ISNULL(d.SoMtlItemSpec, f.MtlItemSpec) MtlItemSpec
                            , ISNULL(g.PickQty, 0) PickQty
                            , CASE 
                                WHEN ISNULL(g.PickQty, 0) <= 0 THEN '未出貨'
                                WHEN b.[Status] != 'S' THEN '未確認'
                                ELSE '' 
                            END Anomaly
                            , CASE
                                WHEN ISNULL(h.TotalNoBarcodeQty, 0) > 0 THEN '未刷條碼'
                                ELSE ''
                            END BarcodeAnomaly
                            , CASE
                                WHEN ISNULL(h.TotalSubstituteQty, 0) > 0 THEN '出替代品：' + CAST(TotalSubstituteQty AS nvarchar)
                                ELSE ''
                            END SubstituteAnomaly
                            FROM SCM.DoDetail a
                            INNER JOIN SCM.DeliveryOrder b ON a.DoId = b.DoId
                            INNER JOIN SCM.DeliveryCustomer c ON b.DcId = c.DcId
                            INNER JOIN SCM.SoDetail d ON a.SoDetailId = d.SoDetailId
                            INNER JOIN SCM.SaleOrder e ON d.SoId = e.SoId
                            INNER JOIN PDM.MtlItem f ON d.MtlItemId = f.MtlItemId
                            OUTER APPLY (
                                SELECT SUM(x.ItemQty) PickQty
                                FROM SCM.PickingItem x
                                WHERE x.SoDetailId = a.SoDetailId
                                AND x.DoId = a.DoId
                            ) g
                            OUTER APPLY (
                                SELECT SUM(ha.DoQty) TotalDoQty, SUM(hb.NoBarcodeQty) TotalNoBarcodeQty, SUM(hc.BarcodeQty) TotalBarcodeQty, SUM(hd.SubstituteQty) TotalSubstituteQty
                                FROM SCM.DoDetail ha
                                OUTER APPLY (
                                    SELECT SUM(bba.ItemQty) NoBarcodeQty
                                    FROM SCM.PickingItem bba
                                    WHERE bba.SoDetailId = ha.SoDetailId
                                    AND bba.DoId = ha.DoId
                                    AND bba.ItemStatus = 'N'
                                    AND bba.BarcodeId IS NULL
                                ) hb
                                OUTER APPLY (
                                    SELECT SUM(bca.ItemQty) BarcodeQty
                                    FROM SCM.PickingItem bca
                                    WHERE bca.SoDetailId = ha.SoDetailId
                                    AND bca.DoId = ha.DoId
                                    AND bca.BarcodeId IS NOT NULL
                                ) hc
                                OUTER APPLY (
                                    SELECT SUM(bba.ItemQty) SubstituteQty
                                    FROM SCM.PickingItem bba
                                    WHERE bba.SoDetailId = ha.SoDetailId
                                    AND bba.DoId = ha.DoId
                                    AND bba.ItemStatus = 'S'
                                ) hd
                                WHERE ha.DoDetailId = a.DoDetailId
                            ) h
                            WHERE b.CompanyId = @CompanyId
                            AND b.DoDate >= @StartDate
                            AND b.DoDate <= @EndDate";
                    dynamicParameters.Add("CompanyId", CompanyId);
                    dynamicParameters.Add("StartDate", DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 00:00:00"));
                    dynamicParameters.Add("EndDate", DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 23:59:59"));

                    var result = sqlConnection.Query(sql, dynamicParameters);
                    #endregion

                    #region //信件通知
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
                    dynamicParameters.Add("SettingSchema", "DailyDeliveryMailAdvice");
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
                                                         <td align='center' bgcolor='#CCFFFF' style='font-family: 微軟正黑體, sans-serif;' width='75'></td>
                                                         <td align='center' bgcolor='#CCFFFF' style='font-family: 微軟正黑體, sans-serif;' width='75'></td>
                                                         <td align='center' bgcolor='#CCFFFF' style='font-family: 微軟正黑體, sans-serif;' width='75'></td>
                                                         <td align='center' bgcolor='#CCFFFF' style='font-family: 微軟正黑體, sans-serif;' width='100'>出貨日期</td>
                                                         <td align='center' bgcolor='#CCFFFF' style='font-family: 微軟正黑體, sans-serif;' width='175'>出貨客戶</td>
                                                         <td align='center' bgcolor='#CCFFFF' style='font-family: 微軟正黑體, sans-serif;' width='200'>訂單</td>
                                                         <td align='center' bgcolor='#CCFFFF' style='font-family: 微軟正黑體, sans-serif;' width='75'>訂單數量</td>
                                                         <td align='center' bgcolor='#CCFFFF' style='font-family: 微軟正黑體, sans-serif;' width='75'>定版數量</td>
                                                         <td align='center' bgcolor='#CCFFFF' style='font-family: 微軟正黑體, sans-serif;' width='75'>出貨數量</td>
                                                         <td align='center' bgcolor='#CCFFFF' style='font-family: 微軟正黑體, sans-serif;' width='300'>品名</td>
                                                         <td align='center' bgcolor='#CCFFFF' style='font-family: 微軟正黑體, sans-serif;' width='250'>品號</td>
                                                         <td align='center' bgcolor='#CCFFFF' style='font-family: 微軟正黑體, sans-serif;' width='200'>規格</td>
                                                       </tr>";

                        foreach (var data in result)
                        {
                            replaceHtml += @"          <tr>
                                                         <td align='center' style='font-family: 微軟正黑體, sans-serif;' width='75'>" + data.Anomaly + @"</td>
                                                         <td align='center' style='font-family: 微軟正黑體, sans-serif;' width='75'>" + data.BarcodeAnomaly + @"</td>
                                                         <td align='center' style='font-family: 微軟正黑體, sans-serif;' width='75'>" + data.SubstituteAnomaly + @"</td>
                                                         <td align='center' style='font-family: 微軟正黑體, sans-serif;' width='100'>" + data.DoDate + @"</td>
                                                         <td align='center' style='font-family: 微軟正黑體, sans-serif;' width='175'>" + data.DcShortName + @"</td>
                                                         <td align='center' style='font-family: 微軟正黑體, sans-serif;' width='200'>" + data.SoErpFullNo + @"</td>
                                                         <td align='center' style='font-family: 微軟正黑體, sans-serif;' width='75'>" + data.SoQty + @"</td>
                                                         <td align='center' style='font-family: 微軟正黑體, sans-serif;' width='75'>" + data.DoQty + @"</td>
                                                         <td align='center' style='font-family: 微軟正黑體, sans-serif;' width='75'>" + data.PickQty + @"</td>
                                                         <td align='center' style='font-family: 微軟正黑體, sans-serif;' width='300'>" + data.MtlItemName + @"</td>
                                                         <td align='center' style='font-family: 微軟正黑體, sans-serif;' width='250'>" + data.MtlItemNo + @"</td>
                                                         <td align='center' style='font-family: 微軟正黑體, sans-serif;' width='200'>" + data.MtlItemSpec + @"</td>
                                                       </tr>";
                        }

                        replaceHtml += @"            </table>
                                                   </td>
                                                 </tr>
                                               </table>";

                        mailSubject = mailSubject.Replace("[Date]", DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd"));
                        mailContent = mailContent.Replace("[Date]", DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd"));
                        mailContent = mailContent.Replace("[CustomContent]", replaceHtml);
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

                        if (result.Count() > 0) BaseHelper.MailSend(mailConfig);
                        #endregion
                    }
                    #endregion
                }

                #region //Response
                jsonResponse = JObject.FromObject(new
                {
                    status = "success",
                    msg = "信件發送成功!"
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
        #endregion
    }
}
