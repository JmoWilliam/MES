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

namespace SCMDA
{
    public class PickingSystemDA
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

        public PickingSystemDA()
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
        #region //GetCartonDetail -- 取得所有物流箱資料 -- Zoey 2022.11.16
        public string GetCartonDetail(int DoId, int DoDetailId, int PickingCartonId, string CartonBarcode)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    if (PickingCartonId == -999)
                    {
                        if (DoDetailId <= 0 && DoId <= 0) throw new SystemException("參數錯誤!");

                        #region //判斷出貨單資料是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.DeliveryOrder a
                                INNER JOIN SCM.DoDetail b ON a.DoId = b.DoId
                                WHERE a.CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "DoId", @" AND a.DoId = @DoId", DoId);
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "DoDetailId", @" AND b.DoDetailId = @DoDetailId", DoDetailId);

                        var doResult = sqlConnection.Query(sql, dynamicParameters);
                        if (doResult.Count() <= 0) throw new SystemException("出貨單資料錯誤!");
                        #endregion
                    }

                    sql = @"SELECT a.PickingCartonId, a.PackingId, a.CartonName, a.CartonBarcode
                            , a.CartonRemark, a.TotalWeight, a.UomId, a.PrintStatus, a.PrintCount
                            , ISNULL(b.DoId, -1) DoId
                            , ISNULL(c.Status, '') Status
                            FROM SCM.PickingCarton a
                            OUTER APPLY (
                                SELECT TOP 1 ba.DoId
                                FROM SCM.PickingItem ba
                                WHERE a.PickingCartonId = ba.PickingCartonId
                            ) b
                            LEFT JOIN SCM.DeliveryOrder c ON b.DoId = c.DoId
                            INNER JOIN SCM.Packing d ON a.PackingId = d.PackingId
                            WHERE d.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    if (CartonBarcode.Length <= 0)
                    {
                        sql += @" AND a.PickingCartonId = @PickingCartonId";
                        dynamicParameters.Add("PickingCartonId", PickingCartonId);
                    }
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "DoId", @" AND b.DoId = @DoId", DoId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "CartonBarcode", @" AND a.CartonBarcode = @CartonBarcode", CartonBarcode);

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

        #region //GetPickingCarton -- 取得物流箱資料(不含有出貨單) -- Zoey 2022.10.12
        public string GetPickingCarton(int PickingCartonId, int PackingId, string CartonName, string CartonBarcode, string CartonStatus
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.PickingCartonId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.PackingId, a.CartonName, a.CartonBarcode, a.CartonRemark, a.TotalWeight, a.PrintStatus, a.PrintCount
                          , b.PackingName + ' (' + b.VolumeSpec + ')' PackingNameWithVolume
                          , c.ItemStatus";
                    sqlQuery.mainTables =
                        @"FROM SCM.PickingCarton a
                          INNER JOIN SCM.Packing b ON b.PackingId = a.PackingId
                          OUTER APPLY(
                            SELECT ISNULL((
                                SELECT TOP 1 1
                                FROM SCM.PickingItem aa
                                WHERE aa.PickingCartonId = a.PickingCartonId
                            ), 0) ItemStatus
                          ) c";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND (
                                                EXISTS (
                                                    SELECT TOP 1 1
                                                    FROM SCM.PickingItem aa
                                                    WHERE aa.PickingCartonId = a.PickingCartonId
                                                    AND aa.DoId IS NULL
                                                ) 
                                                OR c.ItemStatus = 0
                                            )
                                            AND b.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    if (CartonStatus == "E") queryCondition += @" AND c.ItemStatus = 0";
                    if (CartonStatus == "C") queryCondition += @" AND c.ItemStatus = 1";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PickingCartonId", @" AND a.PickingCartonId = @PickingCartonId", PickingCartonId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PackingId", @" AND a.PackingId = @PackingId", PackingId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CartonName", @" AND a.CartonName LIKE '%' + @CartonName + '%'", CartonName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CartonBarcode", @" AND a.CartonBarcode LIKE '%' + @CartonBarcode + '%'", CartonBarcode);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.PickingCartonId";
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

        #region //GetChangeCarton -- 取得物流箱資料(含有出貨單) -- Zoey 2022.11.10
        public string GetChangeCarton(string CartonBarcode)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT TOP 1 1
                            FROM SCM.PickingCarton a
                            LEFT JOIN SCM.PickingItem b ON a.PickingCartonId = b.PickingCartonId
                            INNER JOIN SCM.Packing c ON a.PackingId = c.PackingId
                            WHERE b.DoId IS NULL
                            AND a.CartonBarcode = @CartonBarcode
                            AND c.CompanyId = @CompanyId";
                    dynamicParameters.Add("CartonBarcode", CartonBarcode);
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    var result = sqlConnection.Query(sql, dynamicParameters);
                    if (result.Count() <= 0) throw new SystemException("該物流箱已綁定出貨單，無法進行換箱作業!");

                    sql = @"SELECT a.PickingCartonId, a.PackingId, a.CartonName, a.CartonBarcode, a.CartonRemark, a.TotalWeight, a.PrintStatus, a.PrintCount
                            , b.PackingName + ' (' + b.VolumeSpec + ')' PackingNameWithVolume
                            FROM SCM.PickingCarton a
                            INNER JOIN SCM.Packing b ON b.PackingId = a.PackingId
                            WHERE b.CompanyId = @CompanyId
                            AND a.CartonBarcode = @CartonBarcode";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    dynamicParameters.Add("CartonBarcode", CartonBarcode);

                    var result2 = sqlConnection.Query(sql, dynamicParameters);
                    if (result2.Count() <= 0) throw new SystemException("查無此物流箱條碼!");

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        data = result2
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

        #region //GetPickingItem -- 取得揀貨物件資料 -- Zoey 2022.10.24
        public string GetPickingItem(int DoId, int PickingItemId, int PickingCartonId, int SoDetailId, string PickingItemIds)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT DISTINCT a.PickingCartonId, a.SoDetailId
                            , b.CartonName, b.CartonBarcode
                            , d.MtlItemNo, d.MtlItemName
                            , ISNULL(f.PickRegularQty, 0) PickRegularQty, ISNULL(g.PickFreebieQty, 0) PickFreebieQty, ISNULL(h.PickSpareQty, 0) PickSpareQty
                            , ISNULL((
                                SELECT aa.PickingItemId, aa.BarcodeId, aa.SoDetailId, aa.ItemStatus, aa.ItemType, aa.ItemQty
                                , ISNULL(ab.BarcodeNo, '') ItemBarcode, ISNULL(ac.ItemValue, '') LetteringNo
                                , ISNULL(aa.LotNumber, '') LotNumber
                                FROM SCM.PickingItem aa
                                LEFT JOIN MES.Barcode ab ON ab.BarcodeId = aa.BarcodeId
                                OUTER APPLY (
                                    SELECT aca.ItemValue
                                    FROM MES.BarcodeAttribute aca
                                    WHERE aca.ItemNo = 'Lettering'
                                    AND aca.BarcodeId = aa.BarcodeId
                                ) ac
                                WHERE aa.DoId = a.DoId
                                AND ISNULL(aa.PickingCartonId, -999) = ISNULL(a.PickingCartonId, -999)
                                AND aa.SoDetailId = a.SoDetailId";
                    if (PickingItemIds.Length > 0) BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PickingItemIds", @" AND aa.PickingItemId IN @PickingItemIds", PickingItemIds.Split(','));

                    sql += @"   ORDER BY aa.ItemType, aa.ItemStatus
                            FOR JSON PATH, ROOT('data')
                            ) ,'') ItemDetail
                            FROM SCM.PickingItem a
                            LEFT JOIN SCM.PickingCarton b ON b.PickingCartonId = a.PickingCartonId
                            LEFT JOIN SCM.SoDetail c ON c.SoDetailId = a.SoDetailId
                            LEFT JOIN PDM.MtlItem d ON d.MtlItemId = c.MtlItemId
                            LEFT JOIN SCM.DeliveryOrder e ON e.DoId = a.DoId
                            OUTER APPLY(
                                SELECT SUM(x.ItemQty) PickRegularQty
                                FROM SCM.PickingItem x
                                WHERE ISNULL(x.PickingCartonId, -999) = ISNULL(a.PickingCartonId, -999)
                                AND x.ItemType = 1
                                AND x.SoDetailId = a.SoDetailId
                            ) f
                            OUTER APPLY(
                                SELECT SUM(y.ItemQty) PickFreebieQty
                                FROM SCM.PickingItem y
                                WHERE ISNULL(y.PickingCartonId, -999) = ISNULL(a.PickingCartonId, -999)
                                AND y.ItemType = 2
                                AND y.SoDetailId = a.SoDetailId
                            ) g
                            OUTER APPLY(
                                SELECT SUM(z.ItemQty) PickSpareQty
                                FROM SCM.PickingItem z
                                WHERE ISNULL(z.PickingCartonId, -999) = ISNULL(a.PickingCartonId, -999)
                                AND z.ItemType = 3
                                AND z.SoDetailId = a.SoDetailId
                            ) h
                            WHERE 1=1";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "DoId", @" AND a.DoId = @DoId", DoId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PickingItemId", @" AND a.PickingItemId = @PickingItemId", PickingItemId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PickingCartonId", @" AND a.PickingCartonId = @PickingCartonId", PickingCartonId);
                    if (PickingCartonId == -999) sql += @" AND a.PickingCartonId IS NULL";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "SoDetailId", @" AND a.SoDetailId = @SoDetailId", SoDetailId);
                    if (PickingItemIds.Length > 0) BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "PickingItemIds", @" AND a.PickingItemId IN @PickingItemIds", PickingItemIds.Split(','));

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

        #region //GetDeliveryOrder -- 取得出貨單資料 -- Zoey 2022.10.24
        public string GetDeliveryOrder(int DoId, string TypeTwo, string DoErpFullNo, string CartonBarcode, string Status, string StartDate, string EndDate
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sqlQuery.mainKey = "a.DoId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", FORMAT(a.DoDate, 'yyyy-MM-dd HH:mm') DoDate
                          , a.DoErpPrefix + '-' + a.DoErpNo DoErpFullNo
                          , a.Status
                          , b.CustomerNo + '-' + b.CustomerName CustomerWithNo
                          , ISNULL(c.DcName, '') DcName";
                    sqlQuery.mainTables =
                        @"FROM SCM.DeliveryOrder a 
                          LEFT JOIN SCM.Customer b ON b.CustomerId = a.CustomerId
                          LEFT JOIN SCM.DeliveryCustomer c ON c.DcId = a.DcId
                          OUTER APPLY(
                              SELECT ISNULL((
                                  SELECT DISTINCT y.CartonBarcode
                                  FROM SCM.PickingItem x
                                  LEFT JOIN SCM.PickingCarton y ON y.PickingCartonId = x.PickingCartonId
                                  WHERE x.DoId = a.DoId
                              ), '') CartonBarcode
                          ) x
                          OUTER APPLY(
                              SELECT ISNULL((
                                  SELECT TOP 1 z.TypeTwo
                                  FROM SCM.DoDetail x
                                  LEFT JOIN SCM.SoDetail y ON y.SoDetailId = x.SoDetailId
                                  LEFT JOIN PDM.MtlItem z ON y.MtlItemId = z.MtlItemId
                                  WHERE x.DoId = a.DoId
                              ), '') TypeTwo
                          ) y";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DoId", @" AND a.DoId = @DoId", DoId);
                    if (TypeTwo.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "TypeTwo", @" AND y.TypeTwo = @TypeTwo", TypeTwo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DoErpFullNo", @" AND a.DoErpPrefix + '-' + a.DoErpNo LIKE '%' + @DoErpFullNo + '%'", DoErpFullNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "CartonBarcode", @" AND x.CartonBarcode LIKE '%' + @CartonBarcode + '%'", CartonBarcode);
                    if (Status != "A") BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Status", @" AND a.Status = @Status", Status);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "StartDate", @" AND a.DoDate >= @StartDate ", StartDate.Length > 0 ? Convert.ToDateTime(StartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "EndDate", @" AND a.DoDate <= @EndDate ", EndDate.Length > 0 ? Convert.ToDateTime(EndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.DoId DESC";
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

        #region //GetDoDetail -- 取得出貨單物件資料 -- Zoey 2022.11.14  
        public string GetDoDetail(int DoDetailId, int DoId)
        {
            try
            {
                if (DoDetailId <= 0 && DoId <= 0) throw new SystemException("參數錯誤!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.SoDetailId
                            , b.DoDetailId, b.DoQty RegularQty, ISNULL(b.FreebieQty, 0) FreebieQty, ISNULL(b.SpareQty, 0) SpareQty
                            , c.DoId , c.Status, c.DoErpPrefix + '-' + c.DoErpNo DoErpFullNo
                            , d.MtlItemId, d.MtlItemNo, d.MtlItemName, d.MtlItemSpec
                            , ISNULL(e.TypeName, '') DeliveryProcess
                            , ISNULL(f.PickRegularQty, 0) PickRegularQty, ISNULL(g.PickFreebieQty, 0) PickFreebieQty, ISNULL(h.PickSpareQty, 0) PickSpareQty
                            , ISNULL((
                                SELECT aa.PickingItemId, aa.BarcodeId, aa.SoDetailId, aa.ItemStatus, aa.ItemType, aa.ItemQty
                                , ISNULL(ab.BarcodeNo, '') ItemBarcode, ISNULL(ac.ItemValue, '') LetteringNo
                                , ISNULL(aa.LotNumber, '') LotNumber
								, ISNULL(ad.TrayNo, '') TrayNo
                                FROM SCM.PickingItem aa
                                LEFT JOIN MES.Barcode ab ON ab.BarcodeId = aa.BarcodeId
                                OUTER APPLY (
                                    SELECT aca.ItemValue
                                    FROM MES.BarcodeAttribute aca
                                    WHERE aca.ItemNo = 'Lettering'
                                    AND aca.BarcodeId = aa.BarcodeId
                                ) ac
								LEFT JOIN MES.Tray ad ON ad.BarcodeNo = ab.BarcodeNo
                                WHERE 1=1
                                AND aa.SoDetailId = a.SoDetailId
                                AND aa.DoId = b.DoId
                                ORDER BY aa.ItemType, aa.ItemStatus
                                FOR JSON PATH, ROOT('data')
                            ), '') ItemDetail
                            , ISNULL(i.CustomerPurchaseOrder, '') CustomerPurchaseOrder
                            FROM SCM.SoDetail a
                            INNER JOIN SCM.DoDetail b ON b.SoDetailId = a.SoDetailId
                            INNER JOIN SCM.DeliveryOrder c ON c.DoId = b.DoId
                            LEFT JOIN PDM.MtlItem d ON d.MtlItemId = a.MtlItemId
                            LEFT JOIN BAS.[Type] e ON e.TypeNo = b.DeliveryProcess AND e.TypeSchema = 'DeliveryProcess'
                            OUTER APPLY(
                                SELECT ISNULL(SUM(x.ItemQty), 0) PickRegularQty
                                FROM SCM.PickingItem x
                                WHERE x.SoDetailId = a.SoDetailId
                                AND x.ItemType = 1
                                AND x.DoId = b.DoId
                            ) f
                            OUTER APPLY(
                                SELECT ISNULL(SUM(y.ItemQty), 0) PickFreebieQty
                                FROM SCM.PickingItem y
                                WHERE y.SoDetailId = a.SoDetailId
                                AND y.ItemType = 2
                                AND y.DoId = b.DoId
                            ) g
                            OUTER APPLY(
                                SELECT ISNULL(SUM(z.ItemQty), 0) PickSpareQty
                                FROM SCM.PickingItem z
                                WHERE z.SoDetailId = a.SoDetailId
                                AND z.ItemType = 3
                                AND z.DoId = b.DoId
                            ) h
                            INNER JOIN SCM.SaleOrder i ON a.SoId = i.SoId
                            WHERE c.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "DoDetailId", @" AND b.DoDetailId = @DoDetailId", DoDetailId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "DoId", @" AND b.DoId = @DoId", DoId);

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

        #region //GetCartonItem -- 取得物流箱物件資料 -- Zoey 2022.11.14
        public string GetCartonItem(int DoId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT DISTINCT a.PickingCartonId, a.CartonName, a.CartonBarcode, a.TotalWeight, a.UomId
                            , ISNULL(CONVERT(varchar(10), a.TotalWeight) + ' ' + b.UomNo, 0) TotalWeightWithUom
                            , c.DoId, c.Status
                            , ISNULL((
                                SELECT x.ItemType, SUM(x.ItemQty) ItemQty, y.SoDetailId, z.MtlItemName
                                FROM SCM.PickingItem x
                                INNER JOIN SCM.SoDetail y ON y.SoDetailId = x.SoDetailId
                                LEFT JOIN PDM.MtlItem z ON z.MtlItemId = y.MtlItemId
                                WHERE x.PickingCartonId = a.PickingCartonId
                                GROUP BY x.ItemType, y.SoDetailId, z.MtlItemName
                                FOR JSON PATH, ROOT('data')
                                ), '') ItemDetail
                            FROM SCM.PickingCarton a
                            LEFT JOIN PDM.UnitOfMeasure b ON b.UomId = a.UomId
                            OUTER APPLY(
                                SELECT x.DoId, y.Status, x.ItemType
                                FROM SCM.PickingItem x
                                LEFT JOIN SCM.DeliveryOrder y ON y.DoId = x.DoId
                                WHERE x.PickingCartonId = a.PickingCartonId
                            ) c
                            INNER JOIN SCM.Packing d ON a.PackingId = d.PackingId
                            WHERE d.CompanyId = @CompanyId";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "DoId", @" AND c.DoId = @DoId", DoId);
                    sql += @" UNION ALL
                            SELECT DISTINCT -999 PickingCartonId, '無物流箱' CartonName, '' CartonBarcode, 0 TotalWeight, -1 UomId
                            , '' TotalWeightWithUom
                            , a.DoId, a.[Status]
                            , ISNULL((
                                SELECT x.ItemType, SUM(x.ItemQty) ItemQty, y.SoDetailId, z.MtlItemName
                                FROM SCM.PickingItem x
                                INNER JOIN SCM.SoDetail y ON y.SoDetailId = x.SoDetailId
                                LEFT JOIN PDM.MtlItem z ON z.MtlItemId = y.MtlItemId
                                INNER JOIN SCM.SaleOrder w ON y.SoId = w.SoId
                                WHERE x.DoId = a.DoId
                                AND ISNULL(x.PickingCartonId, -999) = -999
                                AND w.CompanyId = @CompanyId
                                GROUP BY x.ItemType, y.SoDetailId, z.MtlItemName
                                FOR JSON PATH, ROOT('data')
                            ), '') ItemDetail
                            FROM (
                                SELECT x.DoId, y.Status, x.ItemType
                                FROM SCM.PickingItem x
                                LEFT JOIN SCM.DeliveryOrder y ON y.DoId = x.DoId
                                WHERE ISNULL(x.PickingCartonId, -999) = -999
                                AND y.CompanyId = @CompanyId
                            ) a
                            WHERE 1=1";
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "DoId", @" AND a.DoId = @DoId", DoId);
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

        #region //GetPickingTask -- 取得揀貨中物件資料 -- Ben Ma 2022.12.06  
        public string GetPickingTask(int DoDetailId, int PickingCartonId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.DoDetailId, a.DoId, a.SoDetailId, a.DoQty RegularQty, ISNULL(a.FreebieQty, 0) FreebieQty, ISNULL(a.SpareQty, 0) SpareQty
                            , b.SoMtlItemName MtlItemName, b.SoMtlItemSpec MtlItemSpec
                            , c.MtlItemNo, c.LotManagement
                            , ISNULL(d.PickRegularQty, 0) PickRegularQty, ISNULL(f.PickFreebieQty, 0) PickFreebieQty, ISNULL(g.PickSpareQty, 0) PickSpareQty
                            , ISNULL((
                                SELECT aa.PickingItemId, aa.BarcodeId, aa.SoDetailId, aa.ItemStatus, aa.ItemType, aa.ItemQty
                                , ISNULL(ab.BarcodeNo, '') ItemBarcode, ISNULL(ac.ItemValue, '') LetteringNo
                                , ISNULL(aa.LotNumber, '') LotNumber
								, ISNULL(ad.TrayNo, '') TrayNo
                                FROM SCM.PickingItem aa
                                LEFT JOIN MES.Barcode ab ON ab.BarcodeId = aa.BarcodeId
                                OUTER APPLY (
                                    SELECT aca.ItemValue
                                    FROM MES.BarcodeAttribute aca
                                    WHERE aca.ItemNo = 'Lettering'
                                    AND aca.BarcodeId = aa.BarcodeId
                                ) ac
                                LEFT JOIN MES.Tray ad ON ab.BarcodeNo = ad.BarcodeNo
                                WHERE 1=1
                                AND ISNULL(aa.PickingCartonId, -999) = @PickingCartonId
                                AND aa.SoDetailId = a.SoDetailId
                                AND aa.DoId = a.DoId
                                FOR JSON PATH, ROOT('data')
                            ), '') ItemDetail
                            FROM SCM.DoDetail a
                            INNER JOIN SCM.SoDetail b ON a.SoDetailId = b.SoDetailId
                            INNER JOIN PDM.MtlItem c ON b.MtlItemId = c.MtlItemId
                            OUTER APPLY(
                                SELECT ISNULL(SUM(x.ItemQty), 0) PickRegularQty
                                FROM SCM.PickingItem x
                                WHERE x.SoDetailId = a.SoDetailId
                                AND x.ItemType = 1
                                AND x.DoId = a.DoId
                            ) d
                            OUTER APPLY(
                                SELECT ISNULL(SUM(y.ItemQty), 0) PickFreebieQty
                                FROM SCM.PickingItem y
                                WHERE y.SoDetailId = a.SoDetailId
                                AND y.ItemType = 2
                                AND y.DoId = a.DoId
                            ) f
                            OUTER APPLY(
                                SELECT ISNULL(SUM(z.ItemQty), 0) PickSpareQty
                                FROM SCM.PickingItem z
                                WHERE z.SoDetailId = a.SoDetailId
                                AND z.ItemType = 3
                                AND z.DoId = a.DoId
                            ) g
                            INNER JOIN SCM.DeliveryOrder e ON a.DoId = e.DoId
                            WHERE e.CompanyId = @CompanyId
                            AND a.DoDetailId = @DoDetailId
                            UNION ALL
                            SELECT a.DoDetailId, a.DoId, a.SoDetailId, a.DoQty, ISNULL(a.FreebieQty, 0) FreebieQty, ISNULL(a.SpareQty, 0) SpareQty
                            , b.SoMtlItemName MtlItemName, b.SoMtlItemSpec MtlItemSpec
                            , c.MtlItemNo, c.LotManagement
                            , ISNULL(d.PickRegularQty, 0) PickRegularQty, ISNULL(f.PickFreebieQty, 0) PickFreebieQty, ISNULL(g.PickSpareQty, 0) PickSpareQty
                            , ISNULL((
                                SELECT aa.PickingItemId, aa.BarcodeId, aa.SoDetailId, aa.ItemStatus, aa.ItemType, aa.ItemQty
                                , ISNULL(ab.BarcodeNo, '') ItemBarcode, ISNULL(ac.ItemValue, '') LetteringNo
                                , ISNULL(aa.LotNumber, '') LotNumber
                                FROM SCM.PickingItem aa
                                LEFT JOIN MES.Barcode ab ON ab.BarcodeId = aa.BarcodeId
                                OUTER APPLY (
                                    SELECT aca.ItemValue
                                    FROM MES.BarcodeAttribute aca
                                    WHERE aca.ItemNo = 'Lettering'
                                    AND aca.BarcodeId = aa.BarcodeId
                                ) ac
                                WHERE 1=1
                                AND ISNULL(aa.PickingCartonId, -999) = @PickingCartonId
                                AND aa.SoDetailId = a.SoDetailId
                                AND aa.DoId = a.DoId
                                FOR JSON PATH, ROOT('data')
                            ), '') ItemDetail
                            FROM SCM.DoDetail a
                            INNER JOIN SCM.SoDetail b ON a.SoDetailId = b.SoDetailId
                            INNER JOIN PDM.MtlItem c ON b.MtlItemId = c.MtlItemId
                            OUTER APPLY(
                                SELECT ISNULL(SUM(x.ItemQty), 0) PickRegularQty
                                FROM SCM.PickingItem x
                                WHERE x.SoDetailId = a.SoDetailId
                                AND x.ItemType = 1
                                AND x.DoId = a.DoId
                            ) d
                            OUTER APPLY(
                                SELECT ISNULL(SUM(y.ItemQty), 0) PickFreebieQty
                                FROM SCM.PickingItem y
                                WHERE y.SoDetailId = a.SoDetailId
                                AND y.ItemType = 2
                                AND y.DoId = a.DoId
                            ) f
                            OUTER APPLY(
                                SELECT ISNULL(SUM(z.ItemQty), 0) PickSpareQty
                                FROM SCM.PickingItem z
                                WHERE z.SoDetailId = a.SoDetailId
                                AND z.ItemType = 3
                                AND z.DoId = a.DoId
                            ) g
                            INNER JOIN SCM.DeliveryOrder e ON a.DoId = e.DoId
                            WHERE a.DoDetailId IN (
                                SELECT DISTINCT y.DoDetailId
                                FROM SCM.PickingItem z
                                INNER JOIN SCM.DoDetail y ON z.DoId = y.DoId AND z.SoDetailId = y.SoDetailId
                                WHERE z.PickingCartonId = @PickingCartonId
                                AND y.DoDetailId != @DoDetailId
                            )
                            AND e.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    dynamicParameters.Add("DoDetailId", DoDetailId);
                    dynamicParameters.Add("PickingCartonId", PickingCartonId);

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

        #region //GetPackageInBarcode -- 取得包裝內容物條碼 -- GPAI 241111
        public string GetPackageInBarcode(string BarcodeNo)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {



                    #region 確認條碼資料
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.BarcodeId, a.MoId, a.BarcodeQty
                                        FROM MES.Barcode a
                                        WHERE  a.BarcodeNo = @BarcodeNo";
                    dynamicParameters.Add("BarcodeNo", BarcodeNo);


                    var resultBarcode = sqlConnection.Query(sql, dynamicParameters);
                    if (resultBarcode.Count() <= 0) throw new SystemException("包裝條碼資料錯誤!");

                    int BarcodeId = -1;
                    foreach (var item in resultBarcode)
                    {
                        BarcodeId = Convert.ToInt32(item.BarcodeId);
                    }
                    #endregion

                    #region 取得內容物條碼資料
                    dynamicParameters = new DynamicParameters();

                    sql = @"DECLARE @InputBarcode INT = @BarcodeId;
                            DECLARE @CurrentBarcode INT;
                            DECLARE @PrevBarcodeId INT;
                            
                            -- 建立表變量來存儲結果
                            DECLARE @BarcodeResults TABLE (
                                BarcodeId INT,
                                BarcodeNo NVARCHAR(100),
                                PrevBarcodeId INT,
                                BarcodeQty INT,
                                isLeaf NVARCHAR(100),
								MtlItemNo NVARCHAR(100),
								MtlItemName NVARCHAR(100)

                            );
                            
                            INSERT INTO @BarcodeResults (BarcodeId, BarcodeNo, PrevBarcodeId, BarcodeQty,isLeaf, MtlItemNo, MtlItemName)
                            SELECT 
                                a.BarcodeId,
                                a.BarcodeNo,
                                a.PrevBarcodeId,
                                a.BarcodeQty,
                                a.isLeaf,
									d.MtlItemNo,
									d.MtlItemName
                            FROM MES.Barcode a
							left JOIN MES.ManufactureOrder b ON a.MoId = b.MoId
							left JOIN MES.WipOrder c ON b.WoId = c.WoId
							left JOIN PDM.MtlItem d ON c.MtlItemId = d.MtlItemId

                            WHERE a.BarcodeId = @InputBarcode;
                            
                            -- PrevBarcodeId
                            SET @PrevBarcodeId = (SELECT PrevBarcodeId FROM MES.Barcode WHERE BarcodeId = @InputBarcode);
                            
                            
                            
                            SET @CurrentBarcode = @InputBarcode;
                            
                            -- 向下查找下層條碼
                            WHILE @CurrentBarcode IS NOT NULL
                            BEGIN
                                -- 所有下層條碼
                                INSERT INTO @BarcodeResults (BarcodeId, BarcodeNo, PrevBarcodeId, BarcodeQty,isLeaf, MtlItemNo, MtlItemName)
                                SELECT 
                                    a.BarcodeId,
                                    a.BarcodeNo,
                                    a.PrevBarcodeId,
                                    a.BarcodeQty,
                                    a.isLeaf,
									d.MtlItemNo,
									d.MtlItemName
                                FROM MES.Barcode a
							left JOIN MES.ManufactureOrder b ON a.MoId = b.MoId
							left JOIN MES.WipOrder c ON b.WoId = c.WoId
							left JOIN PDM.MtlItem d ON c.MtlItemId = d.MtlItemId
                                WHERE a.PrevBarcodeId = @CurrentBarcode;
                            
                                SET @CurrentBarcode = (SELECT TOP 1 BarcodeId FROM MES.Barcode WHERE PrevBarcodeId = @CurrentBarcode);
                            
                                IF @CurrentBarcode IS NULL
                                    BREAK;
                            END
                            
                            -- 查同ParentBarcode其他條碼
                            INSERT INTO @BarcodeResults (BarcodeId, BarcodeNo, PrevBarcodeId, BarcodeQty,isLeaf, MtlItemNo, MtlItemName)
                            SELECT 
                                BarcodeId,
                                BarcodeNo,
                                PrevBarcodeId,
                                BarcodeQty,
                                isLeaf,
									d.MtlItemNo,
									d.MtlItemName
                                FROM MES.Barcode a
							left JOIN MES.ManufactureOrder b ON a.MoId = b.MoId
							left JOIN MES.WipOrder c ON b.WoId = c.WoId
							left JOIN PDM.MtlItem d ON c.MtlItemId = d.MtlItemId
                            WHERE a.PrevBarcodeId = @PrevBarcodeId
                              AND a.BarcodeId <> @InputBarcode;  -- 排除查詢的條碼
                            
                            -- 顯示所有條碼資料
                            SELECT 
                                BarcodeId,
                                BarcodeNo,
                                PrevBarcodeId,
                                BarcodeQty,
                                isLeaf,
								MtlItemNo,
								MtlItemName
                            FROM @BarcodeResults
                            WHERE BarcodeId NOT IN (
							 SELECT PrevBarcodeId FROM @BarcodeResults WHERE PrevBarcodeId IS NOT NULL
							);";
                    dynamicParameters.Add("BarcodeId", BarcodeId);


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

        #region //GetDoDetailForPackage -- 取得出貨單物件資料(FOR PACKAGE) -- Gpai 241111
        public string GetDoDetailForPackage(int DoId)
        {
            try
            {
                if (DoId <= 0) throw new SystemException("參數錯誤!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"WITH DoDetailData AS (
                            SELECT
                                a.SoDetailId, 
                                '1' as PickType,
                                b.DoDetailId, 
                                b.DoQty RegularQty, 
                                ISNULL(b.FreebieQty, 0) FreebieQty, 
                                ISNULL(b.SpareQty, 0) SpareQty,
                                c.DoId, 
                                c.Status, 
                                c.DoErpPrefix + '-' + c.DoErpNo DoErpFullNo,
                                d.MtlItemId, 
                                d.MtlItemNo, 
                                d.MtlItemName, 
                                d.MtlItemSpec,
                                ISNULL(e.TypeName, '') DeliveryProcess,
                                ISNULL(f.PickRegularQty, 0) PickRegularQty, 
                                ISNULL(g.PickFreebieQty, 0) PickFreebieQty, 
                                ISNULL(h.PickSpareQty, 0) PickSpareQty,
	                        	 ISNULL((
                                                        SELECT aa.PickingItemId, aa.BarcodeId, aa.SoDetailId, aa.ItemStatus, aa.ItemType, aa.ItemQty
                                                        , ISNULL(ab.BarcodeNo, '') ItemBarcode, ISNULL(ac.ItemValue, '') LetteringNo
                                                        , ISNULL(aa.LotNumber, '') LotNumber
	                        							, ISNULL(ad.TrayNo, '') TrayNo
                                                        FROM SCM.PickingItem aa
                                                        LEFT JOIN MES.Barcode ab ON ab.BarcodeId = aa.BarcodeId
                                                        OUTER APPLY (
                                                            SELECT aca.ItemValue
                                                            FROM MES.BarcodeAttribute aca
                                                            WHERE aca.ItemNo = 'Lettering'
                                                            AND aca.BarcodeId = aa.BarcodeId
                                                        ) ac
	                        							LEFT JOIN MES.Tray ad ON ad.BarcodeNo = ab.BarcodeNo
                                                        WHERE 1=1
                                                        AND aa.SoDetailId = a.SoDetailId
                                                        AND aa.DoId = b.DoId
                                                        AND aa.ItemType = '1'
                                                        ORDER BY aa.ItemType, aa.ItemStatus
                                                        FOR JSON PATH, ROOT('data')
                                                    ), '') ItemDetail
                            FROM SCM.SoDetail a
                            INNER JOIN SCM.DoDetail b ON b.SoDetailId = a.SoDetailId
                            INNER JOIN SCM.DeliveryOrder c ON c.DoId = b.DoId
                            LEFT JOIN PDM.MtlItem d ON d.MtlItemId = a.MtlItemId
                            LEFT JOIN BAS.[Type] e ON e.TypeNo = b.DeliveryProcess AND e.TypeSchema = 'DeliveryProcess'
                            OUTER APPLY(
                                SELECT ISNULL(SUM(x.ItemQty), 0) PickRegularQty
                                FROM SCM.PickingItem x
                                WHERE x.SoDetailId = a.SoDetailId
                                AND x.ItemType = 1
                                AND x.DoId = b.DoId
                            ) f
                            OUTER APPLY(
                                SELECT ISNULL(SUM(y.ItemQty), 0) PickFreebieQty
                                FROM SCM.PickingItem y
                                WHERE y.SoDetailId = a.SoDetailId
                                AND y.ItemType = 2
                                AND y.DoId = b.DoId
                            ) g
                            OUTER APPLY(
                                SELECT ISNULL(SUM(z.ItemQty), 0) PickSpareQty
                                FROM SCM.PickingItem z
                                WHERE z.SoDetailId = a.SoDetailId
                                AND z.ItemType = 3
                                AND z.DoId = b.DoId
                            ) h
                            WHERE b.DoId = @DoId and b.DoQty > 0

                            UNION ALL

                            SELECT
                                a.SoDetailId, 
                                '2' as PickType,
                                b.DoDetailId, 
                                b.DoQty RegularQty, 
                                ISNULL(b.FreebieQty, 0) FreebieQty, 
                                ISNULL(b.SpareQty, 0) SpareQty,
                                c.DoId, 
                                c.Status, 
                                c.DoErpPrefix + '-' + c.DoErpNo DoErpFullNo,
                                d.MtlItemId, 
                                d.MtlItemNo, 
                                d.MtlItemName, 
                                d.MtlItemSpec,
                                ISNULL(e.TypeName, '') DeliveryProcess,
                                ISNULL(f.PickRegularQty, 0) PickRegularQty, 
                                ISNULL(g.PickFreebieQty, 0) PickFreebieQty, 
                                ISNULL(h.PickSpareQty, 0) PickSpareQty,
	                        	ISNULL((
                                                        SELECT aa.PickingItemId, aa.BarcodeId, aa.SoDetailId, aa.ItemStatus, aa.ItemType, aa.ItemQty
                                                        , ISNULL(ab.BarcodeNo, '') ItemBarcode, ISNULL(ac.ItemValue, '') LetteringNo
                                                        , ISNULL(aa.LotNumber, '') LotNumber
	                        							, ISNULL(ad.TrayNo, '') TrayNo
                                                        FROM SCM.PickingItem aa
                                                        LEFT JOIN MES.Barcode ab ON ab.BarcodeId = aa.BarcodeId
                                                        OUTER APPLY (
                                                            SELECT aca.ItemValue
                                                            FROM MES.BarcodeAttribute aca
                                                            WHERE aca.ItemNo = 'Lettering'
                                                            AND aca.BarcodeId = aa.BarcodeId
                                                        ) ac
	                        							LEFT JOIN MES.Tray ad ON ad.BarcodeNo = ab.BarcodeNo
                                                        WHERE 1=1
                                                        AND aa.SoDetailId = a.SoDetailId
                                                        AND aa.DoId = b.DoId
                                                        AND aa.ItemType = '2'
                                                        ORDER BY aa.ItemType, aa.ItemStatus
                                                        FOR JSON PATH, ROOT('data')
                                                    ), '') ItemDetail
                            FROM SCM.SoDetail a
                            INNER JOIN SCM.DoDetail b ON b.SoDetailId = a.SoDetailId
                            INNER JOIN SCM.DeliveryOrder c ON c.DoId = b.DoId
                            LEFT JOIN PDM.MtlItem d ON d.MtlItemId = a.MtlItemId
                            LEFT JOIN BAS.[Type] e ON e.TypeNo = b.DeliveryProcess AND e.TypeSchema = 'DeliveryProcess'
                            OUTER APPLY(
                                SELECT ISNULL(SUM(x.ItemQty), 0) PickRegularQty
                                FROM SCM.PickingItem x
                                WHERE x.SoDetailId = a.SoDetailId
                                AND x.ItemType = 1
                                AND x.DoId = b.DoId
                            ) f
                            OUTER APPLY(
                                SELECT ISNULL(SUM(y.ItemQty), 0) PickFreebieQty
                                FROM SCM.PickingItem y
                                WHERE y.SoDetailId = a.SoDetailId
                                AND y.ItemType = 2
                                AND y.DoId = b.DoId
                            ) g
                            OUTER APPLY(
                                SELECT ISNULL(SUM(z.ItemQty), 0) PickSpareQty
                                FROM SCM.PickingItem z
                                WHERE z.SoDetailId = a.SoDetailId
                                AND z.ItemType = 3
                                AND z.DoId = b.DoId
                            ) h
                            WHERE b.DoId = @DoId and ISNULL(b.FreebieQty, 0) > 0

                            UNION ALL

                            SELECT
                                a.SoDetailId, 
                                '3' as PickType,
                                b.DoDetailId, 
                                b.DoQty RegularQty, 
                                ISNULL(b.FreebieQty, 0) FreebieQty, 
                                ISNULL(b.SpareQty, 0) SpareQty,
                                c.DoId, 
                                c.Status, 
                                c.DoErpPrefix + '-' + c.DoErpNo DoErpFullNo,
                                d.MtlItemId, 
                                d.MtlItemNo, 
                                d.MtlItemName, 
                                d.MtlItemSpec,
                                ISNULL(e.TypeName, '') DeliveryProcess,
                                ISNULL(f.PickRegularQty, 0) PickRegularQty, 
                                ISNULL(g.PickFreebieQty, 0) PickFreebieQty, 
                                ISNULL(h.PickSpareQty, 0) PickSpareQty,
	                        	ISNULL((
                                                        SELECT aa.PickingItemId, aa.BarcodeId, aa.SoDetailId, aa.ItemStatus, aa.ItemType, aa.ItemQty
                                                        , ISNULL(ab.BarcodeNo, '') ItemBarcode, ISNULL(ac.ItemValue, '') LetteringNo
                                                        , ISNULL(aa.LotNumber, '') LotNumber
	                        							, ISNULL(ad.TrayNo, '') TrayNo
                                                        FROM SCM.PickingItem aa
                                                        LEFT JOIN MES.Barcode ab ON ab.BarcodeId = aa.BarcodeId
                                                        OUTER APPLY (
                                                            SELECT aca.ItemValue
                                                            FROM MES.BarcodeAttribute aca
                                                            WHERE aca.ItemNo = 'Lettering'
                                                            AND aca.BarcodeId = aa.BarcodeId
                                                        ) ac
	                        							LEFT JOIN MES.Tray ad ON ad.BarcodeNo = ab.BarcodeNo
                                                        WHERE 1=1
                                                        AND aa.SoDetailId = a.SoDetailId
                                                        AND aa.DoId = b.DoId
                                                        AND aa.ItemType = '3'

                                                        ORDER BY aa.ItemType, aa.ItemStatus
                                                        FOR JSON PATH, ROOT('data')
                                                    ), '') ItemDetail
                            FROM SCM.SoDetail a
                            INNER JOIN SCM.DoDetail b ON b.SoDetailId = a.SoDetailId
                            INNER JOIN SCM.DeliveryOrder c ON c.DoId = b.DoId
                            LEFT JOIN PDM.MtlItem d ON d.MtlItemId = a.MtlItemId
                            LEFT JOIN BAS.[Type] e ON e.TypeNo = b.DeliveryProcess AND e.TypeSchema = 'DeliveryProcess'
                            OUTER APPLY(
                                SELECT ISNULL(SUM(x.ItemQty), 0) PickRegularQty
                                FROM SCM.PickingItem x
                                WHERE x.SoDetailId = a.SoDetailId
                                AND x.ItemType = 1
                                AND x.DoId = b.DoId
                            ) f
                            OUTER APPLY(
                                SELECT ISNULL(SUM(y.ItemQty), 0) PickFreebieQty
                                FROM SCM.PickingItem y
                                WHERE y.SoDetailId = a.SoDetailId
                                AND y.ItemType = 2
                                AND y.DoId = b.DoId
                            ) g
                            OUTER APPLY(
                                SELECT ISNULL(SUM(z.ItemQty), 0) PickSpareQty
                                FROM SCM.PickingItem z
                                WHERE z.SoDetailId = a.SoDetailId
                                AND z.ItemType = 3
                                AND z.DoId = b.DoId
                            ) h
                            WHERE b.DoId = @DoId and ISNULL(b.SpareQty, 0) > 0
                            )
                            SELECT * FROM DoDetailData
                            ORDER BY SoDetailId,MtlItemId, PickType;";
                    dynamicParameters.Add("DoId", DoId);

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

        #region //GetPickitem 取得已揀貨條碼 --Gpai 241111
        public string GetPickitem(int MtlItemId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"select a.* 
                            from SCM.PickingItem a
							inner join MES.Barcode b on a.BarcodeId = b.BarcodeId
							inner join MES.ManufactureOrder c on b.MoId = c.MoId
							inner join MES.WipOrder d on c.WoId = d.WoId
							inner join PDM.MtlItem e on d.MtlItemId = e.MtlItemId
                            where e.MtlItemId = @MtlItemId
                            order by a.PickingItemId desc
                            ";
                    dynamicParameters.Add("MtlItemId", MtlItemId);

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


        #region //GetPickingLot 取得已揀貨無條碼批號 --Shintokuro 2024-12-02
        public string GetPickingLot(int PickingItemId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"SELECT a.PickingLotId ,a.PickingItemId,a.LotNumber,a.ItemQty
                            FROM SCM.PickingLot a
                            WHERE a.PickingItemId = @PickingItemId
                            ";
                    dynamicParameters.Add("PickingItemId", PickingItemId);

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

        #region //GetDeliveryLotList 取得已揀貨無條碼批號 --Shintokuro 2024-12-02
        public string GetDeliveryLotList(int DoDetailId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    sql = @"-----------出貨轉暫出----
                            SELECT LotNumber, SUM(ItemQty) AS ItemQty,DoDetailId,ItemType
                            FROM (
                                -- 子查詢 1：非空 LotNumber 的數據
                                SELECT a.DoDetailId,b.LotNumber, b.ItemQty,b.ItemType
                                FROM SCM.DoDetail a
                                INNER JOIN SCM.PickingItem b ON a.DoId = b.DoId AND a.SoDetailId = b.SoDetailId
                                WHERE a.DoDetailId = @DoDetailId
                                AND b.LotNumber != ''
                                UNION ALL
                                -- 子查詢 2：空 LotNumber 的數據
                                SELECT c.DoDetailId,a.LotNumber, a.ItemQty,b.ItemType
                                FROM SCM.PickingLot a
                                INNER JOIN SCM.PickingItem b ON a.PickingItemId = b.PickingItemId
                                INNER JOIN SCM.DoDetail c ON b.DoId = c.DoId AND b.SoDetailId = c.SoDetailId
                                WHERE b.LotNumber = ''
                                AND c.DoDetailId = @DoDetailId
                            ) AS CombinedResults
                            GROUP BY LotNumber,DoDetailId,ItemType
                            ORDER BY LotNumber
                            ";
                    dynamicParameters.Add("DoDetailId", DoDetailId);

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
        #region //AddPickingCarton -- 物流箱資料新增 -- Zoey 2022.10.12
        public string AddPickingCarton(int PackingId, string CartonName, string CartonBarcode, string CartonRemark)
        {
            try
            {
                if (CartonBarcode.Length <= 0) throw new SystemException("【物流箱條碼】不能為空!");
                if (CartonBarcode.Length > 50) throw new SystemException("【物流箱條碼】長度錯誤!");
                if (CartonName.Length <= 0) throw new SystemException("【物流箱名稱】不能為空!");
                if (CartonName.Length > 100) throw new SystemException("【物流箱名稱】長度錯誤!");
                if (PackingId <= 0) throw new SystemException("【物流箱規格】不能為空!");
                if (CartonRemark.Length > 100) throw new SystemException("【備註】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷物流箱條碼是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.PickingCarton
                                WHERE CartonBarcode = @CartonBarcode";
                        dynamicParameters.Add("CartonBarcode", CartonBarcode);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0) throw new SystemException("【物流箱條碼】重複，請重新輸入!");
                        #endregion

                        #region //判斷包材資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.Packing
                                WHERE CompanyId = @CompanyId
                                AND PackingId = @PackingId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("PackingId", PackingId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("包材資料錯誤!");
                        #endregion

                        #region //撈取UomId
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT UomId
                                FROM PDM.UnitOfMeasure
                                WHERE UomNo = 'KG'
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);

                        int uomId = -1;
                        foreach (var item in result3)
                        {
                            uomId = item.UomId;
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.PickingCarton (PackingId, CartonName, CartonBarcode, CartonRemark, TotalWeight, UomId, PrintStatus, PrintCount
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.PickingCartonId
                                VALUES (@PackingId, @CartonName, @CartonBarcode, @CartonRemark, @TotalWeight, @UomId, @PrintStatus, @PrintCount
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                PackingId,
                                CartonName,
                                CartonBarcode,
                                CartonRemark,
                                TotalWeight = 0,
                                UomId = uomId,
                                PrintStatus = 'N',
                                PrintCount = 0,
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

        #region //AddPickingItem -- 揀貨物件資料新增 -- Zoey 2022.10.24
        public string AddPickingItem(int DoDetailId, int PickingCartonId, string InputStatus
            , string Barcode, int NoBarcodeQty, int SubstituteQty, string ItemType
            , string NoBarcodeLotNumber, string SubstituteLotNumber)
        {
            try
            {
                if (InputStatus.Length <= 0) throw new SystemException("【輸入狀態】不能為空!");
                if (!Regex.IsMatch(InputStatus, "^(B|N|S)$", RegexOptions.IgnoreCase)) throw new SystemException("【輸入狀態】錯誤!");
                if (!Regex.IsMatch(ItemType, "^(1|2|3)$", RegexOptions.IgnoreCase)) throw new SystemException("【揀貨產品類型】錯誤!");
                int trayId = -1;

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

                        var resultCompany = sqlConnection.Query(sql, dynamicParameters);
                        if (resultCompany.Count() <= 0) throw new SystemException("【公司】資料錯誤!");

                        foreach (var item in resultCompany)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        }
                        #endregion

                        #region //判斷出貨物件資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.SoDetailId, a.DoId, a.DoQty
                                , ISNULL(a.FreebieQty, 0) FreebieQty, ISNULL(a.SpareQty, 0) SpareQty
                                , ISNULL(b.PickRegularQty, 0) PickRegularQty, ISNULL(c.PickFreebieQty, 0) PickFreebieQty, ISNULL(d.PickSpareQty, 0) PickSpareQty
                                , e.DoDate
                                , g.LotManagement, g.MtlItemId, g.MtlItemNo
                                FROM SCM.DoDetail a
                                OUTER APPLY (
                                    SELECT SUM(ba.ItemQty) PickRegularQty
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
                                INNER JOIN SCM.SoDetail f ON a.SoDetailId = f.SoDetailId
                                INNER JOIN PDM.MtlItem g ON f.MtlItemId = g.MtlItemId
                                WHERE e.CompanyId = @CompanyId
                                AND a.DoDetailId = @DoDetailId
                                ORDER BY a.SoDetailId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("DoDetailId", DoDetailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("出貨物件資料錯誤!");

                        int SoDetailId = -1,
                            DoId = -1,
                            DoQty = 0,
                            FreebieQty = 0,
                            SpareQty = 0,
                            PickRegularQty = 0,
                            PickFreebieQty = 0,
                            PickSpareQty = 0,
                            MtlItemId = 0;

                        string LotManagement = "",
                               MtlItemNo = "";

                        DateTime DoDate = default(DateTime);

                        foreach (var item in result)
                        {
                            SoDetailId = Convert.ToInt32(item.SoDetailId);
                            DoId = Convert.ToInt32(item.DoId);
                            DoQty = Convert.ToInt32(item.DoQty);
                            FreebieQty = Convert.ToInt32(item.FreebieQty);
                            SpareQty = Convert.ToInt32(item.SpareQty);
                            PickRegularQty = Convert.ToInt32(item.PickRegularQty);
                            PickFreebieQty = Convert.ToInt32(item.PickFreebieQty);
                            PickSpareQty = Convert.ToInt32(item.PickSpareQty);
                            LotManagement = item.LotManagement;
                            MtlItemId = item.MtlItemId;
                            MtlItemNo = item.MtlItemNo;
                            DoDate = Convert.ToDateTime(item.DoDate);
                        }

                        if (PickRegularQty > DoQty) throw new SystemException("已超出【正常品】可揀貨數量!");
                        if (PickFreebieQty > FreebieQty) throw new SystemException("已超出【贈品】可揀貨數量!");
                        if (PickSpareQty > SpareQty) throw new SystemException("已超出【備品】可揀貨數量!");
                        #endregion

                        using (SqlConnection erpSqlConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //判斷是否有尚未結案之訂單
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(a.TD001)) + '-' + LTRIM(RTRIM(a.TD002)) SoErpFullNo, a.TD003 SoSequence
                                    , a.TD008 SoQty, a.TD048
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
                                    INNER JOIN COPTC d ON a.TD001 = d.TC001 AND a.TD002 = d.TC002
                                    WHERE a.TD004 = @MtlItemNo
                                    AND a.TD016 = 'N'
                                    AND a.TD048 < @DoDate
                                    AND d.TC027 = 'Y'
                                    AND (a.TD008 - b1.TG009 + b1.TG021) > 0";

                            //AND ((a.TD008 - b1.TG009 + b1.TG021) > 0 OR (a.TD008 - c3.TH008 + c4.TJ007) > 0) 移除銷貨
                            dynamicParameters.Add("MtlItemNo", MtlItemNo);
                            dynamicParameters.Add("DoDate", DoDate.ToString("yyyyMMdd"));

                            var resultErpAccounting = erpSqlConnection.Query(sql, dynamicParameters);
                            if (resultErpAccounting.Count() > 0) throw new SystemException("有未結案之訂單，請優先撿貨!");
                            #endregion
                        }

                        #region //判斷物流箱資料是否正確
                        if (PickingCartonId != -999)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 b.BindStatus
                                    FROM SCM.PickingCarton a
                                    OUTER APPLY (
                                        SELECT ISNULL((
                                            SELECT TOP 1 1
                                            FROM SCM.PickingItem z
                                            WHERE z.PickingCartonId = a.PickingCartonId
                                            AND z.DoId != @DoId
                                        ), 0) BindStatus
                                    ) b
                                    INNER JOIN SCM.Packing c ON a.PackingId = c.PackingId
                                    WHERE c.CompanyId = @CompanyId
                                    AND a.PickingCartonId = @PickingCartonId";
                            dynamicParameters.Add("DoId", DoId);
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("PickingCartonId", PickingCartonId);

                            var result2 = sqlConnection.Query(sql, dynamicParameters);
                            if (result2.Count() <= 0) throw new SystemException("物流箱資料錯誤!");

                            int BindStatus = -1;
                            foreach (var item in result2)
                            {
                                BindStatus = Convert.ToInt32(item.BindStatus);
                            }

                            if (BindStatus != 0) throw new SystemException("物流箱已綁定其他出貨單!");
                        }
                        #endregion

                        var inventoryHold = new InventoryTempHold();

                        int rowsAffected = 0;
                        switch (InputStatus)
                        {
                            case "B":
                                if (Barcode.Length <= 0) throw new SystemException("【物件條碼】不能為空!");

                                #region //是否為Tray盤
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 BarcodeNo, TrayNo, TrayId
                                        FROM MES.Tray
                                        WHERE TrayNo = @TrayNo";
                                dynamicParameters.Add("TrayNo", Barcode);
                                var trayReult = sqlConnection.Query(sql, dynamicParameters);

                                if (trayReult.Count() > 0)
                                {
                                    foreach (var item in trayReult)
                                    {
                                        Barcode = item.BarcodeNo;
                                        trayId = item.TrayId;
                                    }
                                }
                                #endregion

                                #region //判斷物件條碼資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.BarcodeId, a.MoId, a.BarcodeQty
                                        FROM MES.Barcode a
                                        WHERE a.BarcodeStatus = @BarcodeStatus
                                        AND a.BarcodeNo = @BarcodeNo
                                        AND EXISTS (
                                            SELECT TOP 1 1
                                            FROM MES.ManufactureOrder z
                                            INNER JOIN MES.WipOrder y ON z.WoId = y.WoId
                                            INNER JOIN SCM.SoDetail x ON y.MtlItemId = x.MtlItemId
                                            INNER JOIN SCM.DoDetail w ON x.SoDetailId = w.SoDetailId
                                            WHERE z.MoId = a.MoId
                                            AND w.DoDetailId = @DoDetailId
                                            AND y.CompanyId = @CompanyId
                                        )";
                                dynamicParameters.Add("BarcodeStatus", "8");
                                dynamicParameters.Add("BarcodeNo", Barcode);
                                dynamicParameters.Add("DoDetailId", DoDetailId);
                                dynamicParameters.Add("CompanyId", CurrentCompany);

                                var resultBarcode = sqlConnection.Query(sql, dynamicParameters);
                                if (resultBarcode.Count() <= 0) throw new SystemException("物件條碼資料錯誤!");

                                int BarcodeId = -1,
                                    MoId = -1,
                                    BarcodeQty = 0;
                                foreach (var item in resultBarcode)
                                {
                                    BarcodeId = Convert.ToInt32(item.BarcodeId);
                                    MoId = Convert.ToInt32(item.MoId);
                                    BarcodeQty = Convert.ToInt32(item.BarcodeQty);
                                }

                                switch (ItemType)
                                {
                                    case "1":
                                        if ((PickRegularQty + BarcodeQty) > DoQty) throw new SystemException("已超過【正常品】可出貨之數量!");
                                        break;
                                    case "2":
                                        if ((PickFreebieQty + BarcodeQty) > FreebieQty) throw new SystemException("已超過【贈品】可出貨之數量!");
                                        break;
                                    case "3":
                                        if ((PickSpareQty + BarcodeQty) > SpareQty) throw new SystemException("已超過【備品】可出貨之數量!");
                                        break;
                                }
                                #endregion

                                #region //判斷是否已經揀貨
                                sql = @"SELECT TOP 1 1
                                        FROM SCM.PickingItem a
                                        WHERE a.BarcodeId = @BarcodeId
                                        AND a.SoDetailId = @SoDetailId
                                        AND a.DoId = @DoId";
                                dynamicParameters.Add("BarcodeId", BarcodeId);
                                dynamicParameters.Add("SoDetailId", SoDetailId);
                                dynamicParameters.Add("DoId", DoId);

                                //sql = @"SELECT TOP 1 1
                                //        FROM SCM.PickingItem a
                                //        WHERE a.BarcodeId = @BarcodeId";
                                //dynamicParameters.Add("BarcodeId", BarcodeId);

                                var resultPicking = sqlConnection.Query(sql, dynamicParameters);
                                if (resultPicking.Count() > 0) throw new SystemException("該物件條碼已經揀貨!");
                                #endregion

                                #region //若需批號管理，取得相關批號資料
                                string lotNumber = "";
                                if (LotManagement != "N")
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT b.LotNumber
                                            FROM MES.MweBarcode a 
                                            INNER JOIN MES.MweDetail b ON a.MweDetailId = b.MweDetailId
                                            WHERE a.BarcodeId = @BarcodeId
                                            AND b.ConfirmStatus = 'Y'";
                                    dynamicParameters.Add("BarcodeId", BarcodeId);

                                    var MweBarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                                    if (MweBarcodeResult.Count() <= 0) throw new SystemException("查無入庫紀錄!!");

                                    foreach (var item in MweBarcodeResult)
                                    {
                                        if (item.LotNumber == null) throw new SystemException("此入庫條碼尚未設定批號!!");
                                        lotNumber = item.LotNumber;
                                    }
                                }
                                #endregion

                                #region //判斷庫存 2024.11.22.新增庫存卡控
                                if (CurrentCompany != 4 && CurrentCompany != 8)
                                {
                                    // 檢查條碼是否已被暫扣
                                    if (inventoryHold.IsBarcodeHeld(sqlConnection, Barcode, SoDetailId))
                                    {
                                        throw new SystemException("此條碼已被暫扣!");
                                    }

                                    // 檢查庫存是否足夠
                                    if (!inventoryHold.CheckInventoryAvailable(new SqlConnection(ErpConnectionStrings), sqlConnection,
                                        MtlItemNo, BarcodeQty))
                                    {
                                        throw new SystemException("A倉庫存不足（含暫扣數量）!");
                                    }

                                    // 建立暫扣
                                    inventoryHold.CreateBarcodeHold(sqlConnection, MtlItemNo, Barcode,
                                        BarcodeQty, SoDetailId, DoId, LastModifiedBy);
                                }
                                #endregion

                                #region //新增揀貨資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO SCM.PickingItem (PickingCartonId, SoDetailId, DoId, MoId, BarcodeId, LotNumber, ItemType, ItemQty, ItemStatus
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.PickingItemId
                                        VALUES (@PickingCartonId, @SoDetailId, @DoId, @MoId, @BarcodeId, @LotNumber, @ItemType, @ItemQty, @ItemStatus
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        PickingCartonId = PickingCartonId != -999 ? PickingCartonId : (int?)null,
                                        SoDetailId,
                                        DoId,
                                        MoId,
                                        BarcodeId,
                                        lotNumber,
                                        ItemType,
                                        ItemQty = BarcodeQty,
                                        ItemStatus = "B",
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });
                                var insertBarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertBarcodeResult.Count();
                                #endregion

                                if (trayReult.Count() > 0)
                                {
                                    #region //更新 Tray 解除綁定 BarcodeNo
                                    string nullstring = null;
                                    int oldTrayBarcodeLogId = -1;
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 BarcodeNo,  TrayId, TrayBarcodeLogId
                                        FROM MES.TrayBarcodeLog
                                        WHERE TrayId = @TrayId AND BarcodeNo = @BarcodeNo AND ReMoveBindDate IS NULL";
                                    dynamicParameters.Add("TrayId", trayId);
                                    dynamicParameters.Add("BarcodeNo", Barcode);


                                    var trayLogReult = sqlConnection.Query(sql, dynamicParameters);
                                    if (trayLogReult.Count() > 0)
                                    {
                                        foreach (var item in trayLogReult)
                                        {
                                            oldTrayBarcodeLogId = item.TrayBarcodeLogId;
                                        }
                                    }

                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE MES.Tray
                                SET 
                                BarcodeNo = @BarcodeNo,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE TrayId = @TrayId";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            BarcodeNo = nullstring,
                                            LastModifiedDate,
                                            LastModifiedBy,
                                            TrayId = trayId
                                        });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //更新 TrayBarcodeLog
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE MES.TrayBarcodeLog
                                SET 
                                ReMoveBindDate = @LastModifiedDate,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE TrayBarcodeLogId = @TrayBarcodeLogId";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            ReMoveBindDate = LastModifiedDate,
                                            LastModifiedDate,
                                            LastModifiedBy,
                                            TrayBarcodeLogId = oldTrayBarcodeLogId
                                        });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }

                                break;
                            case "N":
                                if (NoBarcodeQty <= 0) throw new SystemException("請輸入【物件數量】!");

                                if (LotManagement != "N" && NoBarcodeLotNumber.Length <= 0) throw new SystemException("此品號需進行批號管控，批號不能為空!!");

                                if (LotManagement != "N")
                                {
                                    #region //確認此品號批號資料是否正確
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM SCM.LotNumber a 
                                            WHERE a.MtlItemId = @MtlItemId 
                                            ANd a.LotNumberNo = @LotNumber";
                                    dynamicParameters.Add("MtlItemId", MtlItemId);
                                    dynamicParameters.Add("LotNumber", NoBarcodeLotNumber);

                                    var LotNumberResult = sqlConnection.Query(sql, dynamicParameters);

                                    if (LotNumberResult.Count() <= 0) throw new SystemException("此批號查無相關資料!!");
                                    #endregion
                                }

                                #region //判斷庫存 2024.11.22.新增庫存卡控
                                if (CurrentCompany != 4 && CurrentCompany != 8)
                                {
                                    // 檢查庫存是否足夠
                                    if (!inventoryHold.CheckInventoryAvailable(new SqlConnection(ErpConnectionStrings), sqlConnection,
                                        MtlItemNo, NoBarcodeQty))
                                    {
                                        throw new SystemException("A倉庫存不足（含暫扣數量）!");
                                    }

                                    // 建立暫扣
                                    inventoryHold.CreateNoBarcodeHold(sqlConnection, MtlItemNo, NoBarcodeQty,
                                        SoDetailId, DoId, LastModifiedBy);
                                }
                                #endregion

                                switch (ItemType)
                                {
                                    case "1":
                                        if ((PickRegularQty + NoBarcodeQty) > DoQty) throw new SystemException("已超過【正常品】可出貨之數量!");
                                        break;
                                    case "2":
                                        if ((PickFreebieQty + NoBarcodeQty) > FreebieQty) throw new SystemException("已超過【贈品】可出貨之數量!");
                                        break;
                                    case "3":
                                        if ((PickSpareQty + NoBarcodeQty) > SpareQty) throw new SystemException("已超過【備品】可出貨之數量!");
                                        break;
                                }

                                #region //判斷物件中是否有無條碼資料
                                int PickingItemId = -1;
                                sql = @"SELECT TOP 1 a.PickingItemId
                                        FROM SCM.PickingItem a
                                        WHERE ISNULL(a.PickingCartonId, -999) = @PickingCartonId
                                        AND a.SoDetailId = @SoDetailId
                                        AND a.DoId = @DoId
                                        AND a.ItemType = @ItemType
                                        AND a.ItemStatus = @ItemStatus";
                                dynamicParameters.Add("PickingCartonId", PickingCartonId);
                                dynamicParameters.Add("SoDetailId", SoDetailId);
                                dynamicParameters.Add("DoId", DoId);
                                dynamicParameters.Add("ItemType", ItemType);
                                dynamicParameters.Add("ItemStatus", "N");

                                var resultNoBarcode = sqlConnection.Query(sql, dynamicParameters);

                                bool statusNoBarcode = resultNoBarcode.Count() > 0;
                                foreach (var item in resultNoBarcode)
                                {
                                    PickingItemId = item.PickingItemId;
                                }

                                #endregion

                                if (statusNoBarcode)
                                {
                                    if (LotManagement != "N")
                                    {
                                        #region //修改揀貨資料
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"UPDATE SCM.PickingItem SET
                                                ItemQty = ItemQty + @ItemQty,
                                                LotNumber = '',
                                                LastModifiedDate = @LastModifiedDate,
                                                LastModifiedBy = @LastModifiedBy
                                                WHERE ISNULL(PickingCartonId, -999) = @PickingCartonId
                                                AND SoDetailId = @SoDetailId
                                                AND DoId = @DoId
                                                AND ItemType = @ItemType
                                                AND ItemStatus = @ItemStatus";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                ItemQty = NoBarcodeQty,
                                                NoBarcodeLotNumber,
                                                LastModifiedDate,
                                                LastModifiedBy,
                                                PickingCartonId,
                                                SoDetailId,
                                                DoId,
                                                ItemType,
                                                ItemStatus = "N"
                                            });
                                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                        #endregion


                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT TOP 1 a.PickingItemId
                                                FROM SCM.PickingLot a
                                                WHERE 1=1
                                                AND a.PickingItemId = @PickingItemId
                                                AND a.LotNumber = @LotNumber";
                                        dynamicParameters.Add("PickingItemId", PickingItemId);
                                        dynamicParameters.Add("LotNumber", NoBarcodeLotNumber);

                                        var resultPickingLot = sqlConnection.Query(sql, dynamicParameters);
                                        if (resultPickingLot.Count() > 0)
                                        {
                                            #region //修改撿貨無條碼批號資料
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"UPDATE SCM.PickingLot SET
                                                    ItemQty = ItemQty + @ItemQty,
                                                    LastModifiedDate = @LastModifiedDate,
                                                    LastModifiedBy = @LastModifiedBy
                                                    WHERE LotNumber = @NoBarcodeLotNumber
                                                    AND PickingItemId = @PickingItemId
                                                    ";
                                            dynamicParameters.AddDynamicParams(
                                                new
                                                {
                                                    ItemQty = NoBarcodeQty,
                                                    LastModifiedDate,
                                                    LastModifiedBy,
                                                    NoBarcodeLotNumber,
                                                    PickingItemId
                                                });
                                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                            #endregion
                                        }
                                        else
                                        {
                                            #region //新增修改撿貨無條碼批號資料
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"INSERT INTO SCM.PickingLot (PickingItemId, LotNumber, ItemQty
                                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                                    OUTPUT INSERTED.PickingItemId
                                                    VALUES (@PickingItemId, @LotNumber, @ItemQty
                                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                            dynamicParameters.AddDynamicParams(
                                                new
                                                {
                                                    PickingItemId,
                                                    LotNumber = NoBarcodeLotNumber,
                                                    ItemQty = NoBarcodeQty,
                                                    CreateDate,
                                                    LastModifiedDate,
                                                    CreateBy,
                                                    LastModifiedBy
                                                });
                                            var insertNoBarcodeResult1 = sqlConnection.Query(sql, dynamicParameters);

                                            rowsAffected += insertNoBarcodeResult1.Count();
                                            #endregion
                                        }
                                    }
                                    else
                                    {
                                        #region //修改揀貨資料
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"UPDATE SCM.PickingItem SET
                                                ItemQty = ItemQty + @ItemQty,
                                                LotNumber = @NoBarcodeLotNumber,
                                                LastModifiedDate = @LastModifiedDate,
                                                LastModifiedBy = @LastModifiedBy
                                                WHERE ISNULL(PickingCartonId, -999) = @PickingCartonId
                                                AND SoDetailId = @SoDetailId
                                                AND DoId = @DoId
                                                AND ItemType = @ItemType
                                                AND ItemStatus = @ItemStatus";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                ItemQty = NoBarcodeQty,
                                                NoBarcodeLotNumber,
                                                LastModifiedDate,
                                                LastModifiedBy,
                                                PickingCartonId,
                                                SoDetailId,
                                                DoId,
                                                ItemType,
                                                ItemStatus = "N"
                                            });
                                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                        #endregion
                                    }
                                }
                                else
                                {
                                    if (LotManagement != "N")
                                    {
                                        #region //新增揀貨資料
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO SCM.PickingItem (PickingCartonId, SoDetailId, DoId, ItemType, ItemQty, ItemStatus, LotNumber
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.PickingItemId
                                            VALUES (@PickingCartonId, @SoDetailId, @DoId, @ItemType, @ItemQty, @ItemStatus, @LotNumber
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                PickingCartonId = PickingCartonId != -999 ? PickingCartonId : (int?)null,
                                                SoDetailId,
                                                DoId,
                                                ItemType,
                                                ItemQty = NoBarcodeQty,
                                                ItemStatus = "N",
                                                LotNumber = "",
                                                CreateDate,
                                                LastModifiedDate,
                                                CreateBy,
                                                LastModifiedBy
                                            });
                                        var insertNoBarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                                        rowsAffected += insertNoBarcodeResult.Count();

                                        foreach (var item in insertNoBarcodeResult)
                                        {
                                            #region //新增修改撿貨無條碼批號資料
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"INSERT INTO SCM.PickingLot (PickingItemId, LotNumber, ItemQty
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.PickingItemId
                                            VALUES (@PickingItemId, @LotNumber, @ItemQty
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                            dynamicParameters.AddDynamicParams(
                                                new
                                                {
                                                    PickingItemId = item.PickingItemId,
                                                    LotNumber = NoBarcodeLotNumber,
                                                    ItemQty = NoBarcodeQty,
                                                    CreateDate,
                                                    LastModifiedDate,
                                                    CreateBy,
                                                    LastModifiedBy
                                                });
                                            insertNoBarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                                            rowsAffected += insertNoBarcodeResult.Count();
                                            #endregion
                                        }
                                        #endregion
                                    }
                                    else
                                    {
                                        #region //新增揀貨資料
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO SCM.PickingItem (PickingCartonId, SoDetailId, DoId, ItemType, ItemQty, ItemStatus, LotNumber
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.PickingItemId
                                            VALUES (@PickingCartonId, @SoDetailId, @DoId, @ItemType, @ItemQty, @ItemStatus, @LotNumber
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                PickingCartonId = PickingCartonId != -999 ? PickingCartonId : (int?)null,
                                                SoDetailId,
                                                DoId,
                                                ItemType,
                                                ItemQty = NoBarcodeQty,
                                                ItemStatus = "N",
                                                LotNumber = NoBarcodeLotNumber,
                                                CreateDate,
                                                LastModifiedDate,
                                                CreateBy,
                                                LastModifiedBy
                                            });
                                        var insertNoBarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                                        rowsAffected += insertNoBarcodeResult.Count();
                                        #endregion
                                    }
                                }
                                break;
                            case "S":
                                if (SubstituteQty <= 0) throw new SystemException("請輸入【物件數量】!");

                                if (LotManagement != "N" && SubstituteLotNumber.Length <= 0) throw new SystemException("此品號需進行批號管控，批號不能為空!!");

                                if (LotManagement != "N")
                                {
                                    #region //確認此品號批號資料是否正確
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM SCM.LotNumber a 
                                            WHERE a.MtlItemId = @MtlItemId 
                                            ANd a.LotNumberNo = @LotNumber";
                                    dynamicParameters.Add("MtlItemId", MtlItemId);
                                    dynamicParameters.Add("LotNumber", SubstituteLotNumber);

                                    var LotNumberResult2 = sqlConnection.Query(sql, dynamicParameters);

                                    if (LotNumberResult2.Count() <= 0) throw new SystemException("此批號查無相關資料!!");
                                    #endregion
                                }

                                switch (ItemType)
                                {
                                    case "1":
                                        if ((PickRegularQty + SubstituteQty) > DoQty) throw new SystemException("已超過【正常品】可出貨之數量!");
                                        break;
                                    case "2":
                                        if ((PickFreebieQty + SubstituteQty) > FreebieQty) throw new SystemException("已超過【贈品】可出貨之數量!");
                                        break;
                                    case "3":
                                        if ((PickSpareQty + SubstituteQty) > SpareQty) throw new SystemException("已超過【備品】可出貨之數量!");
                                        break;
                                }

                                #region //判斷物件中是否有替代品資料
                                sql = @"SELECT TOP 1 1
                                        FROM SCM.PickingItem a
                                        WHERE ISNULL(a.PickingCartonId, -999) = @PickingCartonId
                                        AND a.SoDetailId = @SoDetailId
                                        AND a.DoId = @DoId
                                        AND a.ItemType = @ItemType
                                        AND a.ItemStatus = @ItemStatus";
                                dynamicParameters.Add("PickingCartonId", PickingCartonId);
                                dynamicParameters.Add("SoDetailId", SoDetailId);
                                dynamicParameters.Add("DoId", DoId);
                                dynamicParameters.Add("ItemType", ItemType);
                                dynamicParameters.Add("ItemStatus", "S");

                                var resultSubstitute = sqlConnection.Query(sql, dynamicParameters);

                                bool statusSubstitute = resultSubstitute.Count() > 0;
                                #endregion

                                if (statusSubstitute)
                                {
                                    #region //修改揀貨資料
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE SCM.PickingItem SET
                                            ItemQty = ItemQty + @ItemQty,
                                            LotNumber = @SubstituteLotNumber,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE ISNULL(PickingCartonId, -999) = @PickingCartonId
                                            AND SoDetailId = @SoDetailId
                                            AND DoId = @DoId
                                            AND ItemType = @ItemType
                                            AND ItemStatus = @ItemStatus";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            ItemQty = SubstituteQty,
                                            SubstituteLotNumber,
                                            LastModifiedDate,
                                            LastModifiedBy,
                                            PickingCartonId,
                                            SoDetailId,
                                            DoId,
                                            ItemType,
                                            ItemStatus = "S"
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                                else
                                {
                                    #region //新增揀貨資料
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO SCM.PickingItem (PickingCartonId, SoDetailId, DoId, ItemType, ItemQty, ItemStatus, LotNumber
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.PickingItemId
                                            VALUES (@PickingCartonId, @SoDetailId, @DoId, @ItemType, @ItemQty, @ItemStatus, @LotNumber
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            PickingCartonId = PickingCartonId != -999 ? PickingCartonId : (int?)null,
                                            SoDetailId,
                                            DoId,
                                            ItemType,
                                            ItemQty = SubstituteQty,
                                            ItemStatus = "S",
                                            LotNumber = SubstituteLotNumber,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });
                                    var insertNoBarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertNoBarcodeResult.Count();
                                    #endregion
                                }
                                break;
                        }

                        #region //更新出貨單狀態
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.DeliveryOrder SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE DoId = @DoId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = "P",
                                LastModifiedDate,
                                LastModifiedBy,
                                DoId
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

        #region //AddPackagePickingItem -- 包裝條碼揀貨物件資料新增 -- GPAI 241111
        public string AddPackagePickingItem(int DoDetailId, int PickingCartonId, string InputStatus
            , string Barcode, int NoBarcodeQty, int SubstituteQty, string ItemType, string PickJson)
        {
            try
            {
                if (ItemType.Length <= 0) throw new SystemException("【輸入狀態】不能為空!");
                InputStatus = "B";
                int trayId = -1;
                List<DataModel> PickdataModels = new List<DataModel>();
                var parsedData = JsonConvert.DeserializeObject<DataModel>(PickJson);
                PickdataModels.Add(parsedData);

                // 解析 JSON 並加入列表

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

                        var resultCompany = sqlConnection.Query(sql, dynamicParameters);
                        if (resultCompany.Count() <= 0) throw new SystemException("【公司】資料錯誤!");

                        foreach (var item in resultCompany)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        }
                        #endregion

                        #region //判斷出貨物件資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.SoDetailId, a.DoId, a.DoQty
                                , ISNULL(a.FreebieQty, 0) FreebieQty, ISNULL(a.SpareQty, 0) SpareQty
                                , ISNULL(b.PickRegularQty, 0) PickRegularQty, ISNULL(c.PickFreebieQty, 0) PickFreebieQty, ISNULL(d.PickSpareQty, 0) PickSpareQty
                                , e.DoDate
                                , g.LotManagement, g.MtlItemId, g.MtlItemNo
                                FROM SCM.DoDetail a
                                OUTER APPLY (
                                    SELECT SUM(ba.ItemQty) PickRegularQty
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
                                INNER JOIN SCM.SoDetail f ON a.SoDetailId = f.SoDetailId
                                INNER JOIN PDM.MtlItem g ON f.MtlItemId = g.MtlItemId
                                WHERE e.CompanyId = @CompanyId
                                AND f.SoDetailId = @DoDetailId
                                ORDER BY a.SoDetailId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("DoDetailId", DoDetailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("出貨物件資料錯誤!");

                        int SoDetailId = -1,
                            DoId = -1,
                            DoQty = 0,
                            FreebieQty = 0,
                            SpareQty = 0,
                            PickRegularQty = 0,
                            PickFreebieQty = 0,
                            PickSpareQty = 0,
                            MtlItemId = 0;

                        string LotManagement = "",
                               MtlItemNo = "";

                        DateTime DoDate = default(DateTime);

                        foreach (var item in result)
                        {
                            SoDetailId = Convert.ToInt32(item.SoDetailId);
                            DoId = Convert.ToInt32(item.DoId);
                            DoQty = Convert.ToInt32(item.DoQty);
                            FreebieQty = Convert.ToInt32(item.FreebieQty);
                            SpareQty = Convert.ToInt32(item.SpareQty);
                            PickRegularQty = Convert.ToInt32(item.PickRegularQty);
                            PickFreebieQty = Convert.ToInt32(item.PickFreebieQty);
                            PickSpareQty = Convert.ToInt32(item.PickSpareQty);
                            LotManagement = item.LotManagement;
                            MtlItemId = item.MtlItemId;
                            MtlItemNo = item.MtlItemNo;
                            DoDate = Convert.ToDateTime(item.DoDate);
                        }

                        if (PickRegularQty > DoQty) throw new SystemException("已超出【正常品】可揀貨數量!");
                        if (PickFreebieQty > FreebieQty) throw new SystemException("已超出【贈品】可揀貨數量!");
                        if (PickSpareQty > SpareQty) throw new SystemException("已超出【備品】可揀貨數量!");
                        #endregion

                        using (SqlConnection erpSqlConnection = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //判斷是否有尚未結案之訂單
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(a.TD001)) + '-' + LTRIM(RTRIM(a.TD002)) SoErpFullNo, a.TD003 SoSequence
                                    , a.TD008 SoQty, a.TD048
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
                                    INNER JOIN COPTC d ON a.TD001 = d.TC001 AND a.TD002 = d.TC002
                                    WHERE a.TD004 = @MtlItemNo
                                    AND a.TD016 = 'N'
                                    AND a.TD048 < @DoDate
                                    AND d.TC027 = 'Y'
                                    AND ((a.TD008 - b1.TG009 + b1.TG021) > 0 OR (a.TD008 - c3.TH008 + c4.TJ007) > 0)";
                            dynamicParameters.Add("MtlItemNo", MtlItemNo);
                            dynamicParameters.Add("DoDate", DoDate.ToString("yyyyMMdd"));

                            var resultErpAccounting = erpSqlConnection.Query(sql, dynamicParameters);
                            if (resultErpAccounting.Count() > 0) throw new SystemException("有未結案之訂單，請優先撿貨!");
                            #endregion
                        }

                        #region //判斷物流箱資料是否正確
                        if (PickingCartonId != -999)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 b.BindStatus
                                    FROM SCM.PickingCarton a
                                    OUTER APPLY (
                                        SELECT ISNULL((
                                            SELECT TOP 1 1
                                            FROM SCM.PickingItem z
                                            WHERE z.PickingCartonId = a.PickingCartonId
                                            AND z.DoId != @DoId
                                        ), 0) BindStatus
                                    ) b
                                    INNER JOIN SCM.Packing c ON a.PackingId = c.PackingId
                                    WHERE c.CompanyId = @CompanyId
                                    AND a.PickingCartonId = @PickingCartonId";
                            dynamicParameters.Add("DoId", DoId);
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("PickingCartonId", PickingCartonId);

                            var result2 = sqlConnection.Query(sql, dynamicParameters);
                            if (result2.Count() <= 0) throw new SystemException("物流箱資料錯誤!");

                            int BindStatus = -1;
                            foreach (var item in result2)
                            {
                                BindStatus = Convert.ToInt32(item.BindStatus);
                            }

                            if (BindStatus != 0) throw new SystemException("物流箱已綁定其他出貨單!");
                        }
                        #endregion
                        var inventoryHold = new InventoryTempHold();

                        int rowsAffected = 0;
                        switch (InputStatus)
                        {
                            case "B":
                                if (PickdataModels.Count() <= 0) throw new SystemException("【物件條碼】不能為空!");


                                foreach (var item in PickdataModels)
                                {


                                    foreach (var itemm in item.solution)
                                    {
                                        #region //是否為Tray盤
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT TOP 1 BarcodeNo, TrayNo, TrayId
                                        FROM MES.Tray
                                        WHERE TrayNo = @TrayNo";
                                        dynamicParameters.Add("TrayNo", itemm.barcodeno);
                                        var trayReult = sqlConnection.Query(sql, dynamicParameters);

                                        if (trayReult.Count() > 0)
                                        {
                                            foreach (var item2 in trayReult)
                                            {
                                                Barcode = item2.BarcodeNo;
                                                trayId = item2.TrayId;
                                            }
                                        }
                                        #endregion

                                        #region //判斷物件條碼資料是否正確
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT a.BarcodeId, a.MoId, a.BarcodeQty
                                        FROM MES.Barcode a
                                        WHERE a.BarcodeStatus = @BarcodeStatus
                                        AND a.BarcodeNo = @BarcodeNo
                                        AND EXISTS (
                                            SELECT TOP 1 1
                                            FROM MES.ManufactureOrder z
                                            INNER JOIN MES.WipOrder y ON z.WoId = y.WoId
                                            INNER JOIN SCM.SoDetail x ON y.MtlItemId = x.MtlItemId
                                            INNER JOIN SCM.DoDetail w ON x.SoDetailId = w.SoDetailId
                                            WHERE z.MoId = a.MoId
                                            AND x.SoDetailId = @DoDetailId
                                            AND y.CompanyId = @CompanyId
                                        )";
                                        dynamicParameters.Add("BarcodeStatus", "8");
                                        dynamicParameters.Add("BarcodeNo", itemm.barcodeno);
                                        dynamicParameters.Add("DoDetailId", DoDetailId);
                                        dynamicParameters.Add("CompanyId", CurrentCompany);

                                        var resultBarcode = sqlConnection.Query(sql, dynamicParameters);
                                        if (resultBarcode.Count() <= 0) throw new SystemException("物件條碼資料錯誤!");

                                        int BarcodeId = -1,
                                            MoId = -1,
                                            BarcodeQty = 0;
                                        foreach (var item3 in resultBarcode)
                                        {
                                            BarcodeId = Convert.ToInt32(item3.BarcodeId);
                                            MoId = Convert.ToInt32(item3.MoId);
                                            BarcodeQty = Convert.ToInt32(item3.BarcodeQty);
                                        }

                                        switch (ItemType)
                                        {
                                            case "1":
                                                if ((PickRegularQty + BarcodeQty) > DoQty) throw new SystemException("已超過【正常品】可出貨之數量!");
                                                break;
                                            case "2":
                                                if ((PickFreebieQty + BarcodeQty) > FreebieQty) throw new SystemException("已超過【贈品】可出貨之數量!");
                                                break;
                                            case "3":
                                                if ((PickSpareQty + BarcodeQty) > SpareQty) throw new SystemException("已超過【備品】可出貨之數量!");
                                                break;
                                        }
                                        #endregion

                                        #region //判斷是否已經揀貨
                                        sql = @"SELECT TOP 1 1
                                        FROM SCM.PickingItem a
                                        WHERE a.BarcodeId = @BarcodeId
                                        AND a.SoDetailId = @SoDetailId
                                        AND a.DoId = @DoId";
                                        dynamicParameters.Add("BarcodeId", BarcodeId);
                                        dynamicParameters.Add("SoDetailId", SoDetailId);
                                        dynamicParameters.Add("DoId", DoId);

                                        //sql = @"SELECT TOP 1 1
                                        //        FROM SCM.PickingItem a
                                        //        WHERE a.BarcodeId = @BarcodeId";
                                        //dynamicParameters.Add("BarcodeId", BarcodeId);

                                        var resultPicking = sqlConnection.Query(sql, dynamicParameters);
                                        if (resultPicking.Count() > 0) throw new SystemException("該物件條碼已經揀貨!");
                                        #endregion

                                        #region //若需批號管理，取得相關批號資料
                                        string lotNumber = "";
                                        if (LotManagement != "N")
                                        {
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"SELECT b.LotNumber
                                            FROM MES.MweBarcode a 
                                            INNER JOIN MES.MweDetail b ON a.MweDetailId = b.MweDetailId
                                            WHERE a.BarcodeId = @BarcodeId
                                            AND b.ConfirmStatus = 'Y'";
                                            dynamicParameters.Add("BarcodeId", BarcodeId);

                                            var MweBarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                                            if (MweBarcodeResult.Count() <= 0) throw new SystemException("查無入庫紀錄!!");

                                            foreach (var item4 in MweBarcodeResult)
                                            {
                                                if (item4.LotNumber == null) throw new SystemException("此入庫條碼尚未設定批號!!");
                                                lotNumber = item4.LotNumber;
                                            }
                                        }
                                        #endregion

                                        #region //判斷庫存 2024.11.22.新增庫存卡控
                                        if (CurrentCompany != 4 && CurrentCompany != 8)
                                        {
                                            // 檢查條碼是否已被暫扣
                                            if (inventoryHold.IsBarcodeHeld(sqlConnection, itemm.barcodeno, SoDetailId))
                                            {
                                                throw new SystemException("此條碼已被暫扣!");
                                            }

                                            // 檢查庫存是否足夠
                                            if (!inventoryHold.CheckInventoryAvailable(new SqlConnection(ErpConnectionStrings), sqlConnection,
                                                MtlItemNo, BarcodeQty))
                                            {
                                                throw new SystemException("A倉庫存不足（含暫扣數量）!");
                                            }

                                            // 建立暫扣
                                            inventoryHold.CreateBarcodeHold(sqlConnection, MtlItemNo, itemm.barcodeno,
                                                BarcodeQty, SoDetailId, DoId, LastModifiedBy);
                                        }
                                        #endregion

                                        #region //新增揀貨資料
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO SCM.PickingItem (PickingCartonId, SoDetailId, DoId, MoId, BarcodeId, LotNumber, ItemType, ItemQty, ItemStatus
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.PickingItemId
                                        VALUES (@PickingCartonId, @SoDetailId, @DoId, @MoId, @BarcodeId, @LotNumber, @ItemType, @ItemQty, @ItemStatus
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                PickingCartonId = PickingCartonId != -999 ? PickingCartonId : (int?)null,
                                                SoDetailId,
                                                DoId,
                                                MoId,
                                                BarcodeId,
                                                lotNumber,
                                                ItemType,
                                                ItemQty = BarcodeQty,
                                                ItemStatus = "B",
                                                CreateDate,
                                                LastModifiedDate,
                                                CreateBy,
                                                LastModifiedBy
                                            });
                                        var insertBarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                                        rowsAffected += insertBarcodeResult.Count();
                                        #endregion

                                        if (trayReult.Count() > 0)
                                        {
                                            #region //更新 Tray 解除綁定 BarcodeNo
                                            string nullstring = null;
                                            int oldTrayBarcodeLogId = -1;
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"SELECT TOP 1 BarcodeNo,  TrayId, TrayBarcodeLogId
                                        FROM MES.TrayBarcodeLog
                                        WHERE TrayId = @TrayId AND BarcodeNo = @BarcodeNo AND ReMoveBindDate IS NULL";
                                            dynamicParameters.Add("TrayId", trayId);
                                            dynamicParameters.Add("BarcodeNo", Barcode);


                                            var trayLogReult = sqlConnection.Query(sql, dynamicParameters);
                                            if (trayLogReult.Count() > 0)
                                            {
                                                foreach (var item5 in trayLogReult)
                                                {
                                                    oldTrayBarcodeLogId = item5.TrayBarcodeLogId;
                                                }
                                            }

                                            dynamicParameters = new DynamicParameters();
                                            sql = @"UPDATE MES.Tray
                                SET 
                                BarcodeNo = @BarcodeNo,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE TrayId = @TrayId";
                                            dynamicParameters.AddDynamicParams(
                                                new
                                                {
                                                    BarcodeNo = nullstring,
                                                    LastModifiedDate,
                                                    LastModifiedBy,
                                                    TrayId = trayId
                                                });

                                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                            #endregion

                                            #region //更新 TrayBarcodeLog
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"UPDATE MES.TrayBarcodeLog
                                SET 
                                ReMoveBindDate = @LastModifiedDate,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE TrayBarcodeLogId = @TrayBarcodeLogId";
                                            dynamicParameters.AddDynamicParams(
                                                new
                                                {
                                                    ReMoveBindDate = LastModifiedDate,
                                                    LastModifiedDate,
                                                    LastModifiedBy,
                                                    TrayBarcodeLogId = oldTrayBarcodeLogId
                                                });

                                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                            #endregion
                                        }
                                    }
                                }
                                break;
                        }

                        #region //更新出貨單狀態
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.DeliveryOrder SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE DoId = @DoId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = "P",
                                LastModifiedDate,
                                LastModifiedBy,
                                DoId
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

        #region //Update
        #region //UpdatePickingCarton -- 物流箱資料更新 -- Zoey 2022.10.12
        public string UpdatePickingCarton(int PickingCartonId, int PackingId, string CartonName, string CartonBarcode, string CartonRemark)
        {
            try
            {
                if (CartonBarcode.Length <= 0) throw new SystemException("【物流箱條碼】不能為空!");
                if (CartonBarcode.Length > 50) throw new SystemException("【物流箱條碼】長度錯誤!");
                if (CartonName.Length <= 0) throw new SystemException("【物流箱名稱】不能為空!");
                if (CartonName.Length > 100) throw new SystemException("【物流箱名稱】長度錯誤!");
                if (PackingId <= 0) throw new SystemException("【物流箱規格】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷物流箱資訊是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.PickingCarton a
                                INNER JOIN SCM.Packing b ON a.PackingId = b.PackingId
                                WHERE a.PickingCartonId = @PickingCartonId
                                AND b.CompanyId = @CompanyId";
                        dynamicParameters.Add("PickingCartonId", PickingCartonId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("物流箱資料錯誤!");
                        #endregion

                        #region //判斷物流箱條碼是否重複
                        sql = @"SELECT TOP 1 1
                                FROM SCM.PickingCarton
                                WHERE CartonBarcode = @CartonBarcode
                                AND PickingCartonId != @PickingCartonId";
                        dynamicParameters.Add("CartonBarcode", CartonBarcode);
                        dynamicParameters.Add("PickingCartonId", PickingCartonId);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【物流箱條碼】重複，請重新輸入!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.PickingCarton SET
                                PackingId = @PackingId,
                                CartonName = @CartonName,
                                CartonRemark = @CartonRemark,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE PickingCartonId = @PickingCartonId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                PackingId,
                                CartonName,
                                CartonBarcode,
                                CartonRemark,
                                LastModifiedDate,
                                LastModifiedBy,
                                PickingCartonId
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

        #region //UpdatePickingItem -- 揀貨物件資料更新(物件換箱) -- Zoey 2022.10.24
        public string UpdatePickingItem(string PickingItemIds, string CartonBarcode)
        {
            try
            {
                if (PickingItemIds.Length <= 0) throw new SystemException("【物件條碼】不能為空!");
                if (CartonBarcode.Length <= 0) throw new SystemException("【物流箱條碼】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        var itemStatus = "";
                        int soDetailId = -1;
                        int itemQty = 0;
                        int doId = -1;
                        var pickingItem = PickingItemIds.Split(',');
                        int rowsAffected = 0;

                        #region //撈取原箱ID
                        int oriCarton = -1;
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT DISTINCT ISNULL(a.PickingCartonId, -999) PickingCartonId
                                FROM SCM.PickingItem a
                                WHERE a.PickingItemId IN @PickingItemId";
                        dynamicParameters.Add("PickingItemId", pickingItem);
                        var result = sqlConnection.Query(sql, dynamicParameters);

                        if (result.Count() <= 0) throw new SystemException("原物流箱資料錯誤!");

                        foreach (var item in result)
                        {
                            oriCarton = item.PickingCartonId;
                        }
                        #endregion

                        #region //撈取新箱ID
                        int newCarton = -1;
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT PickingCartonId
                                FROM SCM.PickingCarton
                                WHERE CartonBarcode = @CartonBarcode";
                        dynamicParameters.Add("CartonBarcode", CartonBarcode);
                        var result2 = sqlConnection.Query(sql, dynamicParameters);

                        if (result2.Count() <= 0) throw new SystemException("新物流箱資料錯誤!");

                        foreach (var item in result2)
                        {
                            newCarton = item.PickingCartonId;
                        }
                        #endregion

                        for (int i = 0; i < pickingItem.Length; i++)
                        {
                            string PickingLot = "";

                            #region //判斷物件條碼資訊是否正確/撈取原箱物件資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.ItemStatus, a.SoDetailId, a.ItemQty, ISNULL(a.DoId, -1) DoId,ISNULL(x.PickingLot,'N') PickingLot
                                    FROM SCM.PickingItem a
                                    OUTER APPLY(
                                        SELECT TOP 1 'Y' PickingLot 
                                        FROM SCM.PickingLot x1
                                        WHERE x1.PickingItemId = a.PickingItemId
                                    ) x
                                    WHERE a.PickingItemId = @PickingItemId";
                            dynamicParameters.Add("PickingItemId", pickingItem[i]);

                            var result3 = sqlConnection.Query(sql, dynamicParameters);
                            if (result3.Count() <= 0) throw new SystemException("物件條碼錯誤!");
                            foreach (var item in result3)
                            {
                                itemStatus = item.ItemStatus;
                                soDetailId = item.SoDetailId;
                                itemQty = item.ItemQty;
                                doId = item.DoId;
                                PickingLot = item.PickingLot;
                            }
                            #endregion

                            if (itemStatus == "B")
                            {
                                #region //修改有條碼物件箱號
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE SCM.PickingItem SET
                                        PickingCartonId = @PickingCartonId,
                                        DoId = NULL,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE PickingItemId = @PickingItemId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        PickingCartonId = newCarton,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        PickingItemId = pickingItem[i]
                                    });
                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }
                            else
                            {
                                #region //判斷新箱是否有物件
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT PickingItemId, ItemStatus
                                        FROM SCM.PickingItem
                                        WHERE ISNULL(DoId, -1) = @DoId
                                        AND SoDetailId = @SoDetailId
                                        AND PickingCartonId = @PickingCartonId";
                                dynamicParameters.Add("DoId", doId);
                                dynamicParameters.Add("SoDetailId", soDetailId);
                                dynamicParameters.Add("PickingCartonId", newCarton);
                                var exisItemStatus = sqlConnection.Query(sql, dynamicParameters).ToList();
                                #endregion

                                if (exisItemStatus.Count() > 0 && exisItemStatus.Exists(x => x.ItemStatus == itemStatus))
                                {
                                    var currentItem = exisItemStatus.Where(x => x.ItemStatus == itemStatus).FirstOrDefault();

                                    #region //修改無條碼物件/替代品數量
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE SCM.PickingItem SET
                                            ItemQty = ItemQty + @ItemQty,
                                            DoId = NULL,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE PickingItemId = @PickingItemId";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            itemQty,
                                            LastModifiedDate,
                                            LastModifiedBy,
                                            currentItem.PickingItemId
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //修改-將原無條件批號提換成新的箱子
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE SCM.PickingLot SET
                                            PickingItemId = PickingItemId,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE PickingItemId = @PickingItemId";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            itemQty,
                                            LastModifiedDate,
                                            LastModifiedBy,
                                            currentItem.PickingItemId
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //刪除原箱物件
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE SCM.PickingItem
                                            WHERE PickingItemId = @PickingItemId";
                                    dynamicParameters.Add("PickingItemId", pickingItem[i]);

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                                else
                                {
                                    #region //修改無條碼物件/替代品箱號
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE SCM.PickingItem SET
                                            PickingCartonId = @PickingCartonId,
                                            DoId = NULL,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE PickingItemId = @PickingItemId";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            PickingCartonId = newCarton,
                                            LastModifiedDate,
                                            LastModifiedBy,
                                            PickingItemId = pickingItem[i]
                                        });
                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
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

        #region //UpdateDeleiveryStatus -- 出貨單狀態更新 -- Zoey 2022.10.24
        public string UpdateDeleiveryStatus(string DoIds, string Status)
        {
            try
            {
                if (DoIds.Length <= 0) throw new SystemException("請選取出貨單");
                if (Status.Length <= 0) throw new SystemException("出貨單狀態錯誤!");

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

                        var resultCompany = sqlConnection.Query(sql, dynamicParameters);
                        if (resultCompany.Count() <= 0) throw new SystemException("【公司】資料錯誤!");

                        foreach (var item in resultCompany)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        }
                        #endregion

                        var doId = DoIds.Split(',');
                        int rowsAffected = 0;

                        for (int i = 0; i < doId.Length; i++)
                        {
                            #region //判斷出貨單資訊是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 Status, TransferStatus
                                    FROM SCM.DeliveryOrder
                                    WHERE DoId = @DoId
                                    AND CompanyId = @CompanyId";
                            dynamicParameters.Add("DoId", doId[i]);
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("出貨單資料錯誤!");

                            string originalStatus = "";
                            foreach (var item in result)
                            {
                                originalStatus = item.Status;
                            }
                            #endregion

                            switch (Status)
                            {
                                case "P":
                                    if (originalStatus != "R") throw new SystemException("操作不允許!");
                                    break;
                                case "R":
                                    switch (originalStatus)
                                    {
                                        case "P":
                                            #region //該出貨單至少要有出貨數量
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"SELECT TOP 1 1
                                                    FROM SCM.PickingItem a
                                                    WHERE a.DoId = @DoId";
                                            dynamicParameters.Add("DoId", doId[i]);

                                            var result3 = sqlConnection.Query(sql, dynamicParameters);
                                            if (result3.Count() <= 0) throw new SystemException("該出貨單無揀貨數量，無法完成揀貨!");
                                            #endregion
                                            break;
                                        case "S":
                                            #region //修改條碼狀態為完工入庫
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"UPDATE MES.Barcode SET
                                                    BarcodeStatus = @BarcodeStatus,
                                                    LastModifiedDate = @LastModifiedDate,
                                                    LastModifiedBy = @LastModifiedBy
                                                    FROM MES.Barcode a 
                                                    LEFT JOIN SCM.PickingItem b ON b.BarcodeId = a.BarcodeId
                                                    WHERE b.DoId = @DoId";
                                            dynamicParameters.AddDynamicParams(
                                                new
                                                {
                                                    BarcodeStatus = "8", //製令完工入庫
                                                    LastModifiedDate,
                                                    LastModifiedBy,
                                                    DoId = doId[i]
                                                });

                                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                            #endregion

                                            #region //釋放庫存佔存表 2024.11.22.新增庫存卡控
                                            if (CurrentCompany != 4)
                                            {
                                                var inventoryHold = new InventoryTempHold();

                                                // 檢查是否可以恢復暫扣（庫存是否足夠）
                                                if (!inventoryHold.CanRestoreHolds(new SqlConnection(ErpConnectionStrings), sqlConnection, Convert.ToInt32(doId[i])))
                                                {
                                                    throw new SystemException("目前庫存不足，無法反確認出貨!");
                                                }

                                                // 恢復暫扣記錄
                                                inventoryHold.RestoreReleasedHolds(sqlConnection, Convert.ToInt32(doId[i]), LastModifiedBy);
                                            }
                                            #endregion
                                            break;
                                        default:
                                            throw new SystemException("操作不允許!");
                                    }
                                    break;
                                case "S":
                                    if (originalStatus == "R")
                                    {
                                        #region //判斷物流箱重量是否大於0
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT DISTINCT a.PickingCartonId, a.TotalWeight 
                                                FROM SCM.PickingCarton a
                                                LEFT JOIN SCM.PickingItem b ON b.PickingCartonId = a.PickingCartonId
                                                WHERE b.DoId = @DoId
                                                AND a.TotalWeight < 0";
                                        dynamicParameters.Add("DoId", doId[i]);

                                        var result4 = sqlConnection.Query(sql, dynamicParameters);
                                        if (result4.Count() > 0) throw new SystemException("請先確認物流箱重量!");
                                        #endregion

                                        #region //修改條碼狀態為已出貨
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"UPDATE MES.Barcode SET
                                                BarcodeStatus = @BarcodeStatus,
                                                LastModifiedDate = @LastModifiedDate,
                                                LastModifiedBy = @LastModifiedBy
                                                FROM 
                                                MES.Barcode a LEFT JOIN SCM.PickingItem b ON b.BarcodeId = a.BarcodeId
                                                WHERE b.DoId =  @DoId";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                BarcodeStatus = "10", //已出貨
                                                LastModifiedDate,
                                                LastModifiedBy,
                                                DoId = doId[i]
                                            });

                                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                        #endregion

                                        #region //撈取訂單單頭Id
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT DISTINCT c.SoId
                                                FROM SCM.DeliveryOrder a 
                                                INNER JOIN SCM.DoDetail b ON b.DoId = a.DoId
                                                INNER JOIN SCM.SoDetail c ON c.SoDetailId = b.SoDetailId
                                                WHERE a.DoId = @DoId";
                                        dynamicParameters.Add("DoId", doId[i]);

                                        var resultSo = sqlConnection.Query(sql, dynamicParameters);
                                        if (resultSo.Count() <= 0) throw new SystemException("無法取得訂單資料!");
                                        #endregion

                                        #region //釋放庫存佔存表 2024.11.22.新增庫存卡控
                                        if (CurrentCompany != 4 && CurrentCompany != 8)
                                        {
                                            var inventoryHold = new InventoryTempHold();

                                            // 檢查整張出貨單的庫存是否足夠
                                            if (!inventoryHold.CheckDoInventoryAvailable(new SqlConnection(ErpConnectionStrings), sqlConnection, Convert.ToInt32(doId[i])))
                                            {
                                                throw new SystemException("A倉庫存不足，無法確認出貨!");
                                            }

                                            // 釋放此出貨單的所有暫扣庫存
                                            inventoryHold.ReleaseByDoId(sqlConnection, Convert.ToInt32(doId[i]), LastModifiedBy);
                                        }
                                        #endregion
                                    }
                                    else
                                    {
                                        throw new SystemException("操作不允許!");
                                    }
                                    break;
                                default:
                                    throw new SystemException("操作不允許!");
                            }

                            #region //修改出貨單狀態
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.DeliveryOrder SET
                                    Status = @Status,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE DoId = @DoId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    Status,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    DoId = doId[i]
                                });

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

        #region //UpdateCartonWeight -- 物流箱重量更新 -- Zoey 2022.11.15
        public string UpdateCartonWeight(int PickingCartonId, double TotalWeight, int UomId)
        {
            try
            {
                if (TotalWeight < 0) throw new SystemException("【物流箱重量】不可小於0!");
                if (UomId <= 0) throw new SystemException("【重量單位】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷物流箱資訊是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.PickingCarton
                                WHERE PickingCartonId = @PickingCartonId";
                        dynamicParameters.Add("PickingCartonId", PickingCartonId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("物流箱資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.PickingCarton SET
                                TotalWeight = @TotalWeight,
                                UomId = @UomId,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE PickingCartonId = @PickingCartonId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                TotalWeight,
                                UomId,
                                LastModifiedDate,
                                LastModifiedBy,
                                PickingCartonId
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

        #region //UpdateDoCarton -- 出貨單物流箱更新(綁定物流箱) -- Zoey 2022.11.16
        public string UpdateDoCarton(string CartonBarcode, int DoId)
        {
            try
            {
                if (CartonBarcode.Length <= 0) throw new SystemException("【物流箱條碼】不能為空!");
                if (DoId <= 0) throw new SystemException("查無此出貨單！");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        int pickingCartonId = -1;
                        int pickResultRegular = -1;
                        int pickResultFreebie = -1;
                        int pickResultSpare = -1;

                        #region //判斷出貨單資訊是否正確
                        sql = @"SELECT TOP 1 1
                                FROM SCM.DeliveryOrder
                                WHERE DoId = @DoId
                                AND TransferStatus = @TransferStatus
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("DoId", DoId);
                        dynamicParameters.Add("TransferStatus", "N");
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("出貨單資料錯誤!");
                        #endregion
                        
                        #region //判斷物流箱條碼資訊是否正確/撈物流箱Id
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.PickingCartonId
                                FROM SCM.PickingCarton a
                                INNER JOIN SCM.PickingItem b ON b.PickingCartonId = a.PickingCartonId
                                WHERE CartonBarcode = @CartonBarcode
                                AND b.DoId IS NULL";
                        dynamicParameters.Add("CartonBarcode", CartonBarcode);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() <= 0) throw new SystemException("物流箱條碼錯誤或該物流箱已綁定出貨單!");

                        foreach (var item in result2)
                        {
                            pickingCartonId = item.PickingCartonId;
                        }
                        #endregion

                        #region //判斷物流箱內物件是否符合該出貨單
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.SoDetailId 
                                FROM SCM.PickingItem a
                                WHERE a.PickingCartonId = @PickingCartonId
                                EXCEPT
                                SELECT a.SoDetailId 
                                FROM SCM.DoDetail a
                                INNER JOIN SCM.DeliveryOrder b ON a.DoId = b.DoId
                                WHERE a.DoId = @DoId
                                AND b.CompanyId = @CompanyId";
                        dynamicParameters.Add("PickingCartonId", pickingCartonId);
                        dynamicParameters.Add("DoId", DoId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result3 = sqlConnection.Query(sql, dynamicParameters);
                        if (result3.Count() > 0) throw new SystemException("箱內物件不符合該出貨單!");
                        #endregion

                        #region //判斷物流箱內物件數量是否符合該出貨單
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT SUM(CanNotPickRegularQty) CanNotPickRegularQty, SUM(CanNotPickFreebieQty) CanNotPickFreebieQty, SUM(CanNotPickSpareQty) CanNotPickSpareQty
                                FROM
                                (
                                    SELECT CASE WHEN ((SUM(a.DoQty) - SUM(b.RegularQty)) - e.PickRegularQty) < 0 THEN 1 ELSE 0 END CanNotPickRegularQty
                                    , CASE WHEN ((SUM(a.FreebieQty) - SUM(c.FreebieQty)) - f.PickFreebieQty) < 0 THEN 1 ELSE 0 END CanNotPickFreebieQty
                                    , CASE WHEN ((SUM(a.SpareQty) - SUM(d.SpareQty)) - g.PickSpareQty) < 0 THEN 1 ELSE 0 END CanNotPickSpareQty
                                    FROM SCM.DoDetail a
                                    OUTER APPLY (
                                        SELECT ISNULL(SUM(b1.ItemQty), 0) RegularQty
                                        FROM SCM.PickingItem b1
                                        WHERE b1.DoId = a.DoId
                                        AND b1.ItemType = 1
                                        AND b1.SoDetailId = a.SoDetailId
                                    ) b
                                    OUTER APPLY (
                                        SELECT ISNULL(SUM(b2.ItemQty), 0) FreebieQty
                                        FROM SCM.PickingItem b2
                                        WHERE b2.DoId = a.DoId
                                        AND b2.ItemType = 2
                                        AND b2.SoDetailId = a.SoDetailId
                                    ) c
                                    OUTER APPLY (
                                        SELECT ISNULL(SUM(b3.ItemQty), 0) SpareQty
                                        FROM SCM.PickingItem b3
                                        WHERE b3.DoId = a.DoId
                                        AND b3.ItemType = 3
                                        AND b3.SoDetailId = a.SoDetailId
                                    ) d
                                    OUTER APPLY (
                                        SELECT ISNULL(SUM(x.ItemQty), 0) PickRegularQty
                                        FROM SCM.PickingItem x
                                        WHERE x.PickingCartonId = @PickingCartonId
                                        AND x.SoDetailId = a.SoDetailId
                                        AND x.ItemType = 1
                                    ) e
                                    OUTER APPLY (
                                        SELECT ISNULL(SUM(y.ItemQty), 0) PickFreebieQty
                                        FROM SCM.PickingItem y
                                        WHERE y.PickingCartonId = @PickingCartonId
                                        AND y.SoDetailId = a.SoDetailId
                                        AND y.ItemType = 2
                                    ) f
                                    OUTER APPLY (
                                        SELECT ISNULL(SUM(z.ItemQty), 0) PickSpareQty
                                        FROM SCM.PickingItem z
                                        WHERE z.PickingCartonId = @PickingCartonId
                                        AND z.SoDetailId = a.SoDetailId
                                        AND z.ItemType = 3
                                    ) g
                                    INNER JOIN SCM.DeliveryOrder i ON a.DoId = i.DoId
                                    WHERE a.DoId = @DoId
                                    AND i.CompanyId = @CompanyId
                                    GROUP BY a.SoDetailId, e.PickRegularQty, f.PickFreebieQty, g.PickSpareQty
                                ) result";
                        dynamicParameters.Add("PickingCartonId", pickingCartonId);
                        dynamicParameters.Add("DoId", DoId);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var result4 = sqlConnection.Query(sql, dynamicParameters);
                        foreach (var item in result4)
                        {
                            pickResultRegular = item.CanNotPickRegularQty;
                            pickResultFreebie = item.CanNotPickFreebieQty;
                            pickResultSpare = item.CanNotPickSpareQty;
                        }
                        if (pickResultRegular > 0) throw new SystemException("箱內【正常品】物件數量大於該出貨單!");
                        if (pickResultFreebie > 0) throw new SystemException("箱內【贈品】物件數量大於該出貨單!");
                        if (pickResultSpare > 0) throw new SystemException("箱內【備品】物件數量大於該出貨單!");
                        #endregion

                        int rowsAffected = 0;
                        #region //修改物流箱的出貨單
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.PickingItem SET
                                DoId = @DoId,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE PickingCartonId = @PickingCartonId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                DoId,
                                LastModifiedDate,
                                LastModifiedBy,
                                PickingCartonId = pickingCartonId
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //更新出貨單狀態
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.DeliveryOrder SET
                                Status = @Status,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE DoId = @DoId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                Status = "P",
                                LastModifiedDate,
                                LastModifiedBy,
                                DoId
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
        #region //DeletePickingCarton -- 物流箱資料刪除 -- Zoey 2022.10.12
        public string DeletePickingCarton(string PickingCartonIds)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        var carton = PickingCartonIds.Split(',');

                        for (int i = 0; i < carton.Length; i++)
                        {
                            #region //判斷物流箱是否正確
                            sql = @"SELECT TOP 1 1
                                    FROM SCM.PickingCarton a
                                    INNER JOIN SCM.Packing b ON a.PackingId = b.PackingId
                                    WHERE a.PickingCartonId = @PickingCartonId
                                    AND b.CompanyId = @CompanyId";
                            dynamicParameters.Add("PickingCartonId", carton[i]);
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("物流箱資料錯誤!");
                            #endregion

                            #region //判斷物流箱是否為空箱
                            sql = @"SELECT TOP 1 1
                                    FROM SCM.PickingItem
                                    WHERE PickingCartonId = @PickingCartonId";
                            dynamicParameters.Add("PickingCartonId", carton[i]);

                            var result2 = sqlConnection.Query(sql, dynamicParameters);
                            if (result2.Count() > 0) throw new SystemException("物流箱內尚有物件，無法刪除!!");
                            #endregion
                        }

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.PickingCarton
                                WHERE PickingCartonId IN @PickingCartonId";
                        dynamicParameters.Add("PickingCartonId", carton);

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

        #region //DeletePickingItem -- 揀貨物件資料刪除 -- Zoey 2022.10.24
        public string DeletePickingItem(string PickingItemIds)
        {
            try
            {
                int doId = -1;
                List<int> soDetails = new List<int>();

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        var pickingItem = PickingItemIds.Split(',');

                        var inventoryHold = new InventoryTempHold();

                        for (int i = 0; i < pickingItem.Length; i++)
                        {
                            #region //判斷物件條碼是否正確
                            sql = @"SELECT TOP 1 SoDetailId, DoId, BarcodeId, ItemStatus
                                    FROM SCM.PickingItem
                                    WHERE PickingItemId = @PickingItemId";
                            dynamicParameters.Add("PickingItemId", pickingItem[i]);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("物件條碼錯誤!");
                            #endregion

                            #region //釋放暫存 2024.11.22.新增庫存卡控
                            if (CurrentCompany != 4)
                            {
                                foreach (var item in result)
                                {
                                    inventoryHold.ReleaseByDoDetailId(sqlConnection, Convert.ToInt32(item.DoId), Convert.ToInt32(item.SoDetailId), item.ItemStatus, LastModifiedBy);
                                }
                            }
                            #endregion
                        }

                        int rowsAffected = 0;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.PickingLot
                                WHERE PickingItemId IN @PickingItemId";
                        dynamicParameters.Add("PickingItemId", pickingItem);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.PickingItem
                                WHERE PickingItemId IN @PickingItemId";
                        dynamicParameters.Add("PickingItemId", pickingItem);

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

        #region //DeletePickingLot -- 揀貨物件資料刪除 -- Shintokuro 2024.12.02
        public string DeletePickingLot(int PickingLotId)
        {
            try
            {
                int doId = -1;
                int rowsAffected = 0;

                List<int> soDetails = new List<int>();

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {

                        var inventoryHold = new InventoryTempHold();
                        int PickingItemId = -1;
                        int ItemQty = -1;
                        int num = -1;
                        int SoDetailId = -1;
                        int DoId = -1;
                        string ItemStatus = "";
                        #region //判斷物件條碼是否正確
                        sql = @"SELECT a.ItemQty,x.num,a.PickingItemId,b.SoDetailId, b.DoId,b.ItemStatus
                                FROM SCM.PickingLot a
                                INNER JOIN SCM.PickingItem b on a.PickingItemId = b.PickingItemId
                                OUTER APPLY(
                                    SELECT COUNT(*) num
                                    FROM SCM.PickingLot x1
                                    WHERE x1.PickingItemId = a.PickingItemId
                                ) x
                                WHERE a.PickingLotId = @PickingLotId";
                        dynamicParameters.Add("PickingLotId", PickingLotId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("物件條碼錯誤!");
                        foreach (var item in result)
                        {
                            PickingItemId = item.PickingItemId;
                            ItemQty = item.ItemQty;
                            num = item.num;
                            SoDetailId = item.SoDetailId;
                            DoId = item.DoId;
                            ItemStatus = item.ItemStatus;
                        }
                        #endregion


                        #region //修改撿貨無條碼資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.PickingItem SET
                                ItemQty = ItemQty - @ItemQty,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE PickingItemId = @PickingItemId
                                ";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ItemQty,
                                LastModifiedDate,
                                LastModifiedBy,
                                PickingItemId
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //修改撿貨暫扣資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.InventoryTempHold SET
                                HoldQty = HoldQty - @ItemQty,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE DoId = @DoId
                                AND SoDetailId = @SoDetailId
                                AND InputStatus = @InputStatus
                                ";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ItemQty,
                                LastModifiedDate,
                                LastModifiedBy,
                                DoId,
                                SoDetailId,
                                InputStatus = ItemStatus
                            });
                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion




                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.PickingLot
                                WHERE PickingLotId = @PickingLotId";
                        dynamicParameters.Add("PickingLotId", PickingLotId);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                        if (num == 1)
                        {
                            #region //釋放暫存 2024.11.22.新增庫存卡控
                            if (CurrentCompany != 4)
                            {
                                inventoryHold.ReleaseByDoDetailId(sqlConnection, Convert.ToInt32(DoId), Convert.ToInt32(SoDetailId), ItemStatus, LastModifiedBy);

                            }
                            #endregion

                            //最後一個刪除,也刪除單身
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE SCM.PickingItem
                                WHERE PickingItemId = @PickingItemId";
                            dynamicParameters.Add("PickingItemId", PickingItemId);

                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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
        #endregion
    }

    public class InventoryTempHold //2024.11.22.新增庫存卡控
    {
        // 庫存暫扣檢查
        public bool CheckInventoryAvailable(SqlConnection erpConn, SqlConnection mainConn, string mtlItemNo, decimal quantity)
        {
            // 查詢A倉實際庫存總和
            var sql = @"
            SELECT ISNULL(SUM(x.MC007), 0) AS TotalInventoryQty
            FROM INVMC x
            INNER JOIN CMSMC y ON x.MC002 = y.MC001
            WHERE LTRIM(RTRIM(x.MC001)) = @MtlItemNo
            AND LTRIM(RTRIM(y.MC001)) LIKE 'A%'";

            var actualInventory = erpConn.QueryFirstOrDefault<decimal>(sql,
                new { MtlItemNo = mtlItemNo });

            // 查詢暫扣數量
            sql = @"
            SELECT ISNULL(SUM(HoldQty), 0)
            FROM SCM.InventoryTempHold
            WHERE MtlItemNo = @MtlItemNo 
            AND IsReleased = 0";

            var tempHoldQty = mainConn.QueryFirstOrDefault<decimal>(sql,
                new { MtlItemNo = mtlItemNo });

            return (actualInventory - tempHoldQty) >= quantity;
        }

        // 檢查整張出貨單的庫存是否足夠
        public bool CheckDoInventoryAvailable(SqlConnection erpConn, SqlConnection mainConn, int doId)
        {
            // 1. 取得出貨單所有品項及數量
            var sql = @"
            SELECT MtlItemNo, InputStatus,
                   SUM(HoldQty) as TotalQty
            FROM SCM.InventoryTempHold
            WHERE DoId = @DoId
            AND InputStatus IN ('B', 'N')
            AND IsReleased = 0
            GROUP BY MtlItemNo, InputStatus";

            var holdItems = mainConn.Query<DoHoldItem>(sql, new { DoId = doId }).ToList();

            // 如果沒有未釋放的暫扣，直接返回true
            if (!holdItems.Any())
            {
                return true;
            }

            // 2. 檢查每個品項的庫存
            foreach (var item in holdItems)
            {
                // 查詢A倉實際庫存總和
                sql = @"
                SELECT ISNULL(SUM(x.MC007), 0) AS TotalInventoryQty
                FROM INVMC x
                INNER JOIN CMSMC y ON x.MC002 = y.MC001
                WHERE LTRIM(RTRIM(x.MC001)) = @MtlItemNo
                AND LTRIM(RTRIM(y.MC001)) LIKE 'A%'";

                var actualInventory = erpConn.QueryFirstOrDefault<decimal>(sql,
                    new { MtlItemNo = item.MtlItemNo });

                // 查詢其他出貨單的暫扣數量
                sql = @"
                SELECT ISNULL(SUM(HoldQty), 0)
                FROM SCM.InventoryTempHold
                WHERE MtlItemNo = @MtlItemNo 
                AND DoId != @DoId
                AND IsReleased = 0
                AND InputStatus IN ('B', 'N')";

                var otherHoldQty = mainConn.QueryFirstOrDefault<decimal>(sql,
                    new { MtlItemNo = item.MtlItemNo, DoId = doId });

                // 檢查可用庫存是否足夠
                if ((actualInventory - otherHoldQty) < item.TotalQty)
                {
                    return false;
                }
            }

            return true;
        }

        // 建立或更新無條碼庫存暫扣
        public int CreateNoBarcodeHold(SqlConnection conn, string mtlItemNo, int noBarcodeQty, int soDetailId, int doId, int userId)
        {
            // 檢查是否已存在未釋放的無條碼暫扣記錄
            var sql = @"
            SELECT TempHoldId, HoldQty
            FROM SCM.InventoryTempHold
            WHERE DoId = @DoId
            AND SoDetailId = @SoDetailId
            AND MtlItemNo = @MtlItemNo
            AND InputStatus = 'N'
            AND IsReleased = 0";

            var existingHold = conn.QueryFirstOrDefault<TempHoldRecord>(sql, new
            {
                DoId = doId,
                SoDetailId = soDetailId,
                MtlItemNo = mtlItemNo
            });

            var now = DateTime.Now;
            int tempHoldId;

            if (existingHold != null)
            {
                // 更新既有記錄
                sql = @"
                UPDATE SCM.InventoryTempHold
                SET HoldQty = HoldQty + @AddQty,
                    LastModifiedDate = @LastModifiedDate,
                    LastModifiedBy = @LastModifiedBy
                OUTPUT INSERTED.TempHoldId
                WHERE TempHoldId = @TempHoldId";

                tempHoldId = conn.QuerySingle<int>(sql, new
                {
                    AddQty = noBarcodeQty,
                    LastModifiedDate = now,
                    LastModifiedBy = userId,
                    TempHoldId = existingHold.TempHoldId
                });
            }
            else
            {
                // 建立新記錄
                sql = @"
                INSERT INTO SCM.InventoryTempHold (
                    MtlItemNo, HoldQty, SoDetailId, DoId,
                    CreateDate, CreateBy, LastModifiedDate, LastModifiedBy,
                    InputStatus, IsReleased
                )
                OUTPUT INSERTED.TempHoldId
                VALUES (
                    @MtlItemNo, @HoldQty, @SoDetailId, @DoId,
                    @CreateDate, @CreateBy, @LastModifiedDate, @LastModifiedBy,
                    'N', 0
                )";

                tempHoldId = conn.QuerySingle<int>(sql, new
                {
                    MtlItemNo = mtlItemNo,
                    HoldQty = noBarcodeQty,
                    SoDetailId = soDetailId,
                    DoId = doId,
                    CreateDate = now,
                    CreateBy = userId,
                    LastModifiedDate = now,
                    LastModifiedBy = userId
                });
            }

            return tempHoldId;
        }

        // 建立條碼庫存暫扣
        public int CreateBarcodeHold(SqlConnection conn, string mtlItemNo, string barcode, int barcodeQty, int soDetailId, int doId, int userId)
        {
            var now = DateTime.Now;
            var sql = @"
            INSERT INTO SCM.InventoryTempHold (
                MtlItemNo, HoldQty, SoDetailId, DoId, BarCodeNo,
                CreateDate, CreateBy, LastModifiedDate, LastModifiedBy,
                InputStatus, IsReleased
            )
            OUTPUT INSERTED.TempHoldId
            VALUES (
                @MtlItemNo, @HoldQty, @SoDetailId, @DoId, @BarCodeNo,
                @CreateDate, @CreateBy, @LastModifiedDate, @LastModifiedBy,
                'B', 0
            )";

            return conn.QuerySingle<int>(sql, new
            {
                MtlItemNo = mtlItemNo,
                HoldQty = barcodeQty,
                SoDetailId = soDetailId,
                DoId = doId,
                BarCodeNo = barcode,
                CreateDate = now,
                CreateBy = userId,
                LastModifiedDate = now,
                LastModifiedBy = userId
            });
        }

        // 恢復已釋放的暫扣（用於反確認）
        public void RestoreReleasedHolds(SqlConnection conn, int doId, int userId)
        {
            var now = DateTime.Now;

            // 1. 查詢需要恢復的記錄
            var sql = @"
                WITH LatestHolds AS (
                    SELECT TempHoldId,
                           ROW_NUMBER() OVER (
                               PARTITION BY MtlItemNo, BarCodeNo, InputStatus, SoDetailId 
                               ORDER BY CreateDate DESC
                           ) AS RowNum
                    FROM SCM.InventoryTempHold
                    WHERE DoId = @DoId
                    AND IsReleased = 1
                )
                SELECT TempHoldId
                FROM LatestHolds
                WHERE RowNum = 1";

            var holdIds = conn.Query<int>(sql, new { DoId = doId }).ToList();

            if (holdIds.Any())
            {
                // 2. 恢復最新的已釋放記錄
                sql = @"
                    UPDATE SCM.InventoryTempHold
                    SET IsReleased = 0,
                        ReleaseDate = NULL,
                        LastModifiedDate = @LastModifiedDate,
                        LastModifiedBy = @LastModifiedBy
                    WHERE TempHoldId IN @HoldIds";

                conn.Execute(sql, new
                {
                    HoldIds = holdIds,
                    LastModifiedDate = now,
                    LastModifiedBy = userId
                });

                // 3. 刪除其他舊記錄
                sql = @"
                    DELETE FROM SCM.InventoryTempHold
                    WHERE DoId = @DoId
                    AND TempHoldId NOT IN @HoldIds";

                conn.Execute(sql, new
                {
                    DoId = doId,
                    HoldIds = holdIds
                });
            }
        }

        // 釋放指定出貨單的所有暫扣
        public void ReleaseByDoId(SqlConnection conn, int doId, int userId)
        {
            var sql = @"
            UPDATE SCM.InventoryTempHold
            SET IsReleased = 1,
                ReleaseDate = @ReleaseDate,
                LastModifiedDate = @LastModifiedDate,
                LastModifiedBy = @LastModifiedBy
            WHERE DoId = @DoId
            AND IsReleased = 0";

            var now = DateTime.Now;
            conn.Execute(sql, new
            {
                DoId = doId,
                ReleaseDate = now,
                LastModifiedDate = now,
                LastModifiedBy = userId
            });
        }

        // 釋放特定出貨單特定訂單物件的暫扣
        public void ReleaseByDoDetailId(SqlConnection conn, int doId, int soDetailId, string inputStatus, int userId)
        {
            var sql = @"
            UPDATE SCM.InventoryTempHold
            SET IsReleased = 1,
                ReleaseDate = @ReleaseDate,
                LastModifiedDate = @LastModifiedDate,
                LastModifiedBy = @LastModifiedBy
            WHERE DoId = @DoId
            AND SoDetailId = @SoDetailId
            AND InputStatus = @InputStatus
            AND IsReleased = 0";

            var now = DateTime.Now;
            conn.Execute(sql, new
            {
                DoId = doId,
                SoDetailId = soDetailId,
                InputStatus = inputStatus,
                ReleaseDate = now,
                LastModifiedDate = now,
                LastModifiedBy = userId
            });
        }

        // 檢查條碼是否已被暫扣
        public bool IsBarcodeHeld(SqlConnection conn, string barcode, int soDetailId)
        {
            var sql = @"
            SELECT COUNT(1)
            FROM SCM.InventoryTempHold
            WHERE BarCodeNo = @BarCodeNo
            AND SoDetailId = @SoDetailId
            AND IsReleased = 0";

            var count = conn.QueryFirstOrDefault<int>(sql, new
            {
                BarCodeNo = barcode,
                SoDetailId = soDetailId
            });
            return count > 0;
        }

        // 清理舊的暫扣記錄（保留最新一筆）
        public void CleanupOldHolds(SqlConnection conn, int doId)
        {
            using (var transaction = conn.BeginTransaction())
            {
                try
                {
                    // 1. 獲取最新一批的記錄ID
                    var sql = @"
                WITH LatestHolds AS (
                    SELECT TempHoldId, MtlItemNo, BarCodeNo, InputStatus, SoDetailId,
                           ROW_NUMBER() OVER (
                               PARTITION BY MtlItemNo, BarCodeNo, InputStatus, SoDetailId 
                               ORDER BY CreateDate DESC
                           ) AS RowNum
                    FROM SCM.InventoryTempHold
                    WHERE DoId = @DoId
                )
                SELECT TempHoldId
                FROM LatestHolds
                WHERE RowNum = 1";

                    var latestIds = conn.Query<int>(sql, new { DoId = doId }, transaction).ToList();

                    // 2. 刪除舊記錄
                    sql = @"
                DELETE FROM SCM.InventoryTempHold
                WHERE DoId = @DoId
                AND TempHoldId NOT IN @LatestIds";

                    conn.Execute(sql, new
                    {
                        DoId = doId,
                        LatestIds = latestIds
                    }, transaction);

                    transaction.Commit();
                }
                catch
                {
                    transaction.Rollback();
                    throw;
                }
            }
        }

        // 檢查是否可以恢復暫扣（檢查庫存是否足夠）
        public bool CanRestoreHolds(SqlConnection erpConn, SqlConnection mainConn, int doId)
        {
            // 1. 取得最新一批需要恢復的暫扣記錄
            var sql = @"
            WITH LatestHolds AS (
                SELECT TempHoldId, MtlItemNo, HoldQty, InputStatus,
                       ROW_NUMBER() OVER (
                           PARTITION BY MtlItemNo, BarCodeNo, InputStatus, SoDetailId 
                           ORDER BY CreateDate DESC
                       ) AS RowNum
                FROM SCM.InventoryTempHold
                WHERE DoId = @DoId
                AND IsReleased = 1
                AND DATEADD(DAY, 7, ReleaseDate) >= @CurrentDate
                AND InputStatus IN ('B', 'N')
            )
            SELECT MtlItemNo, InputStatus,
                   SUM(HoldQty) as TotalQty
            FROM LatestHolds
            WHERE RowNum = 1
            GROUP BY MtlItemNo, InputStatus";

            var holdItems = mainConn.Query<DoHoldItem>(sql, new
            {
                DoId = doId,
                CurrentDate = DateTime.Now
            }).ToList();

            // 2. 檢查每個品項的庫存
            foreach (var item in holdItems)
            {
                // 查詢A倉實際庫存總和
                sql = @"
                SELECT ISNULL(SUM(x.MC007), 0) AS TotalInventoryQty
                FROM INVMC x
                INNER JOIN CMSMC y ON x.MC002 = y.MC001
                WHERE LTRIM(RTRIM(x.MC001)) = @MtlItemNo
                AND LTRIM(RTRIM(y.MC001)) LIKE 'A%'";

                var actualInventory = erpConn.QueryFirstOrDefault<decimal>(sql,
                    new { MtlItemNo = item.MtlItemNo });

                // 查詢目前未釋放的暫扣數量
                sql = @"
                SELECT ISNULL(SUM(HoldQty), 0)
                FROM SCM.InventoryTempHold
                WHERE MtlItemNo = @MtlItemNo 
                AND IsReleased = 0
                AND InputStatus IN ('B', 'N')";

                var currentHoldQty = mainConn.QueryFirstOrDefault<decimal>(sql,
                    new { MtlItemNo = item.MtlItemNo });

                if ((actualInventory - currentHoldQty) < item.TotalQty)
                {
                    return false;
                }
            }

            return true;
        }
    }
}
