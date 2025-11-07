using Dapper;
using Helpers;
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
    public class InventoryTransactionDA
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

        public InventoryTransactionDA()
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
        #region //GetInventoryTransaction -- 取得庫存異動單據資料 -- Ann 2023-12-12
        public string GetInventoryTransaction(int ItId, string ItErpPrefix, string ItErpNo, string ItErpFullNo, string ItDate
            , string DocDate, string MtlItemNo, string MtlItemName, int InventoryId, int ToInventoryId, string ConfirmStatus
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //若沒有特別指定搜尋倉別，則確認此設備是否有綁定倉別
                    if (InventoryId <= 0)
                    {
                        string clientIp = BaseHelper.ClientIP();
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ISNULL(a.InventoryId, -1) InventoryId
                                FROM MES.Device a 
                                WHERE a.DeviceIdentifierCode = @DeviceIdentifierCode";
                        dynamicParameters.Add("DeviceIdentifierCode", clientIp);

                        var DeviceResult = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in DeviceResult)
                        {
                            InventoryId = item.InventoryId;
                        }
                    }
                    #endregion

                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.ItId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.ItErpPrefix, a.ItErpNo, FORMAT(a.ItDate, 'yyyy-MM-dd') ItDate, FORMAT(a.DocDate, 'yyyy-MM-dd') DocDate, a.DepartmentId
                        , a.Remark, a.TotalQty, a.Amount, a.ConfirmStatus, a.TransferStatus, FORMAT(a.TransferDate, 'yyyy-MM-dd') TransferDate, a.Remark
                        , FORMAT(a.CreateDate, 'yyyy-MM-dd') CreateDate, FORMAT(a.LastModifiedDate, 'yyyy-MM-dd') LastModifiedDate
                        , ISNULL(c.UserNo, '') ConfirmUserNo, ISNULL(c.UserName, '') ConfirmUserName
                        , ISNULL(b.DepartmentNo, '') DepartmentNo, ISNULL(b.DepartmentName, '') DepartmentName
                        , ISNULL(d.UserNo, '') CreateUserNo, ISNULL(d.UserName, '') CreateUserName";
                    sqlQuery.mainTables =
                        @"FROM SCM.InventoryTransaction a  
                        LEFT JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                        LEFT JOIN BAS.[User] c ON a.ConfirmUserId = c.UserId
                        LEFT JOIN BAS.[User] d ON a.CreateBy = d.UserId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ItId", @" AND a.ItId = @ItId", ItId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ItErpPrefix", @" AND a.ItErpPrefix LIKE '%' + @ItErpPrefix + '%'", ItErpPrefix);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ItErpNo", @" AND a.ItErpNo LIKE '%' + @ItErpNo + '%'", ItErpNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ItErpFullNo", @" AND (a.ItErpPrefix + '-' + a.ItErpNo) LIKE '%' + @ItErpFullNo + '%'", ItErpFullNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ItDate", @" AND a.ItDate = @ItDate", ItDate);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DocDate", @" AND a.DocDate = @DocDate", DocDate);
                    if (ConfirmStatus.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ConfirmStatus", @" AND a.ConfirmStatus IN @ConfirmStatus", ConfirmStatus.Split(','));
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemNo", @" AND EXISTS (
                                                                                                            SELECT TOP 1 1 FROM SCM.ItDetail x
                                                                                                            INNER JOIN PDM.MtlItem xa ON x.MtlItemId = xa.MtlItemId 
                                                                                                            WHERE x.ItId = a.ItId
                                                                                                            AND xa.MtlItemNo = @MtlItemNo
                                                                                                        )", MtlItemNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemName", @" AND EXISTS (
                                                                                                            SELECT TOP 1 1 FROM SCM.ItDetail x 
                                                                                                            WHERE x.ItId = a.ItId
                                                                                                            AND x.ItMtlItemName LIKE '%' + @MtlItemName + '%'
                                                                                                        )", MtlItemName);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "InventoryId", @" AND (EXISTS (
                                                                                                            SELECT TOP 1 1 FROM SCM.ItDetail x
                                                                                                            WHERE x.ItId = a.ItId
                                                                                                            AND x.InventoryId = @InventoryId
                                                                                                        ) OR NOT EXISTS (SELECT TOP 1 1 FROM SCM.ItDetail x WHERE x.ItId = a.ItId))", InventoryId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ToInventoryId", @" AND EXISTS (
                                                                                                            SELECT TOP 1 1 FROM SCM.ItDetail x
                                                                                                            WHERE x.ItId = a.ItId
                                                                                                            AND x.ToInventoryId = @ToInventoryId
                                                                                                        )", ToInventoryId);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.ItId DESC";
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

        #region //GetItDetail -- 取得庫存異動單據詳細資料 -- Ann 2023-12-12
        public string GetItDetail(int ItDetailId, int ItId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.ItDetailId, a.ItId, a.ItSequence, a.MtlItemId, a.ItMtlItemName, a.ItMtlItemSpec
                            , a.ItQty, a.InvQty, a.UnitCost, a.Amount, a.InventoryId, a.ToInventoryId
                            , ISNULL(a.StorageLocation, '') StorageLocation, ISNULL(a.ToStorageLocation, '') ToStorageLocation, a.ItRemark
                            , b.MtlItemNo, b.InventoryUomId UomId
                            , c.InventoryNo, c.InventoryName
                            , ISNULL(d.InventoryNo, '') ToInventoryNo, ISNULL(d.InventoryName, '') ToInventoryName
                            , e.ItErpPrefix, e.ItErpNo
                            , f.UomNo
                            FROM SCM.ItDetail a
                            INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId 
                            INNER JOIN SCM.Inventory c ON a.InventoryId = c.InventoryId
                            LEFT JOIN SCM.Inventory d ON a.ToInventoryId = d.InventoryId
                            INNER JOIN SCM.InventoryTransaction e ON a.ItId = e.ItId
                            INNER JOIN PDM.UnitOfMeasure f ON b.InventoryUomId = f.UomId
                            WHERE 1=1";

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ItDetailId", @" AND a.ItDetailId = @ItDetailId", ItDetailId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "ItId", @" AND a.ItId = @ItId", ItId);

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

        #region //GetMtlItem -- 取得品號資料 -- Ann 2023-12-11
        public string GetMtlItem(int MtlItemId, string MtlItemNo)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.MtlItemId, a.MtlItemNo, a.MtlItemName, a.MtlItemSpec
                            , b.UomId, b.UomNo, b.UomName
                            , c.InventoryId, c.InventoryNo, c.InventoryName
                            FROM PDM.MtlItem a 
                            INNER JOIN PDM.UnitOfMeasure b ON a.InventoryUomId = b.UomId
                            INNER JOIN SCM.Inventory c ON a.InventoryId = c.InventoryId
                            WHERE a.CompanyId = @CurrentCompany";
                    dynamicParameters.Add("CurrentCompany", CurrentCompany);

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemId", @" AND a.MtlItemId = @MtlItemId", MtlItemId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemNo", @" AND a.MtlItemNo = @MtlItemNo", MtlItemNo);

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

        #region //GetInventoryQty -- 取得品號資料 -- Ann 2023-12-11
        public string GetInventoryQty(int MtlItemId, int InventoryId)
        {
            if (MtlItemId <= 0) throw new SystemException("【品號】不能為空!");
            if (InventoryId <= 0) throw new SystemException("【庫別】不能為空!");

            string MtlItemNo = "";
            string InventoryNo = "";

            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
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
                    }
                    #endregion

                    using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                    {
                        #region //取得MES品號資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MtlItemNo
                                FROM PDM.MtlItem a 
                                WHERE a.MtlItemId = @MtlItemId";

                        dynamicParameters.Add("MtlItemId", MtlItemId);

                        var MtlItemResult = sqlConnection.Query(sql, dynamicParameters);

                        if (MtlItemResult.Count() <= 0) throw new SystemException("品號資料錯誤!!");

                        foreach (var item in MtlItemResult)
                        {
                            MtlItemNo = item.MtlItemNo;
                        }
                        #endregion

                        #region //取得MES庫別資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.InventoryNo
                                FROM SCM.Inventory a 
                                WHERE a.InventoryId = @InventoryId";

                        dynamicParameters.Add("InventoryId", InventoryId);

                        var InventoryResult = sqlConnection.Query(sql, dynamicParameters);

                        if (InventoryResult.Count() <= 0) throw new SystemException("品號資料錯誤!!");

                        foreach (var item in InventoryResult)
                        {
                            InventoryNo = item.InventoryNo;
                        }
                        #endregion

                        #region //取得ERP庫存資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ISNULL(a.MC007, 0) InventoryQty
                                FROM INVMC a
                                WHERE a.MC001 = @MtlItemNo
                                AND a.MC002 = @InventoryNo";
                        dynamicParameters.Add("MtlItemNo", MtlItemNo);
                        dynamicParameters.Add("InventoryNo", InventoryNo);

                        var INVMCResult = sqlConnection2.Query(sql, dynamicParameters);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            data = INVMCResult
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

        #region //GetMaterialRequisition -- 取得領退料單據資料 -- Ann 2023-12-19
        public string GetMaterialRequisition(int MrId, string MrErpPrefix, string MrErpNo, string MrErpFullNo, string MrDate
            , string DocDate, string MtlItemNo, string MtlItemName, string WoErpFullNo
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.MrId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.RequesitionNo, a.MrErpPrefix, a.MrErpNo, FORMAT(a.MrDate, 'yyyy-MM-dd') MrDate
                        , FORMAT(a.DocDate, 'yyyy-MM-dd') DocDate, a.ProductionLine, a.Remark, a.TransferStatus, a.ConfirmStatus
                        , FORMAT(a.CreateDate, 'yyyy-MM-dd') CreateDate, FORMAT(a.LastModifiedDate, 'yyyy-MM-dd') LastModifiedDate
                        , ISNULL(b.UserNo, '') ConfirmUserNo,  ISNULL(b.UserName, '') ConfirmUserName
                        , c.UserNo CreateUserNo, c.UserName CreateUserName";
                    sqlQuery.mainTables =
                        @"FROM MES.MaterialRequisition a 
                        LEFT JOIN BAS.[User] b ON a.ConfirmUserId = b.UserId
                        INNER JOIN BAS.[User] c ON a.CreateBy = c.UserId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MrId", @" AND a.MrId = @MrId", MrId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MrErpPrefix", @" AND a.MrErpPrefix LIKE '%' + @MrErpPrefix + '%'", MrErpPrefix);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MrErpNo", @" AND a.MrErpNo LIKE '%' + @MrErpNo + '%'", MrErpNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MrErpFullNo", @" AND (a.MrErpPrefix + '-' + a.MrErpNo) LIKE '%' + @MrErpFullNo + '%'", MrErpFullNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MrDate", @" AND a.MrDate = @MrDate", MrDate);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DocDate", @" AND a.DocDate = @DocDate", DocDate);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "WoErpFullNo", @" AND EXISTS (
                                                                                                            SELECT TOP 1 1 FROM MES.MrDetail x
                                                                                                            INNER JOIN MES.ManufactureOrder xa ON x.MoId = xa.MoId 
                                                                                                            WHERE x.MrId = a.MrId
                                                                                                            AND (x.WoErpPrefix + '-' + x.WoErpNo + '(' + CONVERT(VARCHAR(10), xa.WoSeq) + ')') LIKE '%' + @WoErpFullNo + '%'
                                                                                                        )", WoErpFullNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemNo", @" AND EXISTS (
                                                                                                            SELECT TOP 1 1 FROM MES.MrDetail x
                                                                                                            INNER JOIN PDM.MtlItem xa ON x.MtlItemId = xa.MtlItemId 
                                                                                                            WHERE x.MrId = a.MrId
                                                                                                            AND xa.MtlItemNo LIKE '%' + @MtlItemNo + '%'
                                                                                                        )", MtlItemNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemName", @" AND EXISTS (
                                                                                                            SELECT TOP 1 1 FROM MES.MrDetail x 
                                                                                                            INNER JOIN PDM.MtlItem xa ON x.MtlItemId = xa.MtlItemId
                                                                                                            WHERE x.MrId = a.MrId
                                                                                                            AND xa.MtlItemName LIKE '%' + @MtlItemName + '%'
                                                                                                        )", MtlItemName);

                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.DocDate DESC";
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

        #region //GetMrDetail -- 取得庫存異動單據詳細資料 -- Ann 2023-12-19
        public string GetMrDetail(int MrDetailId, int MrId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.MrDetailId, a.MrId, a.MrSequence, a.MtlItemId, a.Quantity, a.ActualQuantity, a.UomId, a.Unit
                            , a.InventoryId, a.MoId, a.WoErpPrefix, a.WoErpNo, a.DetailDesc, a.Remark, a.ConfirmStatus, a.ProjectCode
                            , a.SubstitutionMtlItemNo, ISNULL(a.StorageLocation, '') StorageLocation
                            , b.MtlItemNo, b.MtlItemName, b.MtlItemSpec
                            , c.InventoryNo, c.InventoryName
                            , d.WoSeq, d.Quantity MesQuantity
                            , e.UomNo, e.UomName
                            , f.MrErpPrefix, f.MrErpNo, f.DocType
                            , h.MtlItemNo MoMtlItemNo, h.MtlItemName MoMtlItemName, h.MtlItemSpec MoMtlItemSpec
                            , (
                                SELECT ISNULL(SUM(x.BarcodeQty), 0) TotalBarcodeQty
                                FROM MES.Barcode x
                                INNER JOIN MES.MrBarcodeRegister xa ON x.BarcodeNo = xa.BarcodeNo
                                WHERE xa.MrDetailId = a.MrDetailId
                            ) TotalBarcodeQty
                            FROM MES.MrDetail a 
                            INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
                            INNER JOIN SCM.Inventory c ON a.InventoryId = c.InventoryId
                            INNER JOIN MES.ManufactureOrder d ON a.MoId = d.MoId
                            INNER JOIN PDM.UnitOfMeasure e ON a.UomId = e.UomId
                            INNER JOIN MES.MaterialRequisition f ON a.MrId = f.MrId
                            INNER JOIN MES.WipOrder g ON d.WoId = g.WoId
                            INNER JOIN PDM.MtlItem h ON g.MtlItemId = h.MtlItemId
                            WHERE 1=1";

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MrDetailId", @" AND a.MrDetailId = @MrDetailId", MrDetailId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MrId", @" AND a.MrId = @MrId", MrId);

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

        #region //GetManufactureOrder -- 取得製令相關資料 -- Ann 2023-12-20
        public string GetManufactureOrder(string WoErpFullNo)
        {
            try
            {
                if (WoErpFullNo.Length <= 0) throw new SystemException("請至少輸入一筆製令進行查詢!!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.MoId, a.WoId, a.Quantity, a.WoSeq
                            , b.WoErpPrefix, b.WoErpNo
                            , c.MtlItemNo, c.MtlItemName, c.MtlItemSpec
                            FROM MES.ManufactureOrder a 
                            INNER JOIN MES.WipOrder b ON a.WoId = b.WoId
                            INNER JOIN PDM.MtlItem c ON b.MtlItemId = c.MtlItemId
                            WHERE b.CompanyId = @CompanyId";

                    if (WoErpFullNo.IndexOf("-") != -1)
                    {
                        if (WoErpFullNo.IndexOf("(") == -1)
                        {
                            throw new SystemException("請連同括號一起輸入，EX: 5101-20231220001(1)!!");
                        }

                        sql += @" AND b.WoErpPrefix + '-' + b.WoErpNo + '(' + CONVERT(VARCHAR(10), a.WoSeq) + ')' = @WoErpFullNo";
                    }
                    else
                    {
                        sql += @" AND a.MoId = @WoErpFullNo";
                    }
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    dynamicParameters.Add("WoErpFullNo", WoErpFullNo);

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

        #region //GetMrWipOrder -- 取得領料單製令綁定相關資料 -- Ann 2023-12-20
        public string GetMrWipOrder(int MrId, int MoId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.MrId, a.MoId, a.MrType, a.Quantity, a.InventoryId, a.RequisitionCode, a.NegativeStatus, a.Remark
                            , a.MaterialCategory, a.SubinventoryType, a.LineSeq, a.UomId, a.StorageLocation
                            , b.InputQty, b.Quantity MesQuantity, b.WoSeq
                            , c.WoErpPrefix, c.WoErpNo, c.PlanQty, c.RequisitionSetQty, c.StockInQty
                            , d.TypeName MrTypeNo
                            , e.InventoryNo, e.InventoryName
                            , f.TypeName RequisitionCodeNo
                            , g.TypeName MaterialCategoryNo
                            , h.TypeName SubinventoryTypeNo
                            , i.MtlItemNo, i.MtlItemName, i.MtlItemSpec
                            , j.DocType
                            , k1.UomNo, k1.UomName
                            FROM MES.MrWipOrder a
                            INNER JOIN MES.ManufactureOrder b ON a.MoId = b.MoId
                            INNER JOIN MES.WipOrder c ON c.WoId = b.WoId
                            INNER JOIN BAS.[Type] d ON a.MrType = d.TypeNo AND d.TypeSchema = 'MrWipOrder.MrType'
                            INNER JOIN SCM.Inventory e ON a.InventoryId = e.InventoryId
                            INNER JOIN BAS.[Type] f ON a.RequisitionCode = f.TypeNo AND f.TypeSchema = 'MrWipOrder.RequisitionCode'
                            INNER JOIN BAS.[Type] g ON a.MaterialCategory = g.TypeNo AND g.TypeSchema = 'MrWipOrder.MaterialCategory'
                            INNER JOIN BAS.[Type] h ON a.SubinventoryType = h.TypeNo AND h.TypeSchema = 'MrWipOrder.SubinventoryType'
                            INNER JOIN PDM.MtlItem i ON c.MtlItemId = i.MtlItemId
                            INNER JOIN MES.MaterialRequisition j ON a.MrId = j.MrId
                            INNER JOIN PDM.UnitOfMeasure k1 ON a.UomId = k1.UomId
                            WHERE c.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);

                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MrId", @" AND a.MrId = @MrId", MrId);
                    BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MoId", @" AND a.MoId = @MoId", MoId);

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

        #region //GetMrBarcodeRegister -- 取得領料單條碼註冊資料 -- Ann 2023-12-21
        public string GetMrBarcodeRegister(int BarcodeRegisterId, int MrDetailId, string BarcodeNo
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.BarcodeRegisterId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.MrDetailId, a.BarcodeNo, FORMAT(a.CreateDate, 'yyyy-MM-dd') CreateDate
                        , b.UserNo, b.UserName
                        , d.TrayBarcode
                        , e.TrayNo
                        , ISNULL(f.BarcodeQty, 1) BarcodeQty
                        , (
                            SELECT ISNULL(SUM(x.BarcodeQty), 0) TotalBarcodeQty
                            FROM MES.Barcode x
                            INNER JOIN MES.MrBarcodeRegister xa ON x.BarcodeNo = xa.BarcodeNo
                            WHERE xa.MrDetailId = a.MrDetailId
                        ) TotalBarcodeQty
                        , (
                            SELECT (xa.TypeName + ':' + x.ItemValue) ItemValue
                            FROM MES.BarcodeAttribute x
                            INNER JOIN BAS.[Type] xa ON x.ItemNo = xa.TypeNo AND xa.TypeSchema = 'RoutingProcessItem.ItemNo'
                            WHERE x.BarcodeId = f.BarcodeId
                            FOR JSON PATH, ROOT('data')
                        ) ItemValue";
                    sqlQuery.mainTables =
                        @"FROM MES.MrBarcodeRegister a
                        INNER JOIN BAS.[User] b ON a.CreateBy = b.UserId
                        INNER JOIN MES.MrDetail c ON a.MrDetailId = c.MrDetailId
                        INNER JOIN MES.MoSetting d ON c.MoId = d.MoId
                        LEFT JOIN MES.Tray e ON a.BarcodeNo = e.BarcodeNo
                        LEFT JOIN MES.Barcode f ON a.BarcodeNo = f.BarcodeNo";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND 1=1";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "BarcodeRegisterId", @" AND a.BarcodeRegisterId = @BarcodeRegisterId", BarcodeRegisterId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MrDetailId", @" AND a.MrDetailId = @MrDetailId", MrDetailId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "BarcodeNo", @" AND a.BarcodeNo LIKE '%' + @BarcodeNo + '%'", BarcodeNo);
                    sqlQuery.conditions = queryCondition;

                    if (OrderBy == "Create") OrderBy = "a.CreateDate DESC";
                    else if (OrderBy == "BarcodeNo") OrderBy = "a.BarcodeNo";
                    else OrderBy = "a.CreateDate DESC";

                    sqlQuery.orderBy = OrderBy;
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

        #region //GetMrBarcodeReRegister -- 取得領料單條碼退料資料 -- Ann 2023-12-21
        public string GetMrBarcodeReRegister(int BarcodeReRegisterId, int MrDetailId, string BarcodeNo
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.BarcodeReRegisterId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.MrDetailId, a.BarcodeNo, FORMAT(a.CreateDate, 'yyyy-MM-dd') CreateDate
                        , b.UserNo, b.UserName
                        , (
                            SELECT (xa.TypeName + ':' + x.ItemValue) ItemValue
                            FROM MES.BarcodeAttribute x
                            INNER JOIN BAS.[Type] xa ON x.ItemNo = xa.TypeNo AND xa.TypeSchema = 'RoutingProcessItem.ItemNo'
                            WHERE x.BarcodeId = c.BarcodeId
                            FOR JSON PATH, ROOT('data')
                        ) ItemValue";
                    sqlQuery.mainTables =
                        @"FROM MES.MrBarcodeReRegister a
                        INNER JOIN BAS.[User] b ON a.CreateBy = b.UserId
                        INNER JOIN MES.Barcode c ON a.BarcodeNo = c.BarcodeNo";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND 1=1";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "BarcodeReRegisterId", @" AND a.BarcodeReRegisterId = @BarcodeReRegisterId", BarcodeReRegisterId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MrDetailId", @" AND a.MrDetailId = @MrDetailId", MrDetailId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "BarcodeNo", @" AND a.BarcodeNo LIKE '%' + @BarcodeNo + '%'", BarcodeNo);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.BarcodeReRegisterId";
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

        #region //GetStorageLocation -- 取得庫別儲位相關資料 -- Ann 2023-12-25
        public string GetStorageLocation(int InventoryId, string StorageLocation, string MtlItemNo)
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
                        #region //確認庫別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.InventoryNo
                                FROM SCM.Inventory a 
                                WHERE a.CompanyId = @CompanyId
                                AND a.InventoryId = @InventoryId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("InventoryId", InventoryId);

                        var InventoryResult = sqlConnection.Query(sql, dynamicParameters);

                        if (InventoryResult.Count() <= 0) throw new SystemException("庫別資料錯誤!!");

                        string InventoryNo = "";
                        foreach (var item in InventoryResult)
                        {
                            InventoryNo = item.InventoryNo;
                        }
                        #endregion

                        #region //取得ERP庫別儲位資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT RTRIM(LTRIM(a.NL001)) InventoryNo, RTRIM(LTRIM(a.NL002)) StorageLocation
                                , RTRIM(LTRIM(a.NL003)) StorageLocationName, RTRIM(LTRIM(a.NL004)) Remark
                                , RTRIM(LTRIM(b.MC002)) InventoryName
                                , c.MM005 StorageLocationQty
                                FROM CMSNL a 
                                INNER JOIN CMSMC b ON a.NL001 = b.MC001
                                LEFT JOIN INVMM c ON a.NL002 = c.MM003
                                WHERE 1=1";
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "InventoryNo", @" AND a.NL001 = @InventoryNo", InventoryNo);
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "StorageLocation", @" AND a.NL002 = @StorageLocation", StorageLocation);
                        BaseHelper.SqlParameter(ref sql, ref dynamicParameters, "MtlItemNo", @" AND c.MM001 = @MtlItemNo", MtlItemNo);

                        var result = sqlConnection2.Query(sql, dynamicParameters);

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            data = result
                        });
                        #endregion
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

        #region //GetInventoryLocationFlag -- 取得庫別儲位設定 -- Ann 2023-12-26
        public string GetInventoryLocationFlag(int InventoryId)
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
                        #region //確認庫別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.InventoryNo
                                FROM SCM.Inventory a 
                                WHERE a.CompanyId = @CompanyId
                                AND a.InventoryId = @InventoryId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("InventoryId", InventoryId);

                        var InventoryResult = sqlConnection.Query(sql, dynamicParameters);

                        if (InventoryResult.Count() <= 0) throw new SystemException("庫別資料錯誤!!");

                        string InventoryNo = "";
                        foreach (var item in InventoryResult)
                        {
                            InventoryNo = item.InventoryNo;
                        }
                        #endregion

                        #region //取得ERP庫別儲位資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT LTRIM(RTRIM(MC009)) LocationFlag
                                FROM CMSMC
                                WHERE MC001 = @MC001";
                        dynamicParameters.Add("MC001", InventoryNo);

                        var result = sqlConnection2.Query(sql, dynamicParameters);

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            data = result
                        });
                        #endregion
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

        #region //GetLocationQty -- 取得庫別儲位庫存數 -- Ann 2024-01-19
        public string GetLocationQty(int MtlItemId, int InventoryId, string Location)
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
                        #region //確認品號資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MtlItemNo
                                FROM PDM.MtlItem a 
                                WHERE a.MtlItemId = @MtlItemId";
                        dynamicParameters.Add("MtlItemId", MtlItemId);

                        var MtlItemResult = sqlConnection.Query(sql, dynamicParameters);

                        string MtlItemNo = "";
                        foreach (var item in MtlItemResult)
                        {
                            MtlItemNo = item.MtlItemNo;
                        }
                        #endregion

                        #region //確認庫別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.InventoryNo
                                FROM SCM.Inventory a 
                                WHERE a.CompanyId = @CompanyId
                                AND a.InventoryId = @InventoryId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);
                        dynamicParameters.Add("InventoryId", InventoryId);

                        var InventoryResult = sqlConnection.Query(sql, dynamicParameters);

                        if (InventoryResult.Count() <= 0) throw new SystemException("庫別資料錯誤!!");

                        string InventoryNo = "";
                        foreach (var item in InventoryResult)
                        {
                            InventoryNo = item.InventoryNo;
                        }
                        #endregion

                        #region //取得ERP庫別儲位庫存數
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT MM005 LocationQty
                                FROM INVMM
                                WHERe MM001 = @MtlItemNo
                                AND MM002 = @InventoryNo
                                AND MM003 = @Location";
                        dynamicParameters.Add("MtlItemNo", MtlItemNo);
                        dynamicParameters.Add("InventoryNo", InventoryNo);
                        dynamicParameters.Add("Location", Location);

                        var INVMMResult = sqlConnection2.Query(sql, dynamicParameters);

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            data = INVMMResult
                        });
                        #endregion
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

        #region //GetAuthority -- 取得庫存異動相關權限資料 -- Ann 2024-01-30
        public string GetAuthority()
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.DetailCode
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
                            INNER JOIN BAS.[Module] e ON b.ModuleId = e.ModuleId
                            WHERE a.[Status] = 'A'
                            AND b.[Status] = 'A'
                            AND b.FunctionCode = 'InventoryTransaction'
                            AND e.ModuleCode = 'ScmPlatform'
                            AND c.Authority > 0
                            ORDER BY a.SortNumber";
                    dynamicParameters.Add("UserId", CurrentUser);
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

        #region //GetDeviceInventory -- 取得裝置庫別資料 -- Ann 2024-02-22
        public string GetDeviceInventory(string MtlItemNo)
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
                        #region //取得MES裝置庫別資料
                        string clientIp = BaseHelper.ClientIP();
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.InventoryId
                                , c.InventoryNo, c.InventoryName
                                , c.InventoryNo + ' ' + c.InventoryName InventoryWithNo
                                FROM MES.DeciveInventory a 
                                INNER JOIN MES.Device b ON a.DeviceId = b.DeviceId
                                INNER JOIN SCM.Inventory c ON a.InventoryId = c.InventoryId
                                WHERE b.DeviceIdentifierCode = @DeviceIdentifierCode
                                AND b.CompanyId = @CompanyId
                                ORDER BY a.SortNumber";
                        dynamicParameters.Add("DeviceIdentifierCode", clientIp);
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        List<DeciveInventory> deciveInventories = sqlConnection.Query<DeciveInventory>(sql, dynamicParameters).ToList();
                        #endregion

                        #region //取得品號庫存資料
                        foreach (var item in deciveInventories)
                        {
                            deciveInventories[deciveInventories.IndexOf(item)].InventoryQty = 0;

                            if (MtlItemNo.Length > 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @" SELECT ISNULL(MC007, 0) InventoryQty
                                         FROM INVMC
                                         WHERE MC001 = @MtlItemNo
                                         AND MC002 = @InventoryNo";
                                dynamicParameters.Add("MtlItemNo", MtlItemNo);
                                dynamicParameters.Add("InventoryNo", item.InventoryNo);

                                var INVMCResult = sqlConnection2.Query(sql, dynamicParameters);

                                foreach (var item2 in INVMCResult)
                                {
                                    deciveInventories[deciveInventories.IndexOf(item)].InventoryQty = Convert.ToDouble(item2.InventoryQty);
                                }
                            }
                        }
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            data = deciveInventories
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

        #region //GetLoginUserInfo -- 取得使用者資料 -- Ann 2024-03-15
        public string GetLoginUserInfo()
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT UserNo, UserName
                            FROM BAS.[User]
                            WHERE UserId = @UserId";
                    dynamicParameters.Add("UserId", CurrentUser);

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
        #region //AddInventoryTransaction -- 新增庫存異動單資料 -- Ann 2023-12-15
        public string AddInventoryTransaction(string ItErpPrefix, string ItErpNo, string ItDate, string DocDate, int DepartmentId, string Remark)
        {
            try
            {
                int UserId = CreateBy;
                if (ItErpPrefix.Length <= 0) throw new SystemException("【單別】不能為空!");
                if (ItDate.Length <= 0) throw new SystemException("【異動日期】不能為空!");
                if (DocDate.Length <= 0) throw new SystemException("【單據日期】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
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
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //比對ERP關帳日期
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MA013)) MA013
                                    FROM CMSMA";
                            var cmsmaResult = sqlConnection2.Query(sql, dynamicParameters);
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

                            #region //單號未拋轉前先取亂碼
                            ItErpNo = BaseHelper.RandomCode(11);
                            #endregion

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO SCM.InventoryTransaction (CompanyId, ItErpPrefix, ItErpNo, ItDate, DocDate, DepartmentId
                                    , Remark, TotalQty, Amount, ConfirmStatus, ConfirmUserId, TransferStatus, TransferDate
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.ItId, INSERTED.ItErpNo, INSERTED.CompanyId, INSERTED.CreateBy
                                    VALUES (@CompanyId, @ItErpPrefix, @ItErpNo, @ItDate, @DocDate, @DepartmentId
                                    , @Remark, @TotalQty, @Amount, @ConfirmStatus, @ConfirmUserId, @TransferStatus, @TransferDate
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    CompanyId = CurrentCompany,
                                    ItErpPrefix,
                                    ItErpNo,
                                    ItDate,
                                    DocDate,
                                    DepartmentId,
                                    Remark,
                                    TotalQty = 0,
                                    Amount = 0,
                                    ConfirmStatus = "N",
                                    ConfirmUserId = (int?)null,
                                    TransferStatus = "N",
                                    TransferDate = (DateTime?)null,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy = UserId,
                                    LastModifiedBy = UserId
                                });

                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            int rowsAffected = insertResult.Count();

                            foreach (var item in insertResult)
                            {
                                if (BaseHelper.CheckUserAuthority(item.CreateBy, item.CompanyId, "A", "InventoryTransaction", "add", sqlConnection).Equals("N")) throw new SystemException("人員權限檢核有異常，請嘗試重新登入!");
                            }

                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                msg = "(" + rowsAffected + " rows affected)",
                                data = insertResult
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

        #region //AddItDetail -- 新增庫存異動單詳細資料 -- Ann 2023-12-15
        public string AddItDetail(int ItId, int MtlItemId, string ItMtlItemName, string ItMtlItemSpec
            , double ItQty, int UomId, int InventoryId, int ToInventoryId, string ItRemark, string StorageLocation, string ToStorageLocation)
        {
            try
            {
                int UserId = CreateBy;
                if (ItMtlItemName.Length <= 0) throw new SystemException("【品名】不能為空!");
                if (ItQty <= 0) throw new SystemException("【異動數量】不能為空!");
                if (InventoryId == ToInventoryId && StorageLocation == ToStorageLocation) throw new SystemException("轉出及轉入庫不能相同!!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
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
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //確認庫存異動單據資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.ItErpPrefix, a.DocDate, a.ConfirmStatus
                                    FROM SCM.InventoryTransaction a 
                                    WHERE a.ItId = @ItId";
                            dynamicParameters.Add("ItId", ItId);

                            var ItResult = sqlConnection.Query(sql, dynamicParameters);

                            if (ItResult.Count() <= 0) throw new SystemException("庫存異動單據資料錯誤!!");

                            DateTime DocDate = new DateTime();
                            string ItErpPrefix = "";
                            foreach (var item in ItResult)
                            {
                                if (item.ItErpPrefix == "1201" && ToInventoryId <= 0) throw new SystemException("轉撥單據需填寫轉入庫資訊!!");
                                if (item.ConfirmStatus != "N") throw new SystemException("目前單據狀態不可修改!!");
                                DocDate = item.DocDate;
                                ItErpPrefix = item.ItErpPrefix;
                            }
                            #endregion

                            #region //比對ERP關帳日期
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MA013)) MA013
                                    FROM CMSMA";
                            var cmsmaResult = sqlConnection2.Query(sql, dynamicParameters);
                            if (cmsmaResult.Count() < 0) throw new SystemException("EPR關帳日期資料錯誤!!");

                            foreach (var item in cmsmaResult)
                            {
                                string eprDate = item.MA013;
                                string erpYear = eprDate.Substring(0, 4);
                                string erpMonth = eprDate.Substring(4, 2);
                                string erpDay = eprDate.Substring(6, 2);
                                string erpFullDate = erpYear + "-" + erpMonth + "-" + erpDay;
                                DateTime erpDateTime = Convert.ToDateTime(erpFullDate);
                                int compare = DocDate.CompareTo(erpDateTime);
                                if (compare <= 0) throw new SystemException("ERP已關帳(" + erpFullDate + ")，無法開立此日期(" + DocDate + ")之單據!!");
                            }
                            #endregion

                            #region //判斷單位資料是否有誤
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM PDM.UnitOfMeasure a 
                                    WHERE a.UomId = @UomId";
                            dynamicParameters.Add("UomId", UomId);

                            var UomResult = sqlConnection.Query(sql, dynamicParameters);

                            if (UomResult.Count() <= 0) throw new SystemException("單位資料錯誤!!");
                            #endregion

                            #region //判斷品號資料是否有誤
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT MtlItemNo, TransferStatus
                                    FROM PDM.MtlItem
                                    WHERE MtlItemId = @MtlItemId";
                            dynamicParameters.Add("MtlItemId", MtlItemId);

                            var result3 = sqlConnection.Query(sql, dynamicParameters);
                            if (result3.Count() <= 0) throw new SystemException("品號資料有誤!");

                            string MtlItemNo = "";
                            foreach (var item in result3)
                            {
                                if (item.TransferStatus != "Y") throw new SystemException("此品號尚未拋轉到ERP，無法使用!!");
                                MtlItemNo = item.MtlItemNo;
                            }
                            #endregion

                            #region //確認ERP品號相關資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MB004)) MB004
                                    , LTRIM(RTRIM(MB030)) MB030
                                    , LTRIM(RTRIM(MB031)) MB031
                                    FROM INVMB
                                    WHERE MB001 = @MB001";
                            dynamicParameters.Add("MB001", MtlItemNo);

                            var INVMBResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (INVMBResult.Count() <= 0) throw new SystemException("ERP品號基本資料錯誤!!");

                            foreach (var item in INVMBResult)
                            {
                                #region //判斷ERP品號生效日與失效日
                                if (item.MB030 != "" && item.MB030 != null)
                                {
                                    #region //判斷日期需大於或等於生效日
                                    string EffectiveDate = item.MB030;
                                    string effYear = EffectiveDate.Substring(0, 4);
                                    string effMonth = EffectiveDate.Substring(4, 2);
                                    string effDay = EffectiveDate.Substring(6, 2);
                                    DateTime effFullDate = Convert.ToDateTime(effYear + "-" + effMonth + "-" + effDay);
                                    int effresult = DateTime.Compare(DocDate, effFullDate);
                                    if (effresult < 0) throw new SystemException("此品號尚未生效，無法使用!!");
                                    #endregion
                                }

                                if (item.MB031 != "" && item.MB031 != null)
                                {
                                    #region //判斷日期需小於或等於失效日
                                    string ExpirationDate = item.MB031;
                                    string effYear = ExpirationDate.Substring(0, 4);
                                    string effMonth = ExpirationDate.Substring(4, 2);
                                    string effDay = ExpirationDate.Substring(6, 2);
                                    DateTime effFullDate = Convert.ToDateTime(effYear + "-" + effMonth + "-" + effDay);
                                    int effresult = DateTime.Compare(DocDate, effFullDate);
                                    if (effresult > 0) throw new SystemException("此品號已失效，無法使用!!");
                                    #endregion
                                }
                                #endregion
                            }
                            #endregion

                            #region //判斷轉出庫存資料是否有誤
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 InventoryNo
                                    FROM SCM.Inventory
                                    WHERE InventoryId = @InventoryId";
                            dynamicParameters.Add("InventoryId", InventoryId);

                            var InventoryResult = sqlConnection.Query(sql, dynamicParameters);
                            if (InventoryResult.Count() <= 0) throw new SystemException("轉出庫存資料有誤!");

                            string InventoryNo = "";
                            foreach (var item in InventoryResult)
                            {
                                InventoryNo = item.InventoryNo;
                            }
                            #endregion

                            #region //判斷轉入庫存資料是否有誤
                            string ToInventoryNo = "";
                            if (ToInventoryId > 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 InventoryNo ToInventoryNo
                                        FROM SCM.Inventory
                                        WHERE InventoryId = @InventoryId";
                                dynamicParameters.Add("InventoryId", ToInventoryId);

                                var ToInventoryResult = sqlConnection.Query(sql, dynamicParameters);
                                if (ToInventoryResult.Count() <= 0) throw new SystemException("轉入庫存資料有誤!");

                                foreach (var item in ToInventoryResult)
                                {
                                    ToInventoryNo = item.ToInventoryNo;
                                }
                            }
                            #endregion

                            #region //確認轉出庫別儲位管理資料正確性
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MC009)) LocationFlag
                                    FROM CMSMC
                                    WHERE MC001 = @MC001";
                            dynamicParameters.Add("MC001", InventoryNo);

                            var CMSMCResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (CMSMCResult.Count() <= 0) throw new SystemException("轉出庫別【" + InventoryNo + "】資料錯誤!!");

                            string LocationFlag = "N";
                            foreach (var item in CMSMCResult)
                            {
                                LocationFlag = item.LocationFlag;

                                if (item.LocationFlag == "Y" && StorageLocation.Length <= 0)
                                {
                                    throw new SystemException("轉出庫別【" + InventoryNo + "】需設定儲位!!");
                                }
                                else if (item.LocationFlag == "Y")
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT RTRIM(LTRIM(a.NL002)) NL002
                                            FROM CMSNL a 
                                            WHERE a.NL001 = @InventoryNo
                                            AND a.NL002 = @StorageLocation";
                                    dynamicParameters.Add("InventoryNo", InventoryNo);
                                    dynamicParameters.Add("StorageLocation", StorageLocation);

                                    var CMSNLResult = sqlConnection2.Query(sql, dynamicParameters);

                                    if (CMSNLResult.Count() <= 0) throw new SystemException("儲位【" + StorageLocation + "】資料有誤!!");
                                }
                            }
                            #endregion

                            #region //確認轉入庫別儲位管理資料正確性
                            if (ToInventoryNo.Length > 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT LTRIM(RTRIM(MC009)) LocationFlag
                                        FROM CMSMC
                                        WHERE MC001 = @MC001";
                                dynamicParameters.Add("MC001", ToInventoryNo);

                                var CMSMCResult2 = sqlConnection2.Query(sql, dynamicParameters);

                                if (CMSMCResult2.Count() <= 0) throw new SystemException("轉入庫別【" + ToInventoryNo + "】資料錯誤!!");

                                foreach (var item in CMSMCResult2)
                                {
                                    if (item.LocationFlag == "Y" && ToStorageLocation.Length <= 0)
                                    {
                                        throw new SystemException("轉入庫別【" + ToInventoryNo + "】需設定儲位!!");
                                    }
                                    else if (item.LocationFlag == "Y")
                                    {
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT RTRIM(LTRIM(a.NL002)) NL002
                                                FROM CMSNL a 
                                                WHERE a.NL001 = @InventoryNo
                                                AND a.NL002 = @StorageLocation";
                                        dynamicParameters.Add("InventoryNo", ToInventoryNo);
                                        dynamicParameters.Add("StorageLocation", ToStorageLocation);

                                        var CMSNLResult = sqlConnection2.Query(sql, dynamicParameters);

                                        if (CMSNLResult.Count() <= 0) throw new SystemException("儲位資料有誤!!");
                                    }
                                }
                            }
                            #endregion

                            #region //取單身序號
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT MAX(ItSequence) ItSequence
                                    FROM SCM.ItDetail a 
                                    WHERE a.ItId = @ItId";
                            dynamicParameters.Add("ItId", ItId);

                            var ItDetailResult = sqlConnection.Query(sql, dynamicParameters);

                            string ItSequence = "";
                            foreach (var item in ItDetailResult)
                            {
                                int maxSeq = Convert.ToInt32(item.ItSequence) + 1;
                                ItSequence = maxSeq.ToString("D4");
                            }
                            #endregion

                            #region //部門費用領料、轉撥單需檢查庫存量是否足夠
                            if (ItErpPrefix == "1101" || ItErpPrefix == "1201")
                            {
                                #region //檢核基本庫別庫存
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT ISNULL(MC007, 0) MC007
                                        FROM INVMC
                                        WHERE MC001 = @MtlItemNo
                                        AND MC002 = @InventoryNo";
                                dynamicParameters.Add("MtlItemNo", MtlItemNo);
                                dynamicParameters.Add("InventoryNo", InventoryNo);

                                var INVMCResult = sqlConnection2.Query(sql, dynamicParameters);

                                if (INVMCResult.Count() <= 0) throw new SystemException("品號【" + MtlItemNo + "】查無庫別【" + InventoryNo + "】相關資料!!");

                                foreach (var item in INVMCResult)
                                {
                                    if (ItQty > Convert.ToDouble(item.MC007)) throw new SystemException("異動數量已超過庫別【" + InventoryNo + "】庫存量【" + item.MC007 + "】!!");
                                }
                                #endregion

                                #region //檢核儲位庫別庫存
                                if (LocationFlag == "Y")
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT MM005
                                            FROM INVMM
                                            WHERe MM001 = @MtlItemNo
                                            AND MM002 = @InventoryNo
                                            AND MM003 = @StorageLocation";
                                    dynamicParameters.Add("MtlItemNo", MtlItemNo);
                                    dynamicParameters.Add("InventoryNo", InventoryNo);
                                    dynamicParameters.Add("StorageLocation", StorageLocation);

                                    var INVMMResult = sqlConnection2.Query(sql, dynamicParameters);

                                    if (INVMMResult.Count() <= 0) throw new SystemException("品號【" + MtlItemNo + "】庫別【" + InventoryNo  + "】儲位【" + StorageLocation + "】庫存量不足!!");

                                    foreach (var item in INVMMResult)
                                    {
                                        if (ItQty > Convert.ToDouble(item.MM005))
                                        {
                                            throw new SystemException("品號【" + MtlItemNo + "】庫別【" + InventoryNo + "】儲位【" + StorageLocation + "】庫存量不足!!");
                                        }
                                    }
                                }
                                #endregion
                            }
                            #endregion

                            #region //取得本國幣別
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 LTRIM(RTRIM(MA003)) MA003 FROM CMSMA";

                            var CMSMAResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (CMSMAResult.Count() <= 0) throw new SystemException("共用參數檔案資料錯誤!!");

                            string Currency = "";
                            foreach (var item in CMSMAResult)
                            {
                                Currency = item.MA003;
                            }
                            #endregion

                            #region //小數點後取位
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.MF005, a.MF006
                                    FROM CMSMF a 
                                    WHERE a.MF001 = @Currency";
                            dynamicParameters.Add("Currency", Currency);

                            var CMSMFResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (CMSMFResult.Count() <= 0) throw new SystemException("金額小數點取位資料錯誤!!");

                            int unitDecimal = 0;
                            int amountDecimal = 0;
                            foreach (var item in CMSMFResult)
                            {
                                unitDecimal = Convert.ToInt32(item.MF005);
                                amountDecimal = Convert.ToInt32(item.MF006);
                            }
                            #endregion

                            #region //計算單位成本、金額
                            //公式: INVMB 【MB065 庫存金額】/【MB064 庫存數量】
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT MB064, MB065
                                    FROM INVMB
                                    WHERE MB001 = @MtlItemNo";
                            dynamicParameters.Add("MtlItemNo", MtlItemNo);

                            var INVMBResult2 = sqlConnection2.Query(sql, dynamicParameters);

                            if (INVMBResult2.Count() <= 0) throw new SystemException("ERP品號資料錯誤!!");

                            double UnitCost = -1;
                            double Amount = -1;
                            foreach (var item in INVMBResult2)
                            {
                                if (item.MB065 <= 0 && item.MB064 <= 0)
                                {
                                    UnitCost = 0;
                                }
                                else
                                {
                                    UnitCost = Convert.ToDouble(Math.Round(item.MB065 / item.MB064, unitDecimal));
                                }
                                Amount = Convert.ToDouble(Math.Round(ItQty * UnitCost, amountDecimal));
                            }
                            #endregion

                            #region //INSERT SCM.ItDetail
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO SCM.ItDetail (ItId, ItSequence, MtlItemId, ItMtlItemName, ItMtlItemSpec
                                    , ItQty, UomId, InvQty, UnitCost, Amount, InventoryId, ToInventoryId, StorageLocation, ToStorageLocation, ItRemark
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.ItDetailId
                                    VALUES (@ItId, @ItSequence, @MtlItemId, @ItMtlItemName, @ItMtlItemSpec
                                    , @ItQty, @UomId, @InvQty, @UnitCost, @Amount, @InventoryId, @ToInventoryId, @StorageLocation, @ToStorageLocation, @ItRemark
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    ItId,
                                    ItSequence,
                                    MtlItemId,
                                    ItMtlItemName,
                                    ItMtlItemSpec,
                                    ItQty,
                                    UomId,
                                    InvQty = 0,
                                    UnitCost,
                                    Amount,
                                    InventoryId,
                                    ToInventoryId = ToInventoryId > 0 ? ToInventoryId : (int?)null,
                                    StorageLocation,
                                    ToStorageLocation,
                                    ItRemark,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy = UserId,
                                    LastModifiedBy = UserId
                                });

                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            int rowsAffected = insertResult.Count();
                            #endregion

                            #region //重新計算目前總數量、總金額
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT SUM(a.ItQty) TotalQty, SUM(a.Amount) TotalAmount
                                    FROM SCM.ItDetail a 
                                    WHERE ItId = @ItId";
                            dynamicParameters.Add("ItId", ItId);

                            var ItDetailResult2 = sqlConnection.Query(sql, dynamicParameters);

                            if (ItDetailResult2.Count() <= 0) throw new SystemException("庫存異動單據詳細資料錯誤!!");

                            double TotalQty = 0;
                            double TotalAmount = 0;
                            foreach (var item in ItDetailResult2)
                            {
                                TotalQty = Convert.ToDouble(item.TotalQty);
                                TotalAmount = Math.Round(Convert.ToDouble(item.TotalAmount), amountDecimal);
                            }
                            #endregion

                            #region //Update SCM.InventoryTransaction 總數量、總金額
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.InventoryTransaction SET
                                    TotalQty += @ItQty,
                                    Amount += @Amount,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE ItId = @ItId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    ItQty,
                                    Amount,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    ItId
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
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

        #region //AddMaterialRequisition -- 新增領退料單資料 -- Ann 2023-12-19
        public string AddMaterialRequisition(string MrErpPrefix, string MrDate, string DocDate, string ProductionLine
            , string ContractManufacturer, string Remark)
        {
            try
            {
                int UserId = CreateBy;
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
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
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            if (MrErpPrefix.Length <= 0) throw new SystemException("【ERP領退料單別】不能為空!");
                            if (MrDate.Length < 0) throw new SystemException("【領退料日期】不能為空!");
                            if (DocDate.Length < 0) throw new SystemException("【單據建立日期】不能為空!");

                            #region //比對ERP關帳日期
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MA013)) MA013
                                    FROM CMSMA";
                            var cmsmaResult = sqlConnection2.Query(sql, dynamicParameters);
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

                            #region //組合MES領料單號
                            string RequesitionNo = DateTime.Now.ToString("yyyy") + DateTime.Now.ToString("MM") + DateTime.Now.ToString("dd") + DateTime.Now.ToString("HH") + DateTime.Now.ToString("mm") + DateTime.Now.ToString("ss");
                            #endregion

                            #region //判斷 公司+MES領料單號 是否重複
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.MaterialRequisition
                                    WHERE CompanyId = @CompanyId
                                    AND RequesitionNo = @RequesitionNo";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("RequesitionNo", RequesitionNo);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() > 0) throw new SystemException("【公司 + MES領料單號】重複，請重新輸入!");
                            #endregion

                            #region //隨機取號ERP單號
                            string MrErpNo = BaseHelper.RandomCode(11);
                            #endregion

                            #region //INSERT MES.MaterialRequisition
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.MaterialRequisition (CompanyId, RequesitionNo, MrErpPrefix, MrErpNo, MrDate
                                    , DocDate, DocType, ProductionLine, Remark, JournalStatus, PriorityType, NegativeStatus, SignupStatus
                                    , SourceType, ConfirmStatus, ContractManufacturer
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.MrId, INSERTED.MrErpPrefix, INSERTED.MrErpNo, INSERTED.CreateBy, INSERTED.CompanyId
                                    VALUES (@CompanyId, @RequesitionNo, @MrErpPrefix, @MrErpNo, @MrDate
                                    , @DocDate, @DocType, @ProductionLine, @Remark, @JournalStatus, @PriorityType, @NegativeStatus, @SignupStatus
                                    , @SourceType, @ConfirmStatus, @ContractManufacturer
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    CompanyId = CurrentCompany,
                                    RequesitionNo,
                                    MrErpPrefix,
                                    MrErpNo,
                                    MrDate,
                                    DocDate,
                                    DocType = MrErpPrefix.Substring(0, 2),
                                    ProductionLine,
                                    Remark,
                                    JournalStatus = "N",
                                    PriorityType = "1",
                                    NegativeStatus = "N",
                                    SignupStatus = "0",
                                    SourceType = "1",
                                    ConfirmStatus = "N",
                                    ContractManufacturer,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy = UserId,
                                    LastModifiedBy = UserId
                                });

                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            int rowsAffected = insertResult.Count();

                            foreach (var item in insertResult)
                            {
                                if (BaseHelper.CheckUserAuthority(item.CreateBy, item.CompanyId, "A", "InventoryTransaction", "add", sqlConnection).Equals("N")) throw new SystemException("人員權限檢核有異常，請嘗試重新登入!");
                            }
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

        #region //AddMrWipOrder -- 新增領退料單設定資料(MrWipOrder + MrDetail) -- Ann 2023-12-20
        public string AddMrWipOrder(int MrId, int MoId, string WoErpPrefix, string WoErpNo, string MrType
            , double Quantity, string RequisitionCode, string NegativeStatus, string Remark, string MaterialCategory, string SubinventoryType
            , string LineSeq, int InventoryId, string StorageLocation)
        {
            try
            {
                if (WoErpPrefix.Length < 0) throw new SystemException("【製令單別】不能為空!");
                if (WoErpPrefix.Length > 4) throw new SystemException("【製令單別】長度錯誤!");
                if (WoErpNo.Length < 0) throw new SystemException("【製令單號】不能為空!");
                if (WoErpNo.Length > 11) throw new SystemException("【製令單號】長度錯誤!");
                if (MrType.Length < 0) throw new SystemException("【領料方式】長度錯誤!");
                if (Quantity < 0) throw new SystemException("【領退料套數】不能小於0!");
                if (RequisitionCode.Length < 0) throw new SystemException("【領料碼】不能為空!");
                if (NegativeStatus.Length < 0) throw new SystemException("【庫存不足照領】不能為空!");
                if (Remark.Length > 255) throw new SystemException("【備註】長度錯誤!");
                if (MaterialCategory.Length < 0) throw new SystemException("【材料型態】不能為空!");
                if (SubinventoryType.Length < 0) throw new SystemException("【庫別型態】不能為空!");
                if (SubinventoryType.Length < 0) throw new SystemException("【輸入序號】不能為空!");
                if (SubinventoryType.Length > 4) throw new SystemException("【輸入序號】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
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
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //判斷領退料單資料是否有誤
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT MrDate, ConfirmStatus, MrErpPrefix, MrErpNo
                                    FROM MES.MaterialRequisition
                                    WHERE CompanyId = @CompanyId
                                    AND MrId = @MrId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("MrId", MrId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("領退料單資料有誤!");

                            DateTime MrDate = new DateTime();
                            string MrErpPrefix = "";
                            string MrErpNo = "";
                            foreach (var item in result)
                            {
                                if (item.ConfirmStatus != "N") throw new SystemException("領退料單已確認，無法更改!");
                                MrDate = item.MrDate;
                                MrErpPrefix = item.MrErpPrefix;
                                MrErpNo = item.MrErpNo;
                            }
                            #endregion

                            #region //查詢單據類別設定資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.MQ010, LTRIM(RTRIM(a.MQ019)) MQ019, LTRIM(RTRIM(a.MQ054)) MQ054
                                    FROM CMSMQ a 
                                    WHERE MQ001 = @MrErprPrefix";
                            dynamicParameters.Add("MrErprPrefix", MrErpPrefix);

                            var CMSMQResult = sqlConnection2.Query(sql, dynamicParameters);
                            if (CMSMQResult.Count() <= 0) throw new SystemException("ERP單別資料有誤!!");

                            int MQ010 = -1;
                            string CheckMoFlag = "";
                            string ExcessFlag = "";
                            foreach (var item in CMSMQResult)
                            {
                                MQ010 = Convert.ToInt32(item.MQ010);
                                CheckMoFlag = item.MQ019;
                                ExcessFlag = item.MQ054;
                                if (ExcessFlag == "N" && Remark == "") throw new SystemException("超領單別備註不可為空!!");
                            }
                            #endregion

                            #region //確認製令資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.Quantity
                                    , c.InventoryId, c.InventoryUomId UomId
                                    , d.InventoryNo
                                    , e.UomNo
                                    FROM MES.ManufactureOrder a
                                    INNER JOIN MES.WipOrder b ON a.WoId = b.WoId
                                    INNER JOIN PDM.MtlItem c ON b.MtlItemId = c.MtlItemId
                                    INNER JOIN SCM.Inventory d ON c.InventoryId = d.InventoryId
                                    INNER JOIN PDM.UnitOfMeasure e ON c.InventoryUomId = e.UomId
                                    WHERE a.MoId = @MoId";
                            dynamicParameters.Add("MoId", MoId);

                            var ManufactureOrderResult = sqlConnection.Query(sql, dynamicParameters);

                            if (ManufactureOrderResult.Count() <= 0) throw new SystemException("製令資料錯誤!!");

                            int MoQuantity = 0;
                            int UomId = -1;
                            string UomNo = "";
                            foreach (var item in ManufactureOrderResult)
                            {
                                MoQuantity = item.Quantity;
                                UomId = item.UomId;
                                UomNo = item.UomNo;
                            }
                            #endregion

                            #region //查詢ERP製令資訊
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.TA015 PlanQty, a.TA016 RequisitionSetQty
                                    , a.TA017 StockInQty, FORMAT(CONVERT(DATETIME, LTRIM(RTRIM(a.TA014))), 'yyyy-MM-dd') ActualEnd
                                    FROM MOCTA a
                                    WHERE a.TA001 = @WoErpPrefix
                                    AND a.TA002 = @WoErpNo";
                            dynamicParameters.Add("WoErpPrefix", WoErpPrefix);
                            dynamicParameters.Add("WoErpNo", WoErpNo);

                            var ErpWoInfoResult = sqlConnection2.Query(sql, dynamicParameters);
                            if (ErpWoInfoResult.Count() <= 0) throw new SystemException("ERP製令(MOCTA)資料有誤!");

                            double PlanQty = -1;
                            double RequisitionSetQty = -1;
                            double StockInQty = -1;
                            string ActualEnd = "";
                            foreach (var item in ErpWoInfoResult)
                            {
                                PlanQty = Convert.ToDouble(item.PlanQty);
                                RequisitionSetQty = Convert.ToDouble(item.RequisitionSetQty);
                                StockInQty = Convert.ToDouble(item.StockInQty);
                                ActualEnd = item.ActualEnd;
                            }
                            #endregion

                            #region //搜尋此製令所有領料單資訊，並找出實際已核單領料套數
                            int ErpQuantity = 0;
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT b.MrErpPrefix, b.MrErpNo
                                    , c.Quantity
                                    , d.WoErpPrefix, d.WoErpNo
                                    FROM MES.MrWipOrder a
                                    INNER JOIN MES.MaterialRequisition b ON a.MrId = b.MrId
                                    INNER JOIN MES.ManufactureOrder c ON a.MoId = c.MoId
                                    INNER JOIN MES.WipOrder d ON c.WoId = d.WoId
                                    WHERE a.MoId = @MoId";
                            dynamicParameters.Add("MoId", MoId);

                            var MrWipOrderResult = sqlConnection.Query(sql, dynamicParameters);

                            if (MrWipOrderResult.Count() > 0)
                            {
                                foreach (var item2 in MrWipOrderResult)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.TD006 Quantity
                                            , b.MQ010
                                            , (
                                                SELECT TOP 1 1
                                                FROM MOCTE x
                                                WHERE x.TE001 = a.TD001
                                                AND x.TE002 = a.TD002
                                                AND x.TE019 = 'Y'
                                            ) CheckConfirm
                                            FROM MOCTD a
                                            INNER JOIN CMSMQ b ON a.TD001 = b.MQ001
                                            WHERE a.TD001 = @MrErpPrefix
                                            AND a.TD002 = @MrErpNo
                                            AND a.TD003 = @WoErpPrefix
                                            AND a.TD004 = @WoErpNo";
                                    dynamicParameters.Add("MrErpPrefix", item2.MrErpPrefix);
                                    dynamicParameters.Add("MrErpNo", item2.MrErpNo);
                                    dynamicParameters.Add("WoErpPrefix", item2.WoErpPrefix);
                                    dynamicParameters.Add("WoErpNo", item2.WoErpNo);

                                    var ErpMrWipOrderResult = sqlConnection2.Query(sql, dynamicParameters);

                                    foreach (var item3 in ErpMrWipOrderResult)
                                    {
                                        if (item3.CheckConfirm != null)
                                        {
                                            if (item3.MQ010 > 0)
                                            {
                                                ErpQuantity -= item3.Quantity;
                                            }
                                            else
                                            {
                                                ErpQuantity += item3.Quantity;
                                            }
                                        }
                                    }
                                }
                            }
                            #endregion

                            #region //判斷MES製令資料是否有誤
                            if (MQ010 < 0) //領料
                            {
                                if (Quantity > (MoQuantity - ErpQuantity) && ExcessFlag == "Y") throw new SystemException("領取數量超過未領用量(" + (MoQuantity - ErpQuantity).ToString() + ")!");
                            }
                            else if (MQ010 > 0) //退料
                            {
                                if (Quantity > ErpQuantity) throw new SystemException("退料數量超過剩餘已領套數(" + ErpQuantity.ToString() + ")!");
                            }
                            else
                            {
                                throw new SystemException("此領料單別(" + MrErpPrefix + ")設定資料有誤，請通知開發人員!!");
                            }

                            #region //確認領退料日期是否小於製令實際完工日期
                            DateTime actualEnd = Convert.ToDateTime(ActualEnd);
                            int compare = actualEnd.CompareTo(MrDate);
                            //if (compare < 0) throw new SystemException("領退料日期不能大於製令實際完工日期!");
                            #endregion
                            #endregion

                            #region //判斷 領退料單 + MES製令 + 領退料方式 是否重複
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.MrWipOrder
                                    WHERE MrId = @MrId
                                    AND MoId = @MoId
                                    AND MrType = @MrType";
                            dynamicParameters.Add("MrId", MrId);
                            dynamicParameters.Add("MoId", MoId);
                            dynamicParameters.Add("MrType", MrType);

                            var result4 = sqlConnection.Query(sql, dynamicParameters);
                            if (result4.Count() > 0) throw new SystemException("【領退料單 + MES製令 + 領退料方式】重複，請重新輸入!");
                            #endregion

                            #region //取單身序號
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT MAX(LineSeq) LineSeq
                                    FROM MES.MrWipOrder a 
                                    WHERE a.MrId = @MrId
                                    AND a.MoId = @MoId";
                            dynamicParameters.Add("MrId", MrId);
                            dynamicParameters.Add("MoId", MoId);

                            var MrWipOrderResult2 = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item in MrWipOrderResult2)
                            {
                                int maxSeq = Convert.ToInt32(item.LineSeq) + 1;
                                LineSeq = maxSeq.ToString("D4");
                            }
                            #endregion

                            #region //確認庫別資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.InventoryNo
                                    FROM SCM.Inventory a 
                                    WHERE a.InventoryId = @InventoryId";
                            dynamicParameters.Add("InventoryId", InventoryId);

                            var InventoryResult = sqlConnection.Query(sql, dynamicParameters);

                            if (InventoryResult.Count() <= 0) throw new SystemException("庫別資料錯誤!!");

                            string InventoryNo = "";
                            foreach (var item in InventoryResult)
                            {
                                InventoryNo = item.InventoryNo;
                            }
                            #endregion

                            #region //確認庫別儲位管理資料正確性
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MC009)) LocationFlag
                                    FROM CMSMC
                                    WHERE MC001 = @MC001";
                            dynamicParameters.Add("MC001", InventoryNo);

                            var CMSMCResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (CMSMCResult.Count() <= 0) throw new SystemException("轉出庫別【" + InventoryNo + "】資料錯誤!!");

                            string LocationFlag = "N";
                            foreach (var item in CMSMCResult)
                            {
                                LocationFlag = item.LocationFlag;

                                if (item.LocationFlag == "Y" && StorageLocation.Length <= 0)
                                {
                                    throw new SystemException("轉出庫別【" + InventoryNo + "】需設定儲位!!");
                                }
                                else if (item.LocationFlag == "Y")
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT RTRIM(LTRIM(a.NL002)) NL002
                                            FROM CMSNL a 
                                            WHERE a.NL001 = @InventoryNo
                                            AND a.NL002 = @StorageLocation";
                                    dynamicParameters.Add("InventoryNo", InventoryNo);
                                    dynamicParameters.Add("StorageLocation", StorageLocation);

                                    var CMSNLResult = sqlConnection2.Query(sql, dynamicParameters);

                                    if (CMSNLResult.Count() <= 0) throw new SystemException("儲位【" + StorageLocation + "】資料有誤!!");
                                }
                            }
                            #endregion

                            #region //新增MES.MrWipOrder
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO MES.MrWipOrder (MrId, MoId, WoErpPrefix, WoErpNo, MrType, Quantity
                                    , InventoryId, StorageLocation, SubInventoryCode, RequisitionCode, NegativeStatus, Remark
                                    , MaterialCategory, SubinventoryType, LineSeq, UomId
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.MrId, INSERTED.MoId
                                    VALUES (@MrId, @MoId, @WoErpPrefix, @WoErpNo, @MrType, @Quantity
                                    , @InventoryId, @StorageLocation, @SubInventoryCode, @RequisitionCode, @NegativeStatus, @Remark
                                    , @MaterialCategory, @SubinventoryType, @LineSeq, @UomId
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MrId,
                                    MoId,
                                    WoErpPrefix,
                                    WoErpNo,
                                    MrType,
                                    Quantity,
                                    InventoryId,
                                    StorageLocation,
                                    SubInventoryCode = InventoryNo,
                                    RequisitionCode,
                                    NegativeStatus,
                                    Remark,
                                    MaterialCategory,
                                    SubinventoryType,
                                    LineSeq,
                                    UomId,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            int rowsAffected = insertResult.Count();
                            #endregion

                            #region //新增MES.MrDetail

                            #region //確認沒有條碼已開工
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.BarcodeNo, FORMAT(a.CreateDate, 'yyyy-MM-dd HH:mm:ss')  BarcodeCreateDate
                                    FROM MES.MrBarcodeRegister a 
                                    INNER JOIN MES.MrDetail b ON a.MrDetailId = b.MrDetailId
                                    WHERE b.MrId = @MrId
                                    AND b.MoId = @MoId";
                            dynamicParameters.Add("MrId", MrId);
                            dynamicParameters.Add("MoId", MoId);

                            var MrBarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item in MrBarcodeResult)
                            {
                                #region //比對條碼歷程時間判斷是否為舊條碼
                                DateTime BarcodeCreateDate = Convert.ToDateTime(item.BarcodeCreateDate);
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 b.FinishDate
                                        FROM MES.Barcode a
                                        INNER JOIN MES.BarcodeProcess b ON a.BarcodeId = b.BarcodeId
                                        WHERE a.BarcodeNo = @BarcodeNo
                                        AND b.MoId = @MoId
                                        ORDER BY FinishDate DESC";
                                dynamicParameters.Add("BarcodeNo", item.BarcodeNo);
                                dynamicParameters.Add("MoId", MoId);

                                var CheckBarcodeProcessResult = sqlConnection.Query(sql, dynamicParameters);

                                if (CheckBarcodeProcessResult.Count() > 0)
                                {
                                    foreach (var item2 in CheckBarcodeProcessResult)
                                    {
                                        DateTime BarcodeProcessDate = item2.FinishDate;
                                        int dateResult = DateTime.Compare(BarcodeProcessDate, BarcodeCreateDate);
                                        if (dateResult > 0) throw new SystemException("條碼【" + item.BarcodeNo + "】已有過站紀錄，無法重整單身資料!!");
                                    }
                                }
                                #endregion
                            }
                            #endregion

                            #region //檢查此發料單單身是否有條碼轉移紀錄
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.LogId, a.MoId, a.BarcodeNo, a.CurrentMoProcessId, a.NextMoProcessId, a.CurrentProdStatus, a.BarcodeStatus
                                    , b.BarcodeStatus CurrentBarcodeStatus
                                    , c.StatusName
                                    FROM MES.MrBarcodeTransferLog a
                                    INNER JOIN MES.Barcode b ON a.BarcodeNo = b.BarcodeNo
                                    INNER JOIN BAS.[Status] c ON b.BarcodeStatus = c.StatusNo AND c.StatusSchema = 'Barcode.BarcodeStatus'
                                    INNER JOIN MES.MrDetail d ON a.MrDetailId = d.MrDetailId
                                    WHERE d.MrId = @MrId
                                    AND d.MoId = @MoId";
                            dynamicParameters.Add("MrId", MrId);
                            dynamicParameters.Add("MoId", MoId);

                            var TransferLogResult = sqlConnection.Query(sql, dynamicParameters);

                            if (TransferLogResult.Count() > 0)
                            {
                                foreach (var item2 in TransferLogResult)
                                {
                                    if (item2.CurrentBarcodeStatus != "1" && item2.CurrentBarcodeStatus != "3")
                                    {
                                        throw new SystemException("此條碼【" + item2.BarcodeNo + "】狀態【" + item2.StatusName + "】無法刪除，請先進行退料或解除綁定!");
                                    }

                                    #region //UPDATE MES.Barcode回原先資料
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE MES.Barcode SET
                                            MoId = @MoId,
                                            CurrentMoProcessId = @CurrentMoProcessId,
                                            NextMoProcessId = @NextMoProcessId,
                                            CurrentProdStatus = @CurrentProdStatus,
                                            BarcodeStatus = @BarcodeStatus,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE BarcodeNo = @BarcodeNo";
                                    dynamicParameters.AddDynamicParams(
                                      new
                                      {
                                          item2.MoId,
                                          item2.CurrentMoProcessId,
                                          item2.NextMoProcessId,
                                          item2.CurrentProdStatus,
                                          item2.BarcodeStatus,
                                          LastModifiedDate,
                                          LastModifiedBy,
                                          item2.BarcodeNo
                                      });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //將轉移LOG移除
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE FROM MES.MrBarcodeTransferLog
                                            WHERE LogId = @LogId";
                                    dynamicParameters.Add("LogId", item2.LogId);

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                            }
                            else
                            {
                                #region //刪除MES.Barcode資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE a
                                        FROM MES.Barcode a
                                        INNER JOIN MES.MrBarcodeRegister b ON a.BarcodeNo = b.BarcodeNo
                                        INNER JOIN MES.MrDetail c ON b.MrDetailId = c.MrDetailId
                                        WHERE c.MrId = @MrId
                                        AND c.MoId = @MoId";
                                dynamicParameters.Add("MrId", MrId);
                                dynamicParameters.Add("MoId", MoId);

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }
                            #endregion

                            #region //DELETE MES.MrBarcodeRegister
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE a 
                                    FROM MES.MrBarcodeRegister a
                                    INNER JOIN MES.MrDetail b ON a.MrDetailId = b.MrDetailId
                                    WHERE b.MrId = @MrId
                                    AND b.MoId = @MoId";
                            dynamicParameters.Add("MrId", MrId);
                            dynamicParameters.Add("MoId", MoId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //DELETE MES.MrDetail
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE a 
                                    FROM MES.MrDetail a
                                    WHERE a.MrId = @MrId
                                    AND a.MoId = @MoId";
                            dynamicParameters.Add("MrId", MrId);
                            dynamicParameters.Add("MoId", MoId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //查看目前MrDetail序號
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT MAX(MrSequence) MrSequence
                                    FROM MES.MrDetail
                                    WHERE MrId = @MrId";
                            dynamicParameters.Add("MrId", MrId);

                            var result5 = sqlConnection.Query(sql, dynamicParameters);
                            int currentMrSequence = -1;
                            foreach (var item in result5)
                            {
                                currentMrSequence = Convert.ToInt32(item.MrSequence);
                            }

                            string MrSequence = (currentMrSequence + 1).ToString().PadLeft(4, '0');
                            #endregion

                            #region //取得MES製令用料設定
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT ISNULL(a.CompositionQuantity, 1) CompositionQuantity, ISNULL(a.Base, 1) Base, ISNULL(a.LossRate, 0) LossRate, a.BarcodeCtrl, a.MainBarcode, a.ControlType, a.DecompositionFlag
                                    , b.MtlItemId, b.DemandRequisitionQty, b.RequisitionQty, b.SubstituteMtlItemId, b.MaterialProperties, b.SubstituteStatus
                                    , c.WoErpPrefix, c.WoErpNo, c.PlanQty, c.StockInQty
                                    , d.UomId, d.UomNo
                                    , e.MtlItemNo, e.MtlItemName
                                    , f.InventoryId, f.InventoryNo
                                    , g.MtlItemNo SubstitutionMtlItemNo, g.MtlItemName SubstitutionMtlItemName
                                    , h.TypeName
                                    , i.MtlItemNo
                                    FROM MES.MoMtlSetting a
                                    INNER JOIN MES.WoDetail b ON a.WoDetailId = b.WoDetailId
                                    INNER JOIN MES.WipOrder c ON b.WoId = c.WoId
                                    INNER JOIN PDM.UnitOfMeasure d ON a.UomId = d.UomId
                                    INNER JOIN PDM.MtlItem e ON b.MtlItemId = e.MtlItemId
                                    INNER JOIN SCM.Inventory f ON b.InventoryId = f.InventoryId
                                    INNER JOIN PDM.MtlItem g ON b.SubstituteMtlItemId = g.MtlItemId
                                    INNER JOIN BAS.[Type] h ON a.ControlType = h.TypeNo AND h.TypeSchema = 'MoMtlSetting.ControlType'
                                    INNER JOIN PDM.MtlItem i ON b.MtlItemId = i.MtlItemId
                                    WHERE a.MoId = @MoId
                                    ORDER BY a.MainBarcode DESC";
                            dynamicParameters.Add("MoId", MoId);

                            var result6 = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item in result6)
                            {
                                if (item.SubstituteStatus == "Y" && (item.MtlItemId != item.SubstituteMtlItemId))
                                {
                                    continue;
                                }

                                #region //依據材料性質判斷要帶出的用料MaterialCategory
                                string MaterialProperties = "";
                                if (item.MaterialProperties == null || item.MaterialProperties == "")
                                {
                                    throw new SystemException("材料性質設定錯誤，請確認工單資料是否正確!!");
                                }
                                else
                                {
                                    if (item.MaterialProperties != "1" && item.MaterialProperties != "2" && item.MaterialProperties != "5") continue;
                                    MaterialProperties = item.MaterialProperties;
                                }

                                if (MaterialCategory != "1" && MaterialCategory != "*" && MaterialCategory != "2" && MaterialCategory != "5")
                                {
                                    MaterialCategory = "*";
                                }

                                if (MaterialCategory != "*")
                                {
                                    if (MaterialCategory != MaterialProperties) continue;
                                }
                                #endregion

                                #region //取得ERP製令單身需領/已領用量
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TB004 DemandRequisitionQty, TB005 RequisitionQty
                                        FROM MOCTB
                                        WHERE TB001 = @WoErpPrefix
                                        AND TB002 = @WoErpNo
                                        AND TB003 = @MtlItemNo";
                                dynamicParameters.Add("WoErpPrefix", WoErpPrefix);
                                dynamicParameters.Add("WoErpNo", WoErpNo);
                                dynamicParameters.Add("MtlItemNo", item.MtlItemNo);

                                var ErpWoDetailResult = sqlConnection2.Query(sql, dynamicParameters);

                                if (ErpWoDetailResult.Count() <= 0) throw new SystemException("查無ERP製令單身資料!!");

                                double DemandRequisitionQty = -1;
                                double RequisitionQty = -1;
                                foreach (var item2 in ErpWoDetailResult)
                                {
                                    DemandRequisitionQty = Convert.ToDouble(item2.DemandRequisitionQty);
                                    RequisitionQty = Convert.ToDouble(item2.RequisitionQty);
                                }
                                #endregion

                                if (MQ010 > 0)
                                {
                                    if (RequisitionQty <= 0) continue;
                                }

                                #region //若為補足需領量，從EPR查是否有剩餘的料要領
                                if (MrType == "2" || MrType == "3")
                                {
                                    if (DemandRequisitionQty - RequisitionQty <= 0) continue;
                                }
                                #endregion

                                #region //若為退已領用量，從ERP查是否有已領量
                                if (MrType == "5")
                                {
                                    if (RequisitionQty <= 0) continue;
                                }
                                #endregion

                                #region //計算一套可以領/退多少料
                                //領料公式: 未領用量 / 總套數(預計生產量) / 單位用量 / 底數
                                //退料公式: 剩餘生產量 / 單位用量
                                double MrDetailQty = -1; //預計領料量
                                double ActualQuantity = -1; //實際領料量
                                if (item.BomDetailId != null)
                                {
                                    double unitQuantity = (DemandRequisitionQty - RequisitionQty) / item.PlanQty / item.CompositionQuantity;
                                    if (unitQuantity * Quantity * item.CompositionQuantity > (DemandRequisitionQty - RequisitionQty)) MrDetailQty = DemandRequisitionQty - RequisitionQty;
                                    else MrDetailQty = unitQuantity * Quantity * item.CompositionQuantity;
                                }
                                else
                                {
                                    switch (MrType)
                                    {
                                        //確認公式是否為: ((領料數量*單位用量)/基數) + ((領料數量*單位用量/基數)*損耗率) 再四捨五入
                                        case "1":
                                            //MrDetailQty = ((Quantity * Convert.ToDouble(item.CompositionQuantity)) / Convert.ToDouble(item.Base)) + (((Quantity * Convert.ToDouble(item.CompositionQuantity)) / Convert.ToDouble(item.Base)) * item.LossRate);
                                            MrDetailQty = (Quantity * Convert.ToDouble(item.CompositionQuantity)) / Convert.ToDouble(item.Base);
                                            break;
                                        case "2":
                                            MrDetailQty = DemandRequisitionQty - RequisitionQty;
                                            break;
                                        case "3":
                                            MrDetailQty = DemandRequisitionQty - RequisitionQty;
                                            break;
                                        case "4":
                                            MrDetailQty = (Quantity * item.CompositionQuantity) / item.Base;
                                            break;
                                        case "5":
                                            MrDetailQty = RequisitionQty;
                                            break;
                                    }
                                }

                                if (item.DecompositionFlag == "Y")
                                {
                                    ActualQuantity = MrDetailQty;
                                }
                                else if (item.BarcodeCtrl == "Y" && item.UomNo == "PCS")
                                {
                                    ActualQuantity = 0;
                                }
                                else
                                {
                                    ActualQuantity = MrDetailQty;
                                }
                                #endregion

                                #region //查詢目前此庫別庫存量，根據庫存是否足夠進行不同流程
                                string DetailDesc = "";

                                double InventoryQty = 0;
                                dynamicParameters = new DynamicParameters();
                                if (LocationFlag == "Y")
                                {
                                    #region //庫別需儲位管理
                                    sql = @"SELECT a.MM005 InventoryQty
                                            FROM INVMM a
                                            WHERE a.MM001 = @MM001
                                            AND a.MM002 = @MM002
                                            AND a.MM003 = @MM003";
                                    dynamicParameters.Add("MM001", item.MtlItemNo);
                                    dynamicParameters.Add("MM002", InventoryNo);
                                    dynamicParameters.Add("MM003", StorageLocation);
                                    #endregion
                                }
                                else
                                {
                                    sql = @"SELECT a.MC007 InventoryQty
                                            FROM INVMC a
                                            WHERE a.MC001 = @MC001
                                            AND a.MC002 = @MC002";
                                    dynamicParameters.Add("MC001", item.MtlItemNo);
                                    dynamicParameters.Add("MC002", InventoryNo);
                                }

                                var InventoryResult2 = sqlConnection2.Query(sql, dynamicParameters);

                                if (InventoryResult2.Count() > 0)
                                {
                                    foreach (var item2 in InventoryResult2)
                                    {
                                        InventoryQty = Convert.ToDouble(item2.InventoryQty);

                                        #region //計算同料號同庫別，但未確認之領料單領料數量
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT ISNULL(SUM(a.TE005), 0) TE005
                                                FROM MOCTE a
                                                WHERE a.TE004 = @TE004
                                                AND a.TE008 = @TE008
                                                AND a.TE019 = 'N'
                                                AND a.TE001 != @TE001
                                                AND a.TE002 != @TE002";
                                        dynamicParameters.Add("TE004", item.MtlItemNo);
                                        dynamicParameters.Add("TE008", InventoryNo);
                                        dynamicParameters.Add("TE001", MrErpPrefix);
                                        dynamicParameters.Add("TE002", MrErpNo);

                                        var MtlItemQtyResult = sqlConnection2.Query(sql, dynamicParameters);

                                        double? TotalMtlItemQty = 0;
                                        foreach (var item3 in MtlItemQtyResult)
                                        {
                                            TotalMtlItemQty = Convert.ToDouble(item3.TE005);
                                        }
                                        #endregion

                                        #region //尋找本身領料單內是否有同品號數量
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT ISNULL(SUM(a.Quantity), 0) TE005
                                                FROM MES.MrDetail a
                                                WHERE a.MrId = @MrId
                                                AND a.MtlItemId = @MtlItemId
                                                AND a.InventoryId = @InventoryId";
                                        dynamicParameters.Add("MrId", MrId);
                                        dynamicParameters.Add("MtlItemId", item.MtlItemId);
                                        dynamicParameters.Add("InventoryId", InventoryId);

                                        var ThisMtlItemQtyResult = sqlConnection.Query(sql, dynamicParameters);

                                        foreach (var item3 in ThisMtlItemQtyResult)
                                        {
                                            TotalMtlItemQty += Convert.ToDouble(item3.TE005);
                                        }
                                        #endregion

                                        double? CurrentMtlItemQty = Convert.ToDouble(InventoryQty);
                                        double? AvailabilityQty = CurrentMtlItemQty - TotalMtlItemQty/* - Convert.ToDouble(item.Quantity)*/;
                                        if (AvailabilityQty >= 0)
                                        {
                                            DetailDesc = "可用量: " + AvailabilityQty;
                                        }
                                        else
                                        {
                                            DetailDesc = "庫存不足，需領用量: " + item.Quantity;
                                        }
                                    }
                                }
                                else
                                {
                                    DetailDesc = "庫存不足，需領用量: " + item.Quantity;
                                }
                                #endregion

                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO MES.MrDetail (MrId, MrSequence, MtlItemId, Quantity, ActualQuantity, UomId, Unit
                                        , InventoryId, SubInventoryCode, ProcessCode, LotNumber, MoId, WoErpPrefix, WoErpNo
                                        , DetailDesc, Remark, MaterialCategory, ConfirmStatus, ProjectCode, BondedStatus
                                        , SubstituteStatus, OfficialItemStatus, SubstitutionId, SubstitutionMtlItemNo, SubstituteProcessCode
                                        , SubstituteQty, SubstituteRate, StorageLocation
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.MrDetailId
                                        VALUES (@MrId, @MrSequence, @MtlItemId, @Quantity, @ActualQuantity, @UomId, @Unit
                                        , @InventoryId, @SubInventoryCode, @ProcessCode, @LotNumber, @MoId, @WoErpPrefix, @WoErpNo
                                        , @DetailDesc, @Remark, @MaterialCategory, @ConfirmStatus, @ProjectCode, @BondedStatus
                                        , @SubstituteStatus, @OfficialItemStatus, @SubstitutionId, @SubstitutionMtlItemNo, @SubstituteProcessCode
                                        , @SubstituteQty, @SubstituteRate, @StorageLocation
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        MrId,
                                        MrSequence,
                                        item.MtlItemId,
                                        Quantity = Math.Round(MrDetailQty * 1000) / 1000,
                                        ActualQuantity = 0,
                                        item.UomId,
                                        Unit = item.UomNo,
                                        InventoryId,
                                        SubInventoryCode = InventoryNo,
                                        ProcessCode = "****",
                                        LotNumber = "",
                                        MoId,
                                        WoErpPrefix,
                                        WoErpNo,
                                        DetailDesc,
                                        Remark = ExcessFlag == "N" ? Remark : "",
                                        MaterialCategory = MaterialProperties,
                                        ConfirmStatus = "N",
                                        ProjectCode = "",
                                        BondedStatus = "N",
                                        SubstituteStatus = "N",
                                        OfficialItemStatus = "2",
                                        SubstitutionId = item.SubstitutionId != null ? (int?)item.SubstituteMtlItemId : null,
                                        item.SubstitutionMtlItemNo,
                                        SubstituteProcessCode = "****",
                                        SubstituteQty = 0,
                                        SubstituteRate = 0,
                                        StorageLocation,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });

                                insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();

                                int nextMrSequence = Convert.ToInt32(MrSequence) + 1;
                                MrSequence = nextMrSequence.ToString().PadLeft(4, '0');
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

        #region //AddMrBarcodeRegister -- 新增領料單條碼註冊資料 -- Ann 2023-12-21
        public string AddMrBarcodeRegister(int MrDetailId, string BarcodeNo)
        {
            try
            {
                BarcodeNo = BarcodeNo.Trim(); //濾掉空格
                int UserId = CreateBy;
                int rowsAffected = 0;
                int OldBarcodeCompany = 0;
                string ErpNo = "";
                string ErpDbName = "";

                if (UserId <= 0) throw new SystemException("使用者ID錯誤，請嘗試重新登入!!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
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

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            if (BarcodeNo.Length < 0) throw new SystemException("【條碼代碼】不能為空!");
                            if (BarcodeNo.Length > 32) throw new SystemException("【條碼代碼】長度錯誤!");

                            BarcodeNo = BarcodeNo.ToString().ToUpper(); //將條碼強制轉成大寫

                            #region //檢查領退料單詳細資料是否有誤
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.ConfirmStatus, a.ActualQuantity, a.MoId, a.Quantity, a.Unit
                                    , b.WoErpPrefix, b.WoErpNo
                                    , c.DemandRequisitionQty, c.RequisitionQty
                                    , d.TrayBarcode, d.Source, d.MrType
                                    , e.MtlItemNo, e.MtlItemName
                                    , FORMAT(f.DocDate, 'MM') DocDate
                                    , g.MtlItemNo MrDetailMtlItemNo
                                    FROM MES.MrDetail a
                                    INNER JOIN MES.WipOrder b ON a.WoErpPrefix = b.WoErpPrefix AND a.WoErpNo = b.WoErpNo
                                    LEFT JOIN MES.WoDetail c ON a.MtlItemId = c.MtlItemId AND c.WoId = b.WoId
                                    INNER JOIN MES.MoSetting d ON a.MoId = d.MoId
                                    LEFT JOIN PDM.MtlItem e ON c.MtlItemId = e.MtlItemId
                                    INNER JOIN MES.MaterialRequisition f ON f.MrId = a.MrId
                                    INNER JOIN PDM.MtlItem g ON a.MtlItemId = g.MtlItemId
                                    WHERE MrDetailId = @MrDetailId";
                            dynamicParameters.Add("MrDetailId", MrDetailId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("領退料單詳細資料有誤!");

                            double? ActualQuantity = -1;
                            int MoId = -1;
                            string Unit = "";
                            double? DemandRequisitionQty = -1;
                            double? RequisitionQty = -1;
                            string TrayBarcode = "";
                            string Source = "";
                            double? Quantity = -1;
                            string MtlItemNo = "";
                            string MtlItemName = "";
                            string MrType = "";
                            string DocDate = "";
                            string WoErpPrefix = "";
                            string WoErpNo = "";
                            string MrDetailMtlItemNo = "";
                            foreach (var item in result)
                            {
                                if (item.ConfirmStatus != "N") throw new SystemException("領退料單已被確認，無法更改!");
                                Unit = item.Unit;
                                Quantity = item.Quantity;
                                ActualQuantity = item.ActualQuantity;
                                MoId = item.MoId;
                                TrayBarcode = item.TrayBarcode;
                                Source = item.Source;
                                DemandRequisitionQty = item.DemandRequisitionQty;
                                RequisitionQty = item.RequisitionQty;
                                MtlItemNo = item.MtlItemNo;
                                MtlItemName = item.MtlItemName;
                                MrType = item.MrType;
                                DocDate = item.DocDate;
                                WoErpPrefix = item.WoErpPrefix;
                                WoErpNo = item.WoErpNo;
                                MrDetailMtlItemNo = item.MrDetailMtlItemNo;
                            }
                            #endregion

                            #region //確認ERP工單(MOCTA)狀態
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(TA011)) Status
                                    FROM MOCTA
                                    WHERE TA001 = @TA001
                                    AND TA002 = @TA002";
                            dynamicParameters.Add("TA001", WoErpPrefix);
                            dynamicParameters.Add("TA002", WoErpNo);
                            var WoErpStatusReesult = sqlConnection2.Query(sql, dynamicParameters);

                            foreach (var item2 in WoErpStatusReesult)
                            {
                                if (item2.Status == "Y" || item2.Status == "y") throw new SystemException("ERP製令狀態無法進行領料!!");
                            }
                            #endregion

                            #region //確認是否為Tray模式
                            if (TrayBarcode == "Y" && Source == "MR")
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.BarcodeNo, a.TrayNo
                                        FROM MES.Tray a
                                        WHERE a.TrayNo = @TrayNo";
                                dynamicParameters.Add("TrayNo", BarcodeNo);

                                var TrayResult = sqlConnection.Query(sql, dynamicParameters);

                                if (TrayResult.Count() < 0) throw new SystemException("查無此Tray盤資訊!");

                                string TrayNo = "";
                                foreach (var item in TrayResult)
                                {
                                    if (item.BarcodeNo == null) throw new SystemException("此Tray盤尚未綁定條碼!");
                                    TrayNo = item.TrayNo;
                                    BarcodeNo = item.BarcodeNo;
                                }
                            }
                            #endregion

                            #region //確認是否為包裝條碼領料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.PackageBarcodeId, a.Status
                                    , b.StatusName
                                    FROM MES.PackageBarcode a
                                    INNER JOIN BAS.Status b ON a.Status = b.StatusNo AND b.StatusSchema = 'PackageBarcode.Status'
                                    WHERE a.PackageBarcodeNo = @PackageBarcodeNo";
                            dynamicParameters.Add("PackageBarcodeNo", BarcodeNo);

                            var PackageBarcodeResult = sqlConnection.Query(sql, dynamicParameters);
                            #endregion

                            if (PackageBarcodeResult.Count() > 0)
                            {
                                foreach (var item in PackageBarcodeResult)
                                {
                                    if (item.Status != "2") throw new SystemException("此包裝條碼【" + BarcodeNo + "】狀態為【" + item.StatusName + "】，不可進行發料!!");

                                    #region //找出此包裝條碼底下所有產品條碼，並進行後續條碼控管流程
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT b.BarcodeNo
                                            FROM MES.PackageBarcodeReference a
                                            INNER JOIN MES.Barcode b ON a.BarcodeId = b.BarcodeId
                                            WHERE a.PackageBarcodeId = @PackageBarcodeId";
                                    dynamicParameters.Add("PackageBarcodeId", item.PackageBarcodeId);

                                    var PackageBarcodeReferenceResult = sqlConnection.Query(sql, dynamicParameters);

                                    if (PackageBarcodeReferenceResult.Count() <= 0) throw new SystemException("此包裝條碼【" + BarcodeNo + "】尚未綁定任何一筆產品條碼!!");

                                    foreach (var item2 in PackageBarcodeReferenceResult)
                                    {
                                        BarcodeNo = item2.BarcodeNo;
                                        MrBarcodeFunction();
                                    }
                                    #endregion

                                    #region //更新包裝條碼狀態為【已領料】
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE MES.PackageBarcode SET
                                            [Status] = '3',
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE PackageBarcodeId = @PackageBarcodeId";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            LastModifiedDate,
                                            LastModifiedBy,
                                            item.PackageBarcodeId
                                        });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                            }
                            else
                            {
                                MrBarcodeFunction();
                            }

                            void MrBarcodeFunction()
                            {
                                #region //判斷 領退料單詳細資料+條碼代碼 是否重複
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM MES.MrBarcodeRegister
                                        WHERE MrDetailId = @MrDetailId
                                        AND BarcodeNo = @BarcodeNo";
                                dynamicParameters.Add("MrDetailId", MrDetailId);
                                dynamicParameters.Add("BarcodeNo", BarcodeNo);

                                var result2 = sqlConnection.Query(sql, dynamicParameters);
                                if (result2.Count() > 0) throw new SystemException("【領退料單詳細資料 + 條碼代碼】重複，請重新輸入!");
                                #endregion

                                #region //確認此條碼是否已註冊且尚未開工
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM MES.MrBarcodeRegister a
                                        WHERE a.BarcodeNo = @BarcodeNo
                                        AND a.MrDetailId = @MrDetailId";
                                dynamicParameters.Add("BarcodeNo", BarcodeNo);
                                dynamicParameters.Add("MrDetailId", MrDetailId);

                                var MrBarcodeRegisterResult = sqlConnection.Query(sql, dynamicParameters);
                                if (MrBarcodeRegisterResult.Count() > 0) throw new SystemException("此條碼【" + BarcodeNo + "】已綁定!");
                                #endregion

                                #region //確認及更改BARCODE資訊

                                #region //先查詢現在製令首站
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 a.MoProcessId
                                        FROM MES.MoProcess a
                                        WHERE a.MoId = @MoId
                                        ORDER BY a.SortNumber";
                                dynamicParameters.Add("MoId", MoId);

                                var MoProcessIdResult = sqlConnection.Query(sql, dynamicParameters);

                                if (MoProcessIdResult.Count() <= 0) throw new SystemException("查詢新製令首站時發生錯誤!!");

                                int CurrentMoProcessId = -1;
                                foreach (var item3 in MoProcessIdResult)
                                {
                                    CurrentMoProcessId = item3.MoProcessId;
                                }
                                #endregion

                                #region //查詢此發料條碼是否為新條碼或舊條碼
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.MoId, a.BarcodeQty, a.CurrentMoProcessId, a.NextMoProcessId, a.CurrentProdStatus, a.BarcodeStatus
                                        , b.StatusName
                                        FROM MES.Barcode a
                                        INNER JOIN BAS.[Status] b ON a.BarcodeStatus = b.StatusNo AND b.StatusSchema = 'Barcode.BarcodeStatus'
                                        WHERE a.BarcodeNo = @BarcodeNo";
                                dynamicParameters.Add("BarcodeNo", BarcodeNo);

                                var CheckBarcodeResult = sqlConnection.Query(sql, dynamicParameters);
                                #endregion

                                #region //判斷此料號是否為主條碼及是否為可切割料號
                                string MainBarcode = "";
                                string DecompositionFlag = "";
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT d.MainBarcode, d.DecompositionFlag
                                        FROM MES.MrDetail a
                                        INNER JOIN MES.ManufactureOrder b ON a.MoId = b.MoId
                                        INNER JOIN MES.WoDetail c ON a.MtlItemId = c.MtlItemId AND c.WoId = b.WoId
                                        INNER JOIN MES.MoMtlSetting d ON c.WoDetailId = d.WoDetailId
                                        WHERE a.MrDetailId = @MrDetailId";
                                dynamicParameters.Add("MrDetailId", MrDetailId);

                                var MainBarcodeResult = sqlConnection.Query(sql, dynamicParameters);
                                foreach (var item3 in MainBarcodeResult)
                                {
                                    MainBarcode = item3.MainBarcode;
                                    DecompositionFlag = item3.DecompositionFlag;
                                }
                                #endregion

                                if (CheckBarcodeResult.Count() <= 0)
                                {
                                    #region //新條碼流程

                                    #region //INSERT MES.Barcode
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO MES.Barcode (BarcodeNo, CurrentMoProcessId, NextMoProcessId, MoId, BarcodeQty, BarcodeProcessId, CurrentProdStatus, BarcodeStatus
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            VALUES (@BarcodeNo, @CurrentMoProcessId, @NextMoProcessId, @MoId, @BarcodeQty, @BarcodeProcessId, @CurrentProdStatus, @BarcodeStatus
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            BarcodeNo,
                                            CurrentMoProcessId,
                                            NextMoProcessId = CurrentMoProcessId,
                                            MoId,
                                            BarcodeQty = 1,
                                            BarcodeProcessId = -1,
                                            CurrentProdStatus = "P",
                                            BarcodeStatus = "1",
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy = UserId,
                                            LastModifiedBy = UserId
                                        });

                                    var insertResult2 = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult2.Count();
                                    #endregion

                                    if (Unit == "PCS" && (DecompositionFlag == "N" || DecompositionFlag == "")) //因鎢鋼模仁流程，故只能先以單位管控
                                    {
                                        if (MtlItemName == null) MtlItemName = MrDetailMtlItemNo;
                                        if ((ActualQuantity + 1) > Quantity) throw new SystemException("此條碼數量(1)已超過【" + MtlItemName + "】未領用量(" + (Quantity - ActualQuantity) + ")");
                                        ActualQuantity++;
                                    }

                                    #endregion
                                }
                                else
                                {
                                    #region //舊條碼流程
                                    foreach (var item4 in CheckBarcodeResult)
                                    {
                                        #region //檢查條碼狀態
                                        if (item4.CurrentProdStatus != "P" && item4.CurrentProdStatus != "R") throw new SystemException("此條碼【" + BarcodeNo + "】非良品條碼，無法領取!!");
                                        #endregion

                                        #region //取得條碼公司別
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT c.CompanyId
                                                  FROM MES.Barcode a
                                                       INNER JOIN MES.ManufactureOrder b ON a.MoId = b.MoId
	                                                   INNER JOIN MES.WipOrder c ON b.WoId = c.WoId
                                                 WHERE a.BarcodeNo = @BarcodeNo";

                                        dynamicParameters.Add("BarcodeNo", BarcodeNo);
                                        var OldCompanyResult = sqlConnection.Query(sql, dynamicParameters);
                                        if (OldCompanyResult.Count() > 0)
                                        {
                                            foreach (var item in OldCompanyResult)
                                            {
                                                OldBarcodeCompany = Convert.ToInt32(item.CompanyId);
                                            }
                                        }
                                        else
                                        {
                                            OldBarcodeCompany = CurrentCompany;
                                        }
                                        #endregion

                                        #region //檢查此條碼是否已完工狀態
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT TOP 1 d.WoErpPrefix + '-' + d.WoErpNo WoErpFullNo, e.ProcessAlias
                                                FROM MES.BarcodeProcess a 
                                                INNER JOIN MES.Barcode b ON a.BarcodeId = b.BarcodeId
                                                INNER JOIN MES.ManufactureOrder c ON a.MoId = c.MoId
                                                INNER JOIN MES.WipOrder d ON c.WoId = d.WoId
                                                INNER JOIN MES.MoProcess e ON a.MoProcessId = e.MoProcessId
                                                WHERE b.BarcodeNo = @BarcodeNo
                                                AND a.FinishDate IS NULL";
                                        dynamicParameters.Add("BarcodeNo", BarcodeNo);

                                        var BarcodeProcessResult = sqlConnection.Query(sql, dynamicParameters);

                                        if (BarcodeProcessResult.Count() > 0)
                                        {
                                            foreach (var item in BarcodeProcessResult)
                                            {
                                                throw new SystemException("此條碼【" + BarcodeNo + "】尚在製令【" + item.WoErpFullNo + "】【" + item.ProcessAlias + "】未完工!!");
                                            }
                                        }
                                        #endregion

                                        #region //確認條碼狀態是否可以領料
                                        if (item4.BarcodeStatus != "8" && item4.BarcodeStatus != "10" && item4.BarcodeStatus != "11")
                                        {
                                            throw new SystemException("此條碼【" + BarcodeNo + "】【目前製令:" + item4.MoId + "】已存在，且狀態【" + item4.StatusName + "】不可重新發料!!");
                                        }
                                        #endregion

                                        #region //特定生產模式下，檢核前後製令用料關聯性

                                        #endregion

                                        #region //後續條碼領料流程

                                        #region //檢查是否為批量條碼未過首站
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT b.BarcodeId
                                                FROM MES.BarcodePrint a
                                                LEFT JOIN MES.Barcode b ON a.BarcodeNo = b.BarcodeNo
                                                WHERE a.BarcodeNo = @BarcodeNo";
                                        dynamicParameters.Add("BarcodeNo", BarcodeNo);

                                        var result4 = sqlConnection.Query(sql, dynamicParameters);

                                        if (result4.Count() > 0)
                                        {
                                            foreach (var item2 in result4)
                                            {
                                                if (item2.BarcodeId == null) throw new SystemException("此條碼【" + BarcodeNo + "】已註冊為批量條碼，無法發料!");
                                            }
                                        }
                                        #endregion

                                        #region //確認當月物料卡控機制
                                        if (MrType == "N")
                                        {
                                            #region //取得舊條碼ERP製令單據日期
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"SELECT FORMAT(a.DocDate, 'yyyy-MM') OldDocDate
                                                    FROM MES.WipOrder a 
                                                    INNER JOIN MES.ManufactureOrder b ON a.WoId = b.WoId
                                                    WHERE b.MoId = @MoId";
                                            dynamicParameters.Add("BarcodeNo", BarcodeNo);
                                            dynamicParameters.Add("MoId", MoId);

                                            var OldDocDateResult = sqlConnection.Query(sql, dynamicParameters);

                                            string OldDocDate = "";
                                            foreach (var item in OldDocDateResult)
                                            {
                                                OldDocDate = item.OldDocDate;
                                            }
                                            #endregion

                                            #region //取得現在領料單ERP製令單據日期
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"SELECT FORMAT(c.DocDate, 'yyyy-MM') NewDocDate
                                                    FROM MES.MrDetail a 
                                                    INNER JOIN MES.ManufactureOrder b ON a.MoId = b.MoId
                                                    INNER JOIN MES.WipOrder c ON b.WoId = c.WoId
                                                    WHERE a.MrDetailId = @MrDetailId";
                                            dynamicParameters.Add("MrDetailId", MrDetailId);

                                            var NewDocDateResult = sqlConnection.Query(sql, dynamicParameters);

                                            string NewDocDate = "";
                                            foreach (var item in NewDocDateResult)
                                            {
                                                NewDocDate = item.NewDocDate;
                                            }
                                            #endregion

                                            if (OldDocDate != NewDocDate) throw new SystemException("此製令已設定無法領取非當月物料!!");
                                        }
                                        #endregion

                                        #region //UPDAET BARCODE STATUS
                                        int BarcodeStatusRowsAffected = 0;
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"UPDATE MES.Barcode SET
                                                MoId = @MoId,
                                                CurrentMoProcessId = @CurrentMoProcessId,
                                                NextMoProcessId = @NextMoProcessId,
                                                BarcodeQty = @BarcodeQty,
                                                BarcodeStatus = @BarcodeStatus,
                                                CurrentProdStatus = @CurrentProdStatus,
                                                LastModifiedDate = @LastModifiedDate,
                                                LastModifiedBy = @LastModifiedBy
                                                WHERE BarcodeNo = @BarcodeNo";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                MoId,
                                                CurrentMoProcessId,
                                                NextMoProcessId = CurrentMoProcessId,
                                                item4.BarcodeQty,
                                                BarcodeStatus = MainBarcode == "Y" ? "1" : "3",
                                                CurrentProdStatus = "P",
                                                LastModifiedDate,
                                                LastModifiedBy = UserId,
                                                BarcodeNo
                                            });

                                        BarcodeStatusRowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                        if (BarcodeStatusRowsAffected <= 0) throw new SystemException("條碼狀態未被更改成功，請重新操作一次!!");
                                        #endregion

                                        #region //INSERT LOG
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO MES.MrBarcodeTransferLog (MrDetailId, MoId, BarcodeNo, CurrentMoProcessId, NextMoProcessId, CurrentProdStatus, BarcodeStatus
                                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                                OUTPUT INSERTED.LogId
                                                VALUES (@MrDetailId, @MoId, @BarcodeNo, @CurrentMoProcessId, @NextMoProcessId, @CurrentProdStatus, @BarcodeStatus
                                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                MrDetailId,
                                                item4.MoId,
                                                BarcodeNo,
                                                item4.CurrentMoProcessId,
                                                item4.NextMoProcessId,
                                                item4.CurrentProdStatus,
                                                item4.BarcodeStatus,
                                                CreateDate,
                                                LastModifiedDate,
                                                CreateBy = UserId,
                                                LastModifiedBy = UserId
                                            });

                                        var insertResult2 = sqlConnection.Query(sql, dynamicParameters);

                                        rowsAffected += insertResult2.Count();

                                        if (insertResult2.Count() < 1) throw new SystemException("LOG紀錄新增失敗!!");
                                        #endregion

                                        if (Unit == "PCS" && DecompositionFlag == "N" || DecompositionFlag == "") //因鎢鋼模仁流程，故只能先以單位管控
                                        {
                                            if ((ActualQuantity + item4.BarcodeQty) > Quantity) throw new SystemException("此條碼【" + BarcodeNo + "】數量(" + item4.BarcodeQty + ")已超過未領用量(" + (Quantity - ActualQuantity) + ")");
                                            ActualQuantity += item4.BarcodeQty;
                                        }
                                        #endregion
                                    }
                                    #endregion
                                }
                                #endregion

                                #region //INSERT MES.MrBarcodeRegister
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO MES.MrBarcodeRegister (MrDetailId, BarcodeNo
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        VALUES (@MrDetailId, @BarcodeNo
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        MrDetailId,
                                        BarcodeNo,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy = UserId,
                                        LastModifiedBy = UserId
                                    });

                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();
                                #endregion

                                #region //更新MrDetail領取套數
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE MES.MrDetail SET
                                        ActualQuantity = @ActualQuantity,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE MrDetailId = @MrDetailId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        ActualQuantity,
                                        LastModifiedDate,
                                        LastModifiedBy = UserId,
                                        MrDetailId
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

        #region //AddMrBarcodeReRegister -- 新增領料單條碼退料資料 -- Ann 2023-12-21
        public string AddMrBarcodeReRegister(int MrDetailId, string BarcodeNo)
        {
            try
            {
                BarcodeNo = BarcodeNo.Trim(); //濾掉空格
                int UserId = CreateBy;

                if (UserId <= 0) throw new SystemException("使用者ID錯誤，請嘗試重新登入!!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (BarcodeNo.Length < 0) throw new SystemException("【條碼代碼】不能為空!");
                        if (BarcodeNo.Length > 32) throw new SystemException("【條碼代碼】長度錯誤!");

                        #region //檢查領退料單詳細資料是否有誤
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ConfirmStatus, a.ActualQuantity, a.MoId, a.Quantity
                                , b.WoErpPrefix, b.WoErpNo
                                , c.DemandRequisitionQty, c.RequisitionQty
                                FROM MES.MrDetail a
                                INNER JOIN MES.WipOrder b ON a.WoErpPrefix = b.WoErpPrefix AND a.WoErpNo = b.WoErpNo
                                INNER JOIN MES.WoDetail c ON a.MtlItemId = c.MtlItemId AND c.WoId = b.WoId
                                WHERE MrDetailId = @MrDetailId";
                        dynamicParameters.Add("MrDetailId", MrDetailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("領退料單詳細資料有誤!");

                        double ActualQuantity = -1;
                        double Quantity = -1;
                        int MoId = -1;
                        string WoErpPrefix = "";
                        string WoErpNo = "";
                        foreach (var item in result)
                        {
                            if (item.ConfirmStatus != "N") throw new SystemException("領退料單已被確認，無法更改!");
                            ActualQuantity = item.ActualQuantity;
                            MoId = item.MoId;
                            WoErpPrefix = item.WoErpPrefix;
                            WoErpNo = item.WoErpNo;
                            Quantity = item.Quantity;
                        }
                        #endregion

                        #region //確認條碼資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.BarcodeStatus, a.BarcodeQty, a.CurrentMoProcessId, a.NextMoProcessId, a.CurrentProdStatus, a.MoId
                                , b.StatusName
                                , (
                                    SELECT TOP 1 1
                                    FROM MES.BarcodeProcess x 
                                    WHERE x.BarcodeId = a.BarcodeId
                                    AND x.FinishDate IS NULL
                                ) BarcodeProcess
                                FROM MES.Barcode a 
                                INNER JOIN BAS.[Status] b ON a.BarcodeStatus = b.StatusNo AND b.StatusSchema = 'Barcode.BarcodeStatus'
                                WHERE a.BarcodeNo = @BarcodeNo";
                        dynamicParameters.Add("BarcodeNo", BarcodeNo);

                        var BarcodeResult = sqlConnection.Query(sql, dynamicParameters);

                        if (BarcodeResult.Count() <= 0) throw new SystemException("條碼【" + BarcodeNo + "】資料錯誤!!");

                        string BarcodeStatus = "";
                        int CurrentMoProcessId = -1;
                        int NextMoProcessId = -1;
                        string CurrentProdStatus = "";
                        int OldMoId = -1;
                        int BarcodeQty = 0;
                        foreach (var item in BarcodeResult)
                        {
                            if (item.BarcodeProcess != null) throw new SystemException("條碼【" + BarcodeNo + "】目前非完工狀態!!");
                            if (item.MoId != MoId) throw new SystemException("條碼【" + BarcodeNo + "】非此製令【" + item.WoErpPrefix + "-" + item.WoErpNo + "】所綁定!!");
                            if (ActualQuantity + item.BarcodeQty > Quantity) throw new SystemException("退料數量已達上限!");
                            if (item.BarcodeStatus == "4") throw new SystemException("條碼【" + BarcodeNo + "】目前尚為上料綁定狀態，請先由上料系統解綁後再進行退料!!");
                            BarcodeStatus = item.BarcodeStatus;
                            CurrentMoProcessId = item.CurrentMoProcessId;
                            NextMoProcessId = item.NextMoProcessId;
                            CurrentProdStatus = item.CurrentProdStatus;
                            OldMoId = item.MoId;
                            BarcodeQty = item.BarcodeQty;
                        }
                        #endregion

                        #region //判斷 領退料單詳細資料+條碼代碼 是否重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM MES.MrBarcodeReRegister
                                WHERE MrDetailId = @MrDetailId
                                AND BarcodeNo = @BarcodeNo";
                        dynamicParameters.Add("MrDetailId", MrDetailId);
                        dynamicParameters.Add("BarcodeNo", BarcodeNo);

                        var result2 = sqlConnection.Query(sql, dynamicParameters);
                        if (result2.Count() > 0) throw new SystemException("【領退料單詳細資料 + 條碼代碼】重複，請重新輸入!");
                        #endregion

                        #region //INSERT MES.MrBarcodeReRegister
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.MrBarcodeReRegister (MrDetailId, BarcodeNo
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                VALUES (@MrDetailId, @BarcodeNo
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                MrDetailId,
                                BarcodeNo,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });

                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        int rowsAffected = insertResult.Count();
                        #endregion

                        #region //更新MrDetail領取套數
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.MrDetail SET
                                ActualQuantity = @ActualQuantity,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE MrDetailId = @MrDetailId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ActualQuantity = ActualQuantity + BarcodeQty,
                                LastModifiedDate,
                                LastModifiedBy,
                                MrDetailId
                            });

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //INSERT MES.MrBarcodeTransferLog
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO MES.MrBarcodeTransferLog (MrDetailId, MoId, BarcodeNo, CurrentMoProcessId, NextMoProcessId, CurrentProdStatus, BarcodeStatus
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                VALUES (@MrDetailId, @MoId, @BarcodeNo, @CurrentMoProcessId, @NextMoProcessId, @CurrentProdStatus, @BarcodeStatus
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                MrDetailId,
                                MoId,
                                BarcodeNo,
                                CurrentMoProcessId,
                                NextMoProcessId,
                                CurrentProdStatus,
                                BarcodeStatus,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy = UserId,
                                LastModifiedBy = UserId
                            });

                        var insertResult2 = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += insertResult2.Count();
                        #endregion

                        #region //更新MES.Barcode Status
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.Barcode SET
                                BarcodeStatus = '8',
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE BarcodeNo = @BarcodeNo";
                        dynamicParameters.AddDynamicParams(
                          new
                          {
                              LastModifiedDate,
                              LastModifiedBy,
                              BarcodeNo
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
        #region //UpdateInventoryTransaction -- 更新庫存異動單據 -- Ann 2023-12-15
        public string UpdateInventoryTransaction(int ItId, string ItDate, string DocDate, int DepartmentId, string Remark)
        {
            try
            {
                if (ItDate.Length <= 0) throw new SystemException("【異動日期】不能為空!");
                if (DocDate.Length <= 0) throw new SystemException("【單據日期】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
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
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //比對ERP關帳日期
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MA013)) MA013
                                    FROM CMSMA";
                            var cmsmaResult = sqlConnection2.Query(sql, dynamicParameters);
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

                            #region //確認庫存異動單據資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.ConfirmStatus
                                    FROM SCM.InventoryTransaction a 
                                    WHERE a.ItId = @ItId";
                            dynamicParameters.Add("ItId", ItId);

                            var ItResult = sqlConnection.Query(sql, dynamicParameters);

                            if (ItResult.Count() <= 0) throw new SystemException("庫存異動單據資料有誤!!");

                            foreach (var item in ItResult)
                            {
                                if (item.ConfirmStatus != "N") throw new SystemException("單據狀態不可修改!!");
                            }
                            #endregion

                            #region //Update SCM.InventoryTransaction
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.InventoryTransaction SET
                                    ItDate = @ItDate,
                                    DocDate = @DocDate,
                                    DepartmentId = @DepartmentId,
                                    Remark = @Remark,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE ItId = @ItId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    ItDate,
                                    DocDate,
                                    DepartmentId,
                                    Remark,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    ItId
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

        #region //UpdateItDetail -- 更新庫存異動單詳細資料 -- Ann 2023-12-15
        public string UpdateItDetail(int ItDetailId, int ItId, int MtlItemId, string ItMtlItemName, string ItMtlItemSpec
            , double ItQty, int UomId, int InventoryId, int ToInventoryId, string ItRemark, string StorageLocation, string ToStorageLocation)
        {
            try
            {
                int UserId = CreateBy;
                if (ItMtlItemName.Length <= 0) throw new SystemException("【品名】不能為空!");
                if (ItQty <= 0) throw new SystemException("【異動數量】不能為空!");
                if (InventoryId == ToInventoryId && StorageLocation == ToStorageLocation) throw new SystemException("轉出及轉入庫不能相同!!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
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
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //確認庫存異動單據詳細資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM SCM.ItDetail a
                                    WHERE a.ItDetailId = @ItDetailId";
                            dynamicParameters.Add("ItDetailId", ItDetailId);

                            var ItDetailResult = sqlConnection.Query(sql, dynamicParameters);

                            if (ItDetailResult.Count() <= 0) throw new SystemException("庫存異動單據詳細資料有誤!!");
                            #endregion

                            #region //確認庫存異動單據資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.ItErpPrefix, a.DocDate, a.ConfirmStatus
                                    FROM SCM.InventoryTransaction a 
                                    WHERE a.ItId = @ItId";
                            dynamicParameters.Add("ItId", ItId);

                            var ItResult = sqlConnection.Query(sql, dynamicParameters);

                            if (ItResult.Count() <= 0) throw new SystemException("庫存異動單據資料錯誤!!");

                            DateTime DocDate = new DateTime();
                            string ItErpPrefix = "";
                            foreach (var item in ItResult)
                            {
                                if (item.ItErpPrefix == "1201" && ToInventoryId <= 0) throw new SystemException("轉撥單據需填寫轉入庫資訊!!");
                                if (item.ConfirmStatus != "N") throw new SystemException("目前單據狀態不可修改!!");
                                DocDate = item.DocDate;
                                ItErpPrefix = item.ItErpPrefix;
                            }
                            #endregion

                            #region //比對ERP關帳日期
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MA013)) MA013
                                    FROM CMSMA";
                            var cmsmaResult = sqlConnection2.Query(sql, dynamicParameters);
                            if (cmsmaResult.Count() < 0) throw new SystemException("EPR關帳日期資料錯誤!!");

                            foreach (var item in cmsmaResult)
                            {
                                string eprDate = item.MA013;
                                string erpYear = eprDate.Substring(0, 4);
                                string erpMonth = eprDate.Substring(4, 2);
                                string erpDay = eprDate.Substring(6, 2);
                                string erpFullDate = erpYear + "-" + erpMonth + "-" + erpDay;
                                DateTime erpDateTime = Convert.ToDateTime(erpFullDate);
                                int compare = DocDate.CompareTo(erpDateTime);
                                if (compare <= 0) throw new SystemException("ERP已關帳(" + erpFullDate + ")，無法開立此日期(" + DocDate + ")之單據!!");
                            }
                            #endregion

                            #region //判斷單位資料是否有誤
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM PDM.UnitOfMeasure a 
                                    WHERE a.UomId = @UomId";
                            dynamicParameters.Add("UomId", UomId);

                            var UomResult = sqlConnection.Query(sql, dynamicParameters);

                            if (UomResult.Count() <= 0) throw new SystemException("單位資料錯誤!!");
                            #endregion

                            #region //判斷品號資料是否有誤
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT MtlItemNo, TransferStatus
                                    FROM PDM.MtlItem
                                    WHERE MtlItemId = @MtlItemId";
                            dynamicParameters.Add("MtlItemId", MtlItemId);

                            var result3 = sqlConnection.Query(sql, dynamicParameters);
                            if (result3.Count() <= 0) throw new SystemException("品號資料有誤!");

                            string MtlItemNo = "";
                            foreach (var item in result3)
                            {
                                if (item.TransferStatus != "Y") throw new SystemException("此品號尚未拋轉到ERP，無法使用!!");
                                MtlItemNo = item.MtlItemNo;
                            }
                            #endregion

                            #region //確認ERP品號相關資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MB004)) MB004
                                    , LTRIM(RTRIM(MB030)) MB030
                                    , LTRIM(RTRIM(MB031)) MB031
                                    FROM INVMB
                                    WHERE MB001 = @MB001";
                            dynamicParameters.Add("MB001", MtlItemNo);

                            var INVMBResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (INVMBResult.Count() <= 0) throw new SystemException("ERP品號基本資料錯誤!!");

                            foreach (var item in INVMBResult)
                            {
                                #region //判斷ERP品號生效日與失效日
                                if (item.MB030 != "" && item.MB030 != null)
                                {
                                    #region //判斷日期需大於或等於生效日
                                    string EffectiveDate = item.MB030;
                                    string effYear = EffectiveDate.Substring(0, 4);
                                    string effMonth = EffectiveDate.Substring(4, 2);
                                    string effDay = EffectiveDate.Substring(6, 2);
                                    DateTime effFullDate = Convert.ToDateTime(effYear + "-" + effMonth + "-" + effDay);
                                    int effresult = DateTime.Compare(DocDate, effFullDate);
                                    if (effresult < 0) throw new SystemException("此品號尚未生效，無法使用!!");
                                    #endregion
                                }

                                if (item.MB031 != "" && item.MB031 != null)
                                {
                                    #region //判斷日期需小於或等於失效日
                                    string ExpirationDate = item.MB031;
                                    string effYear = ExpirationDate.Substring(0, 4);
                                    string effMonth = ExpirationDate.Substring(4, 2);
                                    string effDay = ExpirationDate.Substring(6, 2);
                                    DateTime effFullDate = Convert.ToDateTime(effYear + "-" + effMonth + "-" + effDay);
                                    int effresult = DateTime.Compare(DocDate, effFullDate);
                                    if (effresult > 0) throw new SystemException("此品號已失效，無法使用!!");
                                    #endregion
                                }
                                #endregion
                            }
                            #endregion

                            #region //判斷轉出庫存資料是否有誤
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 InventoryNo
                                    FROM SCM.Inventory
                                    WHERE InventoryId = @InventoryId";
                            dynamicParameters.Add("InventoryId", InventoryId);

                            var InventoryResult = sqlConnection.Query(sql, dynamicParameters);
                            if (InventoryResult.Count() <= 0) throw new SystemException("轉出庫存資料有誤!");

                            string InventoryNo = "";
                            foreach (var item in InventoryResult)
                            {
                                InventoryNo = item.InventoryNo;
                            }
                            #endregion

                            #region //判斷轉入庫存資料是否有誤
                            string ToInventoryNo = "";
                            if (ToInventoryId > 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 InventoryNo ToInventoryNo
                                        FROM SCM.Inventory
                                        WHERE InventoryId = @InventoryId";
                                dynamicParameters.Add("InventoryId", ToInventoryId);

                                var ToInventoryResult = sqlConnection.Query(sql, dynamicParameters);
                                if (ToInventoryResult.Count() <= 0) throw new SystemException("轉入庫存資料有誤!");

                                foreach (var item in ToInventoryResult)
                                {
                                    ToInventoryNo = item.ToInventoryNo;
                                }
                            }
                            #endregion

                            #region //確認轉出庫別儲位管理資料正確性
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MC009)) LocationFlag
                                    FROM CMSMC
                                    WHERE MC001 = @MC001";
                            dynamicParameters.Add("MC001", InventoryNo);

                            var CMSMCResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (CMSMCResult.Count() <= 0) throw new SystemException("轉出庫別【" + InventoryNo + "】資料錯誤!!");

                            string LocationFlag = "N";
                            foreach (var item in CMSMCResult)
                            {
                                LocationFlag = item.LocationFlag;

                                if (item.LocationFlag == "Y" && StorageLocation.Length <= 0)
                                {
                                    throw new SystemException("轉出庫別【" + InventoryNo + "】需設定儲位!!");
                                }
                                else if (item.LocationFlag == "Y")
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT RTRIM(LTRIM(a.NL002)) NL002
                                            FROM CMSNL a 
                                            WHERE a.NL001 = @InventoryNo
                                            AND a.NL002 = @StorageLocation";
                                    dynamicParameters.Add("InventoryNo", InventoryNo);
                                    dynamicParameters.Add("StorageLocation", StorageLocation);

                                    var CMSNLResult = sqlConnection2.Query(sql, dynamicParameters);

                                    if (CMSNLResult.Count() <= 0) throw new SystemException("儲位【" + StorageLocation + "】資料有誤!!");
                                }
                            }
                            #endregion

                            #region //確認轉入庫別儲位管理資料正確性
                            if (ToInventoryNo.Length > 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT LTRIM(RTRIM(MC009)) LocationFlag
                                        FROM CMSMC
                                        WHERE MC001 = @MC001";
                                dynamicParameters.Add("MC001", ToInventoryNo);

                                var CMSMCResult2 = sqlConnection2.Query(sql, dynamicParameters);

                                if (CMSMCResult2.Count() <= 0) throw new SystemException("轉入庫別【" + ToInventoryNo + "】資料錯誤!!");

                                foreach (var item in CMSMCResult2)
                                {
                                    if (item.LocationFlag == "Y" && ToStorageLocation.Length <= 0)
                                    {
                                        throw new SystemException("轉入庫別【" + ToInventoryNo + "】需設定儲位!!");
                                    }
                                    else if (item.LocationFlag == "Y")
                                    {
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT RTRIM(LTRIM(a.NL002)) NL002
                                                FROM CMSNL a 
                                                WHERE a.NL001 = @InventoryNo
                                                AND a.NL002 = @StorageLocation";
                                        dynamicParameters.Add("InventoryNo", ToInventoryNo);
                                        dynamicParameters.Add("StorageLocation", ToStorageLocation);

                                        var CMSNLResult = sqlConnection2.Query(sql, dynamicParameters);

                                        if (CMSNLResult.Count() <= 0) throw new SystemException("儲位資料有誤!!");
                                    }
                                }
                            }
                            #endregion

                            #region //部門費用領料、轉撥單需檢查庫存量是否足夠
                            if (ItErpPrefix == "1101" || ItErpPrefix == "1201")
                            {
                                #region //檢核基本庫別庫存
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT ISNULL(MC007, 0) MC007
                                        FROM INVMC
                                        WHERE MC001 = @MtlItemNo
                                        AND MC002 = @InventoryNo";
                                dynamicParameters.Add("MtlItemNo", MtlItemNo);
                                dynamicParameters.Add("InventoryNo", InventoryNo);

                                var INVMCResult = sqlConnection2.Query(sql, dynamicParameters);

                                if (INVMCResult.Count() <= 0) throw new SystemException("品號【" + MtlItemNo + "】查無庫別【" + InventoryNo + "】相關資料!!");

                                foreach (var item in INVMCResult)
                                {
                                    if (ItQty > Convert.ToDouble(item.MC007)) throw new SystemException("異動數量已超過庫別【" + InventoryNo + "】庫存量【" + item.MC007 + "】!!");
                                }
                                #endregion

                                #region //檢核儲位庫別庫存
                                if (LocationFlag == "Y")
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT MM005
                                            FROM INVMM
                                            WHERe MM001 = @MtlItemNo
                                            AND MM002 = @InventoryNo
                                            AND MM003 = @StorageLocation";
                                    dynamicParameters.Add("MtlItemNo", MtlItemNo);
                                    dynamicParameters.Add("InventoryNo", InventoryNo);
                                    dynamicParameters.Add("StorageLocation", StorageLocation);

                                    var INVMMResult = sqlConnection2.Query(sql, dynamicParameters);

                                    if (INVMMResult.Count() <= 0) throw new SystemException("品號【" + MtlItemNo + "】庫別【" + InventoryNo + "】儲位【" + StorageLocation + "】庫存量不足!!");

                                    foreach (var item in INVMMResult)
                                    {
                                        if (ItQty > Convert.ToDouble(item.MM005))
                                        {
                                            throw new SystemException("品號【" + MtlItemNo + "】庫別【" + InventoryNo + "】儲位【" + StorageLocation + "】庫存量【" + Convert.ToDouble(item.MM005) + "】不足!!");
                                        }
                                    }
                                }
                                #endregion
                            }
                            #endregion

                            #region //取得本國幣別
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 LTRIM(RTRIM(MA003)) MA003 FROM CMSMA";

                            var CMSMAResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (CMSMAResult.Count() <= 0) throw new SystemException("共用參數檔案資料錯誤!!");

                            string Currency = "";
                            foreach (var item in CMSMAResult)
                            {
                                Currency = item.MA003;
                            }
                            #endregion

                            #region //小數點後取位
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.MF005, a.MF006
                                    FROM CMSMF a 
                                    WHERE a.MF001 = @Currency";
                            dynamicParameters.Add("Currency", Currency);

                            var CMSMFResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (CMSMFResult.Count() <= 0) throw new SystemException("金額小數點取位資料錯誤!!");

                            int unitDecimal = 0;
                            int amountDecimal = 0;
                            foreach (var item in CMSMFResult)
                            {
                                unitDecimal = Convert.ToInt32(item.MF005);
                                amountDecimal = Convert.ToInt32(item.MF006);
                            }
                            #endregion

                            #region //計算單位成本、金額
                            //公式: INVMB 【MB065 庫存金額】/【MB064 庫存數量】
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT MB064, MB065
                                    FROM INVMB
                                    WHERE MB001 = @MtlItemNo";
                            dynamicParameters.Add("MtlItemNo", MtlItemNo);

                            var INVMBResult2 = sqlConnection2.Query(sql, dynamicParameters);

                            if (INVMBResult2.Count() <= 0) throw new SystemException("ERP品號資料錯誤!!");

                            double UnitCost = -1;
                            double Amount = -1;
                            foreach (var item in INVMBResult2)
                            {
                                UnitCost = Convert.ToDouble(Math.Round(item.MB065 / item.MB064, unitDecimal));
                                Amount = Convert.ToDouble(Math.Round(ItQty * UnitCost, amountDecimal));
                            }
                            #endregion

                            #region //Update SCM.ItDetail
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.ItDetail SET
                                    MtlItemId = @MtlItemId,
                                    ItMtlItemName = @ItMtlItemName,
                                    ItMtlItemSpec = @ItMtlItemSpec,
                                    ItQty = @ItQty,
                                    UomId = @UomId,
                                    UnitCost = @UnitCost,
                                    Amount = @Amount,
                                    InventoryId = @InventoryId,
                                    ToInventoryId = @ToInventoryId,
                                    StorageLocation = @StorageLocation,
                                    ToStorageLocation = @ToStorageLocation,
                                    ItRemark = @ItRemark,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE ItDetailId = @ItDetailId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MtlItemId,
                                    ItMtlItemName,
                                    ItMtlItemSpec,
                                    ItQty,
                                    UomId,
                                    UnitCost,
                                    Amount,
                                    InventoryId,
                                    ToInventoryId = ToInventoryId > 0 ? ToInventoryId : (int?)null,
                                    StorageLocation,
                                    ToStorageLocation,
                                    ItRemark,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    ItDetailId
                                });

                            int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //重新計算目前總數量、總金額
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT SUM(a.ItQty) TotalQty, SUM(a.Amount) TotalAmount
                                    FROM SCM.ItDetail a 
                                    WHERE ItId = @ItId";
                            dynamicParameters.Add("ItId", ItId);

                            var ItDetailResult2 = sqlConnection.Query(sql, dynamicParameters);

                            if (ItDetailResult2.Count() <= 0) throw new SystemException("庫存異動單據詳細資料錯誤!!");

                            double TotalQty = 0;
                            double TotalAmount = 0;
                            foreach (var item in ItDetailResult2)
                            {
                                TotalQty = Convert.ToDouble(item.TotalQty);
                                TotalAmount = Math.Round(Convert.ToDouble(item.TotalAmount), amountDecimal);
                            }
                            #endregion

                            #region //Update SCM.InventoryTransaction 總數量、總金額
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.InventoryTransaction SET
                                    TotalQty = @TotalQty,
                                    Amount = @TotalAmount,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE ItId = @ItId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    TotalQty,
                                    TotalAmount,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    ItId
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

        #region //UpdateMaterialRequisition -- 更新領退料單資料 -- Ann 2023-12-19
        public string UpdateMaterialRequisition(int MrId, string MrErpPrefix, string MrDate, string DocDate, string ProductionLine
            , string ContractManufacturer, string Remark)
        {
            try
            {
                if (MrErpPrefix.Length <= 0) throw new SystemException("【ERP領退料單別】不能為空!");
                if (MrDate.Length < 0) throw new SystemException("【領退料日期】不能為空!");
                if (DocDate.Length < 0) throw new SystemException("【單據建立日期】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
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
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //比對ERP關帳日期
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MA013)) MA013
                                    FROM CMSMA";
                            var cmsmaResult = sqlConnection2.Query(sql, dynamicParameters);
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

                            #region //判斷領退料單資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT MrErpPrefix, MrErpNo, ConfirmStatus, ConfirmUserId
                                    FROM MES.MaterialRequisition
                                    WHERE MrId = @MrId";
                            dynamicParameters.Add("MrId", MrId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("領退料單資料錯誤!");

                            string MrErpNo = "";
                            string ConfirmStatus = "";
                            int ConfirmUserId = -1;
                            foreach (var item in result)
                            {
                                if (item.ConfirmStatus != "N") throw new SystemException("領退料單已確認，無法更改!");
                                MrErpNo = item.MrErpNo;
                                ConfirmStatus = item.ConfirmStatus;

                                if (item.ConfirmUserId == null) ConfirmUserId = -1;
                                else ConfirmUserId = item.ConfirmUserId;
                            }
                            #endregion

                            #region //若已拋轉ERP，判斷單據目前狀態
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(TC009)) TC009
                                    FROM MOCTC 
                                    WHERE TC001 = @MrErpPrefix
                                    AND TC002 = @MrErpNo";
                            dynamicParameters.Add("MrErpPrefix", MrErpPrefix);
                            dynamicParameters.Add("MrErpNo", MrErpNo);

                            var MOCTCResult = sqlConnection2.Query(sql, dynamicParameters);

                            foreach (var item in MOCTCResult)
                            {
                                if (item.TC009 != "N") throw new SystemException("ERP單據狀態不可修改!!");
                            }
                            #endregion

                            #region //UPDATE MES.MaterialRequisition
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE MES.MaterialRequisition SET
                                    MrDate = @MrDate,
                                    DocDate = @DocDate,
                                    ProductionLine = @ProductionLine,
                                    Remark = @Remark,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE MrId = @MrId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MrDate,
                                    DocDate,
                                    ProductionLine,
                                    Remark,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    MrId
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

        #region //UpdateMrWipOrder -- 更新領退料單設定資料(重計單身) -- Ann 2023-12-20
        public string UpdateMrWipOrder(int MrId, int MoId, string WoErpPrefix, string WoErpNo, string MrType, double Quantity, string RequisitionCode, string NegativeStatus, string Remark
            , string MaterialCategory, string SubinventoryType, string LineSeq, int InventoryId, string StorageLocation)
        {
            try
            {
                if (WoErpPrefix.Length < 0) throw new SystemException("【製令單別】不能為空!");
                if (WoErpPrefix.Length > 4) throw new SystemException("【製令單別】長度錯誤!");
                if (WoErpNo.Length < 0) throw new SystemException("【製令單號】不能為空!");
                if (MrType.Length < 0) throw new SystemException("【領料方式】長度錯誤!");
                if (Quantity < 0) throw new SystemException("【領退料套數】不能小於0!");
                if (RequisitionCode.Length < 0) throw new SystemException("【領料碼】不能為空!");
                if (NegativeStatus.Length < 0) throw new SystemException("【庫存不足照領】不能為空!");
                if (Remark.Length > 255) throw new SystemException("【備註】長度錯誤!");
                if (MaterialCategory.Length < 0) throw new SystemException("【材料型態】不能為空!");
                if (SubinventoryType.Length < 0) throw new SystemException("【庫別型態】不能為空!");
                if (LineSeq.Length < 0) throw new SystemException("【輸入序號】不能為空!");
                if (LineSeq.Length > 4) throw new SystemException("【輸入序號】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
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
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //判斷領退料單資料是否有誤
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT ConfirmStatus, MrErpPrefix, MrErpNo
                                    FROM MES.MaterialRequisition
                                    WHERE CompanyId = @CompanyId
                                    AND MrId = @MrId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("MrId", MrId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("領退料單資料有誤!");

                            string MrErpPrefix = "";
                            string MrErpNo = "";
                            foreach (var item in result)
                            {
                                if (item.ConfirmStatus != "N") throw new SystemException("領退料單已確認，無法更改!");
                                MrErpPrefix = item.MrErpPrefix;
                                MrErpNo = item.MrErpNo;
                            }
                            #endregion

                            #region //查詢單據類別設定資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.MQ010, LTRIM(RTRIM(a.MQ054)) MQ054
                                    FROM CMSMQ a 
                                    WHERE MQ001 = @MrErpPrefix";
                            dynamicParameters.Add("MrErpPrefix", MrErpPrefix);

                            var CMSMQResult = sqlConnection2.Query(sql, dynamicParameters);
                            if (CMSMQResult.Count() <= 0) throw new SystemException("ERP單別資料有誤!!");

                            int MQ010 = -1;
                            string ExcessFlag = "";
                            foreach (var item in CMSMQResult)
                            {
                                MQ010 = Convert.ToInt32(item.MQ010);
                                ExcessFlag = item.MQ054;
                            }
                            #endregion

                            #region //確認製令資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.Quantity
                                    , c.InventoryId, c.InventoryUomId UomId
                                    , d.InventoryNo
                                    , e.UomNo
                                    FROM MES.ManufactureOrder a
                                    INNER JOIN MES.WipOrder b ON a.WoId = b.WoId
                                    INNER JOIN PDM.MtlItem c ON b.MtlItemId = c.MtlItemId
                                    INNER JOIN SCM.Inventory d ON c.InventoryId = d.InventoryId
                                    INNER JOIN PDM.UnitOfMeasure e ON c.InventoryUomId = e.UomId
                                    WHERE a.MoId = @MoId";
                            dynamicParameters.Add("MoId", MoId);

                            var ManufactureOrderResult = sqlConnection.Query(sql, dynamicParameters);

                            if (ManufactureOrderResult.Count() <= 0) throw new SystemException("製令資料錯誤!!");

                            int MoQuantity = 0;
                            int UomId = -1;
                            string UomNo = "";
                            foreach (var item in ManufactureOrderResult)
                            {
                                MoQuantity = item.Quantity;
                                UomId = item.UomId;
                                UomNo = item.UomNo;
                            }
                            #endregion

                            #region //查詢ERP製令資訊
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(a.TA015)) PlanQty, LTRIM(RTRIM(a.TA016)) RequisitionSetQty
                                    , LTRIM(RTRIM(a.TA017)) StockInQty
                                    FROM MOCTA a
                                    WHERE a.TA001 = @WoErpPrefix
                                    AND a.TA002 = @WoErpNo";
                            dynamicParameters.Add("WoErpPrefix", WoErpPrefix);
                            dynamicParameters.Add("WoErpNo", WoErpNo);

                            var ErpWoInfoResult = sqlConnection2.Query(sql, dynamicParameters);
                            if (ErpWoInfoResult.Count() <= 0) throw new SystemException("ERP製令(MOCTA)資料有誤!");

                            double PlanQty = -1;
                            double RequisitionSetQty = -1;
                            double StockInQty = -1;
                            foreach (var item in ErpWoInfoResult)
                            {
                                PlanQty = Convert.ToDouble(item.PlanQty);
                                RequisitionSetQty = Convert.ToDouble(item.RequisitionSetQty);
                                StockInQty = Convert.ToDouble(item.StockInQty);
                            }
                            #endregion

                            #region //搜尋此製令所有領料單資訊，並找出實際已核單領料套數
                            int ErpQuantity = 0;
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT b.MrErpPrefix, b.MrErpNo
                                    , c.Quantity
                                    , d.WoErpPrefix, d.WoErpNo
                                    FROM MES.MrWipOrder a
                                    INNER JOIN MES.MaterialRequisition b ON a.MrId = b.MrId
                                    INNER JOIN MES.ManufactureOrder c ON a.MoId = c.MoId
                                    INNER JOIN MES.WipOrder d ON c.WoId = d.WoId
                                    WHERE a.MoId = @MoId";
                            dynamicParameters.Add("MoId", MoId);

                            var MrWipOrderResult = sqlConnection.Query(sql, dynamicParameters);

                            if (MrWipOrderResult.Count() > 0)
                            {
                                foreach (var item2 in MrWipOrderResult)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.TD006 Quantity
                                            , b.MQ010
                                            , (
                                                SELECT TOP 1 1
                                                FROM MOCTE x
                                                WHERE x.TE001 = a.TD001
                                                AND x.TE002 = a.TD002
                                                AND x.TE019 = 'Y'
                                            ) CheckConfirm
                                            FROM MOCTD a
                                            INNER JOIN CMSMQ b ON a.TD001 = b.MQ001
                                            WHERE a.TD001 = @MrErpPrefix
                                            AND a.TD002 = @MrErpNo
                                            AND a.TD003 = @WoErpPrefix
                                            AND a.TD004 = @WoErpNo";
                                    dynamicParameters.Add("MrErpPrefix", item2.MrErpPrefix);
                                    dynamicParameters.Add("MrErpNo", item2.MrErpNo);
                                    dynamicParameters.Add("WoErpPrefix", item2.WoErpPrefix);
                                    dynamicParameters.Add("WoErpNo", item2.WoErpNo);

                                    var ErpMrWipOrderResult = sqlConnection2.Query(sql, dynamicParameters);

                                    foreach (var item3 in ErpMrWipOrderResult)
                                    {
                                        if (item3.CheckConfirm != null)
                                        {
                                            if (item3.MQ010 > 0)
                                            {
                                                ErpQuantity -= item3.Quantity;
                                            }
                                            else
                                            {
                                                ErpQuantity += item3.Quantity;
                                            }
                                        }
                                    }
                                }
                            }
                            #endregion

                            #region //判斷MES製令資料是否有誤
                            if (MQ010 < 0) //領料
                            {
                                if (Quantity > (MoQuantity - ErpQuantity) && ExcessFlag == "Y") throw new SystemException("領取數量超過未領用量(" + (MoQuantity - ErpQuantity).ToString() + ")!");
                            }
                            else if (MQ010 > 0) //退料
                            {
                                if (Quantity > ErpQuantity) throw new SystemException("退料數量超過剩餘已領套數(" + ErpQuantity.ToString() + ")!");
                            }
                            else
                            {
                                throw new SystemException("此領料單別(" + MrErpPrefix + ")設定資料有誤，請通知開發人員!!");
                            }
                            #endregion

                            #region //確認庫別資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.InventoryNo
                                    FROM SCM.Inventory a 
                                    WHERE a.InventoryId = @InventoryId";
                            dynamicParameters.Add("InventoryId", InventoryId);

                            var InventoryResult = sqlConnection.Query(sql, dynamicParameters);

                            if (InventoryResult.Count() <= 0) throw new SystemException("庫別資料錯誤!!");

                            string InventoryNo = "";
                            foreach (var item in InventoryResult)
                            {
                                InventoryNo = item.InventoryNo;
                            }
                            #endregion

                            #region //確認庫別儲位管理資料正確性
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MC009)) LocationFlag
                                    FROM CMSMC
                                    WHERE MC001 = @MC001";
                            dynamicParameters.Add("MC001", InventoryNo);

                            var CMSMCResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (CMSMCResult.Count() <= 0) throw new SystemException("轉出庫別【" + InventoryNo + "】資料錯誤!!");

                            string LocationFlag = "N";
                            foreach (var item in CMSMCResult)
                            {
                                LocationFlag = item.LocationFlag;

                                if (item.LocationFlag == "Y" && StorageLocation.Length <= 0)
                                {
                                    throw new SystemException("轉出庫別【" + InventoryNo + "】需設定儲位!!");
                                }
                                else if (item.LocationFlag == "Y")
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT RTRIM(LTRIM(a.NL002)) NL002
                                            FROM CMSNL a 
                                            WHERE a.NL001 = @InventoryNo
                                            AND a.NL002 = @StorageLocation";
                                    dynamicParameters.Add("InventoryNo", InventoryNo);
                                    dynamicParameters.Add("StorageLocation", StorageLocation);

                                    var CMSNLResult = sqlConnection2.Query(sql, dynamicParameters);

                                    if (CMSNLResult.Count() <= 0) throw new SystemException("儲位【" + StorageLocation + "】資料有誤!!");
                                }
                            }
                            #endregion

                            #region //UPDATE MES.MrWipOrder
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE MES.MrWipOrder SET
                                    WoErpPrefix = @WoErpPrefix,
                                    WoErpNo = @WoErpNo,
                                    MrType = @MrType,
                                    Quantity = @Quantity,
                                    InventoryId = @InventoryId,
                                    StorageLocation = @StorageLocation,
                                    RequisitionCode = @RequisitionCode,
                                    NegativeStatus = @NegativeStatus,
                                    Remark = @Remark,
                                    MaterialCategory = @MaterialCategory,
                                    SubinventoryType = @SubinventoryType,
                                    LineSeq = @LineSeq,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE MrId = @MrId
                                    AND MoId = @MoId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    WoErpPrefix,
                                    WoErpNo,
                                    MrType,
                                    Quantity,
                                    InventoryId,
                                    StorageLocation,
                                    InventoryNo,
                                    RequisitionCode,
                                    NegativeStatus,
                                    Remark,
                                    MaterialCategory,
                                    SubinventoryType,
                                    LineSeq,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    MrId,
                                    MoId
                                });

                            int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //刪除原本單身
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE a 
                                    FROM MES.MrBarcodeRegister a
                                    INNER JOIN MES.MrDetail b ON a.MrDetailId = b.MrDetailId
                                    WHERE b.MrId = @MrId
                                    AND b.MoId = @MoId";
                            dynamicParameters.Add("MrId", MrId);
                            dynamicParameters.Add("MoId", MoId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE a 
                                    FROM MES.MrDetail a
                                    WHERE a.MrId = @MrId
                                    AND a.MoId = @MoId";
                            dynamicParameters.Add("MrId", MrId);
                            dynamicParameters.Add("MoId", MoId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //查看目前MrDetail序號
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT MAX(MrSequence) MrSequence
                                    FROM MES.MrDetail
                                    WHERE MrId = @MrId";
                            dynamicParameters.Add("MrId", MrId);

                            var result5 = sqlConnection.Query(sql, dynamicParameters);
                            int currentMrSequence = -1;
                            foreach (var item in result5)
                            {
                                currentMrSequence = Convert.ToInt32(item.MrSequence);
                            }

                            string MrSequence = (currentMrSequence + 1).ToString().PadLeft(4, '0');
                            #endregion

                            #region //同步更新MrDetail
                            #region //取得MES製令用料設定
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT ISNULL(a.CompositionQuantity, 1) CompositionQuantity, ISNULL(a.Base, 1) Base, ISNULL(a.LossRate, 0) LossRate, a.BarcodeCtrl, a.MainBarcode, a.ControlType, a.DecompositionFlag
                                    , b.MtlItemId, b.DemandRequisitionQty, b.RequisitionQty, b.SubstituteMtlItemId, b.MaterialProperties
                                    , c.WoErpPrefix, c.WoErpNo, c.PlanQty, c.StockInQty
                                    , d.UomId, d.UomNo
                                    , e.MtlItemNo, e.MtlItemName
                                    , f.InventoryId, f.InventoryNo
                                    , g.MtlItemNo SubstitutionMtlItemNo, g.MtlItemName SubstitutionMtlItemName
                                    , h.TypeName
                                    , i.MtlItemNo
                                    FROM MES.MoMtlSetting a
                                    INNER JOIN MES.WoDetail b ON a.WoDetailId = b.WoDetailId
                                    INNER JOIN MES.WipOrder c ON b.WoId = c.WoId
                                    INNER JOIN PDM.UnitOfMeasure d ON a.UomId = d.UomId
                                    INNER JOIN PDM.MtlItem e ON b.MtlItemId = e.MtlItemId
                                    INNER JOIN SCM.Inventory f ON b.InventoryId = f.InventoryId
                                    INNER JOIN PDM.MtlItem g ON b.SubstituteMtlItemId = g.MtlItemId
                                    INNER JOIN BAS.[Type] h ON a.ControlType = h.TypeNo AND h.TypeSchema = 'MoMtlSetting.ControlType'
                                    INNER JOIN PDM.MtlItem i ON b.MtlItemId = i.MtlItemId
                                    WHERE a.MoId = @MoId
                                    ORDER BY a.MainBarcode DESC";
                            dynamicParameters.Add("MoId", MoId);

                            var result4 = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item in result4)
                            {
                                #region //依據材料性質判斷要帶出的用料MaterialCategory
                                string MaterialProperties = "";
                                if (item.MaterialProperties == null)
                                {
                                    throw new SystemException("材料性質設定錯誤，請確認工單資料是否正確!!");
                                }
                                else
                                {
                                    if (item.MaterialProperties != "1" && item.MaterialProperties != "2" && item.MaterialProperties != "5") continue;
                                    MaterialProperties = item.MaterialProperties;
                                }

                                if (MaterialCategory != "1" && MaterialCategory != "*" && MaterialCategory != "2" && MaterialCategory != "5")
                                {
                                    MaterialCategory = "*";
                                }

                                if (MaterialCategory != "*")
                                {
                                    if (MaterialCategory != MaterialProperties) continue;
                                }
                                #endregion

                                #region //取得ERP製令單身需領/已領用量
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TB004 DemandRequisitionQty, TB005 RequisitionQty
                                        FROM MOCTB
                                        WHERE TB001 = @WoErpPrefix
                                        AND TB002 = @WoErpNo
                                        AND TB003 = @MtlItemNo";
                                dynamicParameters.Add("WoErpPrefix", WoErpPrefix);
                                dynamicParameters.Add("WoErpNo", WoErpNo);
                                dynamicParameters.Add("MtlItemNo", item.MtlItemNo);

                                var ErpWoDetailResult = sqlConnection2.Query(sql, dynamicParameters);

                                if (ErpWoDetailResult.Count() <= 0) throw new SystemException("查無ERP製令單身資料!!");

                                double DemandRequisitionQty = -1;
                                double RequisitionQty = -1;
                                foreach (var item2 in ErpWoDetailResult)
                                {
                                    DemandRequisitionQty = Convert.ToDouble(item2.DemandRequisitionQty);
                                    RequisitionQty = Convert.ToDouble(item2.RequisitionQty);
                                }
                                #endregion

                                if (MQ010 > 0)
                                {
                                    if (RequisitionQty <= 0) continue;
                                }

                                #region //若為補足需領量，從EPR查是否有剩餘的料要領
                                if (MrType == "2" || MrType == "3")
                                {
                                    if (DemandRequisitionQty - RequisitionQty <= 0) continue;
                                }
                                #endregion

                                #region //若為退已領用量，從ERP查是否有已領量
                                if (MrType == "5")
                                {
                                    if (RequisitionQty <= 0) continue;
                                }
                                #endregion

                                #region //計算一套可以領/退多少料
                                //領料公式: 未領用量 / 總套數(預計生產量) / 單位用量 / 底數
                                //退料公式: 剩餘生產量 / 單位用量
                                double MrDetailQty = -1; //預計領料量
                                double ActualQuantity = -1; //實際領料量
                                if (item.BomDetailId != null)
                                {
                                    double unitQuantity = (DemandRequisitionQty - RequisitionQty) / item.PlanQty / item.CompositionQuantity;
                                    if (unitQuantity * Quantity * item.CompositionQuantity > (DemandRequisitionQty - RequisitionQty)) MrDetailQty = DemandRequisitionQty - RequisitionQty;
                                    else MrDetailQty = unitQuantity * Quantity * item.CompositionQuantity;
                                }
                                else
                                {
                                    switch (MrType)
                                    {
                                        //確認公式是否為: ((領料數量*單位用量)/基數) + ((領料數量*單位用量/基數)*損耗率) 再四捨五入
                                        case "1":
                                            //MrDetailQty = ((Quantity * Convert.ToDouble(item.CompositionQuantity)) / Convert.ToDouble(item.Base)) + (((Quantity * Convert.ToDouble(item.CompositionQuantity)) / Convert.ToDouble(item.Base)) * item.LossRate);
                                            MrDetailQty = (Quantity * Convert.ToDouble(item.CompositionQuantity)) / Convert.ToDouble(item.Base);
                                            break;
                                        case "2":
                                            MrDetailQty = DemandRequisitionQty - RequisitionQty;
                                            break;
                                        case "3":
                                            MrDetailQty = DemandRequisitionQty - RequisitionQty;
                                            break;
                                        case "4":
                                            MrDetailQty = (Quantity * item.CompositionQuantity) / item.Base;
                                            break;
                                        case "5":
                                            MrDetailQty = RequisitionQty;
                                            break;
                                    }
                                }

                                if (item.DecompositionFlag == "Y")
                                {
                                    ActualQuantity = MrDetailQty;
                                }
                                else if (item.BarcodeCtrl == "Y" && item.UomNo == "PCS")
                                {
                                    ActualQuantity = 0;
                                }
                                else
                                {
                                    ActualQuantity = MrDetailQty;
                                }
                                #endregion

                                #region //查詢目前此庫別庫存量，根據庫存是否足夠進行不同流程
                                string DetailDesc = "";

                                double InventoryQty = 0;
                                dynamicParameters = new DynamicParameters();
                                if (LocationFlag == "Y")
                                {
                                    #region //庫別需儲位管理
                                    sql = @"SELECT a.MM005 InventoryQty
                                            FROM INVMM a
                                            WHERE a.MM001 = @MM001
                                            AND a.MM002 = @MM002
                                            AND a.MM003 = @MM003";
                                    dynamicParameters.Add("MM001", item.MtlItemNo);
                                    dynamicParameters.Add("MM002", InventoryNo);
                                    dynamicParameters.Add("MM003", StorageLocation);
                                    #endregion
                                }
                                else
                                {
                                    sql = @"SELECT a.MC007 InventoryQty
                                            FROM INVMC a
                                            WHERE a.MC001 = @MC001
                                            AND a.MC002 = @MC002";
                                    dynamicParameters.Add("MC001", item.MtlItemNo);
                                    dynamicParameters.Add("MC002", InventoryNo);
                                }

                                var InventoryResult2 = sqlConnection2.Query(sql, dynamicParameters);

                                if (InventoryResult2.Count() > 0)
                                {
                                    foreach (var item2 in InventoryResult2)
                                    {
                                        InventoryQty = Convert.ToDouble(item2.InventoryQty);

                                        #region //計算同料號同庫別，但未確認之領料單領料數量
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT ISNULL(SUM(a.TE005), 0) TE005
                                                FROM MOCTE a
                                                WHERE a.TE004 = @TE004
                                                AND a.TE008 = @TE008
                                                AND a.TE019 = 'N'
                                                AND a.TE001 != @TE001
                                                AND a.TE002 != @TE002";
                                        dynamicParameters.Add("TE004", item.MtlItemNo);
                                        dynamicParameters.Add("TE008", InventoryNo);
                                        dynamicParameters.Add("TE001", MrErpPrefix);
                                        dynamicParameters.Add("TE002", MrErpNo);

                                        var MtlItemQtyResult = sqlConnection2.Query(sql, dynamicParameters);

                                        double? TotalMtlItemQty = 0;
                                        foreach (var item3 in MtlItemQtyResult)
                                        {
                                            TotalMtlItemQty = Convert.ToDouble(item3.TE005);
                                        }
                                        #endregion

                                        #region //尋找本身領料單內是否有同品號數量
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT ISNULL(SUM(a.Quantity), 0) TE005
                                                FROM MES.MrDetail a
                                                WHERE a.MrId = @MrId
                                                AND a.MtlItemId = @MtlItemId
                                                AND a.InventoryId = @InventoryId";
                                        dynamicParameters.Add("MrId", MrId);
                                        dynamicParameters.Add("MtlItemId", item.MtlItemId);
                                        dynamicParameters.Add("InventoryId", InventoryId);

                                        var ThisMtlItemQtyResult = sqlConnection.Query(sql, dynamicParameters);

                                        foreach (var item3 in ThisMtlItemQtyResult)
                                        {
                                            TotalMtlItemQty += Convert.ToDouble(item3.TE005);
                                        }
                                        #endregion

                                        double? CurrentMtlItemQty = Convert.ToDouble(InventoryQty);
                                        double? AvailabilityQty = CurrentMtlItemQty - TotalMtlItemQty/* - Convert.ToDouble(item.Quantity)*/;
                                        if (AvailabilityQty >= 0)
                                        {
                                            DetailDesc = "可用量: " + AvailabilityQty;
                                        }
                                        else
                                        {
                                            DetailDesc = "庫存不足，需領用量: " + item.Quantity;
                                        }
                                    }
                                }
                                else
                                {
                                    DetailDesc = "庫存不足，需領用量: " + item.Quantity;
                                }
                                #endregion

                                #region //INSERT MES.MrDetail
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO MES.MrDetail (MrId, MrSequence, MtlItemId, Quantity, ActualQuantity, UomId, Unit
                                        , InventoryId, SubInventoryCode, ProcessCode, LotNumber, MoId, WoErpPrefix, WoErpNo
                                        , DetailDesc, Remark, MaterialCategory, ConfirmStatus, ProjectCode, BondedStatus
                                        , SubstituteStatus, OfficialItemStatus, SubstitutionId, SubstitutionMtlItemNo, SubstituteProcessCode
                                        , SubstituteQty, SubstituteRate, StorageLocation
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.MrDetailId
                                        VALUES (@MrId, @MrSequence, @MtlItemId, @Quantity, @ActualQuantity, @UomId, @Unit
                                        , @InventoryId, @SubInventoryCode, @ProcessCode, @LotNumber, @MoId, @WoErpPrefix, @WoErpNo
                                        , @DetailDesc, @Remark, @MaterialCategory, @ConfirmStatus, @ProjectCode, @BondedStatus
                                        , @SubstituteStatus, @OfficialItemStatus, @SubstitutionId, @SubstitutionMtlItemNo, @SubstituteProcessCode
                                        , @SubstituteQty, @SubstituteRate, @StorageLocation
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        MrId,
                                        MrSequence,
                                        item.MtlItemId,
                                        Quantity = Math.Round(MrDetailQty * 1000) / 1000,
                                        ActualQuantity = 0,
                                        item.UomId,
                                        Unit = item.UomNo,
                                        InventoryId,
                                        SubInventoryCode = item.InventoryNo,
                                        ProcessCode = "****",
                                        LotNumber = "",
                                        MoId,
                                        WoErpPrefix,
                                        WoErpNo,
                                        DetailDesc,
                                        Remark = ExcessFlag == "N" ? Remark : "",
                                        MaterialCategory = MaterialProperties,
                                        ConfirmStatus = "N",
                                        ProjectCode = "",
                                        BondedStatus = "N",
                                        SubstituteStatus = "N",
                                        OfficialItemStatus = "2",
                                        SubstitutionId = item.SubstitutionId != null ? (int?)item.SubstituteMtlItemId : null,
                                        item.SubstitutionMtlItemNo,
                                        SubstituteProcessCode = "****",
                                        SubstituteQty = 0,
                                        SubstituteRate = 0,
                                        StorageLocation,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });

                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();

                                int nextMrSequence = Convert.ToInt32(MrSequence) + 1;
                                MrSequence = nextMrSequence.ToString().PadLeft(4, '0');
                                #endregion
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

        #region //UpdateMrWipOrderNoRC -- 更新領退料單設定資料(不重計單身) -- Ann 2023-12-20
        public string UpdateMrWipOrderNoRC(int MrId, int MoId, string WoErpPrefix, string WoErpNo, string MrType, double Quantity, string RequisitionCode, string NegativeStatus, string Remark
            , string MaterialCategory, string SubinventoryType
            , string LineSeq, int InventoryId, string StorageLocation)
        {
            try
            {
                if (WoErpPrefix.Length < 0) throw new SystemException("【製令單別】不能為空!");
                if (WoErpPrefix.Length > 4) throw new SystemException("【製令單別】長度錯誤!");
                if (WoErpNo.Length < 0) throw new SystemException("【製令單號】不能為空!");
                if (MrType.Length < 0) throw new SystemException("【領料方式】長度錯誤!");
                if (Quantity < 0) throw new SystemException("【領退料套數】不能小於0!");
                if (RequisitionCode.Length < 0) throw new SystemException("【領料碼】不能為空!");
                if (NegativeStatus.Length < 0) throw new SystemException("【庫存不足照領】不能為空!");
                if (Remark.Length > 255) throw new SystemException("【備註】長度錯誤!");
                if (MaterialCategory.Length < 0) throw new SystemException("【材料型態】不能為空!");
                if (SubinventoryType.Length < 0) throw new SystemException("【庫別型態】不能為空!");
                if (LineSeq.Length < 0) throw new SystemException("【輸入序號】不能為空!");
                if (LineSeq.Length > 4) throw new SystemException("【輸入序號】長度錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
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
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //判斷領退料單資料是否有誤
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT ConfirmStatus, MrErpPrefix, MrErpNo
                                    FROM MES.MaterialRequisition
                                    WHERE CompanyId = @CompanyId
                                    AND MrId = @MrId";
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            dynamicParameters.Add("MrId", MrId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("領退料單資料有誤!");

                            string MrErpPrefix = "";
                            string MrErpNo = "";
                            foreach (var item in result)
                            {
                                if (item.ConfirmStatus != "N") throw new SystemException("領退料單已確認，無法更改!");
                                MrErpPrefix = item.MrErpPrefix;
                                MrErpNo = item.MrErpNo;
                            }
                            #endregion

                            #region //查詢單據類別設定資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.MQ010, LTRIM(RTRIM(a.MQ054)) MQ054
                                    FROM CMSMQ a 
                                    WHERE MQ001 = @MrErpPrefix";
                            dynamicParameters.Add("MrErpPrefix", MrErpPrefix);

                            var CMSMQResult = sqlConnection2.Query(sql, dynamicParameters);
                            if (CMSMQResult.Count() <= 0) throw new SystemException("ERP單別資料有誤!!");

                            int MQ010 = -1;
                            string ExcessFlag = "";
                            foreach (var item in CMSMQResult)
                            {
                                MQ010 = Convert.ToInt32(item.MQ010);
                                ExcessFlag = item.MQ054;
                            }
                            #endregion

                            #region //確認製令資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.Quantity
                                    , c.InventoryId, c.InventoryUomId UomId
                                    , d.InventoryNo
                                    , e.UomNo
                                    FROM MES.ManufactureOrder a
                                    INNER JOIN MES.WipOrder b ON a.WoId = b.WoId
                                    INNER JOIN PDM.MtlItem c ON b.MtlItemId = c.MtlItemId
                                    INNER JOIN SCM.Inventory d ON c.InventoryId = d.InventoryId
                                    INNER JOIN PDM.UnitOfMeasure e ON c.InventoryUomId = e.UomId
                                    WHERE a.MoId = @MoId";
                            dynamicParameters.Add("MoId", MoId);

                            var ManufactureOrderResult = sqlConnection.Query(sql, dynamicParameters);

                            if (ManufactureOrderResult.Count() <= 0) throw new SystemException("製令資料錯誤!!");

                            int MoQuantity = 0;
                            int UomId = -1;
                            string UomNo = "";
                            foreach (var item in ManufactureOrderResult)
                            {
                                MoQuantity = item.Quantity;
                                UomId = item.UomId;
                                UomNo = item.UomNo;
                            }
                            #endregion

                            #region //查詢ERP製令資訊
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(a.TA015)) PlanQty, LTRIM(RTRIM(a.TA016)) RequisitionSetQty
                                    , LTRIM(RTRIM(a.TA017)) StockInQty
                                    FROM MOCTA a
                                    WHERE a.TA001 = @WoErpPrefix
                                    AND a.TA002 = @WoErpNo";
                            dynamicParameters.Add("WoErpPrefix", WoErpPrefix);
                            dynamicParameters.Add("WoErpNo", WoErpNo);

                            var ErpWoInfoResult = sqlConnection2.Query(sql, dynamicParameters);
                            if (ErpWoInfoResult.Count() <= 0) throw new SystemException("ERP製令(MOCTA)資料有誤!");

                            double PlanQty = -1;
                            double RequisitionSetQty = -1;
                            double StockInQty = -1;
                            foreach (var item in ErpWoInfoResult)
                            {
                                PlanQty = Convert.ToDouble(item.PlanQty);
                                RequisitionSetQty = Convert.ToDouble(item.RequisitionSetQty);
                                StockInQty = Convert.ToDouble(item.StockInQty);
                            }
                            #endregion

                            #region //搜尋此製令所有領料單資訊，並找出實際已核單領料套數
                            int ErpQuantity = 0;
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT b.MrErpPrefix, b.MrErpNo
                                    , c.Quantity
                                    , d.WoErpPrefix, d.WoErpNo
                                    FROM MES.MrWipOrder a
                                    INNER JOIN MES.MaterialRequisition b ON a.MrId = b.MrId
                                    INNER JOIN MES.ManufactureOrder c ON a.MoId = c.MoId
                                    INNER JOIN MES.WipOrder d ON c.WoId = d.WoId
                                    WHERE a.MoId = @MoId";
                            dynamicParameters.Add("MoId", MoId);

                            var MrWipOrderResult = sqlConnection.Query(sql, dynamicParameters);

                            if (MrWipOrderResult.Count() > 0)
                            {
                                foreach (var item2 in MrWipOrderResult)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.TD006 Quantity
                                            , b.MQ010
                                            , (
                                                SELECT TOP 1 1
                                                FROM MOCTE x
                                                WHERE x.TE001 = a.TD001
                                                AND x.TE002 = a.TD002
                                                AND x.TE019 = 'Y'
                                            ) CheckConfirm
                                            FROM MOCTD a
                                            INNER JOIN CMSMQ b ON a.TD001 = b.MQ001
                                            WHERE a.TD001 = @MrErpPrefix
                                            AND a.TD002 = @MrErpNo
                                            AND a.TD003 = @WoErpPrefix
                                            AND a.TD004 = @WoErpNo";
                                    dynamicParameters.Add("MrErpPrefix", item2.MrErpPrefix);
                                    dynamicParameters.Add("MrErpNo", item2.MrErpNo);
                                    dynamicParameters.Add("WoErpPrefix", item2.WoErpPrefix);
                                    dynamicParameters.Add("WoErpNo", item2.WoErpNo);

                                    var ErpMrWipOrderResult = sqlConnection2.Query(sql, dynamicParameters);

                                    foreach (var item3 in ErpMrWipOrderResult)
                                    {
                                        if (item3.CheckConfirm != null)
                                        {
                                            if (item3.MQ010 > 0)
                                            {
                                                ErpQuantity -= item3.Quantity;
                                            }
                                            else
                                            {
                                                ErpQuantity += item3.Quantity;
                                            }
                                        }
                                    }
                                }
                            }
                            #endregion

                            #region //判斷MES製令資料是否有誤
                            if (MQ010 < 0) //領料
                            {
                                if (Quantity > (MoQuantity - ErpQuantity) && ExcessFlag == "Y") throw new SystemException("領取數量超過未領用量(" + (MoQuantity - ErpQuantity).ToString() + ")!");
                            }
                            else if (MQ010 > 0) //退料
                            {
                                if (Quantity > ErpQuantity) throw new SystemException("退料數量超過剩餘已領套數(" + ErpQuantity.ToString() + ")!");
                            }
                            else
                            {
                                throw new SystemException("此領料單別(" + MrErpPrefix + ")設定資料有誤，請通知開發人員!!");
                            }
                            #endregion


                            #region //確認庫別資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.InventoryNo
                                    FROM SCM.Inventory a 
                                    WHERE a.InventoryId = @InventoryId";
                            dynamicParameters.Add("InventoryId", InventoryId);

                            var InventoryResult = sqlConnection.Query(sql, dynamicParameters);

                            if (InventoryResult.Count() <= 0) throw new SystemException("庫別資料錯誤!!");

                            string InventoryNo = "";
                            foreach (var item in InventoryResult)
                            {
                                InventoryNo = item.InventoryNo;
                            }
                            #endregion

                            #region //確認庫別儲位管理資料正確性
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MC009)) LocationFlag
                                    FROM CMSMC
                                    WHERE MC001 = @MC001";
                            dynamicParameters.Add("MC001", InventoryNo);

                            var CMSMCResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (CMSMCResult.Count() <= 0) throw new SystemException("轉出庫別【" + InventoryNo + "】資料錯誤!!");

                            string LocationFlag = "N";
                            foreach (var item in CMSMCResult)
                            {
                                LocationFlag = item.LocationFlag;

                                if (item.LocationFlag == "Y" && StorageLocation.Length <= 0)
                                {
                                    throw new SystemException("轉出庫別【" + InventoryNo + "】需設定儲位!!");
                                }
                                else if (item.LocationFlag == "Y")
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT RTRIM(LTRIM(a.NL002)) NL002
                                            FROM CMSNL a 
                                            WHERE a.NL001 = @InventoryNo
                                            AND a.NL002 = @StorageLocation";
                                    dynamicParameters.Add("InventoryNo", InventoryNo);
                                    dynamicParameters.Add("StorageLocation", StorageLocation);

                                    var CMSNLResult = sqlConnection2.Query(sql, dynamicParameters);

                                    if (CMSNLResult.Count() <= 0) throw new SystemException("儲位【" + StorageLocation + "】資料有誤!!");
                                }
                            }
                            #endregion

                            #region //UPDATE MES.MrWipOrder
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE MES.MrWipOrder SET
                                    WoErpPrefix = @WoErpPrefix,
                                    WoErpNo = @WoErpNo,
                                    MrType = @MrType,
                                    Quantity = @Quantity,
                                    InventoryId = @InventoryId,
                                    StorageLocation = @StorageLocation,
                                    RequisitionCode = @RequisitionCode,
                                    NegativeStatus = @NegativeStatus,
                                    Remark = @Remark,
                                    MaterialCategory = @MaterialCategory,
                                    SubinventoryType = @SubinventoryType,
                                    LineSeq = @LineSeq,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE MrId = @MrId
                                    AND MoId = @MoId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    WoErpPrefix,
                                    WoErpNo,
                                    MrType,
                                    Quantity,
                                    InventoryId,
                                    InventoryNo,
                                    StorageLocation,
                                    RequisitionCode,
                                    NegativeStatus,
                                    Remark,
                                    MaterialCategory,
                                    SubinventoryType,
                                    LineSeq,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    MrId,
                                    MoId
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

        #region //UpdateMrDetail -- 更新領退料單詳細資料 -- Ann 2023-12-21
        public string UpdateMrDetail(int MrDetailId, int MoId, int MtlItemId, double Quantity, int UomId, int InventoryId, string Remark, string DetailDesc, string StorageLocation)
        {
            try
            {
                if (CreateBy <= 0) throw new SystemException("【使用者ID】錯誤，請嘗試重新整理頁面!!");
                if (Quantity < 0) throw new SystemException("【領料數量】不能為空!");
                if (DetailDesc.Length > 60) throw new SystemException("【領料說明】長度錯誤!");
                if (Remark.Length > 255) throw new SystemException("【備註】長度錯誤!");

                string WoErpPrefix = "";
                string WoErpNo = "";
                string SubstitutionMtlItemNo = "";
                double? MrSubstituteRate = -1;
                decimal? MQ010 = 0;

                double epsilon = 0.0001;

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
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
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //判斷領退料單詳細資料是否有誤
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.MrId, a.MoId, a.ProcessCode, a.ConfirmStatus, a.OfficialItemStatus
                                    , a.SubstitutionId, a.SubstitutionMtlItemNo, a.SubstituteProcessCode, a.SubstituteQty, a.SubstituteRate
                                    , a.Quantity, a.ActualQuantity
                                    , b.MrType
                                    , c.WoSeq
                                    , e.BarcodeCtrl, e.DecompositionFlag
                                    , f.DocType, f.MrErpPrefix, FORMAT(f.DocDate, 'yyyy-MM-dd HH:mm:ss') DocDate
                                    , g.WoErpPrefix, g.WoErpNo
                                    , h.MtlItemNo
                                    , i.MtlItemNo BillMtlItemNo
                                    FROM MES.MrDetail a
                                    LEFT JOIN MES.MrWipOrder b ON a.MrId = b.MrId AND a.MoId = b.MoId
                                    INNER JOIN MES.ManufactureOrder c ON a.MoId = c.MoId
                                    LEFT JOIN MES.WoDetail d ON c.WoId = d.WoId AND d.MtlItemId = a.MtlItemId
                                    LEFT JOIN MES.MoMtlSetting e ON e.MoId = a.MoId AND d.WoDetailId = e.WoDetailId
                                    INNER JOIN MES.MaterialRequisition f ON a.MrId = f.MrId
                                    INNER JOIN MES.WipOrder g ON c.WoId = g.WoId
                                    INNER JOIN PDM.MtlItem h ON a.MtlItemId = h.MtlItemId
                                    INNER JOIN PDM.MtlItem i ON g.MtlItemId = i.MtlItemId
                                    WHERE MrDetailId = @MrDetailId";
                            dynamicParameters.Add("MrDetailId", MrDetailId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("領退料單詳細資料有誤!");

                            int MrId = -1;
                            string ProcessCode = "";
                            string ConfirmStatus = "";
                            string OfficialItemStatus = "";
                            string SubstituteProcessCode = "";
                            int SubstituteQty = 0;
                            string MrType = "";
                            int WoSeq = -1;
                            double ActualQuantity = -1;
                            string BarcodeCtrl = "N";
                            string DecompositionFlag = "N";
                            string DocType = "";
                            string ExcessFlag = "";
                            string CheckMoFlag = "";
                            string MrSubstitutionMtlItemNo = "";
                            string MrMtlItemNo = "";
                            double MrSubstituteQty = 0;
                            int? MrSubstitutionId = -1;
                            string BillMtlItemNo = "";
                            DateTime DocDate = new DateTime();
                            double CurrentQuantity = -1;
                            foreach (var item in result)
                            {
                                if (item.ConfirmStatus != "N")
                                {
                                    throw new SystemException("領退料單已確認，無法更改!");
                                }

                                #region //確認此領料單是否為超領單別
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT MQ010, LTRIM(RTRIM(a.MQ019)) MQ019, LTRIM(RTRIM(a.MQ054)) MQ054
                                        FROM CMSMQ a
                                        WHERE a.MQ001 = @MQ001";
                                dynamicParameters.Add("MQ001", item.MrErpPrefix);
                                var CMSMQResult = sqlConnection2.Query(sql, dynamicParameters);

                                if (CMSMQResult.Count() <= 0) throw new SystemException("ERP尚未維護此領料單單別(" + item.MrErpPrefix + ")!!");

                                foreach (var item2 in CMSMQResult)
                                {
                                    if (item.BarcodeCtrl == null && item2.MQ054 != "N")
                                    {
                                        if (item.SubstitutionMtlItemNo == item.MtlItemNo && item.SubstituteQty == 0)
                                        {
                                            throw new SystemException("此領料單非超領單別，但查無製令用料設定!!");
                                        }
                                    }

                                    CheckMoFlag = item2.MQ019;
                                    ExcessFlag = item2.MQ054;
                                    MQ010 = item2.MQ010;
                                }
                                #endregion

                                MrId = item.MrId;
                                ProcessCode = item.ProcessCode;
                                ConfirmStatus = item.ConfirmStatus;
                                OfficialItemStatus = item.OfficialItemStatus;
                                SubstituteProcessCode = item.SubstituteProcessCode;
                                SubstituteQty = item.SubstituteQty;
                                MrType = item.MrType;
                                WoSeq = item.WoSeq;
                                ActualQuantity = item.ActualQuantity;
                                BarcodeCtrl = item.BarcodeCtrl;
                                DecompositionFlag = item.DecompositionFlag;
                                DocType = item.DocType;
                                MoId = item.MoId;
                                WoErpPrefix = item.WoErpPrefix;
                                WoErpNo = item.WoErpNo;
                                MrSubstitutionId = item.SubstitutionId;
                                MrSubstitutionMtlItemNo = item.SubstitutionMtlItemNo;
                                MrMtlItemNo = item.MtlItemNo;
                                MrSubstituteQty = item.SubstituteQty;
                                MrSubstituteRate = item.SubstituteRate;
                                BillMtlItemNo = item.BillMtlItemNo;
                                DocDate = Convert.ToDateTime(item.DocDate);
                                CurrentQuantity = item.Quantity;

                                if (Quantity < ActualQuantity) throw new SystemException("數量低於已綁定條碼數量，無法更改!!");
                            }
                            #endregion

                            #region //判斷品號資料是否有誤
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT MtlItemNo
                                    FROM PDM.MtlItem
                                    WHERE MtlItemId = @MtlItemId";
                            dynamicParameters.Add("MtlItemId", MtlItemId);

                            var result3 = sqlConnection.Query(sql, dynamicParameters);
                            if (result3.Count() <= 0) throw new SystemException("品號資料有誤!");

                            string MtlItemNo = "";
                            foreach (var item in result3)
                            {
                                MtlItemNo = item.MtlItemNo;
                            }
                            #endregion

                            #region //確認更改的數量是否小於目前實際領退料數量
                            dynamicParameters = new DynamicParameters();
                            sql = @"";
                            #endregion

                            #region //若MQ019為Y，檢查用料是否為製令BOM
                            if (CheckMoFlag == "Y")
                            {
                                if (MoId <= 0) throw new SystemException("此單別需要指定製令!!");

                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM MOCTB
                                        WHERE TB001 = @WoErpPrefix
                                        AND TB002 = @WoErpNo
                                        AND TB003 = @MtlItemNo";
                                dynamicParameters.Add("WoErpPrefix", WoErpPrefix);
                                dynamicParameters.Add("WoErpNo", WoErpNo);
                                dynamicParameters.Add("MtlItemNo", MtlItemNo);

                                var MOCTBResult = sqlConnection2.Query(sql, dynamicParameters);

                                if (MOCTBResult.Count() <= 0)
                                {
                                    #region //檢查此品號是否為此製令下之替代料，是的話才允許非BOM結構下領料
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM BOMMB
                                            WHERE MB002 = @BillMtlItemNo
                                            AND MB004 = @SubMtlItemNo";
                                    dynamicParameters.Add("BillMtlItemNo", BillMtlItemNo);
                                    dynamicParameters.Add("SubMtlItemNo", MtlItemNo);

                                    var BOMMBResult = sqlConnection2.Query(sql, dynamicParameters);

                                    if (BOMMBResult.Count() <= 0) throw new SystemException("品號【" + MtlItemNo + "】並未在此製令用料中!!");
                                    #endregion
                                }
                            }
                            #endregion

                            #region //若MQ054為Y，則檢查是否超領
                            if (ExcessFlag == "Y")
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TB004 DemandRequisitionQty, TB005 RequisitionQty
                                        FROM MOCTB
                                        WHERE TB001 = @WoErpPrefix
                                        AND TB002 = @WoErpNo
                                        AND TB003 = @MtlItemNo";
                                dynamicParameters.Add("WoErpPrefix", WoErpPrefix);
                                dynamicParameters.Add("WoErpNo", WoErpNo);
                                dynamicParameters.Add("MtlItemNo", MtlItemNo);

                                var ErpWoDetailResult = sqlConnection2.Query(sql, dynamicParameters);

                                if (ErpWoDetailResult.Count() <= 0)
                                {
                                    if (SubstitutionMtlItemNo == MtlItemNo && SubstituteQty == 0)
                                    {
                                        throw new SystemException("查無ERP製令單身資料!!");
                                    }
                                }
                                else
                                {
                                    double thisDemandRequisitionQty = -1;
                                    double thisRequisitionQty = -1;
                                    foreach (var item2 in ErpWoDetailResult)
                                    {
                                        thisDemandRequisitionQty = Convert.ToDouble(item2.DemandRequisitionQty);
                                        thisRequisitionQty = Convert.ToDouble(item2.RequisitionQty);
                                    }

                                    if (MQ010 < 0)
                                    {
                                        if (Quantity > (thisDemandRequisitionQty - thisRequisitionQty))
                                        {
                                            if (Math.Abs(Quantity - (thisDemandRequisitionQty - thisRequisitionQty)) > epsilon)
                                            {
                                                throw new SystemException("此用料已領最大領料數，無法超領!!");
                                            }
                                        }
                                    }
                                }
                            }
                            #endregion

                            #region //判斷單位資料是否有誤
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 UomNo
                                    FROM PDM.UnitOfMeasure
                                    WHERE UomId = @UomId";
                            dynamicParameters.Add("UomId", UomId);

                            var result4 = sqlConnection.Query(sql, dynamicParameters);
                            if (result4.Count() <= 0) throw new SystemException("單位資料有誤!");

                            string UomNo = "";
                            foreach (var item in result4)
                            {
                                UomNo = item.UomNo;
                            }
                            #endregion

                            #region //確認ERP品號相關資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MB004)) MB004
                                    , LTRIM(RTRIM(MB030)) MB030
                                    , LTRIM(RTRIM(MB031)) MB031
                                    FROM INVMB
                                    WHERE MB001 = @MB001";
                            dynamicParameters.Add("MB001", MtlItemNo);

                            var INVMBResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (INVMBResult.Count() <= 0) throw new SystemException("ERP品號基本資料錯誤!!");

                            foreach (var item in INVMBResult)
                            {
                                #region //比對單位與ERP品號基本資料單位是否相同
                                if (item.MB004 != UomNo) throw new SystemException("物料單位(" + UomNo + ")與ERP品號庫存單位(" + item.MB004 + ")不同!!");
                                #endregion

                                #region //判斷ERP品號生效日與失效日
                                if (item.MB030 != "" && item.MB030 != null)
                                {
                                    #region //判斷日期需大於或等於生效日
                                    string EffectiveDate = item.MB030;
                                    string effYear = EffectiveDate.Substring(0, 4);
                                    string effMonth = EffectiveDate.Substring(4, 2);
                                    string effDay = EffectiveDate.Substring(6, 2);
                                    DateTime effFullDate = Convert.ToDateTime(effYear + "-" + effMonth + "-" + effDay);
                                    int effresult = DateTime.Compare(DocDate, effFullDate);
                                    if (effresult < 0) throw new SystemException("此品號尚未生效，無法使用!!");
                                    #endregion
                                }

                                if (item.MB031 != "" && item.MB031 != null)
                                {
                                    #region //判斷日期需小於或等於失效日
                                    string ExpirationDate = item.MB031;
                                    string effYear = ExpirationDate.Substring(0, 4);
                                    string effMonth = ExpirationDate.Substring(4, 2);
                                    string effDay = ExpirationDate.Substring(6, 2);
                                    DateTime effFullDate = Convert.ToDateTime(effYear + "-" + effMonth + "-" + effDay);
                                    int effresult = DateTime.Compare(DocDate, effFullDate);
                                    if (effresult > 0) throw new SystemException("此品號已失效，無法使用!! " + MtlItemNo);
                                    #endregion
                                }
                                #endregion
                            }
                            #endregion

                            #region //判斷庫別資料是否有誤
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 InventoryNo
                                    FROM SCM.Inventory
                                    WHERE InventoryId = @InventoryId";
                            dynamicParameters.Add("InventoryId", InventoryId);

                            var result5 = sqlConnection.Query(sql, dynamicParameters);
                            if (result5.Count() <= 0) throw new SystemException("庫別資料有誤!");

                            string InventoryNo = "";
                            foreach (var item in result5)
                            {
                                InventoryNo = item.InventoryNo;
                            }
                            #endregion

                            #region //檢查此料號是否已在ERP製令單身存在，並且為取替代料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.TB023
                                    FROM MOCTB a 
                                    WHERE a.TB001 = @WoErpPrefix
                                    AND a.TB002 = @WoErpNo
                                    AND a.TB003 = @MtlItemNo
                                    AND a.TB023 != a.TB003
                                    AND a.TB027 = '2'";
                            dynamicParameters.Add("WoErpPrefix", WoErpPrefix);
                            dynamicParameters.Add("WoErpNo", WoErpNo);
                            dynamicParameters.Add("MtlItemNo", MtlItemNo);

                            var CheckSubResult = sqlConnection2.Query(sql, dynamicParameters);

                            string BomMtlItemNo = "";
                            int BomMtlItemId = -1;
                            if (CheckSubResult.Count() > 0)
                            {
                                foreach (var item in CheckSubResult)
                                {
                                    BomMtlItemNo = item.TB023;

                                    #region //取得MES原元件資料
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.MtlItemId
                                            FROM PDM.MtlItem a 
                                            WHERE a.MtlItemNo = @MtlItemNo";
                                    dynamicParameters.Add("MtlItemNo", BomMtlItemNo);

                                    var BomMtlItemResult = sqlConnection.Query(sql, dynamicParameters);

                                    foreach (var item2 in BomMtlItemResult)
                                    {
                                        BomMtlItemId = item2.MtlItemId;
                                    }
                                    #endregion
                                }
                            }
                            else
                            {
                                #region //檢查此料號是否為製令中某元件替代料
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT LTRIM(RTRIM(a.TB003)) TB003
                                        , LTRIM(RTRIM(b.TA006)) TA006
                                        FROM MOCTB a
                                        INNER JOIN MOCTA b ON a.TB001 = b.TA001 AND a.TB002 = b.TA002
                                        WHERE TB001 = @WoErpPrefix
                                        AND TB002 = @WoErpNo";
                                dynamicParameters.Add("WoErpPrefix", WoErpPrefix);
                                dynamicParameters.Add("WoErpNo", WoErpNo);

                                var CheckMOCTBResult = sqlConnection2.Query(sql, dynamicParameters);

                                string TempBomMtlItemNo = "";
                                foreach (var item in CheckMOCTBResult)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM BOMMB a 
                                            INNER JOIN INVMB b ON a.MB004 = b.MB001
                                            LEFT JOIN INVMC c ON c.MC001 = a.MB004
                                            INNER JOIN INVMB d ON a.MB001 = d.MB001
                                            WHERE a.MB002 = @BillMtlItemNo
                                            AND a.MB001 = @BomMtlItemNo
                                            AND a.MB004 = @MtlItemNo";
                                    dynamicParameters.Add("BillMtlItemNo", item.TA006);
                                    dynamicParameters.Add("BomMtlItemNo", item.TB003);
                                    dynamicParameters.Add("MtlItemNo", MtlItemNo);

                                    var CheckBOMMBResult = sqlConnection2.Query(sql, dynamicParameters);

                                    if (CheckBOMMBResult.Count() > 0)
                                    {
                                        TempBomMtlItemNo = item.TB003;
                                        break;
                                    }
                                }
                                #endregion

                                if (TempBomMtlItemNo != "")
                                {
                                    BomMtlItemNo = TempBomMtlItemNo;

                                    #region //取得MES元件ID
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.MtlItemId
                                            FROM PDM.MtlItem a 
                                            WHERE a.MtlItemNo = @BomMtlItemNo";
                                    dynamicParameters.Add("BomMtlItemNo", BomMtlItemNo);

                                    var BomMtlItemResult = sqlConnection.Query(sql, dynamicParameters);

                                    if (BomMtlItemResult.Count() <= 0) throw new SystemException("查詢MES元件ID時錯誤!!");

                                    foreach (var item in BomMtlItemResult)
                                    {
                                        BomMtlItemId = item.MtlItemId;
                                    }
                                    #endregion
                                }
                                else
                                {
                                    BomMtlItemNo = MtlItemNo;
                                    BomMtlItemId = MtlItemId;
                                }
                            }
                            #endregion

                            #region //確認儲位資料是否正確
                            if (StorageLocation.Length > 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT RTRIM(LTRIM(a.NL002)) NL002
                                        FROM CMSNL a 
                                        WHERE a.NL001 = @InventoryNo
                                        AND a.NL002 = @StorageLocation";
                                dynamicParameters.Add("InventoryNo", InventoryNo);
                                dynamicParameters.Add("StorageLocation", StorageLocation);

                                var CMSNLResult = sqlConnection2.Query(sql, dynamicParameters);

                                if (CMSNLResult.Count() <= 0) throw new SystemException("儲位資料有誤!!");
                            }
                            #endregion

                            #region //UPDATE MES.MrDetail
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE MES.MrDetail SET
                                    MtlItemId = @MtlItemId,
                                    Quantity = @Quantity,
                                    UomId = @UomId,
                                    Unit = @Unit,
                                    InventoryId = @InventoryId,
                                    SubInventoryCode = @SubInventoryCode,
                                    MoId = @MoId,
                                    WoErpPrefix = @WoErpPrefix,
                                    WoErpNo = @WoErpNo,
                                    DetailDesc = @DetailDesc,
                                    Remark = @Remark,
                                    StorageLocation = @StorageLocation,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE MrDetailId = @MrDetailId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MtlItemId,
                                    Quantity,
                                    UomId,
                                    Unit = UomNo,
                                    InventoryId,
                                    SubInventoryCode = InventoryNo,
                                    MoId,
                                    WoErpPrefix,
                                    WoErpNo,
                                    DetailDesc,
                                    Remark,
                                    StorageLocation,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    MrDetailId
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
        #region //DeleteInventoryTransaction -- 刪除庫存異動單據 -- Ann 2023-12-18
        public string DeleteInventoryTransaction(int ItId)
        {
            try
            {
                string ErpNo = "";
                int rowsAffected = -1;
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
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
                            ErpNo = item.ErpNo;
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //判斷庫存異動單據資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.ItErpPrefix, a.ItErpNo, a.TransferStatus
                                    FROM SCM.InventoryTransaction a 
                                    WHERE a.ItId = @ItId";
                            dynamicParameters.Add("ItId", ItId);

                            var InventoryTransactionResult = sqlConnection.Query(sql, dynamicParameters);
                            if (InventoryTransactionResult.Count() <= 0) throw new SystemException("庫存異動單據資料錯誤!");

                            string ItErpPrefix = "";
                            string ItErpNo = "";
                            foreach (var item in InventoryTransactionResult)
                            {
                                if (item.TransferStatus != "N") throw new SystemException("此單據已拋轉ERP，無法刪除!!");
                                ItErpPrefix = item.ItErpPrefix;
                                ItErpNo = item.ItErpNo;
                            }
                            #endregion

                            #region //核對庫存異動單據是否已存在於ERP
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM INVTA
                                    WHERE TA001 = @ItErpPrefix
                                    AND TA002 = @ItErpNo";
                            dynamicParameters.Add("ItErpPrefix", ItErpPrefix);
                            dynamicParameters.Add("ItErpNo", ItErpNo);

                            var INVTAResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (INVTAResult.Count() > 0) throw new SystemException("此庫存異動單據已拋轉到ERP，無法刪除!!");
                            #endregion

                            #region //DELETE相關單據
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE a 
                                    FROM SCM.ItDetail a
                                    WHERE a.ItId = @ItId";
                            dynamicParameters.Add("ItId", ItId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE a 
                                    FROM SCM.InventoryTransaction a
                                    WHERE a.ItId = @ItId";
                            dynamicParameters.Add("ItId", ItId);

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

        #region //DeleteItDetail -- 刪除庫存異動單據詳細資料 -- Ann 2023-12-18
        public string DeleteItDetail(int ItDetailId)
        {
            try
            {
                string ErpNo = "";
                int rowsAffected = -1;
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
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
                            ErpNo = item.ErpNo;
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //判斷庫存異動單據詳細資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.ItId, b.ItErpPrefix, b.ItErpNo, b.ConfirmStatus
                                    FROM SCM.ItDetail a 
                                    INNER JOIN SCM.InventoryTransaction b ON a.ItId = b.ItId
                                    WHERE a.ItDetailId = @ItDetailId";
                            dynamicParameters.Add("ItDetailId", ItDetailId);

                            var ItDetailResult = sqlConnection.Query(sql, dynamicParameters);
                            if (ItDetailResult.Count() <= 0) throw new SystemException("庫存異動單據詳細資料錯誤!");

                            int ItId = -1;
                            string ItErpPrefix = "";
                            string ItErpNo = "";
                            foreach (var item in ItDetailResult)
                            {
                                if (item.ConfirmStatus != "N") throw new SystemException("此庫存異動單據狀態無法刪除!!");
                                ItId = item.ItId;
                                ItErpPrefix = item.ItErpPrefix;
                                ItErpNo = item.ItErpNo;
                            }
                            #endregion

                            #region //DELETE相關單據
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE a 
                                    FROM SCM.ItDetail a
                                    WHERE a.ItDetailId = @ItDetailId";
                            dynamicParameters.Add("ItDetailId", ItDetailId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //取得本國幣別
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 LTRIM(RTRIM(MA003)) MA003 FROM CMSMA";

                            var CMSMAResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (CMSMAResult.Count() <= 0) throw new SystemException("共用參數檔案資料錯誤!!");

                            string Currency = "";
                            foreach (var item in CMSMAResult)
                            {
                                Currency = item.MA003;
                            }
                            #endregion

                            #region //小數點後取位
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.MF005, a.MF006
                                    FROM CMSMF a 
                                    WHERE a.MF001 = @Currency";
                            dynamicParameters.Add("Currency", Currency);

                            var CMSMFResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (CMSMFResult.Count() <= 0) throw new SystemException("金額小數點取位資料錯誤!!");

                            int unitDecimal = 0;
                            int amountDecimal = 0;
                            foreach (var item in CMSMFResult)
                            {
                                unitDecimal = Convert.ToInt32(item.MF005);
                                amountDecimal = Convert.ToInt32(item.MF006);
                            }
                            #endregion

                            #region //重新計算目前總數量、總金額
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT SUM(a.ItQty) TotalQty, SUM(a.Amount) TotalAmount
                                    FROM SCM.ItDetail a 
                                    WHERE ItId = @ItId";
                            dynamicParameters.Add("ItId", ItId);

                            var ItDetailResult2 = sqlConnection.Query(sql, dynamicParameters);

                            if (ItDetailResult2.Count() <= 0) throw new SystemException("庫存異動單據詳細資料錯誤!!");

                            double TotalQty = 0;
                            double TotalAmount = 0;
                            foreach (var item in ItDetailResult2)
                            {
                                TotalQty = Convert.ToDouble(item.TotalQty);
                                TotalAmount = Math.Round(Convert.ToDouble(item.TotalAmount), amountDecimal);
                            }
                            #endregion

                            #region //Update SCM.InventoryTransaction 總數量、總金額
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.InventoryTransaction SET
                                    TotalQty = @TotalQty,
                                    Amount = @TotalAmount,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE ItId = @ItId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    TotalQty,
                                    TotalAmount,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    ItId
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

        #region //DeleteMaterialRequisition -- 刪除領退料單資料 -- Ann 2023-12-21
        public string DeleteMaterialRequisition(int MrId)
        {
            try
            {
                string ErpNo = "";
                int rowsAffected = -1;
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
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
                            ErpNo = item.ErpNo;
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //判斷領退料單資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.MrErpPrefix, a.MrErpNo, a.ConfirmStatus, a.ConfirmUserId, FORMAT(a.DocDate, 'yyyyMMdd') DocDate
                                    , b.UserNo
                                    FROM MES.MaterialRequisition a
                                    LEFT JOIN BAS.[User] b ON a.ConfirmUserId = b.UserId
                                    WHERE a.MrId = @MrId";
                            dynamicParameters.Add("MrId", MrId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("領退料單資料錯誤!");

                            string MrErpPrefix = "";
                            string MrErpNo = "";
                            string DocDate = "";
                            string UserNo = "";
                            foreach (var item in result)
                            {
                                if (item.ConfirmStatus != "N" || item.UserNo != null) throw new SystemException("領退料單已有核單紀錄，無法刪除!");
                                MrErpPrefix = item.MrErpPrefix;
                                MrErpNo = item.MrErpNo;
                                DocDate = item.DocDate;
                                UserNo = item.UserNo;
                            }
                            #endregion

                            string ErpFullNo = MrErpPrefix + "-" + MrErpNo;

                            #region //核對領料單目前狀態是否可以刪除
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MOCTC
                                    WHERE TC001 = @TC001
                                    AND TC002 = @TC002";
                            dynamicParameters.Add("TC001", MrErpPrefix);
                            dynamicParameters.Add("TC002", MrErpNo);

                            var MOCTCResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (MOCTCResult.Count() > 0) throw new SystemException("此領料單已拋轉到ERP，無法刪除!!");
                            #endregion

                            #region //檢查領料單BARCODE是否已經過站
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.MrDetail a
                                    INNER JOIN MES.MrBarcodeRegister b ON a.MrDetailId = b.MrDetailId
                                    INNER JOIN MES.Barcode c ON b.BarcodeNo = c.BarcodeNo
                                    INNER JOIN MES.BarcodeProcess d ON c.BarcodeId = d.BarcodeId AND d.MoId = a.MoId
                                    WHERE a.MrId = @MrId";
                            dynamicParameters.Add("MrId", MrId);

                            var CheckBarcodeProcessResult = sqlConnection.Query(sql, dynamicParameters);

                            if (CheckBarcodeProcessResult.Count() > 0) throw new SystemException("此領料單已有條碼已過站，無法刪除!");
                            #endregion

                            #region //檢查此發料單單身是否有條碼轉移紀錄
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.LogId, a.MoId, a.BarcodeNo, a.CurrentMoProcessId, a.NextMoProcessId, a.CurrentProdStatus, a.BarcodeStatus
                                    , c.BarcodeStatus CurrentBarcodeStatus
                                    , d.StatusName
                                    FROM MES.MrBarcodeTransferLog a
                                    INNER JOIN MES.MrDetail b ON a.MrDetailId = b.MrDetailId
                                    INNER JOIN MES.Barcode c ON a.BarcodeNo = c.BarcodeNo
                                    INNER JOIN BAS.[Status] d ON c.BarcodeStatus = d.StatusNo AND d.StatusSchema = 'Barcode.BarcodeStatus'
                                    WHERE b.MrId = @MrId";
                            dynamicParameters.Add("MrId", MrId);

                            var TransferLogResult = sqlConnection.Query(sql, dynamicParameters);

                            if (TransferLogResult.Count() > 0)
                            {
                                foreach (var item2 in TransferLogResult)
                                {
                                    if (item2.CurrentBarcodeStatus != "1" && item2.CurrentBarcodeStatus != "3")
                                    {
                                        throw new SystemException("此條碼【" + item2.BarcodeNo + "】狀態【" + item2.StatusName + "】無法刪除，請先進行退料或解除綁定!");
                                    }

                                    #region //UPDATE MES.Barcode回原先資料
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE MES.Barcode SET
                                            MoId = @MoId,
                                            CurrentMoProcessId = @CurrentMoProcessId,
                                            NextMoProcessId = @NextMoProcessId,
                                            CurrentProdStatus = @CurrentProdStatus,
                                            BarcodeStatus = @BarcodeStatus,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE BarcodeNo = @BarcodeNo";
                                    dynamicParameters.AddDynamicParams(
                                      new
                                      {
                                          item2.MoId,
                                          item2.CurrentMoProcessId,
                                          item2.NextMoProcessId,
                                          item2.CurrentProdStatus,
                                          item2.BarcodeStatus,
                                          LastModifiedDate,
                                          LastModifiedBy,
                                          item2.BarcodeNo
                                      });

                                    rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //將轉移LOG移除
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE FROM MES.MrBarcodeTransferLog
                                            WHERE LogId = @LogId";
                                    dynamicParameters.Add("LogId", item2.LogId);

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                            }
                            else
                            {
                                #region //刪除MES.Barcode資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE a
                                        FROM MES.Barcode a
                                        INNER JOIN MES.MrBarcodeRegister b ON a.BarcodeNo = b.BarcodeNo
                                        INNER JOIN MES.MrDetail c ON b.MrDetailId = c.MrDetailId
                                        WHERE c.MrId = @MrId";
                                dynamicParameters.Add("MrId", MrId);

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }
                            #endregion

                            #region //刪除MES單據
                            #region //先刪除關聯Table
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE a 
                                    FROM MES.MrBarcodeRegister a
                                    INNER JOIN MES.MrDetail b ON a.MrDetailId = b.MrDetailId
                                    WHERE b.MrId = @MrId";
                            dynamicParameters.Add("MrId", MrId);

                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE a 
                                    FROM MES.MrBarcodeReRegister a
                                    INNER JOIN MES.MrDetail b ON a.MrDetailId = b.MrDetailId
                                    WHERE b.MrId = @MrId";
                            dynamicParameters.Add("MrId", MrId);

                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE FROM MES.MrWipOrder
                                    WHERE MrId = @MrId";
                            dynamicParameters.Add("MrId", MrId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE FROM MES.MrDetail
                                    WHERE MrId = @MrId";
                            dynamicParameters.Add("MrId", MrId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //刪除主要table
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE MES.MaterialRequisition
                                    WHERE MrId = @MrId";
                            dynamicParameters.Add("MrId", MrId);

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

        #region //DeleteMrWipOrder -- 刪除領退料單設定資料 -- Ann 2023-12-21
        public string DeleteMrWipOrder(int MrId, int MoId)
        {
            int rowsAffected = 0;
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
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
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //判斷領退料單資料是否正確
                            sql = @"SELECT a.MrErpPrefix, a.MrErpNo, a.ConfirmStatus, a.ConfirmUserId   
                                    , b.UserNo
                                    FROM MES.MaterialRequisition a
                                    LEFT JOIN BAS.[User] b ON a.ConfirmUserId = b.UserId
                                    WHERE a.MrId = @MrId";
                            dynamicParameters.Add("MrId", MrId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("領退料單資料錯誤!");

                            string MrErpPrefix = "";
                            string MrErpNo = "";
                            foreach (var item in result)
                            {
                                if (item.ConfirmStatus != "N") throw new SystemException("領退料單已確認，無法刪除!");
                                MrErpPrefix = item.MrErpPrefix;
                                MrErpNo = item.MrErpNo;
                            }
                            #endregion

                            #region //核對領料單目前狀態是否可以刪除
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TC009
                                    FROM MOCTC
                                    WHERE TC001 = @TC001
                                    AND TC002 = @TC002";
                            dynamicParameters.Add("TC001", MrErpPrefix);
                            dynamicParameters.Add("TC002", MrErpNo);

                            var MOCTCResult = sqlConnection2.Query(sql, dynamicParameters);

                            foreach (var item in MOCTCResult)
                            {
                                if (item.TC009 != "N") throw new SystemException("此領料ERP狀態無法刪除!!");
                            }
                            #endregion

                            #region //判斷領退料單設定資料是否正確
                            sql = @"SELECT TOP 1 1
                                    FROM MES.MrWipOrder
                                    WHERE MrId = @MrId
                                    AND MoId = @MoId";
                            dynamicParameters.Add("MrId", MrId);
                            dynamicParameters.Add("MoId", MoId);

                            var result2 = sqlConnection.Query(sql, dynamicParameters);
                            if (result2.Count() <= 0) throw new SystemException("領退料單設定資料錯誤!");
                            #endregion

                            #region //檢查領料單BARCODE是否已經過站
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.MrBarcodeRegister a 
                                    INNER JOIN MES.MrDetail b ON a.MrDetailId = b.MrDetailId
                                    INNER JOIN MES.MrWipOrder c ON b.MrId = c.MrId AND b.MoId = c.MoId
                                    INNER JOIN MES.Barcode d ON a.BarcodeNo = d.BarcodeNo
                                    INNER JOIN MES.BarcodeProcess e ON d.BarcodeId = e.BarcodeId
                                    WHERE b.MrId = @MrId
                                    AND b.MoId = @MoId";
                            dynamicParameters.Add("MrId", MrId);
                            dynamicParameters.Add("MoId", MoId);

                            var CheckBarcodeProcessResult = sqlConnection.Query(sql, dynamicParameters);

                            if (CheckBarcodeProcessResult.Count() > 0) throw new SystemException("此領料單已有條碼已過站，無法刪除!");
                            #endregion

                            #region //檢查此發料單單身是否有條碼轉移紀錄
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.LogId, a.MoId, a.BarcodeNo, a.CurrentMoProcessId, a.NextMoProcessId, a.CurrentProdStatus, a.BarcodeStatus
                                    , c.BarcodeStatus CurrentBarcodeStatus
                                    , d.StatusName
                                    FROM MES.MrBarcodeTransferLog a
                                    INNER JOIN MES.MrDetail b ON a.MrDetailId = b.MrDetailId
                                    INNER JOIN MES.Barcode c ON a.BarcodeNo = c.BarcodeNo
                                    INNER JOIN BAS.[Status] d ON c.BarcodeStatus = d.StatusNo AND d.StatusSchema = 'Barcode.BarcodeStatus'
                                    WHERE b.MrId = @MrId
                                    AND b.MoId = @MoId";
                            dynamicParameters.Add("MrId", MrId);
                            dynamicParameters.Add("MoId", MoId);

                            var TransferLogResult = sqlConnection.Query(sql, dynamicParameters);

                            if (TransferLogResult.Count() > 0)
                            {
                                foreach (var item2 in TransferLogResult)
                                {
                                    if (item2.CurrentBarcodeStatus != "1" && item2.CurrentBarcodeStatus != "3" && item2.CurrentBarcodeStatus != "10")
                                    {
                                        throw new SystemException("此條碼【" + item2.BarcodeNo + "】狀態【" + item2.StatusName + "】無法刪除，請先進行退料或解除綁定!");
                                    }

                                    #region //UPDATE MES.Barcode回原先資料
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE MES.Barcode SET
                                            MoId = @MoId,
                                            CurrentMoProcessId = @CurrentMoProcessId,
                                            NextMoProcessId = @NextMoProcessId,
                                            CurrentProdStatus = @CurrentProdStatus,
                                            BarcodeStatus = @BarcodeStatus,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE BarcodeNo = @BarcodeNo";
                                    dynamicParameters.AddDynamicParams(
                                      new
                                      {
                                          item2.MoId,
                                          item2.CurrentMoProcessId,
                                          item2.NextMoProcessId,
                                          item2.CurrentProdStatus,
                                          item2.BarcodeStatus,
                                          LastModifiedDate,
                                          LastModifiedBy,
                                          item2.BarcodeNo
                                      });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //將轉移LOG移除
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE FROM MES.MrBarcodeTransferLog
                                            WHERE LogId = @LogId";
                                    dynamicParameters.Add("LogId", item2.LogId);

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                            }
                            else
                            {
                                #region //刪除MES.Barcode資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE a
                                        FROM MES.Barcode a
                                        INNER JOIN MES.MrBarcodeRegister b ON a.BarcodeNo = b.BarcodeNo
                                        INNER JOIN MES.MrDetail c ON b.MrDetailId = c.MrDetailId
                                        WHERE c.MrId = @MrId AND c.MoId = @MoId";
                                dynamicParameters.Add("MrId", MrId);
                                dynamicParameters.Add("MoId", MoId);

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }
                            #endregion

                            #region //先刪除關聯Table
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE a 
                                    FROM MES.MrBarcodeRegister a
                                    INNER JOIN MES.MrDetail b ON a.MrDetailId = b.MrDetailId
                                    WHERE b.MrId = @MrId AND b.MoId = @MoId";
                            dynamicParameters.Add("MrId", MrId);
                            dynamicParameters.Add("MoId", MoId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE a 
                                    FROM MES.MrBarcodeReRegister a
                                    INNER JOIN MES.MrDetail b ON a.MrDetailId = b.MrDetailId
                                    WHERE b.MrId = @MrId AND b.MoId = @MoId";
                            dynamicParameters.Add("MrId", MrId);
                            dynamicParameters.Add("MoId", MoId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE FROM MES.MrDetail
                                    WHERE MrId = @MrId
                                    AND MoId = @MoId";
                            dynamicParameters.Add("MrId", MrId);
                            dynamicParameters.Add("MoId", MoId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //刪除主要table
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE MES.MrWipOrder
                                    WHERE MrId = @MrId
                                    AND MoId = @MoId";
                            dynamicParameters.Add("MrId", MrId);
                            dynamicParameters.Add("MoId", MoId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //重新排序MrWipOrder
                            dynamicParameters = new DynamicParameters();
                            sql = @"With UpdateSort As
                                    (
                                      SELECT LineSeq,
                                      ROW_NUMBER() OVER(ORDER BY LineSeq) NewMrWipOrderSort
                                      FROM MES.MrWipOrder
                                      WHERE MrId = @MrId
                                    )
                                    UPDATE MES.MrWipOrder 
                                    SET LineSeq = RIGHT(REPLICATE('0', 4) + CAST(NewMrWipOrderSort as NVARCHAR), 4)
                                    FROM MES.MrWipOrder
                                    INNER JOIN UpdateSort ON MES.MrWipOrder.LineSeq = UpdateSort.LineSeq
                                    WHERE MrId = @MrId";
                            dynamicParameters.Add("MrId", MrId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //重新排序MrDetail
                            dynamicParameters = new DynamicParameters();
                            sql = @"With UpdateSort As
                                    (
                                      SELECT MrSequence,
                                      ROW_NUMBER() OVER(ORDER BY MrSequence) NewMrDetailSort
                                      FROM MES.MrDetail
                                      WHERE MrId = @MrId
                                    )
                                    UPDATE MES.MrDetail 
                                    SET MrSequence = RIGHT(REPLICATE('0', 4) + CAST(NewMrDetailSort as NVARCHAR), 4)
                                    FROM MES.MrDetail
                                    INNER JOIN UpdateSort ON MES.MrDetail.MrSequence = UpdateSort.MrSequence
                                    WHERE MrId = @MrId";
                            dynamicParameters.Add("MrId", MrId);

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

        #region //DeleteMrDetail -- 刪除領退料單詳細資料 -- Ann 2023-12-21
        public string DeleteMrDetail(int MrDetailId)
        {
            int rowsAffected = 0;
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
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
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //判斷領退料單詳細資料是否正確
                            sql = @"SELECT a.MrId, a.ConfirmStatus, a.MoId
                                    , b.MrErpPrefix, b.MrErpNo
                                    , c.UserNo
                                    FROM MES.MrDetail a
                                    INNER JOIN MES.MaterialRequisition b ON a.MrId = b.MrId
                                    LEFT JOIN BAS.[User] c ON b.ConfirmUserId = c.UserId
                                    WHERE a.MrDetailId = @MrDetailId";
                            dynamicParameters.Add("MrDetailId", MrDetailId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("領退料單詳細資料錯誤!");

                            int MrId = -1;
                            string MrErpPrefix = "";
                            string MrErpNo = "";
                            int MoId = -1;
                            foreach (var item in result)
                            {
                                if (item.ConfirmStatus != "N") throw new SystemException("狀態已確認，無法刪除!");
                                MrId = item.MrId;
                                MrErpPrefix = item.MrErpPrefix;
                                MrErpNo = item.MrErpNo;
                                MoId = item.MoId;
                            }
                            #endregion

                            #region //製令自動帶出狀況下，刪除領料單單身相關卡控
                            #region //先計算此製令在領料單身下剩餘筆數
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECt COUNT(1) Count
                                    FROM MES.MrDetail a 
                                    WHERE a.MrId = @MrId
                                    AND a.MoId = @MoId";
                            dynamicParameters.Add("MrId", MrId);
                            dynamicParameters.Add("MoId", MoId);

                            var MrDetailCountResult = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item in MrDetailCountResult)
                            {
                                if (item.Count <= 1)
                                {
                                    #region //確認是否有製令自動帶出相關資料
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM MES.MrWipOrder a
                                            WHERE a.MrId = @MrId
                                            AND a.MoId = @MoId ";
                                    dynamicParameters.Add("MrId", MrId);
                                    dynamicParameters.Add("MoId", MoId);

                                    var MrWipOrderResult = sqlConnection.Query(sql, dynamicParameters);

                                    if (MrWipOrderResult.Count() > 0) throw new SystemException("若需完整刪除單身，請從製令自動帶出來進行刪除!!");
                                    #endregion
                                }
                            }
                            #endregion
                            #endregion

                            #region //核對領料單目前狀態是否可以刪除
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TC009
                                    FROM MOCTC
                                    WHERE TC001 = @TC001
                                    AND TC002 = @TC002";
                            dynamicParameters.Add("TC001", MrErpPrefix);
                            dynamicParameters.Add("TC002", MrErpNo);

                            var MOCTCResult = sqlConnection2.Query(sql, dynamicParameters);

                            foreach (var item in MOCTCResult)
                            {
                                if (item.TC009 != "N") throw new SystemException("此領料ERP狀態無法刪除!!");
                            }
                            #endregion

                            #region //檢查領料單BARCODE是否已經過站
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM MES.MrDetail a
                                    INNER JOIN MES.MrBarcodeRegister b ON a.MrDetailId = b.MrDetailId
                                    INNER JOIN MES.Barcode c ON b.BarcodeNo = c.BarcodeNo
                                    INNER JOIN MES.BarcodeProcess d ON c.BarcodeId = d.BarcodeId AND d.MoId = a.MoId
                                    WHERE a.MrDetailId = @MrDetailId";
                            dynamicParameters.Add("MrDetailId", MrDetailId);

                            var CheckBarcodeProcessResult = sqlConnection.Query(sql, dynamicParameters);

                            if (CheckBarcodeProcessResult.Count() > 0) throw new SystemException("此領料單已有條碼已過站，無法刪除!");
                            #endregion

                            #region //檢查此發料單單身是否有條碼轉移紀錄
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.LogId, a.MoId, a.BarcodeNo, a.CurrentMoProcessId, a.NextMoProcessId, a.CurrentProdStatus, a.BarcodeStatus
                                    , b.BarcodeStatus CurrentBarcodeStatus
                                    , c.StatusName
                                    FROM MES.MrBarcodeTransferLog a
                                    INNER JOIN MES.Barcode b ON a.BarcodeNo = b.BarcodeNo
                                    INNER JOIN BAS.[Status] c ON b.BarcodeStatus = c.StatusNo AND c.StatusSchema = 'Barcode.BarcodeStatus'
                                    WHERE a.MrDetailId = @MrDetailId";
                            dynamicParameters.Add("MrDetailId", MrDetailId);

                            var TransferLogResult = sqlConnection.Query(sql, dynamicParameters);

                            if (TransferLogResult.Count() > 0)
                            {
                                foreach (var item2 in TransferLogResult)
                                {
                                    if (item2.CurrentBarcodeStatus != "1" && item2.CurrentBarcodeStatus != "3")
                                    {
                                        throw new SystemException("此條碼【" + item2.BarcodeNo + "】狀態【" + item2.StatusName + "】無法刪除，請先進行退料或解除綁定!");
                                    }

                                    #region //UPDATE MES.Barcode回原先資料
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE MES.Barcode SET
                                            MoId = @MoId,
                                            CurrentMoProcessId = @CurrentMoProcessId,
                                            NextMoProcessId = @NextMoProcessId,
                                            CurrentProdStatus = @CurrentProdStatus,
                                            BarcodeStatus = @BarcodeStatus,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE BarcodeNo = @BarcodeNo";
                                    dynamicParameters.AddDynamicParams(
                                      new
                                      {
                                          item2.MoId,
                                          item2.CurrentMoProcessId,
                                          item2.NextMoProcessId,
                                          item2.CurrentProdStatus,
                                          item2.BarcodeStatus,
                                          LastModifiedDate,
                                          LastModifiedBy,
                                          item2.BarcodeNo
                                      });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //將轉移LOG移除
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"DELETE FROM MES.MrBarcodeTransferLog
                                            WHERE LogId = @LogId";
                                    dynamicParameters.Add("LogId", item2.LogId);

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                            }
                            else
                            {
                                #region //刪除MES.Barcode資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE a
                                        FROM MES.Barcode a
                                        INNER JOIN MES.MrBarcodeRegister b ON a.BarcodeNo = b.BarcodeNo
                                        WHERE b.MrDetailId = @MrDetailId";
                                dynamicParameters.Add("MrDetailId", MrDetailId);

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }
                            #endregion

                            #region //先刪除關聯Table
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE a 
                                    FROM MES.MrBarcodeRegister a
                                    INNER JOIN MES.MrDetail b ON a.MrDetailId = b.MrDetailId
                                    WHERE b.MrDetailId = @MrDetailId";
                            dynamicParameters.Add("MrDetailId", MrDetailId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE a 
                                    FROM MES.MrBarcodeReRegister a
                                    INNER JOIN MES.MrDetail b ON a.MrDetailId = b.MrDetailId
                                    WHERE b.MrDetailId = @MrDetailId";
                            dynamicParameters.Add("MrDetailId", MrDetailId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //刪除主要table
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE MES.MrDetail
                                    WHERE MrDetailId = @MrDetailId";
                            dynamicParameters.Add("MrDetailId", MrDetailId);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //重新排序MrDetail
                            dynamicParameters = new DynamicParameters();
                            sql = @"With UpdateSort As
                                    (
                                      SELECT MrSequence,
                                      ROW_NUMBER() OVER(ORDER BY MrSequence) NewMrDetailSort
                                      FROM MES.MrDetail
                                      WHERE MrId = @MrId
                                    )
                                    UPDATE MES.MrDetail 
                                    SET MrSequence = RIGHT(REPLICATE('0', 4) + CAST(NewMrDetailSort as NVARCHAR), 4)
                                    FROM MES.MrDetail
                                    INNER JOIN UpdateSort ON MES.MrDetail.MrSequence = UpdateSort.MrSequence
                                    WHERE MrId = @MrId";
                            dynamicParameters.Add("MrId", MrId);

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

        #region //DeleteMrBarcodeRegister -- 刪除領料單條碼註冊資料 -- Ann 2023-12-21
        public string DeleteMrBarcodeRegister(int BarcodeRegisterId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();

                        #region //判斷物料條碼控管資料是否正確
                        sql = @"SELECT a.BarcodeNo, a.CreateDate
                                , b.MrDetailId, b.ActualQuantity, b.Unit, b.MoId
                                , c.ConfirmStatus, c.MrId
                                , d.BarcodeId, d.BarcodeQty
                                , e.BarcodeProcessId
                                FROM MES.MrBarcodeRegister a
                                INNER JOIN MES.MrDetail b ON a.MrDetailId = b.MrDetailId
                                INNER JOIN MES.MaterialRequisition c ON b.MrId = c.MrId
                                LEFT JOIN MES.Barcode d ON a.BarcodeNo = d.BarcodeNo
                                LEFT JOIN MES.BarcodeProcess e ON d.BarcodeId = e.BarcodeId
                                WHERE BarcodeRegisterId = @BarcodeRegisterId";
                        dynamicParameters.Add("BarcodeRegisterId", BarcodeRegisterId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("物料條碼控管資料錯誤!");

                        string BarcodeNo = "";
                        int MrDetailId = -1;
                        double? ActualQuantity = -1;
                        int? BarcodeId = -1;
                        double? BarcodeQty = -1;
                        int rowsAffected = 0;
                        string Unit = "";
                        int MoId = -1;
                        int MrId = -1;
                        DateTime BarcodeCreateDate = new DateTime();
                        foreach (var item in result)
                        {
                            if (item.ConfirmStatus != "N") throw new SystemException("領退料單已確認，無法更改!");
                            if (item.ActualQuantity < 0) throw new SystemException("領取套數錯誤!");
                            BarcodeNo = item.BarcodeNo;
                            MrDetailId = item.MrDetailId;
                            ActualQuantity = item.ActualQuantity;
                            BarcodeId = item.BarcodeId != null ? item.BarcodeId : (int?)null;
                            BarcodeQty = item.BarcodeQty != null ? item.BarcodeQty : (double?)null;
                            Unit = item.Unit;
                            MoId = item.MoId;
                            MrId = item.MrId;
                            BarcodeCreateDate = item.CreateDate;
                        }
                        #endregion

                        if (BarcodeId != null) ActualQuantity -= BarcodeQty;
                        else ActualQuantity -= 1;

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.MrBarcodeRegister
                                WHERE BarcodeRegisterId = @BarcodeRegisterId";
                        dynamicParameters.Add("BarcodeRegisterId", BarcodeRegisterId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //檢查此BARCODE是否已經過站
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 b.FinishDate
                                FROM MES.Barcode a
                                INNER JOIN MES.BarcodeProcess b ON a.BarcodeId = b.BarcodeId
                                WHERE a.BarcodeNo = @BarcodeNo
                                AND b.MoId = @MoId
                                ORDER BY FinishDate DESC";
                        dynamicParameters.Add("BarcodeNo", BarcodeNo);
                        dynamicParameters.Add("MoId", MoId);

                        var CheckBarcodeProcessResult = sqlConnection.Query(sql, dynamicParameters);

                        if (CheckBarcodeProcessResult.Count() > 0)
                        {
                            foreach (var item in CheckBarcodeProcessResult)
                            {
                                DateTime BarcodeProcessDate = item.FinishDate;
                                int dateResult = DateTime.Compare(BarcodeProcessDate, BarcodeCreateDate);
                                if (dateResult > 0) throw new SystemException("條碼【" + BarcodeNo + "】已有過站紀錄，無法刪除!!");
                            }
                        }
                        #endregion

                        #region //檢查此發料單單身是否有條碼轉移紀錄
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.LogId, a.MoId, a.BarcodeNo, a.CurrentMoProcessId, a.NextMoProcessId, a.CurrentProdStatus, a.BarcodeStatus
                                , b.BarcodeStatus CurrentBarcodeStatus
                                , c.StatusName
                                FROM MES.MrBarcodeTransferLog a
                                INNER JOIN MES.Barcode b ON a.BarcodeNo = b.BarcodeNo
                                INNER JOIN BAS.[Status] c ON b.BarcodeStatus = c.StatusNo AND c.StatusSchema = 'Barcode.BarcodeStatus'
                                WHERE a.MrDetailId = @MrDetailId
                                AND a.BarcodeNo = @BarcodeNo
                                ORDER BY a.CreateDate DESC";
                        dynamicParameters.Add("MrDetailId", MrDetailId);
                        dynamicParameters.Add("BarcodeNo", BarcodeNo);

                        var TransferLogResult = sqlConnection.Query(sql, dynamicParameters);

                        if (TransferLogResult.Count() > 0)
                        {
                            foreach (var item2 in TransferLogResult)
                            {
                                if (item2.CurrentBarcodeStatus != "1" && item2.CurrentBarcodeStatus != "3")
                                {
                                    throw new SystemException("此條碼【" + item2.BarcodeNo + "】狀態【" + item2.StatusName + "】無法刪除，請先進行退料或解除綁定!");
                                }

                                #region //UPDATE MES.Barcode回原先資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE MES.Barcode SET
                                        MoId = @MoId,
                                        CurrentMoProcessId = @CurrentMoProcessId,
                                        NextMoProcessId = @NextMoProcessId,
                                        CurrentProdStatus = @CurrentProdStatus,
                                        BarcodeStatus = @BarcodeStatus,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE BarcodeNo = @BarcodeNo";
                                dynamicParameters.AddDynamicParams(
                                  new
                                  {
                                      item2.MoId,
                                      item2.CurrentMoProcessId,
                                      item2.NextMoProcessId,
                                      item2.CurrentProdStatus,
                                      item2.BarcodeStatus,
                                      LastModifiedDate,
                                      LastModifiedBy,
                                      BarcodeNo
                                  });

                                rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region //將轉移LOG移除
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE FROM MES.MrBarcodeTransferLog
                                        WHERE LogId = @LogId";
                                dynamicParameters.Add("LogId", item2.LogId);

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }
                        }
                        else
                        {
                            #region //刪除MES.Barcode資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"DELETE a
                                    FROM MES.Barcode a
                                    WHERE a.MoId = @MoId
                                    AND a.BarcodeNo = @BarcodeNo";
                            dynamicParameters.Add("MoId", MoId);
                            dynamicParameters.Add("BarcodeNo", BarcodeNo);

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }
                        #endregion

                        #region //判斷此料號是否為可切割料號
                        string DecompositionFlag = "";
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT d.DecompositionFlag
                                FROM MES.MrDetail a
                                INNER JOIN MES.ManufactureOrder b ON a.MoId = b.MoId
                                INNER JOIN MES.WoDetail c ON a.MtlItemId = c.MtlItemId AND c.WoId = b.WoId
                                INNER JOIN MES.MoMtlSetting d ON c.WoDetailId = d.WoDetailId
                                WHERE a.MrDetailId = @MrDetailId";
                        dynamicParameters.Add("MrDetailId", MrDetailId);

                        var MainBarcodeResult = sqlConnection.Query(sql, dynamicParameters);
                        foreach (var item3 in MainBarcodeResult)
                        {
                            DecompositionFlag = item3.DecompositionFlag;
                        }
                        #endregion

                        if (Unit == "PCS" && (DecompositionFlag == "N" || DecompositionFlag == ""))
                        {
                            #region //更新MrDetail領取套數
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE MES.MrDetail SET
                                    ActualQuantity = @ActualQuantity,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE MrDetailId = @MrDetailId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    ActualQuantity,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    MrDetailId
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

        #region //DeleteMrBarcodeReRegister -- 刪除領料單條碼退料資料 -- Ann 2023-12-21
        public string DeleteMrBarcodeReRegister(int BarcodeReRegisterId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        dynamicParameters = new DynamicParameters();

                        #region //判斷物料條碼控管資料是否正確
                        sql = @"SELECT a.BarcodeNo
                                , b.MrDetailId, b.Quantity, b.ActualQuantity
                                , c.ConfirmStatus, c.TransferStatus
                                , d.BarcodeId, d.BarcodeQty, d.BarcodeStatus
                                FROM MES.MrBarcodeReRegister a
                                INNER JOIN MES.MrDetail b ON a.MrDetailId = b.MrDetailId
                                INNER JOIN MES.MaterialRequisition c ON b.MrId = c.MrId
                                LEFT JOIN MES.Barcode d ON a.BarcodeNo = d.BarcodeNo
                                WHERE BarcodeReRegisterId = @BarcodeReRegisterId";
                        dynamicParameters.Add("BarcodeReRegisterId", BarcodeReRegisterId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("物料條碼控管資料錯誤!");

                        string BarcodeNo = "";
                        int MrDetailId = -1;
                        double? ActualQuantity = -1;
                        int? BarcodeId = -1;
                        double? BarcodeQty = -1;
                        int rowsAffected = 0;
                        string TransferStatus = "";
                        foreach (var item in result)
                        {
                            BarcodeQty = item.BarcodeQty != null ? item.BarcodeQty : 1;
                            if (item.ConfirmStatus != "N") throw new SystemException("領退料單已確認，無法更改!");
                            if (item.ActualQuantity <= 0) throw new SystemException("領取套數錯誤!");
                            BarcodeNo = item.BarcodeNo;
                            MrDetailId = item.MrDetailId;
                            ActualQuantity = item.ActualQuantity;
                            BarcodeId = item.BarcodeId != null ? item.BarcodeId : (int?)null;
                            TransferStatus = item.TransferStatus;
                        }
                        #endregion

                        if (BarcodeId != null) ActualQuantity -= BarcodeQty;
                        else throw new SystemException("查無此條碼資訊!");

                        #region //檢查此發料單單身是否有條碼轉移紀錄
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.LogId, a.MoId, a.BarcodeNo, a.CurrentMoProcessId, a.NextMoProcessId, a.CurrentProdStatus, a.BarcodeStatus
                                , b.BarcodeStatus CurrentBarcodeStatus
                                , c.StatusName
                                FROM MES.MrBarcodeTransferLog a
                                INNER JOIN MES.Barcode b ON a.BarcodeNo = b.BarcodeNo
                                INNER JOIN BAS.[Status] c ON b.BarcodeStatus = c.StatusNo AND c.StatusSchema = 'Barcode.BarcodeStatus'
                                WHERE a.MrDetailId = @MrDetailId
                                AND a.BarcodeNo = @BarcodeNo
                                ORDER BY a.CreateDate DESC";
                        dynamicParameters.Add("MrDetailId", MrDetailId);
                        dynamicParameters.Add("BarcodeNo", BarcodeNo);

                        var TransferLogResult = sqlConnection.Query(sql, dynamicParameters);

                        if (TransferLogResult.Count() > 0)
                        {
                            foreach (var item2 in TransferLogResult)
                            {
                                if (item2.CurrentBarcodeStatus != "8")
                                {
                                    throw new SystemException("此條碼【" + item2.BarcodeNo + "】狀態【" + item2.StatusName + "】無法退料!!");
                                }

                                #region //UPDATE MES.Barcode回原先資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE MES.Barcode SET
                                        MoId = @MoId,
                                        CurrentMoProcessId = @CurrentMoProcessId,
                                        NextMoProcessId = @NextMoProcessId,
                                        CurrentProdStatus = @CurrentProdStatus,
                                        BarcodeStatus = @BarcodeStatus,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE BarcodeNo = @BarcodeNo";
                                dynamicParameters.AddDynamicParams(
                                  new
                                  {
                                      item2.MoId,
                                      item2.CurrentMoProcessId,
                                      item2.NextMoProcessId,
                                      item2.CurrentProdStatus,
                                      item2.BarcodeStatus,
                                      LastModifiedDate,
                                      LastModifiedBy,
                                      BarcodeNo
                                  });

                                rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region //將轉移LOG移除
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE FROM MES.MrBarcodeTransferLog
                                        WHERE LogId = @LogId";
                                dynamicParameters.Add("LogId", item2.LogId);

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }
                        }
                        else
                        {
                            //throw new SystemException("此條碼【" + BarcodeNo + "】查無紀錄LOG，請聯繫系統開發室!!");
                        }
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE MES.MrBarcodeReRegister
                                WHERE BarcodeReRegisterId = @BarcodeReRegisterId";
                        dynamicParameters.Add("BarcodeReRegisterId", BarcodeReRegisterId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //更新MrDetail領取套數
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE MES.MrDetail SET
                                ActualQuantity = @ActualQuantity,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE MrDetailId = @MrDetailId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ActualQuantity,
                                LastModifiedDate,
                                LastModifiedBy,
                                MrDetailId
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
        #endregion
    }
}
