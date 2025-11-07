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
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Transactions;
using System.Web;

namespace SCMDA
{
    public class PurchaseRequisitionDA
    {
        public string MainConnectionStrings = "";
        public string ErpConnectionStrings = "";
        public string HrmConnectionStrings = "";
        public string ErpSysDbConnectionStrings = "";
        public string BpmDbConnectionStrings = "";

        public string BpmServerPath = "";
        public string BpmAccount = "";
        public string BpmPassword = "";

        public int CurrentCompany = -1;
        public int CurrentUser = -1;
        public int CreateBy = -1;
        public int LastModifiedBy = -1;
        public DateTime CreateDate = default(DateTime);
        public DateTime LastModifiedDate = default(DateTime);

        public string sql = "";
        public JObject jsonResponse = new JObject();
        public Logger logger = LogManager.GetCurrentClassLogger();
        public SqlQuery sqlQuery = new SqlQuery();
        public DynamicParameters dynamicParameters = new DynamicParameters();
        public BpmHelper bpmHelper = new BpmHelper();

        public PurchaseRequisitionDA()
        {
            MainConnectionStrings = ConfigurationManager.AppSettings["MainDb"];
            HrmConnectionStrings = ConfigurationManager.AppSettings["HrmDb"];
            ErpSysDbConnectionStrings = ConfigurationManager.AppSettings["ErpSysDb"];
            BpmDbConnectionStrings = ConfigurationManager.AppSettings["BpmDb"];

            BpmServerPath = ConfigurationManager.AppSettings["BpmServerPath"];
            BpmAccount = ConfigurationManager.AppSettings["BpmAccount"];
            BpmPassword = ConfigurationManager.AppSettings["BpmPassword"];

            GetUserInfo();
            CreateDate = DateTime.Now;
            LastModifiedDate = DateTime.Now;
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

        #region //INSERT UserOperateLog 紀錄使用者操作流程
        private void AddUserOperateLog()
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        string FullFunctionName = HttpContext.Current.Request.Path;
                        string IpAddress = BaseHelper.ClientIP();
                        dynamicParameters = new DynamicParameters();

                        sql = @"INSERT INTO BAS.UserOperateLog (FullFunctionName, IpAddress
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                VALUES (@FullFunctionName, @IpAddress
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                FullFunctionName,
                                IpAddress,
                                CreateDate,
                                LastModifiedDate,
                                CreateBy,
                                LastModifiedBy
                            });

                        var insertResult = sqlConnection.Query(sql, dynamicParameters);
                    }
                    transactionScope.Complete();
                }
            }
            catch (Exception e)
            {
                logger.Error(e.Message);
            }
        }
        #endregion

        #region //Get
        #region //GetPurchaseRequisition -- 取得請購單資料 -- Ann 2023-01-04
        public string GetPurchaseRequisition(int PrId, string PrNo, string PrErpPrefix, string PrErpNo, string Edition, string PrStatus
            , string PrDateStartDate, string PrDateEndDate, string MtlItemNo, int DepartmentId, string PrErpFullNo, string BpmNo
            , int UserId, string PoErpFullNo
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
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
                        @"FROM SCM.PurchaseRequisition a
                        INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                        INNER JOIN BAS.Department c ON a.DepartmentId = c.DepartmentId
                        INNER JOIN BAS.[Status] d ON a.PrStatus = d.StatusNo AND d.StatusSchema = 'PurchaseRequisition.PrStatus'
                        INNER JOIN BAS.[Type] e ON a.Priority = e.TypeNo AND e.TypeSchema = 'PR.Priority'";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PrId", @" AND a.PrId = @PrId", PrId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PrErpPrefix", @" AND a.PrErpPrefix = @PrErpPrefix", PrErpPrefix);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PrErpFullNo", @" AND (a.PrErpPrefix + '-' + a.PrErpNo) LIKE '%' + @PrErpFullNo + '%'", PrErpFullNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "DepartmentId", @" AND a.DepartmentId = @DepartmentId", DepartmentId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserId", @" AND a.UserId = @UserId", UserId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PrErpNo", @" AND a.PrErpNo LIKE '%' + @PrErpNo + '%'", PrErpNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Edition", @" AND a.Edition LIKE '%' + @Edition + '%'", Edition);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "BpmNo", @" AND a.BpmNo LIKE '%' + @BpmNo + '%'", BpmNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PrNo", @" AND a.PrNo LIKE '%' + @PrNo + '%'", PrNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PrDateStartDate", @" AND a.PrDate >= @PrDateStartDate", PrDateStartDate.Length > 0 ? Convert.ToDateTime(PrDateStartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PrDateEndDate", @" AND a.PrDate <= @PrDateEndDate", PrDateEndDate.Length > 0 ? Convert.ToDateTime(PrDateEndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemNo", @" AND EXISTS (SELECT TOP 1 1 
                                                                                                                   FROM SCM.PrDetail x
                                                                                                                   INNER JOIN PDM.MtlItem xa ON x.MtlItemId = xa.MtlItemId
                                                                                                                   WHERE x.PrId = a.PrId 
                                                                                                                   AND xa.MtlItemNo LIKE '%' + @MtlItemNo + '%')", MtlItemNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PoErpFullNo", @" AND EXISTS (SELECT TOP 1 1 
                                                                                                                     FROM SCM.PoDetail x
                                                                                                                     WHERE x.PrErpPrefix = a.PrErpPrefix AND x.PrErpNo = a.PrErpNo 
                                                                                                                     AND x.PoErpPrefix + '-' + x.PoErpNo LIKE '%' + @PoErpFullNo + '%')", PoErpFullNo);
                    if (PrStatus.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PrStatus", @" AND a.PrStatus IN @PrStatus", PrStatus.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.PrId DESC";
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

        #region //GetPrSequence -- 取得請購單序號 -- Ann 2023-01-04
        public string GetPrSequence(int PrId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sql = @"SELECT COUNT(1) Count
                            FROM SCM.PrDetail
                            WHERE PrId = @PrId";
                    dynamicParameters.Add("PrId", PrId);

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

        #region //GetPrDetail -- 取得請購單詳細資料 -- Ann 2023-01-06
        public async Task<string> GetPrDetail(int PrDetailId, int PrId, string ConfirmStatus
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.PrDetailId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.PrId, a.PrSequence, a.MtlItemId, a.PrMtlItemName, a.PrMtlItemSpec, a.InventoryId, a.PrUomId, a.PrQty, FORMAT(a.DemandDate, 'yyyy-MM-dd') DemandDate
                        , a.SupplierId, a.PrCurrency, a.PrExchangeRate, a.PrUnitPrice, a.PrPrice, a.PrPriceTw, a.UrgentMtl, a.ProductionPlan, a.Project
                        , a.BudgetDepartmentNo, a.BudgetDepartmentSubject, a.SoDetailId, ISNULL(FORMAT(a.DeliveryDate, 'yyyy-MM-dd'), '') DeliveryDate, a.PoUserId, a.PoUomId, a.PoQty, a.PoCurrency, a.PoUnitPrice
                        , a.PoPrice, a.PoErpPrefixNo, a.LockStaus, a.PoStaus, a.PartialPurchaseStaus, a.InquiryStatus, a.TaxNo, a.Taxation, a.BusinessTaxRate, a.DetailMultiTax
                        , a.TradeTerm, a.PrPriceQty, a.PrPriceUomId, a.DiscountRate, a.DiscountAmount, a.MtlInventory, a.MtlInventoryQty, a.ConfirmStatus, a.ConfirmUserId
                        , a.ClosureStatus, ISNULL(a.PrDetailRemark, '') PrDetailRemark, a.PoRemark
                        , b.InventoryNo, b.InventoryName
                        , c.UomNo PrUomNo, c.UomName PrUomName
                        , ISNULL(d.SupplierNo, '') SupplierNo, ISNULL(d.SupplierName, '') SupplierName
                        , e.SoSequence, ISNULL(e.CustomerMtlItemNo, '') CustomerMtlItemNo
                        , f.SoErpPrefix, f.SoErpNo
                        , g.MtlItemNo, g.MtlModify
                        , h.CustomerName
                        , i.TypeName InquiryStatusName
                        , (
                            SELECT x.FileId
                            FROM SCM.PrFile x
                            WHERE PrId = a.PrId
                            AND PrDetailId = a.PrDetailId
                            FOR JSON PATH, ROOT('data')
                        ) PrFile";
                    sqlQuery.mainTables =
                        @"FROM SCM.PrDetail a
                        INNER JOIN SCM.Inventory b ON a.InventoryId = b.InventoryId
                        INNER JOIN PDM.UnitOfMeasure c ON a.PrUomId = c.UomId
                        LEFT JOIN SCM.Supplier d ON a.SupplierId = d.SupplierId
                        LEFT JOIN SCM.SoDetail e ON a.SoDetailId = e.SoDetailId
                        LEFT JOIN SCM.SaleOrder f ON e.SoId = f.SoId
                        INNER JOIN PDM.MtlItem g ON a.MtlItemId = g.MtlItemId
                        LEFT JOIN SCM.Customer h ON f.CustomerId = h.CustomerId
                        INNER JOIN BAS.[Type] i ON a.InquiryStatus = i.TypeNo AND TypeSchema = 'PrDetail.InquiryStatus'";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PrDetailId", @" AND a.PrDetailId = @PrDetailId", PrDetailId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PrId", @" AND a.PrId = @PrId", PrId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ConfirmStatus", @" AND a.ConfirmStatus = @ConfirmStatus", ConfirmStatus);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.PrDetailId";
                    sqlQuery.pageIndex = PageIndex;
                    sqlQuery.pageSize = PageSize;

                    var result = await BaseHelper.SqlQueryAsync(sqlConnection, dynamicParameters, sqlQuery);

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

        #region //GetPrSeq -- 取得請購單序號 -- Ann 2023-01-10
        public string GetPrSeq()
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT ISNULL(MAX(a.PrNo), 0000) MaxPrNo
                            FROM SCM.PurchaseRequisition a
                            WHERE FORMAT(a.CreateDate, 'yyyy-MM-dd') = @CreateDate";
                    dynamicParameters.Add("CreateDate", DateTime.Now.ToString("yyyy-MM-dd"));
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

        #region //GetPrModification -- 取得請購變更單資料 -- Ann 2023-02-07
        public string GetPrModification(int PrmId, string PrmStatus, string PrDateStartDate, string PrDateEndDate, string ModiDateStartDate, string ModiDateEndDate, int UserId
            , string PrErpFullNo, string PrNo, string Edition, string MtlItemNo
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.PrmId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.PrId, a.DepartmentId, a.UserId, a.BpmNo, a.Edition, FORMAT(a.ModiDate, 'yyyy-MM-dd') ModiDate, FORMAT(a.DocDate, 'yyyy-MM-dd') DocDate
                        , a.ModiReason, a.PrmRemark, a.TotalQty, a.Amount, a.BudgetDepartmentId, a.OriDepartmentId, a.OriUserId, a.OriEdition, FORMAT(a.OriPrDate, 'yyyy-MM-dd') OriPrDate
                        , a.OriPrRemark, a.OriAmount, a.OriBudgetDepartmentId, a.PrmStatus, a.SignupStaus, a.ConfirmStatus, a.ConfirmUserId, a.WholeClosureStatus
                        , a.BpmTransferStatus, a.BpmTransferUserId, a.BpmTransferDate, a.TransferStatus, a.TransferDate
                        , b.DepartmentNo, b.DepartmentName
                        , c.UserNo, c.UserName
                        , d.StatusNo, d.StatusName
                        , (
                            SELECT e.FileId
                            FROM SCM.PrmFile e
                            WHERe e.PrmId = a.PrmId
                            FOR JSON PATH, ROOT('data')
                        ) PrmFile
                        , f.PrNo, f.PrErpPrefix, f.PrErpNo, f.BpmNo, f.TotalQty OriTotalQty, FORMAT(f.PrDate, 'yyyy-MM-dd') PrDate
                        , g.TypeNo Priority";
                    sqlQuery.mainTables =
                        @"FROM SCM.PrModification a
                        INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                        INNER JOIN BAS.[User] c ON a.UserId = c.UserId
                        INNER JOIN BAS.[Status] d ON a.PrmStatus = d.StatusNo AND d.StatusSchema = 'PrModification.PrmStatus'
                        INNER JOIN SCM.PurchaseRequisition f ON a.PrId = f.PrId
                        INNER JOIN BAS.[Type] g ON a.Priority = g.TypeNo AND g.TypeSchema = 'PR.Priority'";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"AND a.CompanyId = @CompanyId";
                    dynamicParameters.Add("CompanyId", CurrentCompany);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PrmId", @" AND a.PrmId = @PrmId", PrmId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "UserId", @" AND a.UserId = @UserId", UserId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PrDateStartDate", @" AND f.PrDate >= @PrDateStartDate", PrDateStartDate.Length > 0 ? Convert.ToDateTime(PrDateStartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PrDateEndDate", @" AND f.PrDate <= @PrDateEndDate", PrDateEndDate.Length > 0 ? Convert.ToDateTime(PrDateEndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ModiDateStartDate", @" AND a.ModiDate >= @ModiDateStartDate", ModiDateStartDate.Length > 0 ? Convert.ToDateTime(ModiDateStartDate).ToString("yyyy-MM-dd 00:00:00") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "ModiDateEndDate", @" AND a.ModiDate <= @ModiDateEndDate", ModiDateEndDate.Length > 0 ? Convert.ToDateTime(ModiDateEndDate).ToString("yyyy-MM-dd 23:59:59") : "");
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PrErpFullNo", @" AND (f.PrErpPrefix + '-' + f.PrErpNo) LIKE '%' + @PrErpFullNo + '%'", PrErpFullNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PrNo", @" AND f.PrNo LIKE '%' + @PrNo + '%'", PrNo);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "Edition", @" AND a.Edition = @Edition", Edition);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemNo", @" AND EXISTS (SELECT TOP 1 1 
                                                                                                                   FROM SCM.PrmDetail x
                                                                                                                   INNER JOIN PDM.MtlItem xa ON x.MtlItemId = xa.MtlItemId
                                                                                                                   WHERE x.PrmId = a.PrmId 
                                                                                                                   AND xa.MtlItemNo LIKE '%' + @MtlItemNo + '%')", MtlItemNo);
                    if (PrmStatus.Length > 0) BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PrmStatus", @" AND a.PrmStatus IN @PrmStatus", PrmStatus.Split(','));
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.PrmId DESC";
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

        #region //GetPrmDetail -- 取得請購變更單詳細資料 -- Ann 2023-02-08
        public string GetPrmDetail(int PrmDetailId, int PrmId
            , string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sqlQuery.mainKey = "a.PrmDetailId";
                    sqlQuery.auxKey = "";
                    sqlQuery.columns =
                        @", a.PrmId, a.PrDetailId, a.PrmSequence, a.MtlItemId, a.PrMtlItemName, a.PrMtlItemSpec, a.InventoryId
                        , a.PrUomId, a.PrQty, FORMAT(a.DemandDate, 'yyyy-MM-dd') DemandDate, a.SupplierId, a.PrCurrency, a.PrExchangeRate, a.PrUnitPrice, a.PrPrice
                        , a.PrPriceTw, a.UrgentMtl, a.ProductionPlan, a.Project, a.BudgetDepartmentNo, a.BudgetDepartmentSubject
                        , a.SoDetailId, FORMAT(a.DeliveryDate, 'yyyy-MM-dd') DeliveryDate, a.PoUserId, a.PoUomId, a.PoQty, a.PoCurrency, a.PoUnitPrice, a.PoPrice, a.LockStaus
                        , a.PoStaus, a.TaxNo, a.Taxation, a.BusinessTaxRate, a.DetailMultiTax, a.TradeTerm, a.PrPriceQty, a.PrPriceUomId
                        , a.DiscountRate, a.DiscountAmount, a.MtlInventory, a.MtlInventoryQty, a.ConfirmStatus, a.ConfirmUserId
                        , a.ClosureStatus, a.PrDetailRemark, a.PoRemark, a.ModiReason, a.OriPrSequence, a.OriMtlItemId
                        , a.OriPrMtlItemName, a.OriPrMtlItemSpec, a.OriInventoryId, a.OriPrUomId, a.OriPrQty, a.OriDemandDate
                        , a.OriSupplierId, a.OriPrCurrency, a.OriPrExchangeRate, a.OriPrUnitPrice, a.OriPrPrice, a.OriPrPriceTw, a.OriUrgentMtl
                        , a.OriProductionPlan, a.OriProject, a.OriBudgetDepartmentNo, a.OriBudgetDepartmentSubject, a.OriSoDetailId
                        , a.OriDeliveryDate, a.OriPoUserId, a.OriPoUomId, a.OriPoQty, a.OriPoCurrency, a.OriPoUnitPrice, a.OriPoPrice
                        , a.OriLockStaus, a.OriPoStaus, a.OriTaxNo, a.OriTaxation, a.OriBusinessTaxRate, a.OriDetailMultiTax
                        , a.OriTradeTerm, a.OriPrPriceQty, a.OriPrPriceUomId, a.OriDiscountRate, a.OriDiscountAmount, a.OriClosureStatus
                        , a.OriPrDetailRemark, a.OriPoRemark
                        , b.InventoryNo, b.InventoryName
                        , c.UomNo PrUomNo, c.UomName PrUomName
                        , ISNULL(d.SupplierNo, '') SupplierNo, ISNULL(d.SupplierName, '') SupplierName
                        , e.SoSequence, ISNULL(e.CustomerMtlItemNo, '') CustomerMtlItemNo
                        , f.SoErpPrefix, f.SoErpNo
                        , g.MtlItemNo, g.MtlModify
                        , h.CustomerName
                        , i.PrSequence";
                    sqlQuery.mainTables =
                        @"FROM SCM.PrmDetail a
                        INNER JOIN SCM.Inventory b ON a.InventoryId = b.InventoryId
                        INNER JOIN PDM.UnitOfMeasure c ON a.PrUomId = c.UomId
                        LEFT JOIN SCM.Supplier d ON a.SupplierId = d.SupplierId
                        LEFT JOIN SCM.SoDetail e ON a.SoDetailId = e.SoDetailId
                        LEFT JOIN SCM.SaleOrder f ON e.SoId = f.SoId
                        INNER JOIN PDM.MtlItem g ON a.MtlItemId = g.MtlItemId
                        LEFT JOIN SCM.Customer h ON f.CustomerId = h.CustomerId
                        INNER JOIN SCM.PrDetail i ON a.PrDetailId = i.PrDetailId";
                    sqlQuery.auxTables = "";
                    string queryCondition = @"";
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PrmDetailId", @" AND a.PrmDetailId = @PrmDetailId", PrmDetailId);
                    BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PrmId", @" AND a.PrmId = @PrmId", PrmId);
                    sqlQuery.conditions = queryCondition;
                    sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.PrmDetailId";
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

        #region //GetPrmSequence -- 取得請購變更單序號 -- Ann 2023-02-08
        public string GetPrmSequence(int PrmId)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    DynamicParameters dynamicParameters = new DynamicParameters();
                    sql = @"SELECT COUNT(1) Count
                            FROM SCM.PrmDetail
                            WHERE PrmId = @PrmId";
                    dynamicParameters.Add("PrmId", PrmId);

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

        #region //GetMtlItemTotalUomInfo -- 取得品號所有可用單位資料 -- Ann 2023-06-08
        public string GetMtlItemTotalUomInfo(int MtlItemId)
        {
            try
            {
                if (MtlItemId <= 0) throw new SystemException("【品號】不能為空!");

                List<MtlItemUomInfo> mtlItemUomInfos = new List<MtlItemUomInfo>();

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
                                , b.UomId, b.UomNo
                                , (b.UomNo + ' ' + b.UomName) UomText
                                FROM PDM.MtlItem a
                                INNER JOIN PDM.UnitOfMeasure b ON a.InventoryUomId = b.UomId
                                WHERE a.MtlItemId = @MtlItemId";
                        dynamicParameters.Add("MtlItemId", MtlItemId);

                        var MtlItemResult = sqlConnection.Query(sql, dynamicParameters);

                        string MtlItemNo = "";
                        string UomText = "";
                        foreach (var item in MtlItemResult)
                        {
                            MtlItemNo = item.MtlItemNo;
                            UomText = item.UomText;

                            MtlItemUomInfo uomInfos = new MtlItemUomInfo()
                            {
                                UomId = item.UomId,
                                UomText = item.UomText,
                                UomNo = item.UomNo
                            };
                            mtlItemUomInfos.Add(uomInfos);
                        }
                        #endregion

                        #region //找ERP INVMD
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT LTRIM(RTRIM(MD002)) MD002
                                FROM INVMD
                                WHERE MD001 = @MD001";
                        dynamicParameters.Add("MD001", MtlItemNo);

                        var INVMDResult = sqlConnection2.Query(sql, dynamicParameters);

                        foreach (var item in INVMDResult)
                        {
                            bool checkInventoryFlag = false;
                            foreach (var item2 in mtlItemUomInfos)
                            {
                                if (item2.UomNo == item.MD002)
                                {
                                    checkInventoryFlag = true;
                                    break;
                                }
                            }
                            
                            if (checkInventoryFlag != false)
                            {
                                continue;
                            }

                            #region //合併回來MES
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.UomId
                                    , a.UomNo
                                    , (a.UomNo + ' ' + a.UomName) UomText
                                    FROM PDM.UnitOfMeasure a 
                                    WHERE a.UomNo = @UomNo
                                    AND a.CompanyId = @CompanyId";
                            dynamicParameters.Add("UomNo", item.MD002);
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            var UnitOfMeasureResult = sqlConnection.Query(sql, dynamicParameters);

                            if (UnitOfMeasureResult.Count() <= 0) throw new SystemException("ERP轉換單位【" + item.MD002 + "】資料錯誤!!!");

                            foreach (var item2 in UnitOfMeasureResult)
                            {
                                MtlItemUomInfo uomInfos = new MtlItemUomInfo()
                                {
                                    UomId = item2.UomId,
                                    UomText = item2.UomText,
                                    UomNo = item2.UomNo
                                };
                                mtlItemUomInfos.Add(uomInfos);
                            }
                            #endregion
                        }
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            data = mtlItemUomInfos
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

        #region //GetInventoryQty -- 取得品號庫存資料 -- Ann 2023-06-09
        public string GetInventoryQty(int MtlItemId, int InventoryId)
        {
            try
            {
                if (MtlItemId <= 0) throw new SystemException("【品號】不能為空!");
                if (InventoryId <= 0) throw new SystemException("【庫別】不能為空!");

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
                                WHERE a.InventoryId = @InventoryId";
                        dynamicParameters.Add("InventoryId", InventoryId);

                        var InventoryResult = sqlConnection.Query(sql, dynamicParameters);

                        string InventoryNo = "";
                        foreach (var item in InventoryResult)
                        {
                            InventoryNo = item.InventoryNo;
                        }
                        #endregion

                        #region //取得ERP庫存資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT SUM(MC007) MC007
                                FROM INVMC
                                WHERE MC001 = @MC001";
                        dynamicParameters.Add("MC001", MtlItemNo);

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

        #region //GetHistoryPrice -- 取得特定品號歷史價格 -- Ann 2025-06-12
        public string GetHistoryPrice(string MtlItemNo)
        {
            try
            {
                if (MtlItemNo.Length <= 0) throw new SystemException("【品號】不能為空!");

                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    dynamicParameters = new DynamicParameters();
                    sql = @"WITH MaxPrice AS (
                                SELECT TOP 1 a.PrUnitPrice AS MaxPrUnitPrice,
                                             a.PrSequence AS MaxPrSeq,
                                             a.PrCurrency AS MaxPrCurrency,
                                             FORMAT(a.CreateDate, 'yyyy-MM-dd') AS MaxCreateDate,
                                             c.PrErpPrefix + '-' + c.PrErpNo MaxPrErpFullNo
                                FROM SCM.PrDetail a
                                INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
                                INNER JOIN SCM.PurchaseRequisition c ON a.PrId = c.PrId
                                WHERE b.MtlItemNo = @MtlItemNo
                                AND c.ConfirmStatus = 'Y'
                                ORDER BY a.PrUnitPrice DESC
                            ),
                            MinPrice AS (
                                SELECT TOP 1 a.PrUnitPrice AS MinPrUnitPrice,
                                             a.PrSequence AS MinPrSeq,
                                             a.PrCurrency AS MinPrCurrency,
                                             FORMAT(a.CreateDate, 'yyyy-MM-dd') AS MinCreateDate,
                                             c.PrErpPrefix + '-' + c.PrErpNo MinPrErpFullNo
                                FROM SCM.PrDetail a
                                INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
                                INNER JOIN SCM.PurchaseRequisition c ON a.PrId = c.PrId
                                WHERE b.MtlItemNo = @MtlItemNo
                                AND c.ConfirmStatus = 'Y'
                                ORDER BY a.PrUnitPrice ASC
                            ),
                            TotalCount AS (
                                SELECT COUNT(*) AS TotalCount
                                FROM SCM.PrDetail a
                                INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
                                WHERE b.MtlItemNo = @MtlItemNo
                            )
                            SELECT 
                                MaxPrUnitPrice,
                                MaxPrCurrency,
                                MaxCreateDate,
                                MaxPrErpFullNo,
                                MaxPrSeq,
                                MinPrUnitPrice,
                                MinPrCurrency,
                                MinCreateDate,
                                MinPrErpFullNo,
                                MinPrSeq,
                                TotalCount
                            FROM MaxPrice, MinPrice, TotalCount;";
                    dynamicParameters.Add("MtlItemNo", MtlItemNo);
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

        #region //GetErpLocalCurrency -- 取得ERP本幣幣別 -- Ann 2023-06-13
        public string GetErpLocalCurrency()
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
                        sql = @"SELECT TOP 1 MA003 FROM CMSMA";

                        var CMSMAResult = sqlConnection2.Query(sql, dynamicParameters);
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            data = CMSMAResult
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

        #region//GetPrConfirmedNotProcured --請購單已確認但未採購的
        public string GetPrConfirmedNotProcured(string OrderBy, int PageIndex, int PageSize)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //確認公司別DB
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.ErpNo, a.ErpDb
                            FROM BAS.Company a
                            WHERE CompanyNo = @CompanyNo";
                    dynamicParameters.Add("CompanyNo", "JMO-CT");

                    var companyResult = sqlConnection.Query(sql, dynamicParameters);

                    if (companyResult.Count() <= 0) throw new SystemException("公司別錯誤!");

                    foreach (var item in companyResult)
                    {
                        ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                    }
                    #endregion

                    using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                    {
                        #region
                        DynamicParameters dynamicParameters = new DynamicParameters();
                        sqlQuery.mainKey = "PA.TA001,PA.TA002,PB.TB003";
                        sqlQuery.auxKey = "";
                        sqlQuery.columns =
                            @",PA.TA003 AS '請購日期', rtrim(Q.MQ001)+''+Q.MQ002 AS '單據名稱', 
	                            (CASE WHEN PA.TA007='Y' THEN '已確認'
		                            WHEN PA.TA007='N' THEN '未確認'
		                            WHEN PA.TA007='V' THEN '作廢' END) AS '確認碼', PA.TA014 AS '確認者編號', ISNULL(CU1.MV002,CA1.MF002) AS '確認者',
	                            PA.TA004 AS '部門編號', DP.ME002 AS '請購部門', PA.TA012 AS '人員工號', ISNULL(U1.MV002,U2.MF002) AS '請購人員', PA.TA006 AS '請購備註',
	                            PB.TB004 AS '品號', PB.TB005 AS '品名', PB.TB006 AS '規格', 
	                            CONCAT(G2.MB005, ISNULL(G2.MA1,'')) AS '會計分類名稱',CONCAT(G2.MB006, ISNULL(G2.MA2,'')) AS '倉管分類名稱', 
	                            CONCAT(G2.MB007, ISNULL(G2.MA3,'')) AS '業務分類名稱',CONCAT(G2.MB008, ISNULL(G2.MA4,'')) AS '生管分類名稱',
	                            PB.TB007 AS '請購單位', PB.TB009 AS '請購數量', 
	                            PB.TB043 AS '專案代號', ISNULL(CB.NB002,'') AS '專案名稱', PB.TB012 AS '品項備註', 
	                            PA.CREATE_DATE AS '請購單創建日',PA.CREATOR AS '請購單創建者編號', ISNULL(U3.MV002,U4.MF002) AS '請購單創建者',
	                            (CASE WHEN PB.TB039='N' THEN '未結案'
		                            WHEN PB.TB039='Y' THEN '自動結案'
		                            WHEN PB.TB039='y' THEN '指定結案' END) AS '結案碼',
	                            (CASE WHEN PB.TB026='1' THEN '內含'
		                            WHEN PB.TB026='2' THEN '外加'
		                            WHEN PB.TB026='3' THEN '零稅率'
		                            WHEN PB.TB026='4' THEN '免稅'
		                            WHEN PB.TB026='9' THEN '不計稅' END) AS '課稅別', PB.TB016 AS '幣別', PC.TC006 AS '匯率', PB.TB063 AS '營業稅率',
	                            PB.TB013 AS '人員帳號',ISNULL(U5.MV002,U6.MF002) AS '採購人員',PB.TB010 AS '廠商編號', MA.MA002 AS '廠商簡稱',
	                            PB.TB014 AS '採購數量',PB.TB015 AS '採購單位',PB.TB016 AS '採購幣別',PB.TB017 AS '採購單價',PB.TB018 AS '採購金額',
	                            CASE WHEN PB.TB026='1' THEN ROUND((PB.TB018/(1+PB.TB063)),2)
		                            WHEN PB.TB026='2' THEN ROUND((PB.TB018*1),2)
		                            WHEN PB.TB026='3' THEN PB.TB018*PC.TC006
		                            WHEN PB.TB026='4' THEN PB.TB018*PC.TC006 ELSE PB.TB018 END AS '本幣未稅採購金額'";
                        sqlQuery.mainTables =
                            @"FROM PURTA PA
	                            LEFT JOIN CMSMQ Q ON Q.MQ001=PA.TA001
	                            LEFT JOIN PURTB PB ON PA.TA001=PB.TB001 AND PA.TA002=PB.TB002
	                            LEFT JOIN (SELECT I.TB004,i.MB001,i.MB002,i.MB003,i.MB005,T1.MA003 AS MA1,i.MB006,T2.MA003 AS MA2,i.MB007,T3.MA003 AS MA3,i.MB008,T4.MA003 AS MA4
					                            FROM PURTB I
						                            LEFT JOIN INVMB i ON I.TB004=i.MB001
						                            LEFT JOIN INVMA T1 ON i.MB005=T1.MA002
						                            LEFT JOIN INVMA T2 ON i.MB006=T2.MA002
						                            LEFT JOIN INVMA T3 ON i.MB007=T3.MA002
						                            LEFT JOIN INVMA T4 ON i.MB008=T4.MA002
					                            WHERE TB025='Y'
					                            GROUP BY I.TB004,i.MB001,i.MB002,i.MB003,i.MB005,T1.MA003,i.MB006,T2.MA003,i.MB007,T3.MA003,i.MB008,T4.MA003
			　                            ) G2 ON G2.TB004=PB.TB004 
	                            LEFT JOIN CMSMV CU1 ON CU1.MV001=PA.TA014
	                            LEFT JOIN ADMMF CA1 ON CA1.MF001=PA.TA014
	                            LEFT JOIN PURMA MA ON PB.TB010=MA.MA001
	                            LEFT JOIN CMSMV U1 ON PA.TA012=U1.MV001
	                            LEFT JOIN ADMMF U2 ON PA.TA012=U2.MF001
	                            LEFT JOIN CMSME DP ON PA.TA004=DP.ME001
	                            LEFT JOIN CMSMV U3 ON PA.CREATOR=U3.MV001
	                            LEFT JOIN ADMMF U4 ON PA.CREATOR=U4.MF001
	                            LEFT JOIN CMSMV U5 ON PB.TB013=U5.MV001
	                            LEFT JOIN ADMMF U6 ON PB.TB013=U6.MF001
	                            LEFT JOIN PURTD PD ON PD.TD026=PB.TB001 AND PD.TD027=PB.TB002 AND PD.TD028=PB.TB003 AND PD.TD018='Y'
	                            LEFT JOIN PURTC PC ON PC.TC001=PD.TD001 AND PC.TC002=PD.TD002 AND PC.TC014='Y'
	                            LEFT JOIN CMSNB CB ON CB.NB001=PB.TB043";
                        sqlQuery.auxTables = "";
                        string queryCondition = @"AND PA.TA007='Y' AND PB.TB025='Y' AND PB.TB039='N' AND PB.TB021='N' AND PA.CREATE_DATE >= DATEADD(YEAR,-1,GETDATE())";
                        
                        sqlQuery.conditions = queryCondition;
                        sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "PA.TA001,PA.TA002,PB.TB003";
                        sqlQuery.pageIndex = PageIndex;
                        sqlQuery.pageSize = PageSize;
                        var result = BaseHelper.SqlQuery(sqlConnection2, dynamicParameters, sqlQuery);

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

        #region//GetSuggestionsForPurchase --粗胚建議請購清單
        public string GetSuggestionsForPurchase(string OrderBy, int PageIndex, int PageSize
            , string MtlItemNo, string MB005, string MB006, string MB007, string MB008,string Condition)
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                {
                    #region //確認公司別DB
                    dynamicParameters = new DynamicParameters();
                    sql = @"SELECT a.ErpNo, a.ErpDb
                            FROM BAS.Company a
                            WHERE CompanyNo = @CompanyNo";
                    dynamicParameters.Add("CompanyNo", "JMO-CT");

                    var companyResult = sqlConnection.Query(sql, dynamicParameters);

                    if (companyResult.Count() <= 0) throw new SystemException("公司別錯誤!");

                    foreach (var item in companyResult)
                    {
                        ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                    }
                    #endregion

                    using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                    {
                        string queryCondition = @"";
                        switch (Condition) {
                            case "MC004MC007":
                                queryCondition = " AND CB.MC004 > CB.MC007";
                                break;
                            default:
                                queryCondition = "";
                                break;
                        } 

                        #region
                        DynamicParameters dynamicParameters = new DynamicParameters();
                        sqlQuery.mainKey = "i.MB001";
                        sqlQuery.auxKey = "";
                        sqlQuery.columns =
                            @",i.MB002 品名,i.MB003 規格,i.MB004 庫存單位,
                                i.MB005+T1.MA003 第一分類, i.MB006+T2.MA003 第二分類, i.MB007+T3.MA003 第三分類, i.MB008+T4.MA003 第四分類,
                                i.MB039 最低補量, i.MB040 補貨倍量, CB.MC004 安全存量, CB.MC007 庫存數量,
                                CASE WHEN (i.MB039+(i.MB040*1)-CB.MC007)>CB.MC004 THEN (i.MB039+(i.MB040*1)-CB.MC007)
		                            WHEN (i.MB039+(i.MB040*2)-CB.MC007)>CB.MC004 THEN (i.MB039+(i.MB040*2)-CB.MC007)
		                            WHEN (i.MB039+(i.MB040*3)-CB.MC007)>CB.MC004 THEN (i.MB039+(i.MB040*3)-CB.MC007)
		                            WHEN (i.MB039+(i.MB040*4)-CB.MC007)>CB.MC004 THEN (i.MB039+(i.MB040*4)-CB.MC007)
		                            WHEN (i.MB039+(i.MB040*5)-CB.MC007)>CB.MC004 THEN (i.MB039+(i.MB040*5)-CB.MC007)
		                            WHEN (i.MB039+(i.MB040*6)-CB.MC007)>CB.MC004 THEN (i.MB039+(i.MB040*6)-CB.MC007)
		                            WHEN (i.MB039+(i.MB040*7)-CB.MC007)>CB.MC004 THEN (i.MB039+(i.MB040*7)-CB.MC007)
		                            WHEN (i.MB039+(i.MB040*8)-CB.MC007)>CB.MC004 THEN (i.MB039+(i.MB040*8)-CB.MC007)
		                            WHEN (i.MB039+(i.MB040*9)-CB.MC007)>CB.MC004 THEN (i.MB039+(i.MB040*9)-CB.MC007)
		                            WHEN (i.MB039+(i.MB040*10)-CB.MC007)>CB.MC004 THEN (i.MB039+(i.MB040*10)-CB.MC007) ELSE 0 END 建議採購數";
                        sqlQuery.mainTables =
                            @"FROM INVMB i
	                            LEFT JOIN INVMA T1 ON i.MB005=T1.MA002
	                            LEFT JOIN INVMA T2 ON i.MB006=T2.MA002
	                            LEFT JOIN INVMA T3 ON i.MB007=T3.MA002
	                            LEFT JOIN INVMA T4 ON i.MB008=T4.MA002
	                            INNER JOIN (SELECT C.MC001,SUM(C.MC004) MC004,SUM(C.MC007) MC007 
                                            FROM INVMC C
					                        LEFT JOIN INVMB B ON B.MB001=C.MC001
				                            GROUP BY C.MC001,C.MC004,C.MC007
	                            ) CB ON CB.MC001=i.MB001";
                        sqlQuery.auxTables = "";
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemNo", @" AND i.MB001 = @MtlItemNo", MtlItemNo);
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MB005", @" AND i.MB005 = @MB005", MB005);
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MB006", @" AND i.MB006 = @MB006", MB006);
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MB007", @" AND i.MB007 = @MB007", MB007);
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MB008", @" AND i.MB008 = @MB008", MB008);
                        sqlQuery.conditions = queryCondition;
                        sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "i.MB001";
                        sqlQuery.pageIndex = PageIndex;
                        sqlQuery.pageSize = PageSize;
                        var result = BaseHelper.SqlQuery(sqlConnection2, dynamicParameters, sqlQuery);
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

        #region//GetPurchaseRequisitionBPM --請購單BPM簽核歷程
        public string GetPurchaseRequisitionBPM(string OrderBy , int PageIndex , int PageSize , string PrNo )
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
                        #region
                        DynamicParameters dynamicParameters = new DynamicParameters();
                        sqlQuery.mainKey = "a.TskID";
                        sqlQuery.auxKey = "";
                        sqlQuery.columns =
                            @",c.ITEM16 AS BPM單號, 
                                    c.ITEM126 AS MES單號, 
                                    a11.UserName AS ApproverName,
                                    DATEADD(MS, CONVERT(BIGINT, a.StartTime) % 60000,
                                        DATEADD(MI, CONVERT(BIGINT, a.StartTime) / 60000, '1970-01-01 08:00:00.000')) AS 開始時間,
                                    DATEADD(MS, CONVERT(BIGINT, a.EndTime) % 60000,
                                        DATEADD(MI, CONVERT(BIGINT, a.EndTime) / 60000, '1970-01-01 08:00:00.000')) AS ApprovalTime,
                                    CONCAT(a7.DISPLAY_NAME, 
                                        CASE 
                                            WHEN a.IapSignResult = 'agree' THEN '核准'
                                            WHEN a.IapSignResult = 'disagree' THEN '抽單'
                                            WHEN a.IapSignResult = 'retrieved' THEN '退回'
                                            WHEN a.IapSignResult = 'nocomment' THEN '未審核'
                                            WHEN a.IapSignResult='' AND a.State = 'complete' THEN '核准'
                                        END) AS StepName,
                                    CONCAT(a6.NOTE, a.IapSignMessage) AS ApprovalComment,
                                    CASE 
                                        WHEN a.State = 'complete' THEN '已完成'
                                        WHEN a.State = 'queue' THEN '待關卡人員認領中'
                                        WHEN a.State = 'retrieved' THEN '已抽單'
                                    END AS ApprovalStatus,
                                    '250110' AS source_version
                                    ";
                        sqlQuery.mainTables =
                            @"FROM BPM.JMO.dbo.Task a
                                LEFT JOIN BPM.JMO.dbo.Mem_GenInf a1 ON a.MemID = a1.MemID
                                LEFT JOIN BPM.JMO.dbo.Mem_GenInf a11 ON a.ExeID = a11.MemID
                                LEFT JOIN BPM.JMO.dbo.Iap_GenInf a4 ON a4.ProID = a.ProID
                                LEFT JOIN BPM.JMO.dbo.Pro_GenInf a5 ON a5.ProID = a.ProID
                                LEFT JOIN BPM.JMO.dbo.PRO_SIGN_INS a6 ON a6.TASK_ID = a.TskID
                                LEFT JOIN BPM.JMO.dbo.PRO_SIGN_STATE a7 ON a7.AST_ID = a6.AST_ID
                                LEFT JOIN BPM.JMO.dbo.Task_ArtIns b ON b.TskID = a.TskID
                                LEFT JOIN BPM.JMO.dbo.ART00161720062813064_Ins c ON c.InsID = b.InsID";
                        sqlQuery.auxTables = "";
                        string queryCondition = @" AND a11.UserName != 'System Administrator'";
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "PrNo", @" AND c.ITEM126 = @PrNo", PrNo);
                        sqlQuery.conditions = queryCondition;
                        sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : " a.StartTime";
                        sqlQuery.pageIndex = PageIndex;
                        sqlQuery.pageSize = PageSize;
                        var result = BaseHelper.SqlQuery(sqlConnection2, dynamicParameters, sqlQuery);

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

        #region //Add
        #region //AddPurchaseRequisition -- 新增請購單資料 -- Ann 2023-01-04
        public string AddPurchaseRequisition(string PrErpPrefix, string DocDate, string PrDate, string PrRemark, string PrFile, int UserId, int DepartmentId, string Priority, string Source, string BomType)
        {
            try
            {
                string PrNo = "";
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
                            if (PrErpPrefix.Length < 0) throw new SystemException("【請購單別】不能為空!");
                            if (DocDate.Length < 0) throw new SystemException("【請購單據日期】不能為空!");
                            if (PrDate.Length < 0) throw new SystemException("【請購日期】不能為空!");
                            if (Priority.Length < 0) throw new SystemException("【優先度】不能為空!");
                            if (BomType.Length < 0) throw new SystemException("【請購物料類型】不能為空!");

                            #region //取號PrNo
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT ISNULL(MAX(a.PrNo), '000') MaxPrNo
                                    FROM SCM.PurchaseRequisition a
                                    WHERE FORMAT(a.CreateDate, 'yyyy-MM-dd') = @CreateDate
                                    AND a.CompanyId = @CurrentCompany";
                            dynamicParameters.Add("CreateDate", DateTime.Now.ToString("yyyy-MM-dd"));
                            dynamicParameters.Add("CurrentCompany", CurrentCompany);

                            var PrNoResult = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item in PrNoResult)
                            {
                                string PrDocDate = DateTime.Now.ToString("yyyyMMdd");
                                string MaxPrNo = item.MaxPrNo;
                                string MaxNo = MaxPrNo.Substring(MaxPrNo.Length - 3);
                                PrNo = "PR" + PrDocDate + (Convert.ToInt32(MaxNo) + 1).ToString().PadLeft(3, '0');
                            }
                            #endregion

                            #region //確認請購人員資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.[User] a
                                    WHERE a.UserId = @UserId";
                            dynamicParameters.Add("UserId", UserId);

                            var UserResult = sqlConnection.Query(sql, dynamicParameters);
                            if (UserResult.Count() <= 0) throw new SystemException("【使用者】資料錯誤，請重新輸入!");
                            #endregion

                            #region //確認請購部門資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.Department a
                                    WHERE a.DepartmentId = @DepartmentId
                                    AND a.Status = 'A'";
                            dynamicParameters.Add("DepartmentId", DepartmentId);

                            var DepartmentResult = sqlConnection.Query(sql, dynamicParameters);
                            if (DepartmentResult.Count() <= 0) throw new SystemException("【請購部門】資料錯誤或狀態非啟用中，請重新輸入!");
                            #endregion

                            #region //比對ERP關帳日期(庫存關帳)
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MA012)) MA012
                                    FROM CMSMA";
                            var cmsmaResult = sqlConnection2.Query(sql, dynamicParameters);
                            if (cmsmaResult.Count() < 0) throw new SystemException("EPR關帳日期資料錯誤!!");

                            foreach (var item in cmsmaResult)
                            {
                                string eprDate = item.MA012;
                                string erpYear = eprDate.Substring(0, 4);
                                string erpMonth = eprDate.Substring(4, 2);
                                string erpFullDate = erpYear + "-" + erpMonth;
                                DateTime erpDateTime = Convert.ToDateTime(erpFullDate);
                                DateTime DocDateDateTime = Convert.ToDateTime(DocDate.Split('-')[0] + "-" + DocDate.Split('-')[1]);
                                int compare = DocDateDateTime.CompareTo(erpDateTime);
                                if (compare <= 0) throw new SystemException("ERP已關帳(" + erpFullDate + ")，無法開立此日期(" + DocDate + ")之單據!!");
                            }
                            #endregion

                            string PrErpNo = BaseHelper.RandomCode(11);

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO SCM.PurchaseRequisition (PrNo, CompanyId, DepartmentId, UserId, PrErpPrefix, PrErpNo, Edition, PrDate, DocDate, PrRemark
                                    , TotalQty, Amount, PrStatus, SignupStaus, LockStaus, BomType, ConfirmStatus, BpmTransferStatus, TransferStatus, Priority, Source
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.PrId, INSERTED.PrErpPrefix, INSERTED.PrErpNo, INSERTED.PrNo
                                    VALUES (@PrNo, @CompanyId, @DepartmentId, @UserId, @PrErpPrefix, @PrErpNo, @Edition, @PrDate, @DocDate, @PrRemark
                                    , @TotalQty, @Amount, @PrStatus, @SignupStaus, @LockStaus, @BomType, @ConfirmStatus, @BpmTransferStatus, @TransferStatus, @Priority, @Source
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    PrNo,
                                    CompanyId = CurrentCompany,
                                    DepartmentId,
                                    UserId,
                                    PrErpPrefix,
                                    PrErpNo,
                                    Edition = "0000",
                                    PrDate,
                                    DocDate,
                                    PrRemark,
                                    TotalQty = 0,
                                    Amount = 0,
                                    PrStatus = "N",
                                    SignupStaus = "0",
                                    LockStaus = "N",
                                    BomType,
                                    ConfirmStatus = "N",
                                    BpmTransferStatus = "N",
                                    TransferStatus = "N",
                                    Priority,
                                    Source,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            int rowsAffected = insertResult.Count();

                            int PrId = -1;
                            foreach (var item in insertResult)
                            {
                                PrId = item.PrId;
                            }

                            #region //新增File
                            if (PrFile.Length > 0)
                            {
                                string[] prFiles = PrFile.Split(',');
                                foreach (var file in prFiles)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO SCM.PrFile (PrId, FileId
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.PrFileId
                                            VALUES (@PrId, @FileId
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            PrId,
                                            FileId = file,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    var insertResult3 = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult3.Count();
                                }
                            }
                            else
                            {
                                throw new SystemException("請購單必須上傳附件!");
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

        #region //AddPrDetail -- 新增請購單詳細資料 -- Ann 2023-01-05
        public string AddPrDetail(int PrId, string PrSequence, int MtlItemId, string PrMtlItemName, string PrMtlItemSpec, int InventoryId, int PrUomId, int PrQty, string DemandDate
            , int SupplierId, string PrCurrency, string PrExchangeRate, double PrUnitPrice, double PrPrice, double PrPriceTw
            , string UrgentMtl, string ProductionPlan, string Project, int SoDetailId, string PrDetailRemark, string PrFile)
        {
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
                            if (PrQty <= 0) throw new SystemException("【請購數量】不能為空!");
                            if (DemandDate.Length <= 0) throw new SystemException("【需求日期】不能為空!");
                            if (PrExchangeRate.Length <= 0) throw new SystemException("【匯率】不能為空!");
                            //if (PrPrice <= 0) throw new SystemException("【請購金額】不能為空!");
                            //if (PrPriceTw <= 0) throw new SystemException("【本幣金額】不能為空!");
                            if (UrgentMtl.Length <= 0) throw new SystemException("【是否急料】不能為空!");
                            if (ProductionPlan.Length <= 0) throw new SystemException("【是否納入生產計畫】不能為空!");
                            if (PrMtlItemName.Length <= 0) throw new SystemException("【品名】不能為空!");

                            #region //確認請購單資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.PrErpPrefix, ISNULL(a.TotalQty, 0) TotalQty, ISNULL(a.Amount, 0) Amount
                                    , a.UserId
                                    FROM SCM.PurchaseRequisition a
                                    WHERE a.PrId = @PrId";
                            dynamicParameters.Add("PrId", PrId);

                            var purchaseRequisitionResult = sqlConnection.Query(sql, dynamicParameters);

                            if (purchaseRequisitionResult.Count() <= 0) throw new SystemException("【請購單】資料錯誤!");

                            double TotalQty = -1;
                            double Amount = -1;
                            string PrErpPrefix = "";
                            foreach (var item in purchaseRequisitionResult)
                            {
                                if (item.UserId != CreateBy) throw new SystemException("無法修改其他請購人之單據!");
                                TotalQty = item.TotalQty;
                                Amount = item.Amount;
                                PrErpPrefix = item.PrErpPrefix;
                            }
                            #endregion

                            #region //確認請購單序號是否重複
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM SCM.PrDetail a
                                    WHERE a.PrId = @PrId
                                    AND a.PrSequence = @PrSequence";
                            dynamicParameters.Add("PrId", PrId);
                            dynamicParameters.Add("PrSequence", PrSequence);

                            var result = sqlConnection.Query(sql, dynamicParameters);

                            if (result.Count() > 0) throw new SystemException("【請購單序號】重複!");
                            #endregion

                            #region //判斷品號資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.MtlItemNo
                                    FROM PDM.MtlItem a
                                    WHERE a.MtlItemId = @MtlItemId";
                            dynamicParameters.Add("MtlItemId", MtlItemId);

                            var result2 = sqlConnection.Query(sql, dynamicParameters);

                            if (result2.Count() <= 0) throw new SystemException("【品號】資料錯誤!");

                            string MtlItemNo = "";
                            foreach (var item in result2)
                            {
                                MtlItemNo = item.MtlItemNo;
                            }
                            #endregion

                            #region //判斷ERP品號生效日與失效日
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MB030)) MB030, LTRIM(RTRIM(MB031)) MB031
                                    FROM INVMB
                                    WHERE MB001 = @MB001";
                            dynamicParameters.Add("MB001", MtlItemNo);

                            var INVMBResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (INVMBResult.Count() <= 0) throw new SystemException("此品號不存在於ERP中!!");

                            foreach (var item in INVMBResult)
                            {
                                if (item.MB030 != "" && item.MB030 != null)
                                {
                                    #region //判斷日期需大於或等於生效日
                                    string EffectiveDate = item.MB030;
                                    string effYear = EffectiveDate.Substring(0, 4);
                                    string effMonth = EffectiveDate.Substring(4, 2);
                                    string effDay = EffectiveDate.Substring(6, 2);
                                    DateTime effFullDate = Convert.ToDateTime(effYear + "-" + effMonth + "-" + effDay);
                                    int effresult = DateTime.Compare(CreateDate, effFullDate);
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
                                    int effresult = DateTime.Compare(CreateDate, effFullDate);
                                    if (effresult > 0) throw new SystemException("此品號已失效，無法使用!!");
                                    #endregion
                                }
                            }
                            #endregion

                            #region //判斷庫別資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.InventoryNo, a.InventoryName
                                    FROM SCM.Inventory a
                                    WHERE a.InventoryId = @InventoryId";
                            dynamicParameters.Add("InventoryId", InventoryId);

                            var result3 = sqlConnection.Query(sql, dynamicParameters);

                            if (result3.Count() <= 0) throw new SystemException("【庫別】資料錯誤!");

                            string InventoryNo = "";
                            string InventoryName = "";
                            foreach (var item in result3)
                            {
                                InventoryNo = item.InventoryNo;
                                InventoryName = item.InventoryName;
                            }
                            #endregion

                            #region //判斷單位資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM PDM.UnitOfMeasure a
                                    WHERE a.UomId = @UomId";
                            dynamicParameters.Add("UomId", PrUomId);

                            var result4 = sqlConnection.Query(sql, dynamicParameters);

                            if (result4.Count() <= 0) throw new SystemException("【單位】資料錯誤!");
                            #endregion

                            #region //取得ERP庫存資訊
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(a.MC007)) InventoryQty
                                    FROM INVMC a
                                    WHERE a.MC001 = @MC001
                                    AND a.MC002 = @MC002";
                            dynamicParameters.Add("MC001", MtlItemNo);
                            dynamicParameters.Add("MC002", InventoryNo);

                            var result5 = sqlConnection2.Query(sql, dynamicParameters);

                            double InventoryQty = 0;
                            string mtlInventory = "目前尚無資料";
                            if (result5.Count() > 0)
                            {
                                foreach (var item in result5)
                                {
                                    InventoryQty = Convert.ToDouble(item.InventoryQty);
                                    #region //組MtlInventory
                                    List<MtlInventory> mtlInventories = new List<MtlInventory>
                                    {
                                        new MtlInventory
                                        {
                                            WAREHOUSE_NO = InventoryNo,
                                            WAREHOUSE_NAME = InventoryName,
                                            WAREHOUSE_QTY = InventoryQty
                                        }
                                    };
                                    #endregion

                                    mtlInventory = JsonConvert.SerializeObject(mtlInventories);
                                    mtlInventory = "{\"data\":" + mtlInventory + "}";
                                }
                            }
                            #endregion

                            #region //檢查供應商資料是否正確
                            string TaxNo = "";
                            string Taxation = "";
                            string TradeTerm = "";
                            double? BusinessTaxRate = -1;
                            string HideSupplier = "";
                            if (SupplierId > 0)
                            {
                                #region //供應商
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.TaxNo, a.Taxation, a.TradeTerm, a.HideSupplier
                                        FROM SCM.Supplier a
                                        WHERE a.SupplierId = @SupplierId";
                                dynamicParameters.Add("SupplierId", SupplierId);

                                var result6 = sqlConnection.Query(sql, dynamicParameters);

                                if (result6.Count() <= 0) throw new SystemException("【供應商】資料錯誤!");

                                foreach (var item in result6)
                                {
                                    TaxNo = item.TaxNo;
                                    Taxation = item.Taxation;
                                    TradeTerm = item.TradeTerm;
                                    HideSupplier = item.HideSupplier;
                                }
                                #endregion

                                #region //查詢營業稅額資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT ROUND(LTRIM(RTRIM(NN004)),2) BusinessTaxRate
                                        FROM CMSNN 
                                        WHERE NN001 = @TaxNo";
                                dynamicParameters.Add("TaxNo", TaxNo);

                                var businessTaxRateResult = sqlConnection2.Query(sql, dynamicParameters);

                                if (businessTaxRateResult.Count() <= 0) throw new SystemException("【課稅別】資料錯誤!");

                                foreach (var item in businessTaxRateResult)
                                {
                                    BusinessTaxRate = item.BusinessTaxRate;
                                }
                                #endregion
                            }
                            #endregion

                            #region //確認集團內/集團外邏輯
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 b.HideSupplier
                                    FROM SCM.PrDetail a 
                                    INNER JOIN SCM.Supplier b ON a.SupplierId = b.SupplierId
                                    WHERE a.PrId = @PrId
                                    ORDER BY a.PrSequence";
                            dynamicParameters.Add("PrId", PrId);

                            var PrDetailResult = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item in PrDetailResult)
                            {
                                string HideSupplierString = "";
                                if (item.HideSupplier == "Y")
                                {
                                    HideSupplierString = "集團內";
                                }
                                else
                                {
                                    HideSupplierString = "集團外";
                                }
                                if (item.HideSupplier != HideSupplier) throw new SystemException("此請購單第一筆單身為【" + HideSupplierString + "】，無法新增!!");
                            }
                            #endregion

                            #region //判斷幣別資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM CMSMF
                                    WHERE MF001 = @MF001";
                            dynamicParameters.Add("MF001", PrCurrency);

                            var result7 = sqlConnection2.Query(sql, dynamicParameters);
                            if (result7.Count() <= 0) throw new SystemException("【幣別】資料有誤!");
                            #endregion

                            #region //判斷MES專案代碼資料及專案預算卡控
                            double budgetAmount = 0;
                            if (Project.Length > 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.LocalBudgetAmount
                                        FROM SCM.ProjectDetail a 
                                        INNER JOIN SCM.Project b ON a.ProjectId = b.ProjectId
                                        WHERE a.ProjectType = '1'
                                        AND b.ProjectNo = @ProjectNo
                                        AND b.CompanyId = @CompanyId
                                        AND a.Status = 'Y'";
                                dynamicParameters.Add("ProjectNo", Project);
                                dynamicParameters.Add("CompanyId", CurrentCompany);

                                var ProjectDetailResult = sqlConnection.Query(sql, dynamicParameters);

                                foreach (var item in ProjectDetailResult)
                                {
                                    budgetAmount = item.LocalBudgetAmount;
                                }

                                #region //專案預算卡控
                                //此專案掛鉤的全部請購單金額
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT ISNULL(SUM(a.PrPriceTw), 0) TotalLocalAmount
                                        FROM SCM.PrDetail a 
                                        INNER JOIN SCM.Project b ON a.Project = b.ProjectNo
                                        INNER JOIN SCM.PurchaseRequisition c ON a.PrId = c.PrId
                                        WHERE b.ProjectNo = @ProjectNo
                                        AND b.CompanyId = @CompanyId
                                        AND (c.PrStatus IN ('Y', 'P')
                                        OR a.PrId = @PrId)";
                                dynamicParameters.Add("ProjectNo", Project);
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                dynamicParameters.Add("PrId", PrId);

                                var ProjectAmountResult = sqlConnection.Query(sql, dynamicParameters);

                                foreach (var item in ProjectAmountResult)
                                {
                                    double totalLocalAmount = item.TotalLocalAmount;
                                    if (totalLocalAmount + PrPriceTw > budgetAmount)
                                    {
                                        throw new SystemException("此次請購金額合計已超過專案預算金額(" + budgetAmount + ")!");
                                    }
                                }
                                #endregion
                            }
                            #endregion

                            #region //判斷ERP專案代碼資料是否正確
                            if (Project.Length > 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM CMSNB
                                        WHERE NB001 = @NB001
                                        AND NB006 = 'N'";
                                dynamicParameters.Add("NB001", Project);

                                var result8 = sqlConnection2.Query(sql, dynamicParameters);
                                if (result8.Count() <= 0) throw new SystemException("【ERP專案代碼】資料有誤!");
                            }
                            #endregion

                            #region //確認訂單資料是否正確
                            if (SoDetailId > 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT MtlItemId, ConfirmStatus
                                        FROM SCM.SoDetail
                                        WHERE SoDetailId = @SoDetailId";
                                dynamicParameters.Add("SoDetailId", SoDetailId);

                                var result9 = sqlConnection.Query(sql, dynamicParameters);
                                if (result9.Count() <= 0) throw new SystemException("【訂單】資料有誤!");

                                foreach (var item in result9)
                                {
                                    if (item.ConfirmStatus != "Y") throw new SystemException("訂單尚未核單，無法綁定!!");
                                    //if (item.MtlItemId != MtlItemId) throw new SystemException("訂單品號與請購品號不同!!");
                                }
                            }
                            #endregion

                            #region //檢核特定相關卡控(目前僅晶彩邏輯)
                            if (CurrentCompany == 4)
                            {
                                double localPrUnitPrice = PrUnitPrice * Convert.ToDouble(PrExchangeRate);

                                #region //單別3109、3108特別卡控
                                if (PrErpPrefix == "3108" || PrErpPrefix == "3109")
                                {
                                    #region //請購單價不能超過1000(本幣)
                                    if (localPrUnitPrice > 1000) throw new SystemException("違反請購單別【3108、3109】卡控:<br>請購單價不能超過【1000】!!<br>此次請購單價【" + localPrUnitPrice + "】已超過!!");
                                    #endregion

                                    #region //所有單身合計金額(本幣)不能超過20000
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT ISNULL(SUM(PrPriceTw), 0) TotalPrice
                                            FROM SCM.PrDetail
                                            WHERE PrId = @PrId";
                                    dynamicParameters.Add("PrId", PrId);

                                    var TotalPriceResult = sqlConnection.Query(sql, dynamicParameters);

                                    double totalPrice = 0;
                                    foreach (var item in TotalPriceResult)
                                    {
                                        totalPrice = item.TotalPrice;
                                        if (totalPrice + PrPriceTw > 20000) throw new SystemException("違反請購單別【3108、3109】卡控:<br>請購合計金額不能超過20000!!<br>目前請購合計金額【" + (totalPrice + PrPriceTw) + "】已超過!!");
                                    }
                                    #endregion
                                }
                                #endregion
                            }
                            #endregion

                            #region //確認品號是否需要上傳相關文件才能進行請購(除晶彩外)
                            if (CurrentCompany == 11)
                            {
                                string pattern = @"^.*-T$";
                                if (Regex.IsMatch(MtlItemNo, pattern))
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM PDM.MtlItem a 
                                            OUTER APPLY (
                                                SELECT x.MtDocId, x.DocName
                                                FROM PDM.MtlItemDocSetting x 
                                                WHERE x.PurchaseMandatory = 'Y'
                                            ) x
                                            LEFT JOIN PDM.MtlItemFile b ON a.MtlItemNo = b.MtlItemNo AND b.MtDocId = x.MtDocId
                                            WHERE a.MtlItemNo = @MtlItemNo
                                            AND b.MtDocId IS NULL";
                                    dynamicParameters.Add("MtlItemNo", MtlItemNo);

                                    var MtlItemFileResult = sqlConnection.Query(sql, dynamicParameters);

                                    if (MtlItemFileResult.Count() > 0) throw new SystemException("品號【" + MtlItemNo + "】需上傳相關文件才可進行請購!!");
                                }
                            }
                            #endregion

                            #region //INSERT SCM.PrDetail
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO SCM.PrDetail (PrId, PrSequence, MtlItemId, PrMtlItemName, PrMtlItemSpec, InventoryId, PrUomId, PrQty, DemandDate, SupplierId, PrCurrency, PrExchangeRate, PrUnitPrice, PrPrice, PrPriceTw, UrgentMtl, ProductionPlan, Project, SoDetailId
                                    , PoUserId, PoUomId, PoQty, PoCurrency, PoUnitPrice, PoPrice, LockStaus, PoStaus, PartialPurchaseStaus, InquiryStatus, TaxNo, Taxation, BusinessTaxRate, DetailMultiTax, TradeTerm, PrPriceQty, PrPriceUomId, MtlInventory, MtlInventoryQty, ConfirmStatus, ClosureStatus, PrDetailRemark, PoRemark
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.PrDetailId
                                    VALUES (@PrId, @PrSequence, @MtlItemId, @PrMtlItemName, @PrMtlItemSpec, @InventoryId, @PrUomId, @PrQty, @DemandDate, @SupplierId, @PrCurrency, @PrExchangeRate, @PrUnitPrice, @PrPrice, @PrPriceTw, @UrgentMtl, @ProductionPlan, @Project, @SoDetailId
                                    , @PoUserId, @PoUomId, @PoQty, @PoCurrency, @PoUnitPrice, @PoPrice, @LockStaus, @PoStaus, @PartialPurchaseStaus, @InquiryStatus, @TaxNo, @Taxation, @BusinessTaxRate, @DetailMultiTax, @TradeTerm, @PrPriceQty, @PrPriceUomId, @MtlInventory, @MtlInventoryQty, @ConfirmStatus, @ClosureStatus, @PrDetailRemark, @PoRemark
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    PrId,
                                    PrSequence,
                                    MtlItemId,
                                    PrMtlItemName,
                                    PrMtlItemSpec,
                                    InventoryId,
                                    PrUomId,
                                    PrQty,
                                    DemandDate,
                                    SupplierId,
                                    PrCurrency,
                                    PrExchangeRate,
                                    PrUnitPrice,
                                    PrPrice,
                                    PrPriceTw,
                                    UrgentMtl,
                                    ProductionPlan,
                                    Project,
                                    SoDetailId,
                                    PoUserId = (int?)null,
                                    PoUomId = PrUomId,
                                    PoQty = PrQty,
                                    PoCurrency = PrCurrency,
                                    PoUnitPrice = PrUnitPrice,
                                    PoPrice = PrPrice,
                                    LockStaus = "N",
                                    PoStaus = "N",
                                    PartialPurchaseStaus = "N",
                                    InquiryStatus = "1",
                                    TaxNo = TaxNo != "" ? TaxNo : null,
                                    Taxation = Taxation != "" ? Taxation : null,
                                    BusinessTaxRate = BusinessTaxRate > 0 ? BusinessTaxRate : (double?)null,
                                    DetailMultiTax = "N",
                                    TradeTerm = TradeTerm != "" ? TradeTerm : null,
                                    PrPriceQty = PrQty,
                                    PrPriceUomId = PrUomId,
                                    MtlInventory = mtlInventory,
                                    MtlInventoryQty = InventoryQty,
                                    ConfirmStatus = "N",
                                    ClosureStatus = "N",
                                    PrDetailRemark,
                                    PoRemark = PrDetailRemark,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            int rowsAffected = insertResult.Count();

                            int PrDetailId = -1;
                            foreach (var item in insertResult)
                            {
                                PrDetailId = item.PrDetailId;
                            }
                            #endregion

                            #region //更新單頭總採購數量及金額
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.PurchaseRequisition SET
                                    TotalQty = TotalQty + @PrQty,
                                    Amount = Amount + @PrPriceTw,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE PrId = @PrId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    PrQty,
                                    PrPriceTw,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    PrId
                                });

                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //新增File
                            if (PrFile.Length > 0)
                            {
                                string[] prFiles = PrFile.Split(',');
                                foreach (var file in prFiles)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO SCM.PrFile (PrId, PrDetailId, FileId
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.PrFileId
                                            VALUES (@PrId, @PrDetailId, @FileId
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            PrId,
                                            PrDetailId,
                                            FileId = file,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    var insertResult3 = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult3.Count();
                                }
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

        #region //AddPrModification -- 新增請購變更單資料 -- Ann 2023-02-07
        public string AddPrModification(int PrId, int UserId, int DepartmentId, string DocDate, string ModiReason, string PrmRemark, string PrmFile, string Priority)
        {
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
                            if (DocDate.Length < 0) throw new SystemException("【請購單據日期】不能為空!");

                            #region //比對ERP關帳日期(庫存關帳)
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MA012)) MA012
                                    FROM CMSMA";
                            var cmsmaResult = sqlConnection2.Query(sql, dynamicParameters);
                            if (cmsmaResult.Count() < 0) throw new SystemException("EPR關帳日期資料錯誤!!");

                            foreach (var item in cmsmaResult)
                            {
                                string eprDate = item.MA012;
                                string erpYear = eprDate.Substring(0, 4);
                                string erpMonth = eprDate.Substring(4, 2);
                                string erpFullDate = erpYear + "-" + erpMonth;
                                DateTime erpDateTime = Convert.ToDateTime(erpFullDate);
                                DateTime DocDateDateTime = Convert.ToDateTime(DocDate.Split('-')[0] + "-" + DocDate.Split('-')[1]);
                                int compare = DocDateDateTime.CompareTo(erpDateTime);
                                if (compare <= 0) throw new SystemException("ERP已關帳(" + erpFullDate + ")，無法開立此日期(" + DocDate + ")之單據!!");
                            }
                            #endregion

                            #region //確認是否已經有此筆請購單之變更紀錄
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 (MAX(b.Edition) + 1) Edition
                                    FROM SCM.PrModification a
                                    INNER JOIN SCM.PurchaseRequisition b ON a.PrId = b.PrId
                                    WHERE a.PrId = @PrId";
                            dynamicParameters.Add("PrId", PrId);

                            var CheckPrIdResult = sqlConnection.Query(sql, dynamicParameters);

                            string Edition = "";
                            foreach (var item in CheckPrIdResult)
                            {
                                if (item.Edition == null)
                                {
                                    Edition = "0001";
                                }
                                else
                                {
                                    Edition = item.Edition.ToString().PadLeft(4, '0');
                                }
                            }
                            #endregion

                            #region //確認請購單及版次是否重複
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM SCM.PrModification a
                                    WHERE a.PrId = @PrId
                                    AND Edition = @Edition
                                    AND a.PrmStatus != 'S'";
                            dynamicParameters.Add("PrId", PrId);
                            dynamicParameters.Add("Edition", Edition);

                            var CheckPrEditionResult = sqlConnection.Query(sql, dynamicParameters);

                            if (CheckPrEditionResult.Count() > 0) throw new SystemException("已存在此請購單版次!");
                            #endregion

                            #region //確認原請購單資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.BpmNo, a.DepartmentId OriDepartmentId, a.UserId OriUserId
                                    , a.Edition OriEdition, a.PrDate OriPrDate, a.PrRemark OriPrRemark, a.Amount OriAmount
                                    , a.BudgetDepartmentId OriBudgetDepartmentId
                                    , a.PrErpPrefix, a.PrErpNo
                                    FROM SCM.PurchaseRequisition a
                                    WHERE a.PrId = @PrId
                                    AND a.CompanyId = @CompanyId";
                            dynamicParameters.Add("PrId", PrId);
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            var PrResult = sqlConnection.Query(sql, dynamicParameters);
                            if (PrResult.Count() <= 0) throw new SystemException("【請購單】資料錯誤!");

                            string BpmNo = "";
                            int OriDepartmentId = -1;
                            int OriUserId = -1;
                            string OriEdition = "";
                            DateTime OriPrDate = new DateTime();
                            string OriPrRemark = "";
                            double OriAmount = -1;
                            int? OriBudgetDepartmentId = -1;
                            string PrErpPrefix = "";
                            string PrErpNo = "";
                            foreach (var item in PrResult)
                            {
                                BpmNo = item.BpmNo;
                                OriDepartmentId = item.OriDepartmentId;
                                OriUserId = item.OriUserId;
                                OriEdition = item.OriEdition;
                                OriPrDate = item.OriPrDate;
                                OriPrRemark = item.OriPrRemark;
                                OriAmount = item.OriAmount;
                                OriBudgetDepartmentId = item.OriBudgetDepartmentId;
                                PrErpPrefix = item.PrErpPrefix;
                                PrErpNo = item.PrErpNo;

                            }
                            #endregion

                            #region //確認請購人員資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.[User] a
                                    WHERE a.UserId = @UserId";
                            dynamicParameters.Add("UserId", UserId);

                            var UserResult = sqlConnection.Query(sql, dynamicParameters);
                            if (UserResult.Count() <= 0) throw new SystemException("【使用者】資料錯誤，請重新輸入!");
                            #endregion

                            #region //確認請購部門資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.Department a
                                    WHERE a.DepartmentId = @DepartmentId";
                            dynamicParameters.Add("DepartmentId", DepartmentId);

                            var DepartmentResult = sqlConnection.Query(sql, dynamicParameters);
                            if (DepartmentResult.Count() <= 0) throw new SystemException("【請購部門】資料錯誤，請重新輸入!");
                            #endregion

                            #region //該請購單是否有開立採購單
                            dynamicParameters = new DynamicParameters();
                            sql = @"select * 
                                    from PURTD
                                    where TD026 = @PrErpPrefix
                                    and TD027 = @PrErpNo";
                            dynamicParameters.Add("PrErpPrefix", PrErpPrefix);
                            dynamicParameters.Add("PrErpNo", PrErpNo);

                            var PoERPResult = sqlConnection2.Query(sql, dynamicParameters);
                            if (PoERPResult.Count() > 0) throw new SystemException("該請購單已有採購單，無法開立變更單");

                            #endregion


                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO SCM.PrModification (PrId, CompanyId, DepartmentId, UserId, BpmNo, Edition, ModiDate, DocDate, ModiReason, PrmRemark, TotalQty
                                    , Amount, BudgetDepartmentId, OriDepartmentId, OriUserId, OriEdition, OriPrDate, OriPrRemark, OriAmount, OriBudgetDepartmentId
                                    , PrmStatus, SignupStaus, ConfirmStatus, ConfirmUserId, WholeClosureStatus, BpmTransferStatus, BpmTransferUserId, BpmTransferDate
                                    , TransferStatus, TransferDate, Priority
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.PrmId
                                    VALUES (@PrId, @CompanyId, @DepartmentId, @UserId, @BpmNo, @Edition, @ModiDate, @DocDate, @ModiReason, @PrmRemark, @TotalQty
                                    , @Amount, @BudgetDepartmentId, @OriDepartmentId, @OriUserId, @OriEdition, @OriPrDate, @OriPrRemark, @OriAmount, @OriBudgetDepartmentId
                                    , @PrmStatus, @SignupStaus, @ConfirmStatus, @ConfirmUserId, @WholeClosureStatus, @BpmTransferStatus, @BpmTransferUserId, @BpmTransferDate
                                    , @TransferStatus, @TransferDate, @Priority
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    PrId,
                                    CompanyId = CurrentCompany,
                                    DepartmentId,
                                    UserId,
                                    BpmNo,
                                    Edition,
                                    ModiDate = DateTime.Now.ToString("yyyy-MM-dd"),
                                    DocDate,
                                    ModiReason,
                                    PrmRemark,
                                    TotalQty = 0,
                                    Amount = 0,
                                    BudgetDepartmentId = (int?)null,
                                    OriDepartmentId,
                                    OriUserId,
                                    OriEdition,
                                    OriPrDate,
                                    OriPrRemark,
                                    OriAmount,
                                    OriBudgetDepartmentId,
                                    PrmStatus = "N",
                                    SignupStaus = "0",
                                    ConfirmStatus = "N",
                                    ConfirmUserId = (int?)null,
                                    WholeClosureStatus = "N",
                                    BpmTransferStatus = "N",
                                    BpmTransferUserId = (int?)null,
                                    BpmTransferDate = (DateTime?)null,
                                    TransferStatus = "N",
                                    TransferDate = (DateTime?)null,
                                    Priority,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            int rowsAffected = insertResult.Count();

                            int PrmId = -1;
                            foreach (var item in insertResult)
                            {
                                PrmId = item.PrmId;
                            }

                            #region //新增File
                            if (PrmFile.Length > 0)
                            {
                                string[] prmFiles = PrmFile.Split(',');
                                foreach (var file in prmFiles)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO SCM.PrmFile (PrmId, FileId
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.PrmFileId
                                            VALUES (@PrmId, @FileId
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            PrmId,
                                            FileId = file,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    var insertResult3 = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult3.Count();
                                }
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

        #region //AddPrmDetail -- 新增請購變更單詳細資料 -- Ann 2023-02-08//
        public string AddPrmDetail(int PrmId, int PrDetailId, string PrmSequence, int MtlItemId, string PrMtlItemName, string PrMtlItemSpec, int InventoryId, int PrUomId, int PrQty, string DemandDate
            , int SupplierId, string PrCurrency, string PrExchangeRate, double PrUnitPrice, double PrPrice, double PrPriceTw
            , string UrgentMtl, string ProductionPlan, string Project, int SoDetailId, string PoRemark, string PrDetailRemark, string ModiReason, string ClosureStatus)
        {
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
                            if (PrQty < 0) throw new SystemException("【請購數量】不能為空!");
                            if (PrQty == 0 && ClosureStatus != "y") throw new SystemException("若【請購數量】為0，此單據需指定結案!");
                            if (DemandDate.Length <= 0) throw new SystemException("【需求日期】不能為空!");
                            if (PrExchangeRate.Length <= 0) throw new SystemException("【匯率】不能為空!");
                            //if (PrPrice <= 0) throw new SystemException("【請購金額】不能為空!");
                            //if (PrPriceTw <= 0) throw new SystemException("【本幣金額】不能為空!");
                            if (UrgentMtl.Length <= 0) throw new SystemException("【是否急料】不能為空!");
                            if (ProductionPlan.Length <= 0) throw new SystemException("【是否納入生產計畫】不能為空!");
                            if (PrMtlItemName.Length <= 0) throw new SystemException("【品名】不能為空!");

                            #region //檢查資料是否與原請購完全相同
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM SCM.PrDetail a
                                    WHERE a.PrDetailId = @PrDetailId
                                    AND a.MtlItemId = @MtlItemId
                                    AND a.PrMtlItemName = @PrMtlItemName
                                    AND a.PrMtlItemSpec = @PrMtlItemSpec
                                    AND a.InventoryId = @InventoryId
                                    AND a.PrUomId = @PrUomId
                                    AND a.PrQty = @PrQty
                                    AND FORMAT(a.DemandDate, 'yyyy-MM-dd') = @DemandDate
                                    AND a.SupplierId = @SupplierId
                                    AND a.PrCurrency = @PrCurrency
                                    AND a.PrExchangeRate = @PrExchangeRate
                                    AND a.PrUnitPrice = @PrUnitPrice
                                    AND a.PrPrice = @PrPrice
                                    AND a.PrPriceTw = @PrPriceTw
                                    AND a.UrgentMtl = @UrgentMtl
                                    AND a.ProductionPlan = @ProductionPlan
                                    AND a.Project = @Project
                                    AND a.SoDetailId = @SoDetailId
                                    AND a.ClosureStatus = @ClosureStatus
                                    AND a.PrDetailRemark = @PrDetailRemark";
                            dynamicParameters.Add("PrDetailId", PrDetailId);
                            dynamicParameters.Add("MtlItemId", MtlItemId);
                            dynamicParameters.Add("PrMtlItemName", PrMtlItemName);
                            dynamicParameters.Add("PrMtlItemSpec", PrMtlItemSpec);
                            dynamicParameters.Add("InventoryId", InventoryId);
                            dynamicParameters.Add("PrUomId", PrUomId);
                            dynamicParameters.Add("PrQty", PrQty);
                            dynamicParameters.Add("DemandDate", DemandDate);
                            dynamicParameters.Add("SupplierId", SupplierId);
                            dynamicParameters.Add("PrCurrency", PrCurrency);
                            dynamicParameters.Add("PrExchangeRate", PrExchangeRate);
                            dynamicParameters.Add("PrUnitPrice", PrUnitPrice);
                            dynamicParameters.Add("PrPrice", PrPrice);
                            dynamicParameters.Add("PrPriceTw", PrPriceTw);
                            dynamicParameters.Add("UrgentMtl", UrgentMtl);
                            dynamicParameters.Add("ProductionPlan", ProductionPlan);
                            dynamicParameters.Add("Project", Project);
                            dynamicParameters.Add("SoDetailId", SoDetailId);
                            dynamicParameters.Add("ClosureStatus", ClosureStatus);
                            dynamicParameters.Add("PrDetailRemark", PrDetailRemark);

                            var CheckChangeResult = sqlConnection.Query(sql, dynamicParameters);

                            if (CheckChangeResult.Count() > 0) throw new SystemException("請購單內容無變更!");
                            #endregion

                            #region //該請購單是否有開立採購單
                            dynamicParameters = new DynamicParameters();
                            sql = @"select b.PrErpPrefix, b.PrErpNo, a.PrSequence
                                    from SCM.PrDetail a
                                    inner join SCM.PurchaseRequisition b on a.PrId = b.PrId
                                    where a.PrDetailId = @PrDetailId";
                            dynamicParameters.Add("PrDetailId", PrDetailId);

                            var PrDetailMesResult = sqlConnection.Query(sql, dynamicParameters);

                            string PrErpPrefix = "";
                            string PrErpNo = "";
                            string PrSequence = "";

                            foreach (var item in PrDetailMesResult)
                            {
                                PrErpPrefix = item.PrErpPrefix;
                                PrErpNo = item.PrErpNo;
                                PrSequence = item.PrSequence;

                            }

                            dynamicParameters = new DynamicParameters();
                            sql = @"select * 
                                    from PURTD
                                    where TD026 = @PrErpPrefix
                                    and TD027 = @PrErpNo
                                    and TD028 = @PrSequence";
                            dynamicParameters.Add("PrErpPrefix", PrErpPrefix);
                            dynamicParameters.Add("PrErpNo", PrErpNo);
                            dynamicParameters.Add("PrSequence", PrSequence);

                            var PoERPResult = sqlConnection2.Query(sql, dynamicParameters);
                            if (PoERPResult.Count() > 0) throw new SystemException("該請購單已有採購單，無法開立變更單");

                            #endregion

                            #region //確認請購變更單資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.PrId, ISNULL(a.TotalQty, 0) TotalQty, ISNULL(a.Amount, 0) Amount
                                    , a.UserId
                                    FROM SCM.PrModification a
                                    WHERE a.PrmId = @PrmId";
                            dynamicParameters.Add("PrmId", PrmId);

                            var prModificationResult = sqlConnection.Query(sql, dynamicParameters);

                            if (prModificationResult.Count() <= 0) throw new SystemException("【請購變更單】資料錯誤!");

                            int PrId = -1;
                            double TotalQty = -1;
                            double Amount = -1;
                            foreach (var item in prModificationResult)
                            {
                                if (item.UserId != CreateBy) throw new SystemException("無法修改其他請購人之單據!");
                                PrId = item.PrId;
                                TotalQty = item.TotalQty;
                                Amount = item.Amount;
                            }
                            #endregion

                            #region //確認請購單詳細資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.PrSequence OriPrSequence, a.MtlItemId OriMtlItemId, a.PrMtlItemName OriPrMtlItemName
                                    , a.PrMtlItemSpec OriPrMtlItemSpec, a.InventoryId OriInventoryId, a.PrUomId OriPrUomId
                                    , a.PrQty OriPrQty, a.DemandDate OriDemandDate, a.SupplierId OriSupplierId, a.PrCurrency OriPrCurrency
                                    , a.PrExchangeRate OriPrExchangeRate, a.PrUnitPrice OriPrUnitPrice, a.PrPrice OriPrPrice
                                    , a.PrPriceTw OriPrPriceTw, a.UrgentMtl OriUrgentMtl, a.ProductionPlan OriProductionPlan
                                    , a.Project OriProject, ISNULL(a.BudgetDepartmentNo, '') OriBudgetDepartmentNo
                                    , ISNULL(a.BudgetDepartmentSubject, '') OriBudgetDepartmentSubject, a.SoDetailId OriSoDetailId
                                    , ISNULL(FORMAT(a.DeliveryDate, 'yyyy-MM-dd'), null) OriDeliveryDate, a.PoUserId OriPoUserId
                                    , a.PoUomId OriPoUomId, a.PoQty OriPoQty, ISNULL(a.PoCurrency, '') OriPoCurrency
                                    , a.PoUnitPrice OriPoUnitPrice, a.PoPrice OriPoPrice, a.LockStaus OriLockStaus
                                    , ISNULL(a.PoStaus, '') OriPoStaus, ISNULL(a.TaxNo, '') OriTaxNo, a.Taxation OriTaxation
                                    , a.BusinessTaxRate OriBusinessTaxRate, ISNULL(a.DetailMultiTax, '') OriDetailMultiTax
                                    , ISNULL(a.TradeTerm, '') OriTradeTerm, a.PrPriceQty OriPrPriceQty, a.PrPriceUomId OriPrPriceUomId
                                    , a.DiscountRate OriDiscountRate, a.DiscountAmount OriDiscountAmount, a.ClosureStatus OriClosureStatus
                                    , ISNULL(a.PrDetailRemark, '') OriPrDetailRemark, ISNULL(a.PoRemark, '') OriPoRemark
                                    , a.PrMtlItemName, a.PrMtlItemSpec
                                    FROM SCM.PrDetail a
                                    WHERE a.PrDetailId = @PrDetailId";
                            dynamicParameters.Add("PrDetailId", PrDetailId);

                            var PrDetailResult = sqlConnection.Query(sql, dynamicParameters);

                            if (PrDetailResult.Count() <= 0) throw new SystemException("【請購單詳細資料】資料錯誤!");

                            string OriPrSequence = "";
                            int OriMtlItemId = -1;
                            string OriPrMtlItemName = "";
                            string OriPrMtlItemSpec = "";
                            int OriInventoryId = -1;
                            int OriPrUomId = -1;
                            double OriPrQty = -1;
                            DateTime OriDemandDate = new DateTime();
                            int OriSupplierId = -1;
                            string OriPrCurrency = "";
                            double OriPrExchangeRate = -1;
                            double OriPrUnitPrice = -1;
                            double OriPrPrice = -1;
                            double OriPrPriceTw = -1;
                            string OriUrgentMtl = "";
                            string OriProductionPlan = "";
                            string OriProject = "";
                            string OriBudgetDepartmentNo = "";
                            string OriBudgetDepartmentSubject = "";
                            int? OriSoDetailId = -1;
                            DateTime? OriDeliveryDate = new DateTime();
                            int? OriPoUserId = -1;
                            int? OriPoUomId = -1;
                            double? OriPoQty = -1;
                            string OriPoCurrency = "";
                            double? OriPoUnitPrice = -1;
                            double? OriPoPrice = -1;
                            string OriLockStaus = "";
                            string OriPoStaus = "";
                            string OriTaxNo = "";
                            string OriTaxation = "";
                            double? OriBusinessTaxRate = -1;
                            string OriDetailMultiTax = "";
                            string OriTradeTerm = "";
                            int OriPrPriceQty = -1;
                            int OriPrPriceUomId = -1;
                            double? OriDiscountRate = -1;
                            double? OriDiscountAmount = -1;
                            string OriClosureStatus = "";
                            string OriPrDetailRemark = "";
                            string OriPoRemark = "";
                            foreach (var item in PrDetailResult)
                            {
                                OriPrSequence = item.OriPrSequence;
                                OriMtlItemId = item.OriMtlItemId;
                                OriPrMtlItemName = item.OriPrMtlItemName;
                                OriPrMtlItemSpec = item.OriPrMtlItemSpec;
                                OriInventoryId = item.OriInventoryId;
                                OriPrUomId = item.OriPrUomId;
                                OriPrQty = item.OriPrQty;
                                OriDemandDate = item.OriDemandDate;
                                OriSupplierId = item.OriSupplierId;
                                OriPrCurrency = item.OriPrCurrency;
                                OriPrExchangeRate = item.OriPrExchangeRate;
                                OriPrUnitPrice = item.OriPrUnitPrice;
                                OriPrPrice = item.OriPrPrice;
                                OriPrPriceTw = item.OriPrPriceTw;
                                OriUrgentMtl = item.OriUrgentMtl;
                                OriProductionPlan = item.OriProductionPlan;
                                OriProject = item.OriProject;
                                OriBudgetDepartmentNo = item.OriBudgetDepartmentNo;
                                OriBudgetDepartmentSubject = item.OriBudgetDepartmentSubject;
                                OriSoDetailId = item.OriSoDetailId;
                                OriDeliveryDate = item.OriDeliveryDate;
                                OriPoUserId = item.OriPoUserId;
                                OriPoUomId = item.OriPoUomId;
                                OriPoQty = item.OriPoQty;
                                OriPoCurrency = item.OriPoCurrency;
                                OriPoUnitPrice = item.OriPoUnitPrice;
                                OriPoPrice = item.OriPoPrice;
                                OriLockStaus = item.OriLockStaus;
                                OriPoStaus = item.OriPoStaus;
                                OriTaxNo = item.OriTaxNo;
                                OriTaxation = item.OriTaxation;
                                OriBusinessTaxRate = item.OriBusinessTaxRate;
                                OriDetailMultiTax = item.OriDetailMultiTax;
                                OriTradeTerm = item.OriTradeTerm;
                                OriPrPriceQty = item.OriPrPriceQty;
                                OriPrPriceUomId = item.OriPrPriceUomId;
                                OriDiscountRate = item.OriDiscountRate;
                                OriDiscountAmount = item.OriDiscountAmount;
                                OriClosureStatus = item.OriClosureStatus;
                                OriPrDetailRemark = item.OriPrDetailRemark;
                                OriPoRemark = item.OriPoRemark;
                            }
                            #endregion

                            #region //確認請購變更單序號是否重複
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM SCM.PrmDetail a
                                    WHERE a.PrmId = @PrmId
                                    AND a.PrmSequence = @PrmSequence";
                            dynamicParameters.Add("PrmId", PrmId);
                            dynamicParameters.Add("PrmSequence", PrmSequence);

                            var result = sqlConnection.Query(sql, dynamicParameters);

                            if (result.Count() > 0) throw new SystemException("【請購變更單序號】重複!");
                            #endregion

                            #region //檢查同張請購變更單中是否已存在相同請購單身變更單
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM SCM.PrmDetail a
                                    WHERE a.PrmId = @PrmId
                                    AND a.PrDetailId = @PrDetailId";
                            dynamicParameters.Add("PrmId", PrmId);
                            dynamicParameters.Add("PrDetailId", PrDetailId);

                            var PrmDetailResult = sqlConnection.Query(sql, dynamicParameters);

                            if (PrmDetailResult.Count() > 0) throw new SystemException("同樣單據中已存在相同請購變更單單身!!");
                            #endregion

                            #region //判斷品號資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.MtlItemNo, a.MtlItemName, a.MtlItemSpec
                                    FROM PDM.MtlItem a
                                    WHERE a.MtlItemId = @MtlItemId";
                            dynamicParameters.Add("MtlItemId", MtlItemId);

                            var result2 = sqlConnection.Query(sql, dynamicParameters);

                            if (result2.Count() <= 0) throw new SystemException("【品號】資料錯誤!");

                            string MtlItemNo = "";
                            string MtlItemName = "";
                            string MtlItemSpec = "";
                            foreach (var item in result2)
                            {
                                MtlItemNo = item.MtlItemNo;
                                MtlItemName = item.MtlItemName;
                                MtlItemSpec = item.MtlItemSpec;
                            }
                            #endregion

                            #region //判斷ERP品號生效日與失效日
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MB030)) MB030, LTRIM(RTRIM(MB031)) MB031
                                    FROM INVMB
                                    WHERE MB001 = @MB001";
                            dynamicParameters.Add("MB001", MtlItemNo);

                            var INVMBResult = sqlConnection2.Query(sql, dynamicParameters);

                            foreach (var item in INVMBResult)
                            {
                                if (item.MB030 != "" && item.MB030 != null)
                                {
                                    #region //判斷日期需大於或等於生效日
                                    string EffectiveDate = item.MB030;
                                    string effYear = EffectiveDate.Substring(0, 4);
                                    string effMonth = EffectiveDate.Substring(4, 2);
                                    string effDay = EffectiveDate.Substring(6, 2);
                                    DateTime effFullDate = Convert.ToDateTime(effYear + "-" + effMonth + "-" + effDay);
                                    int effresult = DateTime.Compare(CreateDate, effFullDate);
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
                                    int effresult = DateTime.Compare(CreateDate, effFullDate);
                                    if (effresult > 0) throw new SystemException("此品號已失效，無法使用!!");
                                    #endregion
                                }
                            }
                            #endregion

                            #region //判斷庫別資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.InventoryNo, a.InventoryName
                                    FROM SCM.Inventory a
                                    WHERE a.InventoryId = @InventoryId";
                            dynamicParameters.Add("InventoryId", InventoryId);

                            var result3 = sqlConnection.Query(sql, dynamicParameters);

                            if (result3.Count() <= 0) throw new SystemException("【庫別】資料錯誤!");

                            string InventoryNo = "";
                            string InventoryName = "";
                            foreach (var item in result3)
                            {
                                InventoryNo = item.InventoryNo;
                                InventoryName = item.InventoryName;
                            }
                            #endregion

                            #region //判斷單位資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM PDM.UnitOfMeasure a
                                    WHERE a.UomId = @UomId";
                            dynamicParameters.Add("UomId", PrUomId);

                            var result4 = sqlConnection.Query(sql, dynamicParameters);

                            if (result4.Count() <= 0) throw new SystemException("【單位】資料錯誤!");
                            #endregion

                            #region //取得ERP庫存資訊
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(a.MC007)) InventoryQty
                                    FROM INVMC a
                                    WHERE a.MC001 = @MC001
                                    AND a.MC002 = @MC002";
                            dynamicParameters.Add("MC001", MtlItemNo);
                            dynamicParameters.Add("MC002", InventoryNo);

                            var result5 = sqlConnection2.Query(sql, dynamicParameters);

                            double InventoryQty = 0;
                            string mtlInventory = "目前尚無資料";
                            if (result5.Count() > 0)
                            {
                                foreach (var item in result5)
                                {
                                    InventoryQty = Convert.ToDouble(item.InventoryQty);
                                    #region //組MtlInventory
                                    List<MtlInventory> mtlInventories = new List<MtlInventory>
                                    {
                                        new MtlInventory
                                        {
                                            WAREHOUSE_NO = InventoryNo,
                                            WAREHOUSE_NAME = InventoryName,
                                            WAREHOUSE_QTY = InventoryQty
                                        }
                                    };
                                    #endregion

                                    mtlInventory = JsonConvert.SerializeObject(mtlInventories);
                                    mtlInventory = "{\"data\":" + mtlInventory + "}";
                                }
                            }
                            #endregion

                            #region //檢查供應商資料是否正確
                            string TaxNo = "";
                            string Taxation = "";
                            string TradeTerm = "";
                            double? BusinessTaxRate = -1;
                            string HideSupplier = "";
                            if (SupplierId > 0)
                            {
                                #region //供應商
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.TaxNo, a.Taxation, a.TradeTerm, a.HideSupplier
                                        FROM SCM.Supplier a
                                        WHERE a.SupplierId = @SupplierId";
                                dynamicParameters.Add("SupplierId", SupplierId);

                                var result6 = sqlConnection.Query(sql, dynamicParameters);

                                if (result6.Count() <= 0) throw new SystemException("【供應商】資料錯誤!");

                                foreach (var item in result6)
                                {
                                    TaxNo = item.TaxNo;
                                    Taxation = item.Taxation;
                                    TradeTerm = item.TradeTerm;
                                    HideSupplier = item.HideSupplier;
                                }
                                #endregion

                                #region //查詢營業稅額資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT ROUND(LTRIM(RTRIM(NN004)),2) BusinessTaxRate
                                        FROM CMSNN 
                                        WHERE NN001 = @TaxNo";
                                dynamicParameters.Add("TaxNo", TaxNo);

                                var businessTaxRateResult = sqlConnection2.Query(sql, dynamicParameters);

                                if (businessTaxRateResult.Count() <= 0) throw new SystemException("【課稅別】資料錯誤!");

                                foreach (var item in businessTaxRateResult)
                                {
                                    BusinessTaxRate = item.BusinessTaxRate;
                                }
                                #endregion
                            }
                            #endregion

                            #region //確認集團內/集團外邏輯
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 b.HideSupplier
                                    FROM SCM.PrDetail a 
                                    INNER JOIN SCM.Supplier b ON a.SupplierId = b.SupplierId
                                    WHERE a.PrId = @PrId
                                    ORDER BY a.PrSequence";
                            dynamicParameters.Add("PrId", PrId);

                            var PrDetailResult2 = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item in PrDetailResult2)
                            {
                                string HideSupplierString = "";
                                if (item.HideSupplier == "Y")
                                {
                                    HideSupplierString = "集團內";
                                }
                                else
                                {
                                    HideSupplierString = "集團外";
                                }
                                if (item.HideSupplier != HideSupplier) throw new SystemException("此請購單第一筆單身為【" + HideSupplierString + "】，無法新增!!");
                            }
                            #endregion

                            #region //判斷幣別資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM CMSMF
                                    WHERE MF001 = @MF001";
                            dynamicParameters.Add("MF001", PrCurrency);

                            var result7 = sqlConnection2.Query(sql, dynamicParameters);
                            if (result7.Count() <= 0) throw new SystemException("【幣別】資料有誤!");
                            #endregion

                            #region //判斷專案代碼資料是否正確
                            if (Project.Length > 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                    FROM CMSNB
                                    WHERE NB001 = @NB001";
                                dynamicParameters.Add("NB001", Project);

                                var result8 = sqlConnection2.Query(sql, dynamicParameters);
                                if (result8.Count() <= 0) throw new SystemException("【專案代碼】資料有誤!");
                            }
                            #endregion

                            #region //確認訂單資料是否正確
                            if (SoDetailId > 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT MtlItemId
                                        FROM SCM.SoDetail
                                        WHERE SoDetailId = @SoDetailId";
                                dynamicParameters.Add("SoDetailId", SoDetailId);

                                var result9 = sqlConnection.Query(sql, dynamicParameters);
                                if (result9.Count() <= 0) throw new SystemException("【訂單】資料有誤!");

                                foreach (var item in result9)
                                {
                                    //if (item.MtlItemId != MtlItemId) throw new SystemException("訂單品號與請購品號不同!!");
                                }
                            }
                            #endregion

                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO SCM.PrmDetail (PrmId, PrDetailId, PrmSequence, MtlItemId, PrMtlItemName, PrMtlItemSpec, InventoryId, PrUomId 
                                    , PrQty, DemandDate, SupplierId, PrCurrency, PrExchangeRate, PrUnitPrice, PrPrice, PrPriceTw, UrgentMtl
                                    , PoUomId, PoQty, PoCurrency, PoUnitPrice, PoPrice
                                    , ProductionPlan, Project, SoDetailId, DeliveryDate, LockStaus, PoStaus, TaxNo, Taxation, BusinessTaxRate
                                    , DetailMultiTax, TradeTerm, PrPriceQty, PrPriceUomId, DiscountRate, DiscountAmount, MtlInventory, MtlInventoryQty
                                    , ConfirmStatus, ConfirmUserId, ClosureStatus, PrDetailRemark, PoRemark, ModiReason, OriPrSequence, OriMtlItemId
                                    , OriPrMtlItemName, OriPrMtlItemSpec, OriInventoryId, OriPrUomId, OriPrQty, OriDemandDate, OriSupplierId
                                    , OriPrCurrency, OriPrExchangeRate, OriPrUnitPrice, OriPrPrice, OriPrPriceTw, OriUrgentMtl, OriProductionPlan
                                    , OriProject, OriBudgetDepartmentNo, OriBudgetDepartmentSubject, OriSoDetailId, OriDeliveryDate, OriPoUserId
                                    , OriPoUomId, OriPoQty, OriPoCurrency, OriPoUnitPrice, OriPoPrice, OriLockStaus, OriPoStaus, OriTaxNo, OriTaxation
                                    , OriBusinessTaxRate, OriDetailMultiTax, OriTradeTerm, OriPrPriceQty, OriPrPriceUomId, OriDiscountRate
                                    , OriDiscountAmount, OriClosureStatus, OriPrDetailRemark, OriPoRemark
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.PrmDetailId
                                    VALUES (@PrmId, @PrDetailId, @PrmSequence, @MtlItemId, @PrMtlItemName, @PrMtlItemSpec, @InventoryId, @PrUomId 
                                    , @PrQty, @DemandDate, @SupplierId, @PrCurrency, @PrExchangeRate, @PrUnitPrice, @PrPrice, @PrPriceTw, @UrgentMtl
                                    , @PoUomId, @PoQty, @PoCurrency, @PoUnitPrice, @PoPrice
                                    , @ProductionPlan, @Project, @SoDetailId, @DeliveryDate, @LockStaus, @PoStaus, @TaxNo, @Taxation, @BusinessTaxRate
                                    , @DetailMultiTax, @TradeTerm, @PrPriceQty, @PrPriceUomId, @DiscountRate, @DiscountAmount, @MtlInventory, @MtlInventoryQty
                                    , @ConfirmStatus, @ConfirmUserId, @ClosureStatus, @PrDetailRemark, @PoRemark, @ModiReason, @OriPrSequence, @OriMtlItemId
                                    , @OriPrMtlItemName, @OriPrMtlItemSpec, @OriInventoryId, @OriPrUomId, @OriPrQty, @OriDemandDate, @OriSupplierId
                                    , @OriPrCurrency, @OriPrExchangeRate, @OriPrUnitPrice, @OriPrPrice, @OriPrPriceTw, @OriUrgentMtl, @OriProductionPlan
                                    , @OriProject, @OriBudgetDepartmentNo, @OriBudgetDepartmentSubject, @OriSoDetailId, @OriDeliveryDate, @OriPoUserId
                                    , @OriPoUomId, @OriPoQty, @OriPoCurrency, @OriPoUnitPrice, @OriPoPrice, @OriLockStaus, @OriPoStaus, @OriTaxNo, @OriTaxation
                                    , @OriBusinessTaxRate, @OriDetailMultiTax, @OriTradeTerm, @OriPrPriceQty, @OriPrPriceUomId, @OriDiscountRate
                                    , @OriDiscountAmount, @OriClosureStatus, @OriPrDetailRemark, @OriPoRemark
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    PrmId,
                                    PrDetailId,
                                    PrmSequence,
                                    MtlItemId,
                                    PrMtlItemName,
                                    PrMtlItemSpec,
                                    InventoryId,
                                    PrUomId,
                                    PrQty,
                                    DemandDate,
                                    SupplierId,
                                    PrCurrency,
                                    PrExchangeRate,
                                    PrUnitPrice,
                                    PrPrice,
                                    PrPriceTw,
                                    UrgentMtl,
                                    PoUomId = PrUomId,
                                    PoQty = PrQty,
                                    PoCurrency = PrCurrency,
                                    PoUnitPrice = PrUnitPrice,
                                    PoPrice = PrPrice,
                                    ProductionPlan,
                                    Project,
                                    SoDetailId,
                                    DeliveryDate = (DateTime?)null,
                                    LockStaus = "N",
                                    PoStaus = "N",
                                    TaxNo = TaxNo != "" ? TaxNo : null,
                                    Taxation = Taxation != "" ? Taxation : null,
                                    BusinessTaxRate = BusinessTaxRate > 0 ? BusinessTaxRate : (double?)null,
                                    DetailMultiTax = "N",
                                    TradeTerm = TradeTerm != "" ? TradeTerm : null,
                                    PrPriceQty = PrQty,
                                    PrPriceUomId = PrUomId,
                                    DiscountRate = 0,
                                    DiscountAmount = 0,
                                    MtlInventory = mtlInventory,
                                    MtlInventoryQty = InventoryQty,
                                    ConfirmStatus = "N",
                                    ConfirmUserId = (int?)null,
                                    ClosureStatus,
                                    PrDetailRemark,
                                    PoRemark,
                                    ModiReason,
                                    OriPrSequence,
                                    OriMtlItemId,
                                    OriPrMtlItemName,
                                    OriPrMtlItemSpec,
                                    OriInventoryId,
                                    OriPrUomId,
                                    OriPrQty,
                                    OriDemandDate,
                                    OriSupplierId,
                                    OriPrCurrency,
                                    OriPrExchangeRate,
                                    OriPrUnitPrice,
                                    OriPrPrice,
                                    OriPrPriceTw,
                                    OriUrgentMtl,
                                    OriProductionPlan,
                                    OriProject,
                                    OriBudgetDepartmentNo,
                                    OriBudgetDepartmentSubject,
                                    OriSoDetailId,
                                    OriDeliveryDate,
                                    OriPoUserId,
                                    OriPoUomId,
                                    OriPoQty,
                                    OriPoCurrency,
                                    OriPoUnitPrice,
                                    OriPoPrice,
                                    OriLockStaus,
                                    OriPoStaus,
                                    OriTaxNo,
                                    OriTaxation,
                                    OriBusinessTaxRate,
                                    OriDetailMultiTax,
                                    OriTradeTerm,
                                    OriPrPriceQty,
                                    OriPrPriceUomId,
                                    OriDiscountRate,
                                    OriDiscountAmount,
                                    OriClosureStatus,
                                    OriPrDetailRemark,
                                    OriPoRemark,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            int rowsAffected = insertResult.Count();

                            #region //更新單頭總採購數量及金額
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.PrModification SET
                                    TotalQty = TotalQty + @PrQty,
                                    Amount = Amount + @PrPriceTw,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE PrmId = @PrmId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    PrQty,
                                    PrPriceTw,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    PrmId
                                });

                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
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

        #region //AddPrLog -- 新增請購單LOG紀錄 -- Ann 2023-12-06
        public string AddPrLog(int PrId, string BpmNo, string Status, string RootId, string ConfirmUser)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認請購單資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.BpmTransferDate
                                FROM SCM.PurchaseRequisition a
                                WHERE PrId = @PrId";
                        dynamicParameters.Add("PrId", PrId);

                        var PurchaseRequisitionResult = sqlConnection.Query(sql, dynamicParameters);

                        if (PurchaseRequisitionResult.Count() <= 0) throw new SystemException("請購單資料錯誤!!");

                        DateTime BpmTransferDate = new DateTime();
                        foreach (var item in PurchaseRequisitionResult)
                        {
                            BpmTransferDate = item.BpmTransferDate;
                        }
                        #endregion

                        #region //INSERT SCM.PrLog
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.PrLog (PrId, RootId, BpmNo, TransferBpmDate, BpmStatus, ConfirmUser
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.PrLogId
                                VALUES (@PrId, @RootId, @BpmNo, @TransferBpmDate, @BpmStatus, @ConfirmUser
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                          new
                          {
                              PrId,
                              RootId,
                              BpmNo,
                              TransferBpmDate = BpmTransferDate,
                              BpmStatus = Status,
                              ConfirmUser,
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

        #region //AddPrFile -- 新增請購單附檔 -- Ann 2024-08-06
        public string AddPrFile(int PrId, string PrFile)
        {
            try
            {
                if (PrFile.Length < 0) throw new SystemException("【請購附檔】為空，無法新增!!");
                int rowsAffected = 0;
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //確認請購人員資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.PrStatus
                                FROM SCM.PurchaseRequisition a 
                                WHERE a.PrId = @PrId";
                        dynamicParameters.Add("PrId", PrId);

                        var PurchaseRequisitionResult = sqlConnection.Query(sql, dynamicParameters);
                        if (PurchaseRequisitionResult.Count() <= 0) throw new SystemException("【請購單】資料錯誤，請重新輸入!");

                        foreach (var item in PurchaseRequisitionResult)
                        {
                            if (item.PrStatus != "N") throw new SystemException("【請購單】狀態不可重複，請重新輸入!");
                        }
                        #endregion

                        #region //先刪除原先檔案
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE FROM SCM.PrFile
                                WHERE PrId = @PrId";
                        dynamicParameters.Add("PrId", PrId);

                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        string[] prFiles = PrFile.Split(',');
                        foreach (var file in prFiles)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO SCM.PrFile (PrId, FileId
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.PrFileId
                                    VALUES (@PrId, @FileId
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    PrId,
                                    FileId = file,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();

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

        #region //ApiAddPrData -- Api自動開立請購單 -- Yi 2023.09.27
        public string ApiAddPrData(string CompanyNo, string PrInfoJson)
        {
            try
            {
                if (!PrInfoJson.TryParseJson(out JObject tempJObject)) throw new SystemException("資料格式錯誤");

                PrInfo prInfos = JsonConvert.DeserializeObject<PrInfo>(PrInfoJson);

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    int rowsAffected = 0, companyId = -1, userId = -1, departmentId = -1, prId = -1, umoId = -1, unitDecimal = 0, totalDecimal = 0;
                    string userNo = "", userName = "", departmentNo = "", departmentName = "", prErpNo = "", mtlItemNo = "", inventoryNo = "", inventoryName = ""
                        , umoNo = "", taxNo = "", taxation = "", tradeTerm = "", currency = "", supplierName = "", mtlItemName = "", mtlItemSpec = ""
                        , mtlInventory = "目前尚無資料", companyNo = "";

                    double prExchangeRate = 0, inventoryQty = 0;
                    double? businessTaxRate = -1;

                    //查詢MES相關資料Amount
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷公司資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 CompanyId, CompanyName, ErpDb, CompanyNo, ErpNo
                                FROM BAS.Company
                                WHERE CompanyNo = @CompanyNo";
                        dynamicParameters.Add("CompanyNo", CompanyNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【公司】資料錯誤!");

                        foreach (var item in result)
                        {
                            companyId = Convert.ToInt32(item.CompanyId);
                            companyNo = item.CompanyNo;
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        }
                        #endregion

                        #region //確認請購人員資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 UserId
                                FROM BAS.[User]
                                WHERE UserNo = @UserNo";
                        dynamicParameters.Add("UserNo", prInfos.UserNo);

                        var resultUser = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUser.Count() <= 0) throw new SystemException("【使用者】資料錯誤!");

                        foreach (var item in resultUser)
                        {
                            userId = Convert.ToInt32(item.UserId);
                        }
                        #endregion

                        #region //判斷使用者資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 d.UserNo, d.UserName, d.DepartmentId, d.DepartmentNo, d.DepartmentName
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
                                    SELECT da.UserNo, da.UserName, db.DepartmentId, db.DepartmentNo, db.DepartmentName
                                    FROM BAS.[User] da
                                    INNER JOIN BAS.Department db ON da.DepartmentId = db.DepartmentId
                                    WHERE da.UserId = @UserId
                                    AND da.UserStatus = 'F'
                                ) d
                                WHERE a.[Status] = @Status
                                AND b.[Status] = @Status
                                AND b.FunctionCode = @FunctionCode
                                AND a.DetailCode = @DetailCode
                                AND c.Authority > 0";
                        dynamicParameters.Add("UserId", userId);
                        dynamicParameters.Add("CompanyId", companyId);
                        dynamicParameters.Add("Status", "A");
                        dynamicParameters.Add("FunctionCode", "PurchaseRequisition");
                        dynamicParameters.Add("DetailCode", "add");

                        var resultUserInfo = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUserInfo.Count() <= 0) throw new SystemException("使用者資料錯誤!");

                        foreach (var item in resultUserInfo)
                        {
                            userNo = item.UserNo;
                            userName = item.UserName;
                            departmentNo = item.DepartmentNo;
                            departmentName = item.DepartmentName;
                        }
                        #endregion

                        #region //確認此PrNo未重複
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.PurchaseRequisition a
                                WHERE a.PrNo = @PrNo";
                        dynamicParameters.Add("PrNo", prInfos.PrNo);

                        var prNoResult = sqlConnection.Query(sql, dynamicParameters);
                        if (prNoResult.Count() > 0) throw new SystemException("【請購單編號】重複，請重新輸入!");
                        #endregion

                        #region //隨機取得單號資料
                        bool checkPrErpNo = true;
                        while (checkPrErpNo)
                        {
                            prErpNo = BaseHelper.RandomCode(11);

                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM SCM.PurchaseRequisition
                                    WHERE PrErpPrefix = @PrErpPrefix
                                    AND PrErpNo = @PrErpNo";
                            dynamicParameters.Add("PrErpPrefix", prInfos.PrErpPrefix);
                            dynamicParameters.Add("PrErpNo", prErpNo);

                            var resultSoErpNo = sqlConnection.Query(sql, dynamicParameters);
                            checkPrErpNo = resultSoErpNo.Count() > 0;
                        }
                        #endregion

                        #region //確認請購部門資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 a.DepartmentId
                                FROM BAS.Department a
                                WHERE a.DepartmentNo = @DepartmentNo";
                        dynamicParameters.Add("DepartmentNo", prInfos.DepartmentNo);

                        var DepartmentResult = sqlConnection.Query(sql, dynamicParameters);
                        if (DepartmentResult.Count() <= 0) throw new SystemException("【請購部門】資料錯誤，請重新輸入!");

                        foreach (var item in DepartmentResult)
                        {
                            departmentId = item.DepartmentId;
                        }
                        #endregion

                        #region //確認品號資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MtlItemId, a.MtlItemNo, a.MtlItemName, a.MtlItemSpec, a.SaleUomId
                                FROM PDM.MtlItem a
                                WHERE a.MtlItemNo IN @MtlItemNo";
                        dynamicParameters.Add("MtlItemNo", prInfos.Details.Select(x => x.MtlItemNo).ToArray());

                        List<MtlItem> mtlItems = sqlConnection.Query<MtlItem>(sql, dynamicParameters).ToList();
                        if (mtlItems.Count() != prInfos.Details.Count()) throw new SystemException("【品號】資料錯誤，請重新輸入!");

                        //判斷兩者找出品號筆數是否一致
                        prInfos.Details = prInfos.Details.Join(mtlItems, x => x.MtlItemNo, y => y.MtlItemNo, (x, y) => { x.MtlItemId = (int)y.MtlItemId; return x; }).Select(x => x).ToList();

                        foreach (var item in mtlItems)
                        {
                            mtlItemNo = item.MtlItemNo;
                            mtlItemName = item.MtlItemName;
                            mtlItemSpec = item.MtlItemSpec;
                            umoId = item.SaleUomId;
                        }
                        #endregion

                        #region //判斷庫別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.InventoryId, a.InventoryNo, a.InventoryName, a.CompanyId
                                FROM SCM.Inventory a
                                WHERE a.InventoryNo IN @InventoryNo
                                AND a.CompanyId = @CompanyId";
                        dynamicParameters.Add("InventoryNo", prInfos.Details.Select(x => x.InventoryNo).ToArray());
                        dynamicParameters.Add("CompanyId", companyId);

                        List<Inventory> inventorys = sqlConnection.Query<Inventory>(sql, dynamicParameters).ToList();
                        if (inventorys.Count() != prInfos.Details.Select(x => x.InventoryNo).Distinct().Count()) throw new SystemException("【庫別】資料錯誤，請重新輸入!");

                        //判斷兩者找出品號筆數是否一致
                        prInfos.Details = prInfos.Details.Join(inventorys, x => x.InventoryNo, y => y.InventoryNo, (x, y) => { x.InventoryId = (int)y.InventoryId; return x; }).Select(x => x).ToList();

                        foreach (var item in inventorys)
                        {
                            inventoryNo = item.InventoryNo;
                            inventoryName = item.InventoryName;
                        }
                        #endregion

                        #region //判斷單位資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 b.MtlItemNo, b.MtlItemName, a.UomNo
                                FROM PDM.UnitOfMeasure a
                                INNER JOIN PDM.MtlItem b ON a.UomId = b.SaleUomId
                                WHERE a.UomId = @UomId";
                        dynamicParameters.Add("UomId", umoId);

                        var umoResult = sqlConnection.Query(sql, dynamicParameters);
                        if (umoResult.Count() <= 0) throw new SystemException("【單位】資料錯誤!");

                        foreach (var item in umoResult)
                        {
                            umoNo = item.UomNo;
                        }
                        #endregion

                        #region //判斷供應商資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.SupplierId, a.SupplierNo, a.TaxNo, a.Taxation, a.TradeTerm, a.SupplierName, a.Currency
                                FROM SCM.Supplier a
                                WHERE a.SupplierNo IN @SupplierNo";
                        dynamicParameters.Add("SupplierNo", prInfos.Details.Select(x => x.SupplierNo).ToArray());

                        List<Supplier> suppliers = sqlConnection.Query<Supplier>(sql, dynamicParameters).ToList();
                        if (suppliers.Count() != prInfos.Details.Select(x => x.SupplierNo).Distinct().Count()) throw new SystemException("【供應商】資料錯誤!");

                        prInfos.Details = prInfos.Details.Join(suppliers, x => x.SupplierNo, y => y.SupplierNo, (x, y) => { x.SupplierId = (int)y.SupplierId; return x; }).Select(x => x).ToList();

                        foreach (var item in suppliers)
                        {
                            taxNo = item.TaxNo;
                            taxation = item.Taxation;
                            tradeTerm = item.TradeTerm;
                            currency = item.Currency;
                            supplierName = item.SupplierName;
                        }
                        #endregion
                    }

                    //查詢ERP相關資料
                    using (SqlConnection sqlConnection = new SqlConnection(ErpConnectionStrings))
                    {
                        #region //判斷ERP品號生效日與失效日
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT LTRIM(RTRIM(MB030)) MB030, LTRIM(RTRIM(MB031)) MB031
                                FROM INVMB
                                WHERE MB001 = @MtlItemNo";
                        dynamicParameters.Add("MtlItemNo", mtlItemNo);

                        var invmbResults = sqlConnection.Query(sql, dynamicParameters);
                        if (invmbResults.Count() <= 0) throw new SystemException("此品號不存在於ERP中!");

                        foreach (var item in invmbResults)
                        {
                            if (item.MB030 != "" && item.MB030 != null)
                            {
                                #region //判斷日期需大於或等於生效日
                                string EffectiveDate = item.MB030;
                                string effYear = EffectiveDate.Substring(0, 4);
                                string effMonth = EffectiveDate.Substring(4, 2);
                                string effDay = EffectiveDate.Substring(6, 2);
                                DateTime effFullDate = Convert.ToDateTime(effYear + "-" + effMonth + "-" + effDay);
                                int effresult = DateTime.Compare(CreateDate, effFullDate);
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
                                int effresult = DateTime.Compare(CreateDate, effFullDate);
                                if (effresult > 0) throw new SystemException("此品號已失效，無法使用!!");
                                #endregion
                            }
                        }
                        #endregion

                        #region //取得ERP庫存資訊
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT LTRIM(RTRIM(a.MC007)) InventoryQty
                                FROM INVMC a
                                WHERE a.MC001 = @MtlItemNo
                                AND a.MC002 = @InventoryNo";
                        dynamicParameters.Add("MtlItemNo", mtlItemNo);
                        dynamicParameters.Add("InventoryNo", inventoryNo);

                        var invmcResults = sqlConnection.Query(sql, dynamicParameters);

                        
                        if (invmcResults.Count() > 0)
                        {
                            foreach (var item in invmcResults)
                            {
                                inventoryQty = Convert.ToDouble(item.InventoryQty);

                                #region //組MtlInventory
                                List<MtlInventory> mtlInventories = new List<MtlInventory>
                                    {
                                        new MtlInventory
                                        {
                                            WAREHOUSE_NO = inventoryNo,
                                            WAREHOUSE_NAME = inventoryName,
                                            WAREHOUSE_QTY = inventoryQty
                                        }
                                    };
                                #endregion

                                mtlInventory = JsonConvert.SerializeObject(mtlInventories);
                                mtlInventory = "{\"data\":" + mtlInventory + "}";
                            }
                        }
                        #endregion

                        #region //查詢營業稅額資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ROUND(LTRIM(RTRIM(NN004)),2) BusinessTaxRate
                                FROM CMSNN 
                                WHERE NN001 = @TaxNo";
                        dynamicParameters.Add("TaxNo", taxNo);

                        var cmsnnResults = sqlConnection.Query(sql, dynamicParameters);
                        if (cmsnnResults.Count() <= 0) throw new SystemException("【課稅別】資料錯誤!");

                        foreach (var item in cmsnnResults)
                        {
                            businessTaxRate = item.BusinessTaxRate;
                        }
                        #endregion

                        #region //查詢ERP是否有此帳號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT LTRIM(RTRIM(a.MF004)) USR_GROUP
                                FROM ADMMF a
                                WHERE MF001 = @UserNo";
                        dynamicParameters.Add("UserNo", userNo);

                        var admmfResults = sqlConnection.Query(sql, dynamicParameters);
                        if (admmfResults.Count() <= 0) throw new SystemException("查無ERP帳號權限，無法進行拋轉及核單!");
                        #endregion

                        #region //判斷幣別及幣別小數點進位資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 LTRIM(RTRIM(MF001)) CurrencyNo, LTRIM(RTRIM(MF002)) CurrencyName
                                , LTRIM(RTRIM(MF003)) UnitDecimal, LTRIM(RTRIM(MF004)) TotalDecimal
                                FROM CMSMF
                                WHERE MF001 = @PrCurrency";
                        dynamicParameters.Add("PrCurrency", currency);

                        var cmsmfUnitResults = sqlConnection.Query(sql, dynamicParameters);
                        if (cmsmfUnitResults.Count() <= 0) throw new SystemException("【幣別或幣別小數點進位】資料有誤!");

                        foreach (var item in cmsmfUnitResults)
                        {
                            unitDecimal = Convert.ToInt32(item.UnitDecimal);
                            totalDecimal = Convert.ToInt32(item.TotalDecimal);
                        }
                        #endregion

                        #region //判斷匯率資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 LTRIM(RTRIM(MG001)) Currency, ROUND(LTRIM(RTRIM(MG004)),3) ExchangeRateNameForMwe
                                , ROUND(LTRIM(RTRIM(MG005)),3) ExchangeRateName, LTRIM(RTRIM(MG002)) StartDate
                                FROM CMSMG
                                WHERE MG001 = @PrCurrency";
                        dynamicParameters.Add("PrCurrency", currency);

                        var cmsmgResults = sqlConnection.Query(sql, dynamicParameters);
                        if (cmsmgResults.Count() <= 0) throw new SystemException("【匯率】資料有誤!");

                        foreach (var item in cmsmgResults)
                        {
                            prExchangeRate = item.ExchangeRateName;
                        }
                        #endregion
                    }

                    //寫入MES相關資料
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //單號自動取號
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT CONVERT(int, RIGHT(ISNULL(MAX(PrNo), '000'), 3)) + 1 CurrentNum
                                FROM SCM.PurchaseRequisition
                                WHERE PrNo LIKE @PrNo";
                        dynamicParameters.Add("PrNo", string.Format("{0}{1}___", "PR", DateTime.Now.ToString("yyyyMMdd")));
                        int currentNum = sqlConnection.QueryFirstOrDefault(sql, dynamicParameters).CurrentNum;

                        string prNo = string.Format("{0}{1}{2}", "PR", DateTime.Now.ToString("yyyyMMdd"), string.Format("{0:000}", currentNum));
                        #endregion

                        #region //取得請購單身資料
                        prInfos.Details
                            .ToList()
                            .ForEach(x =>
                            {
                                x.PrUomId = umoId;                                                                                  //單位(對應PDM.MtlItem.SaleUomId欄位)
                                //x.PrUnitPrice = Math.Round(x.PrUnitPrice, unitDecimal);                                           //TB049 請購單位金額
                                x.PrPrice = Math.Round(x.PrQty * Math.Round(x.PrUnitPrice, unitDecimal), totalDecimal);             //TB051 請購金額
                                x.PrPriceTw = Math.Round(x.PrPrice * prExchangeRate);   //TB045 請購本幣金額
                                x.PrExchangeRate = prExchangeRate;          //TB044 請購匯率
                                x.PoQty = x.PrQty;                          //TB014 採購數量
                                x.PoUomId = umoId;                          //TB015 採購單位
                                x.PoCurrency = currency;                    //TB016 採購幣別
                                x.PoUnitPrice = x.PrUnitPrice;              //TB017 採購單位金額
                                x.PoPrice = x.PrPrice;                      //TB018 採購金額
                                x.PrPriceQty = x.PrQty;                     //TB065 計價數量
                                x.PrPriceUomId = umoId;                     //TB066 計價單位
                                x.SoDetailId = -1;
                                x.LockStaus = "N";                          //TB020 鎖定碼
                                x.PoStaus = "N";                            //TB021 採購碼
                                x.PartialPurchaseStaus = "N";               //TB046 分批採購
                                x.InquiryStatus = "1";                      //TB040 詢價碼
                                x.TaxNo = taxNo;                            //TB057 稅別碼
                                x.Taxation = taxation;                      //TB026 課稅別
                                x.BusinessTaxRate = businessTaxRate;        //TB063 營業稅率
                                x.DetailMultiTax = "N";                     //TB064 單身多稅率
                                x.TradeTerm = tradeTerm;                    //TB058 交易條件
                                x.MtlInventory = mtlInventory;              //品號庫存狀況
                                x.MtlInventoryQty = inventoryQty;           //品號庫存總數量
                                x.ConfirmStatus = "N";                      //TB025 確認碼
                                x.ClosureStatus = "N";                      //TB039 結案碼
                                x.CreateDate = CreateDate;
                                x.LastModifiedDate = LastModifiedDate;
                                x.CreateBy = userId;
                                x.LastModifiedBy = userId;
                            });
                        #endregion

                        #region //請購單頭新增(同時依據單身資料更新TotalQty及Amount)
                        dynamicParameters = new DynamicParameters();
                        sql = @"INSERT INTO SCM.PurchaseRequisition (PrNo, CompanyId, DepartmentId, UserId, PrErpPrefix, PrErpNo, Edition, PrDate, DocDate, PrRemark
                                , TotalQty, Amount, PrStatus, SignupStaus, LockStaus, ConfirmStatus, BpmTransferStatus, TransferStatus, Priority, Source
                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                OUTPUT INSERTED.PrId, INSERTED.PrErpPrefix, INSERTED.PrErpNo
                                VALUES (@PrNo, @CompanyId, @DepartmentId, @UserId, @PrErpPrefix, @PrErpNo, @Edition, @PrDate, @DocDate, @PrRemark
                                , @TotalQty, @Amount, @PrStatus, @SignupStaus, @LockStaus, @ConfirmStatus, @BpmTransferStatus, @TransferStatus, @Priority, @Source
                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                PrNo = prNo,
                                CompanyId = companyId,
                                DepartmentId = departmentId,
                                UserId = userId,
                                prInfos.PrErpPrefix,
                                PrErpNo = prErpNo,
                                Edition = "0000",
                                prInfos.PrDate,
                                prInfos.DocDate,
                                prInfos.PrRemark,
                                TotalQty = prInfos.Details.Sum(x => x.PrQty), //同時計算單身數量
                                Amount = prInfos.Details.Sum(x => x.PrPrice), //同時加總單身金額
                                PrStatus = "N",
                                SignupStaus = "0",
                                LockStaus = "N",
                                ConfirmStatus = "N",
                                prInfos.Priority,
                                Source = "/ApiAddPrData",
                                BpmTransferStatus = "N",
                                BpmTransferUserId = userId,
                                BpmTransferDate = CreateDate,
                                TransferStatus = "N",
                                CreateDate,
                                LastModifiedDate,
                                CreateBy = userId,
                                LastModifiedBy = userId
                            });

                        var insertResult = sqlConnection.Query(sql, dynamicParameters);

                        rowsAffected += insertResult.Count();

                        foreach (var item in insertResult)
                        {
                            prId = item.PrId;
                        }
                        #endregion

                        #region //請購單身新增
                        prInfos.Details
                            .ToList()
                            .ForEach(x =>
                            {
                                x.PrId = prId;
                            });

                        foreach (var detail in prInfos.Details)
                        {
                            sql = @"INSERT INTO SCM.PrDetail (PrId, PrSequence, MtlItemId, PrMtlItemName, PrMtlItemSpec, InventoryId, PrUomId, PrQty
                                    , DemandDate, SupplierId, PrCurrency, PrExchangeRate, PrUnitPrice, PrPrice, PrPriceTw, UrgentMtl
                                    , ProductionPlan, SoDetailId, PoQty, PoUomId, PoCurrency, PoUnitPrice, PoPrice, PrPriceQty, PrPriceUomId
                                    , LockStaus, PoStaus, PartialPurchaseStaus, InquiryStatus, TaxNo, Taxation, BusinessTaxRate, DetailMultiTax
                                    , TradeTerm, MtlInventory, MtlInventoryQty, ConfirmStatus, ClosureStatus, PrDetailRemark
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.PrDetailId, INSERTED.PrSequence
                                    VALUES (@PrId, @PrSequence, @MtlItemId, @PrMtlItemName, @PrMtlItemSpec, @InventoryId, @PrUomId, @PrQty
                                    , @DemandDate, @SupplierId, @PrCurrency, @PrExchangeRate, @PrUnitPrice, @PrPrice, @PrPriceTw, @UrgentMtl
                                    , @ProductionPlan, @SoDetailId, @PoQty, @PoUomId, @PoCurrency, @PoUnitPrice, @PoPrice, @PrPriceQty, @PrPriceUomId
                                    , @LockStaus, @PoStaus, @PartialPurchaseStaus, @InquiryStatus, @TaxNo, @Taxation, @BusinessTaxRate, @DetailMultiTax
                                    , @TradeTerm, @MtlInventory, @MtlInventoryQty, @ConfirmStatus, @ClosureStatus, @PrDetailRemark
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            var insertDetailResult = sqlConnection.Query(sql, detail);

                            rowsAffected += insertDetailResult.Count();

                            foreach (var item in insertDetailResult)
                            {
                                prInfos.Details
                                    .Where(x => x.PrSequence == item.PrSequence.ToString())
                                    .ToList()
                                    .ForEach(x =>
                                    {
                                        x.PrDetailId = Convert.ToInt32(item.PrDetailId);
                                    });
                            }
                        }
                        #endregion

                        #region //判斷資料長度
                        DateTime tempDate = default(DateTime);
                        if (companyId <= 0) throw new SystemException("【所屬公司】不能為空!");
                        if (prInfos.PrErpPrefix.Length <= 0) throw new SystemException("【請購單別】不能為空!");
                        if (prInfos.UserNo.Length <= 0) throw new SystemException("【請購人員】不能為空!");
                        if (prInfos.DepartmentNo.Length <= 0) throw new SystemException("【請購部門】不能為空!");
                        if (prInfos.DocDate.Length <= 0) throw new SystemException("【請購單據日期】不能為空!");
                        if (!DateTime.TryParse(prInfos.DocDate, out tempDate)) throw new SystemException("【請購單據日期】格式錯誤!");
                        if (prInfos.PrDate.Length <= 0) throw new SystemException("【請購日期】不能為空!");
                        if (!DateTime.TryParse(prInfos.PrDate, out tempDate)) throw new SystemException("【請購日期】格式錯誤!");
                        if (prInfos.Priority.Length <= 0) throw new SystemException("【優先度】不能為空!");

                        if (prInfos.Details.Count <= 0) throw new SystemException("【請購單單身】不能為空!");
                        if (prInfos.Details.Any(x => x.MtlItemNo.Length <= 0)) throw new SystemException("【品號】不能為空!");
                        if (prInfos.Details.Any(x => x.MtlItemNo.Length > 40)) throw new SystemException("【品號】長度錯誤!");
                        if (prInfos.Details.Any(x => x.DemandDate.Length <= 0)) throw new SystemException("【需求日期】不能為空!");
                        if (prInfos.Details.Any(x => !DateTime.TryParse(x.DemandDate, out tempDate))) throw new SystemException("【需求日期】不能為空!");
                        if (prInfos.Details.Any(x => x.SupplierNo.Length <= 0)) throw new SystemException("【供應商】不能為空!");
                        if (prInfos.Details.Any(x => x.SupplierNo.Length > 10)) throw new SystemException("【供應商】長度錯誤!");
                        if (prInfos.Details.Any(x => x.PrCurrency.Length <= 0)) throw new SystemException("【幣別】不能為空!");
                        if (prInfos.Details.Any(x => x.PrCurrency.Length > 10)) throw new SystemException("【幣別】長度錯誤!");
                        if (prInfos.Details.Any(x => x.InventoryNo.Length <= 0)) throw new SystemException("【主要庫別】不能為空!");
                        if (prInfos.Details.Any(x => x.InventoryNo.Length > 10)) throw new SystemException("【主要庫別】長度錯誤!");
                        if (prInfos.Details.Any(x => x.PrQty <= 0)) throw new SystemException("【請購數量】不能為空!");
                        if (prInfos.Details.Any(x => x.PrUnitPrice <= 0)) throw new SystemException("【請購單價】不能為空!");
                        if (prInfos.Details.Any(x => x.UrgentMtl.Length <= 0)) throw new SystemException("【急料】不能為空!");
                        if (prInfos.Details.Any(x => x.ProductionPlan.Length <= 0)) throw new SystemException("【納入生產計畫】不能為空!");
                        #endregion
                    }

                    //回傳相關參數至Python
                    var data = new
                    {
                        PrId = prId,
                        prInfos.PrErpPrefix,
                        PrErpNo = prErpNo,
                        Details = from a in prInfos.Details
                                  select new
                                  {
                                      a.PrDetailId,
                                      a.PrSequence,
                                      a.MtlItemNo
                                  }
                    };

                    #region //Response
                    jsonResponse = JObject.FromObject(new
                    {
                        status = "success",
                        msg = "(" + rowsAffected + " rows affected)",
                        data
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

        #region //ApiAddPrFile -- Api請購單附檔新增 -- Yi 2023.10.16
        public string ApiAddPrFile(string CompanyNo, string UserNo, int Id, string Type, string ClientIP, List<FileModel> files)
        {
            try
            {
                if (!Regex.IsMatch(Type, "^(info|detail)$", RegexOptions.IgnoreCase)) throw new SystemException("【單據類型】錯誤!");

                using (TransactionScope transactionScope = new TransactionScope())
                {
                    #region //相關參數設定
                    int rowsAffected = 0, companyId = -1, userId = -1, fileId = -1, prId = -1, prDetailId = -1;
                    string companyNo = "";
                    #endregion

                    //查詢MES相關資料
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷公司資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 CompanyId, CompanyName, ErpDb, CompanyNo, ErpNo
                                FROM BAS.Company
                                WHERE CompanyNo = @CompanyNo";
                        dynamicParameters.Add("CompanyNo", CompanyNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【公司】資料錯誤!");

                        foreach (var item in result)
                        {
                            companyId = Convert.ToInt32(item.CompanyId);
                            companyNo = item.CompanyNo;
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        }
                        #endregion

                        #region //確認請購人員資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 UserId
                                FROM BAS.[User]
                                WHERE UserNo = @UserNo";
                        dynamicParameters.Add("UserNo", UserNo);

                        var resultUser = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUser.Count() <= 0) throw new SystemException("【使用者】資料錯誤!");

                        foreach (var item in resultUser)
                        {
                            userId = Convert.ToInt32(item.UserId);
                        }
                        #endregion

                        //判斷傳入之ID，隸屬單頭或單身
                        switch (Type)
                        {
                            case "info":
                                #region //判斷請購單【單頭】資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM SCM.PurchaseRequisition a
                                        WHERE a.PrId = @PrId";
                                dynamicParameters.Add("PrId", Id);

                                var resultPrInfo = sqlConnection.Query(sql, dynamicParameters);
                                if (resultPrInfo.Count() <= 0) throw new SystemException("【請購單單頭】資料錯誤!");
                                #endregion

                                #region //新增BAS.[File]及SCM.PrFile
                                files.ForEach(file =>
                                {
                                    sql = @"INSERT INTO BAS.[File] (CompanyId, FileName, FileContent, FileExtension, FileSize
                                          , ClientIP, Source, DeleteStatus
                                          , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                          OUTPUT INSERTED.FileId
                                          VALUES (@CompanyId, @FileName, @FileContent, @FileExtension, @FileSize
                                          , @ClientIP, @Source, @DeleteStatus
                                          , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            CompanyId = companyId,
                                            file.FileName,
                                            file.FileContent,
                                            file.FileExtension,
                                            file.FileSize,
                                            ClientIP,
                                            Source = "/ApiAddPrFile",
                                            DeleteStatus = "N",
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy = userId,
                                            LastModifiedBy = userId
                                        });
                                    var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected = insertResult.Count();

                                    foreach (var innerItem in insertResult)
                                    {
                                        fileId = Convert.ToInt32(innerItem.FileId);
                                    }

                                    #region //新增SCM.PrFile
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO SCM.PrFile (PrId, FileId
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.PrFileId
                                            VALUES (@PrId, @FileId
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            PrId = Id,
                                            FileId = fileId,
                                            PrDetailId = prDetailId <= 0 ? (int?)null : prDetailId,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy = userId,
                                            LastModifiedBy = userId
                                        });

                                    var insertResultFile = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResultFile.Count();
                                    #endregion
                                });
                                #endregion

                                break;
                            case "detail":
                                #region //判斷請購【單身】資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 a.PrId, b.PrDetailId
                                        FROM SCM.PurchaseRequisition a
                                        INNER JOIN SCM.PrDetail b ON a.PrId = b.PrId
                                        WHERE b.PrDetailId = @PrDetailId";
                                dynamicParameters.Add("PrDetailId", Id);

                                var resultPrDetail = sqlConnection.Query(sql, dynamicParameters);
                                if (resultPrDetail.Count() <= 0) throw new SystemException("【請購單單身】資料錯誤!");

                                foreach (var item in resultPrDetail)
                                {
                                    prId = item.PrId;
                                }
                                #endregion

                                #region //新增BAS.[File]及SCM.PrFile
                                files.ForEach(file =>
                                {
                                    sql = @"INSERT INTO BAS.[File] (CompanyId, FileName, FileContent, FileExtension, FileSize
                                          , ClientIP, Source, DeleteStatus
                                          , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                          OUTPUT INSERTED.FileId
                                          VALUES (@CompanyId, @FileName, @FileContent, @FileExtension, @FileSize
                                          , @ClientIP, @Source, @DeleteStatus
                                          , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            CompanyId = companyId,
                                            file.FileName,
                                            file.FileContent,
                                            file.FileExtension,
                                            file.FileSize,
                                            ClientIP,
                                            Source = "/ApiAddPrFile",
                                            DeleteStatus = "N",
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy = userId,
                                            LastModifiedBy = userId
                                        });
                                    var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected = insertResult.Count();

                                    foreach (var innerItem in insertResult)
                                    {
                                        fileId = Convert.ToInt32(innerItem.FileId);
                                    }

                                    #region //新增SCM.PrFile
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO SCM.PrFile (PrId, FileId, PrDetailId
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.PrFileId
                                            VALUES (@PrId, @FileId, @PrDetailId
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            PrId = prId,
                                            FileId = fileId,
                                            PrDetailId = Id,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy = userId,
                                            LastModifiedBy = userId
                                        });

                                    var insertResultFile = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResultFile.Count();
                                    #endregion
                                });
                                #endregion

                                break;
                        }
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

        #region //ApiPrTransferBpm -- 拋轉請購單據至BPM -- Yi 2023.10.16
        public string ApiPrTransferBpm(string CompanyNo, string UserNo, int PrId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    #region //相關參數設定
                    int rowsAffected = 0, companyId = -1, userId = -1;
                    string departmentNo = "", departmentName = "", prNo = ""
                        , docDate = "", prDate = "", prRemark = "", prErpPrefix = "", priority = ""
                        , mtlInventory = "目前尚無資料", proId = "", token = "", tempToken = "", companyNo = "", supplierName = ""
                        , bpmUserId = "", bpmRoleId = "", bpmUserNo = "", userName = "", bpmDepNo = "", bpmDepName = "", hideSupplier = "Y";
                    double amount = 0, CheckMin = 0, totalAmountTw = 0;

                    JArray attachFilePath = new JArray();
                    JArray tabPrDetailData = new JArray();
                    DateTime nowDate = DateTime.Now;
                    #endregion

                    IEnumerable<dynamic> prDetailResult;
                    IEnumerable<dynamic> systemTokenResult;

                    //查詢MES相關資料
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷公司資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 CompanyId, CompanyName, ErpDb, CompanyNo, ErpNo
                                FROM BAS.Company
                                WHERE CompanyNo = @CompanyNo";
                        dynamicParameters.Add("CompanyNo", CompanyNo);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("【公司】資料錯誤!");

                        foreach (var item in result)
                        {
                            companyId = Convert.ToInt32(item.CompanyId);
                            companyNo = item.CompanyNo;
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                        }
                        #endregion

                        #region //判斷請購單資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.PrNo, a.PrStatus, FORMAT(a.DocDate, 'yyyy-MM-dd') DocDate, a.PrErpPrefix
                                , FORMAT(a.PrDate, 'yyyy-MM-dd') PrDate, ISNULL(a.PrRemark, '') PrRemark, ISNULL(a.Amount, 0) Amount
                                , a.UserId, a.Priority
                                , b.DepartmentNo, b.DepartmentName
                                FROM SCM.PurchaseRequisition a
                                INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                                WHERE a.PrId = @PrId";
                        dynamicParameters.Add("PrId", PrId);

                        var resultPrInfo = sqlConnection.Query(sql, dynamicParameters);
                        if (resultPrInfo.Count() <= 0) throw new SystemException("請購單資料錯誤!");

                        foreach (var item in resultPrInfo)
                        {
                            if (item.PrStatus != "N" && item.PrStatus != "E" && item.PrStatus != "R") throw new SystemException("請購單狀態無法更改!");
                            prNo = item.PrNo;
                            docDate = item.DocDate;
                            prDate = item.PrDate;
                            prRemark = item.PrRemark;
                            amount = item.Amount;
                            userId = item.UserId;
                            departmentNo = item.DepartmentNo;
                            departmentName = item.DepartmentName;
                            prErpPrefix = item.PrErpPrefix;
                            priority = item.Priority;
                        }
                        #endregion

                        #region //確認請購人員資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 UserId
                                FROM BAS.[User]
                                WHERE UserNo = @UserNo";
                        dynamicParameters.Add("UserNo", UserNo);

                        var resultUser = sqlConnection.Query(sql, dynamicParameters);
                        if (resultUser.Count() <= 0) throw new SystemException("【使用者】資料錯誤!");

                        foreach (var item in resultUser)
                        {
                            userId = Convert.ToInt32(item.UserId);
                        }
                        #endregion

                        #region //取得單身資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.PrDetailId, a.PrSequence, a.PrMtlItemName, a.PrMtlItemSpec, a.PrQty
                                , FORMAT(a.DemandDate, 'yyyy-MM-dd') DemandDate, a.PrCurrency, a.PrExchangeRate
                                , a.PrUnitPrice, a.PrPrice, a.PrPriceTw, a.UrgentMtl, a.ProductionPlan, a.Project
                                , a.PrDetailRemark, a.MtlInventory, a.MtlInventoryQty
                                , b.MtlItemNo
                                , c.InventoryNo, c.InventoryName
                                , d.UomNo
                                , ISNULL(e.SupplierNo, '') SupplierNo, ISNULL(e.SupplierName, '') SupplierName
                                , f.SoSequence
                                , g.SoErpPrefix, g.SoErpNo
                                FROM SCM.PrDetail a
                                INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
                                INNER JOIN SCM.Inventory c ON a.InventoryId = c.InventoryId
                                INNER JOIN PDM.UnitOfMeasure d ON a.PrUomId = d.UomId
                                LEFT JOIN SCM.Supplier e ON a.SupplierId = e.SupplierId
                                LEFT JOIN SCM.SoDetail f ON a.SoDetailId = f.SoDetailId
                                LEFT JOIN SCM.SaleOrder g ON f.SoId = g.SoId
                                WHERE a.PrId = @PrId";
                        dynamicParameters.Add("PrId", PrId);

                        prDetailResult = sqlConnection.Query(sql, dynamicParameters);
                        if (prDetailResult.Count() <= 0) throw new SystemException("請購單身資料錯誤!");

                        foreach (var item in prDetailResult)
                        {
                            supplierName = item.SupplierName;
                        }
                        #endregion

                        #region //取得BPM TOKEN
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.IpAddress, a.CompanyId, a.Token, FORMAT(a.VerifyDate, 'yyyy-MM-dd HH:mm:ss') VerifyDate
                                FROM BPM.SystemToken a
                                WHERE IpAddress = @IpAddress
                                AND CompanyId = @CompanyId";
                        dynamicParameters.Add("IpAddress", BpmServerPath);
                        dynamicParameters.Add("CompanyId", companyId);

                        systemTokenResult = sqlConnection.Query(sql, dynamicParameters);
                        if (systemTokenResult.Count() <= 0) throw new SystemException("查無此憑證!");
                        #endregion

                        #region //依公司別取得ProId
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.TypeName ProId
                                FROM BAS.[Type] a
                                WHERE a.TypeSchema = 'BPM.PrProId'
                                AND a.TypeNo = @CompanyNo";
                        dynamicParameters.Add("CompanyNo", companyNo);

                        var ProIdResult = sqlConnection.Query(sql, dynamicParameters);

                        if (ProIdResult.Count() <= 0) throw new SystemException("此公司別【" + companyNo + "】尚未建立拋轉起單碼，請聯繫資訊人員!!");

                        foreach (var item in ProIdResult)
                        {
                            proId = item.ProId;
                        }
                        #endregion

                        #region //組檔案JSON
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.FileId, b.FileName, b.FileExtension, b.FileSize, b.FileContent
                                FROM SCM.PrFile a
                                INNER JOIN BAS.[File] b ON a.FileId = b.FileId
                                WHERE a.PrId = @PrId";
                        dynamicParameters.Add("PrId", PrId);

                        var prFileResult = sqlConnection.Query(sql, dynamicParameters);

                        foreach (var item in prFileResult)
                        {
                            attachFilePath.Add(JObject.FromObject(new
                            {
                                fileId = item.FileId,
                                fileName = item.FileName + item.FileExtension
                            }));
                        }
                        #endregion
                    }

                    //查詢BPM相關資料
                    using (SqlConnection sqlConnection = new SqlConnection(BpmDbConnectionStrings))
                    {
                        #region //確認TOKEN是否重取
                        foreach (var item in systemTokenResult)
                        {
                            tempToken = item.Token;
                            DateTime verifyDate = Convert.ToDateTime(item.VerifyDate);
                            CheckMin = (nowDate - verifyDate).TotalMinutes;
                            if (CheckMin >= 30)
                            {
                                #region //取得新BPM TOKEN
                                string tokenResponse = BpmHelper.GetBpmToken(BpmServerPath, BpmAccount, BpmPassword);
                                var tokenJson = JObject.Parse(tokenResponse);
                                foreach (var item2 in tokenJson)
                                {
                                    if (item2.Key == "status")
                                    {
                                        if (item2.Value.ToString() != "success") throw new SystemException("取得token失敗!");
                                    }
                                    else if (item2.Key == "data")
                                    {
                                        token = item2.Value.ToString();
                                    }
                                }
                                #endregion
                            }
                            else
                            {
                                token = item.Token;
                            }
                        }
                        #endregion

                        #region //取得BpmUser資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"WITH BasicUserInfo(MemID, LoginID, UserName, MainRoleID, RolID, RolName, ParentRol) AS(
                                SELECT a.MemID, a.LoginID, a.UserName, a.MainRoleID
                                , b.RolID, b.Name, b.DepID AS ParentRol
                                FROM Mem_GenInf a
                                LEFT JOIN Rol_GenInf b ON b.RolID = a.MainRoleID
                                WHERE a.LoginID = @LoginID
                                UNION ALL
                                SELECT a.MemID, a.LoginID, a.UserName, a.MainRoleID
                                , b.RolID, b.Name, b.DepID AS ParentRol
                                FROM BasicUserInfo a, Rol_GenInf b
                                WHERE a.ParentRol = b.RolID
                                )
                                SELECT a.MemID, a.LoginID, a.UserName, a.MainRoleID
                                , b.Name AS RolName
                                , c.DepID, c.ID AS DepNo, c.Name AS DepName
                                FROM BasicUserInfo a
                                LEFT JOIN Rol_GenInf AS parentRol_GenInf ON a.RolID = parentRol_GenInf.RolID
                                LEFT JOIN Dep_GenInf c ON parentRol_GenInf.DepID = c.DepID
                                , Rol_GenInf b
                                WHERE a.MainRoleID = b.RolID
                                AND c.DepID IS NOT NULL
                                UNION
                                SELECT a.MemID, a.LoginID, a.UserName, a.MainRoleID
                                , b.Name AS RolName
                                , c.ComID AS DepID, c.ID AS DepNo, c.Name AS DepName
                                FROM Mem_GenInf a
                                LEFT JOIN Rol_GenInf b ON a.MainRoleID = b.RolID
                                LEFT JOIN Company c ON b.DepID = c.ComID
                                WHERE c.ComID IS NOT NULL
                                AND a.LoginID = @LoginID
                                ORDER BY a.LoginID";
                        dynamicParameters.Add("LoginID", UserNo);

                        var MemGenInfoResult = sqlConnection.Query(sql, dynamicParameters);

                        if (MemGenInfoResult.Count() <= 0) throw new SystemException("取得BPM使用者資訊時發生錯誤!!");

                        foreach (var item in MemGenInfoResult)
                        {
                            bpmUserId = item.MemID;
                            bpmRoleId = item.MainRoleID;
                            bpmUserNo = item.LoginID;
                            userName = item.UserName;
                            bpmDepNo = item.DepNo;
                            bpmDepName = item.DepName;
                        }
                        #endregion

                        #region //組單身資料
                        foreach (var item in prDetailResult)
                        {
                            if (supplierName != "")
                            {
                                if (supplierName.IndexOf("中揚") == -1 && supplierName.IndexOf("中扬") == -1 && supplierName.IndexOf("晶彩") == -1 && supplierName.IndexOf("群英") == -1 && supplierName.IndexOf("紘立") == -1)
                                {
                                    hideSupplier = "N";
                                }
                            }
                            else
                            {
                                hideSupplier = "N";
                            }

                            totalAmountTw = totalAmountTw + item.PrPriceTw;

                            #region //處理MtlInventoryStatus
                            string mtlInventoryStatus = "目前尚無資料";
                            if (mtlInventory != "目前尚無資料")
                            {
                                JObject jo = JObject.Parse(item.MtlInventory);
                                JToken jtoken = jo["data"];

                                for (int i = 0; i < jtoken.Count(); i++)
                                {
                                    mtlInventoryStatus += jtoken[i]["WAREHOUSE_NO"] + ":" + jtoken[i]["WAREHOUSE_NAME"] + ":" + jtoken[i]["WAREHOUSE_QTY"] + ";";
                                }
                            }
                            #endregion

                            tabPrDetailData.Add(JObject.FromObject(new
                            {
                                PRNumber_SN = item.PrSequence,
                                ItemNo = item.MtlItemNo,
                                ItemName = item.PrMtlItemName,
                                Spec = item.PrMtlItemSpec,
                                WareHouse = item.InventoryNo + "-" + item.InventoryName,
                                Qty = item.PrQty.ToString(),
                                Unit = item.UomNo,
                                RequiredDate = item.DemandDate,
                                Supplier = item.SupplierNo + "-" + item.SupplierName,
                                Currency = item.PrCurrency,
                                ExchangeRate = item.PrExchangeRate,
                                Price = item.PrUnitPrice.ToString(),
                                Amount = item.PrPrice.ToString(),
                                AmountInLocalCurrency = item.PrPriceTw.ToString(),
                                Ckb_UrgentMaterial = item.UrgentMtl,
                                Ckb_MRP = item.ProductionPlan,
                                Project = "",
                                LineRemark = item.PrDetailRemark,
                                ReferenceDoc = item.SoErpPrefix != null ? item.SoErpPrefix + "-" + item.SoErpNo + "-" + item.SoSequence : "",
                                MtlInventoryStatus = mtlInventoryStatus,
                                item.MtlInventoryQty
                            }));
                        }
                        #endregion
                    }

                    //寫入MES相關資料
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (CheckMin >= 30)
                        {
                            #region //將新的TOKEN更新回MES
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE BPM.SystemToken SET
                                    Token = @Token,
                                    VerifyDate = @VerifyDate,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE IpAddress = @IpAddress
                                    AND CompanyId = @CompanyId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    Token = token,
                                    VerifyDate = nowDate,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    IpAddress = BpmServerPath,
                                    CompanyId = companyId
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }
                        else
                        {
                            token = tempToken;
                        }

                        #region //組請購BPM資料
                        string memId = bpmUserId;
                        string rolId = bpmRoleId;
                        string startMethod = "NoOpFirst";

                        JObject artInsAppData = JObject.FromObject(new
                        {
                            Title = "MES請購單 = " + prNo + ", ERP請購單別: " + prErpPrefix,
                            PRNumber_MES = prNo,
                            CreateDate = docDate,
                            Requisitioner = bpmUserNo,
                            RequisitionerName = userName,
                            RequisitionDeptID = departmentNo,
                            RequisitionDept = departmentName,
                            RequisitionDate = prDate,
                            HeaderRemark = prRemark,
                            AmountInLocalCurrencyTotal = totalAmountTw.ToString(),
                            tabPR_Line_Data = JsonConvert.SerializeObject(tabPrDetailData),
                            attachFilePath = JsonConvert.SerializeObject(attachFilePath),
                            dbTable = "SCM.PurchaseRequisition",
                            mesID = PrId,
                            company = CompanyNo,
                            HideSupplier = hideSupplier,
                            PRNumber_FormType = prErpPrefix,
                            priority
                        });
                        #endregion

                        string sData = BpmHelper.PostFormToBpm(token, proId, memId, rolId, startMethod, artInsAppData, BpmServerPath);

                        if (sData == "true")
                        {
                            #region //更改BPM拋轉狀態
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.PurchaseRequisition SET
                                    PrStatus = 'P',
                                    BpmTransferStatus = 'Y',
                                    BpmTransferUserId = @BpmTransferUserId,
                                    BpmTransferDate = @BpmTransferDate,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE PrId = @PrId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    BpmTransferUserId = userId,
                                    BpmTransferDate = CreateDate,
                                    LastModifiedDate,
                                    LastModifiedBy = userId,
                                    PrId
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion
                        }
                        else
                        {
                            #region //更改BPM拋轉狀態
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.PurchaseRequisition SET
                                    PrStatus = 'R',
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE PrId = @PrId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    LastModifiedDate,
                                    LastModifiedBy = userId,
                                    PrId
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            throw new SystemException("請購單拋轉BPM失敗!");
                        }
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
        #region //UpdatePurchaseRequisition -- 更新請購單資料 -- Ann 2023-01-09
        public string UpdatePurchaseRequisition(int PrId, string PrErpPrefix, string DocDate, string PrDate, string PrRemark, string PrFile, int UserId, int DepartmentId, string Priority, string BomType)
        {
            try
            {
                if (PrErpPrefix.Length < 0) throw new SystemException("【請購單別】不能為空!");
                if (DocDate.Length < 0) throw new SystemException("【請購單據日期】不能為空!");
                if (PrDate.Length < 0) throw new SystemException("【請購日期】不能為空!");
                if (Priority.Length < 0) throw new SystemException("【優先度】不能為空!");
                if (BomType.Length < 0) throw new SystemException("【請購物料類型】不能為空!");

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
                            #region //比對ERP關帳日期(庫存關帳)
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MA012)) MA012
                                    FROM CMSMA";
                            var cmsmaResult = sqlConnection2.Query(sql, dynamicParameters);
                            if (cmsmaResult.Count() < 0) throw new SystemException("EPR關帳日期資料錯誤!!");

                            foreach (var item in cmsmaResult)
                            {
                                string eprDate = item.MA012;
                                string erpYear = eprDate.Substring(0, 4);
                                string erpMonth = eprDate.Substring(4, 2);
                                string erpFullDate = erpYear + "-" + erpMonth;
                                DateTime erpDateTime = Convert.ToDateTime(erpFullDate);
                                DateTime DocDateDateTime = Convert.ToDateTime(DocDate.Split('-')[0] + "-" + DocDate.Split('-')[1]);
                                int compare = DocDateDateTime.CompareTo(erpDateTime);
                                if (compare <= 0) throw new SystemException("ERP已關帳(" + erpFullDate + ")，無法開立此日期(" + DocDate + ")之單據!!");
                            }
                            #endregion

                            #region //判斷請購單資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.PrErpPrefix, a.PrStatus, a.UserId
                                    FROM SCM.PurchaseRequisition a
                                    LEFT JOIN SCM.PrFile b ON a.PrId = b.PrId
                                    WHERE a.PrId = @PrId";
                            dynamicParameters.Add("PrId", PrId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("請購單資料錯誤!");

                            foreach (var item in result)
                            {
                                if (item.PrStatus != "N" && item.PrStatus != "E") throw new SystemException("請購單狀態無法更改!");
                                if (item.UserId != CreateBy) throw new SystemException("無法修改其他請購人之單據!");
                                if (item.PrErpPrefix != PrErpPrefix) throw new SystemException("請購單別不可修改!!");
                            }
                            #endregion

                            #region //確認請購人員資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.[User] a
                                    WHERE a.UserId = @UserId";
                            dynamicParameters.Add("UserId", UserId);

                            var UserResult = sqlConnection.Query(sql, dynamicParameters);
                            if (UserResult.Count() <= 0) throw new SystemException("【使用者】資料錯誤，請重新輸入!");
                            #endregion

                            #region //確認請購部門資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.Department a
                                    WHERE a.DepartmentId = @DepartmentId";
                            dynamicParameters.Add("DepartmentId", DepartmentId);

                            var DepartmentResult = sqlConnection.Query(sql, dynamicParameters);
                            if (DepartmentResult.Count() <= 0) throw new SystemException("【請購部門】資料錯誤，請重新輸入!");
                            #endregion

                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.PurchaseRequisition SET
                                    CompanyId = @CompanyId,
                                    DepartmentId = @DepartmentId,
                                    UserId = @UserId,
                                    PrErpPrefix = @PrErpPrefix,
                                    PrDate = @PrDate,
                                    DocDate = @DocDate,
                                    PrRemark = @PrRemark,
                                    Priority = @Priority,
                                    BomType = @BomType,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE PrId = @PrId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    CompanyId = CurrentCompany,
                                    DepartmentId,
                                    UserId,
                                    PrErpPrefix,
                                    PrDate,
                                    DocDate,
                                    PrRemark,
                                    Priority,
                                    BomType,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    PrId
                                });

                            int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                            if (PrFile.Length > 0)
                            {
                                #region //先將原本的砍掉
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE FROM SCM.PrFile
                                        WHERE PrId = @PrId";
                                dynamicParameters.Add("PrId", PrId);

                                rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                string[] prFiles = PrFile.Split(',');
                                foreach (var file in prFiles)
                                {
                                    #region //更新SCM.PrFile
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO SCM.PrFile (PrId, FileId
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.PrFileId
                                            VALUES (@PrId, @FileId
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            PrId,
                                            FileId = file,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    var insertResult3 = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult3.Count();
                                    #endregion
                                }
                            }
                            else
                            {
                                throw new SystemException("請購單必須上傳附件!");
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

        #region //UpdatePrDetail -- 更新請購單詳細資料 -- Ann 2023-01-09
        public string UpdatePrDetail(int PrDetailId, int PrId, string PrSequence, int MtlItemId, string PrMtlItemName, string PrMtlItemSpec, int InventoryId, int PrUomId, int PrQty, string DemandDate
            , int SupplierId, string PrCurrency, string PrExchangeRate, double PrUnitPrice, double PrPrice, double PrPriceTw
            , string UrgentMtl, string ProductionPlan, string Project, int SoDetailId, string PoRemark, string PrDetailRemark, string PrFile)
        {
            try
            {
                AddUserOperateLog();
                if (PrQty <= 0) throw new SystemException("【請購數量】不能為空!");
                if (DemandDate.Length <= 0) throw new SystemException("【需求日期】不能為空!");
                if (PrExchangeRate.Length <= 0) throw new SystemException("【匯率】不能為空!");
                //if (PrPrice <= 0) throw new SystemException("【請購金額】不能為空!");
                //if (PrPriceTw <= 0) throw new SystemException("【本幣金額】不能為空!");
                if (UrgentMtl.Length <= 0) throw new SystemException("【是否急料】不能為空!");
                if (ProductionPlan.Length <= 0) throw new SystemException("【是否納入生產計畫】不能為空!");
                if (PrMtlItemName.Length <= 0) throw new SystemException("【品名】不能為空!");
                //if (PrMtlItemSpec.Length <= 0) throw new SystemException("【規格】不能為空!");

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
                            #region //判斷請購單詳細資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.PrQty, a.PrPrice, a.PrPriceTw
                                    FROM SCM.PrDetail a
                                    WHERE a.PrDetailId = @PrDetailId";
                            dynamicParameters.Add("PrDetailId", PrDetailId);

                            var prDetailIdResult = sqlConnection.Query(sql, dynamicParameters);
                            if (prDetailIdResult.Count() <= 0) throw new SystemException("請購單詳細資料錯誤!");

                            double OrgPrQty = -1;
                            double OrgPrPrice = -1;
                            double OrgPrPriceTw = -1;
                            foreach (var item in prDetailIdResult)
                            {
                                OrgPrQty = item.PrQty;
                                OrgPrPrice = item.PrPrice;
                                OrgPrPriceTw = item.PrPriceTw;
                            }
                            #endregion

                            #region //判斷請購單資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.PrErpPrefix, a.PrStatus, a.UserId
                                    FROM SCM.PurchaseRequisition a
                                    WHERE a.PrId = @PrId";
                            dynamicParameters.Add("PrId", PrId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("請購單資料錯誤!");

                            string PrErpPrefix = "";
                            foreach (var item in result)
                            {
                                if (item.PrStatus != "N" && item.PrStatus != "E") throw new SystemException("請購單狀態無法更改!");
                                if (item.UserId != CreateBy) throw new SystemException("無法修改其他請購人之單據!");
                                PrErpPrefix = item.PrErpPrefix;
                            }
                            #endregion

                            #region //判斷品號資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.MtlItemNo, a.MtlItemName, a.MtlItemSpec
                                    FROM PDM.MtlItem a
                                    WHERE a.MtlItemId = @MtlItemId";
                            dynamicParameters.Add("MtlItemId", MtlItemId);

                            var result2 = sqlConnection.Query(sql, dynamicParameters);

                            if (result2.Count() <= 0) throw new SystemException("【品號】資料錯誤!");

                            string MtlItemNo = "";
                            foreach (var item in result2)
                            {
                                //if (PrMtlItemSpec.Length <= 0 && item.MtlItemSpec != "") throw new SystemException("【規格】不能為空!");
                                MtlItemNo = item.MtlItemNo;
                            }
                            #endregion

                            #region //判斷ERP品號生效日與失效日
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MB030)) MB030, LTRIM(RTRIM(MB031)) MB031
                                    FROM INVMB
                                    WHERE MB001 = @MB001";
                            dynamicParameters.Add("MB001", MtlItemNo);

                            var INVMBResult = sqlConnection2.Query(sql, dynamicParameters);

                            foreach (var item in INVMBResult)
                            {
                                if (item.MB030 != "" && item.MB030 != null)
                                {
                                    #region //判斷日期需大於或等於生效日
                                    string EffectiveDate = item.MB030;
                                    string effYear = EffectiveDate.Substring(0, 4);
                                    string effMonth = EffectiveDate.Substring(4, 2);
                                    string effDay = EffectiveDate.Substring(6, 2);
                                    DateTime effFullDate = Convert.ToDateTime(effYear + "-" + effMonth + "-" + effDay);
                                    int effresult = DateTime.Compare(CreateDate, effFullDate);
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
                                    int effresult = DateTime.Compare(CreateDate, effFullDate);
                                    if (effresult > 0) throw new SystemException("此品號已失效，無法使用!!");
                                    #endregion
                                }
                            }
                            #endregion

                            #region //判斷庫別資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.InventoryNo, a.InventoryName
                                    FROM SCM.Inventory a
                                    WHERE a.InventoryId = @InventoryId";
                            dynamicParameters.Add("InventoryId", InventoryId);

                            var result3 = sqlConnection.Query(sql, dynamicParameters);

                            if (result3.Count() <= 0) throw new SystemException("【庫別】資料錯誤!");

                            string InventoryNo = "";
                            string InventoryName = "";
                            foreach (var item in result3)
                            {
                                InventoryNo = item.InventoryNo;
                                InventoryName = item.InventoryName;
                            }
                            #endregion

                            #region //判斷單位資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM PDM.UnitOfMeasure a
                                    WHERE a.UomId = @UomId";
                            dynamicParameters.Add("UomId", PrUomId);

                            var result4 = sqlConnection.Query(sql, dynamicParameters);

                            if (result4.Count() <= 0) throw new SystemException("【單位】資料錯誤!");
                            #endregion

                            #region //取得ERP庫存資訊
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(a.MC007)) InventoryQty
                                    FROM INVMC a
                                    WHERE a.MC001 = @MC001
                                    AND a.MC002 = @MC002";
                            dynamicParameters.Add("MC001", MtlItemNo);
                            dynamicParameters.Add("MC002", InventoryNo);

                            var result5 = sqlConnection2.Query(sql, dynamicParameters);

                            double InventoryQty = 0;
                            string mtlInventory = "目前尚無資料";
                            if (result5.Count() > 0)
                            {
                                foreach (var item in result5)
                                {
                                    InventoryQty = Convert.ToDouble(item.InventoryQty);
                                    #region //組MtlInventory
                                    List<MtlInventory> mtlInventories = new List<MtlInventory>
                                    {
                                        new MtlInventory
                                        {
                                            WAREHOUSE_NO = InventoryNo,
                                            WAREHOUSE_NAME = InventoryName,
                                            WAREHOUSE_QTY = InventoryQty
                                        }
                                    };
                                    #endregion

                                    mtlInventory = JsonConvert.SerializeObject(mtlInventories);
                                    mtlInventory = "{\"data\":" + mtlInventory + "}";
                                }
                            }
                            #endregion

                            #region //檢查供應商資料是否正確
                            string TaxNo = "";
                            string Taxation = "";
                            string TradeTerm = "";
                            double? BusinessTaxRate = -1;
                            string HideSupplier = "";
                            if (SupplierId > 0)
                            {
                                #region //供應商
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.TaxNo, a.Taxation, a.TradeTerm, a.HideSupplier
                                        FROM SCM.Supplier a
                                        WHERE a.SupplierId = @SupplierId";
                                dynamicParameters.Add("SupplierId", SupplierId);

                                var result6 = sqlConnection.Query(sql, dynamicParameters);

                                if (result6.Count() <= 0) throw new SystemException("【供應商】資料錯誤!");

                                foreach (var item in result6)
                                {
                                    TaxNo = item.TaxNo;
                                    Taxation = item.Taxation;
                                    TradeTerm = item.TradeTerm;
                                    HideSupplier = item.HideSupplier;
                                }
                                #endregion

                                #region //查詢營業稅額資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT ROUND(LTRIM(RTRIM(NN004)),2) BusinessTaxRate
                                        FROM CMSNN 
                                        WHERE NN001 = @TaxNo";
                                dynamicParameters.Add("TaxNo", TaxNo);

                                var businessTaxRateResult = sqlConnection2.Query(sql, dynamicParameters);

                                if (businessTaxRateResult.Count() <= 0) throw new SystemException("【課稅別】資料錯誤!");

                                foreach (var item in businessTaxRateResult)
                                {
                                    BusinessTaxRate = item.BusinessTaxRate;
                                }
                                #endregion
                            }
                            #endregion

                            #region //確認集團內/集團外邏輯
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 b.HideSupplier
                                    FROM SCM.PrDetail a 
                                    INNER JOIN SCM.Supplier b ON a.SupplierId = b.SupplierId
                                    WHERE a.PrId = @PrId
                                    AND a.PrDetailId != @PrDetailId
                                    ORDER BY a.PrSequence";
                            dynamicParameters.Add("PrId", PrId);
                            dynamicParameters.Add("PrDetailId", PrDetailId);

                            var PrDetailResult2 = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item in PrDetailResult2)
                            {
                                string HideSupplierString = "";
                                if (item.HideSupplier == "Y")
                                {
                                    HideSupplierString = "集團內";
                                }
                                else
                                {
                                    HideSupplierString = "集團外";
                                }
                                if (item.HideSupplier != HideSupplier) throw new SystemException("此請購單第一筆單身為【" + HideSupplierString + "】，無法新增!!");
                            }
                            #endregion

                            #region //判斷幣別資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM CMSMF
                                    WHERE MF001 = @MF001";
                            dynamicParameters.Add("MF001", PrCurrency);

                            var result7 = sqlConnection2.Query(sql, dynamicParameters);
                            if (result7.Count() <= 0) throw new SystemException("【幣別】資料有誤!");
                            #endregion

                            #region //判斷MES專案代碼資料及專案預算卡控
                            double budgetAmount = 0;
                            if (Project.Length > 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.LocalBudgetAmount
                                        FROM SCM.ProjectDetail a 
                                        INNER JOIN SCM.Project b ON a.ProjectId = b.ProjectId
                                        WHERE a.ProjectType = '1'
                                        AND b.ProjectNo = @ProjectNo
                                        AND b.CompanyId = @CompanyId
                                        AND a.Status = 'Y'";
                                dynamicParameters.Add("ProjectNo", Project);
                                dynamicParameters.Add("CompanyId", CurrentCompany);

                                var ProjectDetailResult = sqlConnection.Query(sql, dynamicParameters);

                                foreach (var item in ProjectDetailResult)
                                {
                                    budgetAmount = item.LocalBudgetAmount;
                                }

                                #region //專案預算卡控
                                //此專案掛鉤的全部請購單金額
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT ISNULL(SUM(a.PrPriceTw), 0) TotalLocalAmount
                                        FROM SCM.PrDetail a 
                                        INNER JOIN SCM.Project b ON a.Project = b.ProjectNo
                                        INNER JOIN SCM.PurchaseRequisition c ON a.PrId = c.PrId
                                        WHERE b.ProjectNo = @ProjectNo
                                        AND b.CompanyId = @CompanyId
                                        AND (c.PrStatus IN ('Y', 'P')
                                        OR a.PrId = @PrId)";
                                dynamicParameters.Add("ProjectNo", Project);
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                dynamicParameters.Add("PrId", PrId);

                                var ProjectAmountResult = sqlConnection.Query(sql, dynamicParameters);

                                foreach (var item in ProjectAmountResult)
                                {
                                    double totalLocalAmount = item.TotalLocalAmount;
                                    if (totalLocalAmount + PrPriceTw > budgetAmount)
                                    {
                                        throw new SystemException("此次請購金額合計已超過專案預算金額(" + budgetAmount + ")!");
                                    }
                                }
                                #endregion
                            }
                            #endregion

                            #region //判斷專案代碼資料是否正確
                            if (Project.Length > 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                    FROM CMSNB
                                    WHERE NB001 = @NB001";
                                dynamicParameters.Add("NB001", Project);

                                var result8 = sqlConnection2.Query(sql, dynamicParameters);
                                if (result8.Count() <= 0) throw new SystemException("【專案代碼】資料有誤!");
                            }
                            #endregion

                            #region //確認訂單資料是否正確
                            if (SoDetailId > 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT MtlItemId, ConfirmStatus
                                        FROM SCM.SoDetail
                                        WHERE SoDetailId = @SoDetailId";
                                dynamicParameters.Add("SoDetailId", SoDetailId);

                                var result9 = sqlConnection.Query(sql, dynamicParameters);
                                if (result9.Count() <= 0) throw new SystemException("【訂單】資料有誤!");

                                foreach (var item in result9)
                                {
                                    if (item.ConfirmStatus != "Y") throw new SystemException("訂單尚未核單，無法綁定!!");
                                    //if (item.MtlItemId != MtlItemId) throw new SystemException("訂單品號與請購品號不同!!");
                                }
                            }
                            #endregion

                            #region //確認品號是否需要上傳相關文件才能進行請購(除晶彩外)
                            if (CurrentCompany == 11)
                            {
                                string pattern = @"^.*-T$";
                                if (Regex.IsMatch(MtlItemNo, pattern))
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM PDM.MtlItem a 
                                            OUTER APPLY (
                                                SELECT x.MtDocId, x.DocName
                                                FROM PDM.MtlItemDocSetting x 
                                                WHERE x.PurchaseMandatory = 'Y'
                                            ) x
                                            LEFT JOIN PDM.MtlItemFile b ON a.MtlItemNo = b.MtlItemNo AND b.MtDocId = x.MtDocId
                                            WHERE a.MtlItemNo = @MtlItemNo
                                            AND b.MtDocId IS NULL";
                                    dynamicParameters.Add("MtlItemNo", MtlItemNo);

                                    var MtlItemFileResult = sqlConnection.Query(sql, dynamicParameters);

                                    if (MtlItemFileResult.Count() > 0) throw new SystemException("品號【" + MtlItemNo + "】需上傳相關文件才可進行請購!!");
                                }
                            }
                            #endregion

                            #region //檢核特定相關卡控(目前僅晶彩邏輯)
                            if (CurrentCompany == 4)
                            {
                                double localPrUnitPrice = PrUnitPrice * Convert.ToDouble(PrExchangeRate);

                                #region //單別3109、3108特別卡控
                                if (PrErpPrefix == "3108" || PrErpPrefix == "3109")
                                {
                                    #region //請購單價不能超過1000(本幣)
                                    if (localPrUnitPrice > 1000) throw new SystemException("違反請購單別【3108、3109】卡控:<br>請購單價不能超過【1000】!!<br>此次請購單價【" + localPrUnitPrice + "】已超過!!");
                                    #endregion

                                    #region //所有單身合計金額(本幣)不能超過20000
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT ISNULL(SUM(PrPriceTw), 0) TotalPrice
                                            FROM SCM.PrDetail
                                            WHERE PrId = @PrId
                                            AND PrDetailId != @PrDetailId";
                                    dynamicParameters.Add("PrId", PrId);
                                    dynamicParameters.Add("PrDetailId", PrDetailId);

                                    var TotalPriceResult = sqlConnection.Query(sql, dynamicParameters);

                                    double totalPrice = 0;
                                    foreach (var item in TotalPriceResult)
                                    {
                                        totalPrice = item.TotalPrice;
                                        if (totalPrice + PrPriceTw > 20000) throw new SystemException("違反請購單別【3108、3109】卡控:<br>請購合計金額不能超過20000!!<br>目前請購合計金額【" + (totalPrice + PrPriceTw) + "】已超過!!");
                                    }
                                    #endregion
                                }
                                #endregion
                            }
                            #endregion

                            #region //Update SCM.PrDetail
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.PrDetail SET
                                    MtlItemId = @MtlItemId,
                                    PrMtlItemName = @PrMtlItemName,
                                    PrMtlItemSpec = @PrMtlItemSpec,
                                    InventoryId = @InventoryId,
                                    PrUomId = @PrUomId,
                                    PrQty = @PrQty,
                                    DemandDate = @DemandDate,
                                    SupplierId = @SupplierId,
                                    PrCurrency = @PrCurrency,
                                    PrExchangeRate = @PrExchangeRate,
                                    PrUnitPrice = @PrUnitPrice,
                                    PrPrice = @PrPrice,
                                    PrPriceTw = @PrPriceTw,
                                    UrgentMtl = @UrgentMtl,
                                    ProductionPlan = @ProductionPlan,
                                    Project = @Project,
                                    PrPriceQty = @PrQty,
                                    PrPriceUomId = @PrUomId,
                                    SoDetailId = @SoDetailId,
                                    MtlInventory = @MtlInventory,
                                    MtlInventoryQty = @MtlInventoryQty,
                                    PoUomId = @PoUomId,
                                    PoQty = @PoQty,
                                    PoCurrency = @PoCurrency,
                                    PoUnitPrice = @PoUnitPrice,
                                    PoPrice = @PoPrice,
                                    PrDetailRemark = @PrDetailRemark,
                                    TaxNo = @TaxNo,
                                    Taxation = @Taxation,
                                    TradeTerm = @TradeTerm,
                                    BusinessTaxRate = @BusinessTaxRate,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE PrDetailId = @PrDetailId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    MtlItemId,
                                    PrMtlItemName,
                                    PrMtlItemSpec,
                                    InventoryId,
                                    PrUomId,
                                    PrQty,
                                    DemandDate,
                                    SupplierId,
                                    PrCurrency,
                                    PrExchangeRate,
                                    PrUnitPrice,
                                    PrPrice,
                                    PrPriceTw,
                                    UrgentMtl,
                                    ProductionPlan,
                                    Project,
                                    PrPriceQty = PrQty,
                                    PrPriceUomId = PrUomId,
                                    SoDetailId,
                                    MtlInventory = mtlInventory,
                                    MtlInventoryQty = InventoryQty,
                                    PoUomId = PrUomId,
                                    PoQty = PrQty,
                                    PoCurrency = PrCurrency,
                                    PoUnitPrice = PrUnitPrice,
                                    PoPrice = PrPrice,
                                    PrDetailRemark,
                                    TaxNo = TaxNo != "" ? TaxNo : null,
                                    Taxation = Taxation != "" ? Taxation : null,
                                    TradeTerm = TradeTerm != "" ? TradeTerm : null,
                                    BusinessTaxRate = BusinessTaxRate > 0 ? BusinessTaxRate : (double?)null,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    PrDetailId
                                });

                            int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //更新請購單頭
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.PurchaseRequisition SET
                                    TotalQty = TotalQty - @OrgPrQty + @PrQty,
                                    Amount = Amount - @OrgPrPriceTw + @PrPriceTw,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE PrId = @PrId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    OrgPrQty,
                                    PrQty,
                                    OrgPrPriceTw,
                                    PrPriceTw,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    PrId
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            if (PrFile.Length > 0)
                            {
                                #region //先將原本的砍掉
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE FROM SCM.PrFile
                                        WHERE PrId = @PrId
                                        AND PrDetailId = @PrDetailId";
                                dynamicParameters.Add("PrId", PrId);
                                dynamicParameters.Add("PrDetailid", PrDetailId);

                                rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                string[] prFiles = PrFile.Split(',');
                                foreach (var file in prFiles)
                                {
                                    #region //更新SCM.PrFile
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO SCM.PrFile (PrId, PrDetailId, FileId
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.PrFileId
                                            VALUES (@PrId, @PrDetailid, @FileId
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            PrId,
                                            PrDetailId,
                                            FileId = file,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    var insertResult3 = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult3.Count();
                                    #endregion
                                }
                            }
                            else
                            {
                                #region //將原本的砍掉
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE FROM SCM.PrFile
                                        WHERE PrId = @PrId
                                        AND PrDetailId = @PrDetailId";
                                dynamicParameters.Add("PrId", PrId);
                                dynamicParameters.Add("PrDetailId", PrDetailId);

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

        #region //UpdatePrTransferBpm -- 拋轉請購單據至BPM -- Ann 2023-01-10
        public string UpdatePrTransferBpm(int PrId)
        {
            try
            {
                string token = "";
                int rowsAffected = 0;
                string CompanyNo = "";
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //取得USER資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserNo
                                FROM BAS.[User] a
                                WHERE a.UserId = @UserId";
                        dynamicParameters.Add("UserId", CreateBy);

                        var UserResult = sqlConnection.Query(sql, dynamicParameters);

                        string UserNo = "";
                        foreach (var item in UserResult)
                        {
                            UserNo = item.UserNo;
                        }
                        #endregion

                        #region //確認公司別DB
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.CompanyNo, a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var companyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (companyResult.Count() <= 0) throw new SystemException("公司別錯誤!");

                        foreach (var item in companyResult)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            CompanyNo = item.CompanyNo;
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //查詢ERP是否有此帳號
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(a.MF004)) USR_GROUP
                                    FROM ADMMF a
                                    WHERE MF001 = @MF001";
                            dynamicParameters.Add("MF001", UserNo);

                            var ADMMFResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (ADMMFResult.Count() <= 0) throw new SystemException("查無ERP帳號權限，無法進行拋轉及核單!");
                            #endregion

                            #region //判斷請購單資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.PrNo, a.PrStatus, FORMAT(a.DocDate, 'yyyy-MM-dd') DocDate, a.PrErpPrefix
                                    , FORMAT(a.PrDate, 'yyyy-MM-dd') PrDate, a.PrRemark, ISNULL(a.Amount, 0) Amount
                                    , a.UserId, a.Priority, a.BomType
                                    , b.DepartmentNo, b.DepartmentName
                                    FROM SCM.PurchaseRequisition a
                                    INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                                    WHERE a.PrId = @PrId";
                            dynamicParameters.Add("PrId", PrId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("請購單資料錯誤!");

                            string PrNo = "";
                            string DocDate = "";
                            string PrDate = "";
                            string PrRemark = "";
                            double Amount = -1;
                            string DepartmentNo = "";
                            string DepartmentName = "";
                            string PrErpPrefix = "";
                            string Priority = "";
                            string BomType = "";
                            foreach (var item in result)
                            {
                                if (item.PrStatus != "N" && item.PrStatus != "E" && item.PrStatus != "R") throw new SystemException("請購單狀態無法更改!");
                                if (item.UserId != CreateBy) throw new SystemException("無法修改其他請購人之單據!");
                                PrNo = item.PrNo;
                                DocDate = item.DocDate;
                                PrDate = item.PrDate;
                                PrRemark = item.PrRemark;
                                Amount = item.Amount;
                                DepartmentNo = item.DepartmentNo;
                                DepartmentName = item.DepartmentName;
                                PrErpPrefix = item.PrErpPrefix;
                                Priority = item.Priority;
                                BomType = item.BomType;
                            }
                            #endregion

                            #region //比對ERP關帳日期(庫存關帳)
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MA012)) MA012
                                    FROM CMSMA";
                            var cmsmaResult = sqlConnection2.Query(sql, dynamicParameters);
                            if (cmsmaResult.Count() < 0) throw new SystemException("EPR關帳日期資料錯誤!!");

                            foreach (var item in cmsmaResult)
                            {
                                string eprDate = item.MA012;
                                string erpYear = eprDate.Substring(0, 4);
                                string erpMonth = eprDate.Substring(4, 2);
                                string erpFullDate = erpYear + "-" + erpMonth;
                                DateTime erpDateTime = Convert.ToDateTime(erpFullDate);
                                DateTime DocDateDateTime = Convert.ToDateTime(DocDate.Split('-')[0] + "-" + DocDate.Split('-')[1]);
                                int compare = DocDateDateTime.CompareTo(erpDateTime);
                                if (compare <= 0) throw new SystemException("ERP已關帳(" + erpFullDate + ")，無法開立此日期(" + DocDate + ")之單據!!");
                            }
                            #endregion

                            #region //取得BPM TOKEN
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.IpAddress, a.CompanyId, a.Token, FORMAT(a.VerifyDate, 'yyyy-MM-dd HH:mm:ss') VerifyDate
                                    FROM BPM.SystemToken a
                                    WHERE IpAddress = @IpAddress
                                    AND CompanyId = @CompanyId";
                            dynamicParameters.Add("IpAddress", BpmServerPath);
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            var systemTokenResult = sqlConnection.Query(sql, dynamicParameters);
                            if (systemTokenResult.Count() <= 0) throw new SystemException("查無此憑證!");

                            foreach (var item in systemTokenResult)
                            {
                                DateTime verifyDate = Convert.ToDateTime(item.VerifyDate);
                                DateTime nowDate = DateTime.Now;
                                var CheckMin = (nowDate - verifyDate).TotalMinutes;
                                if (CheckMin >= 30)
                                {
                                    #region //取得新BPM TOKEN
                                    string tokenResponse = BpmHelper.GetBpmToken(BpmServerPath, BpmAccount, BpmPassword);
                                    var tokenJson = JObject.Parse(tokenResponse);
                                    foreach (var item2 in tokenJson)
                                    {
                                        if (item2.Key == "status")
                                        {
                                            if (item2.Value.ToString() != "success") throw new SystemException("取得token失敗!");
                                        }
                                        else if (item2.Key == "data")
                                        {
                                            token = item2.Value.ToString();
                                        }
                                    }
                                    #endregion

                                    #region //將新的TOKEN更新回MES
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE BPM.SystemToken SET
                                            Token = @Token,
                                            VerifyDate = @VerifyDate,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE IpAddress = @IpAddress
                                            AND CompanyId = @CompanyId";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            Token = token,
                                            VerifyDate = nowDate,
                                            LastModifiedDate,
                                            LastModifiedBy,
                                            IpAddress = BpmServerPath,
                                            CompanyId = CurrentCompany
                                        });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                                else
                                {
                                    token = item.Token;
                                }
                            }
                            #endregion

                            #region //取得BpmUser資料
                            string BpmUserId = "";
                            string BpmRoleId = "";
                            string BpmUserNo = "";
                            string UserName = "";
                            string BpmDepNo = "";
                            string BpmDepName = "";
                            using (SqlConnection sqlConnection3 = new SqlConnection(BpmDbConnectionStrings))
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"WITH BasicUserInfo(MemID, LoginID, UserName, MainRoleID, RolID, RolName, ParentRol) AS(
                                        SELECT a.MemID, a.LoginID, a.UserName, a.MainRoleID
                                        , b.RolID, b.Name, b.DepID AS ParentRol
                                        FROM Mem_GenInf a
                                        LEFT JOIN Rol_GenInf b ON b.RolID = a.MainRoleID
                                        WHERE a.LoginID = @LoginID
                                        UNION ALL
                                        SELECT a.MemID, a.LoginID, a.UserName, a.MainRoleID
                                        , b.RolID, b.Name, b.DepID AS ParentRol
                                        FROM BasicUserInfo a, Rol_GenInf b
                                        WHERE a.ParentRol = b.RolID
                                        )
                                        SELECT a.MemID, a.LoginID, a.UserName, a.MainRoleID
                                        , b.Name AS RolName
                                        , c.DepID, c.ID AS DepNo, c.Name AS DepName
                                        FROM BasicUserInfo a
                                        LEFT JOIN Rol_GenInf AS parentRol_GenInf ON a.RolID = parentRol_GenInf.RolID
                                        LEFT JOIN Dep_GenInf c ON parentRol_GenInf.DepID = c.DepID
                                        , Rol_GenInf b
                                        WHERE a.MainRoleID = b.RolID
                                        AND c.DepID IS NOT NULL
                                        UNION
                                        SELECT a.MemID, a.LoginID, a.UserName, a.MainRoleID
                                        , b.Name AS RolName
                                        , c.ComID AS DepID, c.ID AS DepNo, c.Name AS DepName
                                        FROM Mem_GenInf a
                                        LEFT JOIN Rol_GenInf b ON a.MainRoleID = b.RolID
                                        LEFT JOIN Company c ON b.DepID = c.ComID
                                        WHERE c.ComID IS NOT NULL
                                        AND a.LoginID = @LoginID
                                        ORDER BY a.LoginID";
                                dynamicParameters.Add("LoginID", UserNo);

                                var MemGenInfResult = sqlConnection3.Query(sql, dynamicParameters);

                                if (MemGenInfResult.Count() <= 0) throw new SystemException("取得BPM使用者資訊時發生錯誤!!");

                                foreach (var item in MemGenInfResult)
                                {
                                    BpmUserId = item.MemID;
                                    BpmRoleId = item.MainRoleID;
                                    BpmUserNo = item.LoginID;
                                    UserName = item.UserName;
                                    BpmDepNo = item.DepNo;
                                    BpmDepName = item.DepName;
                                }
                            }
                            #endregion

                            #region //取得單身資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.PrDetailId, a.PrSequence, a.PrMtlItemName, a.PrMtlItemSpec, a.PrQty
                                    , FORMAT(a.DemandDate, 'yyyy-MM-dd') DemandDate, a.PrCurrency, a.PrExchangeRate
                                    , a.PrUnitPrice, a.PrPrice, a.PrPriceTw, a.UrgentMtl, a.ProductionPlan, a.Project
                                    , a.PrDetailRemark, a.MtlInventory, a.MtlInventoryQty, a.SoDetailId
                                    , b.MtlItemNo
                                    , c.InventoryNo, c.InventoryName
                                    , d.UomNo
                                    , ISNULL(e.SupplierNo, '') SupplierNo, ISNULL(e.SupplierName, '') SupplierName, ISNULL(e.HideSupplier, 'N') HideSupplier
                                    , f.SoSequence
                                    , g.SoErpPrefix, g.SoErpNo
                                    FROM SCM.PrDetail a
                                    INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
                                    INNER JOIN SCM.Inventory c ON a.InventoryId = c.InventoryId
                                    INNER JOIN PDM.UnitOfMeasure d ON a.PrUomId = d.UomId
                                    LEFT JOIN SCM.Supplier e ON a.SupplierId = e.SupplierId
                                    LEFT JOIN SCM.SoDetail f ON a.SoDetailId = f.SoDetailId
                                    LEFT JOIN SCM.SaleOrder g ON f.SoId = g.SoId
                                    WHERE a.PrId = @PrId
                                    ORDER BY a.PrSequence";
                            dynamicParameters.Add("PrId", PrId);

                            var prDetailResult = sqlConnection.Query(sql, dynamicParameters);
                            if (prDetailResult.Count() <= 0) throw new SystemException("請購單身資料錯誤!");
                            #endregion

                            #region //組單身資料
                            JArray tabPR_Line_Data = new JArray();
                            string ProjectName = "";
                            double TotalAmountTw = 0;
                            string HideSupplier = prDetailResult.FirstOrDefault().HideSupplier;
                            double maxPrUnitPrice = 0;
                            foreach (var item in prDetailResult)
                            {
                                double localPrUnitPrice = item.PrUnitPrice * Convert.ToDouble(item.PrExchangeRate);

                                #region //取得最大單價
                                if (localPrUnitPrice > maxPrUnitPrice)
                                {
                                    maxPrUnitPrice = localPrUnitPrice;
                                }
                                #endregion

                                #region //檢核特定相關卡控(目前僅晶彩邏輯)
                                if (CurrentCompany == 4)
                                {
                                    #region //單別3109、3108特別卡控
                                    if (PrErpPrefix == "3108" || PrErpPrefix == "3109")
                                    {
                                        #region //請購單價不能超過1000(本幣)
                                        if (localPrUnitPrice > 1000) throw new SystemException("違反請購單別【3108、3109】卡控:<br>請購單價不能超過【1000】!!<br>此次請購單價【" + localPrUnitPrice + "】已超過!!");
                                        #endregion

                                        #region //所有單身合計金額(本幣)不能超過20000
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT ISNULL(SUM(PrPriceTw), 0) TotalPrice
                                                FROM SCM.PrDetail
                                                WHERE PrId = @PrId";
                                        dynamicParameters.Add("PrId", PrId);

                                        var TotalPriceResult = sqlConnection.Query(sql, dynamicParameters);

                                        double totalPrice = 0;
                                        foreach (var item2 in TotalPriceResult)
                                        {
                                            totalPrice = item2.TotalPrice;
                                            if (totalPrice > 20000) throw new SystemException("違反請購單別【3108、3109】卡控:<br>請購合計金額不能超過20000!!<br>目前請購合計金額【" + (totalPrice) + "】已超過!!");
                                        }
                                        #endregion
                                    }
                                    #endregion
                                }
                                #endregion

                                TotalAmountTw = TotalAmountTw + item.PrPriceTw;

                                #region //取得專案資料
                                if (item.Project != null)
                                {
                                    sql = @"SELECT LTRIM(RTRIM(NB001)) Project,LTRIM(RTRIM(NB002)) ProjectName
                                            FROM CMSNB
                                            WHERE NB001 = @Project";
                                    dynamicParameters.Add("Project", item.Project);
                                    var projectResult = sqlConnection2.Query(sql, dynamicParameters);
                                    foreach (var item2 in projectResult)
                                    {
                                        ProjectName = item2.ProjectName;
                                    }
                                }
                                #endregion

                                #region //取得ERP庫存資訊
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT LTRIM(RTRIM(a.MC007)) InventoryQty
                                        FROM INVMC a
                                        WHERE a.MC001 = @MC001
                                        AND a.MC002 = @MC002";
                                dynamicParameters.Add("MC001", item.MtlItemNo);
                                dynamicParameters.Add("MC002", item.InventoryNo);

                                var result5 = sqlConnection2.Query(sql, dynamicParameters);

                                double InventoryQty = 0;
                                string mtlInventory = "目前尚無資料";
                                if (result5.Count() > 0)
                                {
                                    foreach (var item2 in result5)
                                    {
                                        InventoryQty = Convert.ToDouble(item2.InventoryQty);
                                        #region //組MtlInventory
                                        List<MtlInventory> mtlInventories = new List<MtlInventory>
                                    {
                                        new MtlInventory
                                        {
                                            WAREHOUSE_NO = item.InventoryNo,
                                            WAREHOUSE_NAME = item.InventoryName,
                                            WAREHOUSE_QTY = InventoryQty
                                        }
                                    };
                                        #endregion

                                        mtlInventory = JsonConvert.SerializeObject(mtlInventories);
                                        mtlInventory = "{\"data\":" + mtlInventory + "}";
                                    }
                                }
                                #endregion

                                #region //Update SCM.PrDetail MtlInventory、MtlInventoryQty
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE SCM.PrDetail SET
                                        MtlInventory = @MtlInventory,
                                        MtlInventoryQty = @MtlInventoryQty,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE PrDetailId = @PrDetailId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        MtlInventory = mtlInventory,
                                        MtlInventoryQty = InventoryQty,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        item.PrDetailId
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                #region //處理MtlInventoryStatus
                                string mtlInventoryStatus = "目前尚無資料";
                                if (mtlInventory != "目前尚無資料")
                                {
                                    JObject jo = JObject.Parse(mtlInventory);
                                    JToken jtoken = jo["data"];

                                    for (int i = 0; i < jtoken.Count(); i++)
                                    {
                                        mtlInventoryStatus += jtoken[i]["WAREHOUSE_NO"] + ":" + jtoken[i]["WAREHOUSE_NAME"] + ":" + jtoken[i]["WAREHOUSE_QTY"] + ";";
                                    }
                                }
                                #endregion

                                #region //若有訂單，計算此訂單下物料相關資料
                                double soPrTotalAmount = 0;
                                double soTotalAmount = 0;
                                if (item.SoDetailId != null && item.SoDetailId > 0)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.SoId
                                            , b.MtlItemNo
                                            FROM SCM.SoDetail a
                                            INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
                                            WHERE a.SoDetailId = @SoDetailId";
                                    dynamicParameters.Add("SoDetailId", item.SoDetailId);

                                    var SoDetailResult = sqlConnection.Query(sql, dynamicParameters);

                                    int SoId = SoDetailResult.FirstOrDefault().SoId;

                                    #region //此訂單下所有物料有開請購單的金額合計(包括下層物料品號)
                                    foreach (var item2 in SoDetailResult)
                                    {
                                        string mtlItemNo = item2.MtlItemNo;

                                        #region //遞迴算出所有相關品號
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"WITH RecursiveCTE AS (
                                                    SELECT LTRIM(RTRIM(d.MB001)) MtlItemNo, LTRIM(RTRIM(d.MB002)) MtlItemName, LTRIM(RTRIM(d.MB003)) MtlItemSpec
                                                    FROM BOMMD a 
                                                    INNER JOIN BOMMC b ON b.MC001 = a.MD001
                                                    INNER JOIN INVMB c ON c.MB001 = b.MC001
                                                    INNER JOIN INVMB d ON d.MB001 = a.MD003
                                                    WHERE c.MB001 = @MtlItemNo
                                                    AND GETDATE() >= ISNULL(CONVERT(DATETIME, NULLIF(LTRIM(RTRIM(d.MB030)), ''), 120), GETDATE())
                                                    AND GETDATE() <= ISNULL(CONVERT(DATETIME, NULLIF(LTRIM(RTRIM(d.MB031)), ''), 120), GETDATE())
                                                    UNION ALL
                                                    SELECT LTRIM(RTRIM(xd.MB001)) MtlItemNo, LTRIM(RTRIM(xd.MB002)) MtlItemName, LTRIM(RTRIM(xd.MB003)) MtlItemSpec
                                                    FROM BOMMD x
                                                    INNER JOIN BOMMC xb ON xb.MC001 = x.MD001
                                                    INNER JOIN RecursiveCTE xa ON xb.MC001 = xa.MtlItemNo
                                                    INNER JOIN INVMB xc ON xc.MB001 = xb.MC001
                                                    INNER JOIN INVMB xd ON xd.MB001 = x.MD003
                                                    WHERE GETDATE() >= ISNULL(CONVERT(DATETIME, NULLIF(LTRIM(RTRIM(xd.MB030)), ''), 120), GETDATE())
                                                    AND GETDATE() <= ISNULL(CONVERT(DATETIME, NULLIF(LTRIM(RTRIM(xd.MB031)), ''), 120), GETDATE())
                                                )
                                                SELECT MtlItemNo, MtlItemName, MtlItemSpec
                                                , 2 AS SortOrder
                                                FROM RecursiveCTE
                                                UNION ALL
                                                SELECT LTRIM(RTRIM(x.MB001)) MtlItemNo, LTRIM(RTRIM(x.MB002)) MtlItemName, LTRIM(RTRIM(x.MB003)) MtlItemSpec
                                                , 1 AS SortOrder
                                                FROM INVMB x 
                                                WHERE x.MB001 = @MtlItemNo
                                                AND GETDATE() >= ISNULL(CONVERT(DATETIME, NULLIF(LTRIM(RTRIM(x.MB030)), ''), 120), GETDATE())
                                                AND GETDATE() <= ISNULL(CONVERT(DATETIME, NULLIF(LTRIM(RTRIM(x.MB031)), ''), 120), GETDATE())
                                                ORDER BY SortOrder";
                                        dynamicParameters.Add("MtlItemNo", mtlItemNo);

                                        var mtlItemResult = sqlConnection2.Query(sql, dynamicParameters);
                                        #endregion

                                        #region //根據此次遞迴結果，計算有註冊此訂單的請購單
                                        foreach (var item3 in mtlItemResult)
                                        {
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"SELECT ISNULL(SUM(a.PrPriceTw), 0) TotalAmount
                                                    FROM SCM.PrDetail a 
                                                    INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
                                                    WHERE a.SoDetailId = @SoDetailId
                                                    AND b.MtlItemNo = @MtlItemNo";
                                            dynamicParameters.Add("SoDetailId", item.SoDetailId);
                                            dynamicParameters.Add("MtlItemNo", item3.MtlItemNo);

                                            var prTotalAmountResult = sqlConnection.Query(sql, dynamicParameters);

                                            foreach (var item4 in prTotalAmountResult)
                                            {
                                                soPrTotalAmount += item4.TotalAmount;
                                            }
                                        }
                                        #endregion
                                    }
                                    #endregion

                                    #region //此訂單所有物料訂單金額合計
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT SUM(a.Amount) TotalAmount
                                            FROM SCM.SoDetail a 
                                            WHERE a.SoId = @SoId";
                                    dynamicParameters.Add("SoId", SoId);

                                    var soTotalAmountResult = sqlConnection.Query(sql, dynamicParameters);

                                    soTotalAmount = soTotalAmountResult.FirstOrDefault().TotalAmount;
                                    #endregion
                                }
                                #endregion

                                tabPR_Line_Data.Add(JObject.FromObject(new
                                {
                                    PRNumber_SN = item.PrSequence,
                                    ItemNo = item.MtlItemNo,
                                    ItemName = item.PrMtlItemName,
                                    Spec = item.PrMtlItemSpec,
                                    WareHouse = item.InventoryNo + "-" + item.InventoryName,
                                    Qty = item.PrQty.ToString(),
                                    Unit = item.UomNo,
                                    RequiredDate = item.DemandDate,
                                    Supplier = item.SupplierNo + "-" + item.SupplierName,
                                    Currency = item.PrCurrency,
                                    ExchangeRate = item.PrExchangeRate,
                                    Price = item.PrUnitPrice.ToString(),
                                    Amount = item.PrPrice.ToString(),
                                    AmountInLocalCurrency = item.PrPriceTw.ToString(),
                                    Ckb_UrgentMaterial = item.UrgentMtl,
                                    Ckb_MRP = item.ProductionPlan,
                                    Project = item.Project != null ? item.Project + " (" + ProjectName + ")" : "",
                                    LineRemark = item.PrDetailRemark,
                                    ReferenceDoc = item.SoErpPrefix != null ? item.SoErpPrefix + "-" + item.SoErpNo + "-" + item.SoSequence : "",
                                    MtlInventoryStatus = mtlInventoryStatus,
                                    item.MtlInventoryQty,
                                    SoPrTotalAmount = soPrTotalAmount.ToString(),
                                    SoTotalAmount = soTotalAmount.ToString(),
                                    MtlCostRatio = (soPrTotalAmount / soTotalAmount).ToString() + "%"
                                }));
                            }
                            #endregion

                            #region //組檔案JSON
                            JArray attachFilePath = new JArray();
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.FileId
                                    , b.FileName, b.FileExtension, b.FileSize, b.FileContent
                                    FROM SCM.PrFile a
                                    INNER JOIN BAS.[File] b ON a.FileId = b.FileId
                                    WHERE a.PrId = @PrId";
                            dynamicParameters.Add("PrId", PrId);

                            var prFileResult = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item in prFileResult)
                            {
                                attachFilePath.Add(JObject.FromObject(new
                                {
                                    fileId = item.FileId,
                                    fileName = item.FileName + item.FileExtension
                                }));
                            }
                            #endregion

                            #region //依公司別取得ProId
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.TypeName ProId
                                    FROM BAS.[Type] a
                                    WHERE a.TypeSchema = 'BPM.PrProId'
                                    AND a.TypeNo = @CompanyNo";
                            dynamicParameters.Add("CompanyNo", CompanyNo);

                            var ProIdResult = sqlConnection.Query(sql, dynamicParameters);

                            if (ProIdResult.Count() <= 0) throw new SystemException("此公司別【" + CompanyNo + "】尚未建立拋轉起單碼，請聯繫資訊人員!!");

                            string proId = "";
                            foreach (var item in ProIdResult)
                            {
                                proId = item.ProId;
                            }
                            #endregion

                            #region //組請購BPM資料
                            //string proId = "PRO03371647433833564"; 測試區原流程ID
                            //proId = "PRO00641720062815892";
                            string memId = BpmUserId;
                            string rolId = BpmRoleId;
                            string startMethod = "NoOpFirst";

                            JObject artInsAppData = JObject.FromObject(new
                            {
                                Title = "MES請購單 = " + PrNo + ", ERP請購單別: "+ PrErpPrefix,
                                PRNumber_MES = PrNo,
                                CreateDate = DocDate,
                                Requisitioner = BpmUserNo,
                                RequisitionerName = UserName,
                                RequisitionDeptID = DepartmentNo,
                                RequisitionDept = DepartmentName,
                                RequisitionDate = PrDate,
                                HeaderRemark = PrRemark,
                                AmountInLocalCurrencyTotal = TotalAmountTw.ToString(),
                                tabPR_Line_Data = JsonConvert.SerializeObject(tabPR_Line_Data),
                                attachFilePath = JsonConvert.SerializeObject(attachFilePath),
                                dbTable = "SCM.PurchaseRequisition",
                                mesID = PrId,
                                company = CompanyNo,
                                HideSupplier,
                                PRNumber_FormType = PrErpPrefix,
                                priority = Priority,
                                BomType,
                                MaxUnitPrice = maxPrUnitPrice
                            });
                            #endregion

                            string sData = BpmHelper.PostFormToBpm(token, proId, memId, rolId, startMethod, artInsAppData, BpmServerPath);

                            if (sData == "true")
                            {
                                #region //更改BPM拋轉狀態
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE SCM.PurchaseRequisition SET
                                        PrStatus = 'P',
                                        BpmTransferStatus = 'Y',
                                        BpmTransferUserId = @BpmTransferUserId,
                                        BpmTransferDate = @BpmTransferDate,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE PrId = @PrId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        BpmTransferUserId = CreateBy,
                                        BpmTransferDate = CreateDate,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        PrId
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }
                            else
                            {
                                #region //更改BPM拋轉狀態
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE SCM.PurchaseRequisition SET
                                        PrStatus = 'R',
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE PrId = @PrId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        PrId
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                throw new SystemException("請購單拋轉BPM失敗!");
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

        #region //UpdatePrTransferErp -- 拋轉請購單據至ERP -- Ann 2023-02-03
        public string UpdatePrTransferErp(int PrId, string BpmNo, string BpmStatus, string ComfirmUser, string ErpFolderRoot, string CompanyNo)
        {
            int rowsAffected = 0;
            string ErpDocPath = "";
            string PrErpNo = "";

            int currentNum = 0, yearLength = 0, lineLength = 0;
            string encode = "", paymentTerm = "", factory = "";
            DateTime referenceTime = default(DateTime);
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (CompanyNo.Length <= 0) throw new SystemException("公司別不能為空!!");

                        //CompanyNo = "JMO"; //待BPM可以傳公司別後再修正
                        #region //確認公司別DB
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb, a.SysDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        dynamicParameters.Add("CompanyNo", CompanyNo);

                        var companyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (companyResult.Count() <= 0) throw new SystemException("公司別錯誤!");

                        string ErpNo = "";
                        foreach (var item in companyResult)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            ErpSysDbConnectionStrings = ConfigurationManager.AppSettings[item.SysDb];
                            ErpNo = item.ErpNo;
                        }
                        #endregion

                        #region //確認核單者資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserId ConfirmUserId
                                FROM BAS.[User] a
                                WHERE a.UserNo = @ComfirmUser";
                        dynamicParameters.Add("ComfirmUser", ComfirmUser);

                        var confirmUserResult = sqlConnection.Query(sql, dynamicParameters);

                        if (confirmUserResult.Count() <= 0) throw new SystemException("核單者資料錯誤!");

                        int ConfirmUserId = -1;
                        foreach (var item in confirmUserResult)
                        {
                            ConfirmUserId = item.ConfirmUserId;
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            using (SqlConnection sqlConnection3 = new SqlConnection(ErpSysDbConnectionStrings))
                            {
                                #region //判斷請購單資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.BpmTransferStatus, a.ConfirmStatus, a.PrRemark, a.TotalQty, a.Amount
                                        , a.Edition, a.LockStaus, a.PrNo, a.TransferStatus, a.PrErpPrefix
                                        , FORMAT(a.PrDate, 'yyyyMMdd') PrDate
                                        , FORMAT(a.CreateDate, 'yyyyMMdd') CreateDate
                                        , FORMAT(a.CreateDate, 'HH:mm:ss') CreateTime
                                        , FORMAT(a.DocDate, 'yyyyMMdd') DocDate
                                        , FORMAT(a.DocDate, 'yyyy-MM-dd') MesDocDate
                                        , b.UserNo
                                        , c.DepartmentNo
                                        , d.UserNo PrUserNo
                                        , ISNULL(e.UserNo, '') ConfirmUserNo
                                        FROM SCM.PurchaseRequisition a
                                        INNER JOIN BAS.[User] b ON a.UserId = b.UserId
                                        INNER JOIN BAS.Department c ON a.DepartmentId = c.DepartmentId
                                        INNER JOIN BAS.[User] d ON a.UserId = d.UserId
                                        LEFT JOIN BAS.[User] e ON a.ConfirmUserId = e.UserId
                                        WHERE a.PrId = @PrId";
                                dynamicParameters.Add("PrId", PrId);

                                var result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("請購單資料錯誤!");

                                string UserNo = "";
                                string CreateDate = "";
                                string CreateTime = "";
                                string PrDate = "";
                                string DepartmentNo = "";
                                string PrRemark = "";
                                double TotalQty = -1;
                                string PrUserNo = "";
                                string DocDate = "";
                                string ConfirmUserNo = "";
                                double Amount = -1;
                                string Edition = "";
                                string LockStaus = "";
                                string PrNo = "";
                                string PrErpPrefix = "";
                                string MesDocDate = "";
                                foreach (var item in result)
                                {
                                    if (item.BpmTransferStatus != "Y") throw new SystemException("請購單尚未拋轉BPM，無法拋轉ERP!");
                                    if (item.TransferStatus != "N") throw new SystemException("請購單已拋轉ERP，無法重複拋轉!");
                                    UserNo = item.UserNo;
                                    CreateDate = item.CreateDate;
                                    CreateTime = item.CreateTime;
                                    PrDate = item.PrDate;
                                    DepartmentNo = item.DepartmentNo;
                                    PrRemark = item.PrRemark;
                                    TotalQty = item.TotalQty;
                                    PrUserNo = item.PrUserNo;
                                    DocDate = item.DocDate;
                                    ConfirmUserNo = item.ConfirmUserNo;
                                    Amount = item.Amount;
                                    Edition = item.Edition;
                                    LockStaus = item.LockStaus;
                                    PrNo = item.PrNo;
                                    PrErpPrefix = item.PrErpPrefix;
                                    MesDocDate = item.MesDocDate;
                                }
                                #endregion

                                #region //取得單身資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.PrSequence, a.PrQty, a.PrDetailRemark, ISNULL(a.PoQty, 0) PoQty
                                    , ISNULL(a.PoCurrency, '') PoCurrency, ISNULL(a.PoUnitPrice, 0) PoUnitPrice
                                    , ISNULL(a.PoPrice, 0) PoPrice, a.LockStaus, a.PoStaus, ISNULL(a.PoErpPrefixNo, '') PoErpPrefixNo
                                    , ISNULL(a.PoRemark, '') PoRemark, a.Taxation, a.UrgentMtl, a.ClosureStatus
                                    , a.InquiryStatus, ISNULL(a.Project, '') Project, a.PrExchangeRate, a.PrPriceTw
                                    , a.PartialPurchaseStaus, ISNULL(a.BudgetDepartmentNo, '') BudgetDepartmentNo
                                    , ISNULL(a.BudgetDepartmentSubject, '') BudgetDepartmentSubject, a.PrUnitPrice
                                    , a.PrCurrency, a.PrPrice, a.TaxNo, a.TradeTerm, a.BusinessTaxRate, a.DetailMultiTax
                                    , ISNULL(a.DiscountRate, 0) DiscountRate, ISNULL(a.DiscountAmount, 0) DiscountAmount
                                    , FORMAT(a.DemandDate, 'yyyyMMdd') DemandDate, ISNULL(FORMAT(a.DeliveryDate, 'yyyyMMdd'), '') DeliveryDate
                                    , a.PrMtlItemName, a.PrMtlItemSpec
                                    , b.MtlItemNo
                                    , c.InventoryNo
                                    , d.UomNo
                                    , ISNULL(e.SupplierNo, '') SupplierNo
                                    , f.SoSequence
                                    , g.SoErpPrefix, g.SoErpNo
                                    , ISNULL(h.UserNo, '') PoUserNo
                                    , ISNULL(i.UomNo, '') PoUomNo
                                    FROM SCM.PrDetail a
                                    INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
                                    INNER JOIN SCM.Inventory c ON a.InventoryId = c.InventoryId
                                    INNER JOIN PDM.UnitOfMeasure d ON a.PrUomId = d.UomId
                                    LEFT JOIN SCM.Supplier e ON a.SupplierId = e.SupplierId
                                    LEFT JOIN SCM.SoDetail f ON a.SoDetailId = f.SoDetailId
                                    LEFT JOIN SCM.SaleOrder g ON f.SoId = g.SoId
                                    LEFT JOIN BAS.[User] h ON a.PoUserId = h.UserId
                                    LEFT JOIN PDM.UnitOfMeasure i ON a.PoUomId = i.UomId
                                    WHERE a.PrId = @PrId";
                                dynamicParameters.Add("PrId", PrId);

                                var prDetailResult = sqlConnection.Query(sql, dynamicParameters);
                                if (prDetailResult.Count() <= 0) throw new SystemException("請購單身資料錯誤!");
                                #endregion

                                #region //查詢廠別
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT LTRIM(RTRIM(MB001)) MB001 FROM CMSMB";
                                var CMSMBResult = sqlConnection2.Query(sql, dynamicParameters);

                                if (CMSMBResult.Count() > 1) throw new SystemException("廠別數有多個，請與資訊人員確認!!");

                                string TA010 = "";
                                foreach (var item in CMSMBResult)
                                {
                                    TA010 = item.MB001;
                                }
                                #endregion

                                #region //審核ERP權限並取得USR_GROUP
                                //string USR_GROUP = "ZY07";
                                //string USR_GROUP = BaseHelper.CheckErpAuthority(UserNo, sqlConnection2, "PURI05", "CREATE");
                                #endregion

                                #region //取得USR_GROUP
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT LTRIM(RTRIM(a.MF004)) USR_GROUP
                                        FROM ADMMF a
                                        WHERE MF001 = @MF001";
                                dynamicParameters.Add("MF001", UserNo);

                                var ADMMFResult = sqlConnection2.Query(sql, dynamicParameters);

                                if (ADMMFResult.Count() <= 0) throw new SystemException("查無ERP帳號權限，無法進行拋轉及核單!");

                                string USR_GROUP = "";
                                foreach (var item in result)
                                {
                                    USR_GROUP = item.USR_GROUP;
                                }
                                #endregion

                                #region //管理欄位定義
                                string COMPANY = ErpNo;
                                string CREATOR = UserNo;
                                string CREATE_DATE = CreateDate;
                                string MODIFIER = ComfirmUser;
                                string MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
                                string FLAG = "1";
                                string CREATE_TIME = CreateTime;
                                string CREATE_AP = "MES";
                                string CREATE_PRID = "ERPJ01";
                                string MODI_TIME = DateTime.Now.ToString("HH:mm:ss");
                                string MODI_AP = "MES";
                                string MODI_PRID = "ERPJ01";
                                #endregion

                                #region //INSERT PURTA 請購單單頭

                                #region //依請購日期自動取號
                                referenceTime = DateTime.ParseExact(DocDate, "yyyyMMdd", CultureInfo.InvariantCulture);
                                #region //取得單據設定
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.MQ004, a.MQ005, a.MQ006, a.MQ044
                                        FROM CMSMQ a
                                        WHERE a.COMPANY = @CompanyNo
                                        AND a.MQ001 = @ErpPrefix";
                                dynamicParameters.Add("CompanyNo", ErpNo);
                                dynamicParameters.Add("ErpPrefix", PrErpPrefix);

                                var resultDocSetting = sqlConnection2.Query(sql, dynamicParameters);
                                if (resultDocSetting.Count() <= 0) throw new SystemException("ERP單據設定資料不存在!");

                                foreach (var item in resultDocSetting)
                                {
                                    encode = item.MQ004; //編碼方式
                                    yearLength = Convert.ToInt32(item.MQ005); //年碼數
                                    lineLength = Convert.ToInt32(item.MQ006); //流水號碼數
                                }
                                #endregion

                                #region //取號
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT CONVERT(int, RIGHT(ISNULL(MAX(RTRIM(LTRIM(TA002))), '" + new string('0', lineLength) + @"'), " + lineLength + @")) + 1 CurrentNum
                                        FROM PURTA
                                        WHERE TA001 = @ErpPrefix";
                                dynamicParameters.Add("ErpPrefix", PrErpPrefix);

                                #region //編碼方式
                                string dateFormat = "";
                                switch (encode)
                                {
                                    case "1": //日編
                                        dateFormat = new string('y', yearLength) + "MMdd";
                                        sql += @" AND RTRIM(LTRIM(TA002)) LIKE @ReferenceTime + '" + new string('_', lineLength) + @"'";
                                        dynamicParameters.Add("ReferenceTime", referenceTime.ToString(dateFormat));
                                        PrErpNo = referenceTime.ToString(dateFormat);
                                        break;
                                    case "2": //月編
                                        dateFormat = new string('y', yearLength) + "MM";
                                        sql += @" AND RTRIM(LTRIM(TA002)) LIKE @ReferenceTime + '" + new string('_', lineLength) + @"'";
                                        dynamicParameters.Add("ReferenceTime", referenceTime.ToString(dateFormat));
                                        PrErpNo = referenceTime.ToString(dateFormat);
                                        break;
                                    case "3": //流水號
                                        break;
                                    case "4": //手動編號
                                        break;
                                    default:
                                        throw new SystemException("編碼方式錯誤!");
                                }
                                #endregion

                                currentNum = sqlConnection2.QueryFirstOrDefault(sql, dynamicParameters).CurrentNum;
                                PrErpNo += string.Format("{0:" + new string('0', lineLength) + "}", currentNum);
                                #endregion
                                #endregion

                                #region //檢查單據是否已經存在
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                        FROM PURTA a
                                        WHERE a.TA001 = @PrErpPrefix
                                        AND a.TA002 = @TA002";
                                dynamicParameters.Add("TA002", PrErpNo);
                                dynamicParameters.Add("PrErpPrefix", PrErpPrefix);

                                var checkPrSeqResult = sqlConnection2.Query(sql, dynamicParameters);

                                if (checkPrSeqResult.Count() > 0) throw new SystemException("請購單號重覆!");
                                #endregion

                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO PURTA (COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE
                                        , FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID
                                        , TA001, TA002, TA003, TA004, TA005, TA006, TA007, TA008, TA009, TA010
                                        , TA011, TA012, TA013, TA014, TA015, TA016, TA017, TA018, TA019, TA020, TA021
                                        , TA022, TA023, TA024, TA025, TA026, TA027, TA028, TA550, TA551, UDF01, UDF02)
                                        VALUES (@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE
                                        , @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID, @MODI_TIME, @MODI_AP, @MODI_PRID
                                        , @TA001, @TA002, @TA003, @TA004, @TA005, @TA006, @TA007, @TA008, @TA009, @TA010
                                        , @TA011, @TA012, @TA013, @TA014, @TA015, @TA016, @TA017, @TA018, @TA019, @TA020, @TA021
                                        , @TA022, @TA023, @TA024, @TA025, @TA026, @TA027, @TA028, @TA550, @TA551, @UDF01, @UDF02)";

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
                                      TA001 = PrErpPrefix,
                                      TA002 = PrErpNo,
                                      TA003 = DateTime.Now.ToString("yyyyMMdd"),
                                      TA004 = DepartmentNo,
                                      TA005 = "",
                                      TA006 = PrRemark,
                                      TA007 = BpmStatus,
                                      TA008 = 0,
                                      TA009 = "9",
                                      TA010,
                                      TA011 = TotalQty,
                                      TA012 = PrUserNo,
                                      TA013 = DocDate,
                                      TA014 = ComfirmUser,
                                      TA015 = 0,
                                      TA016 = "3",
                                      TA017 = 0,
                                      TA018 = "",
                                      TA019 = "",
                                      TA020 = Amount,
                                      TA021 = Edition,
                                      TA022 = "",
                                      TA023 = 0,
                                      TA024 = 0,
                                      TA025 = "",
                                      TA026 = "",
                                      TA027 = "",
                                      TA028 = LockStaus,
                                      TA550 = "",
                                      TA551 = 0,
                                      UDF01 = PrNo,
                                      UDF02 = BpmNo
                                  });
                                var insertResult = sqlConnection2.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();

                                #endregion

                                foreach (var item in prDetailResult)
                                {
                                    #region //INSERT PURTB 請購單單身
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO PURTB (COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE
                                        , FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID
                                        , TB001, TB002, TB003, TB004, TB005, TB006, TB007, TB008, TB009, TB010, TB011
                                        , TB012, TB013, TB014, TB015, TB016, TB017, TB018, TB019, TB020, TB021, TB022, TB023
                                        , TB024, TB025, TB026, TB027, TB028, TB029, TB030, TB031, TB032, TB033, TB034, TB035
                                        , TB036, TB037, TB038, TB039, TB040, TB041, TB042, TB043, TB044, TB045, TB046, TB047
                                        , TB048, TB049, TB050, TB051, TB052, TB053, TB054, TB055, TB056, TB057, TB058, TB059
                                        , TB060, TB061, TB062, TB063, TB064, TB065, TB066, TB067, TB068, TB200, TB500, TB501
                                        , TB502, TB503, TB550, TB551)
                                        VALUES (@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE
                                        , @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID, @MODI_TIME, @MODI_AP, @MODI_PRID
                                        , @TB001, @TB002, @TB003, @TB004, @TB005, @TB006, @TB007, @TB008, @TB009, @TB010, @TB011
                                        , @TB012, @TB013, @TB014, @TB015, @TB016, @TB017, @TB018, @TB019, @TB020, @TB021, @TB022, @TB023
                                        , @TB024, @TB025, @TB026, @TB027, @TB028, @TB029, @TB030, @TB031, @TB032, @TB033, @TB034, @TB035
                                        , @TB036, @TB037, @TB038, @TB039, @TB040, @TB041, @TB042, @TB043, @TB044, @TB045, @TB046, @TB047
                                        , @TB048, @TB049, @TB050, @TB051, @TB052, @TB053, @TB054, @TB055, @TB056, @TB057, @TB058, @TB059
                                        , @TB060, @TB061, @TB062, @TB063, @TB064, @TB065, @TB066, @TB067, @TB068, @TB200, @TB500, @TB501
                                        , @TB502, @TB503, @TB550, @TB551)";

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
                                          TB001 = PrErpPrefix,
                                          TB002 = PrErpNo,
                                          TB003 = item.PrSequence,
                                          TB004 = item.MtlItemNo,
                                          TB005 = item.PrMtlItemName,
                                          TB006 = item.PrMtlItemSpec,
                                          TB007 = item.UomNo,
                                          TB008 = item.InventoryNo,
                                          TB009 = item.PrQty,
                                          TB010 = item.SupplierNo,
                                          TB011 = item.DemandDate,
                                          TB012 = item.PrDetailRemark,
                                          TB013 = item.PoUserNo,
                                          TB014 = item.PoQty,
                                          TB015 = item.PoUomNo,
                                          TB016 = item.PoCurrency,
                                          TB017 = item.PoUnitPrice,
                                          TB018 = item.PoPrice,
                                          TB019 = item.DeliveryDate,
                                          TB020 = item.LockStaus,
                                          TB021 = item.PoStaus,
                                          TB022 = item.PoErpPrefixNo,
                                          TB023 = "",
                                          TB024 = item.PoRemark,
                                          TB025 = BpmStatus,
                                          TB026 = item.Taxation,
                                          TB027 = "",
                                          TB028 = "",
                                          TB029 = item.SoErpPrefix != null ? item.SoErpPrefix : "",
                                          TB030 = item.SoErpNo != null ? item.SoErpNo : "",
                                          TB031 = item.SoSequence != null ? item.SoSequence : "",
                                          TB032 = item.UrgentMtl,
                                          TB033 = "",
                                          TB034 = 0,
                                          TB035 = 0,
                                          TB036 = "",
                                          TB037 = "",
                                          TB038 = "",
                                          TB039 = item.ClosureStatus,
                                          TB040 = item.InquiryStatus,
                                          TB041 = "",
                                          TB042 = "2",
                                          TB043 = item.Project,
                                          TB044 = item.PrExchangeRate,
                                          TB045 = item.PrPriceTw,
                                          TB046 = item.PartialPurchaseStaus,
                                          TB047 = item.BudgetDepartmentNo,
                                          TB048 = item.BudgetDepartmentSubject,
                                          TB049 = item.PrUnitPrice,
                                          TB050 = item.PrCurrency,
                                          TB051 = item.PrPrice,
                                          TB052 = 0,
                                          TB053 = 0,
                                          TB054 = "",
                                          TB055 = "",
                                          TB056 = "",
                                          TB057 = item.TaxNo,
                                          TB058 = item.TradeTerm,
                                          TB059 = "",
                                          TB060 = "",
                                          TB061 = "",
                                          TB062 = "",
                                          TB063 = item.BusinessTaxRate,
                                          TB064 = item.DetailMultiTax,
                                          TB065 = item.PrQty,
                                          TB066 = item.UomNo,
                                          TB067 = item.DiscountRate,
                                          TB068 = item.DiscountAmount,
                                          TB200 = "",
                                          TB500 = "",
                                          TB501 = "",
                                          TB502 = "",
                                          TB503 = "",
                                          TB550 = "",
                                          TB551 = 0
                                      });
                                    var insertResult2 = sqlConnection2.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult2.Count();
                                    #endregion

                                    #region //INSERT INVMG(NetChange記錄檔)

                                    #region //先查詢是否已經有此請購單資料，若無INSERT，有則UPDATE
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT MG003
                                            FROM INVMG
                                            WHERE MG001 = @MG001
                                            AND MG002 = @MG002";
                                    dynamicParameters.Add("MG001", item.MtlItemNo);
                                    dynamicParameters.Add("MG002", TA010);

                                    var INVMgResult = sqlConnection2.Query(sql, dynamicParameters);

                                    if (INVMgResult.Count() > 0)
                                    {
                                        #region //UPDATE INVMG
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"UPDATE INVMG SET
                                                MODIFIER = @MODIFIER,
                                                MODI_DATE = @MODI_DATE,
                                                FLAG = @FLAG,
                                                MODI_TIME = @MODI_TIME,
                                                MODI_AP = @MODI_AP,
                                                MG003 = @MG003
                                                WHERE MG001 = @MG001
                                                AND MG002 = @MG002";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                MODIFIER,
                                                MODI_DATE,
                                                FLAG,
                                                MODI_TIME,
                                                MODI_AP,
                                                MG003 = "Y",
                                                MG001 = item.MtlItemNo,
                                                MG002 = TA010
                                            });

                                        rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);
                                        #endregion
                                    }
                                    else
                                    {
                                        #region //INSERT INVMG
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO INVMG (COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE
                                                , FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID
                                                , MG001, MG002, MG003, MG004, MG005, MG006, MG007, MG008, MG009, MG010)
                                                VALUES (@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE
                                                , @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID, @MODI_TIME, @MODI_AP, @MODI_PRID
                                                , @MG001, @MG002, @MG003, @MG004, @MG005, @MG006, @MG007, @MG008, @MG009, @MG010)";

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
                                              MG001 = item.MtlItemNo,
                                              MG002 = TA010,
                                              MG003 = "Y",
                                              MG004 = "",
                                              MG005 = "",
                                              MG006 = 0,
                                              MG007 = 0,
                                              MG008 = "",
                                              MG009 = "",
                                              MG010 = ""
                                          });
                                        var insertResult3 = sqlConnection2.Query(sql, dynamicParameters);

                                        rowsAffected += insertResult3.Count();
                                        #endregion
                                    }
                                    #endregion

                                    #endregion

                                    #region //INSERT PURTY(請購單子單身)
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO PURTY (COMPANY, CREATOR, USR_GROUP, CREATE_DATE
                                            , FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID
                                            , TY001, TY002, TY003, TY004, TY005, TY006, TY007, TY008
                                            , TY009, TY010, TY011, TY012, TY013, TY014, TY015, TY016
                                            , TY017, TY018, TY019, TY020, TY021, TY022, TY023, TY024
                                            , TY025, TY026, TY027, TY028, TY029, TY030, TY031, TY032
                                            , TY033, TY034, TY035, TY036, TY037, TY038, TY039, TY040
                                            , TY041, TY042, TY043)
                                            VALUES (@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE
                                            , @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID
                                            , @TY001, @TY002, @TY003, @TY004, @TY005, @TY006, @TY007, @TY008
                                            , @TY009, @TY010, @TY011, @TY012, @TY013, @TY014, @TY015, @TY016
                                            , @TY017, @TY018, @TY019, @TY020, @TY021, @TY022, @TY023, @TY024
                                            , @TY025, @TY026, @TY027, @TY028, @TY029, @TY030, @TY031, @TY032
                                            , @TY033, @TY034, @TY035, @TY036, @TY037, @TY038, @TY039, @TY040
                                            , @TY041, @TY042, @TY043)";

                                    dynamicParameters.AddDynamicParams(
                                      new
                                      {
                                          COMPANY,
                                          CREATOR,
                                          USR_GROUP,
                                          CREATE_DATE,
                                          FLAG,
                                          CREATE_TIME,
                                          CREATE_AP,
                                          CREATE_PRID,
                                          TY001 = PrErpPrefix,
                                          TY002 = PrErpNo,
                                          TY003 = item.PrSequence,
                                          TY004 = "0001",
                                          TY005 = item.InventoryNo,
                                          TY006 = item.SupplierNo,
                                          TY007 = item.PoQty,
                                          TY008 = item.PoCurrency,
                                          TY009 = item.PoUnitPrice,
                                          TY010 = item.PoPrice,
                                          TY011 = item.DeliveryDate,
                                          TY012 = item.LockStaus,
                                          TY013 = item.PoErpPrefixNo,
                                          TY014 = "",
                                          TY015 = item.PoRemark,
                                          TY016 = item.PoUserNo,
                                          TY017 = item.Taxation,
                                          TY018 = item.UrgentMtl,
                                          TY019 = "",
                                          TY020 = 0,
                                          TY021 = item.ClosureStatus,
                                          TY022 = item.PoStaus,
                                          TY023 = "",
                                          TY024 = "",
                                          TY025 = 0,
                                          TY026 = 0,
                                          TY027 = 0,
                                          TY028 = "",
                                          TY029 = "",
                                          TY030 = "",
                                          TY031 = item.TaxNo,
                                          TY032 = item.TradeTerm,
                                          TY033 = "",
                                          TY034 = "",
                                          TY035 = "",
                                          TY036 = "",
                                          TY037 = "",
                                          TY038 = item.PrPrice,
                                          TY039 = item.BusinessTaxRate,
                                          TY040 = item.DetailMultiTax,
                                          TY041 = item.PrPriceQty,
                                          TY042 = item.DiscountRate,
                                          TY043 = item.DiscountAmount
                                      });
                                    var insertResult4 = sqlConnection2.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult4.Count();
                                    #endregion
                                }

                                #region //請購單附檔
                                string DocID = "";
                                string SeqNo = "";
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.PrFileId, a.PrId, a.FileId
                                        , b.[FileName], b.FileExtension, b.FileContent
                                        , c.UserNo
                                        FROM SCM.PrFile a
                                        INNER JOIN BAS.[File] b ON a.FileId = b.FileId
                                        INNER JOIN BAS.[User] c ON a.CreateBy = c.UserId
                                        WHERE a.PrId = @PrId";
                                dynamicParameters.Add("PrId", PrId);

                                var prFileResult = sqlConnection.Query(sql, dynamicParameters);

                                if (prFileResult.Count() > 0)
                                {
                                    foreach (var item in prFileResult)
                                    {
                                        #region //DMS

                                        #region //DMS取號
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT '001'+ RIGHT(REPLICATE('0', 7) + CONVERT(VARCHAR(7)
                                                , ISNULL(MAX(SUBSTRING(a.DocID, LEN(a.DocID) - 6, 7)), 0) + 1), 7) AS nextSN
                                                FROM dbo.DMS a
                                                --FROM BPM.DMS a
                                                WHERE a.DocID LIKE '001%'";

                                        var dmsResult = sqlConnection3.Query(sql, dynamicParameters);

                                        foreach (var item2 in dmsResult)
                                        {
                                            DocID = item2.nextSN;
                                        }
                                        #endregion

                                        #region //INSERT DMS
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO dbo.DMS (DocID, Revision, Owner, Status, AddDate, AddTime
                                                , DocName, DocType, Description)
                                                VALUES (@DocID, @Revision, @Owner, @Status, @AddDate, @AddTime
                                                , @DocName, @DocType, @Description)";

                                        dynamicParameters.AddDynamicParams(
                                          new
                                          {
                                              DocID,
                                              Revision = "001",
                                              Owner = item.UserNo,
                                              Status = "N",
                                              AddDate = DateTime.Now.ToString("yyyyMMdd"),
                                              AddTime = DateTime.Now.ToString("HH:mm"),
                                              DocName = item.FileName + item.FileExtension,
                                              DocType = item.FileExtension,
                                              Description = item.FileName + item.FileExtension
                                          });
                                        var insertResult2 = sqlConnection3.Query(sql, dynamicParameters);

                                        rowsAffected += insertResult2.Count();
                                        #endregion

                                        #endregion

                                        #region //ATTACH

                                        #region //取號
                                        sql = @"SELECT RIGHT(REPLICATE('0', 3) + CONVERT(VARCHAR(3)
                                                , ISNULL(MAX(SUBSTRING(SeqNo, LEN(SeqNo) - 2, 3)), 0) + 1), 3) AS nextSN
                                                FROM dbo.ATTACH a
                                                --FROM BPM.ATTACH a
                                                WHERE a.CompanyID = @CompanyID
                                                AND a.UserID = @UserID
                                                AND a.Parent = @Parent
                                                AND a.KeyValues = @KeyValues";
                                        dynamicParameters.Add("CompanyID", ErpNo);
                                        dynamicParameters.Add("UserID", item.UserNo);
                                        dynamicParameters.Add("Parent", "PURI05");
                                        dynamicParameters.Add("KeyValues", "" + PrErpPrefix + "||" + PrErpNo);

                                        var attachResult = sqlConnection3.Query(sql, dynamicParameters);

                                        foreach (var item2 in attachResult)
                                        {
                                            SeqNo = item2.nextSN;
                                        }
                                        #endregion

                                        #region //INSERT ATTACH
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO dbo.ATTACH (Parent, KeyValues, CompanyID, UserID, Type, SeqNo
                                                , FileName, DocID, Revision, AddDate, AddTime, KeyFields)
                                                VALUES (@Parent, @KeyValues, @CompanyID, @UserID, @Type, @SeqNo
                                                , @FileName, @DocID, @Revision, @AddDate, @AddTime, @KeyFields)";
                                        dynamicParameters.AddDynamicParams(
                                          new
                                          {
                                              Parent = "PURI05",
                                              KeyValues = "" + PrErpPrefix + "||" + PrErpNo,
                                              CompanyID = ErpNo,
                                              UserID = item.UserNo,
                                              Type = "1",
                                              SeqNo,
                                              FileName = item.FileName + item.FileExtension,
                                              DocID,
                                              Revision = "001",
                                              AddDate = DateTime.Now.ToString("yyyyMMdd"),
                                              AddTime = DateTime.Now.ToString("HH:mm"),
                                              KeyFields = "TA001||TA002"
                                          });
                                        var insertResult3 = sqlConnection3.Query(sql, dynamicParameters);

                                        rowsAffected += insertResult3.Count();
                                        #endregion

                                        #endregion

                                        #region //取得ERP檔案路徑
                                        double no = Convert.ToDouble(DocID.Substring(3));
                                        int folderNo = Convert.ToInt32(Math.Floor(no / 1000)) * 1000;
                                        string folderName = folderNo.ToString().PadLeft(7, '0');
                                        string targetErpFolderPath = Path.Combine(ErpFolderRoot, folderName);
                                        ErpDocPath = Path.Combine(targetErpFolderPath, "001" + DocID.Substring(3) + ".001");
                                        #endregion

                                        #region //COPY附檔至ERP路徑資料夾
                                        if (!Directory.Exists(targetErpFolderPath)) { Directory.CreateDirectory(targetErpFolderPath); }
                                        byte[] fileContent = (byte[])item.FileContent;
                                        File.WriteAllBytes(ErpDocPath, fileContent); // Requires System.IO
                                        #endregion

                                        #region //更新MES請購單附檔資訊
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"UPDATE SCM.PrFile SET
                                                ErpDocId = @ErpDocId,
                                                ErpDocPath = @ErpDocPath,
                                                ErpDocDate = @ErpDocDate,
                                                LastModifiedDate = @LastModifiedDate,
                                                LastModifiedBy = @LastModifiedBy
                                                WHERE PrFileId = @PrFileId";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                ErpDocId = DocID,
                                                ErpDocPath,
                                                ErpDocDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                                                LastModifiedDate,
                                                LastModifiedBy,
                                                item.PrFileId
                                            });

                                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                        #endregion
                                    }
                                }
                                #endregion

                                #region //將ERP請購單資訊回填MES
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE SCM.PurchaseRequisition SET
                                        BpmNo = @BpmNo,
                                        PrErpNo = @PrErpNo,
                                        TransferStatus = 'Y',
                                        TransferDate = @TransferDate,
                                        ConfirmStatus = 'Y',
                                        ConfirmUserId = @ConfirmUserId,
                                        PrDate = @LastModifiedDate,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE PrId = @PrId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        BpmNo,
                                        PrErpNo,
                                        TransferDate = LastModifiedDate,
                                        ConfirmUserId,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        PrId
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE SCM.PrDetail SET
                                        ConfirmStatus = 'Y',
                                        ConfirmUserId = @ConfirmUserId,
                                        DemandDate = @LastModifiedDate,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE PrId = @PrId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        ConfirmUserId,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        PrId
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

        #region //UpdatePrStatus -- 更新請購單狀態 -- Ann 2023-02-06
        public string UpdatePrStatus(int PrId, string BpmNo, string Status, string RootId, string ConfirmUser, string ErpFlag)
        {
            try
            {
                int rowsAffected = 0;
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷請購單資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.BpmTransferDate, a.BpmTransferStatus, a.SignupStaus
                                , a.PrErpPrefix + '-' + a.PrErpNo PrErpFullNo
                                , a.UserId
                                FROM SCM.PurchaseRequisition a
                                WHERE a.PrId = @PrId";
                        dynamicParameters.Add("PrId", PrId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("請購單資料錯誤!");

                        DateTime BpmTransferDate = new DateTime();
                        string BpmTransferStatus = "";
                        string SignupStaus = "";
                        string PrErpFullNo = "";
                        int UserId = -1;
                        foreach (var item in result)
                        {
                            BpmTransferDate = item.BpmTransferDate;
                            BpmTransferStatus = item.BpmTransferStatus;
                            SignupStaus = item.SignupStaus;
                            PrErpFullNo = item.PrErpFullNo;
                            UserId = item.UserId;
                        }
                        #endregion

                        #region //UPDATE SCM.PurchaseRequisition
                        if (RootId.Length > 0) //BPM結束流程(E、Y)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.PurchaseRequisition SET
                                    PrStatus = @PrStatus,
                                    SignupStaus = @SignupStaus,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @UserId
                                    WHERE PrId = @PrId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    PrStatus = ErpFlag != "F" ? Status : "F",
                                    SignupStaus = Status == "P" ? "3" : SignupStaus,
                                    LastModifiedDate,
                                    UserId,
                                    PrId
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        }
                        else //BPM流程開始(P)、拋轉失敗(R)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.PurchaseRequisition SET
                                    PrStatus = @PrStatus,
                                    BpmTransferStatus = @BpmTransferStatus,
                                    BpmTransferDate = @BpmTransferDate,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @UserId
                                    WHERE PrId = @PrId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    PrStatus = Status,
                                    BpmTransferStatus = Status == "P" ? "Y" : BpmTransferStatus,
                                    BpmTransferDate = DateTime.Now.ToString("yyyy-MM-dd"),
                                    LastModifiedDate,
                                    UserId,
                                    PrId
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        }
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            msg = "(" + rowsAffected + " rows affected)",
                            data = PrErpFullNo
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

        #region //UpdatePrModification -- 更新請購變更單資料 -- Ann 2023-02-08
        public string UpdatePrModification(int PrmId, int PrId, int UserId, int DepartmentId, string DocDate, string ModiReason, string PrmRemark, string PrmFile, string Priority)
        {
            try
            {
                if (DocDate.Length < 0) throw new SystemException("【請購單據日期】不能為空!");

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
                            #region //比對ERP關帳日期(庫存關帳)
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MA012)) MA012
                                    FROM CMSMA";
                            var cmsmaResult = sqlConnection2.Query(sql, dynamicParameters);
                            if (cmsmaResult.Count() < 0) throw new SystemException("EPR關帳日期資料錯誤!!");

                            foreach (var item in cmsmaResult)
                            {
                                string eprDate = item.MA012;
                                string erpYear = eprDate.Substring(0, 4);
                                string erpMonth = eprDate.Substring(4, 2);
                                string erpFullDate = erpYear + "-" + erpMonth;
                                DateTime erpDateTime = Convert.ToDateTime(erpFullDate);
                                DateTime DocDateDateTime = Convert.ToDateTime(DocDate.Split('-')[0] + "-" + DocDate.Split('-')[1]);
                                int compare = DocDateDateTime.CompareTo(erpDateTime);
                                if (compare <= 0) throw new SystemException("ERP已關帳(" + erpFullDate + ")，無法開立此日期(" + DocDate + ")之單據!!");
                            }
                            #endregion

                            #region //判斷請購變更單資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.PrmStatus, a.UserId
                                    FROM SCM.PrModification a
                                    WHERE a.PrmId = @PrmId";
                            dynamicParameters.Add("PrmId", PrmId);

                            var PrmResult = sqlConnection.Query(sql, dynamicParameters);
                            if (PrmResult.Count() <= 0) throw new SystemException("請購變更單資料錯誤!");

                            foreach (var item in PrmResult)
                            {
                                if (item.PrmStatus != "N" && item.PrmStatus != "E") throw new SystemException("請購變更單狀態無法更改!");
                                if (item.UserId != CreateBy) throw new SystemException("無法修改其他請購人之單據!");
                            }
                            #endregion

                            #region //判斷原請購單資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.BpmNo, a.DepartmentId OriDepartmentId, a.UserId OriUserId
                                    , a.Edition OriEdition, a.PrDate OriPrDate, a.PrRemark OriPrRemark, a.Amount OriAmount
                                    , a.BudgetDepartmentId OriBudgetDepartmentId
									, a.PrErpPrefix, a.PrErpNo
                                    FROM SCM.PurchaseRequisition a
                                    WHERE a.PrId = @PrId";
                            dynamicParameters.Add("PrId", PrId);
                            string PrErpPrefix = "";
                            string PrErpNo = "";

                            var PrResult = sqlConnection.Query(sql, dynamicParameters);
                            if (PrResult.Count() <= 0) throw new SystemException("請購單資料錯誤!");
                            foreach (var item in PrResult)
                            {
                                PrErpPrefix = item.PrErpPrefix;
                                PrErpNo = item.PrErpNo;

                            }
                            #endregion

                            #region //該請購單是否有開立採購單

                            dynamicParameters = new DynamicParameters();
                            sql = @"select * 
                                    from PURTD
                                    where TD026 = @PrErpPrefix
                                    and TD027 = @PrErpNo";
                            dynamicParameters.Add("PrErpPrefix", PrErpPrefix);
                            dynamicParameters.Add("PrErpNo", PrErpNo);

                            var PoERPResult = sqlConnection2.Query(sql, dynamicParameters);
                            if (PoERPResult.Count() > 0) throw new SystemException("該請購單已有採購單，無法開立變更單");

                            #endregion


                            #region //確認請購人員資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.[User] a
                                    WHERE a.UserId = @UserId";
                            dynamicParameters.Add("UserId", UserId);

                            var UserResult = sqlConnection.Query(sql, dynamicParameters);
                            if (UserResult.Count() <= 0) throw new SystemException("【使用者】資料錯誤，請重新輸入!");
                            #endregion

                            #region //確認請購部門資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM BAS.Department a
                                    WHERE a.DepartmentId = @DepartmentId";
                            dynamicParameters.Add("DepartmentId", DepartmentId);

                            var DepartmentResult = sqlConnection.Query(sql, dynamicParameters);
                            if (DepartmentResult.Count() <= 0) throw new SystemException("【請購部門】資料錯誤，請重新輸入!");
                            #endregion

                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.PrModification SET
                                    CompanyId = @CompanyId,
                                    PrId = @PrId,
                                    DepartmentId = @DepartmentId,
                                    UserId = @UserId,
                                    DocDate = @DocDate,
                                    ModiReason = @ModiReason,
                                    PrmRemark = @PrmRemark,
                                    Priority = @Priority,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE PrmId = @PrmId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    CompanyId = CurrentCompany,
                                    PrId,
                                    DepartmentId,
                                    UserId,
                                    DocDate,
                                    ModiReason,
                                    PrmRemark,
                                    Priority,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    PrmId
                                });

                            int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                            if (PrmFile.Length > 0)
                            {
                                #region //先將原本的砍掉
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE FROM SCM.PrmFile
                                        WHERE PrmId = @PrmId";
                                dynamicParameters.Add("PrmId", PrmId);

                                rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                string[] prmFiles = PrmFile.Split(',');
                                foreach (var file in prmFiles)
                                {
                                    #region //更新SCM.PrmFile
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO SCM.PrmFile (PrmId, FileId
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.PrmFileId
                                            VALUES (@PrmId, @FileId
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            PrmId,
                                            FileId = file,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    var insertResult3 = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult3.Count();
                                    #endregion
                                }
                            }
                            else
                            {
                                #region //將原本的砍掉
                                dynamicParameters = new DynamicParameters();
                                sql = @"DELETE FROM SCM.PrmFile
                                        WHERE PrmId = @PrmId";
                                dynamicParameters.Add("PrmId", PrmId);

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

        #region //UpdatePrmDetail -- 更新請購變更單詳細資料 -- Ann 2023-02-09
        public string UpdatePrmDetail(int PrmDetailId, int PrDetailId, string PrmSequence, int MtlItemId, string PrMtlItemName, string PrMtlItemSpec, int InventoryId, int PrUomId, int PrQty, string DemandDate
            , int SupplierId, string PrCurrency, string PrExchangeRate, double PrUnitPrice, double PrPrice, double PrPriceTw
            , string UrgentMtl, string ProductionPlan, string Project, int SoDetailId, string PoRemark, string PrDetailRemark, string ModiReason, string ClosureStatus)
        {
            try
            {
                if (PrQty < 0) throw new SystemException("【請購數量】不能為空!");
                if (PrQty == 0 && ClosureStatus != "y") throw new SystemException("若【請購數量】為0，此單據需指定結案!");
                if (DemandDate.Length <= 0) throw new SystemException("【需求日期】不能為空!");
                if (PrExchangeRate.Length <= 0) throw new SystemException("【匯率】不能為空!");
                if (PrPrice <= 0) throw new SystemException("【請購金額】不能為空!");
                if (PrPriceTw <= 0) throw new SystemException("【本幣金額】不能為空!");
                if (UrgentMtl.Length <= 0) throw new SystemException("【是否急料】不能為空!");
                if (ProductionPlan.Length <= 0) throw new SystemException("【是否納入生產計畫】不能為空!");
                if (PrmSequence.Length <= 0) throw new SystemException("【請購變更單序號】不能為空!");
                if (PrMtlItemName.Length <= 0) throw new SystemException("【品名】不能為空!");
                //if (PrMtlItemSpec.Length <= 0) throw new SystemException("【規格】不能為空!");

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
                            #region //判斷請購變更單詳細資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.PrmId, a.PrQty, a.PrPrice
                                    , b.PrmStatus, b.UserId
                                    FROM SCM.PrmDetail a
                                    INNER JOIN SCM.PrModification b ON a.PrmId = b.PrmId
                                    WHERE a.PrmDetailId = @PrmDetailId";
                            dynamicParameters.Add("PrmDetailId", PrmDetailId);

                            var prmDetailIdResult = sqlConnection.Query(sql, dynamicParameters);
                            if (prmDetailIdResult.Count() <= 0) throw new SystemException("請購變更單詳細資料錯誤!");

                            int PrmId = -1;
                            double OrgPrQty = -1;
                            double OrgPrPrice = -1;
                            foreach (var item in prmDetailIdResult)
                            {
                                if (item.PrmStatus != "N" && item.PrmStatus != "E") throw new SystemException("請購變更單狀態無法更改!");
                                if (item.UserId != CreateBy) throw new SystemException("無法修改其他請購人之單據!");
                                PrmId = item.PrmId;
                                OrgPrQty = item.PrQty;
                                OrgPrPrice = item.PrPrice;
                            }
                            #endregion

                            #region //確認請購單詳細資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.PrId, a.PrSequence OriPrSequence, a.MtlItemId OriMtlItemId, a.PrMtlItemName OriPrMtlItemName
                                    , a.PrMtlItemSpec OriPrMtlItemSpec, a.InventoryId OriInventoryId, a.PrUomId OriPrUomId
                                    , a.PrQty OriPrQty, a.DemandDate OriDemandDate, a.SupplierId OriSupplierId, a.PrCurrency OriPrCurrency
                                    , a.PrExchangeRate OriPrExchangeRate, a.PrUnitPrice OriPrUnitPrice, a.PrPrice OriPrPrice
                                    , a.PrPriceTw OriPrPriceTw, a.UrgentMtl OriUrgentMtl, a.ProductionPlan OriProductionPlan
                                    , a.Project OriProject, ISNULL(a.BudgetDepartmentNo, '') OriBudgetDepartmentNo
                                    , ISNULL(a.BudgetDepartmentSubject, '') OriBudgetDepartmentSubject, a.SoDetailId OriSoDetailId
                                    , ISNULL(FORMAT(a.DeliveryDate, 'yyyy-MM-dd'), null) OriDeliveryDate, a.PoUserId OriPoUserId
                                    , a.PoUomId OriPoUomId, a.PoQty OriPoQty, ISNULL(a.PoCurrency, '') OriPoCurrency
                                    , a.PoUnitPrice OriPoUnitPrice, a.PoPrice OriPoPrice, a.LockStaus OriLockStaus
                                    , ISNULL(a.PoStaus, '') OriPoStaus, ISNULL(a.TaxNo, '') OriTaxNo, a.Taxation OriTaxation
                                    , a.BusinessTaxRate OriBusinessTaxRate, ISNULL(a.DetailMultiTax, '') OriDetailMultiTax
                                    , ISNULL(a.TradeTerm, '') OriTradeTerm, a.PrPriceQty OriPrPriceQty, a.PrPriceUomId OriPrPriceUomId
                                    , a.DiscountRate OriDiscountRate, a.DiscountAmount OriDiscountAmount, a.ClosureStatus OriClosureStatus
                                    , ISNULL(a.PrDetailRemark, '') OriPrDetailRemark, ISNULL(a.PoRemark, '') OriPoRemark
                                    FROM SCM.PrDetail a
                                    WHERE a.PrDetailId = @PrDetailId";
                            dynamicParameters.Add("PrDetailId", PrDetailId);

                            var PrDetailResult = sqlConnection.Query(sql, dynamicParameters);

                            if (PrDetailResult.Count() <= 0) throw new SystemException("【請購單詳細資料】資料錯誤!");

                            string OriPrSequence = "";
                            int OriMtlItemId = -1;
                            string OriPrMtlItemName = "";
                            string OriPrMtlItemSpec = "";
                            int OriInventoryId = -1;
                            int OriPrUomId = -1;
                            double OriPrQty = -1;
                            DateTime OriDemandDate = new DateTime();
                            int OriSupplierId = -1;
                            string OriPrCurrency = "";
                            double OriPrExchangeRate = -1;
                            double OriPrUnitPrice = -1;
                            double OriPrPrice = -1;
                            double OriPrPriceTw = -1;
                            string OriUrgentMtl = "";
                            string OriProductionPlan = "";
                            string OriProject = "";
                            string OriBudgetDepartmentNo = "";
                            string OriBudgetDepartmentSubject = "";
                            int? OriSoDetailId = -1;
                            DateTime? OriDeliveryDate = new DateTime();
                            int? OriPoUserId = -1;
                            int? OriPoUomId = -1;
                            double? OriPoQty = -1;
                            string OriPoCurrency = "";
                            double? OriPoUnitPrice = -1;
                            double? OriPoPrice = -1;
                            string OriLockStaus = "";
                            string OriPoStaus = "";
                            string OriTaxNo = "";
                            string OriTaxation = "";
                            double? OriBusinessTaxRate = -1;
                            string OriDetailMultiTax = "";
                            string OriTradeTerm = "";
                            int OriPrPriceQty = -1;
                            int OriPrPriceUomId = -1;
                            double? OriDiscountRate = -1;
                            double? OriDiscountAmount = -1;
                            string OriClosureStatus = "";
                            string OriPrDetailRemark = "";
                            string OriPoRemark = "";
                            int PrId = -1;
                            foreach (var item in PrDetailResult)
                            {
                                OriPrSequence = item.OriPrSequence;
                                OriMtlItemId = item.OriMtlItemId;
                                OriPrMtlItemName = item.OriPrMtlItemName;
                                OriPrMtlItemSpec = item.OriPrMtlItemSpec;
                                OriInventoryId = item.OriInventoryId;
                                OriPrUomId = item.OriPrUomId;
                                OriPrQty = item.OriPrQty;
                                OriDemandDate = item.OriDemandDate;
                                OriSupplierId = item.OriSupplierId;
                                OriPrCurrency = item.OriPrCurrency;
                                OriPrExchangeRate = item.OriPrExchangeRate;
                                OriPrUnitPrice = item.OriPrUnitPrice;
                                OriPrPrice = item.OriPrPrice;
                                OriPrPriceTw = item.OriPrPriceTw;
                                OriUrgentMtl = item.OriUrgentMtl;
                                OriProductionPlan = item.OriProductionPlan;
                                OriProject = item.OriProject;
                                OriBudgetDepartmentNo = item.OriBudgetDepartmentNo;
                                OriBudgetDepartmentSubject = item.OriBudgetDepartmentSubject;
                                OriSoDetailId = item.OriSoDetailId;
                                OriDeliveryDate = item.OriDeliveryDate;
                                OriPoUserId = item.OriPoUserId;
                                OriPoUomId = item.OriPoUomId;
                                OriPoQty = item.OriPoQty;
                                OriPoCurrency = item.OriPoCurrency;
                                OriPoUnitPrice = item.OriPoUnitPrice;
                                OriPoPrice = item.OriPoPrice;
                                OriLockStaus = item.OriLockStaus;
                                OriPoStaus = item.OriPoStaus;
                                OriTaxNo = item.OriTaxNo;
                                OriTaxation = item.OriTaxation;
                                OriBusinessTaxRate = item.OriBusinessTaxRate;
                                OriDetailMultiTax = item.OriDetailMultiTax;
                                OriTradeTerm = item.OriTradeTerm;
                                OriPrPriceQty = item.OriPrPriceQty;
                                OriPrPriceUomId = item.OriPrPriceUomId;
                                OriDiscountRate = item.OriDiscountRate;
                                OriDiscountAmount = item.OriDiscountAmount;
                                OriClosureStatus = item.OriClosureStatus;
                                OriPrDetailRemark = item.OriPrDetailRemark;
                                OriPoRemark = item.OriPoRemark;
                                PrId = item.PrId;
                            }
                            #endregion

                            #region //判斷品號資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.MtlItemNo, a.MtlItemName, a.MtlItemSpec
                                    FROM PDM.MtlItem a
                                    WHERE a.MtlItemId = @MtlItemId";
                            dynamicParameters.Add("MtlItemId", MtlItemId);

                            var result2 = sqlConnection.Query(sql, dynamicParameters);

                            if (result2.Count() <= 0) throw new SystemException("【品號】資料錯誤!");

                            string MtlItemNo = "";
                            string MtlItemName = "";
                            string MtlItemSpec = "";
                            foreach (var item in result2)
                            {
                                MtlItemNo = item.MtlItemNo;
                                MtlItemName = item.MtlItemName;
                                MtlItemSpec = item.MtlItemSpec;
                            }
                            #endregion

                            #region //判斷ERP品號生效日與失效日
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MB030)) MB030, LTRIM(RTRIM(MB031)) MB031
                                    FROM INVMB
                                    WHERE MB001 = @MB001";
                            dynamicParameters.Add("MB001", MtlItemNo);

                            var INVMBResult = sqlConnection2.Query(sql, dynamicParameters);

                            foreach (var item in INVMBResult)
                            {
                                if (item.MB030 != "" && item.MB030 != null)
                                {
                                    #region //判斷日期需大於或等於生效日
                                    string EffectiveDate = item.MB030;
                                    string effYear = EffectiveDate.Substring(0, 4);
                                    string effMonth = EffectiveDate.Substring(4, 2);
                                    string effDay = EffectiveDate.Substring(6, 2);
                                    DateTime effFullDate = Convert.ToDateTime(effYear + "-" + effMonth + "-" + effDay);
                                    int effresult = DateTime.Compare(CreateDate, effFullDate);
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
                                    int effresult = DateTime.Compare(CreateDate, effFullDate);
                                    if (effresult > 0) throw new SystemException("此品號已失效，無法使用!!");
                                    #endregion
                                }
                            }
                            #endregion

                            #region //判斷庫別資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.InventoryNo, a.InventoryName
                                    FROM SCM.Inventory a
                                    WHERE a.InventoryId = @InventoryId";
                            dynamicParameters.Add("InventoryId", InventoryId);

                            var result3 = sqlConnection.Query(sql, dynamicParameters);

                            if (result3.Count() <= 0) throw new SystemException("【庫別】資料錯誤!");

                            string InventoryNo = "";
                            string InventoryName = "";
                            foreach (var item in result3)
                            {
                                InventoryNo = item.InventoryNo;
                                InventoryName = item.InventoryName;
                            }
                            #endregion

                            #region //判斷單位資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM PDM.UnitOfMeasure a
                                    WHERE a.UomId = @UomId";
                            dynamicParameters.Add("UomId", PrUomId);

                            var result4 = sqlConnection.Query(sql, dynamicParameters);

                            if (result4.Count() <= 0) throw new SystemException("【單位】資料錯誤!");
                            #endregion

                            #region //取得ERP庫存資訊
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(a.MC007)) InventoryQty
                                    FROM INVMC a
                                    WHERE a.MC001 = @MC001
                                    AND a.MC002 = @MC002";
                            dynamicParameters.Add("MC001", MtlItemNo);
                            dynamicParameters.Add("MC002", InventoryNo);

                            var result5 = sqlConnection2.Query(sql, dynamicParameters);

                            double InventoryQty = 0;
                            string mtlInventory = "目前尚無資料";
                            if (result5.Count() > 0)
                            {
                                foreach (var item in result5)
                                {
                                    InventoryQty = Convert.ToDouble(item.InventoryQty);
                                    #region //組MtlInventory
                                    List<MtlInventory> mtlInventories = new List<MtlInventory>
                                    {
                                        new MtlInventory
                                        {
                                            WAREHOUSE_NO = InventoryNo,
                                            WAREHOUSE_NAME = InventoryName,
                                            WAREHOUSE_QTY = InventoryQty
                                        }
                                    };
                                    #endregion

                                    mtlInventory = JsonConvert.SerializeObject(mtlInventories);
                                    mtlInventory = "{\"data\":" + mtlInventory + "}";
                                }
                            }
                            #endregion

                            #region //檢查供應商資料是否正確
                            string TaxNo = "";
                            string Taxation = "";
                            string TradeTerm = "";
                            double? BusinessTaxRate = -1;
                            string HideSupplier = "";
                            if (SupplierId > 0)
                            {
                                #region //供應商
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.TaxNo, a.Taxation, a.TradeTerm, a.HideSupplier
                                        FROM SCM.Supplier a
                                        WHERE a.SupplierId = @SupplierId";
                                dynamicParameters.Add("SupplierId", SupplierId);

                                var result6 = sqlConnection.Query(sql, dynamicParameters);

                                if (result6.Count() <= 0) throw new SystemException("【供應商】資料錯誤!");

                                foreach (var item in result6)
                                {
                                    TaxNo = item.TaxNo;
                                    Taxation = item.Taxation;
                                    TradeTerm = item.TradeTerm;
                                    HideSupplier = item.HideSupplier;
                                }
                                #endregion

                                #region //查詢營業稅額資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT ROUND(LTRIM(RTRIM(NN004)),2) BusinessTaxRate
                                        FROM CMSNN 
                                        WHERE NN001 = @TaxNo";
                                dynamicParameters.Add("TaxNo", TaxNo);

                                var businessTaxRateResult = sqlConnection2.Query(sql, dynamicParameters);

                                if (businessTaxRateResult.Count() <= 0) throw new SystemException("【課稅別】資料錯誤!");

                                foreach (var item in businessTaxRateResult)
                                {
                                    BusinessTaxRate = item.BusinessTaxRate;
                                }
                                #endregion
                            }
                            #endregion

                            #region //確認集團內/集團外邏輯
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 b.HideSupplier
                                    FROM SCM.PrDetail a 
                                    INNER JOIN SCM.Supplier b ON a.SupplierId = b.SupplierId
                                    WHERE a.PrId = @PrId
                                    ORDER BY a.PrSequence";
                            dynamicParameters.Add("PrId", PrId);

                            var PrDetailResult2 = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item in PrDetailResult2)
                            {
                                string HideSupplierString = "";
                                if (item.HideSupplier == "Y")
                                {
                                    HideSupplierString = "集團內";
                                }
                                else
                                {
                                    HideSupplierString = "集團外";
                                }
                                if (item.HideSupplier != HideSupplier) throw new SystemException("此請購單第一筆單身為【" + HideSupplierString + "】，無法新增!!");
                            }
                            #endregion

                            #region //判斷幣別資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM CMSMF
                                    WHERE MF001 = @MF001";
                            dynamicParameters.Add("MF001", PrCurrency);

                            var result7 = sqlConnection2.Query(sql, dynamicParameters);
                            if (result7.Count() <= 0) throw new SystemException("【幣別】資料有誤!");
                            #endregion

                            #region //判斷專案代碼資料是否正確
                            if (Project.Length > 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                    FROM CMSNB
                                    WHERE NB001 = @NB001";
                                dynamicParameters.Add("NB001", Project);

                                var result8 = sqlConnection2.Query(sql, dynamicParameters);
                                if (result8.Count() <= 0) throw new SystemException("【專案代碼】資料有誤!");
                            }
                            #endregion

                            #region //確認訂單資料是否正確
                            if (SoDetailId > 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT MtlItemId
                                        FROM SCM.SoDetail
                                        WHERE SoDetailId = @SoDetailId";
                                dynamicParameters.Add("SoDetailId", SoDetailId);

                                var result9 = sqlConnection.Query(sql, dynamicParameters);
                                if (result9.Count() <= 0) throw new SystemException("【訂單】資料有誤!");

                                foreach (var item in result9)
                                {
                                    if (item.MtlItemId != MtlItemId) throw new SystemException("訂單品號與請購品號不同!!");
                                }
                            }
                            #endregion

                            #region //UPDATE SCM.PrmDetail
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.PrmDetail SET
                                    PrmId = @PrmId,
                                    PrDetailId = @PrDetailId,
                                    PrmSequence = @PrmSequence,
                                    MtlItemId = @MtlItemId,
                                    PrMtlItemName = @PrMtlItemName,
                                    PrMtlItemSpec = @PrMtlItemSpec,
                                    InventoryId = @InventoryId,
                                    PrUomId = @PrUomId,
                                    PrQty = @PrQty,
                                    DemandDate = @DemandDate,
                                    SupplierId = @SupplierId,
                                    PrCurrency = @PrCurrency,
                                    PrExchangeRate = @PrExchangeRate,
                                    PrUnitPrice = @PrUnitPrice,
                                    PrPrice = @PrPrice,
                                    PrPriceTw = @PrPriceTw,
                                    UrgentMtl = @UrgentMtl,
                                    PoUomId = @PrUomId,
                                    PoQty = @PrQty,
                                    PoCurrency = @PrCurrency,
                                    PoUnitPrice = @PrUnitPrice,
                                    PoPrice = @PrPrice,
                                    ProductionPlan = @ProductionPlan,
                                    Project = @Project,
                                    SoDetailId = @SoDetailId,
                                    DeliveryDate = @DeliveryDate,
                                    LockStaus = @LockStaus,
                                    TaxNo = @TaxNo,
                                    Taxation = @Taxation,
                                    BusinessTaxRate = @BusinessTaxRate,
                                    DetailMultiTax = @DetailMultiTax,
                                    TradeTerm = @TradeTerm,
                                    PrPriceQty = @PrPriceQty,
                                    PrPriceUomId = @PrPriceUomId,
                                    DiscountRate = @DiscountRate,
                                    DiscountAmount = @DiscountAmount,
                                    MtlInventory = @MtlInventory,
                                    MtlInventoryQty = @MtlInventoryQty,
                                    ConfirmStatus = @ConfirmStatus,
                                    ConfirmUserId = @ConfirmUserId,
                                    ClosureStatus = @ClosureStatus,
                                    PrDetailRemark = @PrDetailRemark,
                                    PoRemark = @PoRemark,
                                    ModiReason = @ModiReason,
                                    OriPrSequence = @OriPrSequence,
                                    OriMtlItemId = @OriMtlItemId,
                                    OriPrMtlItemName = @OriPrMtlItemName,
                                    OriPrMtlItemSpec = @OriPrMtlItemSpec,
                                    OriInventoryId = @OriInventoryId,
                                    OriPrUomId = @OriPrUomId,
                                    OriPrQty = @OriPrQty,
                                    OriDemandDate = @OriDemandDate,
                                    OriSupplierId = @OriSupplierId,
                                    OriPrCurrency = @OriPrCurrency,
                                    OriPrExchangeRate = @OriPrExchangeRate,
                                    OriPrUnitPrice = @OriPrUnitPrice,
                                    OriPrPrice = @OriPrPrice,
                                    OriPrPriceTw = @OriPrPriceTw,
                                    OriUrgentMtl = @OriUrgentMtl,
                                    OriProductionPlan = @OriProductionPlan,
                                    OriProject = @OriProject,
                                    OriBudgetDepartmentNo = @OriBudgetDepartmentNo,
                                    OriBudgetDepartmentSubject = @OriBudgetDepartmentSubject,
                                    OriSoDetailId = @OriSoDetailId,
                                    OriDeliveryDate = @OriDeliveryDate,
                                    OriPoUserId = @OriPoUserId,
                                    OriPoUomId = @OriPoUomId,
                                    OriPoQty = @OriPoQty,
                                    OriPoCurrency = @OriPoCurrency,
                                    OriPoUnitPrice = @OriPoUnitPrice,
                                    OriPoPrice = @OriPoPrice,
                                    OriLockStaus = @OriLockStaus,
                                    OriPoStaus = @OriPoStaus,
                                    OriTaxNo = @OriTaxNo,
                                    OriTaxation = @OriTaxation,
                                    OriBusinessTaxRate = @OriBusinessTaxRate,
                                    OriDetailMultiTax = @OriDetailMultiTax,
                                    OriTradeTerm = @OriTradeTerm,
                                    OriPrPriceQty = @OriPrPriceQty,
                                    OriPrPriceUomId = @OriPrPriceUomId,
                                    OriDiscountRate = @OriDiscountRate,
                                    OriDiscountAmount = @OriDiscountAmount,
                                    OriClosureStatus = @OriClosureStatus,
                                    OriPrDetailRemark = @OriPrDetailRemark,
                                    OriPoRemark = @OriPoRemark,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE PrmDetailId = @PrmDetailId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    PrmId,
                                    PrDetailId,
                                    PrmSequence,
                                    MtlItemId,
                                    PrMtlItemName,
                                    PrMtlItemSpec,
                                    InventoryId,
                                    PrUomId,
                                    PrQty,
                                    DemandDate,
                                    SupplierId,
                                    PrCurrency,
                                    PrExchangeRate,
                                    PrUnitPrice,
                                    PrPrice,
                                    PrPriceTw,
                                    UrgentMtl,
                                    ProductionPlan,
                                    Project,
                                    SoDetailId,
                                    DeliveryDate = (DateTime?)null,
                                    LockStaus = "N",
                                    PoStaus = "N",
                                    TaxNo = TaxNo != "" ? TaxNo : null,
                                    Taxation = Taxation != "" ? Taxation : null,
                                    BusinessTaxRate = BusinessTaxRate > 0 ? BusinessTaxRate : (double?)null,
                                    DetailMultiTax = "N",
                                    TradeTerm = TradeTerm != "" ? TradeTerm : null,
                                    PrPriceQty = PrQty,
                                    PrPriceUomId = PrUomId,
                                    DiscountRate = 0,
                                    DiscountAmount = 0,
                                    MtlInventory = mtlInventory,
                                    MtlInventoryQty = InventoryQty,
                                    ConfirmStatus = "N",
                                    ConfirmUserId = (int?)null,
                                    ClosureStatus,
                                    PrDetailRemark,
                                    PoRemark,
                                    ModiReason,
                                    OriPrSequence,
                                    OriMtlItemId,
                                    OriPrMtlItemName,
                                    OriPrMtlItemSpec,
                                    OriInventoryId,
                                    OriPrUomId,
                                    OriPrQty,
                                    OriDemandDate,
                                    OriSupplierId,
                                    OriPrCurrency,
                                    OriPrExchangeRate,
                                    OriPrUnitPrice,
                                    OriPrPrice,
                                    OriPrPriceTw,
                                    OriUrgentMtl,
                                    OriProductionPlan,
                                    OriProject,
                                    OriBudgetDepartmentNo,
                                    OriBudgetDepartmentSubject,
                                    OriSoDetailId,
                                    OriDeliveryDate,
                                    OriPoUserId,
                                    OriPoUomId,
                                    OriPoQty,
                                    OriPoCurrency,
                                    OriPoUnitPrice,
                                    OriPoPrice,
                                    OriLockStaus,
                                    OriPoStaus,
                                    OriTaxNo,
                                    OriTaxation,
                                    OriBusinessTaxRate,
                                    OriDetailMultiTax,
                                    OriTradeTerm,
                                    OriPrPriceQty,
                                    OriPrPriceUomId,
                                    OriDiscountRate,
                                    OriDiscountAmount,
                                    OriClosureStatus,
                                    OriPrDetailRemark,
                                    OriPoRemark,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    PrmDetailId
                                });

                            int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //重新計算總請購數量、金額
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT ISNULL(SUM(a.PrQty), 0) TotalPrQty
                                    , ISNULL(SUM(a.PrPriceTw), 0) TotalPrPrice
                                    FROM SCM.PrmDetail a
                                    WHERE a.PrmId = @PrmId";
                            dynamicParameters.Add("PrmId", PrmId);

                            var PrmDetailResult = sqlConnection.Query(sql, dynamicParameters);

                            double TotalPrQty = -1;
                            double TotalPrPrice = -1;
                            foreach (var item in PrmDetailResult)
                            {
                                TotalPrQty = item.TotalPrQty;
                                TotalPrPrice = item.TotalPrPrice;
                            }
                            #endregion

                            #region //更新請購變更單頭
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.PrModification SET
                                    TotalQty = @TotalPrQty,
                                    Amount = @TotalPrPrice,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE PrmId = @PrmId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    TotalPrQty,
                                    TotalPrPrice,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    PrmId
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

        #region //UpdatePrmTransferBpm -- 拋轉請購變更單據至BPM -- Ann 2023-02-09//
        public string UpdatePrmTransferBpm(int PrmId)
        {
            try
            {
                string token = "";
                int rowsAffected = 0;
                string CompanyNo = "";
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //取得USER資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserNo
                                FROM BAS.[User] a
                                WHERE a.UserId = @UserId";
                        dynamicParameters.Add("UserId", CreateBy);

                        var UserResult = sqlConnection.Query(sql, dynamicParameters);

                        string UserNo = "";
                        foreach (var item in UserResult)
                        {
                            UserNo = item.UserNo;
                        }
                        #endregion

                        #region //確認公司別DB
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb, a.CompanyNo
                                FROM BAS.Company a
                                WHERE CompanyId = @CompanyId";
                        dynamicParameters.Add("CompanyId", CurrentCompany);

                        var companyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (companyResult.Count() <= 0) throw new SystemException("公司別錯誤!");

                        foreach (var item in companyResult)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            CompanyNo = item.CompanyNo;
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //查詢ERP是否有此帳號
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(a.MF004)) USR_GROUP
                                    FROM ADMMF a
                                    WHERE MF001 = @MF001";
                            dynamicParameters.Add("MF001", UserNo);

                            var ADMMFResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (ADMMFResult.Count() <= 0) throw new SystemException("查無ERP帳號權限，無法進行拋轉及核單!");
                            #endregion

                            #region //判斷請購變更單資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.PrmStatus, a.Edition PrmEdition, FORMAT(a.DocDate, 'yyyy-MM-dd') DocDate
                                    , FORMAT(a.ModiDate, 'yyyy-MM-dd') ModiDate, a.ModiReason, a.PrmRemark, a.Amount, a.UserId, a.Priority
                                    , b.PrErpPrefix, b.PrErpNo, b.Edition PrEdition, b.PrNo, FORMAT(b.PrDate, 'yyyy-MM-dd') PrDate
                                    , c.DepartmentNo, c.DepartmentName
                                    FROM SCM.PrModification a
                                    INNER JOIN SCM.PurchaseRequisition b ON a.PrId = b.PrId
                                    INNER JOIN BAS.Department c ON a.DepartmentId = c.DepartmentId
                                    WHERE a.PrmId = @PrmId";
                            dynamicParameters.Add("PrmId", PrmId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("請購變更單資料錯誤!");

                            string PrErpPrefix = "";
                            string PrErpNo = "";
                            string PrmEdition = "";
                            string PrEdition = "";
                            string PrNo = "";
                            string DocDate = "";
                            string PrDate = "";
                            string ModiDate = "";
                            string ModiReason = "";
                            string PrmRemark = "";
                            double Amount = -1;
                            string DepartmentNo = "";
                            string DepartmentName = "";
                            string Priority = "";
                            foreach (var item in result)
                            {
                                if (item.PrmStatus != "N" && item.PrmStatus != "E" && item.PrmStatus != "R") throw new SystemException("請購變更單狀態無法拋轉!");
                                if (item.UserId != CreateBy) throw new SystemException("無法修改其他請購人之單據!");
                                PrErpPrefix = item.PrErpPrefix;
                                PrErpNo = item.PrErpNo;
                                PrmEdition = item.PrmEdition;
                                PrEdition = item.PrEdition;
                                PrNo = item.PrNo;
                                DocDate = item.DocDate;
                                PrDate = item.PrDate;
                                ModiDate = item.ModiDate;
                                ModiReason = item.ModiReason;
                                PrmRemark = item.PrmRemark;
                                Amount = item.Amount;
                                DepartmentNo = item.DepartmentNo;
                                DepartmentName = item.DepartmentName;
                                Priority = item.Priority;
                            }
                            #endregion

                            #region //該請購單是否有開立採購單
                            dynamicParameters = new DynamicParameters();
                            sql = @"select * 
                                    from PURTD
                                    where TD026 = @PrErpPrefix
                                    and TD027 = @PrErpNo";
                            dynamicParameters.Add("PrErpPrefix", PrErpPrefix);
                            dynamicParameters.Add("PrErpNo", PrErpNo);

                            var PoERPResult = sqlConnection2.Query(sql, dynamicParameters);
                            if (PoERPResult.Count() > 0) throw new SystemException("該請購單已有採購單，無法開立變更單");

                            #endregion


                            #region //取得BPM TOKEN
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.IpAddress, a.CompanyId, a.Token, FORMAT(a.VerifyDate, 'yyyy-MM-dd HH:mm:ss') VerifyDate
                                    FROM BPM.SystemToken a
                                    WHERE IpAddress = @IpAddress
                                    AND CompanyId = @CompanyId";
                            dynamicParameters.Add("IpAddress", BpmServerPath);
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            var systemTokenResult = sqlConnection.Query(sql, dynamicParameters);
                            if (systemTokenResult.Count() <= 0) throw new SystemException("查無此憑證!");

                            foreach (var item in systemTokenResult)
                            {
                                DateTime verifyDate = Convert.ToDateTime(item.VerifyDate);
                                DateTime nowDate = DateTime.Now;
                                var CheckMin = (nowDate - verifyDate).TotalMinutes;
                                if (CheckMin >= 30)
                                {
                                    #region //取得新BPM TOKEN
                                    string tokenResponse = BpmHelper.GetBpmToken(BpmServerPath, BpmAccount, BpmPassword);
                                    var tokenJson = JObject.Parse(tokenResponse);
                                    foreach (var item2 in tokenJson)
                                    {
                                        if (item2.Key == "status")
                                        {
                                            if (item2.Value.ToString() != "success") throw new SystemException("取得token失敗!");
                                        }
                                        else if (item2.Key == "data")
                                        {
                                            token = item2.Value.ToString();
                                        }
                                    }
                                    #endregion

                                    #region //將新的TOKEN更新回MES
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE BPM.SystemToken SET
                                            Token = @Token,
                                            VerifyDate = @VerifyDate,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE IpAddress = @IpAddress
                                            AND CompanyId = @CompanyId";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            Token = token,
                                            VerifyDate = nowDate,
                                            LastModifiedDate,
                                            LastModifiedBy,
                                            IpAddress = BpmServerPath,
                                            CompanyId = CurrentCompany
                                        });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                                else
                                {
                                    token = item.Token;
                                }
                            }
                            #endregion

                            #region //取得BpmUser資料
                            string BpmUserId = "";
                            string BpmRoleId = "";
                            string BpmUserNo = "";
                            string UserName = "";
                            string BpmDepNo = "";
                            string BpmDepName = "";
                            using (SqlConnection sqlConnection3 = new SqlConnection(BpmDbConnectionStrings))
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"WITH BasicUserInfo(MemID, LoginID, UserName, MainRoleID, RolID, RolName, ParentRol) AS(
                                        SELECT a.MemID, a.LoginID, a.UserName, a.MainRoleID
                                        , b.RolID, b.Name, b.DepID AS ParentRol
                                        FROM Mem_GenInf a
                                        LEFT JOIN Rol_GenInf b ON b.RolID = a.MainRoleID
                                        WHERE a.LoginID = @LoginID
                                        UNION ALL
                                        SELECT a.MemID, a.LoginID, a.UserName, a.MainRoleID
                                        , b.RolID, b.Name, b.DepID AS ParentRol
                                        FROM BasicUserInfo a, Rol_GenInf b
                                        WHERE a.ParentRol = b.RolID
                                        )
                                        SELECT a.MemID, a.LoginID, a.UserName, a.MainRoleID
                                        , b.Name AS RolName
                                        , c.DepID, c.ID AS DepNo, c.Name AS DepName
                                        FROM BasicUserInfo a
                                        LEFT JOIN Rol_GenInf AS parentRol_GenInf ON a.RolID = parentRol_GenInf.RolID
                                        LEFT JOIN Dep_GenInf c ON parentRol_GenInf.DepID = c.DepID
                                        , Rol_GenInf b
                                        WHERE a.MainRoleID = b.RolID
                                        AND c.DepID IS NOT NULL
                                        UNION
                                        SELECT a.MemID, a.LoginID, a.UserName, a.MainRoleID
                                        , b.Name AS RolName
                                        , c.ComID AS DepID, c.ID AS DepNo, c.Name AS DepName
                                        FROM Mem_GenInf a
                                        LEFT JOIN Rol_GenInf b ON a.MainRoleID = b.RolID
                                        LEFT JOIN Company c ON b.DepID = c.ComID
                                        WHERE c.ComID IS NOT NULL
                                        AND a.LoginID = @LoginID
                                        ORDER BY a.LoginID";
                                dynamicParameters.Add("LoginID", UserNo);

                                var MemGenInfResult = sqlConnection3.Query(sql, dynamicParameters);

                                if (MemGenInfResult.Count() <= 0) throw new SystemException("取得BPM使用者資訊時發生錯誤!!");

                                foreach (var item in MemGenInfResult)
                                {
                                    BpmUserId = item.MemID;
                                    BpmRoleId = item.MainRoleID;
                                    BpmUserNo = item.LoginID;
                                    UserName = item.UserName;
                                    BpmDepNo = item.DepNo;
                                    BpmDepName = item.DepName;
                                }
                            }
                            #endregion

                            #region //取得單身資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.PrmDetailId, a.PrDetailId, a.PrmSequence, a.PrMtlItemName, a.PrMtlItemSpec, a.PrQty
                                    , FORMAT(a.DemandDate, 'yyyy-MM-dd') DemandDate, a.PrCurrency, a.PrExchangeRate
                                    , a.PrUnitPrice, a.PrPrice, a.PrPriceTw, a.UrgentMtl, a.ProductionPlan, a.Project
                                    , a.PrDetailRemark, a.MtlInventory, a.MtlInventoryQty, a.OriPrSequence, a.ClosureStatus
                                    , a.ModiReason
                                    , b.MtlItemNo
                                    , c.InventoryNo, c.InventoryName
                                    , d.UomNo
                                    , ISNULL(e.SupplierNo, '') SupplierNo, ISNULL(e.SupplierName, '') SupplierName, ISNULL(e.HideSupplier, 'N') HideSupplier
                                    , f.SoSequence
                                    , g.SoErpPrefix, g.SoErpNo
                                    , h.PrId
                                    FROM SCM.PrmDetail a
                                    INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
                                    INNER JOIN SCM.Inventory c ON a.InventoryId = c.InventoryId
                                    INNER JOIN PDM.UnitOfMeasure d ON a.PrUomId = d.UomId
                                    LEFT JOIN SCM.Supplier e ON a.SupplierId = e.SupplierId
                                    LEFT JOIN SCM.SoDetail f ON a.SoDetailId = f.SoDetailId
                                    LEFT JOIN SCM.SaleOrder g ON f.SoId = g.SoId
                                    INNER JOIN SCM.PrModification h ON a.PrmId = h.PrmId
                                    WHERE a.PrmId = @PrmId
                                    ORDER BY a.PrmSequence";
                            dynamicParameters.Add("PrmId", PrmId);

                            var prmDetailResult = sqlConnection.Query(sql, dynamicParameters);
                            if (prmDetailResult.Count() <= 0) throw new SystemException("請購變更單身資料錯誤!");
                            #endregion

                            #region //組單身資料
                            JArray tabPRM_Line_Data = new JArray();
                            string ProjectName = "";
                            string HideSupplier = prmDetailResult.FirstOrDefault().HideSupplier;
                            double maxPrUnitPrice = 0;
                            List<int> prDetailList = new List<int>();
                            foreach (var item in prmDetailResult)
                            {
                                prDetailList.Add(item.PrDetailId);

                                double localPrUnitPrice = item.PrUnitPrice * Convert.ToDouble(item.PrExchangeRate);

                                #region //取得最大單價
                                if (localPrUnitPrice > maxPrUnitPrice)
                                {
                                    maxPrUnitPrice = localPrUnitPrice;
                                }
                                #endregion

                                #region //檢核特定相關卡控(目前僅晶彩邏輯)
                                if (CurrentCompany == 4)
                                {
                                    #region //單別3109、3108特別卡控
                                    if (PrErpPrefix == "3108" || PrErpPrefix == "3109")
                                    {
                                        #region //請購單價不能超過1000(本幣)
                                        if (localPrUnitPrice > 1000) throw new SystemException("違反請購單別【3108、3109】卡控:<br>請購單價不能超過【1000】!!<br>此次請購單價【" + localPrUnitPrice + "】已超過!!");
                                        #endregion

                                        #region //此修改後單身與原請購單單身合計金額(本幣)不能超過20000
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT 
                                                    SUM(a.PrPriceTw) +
                                                    COALESCE(
                                                        (
                                                            SELECT SUM(x.PrPriceTw)
                                                            FROM SCM.PrDetail x 
                                                            WHERE x.PrId = b.PrId
                                                            AND x.PrDetailId NOT IN (
                                                                SELECT PrDetailId
                                                                FROM SCM.PrmDetail xa
                                                                WHERE xa.PrmId = a.PrmId
                                                            )
                                                        ), 0) TotalPrice
                                                FROM SCM.PrmDetail a
                                                INNER JOIN SCM.PrModification b ON a.PrmId = b.PrmId 
                                                WHERE a.PrmId = @PrmId
                                                GROUP BY a.PrmId, b.PrId";
                                        dynamicParameters.Add("PrmId", PrmId);

                                        var TotalPriceResult = sqlConnection.Query(sql, dynamicParameters);

                                        double totalPrice = 0;
                                        foreach (var item2 in TotalPriceResult)
                                        {
                                            totalPrice = item2.TotalPrice;
                                            if (totalPrice > 20000) throw new SystemException("違反請購單別【3108、3109】卡控:<br>請購合計金額不能超過20000!!<br>目前請購合計金額【" + (totalPrice) + "】已超過!!");
                                        }
                                        #endregion
                                    }
                                    #endregion
                                }
                                #endregion

                                #region //取得專案資料
                                if (item.Project != null)
                                {
                                    sql = @"SELECT LTRIM(RTRIM(NB001)) Project,LTRIM(RTRIM(NB002)) ProjectName
                                            FROM CMSNB
                                            WHERE NB001 = @Project";
                                    dynamicParameters.Add("Project", item.Project);
                                    var projectResult = sqlConnection2.Query(sql, dynamicParameters);
                                    foreach (var item2 in projectResult)
                                    {
                                        ProjectName = item2.ProjectName;
                                    }
                                }
                                #endregion

                                #region //處理MtlInventoryStatus
                                string mtlInventoryStatus = "目前尚無資料";
                                if (item.MtlInventory != "目前尚無資料")
                                {
                                    JObject jo = JObject.Parse(item.MtlInventory);
                                    JToken jtoken = jo["data"];

                                    for (int i = 0; i < jtoken.Count(); i++)
                                    {
                                        mtlInventoryStatus += jtoken[i]["WAREHOUSE_NO"] + ":" + jtoken[i]["WAREHOUSE_NAME"] + ":" + jtoken[i]["WAREHOUSE_QTY"] + ";";
                                    }
                                }
                                #endregion

                                #region //若有訂單，計算此訂單下物料相關資料
                                double soPrTotalAmount = 0;
                                double soTotalAmount = 0;
                                if (item.SoDetailId != null && item.SoDetailId > 0)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.SoId
                                            , b.MtlItemNo
                                            FROM SCM.SoDetail a
                                            INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
                                            WHERE a.SoDetailId = @SoDetailId";
                                    dynamicParameters.Add("SoDetailId", item.SoDetailId);

                                    var SoDetailResult = sqlConnection.Query(sql, dynamicParameters);

                                    int SoId = SoDetailResult.FirstOrDefault().SoId;

                                    #region //此訂單下所有物料有開請購單的金額合計(包括下層物料品號)
                                    foreach (var item2 in SoDetailResult)
                                    {
                                        string mtlItemNo = item2.MtlItemNo;

                                        #region //遞迴算出所有相關品號
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"WITH RecursiveCTE AS (
                                                    SELECT LTRIM(RTRIM(d.MB001)) MtlItemNo, LTRIM(RTRIM(d.MB002)) MtlItemName, LTRIM(RTRIM(d.MB003)) MtlItemSpec
                                                    FROM BOMMD a 
                                                    INNER JOIN BOMMC b ON b.MC001 = a.MD001
                                                    INNER JOIN INVMB c ON c.MB001 = b.MC001
                                                    INNER JOIN INVMB d ON d.MB001 = a.MD003
                                                    WHERE c.MB001 = @MtlItemNo
                                                    AND GETDATE() >= ISNULL(CONVERT(DATETIME, NULLIF(LTRIM(RTRIM(d.MB030)), ''), 120), GETDATE())
                                                    AND GETDATE() <= ISNULL(CONVERT(DATETIME, NULLIF(LTRIM(RTRIM(d.MB031)), ''), 120), GETDATE())
                                                    UNION ALL
                                                    SELECT LTRIM(RTRIM(xd.MB001)) MtlItemNo, LTRIM(RTRIM(xd.MB002)) MtlItemName, LTRIM(RTRIM(xd.MB003)) MtlItemSpec
                                                    FROM BOMMD x
                                                    INNER JOIN BOMMC xb ON xb.MC001 = x.MD001
                                                    INNER JOIN RecursiveCTE xa ON xb.MC001 = xa.MtlItemNo
                                                    INNER JOIN INVMB xc ON xc.MB001 = xb.MC001
                                                    INNER JOIN INVMB xd ON xd.MB001 = x.MD003
                                                    WHERE GETDATE() >= ISNULL(CONVERT(DATETIME, NULLIF(LTRIM(RTRIM(xd.MB030)), ''), 120), GETDATE())
                                                    AND GETDATE() <= ISNULL(CONVERT(DATETIME, NULLIF(LTRIM(RTRIM(xd.MB031)), ''), 120), GETDATE())
                                                )
                                                SELECT MtlItemNo, MtlItemName, MtlItemSpec
                                                , 2 AS SortOrder
                                                FROM RecursiveCTE
                                                UNION ALL
                                                SELECT LTRIM(RTRIM(x.MB001)) MtlItemNo, LTRIM(RTRIM(x.MB002)) MtlItemName, LTRIM(RTRIM(x.MB003)) MtlItemSpec
                                                , 1 AS SortOrder
                                                FROM INVMB x 
                                                WHERE x.MB001 = @MtlItemNo
                                                AND GETDATE() >= ISNULL(CONVERT(DATETIME, NULLIF(LTRIM(RTRIM(x.MB030)), ''), 120), GETDATE())
                                                AND GETDATE() <= ISNULL(CONVERT(DATETIME, NULLIF(LTRIM(RTRIM(x.MB031)), ''), 120), GETDATE())
                                                ORDER BY SortOrder";
                                        dynamicParameters.Add("MtlItemNo", mtlItemNo);

                                        var mtlItemResult = sqlConnection2.Query(sql, dynamicParameters);
                                        #endregion

                                        #region //根據此次遞迴結果，計算有註冊此訂單的請購單
                                        foreach (var item3 in mtlItemResult)
                                        {
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"SELECT ISNULL(SUM(a.PrPriceTw), 0) TotalAmount
                                                    FROM SCM.PrDetail a 
                                                    INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
                                                    WHERE a.SoDetailId = @SoDetailId
                                                    AND b.MtlItemNo = @MtlItemNo";
                                            dynamicParameters.Add("SoDetailId", item.SoDetailId);
                                            dynamicParameters.Add("MtlItemNo", item3.MtlItemNo);

                                            var prTotalAmountResult = sqlConnection.Query(sql, dynamicParameters);

                                            foreach (var item4 in prTotalAmountResult)
                                            {
                                                soPrTotalAmount += item4.TotalAmount;
                                            }
                                        }
                                        #endregion
                                    }
                                    #endregion

                                    #region //此訂單所有物料訂單金額合計
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT SUM(a.Amount) TotalAmount
                                            FROM SCM.SoDetail a 
                                            WHERE a.SoId = @SoId";
                                    dynamicParameters.Add("SoId", SoId);

                                    var soTotalAmountResult = sqlConnection.Query(sql, dynamicParameters);

                                    soTotalAmount = soTotalAmountResult.FirstOrDefault().TotalAmount;
                                    #endregion
                                }
                                #endregion

                                tabPRM_Line_Data.Add(JObject.FromObject(new
                                {
                                    PRNumber_SN = item.PrmSequence,
                                    Ori_PRNumber_SN = item.OriPrSequence,
                                    ItemNo = item.MtlItemNo,
                                    ItemName = item.PrMtlItemName,
                                    Spec = item.PrMtlItemSpec,
                                    WareHouse = item.InventoryNo + "-" + item.InventoryName,
                                    Qty = item.PrQty.ToString(),
                                    Unit = item.UomNo,
                                    RequiredDate = item.DemandDate,
                                    Supplier = item.SupplierNo + "-" + item.SupplierName,
                                    Currency = item.PrCurrency,
                                    ExchangeRate = item.PrExchangeRate,
                                    Price = item.PrUnitPrice.ToString(),
                                    Amount = item.PrPrice.ToString(),
                                    AmountInLocalCurrency = item.PrPriceTw.ToString(),
                                    UrgentMaterial = item.UrgentMtl,
                                    MRP = item.ProductionPlan,
                                    ClosureNo = item.ClosureStatus,
                                    Project = item.Project != null ? item.Project + " (" + ProjectName + ")" : "",
                                    item.ModiReason,
                                    LineRemark = item.PrDetailRemark,
                                    ReferenceDoc = item.SoErpPrefix != null ? item.SoErpPrefix + "-" + item.SoErpNo + "-" + item.SoSequence : "",
                                    MtlInventoryStatus = mtlInventoryStatus,
                                    item.MtlInventoryQty,
                                    SoPrTotalAmount = soPrTotalAmount.ToString(),
                                    SoTotalAmount = soTotalAmount.ToString(),
                                    MtlCostRatio = (soPrTotalAmount / soTotalAmount).ToString() + "%"
                                }));
                            }
                            #endregion

                            #region //組檔案JSON
                            JArray attachFilePath = new JArray();
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.FileId
                                    , b.FileName, b.FileExtension, b.FileSize, b.FileContent
                                    FROM SCM.PrmFile a
                                    INNER JOIN BAS.[File] b ON a.FileId = b.FileId
                                    WHERE a.PrmId = @PrmId";
                            dynamicParameters.Add("PrmId", PrmId);

                            var prmFileResult = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item in prmFileResult)
                            {
                                attachFilePath.Add(JObject.FromObject(new
                                {
                                    fileId = item.FileId,
                                    fileName = item.FileName + item.FileExtension
                                }));
                            }
                            #endregion

                            #region //依公司別取得ProId
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.TypeName ProId
                                    FROM BAS.[Type] a
                                    WHERE a.TypeSchema = 'BPM.PrmProId'
                                    AND a.TypeNo = @CompanyNo";
                            dynamicParameters.Add("CompanyNo", CompanyNo);

                            var ProIdResult = sqlConnection.Query(sql, dynamicParameters);

                            if (ProIdResult.Count() <= 0) throw new SystemException("此公司別【" + CompanyNo + "】尚未建立拋轉起單碼，請聯繫資訊人員!!");

                            string proId = "";
                            foreach (var item in ProIdResult)
                            {
                                proId = item.ProId;
                            }
                            #endregion

                            #region //組請購變更BPM資料
                            //string proId = "PRO03421647434184462"; 原本流程ID
                            //string proId = "PRO05741677634833256";
                            string memId = BpmUserId;
                            string rolId = BpmRoleId;
                            string startMethod = "NoOpFirst";
                            string ErpNoEdition = PrErpPrefix + "-" + PrErpNo + "-" + PrEdition;

                            JObject artInsAppData = JObject.FromObject(new
                            {
                                Title = "ERP請購單-版次 = " + ErpNoEdition,
                                PRNumber_ERP = PrErpPrefix + "-" + PrErpNo,
                                Edition = PrmEdition,
                                PRNumber_MES = PrNo,
                                CreateDate = DocDate,
                                Requisitioner = BpmUserNo,
                                RequisitionerName = UserName,
                                RequisitionDeptID = DepartmentNo,
                                RequisitionDept = DepartmentName,
                                RequisitionDate = PrDate,
                                ChangeDate = ModiDate,
                                ChangeReason = ModiReason,
                                HeaderRemark = PrmRemark,
                                AmountInLocalCurrencyTotal = Amount.ToString(),
                                tabPRM_Line_Data = JsonConvert.SerializeObject(tabPRM_Line_Data),
                                AttachFilePath = JsonConvert.SerializeObject(attachFilePath),
                                dbTable = "SCM.PrModification",
                                mesID = PrmId,
                                company = CompanyNo,
                                HideSupplier,
                                PRNumber_FormType = PrErpPrefix,
                                priority = Priority,
                                MaxUnitPrice = maxPrUnitPrice
                            });
                            #endregion

                            string sData = BpmHelper.PostFormToBpm(token, proId, memId, rolId, startMethod, artInsAppData, BpmServerPath);

                            if (sData == "true")
                            {
                                #region //更改BPM拋轉狀態
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE SCM.PrModification SET
                                        PrmStatus = 'P',
                                        BpmTransferStatus = 'Y',
                                        BpmTransferUserId = @BpmTransferUserId,
                                        BpmTransferDate = @BpmTransferDate,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE PrmId = @PrmId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        BpmTransferUserId = CreateBy,
                                        BpmTransferDate = CreateDate,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        PrmId
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }
                            else
                            {
                                #region //更改BPM拋轉狀態
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE SCM.PrModification SET
                                        PrmStatus = 'R',
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE PrmId = @PrmId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        PrmId
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                #endregion

                                throw new SystemException("請購變更單拋轉BPM失敗!");
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

        #region //UpdatePrmTransferErp -- 拋轉請購變更單據至ERP -- Ann 2023-02-09//
        public string UpdatePrmTransferErp(int PrmId, string BpmNo, string PrmStatus, string ComfirmUser, string ErpFolderRoot, string CompanyNo)
        {
            try
            {
                int rowsAffected = 0;
                string ErpDocPath = "";
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (CompanyNo.Length <= 0) throw new SystemException("公司別不能為空!!");

                        //CompanyNo = "JMO"; //待BPM可以傳公司別後再修正
                        #region //確認公司別DB
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.ErpNo, a.ErpDb
                                FROM BAS.Company a
                                WHERE CompanyNo = @CompanyNo";
                        dynamicParameters.Add("CompanyNo", CompanyNo);

                        var companyResult = sqlConnection.Query(sql, dynamicParameters);

                        if (companyResult.Count() <= 0) throw new SystemException("公司別錯誤!");

                        string ErpNo = "";
                        foreach (var item in companyResult)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            ErpNo = item.ErpNo;
                        }
                        #endregion

                        #region //確認核單者資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.UserId ConfirmUserId
                                FROM BAS.[User] a
                                WHERE a.UserNo = @ComfirmUser";
                        dynamicParameters.Add("ComfirmUser", ComfirmUser);

                        var confirmUserResult = sqlConnection.Query(sql, dynamicParameters);

                        if (confirmUserResult.Count() <= 0) throw new SystemException("核單者資料錯誤!");

                        int ConfirmUserId = -1;
                        foreach (var item in confirmUserResult)
                        {
                            ConfirmUserId = item.ConfirmUserId;
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            using (SqlConnection sqlConnection3 = new SqlConnection(ErpSysDbConnectionStrings))
                            {
                                #region //判斷請購變更單資料是否正確
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.PrId, a.BpmTransferStatus, a.ConfirmStatus, a.PrmRemark, a.TotalQty, a.Amount, a.Edition
                                        , a.WholeClosureStatus, a.ModiReason, a.ConfirmStatus, a.ConfirmStatus, a.DepartmentId, a.UserId
                                        , a.BudgetDepartmentId
                                        , FORMAT(a.ModiDate, 'yyyyMMdd') ModiDate
                                        , FORMAT(a.CreateDate, 'yyyyMMdd') CreateDate
                                        , FORMAT(a.CreateDate, 'HH:mm:ss') CreateTime
                                        , FORMAT(a.DocDate, 'yyyyMMdd') DocDate
                                        , FORMAT(a.DocDate, 'yyyy-MM-dd') MesDocDate
                                        , b.UserNo
                                        , c.DepartmentNo
                                        , d.UserNo PrUserNo
                                        , ISNULL(e.UserNo, '') ConfirmUserNo
                                        , f.PrErpPrefix, f.PrErpNo, FORMAT(f.PrDate, 'yyyyMMdd') OriPrDate, f.Edition OriEdition
                                        , f.PrRemark OriPrRemark, f.Amount OriAmount, f.PrNo
                                        , ISNULL(g.DepartmentNo, '') BudgetDepartmentNo
                                        , h.DepartmentNo OriDepartmentNo
                                        , i.UserNo OriPrUserNo
                                        , ISNULL(j.DepartmentNo, '') OriBudgetDepartmentNo
                                        FROM SCM.PrModification a
                                        INNER JOIN BAS.[User] b ON a.CreateBy = b.UserId
                                        INNER JOIN BAS.Department c ON a.DepartmentId = c.DepartmentId
                                        INNER JOIN BAS.[User] d ON a.UserId = d.UserId
                                        LEFT JOIN BAS.[User] e ON a.ConfirmUserId = e.UserId
                                        INNER JOIN SCM.PurchaseRequisition f ON a.PrId = f.PrId
                                        LEFT JOIN BAS.Department g ON a.BudgetDepartmentId = g.DepartmentId
                                        INNER JOIN BAS.Department h ON f.DepartmentId = h.DepartmentId
                                        INNER JOIN BAS.[User] i ON f.UserId = i.UserId
                                        LEFT JOIN BAS.Department j ON f.BudgetDepartmentId = j.DepartmentId
                                        WHERE a.PrmId = @PrmId";
                                dynamicParameters.Add("PrmId", PrmId);

                                var result = sqlConnection.Query(sql, dynamicParameters);
                                if (result.Count() <= 0) throw new SystemException("請購單資料錯誤!");

                                int PrId = -1;
                                string UserNo = "";
                                string CreateDate = "";
                                string CreateTime = "";
                                string PrDate = "";
                                string DepartmentNo = "";
                                string PrmRemark = "";
                                double TotalQty = -1;
                                string PrUserNo = "";
                                string DocDate = "";
                                string MesDocDate = "";
                                string ConfirmUserNo = "";
                                string Edition = "";
                                string LockStaus = "";
                                string PrNo = "";
                                string PrErpPrefix = "";
                                string PrErpNo = "";
                                string ModiDate = "";
                                string WholeClosureStatus = "";
                                string ModiReason = "";
                                string ConfirmStatus = "";
                                string OriPrDate = "";
                                double Amount = -1;
                                string BudgetDepartmentNo = "";
                                string OriEdition = "";
                                string OriDepartmentNo = "";
                                string OriPrUserNo = "";
                                string OriPrRemark = "";
                                double OriAmount = -1;
                                string OriBudgetDepartmentNo = "";
                                int DepartmentId = -1;
                                int UserId = -1;
                                int? BudgetDepartmentId = -1;
                                foreach (var item in result)
                                {
                                    if (item.BpmTransferStatus != "Y") throw new SystemException("請購變更單尚未拋轉BPM，無法拋轉ERP!");
                                    if (item.ConfirmStatus != "N" && PrmStatus == "Y") throw new SystemException("請購單變更已拋轉ERP，無法重複拋轉!");
                                    PrId = item.PrId;
                                    UserNo = item.UserNo;
                                    CreateDate = item.CreateDate;
                                    CreateTime = item.CreateTime;
                                    PrDate = item.PrDate;
                                    DepartmentNo = item.DepartmentNo;
                                    PrmRemark = item.PrmRemark;
                                    TotalQty = item.TotalQty;
                                    PrUserNo = item.PrUserNo;
                                    DocDate = item.DocDate;
                                    ConfirmUserNo = item.ConfirmUserNo;
                                    Amount = item.Amount;
                                    Edition = item.Edition;
                                    LockStaus = item.LockStaus;
                                    PrNo = item.PrNo;
                                    PrErpPrefix = item.PrErpPrefix;
                                    PrErpNo = item.PrErpNo;
                                    ModiDate = item.ModiDate;
                                    WholeClosureStatus = item.WholeClosureStatus;
                                    ModiReason = item.ModiReason;
                                    ConfirmStatus = item.ConfirmStatus;
                                    OriPrDate = item.OriPrDate;
                                    Amount = item.Amount;
                                    BudgetDepartmentNo = item.BudgetDepartmentNo;
                                    OriEdition = item.OriEdition;
                                    OriDepartmentNo = item.OriDepartmentNo;
                                    OriPrUserNo = item.OriPrUserNo;
                                    OriPrRemark = item.OriPrRemark;
                                    OriAmount = item.OriAmount;
                                    OriBudgetDepartmentNo = item.OriBudgetDepartmentNo;
                                    DepartmentId = item.DepartmentId;
                                    UserId = item.UserId;
                                    BudgetDepartmentId = item.BudgetDepartmentId;
                                    MesDocDate = item.MesDocDate;
                                }
                                #endregion

                                #region //該請購單是否有開立採購單
                                dynamicParameters = new DynamicParameters();
                                sql = @"select * 
                                    from PURTD
                                    where TD026 = @PrErpPrefix
                                    and TD027 = @PrErpNo";
                                dynamicParameters.Add("PrErpPrefix", PrErpPrefix);
                                dynamicParameters.Add("PrErpNo", PrErpNo);

                                var PoERPResult = sqlConnection2.Query(sql, dynamicParameters);
                                if (PoERPResult.Count() > 0) throw new SystemException("該請購單已有採購單，無法開立變更單");

                                #endregion


                                #region //取得單身資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.PrmSequence, a.PrQty, a.PrDetailRemark, ISNULL(a.PoQty, 0) PoQty, a.ProductionPlan
                                        , ISNULL(a.PoCurrency, '') PoCurrency, ISNULL(a.PoUnitPrice, 0) PoUnitPrice
                                        , ISNULL(a.PoPrice, 0) PoPrice, a.LockStaus, a.PoStaus, a.PrDetailId, a.InventoryId
                                        , ISNULL(a.PoRemark, '') PoRemark, a.Taxation, a.UrgentMtl, a.ClosureStatus
                                        , ISNULL(a.Project, '') Project, a.PrExchangeRate, a.PrPriceTw, a.MtlItemId
                                        , ISNULL(a.BudgetDepartmentNo, '') BudgetDepartmentNo, a.PrUomId, a.SupplierId
                                        , ISNULL(a.BudgetDepartmentSubject, '') BudgetDepartmentSubject, a.PrUnitPrice
                                        , a.PrCurrency, a.PrPrice, a.TaxNo, a.TradeTerm, a.BusinessTaxRate, a.DetailMultiTax
                                        , ISNULL(a.DiscountRate, 0) DiscountRate, ISNULL(a.DiscountAmount, 0) DiscountAmount
                                        , FORMAT(a.DemandDate, 'yyyyMMdd') DemandDate, ISNULL(FORMAT(a.DeliveryDate, 'yyyyMMdd'), '') DeliveryDate
                                        , a.ConfirmStatus, a.ModiReason, a.OriPrSequence, a.OriMtlItemId, a.OriPrMtlItemName, a.OriPrMtlItemSpec
                                        , a.OriPrQty, a.OriSupplierId, FORMAT(a.OriDemandDate, 'yyyyMMdd') OriDemandDate, ISNULL(a.OriPoQty, 0) OriPoQty, ISNULL(a.OriPoCurrency, '') OriPoCurrency
                                        , ISNULL(a.OriPoUnitPrice, 0) OriPoUnitPrice, ISNULL(a.OriPoPrice, 0) OriPoPrice, FORMAT(a.OriDeliveryDate, 'yyyyMMdd') OriDeliveryDate
                                        , a.OriTaxation, a.OriUrgentMtl, ISNULL(a.OriPoRemark, '') OriPoRemark, a.OriClosureStatus
                                        , a.OriLockStaus, a.OriPoStaus, ISNULL(a.OriPrDetailRemark, '') OriPrDetailRemark, ISNULL(a.OriProject, '') OriProject
                                        , ISNULL(a.OriPrExchangeRate, 0) OriPrExchangeRate, ISNULL(a.OriPrPriceTw, 0) OriPrPriceTw, ISNULL(a.OriBudgetDepartmentNo, '') OriBudgetDepartmentNo
                                        , ISNULL(a.OriBudgetDepartmentSubject, '') OriBudgetDepartmentSubject, ISNULL(a.OriPrUnitPrice, 0) OriPrUnitPrice, a.OriPrCurrency
                                        , a.OriPrPrice, ISNULL(a.OriTaxNo, '') OriTaxNo, ISNULL(a.OriTradeTerm, '') OriTradeTerm, ISNULL(a.OriDetailMultiTax, '') OriDetailMultiTax
                                        , ISNULL(a.OriBusinessTaxRate, 0) OriBusinessTaxRate, ISNULL(a.OriDiscountRate, 0) OriDiscountRate, ISNULL(a.OriDiscountAmount, 0) OriDiscountAmount
                                        , a.PrMtlItemName, a.PrMtlItemSpec, a.MtlInventoryQty
                                        , b.MtlItemNo
                                        , c.InventoryNo
                                        , d.UomNo
                                        , ISNULL(e.SupplierNo, '') SupplierNo
                                        , f.SoSequence
                                        , g.SoErpPrefix, g.SoErpNo
                                        , ISNULL(h.UserNo, '') PoUserNo
                                        , ISNULL(i.UomNo, '') PoUomNo
                                        , j.UomNo OriPrUomNo
                                        , k.InventoryNo OriInventoryNo
                                        , ISNULL(l.UserNo, '') OriPoUserNo
                                        , ISNULL(m.UomNo, '') OriPoUomNo
                                        , n.SoSequence OriSoSequence
                                        , o.SoErpPrefix OriSoErpPrefix, o.SoErpNo OriSoErpNo
                                        , p.MtlItemNo OriMtlItemNo
                                        , q.SupplierNo OriSupplierNo
                                        , s.PrErpPrefix, s.PrErpNo
                                        FROM SCM.PrmDetail a
                                        INNER JOIN PDM.MtlItem b ON a.MtlItemId = b.MtlItemId
                                        INNER JOIN SCM.Inventory c ON a.InventoryId = c.InventoryId
                                        INNER JOIN PDM.UnitOfMeasure d ON a.PrUomId = d.UomId
                                        LEFT JOIN SCM.Supplier e ON a.SupplierId = e.SupplierId
                                        LEFT JOIN SCM.SoDetail f ON a.SoDetailId = f.SoDetailId
                                        LEFT JOIN SCM.SaleOrder g ON f.SoId = g.SoId
                                        LEFT JOIN BAS.[User] h ON a.PoUserId = h.UserId
                                        LEFT JOIN PDM.UnitOfMeasure i ON a.PoUomId = i.UomId
                                        INNER JOIN PDM.UnitOfMeasure j ON a.OriPrUomId = j.UomId
                                        INNER JOIN SCM.Inventory k ON a.OriInventoryId = k.InventoryId
                                        LEFT JOIN BAS.[User] l ON a.OriPoUserId =l.UserId
                                        LEFT JOIN PDM.UnitOfMeasure m ON a.OriPoUomId = m.UomId
                                        LEFT JOIN SCM.SoDetail n ON a.OriSoDetailId = n.SoDetailId
                                        LEFT JOIN SCM.SaleOrder o ON n.SoId = o.SoId
                                        LEFT JOIN PDM.MtlItem p ON a.OriMtlItemId = p.MtlItemId
                                        LEFT JOIN SCM.Supplier q ON a.OriSupplierId = q.SupplierId
                                        INNER JOIN SCM.PrModification r ON a.PrmId = r.PrmId
                                        INNER JOIN SCM.PurchaseRequisition s ON r.PrId = s.PrId
                                        WHERE a.PrmId = @PrmId";
                                dynamicParameters.Add("PrmId", PrmId);

                                var prmDetailResult = sqlConnection.Query(sql, dynamicParameters);
                                if (prmDetailResult.Count() <= 0) throw new SystemException("請購變更單身資料錯誤!");
                                #endregion

                                #region //查詢廠別
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT LTRIM(RTRIM(MB001)) MB001 FROM CMSMB";
                                var CMSMBResult = sqlConnection2.Query(sql, dynamicParameters);

                                if (CMSMBResult.Count() > 1) throw new SystemException("廠別數有多個，請與資訊人員確認!!");

                                string TU011 = "";
                                foreach (var item in CMSMBResult)
                                {
                                    TU011 = item.MB001;
                                }
                                #endregion

                                #region //管理欄位定義
                                string COMPANY = ErpNo;
                                string CREATOR = UserNo;
                                string CREATE_DATE = CreateDate;
                                string MODIFIER = ComfirmUser;
                                string MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
                                string FLAG = "1";
                                string CREATE_TIME = CreateTime;
                                string CREATE_AP = "MES";
                                string CREATE_PRID = "ERPJ02";
                                string MODI_TIME = DateTime.Now.ToString("HH:mm:ss");
                                string MODI_AP = "MES";
                                string MODI_PRID = "ERPJ02";
                                #endregion

                                string USR_GROUP = "";
                                if (PrmStatus == "Y") //單據新增
                                {
                                    #region //審核ERP權限並取得USR_GROUP
                                    //USR_GROUP = "ZY07";
                                    //USR_GROUP = BaseHelper.CheckErpAuthority(UserNo, sqlConnection2, "PURI16", "CREATE");
                                    #endregion

                                    #region //取得USR_GROUP
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT LTRIM(RTRIM(a.MF004)) USR_GROUP
                                                   FROM ADMMF a
                                                   WHERE MF001 = @MF001";
                                    dynamicParameters.Add("MF001", UserNo);

                                    var ADMMFResult = sqlConnection2.Query(sql, dynamicParameters);

                                    if (ADMMFResult.Count() <= 0) throw new SystemException("查無ERP帳號權限，無法進行拋轉及核單!");

                                    foreach (var item in result)
                                    {
                                        USR_GROUP = item.USR_GROUP;
                                    }
                                    #endregion

                                    #region //INSERT PURTU 請購變更單單頭

                                    #region //寫入前判斷請購變更單是否存在未確認版次
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM PURTU a
                                            WHERE a.TU012 = 'N'
                                            AND a.TU001 = @TU001
                                            AND a.TU002 = @TU002";
                                    dynamicParameters.Add("TU001", PrErpPrefix);
                                    dynamicParameters.Add("TU002", PrErpNo);

                                    var checkPrmConfirmResult = sqlConnection2.Query(sql, dynamicParameters);

                                    if (checkPrmConfirmResult.Count() > 0) throw new SystemException("此單據存在尚未確認之較小版次變更紀錄，請先確認!!");
                                    #endregion

                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO PURTU (COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE
                                            , FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID
                                            , TU001, TU002, TU003, TU004, TU005, TU006, TU007, TU008
                                            , TU009, TU010, TU011, TU012, TU013, TU014, TU015, TU016
                                            , TU017, TU018, TU019, TU020, TU021, TU022, TU023, TU024
                                            , TU025, TU103, TU106, TU107, TU108, TU109, TU120, TU121
                                            , TU122, TU123, TU124, TU125, UDF01, UDF02)
                                            VALUES (@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE
                                            , @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID, @MODI_TIME, @MODI_AP, @MODI_PRID
                                            , @TU001, @TU002, @TU003, @TU004, @TU005, @TU006, @TU007, @TU008
                                            , @TU009, @TU010, @TU011, @TU012, @TU013, @TU014, @TU015, @TU016
                                            , @TU017, @TU018, @TU019, @TU020, @TU021, @TU022, @TU023, @TU024
                                            , @TU025, @TU103, @TU106, @TU107, @TU108, @TU109, @TU120, @TU121
                                            , @TU122, @TU123, @TU124, @TU125, @UDF01, @UDF02)";

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
                                          TU001 = PrErpPrefix,
                                          TU002 = PrErpNo,
                                          TU003 = Edition,
                                          TU004 = ModiDate,
                                          TU005 = DocDate,
                                          TU006 = DepartmentNo,
                                          TU007 = PrUserNo,
                                          TU008 = WholeClosureStatus,
                                          TU009 = ModiReason,
                                          TU010 = 9,
                                          TU011,
                                          TU012 = PrmStatus,
                                          TU013 = ComfirmUser,
                                          TU014 = "3",
                                          TU015 = 0,
                                          TU016 = 0,
                                          TU017 = OriPrDate,
                                          TU018 = PrmRemark,
                                          TU019 = Amount,
                                          TU020 = BudgetDepartmentNo,
                                          TU021 = 0,
                                          TU022 = 0,
                                          TU023 = "",
                                          TU024 = "",
                                          TU025 = "",
                                          TU103 = OriEdition,
                                          TU106 = OriDepartmentNo,
                                          TU107 = OriPrUserNo,
                                          TU108 = OriPrRemark,
                                          TU109 = OriAmount,
                                          TU120 = OriBudgetDepartmentNo,
                                          TU121 = 0,
                                          TU122 = 0,
                                          TU123 = "",
                                          TU124 = "",
                                          TU125 = "",
                                          UDF01 = PrNo,
                                          UDF02 = BpmNo
                                      });
                                    var insertResult = sqlConnection2.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult.Count();

                                    #endregion

                                    foreach (var item in prmDetailResult)
                                    {
                                        #region //INSERT PURTV 請購變更單單身
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO PURTV (COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE
                                                , FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID
                                                , TV001, TV002, TV003, TV004, TV005, TV006, TV007, TV008
                                                , TV009, TV010, TV011, TV012, TV013, TV014, TV015, TV016
                                                , TV017, TV018, TV019, TV020, TV021, TV022, TV023, TV024
                                                , TV025, TV026, TV027, TV028, TV029, TV030, TV031, TV032
                                                , TV033, TV034, TV035, TV036, TV037, TV038, TV039, TV040
                                                , TV041, TV042, TV043, TV044, TV045, TV046, TV047, TV048
                                                , TV049, TV050, TV051, TV052, TV053, TV054, TV055, TV056
                                                , TV057, TV058, TV059, TV060, TV061, TV104, TV105, TV106
                                                , TV107, TV108, TV109, TV110, TV111, TV112, TV113, TV114
                                                , TV115, TV116, TV117, TV118, TV119, TV120, TV121, TV122
                                                , TV123, TV124, TV125, TV126, TV127, TV128, TV129, TV130
                                                , TV131, TV132, TV133, TV134, TV135, TV136, TV137, TV140
                                                , TV141, TV142, TV143, TV144, TV145, TV146, TV147, TV148
                                                , TV149, TV150, TV151, TV152, TV153, TV154, TV155, TV156
                                                , TV157, TV158, TV159, TV160, TV161)
                                                VALUES (@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE
                                                , @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID, @MODI_TIME, @MODI_AP, @MODI_PRID
                                                , @TV001, @TV002, @TV003, @TV004, @TV005, @TV006, @TV007, @TV008
                                                , @TV009, @TV010, @TV011, @TV012, @TV013, @TV014, @TV015, @TV016
                                                , @TV017, @TV018, @TV019, @TV020, @TV021, @TV022, @TV023, @TV024
                                                , @TV025, @TV026, @TV027, @TV028, @TV029, @TV030, @TV031, @TV032
                                                , @TV033, @TV034, @TV035, @TV036, @TV037, @TV038, @TV039, @TV040
                                                , @TV041, @TV042, @TV043, @TV044, @TV045, @TV046, @TV047, @TV048
                                                , @TV049, @TV050, @TV051, @TV052, @TV053, @TV054, @TV055, @TV056
                                                , @TV057, @TV058, @TV059, @TV060, @TV061, @TV104, @TV105, @TV106
                                                , @TV107, @TV108, @TV109, @TV110, @TV111, @TV112, @TV113, @TV114
                                                , @TV115, @TV116, @TV117, @TV118, @TV119, @TV120, @TV121, @TV122
                                                , @TV123, @TV124, @TV125, @TV126, @TV127, @TV128, @TV129, @TV130
                                                , @TV131, @TV132, @TV133, @TV134, @TV135, @TV136, @TV137, @TV140
                                                , @TV141, @TV142, @TV143, @TV144, @TV145, @TV146, @TV147, @TV148
                                                , @TV149, @TV150, @TV151, @TV152, @TV153, @TV154, @TV155, @TV156
                                                , @TV157, @TV158, @TV159, @TV160, @TV161)";

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
                                              TV001 = PrErpPrefix,
                                              TV002 = PrErpNo,
                                              TV003 = Edition,
                                              TV004 = item.PrmSequence,
                                              TV005 = item.MtlItemNo,
                                              TV006 = item.PrMtlItemName,
                                              TV007 = item.PrMtlItemSpec,
                                              TV008 = item.UomNo,
                                              TV009 = item.InventoryNo,
                                              TV010 = item.PrQty,
                                              TV011 = item.SupplierNo,
                                              TV012 = item.DemandDate,
                                              TV013 = item.PoUserNo,
                                              TV014 = item.PoQty,
                                              TV015 = item.PoUomNo,
                                              TV016 = item.PoCurrency,
                                              TV017 = item.PoUnitPrice,
                                              TV018 = item.PoPrice,
                                              TV019 = item.DeliveryDate,
                                              TV020 = "",
                                              TV021 = item.Taxation,
                                              TV022 = "",
                                              TV023 = "",
                                              TV024 = item.SoErpPrefix,
                                              TV025 = item.SoErpNo,
                                              TV026 = item.SoSequence,
                                              TV027 = item.UrgentMtl,
                                              TV028 = "",
                                              TV029 = 0,
                                              TV030 = 0,
                                              TV031 = "",
                                              TV032 = "",
                                              TV033 = item.PoRemark,
                                              TV034 = item.ClosureStatus,
                                              TV035 = item.LockStaus,
                                              TV036 = item.PoStaus,
                                              TV037 = PrmStatus,
                                              TV038 = item.ModiReason,
                                              TV039 = item.PrDetailRemark,
                                              TV040 = "2",
                                              TV041 = item.Project,
                                              TV042 = item.PrExchangeRate,
                                              TV043 = item.PrPriceTw,
                                              TV044 = item.BudgetDepartmentNo,
                                              TV045 = item.BudgetDepartmentSubject,
                                              TV046 = item.PrUnitPrice,
                                              TV047 = item.PrCurrency,
                                              TV048 = item.PrPrice,
                                              TV049 = 0,
                                              TV050 = 0,
                                              TV051 = "",
                                              TV052 = "",
                                              TV053 = "",
                                              TV054 = item.TaxNo,
                                              TV055 = item.TradeTerm,
                                              TV056 = item.DetailMultiTax,
                                              TV057 = item.BusinessTaxRate,
                                              TV058 = item.PrQty,
                                              TV059 = item.UomNo,
                                              TV060 = item.DiscountRate,
                                              TV061 = item.DiscountAmount,
                                              TV104 = item.OriPrSequence,
                                              TV105 = item.OriMtlItemNo,
                                              TV106 = item.OriPrMtlItemName,
                                              TV107 = item.OriPrMtlItemSpec,
                                              TV108 = item.OriPrUomNo,
                                              TV109 = item.OriInventoryNo,
                                              TV110 = item.OriPrQty,
                                              TV111 = item.OriSupplierNo,
                                              TV112 = item.OriDemandDate != null ? item.OriDemandDate : (DateTime?)null,
                                              TV113 = item.OriPoUserNo,
                                              TV114 = item.OriPoQty,
                                              TV115 = item.OriPoUomNo,
                                              TV116 = item.OriPoCurrency,
                                              TV117 = item.OriPoUnitPrice,
                                              TV118 = item.OriPoPrice,
                                              TV119 = item.OriDeliveryDate != null ? item.OriDeliveryDate : (DateTime?)null,
                                              TV120 = "",
                                              TV121 = item.OriTaxation,
                                              TV122 = "",
                                              TV123 = "",
                                              TV124 = item.OriSoErpPrefix,
                                              TV125 = item.OriSoErpNo,
                                              TV126 = item.OriSoSequence,
                                              TV127 = item.OriUrgentMtl,
                                              TV128 = "",
                                              TV129 = 0,
                                              TV130 = 0,
                                              TV131 = "",
                                              TV132 = "",
                                              TV133 = item.OriPoRemark,
                                              TV134 = item.OriClosureStatus,
                                              TV135 = item.OriLockStaus,
                                              TV136 = item.OriPoStaus,
                                              TV137 = item.OriPrDetailRemark,
                                              TV140 = "2",
                                              TV141 = item.OriProject,
                                              TV142 = item.OriPrExchangeRate,
                                              TV143 = item.OriPrPriceTw,
                                              TV144 = item.OriBudgetDepartmentNo,
                                              TV145 = item.OriBudgetDepartmentSubject,
                                              TV146 = item.OriPrUnitPrice,
                                              TV147 = item.OriPrCurrency,
                                              TV148 = item.OriPrPrice,
                                              TV149 = 0,
                                              TV150 = 0,
                                              TV151 = "",
                                              TV152 = "",
                                              TV153 = "",
                                              TV154 = item.OriTaxNo,
                                              TV155 = item.OriTradeTerm,
                                              TV156 = item.OriDetailMultiTax,
                                              TV157 = item.OriBusinessTaxRate,
                                              TV158 = item.OriPrQty,
                                              TV159 = item.OriPrUomNo,
                                              TV160 = item.OriDiscountRate,
                                              TV161 = item.OriDiscountAmount
                                          });

                                        var insertResult2 = sqlConnection2.Query(sql, dynamicParameters);

                                        rowsAffected += insertResult2.Count();
                                        #endregion

                                        #region //UPDATE PURTB 原請購單單身
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"UPDATE PURTB SET
                                                MODIFIER = @MODIFIER,
                                                MODI_DATE = @MODI_DATE,
                                                FLAG = @FLAG,
                                                MODI_TIME = @MODI_TIME,
                                                MODI_AP = @MODI_AP,
                                                MODI_PRID = @MODI_PRID,
                                                TB004 = @TB004,            --品號
                                                TB005 = @TB005,            --品名
                                                TB006 = @TB006,            --規格
                                                TB007 = @TB007,            --請購單位
                                                TB008 = @TB008,            --庫別
                                                TB009 = @TB009,            --請購數量
                                                TB010 = @TB010,            --供應廠商
                                                TB011 = @TB011,            --需求日期
                                                TB012 = @TB012,            --備註
                                                TB013 = @TB013,            --採購人員
                                                TB014 = @TB014,            --採購數量
                                                TB015 = @TB015,            --採購單位
                                                TB016 = @TB016,            --採購幣別
                                                TB017 = @TB017,            --採購單價
                                                TB018 = @TB018,            --採購金額
                                                TB019 = @TB019,            --交貨日
                                                TB020 = @TB020,            --鎖定碼
                                                TB021 = @TB021,            --採購碼
                                                TB024 = @TB024,            --採購備註
                                                TB025 = @TB025,            --確認碼
                                                TB026 = @TB026,            --課稅別
                                                TB029 = @TB029,            --參考單別
                                                TB030 = @TB030,            --參考單號
                                                TB031 = @TB031,            --參考訂單序號
                                                TB032 = @TB032,            --急料
                                                TB039 = @TB039,            --結案碼
                                                TB043 = @TB043,            --專案代號
                                                TB044 = @TB044,            --請購匯率
                                                TB045 = @TB045,            --本幣金額
                                                TB047 = @TB047,            --預算編號
                                                TB048 = @TB048,            --預算科目
                                                TB049 = @TB049,            --請購單價
                                                TB050 = @TB050,            --請購幣別
                                                TB051 = @TB051,            --請購金額
                                                TB057 = @TB057,            --稅別碼
                                                TB058 = @TB058,            --交易條件
                                                TB063 = @TB063,            --營業稅率
                                                TB064 = @TB064,            --單身多稅率
                                                TB065 = @TB065,            --計價數量
                                                TB066 = @TB066,            --計價單位
                                                TB067 = @TB067,            --折扣率
                                                TB068 = @TB068             --折扣金額
                                                WHERE TB001 = @TB001
                                                AND TB002 = @TB002
                                                AND TB003 = @TB003";
                                        dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            MODIFIER,
                                            MODI_DATE,
                                            FLAG,
                                            MODI_TIME,
                                            MODI_AP,
                                            MODI_PRID,
                                            TB004 = item.MtlItemNo,
                                            TB005 = item.PrMtlItemName,
                                            TB006 = item.PrMtlItemSpec,
                                            TB007 = item.UomNo,
                                            TB008 = item.InventoryNo,
                                            TB009 = item.PrQty,
                                            TB010 = item.SupplierNo,
                                            TB011 = item.DemandDate,
                                            TB012 = item.PrDetailRemark,
                                            TB013 = item.PoUserNo,
                                            TB014 = item.PoQty,
                                            TB015 = item.PoUomNo,
                                            TB016 = item.PoCurrency,
                                            TB017 = item.PoUnitPrice,
                                            TB018 = item.PoPrice,
                                            TB019 = item.DeliveryDate,
                                            TB020 = item.LockStaus,
                                            TB021 = item.PoStaus,
                                            TB024 = item.PoRemark,
                                            TB025 = PrmStatus,
                                            TB026 = item.Taxation,
                                            TB029 = item.SoErpPrefix != null ? item.SoErpPrefix : "",
                                            TB030 = item.SoErpNo != null ? item.SoErpNo : "",
                                            TB031 = item.SoSequence != null ? item.SoSequence : "",
                                            TB032 = item.UrgentMtl,
                                            TB039 = item.ClosureStatus,
                                            TB043 = item.Project,
                                            TB044 = item.PrExchangeRate,
                                            TB045 = item.PrPriceTw,
                                            TB047 = item.BudgetDepartmentNo,
                                            TB048 = item.BudgetDepartmentSubject,
                                            TB049 = item.PrUnitPrice,
                                            TB050 = item.PrCurrency,
                                            TB051 = item.PrPrice,
                                            TB057 = item.TaxNo,
                                            TB058 = item.TradeTerm,
                                            TB063 = item.BusinessTaxRate,
                                            TB064 = item.DetailMultiTax,
                                            TB065 = item.PrQty,
                                            TB066 = item.UomNo,
                                            TB067 = item.DiscountRate,
                                            TB068 = item.DiscountAmount,
                                            TB001 = item.PrErpPrefix,
                                            TB002 = item.PrErpNo,
                                            TB003 = item.OriPrSequence
                                        });

                                        rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);
                                        #endregion

                                        #region //UPDATE PURTY 原請購單子單身集
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"UPDATE PURTY SET
                                                MODIFIER = @MODIFIER,
                                                MODI_DATE = @MODI_DATE,
                                                FLAG = @FLAG,
                                                MODI_TIME = @MODI_TIME,
                                                MODI_AP = @MODI_AP,
                                                MODI_PRID = @MODI_PRID,
                                                TY005 = @TY005,            --庫別
                                                TY006 = @TY006,            --供應廠商
                                                TY007 = @TY007,            --採購數量
                                                TY008 = @TY008,            --採購幣別
                                                TY009 = @TY009,            --採購單價
                                                TY010 = @TY010,            --採購金額
                                                TY011 = @TY011,            --交貨日
                                                TY012 = @TY012,            --鎖定碼
                                                TY015 = @TY015,            --採購備註
                                                TY016 = @TY016,            --採購人員
                                                TY017 = @TY017,            --課稅別
                                                TY018 = @TY018,            --急料
                                                TY021 = @TY021,            --結案碼
                                                TY022 = @TY022,            --採購碼
                                                TY031 = @TY031,            --稅別碼
                                                TY032 = @TY032,            --交易條件
                                                TY038 = @TY038,            --請購金額
                                                TY039 = @TY039,            --營業稅率
                                                TY040 = @TY040,            --單身多稅率
                                                TY041 = @TY041,            --計價數量
                                                TY042 = @TY042,            --折扣率
                                                TY043 = @TY043             --折扣金額
                                                WHERE TY001 = @TY001
                                                AND TY002 = @TY002
                                                AND TY003 = @TY003";
                                        dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            MODIFIER,
                                            MODI_DATE,
                                            FLAG,
                                            MODI_TIME,
                                            MODI_AP,
                                            MODI_PRID,
                                            TY005 = item.InventoryNo,
                                            TY006 = item.SupplierNo,
                                            TY007 = item.PoQty,
                                            TY008 = item.PoCurrency,
                                            TY009 = item.PoUnitPrice,
                                            TY010 = item.PoPrice,
                                            TY011 = item.DeliveryDate,
                                            TY012 = item.LockStaus,
                                            TY015 = item.PoRemark,
                                            TY016 = item.PoUserNo,
                                            TY017 = item.Taxation,
                                            TY018 = item.UrgentMtl,
                                            TY021 = item.ClosureStatus,
                                            TY022 = item.PoStaus,
                                            TY031 = item.TaxNo,
                                            TY032 = item.TradeTerm,
                                            TY038 = item.PrPrice,
                                            TY039 = item.BusinessTaxRate,
                                            TY040 = item.DetailMultiTax,
                                            TY041 = item.PrPriceQty,
                                            TY042 = item.DiscountRate,
                                            TY043 = item.DiscountAmount,
                                            TY001 = item.PrErpPrefix,
                                            TY002 = item.PrErpNo,
                                            TY003 = item.OriPrSequence
                                        });

                                        rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);
                                        #endregion

                                        #region //UPDATE SCM.PrDetail 原MES請購單身
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"UPDATE SCM.PrDetail SET
                                                MtlItemId = @MtlItemId,
                                                PrMtlItemName = @PrMtlItemName,
                                                PrMtlItemSpec = @PrMtlItemSpec,
                                                InventoryId = @InventoryId,
                                                PrUomId = @PrUomId,
                                                PrQty = @PrQty,
                                                DemandDate = @DemandDate,
                                                SupplierId = @SupplierId,
                                                PrCurrency = @PrCurrency,
                                                PrExchangeRate = @PrExchangeRate,
                                                PrUnitPrice = @PrUnitPrice,
                                                PrPrice = @PrPrice,
                                                PrPriceTw = @PrPriceTw,
                                                UrgentMtl = @UrgentMtl,
                                                ProductionPlan = @ProductionPlan,
                                                Project = @Project,
                                                PrPriceQty = @PrQty,
                                                PrPriceUomId = @PrUomId,
                                                MtlInventoryQty = @MtlInventoryQty,
                                                PoUomId = @PoUomId,
                                                PoQty = @PoQty,
                                                PoCurrency = @PoCurrency,
                                                PoUnitPrice = @PoUnitPrice,
                                                PoPrice = @PoPrice,
                                                LastModifiedDate = @LastModifiedDate,
                                                LastModifiedBy = @LastModifiedBy
                                                WHERE PrDetailId = @PrDetailId";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                item.MtlItemId,
                                                item.PrMtlItemName,
                                                item.PrMtlItemSpec,
                                                item.InventoryId,
                                                item.PrUomId,
                                                item.PrQty,
                                                item.DemandDate,
                                                item.SupplierId,
                                                item.PrCurrency,
                                                item.PrExchangeRate,
                                                item.PrUnitPrice,
                                                item.PrPrice,
                                                item.PrPriceTw,
                                                item.UrgentMtl,
                                                item.ProductionPlan,
                                                item.Project,
                                                PrPriceQty = item.PrQty,
                                                PrPriceUomId = item.PrUomId,
                                                item.MtlInventoryQty,
                                                PoUomId = item.PrUomId,
                                                PoQty = item.PrQty,
                                                PoCurrency = item.PrCurrency,
                                                PoUnitPrice = item.PrUnitPrice,
                                                PoPrice = item.PrPrice,
                                                LastModifiedDate,
                                                LastModifiedBy,
                                                item.PrDetailId
                                            });

                                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                        #endregion
                                    }

                                    #region //重新計算ERP總請購數量
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT ISNULL(SUM(a.TB009), 0) TotalErpPrQty
                                            , ISNULL(SUM(a.TB045), 0) TotalErpPrPrice
                                            FROM PURTB a
                                            WHERE a.TB001 = @TB001
                                            AND a.TB002 = @TB002";
                                    dynamicParameters.Add("TB001", PrErpPrefix);
                                    dynamicParameters.Add("TB002", PrErpNo);
                                    var TotalErpPrQtyResult = sqlConnection2.Query(sql, dynamicParameters);

                                    double TotalErpPrQty = -1;
                                    double TotalErpPrPrice = -1;
                                    foreach (var item in TotalErpPrQtyResult)
                                    {
                                        TotalErpPrQty = Convert.ToDouble(item.TotalErpPrQty);
                                        TotalErpPrPrice = Convert.ToDouble(item.TotalErpPrPrice);
                                    }
                                    #endregion

                                    #region //UPDATE PURTA 原請購單單頭
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE PURTA SET
                                            MODIFIER = @MODIFIER,
                                            MODI_DATE = @MODI_DATE,
                                            FLAG = @FLAG,
                                            MODI_TIME = @MODI_TIME,
                                            MODI_AP = @MODI_AP,
                                            MODI_PRID = @MODI_PRID,
                                            TA004 = @TA004,                           --請購部門
                                            TA006 = @TA006,                           --備註
                                            TA007 = @TA007,                           --確認碼
                                            TA011 = @TA011,                           --數量合計
                                            TA012 = @TA012,                           --請購人員
                                            TA014 = @TA014,                           --確認者
                                            TA020 = @TA020,                           --本幣金額合計
                                            TA021 = @TA021,                           --版次
                                            TA022 = @TA022                            --預算部門
                                            WHERE TA001 = @TA001
                                            AND TA002 = @TA002";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            MODIFIER,
                                            MODI_DATE,
                                            FLAG,
                                            MODI_TIME,
                                            MODI_AP,
                                            MODI_PRID,
                                            TA004 = DepartmentNo,
                                            TA006 = PrmRemark,
                                            TA007 = PrmStatus,
                                            TA011 = TotalErpPrQty,
                                            TA012 = PrUserNo,
                                            TA014 = ComfirmUser,
                                            TA020 = TotalErpPrPrice,
                                            TA021 = Edition,
                                            TA022 = BudgetDepartmentNo,
                                            TA001 = PrErpPrefix,
                                            TA002 = PrErpNo
                                        });

                                    rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //重新計算MES總請購數量
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT ISNULL(SUM(a.PrQty), 0) TotalPrQty
                                            , ISNULL(SUM(a.PrPriceTw), 0) TotalPrPrice
                                            FROM SCM.PrDetail a
                                            WHERE a.PrId = @PrId";
                                    dynamicParameters.Add("PrId", PrId);

                                    var TotalPrQtyResult = sqlConnection.Query(sql, dynamicParameters);

                                    double TotalPrQty = -1;
                                    double TotalPrPrice = -1;
                                    foreach (var item in TotalPrQtyResult)
                                    {
                                        TotalPrQty = item.TotalPrQty;
                                        TotalPrPrice = item.TotalPrPrice;
                                    }
                                    #endregion

                                    #region //UPDATE MES.PurchaseRequisition 原MES請購單單頭
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE SCM.PurchaseRequisition SET
                                            DepartmentId = @DepartmentId,
                                            UserId = @UserId,
                                            Edition = @Edition,
                                            DocDate = @MesDocDate,
                                            PrRemark = @PrmRemark,
                                            TotalQty = @TotalQty,
                                            Amount = @TotalPrPrice,
                                            BudgetDepartmentId = @BudgetDepartmentId,
                                            ConfirmUserId = @ConfirmUserId,
                                            PrDate = @LastModifiedDate,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE PrId = @PrId";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            CompanyId = CurrentCompany,
                                            DepartmentId,
                                            UserId,
                                            Edition,
                                            MesDocDate,
                                            PrmRemark,
                                            TotalQty = TotalPrQty,
                                            TotalPrPrice,
                                            BudgetDepartmentId,
                                            ConfirmUserId,
                                            LastModifiedDate,
                                            LastModifiedBy,
                                            PrId
                                        });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //UPDATE MES.PrDetail 原MES請購單單身
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE SCM.PrDetail SET
                                            DemandDate = @LastModifiedDate,
                                            LastModifiedDate = @LastModifiedDate,
                                            LastModifiedBy = @LastModifiedBy
                                            WHERE PrId = @PrId";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            LastModifiedDate,
                                            LastModifiedBy,
                                            PrId
                                        });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion
                                }
                                else if (PrmStatus == "E") //單據修改
                                {
                                    #region //先確認ERP是否已經有請購變更單據資料
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM PURTU
                                            WHERE TU001 = @TU001
                                            AND TU002 = @TU002
                                            AND TU003 = @TU003";
                                    dynamicParameters.Add("TU001", PrErpPrefix);
                                    dynamicParameters.Add("TU002", PrErpNo);
                                    dynamicParameters.Add("TU003", Edition);

                                    var CheckPrmResult = sqlConnection2.Query(sql, dynamicParameters);
                                    #endregion

                                    if (CheckPrmResult.Count() > 0)
                                    {
                                        #region //PURTU 請購變更單單頭

                                        #region //UPDATE PURTU
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"UPDATE PURTU SET
                                                MODIFIER = @MODIFIER,
                                                MODI_DATE = @MODI_DATE,
                                                FLAG = @FLAG,
                                                MODI_TIME = @MODI_TIME,
                                                MODI_AP = @MODI_AP,
                                                MODI_PRID = @MODI_PRID,
                                                TU005 = @TU005,
                                                TU006 = @TU006,
                                                TU007 = @TU007,
                                                TU009 = @TU009,
                                                TU012 = 'N',
                                                TU014 = @TU014,
                                                TU018 = @TU018,
                                                TU019 = @TU019,
                                                TU020 = @TU020,
                                                UDF02 = @UDF02
                                                WHERE TU001 = @TU001
                                                AND TU002 = @TU002
                                                AND TU003 = @TU003";
                                        dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            MODIFIER,
                                            MODI_DATE,
                                            FLAG,
                                            MODI_TIME,
                                            MODI_AP,
                                            MODI_PRID,
                                            TU005 = DocDate,
                                            TU006 = DepartmentNo,
                                            TU007 = PrUserNo,
                                            TU009 = ModiReason,
                                            TU014 = "0", // 狀態為E時為修改狀態
                                            TU018 = PrmRemark,
                                            TU019 = Amount,
                                            TU020 = BudgetDepartmentNo,
                                            UDF02 = BpmNo,
                                            TU001 = PrErpPrefix,
                                            TU002 = PrErpNo,
                                            TU003 = Edition
                                        });

                                        rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);
                                        #endregion

                                        #endregion

                                        #region //PURTV 請購變更單單身

                                        #region //先刪除全部單身
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"DELETE FROM PURTV
                                                WHERE TV001 = @TV001
                                                AND TV002 = @TV002
                                                AND TV003 = @TV003";
                                        dynamicParameters.Add("TV001", PrErpPrefix);
                                        dynamicParameters.Add("TV002", PrErpNo);
                                        dynamicParameters.Add("TV003", Edition);
                                        int rowsAffected2 = sqlConnection2.Execute(sql, dynamicParameters);
                                        #endregion

                                        foreach (var item in prmDetailResult)
                                        {
                                            #region //INSERT PURTV 請購變更單單身
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"INSERT INTO PURTV (COMPANY, CREATOR, USR_GROUP, CREATE_DATE, MODIFIER, MODI_DATE
                                            , FLAG, CREATE_TIME, CREATE_AP, CREATE_PRID, MODI_TIME, MODI_AP, MODI_PRID
                                            , TV001, TV002, TV003, TV004, TV005, TV006, TV007, TV008
                                            , TV009, TV010, TV011, TV012, TV013, TV014, TV015, TV016
                                            , TV017, TV018, TV019, TV020, TV021, TV022, TV023, TV024
                                            , TV025, TV026, TV027, TV028, TV029, TV030, TV031, TV032
                                            , TV033, TV034, TV035, TV036, TV037, TV038, TV039, TV040
                                            , TV041, TV042, TV043, TV044, TV045, TV046, TV047, TV048
                                            , TV049, TV050, TV051, TV052, TV053, TV054, TV055, TV056
                                            , TV057, TV058, TV059, TV060, TV061, TV104, TV105, TV106
                                            , TV107, TV108, TV109, TV110, TV111, TV112, TV113, TV114
                                            , TV115, TV116, TV117, TV118, TV119, TV120, TV121, TV122
                                            , TV123, TV124, TV125, TV126, TV127, TV128, TV129, TV130
                                            , TV131, TV132, TV133, TV134, TV135, TV136, TV137, TV140
                                            , TV141, TV142, TV143, TV144, TV145, TV146, TV147, TV148
                                            , TV149, TV150, TV151, TV152, TV153, TV154, TV155, TV156
                                            , TV157, TV158, TV159, TV160, TV161)
                                            VALUES (@COMPANY, @CREATOR, @USR_GROUP, @CREATE_DATE, @MODIFIER, @MODI_DATE
                                            , @FLAG, @CREATE_TIME, @CREATE_AP, @CREATE_PRID, @MODI_TIME, @MODI_AP, @MODI_PRID
                                            , @TV001, @TV002, @TV003, @TV004, @TV005, @TV006, @TV007, @TV008
                                            , @TV009, @TV010, @TV011, @TV012, @TV013, @TV014, @TV015, @TV016
                                            , @TV017, @TV018, @TV019, @TV020, @TV021, @TV022, @TV023, @TV024
                                            , @TV025, @TV026, @TV027, @TV028, @TV029, @TV030, @TV031, @TV032
                                            , @TV033, @TV034, @TV035, @TV036, @TV037, @TV038, @TV039, @TV040
                                            , @TV041, @TV042, @TV043, @TV044, @TV045, @TV046, @TV047, @TV048
                                            , @TV049, @TV050, @TV051, @TV052, @TV053, @TV054, @TV055, @TV056
                                            , @TV057, @TV058, @TV059, @TV060, @TV061, @TV104, @TV105, @TV106
                                            , @TV107, @TV108, @TV109, @TV110, @TV111, @TV112, @TV113, @TV114
                                            , @TV115, @TV116, @TV117, @TV118, @TV119, @TV120, @TV121, @TV122
                                            , @TV123, @TV124, @TV125, @TV126, @TV127, @TV128, @TV129, @TV130
                                            , @TV131, @TV132, @TV133, @TV134, @TV135, @TV136, @TV137, @TV140
                                            , @TV141, @TV142, @TV143, @TV144, @TV145, @TV146, @TV147, @TV148
                                            , @TV149, @TV150, @TV151, @TV152, @TV153, @TV154, @TV155, @TV156
                                            , @TV157, @TV158, @TV159, @TV160, @TV161)";

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
                                                  TV001 = PrErpPrefix,
                                                  TV002 = PrErpNo,
                                                  TV003 = Edition,
                                                  TV004 = item.PrmSequence,
                                                  TV005 = item.MtlItemNo,
                                                  TV006 = item.MtlItemName,
                                                  TV007 = item.MtlItemSpec,
                                                  TV008 = item.UomNo,
                                                  TV009 = item.InventoryNo,
                                                  TV010 = item.PrQty,
                                                  TV011 = item.SupplierNo,
                                                  TV012 = item.DemandDate,
                                                  TV013 = item.PoUserNo,
                                                  TV014 = item.PoQty,
                                                  TV015 = item.PoUomNo,
                                                  TV016 = item.PoCurrency,
                                                  TV017 = item.PoUnitPrice,
                                                  TV018 = item.PoPrice,
                                                  TV019 = item.DeliveryDate,
                                                  TV020 = "",
                                                  TV021 = item.Taxation,
                                                  TV022 = "",
                                                  TV023 = "",
                                                  TV024 = item.SoErpPrefix,
                                                  TV025 = item.SoErpNo,
                                                  TV026 = item.SoSequence,
                                                  TV027 = item.UrgentMtl,
                                                  TV028 = "",
                                                  TV029 = 0,
                                                  TV030 = 0,
                                                  TV031 = "",
                                                  TV032 = "",
                                                  TV033 = item.PoRemark,
                                                  TV034 = item.ClosureStatus,
                                                  TV035 = item.LockStaus,
                                                  TV036 = item.PoStaus,
                                                  TV037 = 'N',
                                                  TV038 = item.ModiReason,
                                                  TV039 = item.PrDetailRemark,
                                                  TV040 = "2",
                                                  TV041 = item.Project,
                                                  TV042 = item.PrExchangeRate,
                                                  TV043 = item.PrPriceTw,
                                                  TV044 = item.BudgetDepartmentNo,
                                                  TV045 = item.BudgetDepartmentSubject,
                                                  TV046 = item.PrUnitPrice,
                                                  TV047 = item.PrCurrency,
                                                  TV048 = item.PrPrice,
                                                  TV049 = 0,
                                                  TV050 = 0,
                                                  TV051 = "",
                                                  TV052 = "",
                                                  TV053 = "",
                                                  TV054 = item.TaxNo,
                                                  TV055 = item.TradeTerm,
                                                  TV056 = item.DetailMultiTax,
                                                  TV057 = item.BusinessTaxRate,
                                                  TV058 = item.PrQty,
                                                  TV059 = item.UomNo,
                                                  TV060 = item.DiscountRate,
                                                  TV061 = item.DiscountAmount,
                                                  TV104 = item.OriPrSequence,
                                                  TV105 = item.OriMtlItemNo,
                                                  TV106 = item.OriPrMtlItemName,
                                                  TV107 = item.OriPrMtlItemSpec,
                                                  TV108 = item.OriPrUomNo,
                                                  TV109 = item.OriInventoryNo,
                                                  TV110 = item.OriPrQty,
                                                  TV111 = item.OriSupplierNo,
                                                  TV112 = item.OriDemandDate != null ? item.OriDemandDate : (DateTime?)null,
                                                  TV113 = item.OriPoUserNo,
                                                  TV114 = item.OriPoQty,
                                                  TV115 = item.OriPoUomNo,
                                                  TV116 = item.OriPoCurrency,
                                                  TV117 = item.OriPoUnitPrice,
                                                  TV118 = item.OriPoPrice,
                                                  TV119 = item.OriDeliveryDate != null ? item.OriDeliveryDate : (DateTime?)null,
                                                  TV120 = "",
                                                  TV121 = item.OriTaxation,
                                                  TV122 = "",
                                                  TV123 = "",
                                                  TV124 = item.OriSoErpPrefix,
                                                  TV125 = item.OriSoErpNo,
                                                  TV126 = item.OriSoSequence,
                                                  TV127 = item.OriUrgentMtl,
                                                  TV128 = "",
                                                  TV129 = 0,
                                                  TV130 = 0,
                                                  TV131 = "",
                                                  TV132 = "",
                                                  TV133 = item.OriPoRemark,
                                                  TV134 = item.OriClosureStatus,
                                                  TV135 = item.OriLockStaus,
                                                  TV136 = item.OriPoStaus,
                                                  TV137 = item.OriPrDetailRemark,
                                                  TV140 = "2",
                                                  TV141 = item.OriProject,
                                                  TV142 = item.OriPrExchangeRate,
                                                  TV143 = item.OriPrPriceTw,
                                                  TV144 = item.OriBudgetDepartmentNo,
                                                  TV145 = item.OriBudgetDepartmentSubject,
                                                  TV146 = item.OriPrUnitPrice,
                                                  TV147 = item.OriPrCurrency,
                                                  TV148 = item.OriPrPrice,
                                                  TV149 = 0,
                                                  TV150 = 0,
                                                  TV151 = "",
                                                  TV152 = "",
                                                  TV153 = "",
                                                  TV154 = item.OriTaxNo,
                                                  TV155 = item.OriTradeTerm,
                                                  TV156 = item.OriDetailMultiTax,
                                                  TV157 = item.OriBusinessTaxRate,
                                                  TV158 = item.OriPrQty,
                                                  TV159 = item.OriPrUomNo,
                                                  TV160 = item.OriDiscountRate,
                                                  TV161 = item.OriDiscountAmount
                                              });
                                            var insertResult2 = sqlConnection2.Query(sql, dynamicParameters);

                                            rowsAffected += insertResult2.Count();
                                            #endregion
                                        }

                                        #endregion
                                    }
                                }

                                #region //請購變更單附檔
                                string DocID = "";
                                string SeqNo = "";
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.PrmFileId, a.PrmId, a.FileId
                                        , b.[FileName], b.FileExtension, b.FileContent
                                        , c.UserNo
                                        FROM SCM.PrmFile a
                                        INNER JOIN BAS.[File] b ON a.FileId = b.FileId
                                        INNER JOIN BAS.[User] c ON a.CreateBy = c.UserId
                                        WHERE a.PrmId = @PrmId";
                                dynamicParameters.Add("PrmId", PrmId);

                                var prmFileResult = sqlConnection.Query(sql, dynamicParameters);

                                if (prmFileResult.Count() > 0)
                                {
                                    foreach (var item in prmFileResult)
                                    {
                                        #region //DMS

                                        #region //DMS取號
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT '001'+ RIGHT(REPLICATE('0', 7) + CONVERT(VARCHAR(7)
                                                , ISNULL(MAX(SUBSTRING(a.DocID, LEN(a.DocID) - 6, 7)), 0) + 1), 7) AS nextSN
                                                FROM dbo.DMS a
                                                --FROM BPM.DMS a
                                                WHERE a.DocID LIKE '001%'";

                                        var dmsResult = sqlConnection3.Query(sql, dynamicParameters);

                                        foreach (var item2 in dmsResult)
                                        {
                                            DocID = item2.nextSN;
                                        }
                                        #endregion

                                        #region //INSERT DMS
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO dbo.DMS (DocID, Revision, Owner, Status, AddDate, AddTime
                                                , DocName, DocType, Description)
                                                VALUES (@DocID, @Revision, @Owner, @Status, @AddDate, @AddTime
                                                , @DocName, @DocType, @Description)";

                                        dynamicParameters.AddDynamicParams(
                                          new
                                          {
                                              DocID,
                                              Revision = "001",
                                              Owner = item.UserNo,
                                              Status = "N",
                                              AddDate = DateTime.Now.ToString("yyyyMMdd"),
                                              AddTime = DateTime.Now.ToString("HH:mm"),
                                              DocName = item.FileName + item.FileExtension,
                                              DocType = item.FileExtension,
                                              Description = item.FileName + item.FileExtension
                                          });
                                        var insertResult2 = sqlConnection3.Query(sql, dynamicParameters);

                                        rowsAffected += insertResult2.Count();
                                        #endregion

                                        #endregion

                                        #region //ATTACH

                                        #region //取號
                                        sql = @"SELECT RIGHT(REPLICATE('0', 3) + CONVERT(VARCHAR(3)
                                                , ISNULL(MAX(SUBSTRING(SeqNo, LEN(SeqNo) - 2, 3)), 0) + 1), 3) AS nextSN
                                                FROM dbo.ATTACH a
                                                --FROM BPM.ATTACH a
                                                WHERE a.CompanyID = @CompanyID
                                                AND a.UserID = @UserID
                                                AND a.Parent = @Parent
                                                AND a.KeyValues = @KeyValues";
                                        dynamicParameters.Add("CompanyID", ErpNo);
                                        dynamicParameters.Add("UserID", item.UserNo);
                                        dynamicParameters.Add("Parent", "PURI16");
                                        dynamicParameters.Add("KeyValues", PrErpPrefix + "||" + PrErpNo + "||" + Edition);

                                        var attachResult = sqlConnection3.Query(sql, dynamicParameters);

                                        foreach (var item2 in attachResult)
                                        {
                                            SeqNo = item2.nextSN;
                                        }
                                        #endregion

                                        #region //INSERT ATTACH
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO dbo.ATTACH (Parent, KeyValues, CompanyID, UserID, Type, SeqNo
                                                , FileName, DocID, Revision, AddDate, AddTime, KeyFields)
                                                VALUES (@Parent, @KeyValues, @CompanyID, @UserID, @Type, @SeqNo
                                                , @FileName, @DocID, @Revision, @AddDate, @AddTime, @KeyFields)";
                                        dynamicParameters.AddDynamicParams(
                                          new
                                          {
                                              Parent = "PURI16",
                                              KeyValues = PrErpPrefix + "||" + PrErpNo + "||" + Edition,
                                              CompanyID = ErpNo,
                                              UserID = item.UserNo,
                                              Type = "",
                                              SeqNo,
                                              FileName = item.FileName + item.FileExtension,
                                              DocID,
                                              Revision = "001",
                                              AddDate = DateTime.Now.ToString("yyyyMMdd"),
                                              AddTime = DateTime.Now.ToString("HH:mm"),
                                              KeyFields = "TU001||TU002||TU003"
                                          });
                                        var insertResult3 = sqlConnection3.Query(sql, dynamicParameters);

                                        rowsAffected += insertResult3.Count();
                                        #endregion

                                        #endregion

                                        #region //取得ERP檔案路徑
                                        double no = Convert.ToDouble(DocID.Substring(3));
                                        int folderNo = Convert.ToInt32(Math.Floor(no / 1000)) * 1000;
                                        string folderName = folderNo.ToString().PadLeft(7, '0');
                                        string targetErpFolderPath = Path.Combine(ErpFolderRoot, folderName);
                                        ErpDocPath = Path.Combine(targetErpFolderPath, "001" + DocID.Substring(3) + ".001");
                                        #endregion

                                        #region //COPY附檔至ERP路徑資料夾
                                        if (!Directory.Exists(targetErpFolderPath)) { Directory.CreateDirectory(targetErpFolderPath); }
                                        byte[] fileContent = (byte[])item.FileContent;
                                        File.WriteAllBytes(ErpDocPath, fileContent); // Requires System.IO
                                        #endregion

                                        #region //更新MES請購單附檔資訊
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"UPDATE SCM.PrmFile SET
                                                ErpDocId = @ErpDocId,
                                                ErpDocPath = @ErpDocPath,
                                                ErpDocDate = @ErpDocDate,
                                                LastModifiedDate = @LastModifiedDate,
                                                LastModifiedBy = @LastModifiedBy
                                                WHERE PrmFileId = @PrmFileId";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                ErpDocId = DocID,
                                                ErpDocPath,
                                                ErpDocDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                                                LastModifiedDate,
                                                LastModifiedBy,
                                                item.PrmFileId
                                            });

                                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                        #endregion
                                    }
                                }
                                #endregion

                                #region //將ERP請購單資訊回填MES
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE SCM.PrModification SET
                                        TransferStatus = @TransferStatus,
                                        TransferDate = @TransferDate,
                                        ConfirmStatus = @ConfirmStatus,
                                        ConfirmUserId = @ConfirmUserId,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE PrmId = @PrmId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        TransferStatus = PrmStatus == "Y" ? "Y" : "N",
                                        TransferDate = PrmStatus == "Y" ? LastModifiedDate : (DateTime?)null,
                                        ConfirmStatus = PrmStatus == "Y" ? "Y" : "N",
                                        ConfirmUserId = PrmStatus == "Y" ? ConfirmUserId : (int?)null,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        PrmId
                                    });

                                rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE SCM.PrmDetail SET
                                        ConfirmStatus = @ConfirmStatus,
                                        ConfirmUserId = @ConfirmUserId,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE PrmId = @PrmId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        ConfirmStatus = PrmStatus == "Y" ? "Y" : "N",
                                        ConfirmUserId = PrmStatus == "Y" ? ConfirmUserId : (int?)null,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        PrmId
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

        #region //UpdatePrmStatus -- 更新請購變更單狀態 -- Ann 2023-02-10
        public string UpdatePrmStatus(int PrmId, string BpmNo, string Status, string RootId, string ConfirmUser, string ErpFlag)
        {
            try
            {
                int rowsAffected = 0;
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷請購變更單資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.PrId, a.BpmTransferDate, a.BpmTransferStatus, a.SignupStaus, a.PrmStatus
                                FROM SCM.PrModification a
                                WHERE a.PrmId = @PrmId";
                        dynamicParameters.Add("PrmId", PrmId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("請購變更單資料錯誤!");

                        DateTime BpmTransferDate = new DateTime();
                        string BpmTransferStatus = "";
                        string SignupStaus = "";
                        int PrId = -1;
                        foreach (var item in result)
                        {
                            if (item.BpmTransferStatus != "Y") throw new SystemException("此請購變更單尚未拋轉BPM!!");
                            BpmTransferDate = item.BpmTransferDate;
                            BpmTransferStatus = item.BpmTransferStatus;
                            SignupStaus = item.SignupStaus;
                            PrId = item.PrId;
                        }
                        #endregion

                        #region //UPDATE SCM.PrModification
                        if (RootId.Length > 0) //BPM結束流程(E、Y)
                        {
                            #region //存LOG
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO SCM.PrmLog (PrmId, RootId, BpmNo, TransferBpmDate, BpmStatus, ConfirmUser
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    VALUES (@PrmId, @RootId, @BpmNo, @TransferBpmDate, @BpmStatus, @ConfirmUser
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                              new
                              {
                                  PrmId,
                                  RootId,
                                  BpmNo,
                                  TransferBpmDate = BpmTransferDate,
                                  BpmStatus = Status,
                                  ConfirmUser,
                                  CreateDate,
                                  LastModifiedDate,
                                  CreateBy,
                                  LastModifiedBy
                              });
                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();
                            #endregion

                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.PrModification SET
                                    PrmStatus = @PrmStatus,
                                    SignupStaus = @SignupStaus,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE PrmId = @PrmId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    PrmStatus = ErpFlag != "F" ? Status : "F",
                                    SignupStaus = Status == "Y" ? "3" : SignupStaus,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    PrmId
                                });

                            rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

                            //if (Status == "Y")
                            //{
                            //    #region //確認是否已經有此筆請購單之變更紀錄
                            //    dynamicParameters = new DynamicParameters();
                            //    sql = @"SELECT TOP 1 (MAX(b.Edition) + 1) Edition
                            //            FROM SCM.PrModification a
                            //            INNER JOIN SCM.PurchaseRequisition b ON a.PrId = b.PrId
                            //            WHERE a.PrId = @PrId";
                            //    dynamicParameters.Add("PrId", PrId);

                            //    var CheckPrIdResult = sqlConnection.Query(sql, dynamicParameters);

                            //    string Edition = "";
                            //    foreach (var item in CheckPrIdResult)
                            //    {
                            //        if (item.Edition == null)
                            //        {
                            //            Edition = "0001";
                            //        }
                            //        else
                            //        {
                            //            Edition = item.Edition.ToString().PadLeft(4, '0');
                            //        }
                            //    }
                            //    #endregion

                            //    #region //更新請購單版次
                            //    dynamicParameters = new DynamicParameters();
                            //    sql = @"UPDATE SCM.PurchaseRequisition SET
                            //            Edition = @Edition,
                            //            LastModifiedDate = @LastModifiedDate,
                            //            LastModifiedBy = @LastModifiedBy
                            //            WHERE PrId = @PrId";
                            //    dynamicParameters.AddDynamicParams(
                            //        new
                            //        {
                            //            Edition,
                            //            LastModifiedDate,
                            //            LastModifiedBy,
                            //            PrId
                            //        });

                            //    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                            //    #endregion
                            //}
                        }
                        else //BPM流程開始(P)、拋轉失敗(R)、ERP拋轉失敗(F)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.PrModification SET
                                    PrmStatus = @PrmStatus,
                                    BpmTransferStatus = @BpmTransferStatus,
                                    BpmTransferDate = @BpmTransferDate,
                                    TransferStatus = @TransferStatus,
                                    TransferDate = @TransferDate,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE PrmId = @PrmId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    PrmStatus = Status,
                                    BpmTransferStatus = Status == "P" ? "Y" : "N",
                                    BpmTransferDate = Status == "P" ? DateTime.Now.ToString("yyyy-MM-dd") : null,
                                    TransferStatus = Status == "P" ? "Y" : "N",
                                    TransferDate = Status == "P" ? DateTime.Now.ToString("yyyy-MM-dd") : null,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    PrmId
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

        #region //UpdatePrVoid -- 請購單作廢 -- Ann 2023-02-13
        public string UpdatePrVoid(int PrId)
        {
            try
            {
                int rowsAffected = 0;
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

                        string ErpNo = "";
                        foreach (var item in companyResult)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            ErpNo = item.ErpNo;
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //判斷請購單資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.PrErpPrefix, a.PrErpNo, a.TransferStatus, a.UserId, a.PrStatus
                                    FROM SCM.PurchaseRequisition a
                                    WHERE a.PrId = @PrId";
                            dynamicParameters.Add("PrId", PrId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("請購單資料錯誤!");

                            string PrErpPrefix = "";
                            string PrErpNo = "";
                            string TransferStatus = "";
                            foreach (var item in result)
                            {
                                if (item.PrStatus != "N" && item.PrStatus != "E" && item.PrStatus != "R" && item.PrStatus != "F") throw new SystemException("請購單狀態無法修改!");
                                if (item.UserId != CreateBy) throw new SystemException("無法修改其他請購人之單據!");
                                PrErpPrefix = item.PrErpPrefix;
                                PrErpNo = item.PrErpNo;
                                TransferStatus = item.TransferStatus;
                            }
                            #endregion

                            if (TransferStatus == "Y")
                            {
                                #region //ERP作廢
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE PURTA SET
                                        TA007 = 'V'
                                        WHERE TA001 = @TA001
                                        AND TA002 = @TA002";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        TA001 = PrErpPrefix,
                                        TA002 = PrErpNo
                                    });

                                rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);
                                #endregion
                            }

                            #region //MES請購單狀態修改
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.PurchaseRequisition SET
                                    PrStatus = @PrStatus
                                    WHERE PrId = @PrId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    PrStatus = TransferStatus == "Y" ? "V" : "S",
                                    PrId
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

        #region //UpdatePrmVoid -- 請購變更單作廢 -- Ann 2023-02-13
        public string UpdatePrmVoid(int PrmId)
        {
            try
            {
                int rowsAffected = 0;
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

                        string ErpNo = "";
                        foreach (var item in companyResult)
                        {
                            ErpConnectionStrings = ConfigurationManager.AppSettings[item.ErpDb];
                            ErpNo = item.ErpNo;
                        }
                        #endregion

                        using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                        {
                            #region //判斷請購單資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.TransferStatus, a.PrmStatus, a.UserId
                                    , b.PrErpPrefix, b.PrErpNo, b.Edition
                                    FROM SCM.PrModification a
                                    INNER JOIN SCM.PurchaseRequisition b ON a.PrId = b.PrId
                                    WHERE a.PrmId = @PrmId";
                            dynamicParameters.Add("PrmId", PrmId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("請購變更單資料錯誤!");

                            string PrErpPrefix = "";
                            string PrErpNo = "";
                            string TransferStatus = "";
                            string Edition = "";
                            foreach (var item in result)
                            {
                                if (item.PrmStatus != "N" && item.PrmStatus != "E" && item.PrmStatus != "R" && item.PrmStatus != "F") throw new SystemException("請購單狀態無法修改!");
                                if (item.UserId != CreateBy) throw new SystemException("無法修改其他請購人之單據!");
                                PrErpPrefix = item.PrErpPrefix;
                                PrErpNo = item.PrErpNo;
                                TransferStatus = item.TransferStatus;
                                Edition = item.Edition;
                            }
                            #endregion

                            if (TransferStatus == "Y")
                            {
                                #region //ERP作廢
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE PURTU SET
                                        TU012 = 'V'
                                        WHERE TU001 = @TU001
                                        AND TU002 = @TU002
                                        AND TU003 = @TU003";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        TU001 = PrErpPrefix,
                                        TU002 = PrErpNo,
                                        TU003 = Edition
                                    });

                                rowsAffected += sqlConnection2.Execute(sql, dynamicParameters);
                                #endregion
                            }

                            #region //MES請購單狀態修改
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.PrModification SET
                                    PrmStatus = @PrmStatus
                                    WHERE PrmId = @PrmId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    PrmStatus = TransferStatus == "Y" ? "V" : "S",
                                    PrmId
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

        #region //UpdatePrDuplicate -- 複製請購單 -- Ann 2023-02-23
        public string UpdatePrDuplicate(int PrId)
        {
            try
            {
                int rowsAffected = 0;
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
                            #region //取PrNo
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT ISNULL(MAX(a.PrNo), 'PR00000000000') MaxPrNo
                                    FROM SCM.PurchaseRequisition a
                                    WHERE FORMAT(a.CreateDate, 'yyyy-MM-dd') = @CreateDate";
                            dynamicParameters.Add("CreateDate", DateTime.Now.ToString("yyyy-MM-dd"));
                            var PrNoResult = sqlConnection.Query(sql, dynamicParameters);

                            string PrNo = "";
                            foreach (var item in PrNoResult)
                            {
                                string MaxPrNo = item.MaxPrNo;
                                string NewMaxPrNoNumber = MaxPrNo.Substring(MaxPrNo.Length - 3);
                                string nowDateTime = DateTime.Now.ToString("yyyyMMdd");
                                PrNo = "PR" + nowDateTime + (Convert.ToInt32(NewMaxPrNoNumber) + 1).ToString().PadLeft(3, '0');
                            }
                            #endregion

                            #region //判斷請購單資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.DepartmentId, a.UserId, a.PrRemark, a.TotalQty, a.Amount, a.BudgetDepartmentId, a.PrErpPrefix, a.Priority
                                    FROM SCM.PurchaseRequisition a
                                    WHERE a.PrId = @PrId";
                            dynamicParameters.Add("PrId", PrId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("請購單資料錯誤!");

                            int DepartmentId = -1;
                            int UserId = -1;
                            string PrRemark = "";
                            double TotalQty = -1;
                            double Amount = -1;
                            int? BudgetDepartmentId = -1;
                            string PrErpPrefix = "";
                            string Priority = "";
                            foreach (var item in result)
                            {
                                DepartmentId = item.DepartmentId;
                                UserId = item.UserId;
                                PrRemark = item.PrRemark;
                                TotalQty = item.TotalQty;
                                Amount = item.Amount;
                                BudgetDepartmentId = item.BudgetDepartmentId;
                                PrErpPrefix = item.PrErpPrefix;
                                Priority = item.Priority;
                            }
                            #endregion

                            #region //取得請購單身資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.PrSequence, a.MtlItemId, a.PrMtlItemName, a.PrMtlItemSpec, a.InventoryId, a.PrUomId, a.PrQty, a.SupplierId, a.PrCurrency, a.PrExchangeRate
                                    , a.PrUnitPrice, a.PrPrice, a.PrPriceTw, a.UrgentMtl, a.ProductionPlan, a.Project, a.BudgetDepartmentNo, a.BudgetDepartmentSubject, a.SoDetailId
                                    , a.DeliveryDate, a.PoUserId, a.PoUomId, a.PoQty, a.PoCurrency, a.PoUnitPrice, a.PoPrice, a.TaxNo, a.Taxation, a.BusinessTaxRate
                                    , a.DetailMultiTax, a.TradeTerm, a.PrPriceQty, a.PrPriceUomId, a.DiscountRate, a.DiscountAmount, a.PrDetailRemark, a.PoRemark
                                    , b.InventoryNo, b.InventoryName
                                    , c.MtlItemNo, c.MtlItemName
                                    FROM SCM.PrDetail a
                                    INNER JOIN SCM.Inventory b ON a.InventoryId = b.InventoryId
                                    INNER JOIN PDM.MtlItem c ON a.MtlItemId = c.MtlItemId
                                    WHERE PrId = @PrId";
                            dynamicParameters.Add("PrId", PrId);

                            var PrDetailResult = sqlConnection.Query(sql, dynamicParameters);
                            #endregion

                            #region //取得請購單附檔資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT b.[FileName], b.FileContent, b.FileExtension, b.FileSize, b.Source, b.DeleteStatus
                                    FROM SCM.PrFile a
                                    INNER JOIN BAS.[File] b ON a.FileId = b.FileId
                                    WHERE PrId = @PrId";
                            dynamicParameters.Add("PrId", PrId);

                            var FileResult = sqlConnection.Query(sql, dynamicParameters);
                            #endregion

                            string PrErpNo = BaseHelper.RandomCode(11);

                            #region //INSERT 請購單單頭
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO SCM.PurchaseRequisition (PrNo, CompanyId, DepartmentId, UserId, PrErpPrefix, PrErpNo, Edition, PrDate, DocDate, PrRemark
                                    , TotalQty, Amount, PrStatus, SignupStaus, LockStaus, ConfirmStatus, BpmTransferStatus, TransferStatus, Priority
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.PrId
                                    VALUES (@PrNo, @CompanyId, @DepartmentId, @UserId, @PrErpPrefix, @PrErpNo, @Edition, @PrDate, @DocDate, @PrRemark
                                    , @TotalQty, @Amount, @PrStatus, @SignupStaus, @LockStaus, @ConfirmStatus, @BpmTransferStatus, @TransferStatus, @Priority
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    PrNo,
                                    CompanyId = CurrentCompany,
                                    DepartmentId,
                                    UserId = CreateBy,
                                    PrErpPrefix,
                                    PrErpNo,
                                    Edition = "0000",
                                    PrDate = DateTime.Now.ToString("yyyy-MM-dd"),
                                    DocDate = DateTime.Now.ToString("yyyy-MM-dd"),
                                    PrRemark,
                                    TotalQty,
                                    Amount,
                                    PrStatus = "N",
                                    SignupStaus = "0",
                                    LockStaus = "N",
                                    ConfirmStatus = "N",
                                    BpmTransferStatus = "N",
                                    TransferStatus = "N",
                                    Priority,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            rowsAffected += insertResult.Count();

                            int newPrId = -1;
                            foreach (var item in insertResult)
                            {
                                newPrId = item.PrId;
                            }
                            #endregion

                            #region //INSERT 請購單單身
                            foreach (var item in PrDetailResult)
                            {
                                #region //取得ERP庫存資訊
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT LTRIM(RTRIM(a.MC007)) InventoryQty
                                        FROM INVMC a
                                        WHERE a.MC001 = @MC001
                                        AND a.MC002 = @MC002";
                                dynamicParameters.Add("MC001", item.MtlItemNo);
                                dynamicParameters.Add("MC002", item.InventoryNo);

                                var result5 = sqlConnection2.Query(sql, dynamicParameters);

                                double InventoryQty = 0;
                                string mtlInventory = "目前尚無資料";
                                if (result5.Count() > 0)
                                {
                                    foreach (var item2 in result5)
                                    {
                                        InventoryQty = Convert.ToDouble(item.InventoryQty);
                                        #region //組MtlInventory
                                        List<MtlInventory> mtlInventories = new List<MtlInventory>
                                    {
                                        new MtlInventory
                                        {
                                            WAREHOUSE_NO = item.InventoryNo,
                                            WAREHOUSE_NAME = item.InventoryName,
                                            WAREHOUSE_QTY = InventoryQty
                                        }
                                    };
                                        #endregion

                                        mtlInventory = JsonConvert.SerializeObject(mtlInventories);
                                        mtlInventory = "{\"data\":" + mtlInventory + "}";
                                    }
                                }
                                #endregion

                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO SCM.PrDetail (PrId, PrSequence, MtlItemId, PrMtlItemName, PrMtlItemSpec, InventoryId, PrUomId, PrQty, DemandDate, SupplierId, PrCurrency, PrExchangeRate, PrUnitPrice, PrPrice, PrPriceTw, UrgentMtl, ProductionPlan, Project, SoDetailId
                                        , PoUserId, PoUomId, PoQty, PoCurrency, PoUnitPrice, PoPrice, LockStaus, PoStaus, PartialPurchaseStaus, InquiryStatus, TaxNo, Taxation, BusinessTaxRate, DetailMultiTax, TradeTerm, PrPriceQty, PrPriceUomId, MtlInventory, MtlInventoryQty, ConfirmStatus, ClosureStatus, PrDetailRemark, PoRemark
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.PrDetailId
                                        VALUES (@PrId, @PrSequence, @MtlItemId, @PrMtlItemName, @PrMtlItemSpec, @InventoryId, @PrUomId, @PrQty, @DemandDate, @SupplierId, @PrCurrency, @PrExchangeRate, @PrUnitPrice, @PrPrice, @PrPriceTw, @UrgentMtl, @ProductionPlan, @Project, @SoDetailId
                                        , @PoUserId, @PoUomId, @PoQty, @PoCurrency, @PoUnitPrice, @PoPrice, @LockStaus, @PoStaus, @PartialPurchaseStaus, @InquiryStatus, @TaxNo, @Taxation, @BusinessTaxRate, @DetailMultiTax, @TradeTerm, @PrPriceQty, @PrPriceUomId, @MtlInventory, @MtlInventoryQty, @ConfirmStatus, @ClosureStatus, @PrDetailRemark, @PoRemark
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        PrId = newPrId,
                                        item.PrSequence,
                                        item.MtlItemId,
                                        item.PrMtlItemName,
                                        item.PrMtlItemSpec,
                                        item.InventoryId,
                                        item.PrUomId,
                                        item.PrQty,
                                        DemandDate = DateTime.Now.ToString("yyyy-MM-dd"),
                                        item.SupplierId,
                                        item.PrCurrency,
                                        item.PrExchangeRate,
                                        item.PrUnitPrice,
                                        item.PrPrice,
                                        item.PrPriceTw,
                                        item.UrgentMtl,
                                        item.ProductionPlan,
                                        item.Project,
                                        item.SoDetailId,
                                        PoUserId = (int?)null,
                                        item.PoUomId,
                                        item.PoQty,
                                        item.PoCurrency,
                                        item.PoUnitPrice,
                                        item.PoPrice,
                                        LockStaus = "N",
                                        PoStaus = "N",
                                        PartialPurchaseStaus = "N",
                                        InquiryStatus = "1",
                                        item.TaxNo,
                                        item.Taxation,
                                        item.BusinessTaxRate,
                                        DetailMultiTax = "N",
                                        item.TradeTerm,
                                        item.PrPriceQty,
                                        item.PrPriceUomId,
                                        MtlInventory = mtlInventory,
                                        MtlInventoryQty = InventoryQty,
                                        ConfirmStatus = "N",
                                        ClosureStatus = "N",
                                        item.PrDetailRemark,
                                        item.PoRemark,
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });

                                var insertResult2 = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult2.Count();
                            }
                            #endregion

                            #region //INSERT 請購單附檔
                            if (FileResult.Count() > 0)
                            {
                                string DeviceIdentifierCode = BaseHelper.ClientIP();
                                foreach (var item in FileResult)
                                {
                                    #region //INSERT BAS.[File]
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO BAS.[File] (CompanyId, FileName, FileContent, FileExtension, FileSize, ClientIP, Source, DeleteStatus
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.FileId
                                            VALUES (@CompanyId, @FileName, @FileContent, @FileExtension, @FileSize, @ClientIP, @Source, @DeleteStatus
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            CompanyId = CurrentCompany,
                                            item.FileName,
                                            item.FileContent,
                                            item.FileExtension,
                                            item.FileSize,
                                            ClientIP = DeviceIdentifierCode,
                                            item.Source,
                                            item.DeleteStatus,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    var insertResult4 = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult4.Count();

                                    int FileId = -1;
                                    foreach (var item2 in insertResult4)
                                    {
                                        FileId = item2.FileId;
                                    }
                                    #endregion

                                    #region //INSERT SCM.PrFile
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO SCM.PrFile (PrId, FileId
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.PrFileId
                                            VALUES (@PrId, @FileId
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            PrId = newPrId,
                                            FileId,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    var insertResult3 = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult3.Count();
                                    #endregion
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

        #region //UpdatePrLogErrorMessage -- 更新請購單LOG錯誤訊息 -- Ann 2024-05-20
        public string UpdatePrLogErrorMessage(int PrLogId, string ErrorMessage)
        {
            try
            {
                int rowsAffected = 0;
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷訂單LOG紀錄資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.PrLog a 
                                WHERE a.PrLogId = @PrLogId";
                        dynamicParameters.Add("PrLogId", PrLogId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("請購單LOG資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.PrLog SET
                                ErrorMessage = @ErrorMessage,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE PrLogId = @PrLogId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ErrorMessage,
                                LastModifiedDate,
                                LastModifiedBy,
                                PrLogId
                            });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

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

        #region //UpdatePrmLogErrorMessage -- 更新請購變更單LOG錯誤訊息 -- Ann 2024-05-20
        public string UpdatePrmLogErrorMessage(int PrmLogId, string ErrorMessage)
        {
            try
            {
                int rowsAffected = 0;
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷訂單LOG紀錄資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT TOP 1 1
                                FROM SCM.PrmLog a 
                                WHERE a.PrmLogId = @PrmLogId";
                        dynamicParameters.Add("PrmLogId", PrmLogId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("請購變更單LOG資料錯誤!");
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.PrmLog SET
                                ErrorMessage = @ErrorMessage,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE PrmLogId = @PrmLogId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                ErrorMessage,
                                LastModifiedDate,
                                LastModifiedBy,
                                PrmLogId
                            });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);

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

        #region //Delete
        #region //DeletePurchaseRequisition -- 刪除請購單據資料 -- Ann 2023-01-17
        public string DeletePurchaseRequisition(int PrId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷請購單資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.PrStatus, a.UserId
                                FROM SCM.PurchaseRequisition a
                                WHERE a.PrId = @PrId";
                        dynamicParameters.Add("PrId", PrId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("請購單資料錯誤!");

                        foreach (var item in result)
                        {
                            if (item.PrStatus != "N") throw new SystemException("請購單狀態無法刪除!");
                            if (item.UserId != CreateBy) throw new SystemException("無法修改其他請購人之單據!");
                        }
                        #endregion

                        #region //先刪除關聯Table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE FROM SCM.PrFile
                                WHERE PrId = @PrId";
                        dynamicParameters.Add("PrId", PrId);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE FROM SCM.PrDetail
                                WHERE PrId = @PrId";
                        dynamicParameters.Add("PrId", PrId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.PurchaseRequisition
                                WHERE PrId = @PrId";
                        dynamicParameters.Add("PrId", PrId);

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

        #region //DeletePrDetail -- 刪除請購單據詳細資料 -- Ann 2023-01-17
        public string DeletePrDetail(int PrDetailId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷請購單詳細資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.PrQty, a.PrPrice, a.PrPriceTw
                                , b.PrStatus, b.UserId, b.PrId
                                FROM SCM.PrDetail a
                                INNER JOIN SCM.PurchaseRequisition b ON a.PrId = b.PrId
                                WHERE a.PrDetailId = @PrDetailId";
                        dynamicParameters.Add("PrDetailId", PrDetailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("請購單資料錯誤!");

                        int PrId = -1;
                        int PrQty = -1;
                        double PrPrice = -1;
                        double PrPriceTw = -1;
                        foreach (var item in result)
                        {
                            if (item.PrStatus != "N" && item.PrStatus != "E") throw new SystemException("請購單狀態無法刪除!");
                            if (item.UserId != CreateBy) throw new SystemException("無法修改其他請購人之單據!");
                            PrId = item.PrId;
                            PrQty = item.PrQty;
                            PrPrice = item.PrPrice;
                            PrPriceTw = item.PrPriceTw;
                        }
                        #endregion

                        #region //刪除附加table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.PrFile
                                WHERE PrDetailId = @PrDetailId";
                        dynamicParameters.Add("PrDetailId", PrDetailId);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.PrDetail
                                WHERE PrDetailId = @PrDetailId";
                        dynamicParameters.Add("PrDetailId", PrDetailId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //更新請購單總數量、總金額
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.PurchaseRequisition SET
                                TotalQty = TotalQty - @PrQty,
                                Amount = Amount - @PrPriceTw,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE PrId = @PrId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                PrQty,
                                PrPriceTw,
                                LastModifiedDate,
                                LastModifiedBy,
                                PrId
                            });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //重新調整單身序號
                        dynamicParameters = new DynamicParameters();
                        sql = @"With UpdateSort As
                                (
                                    SELECT PrSequence,
                                    ROW_NUMBER() OVER(ORDER BY PrSequence) NewPrSequence
                                    FROM SCM.PrDetail
                                    WHERE PrId = @PrId
                                )
                                UPDATE SCM.PrDetail
                                SET PrSequence = Right('0000' + Cast(NewPrSequence as varchar), 4)
                                FROM SCM.PrDetail
                                INNER JOIN UpdateSort ON SCM.PrDetail.PrSequence = UpdateSort.PrSequence
                                WHERE PrId = @PrId";
                        dynamicParameters.Add("PrId", PrId);

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

        #region //DeletePrModification -- 刪除請購變更單據資料 -- Ann 2023-02-09
        public string DeletePrModification(int PrmId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷請購變更單資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.PrmStatus, a.UserId
                                FROM SCM.PrModification a
                                WHERE a.PrmId = @PrmId";
                        dynamicParameters.Add("PrmId", PrmId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("請購變更單資料錯誤!");

                        foreach (var item in result)
                        {
                            if (item.PrmStatus != "N") throw new SystemException("請購變更單已拋轉，無法刪除!");
                            if (item.UserId != CreateBy) throw new SystemException("無法修改其他請購人之單據!");
                        }
                        #endregion

                        #region //先刪除關聯Table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE FROM SCM.PrmDetail
                                WHERE PrmId = @PrmId";
                        dynamicParameters.Add("PrmId", PrmId);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE FROM SCM.PrmFile
                                WHERE PrmId = @PrmId";
                        dynamicParameters.Add("PrmId", PrmId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.PrModification
                                WHERE PrmId = @PrmId";
                        dynamicParameters.Add("PrmId", PrmId);

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

        #region //DeletePrmDetail -- 刪除請購變更單據詳細資料 -- Ann 2023-02-09
        public string DeletePrmDetail(int PrmDetailId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷請購變更單詳細資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.PrQty, a.PrPrice, a.PrPriceTw
                                , b.PrmStatus, b.UserId, b.PrmId
                                FROM SCM.PrmDetail a
                                INNER JOIN SCM.PrModification b ON a.PrmId = b.PrmId
                                WHERE a.PrmDetailId = @PrmDetailId";
                        dynamicParameters.Add("PrmDetailId", PrmDetailId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("請購單資料錯誤!");

                        int PrmId = -1;
                        int PrQty = -1;
                        double PrPrice = -1;
                        double PrPriceTw = -1;
                        foreach (var item in result)
                        {
                            if (item.PrmStatus != "N" && item.PrmStatus != "E") throw new SystemException("請購變更單已拋轉，無法刪除!");
                            if (item.UserId != CreateBy) throw new SystemException("無法修改其他請購人之單據!");
                            PrmId = item.PrmId;
                            PrQty = item.PrQty;
                            PrPrice = item.PrPrice;
                            PrPriceTw = item.PrPriceTw;
                        }
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.PrmDetail
                                WHERE PrmDetailId = @PrmDetailId";
                        dynamicParameters.Add("PrmDetailId", PrmDetailId);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //更新請購單總數量、總金額
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.PrModification SET
                                TotalQty = TotalQty - @PrQty,
                                Amount = Amount - @PrPriceTw,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE PrmId = @PrmId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                PrQty,
                                PrPriceTw,
                                LastModifiedDate,
                                LastModifiedBy,
                                PrmId
                            });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //重新調整單身序號
                        dynamicParameters = new DynamicParameters();
                        sql = @"With UpdateSort As
                                (
                                    SELECT PrmSequence,
                                    ROW_NUMBER() OVER(ORDER BY PrmSequence) NewPrmSequence
                                    FROM SCM.PrmDetail
                                    WHERE PrmId = @PrmId
                                )
                                UPDATE SCM.PrmDetail
                                SET PrmSequence = Right('0000' + Cast(NewPrmSequence as varchar), 4)
                                FROM SCM.PrmDetail
                                INNER JOIN UpdateSort ON SCM.PrmDetail.PrmSequence = UpdateSort.PrmSequence
                                WHERE PrmId = @PrmId";
                        dynamicParameters.Add("PrmId", PrmId);

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
        #region //SendPrErrorMail -- 寄送請購/請購變更異常通知信件 -- Ann 2023-09-06
        public string SendPrErrorMail(int DocId, string DocType)
        {
            try
            {
                int rowsAffected = 0;
                int DocUserId = -1;
                string PrNo = "";
                string Remark = "";
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        if (DocType == "PR")
                        {
                            #region //判斷請購單資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.PrNo, a.CreateBy DocUserId, a.PrRemark
                                    FROM SCM.PurchaseRequisition a
                                    WHERE a.PrId = @PrId";
                            dynamicParameters.Add("PrId", DocId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("請購單資料錯誤!");

                            foreach (var item in result)
                            {
                                DocUserId = item.DocUserId;
                                PrNo = item.PrNo;
                                Remark = item.Remark;
                            }
                            #endregion
                        }
                        else if (DocType == "PRM")
                        {
                            #region //判斷請購變更單資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.CreateBy DocUserId, a.PrmRemark
                                    , b.PrNo
                                    FROM SCM.PrModification a
                                    INNER JOIN SCM.PurchaseRequisition b ON a.PrId = b.PrId
                                    WHERE a.PrmId = @PrmId";
                            dynamicParameters.Add("PrmId", DocId);

                            var result = sqlConnection.Query(sql, dynamicParameters);
                            if (result.Count() <= 0) throw new SystemException("請購變更單資料錯誤!");
                            
                            foreach (var item in result)
                            {
                                DocUserId = item.DocUserId;
                                PrNo = item.PrNo;
                                Remark = item.Remark;
                            }
                            #endregion
                        }
                        else
                        {
                            throw new SystemException("請購單類型錯誤!!");
                        }

                        #region //取得單據建立人員資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT (b.DepartmentName + '-' + a.UserName) + ':' + a.Email UserMailInfo
                                , a.Email
                                FROM BAS.[User] a 
                                INNER JOIN BAS.Department b ON a.DepartmentId = b.DepartmentId
                                WHERE a.UserId = @DocUserId";
                        dynamicParameters.Add("DocUserId", DocUserId);

                        var EmailResult = sqlConnection.Query(sql, dynamicParameters);

                        if (EmailResult.Count() <= 0) throw new SystemException("人員基本資料錯誤!!");

                        string UserMailInfo = "";
                        foreach (var item in EmailResult)
                        {
                            if (item.Email == "")
                            {
                                #region //Response
                                jsonResponse = JObject.FromObject(new
                                {
                                    status = "success",
                                    msg = "(" + rowsAffected + " rows affected)"
                                });
                                #endregion

                                return jsonResponse.ToString();
                            }

                            UserMailInfo = item.UserMailInfo;
                        }
                        #endregion

                        #region //寄送異常信件
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
                        dynamicParameters.Add("SettingSchema", "PrErrorMail");
                        dynamicParameters.Add("SettingNo", "Y");

                        var resultMailTemplate = sqlConnection.Query(sql, dynamicParameters);
                        if (resultMailTemplate.Count() <= 0) throw new SystemException("Mail設定錯誤!");
                        #endregion

                        foreach (var item in resultMailTemplate)
                        {
                            string mailSubject = item.MailSubject,
                                mailContent = HttpUtility.UrlDecode(item.MailContent),
                                mailTo = item.MailTo;

                            mailTo = mailTo + ";" + UserMailInfo;

                            #region //Mail內容
                            mailContent = mailContent.Replace("[FinishDate]", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                            mailContent = mailContent.Replace("[DocType]", DocType == "PR" ? "請購單" : "請購變更單");
                            mailContent = mailContent.Replace("[No]", PrNo);
                            mailContent = mailContent.Replace("[Remark]", Remark);
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

        #region /Modal
        #region //MtlInventory
        public class MtlInventory
        {
            public string WAREHOUSE_NO { get; set; }
            public string WAREHOUSE_NAME { get; set; }
            public double? WAREHOUSE_QTY { get; set; }
        }

        public class MtlItemUomInfo
        {
            public int? UomId { get; set; }
            public string UomText { get; set; }
            public string MtlItemNo { get; set; }
            public string UomNo { get; set; }
        }
        #endregion
        #endregion

        #region //ForPrDetailExcel 批量請購單身用

        #region //GetErpInventoryQtyForExcel -- 取得ERP庫存資料 -- GPAI 2024-01-29
        public string GetErpInventoryQtyForExcel(string MtlItemNo, int InventoryId, string InventoryNo
            , string OrderBy, int PageIndex, int PageSize)
        {
            if (MtlItemNo.Length <= 0) throw new SystemException("品號不能為空!!");
            //if (InventoryId <= 0 && InventoryNo.Length <= 0) throw new SystemException("庫別不能為空!!");

            try
            {
                List<INVMC> iNVMCs = new List<INVMC>();
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
                        #region //找InventoryNo
                        if (InventoryId > 0)
                        {
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.InventoryNo
                                    FROM SCM.Inventory a
                                    WHERE a.InventoryId = @InventoryId
                                    AND CompanyId = @CompanyId";
                            dynamicParameters.Add("InventoryId", InventoryId);
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            var inventoryResult = sqlConnection.Query(sql, dynamicParameters);

                            if (inventoryResult.Count() <= 0) throw new SystemException("查無此庫別資料!");

                            foreach (var item in inventoryResult)
                            {
                                InventoryNo = item.InventoryNo;
                            }
                        }
                        #endregion

                        dynamicParameters = new DynamicParameters();
                        sqlQuery.mainKey = "a.MC001, a.MC002";
                        sqlQuery.auxKey = "";
                        sqlQuery.columns =
                            @", LTRIM(RTRIM(a.MC001)) MtlItemNo, LTRIM(RTRIM(a.MC002)) InventoryNo, LTRIM(RTRIM(a.MC003)) InventorySite, LTRIM(RTRIM(a.MC007)) InventoryQty";
                        sqlQuery.mainTables =
                            @"FROM INVMC a";
                        sqlQuery.auxTables = "";
                        string queryCondition = @"";
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "MtlItemNo", @" AND a.MC001 LIKE '%' + @MtlItemNo + '%'", MtlItemNo);
                        BaseHelper.SqlParameter(ref queryCondition, ref dynamicParameters, "InventoryNo", @" AND a.MC002 = @InventoryNo", InventoryNo);
                        sqlQuery.conditions = queryCondition;
                        sqlQuery.orderBy = OrderBy.Length > 0 ? OrderBy : "a.MC001";
                        sqlQuery.pageIndex = PageIndex;
                        sqlQuery.pageSize = PageSize;

                        iNVMCs = BaseHelper.SqlQuery<INVMC>(sqlConnection2, dynamicParameters, sqlQuery);

                        #region //回MES查資料
                        int i = 0;
                        foreach (var item in iNVMCs)
                        {
                            #region //查庫別資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.InventoryId, a.InventoryName
                                    FROM SCM.Inventory a
                                    WHERE a.InventoryNo = @InventoryNo
                                    AND CompanyId = @CompanyId";
                            dynamicParameters.Add("InventoryNo", item.InventoryNo);
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            var inventoryResult = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item2 in inventoryResult)
                            {
                                iNVMCs[i].InventoryId = item2.InventoryId;
                                iNVMCs[i].InventoryName = item2.InventoryName;
                            }
                            #endregion

                            #region //查品號資料
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.MtlItemName
                                    FROM PDM.MtlItem a
                                    WHERE a.MtlItemNo = @MtlItemNo
                                    AND CompanyId = @CompanyId";
                            dynamicParameters.Add("MtlItemNo", item.MtlItemNo);
                            dynamicParameters.Add("CompanyId", CurrentCompany);
                            var mtlItemResult = sqlConnection.Query(sql, dynamicParameters);

                            foreach (var item2 in mtlItemResult)
                            {
                                iNVMCs[i].MtlItemName = item2.MtlItemName;
                            }
                            #endregion
                            i++;
                        }
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            data = iNVMCs
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

        #region //GetMtlItemTotalUomInfoFoeExcel -- 取得品號所有可用單位資料 -- GPAI 2024-01-29
        public string GetMtlItemTotalUomInfoFoeExcel(string MtlItemNo)
        {
            try
            {
                if (MtlItemNo == "") throw new SystemException("【品號】不能為空!");

                List<MtlItemUomInfo> mtlItemUomInfos = new List<MtlItemUomInfo>();

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
                                , b.UomId, b.UomNo
                                , (b.UomNo + ' ' + b.UomName) UomText
                                FROM PDM.MtlItem a
                                INNER JOIN PDM.UnitOfMeasure b ON a.InventoryUomId = b.UomId
                                WHERE a.MtlItemNo = @MtlItemNo";
                        dynamicParameters.Add("MtlItemNo", MtlItemNo);

                        var MtlItemResult = sqlConnection.Query(sql, dynamicParameters);

                        //string MtlItemNo = "";
                        string UomText = "";
                        foreach (var item in MtlItemResult)
                        {
                            MtlItemNo = item.MtlItemNo;
                            UomText = item.UomText;

                            MtlItemUomInfo uomInfos = new MtlItemUomInfo()
                            {
                                UomId = item.UomId,
                                UomText = item.UomText,
                                UomNo = item.UomNo
                            };
                            mtlItemUomInfos.Add(uomInfos);
                        }
                        #endregion

                        #region //找ERP INVMD
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT LTRIM(RTRIM(MD002)) MD002
                                FROM INVMD
                                WHERE MD001 = @MD001";
                        dynamicParameters.Add("MD001", MtlItemNo);

                        var INVMDResult = sqlConnection2.Query(sql, dynamicParameters);

                        foreach (var item in INVMDResult)
                        {
                            bool checkInventoryFlag = false;
                            foreach (var item2 in mtlItemUomInfos)
                            {
                                if (item2.UomNo == item.MD002)
                                {
                                    checkInventoryFlag = true;
                                    break;
                                }
                            }

                            if (checkInventoryFlag != false)
                            {
                                continue;
                            }

                            #region //合併回來MES
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.UomId
                                    , a.UomNo
                                    , (a.UomNo + ' ' + a.UomName) UomText
                                    FROM PDM.UnitOfMeasure a 
                                    WHERE a.UomNo = @UomNo
                                    AND a.CompanyId = @CompanyId";
                            dynamicParameters.Add("UomNo", item.MD002);
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            var UnitOfMeasureResult = sqlConnection.Query(sql, dynamicParameters);

                            if (UnitOfMeasureResult.Count() <= 0) throw new SystemException("ERP轉換單位【" + item.MD002 + "】資料錯誤!!!");

                            foreach (var item2 in UnitOfMeasureResult)
                            {
                                MtlItemUomInfo uomInfos = new MtlItemUomInfo()
                                {
                                    UomId = item2.UomId,
                                    UomText = item2.UomText,
                                    UomNo = item2.UomNo
                                };
                                mtlItemUomInfos.Add(uomInfos);
                            }
                            #endregion
                        }
                        #endregion

                        #region //Response
                        jsonResponse = JObject.FromObject(new
                        {
                            status = "success",
                            data = mtlItemUomInfos
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

        #region //ADD

        #region //AddPrDetailForExcel -- 新增請購單詳細資料 -- GPAI 2024-01-29
        public string AddPrDetailForExcel(int PrId, string PrSequence, string MtlItemNo, string PrMtlItemName, string PrMtlItemSpec, string InventoryNo, string PrUomNo, int PrQty, string DemandDate
            , string SupplierNo, string PrCurrency, string PrExchangeRate, double PrUnitPrice, double PrPrice, double PrPriceTw
            , string UrgentMtl, string ProductionPlan, string Project, int SoDetailId, string PrDetailRemark, string PrFile)
        {
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
                            if (PrQty <= 0) throw new SystemException("【請購數量】不能為空!");
                            if (DemandDate.Length <= 0) throw new SystemException("【需求日期】不能為空!");
                            if (PrExchangeRate.Length <= 0) throw new SystemException("【匯率】不能為空!");
                            if (PrPrice <= 0) throw new SystemException("【請購金額】不能為空!");
                            if (PrPriceTw <= 0) throw new SystemException("【本幣金額】不能為空!");
                            if (UrgentMtl.Length <= 0) throw new SystemException("【是否急料】不能為空!");
                            if (ProductionPlan.Length <= 0) throw new SystemException("【是否納入生產計畫】不能為空!");
                            if (PrMtlItemName.Length <= 0) throw new SystemException("【品名】不能為空!");

                            #region //確認請購單資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT ISNULL(a.TotalQty, 0) TotalQty, ISNULL(a.Amount, 0) Amount
                                    , a.UserId
                                    FROM SCM.PurchaseRequisition a
                                    WHERE a.PrId = @PrId";
                            dynamicParameters.Add("PrId", PrId);

                            var purchaseRequisitionResult = sqlConnection.Query(sql, dynamicParameters);

                            if (purchaseRequisitionResult.Count() <= 0) throw new SystemException("【請購單】資料錯誤!");

                            double TotalQty = -1;
                            double Amount = -1;
                            foreach (var item in purchaseRequisitionResult)
                            {
                                if (item.UserId != CreateBy) throw new SystemException("無法修改其他請購人之單據!");
                                TotalQty = item.TotalQty;
                                Amount = item.Amount;
                            }
                            #endregion

                            #region //確認請購單序號是否重複
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM SCM.PrDetail a
                                    WHERE a.PrId = @PrId
                                    AND a.PrSequence = @PrSequence";
                            dynamicParameters.Add("PrId", PrId);
                            dynamicParameters.Add("PrSequence", PrSequence);

                            var result = sqlConnection.Query(sql, dynamicParameters);

                            if (result.Count() > 0) throw new SystemException("【請購單序號】重複!");
                            #endregion

                            #region //判斷品號資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.MtlItemNo, a.MtlItemId
                                    FROM PDM.MtlItem a
                                    WHERE a.MtlItemNo = @MtlItemNo";
                            dynamicParameters.Add("MtlItemNo", MtlItemNo);

                            var result2 = sqlConnection.Query(sql, dynamicParameters);

                            if (result2.Count() <= 0) throw new SystemException("【品號】資料錯誤!");
                            if (result2.Count() > 1) throw new SystemException("【品號】資料錯誤!(重複品號)");

                            int MtlItemId = 0;
                            foreach (var item in result2)
                            {
                                MtlItemId = item.MtlItemId;
                            }
                            #endregion

                            #region //判斷ERP品號生效日與失效日
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(MB030)) MB030, LTRIM(RTRIM(MB031)) MB031
                                    FROM INVMB
                                    WHERE MB001 = @MB001";
                            dynamicParameters.Add("MB001", MtlItemNo);

                            var INVMBResult = sqlConnection2.Query(sql, dynamicParameters);

                            if (INVMBResult.Count() <= 0) throw new SystemException("此品號不存在於ERP中!!");

                            foreach (var item in INVMBResult)
                            {
                                if (item.MB030 != "" && item.MB030 != null)
                                {
                                    #region //判斷日期需大於或等於生效日
                                    string EffectiveDate = item.MB030;
                                    string effYear = EffectiveDate.Substring(0, 4);
                                    string effMonth = EffectiveDate.Substring(4, 2);
                                    string effDay = EffectiveDate.Substring(6, 2);
                                    DateTime effFullDate = Convert.ToDateTime(effYear + "-" + effMonth + "-" + effDay);
                                    int effresult = DateTime.Compare(CreateDate, effFullDate);
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
                                    int effresult = DateTime.Compare(CreateDate, effFullDate);
                                    if (effresult > 0) throw new SystemException("此品號已失效，無法使用!!");
                                    #endregion
                                }
                            }
                            #endregion

                            #region //判斷庫別資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT a.InventoryNo, a.InventoryName, a.InventoryId
                                    FROM SCM.Inventory a
                                    WHERE a.InventoryNo = @InventoryNo";
                            dynamicParameters.Add("InventoryNo", InventoryNo);

                            var result3 = sqlConnection.Query(sql, dynamicParameters);

                            if (result3.Count() <= 0) throw new SystemException("【庫別】資料錯誤!");

                            int InventoryId = 0;
                            string InventoryName = "";
                            foreach (var item in result3)
                            {
                                InventoryId = item.InventoryId;
                                InventoryName = item.InventoryName;
                            }
                            #endregion

                            #region //判斷單位資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM PDM.UnitOfMeasure a
                                    WHERE a.UomNo = @UomNo";
                            dynamicParameters.Add("UomNo", PrUomNo);

                            var result4 = sqlConnection.Query(sql, dynamicParameters);

                            if (result4.Count() <= 0) throw new SystemException("【單位】資料錯誤!");//PrUomId
                            int PrUomId = 0;
                            foreach (var item in result4)
                            {
                                PrUomId = item.UomId;
                            }
                            #endregion

                            #region //取得ERP庫存資訊
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT LTRIM(RTRIM(a.MC007)) InventoryQty
                                    FROM INVMC a
                                    WHERE a.MC001 = @MC001
                                    AND a.MC002 = @MC002";
                            dynamicParameters.Add("MC001", MtlItemNo);
                            dynamicParameters.Add("MC002", InventoryNo);

                            var result5 = sqlConnection2.Query(sql, dynamicParameters);

                            double InventoryQty = 0;
                            string mtlInventory = "目前尚無資料";
                            if (result5.Count() > 0)
                            {
                                foreach (var item in result5)
                                {
                                    InventoryQty = Convert.ToDouble(item.InventoryQty);
                                    #region //組MtlInventory
                                    List<MtlInventory> mtlInventories = new List<MtlInventory>
                                    {
                                        new MtlInventory
                                        {
                                            WAREHOUSE_NO = InventoryNo,
                                            WAREHOUSE_NAME = InventoryName,
                                            WAREHOUSE_QTY = InventoryQty
                                        }
                                    };
                                    #endregion

                                    mtlInventory = JsonConvert.SerializeObject(mtlInventories);
                                    mtlInventory = "{\"data\":" + mtlInventory + "}";
                                }
                            }
                            #endregion

                            #region //檢查供應商資料是否正確
                            string TaxNo = "";
                            string Taxation = "";
                            string TradeTerm = "";
                            double? BusinessTaxRate = -1;
                            int SupplierId = 0;
                            if (SupplierNo != "")
                            {
                                #region //供應商
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.TaxNo, a.Taxation, a.TradeTerm, a.SupplierId
                                        FROM SCM.Supplier a
                                        WHERE a.SupplierNo = @SupplierNo";
                                dynamicParameters.Add("SupplierNo", SupplierNo);

                                var result6 = sqlConnection.Query(sql, dynamicParameters);

                                if (result6.Count() <= 0) throw new SystemException("【供應商】資料錯誤!");//SupplierId

                                foreach (var item in result6)
                                {
                                    TaxNo = item.TaxNo;
                                    Taxation = item.Taxation;
                                    TradeTerm = item.TradeTerm;
                                    SupplierId = item.SupplierId;
                                }
                                #endregion

                                #region //查詢營業稅額資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT ROUND(LTRIM(RTRIM(NN004)),2) BusinessTaxRate
                                        FROM CMSNN 
                                        WHERE NN001 = @TaxNo";
                                dynamicParameters.Add("TaxNo", TaxNo);

                                var businessTaxRateResult = sqlConnection2.Query(sql, dynamicParameters);

                                if (businessTaxRateResult.Count() <= 0) throw new SystemException("【課稅別】資料錯誤!");

                                foreach (var item in businessTaxRateResult)
                                {
                                    BusinessTaxRate = item.BusinessTaxRate;
                                }
                                #endregion
                            }
                            #endregion

                            #region //判斷幣別資料是否正確
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT TOP 1 1
                                    FROM CMSMF
                                    WHERE MF001 = @MF001";
                            dynamicParameters.Add("MF001", PrCurrency);

                            var result7 = sqlConnection2.Query(sql, dynamicParameters);
                            if (result7.Count() <= 0) throw new SystemException("【幣別】資料有誤!");
                            #endregion

                            #region //判斷專案代碼資料是否正確
                            if (Project.Length > 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 1
                                    FROM CMSNB
                                    WHERE NB001 = @NB001";
                                dynamicParameters.Add("NB001", Project);

                                var result8 = sqlConnection2.Query(sql, dynamicParameters);
                                if (result8.Count() <= 0) throw new SystemException("【專案代碼】資料有誤!");
                            }
                            #endregion

                            #region //確認訂單資料是否正確
                            if (SoDetailId > 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT MtlItemId, ConfirmStatus
                                        FROM SCM.SoDetail
                                        WHERE SoDetailId = @SoDetailId";
                                dynamicParameters.Add("SoDetailId", SoDetailId);

                                var result9 = sqlConnection.Query(sql, dynamicParameters);
                                if (result9.Count() <= 0) throw new SystemException("【訂單】資料有誤!");

                                foreach (var item in result9)
                                {
                                    if (item.ConfirmStatus != "Y") throw new SystemException("訂單尚未核單，無法綁定!!");
                                    //if (item.MtlItemId != MtlItemId) throw new SystemException("訂單品號與請購品號不同!!");
                                }
                            }
                            #endregion

                            #region //INSERT SCM.PrDetail
                            dynamicParameters = new DynamicParameters();
                            sql = @"INSERT INTO SCM.PrDetail (PrId, PrSequence, MtlItemId, PrMtlItemName, PrMtlItemSpec, InventoryId, PrUomId, PrQty, DemandDate, SupplierId, PrCurrency, PrExchangeRate, PrUnitPrice, PrPrice, PrPriceTw, UrgentMtl, ProductionPlan, Project, SoDetailId
                                    , PoUserId, PoUomId, PoQty, PoCurrency, PoUnitPrice, PoPrice, LockStaus, PoStaus, PartialPurchaseStaus, InquiryStatus, TaxNo, Taxation, BusinessTaxRate, DetailMultiTax, TradeTerm, PrPriceQty, PrPriceUomId, MtlInventory, MtlInventoryQty, ConfirmStatus, ClosureStatus, PrDetailRemark, PoRemark
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.PrDetailId
                                    VALUES (@PrId, @PrSequence, @MtlItemId, @PrMtlItemName, @PrMtlItemSpec, @InventoryId, @PrUomId, @PrQty, @DemandDate, @SupplierId, @PrCurrency, @PrExchangeRate, @PrUnitPrice, @PrPrice, @PrPriceTw, @UrgentMtl, @ProductionPlan, @Project, @SoDetailId
                                    , @PoUserId, @PoUomId, @PoQty, @PoCurrency, @PoUnitPrice, @PoPrice, @LockStaus, @PoStaus, @PartialPurchaseStaus, @InquiryStatus, @TaxNo, @Taxation, @BusinessTaxRate, @DetailMultiTax, @TradeTerm, @PrPriceQty, @PrPriceUomId, @MtlInventory, @MtlInventoryQty, @ConfirmStatus, @ClosureStatus, @PrDetailRemark, @PoRemark
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    PrId,
                                    PrSequence,
                                    MtlItemId,
                                    PrMtlItemName,
                                    PrMtlItemSpec,
                                    InventoryId,
                                    PrUomId,
                                    PrQty,
                                    DemandDate,
                                    SupplierId,
                                    PrCurrency,
                                    PrExchangeRate,
                                    PrUnitPrice,
                                    PrPrice,
                                    PrPriceTw,
                                    UrgentMtl,
                                    ProductionPlan,
                                    Project,
                                    SoDetailId,
                                    PoUserId = (int?)null,
                                    PoUomId = PrUomId,
                                    PoQty = PrQty,
                                    PoCurrency = PrCurrency,
                                    PoUnitPrice = PrUnitPrice,
                                    PoPrice = PrPrice,
                                    LockStaus = "N",
                                    PoStaus = "N",
                                    PartialPurchaseStaus = "N",
                                    InquiryStatus = "1",
                                    TaxNo = TaxNo != "" ? TaxNo : null,
                                    Taxation = Taxation != "" ? Taxation : null,
                                    BusinessTaxRate = BusinessTaxRate > 0 ? BusinessTaxRate : (double?)null,
                                    DetailMultiTax = "N",
                                    TradeTerm = TradeTerm != "" ? TradeTerm : null,
                                    PrPriceQty = PrQty,
                                    PrPriceUomId = PrUomId,
                                    MtlInventory = mtlInventory,
                                    MtlInventoryQty = InventoryQty,
                                    ConfirmStatus = "N",
                                    ClosureStatus = "N",
                                    PrDetailRemark,
                                    PoRemark = PrDetailRemark,
                                    CreateDate,
                                    LastModifiedDate,
                                    CreateBy,
                                    LastModifiedBy
                                });

                            var insertResult = sqlConnection.Query(sql, dynamicParameters);

                            int rowsAffected = insertResult.Count();

                            int PrDetailId = -1;
                            foreach (var item in insertResult)
                            {
                                PrDetailId = item.PrDetailId;
                            }
                            #endregion

                            #region //更新單頭總採購數量及金額
                            dynamicParameters = new DynamicParameters();
                            sql = @"UPDATE SCM.PurchaseRequisition SET
                                    TotalQty = TotalQty + @PrQty,
                                    Amount = Amount + @PrPriceTw,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE PrId = @PrId";
                            dynamicParameters.AddDynamicParams(
                                new
                                {
                                    PrQty,
                                    PrPriceTw,
                                    LastModifiedDate,
                                    LastModifiedBy,
                                    PrId
                                });

                            rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                            #endregion

                            #region //新增File
                            if (PrFile.Length > 0)
                            {
                                string[] prFiles = PrFile.Split(',');
                                foreach (var file in prFiles)
                                {
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO SCM.PrFile (PrId, PrDetailId, FileId
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.PrFileId
                                            VALUES (@PrId, @PrDetailId, @FileId
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            PrId,
                                            PrDetailId,
                                            FileId = file,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    var insertResult3 = sqlConnection.Query(sql, dynamicParameters);

                                    rowsAffected += insertResult3.Count();
                                }
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

        #region //UpdatePrDetailForExcel --
        public string UpdatePrDetailForExcel(int PrId, string ExcelJson)
        {
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


                        List<int> qcitemidList = new List<int>();

                        int rowsAffected = 0;

                        #region //確認請購單資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT ISNULL(a.TotalQty, 0) TotalQty, ISNULL(a.Amount, 0) Amount
                                    , a.UserId
                                    FROM SCM.PurchaseRequisition a
                                    WHERE a.PrId = @PrId";
                        dynamicParameters.Add("PrId", PrId);

                        var purchaseRequisitionResult = sqlConnection.Query(sql, dynamicParameters);

                        if (purchaseRequisitionResult.Count() <= 0) throw new SystemException("【請購單】資料錯誤!");

                        double TotalQty = -1;
                        double Amount = -1;
                        foreach (var item in purchaseRequisitionResult)
                        {
                            if (item.UserId != CreateBy) throw new SystemException("無法修改其他請購人之單據!");
                            TotalQty = item.TotalQty;
                            Amount = item.Amount;
                        }
                        #endregion

                        #region//撈取Id
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT PrDetailId
                                FROM SCM.PrDetail a
                                WHERE a.PrId=@PrId";
                        dynamicParameters.Add("PrId", PrId);
                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() > 0)
                        {
                            foreach (var item in result)
                            {
                                qcitemidList.Add(item.PrDetailId);
                            }
                        }
                        #endregion



                        var spreadsheetJson = JObject.Parse(ExcelJson);

                        #region //解析Spreadsheet Data
                        foreach (var item in spreadsheetJson["spreadsheetInfo"])
                        {

                            string PrSequence = item["PrSequence"] != null ? item["PrSequence"].ToString() : throw new SystemException("【資料維護不完整】PrSequence,請重新確認~~");
                            string MtlItemNo = item["MtlItemNo"] != null ? item["MtlItemNo"].ToString() : throw new SystemException("【資料維護不完整】MtlItemNo,請重新確認~~");
                            string PrMtlItemName = item["PrMtlItemName"] != null ? item["PrMtlItemName"].ToString() : throw new SystemException("【資料維護不完整】PrMtlItemName,請重新確認~~");
                            string PrMtlItemSpec = item["PrMtlItemSpec"] != null ? item["PrMtlItemSpec"].ToString() : throw new SystemException("【資料維護不完整】PrMtlItemSpec,請重新確認~~");
                            string InventoryNo = item["InventoryNo"] != null ? item["InventoryNo"].ToString() : throw new SystemException("【資料維護不完整】InventoryNo,請重新確認~~");
                            string InventoryName = item["InventoryName"] != null ? item["InventoryName"].ToString() : throw new SystemException("【資料維護不完整】InventoryName,請重新確認~~");
                            //decimal InventoryQty = item["InventoryAmout"] != null ? Convert.ToDecimal(item["InventoryAmout"]) : throw new SystemException("【資料維護不完整】MtlInventoryAmoutItemNo,請重新確認~~");
                            string PrUomNo = item["PrUomNo"] != null ? item["PrUomNo"].ToString() : throw new SystemException("【資料維護不完整】PrUomNo,請重新確認~~");
                            string DemandDate = item["DemandDate"] != null ? item["DemandDate"].ToString() : throw new SystemException("【資料維護不完整】DemandDate,請重新確認~~");
                            string SupplierNo = item["SupplierNo"] != null ? item["SupplierNo"].ToString() : "";
                            string PrCurrency = item["PrCurrency"] != null ? item["PrCurrency"].ToString() : throw new SystemException("【資料維護不完整】PrCurrency,請重新確認~~");
                            string PrExchangeRate = item["PrExchangeRate"] != null ? item["PrExchangeRate"].ToString() : throw new SystemException("【資料維護不完整】PrExchangeRate,請重新確認~~");
                            int PrQty = item["PrQty"] != null ? Convert.ToInt32(item["PrQty"]) : throw new SystemException("【資料維護不完整】PrQty,請重新確認~~");
                            decimal PrUnitPrice = item["PrUnitPrice"] != null ? Convert.ToDecimal(item["PrUnitPrice"]) : throw new SystemException("【資料維護不完整】PrUnitPrice,請重新確認~~");
                            decimal PrPrice = item["PrPrice"] != null ? Convert.ToDecimal(item["PrPrice"]) : throw new SystemException("【資料維護不完整】PrPrice,請重新確認~~");
                            decimal PrPriceTw = item["PrPriceTw"] != null ? Convert.ToDecimal(item["PrPriceTw"]) : throw new SystemException("【資料維護不完整】PrPriceTw,請重新確認~~");
                            string UrgentMtl = item["UrgentMtl"] != null ? item["UrgentMtl"].ToString() : throw new SystemException("【資料維護不完整】UrgentMtl,請重新確認~~");
                            string ProductionPlan = item["ProductionPlan"] != null ? item["ProductionPlan"].ToString() : throw new SystemException("【資料維護不完整】ProductionPlan,請重新確認~~");
                            string Project = item["Project"] != null ? item["Project"].ToString() : "";
                            string PrDetailRemark = item["PrDetailRemark"] != null ? item["PrDetailRemark"].ToString() : "";
                            string PoErpPrefixNo = item["PoErpPrefixNo"] != null ? item["PoErpPrefixNo"].ToString() : "";
                            string ConfirmStatus = item["ConfirmStatus"] != null ? item["ConfirmStatus"].ToString() : "";
                            string SoErpFullNo = item["SoErpFullNo"] != null ? item["SoErpFullNo"].ToString() : "";
                            string SoCustomerName = item["SoCustomerName"] != null ? item["SoCustomerName"].ToString() : "";
                            string SoSequence = item["SoSequence"] != null ? item["SoSequence"].ToString() : "";
                            string SoCustomerMtlItemNo = item["SoCustomerMtlItemNo"] != null ? item["SoCustomerMtlItemNo"].ToString() : "";
                            string File = item["File"] != null ? item["File"].ToString() : "";
                            int ID = item["PrdetailID"] != null ? Convert.ToInt32(item["PrdetailID"]) : -1;


                            string PrFile = File != "" ? File.Split(' ')[1]: "";

                            string TaxNo = "";
                            string Taxation = "";
                            string TradeTerm = "";
                            double? BusinessTaxRate = -1;
                            double InventoryQty = 0;
                            string mtlInventory = "目前尚無資料";

                            if (PrDetailRemark.Length > 255) throw new SystemException("【備註】長度錯誤!");

                            //套件的日期值是 從1899/12/30 為第0天 所以將取的數值 從1899/12/30 開始加即可推導出正確日期格式
                            DateTime initialDate = new DateTime(1899, 12, 30);
                            DateTime? demandDate = new DateTime();
                            if (DemandDate != null)
                            {
                                demandDate = initialDate.AddDays(Convert.ToInt32(Convert.ToDouble(DemandDate)));
                            }

                            #region //確認品號資料是否正確
                            int MtlItemId = -1;
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT *
                                FROM PDM.MtlItem a 
                                WHERE a.MtlItemNo = @MtlItemNo";
                            dynamicParameters.Add("MtlItemNo", MtlItemNo);

                            var MtlItemResult = sqlConnection.Query(sql, dynamicParameters);
                            if (!MtlItemResult.Any()) throw new SystemException("找不到【" + MtlItemNo + "】品號資料,請重新確認~~");
                            foreach (var item1 in MtlItemResult)
                            {
                                MtlItemId = item1.MtlItemId;
                            }
                            #endregion

                            #region //確認庫別資料是否正確
                            int InventoryId = -1;
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT *
                                FROM SCM.Inventory a 
                                WHERE a.InventoryNo = @InventoryNo AND a.CompanyId = @CompanyId ";
                            dynamicParameters.Add("InventoryNo", InventoryNo);
                            dynamicParameters.Add("CompanyId", CurrentCompany);


                            var InventoryResult = sqlConnection.Query(sql, dynamicParameters);
                            if (!InventoryResult.Any()) throw new SystemException("找不到【" + InventoryNo + "】庫別資料,請重新確認~~");
                            foreach (var item1 in InventoryResult)
                            {
                                InventoryId = item1.InventoryId;
                            }
                            #endregion

                            #region //確認單位資料是否正確
                            int PrUomId = -1;
                            dynamicParameters = new DynamicParameters();
                            sql = @"SELECT *
                                FROM PDM.UnitOfMeasure a 
                                WHERE a.UomNo = @UomNo AND a.CompanyId = @CompanyId ";
                            dynamicParameters.Add("UomNo", PrUomNo);
                            dynamicParameters.Add("CompanyId", CurrentCompany);

                            var UomResult = sqlConnection.Query(sql, dynamicParameters);
                            if (!UomResult.Any()) throw new SystemException("找不到【" + PrUomNo + "】單位資料,請重新確認~~");
                            foreach (var item1 in UomResult)
                            {
                                PrUomId = item1.UomId;
                            }
                            #endregion

                            #region //確認供應商資料是否正確
                            int SupplierId = -1;

                            if (SupplierNo.Length > 1) {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT *
                                FROM SCM.Supplier a 
                                WHERE a.SupplierNo = @SupplierNo AND a.CompanyId = @CompanyId ";
                                dynamicParameters.Add("SupplierNo", SupplierNo.Split(' ')[0]);
                                dynamicParameters.Add("CompanyId", CurrentCompany);

                                var SupplierIdResult = sqlConnection.Query(sql, dynamicParameters);
                                if (!SupplierIdResult.Any()) throw new SystemException("找不到【" + SupplierNo + "】供應商資料,請重新確認~~");
                                foreach (var item1 in SupplierIdResult)
                                {
                                    SupplierId = item1.SupplierId;
                                }
                            }
                            
                            #endregion

                            #region //確認訂單資料是否正確

                            int SoDetailId = -1;

                            if (SoErpFullNo.Length > 0) {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.SoDetailId, a.MtlItemId
                                        FROM SCM.SoDetail a
										INNER JOIN SCM.SaleOrder b ON a.SoId = b.SoId
                                        WHERE (b.SoErpPrefix + '-' + b.SoErpNo) = @SoErpFullNo AND b.CompanyId = @CompanyId AND a.MtlItemId = @MtlItemId";
                                dynamicParameters.Add("SoErpFullNo", SoErpFullNo);
                                dynamicParameters.Add("CompanyId", CurrentCompany);
                                dynamicParameters.Add("MtlItemId", MtlItemId);


                                var SoDetailIdResult = sqlConnection.Query(sql, dynamicParameters);
                                if (!SoDetailIdResult.Any()) throw new SystemException("找不到【" + SoErpFullNo + "】訂單資料,請重新確認~~");
                                foreach (var item0 in SoDetailIdResult)
                                {
                                    if (item0.MtlItemId != MtlItemId) throw new SystemException("訂單品號與請購品號不同!!");
                                    SoDetailId = item0.SoDetailId;
                                }
                            }
                            


                            #endregion

                            using (SqlConnection sqlConnection2 = new SqlConnection(ErpConnectionStrings))
                            {
                                if (ID > 0)
                                {


                                    #region //判斷請購單詳細資料是否正確
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.PrQty, a.PrPrice, a.PrPriceTw
                                    FROM SCM.PrDetail a
                                    WHERE a.PrDetailId = @PrDetailId";
                                    dynamicParameters.Add("PrDetailId", ID);

                                    var prDetailIdResult = sqlConnection.Query(sql, dynamicParameters);
                                    if (prDetailIdResult.Count() <= 0) throw new SystemException("請購單詳細資料錯誤!");

                                    double OrgPrQty = -1;
                                    double OrgPrPrice = -1;
                                    double OrgPrPriceTw = -1;
                                    foreach (var item2 in prDetailIdResult)
                                    {
                                        OrgPrQty = item2.PrQty;
                                        OrgPrPrice = item2.PrPrice;
                                        OrgPrPriceTw = item2.PrPriceTw;
                                    }
                                    #endregion

                                    #region //判斷請購單資料是否正確
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.PrStatus, a.UserId
                                    FROM SCM.PurchaseRequisition a
                                    WHERE a.PrId = @PrId";
                                    dynamicParameters.Add("PrId", PrId);

                                    var result2 = sqlConnection.Query(sql, dynamicParameters);
                                    if (result2.Count() <= 0) throw new SystemException("請購單資料錯誤!");

                                    foreach (var item2 in result2)
                                    {
                                        if (item2.PrStatus != "N" && item2.PrStatus != "E") throw new SystemException("請購單狀態無法更改!");
                                        if (item2.UserId != CreateBy) throw new SystemException("無法修改其他請購人之單據!");
                                    }
                                    #endregion

                                    #region //判斷品號資料是否正確
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.MtlItemNo, a.MtlItemName, a.MtlItemSpec
                                    FROM PDM.MtlItem a
                                    WHERE a.MtlItemId = @MtlItemId";
                                    dynamicParameters.Add("MtlItemId", MtlItemId);

                                    var result3 = sqlConnection.Query(sql, dynamicParameters);

                                    if (result3.Count() <= 0) throw new SystemException("【品號】資料錯誤!");


                                    #endregion

                                    #region //判斷ERP品號生效日與失效日
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT LTRIM(RTRIM(MB030)) MB030, LTRIM(RTRIM(MB031)) MB031
                                    FROM INVMB
                                    WHERE MB001 = @MB001";
                                    dynamicParameters.Add("MB001", MtlItemNo);

                                    var INVMBResult = sqlConnection2.Query(sql, dynamicParameters);

                                    foreach (var item3 in INVMBResult)
                                    {
                                        if (item3.MB030 != "" && item3.MB030 != null)
                                        {
                                            #region //判斷日期需大於或等於生效日
                                            string EffectiveDate = item3.MB030;
                                            string effYear = EffectiveDate.Substring(0, 4);
                                            string effMonth = EffectiveDate.Substring(4, 2);
                                            string effDay = EffectiveDate.Substring(6, 2);
                                            DateTime effFullDate = Convert.ToDateTime(effYear + "-" + effMonth + "-" + effDay);
                                            int effresult = DateTime.Compare(CreateDate, effFullDate);
                                            if (effresult < 0) throw new SystemException("此品號尚未生效，無法使用!!");
                                            #endregion
                                        }

                                        if (item3.MB031 != "" && item3.MB031 != null)
                                        {
                                            #region //判斷日期需小於或等於失效日
                                            string ExpirationDate = item3.MB031;
                                            string effYear = ExpirationDate.Substring(0, 4);
                                            string effMonth = ExpirationDate.Substring(4, 2);
                                            string effDay = ExpirationDate.Substring(6, 2);
                                            DateTime effFullDate = Convert.ToDateTime(effYear + "-" + effMonth + "-" + effDay);
                                            int effresult = DateTime.Compare(CreateDate, effFullDate);
                                            if (effresult > 0) throw new SystemException("此品號已失效，無法使用!!");
                                            #endregion
                                        }
                                    }
                                    #endregion

                                    #region //判斷庫別資料是否正確
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT a.InventoryNo, a.InventoryName
                                    FROM SCM.Inventory a
                                    WHERE a.InventoryId = @InventoryId";
                                    dynamicParameters.Add("InventoryId", InventoryId);

                                    var result4 = sqlConnection.Query(sql, dynamicParameters);

                                    if (result4.Count() <= 0) throw new SystemException("【庫別】資料錯誤!");


                                    foreach (var item4 in result4)
                                    {
                                        InventoryNo = item4.InventoryNo;
                                        InventoryName = item4.InventoryName;
                                    }
                                    #endregion

                                    #region //判斷單位資料是否正確
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                    FROM PDM.UnitOfMeasure a
                                    WHERE a.UomId = @UomId";
                                    dynamicParameters.Add("UomId", PrUomId);

                                    var result5 = sqlConnection.Query(sql, dynamicParameters);

                                    if (result5.Count() <= 0) throw new SystemException("【單位】資料錯誤!");
                                    #endregion

                                    #region //取得ERP庫存資訊
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT LTRIM(RTRIM(a.MC007)) InventoryQty
                                    FROM INVMC a
                                    WHERE a.MC001 = @MC001
                                    AND a.MC002 = @MC002";
                                    dynamicParameters.Add("MC001", MtlItemNo);
                                    dynamicParameters.Add("MC002", InventoryNo);

                                    var result6 = sqlConnection2.Query(sql, dynamicParameters);


                                    if (result6.Count() > 0)
                                    {
                                        foreach (var item6 in result6)
                                        {
                                            InventoryQty = Convert.ToDouble(item6.InventoryQty);
                                            #region //組MtlInventory
                                            List<MtlInventory> mtlInventories = new List<MtlInventory>
                                    {
                                        new MtlInventory
                                        {
                                            WAREHOUSE_NO = InventoryNo,
                                            WAREHOUSE_NAME = InventoryName,
                                            WAREHOUSE_QTY = InventoryQty
                                        }
                                    };
                                            #endregion

                                            mtlInventory = JsonConvert.SerializeObject(mtlInventories);
                                            mtlInventory = "{\"data\":" + mtlInventory + "}";
                                        }
                                    }
                                    #endregion

                                    #region //檢查供應商資料是否正確

                                    if (SupplierId > 0)
                                    {
                                        #region //供應商
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT a.TaxNo, a.Taxation, a.TradeTerm
                                        FROM SCM.Supplier a
                                        WHERE a.SupplierId = @SupplierId";
                                        dynamicParameters.Add("SupplierId", SupplierId);

                                        var result7 = sqlConnection.Query(sql, dynamicParameters);

                                        if (result7.Count() <= 0) throw new SystemException("【供應商】資料錯誤!");

                                        foreach (var item7 in result7)
                                        {
                                            TaxNo = item7.TaxNo;
                                            Taxation = item7.Taxation;
                                            TradeTerm = item7.TradeTerm;
                                        }
                                        #endregion

                                        #region //查詢營業稅額資料
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT ROUND(LTRIM(RTRIM(NN004)),2) BusinessTaxRate
                                        FROM CMSNN 
                                        WHERE NN001 = @TaxNo";
                                        dynamicParameters.Add("TaxNo", TaxNo);

                                        var businessTaxRateResult = sqlConnection2.Query(sql, dynamicParameters);

                                        if (businessTaxRateResult.Count() <= 0) throw new SystemException("【課稅別】資料錯誤!");

                                        foreach (var item8 in businessTaxRateResult)
                                        {
                                            BusinessTaxRate = item8.BusinessTaxRate;
                                        }
                                        #endregion
                                    }
                                    #endregion

                                    #region //判斷幣別資料是否正確
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                    FROM CMSMF
                                    WHERE MF001 = @MF001";
                                    dynamicParameters.Add("MF001", PrCurrency);

                                    var result9 = sqlConnection2.Query(sql, dynamicParameters);
                                    if (result9.Count() <= 0) throw new SystemException("【幣別】資料有誤!");
                                    #endregion

                                    #region //判斷專案代碼資料是否正確
                                    if (Project.Length > 0)
                                    {
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT TOP 1 1
                                    FROM CMSNB
                                    WHERE NB001 = @NB001";
                                        dynamicParameters.Add("NB001", Project);

                                        var result10 = sqlConnection2.Query(sql, dynamicParameters);
                                        if (result10.Count() <= 0) throw new SystemException("【專案代碼】資料有誤!");
                                    }
                                    #endregion

                                    #region //驗證計算價格是否有誤
                                    //PrCurrency 幣別
                                    //PrExchangeRate 匯率
                                    //PrQty 請購數 self
                                    //PrUnitPrice 單價 self
                                    //string Today = DateTime.Now.ToString("yyyyMMdd");

                                    //匯率
                                    decimal exchangerate = -1;
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 LTRIM(RTRIM(MG001)) Currency, ROUND(LTRIM(RTRIM(MG004)),3) ExchangeRateNameForMwe, ROUND(LTRIM(RTRIM(MG005)),3) ExchangeRateName, LTRIM(RTRIM(MG002)) StartDate
                                    FROM CMSMG 
                                    WHERE 1=1 AND MG001 = @MG001
                                    Order By StartDate DESC";
                                    dynamicParameters.Add("MG001", PrCurrency);
                                    var currencyresult1 = sqlConnection2.Query(sql, dynamicParameters);
                                    if (currencyresult1.Count() > 0)
                                    {
                                        foreach (var currencyitem in currencyresult1)
                                        {
                                            exchangerate = Convert.ToDecimal (currencyitem.ExchangeRateName);
                                        }
                                    }
                                    //幣別小數點進位資料
                                    decimal unitdecimal = -1;
                                    decimal totaldecimal = -1;

                                    sql = @"SELECT LTRIM(RTRIM(MF001)) CurrencyNo, LTRIM(RTRIM(MF002)) CurrencyName
                                    , LTRIM(RTRIM(MF003)) UnitDecimal, LTRIM(RTRIM(MF004)) TotalDecimal
                                    FROM CMSMF
                                    WHERE 1=1  AND MF001 = @MF001";
                                    dynamicParameters.Add("MF001", PrCurrency);
                                    var currencyresult2 = sqlConnection2.Query(sql, dynamicParameters);


                                    if (currencyresult2.Count() > 0) {
                                        foreach (var currencyitem in currencyresult2)
                                        {
                                            unitdecimal = Convert.ToDecimal (currencyitem.UnitDecimal);
                                            totaldecimal = Convert.ToDecimal (currencyitem.TotalDecimal);
                                        }
                                    }

                                    var unitdecimalnumber = "1";
                                    var totaldecimalnumber = "1";

                                    for (var i = 1; i <= unitdecimal; i++)
                                    {
                                        unitdecimalnumber += "0";
                                    }

                                    for (var i = 1; i <= totaldecimal; i++)
                                    {
                                        totaldecimalnumber += "0";
                                    }

                                    //請購金額 prPrice prQty * prUnitPrice * parseInt(TotalDecimalNumber)) / TotalDecimalNumber
                                    decimal correctprPrice = PrQty * PrUnitPrice * Convert.ToInt32(totaldecimalnumber) / Convert.ToInt32(totaldecimalnumber);

                                    //本幣金額 prPriceTw prQty * prUnitPrice * prExchangeRate * parseInt(TotalDecimalNumber)) / TotalDecimalNumber
                                    decimal correctprPriceTw = PrQty * PrUnitPrice * exchangerate * Convert.ToInt32(totaldecimalnumber) / Convert.ToInt32(totaldecimalnumber);
                                    if (correctprPrice != PrPrice) throw new SystemException("【請購金額】資料有誤!");
                                    if (correctprPriceTw != PrPriceTw) throw new SystemException("【本幣金額】資料有誤!");

                                    #endregion

                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE SCM.PrDetail SET
                                    MtlItemId = @MtlItemId,
                                    PrMtlItemName = @PrMtlItemName,
                                    PrMtlItemSpec = @PrMtlItemSpec,
                                    InventoryId = @InventoryId,
                                    PrUomId = @PrUomId,
                                    PrQty = @PrQty,
                                    DemandDate = @DemandDate,
                                    SupplierId = @SupplierId,
                                    PrCurrency = @PrCurrency,
                                    PrExchangeRate = @PrExchangeRate,
                                    PrUnitPrice = @PrUnitPrice,
                                    PrPrice = @PrPrice,
                                    PrPriceTw = @PrPriceTw,
                                    UrgentMtl = @UrgentMtl,
                                    ProductionPlan = @ProductionPlan,
                                    Project = @Project,
                                    PrPriceQty = @PrQty,
                                    PrPriceUomId = @PrUomId,
                                    MtlInventoryQty = @MtlInventoryQty,
                                    PoUomId = @PoUomId,
                                    PoQty = @PoQty,
                                    PoCurrency = @PoCurrency,
                                    PoUnitPrice = @PoUnitPrice,
                                    PoPrice = @PoPrice,
                                    PrDetailRemark = @PrDetailRemark,
                                    TaxNo = @TaxNo,
                                    Taxation = @Taxation,
                                    TradeTerm = @TradeTerm,
                                    BusinessTaxRate = @BusinessTaxRate,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE PrDetailId = @PrDetailId";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            MtlItemId,
                                            PrMtlItemName,
                                            PrMtlItemSpec,
                                            InventoryId,
                                            PrUomId,
                                            PrQty,
                                            demandDate,
                                            SupplierId,
                                            PrCurrency,
                                            PrExchangeRate,
                                            PrUnitPrice,
                                            PrPrice,
                                            PrPriceTw,
                                            UrgentMtl,
                                            ProductionPlan,
                                            Project,
                                            PrPriceQty = PrQty,
                                            PrPriceUomId = PrUomId,
                                            MtlInventoryQty = InventoryQty,
                                            PoUomId = PrUomId,
                                            PoQty = PrQty,
                                            PoCurrency = PrCurrency,
                                            PoUnitPrice = PrUnitPrice,
                                            PoPrice = PrPrice,
                                            PrDetailRemark,
                                            TaxNo = TaxNo != "" ? TaxNo : null,
                                            Taxation = Taxation != "" ? Taxation : null,
                                            TradeTerm = TradeTerm != "" ? TradeTerm : null,
                                            BusinessTaxRate = BusinessTaxRate > 0 ? BusinessTaxRate : (double?)null,
                                            LastModifiedDate,
                                            LastModifiedBy,
                                            PrDetailId = ID
                                        });

                                    rowsAffected = sqlConnection.Execute(sql, dynamicParameters);

                                    #region //更新請購單頭
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE SCM.PurchaseRequisition SET
                                    TotalQty = TotalQty - @OrgPrQty + @PrQty,
                                    Amount = Amount - @OrgPrPriceTw + @PrPriceTw,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE PrId = @PrId";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            OrgPrQty,
                                            PrQty,
                                            OrgPrPriceTw,
                                            PrPriceTw,
                                            LastModifiedDate,
                                            LastModifiedBy,
                                            PrId
                                        });

                                    rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    if (PrFile.Length > 0)
                                    {
                                        #region //先將原本的砍掉
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"DELETE FROM SCM.PrFile
                                        WHERE PrId = @PrId
                                        AND PrDetailId = @PrDetailId";
                                        dynamicParameters.Add("PrId", PrId);
                                        dynamicParameters.Add("PrDetailid", ID);

                                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                        #endregion

                                        string[] prFiles = PrFile.Split(',');
                                        foreach (var file in prFiles)
                                        {
                                            #region //更新SCM.PrFile
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"INSERT INTO SCM.PrFile (PrId, PrDetailId, FileId
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.PrFileId
                                            VALUES (@PrId, @PrDetailid, @FileId
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                            dynamicParameters.AddDynamicParams(
                                                new
                                                {
                                                    PrId,
                                                    PrDetailId = ID,
                                                    FileId = file,
                                                    CreateDate,
                                                    LastModifiedDate,
                                                    CreateBy,
                                                    LastModifiedBy
                                                });

                                            var insertResult3 = sqlConnection.Query(sql, dynamicParameters);

                                            rowsAffected += insertResult3.Count();
                                            #endregion
                                        }
                                    }
                                    else
                                    {
                                        #region //將原本的砍掉
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"DELETE FROM SCM.PrFile
                                        WHERE PrId = @PrId
                                        AND PrDetailId = @PrDetailId";
                                        dynamicParameters.Add("PrId", PrId);
                                        dynamicParameters.Add("PrDetailId", ID);

                                        rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                        #endregion
                                    }


                                }
                                else {
                                    #region //INSERT SCM.PrDetail
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"INSERT INTO SCM.PrDetail (PrId, PrSequence, MtlItemId, PrMtlItemName, PrMtlItemSpec, InventoryId, PrUomId, PrQty, DemandDate, SupplierId, PrCurrency, PrExchangeRate, PrUnitPrice, PrPrice, PrPriceTw, UrgentMtl, ProductionPlan, Project, SoDetailId
                                    , PoUserId, PoUomId, PoQty, PoCurrency, PoUnitPrice, PoPrice, LockStaus, PoStaus, PartialPurchaseStaus, InquiryStatus, TaxNo, Taxation, BusinessTaxRate, DetailMultiTax, TradeTerm, PrPriceQty, PrPriceUomId, MtlInventory, MtlInventoryQty, ConfirmStatus, ClosureStatus, PrDetailRemark, PoRemark
                                    , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                    OUTPUT INSERTED.PrDetailId
                                    VALUES (@PrId, @PrSequence, @MtlItemId, @PrMtlItemName, @PrMtlItemSpec, @InventoryId, @PrUomId, @PrQty, @DemandDate, @SupplierId, @PrCurrency, @PrExchangeRate, @PrUnitPrice, @PrPrice, @PrPriceTw, @UrgentMtl, @ProductionPlan, @Project, @SoDetailId
                                    , @PoUserId, @PoUomId, @PoQty, @PoCurrency, @PoUnitPrice, @PoPrice, @LockStaus, @PoStaus, @PartialPurchaseStaus, @InquiryStatus, @TaxNo, @Taxation, @BusinessTaxRate, @DetailMultiTax, @TradeTerm, @PrPriceQty, @PrPriceUomId, @MtlInventory, @MtlInventoryQty, @ConfirmStatus, @ClosureStatus, @PrDetailRemark, @PoRemark
                                    , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            PrId,
                                            PrSequence,
                                            MtlItemId,
                                            PrMtlItemName,
                                            PrMtlItemSpec,
                                            InventoryId,
                                            PrUomId,
                                            PrQty,
                                            demandDate,
                                            SupplierId,
                                            PrCurrency,
                                            PrExchangeRate,
                                            PrUnitPrice,
                                            PrPrice,
                                            PrPriceTw,
                                            UrgentMtl,
                                            ProductionPlan,
                                            Project,
                                            SoDetailId,
                                            PoUserId = (int?)null,
                                            PoUomId = PrUomId,
                                            PoQty = PrQty,
                                            PoCurrency = PrCurrency,
                                            PoUnitPrice = PrUnitPrice,
                                            PoPrice = PrPrice,
                                            LockStaus = "N",
                                            PoStaus = "N",
                                            PartialPurchaseStaus = "N",
                                            InquiryStatus = "1",
                                            TaxNo = TaxNo != "" ? TaxNo : null,
                                            Taxation = Taxation != "" ? Taxation : null,
                                            BusinessTaxRate = BusinessTaxRate > 0 ? BusinessTaxRate : (double?)null,
                                            DetailMultiTax = "N",
                                            TradeTerm = TradeTerm != "" ? TradeTerm : null,
                                            PrPriceQty = PrQty,
                                            PrPriceUomId = PrUomId,
                                            MtlInventory = mtlInventory,
                                            MtlInventoryQty = InventoryQty,
                                            ConfirmStatus = "N",
                                            ClosureStatus = "N",
                                            PrDetailRemark,
                                            PoRemark = PrDetailRemark,
                                            CreateDate,
                                            LastModifiedDate,
                                            CreateBy,
                                            LastModifiedBy
                                        });

                                    var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                     rowsAffected = insertResult.Count();

                                    int PrDetailId = -1;
                                    foreach (var item12 in insertResult)
                                    {
                                        PrDetailId = item12.PrDetailId;
                                    }
                                    #endregion

                                    #region //更新單頭總採購數量及金額
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"UPDATE SCM.PurchaseRequisition SET
                                    TotalQty = TotalQty + @PrQty,
                                    Amount = Amount + @PrPriceTw,
                                    LastModifiedDate = @LastModifiedDate,
                                    LastModifiedBy = @LastModifiedBy
                                    WHERE PrId = @PrId";
                                    dynamicParameters.AddDynamicParams(
                                        new
                                        {
                                            PrQty,
                                            PrPriceTw,
                                            LastModifiedDate,
                                            LastModifiedBy,
                                            PrId
                                        });

                                    rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                    #endregion

                                    #region //新增File
                                    if (PrFile.Length > 0)
                                    {
                                        string[] prFiles = PrFile.Split(',');
                                        foreach (var file in prFiles)
                                        {
                                            dynamicParameters = new DynamicParameters();
                                            sql = @"INSERT INTO SCM.PrFile (PrId, PrDetailId, FileId
                                            , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                            OUTPUT INSERTED.PrFileId
                                            VALUES (@PrId, @PrDetailId, @FileId
                                            , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                            dynamicParameters.AddDynamicParams(
                                                new
                                                {
                                                    PrId,
                                                    PrDetailId,
                                                    FileId = file,
                                                    CreateDate,
                                                    LastModifiedDate,
                                                    CreateBy,
                                                    LastModifiedBy
                                                });

                                            var insertResult3 = sqlConnection.Query(sql, dynamicParameters);

                                            rowsAffected += insertResult3.Count();
                                        }
                                    }
                                    #endregion
                                }
                            }
                        

                        }
                        #endregion


                        //#region //以上傳資料為最新,故要移除沒有再上傳資料中的
                        //if (qcitemidList.Count() > 0)
                        //{
                        //    foreach (var item in qcitemidList)
                        //    {
                        //        #region //刪子table
                        //        sql = @" DELETE a FROM MES.ProcessParameterQcItem a
                        //         INNER JOIN QMS.QcItem b ON a.QcItemId = b.QcItemId
                        //         WHERE b.QcItemId = @QcItemId AND a.ParameterId = @ParameterId";
                        //        dynamicParameters.Add("QcItemId", item);
                        //        dynamicParameters.Add("ParameterId", ParameterId);


                        //        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        //        #endregion

                        //    }
                        //}
                        //#endregion



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

        #region //DeleteAllPrDetail -- 刪除ALL請購單據詳細資料 -- GPAI 2024-01-29
        public string DeleteAllPrDetail(int PrId)
        {
            try
            {
                using (TransactionScope transactionScope = new TransactionScope())
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //判斷請購單詳細資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.PrQty, a.PrPrice, a.PrPriceTw
                                , b.PrStatus, b.UserId, b.PrId
                                FROM SCM.PrDetail a
                                INNER JOIN SCM.PurchaseRequisition b ON a.PrId = b.PrId
                                WHERE a.PrId = @PrId";
                        dynamicParameters.Add("PrId", PrId);

                        var result = sqlConnection.Query(sql, dynamicParameters);
                        if (result.Count() <= 0) throw new SystemException("請購單資料錯誤!");

                        int PrQty = -1;
                        double PrPrice = -1;
                        double PrPriceTw = -1;
                        foreach (var item in result)
                        {
                            if (item.PrStatus != "N" && item.PrStatus != "E") throw new SystemException("請購單狀態無法刪除!");
                            if (item.UserId != CreateBy) throw new SystemException("無法修改其他請購人之單據!");
                            PrQty = item.PrQty;
                            PrPrice = item.PrPrice;
                            PrPriceTw = item.PrPriceTw;
                        }
                        #endregion

                        #region //刪除附加table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.PrFile
                                WHERE PrId = @PrId";
                        dynamicParameters.Add("PrId", PrId);

                        int rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //刪除主要table
                        dynamicParameters = new DynamicParameters();
                        sql = @"DELETE SCM.PrDetail
                                WHERE PrId = @PrId";
                        dynamicParameters.Add("PrId", PrId);

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        #region //更新請購單總數量、總金額
                        dynamicParameters = new DynamicParameters();
                        sql = @"UPDATE SCM.PurchaseRequisition SET
                                TotalQty = TotalQty - @PrQty,
                                Amount = Amount - @PrPriceTw,
                                LastModifiedDate = @LastModifiedDate,
                                LastModifiedBy = @LastModifiedBy
                                WHERE PrId = @PrId";
                        dynamicParameters.AddDynamicParams(
                            new
                            {
                                PrQty,
                                PrPriceTw,
                                LastModifiedDate,
                                LastModifiedBy,
                                PrId
                            });

                        rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        #endregion

                        //#region //重新調整單身序號
                        //dynamicParameters = new DynamicParameters();
                        //sql = @"With UpdateSort As
                        //        (
                        //            SELECT PrSequence,
                        //            ROW_NUMBER() OVER(ORDER BY PrSequence) NewPrSequence
                        //            FROM SCM.PrDetail
                        //            WHERE PrId = @PrId
                        //        )
                        //        UPDATE SCM.PrDetail
                        //        SET PrSequence = Right('0000' + Cast(NewPrSequence as varchar), 4)
                        //        FROM SCM.PrDetail
                        //        INNER JOIN UpdateSort ON SCM.PrDetail.PrSequence = UpdateSort.PrSequence
                        //        WHERE PrId = @PrId";
                        //dynamicParameters.Add("PrId", PrId);

                        //rowsAffected += sqlConnection.Execute(sql, dynamicParameters);
                        //#endregion

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

        #region //GetInventoryQtyForExcel -- 取得品號庫存資料 -- GPAI 2024-01-29
        public string GetInventoryQtyForExcel(string MtlItemNo, int InventoryId)
        {
            try
            {
                //if (MtlItemNo.Length <= 0) throw new SystemException("【品號】不能為空!");
                if (InventoryId <= 0) throw new SystemException("【庫別】不能為空!");

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
                        sql = @"SELECT a.MtlItemNo, a.MtlItemId
                                FROM PDM.MtlItem a
                                WHERE a.MtlItemNo = @MtlItemNo";
                        dynamicParameters.Add("MtlItemNo", MtlItemNo);

                        var MtlItemResult = sqlConnection.Query(sql, dynamicParameters);

                        int MtlItemId = 0;
                        foreach (var item in MtlItemResult)
                        {
                            MtlItemId = item.MtlItemId;
                        }
                        #endregion

                        #region //確認庫別資料是否正確
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.InventoryNo
                                FROM SCM.Inventory a 
                                WHERE a.InventoryId = @InventoryId";
                        dynamicParameters.Add("InventoryId", InventoryId);

                        var InventoryResult = sqlConnection.Query(sql, dynamicParameters);

                        string InventoryNo = "";
                        foreach (var item in InventoryResult)
                        {
                            InventoryNo = item.InventoryNo;
                        }
                        #endregion

                        #region //取得ERP庫存資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT SUM(MC007) MC007
                                FROM INVMC
                                WHERE MC001 = @MC001";
                        dynamicParameters.Add("MC001", MtlItemNo);

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


        #endregion

        #region //INVMC
        public class INVMC
        {
            public string MtlItemNo { get; set; }
            public string MtlItemName { get; set; }
            public int? InventoryId { get; set; }
            public string InventoryNo { get; set; }
            public string InventoryName { get; set; }
            public string InventorySite { get; set; }
            public double? InventoryQty { get; set; }
            public int? TotalCount { get; set; }
        }
        #endregion

        #endregion

        #region //For批量開立請購單
        #region //BatchAddPurchaseRequisition -- 批量新增請購單 -- Ann 2024-08-02
        public string BatchAddPurchaseRequisition(string UploadJson = "")
        {
            try
            {
                if (UploadJson.Length <= 0) throw new SystemException("上傳檔案內容為空，請嘗試重新操作!!");
                List<string> TotalErrorMessages = new List<string>();
                List<int> PrIdList = new List<int>();
                int rowsAffected = 0;
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
                            JObject uploadJson = JObject.Parse(UploadJson);
                            string tempPrDep = "";
                            string tempPrErpPrefix = "";
                            int PrId = -1;
                            bool mrBool = false; //是否建立新請購單頭
                            Regex regex = new Regex(@"^-?\d+(\.\d+)?$");
                            int count = 1;
                            foreach (var item in uploadJson["uploadInfo"])
                            {
                                mrBool = false;

                                #region //初始化ErrorMessage
                                List<string> ErrorMessages = new List<string>();
                                string defaultErrorMsg = "請購單【" + item["PrErpPrefix"].ToString() + "-" + item["PrErpNo"] + "】<br>";
                                ErrorMessages.Add(defaultErrorMsg);
                                #endregion

                                #region //資料驗證
                                #region //請購單別
                                string PrErpPrefix = "";
                                if (item["PrErpPrefix"] != null)
                                {
                                    if (item["PrErpPrefix"].Type.ToString() != "Null")
                                    {
                                        PrErpPrefix = item["PrErpPrefix"].ToString();

                                        if (PrErpPrefix != tempPrErpPrefix)
                                        {
                                            mrBool = true;
                                            tempPrErpPrefix = PrErpPrefix;
                                        }

                                        #region //確認單別資料是否正確
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT LTRIM(RTRIM(a.MQ001)) MQ001, LTRIM(RTRIM(a.MQ002)) MQ002
                                                , (LTRIM(RTRIM(a.MQ001)) + ' ' + LTRIM(RTRIM(a.MQ002))) PrErpPrefix
                                                FROM CMSMQ a
                                                LEFT JOIN CMSMU b ON a.MQ001 = b.MU001
                                                WHERE 1=1
                                                AND ((b.MU003 = 'DS' AND a.MQ029 = 'Y') OR a.MQ029 = 'N')
                                                AND a.MQ001 >= N''''
                                                AND a.MQ003 = '31'
                                                AND a.MQ001 = @PrErpPrefix";
                                        dynamicParameters.Add("PrErpPrefix", PrErpPrefix);

                                        var CMSMQResult = sqlConnection2.Query(sql, dynamicParameters);

                                        if (CMSMQResult.Count() <= 0)
                                        {
                                            ErrorMessages.Add("(" + ErrorMessages.Count + ") 請購單別【" + PrErpPrefix + "】資料錯誤!!<br>");
                                        }
                                        #endregion
                                    }
                                    else
                                    {
                                        ErrorMessages.Add("(" + ErrorMessages.Count + ") 【請購單別】不能為空!!<br>");
                                    }
                                }
                                else
                                {
                                    ErrorMessages.Add("(" + ErrorMessages.Count + ") 【請購單別】不能為空!!<br>");
                                }
                                #endregion

                                #region //請購單號
                                string PrErpNo = "";
                                if (item["PrErpNo"] != null)
                                {
                                    if (item["PrErpNo"].Type.ToString() != "Null")
                                    {
                                        PrErpNo = item["PrErpNo"].ToString();
                                    }
                                    else
                                    {
                                        PrErpNo = BaseHelper.RandomCode(11);
                                    }
                                }
                                else
                                {
                                    PrErpNo = BaseHelper.RandomCode(11);
                                }
                                #endregion

                                #region //請購部門
                                string DepartmentNo = "";
                                int DepartmentId = -1;
                                if (item["DepartmentNo"] != null)
                                {
                                    if (item["DepartmentNo"].Type.ToString() != "Null")
                                    {
                                        DepartmentNo = item["DepartmentNo"].ToString();

                                        if (DepartmentNo != tempPrDep)
                                        {
                                            mrBool = true;
                                            tempPrDep = DepartmentNo;
                                        }

                                        #region //確認部門資料是否正確
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT a.DepartmentId
                                                FROM BAS.Department a
                                                WHERE a.DepartmentNo = @DepartmentNo
                                                AND a.Status = 'A'
                                                AND a.CompanyId = @CompanyId";
                                        dynamicParameters.Add("DepartmentNo", item["DepartmentNo"].ToString());
                                        dynamicParameters.Add("CompanyId", CurrentCompany);

                                        var DepartmentResult = sqlConnection.Query(sql, dynamicParameters);
                                        if (DepartmentResult.Count() <= 0)
                                        {
                                            ErrorMessages.Add("(" + ErrorMessages.Count + ") 【請購部門】資料錯誤或狀態非啟用中，請重新輸入!!<br>");
                                        }

                                        foreach (var item2 in DepartmentResult)
                                        {
                                            DepartmentId = item2.DepartmentId;
                                        }
                                        #endregion
                                    }
                                    else
                                    {
                                        ErrorMessages.Add("(" + ErrorMessages.Count + ") 請購部門不能為空!!<br>");
                                    }
                                }
                                else
                                {
                                    ErrorMessages.Add("(" + ErrorMessages.Count + ") 請購部門不能為空!!<br>");
                                }
                                #endregion

                                #region //單據日期
                                string DocDate = "";
                                if (item["DocDate"] != null)
                                {
                                    if (item["DocDate"].Type.ToString() != "Null")
                                    {
                                        DocDate = item["DocDate"].ToString();

                                        DateTime parsedDate;
                                        bool docDateIsValid = DateTime.TryParseExact(
                                            item["DocDate"].ToString(),
                                            "yyyy-MM-dd",
                                            CultureInfo.InvariantCulture,
                                            DateTimeStyles.None,
                                            out parsedDate
                                        );
                                        if (!docDateIsValid)
                                        {
                                            ErrorMessages.Add("(" + ErrorMessages.Count + ") 【單據日期】格式錯誤!!<br>");
                                        }
                                    }
                                    else
                                    {
                                        ErrorMessages.Add("(" + ErrorMessages.Count + ") 單據日期不能為空!!<br>");
                                    }
                                }
                                else
                                {
                                    ErrorMessages.Add("(" + ErrorMessages.Count + ") 單據日期不能為空!!<br>");
                                }
                                #endregion

                                #region //需求日期
                                string DemandDate = "";
                                if (item["DemandDate"] != null)
                                {
                                    if (item["DemandDate"].Type.ToString() != "Null")
                                    {
                                        DemandDate = item["DemandDate"].ToString();

                                        DateTime parsedDate;
                                        bool demandDateIsValid = DateTime.TryParseExact(
                                            item["DemandDate"].ToString(),
                                            "yyyy-MM-dd",
                                            CultureInfo.InvariantCulture,
                                            DateTimeStyles.None,
                                            out parsedDate
                                        );
                                        if (!demandDateIsValid)
                                        {
                                            ErrorMessages.Add("(" + ErrorMessages.Count + ") 【需求日期】格式錯誤!!<br>");
                                        }
                                    }
                                    else
                                    {
                                        ErrorMessages.Add("(" + ErrorMessages.Count + ") 需求日期不能為空!!<br>");
                                    }
                                }
                                else
                                {
                                    ErrorMessages.Add("(" + ErrorMessages.Count + ") 需求日期不能為空!!<br>");
                                }
                                #endregion

                                #region //品號
                                string MtlItemNo = "";
                                int MtlItemId = -1;
                                string MtlItemName = "";
                                string MtlItemSpec = "";
                                int UomId = -1;
                                if (item["MtlItemNo"] != null)
                                {
                                    if (item["MtlItemNo"].Type.ToString() != "Null")
                                    {
                                        MtlItemNo = item["MtlItemNo"].ToString();

                                        #region //判斷品號資料是否正確
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT a.MtlItemId, a.MtlItemNo, a.MtlItemName, a.MtlItemSpec
                                                , a.InventoryUomId UomId
                                                FROM PDM.MtlItem a
                                                WHERE a.MtlItemNo = @MtlItemNo
                                                AND a.CompanyId = @CompanyId";
                                        dynamicParameters.Add("MtlItemNo", item["MtlItemNo"].ToString());
                                        dynamicParameters.Add("CompanyId", CurrentCompany);

                                        var MtlItemResult = sqlConnection.Query(sql, dynamicParameters);

                                        if (MtlItemResult.Count() <= 0)
                                        {
                                            ErrorMessages.Add("(" + ErrorMessages.Count + ") 查無此品號資料【" + item["MtlItemNo"].ToString() + "】!!<br>");
                                        }

                                        foreach (var item2 in MtlItemResult)
                                        {
                                            MtlItemId = item2.MtlItemId;
                                            MtlItemNo = item2.MtlItemNo;
                                            MtlItemName = item2.MtlItemName;
                                            MtlItemSpec = item2.MtlItemSpec;
                                            UomId = item2.UomId;
                                        }
                                        #endregion

                                        #region //判斷ERP品號生效日與失效日
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT LTRIM(RTRIM(MB030)) MB030, LTRIM(RTRIM(MB031)) MB031
                                                FROM INVMB
                                                WHERE MB001 = @MB001";
                                        dynamicParameters.Add("MB001", MtlItemNo);

                                        var INVMBResult = sqlConnection2.Query(sql, dynamicParameters);

                                        if (INVMBResult.Count() <= 0)
                                        {
                                            ErrorMessages.Add("(" + ErrorMessages.Count + ") 品號【" + item["MtlItemNo"].ToString() + "】不存在於ERP中!!<br>");
                                        }

                                        foreach (var item2 in INVMBResult)
                                        {
                                            if (item2.MB030 != "" && item2.MB030 != null)
                                            {
                                                #region //判斷日期需大於或等於生效日
                                                string EffectiveDate = item2.MB030;
                                                string effYear = EffectiveDate.Substring(0, 4);
                                                string effMonth = EffectiveDate.Substring(4, 2);
                                                string effDay = EffectiveDate.Substring(6, 2);
                                                DateTime effFullDate = Convert.ToDateTime(effYear + "-" + effMonth + "-" + effDay);
                                                int effresult = DateTime.Compare(CreateDate, effFullDate);
                                                if (effresult < 0)
                                                {
                                                    ErrorMessages.Add("(" + ErrorMessages.Count + ") 品號尚未生效，無法使用!!<br>");
                                                }
                                                #endregion
                                            }

                                            if (item2.MB031 != "" && item2.MB031 != null)
                                            {
                                                #region //判斷日期需小於或等於失效日
                                                string ExpirationDate = item2.MB031;
                                                string effYear = ExpirationDate.Substring(0, 4);
                                                string effMonth = ExpirationDate.Substring(4, 2);
                                                string effDay = ExpirationDate.Substring(6, 2);
                                                DateTime effFullDate = Convert.ToDateTime(effYear + "-" + effMonth + "-" + effDay);
                                                int effresult = DateTime.Compare(CreateDate, effFullDate);
                                                if (effresult > 0)
                                                {
                                                    ErrorMessages.Add("(" + ErrorMessages.Count + ") 品號已失效，無法使用!!<br>");
                                                }
                                                #endregion
                                            }
                                        }
                                        #endregion
                                    }
                                    else
                                    {
                                        ErrorMessages.Add("(" + ErrorMessages.Count + ") 品號不能為空!!<br>");
                                    }
                                }
                                else
                                {
                                    ErrorMessages.Add("(" + ErrorMessages.Count + ") 品號不能為空!!<br>");
                                }
                                #endregion

                                #region //庫別
                                string InventoryNo = "";
                                int InventoryId = -1;
                                string InventoryName = "";
                                double InventoryQty = 0;
                                string mtlInventory = "目前尚無資料";
                                if (item["InventoryNo"] != null)
                                {
                                    if (item["InventoryNo"].Type.ToString() != "Null")
                                    {
                                        InventoryNo = item["InventoryNo"].ToString();

                                        #region //判斷庫別資料是否正確
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT a.InventoryId, a.InventoryNo, a.InventoryName
                                                FROM SCM.Inventory a
                                                WHERE a.InventoryNo = @InventoryNo
                                                AND a.CompanyId = @CompanyId";
                                        dynamicParameters.Add("InventoryNo", InventoryNo);
                                        dynamicParameters.Add("CompanyId", CurrentCompany);

                                        var InventoryResult = sqlConnection.Query(sql, dynamicParameters);

                                        if (InventoryResult.Count() <= 0)
                                        {
                                            ErrorMessages.Add("(" + ErrorMessages.Count + ") 【庫別】資料錯誤!!<br>");
                                        }
                                        
                                        foreach (var item2 in InventoryResult)
                                        {
                                            InventoryId = item2.InventoryId;
                                            InventoryNo = item2.InventoryNo;
                                            InventoryName = item2.InventoryName;
                                        }
                                        #endregion

                                        #region //取得ERP庫存資訊
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT LTRIM(RTRIM(a.MC007)) InventoryQty
                                                FROM INVMC a
                                                WHERE a.MC001 = @MC001
                                                AND a.MC002 = @MC002";
                                        dynamicParameters.Add("MC001", MtlItemNo);
                                        dynamicParameters.Add("MC002", InventoryNo);

                                        var INVMCResult = sqlConnection2.Query(sql, dynamicParameters);

                                        if (INVMCResult.Count() > 0)
                                        {
                                            foreach (var item2 in INVMCResult)
                                            {
                                                InventoryQty = Convert.ToDouble(item2.InventoryQty);
                                                #region //組MtlInventory
                                                List<MtlInventory> mtlInventories = new List<MtlInventory>
                                                {
                                                    new MtlInventory
                                                    {
                                                        WAREHOUSE_NO = InventoryNo,
                                                        WAREHOUSE_NAME = InventoryName,
                                                        WAREHOUSE_QTY = InventoryQty
                                                    }
                                                };
                                                #endregion

                                                mtlInventory = JsonConvert.SerializeObject(mtlInventories);
                                                mtlInventory = "{\"data\":" + mtlInventory + "}";
                                            }
                                        }
                                        #endregion
                                    }
                                    else
                                    {
                                        ErrorMessages.Add("(" + ErrorMessages.Count + ") 庫別不能為空!!<br>");
                                    }
                                }
                                else
                                {
                                    ErrorMessages.Add("(" + ErrorMessages.Count + ") 庫別不能為空!!<br>");
                                }
                                #endregion

                                #region //請購數量
                                double PrQty = 0;
                                if (item["PrQty"] != null)
                                {
                                    if (item["PrQty"].Type.ToString() != "Null")
                                    {
                                        if (regex.IsMatch(item["PrQty"].ToString()))
                                        {
                                            PrQty = Convert.ToDouble(item["PrQty"]);

                                            if (PrQty <= 0)
                                            {
                                                continue;
                                            }
                                        }
                                        else
                                        {
                                            ErrorMessages.Add("(" + ErrorMessages.Count + ") 請購數量格式錯誤!!<br>");
                                        }
                                    }
                                    else
                                    {
                                        ErrorMessages.Add("(" + ErrorMessages.Count + ") 請購數量不能為空!!<br>");
                                    }
                                }
                                else
                                {
                                    ErrorMessages.Add("(" + ErrorMessages.Count + ") 請購數量不能為空!!<br>");
                                }
                                #endregion

                                #region //請購單價
                                double PrUnitPrice = 0;
                                if (item["PrUnitPrice"] != null)
                                {
                                    if (item["PrUnitPrice"].Type.ToString() != "Null")
                                    {
                                        if (regex.IsMatch(item["PrUnitPrice"].ToString()))
                                        {
                                            PrUnitPrice = Convert.ToDouble(item["PrUnitPrice"]);
                                        }
                                        else
                                        {
                                            ErrorMessages.Add("(" + ErrorMessages.Count + ") 請購單價格式錯誤!!<br>");
                                        }
                                    }
                                    else
                                    {
                                        ErrorMessages.Add("(" + ErrorMessages.Count + ") 請購單價不能為空!!<br>");
                                    }
                                }
                                else
                                {
                                    ErrorMessages.Add("(" + ErrorMessages.Count + ") 請購單價不能為空!!<br>");
                                }
                                #endregion

                                #region //參考訂單
                                string SoErpFullNo = "";
                                int SoDetailId = -1;
                                if (item["SoErpFullNo"] != null)
                                {
                                    if (item["SoErpFullNo"].Type.ToString() != "Null")
                                    {
                                        SoErpFullNo = item["SoErpFullNo"].ToString();

                                        if (SoErpFullNo.IndexOf("(") == -1)
                                        {
                                            ErrorMessages.Add("(" + ErrorMessages.Count + ") 訂單格式錯誤!!<br>");
                                        }
                                        else
                                        {
                                            #region //確認訂單資料是否正確
                                            string soErpPrefix = SoErpFullNo.Split('-')[0];
                                            string soErpNo = SoErpFullNo.Split('-')[1].Split('(')[0];
                                            string soSequence = SoErpFullNo.Split('-')[1].Split('(')[1].Split(')')[0];

                                            dynamicParameters = new DynamicParameters();
                                            sql = @"SELECT a.SoDetailId
                                                    FROM SCM.SoDetail a 
                                                    INNER JOIN SCM.SaleOrder b ON a.SoId = b.SoId
                                                    WHERE b.SoErpPrefix = @SoErpPrefix
                                                    AND b.SoErpNo = @SoErpNo 
                                                    AND a.SoSequence = @SoSequence
                                                    AND b.CompanyId = @CompanyId 
                                                    AND b.ConfirmStatus = 'Y'";
                                            dynamicParameters.Add("SoErpPrefix", soErpPrefix);
                                            dynamicParameters.Add("SoErpNo", soErpNo);
                                            dynamicParameters.Add("SoSequence", soSequence);
                                            dynamicParameters.Add("CompanyId", CurrentCompany);

                                            var SoDetailResult = sqlConnection.Query(sql, dynamicParameters);

                                            if (SoDetailResult.Count() <= 0) ErrorMessages.Add("(" + ErrorMessages.Count + ") 查無訂單資料!!<br>");

                                            foreach (var item2 in SoDetailResult)
                                            {
                                                SoDetailId = item2.SoDetailId;
                                            }
                                            #endregion
                                        }
                                    }
                                }
                                #endregion

                                #region //供應商
                                string SupplierNo = "";
                                int SupplierId = -1;
                                string TaxNo = null;
                                string Taxation = null;
                                string TradeTerm = null;
                                double? BusinessTaxRate = null;
                                string HideSupplier = "N";
                                string PrCurrency = "";
                                if (item["SupplierNo"] != null)
                                {
                                    if (item["SupplierNo"].Type.ToString() != "Null")
                                    {
                                        SupplierNo = item["SupplierNo"].ToString();

                                        #region //確認供應商資料是否正確
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT a.SupplierId, a.TaxNo, a.Taxation, a.TradeTerm, a.HideSupplier, a.Currency
                                                FROM SCM.Supplier a
                                                WHERE a.SupplierNo = @SupplierNo
                                                AND a.CompanyId = @CompanyId";
                                        dynamicParameters.Add("SupplierNo", SupplierNo);
                                        dynamicParameters.Add("CompanyId", CurrentCompany);

                                        var SupplierResult = sqlConnection.Query(sql, dynamicParameters);

                                        if (SupplierResult.Count() <= 0)
                                        {
                                            ErrorMessages.Add("(" + ErrorMessages.Count + ") 【供應商】資料錯誤!!<br>");
                                        }

                                        foreach (var item2 in SupplierResult)
                                        {
                                            TaxNo = item2.TaxNo;
                                            Taxation = item2.Taxation;
                                            TradeTerm = item2.TradeTerm;
                                            HideSupplier = item2.HideSupplier;
                                            PrCurrency = item2.Currency;
                                            SupplierId = item2.SupplierId;
                                        }
                                        #endregion

                                        #region //查詢營業稅額資料
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT ROUND(LTRIM(RTRIM(NN004)),2) BusinessTaxRate
                                                FROM CMSNN 
                                                WHERE NN001 = @TaxNo";
                                        dynamicParameters.Add("TaxNo", TaxNo);

                                        var businessTaxRateResult = sqlConnection2.Query(sql, dynamicParameters);

                                        if (businessTaxRateResult.Count() <= 0)
                                        {
                                            ErrorMessages.Add("(" + ErrorMessages.Count + ") 【課稅別】資料錯誤!!<br>");
                                        }

                                        foreach (var item2 in businessTaxRateResult)
                                        {
                                            BusinessTaxRate = item2.BusinessTaxRate;
                                        }
                                        #endregion
                                    }
                                }
                                #endregion

                                #region //取得本幣資料
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 MA003 FROM CMSMA";
                                var CMSMAResult = sqlConnection2.Query(sql, dynamicParameters);

                                if (CMSMAResult.Count() <= 0)
                                {
                                    ErrorMessages.Add("(" + ErrorMessages.Count + ") 取得本幣資料錯誤!!<br>");
                                }
                                
                                foreach (var item2 in CMSMAResult)
                                {
                                    if (PrCurrency.Length <= 0)
                                    {
                                        PrCurrency = item2.MA003;
                                    }
                                }
                                #endregion

                                #region //請購幣別
                                if (item["PrCurrency"] != null)
                                {
                                    if (item["PrCurrency"].Type.ToString() != "Null")
                                    {
                                        PrCurrency = item["PrCurrency"].ToString();

                                        #region //確認幣別資料是否正確
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT LTRIM(RTRIM(MF001)) Currency, LTRIM(RTRIM(MF002)) CurrencyName
                                                FROM CMSMF
                                                WHERE LTRIM(RTRIM(MF001)) = @PrCurrency";
                                        dynamicParameters.Add("PrCurrency", PrCurrency);

                                        var CMSMFResult = sqlConnection2.Query(sql, dynamicParameters);

                                        if (CMSMFResult.Count() <= 0) ErrorMessages.Add("(" + ErrorMessages.Count + ") 查無幣別資料!!<br>");
                                        #endregion
                                    }
                                }
                                #endregion

                                #region //取得請購匯率
                                string Today = DateTime.Now.ToString("yyyyMMdd");
                                sql = @"SELECT TOP 1 ROUND(LTRIM(RTRIM(MG005)),3) MG005
                                        FROM CMSMG 
                                        WHERE MG001 = @MG001
                                        AND MG002 <= @Today
                                        ORDER BY MG002 DESC";
                                dynamicParameters.Add("MG001", PrCurrency);
                                dynamicParameters.Add("Today", Today);

                                var CMSMGResult = sqlConnection2.Query(sql, dynamicParameters);

                                if (CMSMGResult.Count() <= 0)
                                {
                                    ErrorMessages.Add("(" + ErrorMessages.Count + ") 本幣匯率資料錯誤!!<br>");
                                }

                                double Exchange = 0;
                                foreach (var item2 in CMSMGResult)
                                {
                                    Exchange = item2.MG005;
                                }
                                #endregion
                                #endregion

                                #region //建立請購單頭
                                if (mrBool == true || count == 1)
                                {
                                    InsertPurchaseRequisition();
                                }
                                #endregion

                                #region //單身資料驗證
                                #region //確認品號是否需要上傳相關文件才能進行請購(除晶彩外)
                                if (CurrentCompany == 11)
                                {
                                    string pattern = @"^.*-T$";
                                    if (Regex.IsMatch(MtlItemNo, pattern))
                                    {
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT TOP 1 1
                                                FROM PDM.MtlItem a 
                                                OUTER APPLY (
                                                    SELECT x.MtDocId, x.DocName
                                                    FROM PDM.MtlItemDocSetting x 
                                                    WHERE x.PurchaseMandatory = 'Y'
                                                ) x
                                                LEFT JOIN PDM.MtlItemFile b ON a.MtlItemNo = b.MtlItemNo AND b.MtDocId = x.MtDocId
                                                WHERE a.MtlItemNo = @MtlItemNo
                                                AND b.MtDocId IS NULL";
                                            dynamicParameters.Add("MtlItemNo", MtlItemNo);

                                        var MtlItemFileResult = sqlConnection.Query(sql, dynamicParameters);

                                        if (MtlItemFileResult.Count() > 0)
                                        {
                                            ErrorMessages.Add("(" + ErrorMessages.Count + ")品號【" + MtlItemNo + "】 需上傳相關文件才可進行請購!!<br>");
                                        }
                                    }
                                }
                                #endregion

                                #region //確認集團內/集團外邏輯
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT TOP 1 b.HideSupplier
                                        FROM SCM.PrDetail a 
                                        INNER JOIN SCM.Supplier b ON a.SupplierId = b.SupplierId
                                        WHERE a.PrId = @PrId
                                        ORDER BY a.PrSequence";
                                dynamicParameters.Add("PrId", PrId);

                                var HideSupplierResult = sqlConnection.Query(sql, dynamicParameters);

                                foreach (var item2 in HideSupplierResult)
                                {
                                    if (item2.HideSupplier != HideSupplier)
                                    {
                                        ErrorMessages.Add("(" + ErrorMessages.Count + ") 同一批請購單只能統一集團內或集團外!!<br>");
                                    }
                                }
                                #endregion

                                #region //檢核特定相關卡控(目前僅晶彩邏輯)
                                if (CurrentCompany == 4)
                                {
                                    double localPrUnitPrice = PrUnitPrice * Exchange;

                                    #region //單別3109、3108特別卡控
                                    if (PrErpPrefix == "3108" || PrErpPrefix == "3109")
                                    {
                                        #region //請購單價不能超過1000(本幣)
                                        if (localPrUnitPrice > 1000)
                                        {
                                            ErrorMessages.Add("(" + ErrorMessages.Count + ") 違反請購單別【3108、3109】卡控:<br>請購單價不能超過【1000】!!<br>此次請購單價【" + localPrUnitPrice + "】已超過!!<br>");
                                        }
                                        #endregion

                                        #region //所有單身合計金額(本幣)不能超過20000
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"SELECT ISNULL(SUM(PrPriceTw), 0) TotalPrice
                                                FROM SCM.PrDetail
                                                WHERE PrId = @PrId";
                                        dynamicParameters.Add("PrId", PrId);

                                        var TotalPriceResult = sqlConnection.Query(sql, dynamicParameters);

                                        double totalPrice = 0;
                                        foreach (var item2 in TotalPriceResult)
                                        {
                                            totalPrice = item2.TotalPrice;
                                            double PrPriceTw = PrUnitPrice * PrQty * Exchange;
                                            if (PrPriceTw > 20000)
                                            {
                                                ErrorMessages.Add("(" + ErrorMessages.Count + ") 違反請購單別【3108、3109】卡控:<br>單身合計金額(本幣)不能超過【20000】!!<br>目前金額【" + PrPriceTw + "】已超過!!<br>");
                                            }
                                            else
                                            {
                                                if (totalPrice + PrPriceTw > 20000)
                                                {
                                                    InsertPurchaseRequisition();
                                                }
                                            }
                                        }
                                        #endregion
                                    }
                                    #endregion
                                }
                                #endregion

                                if (ErrorMessages.Count > 1)
                                {
                                    ErrorMessages.Add("=======================<br>");
                                    TotalErrorMessages.AddRange(ErrorMessages);
                                    continue;
                                }
                                #endregion

                                #region //新增請購單身段
                                #region //取得目前此請購單身序號
                                string PrSequence = "";
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT MAX(a.PrSequence) PrSequence
                                        FROM SCM.PrDetail a 
                                        WHERE a.PrId = @PrId";
                                dynamicParameters.Add("PrId", PrId);

                                var PrDetailResult = sqlConnection.Query(sql, dynamicParameters);

                                int maxSeq = 0;
                                if (PrDetailResult.Count() <= 0)
                                {
                                    maxSeq = 1;
                                }
                                else
                                {
                                    foreach (var item2 in PrDetailResult)
                                    {
                                        maxSeq = Convert.ToInt32(item2.PrSequence) + 1;
                                    }
                                }

                                PrSequence = maxSeq.ToString("D4");
                                #endregion

                                #region //INSERT SCM.PrDetail
                                dynamicParameters = new DynamicParameters();
                                sql = @"INSERT INTO SCM.PrDetail (PrId, PrSequence, MtlItemId, PrMtlItemName, PrMtlItemSpec, InventoryId, PrUomId, PrQty, DemandDate, SupplierId, PrCurrency, PrExchangeRate, PrUnitPrice, PrPrice, PrPriceTw, UrgentMtl, ProductionPlan, Project, SoDetailId
                                        , PoUserId, PoUomId, PoQty, PoCurrency, PoUnitPrice, PoPrice, LockStaus, PoStaus, PartialPurchaseStaus, InquiryStatus, TaxNo, Taxation, BusinessTaxRate, DetailMultiTax, TradeTerm, PrPriceQty, PrPriceUomId, MtlInventory, MtlInventoryQty, ConfirmStatus, ClosureStatus, PrDetailRemark, PoRemark
                                        , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                        OUTPUT INSERTED.PrDetailId
                                        VALUES (@PrId, @PrSequence, @MtlItemId, @PrMtlItemName, @PrMtlItemSpec, @InventoryId, @PrUomId, @PrQty, @DemandDate, @SupplierId, @PrCurrency, @PrExchangeRate, @PrUnitPrice, @PrPrice, @PrPriceTw, @UrgentMtl, @ProductionPlan, @Project, @SoDetailId
                                        , @PoUserId, @PoUomId, @PoQty, @PoCurrency, @PoUnitPrice, @PoPrice, @LockStaus, @PoStaus, @PartialPurchaseStaus, @InquiryStatus, @TaxNo, @Taxation, @BusinessTaxRate, @DetailMultiTax, @TradeTerm, @PrPriceQty, @PrPriceUomId, @MtlInventory, @MtlInventoryQty, @ConfirmStatus, @ClosureStatus, @PrDetailRemark, @PoRemark
                                        , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        PrId,
                                        PrSequence,
                                        MtlItemId,
                                        PrMtlItemName = MtlItemName,
                                        PrMtlItemSpec = MtlItemSpec,
                                        InventoryId,
                                        PrUomId = UomId,
                                        PrQty,
                                        DemandDate,
                                        SupplierId,
                                        PrCurrency,
                                        PrExchangeRate = Exchange,
                                        PrUnitPrice,
                                        PrPrice = PrUnitPrice * PrQty,
                                        PrPriceTw = PrUnitPrice * PrQty * Exchange,
                                        UrgentMtl = "N",
                                        ProductionPlan = "N",
                                        Project = "",
                                        SoDetailId = SoDetailId > 0 ? SoDetailId : (int?)null,
                                        PoUserId = (int?)null,
                                        PoUomId = UomId,
                                        PoQty = PrQty,
                                        PoCurrency = PrCurrency,
                                        PoUnitPrice = PrUnitPrice,
                                        PoPrice = PrUnitPrice * PrQty,
                                        LockStaus = "N",
                                        PoStaus = "N",
                                        PartialPurchaseStaus = "N",
                                        InquiryStatus = "1",
                                        TaxNo = TaxNo != "" ? TaxNo : null,
                                        Taxation = Taxation != "" ? Taxation : null,
                                        BusinessTaxRate = BusinessTaxRate > 0 ? BusinessTaxRate : (double?)null,
                                        DetailMultiTax = "N",
                                        TradeTerm = TradeTerm != "" ? TradeTerm : null,
                                        PrPriceQty = PrQty,
                                        PrPriceUomId = UomId,
                                        MtlInventory = mtlInventory,
                                        MtlInventoryQty = InventoryQty,
                                        ConfirmStatus = "N",
                                        ClosureStatus = "N",
                                        PrDetailRemark = "",
                                        PoRemark = "",
                                        CreateDate,
                                        LastModifiedDate,
                                        CreateBy,
                                        LastModifiedBy
                                    });

                                var insertResult = sqlConnection.Query(sql, dynamicParameters);

                                rowsAffected += insertResult.Count();

                                int PrDetailId = -1;
                                foreach (var item2 in insertResult)
                                {
                                    PrDetailId = item2.PrDetailId;
                                }
                                #endregion
                                #endregion

                                #region //Insert PurchaseRequisition Fun
                                void InsertPurchaseRequisition()
                                {
                                    #region //確認此請購單號是否已存在
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT TOP 1 1
                                            FROM SCM.PurchaseRequisition a 
                                            WHERE a.PrErpPrefix = @PrErpPrefix
                                            AND a.PrErpNo = @PrErpNo";
                                    dynamicParameters.Add("PrErpPrefix", PrErpPrefix);
                                    dynamicParameters.Add("PrErpNo", PrErpNo);

                                    var PurchaseRequisitionResult2 = sqlConnection.Query(sql, dynamicParameters);

                                    if (PurchaseRequisitionResult2.Count() > 0)
                                    {
                                        ErrorMessages.Add("(" + ErrorMessages.Count + ") 【請購單號】已存在，無法重複新增!!<br>");
                                    }
                                    #endregion

                                    #region //取號PrNo
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT ISNULL(MAX(a.PrNo), '000') MaxPrNo
                                            FROM SCM.PurchaseRequisition a
                                            WHERE FORMAT(a.CreateDate, 'yyyy-MM-dd') = @CreateDate
                                            AND a.CompanyId = @CurrentCompany";
                                    dynamicParameters.Add("CreateDate", DateTime.Now.ToString("yyyy-MM-dd"));
                                    dynamicParameters.Add("CurrentCompany", CurrentCompany);

                                    var PrNoResult = sqlConnection.Query(sql, dynamicParameters);

                                    string PrNo = "";
                                    foreach (var item2 in PrNoResult)
                                    {
                                        string PrDocDate = DateTime.Now.ToString("yyyyMMdd");
                                        string MaxPrNo = item2.MaxPrNo;
                                        string MaxNo = MaxPrNo.Substring(MaxPrNo.Length - 3);
                                        PrNo = "PR" + PrDocDate + (Convert.ToInt32(MaxNo) + 1).ToString().PadLeft(3, '0');
                                    }
                                    #endregion

                                    #region //比對ERP關帳日期(庫存關帳)
                                    dynamicParameters = new DynamicParameters();
                                    sql = @"SELECT LTRIM(RTRIM(MA012)) MA012
                                            FROM CMSMA";
                                    var cmsmaResult = sqlConnection2.Query(sql, dynamicParameters);
                                    if (cmsmaResult.Count() < 0)
                                    {
                                        throw new SystemException("EPR關帳日期資料錯誤!!");
                                    }

                                    foreach (var item2 in cmsmaResult)
                                    {
                                        string eprDate = item2.MA012;
                                        string erpYear = eprDate.Substring(0, 4);
                                        string erpMonth = eprDate.Substring(4, 2);
                                        string erpFullDate = erpYear + "-" + erpMonth;
                                        DateTime erpDateTime = Convert.ToDateTime(erpFullDate);
                                        DateTime DocDateDateTime = Convert.ToDateTime(DocDate.Split('-')[0] + "-" + DocDate.Split('-')[1]);
                                        int compare = DocDateDateTime.CompareTo(erpDateTime);
                                        if (compare <= 0)
                                        {
                                            ErrorMessages.Add("(" + ErrorMessages.Count + ") ERP已關帳(" + erpFullDate + ")，無法開立此日期(" + DocDate + ")之單據!!<br>");
                                        }
                                    }
                                    #endregion

                                    #region //INSERT SCM.PurchaseRequisition
                                    if (ErrorMessages.Count == 1)
                                    {
                                        dynamicParameters = new DynamicParameters();
                                        sql = @"INSERT INTO SCM.PurchaseRequisition (PrNo, CompanyId, DepartmentId, UserId, PrErpPrefix, PrErpNo, Edition, PrDate, DocDate, PrRemark
                                                , TotalQty, Amount, PrStatus, SignupStaus, LockStaus, BomType, ConfirmStatus, BpmTransferStatus, TransferStatus, Priority, Source
                                                , CreateDate, LastModifiedDate, CreateBy, LastModifiedBy)
                                                OUTPUT INSERTED.PrId, INSERTED.PrErpPrefix, INSERTED.PrErpNo, INSERTED.PrNo
                                                VALUES (@PrNo, @CompanyId, @DepartmentId, @UserId, @PrErpPrefix, @PrErpNo, @Edition, @PrDate, @DocDate, @PrRemark
                                                , @TotalQty, @Amount, @PrStatus, @SignupStaus, @LockStaus, @BomType, @ConfirmStatus, @BpmTransferStatus, @TransferStatus, @Priority, @Source
                                                , @CreateDate, @LastModifiedDate, @CreateBy, @LastModifiedBy)";
                                        dynamicParameters.AddDynamicParams(
                                            new
                                            {
                                                PrNo,
                                                CompanyId = CurrentCompany,
                                                DepartmentId,
                                                UserId = CurrentUser,
                                                PrErpPrefix,
                                                PrErpNo,
                                                Edition = "0000",
                                                PrDate = DocDate,
                                                DocDate,
                                                PrRemark = "",
                                                TotalQty = 0,
                                                Amount = 0,
                                                PrStatus = "N",
                                                SignupStaus = "0",
                                                LockStaus = "N",
                                                BomType = "Y",
                                                ConfirmStatus = "N",
                                                BpmTransferStatus = "N",
                                                TransferStatus = "N",
                                                Priority = "2",
                                                Source = "/PurchaseRequisition/BatchAddPurchaseRequisition",
                                                CreateDate,
                                                LastModifiedDate,
                                                CreateBy,
                                                LastModifiedBy
                                            });

                                        var insertResult2 = sqlConnection.Query(sql, dynamicParameters);

                                        rowsAffected += insertResult2.Count();

                                        foreach (var item2 in insertResult2)
                                        {
                                            PrId = item2.PrId;
                                            PrIdList.Add(PrId);
                                        }
                                    }
                                    #endregion
                                }
                                #endregion

                                count++;
                            }

                            if (TotalErrorMessages.Count > 0)
                            {
                                string errorMessages = "";
                                foreach (var msg in TotalErrorMessages)
                                {
                                    errorMessages += msg;
                                }
                                throw new SystemException(errorMessages);
                            }

                            #region //更新單頭總採購數量及金額
                            foreach (var prId in PrIdList)
                            {
                                #region //取得目前請購單身總數量、金額
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT ISNULL(SUM(a.PrQty), 0) TotalQty,  ISNULL(SUM(a.PrPriceTw), 0) TotalPrPriceTw
                                        FROM SCM.PrDetail a 
                                        WHERE a.PrId = @PrId";
                                dynamicParameters.Add("PrId", prId);

                                var PrDetailResult = sqlConnection.Query(sql, dynamicParameters);

                                double TotalQty = 0;
                                double Amount = 0;
                                foreach (var item in PrDetailResult)
                                {
                                    TotalQty = item.TotalQty;
                                    Amount = item.TotalPrPriceTw;
                                }
                                #endregion

                                #region //Update SCM.PurchaseRequisition
                                dynamicParameters = new DynamicParameters();
                                sql = @"UPDATE SCM.PurchaseRequisition SET
                                        TotalQty = @TotalQty,
                                        Amount = @Amount,
                                        LastModifiedDate = @LastModifiedDate,
                                        LastModifiedBy = @LastModifiedBy
                                        WHERE PrId = @PrId";
                                dynamicParameters.AddDynamicParams(
                                    new
                                    {
                                        TotalQty,
                                        Amount,
                                        LastModifiedDate,
                                        LastModifiedBy,
                                        prId
                                    });

                                rowsAffected = sqlConnection.Execute(sql, dynamicParameters);
                                #endregion
                            }
                            #endregion

                            #region //取得目前請購單資料
                            List<BatchPurchaseRequisition> batchPurchaseRequisitions = new List<BatchPurchaseRequisition>();
                            if (PrIdList.Count > 0)
                            {
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT a.PrDetailId, a.PrQty, a.PrUnitPrice, a.PrPrice, a.PrCurrency
                                        , FORMAT(a.DemandDate, 'yyyy-MM-dd') DemandDate
                                        , b.PrErpPrefix, b.PrErpNo, FORMAT(b.DocDate, 'yyyy-MM-dd') DocDate
                                        , c.DepartmentNo + ' ' + c.DepartmentName DepartmentFullNo
                                        , d.MtlItemId, d.MtlItemNo, d.MtlItemName, d.MtlItemSpec
                                        , e.InventoryId, e.InventoryNo, e.InventoryName
                                        , f.UomId, f.UomNo, f.UomName
                                        , h.SoErpPrefix + '-' + h.SoErpNo + '(' + g.SoSequence + ')' SoErpFullNo
                                        , i.SupplierNo + ' ' + i.SupplierShortName SupplierFullNo
                                        FROM SCM.PrDetail a 
                                        INNER JOIN SCM.PurchaseRequisition b ON a.PrId = b.PrId
                                        INNER JOIN BAS.Department c ON b.DepartmentId = c.DepartmentId
                                        INNER JOIN PDM.MtlItem d ON a.MtlItemId = d.MtlItemId
                                        INNER JOIN SCM.Inventory e ON a.InventoryId = e.InventoryId
                                        INNER JOIN PDM.UnitOfMeasure f ON a.PrUomId = f.UomId
                                        LEFT JOIN SCM.SoDetail g ON a.SoDetailId = g.SoDetailId
                                        LEFT JOIN SCM.SaleOrder h ON g.SoId = h.SoId
                                        LEFT JOIN SCM.Supplier i ON a.SupplierId = i.SupplierId
                                        WHERE a.PrId IN (" + string.Join(",", PrIdList) + ")";

                                batchPurchaseRequisitions = sqlConnection.Query<BatchPurchaseRequisition>(sql, dynamicParameters).ToList();
                            }
                            #endregion

                            #region //取得EPR相關資料
                            int index = 0;
                            foreach (var item in batchPurchaseRequisitions)
                            {
                                #region //取得ERP庫存數
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT ISNULL(MC007, 0) MC007 
                                        FROM INVMC 
                                        WHERE MC001 = @MC001 
                                        AND MC002 = @MC002";
                                dynamicParameters.Add("MC001", item.MtlItemNo);
                                dynamicParameters.Add("MC002", item.InventoryNo);

                                var INVMCResult = sqlConnection2.Query(sql, dynamicParameters);

                                double? InventoryQty = 0;
                                foreach (var item2 in INVMCResult)
                                {
                                    InventoryQty = Convert.ToDouble(item2.MC007);
                                }

                                batchPurchaseRequisitions[index].InventoryQty = InventoryQty;
                                #endregion

                                #region //取得ERP在途採購數量
                                dynamicParameters = new DynamicParameters();
                                sql = @"SELECT ISNULL(SUM(TD008 - TD015), 0) PoQty
                                        FROM PURTD 
                                        WHERE TD004 = @MtlItemNo 
                                        AND TD016 = 'N'
                                        AND TD018 = 'Y'";
                                dynamicParameters.Add("MtlItemNo", item.MtlItemNo);

                                var PURTDResult = sqlConnection2.Query(sql, dynamicParameters);

                                double? PoQty = 0;
                                foreach (var item2 in PURTDResult)
                                {
                                    PoQty = Convert.ToDouble(item2.PoQty);
                                }

                                batchPurchaseRequisitions[index].PoQty = PoQty;
                                #endregion

                                index++;
                            }
                            #endregion

                            #region //Response
                            jsonResponse = JObject.FromObject(new
                            {
                                status = "success",
                                msg = "(" + rowsAffected + " rows affected)",
                                data = batchPurchaseRequisitions
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

        #region//Mail
        #region//請購單已確認但未採購的
        public string ApiPrConfirmedNotProcuredRecord(string CompanyNo)
        {
            try
            {
                if (CompanyNo.Length <= 0) throw new SystemException("【公司】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope(TransactionScopeOption.Required, new TimeSpan(0, 180, 0)))
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //信件通知
                        #region //Mail資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MailId,c.MailFrom, d.MailTo, a.MailCc, a.MailBcc, a.MailSubject, a.MailContent
                                , b.Host, b.Port, b.SendMode, b.Account, b.Password
                                FROM BAS.Mail a
                                LEFT JOIN BAS.MailServer b ON b.ServerId = a.ServerId
                                LEFT JOIN (
                                    SELECT a1.MailId,b1.ContactName+':'+b1.Email AS MailFrom
                                    FROM BAS.MailUser a1
                                        INNER JOIN BAS.MailContact b1 ON a1.ContactId=b1.ContactId
                                    WHERE a1.MailUserType='F'   
                                )c ON c.MailId =a.MailId
                                LEFT JOIN (
                                    SELECT x1.MailId,STRING_AGG(y1.UserName+':'+y1.Email,';') AS MailTo
                                    FROM BAS.MailUser x1
                                    INNER JOIN BAS.[User] y1 ON x1.UserId=y1.UserId
                                    WHERE x1.MailUserType='T' AND y1.UserStatus='F'
                                    AND x1.MailId IN (
                                        SELECT z.MailId
                                        FROM BAS.MailSendSetting z
                                        WHERE z.SettingSchema = @SettingSchema
                                        AND z.SettingNo =@SettingNo)
                                    GROUP BY x1.MailId
                                ) d ON d.MailId=a.MailId
                                WHERE a.MailId IN (
                                    SELECT z.MailId
                                    FROM BAS.MailSendSetting z
                                    WHERE z.SettingSchema = @SettingSchema
                                    AND z.SettingNo =@SettingNo
                            )";
                        dynamicParameters.Add("SettingSchema", "PrConfirmedNotProcuredRecord");
                        dynamicParameters.Add("SettingNo", "Y");

                        var resultMailTemplate = sqlConnection.Query(sql, dynamicParameters);
                        if (resultMailTemplate.Count() <= 0) throw new SystemException("Mail設定錯誤!");
                        #endregion

                        foreach (var item in resultMailTemplate)
                        {
                            string mailSubject = item.MailSubject,
                                mailContent = HttpUtility.UrlDecode(item.MailContent);

                            #region //Mail內容                            
                            //string hyperlink = "http://" + HttpContext.Current.Request.Url.Authority + "/PurchaseRequisition/PrConfirmedNotProcuredRecord";
                            //hyperlink = string.Format("<a href=\"{0}\">請點選</a>", hyperlink);
                            //mailSubject = mailSubject.Replace("[Date]", DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd"));
                            //mailContent = mailContent.Replace("[Date]", DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd"));
                            //mailContent = mailContent.Replace("[hyperLink]", hyperlink);
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

        #region//粗胚建議請購清單(無條件)
        public string ApiSuggestionsForPurchaseRecord(string CompanyNo)
        {
            try
            {
                if (CompanyNo.Length <= 0) throw new SystemException("【公司】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope(TransactionScopeOption.Required, new TimeSpan(0, 180, 0)))
                {
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //信件通知
                        #region //Mail資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MailId,c.MailFrom, d.MailTo, a.MailCc, a.MailBcc, a.MailSubject, a.MailContent
                                , b.Host, b.Port, b.SendMode, b.Account, b.Password
                                FROM BAS.Mail a
                                LEFT JOIN BAS.MailServer b ON b.ServerId = a.ServerId
                                LEFT JOIN (
                                    SELECT a1.MailId,b1.ContactName+':'+b1.Email AS MailFrom
                                    FROM BAS.MailUser a1
                                        INNER JOIN BAS.MailContact b1 ON a1.ContactId=b1.ContactId
                                    WHERE a1.MailUserType='F'   
                                )c ON c.MailId =a.MailId
                                LEFT JOIN (
                                    SELECT x1.MailId,STRING_AGG(y1.UserName+':'+y1.Email,';') AS MailTo
                                    FROM BAS.MailUser x1
                                    INNER JOIN BAS.[User] y1 ON x1.UserId=y1.UserId
                                    WHERE x1.MailUserType='T' AND y1.UserStatus='F'
                                    AND x1.MailId IN (
                                        SELECT z.MailId
                                        FROM BAS.MailSendSetting z
                                        WHERE z.SettingSchema = @SettingSchema
                                        AND z.SettingNo =@SettingNo)
                                    GROUP BY x1.MailId
                                ) d ON d.MailId=a.MailId
                                WHERE a.MailId IN (
                                    SELECT z.MailId
                                    FROM BAS.MailSendSetting z
                                    WHERE z.SettingSchema = @SettingSchema
                                    AND z.SettingNo =@SettingNo
                            )";
                        dynamicParameters.Add("SettingSchema", "SuggestionsForPurchaseRecord");
                        dynamicParameters.Add("SettingNo", "Y");

                        var resultMailTemplate = sqlConnection.Query(sql, dynamicParameters);
                        if (resultMailTemplate.Count() <= 0) throw new SystemException("Mail設定錯誤!");
                        #endregion

                        foreach (var item in resultMailTemplate)
                        {
                            string mailSubject = item.MailSubject,
                                mailContent = HttpUtility.UrlDecode(item.MailContent);

                            #region //Mail內容
                            //string queryCondition = "", MB005 = "101", MB007 = "325";
                            //string hyperlink = "http://" + HttpContext.Current.Request.Url.Authority + "/PurchaseRequisition/SuggestionsForPurchaseRecord";
                            //hyperlink = string.Format("<a href=\"{0}?queryCondition={1}&MB005={2}&MB007={3}\">請點選</a>", hyperlink, queryCondition, MB005, MB007);
                            //mailSubject = mailSubject.Replace("[Date]", DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd"));
                            //mailContent = mailContent.Replace("[Date]", DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd"));
                            //mailContent = mailContent.Replace("[hyperLink]", hyperlink);
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

        #region//粗胚建議請購清單(有條件)
        public string ApiSuggestionsForPurchaseConditionRecord(string CompanyNo)
        {
            try
            {
                if (CompanyNo.Length <= 0) throw new SystemException("【公司】不能為空!");

                using (TransactionScope transactionScope = new TransactionScope(TransactionScopeOption.Required, new TimeSpan(0, 180, 0)))
                {   
                    using (SqlConnection sqlConnection = new SqlConnection(MainConnectionStrings))
                    {
                        #region //信件通知
                        #region //Mail資料
                        dynamicParameters = new DynamicParameters();
                        sql = @"SELECT a.MailId,c.MailFrom, d.MailTo, a.MailCc, a.MailBcc, a.MailSubject, a.MailContent
                                , b.Host, b.Port, b.SendMode, b.Account, b.Password
                                FROM BAS.Mail a
                                LEFT JOIN BAS.MailServer b ON b.ServerId = a.ServerId
                                LEFT JOIN (
                                    SELECT a1.MailId,b1.ContactName+':'+b1.Email AS MailFrom
                                    FROM BAS.MailUser a1
                                        INNER JOIN BAS.MailContact b1 ON a1.ContactId=b1.ContactId
                                    WHERE a1.MailUserType='F'   
                                )c ON c.MailId =a.MailId
                                LEFT JOIN (
                                    SELECT x1.MailId,STRING_AGG(y1.UserName+':'+y1.Email,';') AS MailTo
                                    FROM BAS.MailUser x1
                                    INNER JOIN BAS.[User] y1 ON x1.UserId=y1.UserId
                                    WHERE x1.MailUserType='T' AND y1.UserStatus='F'
                                    AND x1.MailId IN (
                                        SELECT z.MailId
                                        FROM BAS.MailSendSetting z
                                        WHERE z.SettingSchema = @SettingSchema
                                        AND z.SettingNo =@SettingNo)
                                    GROUP BY x1.MailId
                                ) d ON d.MailId=a.MailId
                                WHERE a.MailId IN (
                                    SELECT z.MailId
                                    FROM BAS.MailSendSetting z
                                    WHERE z.SettingSchema = @SettingSchema
                                    AND z.SettingNo =@SettingNo
                            )";
                        dynamicParameters.Add("SettingSchema", "SuggestionsForPurchaseRecord");
                        dynamicParameters.Add("SettingNo", "Y");

                        var resultMailTemplate = sqlConnection.Query(sql, dynamicParameters);
                        if (resultMailTemplate.Count() <= 0) throw new SystemException("Mail設定錯誤!");
                        #endregion

                        foreach (var item in resultMailTemplate)
                        {
                            string mailSubject = item.MailSubject,
                                mailContent = HttpUtility.UrlDecode(item.MailContent);

                            #region //Mail內容
                            //string queryCondition = "MC004MC007", MB005 = "101", MB007 = "325";
                            //string hyperlink = "http://" + HttpContext.Current.Request.Url.Authority + "/PurchaseRequisition/SuggestionsForPurchaseRecord";
                            //hyperlink = string.Format("<a href=\"{0}?queryCondition={1}&MB005={2}&MB007={3}\">請點選</a>", hyperlink, queryCondition, MB005, MB007);
                            //mailSubject = mailSubject.Replace("[Date]", DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd"));
                            //mailContent = mailContent.Replace("[Date]", DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd"));
                            //mailContent = mailContent.Replace("[hyperLink]", hyperlink);
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
    }
}
